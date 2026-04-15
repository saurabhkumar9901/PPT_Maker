"""
Stage 4 -- Integration: FastMCP Server + Standalone CLI

Wraps the entire 3-stage pipeline (Ingest -> Orchestrate -> Compile) as
MCP tools for LLM clients (Claude Desktop, Cursor, etc.) and also provides
a standalone CLI mode for direct use.

MCP Tools exposed:
  - ingest_template(template_path) -> design_tokens.json
  - plan_presentation(markdown_path, tokens_path) -> ui_plan.json
  - compile_presentation(tokens_path, plan_path, template_path, output_path) -> .pptx
  - generate_presentation(markdown_path, template_path, output_path) -> full pipeline
"""

import os
import sys
import argparse
from dotenv import load_dotenv

load_dotenv()

# Import pipeline stages
from ingest import ingest as run_ingest
# Deferred import for orchestrator to avoid genai dependency locally
from compiler import compile_presentation as run_compile
from validator import validate_presentation, print_validation_report
from editor import get_inventory, print_inventory, replace_text, reorder_slides
from auto_fixer import run_fixes


# ─── Full Pipeline ────────────────────────────────────────────────────────────

def generate_presentation(markdown_path: str, template_path: str, output_path: str, api_key: str = None) -> str:
    """
    Full pipeline: Ingest -> Plan -> Compile -> .pptx
    """
    output_dir = os.path.dirname(output_path) or "output"
    os.makedirs(output_dir, exist_ok=True)

    # Stage 1: Ingest
    print("=" * 60)
    print("STAGE 1: INGESTION")
    print("=" * 60)
    tokens_path = run_ingest(template_path, output_dir)

    # Stage 2: Orchestrate
    print("\n" + "=" * 60)
    print("STAGE 2: ORCHESTRATION")
    print("=" * 60)
    plan_path = os.path.join(output_dir, "ui_plan.json")
    from orchestrator import plan_presentation as run_orchestrate
    run_orchestrate(markdown_path, tokens_path, plan_path, api_key)

    # Stage 3: Compile (now using python-pptx with native template)
    print("\n" + "=" * 60)
    print("STAGE 3: COMPILATION")
    print("=" * 60)
    run_compile(tokens_path, plan_path, template_path, output_path)

    # Stage 4: Auto-Fixer
    print("\n" + "=" * 60)
    print("STAGE 4: AUTO-FIXING")
    print("=" * 60)
    output_path = run_fixes(output_path, None, output_path)

    # Stage 5: Validate
    print("\n" + "=" * 60)
    print("STAGE 5: VALIDATION")
    print("=" * 60)
    
    with open(tokens_path, 'r', encoding='utf-8') as f:
        import json
        tokens = json.load(f)
        
    report = validate_presentation(output_path, tokens)
    print_validation_report(report)

    return f"Presentation generated at: {output_path}"


# ─── FastMCP Server ──────────────────────────────────────────────────────────

def start_mcp_server():
    """Start the FastMCP server with all tools exposed."""
    from fastmcp import FastMCP

    mcp = FastMCP("PPT Maker Pipeline Server")

    @mcp.tool()
    def ingest_template(template_path: str, output_dir: str = "output") -> str:
        """Stage 1: Extract design tokens (colors, fonts, backgrounds) from a PowerPoint slide master template."""
        try:
            tokens_path = run_ingest(template_path, output_dir)
            return f"Design tokens extracted to: {tokens_path}"
        except Exception as e:
            return f"Ingestion Error: {e}"

    @mcp.tool()
    def plan_presentation_tool(markdown_path: str, tokens_path: str, output_path: str = "output/ui_plan.json") -> str:
        """Stage 2: Send markdown content to Gemini to create a structured UI plan."""
        try:
            result = run_orchestrate(markdown_path, tokens_path, output_path)
            return f"UI plan saved to: {result}"
        except Exception as e:
            return f"Orchestration Error: {e}"

    @mcp.tool()
    def compile_presentation_tool(tokens_path: str, plan_path: str, template_path: str, output_path: str) -> str:
        """Stage 3: Compile design tokens + UI plan into a .pptx file using the native template."""
        try:
            run_compile(tokens_path, plan_path, template_path, output_path)
            return f"Presentation compiled to: {output_path}"
        except Exception as e:
            return f"Compilation Error: {e}"

    @mcp.tool()
    def generate_presentation_tool(markdown_path: str, template_path: str, output_path: str) -> str:
        """Full pipeline: Ingest template -> Plan with Gemini -> Compile .pptx. One-shot presentation generation."""
        try:
            return generate_presentation(markdown_path, template_path, output_path)
        except Exception as e:
            return f"Pipeline Error: {e}"

    print("Starting PPT Maker MCP Server...")
    mcp.run()


# ─── CLI Mode ─────────────────────────────────────────────────────────────────

def cli_mode():
    """Standalone CLI interface for direct pipeline execution."""
    parser = argparse.ArgumentParser(
        description="PPT Maker v2 -- Markdown to PowerPoint Pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Full pipeline:
    python server.py generate --markdown content.md --template master.pptx --output deck.pptx

  Individual stages:
    python server.py ingest --template master.pptx --output output/
    python server.py plan --markdown content.md --tokens output/design_tokens.json
    python server.py compile --tokens output/design_tokens.json --plan output/ui_plan.json --template master.pptx --output deck.pptx

  MCP Server:
    python server.py serve
        """
    )

    subparsers = parser.add_subparsers(dest="command", help="Pipeline command")

    # Full pipeline
    gen_parser = subparsers.add_parser("generate", help="Run full pipeline: Ingest -> Plan -> Compile")
    gen_parser.add_argument("--markdown", required=True, help="Path to input Markdown file")
    gen_parser.add_argument("--template", required=True, help="Path to Slide Master .pptx template")
    gen_parser.add_argument("--output", required=True, help="Path to output .pptx file")
    gen_parser.add_argument("--api-key", required=False, help="Gemini API Key")

    # Stage 1: Ingest
    ing_parser = subparsers.add_parser("ingest", help="Stage 1: Extract design tokens from template")
    ing_parser.add_argument("--template", required=True, help="Path to Slide Master .pptx template")
    ing_parser.add_argument("--output", default="output", help="Output directory")

    # Stage 2: Plan
    plan_parser = subparsers.add_parser("plan", help="Stage 2: Generate UI plan from markdown via Gemini")
    plan_parser.add_argument("--markdown", required=True, help="Path to input Markdown file")
    plan_parser.add_argument("--tokens", required=True, help="Path to design_tokens.json")
    plan_parser.add_argument("--output", default="output/ui_plan.json", help="Path to output ui_plan.json")
    plan_parser.add_argument("--api-key", required=False, help="Gemini API Key")

    # Stage 3: Compile
    comp_parser = subparsers.add_parser("compile", help="Stage 3: Compile tokens + plan into .pptx")
    comp_parser.add_argument("--tokens", required=True, help="Path to design_tokens.json")
    comp_parser.add_argument("--plan", required=True, help="Path to ui_plan.json")
    comp_parser.add_argument("--template", required=True, help="Path to the original .pptx template")
    comp_parser.add_argument("--output", required=True, help="Path to output .pptx file")
    comp_parser.add_argument("--no-validate", action="store_true", help="Skip post-generation validation")

    # Editor Mode
    edit_parser = subparsers.add_parser("edit", help="Surgical edits to an existing presentation")
    edit_parser.add_argument("--pptx", required=True, help="Path to the .pptx file")
    edit_parser.add_argument("--inventory", action="store_true", help="List all text shapes")
    edit_parser.add_argument("--replace", help="JSON string: {\"slide\": 3, \"old\": \"...\", \"new\": \"...\"}")
    edit_parser.add_argument("--reorder", help="Comma-separated 0-indexed slide order")
    edit_parser.add_argument("--output", help="Output path (default: overwrite input)")

    # MCP Server
    subparsers.add_parser("serve", help="Start FastMCP server")

    args = parser.parse_args()

    if args.command == "generate":
        result = generate_presentation(args.markdown, args.template, args.output, args.api_key)
        print("\n" + "=" * 60)
        print(result)

    elif args.command == "ingest":
        run_ingest(args.template, args.output)

    elif args.command == "plan":
        from orchestrator import plan_presentation as run_orchestrate
        run_orchestrate(args.markdown, args.tokens, args.output, args.api_key)

    elif args.command == "compile":
        run_compile(args.tokens, args.plan, args.template, args.output)
        
        # Auto-fix before validating
        run_fixes(args.output, None, args.output)
        
        if not args.no_validate:
            import json
            with open(args.tokens, 'r', encoding='utf-8') as f:
                tokens = json.load(f)
            report = validate_presentation(args.output, tokens)
            print_validation_report(report)
            
    elif args.command == "edit":
        if args.inventory:
            inv = get_inventory(args.pptx)
            print_inventory(inv)
        elif args.replace:
            import json
            repl = json.loads(args.replace)
            if isinstance(repl, dict): repl = [repl]
            res = replace_text(args.pptx, repl, args.output)
            print(f"Replaced {res['total_replacements']} occurrences. Saved to {res['output_path']}")
        elif args.reorder:
            order = [int(x.strip()) for x in args.reorder.split(',')]
            out = reorder_slides(args.pptx, order, args.output)
            print(f"Slides reordered. Saved to {out}")

    elif args.command == "serve":
        start_mcp_server()

    else:
        parser.print_help()


if __name__ == "__main__":
    cli_mode()
