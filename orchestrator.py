"""
Stage 2 -- Orchestration: Markdown -> UI JSON via Gemini Structured Outputs.

Reads a markdown file and design tokens, sends them to Gemini Flash,
and forces the LLM to return a UI JSON schema describing each slide's
layout, content, and visual component types.

Outputs: ui_plan.json
"""

import os
import json
import argparse
from pydantic import BaseModel, Field
from dotenv import load_dotenv
import google.genai as genai
from google.genai import types

load_dotenv()


# ---------------------------------------------------------------------------
#  Pydantic Schemas  --  Every element type the LLM can produce
# ---------------------------------------------------------------------------

class GridItem(BaseModel):
    heading: str = Field(description="Short card heading, max 5 words")
    body: str = Field(description="Card body text, max 15 words")

class GridElement(BaseModel):
    type: str = Field(default="grid", description="Must be 'grid'")
    columns: int = Field(description="Number of columns (2-4)")
    items: list[GridItem]

class TimelineStep(BaseModel):
    label: str = Field(description="Step label e.g. '01', 'Phase 1'")
    title: str = Field(description="Step title")
    description: str = Field(description="Detailed step description paragraph, min 40 words")

class TimelineElement(BaseModel):
    type: str = Field(default="timeline", description="Must be 'timeline'")
    steps: list[TimelineStep]

class HeroElement(BaseModel):
    type: str = Field(default="hero", description="Must be 'hero'")
    heading: str = Field(description="Large hero heading")
    body: str = Field(description="Detailed supporting body text paragraph, min 80 words")
    image_query: str | None = Field(default=None, description="Optional search term for hero background image")
    image_url: str | None = Field(default=None, description="Extract any actual image file URL/path from the markdown if present")

class BulletItem(BaseModel):
    text: str = Field(description="Detailed bullet point text paragraph, min 40 words")
    bold_prefix: str | None = Field(default=None, description="Optional bold prefix before the bullet text")

class BulletsElement(BaseModel):
    type: str = Field(default="bullets", description="Must be 'bullets'")
    items: list[BulletItem]

class ChartSeries(BaseModel):
    name: str
    values: list[float]

class ChartElement(BaseModel):
    type: str = Field(default="chart", description="Must be 'chart'")
    chart_type: str = Field(description="One of: bar, line, pie, column, doughnut, area")
    categories: list[str]
    series: list[ChartSeries]

class TableRow(BaseModel):
    cells: list[str]

class TableElement(BaseModel):
    type: str = Field(default="table", description="Must be 'table'")
    headers: list[str]
    rows: list[TableRow]

class TwoColumnContent(BaseModel):
    heading: str
    body: str

class TwoColumnElement(BaseModel):
    type: str = Field(default="two_column", description="Must be 'two_column'")
    left: TwoColumnContent
    right: TwoColumnContent

class StatItem(BaseModel):
    value: str = Field(description="The statistic value e.g. '42%', '$1.2B'")
    label: str = Field(description="Label for the stat, max 4 words")

class StatsRowElement(BaseModel):
    type: str = Field(default="stats_row", description="Must be 'stats_row'")
    items: list[StatItem]

class QuoteElement(BaseModel):
    type: str = Field(default="quote", description="Must be 'quote'")
    quote: str = Field(description="The quote text")
    attribution: str = Field(description="Who said it")

class ImageTextContent(BaseModel):
    heading: str
    body: str
    image_side: str = Field(default="right", description="'left' or 'right'")
    image_query: str | None = Field(default=None, description="Short 2-3 word search term for an image, e.g. 'corporate office', 'data server'")
    image_url: str | None = Field(default=None, description="Extract any actual image file URL/path from the markdown if present")

class ImageTextElement(BaseModel):
    type: str = Field(default="image_text", description="Must be 'image_text'")
    content: ImageTextContent

class ComparisonColumn(BaseModel):
    title: str
    points: list[str]

class ComparisonElement(BaseModel):
    type: str = Field(default="comparison", description="Must be 'comparison'")
    left: ComparisonColumn
    right: ComparisonColumn

class IconGridItem(BaseModel):
    icon: str = Field(description="Emoji or text icon e.g. a single emoji character")
    title: str = Field(description="Short title")
    description: str = Field(description="Detailed explanation paragraph, min 50 words")

class IconGridElement(BaseModel):
    type: str = Field(default="icon_grid", description="Must be 'icon_grid'")
    columns: int = Field(default=3, description="Number of columns (2-4)")
    items: list[IconGridItem]

class WaterfallStep(BaseModel):
    label: str
    value: float
    is_total: bool = Field(default=False)

class WaterfallElement(BaseModel):
    type: str = Field(default="waterfall", description="Must be 'waterfall'")
    steps: list[WaterfallStep]

class FunnelStep(BaseModel):
    label: str
    value: str = Field(description="Display value e.g. '1000', '85%'")
    description: str | None = Field(default=None)

class FunnelElement(BaseModel):
    type: str = Field(default="funnel", description="Must be 'funnel'")
    steps: list[FunnelStep]

class PyramidLevel(BaseModel):
    label: str
    description: str | None = Field(default=None)

class PyramidElement(BaseModel):
    type: str = Field(default="pyramid", description="Must be 'pyramid'")
    levels: list[PyramidLevel] = Field(description="From top (smallest) to bottom (largest)")

class MatrixQuadrant(BaseModel):
    label: str
    items: list[str]

class MatrixElement(BaseModel):
    type: str = Field(default="matrix", description="Must be 'matrix'")
    title: str | None = Field(default=None)
    x_axis: str = Field(description="Label for horizontal axis")
    y_axis: str = Field(description="Label for vertical axis")
    quadrants: list[MatrixQuadrant] = Field(description="Exactly 4 quadrants: TL, TR, BL, BR")

class SWOTElement(BaseModel):
    type: str = Field(default="swot", description="Must be 'swot'")
    strengths: list[str]
    weaknesses: list[str]
    opportunities: list[str]
    threats: list[str]

class CycleStep(BaseModel):
    title: str
    description: str | None = Field(default=None)

class CycleElement(BaseModel):
    type: str = Field(default="cycle", description="Must be 'cycle'")
    steps: list[CycleStep]

class GaugeElement(BaseModel):
    type: str = Field(default="gauge", description="Must be 'gauge'")
    label: str
    value: float = Field(description="Value 0-100")
    unit: str = Field(default="%")

class KPIItem(BaseModel):
    label: str
    value: str
    trend: str | None = Field(default=None, description="'up', 'down', or 'flat'")
    change: str | None = Field(default=None, description="e.g. '+12%', '-5%'")

class KPIElement(BaseModel):
    type: str = Field(default="kpi_cards", description="Must be 'kpi_cards'")
    items: list[KPIItem]

# Union of all element types -- Gemini will output one of these per slide element
SlideElement = (
    GridElement | TimelineElement | HeroElement | BulletsElement |
    ChartElement | TableElement | TwoColumnElement | StatsRowElement |
    QuoteElement | ImageTextElement | ComparisonElement | IconGridElement |
    WaterfallElement | FunnelElement | PyramidElement | MatrixElement |
    SWOTElement | CycleElement | GaugeElement | KPIElement
)


class SlideSchema(BaseModel):
    layout: str = Field(description="One of: cover, divider, content, chart, thank_you")
    title: str = Field(description="Slide title")
    subtitle: str | None = Field(default=None, description="Optional subtitle (used for cover/divider)")
    elements: list[SlideElement] = Field(default_factory=list, description="List of visual elements on this slide")

class PresentationPlan(BaseModel):
    slides: list[SlideSchema]


# ---------------------------------------------------------------------------
#  Prompt Engineering
# ---------------------------------------------------------------------------

ELEMENT_CATALOG = """
AVAILABLE ELEMENT TYPES (choose the best fit for each slide):

| Type          | Description                                      | Best For                                  |
|---------------|--------------------------------------------------|-------------------------------------------|
| grid          | N-column card grid (2-4 cols)                    | Multi-point content, features, comparisons|
| timeline      | Horizontal process/step flow with connectors     | Processes, roadmaps, sequences            |
| hero          | Single large centered text block                 | Conclusions, key takeaways, openers       |
| bullets       | Styled bullet list with optional bold prefix     | Simple text content, lists                |
| chart         | Embedded chart (bar/line/pie/column/doughnut/area)| Data visualization                       |
| table         | Data table with headers                          | Structured tabular data                   |
| two_column    | 50/50 split with distinct left/right content     | Comparisons, before/after, pros/cons      |
| stats_row     | Row of large statistics with labels              | KPIs, key numbers, metrics                |
| quote         | Blockquote with attribution                      | Expert opinions, testimonials             |
| image_text    | Text on one side, placeholder on other           | Product features, case studies            |
| comparison    | Side-by-side columns with bullet points          | Detailed comparisons, vs. analysis        |
| icon_grid     | Grid of items with emoji icons                   | Features, capabilities, benefits          |
| waterfall     | Waterfall chart data                             | Financial breakdowns, cumulative changes  |
| funnel        | Funnel visualization                             | Sales funnels, conversion pipelines       |
| pyramid       | Pyramid/hierarchy visualization                  | Hierarchies, priority levels              |
| matrix        | 2x2 quadrant matrix                              | Strategic positioning, risk matrices      |
| swot          | SWOT analysis (4-quadrant)                       | Strategic analysis                        |
| cycle         | Circular process flow                            | Recurring processes, feedback loops       |
| gauge         | Single gauge/meter visualization                 | Progress indicators, scores               |
| kpi_cards     | Cards showing KPIs with trends                   | Dashboard-style metric displays           |
"""

def build_prompt(markdown_content: str, tokens: dict) -> str:
    colors_summary = ", ".join(f"{k}={v}" for k, v in tokens["colors"].items())
    fonts_summary = f"Heading: {tokens['fonts']['heading']}, Body: {tokens['fonts']['body']}"
    layout_names = [l["name"] for l in tokens["layouts"]]

    return f"""You are an expert Presentation Designer and Information Architect.
You design slides that look like they came from McKinsey or Accenture — premium, data-rich, infographic-style.

Read the following markdown content and convert it into a STRICT 10-15 slide presentation plan.
Each slide must use exactly one element type from the catalog below.
You must maximize VISUAL VARIETY -- do NOT use the same element type on consecutive slides.

TEMPLATE INFO:
- Slide size: {tokens['dimensions']['width']}" x {tokens['dimensions']['height']}"
- Available layouts: {layout_names}
- Color palette: {colors_summary}
- Fonts: {fonts_summary}

{ELEMENT_CATALOG}

CRITICAL RULES:
1. First slide MUST be layout="cover" with a title and subtitle. No elements needed.
2. Last slide MUST be layout="thank_you". No elements needed.
3. Use 1-2 "divider" layout slides to break the presentation into logical sections.
4. If content has steps/processes, use "timeline" or "cycle".
5. If content has numerical data, use "chart", "stats_row", "kpi_cards", "waterfall", or "gauge".
6. If content has comparisons, use "comparison", "two_column", "matrix", or "swot".
7. If content has hierarchies, use "pyramid" or "funnel".
8. All other content slides use layout="content".
9. Chart slides use layout="chart".

CONTENT DENSITY (CRITICAL — slides must NOT feel empty):
- Each slide should be packed with high-value information. GATHER COMPREHENSIVE DETAILS from the source text. Do not just summarize; extract the detailed explanations, methodologies, and specific examples.
- Grid/icon_grid: USE 4-6 ITEMS. Fill the slide. Give them rich descriptions.
- Bullets: USE 5-8 items per slide. Include sub-points if it adds value.
- Stats_row: USE 3-4 stats. Extract real numbers and detailed context from the source.
- Table: USE 4-8 data rows. Fill with actual detailed data.
- Two_column/comparison: Both sides must be densely populated with equal, substantial content.

TEXT ABUNDANCE & AUTO-FIT (CRITICAL):
- Our compiler uses auto-fit font scaling and dynamic stretching to fill space. YOU MUST WRITE EXTREMELY LONG TEXT to eliminate empty slide space.
- Grid/Icon items: heading max 8 words, body MUST BE 80-120 words. YOU MUST GENERATE EXACTLY 4 OR 6 ITEMS for a grid to ensure symmetrical layouts (no empty 5th slot!).
- Timeline/cycle steps: description MUST BE 40-60 words. Give detailed chronologies. YOU MUST GENERATE EXACTLY 3 OR 4 STEPS for a timeline to ensure column widths remain readable.
- Bullets: each item MUST BE 60-100 words. Use heavy, paragraph-style bullets with extensive context.
- Two-column/Hero: Body MUST BE 150-250 words. Provide comprehensive, unbroken analysis blocks.
- All slide titles: 8-15 words. Keep them descriptive like a headline.
- DO NOT CONDENSE. Extraction should pull massive blocks of context from your source material. Fill all empty negative space!

VARIETY RULE:
- You MUST use at least 6 DIFFERENT element types across the presentation.
- Never use the same element type more than 2 times total.
- Prefer infographic-style elements (grid, icon_grid, stats_row, kpi_cards, timeline) over text-heavy ones (bullets).

SLIDE TITLES:
- Be specific and descriptive: "Acquisition Volume by Fiscal Year" NOT "Overview"
- Include numbers where relevant: "$6.6B Investment in FY24" NOT "Financial Summary"

MARKDOWN CONTENT:
{markdown_content}
"""


# ---------------------------------------------------------------------------
#  Main Pipeline
# ---------------------------------------------------------------------------

def plan_presentation(markdown_path: str, tokens_path: str, output_path: str = "output/ui_plan.json", api_key: str = None) -> str:
    """
    Stage 2: Send markdown + tokens to Gemini -> UI plan JSON.
    Returns the path to the saved ui_plan.json.
    """
    api_key = api_key or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY not provided. Set it in .env or pass --api-key.")

    # Load inputs
    with open(markdown_path, "r", encoding="utf-8") as f:
        markdown_content = f.read()

    with open(tokens_path, "r", encoding="utf-8") as f:
        tokens = json.load(f)

    # Truncate very large markdown to avoid token limits
    MAX_CHARS = 100_000
    if len(markdown_content) > MAX_CHARS:
        print(f"WARNING: Markdown is {len(markdown_content)} chars. Truncating to {MAX_CHARS}.")
        markdown_content = markdown_content[:MAX_CHARS]

    prompt = build_prompt(markdown_content, tokens)

    print("Sending content to Gemini for structured slide planning...")
    print(f"  Markdown: {len(markdown_content):,} chars")
    print(f"  Template: {tokens.get('template_name', 'unknown')}")

    client = genai.Client(api_key=api_key)

    response = client.models.generate_content(
        model="gemini-3-flash-preview",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=PresentationPlan,
            temperature=0.2,
        ),
    )

    plan = response.parsed
    print(f"  Gemini designed {len(plan.slides)} slides.")

    # Log slide types for visibility
    for i, slide in enumerate(plan.slides):
        element_types = [e.type for e in slide.elements] if slide.elements else ["(none)"]
        print(f"    Slide {i+1}: [{slide.layout}] {slide.title} -> {', '.join(element_types)}")

    # Save
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(plan.model_dump(), f, indent=2, ensure_ascii=False)

    print(f"\nUI plan saved to: {output_path}")
    return output_path


# ---------------------------------------------------------------------------
#  CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Stage 2 -- Orchestrate: Markdown -> UI JSON via Gemini."
    )
    parser.add_argument("--markdown", required=True, help="Path to input Markdown file")
    parser.add_argument("--tokens", required=True, help="Path to design_tokens.json from Stage 1")
    parser.add_argument("--output", default="output/ui_plan.json", help="Path to output ui_plan.json")
    parser.add_argument("--api-key", required=False, help="Gemini API Key (defaults to GEMINI_API_KEY env var)")
    args = parser.parse_args()

    plan_presentation(args.markdown, args.tokens, args.output, args.api_key)
