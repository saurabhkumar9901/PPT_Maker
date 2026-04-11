import win32com.client
import os

ppt_path = os.path.abspath(r'output\Accenture_v4g_Aligned.pptx')
out_dir = os.path.abspath(r'output\v4g_images')
os.makedirs(out_dir, exist_ok=True)

powerpoint = win32com.client.Dispatch("Powerpoint.Application")
try:
    ppt = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
except:
    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    ppt = powerpoint.Presentations.Open(ppt_path, WithWindow=False)

for i, slide in enumerate(ppt.Slides):
    slide.Export(os.path.join(out_dir, f"slide_{i+1}.png"), "PNG")
    if i >= 4:
        break

ppt.Close()
powerpoint.Quit()
print("Exported 5 slides")
