"""Capture the real Tk interface using the screenshot skill helper."""
import sys
import time
import subprocess
import shutil
import ctypes
from pathlib import Path

ctypes.windll.shcore.SetProcessDpiAwareness(2)
BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE))
from prototype_tool.gui import PrototypeGUI, TkinterDnD, tk

dest = BASE / 'deliverables/20260916_resume_portfolio/screenshots'
dest.mkdir(parents=True, exist_ok=True)
root = TkinterDnD.Tk() if TkinterDnD else tk.Tk()
app = PrototypeGUI(root)
root.geometry(f'1320x{min(1100, root.winfo_screenheight() - 100)}+20+20')
app.template_path_var.set(str(BASE / '上海-架构师城市沙龙-PPT模板.pptx'))
app.quick_source_path_var.set(str(BASE / 'EverMemOS_Introduction.pptx'))
app.source_files = [str(BASE / 'EverMemOS_Introduction.pptx')]
app._refresh_template_summary()
app._refresh_sources()
app._refresh_output_summary()
app._refresh_quick_selection_summaries()
root.attributes('-topmost', True)
helper = Path.home() / '.codex/skills/screenshot/scripts/take_screenshot.ps1'
for index, name in enumerate(['01_一键替换.png', '02_批量替换.png', '03_模板管理.png']):
    app.notebook.select(index)
    root.update()
    time.sleep(0.8)
    root.update()
    region = f'{root.winfo_rootx()},{root.winfo_rooty()},{root.winfo_width()},{root.winfo_height()}'
    result = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(helper), '-Mode', 'temp', '-Region', region], capture_output=True, text=True, check=True)
    captured = Path(result.stdout.strip().splitlines()[-1])
    shutil.copy2(captured, dest / name)
    print(dest / name)
root.destroy()
