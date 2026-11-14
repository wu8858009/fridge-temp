"""打包 Windows 執行檔：python build.py → dist\\冰箱溫度記錄表.exe"""
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
NAME = "冰箱溫度記錄表"
# 程式正在執行時 dist 裡的 exe 無法覆蓋，可用 python build.py <其他資料夾> 先打包到別處
DIST = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else ROOT / "dist"


def main() -> None:
    subprocess.check_call([sys.executable, str(ROOT / "make_icons.py")])
    shutil.rmtree(ROOT / "build" / NAME, ignore_errors=True)
    subprocess.check_call([
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onefile", "--windowed",
        "--name", NAME,
        "--icon", str(ROOT / "app.ico"),
        "--add-data", f"{ROOT / 'web'};web",
        "--distpath", str(DIST),
        "--workpath", str(ROOT / "build"),
        "--specpath", str(ROOT / "build"),
        str(ROOT / "desktop" / "main.py"),
    ], cwd=ROOT)
    exe = DIST / f"{NAME}.exe"
    print(f"\n完成：{exe}  ({exe.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
