import subprocess
import sys
from pathlib import Path
import xlwings as xw


def generate_report():
    project_root = Path(__file__).resolve().parent.parent.parent

    log_file = project_root / "excel_debug.log"

    with open(log_file, "a") as f:
        f.write("\n=== START ===\n")
        f.write(f"Python: {sys.executable}\n")
        f.write(f"Project Root: {project_root}\n")

    result = subprocess.run(
        [sys.executable, "app.py"],
        cwd=project_root,
        capture_output=True,
        text=True,
    )

    with open(log_file, "a") as f:
        f.write(f"Return Code: {result.returncode}\n")
        f.write(f"STDOUT:\n{result.stdout}\n")
        f.write(f"STDERR:\n{result.stderr}\n")
        f.write("=== END ===\n")

    return {
        "success": result.returncode == 0,
        "stdout": result.stdout,
        "stderr": result.stderr,
    }


if __name__ == "__main__":
    result = generate_report()

    if result["success"]:
        print("Report generated successfully.")
        print(result["stdout"])
    else:
        print("Report generation failed.")
        print(result["stderr"])


def excel_generate_report():
    try:
        result = generate_report()
        print(result)
        return result
    except Exception as e:
        import traceback
        print("ERROR:")
        print(traceback.format_exc())
        raise