import base64
import json
import os
import shutil
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

PROJECT_ROOT = Path(__file__).resolve().parents[2]
COLAB_SCRIPT = PROJECT_ROOT / "app" / "pipeline" / "originals" / "colab101.py"
OUTPUT_JSON = "output.json"
EXCEL_FILE = "financial_report.xlsx"


def _normalize_payload(payload: Any) -> Any:
    if payload is None:
        raise ValueError("payload is required")
    return payload


def _run(cmd, cwd: Path, env: dict):
    p = subprocess.run(cmd, cwd=str(cwd), env=env, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError(
            "Command failed:\n"
            f"cmd={cmd}\n"
            f"returncode={p.returncode}\n"
            f"stdout:\n{p.stdout}\n"
            f"stderr:\n{p.stderr}\n"
        )
    return p


def _default_filename(payload: Any) -> str:
    date_str = ""
    if isinstance(payload, dict):
        cd = payload.get("決算期年月日") or {}
        if isinstance(cd, dict):
            date_str = str(cd.get("今期") or "").replace("/", "").replace("-", "")
        if not date_str:
            date_str = str(payload.get("postingPeriod") or "").replace("/", "").replace("-", "")
    return f"財務分析表{'_' + date_str if date_str else ''}.xlsx"


def run_colab101(payload: Any) -> Dict[str, Any]:
    payload = _normalize_payload(payload)
    run_dir = Path(tempfile.mkdtemp(prefix="cashai02_excel_", dir="/tmp"))
    try:
        (run_dir / OUTPUT_JSON).write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        env = dict(os.environ)
        env["NO_HTML"] = "1"
        env["DISABLE_EXCEL"] = "0"
        env["EXCEL_OUTPUT_PATH"] = EXCEL_FILE
        env["PYTHONIOENCODING"] = "utf-8"

        _run(["python3", str(COLAB_SCRIPT)], cwd=run_dir, env=env)

        excel_path = run_dir / EXCEL_FILE
        if not excel_path.exists():
            raise RuntimeError("Excel file was not generated.")

        b64 = base64.b64encode(excel_path.read_bytes()).decode("utf-8")
        return {
            "ok": True,
            "b64": b64,
            "filename": _default_filename(payload),
            "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "generated_at": datetime.utcnow().isoformat() + "Z",
        }
    finally:
        if os.getenv("DEBUG_KEEP_TMP", "0") != "1":
            shutil.rmtree(run_dir, ignore_errors=True)
