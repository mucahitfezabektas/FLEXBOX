from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parent
BACKEND_PORT = 8000
BACKEND_VENV_PYTHON = ROOT / "backend" / ".venv" / "Scripts" / "python.exe"


def resolve_npm() -> str:
    return shutil.which("npm.cmd") or shutil.which("npm") or "npm.cmd"


def resolve_backend_python() -> str:
    if BACKEND_VENV_PYTHON.exists():
        return str(BACKEND_VENV_PYTHON)
    return sys.executable


def wait_for_backend(process: subprocess.Popen[bytes], timeout_seconds: float = 15.0) -> None:
    deadline = time.time() + timeout_seconds
    healthcheck_command = [
        sys.executable,
        "-c",
        (
            "import sys, urllib.request; "
            "urllib.request.urlopen('http://127.0.0.1:8000/api/health', timeout=1); "
            "sys.exit(0)"
        ),
    ]

    while time.time() < deadline:
        if process.poll() is not None:
            raise RuntimeError("Backend process exited before becoming healthy.")

        result = subprocess.run(
            healthcheck_command,
            cwd=ROOT,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        if result.returncode == 0:
            return
        time.sleep(0.35)

    raise TimeoutError("Backend healthcheck timed out.")


def terminate_process(process: subprocess.Popen[bytes] | None) -> None:
    if process is None or process.poll() is not None:
        return

    process.terminate()
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=5)


def main() -> int:
    env = os.environ.copy()
    env.setdefault("PYTHONUNBUFFERED", "1")
    npm_command = resolve_npm()
    backend_python = resolve_backend_python()

    backend_process: subprocess.Popen[bytes] | None = None

    try:
        backend_process = subprocess.Popen(
            [
                sys.executable,
                "-m",
                "uvicorn",
                "backend.main:app",
                "--host",
                "127.0.0.1",
                "--port",
                str(BACKEND_PORT),
            ] if backend_python == sys.executable else [
                backend_python,
                "-m",
                "uvicorn",
                "backend.main:app",
                "--host",
                "127.0.0.1",
                "--port",
                str(BACKEND_PORT),
            ],
            cwd=ROOT,
            env=env,
        )
        wait_for_backend(backend_process)

        tauri_process = subprocess.run(
            [npm_command, "run", "tauri:dev"],
            cwd=ROOT,
            env=env,
            check=False,
        )
        return tauri_process.returncode
    except (RuntimeError, TimeoutError) as error:
        print(f"run.py error: {error}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        return 130
    finally:
        terminate_process(backend_process)


if __name__ == "__main__":
    raise SystemExit(main())
