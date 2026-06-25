from __future__ import annotations

import argparse
import json
import re
import shutil
import struct
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
BUILD_DIR = ROOT / "build"
SIDE_CAR_DIST_DIR = BUILD_DIR / "sidecar"
RELEASES_DIR = ROOT / "releases"
SRC_TAURI_DIR = ROOT / "src-tauri"
BINARIES_DIR = SRC_TAURI_DIR / "binaries"
ICONS_DIR = SRC_TAURI_DIR / "icons"
BACKEND_VENV_DIR = ROOT / "backend" / ".venv"
BACKEND_VENV_PYTHON = BACKEND_VENV_DIR / "Scripts" / "python.exe"
VERSION_SETTINGS_PATH = ROOT / "version-settings.json"
PACKAGE_JSON_PATH = ROOT / "package.json"
PACKAGE_LOCK_PATH = ROOT / "package-lock.json"
TAURI_CONFIG_PATH = SRC_TAURI_DIR / "tauri.conf.json"
CARGO_TOML_PATH = SRC_TAURI_DIR / "Cargo.toml"
BACKEND_VERSION_PATH = ROOT / "backend" / "version.py"
REMOTE_URL = "https://github.com/mucahitfezabektas/FLEXBOX.git"
SIDE_CAR_NAME = "uniframe-backend"
VERSION_PATTERN = re.compile(r"^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build and publish the UniFrame desktop release.")
    parser.add_argument(
        "--version",
        help="Override version-settings.json and persist the supplied semantic version before the build starts.",
    )
    return parser.parse_args()


def run(command: list[str], *, env: dict[str, str] | None = None, check: bool = True) -> subprocess.CompletedProcess[str]:
    print("$", " ".join(command))
    return subprocess.run(
        command,
        cwd=ROOT,
        env=env,
        text=True,
        check=check,
    )


def resolve_npm() -> str:
    return shutil.which("npm.cmd") or shutil.which("npm") or "npm.cmd"


def validate_version(version: str) -> str:
    if not VERSION_PATTERN.match(version):
        raise ValueError(f"Invalid version value: {version}")
    return version


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, payload: dict) -> None:
    path.write_text(f"{json.dumps(payload, indent=2)}\n", encoding="utf-8")


def load_version_settings() -> dict:
    if not VERSION_SETTINGS_PATH.exists():
        package_payload = read_json(PACKAGE_JSON_PATH)
        payload = {"version": package_payload["version"]}
        write_json(VERSION_SETTINGS_PATH, payload)
        return payload

    payload = read_json(VERSION_SETTINGS_PATH)
    if not isinstance(payload.get("version"), str):
        raise RuntimeError("version-settings.json must contain a string 'version' field.")
    return payload


def update_cargo_version(version: str) -> None:
    cargo_manifest = CARGO_TOML_PATH.read_text(encoding="utf-8")
    updated_manifest, replacements = re.subn(
        r'(?m)^version = "[^"]+"$',
        f'version = "{version}"',
        cargo_manifest,
        count=1,
    )
    if replacements != 1:
        raise RuntimeError("Cargo.toml version field could not be updated.")
    CARGO_TOML_PATH.write_text(updated_manifest, encoding="utf-8")


def sync_version_metadata(version: str) -> None:
    package_payload = read_json(PACKAGE_JSON_PATH)
    package_payload["version"] = version
    write_json(PACKAGE_JSON_PATH, package_payload)

    package_lock_payload = read_json(PACKAGE_LOCK_PATH)
    package_lock_payload["version"] = version
    root_package = package_lock_payload.get("packages", {}).get("")
    if isinstance(root_package, dict):
        root_package["version"] = version
    write_json(PACKAGE_LOCK_PATH, package_lock_payload)

    tauri_payload = read_json(TAURI_CONFIG_PATH)
    tauri_payload["version"] = version
    write_json(TAURI_CONFIG_PATH, tauri_payload)

    update_cargo_version(version)
    BACKEND_VERSION_PATH.write_text(f'APP_VERSION = "{version}"\n', encoding="utf-8")


def resolve_build_version(cli_version: str | None) -> str:
    payload = load_version_settings()
    version = validate_version(cli_version or payload["version"])

    if payload["version"] != version:
        write_json(VERSION_SETTINGS_PATH, {"version": version})

    return version


def require_command(command_name: str, resolved_path: str | None = None) -> str:
    command = resolved_path or shutil.which(command_name)
    if not command:
        raise RuntimeError(f"Required command is missing: {command_name}")
    return command


def check_prerequisites() -> tuple[str, str]:
    npm_command = require_command("npm.cmd", resolve_npm())
    require_command("git")
    require_command("git-lfs")
    require_command("rustc")
    return npm_command, sys.executable


def ensure_app_icon() -> None:
    ICONS_DIR.mkdir(parents=True, exist_ok=True)
    icon_path = ICONS_DIR / "app.ico"
    size = 16
    pixel = bytes((255, 200, 52, 14))  # BGRA
    xor_bitmap = pixel * (size * size)
    and_mask = b"\x00\x00\x00\x00" * size
    dib_header = struct.pack(
        "<IIIHHIIIIII",
        40,
        size,
        size * 2,
        1,
        32,
        0,
        len(xor_bitmap),
        0,
        0,
        0,
        0,
    )
    image_data = dib_header + xor_bitmap + and_mask
    icon_dir = struct.pack("<HHH", 0, 1, 1)
    icon_entry = struct.pack(
        "<BBBBHHII",
        size,
        size,
        0,
        0,
        1,
        32,
        len(image_data),
        22,
    )
    icon_path.write_bytes(icon_dir + icon_entry + image_data)


def ensure_backend_venv() -> str:
    if not BACKEND_VENV_PYTHON.exists():
        run([sys.executable, "-m", "venv", str(BACKEND_VENV_DIR)])
    return str(BACKEND_VENV_PYTHON)


def get_target_triple() -> str:
    commands = [
        ["rustc", "--print", "host-tuple"],
        ["rustc", "-Vv"],
    ]

    for command in commands:
        result = subprocess.run(
            command,
            cwd=ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            continue

        output = result.stdout.strip()
        if command[-1] == "host-tuple" and output:
            return output

        for line in output.splitlines():
            if line.startswith("host:"):
                return line.split(":", 1)[1].strip()

    raise RuntimeError("Rust target triple could not be detected.")


def build_backend_sidecar(backend_python: str, target_triple: str) -> Path:
    BINARIES_DIR.mkdir(parents=True, exist_ok=True)
    SIDE_CAR_DIST_DIR.mkdir(parents=True, exist_ok=True)

    run(
        [
            backend_python,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--onefile",
            "--name",
            SIDE_CAR_NAME,
            "--distpath",
            str(SIDE_CAR_DIST_DIR),
            "backend/sidecar_main.py",
        ]
    )

    source_executable = SIDE_CAR_DIST_DIR / f"{SIDE_CAR_NAME}.exe"
    if not source_executable.exists():
        raise FileNotFoundError(f"PyInstaller output not found: {source_executable}")

    target_executable = BINARIES_DIR / f"{SIDE_CAR_NAME}-{target_triple}.exe"
    shutil.copy2(source_executable, target_executable)
    return target_executable


def ensure_dependencies() -> str:
    npm_command, _ = check_prerequisites()
    backend_python = ensure_backend_venv()
    run([npm_command, "install"])
    run([backend_python, "-m", "pip", "install", "-r", "backend/requirements.txt"])
    return backend_python


def build_tauri() -> Path:
    npm_command = resolve_npm()
    run([npm_command, "run", "tauri:build"])

    nsis_dir = SRC_TAURI_DIR / "target" / "release" / "bundle" / "nsis"
    installers = sorted(nsis_dir.glob("*-setup.exe"))
    if not installers:
        raise FileNotFoundError(f"NSIS installer not found in {nsis_dir}")

    return installers[-1]


def prepare_release_artifact(installer: Path) -> Path:
    RELEASES_DIR.mkdir(parents=True, exist_ok=True)
    destination = RELEASES_DIR / installer.name
    shutil.copy2(installer, destination)
    return destination


def ensure_lfs() -> None:
    run(["git", "lfs", "install"])
    run(["git", "lfs", "track", "releases/*.exe"])


def ensure_origin_remote() -> None:
    result = subprocess.run(
        ["git", "remote", "get-url", "origin"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )

    if result.returncode != 0:
        run(["git", "remote", "add", "origin", REMOTE_URL])
        return

    current_url = result.stdout.strip()
    if current_url != REMOTE_URL:
        run(["git", "remote", "set-url", "origin", REMOTE_URL])


def check_remote_access() -> str | None:
    result = subprocess.run(
        ["git", "ls-remote", REMOTE_URL, "refs/heads/main"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError("GitHub remote could not be reached.")

    output = result.stdout.strip()
    if not output:
        return None

    return output.split()[0]


def ensure_main_branch() -> None:
    current_branch = subprocess.run(
        ["git", "branch", "--show-current"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=True,
    ).stdout.strip()

    if current_branch != "main":
        run(["git", "branch", "-M", "main"])


def commit_and_push(installer_path: Path) -> None:
    status = subprocess.run(
        ["git", "status", "--short"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=True,
    ).stdout.strip()

    if status:
        run(["git", "add", "."])
        run(["git", "commit", "-m", f"Build UniFrame release with {installer_path.name}"], check=False)

    run(["git", "push", "-u", "origin", "main"])


def get_local_head_sha() -> str:
    return subprocess.run(
        ["git", "rev-parse", "HEAD"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=True,
    ).stdout.strip()


def ensure_clean_worktree() -> None:
    status = subprocess.run(
        ["git", "status", "--short"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=True,
    ).stdout.strip()

    if status:
        raise RuntimeError("Working tree is not clean after build and push.")


def main() -> int:
    args = parse_args()
    build_version = resolve_build_version(args.version)
    sync_version_metadata(build_version)
    remote_head_before = check_remote_access()
    ensure_app_icon()
    backend_python = ensure_dependencies()
    target_triple = get_target_triple()
    build_backend_sidecar(backend_python, target_triple)
    installer_path = build_tauri()
    copied_installer = prepare_release_artifact(installer_path)
    ensure_lfs()
    ensure_origin_remote()
    ensure_main_branch()
    commit_and_push(copied_installer)
    remote_head_after = check_remote_access()
    local_head = get_local_head_sha()
    ensure_clean_worktree()
    print(f"Build version: {build_version}")
    print(f"Previous remote HEAD: {remote_head_before or 'none'}")
    print(f"Current remote HEAD: {remote_head_after or 'none'}")
    print(f"Local HEAD: {local_head}")
    print(f"Release ready: {copied_installer} ({copied_installer.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
