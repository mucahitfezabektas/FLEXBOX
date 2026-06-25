import hashlib
import hmac
import json
import os
import secrets
from asyncio import sleep
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel

from backend.version import APP_VERSION


class LoginRequest(BaseModel):
    username: str
    password: str


DEFAULT_USERNAME = "admin"
DEFAULT_PASSWORD = "uniframe"
DEFAULT_ROLE = "Administrator"
AUTH_ITERATIONS = 200_000

app = FastAPI(title="UniFrame API", version=APP_VERSION)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:1420",
        "http://127.0.0.1:1420",
        "http://tauri.localhost",
        "https://tauri.localhost",
        "tauri://localhost",
        "null",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def config_dir() -> Path:
    appdata = os.getenv("APPDATA")
    base_dir = Path(appdata) if appdata else Path.home()
    target_dir = base_dir / "UniFrame"
    target_dir.mkdir(parents=True, exist_ok=True)
    return target_dir


def auth_config_path() -> Path:
    return config_dir() / "auth-config.json"


def hash_password(password: str, salt: str, iterations: int = AUTH_ITERATIONS) -> str:
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        bytes.fromhex(salt),
        iterations,
    )
    return digest.hex()


def build_default_auth_config() -> dict:
    salt = secrets.token_hex(16)
    return {
        "environment": os.getenv("UNIFRAME_ENV", "DEV"),
        "auth_mode": "local-config",
        "users": [
            {
                "username": DEFAULT_USERNAME,
                "role": DEFAULT_ROLE,
                "active": True,
                "iterations": AUTH_ITERATIONS,
                "salt": salt,
                "password_hash": hash_password(DEFAULT_PASSWORD, salt, AUTH_ITERATIONS),
            }
        ],
    }


def load_auth_config() -> dict:
    path = auth_config_path()

    if not path.exists():
        payload = build_default_auth_config()
        path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        return payload

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(payload.get("users"), list) or not payload["users"]:
            raise ValueError("invalid auth store")
        return payload
    except Exception:
        payload = build_default_auth_config()
        path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        return payload


def authenticate_user(username: str, password: str) -> dict | None:
    payload = load_auth_config()

    for user in payload["users"]:
        if not user.get("active", True):
            continue

        if user.get("username", "").lower() != username.lower():
            continue

        salt = user.get("salt", "")
        iterations = int(user.get("iterations", AUTH_ITERATIONS))
        expected_hash = user.get("password_hash", "")
        password_hash = hash_password(password, salt, iterations)

        if hmac.compare_digest(password_hash, expected_hash):
            return user

    return None


@app.get("/api/health")
async def healthcheck():
    return {"status": "ok", "service": "uniframe-api", "version": APP_VERSION}


@app.get("/api/system/context")
async def system_context():
    auth_payload = load_auth_config()
    return {
        "application": "UniFrame Unified Framework",
        "version": APP_VERSION,
        "environment": auth_payload.get("environment", "DEV"),
        "backend": "FastAPI Sidecar",
        "transport": "127.0.0.1:8000",
        "modules": ["auth", "customize", "empty"],
        "authMode": auth_payload.get("auth_mode", "local-config"),
    }


@app.post("/api/auth/login")
async def login(payload: LoginRequest):
    await sleep(0.25)

    user = authenticate_user(payload.username, payload.password)
    if user:
        return {
            "success": True,
            "message": f"{user['username']} kullanicisi dogrulandi.",
            "token": secrets.token_urlsafe(24),
            "role": user.get("role", DEFAULT_ROLE),
        }

    return JSONResponse(
        status_code=401,
        content={"success": False, "message": "Gecersiz bilgiler."},
    )
