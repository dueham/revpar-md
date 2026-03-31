"""
auth_utils.py  —  Shared auth helpers for RevPar MD (AWS Version)
==================================================================
Handles:
  - Password hashing / verification
  - Short-lived session token generation + validation
  - users.yaml read / write (via S3)

All file I/O goes through S3 using helpers in hotel_config.py.

S3 layout:
  revpar-md-data/auth/users.yaml
  revpar-md-data/auth/session_tokens.json
  revpar-md-data/auth/cal_session_tokens.json

Token flow:
  1. Launcher validates credentials → writes a token to S3
  2. Launcher redirects browser to revpar_app on port 8502 with ?token=<uuid>
  3. revpar_app reads the token, validates it (60-second window), then
     stores the user's allowed hotels in st.session_state and deletes the token.
"""

import uuid
import time
import hashlib

from hotel_config import (
    read_json_from_s3,
    write_json_to_s3,
    read_yaml_from_s3,
    write_yaml_to_s3,
)

# ── S3 Key paths for auth files ───────────────────────────────────────────────
USERS_KEY          = "auth/users.yaml"
SESSION_TOKENS_KEY = "auth/session_tokens.json"
CAL_TOKENS_KEY     = "auth/cal_session_tokens.json"

# Token TTLs
TOKEN_TTL     = 60    # seconds — short-lived login redirect token
CAL_TOKEN_TTL = 1800  # seconds — 30-minute calendar session token

# All hotel IDs known to the system (must match hotel_config.py exactly)
ALL_HOTELS = [
    "hampton_lino_lakes",
    "hampton_superior",
    "hilton_checkers_la",
    "hampton_cherry_creek",
    "holiday_inn_express_superior",
    "hotel_indigo_rochester",
]


# ── Password helpers ──────────────────────────────────────────────────────────

def _hash_password(plain: str) -> str:
    """SHA-256 hash with fixed salt prefix."""
    salted = f"revparmd::{plain}"
    return hashlib.sha256(salted.encode()).hexdigest()


def verify_password(plain: str, stored_hash: str) -> bool:
    return _hash_password(plain) == stored_hash


def make_password_hash(plain: str) -> str:
    return _hash_password(plain)


# ── users.yaml helpers ────────────────────────────────────────────────────────

def load_users() -> dict:
    data = read_yaml_from_s3(USERS_KEY)
    if not data:
        return {}
    return data.get("users", {})


def save_users(users: dict) -> None:
    write_yaml_to_s3(USERS_KEY, {"users": users})


def authenticate(username: str, password: str):
    """
    Returns a user dict on success, or None on failure.
    User dict keys: username, display_name, role, hotels
    """
    users    = load_users()
    username = username.strip().lower()
    if username not in users:
        return None
    user   = users[username]
    stored = user.get("password_hash", "")
    if not stored or not verify_password(password, stored):
        return None
    role   = user.get("role", "read_only")
    hotels = ALL_HOTELS if role == "admin" else user.get("hotels", [])
    return {
        "username":     username,
        "display_name": user.get("display_name", username),
        "role":         role,
        "hotels":       hotels,
    }


# ── Session token helpers ─────────────────────────────────────────────────────

def _load_tokens() -> dict:
    data = read_json_from_s3(SESSION_TOKENS_KEY)
    return data if isinstance(data, dict) else {}


def _save_tokens(tokens: dict) -> None:
    write_json_to_s3(SESSION_TOKENS_KEY, tokens)


def _purge_expired(tokens: dict) -> dict:
    now = time.time()
    return {k: v for k, v in tokens.items()
            if now - v["created_at"] < TOKEN_TTL}


def create_session_token(user: dict) -> str:
    """Write a new short-lived token to S3 and return its UUID string."""
    token_id = str(uuid.uuid4())
    tokens   = _load_tokens()
    tokens   = _purge_expired(tokens)
    tokens[token_id] = {
        "username":     user["username"],
        "display_name": user["display_name"],
        "role":         user["role"],
        "hotels":       user["hotels"],
        "created_at":   time.time(),
    }
    _save_tokens(tokens)
    return token_id


def consume_session_token(token_id: str):
    """
    Validates and consumes (deletes) a token from S3.
    Returns user dict on success, None on failure/expiry.
    """
    tokens = _load_tokens()
    tokens = _purge_expired(tokens)

    if token_id not in tokens:
        _save_tokens(tokens)
        return None

    user_data = tokens.pop(token_id)
    _save_tokens(tokens)

    return {
        "username":     user_data["username"],
        "display_name": user_data["display_name"],
        "role":         user_data["role"],
        "hotels":       user_data["hotels"],
    }


# ── Calendar session tokens ───────────────────────────────────────────────────

def _load_cal_tokens() -> dict:
    data = read_json_from_s3(CAL_TOKENS_KEY)
    return data if isinstance(data, dict) else {}


def _save_cal_tokens(tokens: dict) -> None:
    write_json_to_s3(CAL_TOKENS_KEY, tokens)


def create_cal_token(user: dict) -> str:
    """Create a 30-minute session token for cal_date page reloads."""
    token_id = str(uuid.uuid4())
    tokens   = _load_cal_tokens()
    now      = time.time()
    tokens   = {k: v for k, v in tokens.items()
                if now - v["created_at"] < CAL_TOKEN_TTL}
    tokens[token_id] = {
        "username":     user["username"],
        "display_name": user["display_name"],
        "role":         user["role"],
        "hotels":       user["hotels"],
        "created_at":   now,
    }
    _save_cal_tokens(tokens)
    return token_id


def validate_cal_token(token_id: str) -> dict | None:
    """Validate a cal token. Returns user dict or None. Does NOT consume it."""
    tokens = _load_cal_tokens()
    now    = time.time()
    tokens = {k: v for k, v in tokens.items()
              if now - v["created_at"] < CAL_TOKEN_TTL}
    _save_cal_tokens(tokens)
    data = tokens.get(token_id)
    if not data:
        return None
    return {
        "username":     data["username"],
        "display_name": data["display_name"],
        "role":         data["role"],
        "hotels":       data["hotels"],
    }
