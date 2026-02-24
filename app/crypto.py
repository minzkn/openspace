# SPDX-License-Identifier: MIT
# Copyright (c) 2026 JAEHYUK CHO
import os
import base64
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from argon2 import PasswordHasher
from argon2.exceptions import VerifyMismatchError

from .config import settings

_ph = PasswordHasher(time_cost=3, memory_cost=65536, parallelism=4)


class CryptoService:
    def __init__(self, kek_hex: str):
        key_bytes = bytes.fromhex(kek_hex)
        if len(key_bytes) != 32:
            raise ValueError("KEK_KEY must be 64 hex chars (32 bytes)")
        self._kek = key_bytes

    # ------------------------------------------------------------------
    # Password hashing (Argon2id)
    # ------------------------------------------------------------------
    def hash_password(self, password: str) -> str:
        return _ph.hash(password)

    def verify_password(self, stored_hash: str, password: str) -> bool:
        try:
            return _ph.verify(stored_hash, password)
        except Exception:
            return False

    def needs_rehash(self, stored_hash: str) -> bool:
        return _ph.check_needs_rehash(stored_hash)

    # ------------------------------------------------------------------
    # AES-256-GCM field encryption
    # ------------------------------------------------------------------
    def encrypt(self, plaintext: str) -> str:
        nonce = os.urandom(12)
        aesgcm = AESGCM(self._kek)
        ct = aesgcm.encrypt(nonce, plaintext.encode("utf-8"), None)
        return base64.b64encode(nonce + ct).decode("ascii")

    def decrypt(self, ciphertext_b64: str) -> str:
        data = base64.b64decode(ciphertext_b64)
        nonce, ct = data[:12], data[12:]
        aesgcm = AESGCM(self._kek)
        return aesgcm.decrypt(nonce, ct, None).decode("utf-8")

    def is_encrypted(self, value: str) -> bool:
        """Check if value looks like an encrypted blob (base64, len >= 40)."""
        try:
            decoded = base64.b64decode(value)
            return len(decoded) >= 28  # 12 nonce + 16 tag minimum
        except Exception:
            return False


# Singleton
crypto = CryptoService(settings.kek_key)
