import json
import os
import base64
import getpass

from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM


# ---- Configuration ----
ITERATIONS = 150_000
SALT_SIZE = 16        # bytes
IV_SIZE = 12          # bytes (recommended for GCM)
KEY_SIZE = 32         # 256-bit AES


def derive_key(password: str, salt: bytes) -> bytes:
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=KEY_SIZE,
        salt=salt,
        iterations=ITERATIONS,
    )
    return kdf.derive(password.encode("utf-8"))


def encrypt_json(input_path: str, output_path: str):
    password = getpass.getpass("Enter encryption password: ")
    confirm = getpass.getpass("Confirm password: ")

    if password != confirm:
        raise ValueError("Passwords do not match")

    with open(input_path, "r", encoding="utf-8") as f:
        plaintext = f.read().encode("utf-8")

    salt = os.urandom(SALT_SIZE)
    iv = os.urandom(IV_SIZE)

    key = derive_key(password, salt)
    aesgcm = AESGCM(key)

    ciphertext = aesgcm.encrypt(iv, plaintext, None)

    encrypted_payload = {
        "kdf": "PBKDF2",
        "hash": "SHA-256",
        "iterations": ITERATIONS,
        "salt": base64.b64encode(salt).decode("utf-8"),
        "iv": base64.b64encode(iv).decode("utf-8"),
        "ciphertext": base64.b64encode(ciphertext).decode("utf-8"),
    }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(encrypted_payload, f, indent=2)

    print(f"Encrypted file written to: {output_path}")


if __name__ == "__main__":
    encrypt_json(
        input_path="private.json",
        output_path="private.json.enc",
    )
