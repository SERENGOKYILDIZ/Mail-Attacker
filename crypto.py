import base64

# Unique static key for lightweight obfuscation (prevents casual plaintext reading)
_APP_KEY = b"mailflow_secure_key_2026_xyz!"

def encrypt(plaintext: str) -> str:
    """Encrypts a string using basic XOR masking and base64 encoding."""
    if not plaintext:
        return ""
    try:
        pt_bytes = plaintext.encode('utf-8')
        masked = bytearray(
            pt_bytes[i] ^ _APP_KEY[i % len(_APP_KEY)] 
            for i in range(len(pt_bytes))
        )
        return "ENC:" + base64.b64encode(masked).decode('utf-8')
    except Exception:
        return ""

def decrypt(ciphertext: str) -> str:
    """Decrypts a previously encrypted string via XOR unmasking."""
    if not ciphertext or not str(ciphertext).startswith("ENC:"):
        return ciphertext  # Return as-is if it's plaintext (for backward compatibility)
    try:
        raw_b64 = ciphertext[4:]
        masked = base64.b64decode(raw_b64)
        pt_bytes = bytearray(
            masked[i] ^ _APP_KEY[i % len(_APP_KEY)] 
            for i in range(len(masked))
        )
        return pt_bytes.decode('utf-8')
    except Exception:
        return ""
