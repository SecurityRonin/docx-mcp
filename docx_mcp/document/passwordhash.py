"""Password hashing for w:documentProtection / w:writeProtection.

Word stores a protection password as a salted, iterated hash — but it does
not hash the password. It first reduces the password through the Word 97 XOR
verifier, byte-reverses that 32-bit key, renders it as an 8-character
uppercase hex string, and hashes *that* string as UTF-16LE:

    legacy  = xor_hash_password_reversed(password)      # e.g. "7EEDCE64"
    H       = Hash(salt + legacy.encode("utf-16-le"))
    for i in range(spin_count):
        H   = Hash(H + LE32(i))
    w:hash  = base64(H)

The initial hash is *not* counted against the spin count, and the iteration
counter is appended to the running digest, not prepended (prepending is the
agile-encryption key derivation, a different algorithm that happens to share
this shape — Apache POI guards the distinction with an `iteratorFirst` flag).

References
----------
MS-OE376 Part 4 §2.15.1.28 `documentProtection` — the third (hex-string)
stage and the 5,000,000 spin ceiling.
MS-OI29500 Part 1 §17.15.1.29 — the salt is prepended in its binary form,
never as its base64 text.
MS-OFFCRYPTO §2.3.7.1 / §2.3.7.4 — Binary Document Password Verifier
Derivation Methods 1 and 2, the source of the constants below.

This module deliberately mirrors the observable behaviour of Apache POI's
`org.apache.poi.poifs.crypt.CryptoFunctions`, which is the implementation
validated against Word in `tests/test_password_hash.py`. That includes POI's
Java signed-byte arithmetic, which is observable for password bytes >= 0x80
(see `_xor_verifier1`).

ECMA-376 also defines an ISO-conformant attribute set (`w:algorithmName`,
`w:hashValue`, `w:saltValue`, `w:spinCount`) which omits the XOR stage and
hashes the raw password. Word only writes it when a registry opt-in is set,
so it is not implemented here.
"""

from __future__ import annotations

import hashlib

# MS-OFFCRYPTO §2.3.7.1, keyed by password length (1-15).
_INITIAL_CODE_ARRAY = (
    0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE,
    0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A,
    0x4EC3,
)  # fmt: skip

# MS-OFFCRYPTO §2.3.7.1, one row of 7 words per password position.
_ENCRYPTION_MATRIX = (
    (0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09),
    (0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF),
    (0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0),
    (0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40),
    (0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5),
    (0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A),
    (0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9),
    (0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0),
    (0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC),
    (0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10),
    (0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168),
    (0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C),
    (0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD),
    (0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC),
    (0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4),
)

_MAX_PASSWORD_LENGTH = 15

# w:cryptAlgorithmSid values Word understands (MS-OE376 §2.15.1.28).
SID_TO_ALGORITHM: dict[int, str] = {
    1: "MD2",
    2: "MD4",
    3: "MD5",
    4: "SHA-1",
    12: "SHA-256",
    13: "SHA-384",
    14: "SHA-512",
}

ALGORITHM_TO_SID: dict[str, int] = {name: sid for sid, name in SID_TO_ALGORITHM.items()}

# w:cryptProviderType must match the algorithm family or Word rejects the pair.
CRYPT_PROVIDER_TYPES: dict[str, str] = {
    "MD2": "rsaFull",
    "MD4": "rsaFull",
    "MD5": "rsaFull",
    "SHA-1": "rsaFull",
    "SHA-256": "rsaAES",
    "SHA-384": "rsaAES",
    "SHA-512": "rsaAES",
}

_HASHLIB_NAMES: dict[str, str] = {
    "MD2": "md2",
    "MD4": "md4",
    "MD5": "md5",
    "SHA-1": "sha1",
    "SHA-256": "sha256",
    "SHA-384": "sha384",
    "SHA-512": "sha512",
}

# MS-OE376 §2.15.1.28: the legacy spin count is bounded at 5,000,000.
MAX_SPIN_COUNT = 5_000_000


def _utf16_code_units(text: str) -> list[int]:
    """UTF-16 code units, matching Java's String.charAt over the same string."""
    raw = text.encode("utf-16-le", errors="surrogatepass")
    return [raw[i] | (raw[i + 1] << 8) for i in range(0, len(raw), 2)]


def _to_ansi_password(units: list[int]) -> list[int]:
    """Collapse each code unit to one byte: low byte, or high byte if low is 0.

    Lossy above U+00FF by design — this is the Word 97 algorithm, not a
    character encoding.
    """
    out = []
    for unit in units:
        low = unit & 0xFF
        out.append(low if low != 0 else (unit >> 8) & 0xFF)
    return out


def _as_signed_byte(value: int) -> int:
    """Reinterpret the low 8 bits as a Java `byte` (signed)."""
    value &= 0xFF
    return value - 0x100 if value >= 0x80 else value


def _rotate_left_base15(verifier: int) -> int:
    """MS-OFFCRYPTO §2.3.7.1 rotate: 15-bit left rotation through bit 14."""
    carry = 1 if verifier & 0x4000 else 0
    return carry | ((verifier << 1) & 0x7FFF)


def _xor_verifier1(ansi: list[int]) -> int:
    """Binary Document Password Verifier Derivation Method 1 — the low word.

    The XOR against each password byte is done with Java signed-byte
    semantics: a byte >= 0x80 sign-extends and flips the verifier's high
    bits, which the next rotation then propagates. That is observable for
    extended-ANSI passwords, so it is reproduced rather than normalised away.
    """
    if not ansi:
        return 0
    verifier = 0
    for byte in reversed(ansi):
        verifier = _rotate_left_base15(verifier)
        verifier = (verifier ^ _as_signed_byte(byte)) & 0xFFFF
    verifier = _rotate_left_base15(verifier)
    verifier = (verifier ^ len(ansi)) & 0xFFFF
    return (verifier ^ 0xCE4B) & 0xFFFF


def _xor_verifier2(password: str) -> int:
    """Derivation Method 2 — the full 32-bit key (high word << 16 | low word)."""
    units = _utf16_code_units(password)[:_MAX_PASSWORD_LENGTH]
    if not units:
        return 0
    ansi = _to_ansi_password(units)

    high = _INITIAL_CODE_ARRAY[len(ansi) - 1]
    row = _MAX_PASSWORD_LENGTH - len(ansi)
    for byte in ansi:
        char = _as_signed_byte(byte)
        for word in _ENCRYPTION_MATRIX[row]:
            if char & 1:
                high ^= word
            # Java `ch >>>= 1` on a byte: sign-extend, unsigned-shift, renarrow.
            char = _as_signed_byte((char & 0xFFFFFFFF) >> 1)
        row += 1

    return ((high & 0xFFFF) << 16) | _xor_verifier1(ansi)


def xor_hash_password(password: str) -> str:
    """The Word 97 password key as 8 uppercase hex digits, in natural order."""
    return f"{_xor_verifier2(password):08X}"


def xor_hash_password_reversed(password: str) -> str:
    """The Word 97 password key byte-reversed, as 8 uppercase hex digits.

    This string — not the password — is what Word feeds to the salted hash.
    Zero padding is load-bearing: a byte below 0x10 rendered as one nibble
    yields a different, silently wrong hash.
    """
    key = _xor_verifier2(password)
    return (
        f"{key & 0xFF:02X}{(key >> 8) & 0xFF:02X}{(key >> 16) & 0xFF:02X}{(key >> 24) & 0xFF:02X}"
    )


def _new_hasher(algorithm: str):
    name = _HASHLIB_NAMES.get(algorithm)
    if name is None:
        raise ValueError(
            f"Unsupported hash algorithm {algorithm!r}. "
            f"Word accepts: {', '.join(sorted(_HASHLIB_NAMES))}."
        )
    try:
        return hashlib.new(name)
    except ValueError as exc:
        raise ValueError(
            f"Hash algorithm {algorithm!r} is not available in this Python build "
            f"(hashlib name {name!r}): {exc}"
        ) from exc


def ooxml_password_hash(
    password: str,
    salt: bytes,
    spin_count: int,
    algorithm: str = "SHA-512",
) -> bytes:
    """Compute the w:hash value for w:documentProtection.

    Args:
        password: The plaintext password. Silently truncated to 15 characters
            by the legacy verifier stage — that is the algorithm, not a bug.
        salt: The raw salt bytes. Must be the decoded value, never the base64
            text of w:salt.
        spin_count: Iterations of the spin loop. The initial hash is extra and
            is not counted, so 0 is legal and yields that initial hash alone.
        algorithm: One of the names in SID_TO_ALGORITHM.

    Returns:
        The raw digest. Callers base64-encode it for the w:hash attribute.
    """
    if spin_count < 0 or spin_count > MAX_SPIN_COUNT:
        raise ValueError(f"spin_count must be between 0 and {MAX_SPIN_COUNT}, got {spin_count}")
    _new_hasher(algorithm)  # validate before doing any work

    legacy = xor_hash_password_reversed(password).encode("utf-16-le")

    hasher = _new_hasher(algorithm)
    hasher.update(salt)
    hasher.update(legacy)
    digest = hasher.digest()

    for i in range(spin_count):
        hasher = _new_hasher(algorithm)
        hasher.update(digest)
        hasher.update(i.to_bytes(4, "little"))
        digest = hasher.digest()

    return digest
