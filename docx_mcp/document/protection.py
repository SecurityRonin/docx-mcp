"""Protection mixin: document protection settings."""

from __future__ import annotations

import base64
import os

from lxml import etree

from .base import W
from .passwordhash import (
    ALGORITHM_TO_SID,
    CRYPT_PROVIDER_TYPES,
    MAX_SPIN_COUNT,
    ooxml_password_hash,
)

_SALT_BYTES = 16
_DEFAULT_SPIN_COUNT = 100_000


class ProtectionMixin:
    """Document protection operations."""

    def set_document_protection(
        self,
        edit: str,
        *,
        password: str | None = None,
        algorithm: str = "SHA-512",
        spin_count: int = _DEFAULT_SPIN_COUNT,
    ) -> dict:
        """Set document protection in settings.xml.

        Args:
            edit: Protection type — "trackedChanges", "comments", "readOnly",
                  "forms", or "none" (removes protection).
            password: Optional password. Hashed per MS-OE376 §2.15.1.28 so that
                  Word can verify it — see docx_mcp.document.passwordhash.
                  Note the legacy algorithm truncates it to 15 characters.
            algorithm: Hash algorithm name (default SHA-512). Determines both
                  w:cryptAlgorithmSid and w:cryptProviderType.
            spin_count: Iterations of the hash spin loop.
        """
        settings = self._require("word/settings.xml")

        # Remove existing protection
        for old in settings.findall(f"{W}documentProtection"):
            settings.remove(old)

        if edit == "none":
            self._mark("word/settings.xml")
            return {"edit": "none", "enforcement": "0", "has_password": False}

        has_password = False
        if password:
            # Validate before mutating the tree so a rejected call leaves
            # existing protection removed-and-not-replaced only on success.
            sid = ALGORITHM_TO_SID.get(algorithm)
            if sid is None:
                raise ValueError(
                    f"Unsupported hash algorithm {algorithm!r}. "
                    f"Word accepts: {', '.join(sorted(ALGORITHM_TO_SID))}."
                )
            if spin_count < 0 or spin_count > MAX_SPIN_COUNT:
                raise ValueError(
                    f"spin_count must be between 0 and {MAX_SPIN_COUNT}, got {spin_count}"
                )

        prot = etree.SubElement(settings, f"{W}documentProtection")
        prot.set(f"{W}edit", edit)
        prot.set(f"{W}enforcement", "1")

        if password:
            salt = os.urandom(_SALT_BYTES)
            digest = ooxml_password_hash(password, salt, spin_count, algorithm)
            prot.set(f"{W}cryptProviderType", CRYPT_PROVIDER_TYPES[algorithm])
            prot.set(f"{W}cryptAlgorithmClass", "hash")
            prot.set(f"{W}cryptAlgorithmType", "typeAny")
            prot.set(f"{W}cryptAlgorithmSid", str(ALGORITHM_TO_SID[algorithm]))
            prot.set(f"{W}cryptSpinCount", str(spin_count))
            prot.set(f"{W}hash", base64.b64encode(digest).decode())
            prot.set(f"{W}salt", base64.b64encode(salt).decode())
            has_password = True

        self._mark("word/settings.xml")

        return {
            "edit": edit,
            "enforcement": "1",
            "has_password": has_password,
        }
