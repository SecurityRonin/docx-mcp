"""Protection mixin: document protection settings."""

from __future__ import annotations

import base64
import hashlib
import os

from lxml import etree

from .base import W


class ProtectionMixin:
    """Document protection operations."""

    def set_document_protection(
        self,
        edit: str,
        *,
        password: str | None = None,
    ) -> dict:
        """Set document protection in settings.xml.

        Args:
            edit: Protection type — "trackedChanges", "comments", "readOnly",
                  "forms", or "none" (removes protection).
            password: Optional password. Hashed with SHA-512 per OOXML spec.
        """
        settings = self._require("word/settings.xml")

        # Remove existing protection
        for old in settings.findall(f"{W}documentProtection"):
            settings.remove(old)

        if edit == "none":
            self._mark("word/settings.xml")
            return {"edit": "none", "enforcement": "0", "has_password": False}

        prot = etree.SubElement(settings, f"{W}documentProtection")
        prot.set(f"{W}edit", edit)
        prot.set(f"{W}enforcement", "1")

        has_password = False
        if password:
            # OOXML SHA-512 password hashing
            salt = os.urandom(16)
            hash_val = hashlib.sha512(salt + password.encode("utf-16-le")).digest()
            for _ in range(100000):
                hash_val = hashlib.sha512(hash_val + salt).digest()
            prot.set(f"{W}cryptAlgorithmClass", "hash")
            prot.set(f"{W}cryptAlgorithmType", "typeAny")
            prot.set(f"{W}cryptAlgorithmSid", "14")  # SHA-512
            prot.set(f"{W}cryptSpinCount", "100000")
            prot.set(f"{W}hash", base64.b64encode(hash_val).decode())
            prot.set(f"{W}salt", base64.b64encode(salt).decode())
            has_password = True

        self._mark("word/settings.xml")

        return {
            "edit": edit,
            "enforcement": "1",
            "has_password": has_password,
        }
