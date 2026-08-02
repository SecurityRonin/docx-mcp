"""OOXML w:documentProtection password hashing.

Evidence tiering (see CLAUDE.md "Evidence-Based Rigor"):

  T1  TestWordAuthoredVector — the artifact was authored by Microsoft Word and
      the answer key by Apache POI's test suite. Neither is ours. This is the
      vector that matters: it fails if the legacy XOR pre-hash is skipped AND
      it fails if the spin-loop iterator is prepended instead of appended, so
      it pins both degrees of freedom at once.

  T1  TestLegacyXorVerifier — expected values asserted by POI's
      TestDocumentProtection, independent of us.

  T3  TestDerivedProperties — our implementation checked against itself for
      structural properties (truncation, spin-count edges). Legitimate only
      because the value-producing path is already pinned by T1 above.
"""

from __future__ import annotations

import base64
import hashlib

import pytest

from docx_mcp.document.passwordhash import (
    CRYPT_PROVIDER_TYPES,
    SID_TO_ALGORITHM,
    ooxml_password_hash,
    xor_hash_password,
    xor_hash_password_reversed,
)

# ── T1 fixture ──────────────────────────────────────────────────────────────
# Source: word/settings.xml of apache/poi test-data/document/bug56076.docx,
#   <w:documentProtection w:edit="trackedChanges" w:enforcement="1"
#    w:cryptProviderType="rsaFull" w:cryptAlgorithmClass="hash"
#    w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="4"
#    w:cryptSpinCount="100000" w:hash="MUHbcmpC9AnlLsd9v3lW0j30y6E="
#    w:salt="2Z+i7o/0EZyUNakVeWzU/w=="/>
# Answer key: poi-ooxml TestDocumentProtection.bug56076_read(), which asserts
#   document.validateProtectionPassword("Example") is true.
POI_PASSWORD = "Example"
POI_SALT = base64.b64decode("2Z+i7o/0EZyUNakVeWzU/w==")
POI_SPIN_COUNT = 100000
POI_ALGORITHM = "SHA-1"
POI_EXPECTED_HASH = "MUHbcmpC9AnlLsd9v3lW0j30y6E="


class TestWordAuthoredVector:
    """T1 — Word-authored artifact, POI-authored answer key."""

    def test_matches_the_word_authored_hash(self):
        got = ooxml_password_hash(POI_PASSWORD, POI_SALT, POI_SPIN_COUNT, POI_ALGORITHM)
        assert base64.b64encode(got).decode() == POI_EXPECTED_HASH

    def test_a_wrong_password_does_not_collide(self):
        got = ooxml_password_hash("Exampl", POI_SALT, POI_SPIN_COUNT, POI_ALGORITHM)
        assert base64.b64encode(got).decode() != POI_EXPECTED_HASH

    def test_skipping_the_xor_prehash_would_fail(self):
        """Guards the stage the original implementation was missing entirely."""
        digest = hashlib.sha1(POI_SALT + POI_PASSWORD.encode("utf-16-le")).digest()
        for i in range(POI_SPIN_COUNT):
            digest = hashlib.sha1(digest + i.to_bytes(4, "little")).digest()
        assert base64.b64encode(digest).decode() != POI_EXPECTED_HASH

    def test_prepending_the_iterator_would_fail(self):
        """Guards the other degree of freedom the vector discriminates."""
        pw = xor_hash_password_reversed(POI_PASSWORD).encode("utf-16-le")
        digest = hashlib.sha1(POI_SALT + pw).digest()
        for i in range(POI_SPIN_COUNT):
            digest = hashlib.sha1(i.to_bytes(4, "little") + digest).digest()
        assert base64.b64encode(digest).decode() != POI_EXPECTED_HASH

    def test_salt_must_be_the_decoded_bytes_not_the_base64_text(self):
        b64_text = b"2Z+i7o/0EZyUNakVeWzU/w=="
        got = ooxml_password_hash(POI_PASSWORD, b64_text, POI_SPIN_COUNT, POI_ALGORITHM)
        assert base64.b64encode(got).decode() != POI_EXPECTED_HASH


class TestLegacyXorVerifier:
    """T1 — expected values asserted by POI's TestDocumentProtection."""

    def test_example_verifier(self):
        assert xor_hash_password("Example") == "64CEED7E"

    def test_leading_zero_is_padded(self):
        """POI keeps this case because %X instead of %02X silently truncates."""
        assert xor_hash_password("34579") == "0005CB00"

    def test_reversed_form_for_example(self):
        assert xor_hash_password_reversed("Example") == "7EEDCE64"

    def test_reversed_form_is_eight_hex_chars(self):
        for pw in ("a", "34579", "Example", "correct horse battery"):
            got = xor_hash_password_reversed(pw)
            assert len(got) == 8
            assert got == got.upper()
            int(got, 16)

    def test_reversed_form_encodes_to_the_documented_bytes(self):
        """MS-OE376 §2.15.1.28 quotes this exact byte stream for "Example"."""
        got = xor_hash_password_reversed("Example").encode("utf-16-le")
        assert got == bytes.fromhex("37004500450044004300450036003400")


class TestDerivedProperties:
    """T3 — structural properties of our implementation, not independent truth."""

    def test_password_is_truncated_to_fifteen_characters(self):
        assert xor_hash_password("0123456789abcdef") == xor_hash_password("0123456789abcde")

    def test_spin_count_zero_is_the_initial_hash_only(self):
        pw = xor_hash_password_reversed("Example").encode("utf-16-le")
        expected = hashlib.sha1(POI_SALT + pw).digest()
        assert ooxml_password_hash("Example", POI_SALT, 0, "SHA-1") == expected

    def test_spin_count_one_adds_exactly_one_round(self):
        pw = xor_hash_password_reversed("Example").encode("utf-16-le")
        expected = hashlib.sha1(POI_SALT + pw).digest()
        expected = hashlib.sha1(expected + (0).to_bytes(4, "little")).digest()
        assert ooxml_password_hash("Example", POI_SALT, 1, "SHA-1") == expected

    def test_sha512_digest_length(self):
        got = ooxml_password_hash("Example", POI_SALT, 16, "SHA-512")
        assert len(got) == 64

    def test_empty_password_is_accepted(self):
        assert len(ooxml_password_hash("", POI_SALT, 16, "SHA-512")) == 64


class TestAlgorithmTable:
    def test_word_supported_sids_are_mapped(self):
        for sid in (1, 2, 3, 4, 12, 13, 14):
            assert sid in SID_TO_ALGORITHM

    def test_sha512_is_sid_fourteen(self):
        assert SID_TO_ALGORITHM[14] == "SHA-512"

    def test_sha1_is_sid_four(self):
        assert SID_TO_ALGORITHM[4] == "SHA-1"

    def test_legacy_algorithms_use_rsafull(self):
        for alg in ("MD2", "MD4", "MD5", "SHA-1"):
            assert CRYPT_PROVIDER_TYPES[alg] == "rsaFull"

    def test_sha2_algorithms_use_rsaaes(self):
        """Word rejects the pairing if the provider type does not match the sid."""
        for alg in ("SHA-256", "SHA-384", "SHA-512"):
            assert CRYPT_PROVIDER_TYPES[alg] == "rsaAES"

    def test_unknown_algorithm_names_the_offending_value(self):
        with pytest.raises(ValueError, match="SHA-999"):
            ooxml_password_hash("x", POI_SALT, 1, "SHA-999")

    def test_negative_spin_count_is_rejected(self):
        with pytest.raises(ValueError, match="-1"):
            ooxml_password_hash("x", POI_SALT, -1, "SHA-1")
