"""Sheet and workbook protection: the password as the file keeps it (a hash), and the options a sheet is protected
with. What protection refuses is asked by the operations themselves (Sheet.guard)."""

import base64
import hashlib

from ..errors import VBANotImplemented, VBARuntimeError
from ..values import VBAObject

# Protect's Allow... arguments, and the attribute of <sheetProtection> that is "0" when the option is on: the file
# names what protection forbids, VBA what it allows.
ALLOW = {
    "allowformattingcells": "formatCells",
    "allowformattingcolumns": "formatColumns",
    "allowformattingrows": "formatRows",
    "allowinsertingcolumns": "insertColumns",
    "allowinsertingrows": "insertRows",
    "allowinsertinghyperlinks": "insertHyperlinks",
    "allowdeletingcolumns": "deleteColumns",
    "allowdeletingrows": "deleteRows",
    "allowsorting": "sort",
    "allowfiltering": "autoFilter",
    "allowusingpivottables": "pivotTables",
}
SPIN = 100_000  # the rounds of hashing Excel gives a password it saves


def legacy_hash(password):
    """The 16-bit hash Excel saved passwords as before 2013, and still reads."""
    h = 0
    for i, ch in enumerate(password, 1):
        v = ord(ch) << i
        h ^= (v & 0x7FFF) | (v >> 15)
    return f"{h ^ len(password) ^ 0xCE4B:04X}"


def sha_hash(password, salt, spin, algorithm="SHA-512"):
    """The salted hash Excel saves a password as: the hash of salt and password, hashed again with each round's
    number, `spin` times."""
    name = algorithm.replace("-", "").lower()
    if name not in ("sha512", "sha384", "sha256", "sha1"):
        raise VBANotImplemented(f"a password kept with the {algorithm} algorithm")
    h = hashlib.new(name, salt + password.encode("utf-16-le")).digest()
    for i in range(spin):
        h = hashlib.new(name, h + i.to_bytes(4, "little")).digest()
    return base64.b64encode(h).decode()


def _attr(base, prefix):
    """The file's name for a part of a password: hashValue in <sheetProtection>, workbookHashValue in the book's."""
    return prefix + base if prefix else base[0].lower() + base[1:]


_PARTS = ("Password", "AlgorithmName", "HashValue", "SaltValue", "SpinCount")


class Password:
    """A protection's password: as the file keeps it (attrs, its hash), or as a macro gave it (text)."""

    def __init__(self, text=None, attrs=None):
        self.text, self.attrs = text, attrs or {}

    @classmethod
    def read(cls, attrs, prefix=""):
        """From <sheetProtection> (prefix "") or <workbookProtection> (prefix "workbook"); None when it has none."""
        kept = {_attr(p, prefix): attrs[_attr(p, prefix)] for p in _PARTS if _attr(p, prefix) in attrs}
        return cls(attrs=kept) if kept else None

    def matches(self, text, prefix=""):
        if self.text is not None:
            return text == self.text
        part = {p: self.attrs.get(_attr(p, prefix)) for p in _PARTS}
        if part["HashValue"] is not None:
            salt = base64.b64decode(part["SaltValue"] or "")
            hashed = sha_hash(text, salt, int(part["SpinCount"] or 0), part["AlgorithmName"] or "SHA-512")
            return hashed == part["HashValue"]
        return part["Password"] is not None and legacy_hash(text) == part["Password"].upper()

    def to_attrs(self, salt, prefix=""):
        """The attributes the file keeps the password in; a password a macro gave is hashed as Excel hashes it."""
        if self.text is None:
            return dict(self.attrs)
        values = ("SHA-512", sha_hash(self.text, salt, SPIN), base64.b64encode(salt).decode(), str(SPIN))
        return {_attr(p, prefix): v for p, v in zip(_PARTS[1:], values, strict=True)}


class SheetProtection:
    """A sheet's protection: its password (or None) and the Allow... options on. `kept` is the element's attributes
    as the file had them, for as long as nothing about the protection changes."""

    def __init__(self, password=None, allow=(), objects=True, scenarios=True, kept=None):
        self.password, self.allow, self.objects, self.scenarios = password, frozenset(allow), objects, scenarios
        self.kept = kept

    @classmethod
    def read(cls, attrs):
        if attrs.get("sheet") not in ("1", "true"):
            return None
        allow = {vba for vba, xml in ALLOW.items() if attrs.get(xml) in ("0", "false")}
        objects, scenarios = (attrs.get(name) in ("1", "true") for name in ("objects", "scenarios"))
        return cls(Password.read(attrs), allow, objects, scenarios, kept=dict(attrs))

    def attrs(self, salt):
        if self.kept is not None:
            return self.kept
        out = self.password.to_attrs(salt) if self.password is not None else {}
        out["sheet"] = "1"
        if self.objects:
            out["objects"] = "1"
        if self.scenarios:
            out["scenarios"] = "1"
        for vba, xml in ALLOW.items():
            if vba in self.allow:
                out[xml] = "0"
        return out


class BookProtection:
    """A workbook's protection: its password (or None), and whether it locks the structure (adding, deleting, renaming,
    hiding and showing sheets) and the windows. `kept` as for SheetProtection."""

    def __init__(self, password=None, structure=True, windows=False, kept=None):
        self.password, self.structure, self.windows, self.kept = password, structure, windows, kept

    @classmethod
    def read(cls, attrs):
        structure, windows = (attrs.get(name) in ("1", "true") for name in ("lockStructure", "lockWindows"))
        password = Password.read(attrs, "workbook")
        if not (structure or windows):
            return None
        return cls(password, structure, windows, kept=dict(attrs))

    def attrs(self, salt):
        if self.kept is not None:
            return self.kept
        out = self.password.to_attrs(salt, "workbook") if self.password is not None else {}
        if self.structure:
            out["lockStructure"] = "1"
        if self.windows:
            out["lockWindows"] = "1"
        return out


def check_password(password, given, prefix=""):
    """Unprotect: the password given must be the protection's, or error 1004."""
    if password is not None and not password.matches(given, prefix):
        raise VBARuntimeError(1004, "The password you supplied is not correct")


class ProtectionObject(VBAObject):
    """Worksheet.Protection: the Allow... options the sheet is protected with."""

    vba_type_name = "Protection"

    def __init__(self, sheet):
        self.sheet = sheet

    def allows(self, option):
        protection = self.sheet.protection
        return protection is not None and option in protection.allow


for _option in ALLOW:
    setattr(ProtectionObject, "get_" + _option, lambda self, _o=_option: self.allows(_o))
