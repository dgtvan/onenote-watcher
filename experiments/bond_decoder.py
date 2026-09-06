"""Decode 1DS Common-Schema records (Bond Compact Binary v1) from OneNote's OTele telemetry store.

Usage:
    python bond_decoder.py                 # scans %LOCALAPPDATA%\Microsoft\Office\OTele\onenote.exe.db-wal
    python bond_decoder.py <file>          # any file containing records (a .db, a -wal, a raw dump)

Read-only. Prints one block per sync-related record with its Data.* properties decoded.
See docs/reference-data-sources.md for the format.
"""
import datetime
import os
import re
import struct
import sys

(BT_STOP, BT_STOP_BASE, BT_BOOL, BT_UINT8, BT_UINT16, BT_UINT32, BT_UINT64, BT_FLOAT, BT_DOUBLE,
 BT_STRING, BT_STRUCT, BT_LIST, BT_SET, BT_MAP, BT_INT8, BT_INT16, BT_INT32, BT_INT64, BT_WSTRING) = range(19)

RECORD_START = re.compile(rb"\x29\x033\.0\x49")   # field1 string "3.0", then field2 (name) header
SYNC_EVENTS = ("NotebookSyncResult", "SyncScore", "PageSyncSession", "ErrorsWithTag",
               "StuckContentError", "ReadOnlyTriggeredByError")


class Reader:
    def __init__(self, buf, pos=0):
        self.b, self.i = buf, pos

    def u8(self):
        v = self.b[self.i]; self.i += 1; return v

    def varint(self):
        r = sh = 0
        while True:
            c = self.u8(); r |= (c & 0x7F) << sh; sh += 7
            if not c & 0x80:
                return r

    def zigzag(self):
        v = self.varint(); return (v >> 1) ^ -(v & 1)

    def field_header(self):
        h = self.u8(); t, fid = h & 0x1F, h >> 5
        if fid == 6: fid = self.u8()
        elif fid == 7: fid = self.u8() | (self.u8() << 8)
        return fid, t

    def value(self, t):
        if t == BT_BOOL: return bool(self.u8())
        if t == BT_UINT8: return self.u8()
        if t in (BT_UINT16, BT_UINT32, BT_UINT64): return self.varint()
        if t in (BT_INT16, BT_INT32, BT_INT64): return self.zigzag()
        if t == BT_INT8: return struct.unpack("b", bytes([self.u8()]))[0]
        if t == BT_FLOAT: v = struct.unpack_from("<f", self.b, self.i)[0]; self.i += 4; return v
        if t == BT_DOUBLE: v = struct.unpack_from("<d", self.b, self.i)[0]; self.i += 8; return v
        if t == BT_STRING: n = self.varint(); v = self.b[self.i:self.i + n].decode("utf-8", "replace"); self.i += n; return v
        if t == BT_WSTRING: n = self.varint(); v = self.b[self.i:self.i + 2 * n].decode("utf-16le", "replace"); self.i += 2 * n; return v
        if t == BT_STRUCT: return self.struct()
        if t in (BT_LIST, BT_SET):
            h = self.u8(); et, hi = h & 0x1F, h >> 5
            n = (hi - 1) if hi else self.varint()
            if et in (BT_UINT8, BT_INT8):
                v = bytes(self.b[self.i:self.i + n]); self.i += n; return v
            return [self.value(et) for _ in range(n)]
        if t == BT_MAP:
            kt = self.u8() & 0x1F; vt = self.u8() & 0x1F; n = self.varint()
            return {self.value(kt): self.value(vt) for _ in range(n)}
        raise ValueError(f"unknown bond type {t} at offset {self.i}")

    def struct(self):
        d = {}
        while True:
            fid, t = self.field_header()
            if t == BT_STOP: return d
            if t == BT_STOP_BASE: continue
            d[fid] = self.value(t)


def cs_value(v):
    """CsProtocol `Value` struct -> Python scalar. kind(1): 6=bool; string(3) int(4) double(5) guid(6)."""
    if not isinstance(v, dict):
        return v
    kind = v.get(1, 0)
    if 3 in v: return v[3]
    if 4 in v: return bool(v[4]) if kind == 6 else v[4]
    if 5 in v: return v[5]
    if 6 in v: return v[6]
    return False if kind == 6 else 0


def dotnet_ticks(t):
    if isinstance(t, int) and t > 10 ** 17:
        return (datetime.datetime(1, 1, 1) + datetime.timedelta(microseconds=t / 10)).strftime("%Y-%m-%d %H:%M:%S UTC")
    return t


def decode_all(buf):
    """Yield (offset, event_name, time, properties) for every decodable record in buf."""
    for m in RECORD_START.finditer(buf):
        try:
            rec = Reader(buf, m.start()).struct()
        except Exception:
            continue
        name = rec.get(2)
        if not isinstance(name, str):
            continue
        props = {}
        for data in rec.get(70, []) if isinstance(rec.get(70), list) else []:
            if isinstance(data, dict) and isinstance(data.get(1), dict):
                props.update({k: cs_value(v) for k, v in data[1].items()})
        yield m.start(), name, rec.get(3), props


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.environ["LOCALAPPDATA"], r"Microsoft\Office\OTele\onenote.exe.db-wal")
    buf = open(path, "rb").read()
    n = 0
    for off, name, t, props in decode_all(buf):
        if not any(k in name for k in SYNC_EVENTS):
            continue
        n += 1
        print(f"\n=== @{off} {name}  time={dotnet_ticks(t)}")
        for k in sorted(props):
            if k.startswith("Data."):
                v = props[k]
                if k.endswith("Error_Code") and isinstance(v, int):
                    v = f"{v} (0x{v:08X})"
                elif isinstance(v, int) and v > 10 ** 17:
                    v = f"{v} ({dotnet_ticks(v)})"
                print(f"  {k} = {v!r}")
    print(f"\n{n} sync-related records decoded from {path}")


if __name__ == "__main__":
    main()
