"""Dump the notebook/section/page tree with LastModifiedTime from OneNote's FullTextSearchIndex.

Usage:  python inspect_search_index.py
Read-only: opens each *.db with mode=ro. Works whether or not OneNote is running.
See docs/reference-data-sources.md §1.
"""
import datetime
import glob
import os
import sqlite3

INDEX_DIR = os.path.join(os.environ["LOCALAPPDATA"], r"Microsoft\OneNote\16.0\FullTextSearchIndex")
TYPE_NAME = {4: "notebook", 3: "group", 2: "section", 1: "page"}


def filetime(v):
    if not v:
        return "-"
    return (datetime.datetime(1601, 1, 1) + datetime.timedelta(microseconds=v / 10)).strftime("%Y-%m-%d %H:%M:%S UTC")


def main():
    for db in sorted(glob.glob(os.path.join(INDEX_DIR, "*.db"))):
        con = sqlite3.connect(f"file:{db}?mode=ro", uri=True, timeout=2)
        rows = con.execute("select Type, GOID, GOSID, ParentGOID, Title, LastModifiedTime from Entities").fetchall()
        con.close()
        by_parent = {}
        for r in rows:
            by_parent.setdefault(r[3] or "", []).append(r)
        root = next((r for r in rows if r[0] == 4), None)
        if not root:
            continue
        newest = max(r[5] or 0 for r in rows)
        print(f"\n##### {os.path.basename(db)}  ({len(rows)} entities, newest change {filetime(newest)})")

        def walk(parent_goid, depth):
            for r in sorted(by_parent.get(parent_goid, []), key=lambda x: -(x[5] or 0)):
                t, goid, gosid, _, title, lm = r
                if t == 1 and depth > 2:
                    continue  # keep output short: pages listed only under top-level sections
                print(f"{'  ' * depth}{TYPE_NAME.get(t, t):8} {filetime(lm)}  {title or ''!s:40.40}  {gosid}")
                walk(goid, depth + 1)

        print(f"notebook {filetime(root[5])}  {root[4] or ''}  {root[2]}")
        walk(root[1], 1)


if __name__ == "__main__":
    main()
