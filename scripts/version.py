#!/usr/bin/env python3
"""
version.py — Snapshot-based version control for Comply spec workflow

USAGE:
    python3 scripts/version.py snap [tag]            # quick snap: xlsx + SKILL.md
    python3 scripts/version.py snap-full [tag]       # full snap: + output/ tar.gz
    python3 scripts/version.py list                  # list snapshots
    python3 scripts/version.py show <id>             # show snapshot details
    python3 scripts/version.py diff <id1> <id2>      # compare 2 snapshots
    python3 scripts/version.py restore <id>          # restore quick (xlsx + SKILL.md)
    python3 scripts/version.py restore-full <id>     # restore full output/
    python3 scripts/version.py prune --keep N        # keep last N snapshots only
    python3 scripts/version.py auto-snap             # snap if xlsx changed since last

Snapshots stored in _versions/snapshots/<id>/ where <id> = YYYY-MM-DD_HHMMSS[_tag]

Examples:
    python3 scripts/version.py snap "before R349 work"
    python3 scripts/version.py snap-full "release-v1"
    python3 scripts/version.py list
    python3 scripts/version.py diff 2026-05-09_1130 2026-05-09_1200
    python3 scripts/version.py restore 2026-05-09_1130
"""
import argparse
import hashlib
import json
import os
import shutil
import sys
import tarfile
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VERSIONS_DIR = ROOT / "_versions"
SNAPS_DIR = VERSIONS_DIR / "snapshots"
TRACKED_QUICK = [
    ROOT / "output" / "Comply spec Smart Plant 1.xlsx",
    ROOT / "SKILL.md",
]
TRACKED_QUICK_DIRS = [
    ROOT / "knowledge_base",  # KB folder — small, important for agent development
    ROOT / "_continuity",     # session continuity state files — critical for resume
]
EXCLUDE_FROM_FULL = [
    "_archive",
    ".DS_Store",
]


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def make_id(tag: str = "") -> str:
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    if tag:
        clean = "".join(c if c.isalnum() or c in "-_" else "-" for c in tag.strip())[:50]
        return f"{ts}_{clean}"
    return ts


def cmd_snap(args, full: bool = False):
    SNAPS_DIR.mkdir(parents=True, exist_ok=True)
    snap_id = make_id(args.tag or "")
    snap_dir = SNAPS_DIR / snap_id
    snap_dir.mkdir()

    manifest = {
        "id": snap_id,
        "timestamp": datetime.now().isoformat(),
        "tag": args.tag or "",
        "kind": "full" if full else "quick",
        "files": {},
    }

    # Copy tracked files (xlsx + SKILL.md)
    for src in TRACKED_QUICK:
        if not src.exists():
            print(f"  ⚠ skip (missing): {src.relative_to(ROOT)}")
            continue
        rel = str(src.relative_to(ROOT))
        dst = snap_dir / src.name
        shutil.copy2(src, dst)
        manifest["files"][rel] = {
            "size": src.stat().st_size,
            "sha256": sha256_file(src),
            "snapped_as": src.name,
        }
        print(f"  ✓ {rel} ({src.stat().st_size:,} bytes)")

    # Copy tracked dirs (knowledge_base) — small + important
    for src_dir in TRACKED_QUICK_DIRS:
        if not src_dir.exists():
            continue
        rel = str(src_dir.relative_to(ROOT))
        dst = snap_dir / src_dir.name
        shutil.copytree(src_dir, dst, dirs_exist_ok=True)
        # record file list + sizes
        files_info = {}
        for f in src_dir.rglob("*"):
            if f.is_file() and not f.name.startswith(".") and not f.name.startswith("~$"):
                f_rel = str(f.relative_to(ROOT))
                files_info[f_rel] = {
                    "size": f.stat().st_size,
                    "sha256": sha256_file(f),
                }
        manifest["files"][rel + "/"] = {
            "kind": "directory",
            "files": files_info,
            "total_size": sum(v["size"] for v in files_info.values()),
        }
        print(f"  ✓ {rel}/ ({len(files_info)} files, {sum(v['size'] for v in files_info.values()):,} bytes)")

    if full:
        # tar.gz the entire output/ folder (excluding _archive, etc.)
        out_src = ROOT / "output"
        tar_path = snap_dir / "output.tar.gz"
        print("  • compressing output/ → output.tar.gz...")

        def filter_(tarinfo):
            for ex in EXCLUDE_FROM_FULL:
                if ex in tarinfo.name.split("/"):
                    return None
            if tarinfo.name.startswith("output/~$"):
                return None
            return tarinfo

        with tarfile.open(tar_path, "w:gz", compresslevel=6) as tar:
            tar.add(out_src, arcname="output", filter=filter_)
        manifest["files"]["output.tar.gz"] = {
            "size": tar_path.stat().st_size,
            "sha256": sha256_file(tar_path),
            "kind": "tarball",
        }
        print(f"  ✓ output.tar.gz ({tar_path.stat().st_size:,} bytes)")

    # Output tree (just file list of output/)
    tree = []
    for p in sorted((ROOT / "output").rglob("*")):
        if p.is_file() and not p.name.startswith("~$") and p.name != ".DS_Store":
            try:
                rel = str(p.relative_to(ROOT))
                tree.append({"path": rel, "size": p.stat().st_size, "mtime": p.stat().st_mtime})
            except Exception:
                pass
    manifest["output_tree"] = tree

    with open(snap_dir / "manifest.json", "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)

    total_size = sum(p.stat().st_size for p in snap_dir.rglob("*") if p.is_file())
    print(f"\n✓ Snapshot {snap_id} created ({total_size:,} bytes total, {len(tree)} files tracked)")


def cmd_list(args):
    if not SNAPS_DIR.exists():
        print("No snapshots yet. Run: python3 scripts/version.py snap")
        return
    snaps = sorted(p for p in SNAPS_DIR.iterdir() if p.is_dir())
    if not snaps:
        print("No snapshots yet.")
        return
    print(f"{'ID':<40} {'KIND':<6} {'SIZE':>10}  TAG")
    print("-" * 90)
    for s in snaps:
        m_path = s / "manifest.json"
        if not m_path.exists():
            continue
        try:
            m = json.loads(m_path.read_text(encoding="utf-8"))
        except Exception:
            continue
        size = sum(p.stat().st_size for p in s.rglob("*") if p.is_file())
        tag = m.get("tag", "")
        print(f"{m['id']:<40} {m.get('kind','?'):<6} {size:>10,}  {tag}")


def find_snap(snap_id: str) -> Path:
    """Find snapshot by exact ID or prefix."""
    if not SNAPS_DIR.exists():
        return None
    matches = [p for p in SNAPS_DIR.iterdir() if p.is_dir() and (p.name == snap_id or p.name.startswith(snap_id))]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        print(f"Ambiguous prefix '{snap_id}'. Matches:")
        for m in matches:
            print(f"  {m.name}")
        return None
    return None


def cmd_show(args):
    snap = find_snap(args.id)
    if not snap:
        print(f"Snapshot '{args.id}' not found.")
        sys.exit(1)
    m = json.loads((snap / "manifest.json").read_text(encoding="utf-8"))
    print(f"ID:        {m['id']}")
    print(f"Timestamp: {m['timestamp']}")
    print(f"Tag:       {m.get('tag','(none)')}")
    print(f"Kind:      {m.get('kind','?')}")
    print("\nTracked files in snapshot:")
    for path, info in m["files"].items():
        print(f"  {path:<55} {info['size']:>10,} bytes  sha:{info['sha256'][:12]}")
    print(f"\noutput/ tree at snapshot time: {len(m.get('output_tree', []))} files")


def cmd_diff(args):
    s1 = find_snap(args.id1)
    s2 = find_snap(args.id2)
    if not s1 or not s2:
        sys.exit(1)
    m1 = json.loads((s1 / "manifest.json").read_text(encoding="utf-8"))
    m2 = json.loads((s2 / "manifest.json").read_text(encoding="utf-8"))

    print("\n=== Tracked files ===")
    f1 = m1.get("files", {})
    f2 = m2.get("files", {})
    all_keys = set(f1.keys()) | set(f2.keys())
    for k in sorted(all_keys):
        a, b = f1.get(k), f2.get(k)
        if not a:
            print(f"  + {k} (added)")
        elif not b:
            print(f"  - {k} (removed)")
        elif a.get("sha256") != b.get("sha256"):
            print(f"  ~ {k} (changed: {a['size']:,} → {b['size']:,} bytes)")
        else:
            print(f"  = {k} (unchanged)")

    print("\n=== output/ tree diff ===")
    t1 = {x["path"]: x for x in m1.get("output_tree", [])}
    t2 = {x["path"]: x for x in m2.get("output_tree", [])}
    added = [k for k in t2 if k not in t1]
    removed = [k for k in t1 if k not in t2]
    changed = [k for k in t1 if k in t2 and t1[k]["size"] != t2[k]["size"]]
    for k in sorted(added)[:50]:
        print(f"  + {k}")
    for k in sorted(removed)[:50]:
        print(f"  - {k}")
    for k in sorted(changed)[:50]:
        a, b = t1[k]["size"], t2[k]["size"]
        print(f"  ~ {k} ({a:,} → {b:,})")
    print(f"\nTotal: +{len(added)}  -{len(removed)}  ~{len(changed)}  ={len(t1) - len(removed) - len(changed)}")


def cmd_restore(args, full: bool = False):
    snap = find_snap(args.id)
    if not snap:
        sys.exit(1)
    m = json.loads((snap / "manifest.json").read_text(encoding="utf-8"))
    print(f"Restoring snapshot: {m['id']}  tag={m.get('tag','')!r}")
    if not args.yes:
        ans = input("This will OVERWRITE current files. Continue? [y/N] ").strip().lower()
        if ans != "y":
            print("Aborted.")
            return

    # Restore tracked files
    for src_rel, info in m["files"].items():
        if info.get("kind") == "tarball":
            continue
        snap_file = snap / info.get("snapped_as", Path(src_rel).name)
        target = ROOT / src_rel
        if snap_file.exists():
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(snap_file, target)
            print(f"  ✓ restored {src_rel}")
        else:
            print(f"  ⚠ skip (missing in snapshot): {snap_file}")

    if full:
        tar_path = snap / "output.tar.gz"
        if not tar_path.exists():
            print("  ⚠ This is a quick snapshot — no output.tar.gz to restore. Use --quick or take a full snapshot.")
            return
        print("  • extracting output.tar.gz → output/ (overwrites)...")
        # Extract with safe path filter
        with tarfile.open(tar_path, "r:gz") as tar:
            tar.extractall(ROOT)
        print("  ✓ restored output/ from tarball")


def cmd_prune(args):
    snaps = sorted(p for p in SNAPS_DIR.iterdir() if p.is_dir()) if SNAPS_DIR.exists() else []
    keep = args.keep
    if len(snaps) <= keep:
        print(f"Have {len(snaps)} snapshots — nothing to prune (keep={keep})")
        return
    to_remove = snaps[:-keep]
    print(f"Will remove {len(to_remove)} snapshots (keeping last {keep}):")
    for s in to_remove:
        print(f"  - {s.name}")
    if not args.yes:
        ans = input("Continue? [y/N] ").strip().lower()
        if ans != "y":
            print("Aborted.")
            return
    for s in to_remove:
        shutil.rmtree(s)
    print(f"✓ removed {len(to_remove)} snapshots")


def cmd_auto_snap(args):
    """Snap only if xlsx changed since last snapshot."""
    if not SNAPS_DIR.exists():
        return cmd_snap(argparse.Namespace(tag="auto"))
    snaps = sorted(p for p in SNAPS_DIR.iterdir() if p.is_dir())
    if not snaps:
        return cmd_snap(argparse.Namespace(tag="auto"))

    last = snaps[-1]
    m = json.loads((last / "manifest.json").read_text(encoding="utf-8"))
    xlsx_rel = "output/Comply spec Smart Plant 1.xlsx"
    last_sha = m.get("files", {}).get(xlsx_rel, {}).get("sha256")
    cur_sha = sha256_file(ROOT / xlsx_rel) if (ROOT / xlsx_rel).exists() else None

    if last_sha == cur_sha:
        print(f"xlsx unchanged since last snapshot ({last.name}). Skipping.")
        return
    print("xlsx changed — taking auto snapshot")
    cmd_snap(argparse.Namespace(tag="auto"))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("snap", help="Quick snapshot (xlsx + SKILL.md)")
    p.add_argument("tag", nargs="?", default="")

    p = sub.add_parser("snap-full", help="Full snapshot (xlsx + SKILL.md + output/ tar.gz)")
    p.add_argument("tag", nargs="?", default="")

    p = sub.add_parser("list", help="List all snapshots")

    p = sub.add_parser("show", help="Show snapshot details")
    p.add_argument("id")

    p = sub.add_parser("diff", help="Diff 2 snapshots")
    p.add_argument("id1")
    p.add_argument("id2")

    p = sub.add_parser("restore", help="Restore quick (xlsx + SKILL.md)")
    p.add_argument("id")
    p.add_argument("-y", "--yes", action="store_true", help="Skip confirmation")

    p = sub.add_parser("restore-full", help="Restore full (output/ from tarball)")
    p.add_argument("id")
    p.add_argument("-y", "--yes", action="store_true")

    p = sub.add_parser("prune", help="Keep only last N snapshots")
    p.add_argument("--keep", type=int, default=10)
    p.add_argument("-y", "--yes", action="store_true")

    p = sub.add_parser("auto-snap", help="Snap only if xlsx changed since last")

    args = ap.parse_args()
    if args.cmd == "snap":
        cmd_snap(args, full=False)
    elif args.cmd == "snap-full":
        cmd_snap(args, full=True)
    elif args.cmd == "list":
        cmd_list(args)
    elif args.cmd == "show":
        cmd_show(args)
    elif args.cmd == "diff":
        cmd_diff(args)
    elif args.cmd == "restore":
        cmd_restore(args, full=False)
    elif args.cmd == "restore-full":
        cmd_restore(args, full=True)
    elif args.cmd == "prune":
        cmd_prune(args)
    elif args.cmd == "auto-snap":
        cmd_auto_snap(args)


if __name__ == "__main__":
    main()
