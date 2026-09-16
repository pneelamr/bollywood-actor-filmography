"""Upsert collaborator rows into research/<actor>/collaborators.json.

usage: p3_merge_collab.py RESEARCH ROSTER.json FILE1.json [FILE2.json ...] [--dry-run]

Each input file is a JSON array of collaborator rows keyed by person_id. For every
row this script:
  * checks person_id / name / film_ids against the frozen roster and reports drift,
  * checks every source_id resolves in sources.json,
  * checks every film_id resolves in films.json,
  * leaves collaboration_count blank (the build script derives it),
  * fills collaboration_film_ids from the roster when the row omits it.
Rows are upserted by person_id and the file is written sorted by person_id.
Idempotent, like merge_batch.py.

Note: sources.json is NOT modified. The Sources sheet has no person_ids column --
it carries only film_ids and pairing_ids -- so there is no reverse back-link to
write. The collaborator -> source direction is validated here and by the build.
"""
import argparse, json
from pathlib import Path


def split_ids(v):
    if not v:
        return []
    return [x.strip() for x in (v if isinstance(v, list) else str(v).split(";")) if x.strip()]


def dump(path, rows, style_from):
    raw = style_from.read_text()
    indent = 2 if raw.startswith("[\n  ") else (1 if raw.startswith("[\n {") else 1)
    path.write_text(json.dumps(rows, ensure_ascii=False, indent=indent) + "\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("research")
    ap.add_argument("roster")
    ap.add_argument("files", nargs="+")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    R = Path(a.research)
    collab = json.loads((R / "collaborators.json").read_text())
    srcs = json.loads((R / "sources.json").read_text())
    films = json.loads((R / "films.json").read_text())
    roster = {r["person_id"]: r for r in json.loads(Path(a.roster).read_text())}
    S = {s["source_id"]: s for s in srcs}
    FID = {f["film_id"] for f in films}
    by_id = {c["person_id"]: i for i, c in enumerate(collab)}
    problems = []

    for p in a.files:
        for row in json.loads(Path(p).read_text()):
            pid = row.get("person_id")
            r = roster.get(pid)
            if not r:
                problems.append(f"{pid}: not in the frozen roster")
                continue
            if row.get("name") != r["name"]:
                problems.append(f"{pid}: name {row.get('name')!r} != roster {r['name']!r}")
            # collaboration_count is derived by the build script
            if row.pop("collaboration_count", None) is not None:
                problems.append(f"{pid}: collaboration_count is derived; dropped")
            ids = split_ids(row.get("collaboration_film_ids")) or list(r["film_ids"])
            bad = [x for x in ids if x not in FID]
            if bad:
                problems.append(f"{pid}: unknown film ids {bad}")
            if sorted(ids) != sorted(r["film_ids"]):
                problems.append(f"{pid}: film ids differ from roster ({len(ids)} vs {len(r['film_ids'])})")
            row["collaboration_film_ids"] = "; ".join(ids)
            for sid in split_ids(row.get("source_ids")):
                if sid not in S:
                    problems.append(f"{pid}: unknown source {sid}")
            if pid in by_id:
                collab[by_id[pid]] = row
            else:
                by_id[pid] = len(collab)
                collab.append(row)

    collab.sort(key=lambda c: c["person_id"])

    print("\n".join(problems) or "no problems")
    print(f"collaborators {len(collab)}, sources {len(srcs)} (sources unchanged)")
    if not a.dry_run:
        dump(R / "collaborators.json", collab, R / "films.json")


if __name__ == "__main__":
    main()
