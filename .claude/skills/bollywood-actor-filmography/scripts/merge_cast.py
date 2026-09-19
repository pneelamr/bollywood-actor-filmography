"""Upsert ensemble-cast rows into research/<actor>/ensemble_cast.json.

usage: p4_merge_cast.py RESEARCH FILE1.json [FILE2.json ...] [--dry-run]

Each input file is a JSON array of cast rows keyed by cast_id. For every row this
script:
  * checks cast_id is unique and well formed (CAST-###),
  * checks film_id resolves in films.json,
  * checks every source_id resolves in sources.json,
  * checks cast_function and importance against the controlled vocabularies,
  * leaves film / release_year blank (the build derives them from film_id),
  * warns when the actor is the selected actor himself.
Rows are upserted by cast_id and the file is written sorted by cast_id.
Idempotent, like merge_batch.py and p3_merge_collab.py.

Note: sources.json is NOT modified. Cast rows carry source_ids; the reverse
back-link (film_ids on the source row) is already written by the films phase.
"""
import argparse, json, re
from pathlib import Path

FUNCTIONS = {"Co-lead", "Supporting", "Antagonist", "Mentor", "Family", "Friend",
             "Comic support", "Romantic rival", "Authority figure", "Cameo", "Self appearance"}
IMPORTANCE = {"Major", "Medium", "Minor"}
DERIVED = ("film", "release_year")


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
    ap.add_argument("files", nargs="+")
    ap.add_argument("--actor", default="Amitabh Bachchan")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    R = Path(a.research)
    cast = json.loads((R / "ensemble_cast.json").read_text())
    srcs = json.loads((R / "sources.json").read_text())
    films = json.loads((R / "films.json").read_text())
    S = {s["source_id"] for s in srcs}
    FID = {f["film_id"] for f in films}
    by_id = {c["cast_id"]: i for i, c in enumerate(cast)}
    seen = set()
    problems = []

    for p in a.files:
        for row in json.loads(Path(p).read_text()):
            cid = row.get("cast_id")
            if not cid or not re.fullmatch(r"CAST-\d{3,}", cid):
                problems.append(f"{cid!r}: malformed cast_id")
                continue
            if cid in seen:
                problems.append(f"{cid}: duplicated within this batch")
            seen.add(cid)
            if row.get("film_id") not in FID:
                problems.append(f"{cid}: unknown film {row.get('film_id')}")
            if row.get("cast_function") not in FUNCTIONS:
                problems.append(f"{cid}: bad cast_function {row.get('cast_function')!r}")
            if row.get("importance") not in IMPORTANCE:
                problems.append(f"{cid}: bad importance {row.get('importance')!r}")
            if not row.get("actor"):
                problems.append(f"{cid}: no actor")
            if (row.get("actor") or "").strip().lower() == a.actor.lower():
                problems.append(f"{cid}: actor is {a.actor} himself")
            for k in DERIVED:
                if row.pop(k, None) is not None:
                    problems.append(f"{cid}: {k} is derived by the build; dropped")
            for sid in split_ids(row.get("source_ids")):
                if sid not in S:
                    problems.append(f"{cid}: unknown source {sid}")
            if cid in by_id:
                cast[by_id[cid]] = row
            else:
                by_id[cid] = len(cast)
                cast.append(row)

    cast.sort(key=lambda c: int(c["cast_id"].split("-")[1]))
    # one film's rows must not name the same actor twice
    per = {}
    for c in cast:
        per.setdefault((c["film_id"], (c.get("actor") or "").lower()), []).append(c["cast_id"])
    for (fid, actor), ids in per.items():
        if len(ids) > 1:
            problems.append(f"{fid}: {actor} appears in {ids}")

    print("\n".join(problems) or "no problems")
    print(f"ensemble_cast {len(cast)} rows over {len({c['film_id'] for c in cast})} films")
    if not a.dry_run:
        dump(R / "ensemble_cast.json", cast, R / "films.json")


if __name__ == "__main__":
    main()
