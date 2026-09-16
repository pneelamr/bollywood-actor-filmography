"""Apply a batch of hand-written rows to research/<actor> JSON.

usage: merge_batch.py RESEARCH --sources S.json --films A.json [B.json ...] --romances R1.json [R2.json ...] [--dry-run]

Films files map film_id -> fields. Special keys: add_sources (merged into source_ids),
notes_add (appended to research_notes), overwrite (list of keys allowed to replace a
non-empty existing value). Romances and sources rows are upserted by ID. Source rows get
film_ids / pairing_ids back-links for every reference. Idempotent.
"""
import argparse, json
from pathlib import Path


def load(p):
    return json.loads(Path(p).read_text())


def split_ids(v):
    if not v:
        return []
    return [x.strip() for x in (v if isinstance(v, list) else str(v).split(";")) if x.strip()]


def dump(path, rows, style_from):
    raw = style_from.read_text()  # empty files ("[]") carry no style; follow films.json
    indent = 2 if raw.startswith("[\n  ") else (1 if raw.startswith("[\n {") else None)
    path.write_text(json.dumps(rows, ensure_ascii=False, indent=indent) + "\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("research")
    ap.add_argument("--sources", nargs="*", default=[])
    ap.add_argument("--films", nargs="*", default=[])
    ap.add_argument("--romances", nargs="*", default=[])
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    R = Path(a.research)
    films, roms, srcs = load(R / "films.json"), load(R / "romances.json"), load(R / "sources.json")
    F = {f["film_id"]: f for f in films}
    S = {s["source_id"]: s for s in srcs}
    problems = []

    for p in a.sources:
        for row in load(p):
            if row["source_id"] in S:
                S[row["source_id"]].update({k: v for k, v in row.items() if k not in ("film_ids", "pairing_ids")})
            else:
                srcs.append(row)
                S[row["source_id"]] = row

    for p in a.films:
        for fid, patch in load(p).items():
            f = F.get(fid)
            if not f:
                problems.append(f"{fid}: not in films.json")
                continue
            allow = set(patch.get("overwrite", []))
            for k, v in patch.items():
                if k in ("overwrite", "add_sources", "notes_add"):
                    continue
                old = f.get(k)
                if old not in (None, "", []) and old != v and k not in allow:
                    problems.append(f"{fid}.{k}: kept existing {old!r}, patch had {v!r}")
                    continue
                f[k] = v
            ids = split_ids(f.get("source_ids"))
            ids += [x for x in split_ids(patch.get("add_sources")) if x not in ids]
            f["source_ids"] = "; ".join(ids)
            note = patch.get("notes_add")
            if note and note not in (f.get("research_notes") or ""):
                f["research_notes"] = ((f.get("research_notes") or "") + " " + note).strip()
            if any(f.get(k) is not None for k in ("domestic_nett", "domestic_gross", "overseas_gross", "worldwide_gross", "budget")) \
                    and f.get("currency") != "INR crore":
                problems.append(f"{fid}: money fields with currency {f.get('currency')!r}")

    by_id = {r["pairing_id"]: i for i, r in enumerate(roms)}
    for p in a.romances:
        for row in load(p):
            if row["film_id"] not in F:
                problems.append(f"{row['pairing_id']}: unknown film {row['film_id']}")
            film_src = set(split_ids(F.get(row["film_id"], {}).get("source_ids")))
            if not film_src & set(split_ids(row.get("source_ids"))):
                problems.append(f"{row['pairing_id']}: shares no source with its film")
            if row["pairing_id"] in by_id:
                roms[by_id[row["pairing_id"]]] = row
            else:
                by_id[row["pairing_id"]] = len(roms)
                roms.append(row)
    roms.sort(key=lambda r: r["pairing_id"])

    for owner, key, rows, idkey in ((F, "film_ids", films, "film_id"), (None, "pairing_ids", roms, "pairing_id")):
        for r in rows:
            for sid in split_ids(r.get("source_ids")):
                s = S.get(sid)
                if not s:
                    problems.append(f"{r[idkey]}: unknown source {sid}")
                    continue
                links = split_ids(s.get(key))
                if r[idkey] not in links:
                    s[key] = links + [r[idkey]]
    srcs.sort(key=lambda s: s["source_id"])

    print("\n".join(problems) or "no problems")
    print(f"films {len(films)}, romances {len(roms)}, sources {len(srcs)}")
    if not a.dry_run:
        style = R / "films.json"
        dump(R / "romances.json", roms, style)
        dump(R / "sources.json", srcs, style)
        dump(R / "films.json", films, style)


if __name__ == "__main__":
    main()
