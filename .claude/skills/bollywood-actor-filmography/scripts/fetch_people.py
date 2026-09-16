"""Fetch Wikipedia raw wikitext for collaborators.

usage: p3_fetch_people.py ROSTER.json OUTDIR START END [--titles extra_titles.json]

Reads roster entries PERSON-<START>..PERSON-<END>, fetches each name's article as
raw wikitext into OUTDIR/<person_id>.txt, follows #REDIRECT once, and reports which
names resolved to an article and which did not. extra_titles.json maps person_id ->
explicit Wikipedia title, for names whose article title differs from the credit.
"""
import sys, json, re, subprocess, time, urllib.parse
from pathlib import Path

UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"


def raw(title):
    url = "https://en.wikipedia.org/w/index.php?title=%s&action=raw" % urllib.parse.quote(title.replace(" ", "_"), safe="")
    r = subprocess.run(["curl", "-sL", "--max-time", "25", "-A", UA, url], capture_output=True, text=True)
    return r.stdout or ""


def fetch(title):
    """Return (text, resolved_title). Follows a single #REDIRECT hop."""
    t = raw(title)
    m = re.match(r"(?i)\s*#redirect\s*\[\[([^\]|#]+)", t or "")
    if m:
        tgt = m.group(1).strip()
        return raw(tgt), tgt
    return t, title


def main():
    roster = json.load(open(sys.argv[1]))
    out = Path(sys.argv[2]); out.mkdir(parents=True, exist_ok=True)
    lo, hi = int(sys.argv[3]), int(sys.argv[4])
    extra = {}
    if "--titles" in sys.argv:
        extra = json.load(open(sys.argv[sys.argv.index("--titles") + 1]))
    rows = [r for r in roster if lo <= int(r["person_id"].split("-")[1]) <= hi]
    found, missing = [], []
    for r in rows:
        pid = r["person_id"]
        f = out / ("%s.txt" % pid)
        if f.exists() and f.stat().st_size > 200:
            found.append((pid, r["name"], f.stat().st_size, "cached"))
            continue
        title = extra.get(pid, r["name"])
        text, resolved = fetch(title)
        # an article page starts with wikitext; a missing one returns empty or a tiny stub
        if len(text) > 200 and not text.lstrip().lower().startswith("#redirect"):
            f.write_text(text, encoding="utf-8")
            found.append((pid, r["name"], len(text), resolved if resolved != r["name"] else ""))
        else:
            missing.append((pid, r["name"], len(text)))
        time.sleep(0.4)
    print("FOUND %d / MISSING %d" % (len(found), len(missing)))
    for pid, name, size, note in found:
        print("  ok      %-11s %-34s %7d %s" % (pid, name[:34], size, note))
    print()
    for pid, name, size in missing:
        print("  MISSING %-11s %-34s (%d bytes)" % (pid, name[:34], size))


if __name__ == "__main__":
    main()
