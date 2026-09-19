#!/usr/bin/env python3
"""Fill mechanical filmography fields from Wikipedia raw wikitext.

Companion to build_workbook.py; standard library only. WORK is a scratch
directory (downloaded pages go in WORK/wt). RESEARCH is research/<actor-slug>.

  extract_wiki.py table     WORK --page "Amitabh Bachchan filmography" [--section "Acting credits"]
  extract_wiki.py fetch     WORK
  extract_wiki.py films     WORK RESEARCH --actor "Name" [--config scope.json] [--overwrite]
  extract_wiki.py plot      WORK RESEARCH --ids FILM-001-FILM-030 [--chars 2500] [--terms "Bachchan;Amitabh"]
  extract_wiki.py cast      WORK RESEARCH --ids FILM-001,FILM-004 [--max 25]
  extract_wiki.py reviews   WORK RESEARCH --ids ... --terms "Bachchan;Amitabh;chemistry"
  extract_wiki.py boxoffice WORK RESEARCH --ids ...

`table` parses the actor's filmography table into WORK/table.json and prints one
line per row. `fetch` downloads every linked film article, following redirects.
`films` writes the mechanical Films fields and one Wikipedia source row per film
into RESEARCH, keeping any value already present unless --overwrite is given.
The later commands print tight extracts for hand-written fields; they re-download
pages named by the film-article source URLs when WORK is a fresh directory.

--config JSON: {"exclude": {"Title|Year": "reason"},
                "add": [{"year": 2025, "title": "...", "target": "Wiki title", "role": "...", "notes": "..."}],
                "overrides": {"Title|Year": {"field": value}}}
"""
from __future__ import annotations

import argparse
import concurrent.futures as cf
import datetime as dt
import json
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

UA = "FilmographyResearch/1.0 (personal research script)"
RAW_URL = "https://en.wikipedia.org/w/index.php?title={}&action=raw"
WIKI_URL = "https://en.wikipedia.org/wiki/{}"

# ---------------------------------------------------------------- download


def page_file(work: Path, title: str) -> Path:
    return work / "wt" / (re.sub(r"[^\w.-]+", "_", title)[:150] + ".txt")


def download(title: str) -> str | None:
    url = RAW_URL.format(urllib.parse.quote(title.replace(" ", "_"), safe=""))
    for attempt in range(4):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=30) as r:
                return r.read().decode("utf-8")
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None
            time.sleep(2 * (attempt + 1))
        except (urllib.error.URLError, TimeoutError):
            time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"download failed: {title}")


def ensure_page(work: Path, title: str) -> tuple[str | None, str | None]:
    """Return (resolved title, wikitext), downloading and following redirects."""
    for _ in range(3):
        f = page_file(work, title)
        if f.exists():
            text = f.read_text(encoding="utf-8")
        else:
            text = download(title)
            if text is None:
                return None, None
            f.parent.mkdir(parents=True, exist_ok=True)
            f.write_text(text, encoding="utf-8")
        m = re.match(r"\s*#REDIRECT\s*\[\[([^\]|#]+)", text, re.I)
        if not m:
            return title, text
        title = m.group(1).strip()
    return None, None


# ---------------------------------------------------------------- wikitext cleaning

LIST_TPL = {"plainlist", "plain list", "ubl", "unbulleted list", "flatlist", "flat list", "hlist",
            "bulleted list", "collapsible list", "unordered list", "ordered list", "br separated entries",
            "pagelist", "bulleted", "nowrap list", "indented plainlist"}
DROP_TPL = {"efn", "sfn", "sfnp", "refn", "cn", "citation needed", "better source needed", "clarify", "when",
            "dubious", "r", "rp", "increase", "decrease", "steady", "verify", "vague", "update", "by whom",
            "failed verification", "additional citation needed", "who", "which", "specify", "notetag", "ref",
            "nb", "use dmy dates", "use indian english", "hair space", "shy", "anchor", "toc limit", "clear",
            "main", "see also", "further", "portal", "div col", "div col end", "col-begin", "col-end",
            "col-break", "reflist", "efn-la", "efn-lr", "ill-wd", "self-published inline", "full citation needed",
            "page needed", "year needed", "unreliable source?", "pb", "multiple image", "listen", "quote box"}
FIRST_TPL = {"abbr", "tooltip", "nihongo", "lang-hi", "ill", "illm", "interlanguage link", "pending film",
             "small", "big", "nowrap", "nobr", "noitalic", "center", "transliteration", "keypress", "em",
             "strong", "sc", "smallcaps", "resize", "vanchor", "anchored", "linktext", "blockquote", "quote",
             "script", "tq", "text", "lang-en"}
LAST_TPL = {"lang", "transl", "sort", "ipa", "sortname"}
UNITS = {"c": "crore", "cr": "crore", "l": "lakh", "lc": "lakh crore", "m": "million", "b": "billion",
         "k": "thousand", "t": "thousand"}


def split_top(text: str, sep: str = "|") -> list[str]:
    """Split on sep outside {{ }} and [[ ]]."""
    out, buf, depth, i = [], [], 0, 0
    while i < len(text):
        two = text[i:i + 2]
        if two in ("{{", "[["):
            depth += 1; buf.append(two); i += 2; continue
        if two in ("}}", "]]") and depth:
            depth -= 1; buf.append(two); i += 2; continue
        if depth == 0 and text.startswith(sep, i):
            out.append("".join(buf)); buf = []; i += len(sep); continue
        buf.append(text[i]); i += 1
    out.append("".join(buf))
    return out


def matching_end(text: str, start: int, open_: str, close: str) -> int:
    depth, i = 0, start
    while i < len(text):
        if text.startswith(open_, i):
            depth += 1; i += 2; continue
        if text.startswith(close, i):
            depth -= 1; i += 2
            if depth == 0:
                return i
            continue
        i += 1
    return len(text)


def _date_lines(pos: list[str]) -> str:
    groups, cur = [], []
    for a in (p.strip() for p in pos):
        if re.fullmatch(r"\d{1,4}", a):
            if len(cur) == 3:
                groups.append((cur, None)); cur = []
            cur.append(int(a))
        elif a and cur:
            groups.append((cur, a)); cur = []
    if cur:
        groups.append((cur, None))
    lines = []
    for nums, loc in groups:
        s = f"{nums[0]:04d}" + "".join(f"-{n:02d}" for n in nums[1:])
        lines.append(s + (f" ({loc})" if loc else ""))
    return "\n".join(lines)


def _template(inner: str) -> str:
    parts = inner.split("|")
    name = parts[0].strip().lower().replace("_", " ")
    args = parts[1:]
    pos = [a for a in args if not re.match(r"^\s*[\w -]+\s*=", a)]
    kw = {}
    for a in args:
        m = re.match(r"^\s*([\w -]+?)\s*=(.*)$", a, re.S)
        if m:
            kw[m.group(1).lower()] = m.group(2).strip()
    if name.startswith("cite") or name in DROP_TPL or name.startswith("infobox"):
        return ""
    if name in LIST_TPL:
        return "\n" + "\n".join(p.strip() for p in pos if p.strip()) + "\n"
    if name.startswith("film date") or name in {"start date", "release date", "dts", "start date and age"}:
        return _date_lines(pos)
    if name in {"inr", "₹", "indian rupee", "rs", "rs."}:
        return "₹" + (pos[0].strip() if pos else "")
    if name in {"inrconvert", "inr convert"}:
        nums = [p.strip() for p in pos if re.fullmatch(r"[\d.,]+", p.strip())]
        unit = next((UNITS.get(p.strip().lower(), "") for p in pos if p.strip().lower() in UNITS), "")
        return "₹" + "–".join(nums) + (f" {unit}" if unit else "")
    if name in {"us$", "usd", "us dollar", "$"}:
        return "US$" + (pos[0].strip() if pos else "")
    if name in {"gbp", "£"}:
        return "£" + (pos[0].strip() if pos else "")
    if name in {"eur", "€"}:
        return "€" + (pos[0].strip() if pos else "")
    if name == "runtime":
        n = [int(p) for p in pos if p.strip().isdigit()]
        return f"{n[0] * 60 + n[1] if len(n) > 1 else n[0]} minutes" if n else ""
    if name == "based on":
        return "Based on " + (pos[0].strip() if pos else "") + (" by " + ", ".join(p.strip() for p in pos[1:]) if len(pos) > 1 else "")
    if name in {"snd", "spaced ndash", "spaced en dash"}:
        return " – "
    if name in {"ndash", "en dash", "dash"}:
        return "–"
    if name in {"mdash", "em dash"}:
        return "—"
    if name in {"nbsp", "sp"}:
        return " "
    if name in {"est.", "est"}:
        return "est. " + (pos[0].strip() if pos else "")
    if name in {"circa", "c."}:
        return "c. " + (pos[0].strip() if pos else "")
    if name == "aka":
        return "a.k.a."
    if name in FIRST_TPL:
        return pos[0] if pos else kw.get("text", kw.get("1", ""))
    if name in LAST_TPL:
        return pos[-1] if pos else ""
    return " ".join(p.strip() for p in pos)


def _drop_file_links(text: str) -> str:
    while True:
        m = re.search(r"\[\[\s*(?:File|Image|Category)\s*:", text, re.I)
        if not m:
            return text
        text = text[:m.start()] + text[matching_end(text, m.start(), "[[", "]]"):]


def clean(text: str | None) -> str:
    if not text:
        return ""
    t = re.sub(r"<!--.*?-->", "", text, flags=re.S)
    t = re.sub(r"<ref[^>/]*/>", "", t)
    t = re.sub(r"<ref[^>]*>.*?</ref>", "", t, flags=re.S | re.I)
    t = re.sub(r"<ref\b[^<>]*(?=<|$)", "", t, flags=re.M)  # malformed, unterminated ref tags
    t = re.sub(r"<br\s*/?\s*>", "\n", t, flags=re.I)
    t = _drop_file_links(t)
    t = re.sub(r"\[\[(?:[^\[\]|]*)\|([^\[\]]*)\]\]", r"\1", t)
    t = re.sub(r"\[\[:?([^\[\]|#]*)(?:#[^\[\]]*)?\]\]", r"\1", t)
    for _ in range(20):
        new = re.sub(r"\{\{([^{}]*)\}\}", lambda m: _template(m.group(1)), t)
        if new == t:
            break
        t = new
    t = re.sub(r"\[https?://\S+\s+([^\]]+)\]", r"\1", t)
    t = re.sub(r"\[https?://\S+\]", "", t)
    t = re.sub(r"</?[a-zA-Z][^>]*>", "", t)
    t = t.replace("'''", "").replace("''", "").replace("&nbsp;", " ").replace("&ndash;", "–").replace("&amp;", "&")
    t = re.sub(r"[ \t]+", " ", t)
    return re.sub(r"\n\s*\n+", "\n", t).strip()


def items(text: str | None, commas: bool = True) -> list[str]:
    out = []
    for line in clean(text).split("\n"):
        line = re.sub(r"^[*#:;\s]+", "", line).strip(" ,;")
        if not line:
            continue
        parts = re.split(r",\s*|\s+and\s+(?=[A-Z])", line) if commas else [line]
        out += [p.strip(" ,;") for p in parts if p.strip(" ,;")]
    return out


def unique(seq) -> list[str]:
    seen, out = set(), []
    for s in seq:
        if s and s.lower() not in seen:
            seen.add(s.lower()); out.append(s)
    return out


def sentences(text: str) -> list[str]:
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+(?=[A-Z\"'‘“(₹])", text.replace("\n", " ")) if s.strip()]


# ---------------------------------------------------------------- page structure


def template_params(text: str, name_regex: str) -> dict[str, str] | None:
    m = re.search(r"\{\{\s*" + name_regex + r"\s*[|\n}]", text, re.I)
    if not m:
        return None
    body = text[m.start() + 2:matching_end(text, m.start(), "{{", "}}") - 2]
    params = {}
    for part in split_top(body)[1:]:
        if "=" in part:
            k, v = part.split("=", 1)
            params[k.strip().lower().replace(" ", "_")] = v.strip()
    return params


def all_template_params(text: str, name_regex: str) -> list[dict[str, str]]:
    out, pos = [], 0
    while True:
        m = re.search(r"\{\{\s*" + name_regex + r"\s*[|\n}]", text[pos:], re.I)
        if not m:
            return out
        start = pos + m.start()
        out.append(template_params(text[start:], name_regex) or {})
        pos = matching_end(text, start, "{{", "}}")


def section(text: str, name_regex: str, exclude: str | None = None) -> str:
    """Concatenate every section whose heading matches, up to the next heading of equal or higher level."""
    heads = list(re.finditer(r"^(={2,6})\s*(.*?)\s*\1\s*$", text, re.M))
    chunks = []
    for i, h in enumerate(heads):
        title = clean(h.group(2))
        if not re.search(name_regex, title, re.I) or (exclude and re.search(exclude, title, re.I)):
            continue
        level = len(h.group(1))
        end = next((g.start() for g in heads[i + 1:] if len(g.group(1)) <= level), len(text))
        chunks.append(text[h.end():end])
    return "\n".join(chunks)


def lead(text: str) -> str:
    m = re.search(r"^==", text, re.M)
    return text[:m.start()] if m else text


# ---------------------------------------------------------------- filmography table


def parse_wikitable(tbl: str) -> tuple[list[str], list[list[str]]]:
    rows, cur = [], None
    for line in tbl.split("\n")[1:]:
        s = line.strip()
        if s.startswith("|}"):
            break
        if s.startswith("|-"):
            if cur:
                rows.append(cur)
            cur = []
            continue
        if s.startswith("|+"):
            continue
        if s[:1] in ("!", "|"):
            cur = [] if cur is None else cur
            body = s[1:]
            cells = split_top(body, "!!" if s[0] == "!" else "||")
            if s[0] == "!" and len(cells) == 1:
                cells = split_top(body, "||")
            cur += [[s[0], c] for c in cells]
        elif cur:
            cur[-1][1] += "\n" + line
    if cur:
        rows.append(cur)
    header_row = next(r for r in rows if all(k == "!" for k, _ in r))
    headers = [clean(split_attr(c)[1]).lower() for _, c in header_row]
    data = rows[rows.index(header_row) + 1:]
    grid, carry = [], [[0, None] for _ in headers]
    for r in data:
        out, cells, ci = [None] * len(headers), iter(r), 0
        while ci < len(headers):
            if carry[ci][0] > 0:
                out[ci] = carry[ci][1]; carry[ci][0] -= 1; ci += 1; continue
            cell = next(cells, None)
            if cell is None:
                break
            attrs, content = split_attr(cell[1])
            rs = re.search(r'rowspan\s*=\s*"?(\d+)', attrs)
            cs = re.search(r'colspan\s*=\s*"?(\d+)', attrs)
            for _ in range(int(cs.group(1)) if cs else 1):
                if ci >= len(headers):
                    break
                out[ci] = content
                if rs and int(rs.group(1)) > 1:
                    carry[ci] = [int(rs.group(1)) - 1, content]
                ci += 1
        grid.append(out)
    return headers, grid


def split_attr(cell: str) -> tuple[str, str]:
    parts = split_top(cell)
    if len(parts) > 1 and re.fullmatch(r"\s*(?:[\w-]+\s*=\s*(?:\"[^\"]*\"|'[^']*'|[^\s|]+)\s*)+", parts[0]):
        return parts[0], "|".join(parts[1:]).strip()
    return "", cell.strip()


def cmd_table(a) -> int:
    work = Path(a.work)
    title, text = ensure_page(work, a.page)
    if text is None:
        print(f"page not found: {a.page}")
        return 1
    sec = section(text, a.section)
    start = sec.find("{|")
    if start < 0:
        print(f"no table under a heading matching {a.section!r}")
        return 1
    headers, grid = parse_wikitable(sec[start:])
    col = lambda *names: next((i for i, h in enumerate(headers) if any(n in h for n in names)), None)
    yc, tc, rc, nc, lc = col("year"), col("title", "film"), col("role"), col("note"), col("language")
    rows = []
    for g in grid:
        raw = g[tc] or ""
        link = re.search(r"\[\[([^\]|#]+)(?:#[^\]|]*)?(?:\|[^\]]*)?\]\]", raw)
        year = re.search(r"\b(1[89]\d\d|20\d\d)\b", clean(g[yc]) if yc is not None else "")
        notes = clean(g[nc]) if nc is not None else ""
        if lc is not None and clean(g[lc]):
            notes = f"{clean(g[lc])} film; {notes}".strip("; ")
        rows.append({"year": int(year.group(1)) if year else None, "title": clean(raw),
                     "target": link.group(1).strip() if link else None,
                     "role": clean(g[rc]) if rc is not None else "", "notes": notes.replace("\n", "; "),
                     "pending": bool(re.search(r"pending film|\bTBA\b", raw + (g[yc] or ""), re.I))})
    (work / "table.json").write_text(json.dumps({"page": title, "rows": rows}, indent=1, ensure_ascii=False))
    for i, r in enumerate(rows, 1):
        flag = " PENDING" if r["pending"] else ("" if r["target"] else " NOLINK")
        print(f"{i:3} {r['year']} | {r['title']} | {r['role'][:40]} | {r['notes'][:70]}{flag}")
    print(f"{len(rows)} rows -> {work / 'table.json'}")
    return 0


def cmd_fetch(a) -> int:
    work = Path(a.work)
    rows = json.loads((work / "table.json").read_text())["rows"]
    rows += load_config(a.config).get("add", [])
    targets = unique(r["target"] for r in rows if r.get("target"))
    pages, failed = {}, []

    def get(t):
        try:
            return ensure_page(work, t)[0]
        except RuntimeError:
            failed.append(t)
            return None

    with cf.ThreadPoolExecutor(3) as pool:
        for t, resolved in zip(targets, pool.map(get, targets)):
            pages[t] = resolved
    (work / "pages.json").write_text(json.dumps(pages, indent=1, ensure_ascii=False))
    missing = [t for t, r in pages.items() if r is None and t not in failed]
    redirected = {t: r for t, r in pages.items() if r and r != t}
    print(f"{len(pages) - len(missing) - len(failed)} pages downloaded; {len(redirected)} via redirect; "
          f"missing: {missing or 'none'}; failed (rerun fetch): {failed or 'none'}")
    return 1 if failed else 0


# ---------------------------------------------------------------- mechanical fields

MONEY = re.compile(r"(₹|Rs\.?|INR|US\$|\$|£|€)\s*([\d,]+(?:\.\d+)?)"
                   r"(?:\s*(?:–|-|to)\s*(?:₹|Rs\.?|US\$|\$)?\s*([\d,]+(?:\.\d+)?))?"
                   r"\s*(crores?|cr\b|lakhs?|lacs?|million|billion|mn\b|bn\b)?", re.I)


def money(text: str) -> tuple[str, float | str] | None:
    """Return (currency unit, value) for the first figure; a range becomes text."""
    m = MONEY.search(text)
    if not m:
        return None
    cur = {"₹": "INR", "rs": "INR", "rs.": "INR", "inr": "INR", "us$": "USD", "$": "USD", "£": "GBP", "€": "EUR"}[m.group(1).lower()]
    unit = (m.group(4) or "").lower()
    lo = float(m.group(2).replace(",", ""))
    hi = float(m.group(3).replace(",", "")) if m.group(3) else None
    if cur == "INR":
        factor = 1 if unit.startswith("cr") else 0.01 if unit.startswith(("lakh", "lac")) else \
            0.1 if unit in ("million", "mn") else 100 if unit in ("billion", "bn") else 1e-7
        unit_name = "INR crore"
    else:
        factor = 1 if unit in ("million", "mn") else 1000 if unit in ("billion", "bn") else 1e-6
        unit_name = f"{cur} million"
    if hi is not None:
        return unit_name, f"{m.group(1)}{m.group(2)}–{m.group(3)}" + (f" {m.group(4)}" if m.group(4) else "")
    v = round(lo * factor, 4)
    return unit_name, int(v) if v == int(v) else v


def parse_music(raw: str | None) -> tuple[list[str], list[str], list[str]]:
    songs, score, guest, lyrics, bucket = [], [], [], [], "songs"
    kind = lambda label: ("lyrics" if "lyric" in label else
                          "both" if re.search(r"song|soundtrack", label) and re.search(r"score|background", label)
                          else "score" if re.search(r"score|background|bgm", label)
                          else "guest" if re.search(r"guest|additional|recreat|original|remix", label) else "songs")
    for it in items(raw, commas=False):
        m = re.match(r"^([A-Za-z ()/&-]{3,40}?)\s*:\s*(.*)$", it)
        if m:
            bucket = kind(m.group(1).lower())
            it = m.group(2).strip()
            if not it:
                continue
        b = bucket
        pm = re.match(r"^(.*?)\s*\(([^)]*(?:song|score|background|bgm|guest|additional)[^)]*)\)\s*$", it, re.I)
        if pm:
            it, b = pm.group(1), kind(pm.group(2).lower())
        names = [n.strip(" ,;") for n in re.split(r",\s*|\s+and\s+(?=[A-Z])", it) if n.strip(" ,;")]
        if b in ("songs", "both"):
            songs += names
        if b in ("score", "both"):
            score += names
        if b == "guest":
            guest += names
        if b == "lyrics":
            lyrics += names
    return unique(songs), unique(score), unique(guest), unique(lyrics)


def strip_label(s: str) -> str:
    """Drop role labels such as 'Story:' or '(dialogues)'; a credit labelled only as lyrics becomes ''."""
    if re.search(r"\((?:lyrics?|lyricist)\)\s*$", s, re.I):
        return ""
    s = re.sub(r"\s*\([^)]*(?:story|screen|dialog|script|writer|scenario|lyric|adapt|additional|novel|version|idea"
               r"|credited)[^)]*\)", "", s, flags=re.I)
    return re.sub(r"^[A-Za-z &/()-]{3,30}:\s*", "", s).strip(" ;,")


def lyricists(text: str) -> list[str]:
    names = []
    for tl in all_template_params(text, r"track\s*listing"):
        for k, v in tl.items():
            if re.fullmatch(r"all_lyrics|lyrics\d+", k):
                names += items(v)
    music = re.sub(r"<ref[^>]*>.*?</ref>|<ref[^>]*/>", "", section(text, r"soundtrack|music|songs"), flags=re.S | re.I)
    stop = re.compile(r"\b(?:sung|sang|sing|singers?|vocals?|voiced|playback|perform|render|music|compos|releas"
                      r"|label|album|track|featur|record)", re.I)
    for sent in re.split(r"(?<=\.)\s", music):
        i = sent.lower().find("lyric")
        if i < 0:
            continue
        tail = sent[i:]
        s = stop.search(tail, 5)
        tail = tail[:s.start()] if s else tail
        run = re.search(r"(?:\bby|lyricists?|penned|written|:)\s+((?:\[\[[^\]]+\]\][\s,]*(?:and\s+|&\s*)?)+)", tail)
        if run:
            names += [clean(l) for l in re.findall(r"\[\[[^\]]+\]\]", run.group(1))]
    return unique(n for n in names if n and not re.search(r"film|album|soundtrack|music|records|\d{4}", n, re.I))


def cast_list(text: str) -> list[tuple[str, str]]:
    out = []
    for line in section(text, r"^cast\b|cast$|starring").split("\n"):
        if not line.lstrip().startswith("*"):
            continue
        s = clean(line).lstrip("*: ").strip()
        if not s:
            continue
        actor, _, char = s.partition(" as ")
        out.append((actor.strip(), char.strip()))
    return out


def name_match(actor: str, s: str) -> bool:
    return all(tok in s.lower() for tok in actor.lower().split())


def mechanical(text: str, actor: str) -> tuple[dict, list[str]]:
    ib = template_params(text, r"infobox\s+film") or {}
    f, notes = {}, []
    if not ib:
        notes.append("Film article has no infobox.")
    rel = clean(ib.get("released"))
    dates = [(m.group(1), m.group(2) or "") for m in
             re.finditer(r"(\d{4}(?:-\d\d(?:-\d\d)?)?)(?:\s*\(([^)]*)\))?", rel)]
    if not dates:
        for m in re.finditer(r"(\d{1,2})\s+([A-Z][a-z]+)\s+(\d{4})|([A-Z][a-z]+)\s+(\d{1,2}),?\s+(\d{4})", rel):
            d, mon, y = (m.group(1), m.group(2), m.group(3)) if m.group(1) else (m.group(5), m.group(4), m.group(6))
            try:
                dates.append((dt.datetime.strptime(f"{d} {mon[:3]} {y}", "%d %b %Y").date().isoformat(), ""))
            except ValueError:
                pass
    if dates:
        india = [d for d, loc in dates if "india" in loc.lower()]
        f["release_date"] = india[0] if india else dates[0][0]
        if len(dates) > 1:
            notes.append("Infobox release dates: " + "; ".join(f"{d} ({loc})" if loc else d for d, loc in dates) + ".")
    rt = re.search(r"(\d+)\s*min", clean(ib.get("runtime"))) or re.fullmatch(r"\s*(\d+)\s*", clean(ib.get("runtime")))
    if rt:
        f["runtime"] = int(rt.group(1))
    for key, field in (("director", "director"), ("producer", "producers")):
        v = unique(strip_label(s) for s in items(ib.get(key)))
        if v:
            f[field] = "; ".join(v)
    banner = unique(items(ib.get("production_companies") or ib.get("studio") or ib.get("production_company")))
    if banner:
        f["production_banner"] = "; ".join(banner)
    writers = unique(strip_label(s) for k in ("writer", "story", "screenplay", "dialogue", "dialogues", "dialogs")
                     for s in items(ib.get(k)))
    if writers:
        f["writers"] = "; ".join(writers)
    lang = unique(re.sub(r"\s*\(.*?\)", "", s) for s in items(ib.get("language")))
    if lang:
        f["language"] = "; ".join(lang)
    based = clean(ib.get("based_on")).replace("\n", " ")
    remake = re.search(r"remake of (?:the )?([^.;]{3,140})", clean(lead(text)).replace("\n", " "))
    if based:
        f["source_material"] = based if based.lower().startswith("based on") else "Based on " + based
    elif remake:
        f["source_material"] = "Remake of " + remake.group(1).strip()
    songs, score, guest, lyr_ib = parse_music(ib.get("music"))
    if not songs:
        tl = template_params(text, r"track\s*listing") or {}
        songs = unique(items(tl.get("all_music") or tl.get("all_writing")))
    if songs:
        f["song_composer"] = "; ".join(songs)
    if score:
        f["background_score_composer"] = "; ".join(score)
    if guest:
        notes.append("Guest or additional composition credit: " + "; ".join(guest) + ".")
    lyr = unique(lyr_ib + lyricists(text))
    if lyr:
        f["lyricist"] = "; ".join(lyr)
    starring = items(ib.get("starring"))
    pos = next((i for i, s in enumerate(starring, 1) if name_match(actor, s)), None)
    if pos:
        f["billing_position"] = pos
    cur = None
    b = money(clean(ib.get("budget")))
    if b:
        cur, f["budget"] = b
        if re.search(r"\best\b|estimated|approx", clean(ib.get("budget")), re.I):
            notes.append("Budget is an estimate.")
    for line in clean(ib.get("gross")).split("\n"):
        mv = money(line)
        if not mv:
            continue
        low = re.sub(r"<[^>]*>?", "", line).lower()
        field = ("worldwide_gross" if re.search(r"\b(worldwide|world-wide|global)\b", low) else
                 "overseas_gross" if re.search(r"\b(overseas|international)\b", low) else
                 "domestic_nett" if re.search(r"\bnett?\b", low) else
                 "domestic_gross" if re.search(r"\b(india|indian|domestic)\b", low) else None)
        if field is None or field in f:
            notes.append(f"Infobox gross '{line.strip()}' has no nett/gross label and is not entered.")
            continue
        unit, val = mv
        if cur and unit != cur and not isinstance(val, str):
            val = line.strip()
        cur = cur or unit
        f[field] = val
    if cur:
        f["currency"] = cur
    return f, notes


# ---------------------------------------------------------------- research files

MECH = ["title", "release_date", "language", "runtime", "director", "producers", "production_banner", "writers",
        "source_material", "song_composer", "background_score_composer", "lyricist", "character",
        "alternate_identity", "role_type", "billing_position", "budget", "domestic_nett", "domestic_gross",
        "overseas_gross", "worldwide_gross", "currency", "source_ids", "research_notes"]
AWARD_NOTE = re.compile(r"award|nominated|\bwon\b|filmfare|national film", re.I)


def load_config(path: str | None) -> dict:
    return json.loads(Path(path).read_text()) if path else {}


def norm(title: str) -> str:
    return re.sub(r"[^a-z0-9]", "", title.lower())


def role_fields(row: dict, cast: list[tuple[str, str]], actor: str) -> tuple[dict, list[str]]:
    role, notes = row.get("role", "").strip(), row.get("notes", "")
    f, extra = {}, []
    if not role:
        hit = next((c for a, c in cast if name_match(actor, a)), None)
        if hit:
            role = hit
            extra.append(f"Character taken from the cast list ('{hit}').")
    low, nlow = role.lower(), notes.lower()
    role = re.sub(r"\s*\((?:cameo|special appearance|guest appearance|uncredited)[^)]*\)", "", role, flags=re.I)
    if "/" in role:  # double role, disguise or alias
        parts = [p.strip() for p in role.split("/") if p.strip()]
        f["character"], f["alternate_identity"] = parts[0], "; ".join(parts[1:])
    elif role:
        f["character"] = role
    if "narrator" in low or ("narrat" in nlow and not role):
        f["role_type"] = "Narrator"
    elif re.search(r"\bhim ?self\b", low):
        f["role_type"] = "Self appearance"
    elif re.search(r"voice", nlow + " " + low):
        f["role_type"] = "Voice role"
    elif "cameo" in nlow + " " + low:
        f["role_type"] = "Cameo"
    elif re.search(r"special appearance|guest appearance|extended appearance", nlow + " " + low):
        f["role_type"] = "Extended special appearance"
    elif re.search(r"child (?:artist|actor|role)", nlow):
        f["role_type"] = "Child role"
    kept = [n.strip() for n in notes.split(";") if n.strip() and not AWARD_NOTE.search(n)]
    if kept:
        extra.insert(0, "Filmography notes: " + "; ".join(kept) + ".")
    return f, extra


def cmd_films(a) -> int:
    work, research = Path(a.work), Path(a.research)
    cfg = load_config(a.config)
    table = json.loads((work / "table.json").read_text())
    pages = json.loads((work / "pages.json").read_text())
    rows = table["rows"] + cfg.get("add", [])
    today = dt.date.today().isoformat()
    films_path, src_path = research / "films.json", research / "sources.json"
    films = json.loads(films_path.read_text()) if films_path.exists() else []
    sources = json.loads(src_path.read_text()) if src_path.exists() else []
    for s in sources:
        if isinstance(s.get("film_ids"), str):
            s["film_ids"] = [x for x in s["film_ids"].split("; ") if x]
        s.setdefault("film_ids", [])
    by_key = {(norm(r["title"]), str(r.get("release_date", ""))[:4]): r for r in films}
    by_url = {s.get("url"): s for s in sources}
    next_src = max([int(s["source_id"][-3:]) for s in sources] + [0]) + 1
    next_film = max([int(r["film_id"][-3:]) for r in films] + [0]) + 1

    def source_row(title: str, url: str, claims: str, note: str) -> dict:
        nonlocal next_src
        if url not in by_url:
            by_url[url] = {"source_id": f"SOURCE-{next_src:03d}", "title": title, "publisher": "Wikipedia",
                           "url": url, "access_date": today, "claims": claims, "source_type": "Reference encyclopedia",
                           "reliability": "C", "notes": note, "film_ids": []}
            sources.append(by_url[url]); next_src += 1
        return by_url[url]

    fg = source_row(table["page"], WIKI_URL.format(table["page"].replace(" ", "_")),
                    "Feature-film acting credits, character names, and cameo, special-appearance, narration and voice notes.",
                    "Downloaded as raw wikitext.")
    built, report = [], {"excluded": [], "no_article": [], "no_date": [], "year_moved": [], "no_role_type": []}
    for row in rows:
        key = f"{row['title']}|{row['year']}"
        if row.get("pending") or key in cfg.get("exclude", {}):
            report["excluded"].append(key)
            continue
        resolved = pages.get(row.get("target")) if row.get("target") else None
        if resolved and resolved.startswith("List of"):  # red link redirected to a year list
            resolved = None
        text =page_file(work, resolved).read_text(encoding="utf-8") if resolved else ""
        mech, notes = mechanical(text, a.actor) if text else ({}, ["No Wikipedia film article; fields come from the filmography table."])
        if not text:
            report["no_article"].append(key)
        rf, rnotes = role_fields(row, cast_list(text) if text else [], a.actor)
        mech.update(rf)
        mech["title"] = row["title"]
        if "release_date" not in mech:
            if row["year"]:
                mech["release_date"] = str(row["year"])
            report["no_date"].append(key)
        elif row["year"] and mech["release_date"][:4] != str(row["year"]):
            notes.append(f"Filmography table lists {row['year']}; infobox release date is {mech['release_date']}.")
            report["year_moved"].append(f"{key} -> {mech['release_date']}")
        mech.update(cfg.get("overrides", {}).get(key, {}))
        if not mech.get("role_type"):
            report["no_role_type"].append(key)
        src_ids = [fg["source_id"]]
        if resolved:
            claims = "Infobox: " + ", ".join(k.replace("_", " ") for k in MECH[1:] if k in mech and k not in
                                             ("title", "character", "alternate_identity", "role_type")) + "."
            s = source_row(resolved, WIKI_URL.format(resolved.replace(" ", "_")), claims, "Film article; downloaded as raw wikitext.")
            src_ids.append(s["source_id"])
        for extra in cfg.get("extra_sources", {}).get(key, []):
            if extra["url"] not in by_url:
                by_url[extra["url"]] = {"source_id": f"SOURCE-{next_src:03d}", "access_date": today, "film_ids": [], **extra}
                sources.append(by_url[extra["url"]]); next_src += 1
            src_ids.append(by_url[extra["url"]]["source_id"])
        mech["source_ids"] = "; ".join(src_ids)
        mech["research_notes"] = " ".join(rnotes + notes + ([cfg["notes"][key]] if key in cfg.get("notes", {}) else [])) or None
        built.append((row, mech, resolved))

    built.sort(key=lambda b: (b[1].get("release_date") or "9999"))
    out = []
    for row, mech, resolved in built:
        existing = by_key.get((norm(mech["title"]), mech.get("release_date", "")[:4]))
        if existing is None:
            existing = {"film_id": f"FILM-{next_film:03d}"}
            next_film += 1
        for k in MECH + [k for k in mech if k not in MECH]:  # overrides may set other fields
            if mech.get(k) is not None and (a.overwrite or existing.get(k) in (None, "")):
                existing[k] = mech[k]
        out.append(existing)
        sids = existing["source_ids"]
        for sid in (sids if isinstance(sids, list) else sids.split("; ")):
            s = next(x for x in sources if x["source_id"] == sid)
            if existing["film_id"] not in s["film_ids"]:
                s["film_ids"].append(existing["film_id"])
    kept_ids = {r["film_id"] for r in out}
    out += [r for r in films if r["film_id"] not in kept_ids]
    if not films:  # fresh run: number chronologically
        for i, r in enumerate(out, 1):
            old, r["film_id"] = r["film_id"], f"FILM-{i:03d}"
            for s in sources:
                s["film_ids"] = [r["film_id"] if x == old else x for x in s["film_ids"]]
        for s in sources:
            s["film_ids"] = sorted(set(s["film_ids"]))
    out.sort(key=lambda r: (r.get("release_date") or "9999", r["film_id"]))
    research.mkdir(parents=True, exist_ok=True)
    films_path.write_text(json.dumps(out, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    src_path.write_text(json.dumps(sources, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"{len(out)} films -> {films_path}; {len(sources)} sources -> {src_path}")
    fill = {k: sum(1 for r in out if r.get(k) not in (None, "")) for k in MECH[1:]}
    print("filled: " + ", ".join(f"{k} {v}" for k, v in fill.items()))
    for k, v in report.items():
        print(f"{k} ({len(v)}): " + ("; ".join(v) if v else "none"))
    return 0


# ---------------------------------------------------------------- extracts for hand-written fields


def select_films(research: Path, spec: str) -> list[dict]:
    films = json.loads((research / "films.json").read_text())
    ids = []
    for part in spec.split(","):
        m = re.fullmatch(r"\s*FILM-(\d+)\s*-\s*FILM-(\d+)\s*", part)
        ids += [f"FILM-{n:03d}" for n in range(int(m.group(1)), int(m.group(2)) + 1)] if m else [part.strip()]
    return [f for f in films if f["film_id"] in ids]


def film_page(work: Path, research: Path, film: dict) -> str:
    sources = json.loads((research / "sources.json").read_text())
    for s in sources:
        if s.get("film_ids") == [film["film_id"]] and "wikipedia.org/wiki/" in (s.get("url") or ""):
            title = urllib.parse.unquote(s["url"].split("/wiki/", 1)[1]).replace("_", " ")
            return ensure_page(work, title)[1] or ""
    return ""


def for_each(a, fn) -> int:
    work, research = Path(a.work), Path(a.research)
    for film in select_films(research, a.ids):
        text = film_page(work, research, film)
        print(f"== {film['film_id']} {film['title']} ({str(film.get('release_date', ''))[:4]}) "
              f"[{film.get('role_type') or '?'}: {film.get('character') or '?'}]")
        print(fn(text, a) if text else "(no article)")
    return 0


def terms_filter(text: str, terms: str | None, cap: int = 400) -> str:
    keys = [t.strip().lower() for t in (terms or "").split(";") if t.strip()]
    out = []
    for line in text.split("\n"):  # keep headings and cast bullets from running into sentences
        line = line.strip()
        if not line or line.startswith("="):
            continue
        parts = [line] if line.startswith(("*", "#")) else sentences(line)
        out += [p for p in parts if not keys or any(k in p.lower() for k in keys)]
    return "\n".join(p[:cap] for p in out)


def cmd_plot(a) -> int:
    def fn(text, a):
        body = clean(section(text, r"^(plot|synopsis|story|premise|plot summary)$"))
        if a.terms:  # cameos: sentences naming the actor, from the whole article if the plot has none
            hits = terms_filter(body, a.terms)
            return hits or ("(outside plot) " + (terms_filter(clean(text), a.terms) or "no matching sentences"))
        return body.replace("\n", " ")[:a.chars] if body else "(no plot section)"
    return for_each(a, fn)


def cmd_cast(a) -> int:
    return for_each(a, lambda text, a: "; ".join(f"{x} = {c}" if c else x for x, c in cast_list(text)[:a.max]) or "(no cast list)")


def cmd_reviews(a) -> int:
    def fn(text, a):
        body = clean(section(text, r"reception|critic|review|response", exclude=r"box office|commercial"))
        return terms_filter(body, a.terms) or "(no matching sentences)"
    return for_each(a, fn)


def cmd_boxoffice(a) -> int:
    def fn(text, a):
        body = clean(section(text, r"box[- ]office|commercial"))
        keep = [s[:300] for s in sentences(body) if MONEY.search(s) or re.search(r"\b(hit|flop|blockbuster|verdict|average)\b", s, re.I)]
        return "\n".join(keep) or "(no box-office sentences)"
    return for_each(a, fn)


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("table"); s.add_argument("work"); s.add_argument("--page", required=True)
    s.add_argument("--section", default="acting credits|^films?$|filmography"); s.set_defaults(fn=cmd_table)
    s = sub.add_parser("fetch"); s.add_argument("work"); s.add_argument("--config"); s.set_defaults(fn=cmd_fetch)
    s = sub.add_parser("films"); s.add_argument("work"); s.add_argument("research"); s.add_argument("--actor", required=True)
    s.add_argument("--config"); s.add_argument("--overwrite", action="store_true"); s.set_defaults(fn=cmd_films)
    for name, fn in (("plot", cmd_plot), ("cast", cmd_cast), ("reviews", cmd_reviews), ("boxoffice", cmd_boxoffice)):
        s = sub.add_parser(name); s.add_argument("work"); s.add_argument("research"); s.add_argument("--ids", required=True)
        s.add_argument("--terms"); s.add_argument("--chars", type=int, default=2500); s.add_argument("--max", type=int, default=25)
        s.set_defaults(fn=fn)
    a = p.parse_args()
    return a.fn(a)


if __name__ == "__main__":
    sys.exit(main())
