"""Fetch cited non-Wikipedia pages for a batch and keep only paragraphs with given terms."""
import sys, json, re, subprocess, hashlib
from html.parser import HTMLParser
from pathlib import Path
sys.path.insert(0, ".claude/skills/bollywood-actor-filmography/scripts")
import extract_wiki as ex
UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"
class P(HTMLParser):
    def __init__(s): super().__init__(); s.paras=[]; s.buf=None; s.skip=0
    def handle_starttag(s, t, a):
        if t in ("script","style","noscript"): s.skip+=1
        if t in ("p","li","h1","h2"): s.buf=[]
    def handle_endtag(s, t):
        if t in ("script","style","noscript"): s.skip=max(0,s.skip-1)
        if t in ("p","li","h1","h2") and s.buf is not None:
            x=re.sub(r"\s+"," ","".join(s.buf)).strip()
            if len(x)>40: s.paras.append(x)
            s.buf=None
    def handle_data(s, d):
        if s.buf is not None and not s.skip: s.buf.append(d)
def get(url, work):
    f = work/"html"/(hashlib.md5(url.encode()).hexdigest()+".html")
    if not f.exists():
        subprocess.run(["curl","-sL","--max-time","25","-A",UA,"-o",str(f),url])
    return f.read_text(errors="ignore") if f.exists() else ""
def paras(html):
    p=P(); p.feed(html); return p.paras
if __name__=="__main__":
    work=Path(sys.argv[1]); research=Path(sys.argv[2]); spec=json.load(open(sys.argv[3]))
    films={f["film_id"]:f for f in json.load(open(research/"films.json"))}
    for job in spec:
        fid, dom, terms, cap = job["film"], job["domain"], job["terms"], job.get("cap",1500)
        text=ex.film_page(work, research, films[fid]) or ""
        urls=[u for u in re.findall(r"\|\s*url\s*=\s*([^|}\s]+)", text) if dom in u]
        url=job.get("url") or (urls[0] if urls else None)
        if not url: print(f"== {fid} {dom}: URL not found"); continue
        ps=paras(get(url, work))
        keep=[x for x in ps if any(t.lower() in x.lower() for t in terms)]
        out="\n".join(keep)[:cap]
        print(f"== {fid} {films[fid]['title']} | {url} | paras {len(ps)} kept {len(keep)}")
        print(out if out else "(nothing kept)")
