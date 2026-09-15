#!/usr/bin/env python3
"""Palimpsest corpus harvester.

Discovers real-world .docx documents on the public web from two independent
sources and downloads the valid ones into this directory:

  A. Web search engines (DuckDuckGo HTML/Lite, Mojeek, Bing, 4get) driven by
     `filetype:docx <topic>` queries.
  B. The Internet Archive Wayback CDX index, queried per-host across a broad
     spread of government / court / university / NGO hosts, plus a
     keyword-filtered topic sweep.

Discovery and download run as a streaming producer/consumer pipeline: each CDX
response is handed straight to the download pool, so files land continuously
instead of after a full enumeration. manifest.json is rewritten after every
accepted file, so a partial run is never lost.

Only harvesting + validity checking happens here. No document analysis.

Python 3 standard library only.
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import queue
import random
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.parse
import zipfile
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
MANIFEST = os.path.join(HERE, "manifest.json")

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36"
)

TARGET = 110              # keep harvesting until this many accepted files
MIN_BYTES = 5 * 1024
MAX_BYTES = 8 * 1024 * 1024
DL_WORKERS = 8
DISC_WORKERS = 5          # CDX 504s under heavier parallelism
CDX_TRIES = 2             # drop the host after this many failures
PER_DOMAIN_CAP = 7        # corpus diversity: cap accepted files per host
DEADLINE = float(os.environ.get("HARVEST_DEADLINE", 1500))  # wall seconds

# --------------------------------------------------------------------------
# discovery inputs
# --------------------------------------------------------------------------

# Source B hosts. Leafy subdomains first: their slice of the Wayback index is
# small, so they answer in seconds, where apex domains take minutes or 504.
CDX_HOSTS = [
    # --- US federal, leafy ---
    "www.nist.gov", "www.gsa.gov", "www.ftc.gov", "www.sec.gov",
    "www.osha.gov", "www.usgs.gov", "oig.hhs.gov", "www.gao.gov",
    "grants.nih.gov", "www.federalregister.gov", "www.fws.gov",
    "www.nrc.gov", "www.fema.gov", "www.tva.gov", "www.ojp.gov",
    "www.bjs.gov", "www.eia.gov", "www.ncjrs.gov", "www.bia.gov",
    "nifa.usda.gov", "www.ams.usda.gov", "www.fsis.usda.gov",
    "www.ars.usda.gov", "www.nrcs.usda.gov", "www.uspto.gov",
    "ustr.gov", "www.energy.gov", "www.arctic.gov", "www.nsf.gov",
    "www.sba.gov", "www.fcc.gov", "www.doi.gov", "www.hud.gov",
    "www.bls.gov", "www.census.gov", "www.usaid.gov", "www.state.gov",
    "www.epa.gov", "www.noaa.gov", "www.cdc.gov", "www.nih.gov",
    "www.nasa.gov", "www.senate.gov", "www.house.gov",
    # --- courts / state / municipal ---
    "www.courts.state.nc.us", "www.utcourts.gov", "www.courts.ca.gov",
    "www.nycourts.gov", "www.seattle.gov", "www.austintexas.gov",
    "www.nyc.gov", "www.phila.gov", "www.kingcounty.gov",
    "www.portlandoregon.gov", "www.baltimorecity.gov", "www.smcgov.org",
    "www.michigan.gov", "www.mass.gov", "www.oregon.gov", "www.wa.gov",
    "www.virginia.gov", "www.texas.gov", "www.ca.gov", "www.ny.gov",
    # --- universities ---
    "web.mit.edu", "ocw.mit.edu", "registrar.berkeley.edu",
    "grad.berkeley.edu", "hr.harvard.edu", "www.ucop.edu",
    "research.utexas.edu", "policies.northwestern.edu", "www.mit.edu",
    "www.harvard.edu", "www.berkeley.edu", "www.stanford.edu",
    "www.cornell.edu", "www.umich.edu", "www.washington.edu",
    "www.ox.ac.uk", "www.cam.ac.uk", "www.ucl.ac.uk", "www.ed.ac.uk",
    "www.brookings.edu", "www.rand.org", "www.urban.org",
    # --- intergovernmental / NGO / nonprofit ---
    "www.un.org", "www.who.int", "www.worldbank.org", "www.redcross.org",
    "www.aclu.org", "www.unicef.org", "www.unesco.org", "www.ifrc.org",
    "www.oecd.org", "www.ilo.org", "www.unhcr.org", "www.imf.org",
    "www.savethechildren.org", "www.habitat.org", "www.amnesty.org",
    "www.gov.uk", "www.nhs.uk", "europa.eu",
]

# Source A: `filetype:docx <topic>` search queries.
SEARCH_TOPICS = [
    "board meeting minutes", "grant proposal", "court motion",
    "city council agenda", "job description", "annual report",
    "memorandum of understanding", "employment contract", "syllabus",
    "press release", "policy draft", "incident report", "resume",
    "cover letter", "statement of work", "RFP response", "lease agreement",
    "IRB protocol", "thesis template", "meeting agenda", "budget request",
    "strategic plan", "audit report", "nondisclosure agreement",
    "position paper", "staff report", "committee report",
]

# Source B, topic mode: keyword-constrained CDX sweeps, so document *kinds*
# are represented even though the search engines are unreachable here.
CDX_TOPIC_KEYWORDS = [
    ("meeting minutes", "[Mm]inutes"), ("meeting agenda", "[Aa]genda"),
    ("annual report", "[Rr]eport"), ("grant proposal", "[Pp]roposal"),
    ("employment contract", "[Cc]ontract"), ("syllabus", "[Ss]yllabus"),
    ("policy draft", "[Pp]olicy"), ("application form", "[Aa]pplication"),
    ("cover letter", "[Ll]etter"), ("resume", "[Rr]esume"),
    ("strategic plan", "[Pp]lan"), ("memorandum", "[Mm]emo"),
    ("thesis template", "[Tt]emplate"), ("guidelines", "[Gg]uidelines"),
    ("public notice", "[Nn]otice"), ("budget request", "[Bb]udget"),
    ("IRB protocol", "[Pp]rotocol"), ("testimony", "[Tt]estimony"),
    ("job description", "[Dd]escription"),
    ("statement of work", "[Ss]tatement"),
]

CDX_TOPIC_HOSTS = [
    "www.nist.gov", "www.gsa.gov", "www.epa.gov", "www.nih.gov",
    "www.senate.gov", "www.house.gov", "web.mit.edu", "www.harvard.edu",
    "www.un.org", "www.who.int", "www.worldbank.org", "www.ca.gov",
]

_print_lock = threading.Lock()
T0 = time.time()


def log(msg: str) -> None:
    with _print_lock:
        sys.stdout.write("[%5.0fs] %s\n" % (time.time() - T0, msg))
        sys.stdout.flush()


def out_of_time() -> bool:
    return (time.time() - T0) > DEADLINE


def curl(args, timeout):
    """Run curl, return (returncode, stdout bytes)."""
    try:
        p = subprocess.run(["curl", "-sS", "-A", UA] + args,
                           capture_output=True, timeout=timeout)
        return p.returncode, p.stdout
    except subprocess.TimeoutExpired:
        return 28, b""
    except Exception:
        return 1, b""


# --------------------------------------------------------------------------
# source A: web search engines
# --------------------------------------------------------------------------

DOCX_RE = re.compile(rb"""https?://[^\s"'<>\\)\]]+?\.docx""", re.I)


def _raw_docx_links(body: bytes):
    return [m.group(0).decode("utf-8", "replace") for m in DOCX_RE.finditer(body)]


def _unescape(u: str) -> str:
    return u.replace("&amp;", "&").replace("\\/", "/")


def engine_ddg(query, endpoint):
    """DuckDuckGo HTML / Lite. Result links are wrapped as ?uddg=<urlencoded>."""
    rc, body = curl([
        "--get",
        "--data-urlencode", "q=filetype:docx " + query,
        "--data", "kl=us-en",
        "-H", "Accept: text/html,application/xhtml+xml;q=0.9,*/*;q=0.8",
        "-H", "Accept-Language: en-US,en;q=0.9",
        "-H", "Referer: https://duckduckgo.com/",
        endpoint,
    ], timeout=45)
    if rc != 0 or not body:
        raise RuntimeError("curl rc=%d" % rc)
    if b"uddg=" not in body:
        raise RuntimeError("no uddg= in response (rate-limited or challenged)")
    urls = []
    for m in re.finditer(rb"uddg=([^&\"'\s]+)", body):
        urls.append(_unescape(urllib.parse.unquote(
            m.group(1).decode("ascii", "replace"))))
    urls.extend(_raw_docx_links(body))
    return urls


def engine_ddg_html(q):
    return engine_ddg(q, "https://html.duckduckgo.com/html/")


def engine_ddg_lite(q):
    return engine_ddg(q, "https://lite.duckduckgo.com/lite/")


def engine_mojeek(query):
    rc, body = curl(["--get", "--data-urlencode", "q=filetype:docx " + query,
                     "https://www.mojeek.com/search"], timeout=45)
    if rc != 0 or not body:
        raise RuntimeError("curl rc=%d" % rc)
    if b"403 - Forbidden" in body or b"automated queries" in body:
        raise RuntimeError("blocked (403 automated queries)")
    return [_unescape(u) for u in _raw_docx_links(body)]


def engine_brave(query):
    rc, body = curl(["--get", "--data-urlencode", "q=filetype:docx " + query,
                     "-H", "Accept-Language: en-US,en;q=0.9",
                     "https://search.brave.com/search"], timeout=45)
    if rc != 0 or not body:
        raise RuntimeError("curl rc=%d" % rc)
    if b"Too Many Requests" in body or b"captcha" in body.lower():
        raise RuntimeError("blocked (429 / captcha)")
    urls = [_unescape(u) for u in _raw_docx_links(body)]
    if not urls:
        raise RuntimeError("0 docx links")
    return urls


def engine_bing(query):
    rc, body = curl(["--get", "--data-urlencode", "q=filetype:docx " + query,
                     "--data", "count=50",
                     "https://www.bing.com/search"], timeout=45)
    if rc != 0 or not body:
        raise RuntimeError("curl rc=%d" % rc)
    urls = [_unescape(u) for u in _raw_docx_links(body)]
    # Bing wraps outbound links as /ck/a?...&u=a1<base64url>
    for m in re.finditer(rb"u=a1([A-Za-z0-9_\-]{8,})", body):
        s = m.group(1).decode("ascii", "replace")
        try:
            d = base64.urlsafe_b64decode(
                s + "=" * (-len(s) % 4)).decode("utf-8", "replace")
        except Exception:
            continue
        if d.lower().split("?")[0].endswith(".docx"):
            urls.append(d)
    if not urls:
        raise RuntimeError("0 docx links (filetype: operator ignored/blocked)")
    return urls


def engine_4get(query):
    rc, body = curl(["--get", "--data-urlencode", "s=filetype:docx " + query,
                     "https://4get.ca/web"], timeout=45)
    if rc != 0 or not body:
        raise RuntimeError("curl rc=%d" % rc)
    urls = [_unescape(u) for u in _raw_docx_links(body)]
    if not urls:
        raise RuntimeError("0 docx links (scraper returned no results)")
    return urls


ENGINES = [
    ("ddg-html", engine_ddg_html), ("ddg-lite", engine_ddg_lite),
    ("mojeek", engine_mojeek), ("brave", engine_brave),
    ("bing", engine_bing), ("4get", engine_4get),
]


# --------------------------------------------------------------------------
# source B: Wayback CDX
# --------------------------------------------------------------------------

def cdx_query(host, limit=40, name_filter=None, timeout=120):
    pat = (r"filter=original:.*\.docx$" if not name_filter
           else r"filter=original:.*%s.*\.docx$" % name_filter)
    params = ["url=" + host, "matchType=domain", pat,
              "filter=statuscode:200", "collapse=urlkey",
              "limit=%d" % limit, "fl=timestamp,original", "output=json"]
    args = ["-G", "--max-time", str(timeout)]
    for p in params:
        args += ["--data-urlencode", p]
    args.append("https://web.archive.org/cdx/search/cdx")
    rc, body = curl(args, timeout=timeout + 20)
    if rc != 0:
        raise RuntimeError("curl rc=%d" % rc)
    body = body.strip()
    if not body:
        raise RuntimeError("empty response (index scan timed out)")
    if body.startswith(b"<"):
        raise RuntimeError("html error page (gateway timeout)")
    rows = json.loads(body.decode("utf-8", "replace"))
    if rows and rows[0] and rows[0][0] == "timestamp":
        rows = rows[1:]
    return rows


# --------------------------------------------------------------------------
# download + validation
# --------------------------------------------------------------------------

def sanitize(name):
    name = urllib.parse.unquote(name)
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._-")
    if not name.lower().endswith(".docx"):
        name = (name or "doc") + ".docx"
    stem = name[:-5][:70] or "doc"
    return stem + ".docx"


def is_valid_docx(path):
    try:
        size = os.path.getsize(path)
    except OSError:
        return False, "missing"
    if size < MIN_BYTES:
        return False, "too small"
    if size > MAX_BYTES:
        return False, "too large"
    with open(path, "rb") as fh:
        if fh.read(2) != b"PK":
            return False, "not a zip"
    try:
        with zipfile.ZipFile(path) as zf:
            if "word/document.xml" not in zf.namelist():
                return False, "no word/document.xml"
            if zf.testzip() is not None:
                return False, "corrupt zip member"
    except zipfile.BadZipFile:
        return False, "bad zip"
    except Exception:
        return False, "zip error"
    return True, size


def fetch(url, dest):
    rc, _ = curl(["-L", "--max-time", "60", "-o", dest, url], timeout=80)
    return rc == 0 and os.path.exists(dest) and os.path.getsize(dest) > 0


def reg_domain(host):
    host = host.lower().split(":")[0].strip(".")
    parts = host.split(".")
    if host.endswith((".ac.uk", ".gov.uk", ".co.uk", ".org.uk", ".state.nc.us")):
        return ".".join(parts[-3:])
    return ".".join(parts[-2:])


# --------------------------------------------------------------------------
# pipeline
# --------------------------------------------------------------------------

class Harvester:
    def __init__(self):
        self.lock = threading.Lock()
        self.hashes = set()
        self.names = set()
        self.seen_urls = set()
        self.per_domain = defaultdict(int)
        self.manifest = []
        self.by_source = defaultdict(int)
        self.stats = defaultdict(int)
        self.failures = {}
        self.stop = threading.Event()
        self.cand_q = queue.Queue()

    # ---- adopt files a previous/parallel run already landed ---------------
    def adopt_existing(self):
        prior = {}
        if os.path.exists(MANIFEST):
            try:
                with open(MANIFEST) as fh:
                    for row in json.load(fh):
                        if isinstance(row, dict) and row.get("file"):
                            prior[row["file"]] = row
            except Exception:
                pass
        adopted = 0
        for fname in sorted(os.listdir(HERE)):
            if not fname.endswith(".docx"):
                continue
            path = os.path.join(HERE, fname)
            ok, info = is_valid_docx(path)
            if not ok:
                os.remove(path)
                self.stats["adopt_invalid:" + info] += 1
                continue
            with open(path, "rb") as fh:
                digest = hashlib.sha256(fh.read()).hexdigest()
            if digest in self.hashes:
                os.remove(path)
                self.stats["dup_sha256"] += 1
                continue
            row = prior.get(fname, {})
            url = row.get("url", "")
            self.hashes.add(digest)
            self.names.add(fname)
            if url:
                self.seen_urls.add(url.lower())
                self.per_domain[reg_domain(
                    urllib.parse.urlsplit(url).netloc)] += 1
            self.manifest.append({
                "file": fname, "url": url,
                "query": row.get("query", "wayback-cdx"),
                "bytes": info})
            self.by_source["adopted-prior-run"] += 1
            adopted += 1
        if adopted:
            self.write_manifest()
        log("adopted %d pre-existing valid .docx from disk" % adopted)

    # ---- candidate intake -------------------------------------------------
    def offer(self, url, query, source, wayback=None):
        url = url.split("#")[0]
        if not url.lower().split("?")[0].endswith(".docx"):
            return 0
        key = url.lower()
        with self.lock:
            if key in self.seen_urls:
                return 0
            self.seen_urls.add(key)
        self.cand_q.put({"url": url, "query": query,
                         "source": source, "wayback": wayback})
        return 1

    # ---- manifest ---------------------------------------------------------
    def write_manifest(self):
        rows = sorted(self.manifest, key=lambda m: m["file"])
        tmp = MANIFEST + ".tmp"
        with open(tmp, "w") as fh:
            json.dump(rows, fh, indent=1)
        os.replace(tmp, MANIFEST)

    # ---- download worker --------------------------------------------------
    def try_one(self, cand):
        url = cand["url"]
        reg = reg_domain(urllib.parse.urlsplit(url).netloc)
        with self.lock:
            if self.per_domain[reg] >= PER_DOMAIN_CAP:
                self.stats["skip_domain_cap"] += 1
                return

        attempts = []
        if cand.get("wayback"):
            attempts.append(("wb", "https://web.archive.org/web/%sid_/%s"
                             % (cand["wayback"], url)))
        attempts.append(("live", url))

        fd, tmp = tempfile.mkstemp(prefix=".dl_", dir=HERE)
        os.close(fd)
        try:
            for how, u in attempts:
                if self.stop.is_set():
                    return
                if not fetch(u, tmp):
                    self.stats["fetch_fail"] += 1
                    continue
                ok, info = is_valid_docx(tmp)
                if not ok:
                    self.stats["invalid:" + info] += 1
                    continue
                size = info
                with open(tmp, "rb") as fh:
                    digest = hashlib.sha256(fh.read()).hexdigest()
                with self.lock:
                    if digest in self.hashes:
                        self.stats["dup_sha256"] += 1
                        return
                    if self.per_domain[reg] >= PER_DOMAIN_CAP:
                        self.stats["skip_domain_cap"] += 1
                        return
                    base = sanitize(os.path.basename(
                        urllib.parse.urlsplit(url).path))
                    fname = "%s_%s" % (digest[:8], base)
                    if fname in self.names:
                        self.stats["dup_name"] += 1
                        return
                    self.hashes.add(digest)
                    self.names.add(fname)
                    self.per_domain[reg] += 1
                    shutil.move(tmp, os.path.join(HERE, fname))
                    self.manifest.append({
                        "file": fname, "url": url, "query": cand["query"],
                        "bytes": size})
                    self.by_source[cand["source"]] += 1
                    self.write_manifest()
                    n = len(self.manifest)
                    if n >= TARGET:
                        self.stop.set()
                log("  + [%3d] %-44s %7d B  %s/%s" %
                    (n, fname[:44], size, reg, how))
                return
        finally:
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass

    def dl_worker(self):
        while not self.stop.is_set():
            try:
                cand = self.cand_q.get(timeout=2)
            except queue.Empty:
                if self.disc_done.is_set():
                    return
                continue
            try:
                if not out_of_time():
                    self.try_one(cand)
            except Exception as e:
                self.stats["worker_error"] += 1
                self.failures.setdefault("dl", str(e))
            finally:
                self.cand_q.task_done()

    # ---- discovery workers ------------------------------------------------
    def disc_worker(self):
        while not self.stop.is_set() and not out_of_time():
            try:
                label, host, limit, kw = self.disc_q.get_nowait()
            except queue.Empty:
                return
            err = None
            for attempt in range(CDX_TRIES):
                try:
                    rows = cdx_query(host, limit=limit, name_filter=kw)
                    err = None
                    break
                except Exception as e:
                    err = str(e)
                    time.sleep(3 * (attempt + 1) + random.uniform(0, 2))
            if err:
                self.failures.setdefault("cdx:%s%s" % (
                    host, "/" + label if kw else ""), err)
                log("  cdx %-26s %-20s FAIL(%dx): %s" %
                    (host, label[:20], CDX_TRIES, err[:48]))
                continue
            fresh = 0
            for row in rows:
                if len(row) >= 2:
                    fresh += self.offer(row[1], label, "wayback-cdx", row[0])
            log("  cdx %-26s %-20s +%d queued (q=%d, have=%d)" %
                (host, label[:20], fresh, self.cand_q.qsize(),
                 len(self.manifest)))

    # ---- source A ---------------------------------------------------------
    def search_phase(self):
        mode = os.environ.get("HARVEST_SEARCH", "full")
        if mode == "skip":
            log("=== SOURCE A: skipped (engines already proven blocked) ===")
            return
        topics = SEARCH_TOPICS if mode == "full" else SEARCH_TOPICS[:1]
        tries = 2 if mode == "full" else 1
        log("=== SOURCE A: web search (filetype:docx), %d topic(s) ==="
            % len(topics))
        health = {}
        for qi, topic in enumerate(topics, 1):
            alive = [(n, f) for n, f in ENGINES if health.get(n, 0) < 3]
            if not alive:
                log("  all engines dead after %d queries; abandoning source A"
                    % (qi - 1))
                return
            for name, fn in alive:
                err = None
                for attempt in range(tries):
                    try:
                        urls = fn(topic)
                        err = None
                        break
                    except Exception as e:
                        err = str(e)
                        time.sleep(2 * (attempt + 1))
                if err:
                    health[name] = health.get(name, 0) + 1
                    self.failures.setdefault("search:" + name, err)
                    log("  %-9s %-26s FAIL: %s" % (name, topic[:26], err[:52]))
                    continue
                health[name] = 0
                fresh = sum(self.offer(u, "filetype:docx " + topic,
                                       "search:" + name) for u in urls)
                log("  %-9s %-26s +%d docx" % (name, topic[:26], fresh))
            if mode == "full":
                time.sleep(random.uniform(5, 10))

    # ---- orchestration ----------------------------------------------------
    def run(self):
        self.adopt_existing()
        self.search_phase()

        self.disc_q = queue.Queue()
        only = [h for h in os.environ.get("HARVEST_HOSTS", "").split(",") if h]
        hosts = only or CDX_HOSTS
        jobs = [(h, h, 40, None) for h in hosts]
        if not only:
            for i, (label, kw) in enumerate(CDX_TOPIC_KEYWORDS):
                jobs.append((label, CDX_TOPIC_HOSTS[i % len(CDX_TOPIC_HOSTS)],
                             25, kw))
        for j in jobs:
            self.disc_q.put(j)
        log("=== SOURCE B: Wayback CDX streaming (%d jobs) ===" % len(jobs))

        self.disc_done = threading.Event()
        dls = [threading.Thread(target=self.dl_worker, daemon=True)
               for _ in range(DL_WORKERS)]
        for t in dls:
            t.start()
        discs = [threading.Thread(target=self.disc_worker, daemon=True)
                 for _ in range(DISC_WORKERS)]
        for t in discs:
            t.start()
        for t in discs:
            t.join()
        log("  discovery finished; draining %d queued candidates"
            % self.cand_q.qsize())
        self.disc_done.set()
        for t in dls:
            t.join(timeout=max(30, DEADLINE - (time.time() - T0) + 120))
        self.stop.set()

        self.write_manifest()
        log("=== DONE in %.0fs ===" % (time.time() - T0))
        log("accepted: %d" % len(self.manifest))
        for k, v in sorted(self.by_source.items()):
            log("  by source  %-22s %d" % (k, v))
        for k, v in sorted(self.stats.items()):
            log("  rejected   %-22s %d" % (k, v))
        log("failed sources: %d" % len(self.failures))
        for k, v in sorted(self.failures.items()):
            log("  FAIL %-32s %s" % (k, v[:64]))
        log("manifest: %s" % MANIFEST)


if __name__ == "__main__":
    random.seed(20260914)
    Harvester().run()
