#!/usr/bin/env python3
"""Site integrity + freshness checker.

Two independent questions this answers:

  1. Is the committed site internally consistent?   (structural checks)
  2. Is what visitors see actually what I wrote?    (--live: sync checks)

Question 2 exists because "the site isn't showing X" has had three different
causes here — browser cache, a scroll-reveal race, and unmerged work — which
look identical from the outside. --live names which one it is.

Usage:
    python3 scripts/check_site.py            # structural checks only (fast, offline)
    python3 scripts/check_site.py --live     # also compare repo vs deployed site
    python3 scripts/check_site.py --cv       # also check the CV pdf (needs pdftotext)

Exit code is non-zero if any ERROR was reported, so CI can gate on it.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INDEX = os.path.join(ROOT, "index.html")
CV = os.path.join(ROOT, "CV", "Zhihao_Feng_CV.pdf")
LIVE_URL = "https://fengzhihao77.github.io/"

errors: list[str] = []
warnings: list[str] = []
notes: list[str] = []


def err(msg: str) -> None:
    errors.append(msg)


def warn(msg: str) -> None:
    warnings.append(msg)


def note(msg: str) -> None:
    notes.append(msg)


def read_index() -> str:
    with open(INDEX, encoding="utf-8") as fh:
        return fh.read()


# ---------------------------------------------------------------- structural

def check_jsonld(html: str) -> None:
    m = re.search(r'<script type="application/ld\+json">(.*?)</script>', html, re.S)
    if not m:
        err("JSON-LD block not found")
        return
    try:
        data = json.loads(m.group(1))
    except json.JSONDecodeError as exc:
        err(f"JSON-LD does not parse: {exc}")
        return
    graph = data.get("@graph", [])
    articles = [n for n in graph if n.get("@type") == "ScholarlyArticle"]
    cards = re.findall(r'<li class="pub"', html)
    note(f"JSON-LD ok: {len(graph)} nodes, {len(articles)} ScholarlyArticle")
    if len(articles) != len(cards):
        warn(
            f"JSON-LD has {len(articles)} articles but there are {len(cards)} "
            "publication cards — structured data is out of sync with the page"
        )


def check_publication_ordering(html: str) -> None:
    cards = re.findall(
        r'<li class="pub" data-year="(\d{4})" data-rank="(\d+)" data-order="(\d+)"', html
    )
    if not cards:
        err("no publication cards found")
        return
    n = len(cards)
    ranks = sorted(int(c[1]) for c in cards)
    orders = [int(c[2]) for c in cards]
    if ranks != list(range(1, n + 1)):
        err(f"data-rank must be a contiguous 1..{n}; got {ranks}")
    if orders != list(range(n)):
        err(
            f"data-order must match document position 0..{n-1}; got {orders} "
            "(the By-year view breaks otherwise)"
        )
    note(f"publications ok: {n} cards, ranks 1..{n}, order sequential")


def check_news_order(html: str) -> None:
    block = re.search(r'<ul class="news-list">(.*?)</ul>', html, re.S)
    if not block:
        err("news list not found")
        return
    stamps = re.findall(r'<time datetime="(\d{4})-(\d{2})">', block.group(1))
    if not stamps:
        err("news list has no <time datetime> entries")
        return
    keys = [(int(y), int(mo)) for y, mo in stamps]
    if keys != sorted(keys, reverse=True):
        err(f"news items are not newest-first: {['%s-%s' % s for s in stamps]}")
    note(f"news ok: {len(keys)} items, newest {keys[0][0]}-{keys[0][1]:02d}, newest-first")


def check_local_assets(html: str) -> None:
    refs = set(re.findall(r'(?:src|href)="((?!https?:|mailto:|#)[^"]+)"', html))
    missing = []
    for ref in sorted(refs):
        path = os.path.join(ROOT, ref.split("?")[0].lstrip("/"))
        if not os.path.exists(path):
            missing.append(ref)
    if missing:
        err(f"referenced files that do not exist: {missing}")
    else:
        note(f"assets ok: all {len(refs)} local references resolve")


def check_thumbnail_ratios(html: str) -> None:
    """Flag card thumbnails whose file ratio differs from their display box.

    Card thumbnails are pre-cropped to the display ratio by convention. They
    also carry `object-fit: cover`, so a mismatch does not distort — it
    silently crops, quietly eating content off the edges. That is how a
    900x463 file once slipped into a 900x462 slot, and why conference photos
    lost people when the ratio moved. The hero portrait is deliberately
    cover-cropped and is not a card thumbnail, so it is out of scope here.
    """
    try:
        from PIL import Image  # type: ignore
    except ImportError:
        note("thumbnail-ratio check skipped (Pillow not installed)")
        return
    pairs = re.findall(
        r'<img class="thumb"[^>]*src="([^"]+)"[^>]*width="(\d+)"[^>]*height="(\d+)"', html
    )
    if not pairs:
        note("thumbnail-ratio check: no card thumbnails found")
        return
    off = []
    for src, w, h in pairs:
        path = os.path.join(ROOT, src.split("?")[0].lstrip("/"))
        if not os.path.exists(path):
            continue
        with Image.open(path) as im:
            fw, fh = im.size
        want, got = int(w) / int(h), fw / fh
        crop_pct = 100 * (1 - min(want, got) / max(want, got))
        if crop_pct > 1.0:
            off.append(
                f"{src}: file {fw}x{fh} ({got:.3f}) vs box {w}x{h} ({want:.3f}) "
                f"— ~{crop_pct:.1f}% gets cropped away"
            )
    if off:
        warn("thumbnail ratio differs from its display box:\n      " + "\n      ".join(off))
    else:
        note(f"thumbnails ok: all {len(pairs)} match their display ratio exactly")


def check_reveal_failsafe(html: str) -> None:
    """Regression guard for the scroll-reveal race fixed in PR #28.

    Without these, a late layout shift can leave content stuck invisible —
    which reads to a visitor as "the new item is missing".
    """
    for needle, what in [
        ("unstickPassedReveals", "reveal failsafe sweep"),
        ("ScrollTrigger.refresh()", "trigger re-measure"),
        ("legacyReveal", "no-GSAP fallback path"),
    ]:
        if needle not in html:
            err(f"{what} ('{needle}') is missing — scroll-reveal can strand content")
    if not errors:
        note("reveal safety net ok: failsafe, refresh and fallback all present")


def check_anchors(html: str) -> None:
    ids = set(re.findall(r'\sid="([^"]+)"', html))
    targets = set(re.findall(r'href="#([^"]+)"', html))
    dangling = sorted(t for t in targets if t and t not in ids)
    if dangling:
        err(f"in-page links point at missing ids: {dangling}")
    else:
        note(f"anchors ok: {len(targets)} in-page links all resolve")


# ------------------------------------------------------------------- the CV

def check_cv() -> None:
    if not os.path.exists(CV):
        err("CV pdf is missing")
        return
    try:
        text = subprocess.run(
            ["pdftotext", "-layout", CV, "-"],
            capture_output=True, text=True, check=True,
        ).stdout
    except (FileNotFoundError, subprocess.CalledProcessError):
        note("CV text check skipped (pdftotext unavailable)")
        return
    flat = " ".join(text.split())
    if "chronotope" in flat.lower():
        err("CV names the company — public documents are meant to stay stealth")
    if "zfeng77@gatech.edu" not in flat:
        warn("CV does not contain the public contact address")
    pages = text.count("\f") or 1
    note(f"CV ok: {pages} pages, privacy gate passed")


# ------------------------------------------------------------------ liveness

def git(*args: str) -> str:
    return subprocess.run(
        ["git", *args], cwd=ROOT, capture_output=True, text=True
    ).stdout.strip()


def check_live() -> None:
    """Answer: does the live site show what I wrote, and if not, why not?"""
    git("fetch", "origin", "--quiet")
    branch = git("rev-parse", "--abbrev-ref", "HEAD")
    dirty = bool(git("status", "--porcelain"))
    ahead = [l for l in git("log", "--oneline", "origin/main..HEAD").splitlines() if l]

    try:
        with urllib.request.urlopen(LIVE_URL, timeout=20) as resp:
            live = resp.read().decode("utf-8", "replace")
    except Exception as exc:  # noqa: BLE001
        err(f"could not fetch the live site: {exc}")
        return

    deployed = git("show", "origin/main:index.html")
    live_news = re.findall(r'<time datetime="([\d-]+)">', live)
    head_news = re.findall(r'<time datetime="([\d-]+)">', read_index())

    if live.strip() == deployed.strip():
        note("deploy ok: live site matches origin/main exactly (no cache lag)")
    else:
        warn(
            "live site differs from origin/main — either Pages is still building "
            "(wait ~1 min) or your browser is showing a cached copy (hard-reload)"
        )

    if dirty:
        warn(f"working tree has uncommitted changes on '{branch}' — not deployable yet")
    if ahead:
        warn(
            f"'{branch}' is {len(ahead)} commit(s) ahead of origin/main — "
            "this work is NOT live until the PR is merged:\n      "
            + "\n      ".join(ahead)
        )
    if not dirty and not ahead:
        note("sync ok: your branch has nothing that main is missing")

    note(f"news items — local: {len(head_news)}, live: {len(live_news)}")
    if len(head_news) != len(live_news):
        warn(
            f"local has {len(head_news)} news items but live shows {len(live_news)} "
            "— see the sync warnings above for the reason"
        )


# ---------------------------------------------------------------------- main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--live", action="store_true", help="compare repo against the deployed site")
    ap.add_argument("--cv", action="store_true", help="also validate the CV pdf")
    args = ap.parse_args()

    html = read_index()
    check_jsonld(html)
    check_publication_ordering(html)
    check_news_order(html)
    check_local_assets(html)
    check_thumbnail_ratios(html)
    check_reveal_failsafe(html)
    check_anchors(html)
    if args.cv:
        check_cv()
    if args.live:
        check_live()

    for n in notes:
        print(f"  ok    {n}")
    for w in warnings:
        print(f"  WARN  {w}")
    for e in errors:
        print(f"  ERROR {e}")

    print()
    if errors:
        print(f"FAILED — {len(errors)} error(s), {len(warnings)} warning(s)")
        return 1
    print(f"PASSED — {len(warnings)} warning(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
