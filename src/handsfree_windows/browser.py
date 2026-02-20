from __future__ import annotations

"""Browser automation module using Playwright.

Uses a persistent browser profile so login sessions survive between runs.
Default profile dir: ~/.handsfree-windows/browser-profile/
State file (last URL): ~/.handsfree-windows/browser-state.json
"""

import json
import os
from pathlib import Path
from typing import Any, Literal

BrowserType = Literal["chromium", "firefox", "webkit"]

_STATE_FILE = Path.home() / ".handsfree-windows" / "browser-state.json"
_PROFILE_BASE = Path.home() / ".handsfree-windows" / "browser-profiles"


def _profile_dir(browser: str) -> Path:
    p = _PROFILE_BASE / browser
    p.mkdir(parents=True, exist_ok=True)
    return p


def _save_state(url: str, browser: str) -> None:
    _STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    data = {"url": url, "browser": browser}
    _STATE_FILE.write_text(json.dumps(data), encoding="utf-8")


def _load_state() -> dict[str, str]:
    if _STATE_FILE.exists():
        try:
            return json.loads(_STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _launch(browser: BrowserType, headless: bool = False):
    """Launch (or reuse profile of) a Playwright persistent context."""
    from playwright.sync_api import sync_playwright

    pw = sync_playwright().start()

    if browser == "firefox":
        engine = pw.firefox
    elif browser == "webkit":
        engine = pw.webkit
    else:
        engine = pw.chromium

    ctx = engine.launch_persistent_context(
        user_data_dir=str(_profile_dir(browser)),
        headless=headless,
        args=["--start-maximized"] if browser == "chromium" and not headless else [],
        no_viewport=True if not headless else False,
        viewport={"width": 1280, "height": 800} if headless else None,
    )
    return pw, ctx


def _get_page(ctx, url: str | None = None):
    """Get current page (or open a new one, navigating to url)."""
    pages = ctx.pages
    page = pages[0] if pages else ctx.new_page()
    if url:
        page.goto(url, wait_until="domcontentloaded", timeout=30000)
    return page


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def open_url(url: str, browser: BrowserType = "chromium", headless: bool = False) -> dict[str, Any]:
    pw, ctx = _launch(browser, headless=headless)
    try:
        page = _get_page(ctx, url)
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        _save_state(page.url, browser)
        return {"url": page.url, "title": page.title()}
    finally:
        ctx.close()
        pw.stop()


def navigate(url: str) -> dict[str, Any]:
    state = _load_state()
    browser = state.get("browser", "chromium")
    pw, ctx = _launch(browser)
    try:
        page = _get_page(ctx)
        page.goto(url, wait_until="domcontentloaded", timeout=30000)
        _save_state(page.url, browser)
        return {"url": page.url, "title": page.title()}
    finally:
        ctx.close()
        pw.stop()


def snapshot(fmt: str = "aria") -> dict[str, Any]:
    """Return the accessibility tree or visible text of the current page."""
    state = _load_state()
    browser = state.get("browser", "chromium")
    url = state.get("url")
    pw, ctx = _launch(browser, headless=True)
    try:
        page = _get_page(ctx, url)
        page.wait_for_load_state("domcontentloaded", timeout=15000)

        if fmt == "text":
            result = page.evaluate("() => document.body.innerText")
        else:
            # Aria snapshot (Playwright 1.46+)
            try:
                result = page.accessibility.snapshot()
            except Exception:
                result = page.evaluate("() => document.body.innerText")

        return {"url": page.url, "title": page.title(), "content": result}
    finally:
        ctx.close()
        pw.stop()


def click(selector: str | None = None, text: str | None = None, exact: bool = False) -> dict[str, Any]:
    state = _load_state()
    browser = state.get("browser", "chromium")
    url = state.get("url")
    pw, ctx = _launch(browser)
    try:
        page = _get_page(ctx, url)
        page.wait_for_load_state("domcontentloaded", timeout=15000)

        if selector:
            page.click(selector, timeout=10000)
        elif text:
            page.get_by_text(text, exact=exact).first.click(timeout=10000)
        else:
            raise ValueError("Provide --selector or --text")

        _save_state(page.url, browser)
        return {"url": page.url, "action": "clicked"}
    finally:
        ctx.close()
        pw.stop()


def type_text(selector: str, text: str, clear: bool = True) -> dict[str, Any]:
    state = _load_state()
    browser = state.get("browser", "chromium")
    url = state.get("url")
    pw, ctx = _launch(browser)
    try:
        page = _get_page(ctx, url)
        page.wait_for_load_state("domcontentloaded", timeout=15000)

        el = page.locator(selector).first
        if clear:
            el.clear(timeout=10000)
        el.type(text, timeout=10000)

        _save_state(page.url, browser)
        return {"url": page.url, "action": "typed"}
    finally:
        ctx.close()
        pw.stop()


def screenshot(out: str = "screenshot.png", full_page: bool = False) -> dict[str, Any]:
    state = _load_state()
    browser = state.get("browser", "chromium")
    url = state.get("url")
    pw, ctx = _launch(browser, headless=True)
    try:
        page = _get_page(ctx, url)
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        path = str(Path(out).resolve())
        page.screenshot(path=path, full_page=full_page, timeout=15000)
        return {"url": page.url, "saved": path}
    finally:
        ctx.close()
        pw.stop()


def evaluate(js: str) -> dict[str, Any]:
    state = _load_state()
    browser = state.get("browser", "chromium")
    url = state.get("url")
    pw, ctx = _launch(browser, headless=True)
    try:
        page = _get_page(ctx, url)
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        result = page.evaluate(js)
        return {"url": page.url, "result": result}
    finally:
        ctx.close()
        pw.stop()


def get_links() -> dict[str, Any]:
    state = _load_state()
    browser = state.get("browser", "chromium")
    url = state.get("url")
    pw, ctx = _launch(browser, headless=True)
    try:
        page = _get_page(ctx, url)
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        links = page.evaluate("""() => {
            return Array.from(document.querySelectorAll('a[href]'))
                .map(a => ({text: a.innerText.trim(), href: a.href}))
                .filter(l => l.text && l.href)
                .slice(0, 200);
        }""")
        return {"url": page.url, "links": links}
    finally:
        ctx.close()
        pw.stop()


def fill_form(fields: list[dict[str, str]]) -> dict[str, Any]:
    """Fill multiple form fields at once.

    fields: list of {selector: str, text: str} dicts.
    """
    state = _load_state()
    browser = state.get("browser", "chromium")
    url = state.get("url")
    pw, ctx = _launch(browser)
    try:
        page = _get_page(ctx, url)
        page.wait_for_load_state("domcontentloaded", timeout=15000)

        for f in fields:
            sel = f.get("selector") or f.get("css")
            txt = f.get("text", "")
            if sel:
                el = page.locator(sel).first
                el.clear(timeout=5000)
                el.type(txt, timeout=5000)

        _save_state(page.url, browser)
        return {"url": page.url, "fields_filled": len(fields)}
    finally:
        ctx.close()
        pw.stop()
