#!/usr/bin/env python3
"""
Playwright verification for CS Dashboard.
Renders the dashboard, checks for console errors, asserts no AI tells,
and writes screenshots for visual diff.
"""
from __future__ import annotations

import re
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE_URL = "http://localhost:5101/"
OUT = Path("/tmp/cs-dashboard-shots")
OUT.mkdir(parents=True, exist_ok=True)

# Banned characters per design-taste-frontend skill
EM_DASH = "\u2014"
EN_DASH_AS_SEP = re.compile(r"\s[\u2013]\s")  # en-dash used as separator
MIDDLE_DOT = "\u00b7"


def assert_no_em_dash(page_text: str) -> list[str]:
    issues = []
    if EM_DASH in page_text:
        # Limit to first 5 occurrences
        for line in page_text.splitlines():
            if EM_DASH in line:
                issues.append(f"em-dash: {line.strip()[:120]}")
                if len(issues) >= 5:
                    break
    return issues


def assert_middle_dot_restrained(page_text: str) -> list[str]:
    issues = []
    for line in page_text.splitlines():
        count = line.count(MIDDLE_DOT)
        if count > 1:
            issues.append(f"middle-dot x{count}: {line.strip()[:120]}")
    return issues


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1440, "height": 900})

        # Capture console errors
        console_errors: list[str] = []
        page.on("pageerror", lambda exc: console_errors.append(f"pageerror: {exc}"))
        page.on("console", lambda msg: console_errors.append(f"console.{msg.type}: {msg.text}")
                if msg.type == "error" else None)

        # 1. Desktop default
        page.goto(BASE_URL, wait_until="domcontentloaded", timeout=30000)
        time.sleep(5)
        page.screenshot(path=str(OUT / "verify-desktop.png"), full_page=True)

        # 2. Mobile
        page.set_viewport_size({"width": 390, "height": 844})
        page.reload(wait_until="domcontentloaded")
        time.sleep(4)
        page.screenshot(path=str(OUT / "verify-mobile.png"), full_page=True)

        # 3. Tablet
        page.set_viewport_size({"width": 820, "height": 1180})
        page.reload(wait_until="domcontentloaded")
        time.sleep(4)
        page.screenshot(path=str(OUT / "verify-tablet.png"), full_page=True)

        # 4. Filter interaction: switch to BON only
        page.set_viewport_size({"width": 1440, "height": 900})
        page.reload(wait_until="domcontentloaded")
        time.sleep(4)
        page.select_option("#team-select", "bon")
        time.sleep(5)
        page.screenshot(path=str(OUT / "verify-bon.png"), full_page=True)

        # 5. Filter interaction: month mode
        page.select_option("#team-select", "all")
        page.select_option("#mode-select", "date")
        time.sleep(1)
        # Type a date
        page.fill("#month-picker", "2026-05-15")
        time.sleep(5)
        page.screenshot(path=str(OUT / "verify-month-mode.png"), full_page=True)

        # 6. Capture rendered text for content checks
        body_text = page.inner_text("body")

        # 7. Section presence
        sections = page.locator(".section-title").all_inner_texts()

        # Run assertions
        problems: list[str] = []
        problems.extend(assert_no_em_dash(body_text))
        problems.extend(assert_middle_dot_restrained(body_text))

        if console_errors:
            problems.append(f"console errors: {console_errors[:5]}")

        expected_sections = [
            "Khiếu nại & Bảo hành",
            "Khảo sát khách hàng",
            "Hành vi mua hàng",
            "Vận hành & SLA",
        ]
        for s in expected_sections:
            if s not in sections:
                problems.append(f"missing section: {s}")

        # Check that background is light (computed bg color)
        bg = page.evaluate("getComputedStyle(document.body).backgroundColor")
        # Acceptable light: rgb(245,244,240) or similar warm-white
        if "rgb(0, 0, 0)" in bg or "rgb(5, 5, 5)" in bg:
            problems.append(f"body bg is dark: {bg}")

        # Print report
        print("=" * 60)
        print("CS Dashboard verification report")
        print("=" * 60)
        print(f"URL:           {BASE_URL}")
        print(f"Sections found: {len(sections)} -> {sections}")
        print(f"Body bg:       {bg}")
        print(f"Console errors: {len(console_errors)}")
        for e in console_errors[:10]:
            print(f"  - {e}")
        print(f"Problems:      {len(problems)}")
        for p_ in problems:
            print(f"  - {p_}")
        print()
        print(f"Screenshots written to: {OUT}/")
        for f in sorted(OUT.glob("verify-*.png")):
            print(f"  - {f.name}  ({f.stat().st_size//1024} KB)")

        browser.close()

        return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
