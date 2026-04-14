#!/usr/bin/env python3
"""Export Calendar B view as JPEG screenshot."""
import sys
from playwright.sync_api import sync_playwright

OUTPUT = "/Users/marie/Melia/calendar-b-export.jpg"
URL = "http://localhost:58901/travelcore-rm-hub.html"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    page.goto(URL, wait_until="networkidle")

    # Scroll to calendar section and wait for render
    page.evaluate("""() => {
        // Show the week view (Daily B)
        var wv = document.getElementById('weekView');
        if (wv) wv.style.display = 'block';
    }""")
    page.wait_for_timeout(1000)

    # Find the calendar monthly section
    cal_section = page.query_selector(".cal-months-grid")
    if not cal_section:
        # Fallback: try the whole calendar card
        cal_section = page.query_selector(".cal-card")

    if cal_section:
        cal_section.screenshot(path=OUTPUT, type="jpeg", quality=95)
        print(f"Exported calendar to {OUTPUT}")
    else:
        # Full page fallback — scroll to calendar area
        page.evaluate("document.querySelector('.cal-grid-wrap, .cal-months-grid')?.scrollIntoView()")
        page.wait_for_timeout(500)
        page.screenshot(path=OUTPUT, type="jpeg", quality=95, full_page=False)
        print(f"Exported full viewport to {OUTPUT}")

    browser.close()
