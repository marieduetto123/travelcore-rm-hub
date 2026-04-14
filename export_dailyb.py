#!/usr/bin/env python3
"""Export Daily B (7-day week view) as JPEG screenshot."""
from playwright.sync_api import sync_playwright

OUTPUT = "/Users/marie/Melia/dailyb-7day-export.jpg"
URL = "http://localhost:58901/travelcore-rm-hub.html"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    page.goto(URL, wait_until="networkidle")

    # Open the week view (Daily B) — month 3, day 1
    page.evaluate("""() => {
        if (typeof openWeekView === 'function') {
            openWeekView(3, 1);
        }
    }""")
    page.wait_for_timeout(2000)

    # Ensure dailyB groupBy is active
    page.evaluate("""() => {
        if (typeof wvGroupBy !== 'undefined') {
            wvGroupBy = 'dailyB';
        }
        if (typeof buildWeekGrid === 'function') {
            buildWeekGrid(3, 1, 1);
        }
    }""")
    page.wait_for_timeout(2000)

    # Screenshot the weekView section
    wv = page.query_selector("#weekView")
    if wv:
        wv.screenshot(path=OUTPUT, type="jpeg", quality=95)
        print(f"Exported Daily B week view to {OUTPUT}")
    else:
        # Fallback: full page
        page.screenshot(path=OUTPUT, type="jpeg", quality=95, full_page=True)
        print(f"Exported full page fallback to {OUTPUT}")

    browser.close()
