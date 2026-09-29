from html.parser import HTMLParser

from playwright.sync_api import sync_playwright


PAGE_URL = "https://hamptonresearch.com/make-tray.php"
RESULT_URL = "https://hamptonresearch.com/includes/processtray.php"


class FirstTableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.table_depth = 0
        self.found_table = False
        self.rows = []
        self.current_row = None
        self.current_cell = None

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            if not self.found_table:
                self.found_table = True
                self.table_depth = 1
            elif self.table_depth:
                self.table_depth += 1
        elif self.table_depth and tag == "tr":
            self.current_row = []
        elif self.table_depth and tag in {"td", "th"} and self.current_row is not None:
            self.current_cell = []
        elif self.table_depth and tag == "br" and self.current_cell is not None:
            self.current_cell.append("\n")

    def handle_data(self, data):
        if self.current_cell is not None:
            self.current_cell.append(data)

    def handle_endtag(self, tag):
        if tag in {"td", "th"} and self.current_cell is not None:
            cell = "".join(self.current_cell).replace("\xa0", " ").strip()
            self.current_row.append(cell)
            self.current_cell = None
        elif tag == "tr" and self.current_row is not None:
            self.rows.append(self.current_row)
            self.current_row = None
        elif tag == "table" and self.table_depth:
            self.table_depth -= 1


def parse_conditions(rows):
    conditions = {}
    for row in rows[1:]:
        for cell in row[1:]:
            reagents = [
                " ".join(line.split())
                for line in cell.splitlines()
                if line.strip()
            ]
            if reagents:
                conditions[len(conditions)] = " ".join(reversed(reagents))
    return conditions


def scrape_maketray():
    result_url = RESULT_URL.rstrip("/")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(channel="chrome", headless=False)
        try:
            page = browser.new_page()
            page.goto(PAGE_URL, wait_until="domcontentloaded")
            print("Complete the Make Tray form in Chrome and submit it.")
            print("Waiting for the generated tray table...")

            response = page.wait_for_event(
                "response",
                predicate=lambda item: (
                    item.url.split("?", 1)[0].rstrip("/") == result_url
                    and item.request.method == "POST"
                ),
                timeout=0,
            )
            if not response.ok:
                raise RuntimeError(
                    f"Make Tray returned HTTP {response.status}: {response.text()[:1000]}"
                )
            parser = FirstTableParser()
            parser.feed(response.text())
        finally:
            browser.close()

    if not parser.rows:
        raise RuntimeError("The response was received, but no HTML table was found.")

    conditions = parse_conditions(parser.rows)
    return conditions


def main():
    print(scrape_maketray())


if __name__ == "__main__":
    main()
