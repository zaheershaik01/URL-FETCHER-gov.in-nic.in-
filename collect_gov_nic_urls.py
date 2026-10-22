import json
import time
from urllib.parse import urlparse

import requests
from openpyxl import Workbook
from openpyxl.styles import Font

SUFFIXES = ("gov.in", "nic.in")     
MAX_PAGES = 200                
CRAWLS = 1                         
HEADERS = {"User-Agent": "gov-nic-url-collector/3.0 (research script)"}
DOTTED = tuple("." + s for s in SUFFIXES)  


def get(url, params=None, timeout=180):
    """Try up to 4 times. Returns the reply, or None if every try fails."""
    for attempt in range(1, 5):
        try:
            reply = requests.get(url, params=params, headers=HEADERS, timeout=timeout)
            reply.raise_for_status()
            return reply
        except Exception as error:
            print(f"   try {attempt}/4 failed: {type(error).__name__}")
            time.sleep(5 * attempt)          # wait longer after each failure
    return None


def keep(name):
    """Clean a name. Returns it if it ends with gov.in or nic.in, else None."""
    name = name.strip().lower().lstrip("*.").rstrip(".")
    if "@" not in name and (name in SUFFIXES or name.endswith(DOTTED)):
        return name
    return None



# source 1: crt.sh
def from_crtsh():
    found = set()                          
    for suffix in SUFFIXES:
        print(f"[crt.sh] asking for *.{suffix} ...")
        reply = get(f"https://crt.sh/?q=%25.{suffix}&output=json"
                    "&exclude=expired&deduplicate=Y")
        if reply is None:
            continue
        try:
            rows = reply.json()
        except ValueError:                   # server sent an error page, not data
            continue
        for row in rows:
            for name in row["name_value"].split("\n"):   # 1 certificate = many names
                if keep(name):
                    found.add(keep(name))
    return found



# source 2: commoncrawl
def from_common_crawl():
    found = set()
    reply = get("https://index.commoncrawl.org/collinfo.json", timeout=60)
    if reply is None:
        return found

    for crawl in reply.json()[:CRAWLS]:      # newest snapshots first
        api = crawl["cdx-api"]
        for suffix in SUFFIXES:
            print(f"[commoncrawl] asking for *.{suffix} ...")
            query = {"url": f"*.{suffix}", "output": "json", "fl": "url"}

            # how many result pages exist?
            reply = get(api, {**query, "showNumPages": "true"})
            if reply is None:
                continue
            total = reply.json().get("pages", 1)

            # Results are sorted A-Z. Reading only the first pages gives only
            # names starting with a-c, so we pick pages spread over the WHOLE list.
            step = max(1, total // MAX_PAGES)
            for page in list(range(0, total, step))[:MAX_PAGES]:
                reply = get(api, {**query, "page": page})
                if reply is not None:
                    for line in reply.text.splitlines():
                        try:
                            host = urlparse(json.loads(line)["url"]).hostname or ""
                        except (ValueError, KeyError):
                            continue
                        if keep(host):
                            found.add(keep(host))
                print(f"   page {page} done, total found so far: {len(found)}")
                time.sleep(1)                # be polite to the server
    return found



# SAVE: 2 files per suffix
def save(hosts):
    for suffix in SUFFIXES:
        urls = sorted(f"https://{h}" for h in hosts
                      if h == suffix or h.endswith("." + suffix))
        name = suffix.replace(".", "_") + "_urls"        # gov_in_urls / nic_in_urls

        if not urls:                         # never create an empty file
            print(f"No {suffix} URLs found, so {name} files were not written.")
            continue

        with open(name + ".txt", "w", encoding="utf-8") as file:     # text file
            file.write("\n".join(urls))

        book = Workbook()                                            # Excel file
        sheet = book.active
        sheet.title = suffix
        sheet.append(["S.No", "Website URL"])
        for number, url in enumerate(urls, start=1):
            sheet.append([number, url])
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        sheet.freeze_panes = "A2"
        sheet.column_dimensions["B"].width = 55
        try:
            book.save(name + ".xlsx")
        except PermissionError:              # file is open in Excel
            print(f"Close {name}.xlsx in Excel and run again.")
        print(f"Saved {len(urls)} {suffix} URLs -> {name}.txt, {name}.xlsx")



crt = from_crtsh()
crawl = from_common_crawl()
print(f"\ncrt.sh found {len(crt)} | Common Crawl found {len(crawl)}")
save(crt | crawl)                           
print("All done.")
