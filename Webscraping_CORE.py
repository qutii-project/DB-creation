import requests
import pandas as pd
import asyncio
from playwright.async_api import async_playwright
import time

API_KEY = ""
CORE_BASE_URL = "https://api.core.ac.uk/v3/search/works"
JOURNAL_NAME = "Sustainability"
LIMIT = 100
MAX_PAGES = 100       
DELAY = 1             # Delay between requests to avoid rate limits
OUTPUT_FILE = "sustainability_authors_emails.xlsx"

all_articles = []

for page in range(1, MAX_PAGES + 1):
    print(f"Fetching CORE page {page}...")
    params = {
        "q": f'source.name:"{JOURNAL_NAME}" AND year:>=2020',
        "limit": LIMIT,
        "page": page,
        "apiKey": API_KEY
    }

    resp = requests.get(CORE_BASE_URL, params=params)
    if resp.status_code != 200:
        print(f"CORE API error {resp.status_code} on page {page}")
        break

    data = resp.json()
    results = data.get("results", [])
    if not results:
        print("No more results.")
        break

    for item in results:
        authors = item.get("authors", [])
        first_author = authors[0]["name"] if authors else ""
        all_articles.append({
            "Journal": item.get("source", {}).get("name", ""),
            "Title": item.get("title", ""),
            "Full Name": first_author,
            "DOI": item.get("doi", ""),
            "URL": item.get("links", [{}])[0].get("url", ""),
            "Topics": ', '.join(item.get("topics", []))
        })

    time.sleep(DELAY)

# Scrape main author emails
async def scrape_emails(articles):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        for article in articles:
            url = article["URL"]
            email = ""
            try:
                await page.goto(url, timeout=120000)
                await page.wait_for_load_state("networkidle")

               
                # Publisher-specific selectors
                if "mdpi.com" in url:
                    email_elem = await page.query_selector("a.toEncode.emailCaptcha")
                    email = email_elem.get_attribute("href") if email_elem else ""
                elif "wiley.com" in url:
                    email_elem = await page.query_selector("a[href^='mailto:']")
                    email = email_elem.get_attribute("href") if email_elem else ""
                elif "tandfonline.com" in url:
                    email_elem = await page.query_selector("a.corresponding-author-email")
                    email = email_elem.get_attribute("href") if email_elem else ""
                #---

                email = email.replace("mailto:", "") if email else ""
                article["Email"] = email
                print(f"Scraped email for: {article['Full Name']} ({email})")

            except Exception as e:
                article["Email"] = ""
                print(f"Error scraping {url}: {e}")

        await browser.close()

# Run scraping
asyncio.run(scrape_emails(all_articles))

# Step 3: Save to Excel
df = pd.DataFrame(all_articles)
df.to_excel(OUTPUT_FILE, index=False)
print(f"Saved {len(all_articles)} articles with emails to {OUTPUT_FILE}")
