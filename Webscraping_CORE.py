import requests
import pandas as pd
import time

API_KEY = ""
BASE_URL = "https://api.core.ac.uk/v3/search/works"
QUERY = "all"            # Broad query to cover all topics
LIMIT = 100              # Maximum results per page
MAX_PAGES = 200          # Adjust according to the estimated token usage
DELAY = 1                # Seconds to wait between requests (avoid throttling)

all_topics = set()

for page in range(1, MAX_PAGES + 1):
    print(f"Fetching page {page}...")
    params = {
        "q": QUERY,
        "limit": LIMIT,
        "page": page,
        "apiKey": API_KEY
    }

    try:
        response = requests.get(BASE_URL, params=params)
        if response.status_code != 200:
            print(f"Error {response.status_code} on page {page}, stopping.")
            break

        data = response.json()
        results = data.get("results", [])

        if not results:
            print("No more results, stopping.")
            break

        for item in results:
            topics = item.get("topics", [])
            for t in topics:
                all_topics.add(t.strip())

        time.sleep(DELAY)

    except Exception as e:
        print(f"Exception on page {page}: {e}")
        break

unique_topics = sorted(all_topics)

# Save to Excel
df = pd.DataFrame(unique_topics, columns=["Topic"])
df.to_excel("core_all_topics.xlsx", index=False)
print(f"Extraction complete! {len(unique_topics)} unique topics saved to core_all_topics.xlsx")
