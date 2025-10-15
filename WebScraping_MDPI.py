import pandas as pd
import asyncio
from playwright.async_api import async_playwright
import nest_asyncio
import openpyxl
import os
from openpyxl import load_workbook


nest_asyncio.apply()


EXCEL_FILE = "mdpi_articles_sheets_2.xlsx"
LINKS_FILE = "Covered_Articles.xlsx"

# Extract first and last name from full name
def extract_names(full_name):
    name_parts = full_name.split()
    first_name = name_parts[0] if len(name_parts) > 1 else full_name
    last_name = name_parts[-1] if len(name_parts) > 1 else ''
    return first_name, last_name

# Extract country from affiliation (last part after comma)
def extract_country_from_affiliation(affiliation_name):
    return affiliation_name.split(',')[-1].strip()  


def save_extracted_link(article_url):
    df_new = pd.DataFrame({"Article Link": [article_url]})

    try:
        if os.path.exists(LINKS_FILE):
            df_existing = pd.read_excel(LINKS_FILE, engine="openpyxl")
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_combined = df_new

        with pd.ExcelWriter(LINKS_FILE, engine="openpyxl", mode="w") as writer:
            df_combined.to_excel(writer, sheet_name="Extracted_Links", index=False)
        print(f"Saved link: {article_url}")

    except Exception as e:
        print(f"Error saving link {article_url}: {e}")




def save_to_excel(df, file_name, sheet_name):
    """Appends data to an existing sheet if present, else creates a new sheet."""
    try:
        if os.path.exists(file_name):
            with pd.ExcelWriter(file_name, engine="openpyxl", mode="a") as writer:
                book = load_workbook(file_name)
                if sheet_name in book.sheetnames:
                    startrow = book[sheet_name].max_row
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=startrow)
                else:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(file_name, engine="openpyxl", mode="w") as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Data successfully saved to {file_name} in sheet '{sheet_name}'")
    except Exception as e:
        print(f"Error saving to Excel: {e}")


async def extract_article_details(article_url, page):
    try:
        await page.goto(article_url)

        title_element = await page.query_selector('h1.title.hypothesis_container')
        article_title = await title_element.inner_text() if title_element else ''

      
        while True:
            author_elements = await page.query_selector_all('span.inlineblock')
            if author_elements:
                break
            print(f"Waiting for authors to load: {article_url}")
            await page.wait_for_timeout(60000)

        authors_data = []
        for author in author_elements:
            author_name_element = await author.query_selector('div.profile-card-drop')
            profile_name = await author_name_element.inner_text() if author_name_element else ''

            first_name, last_name = extract_names(profile_name.strip())

            sup_element = await author.query_selector('sup')
            sup_text = await sup_element.inner_text() if sup_element else '0'
            sup_numbers = [num.strip() for num in sup_text.replace('*', '').split(',') if num.strip()]
            sup_numbers = sup_numbers if sup_numbers else ['0']

            email_element = await author.query_selector('a.toEncode.emailCaptcha')
            email = await email_element.get_attribute('href') if email_element else ''
            email = email.replace('mailto:', '') if email else ''

            for sup_number in sup_numbers:
                authors_data.append({
                    'Sup Number': sup_number,
                    'Full Name': profile_name.strip(),
                    'First Name': first_name,
                    'Last Name': last_name,
                    'Email': email.strip()
                })

        author_df = pd.DataFrame(authors_data)
        author_df = author_df[author_df['Email'] != '']  # Filter out empty emails

        affiliation_elements = await page.query_selector_all('div.affiliation')
        affiliations = []
        for aff in affiliation_elements:
            sup_element = await aff.query_selector('div.affiliation-item sup')
            sup_text = await sup_element.inner_text() if sup_element else '0'
            if sup_text == '*':
                continue

            aff_name_element = await aff.query_selector('div.affiliation-name')
            affiliation_name = await aff_name_element.inner_text() if aff_name_element else ''
            country = extract_country_from_affiliation(affiliation_name)

            affiliations.append({
                'Sup Number': sup_text,
                'Affiliation': affiliation_name.strip(),
                'Country': country
            })

        aff_df = pd.DataFrame(affiliations)

        merged_df = pd.merge(author_df, aff_df, on="Sup Number", how="inner")
        merged_df['Research Title'] = article_title
        merged_df['Journal'] = article_url
        merged_df['Title'] = ''

        column_order = ["Journal", "Title", "Full Name", "First Name", "Last Name", "Email",
                        "Affiliation", "Country", "Research Title"]
        merged_df = merged_df[column_order]

        return merged_df

    except Exception as e:
        print(f"Failed to scrape {article_url}: {e}")
        return None

async def scrape_all_articles():
    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            page = await browser.new_page()

            base_url = "https://www.mdpi.com/search?sort=pubdate&page_no={}&page_count=100&year_from=2021&year_to=2025&journal=energies&view=default"
            all_data = []
            start_page, end_page =  26,41
            sheet_name = f"Page_{start_page}_{end_page-1}"

            extracted_links = set(pd.read_excel(LINKS_FILE).get("Article Link", [])) if os.path.exists(LINKS_FILE) else set()

            for page_num in range(start_page, end_page):
                print(f"Scraping page {page_num}...")

                search_url = base_url.format(page_num)
                await page.goto(search_url)
                await page.wait_for_load_state("networkidle", timeout=120000)

                for _ in range(5):
                    await page.evaluate("window.scrollBy(0, window.innerHeight)")
                    await page.wait_for_timeout(1000)

                await page.wait_for_selector("a.title-link", timeout=60000)
                article_elements = await page.query_selector_all("a.title-link")
                article_links = ["https://www.mdpi.com" + await a.get_attribute("href") 
                                 for a in article_elements if await a.get_attribute("href")]
            
                for article_url in article_links:
                    if article_url in extracted_links:
                        print(f"Skipping already extracted: {article_url}")
                        continue
                    try:
                        article_df = await extract_article_details(article_url, page)
                        if article_df is not None:
                            all_data.append(article_df)
                            extracted_links.add(article_url)
                    except Exception as e:
                        print(f"Error scraping {article_url}: {e}")
        
    finally:
        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)

            
            sheet_name = f"Page_{start_page}_{end_page-1}"  
            save_to_excel(final_df, "mdpi_articles_sheets_2.xlsx", sheet_name)
        print("Scraping complete! Data saved to mdpi_articles_sheets")
        await browser.close()


task = asyncio.ensure_future(scrape_all_articles())
await task
