import requests
from datetime import datetime
import random
from bs4 import BeautifulSoup as bs
import pandas as pd
import feedparser
from dateutil import parser as date_parser
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os
from requests.exceptions import ChunkedEncodingError

# List of User-Agents
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36",
]

# Function to parse date using dateutil
def parse_date(date_string):
    try:
        parsed_date = date_parser.parse(date_string, fuzzy=True)
        return parsed_date.strftime('%Y-%m-%d %H:%M:%S')
    except (ValueError, TypeError):
        print(f"Error parsing date: {date_string}. Defaulting to current time.")
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

# Function to fetch and filter RSS feeds based on keywords
def rss_url_feed(keyword):
    rss_url = f"https://news.google.com/rss/search?q={keyword}+site:*"
    feed = feedparser.parse(rss_url)
    articles = []
    for entry in feed.entries:
        title = entry.title.lower()
        summary = entry.summary.lower() if 'summary' in entry else ""
        if keyword.lower() in title or keyword.lower() in summary:
            articles.append({
                'title': entry.title,
                'link': entry.link,
                'published': entry.published if 'published' in entry else None
            })
    return articles

# Function to fetch Google search results with retry mechanism
def fetch_google_results(query, start=0):
    url = f"https://www.google.com/search?q={query}&start={start}"
    headers = {"User-Agent": random.choice(USER_AGENTS)}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Raise HTTPError for bad responses
        soup = bs(response.text, "html.parser")
        return soup
    except (requests.exceptions.RequestException, ChunkedEncodingError) as e:
        print(f"Error fetching Google results: {e}. Retrying...")
        return None

# Function to extract search results
def extract_search_results(soup):
    results = []
    if soup:
        for item in soup.select('div.g'):
            title_tag = item.select_one('h3')
            if not title_tag:
                continue
            title = title_tag.get_text()
            link = item.select_one('a')['href']
            snippet_tag = item.select_one('.VwiC3b')
            snippet = snippet_tag.get_text() if snippet_tag else ""
            results.append({
                "title": title,
                "link": link,
                "snippet": snippet,
                "published": None  # No date available from Google search results
            })
    return results

# Function to load existing data from Excel file
def load_existing_data(file_path, sheet_name):
    if os.path.exists(file_path):
        try:
            return pd.read_excel(file_path, sheet_name=sheet_name)
        except ValueError:
            return pd.DataFrame(columns=["Time", "Keyword", "Title", "URL"])
    else:
        return pd.DataFrame(columns=["Time", "Keyword", "Title", "URL"])

# Function to remove rows containing specific keywords from DataFrame
def remove_rows_with_keywords(data, keywords):
    for keyword in keywords:
        data = data[~data['Title'].str.contains(keyword, case=False, na=False)]
        data = data[~data['URL'].str.contains(keyword, case=False, na=False)]
    return data

# Function to prepare final data without duplicates, remove LinkedIn rows, and filter by keyword
def prepare_final_data_without_dupes(excel_file, sheet_name, scraped_data):
    # Load existing data
    existing_data = load_existing_data(excel_file, sheet_name)

    # Apply necessary transformations to scraped data
    scraped_data['Time'] = scraped_data['published'].fillna(datetime.now().strftime('%Y-%m-%d %H:%M:%S')).apply(parse_date)
    scraped_data['Keyword'] = "manastik"
    scraped_data = scraped_data.rename(columns={"title": "Title", "link": "URL"})
    scraped_data = scraped_data[["Time", "Keyword", "Title", "URL"]]  # Ensure correct order of columns

    # Remove rows with specific keywords
    scraped_data = remove_rows_with_keywords(scraped_data, ["linkedin", "LinkedIn", "Linkedin"])
    existing_data = remove_rows_with_keywords(existing_data, ["linkedin", "LinkedIn", "Linkedin"])

    # Drop duplicates in existing data
    existing_data = existing_data.drop_duplicates(subset=["Title", "URL"])

    # Ensure uniqueness in scraped data
    scraped_data = scraped_data.drop_duplicates(subset=["Title", "URL"])

    # Identify unique data not in existing data
    unique_data = scraped_data[
        ~scraped_data[["Title", "URL"]].apply(tuple, axis=1).isin(existing_data[["Title", "URL"]].apply(tuple, axis=1))
    ]

    if unique_data.empty:
        print('No new data is unique.')
        return

    # Append unique data to the existing data
    final_data = pd.concat([existing_data, unique_data], ignore_index=True)

    # Save the final data to the Excel file
    with pd.ExcelWriter(excel_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        final_data.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Appended {len(unique_data)} unique records to the sheet '{sheet_name}'.")

    # Format columns in Excel
    format_excel_columns(excel_file, sheet_name)

# Function to format columns in Excel
def format_excel_columns(excel_file, sheet_name):
    wb = load_workbook(excel_file)
    ws = wb[sheet_name]

    # Set column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    # Make URL column clickable and blue
    for cell in ws['D']:
        if cell.value != "URL":  # Skip header
            cell.hyperlink = cell.value
            cell.style = "Hyperlink"

    wb.save(excel_file)

# Example usage
def main():
    keyword = "manastik"  # Only focusing on 'manastik'
    file_path = "keyword_scrapping_document226.xlsx"

    # Fetch RSS feed results
    rss_articles = rss_url_feed(keyword)
    all_results = [{
        'title': entry['title'],
        'link': entry['link'],
        'published': entry['published']
    } for entry in rss_articles]

    # Fetch Google search results
    for start in range(0, 400, 10):
        soup = fetch_google_results(keyword, start=start)
        if soup is not None:
            results = extract_search_results(soup)
            all_results.extend(results)

    # Convert all_results to DataFrame
    all_results_df = pd.DataFrame(all_results)

    # Prepare final data without duplicates and save to Excel
    prepare_final_data_without_dupes(file_path, keyword.lower(), all_results_df)

if __name__ == "__main__":
    main()
