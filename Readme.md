# Hacker News — Daily Tech Intelligence Scraper

## Business Problem
Tech teams and researchers need to monitor trending topics in the developer community daily. Manual browsing is time-consuming and inconsistent.

## What It Does
Automatically scrapes the top 30 posts from Hacker News and extracts:
- Post title
- Score (upvotes)
- Link
- Username
- Time posted

## Output
Professional formatted Excel report with color-coded columns, auto-sized rows, and frozen headers — ready for sharing or analysis.

## Tech Stack
Python, Selenium, openpyxl, Pandas

## How to Run
1. Install dependencies: `pip install selenium webdriver-manager pandas openpyxl`
2. Run scraper: `python main.py`
3. Run formatter: `python formatter.py`
4. Open `hackernews_report.xlsx`
