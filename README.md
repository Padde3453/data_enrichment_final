People Enricher & Lead Scraper

Overview

This tool is an automated data enrichment script designed to process a list of company websites from an Excel file. It crawls each website to identify key decision-makers (CEOs, Founders, Managers), extracts contact information, and generates personalized outreach content (icebreakers and salutations) using OpenAI's LLM.

Key Capabilities

Intelligent Web Crawling:

Uses Playwright to render modern JavaScript-heavy websites.

Automatically discovers and prioritizes relevant sub-pages (e.g., "Team", "Contact", "Imprint", "About Us") to find personnel data.

Configurable adherence to robots.txt rules.

LLM-Powered Extraction:

Utilizes OpenAI (GPT-4o) to analyze page text and extract structured data about people (Names, Roles, Gender).

Filters for decision-making roles (e.g., "Geschäftsführer", "Inhaber", "Leitung") to ensure high-quality leads.

Generates a concise Company Summary based on the homepage text.

Personalization Engine:

Icebreakers: Generates a unique, non-generic icebreaker sentence based on the company's specific homepage content (Toggleable).

Salutations: Creates culturally correct salutations (Formal/Informal) based on the detected gender and configured language.

Multi-Language Support: Fully supports output in German (de), French (fr), Italian (it), and Spanish (es).

Robust Data Handling:

Excel Integration: Reads from an input sheet and appends enriched data to an output sheet without overwriting existing history.

Email Validation: Extracts emails via Regex and mailto: links, with an option to strictly filter for emails matching the company domain.

Concurrency & Safety: Handles file locking, retries on network failures, and includes rate limiting for API calls.

Installation

Prerequisites:

Python 3.8+

An OpenAI API Key

Install Dependencies:

pip install playwright openai beautifulsoup4 openpyxl tenacity python-dotenv tldextract


Install Playwright Browsers:

playwright install chromium


Configuration

Configuration is managed via a combination of Environment Variables (for keys and dynamic settings) and Script Constants (for core logic).

1. Environment Variables (.env or Key.env)

Create a .env file in the same directory:

Variable

Description

Default

OPENAI_API_KEY

Required. Your OpenAI API key.

None

OPENAI_MODEL

The LLM model to use.

gpt-4o-mini

ENRICH_LANGUAGE

Language for Salutations & Icebreakers. Options: de, fr, it, es.

es

TONE

Tone of voice. Options: formal, informal.

formal

ENABLE_ICEBREAKER

Toggle icebreaker generation. Set to false to disable.

true

2. Script Options (Top of Python File)

You can modify the CONFIG section inside people_enricher_language_v1.py to tune the scraper's behavior:

Constant

Description

Standard Value

EXCEL_PATH

Full path to your input/output Excel file.

.../Data enrichement.xlsx

INPUT_SHEET_NAME

The sheet containing the list of websites to process.

"Analyse"

OUTPUT_SHEET_NAME

The sheet where results will be written.

"Enriched Sheet"

HEADLESS

Run browser in background (True) or visible (False).

True

IGNORE_ROBOTS

If True, ignores robots.txt rules (useful for testing).

True

STRICT_DOMAIN_ONLY

If True, only saves emails that match the website domain.

False

MAX_ROWS_TO_PROCESS

Limit how many rows to process per run (0 = unlimited).

0

LLM_MAX_PAGES

Max number of sub-pages per site to send to LLM (saves costs).

2

Usage

Prepare Excel File: Ensure your Excel file has the input sheet (default: "Analyse") with Company Name in Column A and Website URL in Column B.

Run the Script:

python people_enricher_language_v1.py


Review Output: The script will print the configuration and a countdown, then process sites one by one. Results are saved to the "Enriched Sheet" in real-time.

Troubleshooting

Excel File Locked: If the script errors with "File appears open/locked", close the Excel file and try again.

Missing Imports: Ensure you have installed all requirements. If dotenv warning appears, it is safe to ignore as the script has a fallback loader.

Timeouts: If sites are loading too slowly, increase NAV_TIMEOUT_MS or PAGE_TIMEOUT_S in the script config.
