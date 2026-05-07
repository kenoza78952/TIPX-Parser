# Patent Marketplace Scraping Pipeline

Modular Python-based scraping and monitoring pipeline for tracking patent marketplace listings across multiple IP transaction platforms.

The system automates data extraction, change detection, PostgreSQL storage, Excel report generation, and email-based reporting workflows.

## Features

- Multi-source scraping pipeline covering 8+ patent marketplaces
- Requests + BeautifulSoup scraping for static sites
- Selenium-based scraping for JavaScript-heavy platforms
- PostgreSQL persistence layer with session tracking
- Detection of newly added and closed listings
- Automated Excel report generation
- Scheduled reporting workflows with email delivery
- Structured logging and debug utilities
- Modular scraper architecture with per-source parsers

## Supported Sources

- Vitek IP
- Tangible IP
- Transactions IP
- IP Investments Group
- GTT Group
- IPApproach
- Parallel North IP
- ICAP
- RZV

## Tech Stack

- Python
- Selenium
- BeautifulSoup4
- Requests
- PostgreSQL
- psycopg2
- SQLAlchemy
- pandas
- openpyxl

## Architecture

```text
Marketplace Sources
        ↓
Scraper Modules
        ↓
Parser & Normalization Layer
        ↓
PostgreSQL Storage
        ↓
Change Detection Logic
        ↓
Excel Report Generation
        ↓
Email Notification Pipeline
```

## Project Structure

```text
TIPX-Parser/
│
├── execute.py
├── dbconn.py
├── excel_reports.py
├── email_utils.py
├── debug_utils.py
│
├── vitek.py
├── transactionip.py
├── tangibleip.py
├── ipinvestments.py
├── gttgrp.py
├── icap.py
├── parallelnorth.py
├── rzv.py
│
└── README.md
```
## Notes

Designed for automated patent marketplace monitoring, portfolio tracking, and reporting workflows involving structured IP listing data across multiple external sources.
