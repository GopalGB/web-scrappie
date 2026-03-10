<h1 align="center">web_scrappie</h1>

<p align="center">
  <b>Desktop GUI tool for scraping product images & metadata from e-commerce websites</b>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python 3.8+">
  <img src="https://img.shields.io/badge/license-MIT-green?style=for-the-badge" alt="MIT License">
  <img src="https://img.shields.io/badge/GUI-customtkinter-blue?style=for-the-badge" alt="CustomTkinter">
  <img src="https://img.shields.io/badge/browser-Chrome-orange?style=for-the-badge&logo=googlechrome&logoColor=white" alt="Chrome">
</p>

---

Give it a spreadsheet of URLs, it opens each page in a real browser, grabs every product image and title it can find, and writes everything to a beautifully formatted Excel workbook -- one sheet per category, with optional embedded thumbnails.

## Why?

I got tired of manually copying product images and titles from retail sites. Now I maintain a spreadsheet of URLs and let this handle the rest.

## Features

- **Smart input parsing** -- reads URLs from `.ods`, `.xlsx`, `.xls`, or `.pdf` files
- **Anti-detection** -- uses `undetected-chromedriver` to bypass bot protection
- **Deep page scraping** -- auto-scrolls, clicks "Load More" buttons, waits for lazy content
- **3 extraction methods** -- tries preloaded state (React/Next.js), DOM links, then standalone images
- **Parallel image downloads** -- configurable thread count for fast downloads
- **Excel with thumbnails** -- embeds scaled images directly into cells
- **Formatted output** -- alternating row colors, frozen headers, auto-filters, hyperlinked URLs
- **Auto-installs dependencies** -- just run it, first launch handles everything
- **Handles SSL issues** -- works on restricted/corporate networks

## Quick Start

### Prerequisites

- **Python 3.8+**
- **Google Chrome** installed

### Run

```bash
# Clone
git clone https://github.com/GopalGB/web-scrappie.git
cd web-scrappie

# Run (dependencies install automatically on first launch)
python web_scrappie.py
```

Or install dependencies manually first:

```bash
pip install -r requirements.txt
python web_scrappie.py
```

## Input File Format

Your spreadsheet needs at least two columns. The tool auto-detects which column has categories and which has URLs based on header names.

| Category | URL |
|----------|-----|
| Shoes | `https://example.com/shoes` |
| Bags | `https://example.com/bags` |
| Watches | `https://example.com/watches` |

- If headers aren't recognized, it assumes column 1 = category, column 2 = URL
- Multiple sheets are supported -- each sheet is processed independently
- PDFs: extracts every URL found in text and annotations

## Settings

All configurable from the GUI:

| Setting | Default | Description |
|---------|---------|-------------|
| **Max Scrolls** | 15 | Times to scroll down per page (for lazy-loaded content) |
| **Scroll Pause** | 2.0s | Wait time between scrolls |
| **Page Wait** | 8s | Wait time after initial page load |
| **DL Threads** | 8 | Parallel threads for image downloading |
| **Headless** | Off | Run Chrome invisibly (faster, but some sites block it) |
| **Download Images** | Off | Download images locally & embed thumbnails in Excel |

## How the Scraper Works

The scraper tries three methods in order, using the first one that returns results:

1. **Preloaded State** -- Many React/Next.js sites embed product data in `window.__PRELOADED_STATE__` or `window.__NEXT_DATA__`. Pulls structured data directly. Fastest and most reliable.

2. **Product Links** -- Finds all `<a>` tags containing `<img>` elements. Extracts image source and alt text as the product title.

3. **Standalone Images** -- Falls back to grabbing every `<img>` with alt text, filtering out icons and tiny spacer images.

## Output

The Excel workbook includes:

- **Summary sheet** -- category names, item counts, generation timestamp
- **One sheet per category** -- columns: #, Title, Image URL, Page URL
- **Embedded thumbnails** (if enabled) -- scaled to fit cells without stretching
- **Styling** -- alternating row colors, frozen header rows, auto-filters, clickable hyperlinks

## Tips

- **Getting zero results?** Uncheck headless mode -- some sites aggressively block headless browsers
- **SSL errors?** The tool handles certificate issues automatically, useful on corporate networks
- **Small images skipped** -- files under 500 bytes (tracking pixels, spacers) are filtered out
- **Deduplication** -- results are deduplicated by image URL within each category

## Tech Stack

| Component | Library |
|-----------|---------|
| GUI | [customtkinter](https://github.com/TomSchimansky/CustomTkinter) |
| Browser automation | [selenium](https://www.selenium.dev/) + [undetected-chromedriver](https://github.com/ultrafunkamsterdam/undetected-chromedriver) |
| Spreadsheet I/O | [pandas](https://pandas.pydata.org/) + [openpyxl](https://openpyxl.readthedocs.io/) + [odfpy](https://github.com/eea/odfpy) |
| PDF parsing | [pdfplumber](https://github.com/jsvine/pdfplumber) |
| Image processing | [Pillow](https://pillow.readthedocs.io/) |
| HTTP | [requests](https://docs.python-requests.org/) |

## License

[MIT](LICENSE)

---

<p align="center">Made by <b>Gopal Bagaswar</b></p>
