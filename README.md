# WebArticleArchiver

## Introduction
The **WebArticleArchiver** is an automated tool designed to scrape articles from various web platforms and convert them into high-quality PDF files. This tool first gathers all article links from a specified website, generates an Excel table containing these links along with relevant information (such as titles and publication dates), and then downloads each article and converts it into a PDF file.

## Features
- **Automated Link Scraping**: Automatically scrolls through the webpage to gather all article links.
- **Excel Table Generation**: Creates an Excel table (`articles.xlsx`) with details about each article.
- **PDF Conversion**: Downloads each article and converts it into a high-quality PDF file.
- **Retry Mechanism**: Includes a retry mechanism in case of network issues or loading timeouts.
- **Headless Mode**: Runs the browser in headless mode for better performance and reduced resource usage.
- **Customizable Options**: Allows users to skip certain articles or limit the number of articles to download.

## Requirements
- Python 3.x
- Required Libraries:
  - `beautifulsoup4`
  - `openpyxl`
  - `Pillow`
  - `img2pdf`
  - `selenium`
  - `webdriver-manager`
  - `requests`
  - `pdfkit`

You can install these libraries using pip:




