# Stock Sentiment Analysis and Visualization

This project analyzes sentiment from social media platforms, correlates it with stock prices, and visualizes the results for selected stocks. It consists of five main components:

1. Data Scraping and Cleaning
2. Sentiment Analysis on Reddit Data
3. Stock Price Data Fetching and Correlation Calculation
4. Dashboard Visualization for Stocks
5. Output Results and Visualizations

## Table of Contents
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
  - [1. Data Scraping and Cleaning](#1-data-scraping-and-cleaning)
  - [2. Sentiment Analysis on Reddit Data](#2-sentiment-analysis-on-reddit-data)
  - [3. Stock Price Data Fetching and Correlation Calculation](#3-stock-price-data-fetching-and-correlation-calculation)
  - [4. Dashboard Visualization for Stocks](#4-dashboard-visualization-for-stocks)
  - [5. Output Results and Visualizations](#5-output-results-and-visualizations)
- [File Structure](#file-structure)

## Prerequisites
Before running the scripts, ensure you have the following installed:

- Python 3.7 or higher
- Libraries:
  - `pandas`
  - `yfinance`
  - `scipy`
  - `matplotlib`
  - `seaborn`
  - `vaderSentiment`
  - `openpyxl`
  - `praw`

You can install the required libraries using pip:

```bash
pip install pandas yfinance scikit-learn matplotlib seaborn vaderSentiment openpyxl praw re datetime nltk
```
## Installation
Clone the repository to your local machine:

```bash
git clone <repository-url>
cd <repository-directory>
```
Ensure your Reddit API credentials are set up correctly in the scraping code.

## Usage
1. Data Scraping and Cleaning
    - Open the Data scraping and cleaning.py script.
    - Replace the placeholder values for client_id, client_secret, and user_agent with your Reddit API credentials.
    - Define your search queries in the tickers dictionary.
    - Run the script. It will scrape Reddit posts related to the specified stocks, clean the text data, and save the results to an Excel file named reddit_data.xlsx.


2. Sentiment Analysis on Reddit Data
    - Open the sentiment analysis.py script.
    - Set the input_file_path to reddit_data.xlsx and specify the output_file_path where you want to save the sentiment analysis results.
    - Run the script. This will analyze the sentiment of post titles and body content, categorizing them as Positive, Negative, or Neutral, and will output the results to a new Excel file named sentiment analysis.xlsx.
   

3. Stock Price Data Fetching and Correlation Calculation
    - Open the stock price correlation.py script.
    - Set the sentiment_file_path to the location of sentiment analysis.xlsx.
    - Define the stock tickers you want to analyze in the tickers list.
    - Specify the start and end dates for the stock price data.
    - Run the script. It will fetch daily stock prices for the specified tickers, merge them with sentiment analysis results, and calculate the correlations between sentiment scores and stock prices. The results will be saved in correlation_and_merged_data.xlsx.
   

4. Dashboard Visualization for Stocks
    - Open the Data Visualization.py script.
    - Define the stock tickers and their corresponding company names in the companies dictionary.
    - Run the script. This will create visualizations for the merged stock and sentiment data, generating a dashboard for each stock and saving the visualizations as PNG files named {ticker}_dashboard.png.
   

5. Output Results and Visualizations
    - All results, including merged data and visualizations, will be saved in specified output files. The visualizations will be saved as PNG files for easy sharing.

## File Structure
```bash

/Stock-Sentiment-Analysis/
│
├── Data scraping and cleaning.py              # Script for data scraping and cleaning
├── Sentiment analysis.py          # Script for sentiment analysis
├── Stock price correlation.py     # Script for fetching stock data and correlation
├── Data_Visualization.py      # Script for creating visualizations
├── correlation_and_merged_data.xlsx # Output file with merged data
├── README.md                       # This README file
└── requirements.txt                # List of Python dependencies
```