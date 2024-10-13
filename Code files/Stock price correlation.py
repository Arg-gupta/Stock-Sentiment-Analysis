import pandas as pd
import yfinance as yf
from scipy.stats import pearsonr


def fetch_stock_data(tickers, start_date, end_date):
    """Fetch stock price data for the specified tickers within the given date range."""
    try:
        stock_data = yf.download(tickers, start=start_date, end=end_date, group_by='ticker')
        return stock_data
    except Exception as e:
        print(f"Error fetching stock data: {e}")
        return None


def adjust_sentiment_dates(sentiment_data):
    """Shift weekend sentiment dates to the previous Friday."""
    sentiment_data['Date'] = pd.to_datetime(sentiment_data['Date'])
    sentiment_data['Adjusted_Date'] = sentiment_data['Date'].apply(
        lambda date: date - pd.DateOffset(days=(date.weekday() - 4) % 7) if date.weekday() >= 5 else date
    )
    return sentiment_data


def merge_sentiment_with_stock(sentiment_data, stock_data, ticker):
    """Merge sentiment data with stock price data for a specific ticker."""
    stock_daily = stock_data[ticker][['Close']].reset_index()
    stock_daily['Date'] = stock_daily['Date'].dt.tz_localize(None).dt.date
    sentiment_data = adjust_sentiment_dates(sentiment_data)
    sentiment_data['Adjusted_Date'] = sentiment_data['Adjusted_Date'].dt.date

    merged_data = pd.merge(sentiment_data, stock_daily, left_on='Adjusted_Date', right_on='Date', how='inner')
    merged_data['Returns'] = merged_data['Close'].pct_change() * 100
    return merged_data


def calculate_correlation(data, sentiment_column, stock_column):
    """Calculate the Pearson correlation between sentiment and stock price movements."""
    valid_data = data[[sentiment_column, stock_column]].dropna()
    if not valid_data.empty:
        correlation, p_value = pearsonr(valid_data[sentiment_column], valid_data[stock_column])
        return correlation, p_value
    return None, None


def process_sentiment_analysis(sentiment_file_path, output_file, tickers, sentiment_columns, start_date, end_date):
    """Process sentiment analysis and correlation calculations."""
    # Load the sentiment analysis data
    try:
        xls = pd.ExcelFile(sentiment_file_path)
    except Exception as e:
        print(f"Error loading sentiment data: {e}")
        return

    # Fetch stock price data
    stock_data = fetch_stock_data(tickers, start_date, end_date)
    if stock_data is None:
        return

    correlation_results = []

    # Create an Excel writer to save output
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for ticker in tickers:
            print(f"\nProcessing {ticker}...\n")

            # Load the sentiment data for the stock
            try:
                sentiment_data = pd.read_excel(xls, sheet_name=ticker)
            except Exception as e:
                print(f"Error loading sentiment data for {ticker}: {e}")
                continue

            # Merge sentiment data with daily stock price data
            merged_data = merge_sentiment_with_stock(sentiment_data, stock_data, ticker)

            # Save the merged data for the stock in a separate sheet
            merged_data.to_excel(writer, sheet_name=f'{ticker}_Merged_Data', index=False)

            # Perform correlation analysis on each sentiment component
            for sentiment_column in sentiment_columns:
                correlation, p_value = calculate_correlation(merged_data, sentiment_column, 'Close')
                if correlation is not None:
                    correlation_results.append({
                        'Stock': ticker,
                        'Sentiment_Component': sentiment_column,
                        'Stock_Price_Metric': 'Close',
                        'Correlation': correlation,
                        'P-Value': p_value
                    })

            # Correlation between Mentions and stock price volatility (Returns)
            correlation, p_value = calculate_correlation(merged_data, 'Mentions', 'Returns')
            if correlation is not None:
                correlation_results.append({
                    'Stock': ticker,
                    'Sentiment_Component': 'Mentions',
                    'Stock_Price_Metric': 'Returns',
                    'Correlation': correlation,
                    'P-Value': p_value
                })

        # Convert correlation results to a DataFrame and save it in a separate sheet
        correlation_results_df = pd.DataFrame(correlation_results)
        correlation_results_df.to_excel(writer, sheet_name='Correlation_Results', index=False)

    print(f"Analysis complete! Data saved to {output_file}")

if __name__ == "__main__":
    # Define file paths and parameters
    sentiment_file_path = 'sentiment analysis.xlsx'
    output_file = 'correlation_and_merged_data.xlsx'
    tickers = ['AMC', 'TSLA', 'AAPL', 'GME']
    sentiment_columns = ['Compound_Sentiment', 'Weighted_Sentiment', 'Mentions']
    start_date = '2020-01-01'
    end_date = '2024-10-10'

    # Run the sentiment analysis processing
    process_sentiment_analysis(sentiment_file_path, output_file, tickers, sentiment_columns, start_date, end_date)
