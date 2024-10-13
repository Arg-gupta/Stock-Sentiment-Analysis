import pandas as pd
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer

# Initialize VADER sentiment analyzer
analyzer = SentimentIntensityAnalyzer()


def analyze_sentiment(text):
    """Perform sentiment analysis on the given text and return the compound score."""
    if isinstance(text, str):
        return analyzer.polarity_scores(text)['compound']
    return 0.0


def classify_sentiment(compound_score):
    """Classify sentiment based on the compound score."""
    if compound_score >= 0.05:
        return 'Positive'
    elif compound_score <= -0.05:
        return 'Negative'
    return 'Neutral'


def process_stock_data(file_path, output_file):
    """Load stock data from an Excel file, perform sentiment analysis, and save results."""
    try:
        # Load the Excel file
        excel_data = pd.ExcelFile(file_path)
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return

    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Process each sheet (each stock)
            for sheet_name in excel_data.sheet_names:
                # Load the data for the specific stock
                df = pd.read_excel(excel_data, sheet_name=sheet_name)

                # Apply sentiment analysis to Title and Body
                df['Title_Sentiment'] = df['Title'].apply(analyze_sentiment)
                df['Body_Sentiment'] = df['Body'].apply(analyze_sentiment)

                # Calculate compound sentiment and classify it
                df['Compound_Sentiment'] = (df['Title_Sentiment'] + df['Body_Sentiment']) / 2
                df['Sentiment_Category'] = df['Compound_Sentiment'].apply(classify_sentiment)

                # Calculate engagement metrics
                df['Weighted_Sentiment'] = (df['Score'] + df['Comments']) * df['Compound_Sentiment']
                df['Engagement_Metrics'] = df[['Score', 'Comments']].mean(axis=1)

                # Convert Date column to datetime format
                df['Date'] = pd.to_datetime(df['Date'])

                # Group by week and calculate averages
                df_grouped = df.resample('W', on='Date').agg({
                    'Compound_Sentiment': 'mean',
                    'Weighted_Sentiment': 'mean',
                    'Title': 'count',
                    'Engagement_Metrics': 'mean'
                }).rename(columns={'Title': 'Mentions'})

                # Remove rows with 0 mentions
                df_grouped = df_grouped[df_grouped['Mentions'] > 0]

                # Write the results to the corresponding sheet in the output Excel file
                df_grouped.to_excel(writer, sheet_name=sheet_name)

        print(f"Sentiment analysis and feature extraction completed. Results saved to {output_file}")

    except Exception as e:
        print(f"Error processing stock data: {e}")


def main(input_file, output_file):
    """Main function to run sentiment analysis on stock data."""
    process_stock_data(input_file, output_file)


# Define input and output file paths
if __name__ == "__main__":
    input_file_path = 'reddit_data.xlsx'
    output_file_path = 'sentiment analysis.xlsx'
    main(input_file_path, output_file_path)
