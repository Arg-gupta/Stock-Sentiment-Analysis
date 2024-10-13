import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


# Load the Excel file
def load_excel(file_path):
    """Load an Excel file and return the ExcelFile object."""
    try:
        return pd.ExcelFile(file_path)
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None


def create_dashboard(stock_data, company_name, filename):
    """Create and save improved visualizations for the stock data."""
    # Set up the figure
    fig, axs = plt.subplots(2, 2, figsize=(16, 14), gridspec_kw={'hspace': 0.4, 'wspace': 0.3})
    fig.suptitle(f'{company_name} - Sentiment and Stock Analysis', fontsize=20, weight='bold', y=0.98)

    # 1. Dual Y-Axis: Close Price vs Compound Sentiment
    ax1 = axs[0, 0]
    ax2 = ax1.twinx()
    ax1.plot(stock_data['Date_y'], stock_data['Close'], color=sns.color_palette("coolwarm", 10)[9], label='Close Price')
    ax2.plot(stock_data['Date_y'], stock_data['Compound_Sentiment'], color=sns.color_palette("coolwarm", 10)[2], linestyle='--', label='Compound Sentiment')

    ax1.set_title('Price vs Compound Sentiment', fontsize=14)
    ax1.set_xlabel('Date', fontsize=12)
    ax1.set_ylabel('Close Price', color=sns.color_palette("coolwarm", 10)[9])
    ax2.set_ylabel('Compound Sentiment', color=sns.color_palette("coolwarm", 10)[2])
    ax1.tick_params(axis='x', rotation=45)
    ax1.legend(loc='upper left')
    ax2.legend(loc='upper right')

    # 2. Mentions Over Time
    ax3 = axs[0, 1]
    ax3.bar(stock_data['Date_y'], stock_data['Mentions'], color=sns.color_palette("coolwarm", 10)[3], width=5)
    ax3.set_title('Stock Mentions Over Time', fontsize=14)
    ax3.set_xlabel('Date', fontsize=12)
    ax3.set_ylabel('Mentions')
    ax3.tick_params(axis='x', rotation=45)
    max_mentions = stock_data['Mentions'].max()
    ax3.set_ylim(0, max_mentions * 1.2)

    # 3. Correlation Heatmap for Selected Metrics
    correlation_matrix = stock_data[['Compound_Sentiment', 'Weighted_Sentiment', 'Mentions', 'Close', 'Returns']].corr()
    sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', center=0, ax=axs[1, 0], cbar_kws={'shrink': 0.8})
    axs[1, 0].set_title('Correlation Heatmap', fontsize=14)

    # 4. Dual Y-Axis: Close Price vs Weighted Sentiment
    ax4 = axs[1, 1]
    ax5 = ax4.twinx()
    ax4.plot(stock_data['Date_y'], stock_data['Close'], color=sns.color_palette("coolwarm", 10)[9], label='Close Price')
    ax5.plot(stock_data['Date_y'], stock_data['Weighted_Sentiment'], color=sns.color_palette("coolwarm", 10)[1], linestyle='--', label='Weighted Sentiment')

    ax4.set_title('Price vs Weighted Sentiment', fontsize=14)
    ax4.set_xlabel('Date', fontsize=12)
    ax4.set_ylabel('Close Price', color=sns.color_palette("coolwarm", 10)[9])
    ax5.set_ylabel('Weighted Sentiment', color=sns.color_palette("coolwarm", 10)[1])
    ax4.tick_params(axis='x', rotation=45)
    ax4.legend(loc='upper left')
    ax5.legend(loc='upper right')

    # Adjust subplot spacing
    plt.subplots_adjust(left=0.11, right=0.95, top=0.93, bottom=0.15, hspace=0.45, wspace=0.35)

    # Save the figure to a file
    plt.savefig(filename, dpi=300)  # Save with high resolution (300 DPI)
    plt.close()  # Close the figure to free memory


def main():
    # Load the Excel file
    file_path = "correlation_and_merged_data.xlsx"  # Adjust the path if necessary
    xls = load_excel(file_path)
    if xls is None:
        return

    # Define the stock tickers and their company names
    companies = {
        'AMC': 'AMC Entertainment Holdings',
        'TSLA': 'Tesla, Inc.',
        'AAPL': 'Apple Inc.',
        'GME': 'GameStop Corp.'
    }

    # Loop through each company and generate its dashboard
    for ticker, company_name in companies.items():
        print(f"Creating and saving dashboard for {company_name}...")

        # Load the stock's merged data from the Excel file
        stock_data = pd.read_excel(xls, sheet_name=f'{ticker}_Merged_Data')

        # Convert Date to datetime
        stock_data['Date_y'] = pd.to_datetime(stock_data['Date_y'])

        # Define the filename for saving the dashboard
        filename = f"{ticker}_dashboard.png"

        # Create and save the dashboard for the stock
        create_dashboard(stock_data, company_name, filename)


if __name__ == "__main__":
    main()
