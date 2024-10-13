import praw
import pandas as pd
import re
from datetime import datetime, timezone
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.stem import WordNetLemmatizer

# Download necessary NLTK data only if not already available
nltk_data = ['punkt', 'stopwords', 'wordnet']
for data in nltk_data:
    nltk.download(data, quiet=True)


def preprocess_post_text(text):
    """Preprocesses a Reddit post's text by cleaning, tokenizing, removing stopwords, and lemmatizing."""
    if not isinstance(text, str):
        return ""  # Handle non-string input

    # Convert text to lowercase
    text = text.lower()

    # Remove URLs
    text = re.sub(r'http\S+|www.\S+|https\S+', '', text)

    # Remove special characters, punctuation, and digits
    text = re.sub(r'[^a-zA-Z\s]', '', text)

    # Tokenize the text
    words = word_tokenize(text)

    # Remove stopwords
    stop_words = set(stopwords.words('english'))
    words = [word for word in words if word not in stop_words]

    # Lemmatize the words (convert to base form)
    lemmatizer = WordNetLemmatizer()
    words = [lemmatizer.lemmatize(word) for word in words]

    # Join the words back into a single string
    return ' '.join(words)


def initialize_reddit_client(client_id, client_secret, user_agent):
    """Initializes the Reddit client using PRAW."""
    try:
        reddit = praw.Reddit(
            client_id=client_id,
            client_secret=client_secret,
            user_agent=user_agent
        )
        return reddit
    except Exception as e:
        raise RuntimeError(f"Failed to initialize Reddit API: {e}")


def fetch_reddit_posts(reddit, tickers, subreddits, start_date=datetime(2020, 1, 1)):
    """Fetches Reddit posts matching specific queries from given subreddits."""
    posts_data = {ticker: [] for ticker in tickers.keys()}
    start_date = start_date.replace(tzinfo=None)
    try:
        for ticker, query in tickers.items():
            for subreddit in subreddits:
                subreddit_obj = reddit.subreddit(subreddit)
                for submission in subreddit_obj.search(query, limit=1000):
                    post_datetime = datetime.fromtimestamp(submission.created_utc, tz=None)

                    if post_datetime >= start_date:  # Filter posts from start_date onwards
                        post = {
                            'Title': preprocess_post_text(submission.title),
                            'Score': submission.score,
                            'Date': post_datetime,
                            'Comments': submission.num_comments,
                            'Body': preprocess_post_text(submission.selftext),
                            'Subreddit': subreddit
                        }
                        posts_data[ticker].append(post)
    except Exception as e:
        print(f"Error fetching Reddit posts: {e}")

    return posts_data


def save_posts_to_excel(posts_data, filename='reddit_data.xlsx'):
    """Saves posts data to an Excel file with each ticker's posts on a separate sheet."""
    try:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            for ticker, posts in posts_data.items():
                if posts:  # Only create a sheet if there are posts
                    df = pd.DataFrame(posts)
                    df.to_excel(writer, sheet_name=ticker, index=False)
        print(f"Data successfully saved to '{filename}'")
    except Exception as e:
        print(f"Error saving data to Excel: {e}")


def main():
    # Define Reddit API credentials
    client_id = 'LA3LKSRylCsCaQvtciFvlw'
    client_secret = 'Gi9EikLF4Ai0OKBQWLoVv_mJz7wWEw'
    user_agent = 'CapX/0.1 by Frosty-Farmer-397'

    # Initialize Reddit client
    reddit = initialize_reddit_client(client_id, client_secret, user_agent)

    # Define stock tickers and search queries
    tickers = {
        'GME': 'GameStop shares OR $GME OR GME stock',
        'AMC': 'AMC shares OR $AMC OR AMC stock',
        'TSLA': 'Tesla shares OR $TSLA OR Tesla stock',
        'AAPL': 'Apple shares OR $AAPL OR Apple stock'
    }
    subreddits = ['stocks', 'investing', 'wallstreetbets']

    # Fetch posts for all tickers and subreddits
    posts_data = fetch_reddit_posts(reddit, tickers, subreddits)

    # Save the fetched posts to an Excel file
    save_posts_to_excel(posts_data)


# Run the script only if executed directly
if __name__ == "__main__":
    main()
