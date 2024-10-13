# 1. Executive Summary

**Purpose:** This analysis explores the relationship between social media sentiment and stock price movements for four companies: Apple (AAPL), AMC Entertainment (AMC), Tesla (TSLA), and GameStop (GME) from January 2020 to 2024. The goal is to assess the impact of online sentiment on stock prices and returns.

**Sentiment Measures Used:**

- **Compound Sentiment:** A holistic measure of sentiment strength, ranging from -1 (very negative) to +1 (very positive).
- **Weighted Sentiment:** A measure that adjusts sentiment scores based on the engagement level (e.g., upvotes, comments), giving more weight to influential social media posts.

### Key Findings:

- **AMC:** Weighted sentiment shows a significant correlation with stock price, and mentions correlate strongly with returns.
- **Tesla:** Mentions show a weak correlation with returns, while sentiment does not show a strong correlation with price movements.
- **Apple:** Mentions have a slightly negative correlation with stock price, while sentiment indicators are not strongly correlated.
- **GameStop:** Mentions show a strong positive correlation with both stock price and returns, suggesting strong price sensitivity to social media discussions.

# 2. Detailed Findings

## A. Sentiment-Price Correlation Analysis

### AMC Entertainment Holdings (AMC)
- **Compound Sentiment vs. Close Price:** Correlation = 0.18, p = 0.04 (significant)
- **Weighted Sentiment vs. Close Price:** Correlation = 0.22, p = 0.009 (significant)
- **Mentions vs. Close Price:** Correlation = 0.07, p = 0.41 (not significant)
- **Mentions vs. Returns:** Correlation = 0.56, p = 0.000000000002 (highly significant)

**Key Insight:** AMC shows a statistically significant correlation between Weighted Sentiment and stock price, suggesting that more impactful social media posts (those with higher engagement) align closely with price movements. The correlation between mentions and returns is extremely strong, indicating that spikes in discussion about AMC correspond to increased volatility and price movement.

### Tesla (TSLA)
- **Compound Sentiment vs. Close Price:** Correlation = 0.06, p = 0.45 (not significant)
- **Weighted Sentiment vs. Close Price:** Correlation = 0.004, p = 0.95 (not significant)
- **Mentions vs. Close Price:** Correlation = 0.03, p = 0.69 (not significant)
- **Mentions vs. Returns:** Correlation = 0.17, p = 0.023 (significant)

**Key Insight:** Tesla shows little correlation between sentiment and stock price, with p-values indicating no statistically significant relationship. However, the mentions show a weak but statistically significant relationship with returns, suggesting that social discussions may contribute slightly to volatility in Tesla's stock.

### Apple Inc. (AAPL)
- **Compound Sentiment vs. Close Price:** Correlation = 0.15, p = 0.04 (significant)
- **Weighted Sentiment vs. Close Price:** Correlation = -0.08, p = 0.28 (not significant)
- **Mentions vs. Close Price:** Correlation = -0.19, p = 0.01 (significant)
- **Mentions vs. Returns:** Correlation = 0.04, p = 0.63 (not significant)

**Key Insight:** Apple’s mentions exhibit a significant negative correlation with stock price, suggesting that high levels of social media discussion could signal a potential decrease in price. Sentiment metrics show a weaker correlation, with compound sentiment barely reaching statistical significance.

### GameStop (GME)
- **Compound Sentiment vs. Close Price:** Correlation = 0.01, p = 0.88 (not significant)
- **Weighted Sentiment vs. Close Price:** Correlation = 0.16, p = 0.06 (marginal significance)
- **Mentions vs. Close Price:** Correlation = 0.23, p = 0.008 (significant)
- **Mentions vs. Returns:** Correlation = 0.60, p = 0.00000000000002 (highly significant)

**Key Insight:** GameStop shows a strong correlation between mentions and both stock price and returns, indicating that discussions about GameStop in social media directly influence its price movements. The weighted sentiment shows a borderline significant relationship with the stock price, hinting at a possible connection between influential posts and stock movements.

## B. Stock Mentions and Price Movements

- **AMC:** There is no significant correlation between mentions and close price. However, the correlation between mentions and returns is strong, implying that spikes in social media mentions often precede increased volatility and price fluctuations.
- **Tesla:** Mentions show a weak but significant correlation with returns, suggesting that spikes in discussions about Tesla may lead to slight increases in price volatility. However, no strong correlation was found with the closing price.
- **Apple:** Mentions have a slightly negative correlation with stock price, hinting that more discussions could be a precursor to price drops.
- **GameStop:** GameStop demonstrates a strong positive correlation between mentions and stock price. Spikes in discussions are highly correlated with stock price increases and returns, particularly during major volatility events such as the 2021 short squeeze.

# 3. Key Insights and Recommendations

## A. Insights

- **AMC:** Traders could monitor Weighted Sentiment and mentions as leading indicators for price movements and returns. Spikes in mentions may signal upcoming price volatility.
- **Tesla:** Despite no strong sentiment correlation, mentions could be used as a volatility signal for Tesla, although the effect size is small.
- **Apple:** High levels of mentions could be a negative signal for Apple, as it correlates with slight price decreases.
- **GameStop:** Mentions are a clear predictor of price movements and returns. Investors should pay close attention to spikes in social media discussions to capitalize on price swings.

## B. Trading/Investment Recommendations

- **AMC:** Traders could integrate Weighted Sentiment into their strategies as it correlates significantly with price. High sentiment and mentions could be a signal for buying before an upward price movement.
- **Tesla:** While sentiment indicators are weak, mentions could still be a signal for volatility, so short-term traders may want to capitalize on spikes in discussion.
- **Apple:** Negative correlation between mentions and price suggests caution when Apple is being heavily discussed on social media.
- **GameStop:** Mentions offer clear buy/sell signals, with a high correlation to both price and returns. Traders should consider buying when mentions rise and selling when mentions start to fall off.


# 4. Conclusion

This analysis demonstrates the significant role of social media sentiment in stock price movements, especially for volatile stocks like AMC and GameStop. While sentiment analysis may not be a reliable indicator for all stocks, it holds strong potential as part of a broader trading strategy for certain stocks. Moving forward, traders and investors can benefit from integrating sentiment data with traditional financial analysis to make more informed decisions.
