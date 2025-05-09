import logging
import yfinance as yf
import pandas as pd
from datetime import datetime

logger = logging.getLogger(__name__)


def get_data(ticker: str):
    stock = yf.Ticker(ticker)
    info = stock.info
    info["dividendYield"] = (
        info.get("dividendYield")
        if isinstance(info.get("dividendYield"), (int, float))
        else "NA"
    )
    info["marketCap"] = round(info.get("marketCap", 0) / 1000000)
    info["beta"] = round(info.get("beta", 0), 2)
    description = info.get("longBusinessSummary", "")

    financials = stock.financials
    balance_sheet = stock.balance_sheet
    history = stock.history(period="max")

    # Rename columns dynamically
    fin_cols = financials.shape[1]
    bs_cols = balance_sheet.shape[1]
    financials.columns = [f"N-{i}" if i > 0 else "N" for i in range(fin_cols)]
    balance_sheet.columns = [f"N-{i}" if i > 0 else "N" for i in range(bs_cols)]

    df_combined = pd.concat([financials, balance_sheet], axis=0)
    return df_combined, stock, financials, balance_sheet, info, history, description


def fetch_esg_data(ticker: str) -> pd.DataFrame:
    """Récupère les scores ESG et controverses via yfinance."""
    try:
        esg_df = yf.Ticker(ticker).sustainability
        logger.info(f"Fetched ESG data for {ticker}")
        return esg_df
    except Exception as e:
        logger.error(f"Error fetching ESG data for {ticker}: {e}")
        return pd.DataFrame()


def fetch_news(ticker: str) -> list[dict]:
    """Récupère les dernières actualités."""
    try:
        items = yf.Ticker(ticker).news
        for it in items:
            if "providerPublishTime" in it:
                it["datetime"] = datetime.fromtimestamp(it["providerPublishTime"])
        logger.info(f"Fetched {len(items)} news items for {ticker}")
        return items
    except Exception as e:
        logger.error(f"Error fetching news for {ticker}: {e}")
        return []


def fetch_calendar(ticker: str) -> dict:
    """Récupère le calendrier économique (earnings, dividendes...)."""
    try:
        cal = yf.Ticker(ticker).calendar
        logger.info(f"Fetched calendar for {ticker}")
        return cal
    except Exception as e:
        logger.error(f"Error fetching calendar for {ticker}: {e}")
        return {}


def fetch_forecasts(ticker: str) -> dict:
    """Récupère les prévisions de cours et recommandations analystes."""
    try:
        info = yf.Ticker(ticker).info
        forecasts = {
            "target_mean_price": info.get("targetMeanPrice"),
            "target_low_price": info.get("targetLowPrice"),
            "target_high_price": info.get("targetHighPrice"),
            "recommendation_mean": info.get("recommendationMean"),
            "recommendation_key": info.get("recommendationKey"),
        }
        logger.info(f"Fetched forecasts for {ticker}")
        return forecasts
    except Exception as e:
        logger.error(f"Error fetching forecasts for {ticker}: {e}")
        return {}
