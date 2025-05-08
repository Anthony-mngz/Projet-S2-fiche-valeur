import pandas as pd
import matplotlib.pyplot as plt


def plot_stock_chart(history, stock):
    """Generate and save a stock chart for the given ticker."""
    try:
        data = history
        plt.figure(figsize=(16, 6))  # Slightly larger and proportional figure size
        plt.plot(data["Close"], label=f"{stock} Stock Price")
        plt.title(f"{stock} Stock Price (Maximum Historical Data)", fontsize=18)
        plt.xlabel("Date", fontsize=14)
        plt.ylabel("Price", fontsize=14)
        plt.legend(fontsize=14)
        plt.xticks(fontsize=14)
        plt.yticks(fontsize=14)
        chart_path = f"{stock}_chart.png"
        plt.savefig(chart_path)
        plt.close()
        return chart_path
    except Exception as e:
        print(f"Error generating stock chart for {stock}: {e}")
        return None

def calculate_cagr(history, years):
    """
    Calculate Compound Annual Growth Rate (CAGR) over given years.
    :param history: Pandas DataFrame contenant les données historiques (avec une colonne "Close").
    :param years: Nombre d'années pour le calcul du CAGR.
    :return: CAGR en pourcentage.
    """
    try:
        if "Close" not in history.columns:
            raise ValueError("La colonne 'Close' est absente des données historiques.")

        if not isinstance(history.index, pd.DatetimeIndex):
            history.index = pd.to_datetime(history.index)

        annual_prices = history["Close"].resample('YE').last()

        if len(annual_prices) < years + 1:
            raise ValueError(f"Pas assez de données pour calculer le CAGR sur {years} années.")

        # Calcul du CAGR
        start_price = annual_prices.iloc[-(years + 2)]
        end_price = annual_prices.iloc[-2]

        if start_price <= 0 or end_price <= 0:
            raise ValueError("Les prix de début ou de fin sont invalides pour le calcul du CAGR.")

        cagr = ((end_price / start_price) ** (1 / years)) - 1
        return round(cagr * 100, 2)
    except Exception as e:
        print(f"Erreur lors du calcul du CAGR : {e}")
        return None

def calculate_ratios(history: pd.DataFrame) -> dict:
    """Calculate key financial ratios: overall, 5-year, 3-year CAGR."""
    try:
        years_data = len(history.resample('YE').last()) - 2
        overall = calculate_cagr(history, years_data)
        five_y = calculate_cagr(history, 4)
        three_y = calculate_cagr(history, 2)

        return {
            "overall": f"{overall:.2f}%",
            "5y": f"{five_y:.2f}%",
            "3y": f"{three_y:.2f}%"
        }
    except Exception as e:
        print(f"Error calculating financial ratios: {e}")
        return {}
