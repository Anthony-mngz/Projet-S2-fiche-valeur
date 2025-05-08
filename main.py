import logging
import yaml
import os
from repository import get_data, fetch_esg_data, fetch_news, fetch_calendar, fetch_forecasts
from model import plot_stock_chart, calculate_ratios
from view import update_ppt, convert_ppt_to_pdf, display_pdf, send_report_via_email

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)



def main():
    logger.info("Démarrage du programme")

    # Ticker utilisateur
    ticker = input("Enter the Yahoo Finance stock ticker: ")
    ticker = ticker.upper()

    # Base directory
    base_dir = os.getcwd()
    # 0) Charge le fichier de config YAML
    config_file = os.path.join(base_dir, "config.yaml")
    with open(config_file, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f)

    template_path = os.path.join(base_dir, cfg["template_path"])
    output_ppt_path = os.path.join(base_dir, cfg["output_ppt"].format(ticker=ticker))
    output_pdf_path = os.path.join(base_dir, cfg["output_pdf"].format(ticker=ticker))


    logger.info("1) Récupération des données financières et historique")
    # 1) Données financières et historique
    df_combined, stock, financials, balance_sheet, info, history, description = get_data(ticker)
    logger.info("→ Données financières récupérées")

    logger.info("2) Génération du graphique boursier")
    # 2) Génération du graphique et ratios sur l'historique
    chart_path = plot_stock_chart(history, ticker)
    logger.info(f"→ Chart généré : {chart_path}")
    logger.info("3) Calcul des ratios")
    ratios = calculate_ratios(history)
    logger.info(f"→ Ratios calculés : {ratios}")

    logger.info("4) Fetch ESG data")
    # 3) Extensions: ESG, News, Calendar, Forecasts
    esg_df = fetch_esg_data(ticker)
    logger.info(f"→ ESG récupéré ({len(esg_df)} lignes)")

    logger.info("5) Fetch News")
    news_items = fetch_news(ticker)
    logger.info(f"→ News récupérées ({len(news_items)} items)")

    logger.info("6) Fetch Calendar")
    calendar_df = fetch_calendar(ticker)
    logger.info(f"→ Calendar récupéré ({len(calendar_df)} entrées)")

    logger.info("7) Fetch Forecasts")
    forecasts = fetch_forecasts(ticker)
    logger.info(f"→ Prévisions récupérées : {forecasts}")

    logger.info("8) Mise à jour du template PowerPoint")
    # 4) Mise à jour PPT
    update_ppt(
        template_path,
        output_ppt_path,
        df_combined,
        description,
        info,
        ratios,
        chart_path,
        esg_df,
        news_items,
        calendar_df,
        forecasts
    )
    logger.info(f"→ Présentation mise à jour enregistrée dans {output_ppt_path}")

    logger.info("9) Conversion du PPT en PDF")
    # 5) Conversion en PDF
    convert_ppt_to_pdf(output_ppt_path, output_pdf_path)
    logger.info(f"→ PDF généré : {output_pdf_path}")

    # 6) Affichage du PDF
    display_pdf(output_pdf_path)
    logger.info("Programme terminé avec succès")

    # Envoi automatique du rapport par email
    try:
        logger.info("Envoi du rapport par email")
        send_report_via_email(output_pdf_path, cfg, ticker)
    except Exception as e:
        logger.error(f"Erreur lors de l'envoi de l'email : {e}")

if __name__ == "__main__":
    main()
