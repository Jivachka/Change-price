import openpyxl
import logging
from datetime import datetime


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ExcelReader:
    def __init__(self, file_path, sheet_name):
        self.file_path = file_path
        self.wb = openpyxl.load_workbook(self.file_path)
        self.ws = self.wb[sheet_name]

    def get_value(self, row, column):
        return self.ws.cell(row=row, column=column).value

class ExchangeRateFinder:
    def __init__(self, excel_reader):
        self.excel_reader = excel_reader

    def get_rates_by_date(self, date):
        row = 2
        current_value = self.excel_reader.get_value(row, 2)
        current_date = None

        while current_value:
            if isinstance(current_value, str):
                current_date = current_value
            elif isinstance(current_value, datetime):
                current_date = str(current_value.strftime("%Y-%m-%d"))

            if current_date == date.strftime("%Y-%m-%d"):
                usd = self.excel_reader.get_value(row, 3)
                usd_deferred = self.excel_reader.get_value(row, 4)
                eur = self.excel_reader.get_value(row, 5)
                eur_deferred = self.excel_reader.get_value(row, 6)
                return {"usd": usd, "usd_deferred": usd_deferred, "eur": eur, "eur_deferred": eur_deferred}
            row += 1
            current_value = self.excel_reader.get_value(row, 2)

        return None

def main():
    file_path = "Pricess2.xlsx"
    sheet_name = "Kurs"
    today = datetime.now().date()
    formatted_today = today.strftime("%Y-%m-%d")

    excel_reader = ExcelReader(file_path, sheet_name)
    exchange_rate_finder = ExchangeRateFinder(excel_reader)
    rates = exchange_rate_finder.get_rates_by_date(today)

    if rates:
        logger.info(f"Exchange rates for {formatted_today}:")
        logger.info(f"USD: {rates['usd']}")
        logger.info(f"USD Deferred: {rates['usd_deferred']}")
        logger.info(f"EUR: {rates['eur']}")
        logger.info(f"EUR Deferred: {rates['eur_deferred']}")
    else:
        logger.warning(f"No exchange rates found for {formatted_today}.")

if __name__ == "__main__":
    main()
