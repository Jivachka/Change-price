import openpyxl
import logging
from datetime import datetime


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ExcelHandler:
    def __init__(self, file_path, sheet_name):
        self.file_path = file_path
        self.wb = openpyxl.load_workbook(self.file_path)
        self.ws = self.wb[sheet_name]

    def get_value(self, row, column):
        return self.ws.cell(row=row, column=column).value

    def set_value(self, row, column, value):
        self.ws.cell(row=row, column=column).value = value
        self.wb.save(self.file_path)

class ExchangeRateFinder:
    def __init__(self, excel_handler):
        self.excel_handler = excel_handler

    def get_rates_by_date(self, date):
        row = 2
        current_value = self.excel_handler.get_value(row, 2)
        current_date = None

        while current_value:
            if isinstance(current_value, str):
                current_date = current_value
            elif isinstance(current_value, datetime):
                current_date = str(current_value.strftime("%Y-%m-%d"))

            if current_date == date.strftime("%Y-%m-%d"):
                usd = self.excel_handler.get_value(row, 3)
                usd_deferred = self.excel_handler.get_value(row, 4)
                eur = self.excel_handler.get_value(row, 5)
                eur_deferred = self.excel_handler.get_value(row, 6)
                return {"usd": usd, "usd_deferred": usd_deferred, "eur": eur, "eur_deferred": eur_deferred}
            row += 1
            current_value = self.excel_handler.get_value(row, 2)

        return None

class Logger:
    @staticmethod
    def log_rates(formatted_today, rates):
        if rates:
            logger.info(f"Exchange rates for {formatted_today}:")
            logger.info(f"USD: {rates['usd']}")
            logger.info(f"USD Deferred: {rates['usd_deferred']}")
            logger.info(f"EUR: {rates['eur']}")
            logger.info(f"EUR Deferred: {rates['eur_deferred']}")
        else:
            logger.warning(f"No exchange rates found for {formatted_today}.")

class ExchangeRateRecorder:
    def __init__(self, excel_handler):
        self.excel_handler = excel_handler

    def record_rates(self, formatted_today, rates):
        if rates:
            self.excel_handler.set_value(2, 2, formatted_today)
            self.excel_handler.set_value(3, 4, rates['usd'])
            self.excel_handler.set_value(3, 5, rates['usd_deferred'])
            self.excel_handler.set_value(4, 4, rates['eur'])
            self.excel_handler.set_value(4, 5, rates['eur_deferred'])

def main():
    file_path = "Pricess2.xlsx"
    sheet_name = "Kurs"
    today = datetime.now().date()
    formatted_today = today.strftime("%Y-%m-%d")

    excel_handler = ExcelHandler(file_path, sheet_name)
    exchange_rate_finder = ExchangeRateFinder(excel_handler)
    rates = exchange_rate_finder.get_rates_by_date(today)

    # Логгирование курсов валют
    Logger.log_rates(formatted_today, rates)

    # Запись курсов валют на новый лист
    if rates:
        excel_handler_2 = ExcelHandler(file_path, "Price2")
        exchange_rate_recorder = ExchangeRateRecorder(excel_handler_2)
        exchange_rate_recorder.record_rates(formatted_today, rates)

if __name__ == "__main__":
    main()

