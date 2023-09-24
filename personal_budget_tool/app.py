import os
from datetime import datetime

from config import InputConfig
from data import DataBuilder
from excel import BudgetApp


def main():
    min_date = datetime(2023, 1, 1, 0, 0, 0)
    max_date = datetime(2024, 12, 31, 0, 0, 0)

    inputConfig = InputConfig()

    data = DataBuilder(min_date=min_date, max_date=max_date)
    df = data.get_export_data()

    excelApp = BudgetApp(min_date=min_date, max_date=max_date, df=df)
    excelApp.save_and_close()


if __name__ == "__main__":
    main()
