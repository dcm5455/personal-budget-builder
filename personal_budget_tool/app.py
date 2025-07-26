import os
from datetime import datetime

from config import InputConfig
from data import DataBuilder
from excel import BudgetApp


def main():
    min_date = datetime(2025, 1, 1, 0, 0, 0)
    max_date = datetime(2030, 12, 31, 0, 0, 0)

    input_config = InputConfig()
    input_config.prompt()

    data_builder = DataBuilder(min_date=min_date, max_date=max_date)
    data_builder.build_data_model()
    df = data_builder.get_df()

    excel_app = BudgetApp(min_date=min_date, max_date=max_date, df=df)
    excel_app.save_and_close()


if __name__ == "__main__":
    main()
