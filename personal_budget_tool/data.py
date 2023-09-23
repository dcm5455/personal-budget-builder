import os
import pandas as pd
from datetime import datetime
from typing import Union, Tuple

from constants import Models
from utils import read_dataframe_input, filter_df_between


class DataBuilder:
    def __init__(self, min_date: datetime, max_date: datetime):
        self.min_date = min_date
        self.max_date = max_date
        self.dates = None
        self.items = None
        self.date_items = None
        self._get_dates()
        self._get_date_list()
        self._get_items()
        self._get_date_items()
        self._audit_date_frequencies()
        self._calc_budget_amounts()

    def _get_dates(self):
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        dates = read_dataframe_input(**Models.BudgetDate)
        dates = filter_df_between(
            dates, "date", (self.min_date, self.max_date), Models.BudgetDate.IndexColumn
        )
        self.dates = dates

    def _get_date_list(self):
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        self.date_list = [
            f"{x['year']}-{x['month_number']}-{x['day_number']}"
            for _, x in self.dates.iterrows()
        ]

    def _get_items(self):
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        self.items = read_dataframe_input(**Models.BudgetItem)

    def _get_date_items(self):
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        df = self.dates.merge(self.items, how="cross")
        df = df.sort_values(
            by=["budget_item_id", "date_id"], ascending=[True, True]
        ).reset_index(drop=True)
        df["budget_item_amount"] = 0.00
        self.date_items = df

    def _is_even_week(self, week_number: int) -> bool:
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        return (week_number % 2) == 0

    def _validate_date(self, year: int, month: int, day: int) -> bool:
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        return f"{year}-{month}-{day}" in self.date_list

    def _get_max_day_in_yearmonth(self, year: int, month: int) -> int:
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        return (
            self.date_items.groupby(["year", "month_number"])["day_number"]
            .max()
            .reset_index()
            .query(f"year == {year} & month_number == {month}")["day_number"]
        )

    def _get_date_attribs(self, date: datetime) -> Tuple[datetime, bool]:
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        iter_date = self.dates[self.dates["date"] == date].iloc[0]
        start_day_of_week = iter_date["day_of_week"]
        is_even_week = self._is_even_week(iter_date["week_number"])
        return (start_day_of_week, is_even_week)

    def _audit_date_frequencies(self):
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        for index, row in self.date_items.iterrows():
            if pd.isnull(row["frequency_day"]):
                continue
            if self._validate_date(
                row["year"], row["month_number"], int(row["frequency_day"])
            ):
                continue
            else:
                new_day = self._get_max_day_in_yearmonth(
                    row["year"], row["month_number"]
                )
                self.date_items.at[index, "frequency_day"] = new_day
            print(
                f"Updated {row['item_name']} for {row['month_year']} from {row['frequency_day']} to {new_day}.."
            )

    def _calc_multiplier(self, row: dict) -> float:
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        return (1.00 if row["item_type"] == "Income" else -1.00) * (
            row["seasonality_multiplier"] if row["is_seasonality"] else 1.00
        )

    def _validate_record(self, row: dict) -> bool:
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        if not row["is_active"]:
            return False

        if not pd.isnull(row["end_date"]) and row["date"] > row["end_date"]:
            ## print(f"Skipping end date.. {row['endDate']}, {row['date']}")
            return False

        if not pd.isnull(row["start_date"]) and row["date"] < row["start_date"]:
            ## print(f"Skipping start date.. {row['startDate']}, {row['date']}")
            return False

        return True

    def _calculate_budget_amount(self, row: dict) -> float:
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        if not self._validate_record(row):
            return 0.00

        ## Daily -> Return Budgeted Amount
        if row["frequency_type"] == "Daily":
            return row["item_amount"] * self._calc_multiplier(row)

        ## Weekly ->
        elif row["frequency_type"] == "Weekly":
            ## If Limited to a Day, Ignore other Days
            if (
                not pd.isnull(row["frequency_day"])
                and row["day_of_week"] != row["frequency_day"]
            ):
                return 0.00
            ## If same day as limited, or no limited day (and day_number == 1)
            elif (row["day_of_week"] == row["frequency_day"]) or (
                pd.isnull(row["frequency_day"]) and row["day_of_week"] == 1
            ):
                return row["item_amount"] * self._calc_multiplier(row)
            else:
                return 0.00

        ##Bi-Weekly ->
        elif row["frequency_type"] == "Bi-Weekly":
            ##If no start-date, raise exception
            if pd.isnull(row["start_date"]):
                raise ValueError(
                    "Need startDate for bi-weekly expense to determine alternating weeks."
                )
            ##If start date, validate start day of week & even/odd week increments
            else:
                start_day_of_week, is_even_week = self._get_date_attribs(
                    row["start_date"]
                )
                if (
                    self._is_even_week(row["week_number"]) == is_even_week
                    and row["day_of_week"] == start_day_of_week
                ):
                    return row["item_amount"] * self._calc_multiplier(row)
                else:
                    return 0.00

        ##Monthly ->
        elif row["frequency_type"] == "Monthly":
            ##If limited to day of month
            if not pd.isnull(row["frequency_day"]):
                ##If not same day, return 0
                if row["day_number"] != row["frequency_day"]:
                    return 0.00
                ##Otherwise, calc
                else:
                    return row["item_amount"] * self._calc_multiplier(row)

            ##If not limited to day of month & is first of month, return
            elif pd.isnull(row["frequency_day"]) and row["frequency_day"] == 1:
                return row["item_amount"] * self._calc_multiplier(row)
            ##Otherwise 0
            else:
                return 0.00

        ##Annual ->
        elif row["frequency_type"] == "Annual":
            if pd.isnull(row["frequency_date"]):
                raise ValueError("Issue with frequency_date")
            elif (
                not pd.isnull(row["frequency_date"])
                and row["date"] == row["frequency_date"]
            ):
                return row["item_amount"] * self._calc_multiplier(row)
            else:
                return 0.00

        ##One-Time ->
        elif row["frequency_type"] == "One-Time":
            if pd.isnull(row["frequency_date"]):
                raise ValueError("Issue with frequency_date")
            elif row["date"] == row["frequency_date"]:
                return row["item_amount"] * self._calc_multiplier(row)
            else:
                return 0.00

        print(row)
        print("Issue -- did not make it through an IF")
        return 0.00

    def _calc_budget_amounts(self):
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        self.date_items["budget_item_amount"] = self.date_items.apply(
            lambda x: self._calculate_budget_amount(x), axis=1
        )

    def get_export_data(self) -> pd.DataFrame:
        """One-line description of function

        Multi-line expanded description of function

        Args:
            Arg1: ArgType
                Arg1 Description
            Arg2: ArgType
                Arg2 Description
            Arg3: ArgType
                Arg3 Description
            Arg4: ArgType
                Arg4 Description

        Returns:
            return: val

        """
        df = self.date_items.copy()
        df = df[Models.ExportData.Columns]
        return df
