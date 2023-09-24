import os
import pandas as pd
import warnings
from datetime import datetime
from typing import Union, Tuple

from constants import Models
from utils import read_dataframe_input, filter_df_between

warnings.simplefilter("ignore")


class DataBuilder:
    """Class to interact with data model for budget

    Attributes
    ----------
        min_date : datetime
            Start date for the budget
        max_date: datetime
            End date for the budget

    Methods
    -------
        build_data_model():
            Acquires data from multiple sources, consolidates
        get_df():
            Returns df filtered for model

    """

    def __init__(self, min_date: datetime, max_date: datetime):
        """Initializes DataBuilder class

        Parameters
        ----------
            min_date : datetime
                Start date for the budget
            max_date: datetime
                End date for the budget
        """
        self.min_date = min_date
        self.max_date = max_date

    def _get_dates(self):
        """Reads table of dates from Inputs file

        Reads table from excel file into pd.DataFrame format. Sets self.dates
        """
        dates = read_dataframe_input(**Models.BudgetDate)
        dates = filter_df_between(
            dates, "date", (self.min_date, self.max_date), Models.BudgetDate.IndexColumn
        )
        self.dates = dates

    def _get_date_list(self):
        """One-line description of function

        Multi-line expanded description of function
        """
        self.date_list = [
            f"{x['year']}-{x['month_number']}-{x['day_number']}"
            for _, x in self.dates.iterrows()
        ]

    def _get_items(self):
        """Reads table of items from Inputs file

        Reads table from excel file into pd.DataFrame format. Sets self.items
        """
        self.items = read_dataframe_input(**Models.BudgetItem)

    def _get_date_items(self):
        """Creates cartesian product of dates and items dataframes

        Combines two existing dataframes as a cross join (or cartesian product).
        Sorts by budget_item, then date ascending. Adds a 0-value field for budget amt,
        which we fill in next.
        """
        df = self.dates.merge(self.items, how="cross")
        df = df.sort_values(
            by=["budget_item_id", "date_id"], ascending=[True, True]
        ).reset_index(drop=True)
        df["budget_item_amount"] = 0.00
        self.date_items = df

    def _is_even_week(self, week_number: int) -> bool:
        """Boolean check if week_number is even

        Simple math function to check whether week is even or not
        This is used for frequencies such as 'bi-weekly' so we can understand
            when to budget for something correctly

        Parameters
        ----------
            week_number : int

        Returns
        -------
            bool
                is_even_week
        """
        return (week_number % 2) == 0

    def _validate_date(self, year: int, month: int, day: int) -> bool:
        """Boolean check of date validity

        Create date based on year, month, day params.
        Validate based on the date_list created by earlier dates data

        Parameters
        ----------
            year : int
                year of date
            month : int
                month of date
            day : int
                day in month

        Returns
        -------
            bool
                exists_in_list
        """
        return f"{year}-{month}-{day}" in self.date_list

    def _get_max_day_in_yearmonth(self, year: int, month: int) -> int:
        """Get max day_number for given month & year

        Uses existing date_items DF, filters to given year & month, returns maximum
            day_number based on those params.

        Parameters
        ----------
            year : int
                given_year
            month : int
                given_month

        Returns
        -------
            int
                last_day_in_yearmonth
        """
        return (
            self.date_items.groupby(["year", "month_number"])["day_number"]
            .max()
            .reset_index()
            .query(f"year == {year} & month_number == {month}")["day_number"]
        )

    def _get_date_attribs(self, date: datetime) -> Tuple[int, bool]:
        """Get additional date attributes

        Gets additional date attributes based on passed datetime.
        Start day of week (day # in week based on datetime passed) and
        whether the week was an even week. This is used in bi-weekly freq's
        to understand what cadence to budget on

        Parameters
        ----------
            date : datetime
                start_date provided to base cadence upon

        Returns
        -------
            Tuple[int, bool]
                day_of_week, is_even_week
        """
        iter_date = self.dates[self.dates["date"] == date].iloc[0]
        start_day_of_week = iter_date["day_of_week"]
        is_even_week = self._is_even_week(iter_date["week_number"])
        return (start_day_of_week, is_even_week)

    def _audit_date_frequencies(self):
        """Check validity of dates in date_items

        Frequency defines day of budget in some cases (i.e. always due on 30th of month)
        If month does not have 30 days, we need to account for this and move to
            the max existing day in month.
        Iterates over rows in date_items, skips rows w/o frequency_day
        If row has an invalid date, set to max_day_in_yearmonth
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

    def _calc_multiplier(self, row: dict) -> float:
        """Calculate multiplier for budget_amount

        Calculates multiplier applied to input budget value
        If income, keep number positive else negative (expense)
        If seasonality applied, add seasonality_multiplier on top

        Parameters
        ----------
            row : dict
                row: dict (pd.DataFrame row)

        Returns
        -------
            float
                multiplier
        """
        return (1.00 if row["item_type"] == "Income" else -1.00) * (
            row["seasonality_multiplier"] if row["is_seasonality"] else 1.00
        )

    def _validate_record(self, row: dict) -> bool:
        """Validate basic requirements for date_item

        A few criterae to check before running through freq_types
            1 - if inactive record, return false
            2 - if end_date exists and date is past end, return false
            3 - if start_date exists and date is before start, return false

        Parameters
        ----------
            row : dict
                row: dict (pd.DataFrame row)

        Returns
        -------
            bool
                is valid record
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
        """Calculate budget amount for record in date_items

        Performs validation_check on record
        Checks by frequency_type for matching criteria
        Returns item_amount * calc_multiplier for records we want to budget for
            else 0.00

        Parameters
        ----------
            row : dict
                row of pd.Dataframe

        Returns
        -------
            float
                budget_amount
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
        """Applies function to each row in date_items

        Calculates budget_amount field using above function
        """
        self.date_items["budget_item_amount"] = self.date_items.apply(
            lambda x: self._calculate_budget_amount(x), axis=1
        )

    def build_data_model(self):
        """Runs individual steps to create data model"""
        self._get_dates()
        self._get_date_list()
        self._get_items()
        self._get_date_items()
        self._audit_date_frequencies()
        self._calc_budget_amounts()

    def get_df(self) -> pd.DataFrame:
        """Get DF limited to fields for tool

        Returns
        -------
            pd.DataFrame
                data model
        """
        df = self.date_items.copy()
        df = df[Models.ExportData.Columns]
        return df
