import os
import pandas as pd
import xlwings as xw
from typing import Tuple, Union
from datetime import datetime, date, timedelta

from constants import Defaults, Excel
from utils import get_col_char

TEMPLATE_PATH = "../src/Template.xlsx"
TEMPLATE_SHEET = "Template"
SAVE_PATH = "" "../src/output"


class BudgetApp:
    def __init__(self, min_date: datetime, max_date: datetime, df: pd.DataFrame):
        self.min_date = min_date
        self.max_date = max_date
        self.df = df
        self._xlApp = None
        self._wb = None
        self.month_years = self._get_unique_months()
        self.items = self._get_items()
        self.category_groups = self._get_category_groups()
        self.income_total_row = None
        self.expense_total_rows = []
        self.new_year_cols = []
        self._create_xlInstance()
        self._build_file()
        ## self._save_and_close()

    def _get_unique_months(self) -> list:
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
        return [
            (x["month_number"], x["year"])
            for _, x in self.df[self.df["budget_item_amount"] > 0][
                ["month_number", "year"]
            ]
            .drop_duplicates()
            .reset_index(drop=True)
            .iterrows()
        ]

    def _get_items(self) -> pd.DataFrame:
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
        df_copy = self.df.copy()
        df_copy["budget_item_amount_abs"] = abs(df_copy["budget_item_amount"])
        df_copy = df_copy[
            [
                "item_name",
                "display_group",
                "budget_item_amount",
                "budget_item_amount_abs",
            ]
        ]
        it = (
            df_copy.groupby(["item_name", "display_group"])[
                ["budget_item_amount", "budget_item_amount_abs"]
            ]
            .sum()
            .reset_index()
            .sort_values(by=["budget_item_amount_abs"], ascending=False)
            .reset_index(drop=True)
        )
        return it

    def _get_category_groups(self) -> pd.DataFrame:
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
            self.items.groupby("display_group")["budget_item_amount_abs"]
            .sum()
            .reset_index()
            .sort_values(by=["budget_item_amount_abs"], ascending=False)
            .reset_index(drop=True)
        )

    def _format_range(
        self,
        range: xw.Range,
        value: Union[str, int] = None,
        formula: str = None,
        format: Excel.FormatType = None,
        font_size: int = None,
        font_name: str = None,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        border: bool = False,
        border_pos: Excel.BordersIndex = Excel.BordersIndex.xlEdgeBottom,
        line_style: Excel.LineStyle = Excel.LineStyle.xlContinuous,
        border_weight: Excel.BorderWeight = Excel.BorderWeight.xlThin,
    ):
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
        if value:
            range.value = value
        if formula:
            range.formula = formula
        if format:
            range.number_format = format
        if font_size:
            range.font.size = font_size
        if font_name:
            range.font.name = font_name
        if bold:
            range.font.bold = True
        if italic:
            range.font.italic = True
        if underline:
            range.api.Font.Underline = 2
        if border:
            range.api.Borders(border_pos).LineStyle = line_style
            range.api.Borders(border_pos).Weight = border_weight

    def _create_xlInstance(self):
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
        # self._xlApp = xw.App(visible=True)
        self._wb = xw.Book(TEMPLATE_PATH)
        self._wb.app.activate(steal_focus=True)

    def _update_data(self):
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
        detail = self._wb.sheets["Data"]
        detail.range("A2:T100000").clear_contents()
        detail.range("A2").options(index=False, header=False).value = self.df

    def _create_summary_header(
        self, summary: xw.Sheet, col_index: int, month_year: Tuple[int, int]
    ):
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
        col = get_col_char(col_index)

        ##Hidden references
        summary.range(f"{col}5").formula = f"=DATE({col}7,{col}6,1)"
        summary.range(f"{col}6").value = month_year[0]
        summary.range(f"{col}7").value = month_year[1]

        ##Title
        title_cell = summary.range(f"{col}9")
        self._format_range(
            title_cell, formula=f"={col}5", format=Excel.FormatType.DateMonth, bold=True
        )

        ##Copy header fill cell
        summary.range(f"D2").copy(summary.range(f"{col}2"))

        ##Check if new year, add for later
        if col_index > 4 and month_year[0] == 1:
            self.new_year_cols.append(col_index)

    def _create_category(
        self, summary: xw.Sheet, row_index: int, max_col_index: int, categoryGroup: dict
    ):
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
        max_col_char = get_col_char(max_col_index)

        # Get items
        category_data = self.items[
            self.items["display_group"] == categoryGroup["display_group"]
        ]

        ##Create title cell
        title_cell = summary.range(f"B{row_index}")
        self._format_range(
            title_cell, value=categoryGroup["display_group"], bold=True, underline=True
        )

        ##Add border from title cell to max col width
        title_border = summary.range(f"B{row_index}:{max_col_char}{row_index}")
        self._format_range(title_border, border=True)

        ##Next we iterate over each item, add label
        item_row = row_index
        for _, item in category_data.iterrows():
            item_row += 1
            item_title_cell = summary.range(f"B{item_row}")
            item_title_cell.value = item["item_name"]

            # Add formulas for all months
            item_rng = summary.range(f"D{item_row}:{max_col_char}{item_row}")
            self._format_range(
                item_rng,
                format=Excel.FormatType.Number,
                formula=f"=IFERROR(SUMIFS(Data!$T:$T,Data!$H:$H,@$B:$B,Data!$E:$E,D$7,Data!$D:$D,D$6),0)",
            )

        ##For last row in "group" add bottom border
        item_btm_rng = summary.range(f"B{item_row}:{max_col_char}{item_row}")
        self._format_range(item_btm_rng, border=True)

        ##Identify total row for group
        total_row = item_row + 1
        if categoryGroup["display_group"] == "Income":
            self.income_total_row = total_row
        else:
            self.expense_total_rows.append(total_row)

        # Add text to total label
        total_cell = summary.range(f"B{total_row}")
        self._format_range(
            total_cell, value=f"Total {categoryGroup['display_group']}", bold=True
        )

        # Add border to totals row
        total_btm_rng = summary.range(f"B{total_row}:{max_col_char}{total_row}")
        self._format_range(total_btm_rng, border=True)

        ##Now do rest of totals row
        total_rng = summary.range(f"D{total_row}:{max_col_char}{total_row}")
        self._format_range(
            total_rng,
            formula=f"=SUM(D{row_index+1}:D{total_row-1})",
            format=Excel.FormatType.Number,
            bold=True,
            border=True,
        )

        ##Now add % of income?
        if categoryGroup["display_group"] == "Income":
            return total_row
        else:
            pct_row = total_row + 1
            pct_cell = summary.range(f"B{pct_row}")
            self._format_range(pct_cell, value="""% of Income""", italic=True)
            pct_rng = summary.range(f"D{pct_row}:{max_col_char}{pct_row}")
            self._format_range(
                pct_rng,
                formula=f"=-D{pct_row-1}/D{self.income_total_row}",
                format=Excel.FormatType.Percentage,
                italic=True,
            )
            return pct_row

    def _build_totals(self, summary: xw.Sheet, row_index: int, max_col_index: int):
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
        max_col_char = get_col_char(max_col_index)

        # Title for section
        total_title = summary.range(f"B{row_index}")
        self._format_range(total_title, value="Totals", bold=True, underline=True)
        row_index += 1

        # Income section
        income_title = summary.range(f"B{row_index}")
        self._format_range(income_title, value="Income")
        income_rng = summary.range(f"D{row_index}:{max_col_char}{row_index}")
        self._format_range(
            income_rng,
            formula=f"=D{self.income_total_row}",
            format=Excel.FormatType.Number,
        )
        row_index += 1

        # Expense section
        expense_title = summary.range(f"B{row_index}")
        self._format_range(expense_title, value="Expenses")
        expense_rng = summary.range(f"D{row_index}:{max_col_char}{row_index}")
        self._format_range(
            expense_rng,
            formula="=SUM("
            + (",".join(f"D{c}" for c in self.expense_total_rows))
            + ")",
            format=Excel.FormatType.Number,
        )
        expense_border_rng = summary.range(f"B{row_index}:{max_col_char}{row_index}")
        self._format_range(
            expense_border_rng, border=True, line_style=Excel.LineStyle.xlDouble
        )
        row_index += 1

        # Remaining Section
        remaining_bal_title = summary.range(f"B{row_index}")
        self._format_range(remaining_bal_title, value="Remaining Balance", bold=True)
        remaining_bal_rng = summary.range(f"D{row_index}:{max_col_char}{row_index}")
        self._format_range(
            remaining_bal_rng,
            formula=f"=D{row_index-1}+D{row_index-2}",
            format=Excel.FormatType.Number,
            bold=True,
        )
        remaining_border_rng = summary.range(f"B{row_index}:{max_col_char}{row_index}")
        self._format_range(remaining_border_rng, border=True)
        row_index += 1

        # % remaining section
        remaining_pct_title = summary.range(f"B{row_index}")
        self._format_range(remaining_pct_title, value="Remaining %", italic=True)
        remaining_pct_rng = summary.range(f"D{row_index}:{max_col_char}{row_index}")
        self._format_range(
            remaining_pct_rng,
            formula=f"=D{row_index-1}/D{self.income_total_row}",
            format=Excel.FormatType.Percentage,
            italic=True,
        )

        return row_index

    def _sheet_level_formatting(
        self, summary: xw.Sheet, max_col_index: int, last_row_index: int
    ):
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
        all_cells = summary.range(
            f"A1:{get_col_char(max_col_index+26)}{last_row_index+1000}"
        )
        self._format_range(all_cells, font_name="Arial", font_size=10)
        summary.range("B2").characters[0:15].font.size = 16  # reset title font

        summary.book.app.api.ActiveWindow.Zoom = 80

        for new_year in self.new_year_cols:
            new_year_col = summary.range(
                f"{get_col_char(new_year)}9:{get_col_char(new_year)}{last_row_index}"
            )
            self._format_range(
                new_year_col, border=True, border_pos=Excel.BordersIndex.xlEdgeLeft
            )

    def _build_file(self):
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
        summary = self._wb.sheets["Template"]
        summary.activate()

        ##Edit title
        summary.range("B2").value = (
            summary.range("B2")
            .value.replace("MinDate", self.min_date.strftime("%m/%Y"))
            .replace("MaxDate", self.max_date.strftime("%m/%Y"))
        )

        # Loop to create date headers
        start_col = 3
        for month_year in self.month_years:
            start_col += 1
            self._create_summary_header(summary, start_col, month_year)

        # Iterate over category groups now
        iter_row = 9
        for _, cg in self.category_groups.iterrows():
            iter_row += 2
            iter_row = self._create_category(summary, iter_row, start_col, cg)

        # Create totals
        iter_row += 2
        iter_row = self._build_totals(summary, iter_row, start_col)

        # Sheet-level formatting
        self._sheet_level_formatting(summary, start_col, iter_row)

    def save_and_close(self):
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
        self._wb.save(
            os.path.join(
                SAVE_PATH, f"Budget Tool {datetime.now().strftime('%Y%m%d')}.xlsx"
            )
        )
        self._wb.close()
        self._xlApp.kill()
