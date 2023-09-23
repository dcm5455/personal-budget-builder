from utils import DotDict


class Defaults:
    DateFormats = DotDict({"NumberDate": "%Y-%m-%d", "MonthYear": "%m/%Y"})


class Models:
    BudgetDate = DotDict(
        {
            "Source": {
                "io": "../src/Inputs.xlsx",
                "sheet_name": "Dates",
                "header": 1,
                "usecols": "B:J",
            },
            "Columns": [
                "date",
                "day_of_week",
                "day_number",
                "week_number",
                "week_year",
                "month_number",
                "month_year",
                "year",
                "seasonality_multiplier",
            ],
            "IndexColumn": "date_id",
        }
    )

    BudgetItem = DotDict(
        {
            "Source": {
                "io": "../src/Inputs.xlsx",
                "sheet_name": "Budget Items",
                "header": 1,
                "usecols": "B:P",
            },
            "Columns": [
                "is_active",
                "is_seasonality",
                "company_name",
                "item_name",
                "category_name",
                "category_group",
                "display_group",
                "item_type",
                "item_amount",
                "frequency_type",
                "frequency_day",
                "frequency_date",
                "start_date",
                "end_date",
                "notes",
            ],
            "IndexColumn": "budget_item_id",
            "BoolColumns": ["is_active", "is_seasonality"],
        }
    )

    ExportData = DotDict(
        {
            "Columns": [
                "date_id",
                "date",
                "week_number",
                "month_number",
                "year",
                "is_active",
                "company_name",
                "item_name",
                "category_name",
                "category_group",
                "display_group",
                "item_type",
                "item_amount",
                "frequency_type",
                "frequency_day",
                "frequency_date",
                "start_date",
                "end_date",
                "is_seasonality",
                "seasonality_multiplier",
                "budget_item_amount",
            ]
        }
    )


class Excel:
    BordersIndex = DotDict(
        {
            "xlDiagonalDown": 5,
            "xlDiagonalUp": 6,
            "xlEdgeBottom": 9,
            "xlEdgeLeft": 7,
            "xlEdgeRight": 10,
            "xlEdgeTop": 8,
            "xlInsideHorizontal": 12,
            "xlInsideVertical": 11,
        }
    )
    BorderWeight = DotDict(
        {"xlHairline": 1, "xlMedium": -4138, "xlThick": 4, "xlThin": 2}
    )
    LineStyle = DotDict({"xlContinuous": 1, "xlDouble": -4119})
    Groups: DotDict(
        {
            "Income": ["Income", "Taxes", "Benefits"],
            "Other": [
                "Home & Utilities",
                "Auto & Transport",
                "Food & Dining",
                "Health & Personal",
                "Entertainment & Memberships",
                "Misc.",
            ],
        }
    )
    # IncomeCategoryGroups = DotDict(["Income", "Taxes", "Benefits"])
    FormatType = DotDict(
        {
            "Percentage": """0.0%""",
            "Number": """_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)""",
            "DateMonth": """[$-en-US]mmm-yyyy""",
        }
    )
    # OtherCategoryGroups = DotDict(
    #     [
    #         "Home & Utilities",
    #         "Auto & Transport",
    #         "Food & Dining",
    #         "Health & Personal",
    #         "Entertainment & Memberships",
    #         "Misc.",
    #     ]
    # )
