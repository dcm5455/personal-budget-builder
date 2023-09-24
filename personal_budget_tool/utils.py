import itertools
import pandas as pd


def read_dataframe_input(
    Source: dict,
    Columns: list = None,
    IndexColumn: str = None,
    BoolColumns: list = None,
) -> pd.DataFrame:
    """Fetches a dataframe from a local Excel file.

    Parameters
    ----------
        Source (dict): pandas.read_excel parameters
        Columns (list, optional): column names to use. Defaults to None.
        IndexColumn (str, optional): str name of index column to be created based on ID. Defaults to None.
        BoolColumns (list, optional): list of str columns to be converted to boolean from Y/N text options in Excel. Defaults to None.

    Returns
    -------
        pd.DataFrame:
    """
    df = pd.read_excel(**Source)
    if Columns:
        df.columns = Columns
    if IndexColumn:
        df.insert(0, IndexColumn, df.index + 1)
    if BoolColumns:
        for bool_col in BoolColumns:
            df[bool_col] = df[bool_col].apply(lambda x: True if x == "Y" else False)

    return df


def filter_df_between(
    df: pd.DataFrame, col: str, vals: tuple, index_col: str = None
) -> pd.DataFrame:
    """Filters a dataframe based on a single key and two inclusive boundaries

    Reads a table from a local input .xlsx file. Parameters are specialized based on constants.Models

    Parameters
    ----------
        df (pd.DataFrame): The dataframe to be filtered
        col (str): The key to filter on
        vals (tuple): (min_value, max_value)
        index_col (str, optional): Re-create ID (index) column based on index, name str. Defaults to None.

    Returns
    -------
        pd.DataFrame: filtered dataframe
    """
    df = df[(df[col] >= vals[0]) & (df[col] <= vals[1])].reset_index(drop=True)
    if index_col:
        df[index_col] = df.index + 1
    return df


class DotDict(dict):
    """_summary_

    _extended_summary_

    Parameters
    ----------
        dict (_type_): _description_
    """

    __getattr__ = dict.get
    __setattr__ = dict.__setitem__
    __delattr__ = dict.__delitem__


def get_col_char(i: int) -> str:
    """Converts an integer index to the corresponding Excel Column (char)

    Parameters
    ----------
        i (int): Integer index off column

    Returns
    -------
        str: Excel column equivalent
    """
    string = ""
    while i > 0:
        i, remainder = divmod(i - 1, 26)
        string = chr(65 + remainder) + string
    return string
