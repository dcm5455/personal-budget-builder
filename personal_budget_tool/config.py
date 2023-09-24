import os
import time
import xlwings as xw

INPUTS_PATH = "../src/Inputs.xlsx"


class InputConfig:
    """_summary_

    _extended_summary_
    """

    def __init__(self):
        """_summary_

        _extended_summary_
        """
        self.wb = None
        self.prompt()

    def _connectToWb(self):
        """_summary_

        _extended_summary_
        """
        self._wb = xw.Book(INPUTS_PATH)

    def prompt(self):
        print("Opening Inputs.xlsx..\n")

        self._connectToWb()
        print("Press enter when done editing..")

        self._wb.app.activate(steal_focus=True)
        _ = input("")

        self._wb.save()
        self._wb.close()
