import os
import time
import xlwings as xw

INPUTS_PATH = "../src/Inputs.xlsx"


class InputConfig:
    def __init__(self):
        """Initializes class"""
        self.wb = None
        self.prompt()

    def _connectToWb(self):
        """Creates instance of workbook for Inputs"""
        self.wb = xw.Book(INPUTS_PATH)

    def prompt(self):
        """Open Excel & Prompt to continue"""
        print("Opening Inputs.xlsx..\n")

        self._connectToWb()
        print("Press enter when done editing..")

        self.wb.app.activate(steal_focus=True)
        _ = input("")

        self.wb.save()
        self.wb.close()
