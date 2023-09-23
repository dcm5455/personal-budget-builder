import os
import time
import xlwings as xw

INPUTS_PATH = "../src/Inputs.xlsx"


class InputConfig:
    def __init__(self):
        self._wb = None
        self._prompt()

    def _connectToWb(self):
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
        self._wb = xw.Book(INPUTS_PATH)

    def _prompt(self):
        print("Opening Inputs.xlsx..\n")
        self._connectToWb()
        print("Press enter when done editing..")
        ##time.sleep(2)
        self._wb.app.activate(steal_focus=True)
        _ = input("")
        self._wb.save()
        self._wb.close()
