# Personal-Budget-Tool

Excel + Python collaboration that generates a personal budget tool in Excel based on user inputs. 

![Example Run](src/img/project.gif)

## Overview
App to dynamically build a functional personal budget in Excel & Python based on user input. 

## Quick Start
- Downlad or clone the repository.
- Edit the Budget Items in the [Inputs](src/Inputs.xlsx) file to add your own expenses.
- Run `python personal_budget_tool/app.py` to generate your Excel File

The [Inputs](src/Inputs.xlsx) file acts as a configuration for the items being budgeted for: 
-  Item Name 
-  __Category__ pre-populated list of categories from Intuit
-  Amount
-  __Frequency__ Daily, Monthly, Annual, etc.
    
You can also add more advanced details to each item, such as:
- __Frequency Day__ (i.e. `5` if the bill is monthly, due on the fifth
- Frequency Date
  - (i.e. `2024-04-01` if you're going to have an annual subscription renew next year
- Start/End Dates
  - (i.e. you pay off your car in a few months)
- Seasonality (electricity costs more in the summer)

## `App.py`
### InputConfig
- A small handler used to prompt the user for entries in the Inputs file. 

## Instructions
- Add your desired Budget Items to the [Inputs](src/Inputs.xlsx) file.
  -  Item Name
  -  Category (Dropdown)
  -  Amount
  -  Frequency (Dropdown)
-  Run `python personal_budget_tool/app.py` to generate the tool
