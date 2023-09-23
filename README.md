# Personal-Budget-Tool

Excel + Python app that generates a personal budget tool in Excel based on user input.

![Budget Tool](src/img/Budget.png)

## Overview
App to dynamically build a functional personal budget in Excel & Python based on user input. 

### Quick Start
- Download or clone the repository.
- Edit the Budget Items in the [Inputs](src/Inputs.xlsx) file to add your own expenses.
- Run `python personal_budget_tool/app.py` to generate your Excel File
- Confirm your changes are visibile in the summary & data tabs.

![Example Run](src/img/project.gif)

## Models

### Input (Budget Items)
Each Budget Item is generated from a row in the [Inputs](src/Inputs.xlsx) file.
- `is_active`
  - On/Off switch for each budget item 
- `is_seasonality`
  - Boolean whether seasonality should be applied (ex. electric bill in summer) 
- `company_name`
  - Company associated with budget item 
- `item_name`
  - Budget item name
- `category_name`
  - Dropdown for category associated with budget item 
- `category_group`
  - Higher-level grouping for categories 
- `display_group`
  - Grouping of items in grids in tool 
- `item_type`
  - Income / Expense  
- `item_amount`
  - Budgeted amount   
- `frequency_type`
  - Budgeted frequency (weekly, monthly, annually, etc.) 
- `frequency_day`
  - Day of {x} associated with frequency (i.e. monthly on the 1st day) 
- `frequency_date`
  - Full date associated with frequency (annual on 4/24, etc.) 
- `start_date`
  - Start date for budget item 
- `end_date`
  - End date for budget item 
- `notes`
  
## Application

### InputConfig
A small handler used to prompt the user for entries in the Inputs file. Opens the [Inputs](src/Inputs.xlsx) file to allow you to view/edit prior to building the budget. 

### DataBuilder
- Reads data from Inputs, generates a calendar tied to the budget. Writes data out for use in the output. Stores logic behind working with `frequency` & deciding where budgeted amounts will be allocated by day. 

### BudgetApp
- In short, a massive wrapper for [xlwings](https://github.com/xlwings/xlwings) operations. The template itself is barebones, so all of the styling, formulas and data is coming via this module.

## Dependencies
- Microsoft Excel
- Python 3.x
- xlwings >= 0.30.10
