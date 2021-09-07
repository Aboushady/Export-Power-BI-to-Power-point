# Export-Power-BI-to-Power-point
How to export data from power BI to power point.

## Usage
### In this guide, I will walk you through the following:
  - How to write a basic python script that:
    - Connects to the power Bi's model, and gets its **Meta Data** .
    - Using the **Meta Data**, query power Bi's model, with **DAX** query.
    - Use the results fo these queries to create some visualizations e.g. pie charts and tables.
    - Create a **power point** file, and attach the visuals to it.
  - Direct you to another guide, that explains how we can add a new button to the **Power BI Desktop**'s ribbon, under *External tools* tap.  

## Requirements:
### Software
1- Python.\n
2- Power BI Desktop.\n
3- Power Point.\n
4- Anaconda [Optional].\n
### Packages
1- Python-pptx.\n
2- imgkit.\n
3- Pandas.\n
### Libraries
1- MSOLAP.\n
2- AMO.\n
3- ADOMD.\n

## [Python as an “External Tool” for Power BI Desktop](https://dataveld.com/2020/07/20/python-as-an-external-tool-for-power-bi-desktop-part-1/)
This is a guide on how to run a **python** script from within **Power BI**, 
it will point you to the configuration needed, and where to save the configuration files. In the following Python script I used his script to connect to power bi's model,
so that will save you some time, to only learn how to setup the configuration and save its file.

## EPPT "Export to Power Point" tool:
For the sake of this guide, I created a sample Excel sheet with some random data with the following fields:
Customer | Number of Purchases | Year
-------- | ------------------- | ----
A | 43 | 2020
B | 23 | 2021

That you should make your data source in the power bi's model.
### The goal is to group that table by Year to know total number of purchases per year.


'''python
import sys
import ssas_api as powerbi
import pandas as pd


"""
Part 0:
    Refer to this guid for info regarding the connection to Power BI's data base.
    https://dataveld.com/2020/07/20/python-as-an-external-tool-for-power-bi-desktop-part-1/
"""

print('Power BI Desktop Connection')
print(str(sys.argv[1]))
print(str(sys.argv[2]))

conn = "\nProvider=MSOLAP;Data Source=" + str(sys.argv[1]) + ";Initial Catalog='';"
print(conn)


dax_string = 'EVALUATE ROW("Loading .NET assemblies",1)'
df = powerbi.get_DAX(connection_string=conn, dax_string=dax_string)

print("\nCrossing the streams...")
global System, DataTable, AMO, ADOMD


"""
Part 1:
    Get the databases info, and retrieve the Meta data for the indended one,
    using the Microsoft Analysis services DLLs.
"""


import System
from System.Data import DataTable
import Microsoft.AnalysisServices.Tabular as TOM
import Microsoft.AnalysisServices.AdomdClient as ADOMD
from pptx import Presentation
from pptx.util import Inches
import plotly.express as px
from pandas.plotting import table
import imgkit

print("\nReticulating splines...")


# Connecting to Tabular AnalysisServices.
TOMServer = TOM.Server()
TOMServer.Connect(conn)

print("\nConnectoin Successfully to TOM's Server...")
print()

# Show Databases info
for item in TOMServer.Databases:
    print("Database: ", item.Name)
    print("Compatibility Level: ", item.CompatibilityLevel)
    print("Created: ", item.CreatedTimestamp)

input("Click to cont.")

'''
Retrieve the Meta data for the database we got from the connection earlier,
from TOM's databases.
'''

DatabaseId = str(sys.argv[2])
PowerBIDatabase = TOMServer.Databases[DatabaseId]


'''
