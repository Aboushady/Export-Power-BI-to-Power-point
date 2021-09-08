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

1. Python
2. Power BI Desktop
3. Power Point
4. Anaconda [Optional]

### Packages

1. Python-pptx
2. imgkit
3. Pandas
4. Pythonnet,  package for .NET CLR
5. Python-SSAS, module (ssas_api.py placed in the same folder as the main script you’d like to run)

### Libraries

1. MSOLAP
2. AMO
3. ADOMD

## [Python as an “External Tool” for Power BI Desktop](https://dataveld.com/2020/07/20/python-as-an-external-tool-for-power-bi-desktop-part-1/)
This is a guide on how to run a **python** script from within **Power BI**, 
it will point you to the configuration needed, and where to save the configuration files. Also provides the python script to connect to PBI's model,
Here I am simply reusing and referncing it to showing you some things you can do with this data. Hence, you can save you some time, by only refering to [part 2](https://dataveld.com/2020/07/21/python-as-an-external-tool-in-power-bi-desktop-part-2-create-a-pbitool-json-file/) to learn 
how to setup the configuration PBItool.jason file and where to save it, but I strongly suggest that you read the whole thing.

## EPPT "Export to Power Point" tool:
For the sake of this guide, I created a sample Excel sheet *EPPT_test_data.xlsx*  with some random data with the following fields:

| Customer | Number of Purchases | Year |
| -------- | ------------------- | ---- |
| A | 43 | 2020                         |
| B | 23 | 2021                         |

That you should make your data source in the power bi's model.
*The goal is to group that table by Year to know the total number of purchases per year.*

## Connecting to  Tabular Object Model (TOM)
TOM is the .Net library opens Power BI’s data model to external tools.
The *python-ssas (ssas_api.py)* Python module that facilitates the TOM connection is all the work of Josh Dimarsky–originally for querying and processing Analysis Services.
Everything relies on Josh’s Python module, which has functions to connect to TOM, run DAX queries, etc.
As long as Power BI Desktop is installed, you should not have to manually obtain the required DLLs. You could get them from [Microsoft directly](https://docs.microsoft.com/en-us/analysis-services/client-libraries?view=azure-analysis-services-current) though if needed.


```python
import sys
import ssas_api as powerbi
import pandas as pd


"""
Part 0:
    Refer to this guide for info regarding the connection to Power BI's data model.
    https://dataveld.com/2020/07/20/python-as-an-external-tool-for-power-bi-desktop-part-1/
"""

print('Power BI Desktop Connection')
print(str(sys.argv[1]))
print(str(sys.argv[2]))

conn = "\nProvider=MSOLAP;Data Source=" + str(sys.argv[1]) + ";Initial Catalog='';"
print(conn)

```

**sys.argv[1]** is the argument corresponding to server and **sys.argv[2]** is the database GUID.

### Loading the .Net assemblies.
```python

dax_string = 'EVALUATE ROW("Loading .NET assemblies",1)'
df = powerbi.get_DAX(connection_string=conn, dax_string=dax_string)

```

### Get PBI model's Meta data throguh TOM 
```python
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

print("\nReticulating splines...")


# Connecting to Tabular AnalysisServices.
TOMServer = TOM.Server()
TOMServer.Connect(conn)
```

### Get the database info with the database GUID "sys.argv[2]".

```python

DatabaseId = str(sys.argv[2])
PowerBIDatabase = TOMServer.Databases[DatabaseId]

print("Listing tables...")
for table in PowerBIDatabase.Model.Tables:
    print(table.Name + "\n")


```
