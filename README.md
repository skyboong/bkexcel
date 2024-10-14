# bkexcel

- This program helps enhance user convenience when saving a pandas DataFrame to an Excel file by adding Excel sheets and Excel
- BK Choi
- stsboongkee@gmail.com 
- Oct 10th, 2024 


## install
- pip install git+https://github.com/skyboong/bkexcel.git

## tutorial

```
import pandas as pd
from pandas import DataFrame, Series

from bkexcel import bkexcel as be 


data = [[2012, 13822078.0, 41628038.0, 3.85],
 [2013, 14241744.0, 45059205.0, 3.95],
 [2014, 15275007.0, 48459119.0, 4.08],
 [2015, 16293518.0, 49665854.0, 3.98],
 [2016, 16410047.0, 52995483.0, 3.99],
 [2017, 17737134.0, 61052054.0, 4.29],
 [2018, 18363011.0, 67365704.0, 4.52],
 [2019, 19095480.0, 69951597.0, 4.63],
 [2020, 21581228.0, 71490459.0, 4.8],
 [2021, 24094954.0, 78040289.0, 4.91],
 [2022, 26328329.0, 86317679.0, 5.21]]

df=DataFrame(data, columns=['PY','FUND1','FUND2', 'PCT'])

ex = be.BKExcelWriter(save_file_name=add_timestamp_to_filename("test.xlsx"))     
ex.to_sheet(df=df, sheet_name="Sheet1")
ex.set_settings(x_column='PY', w=3, left_gap=len(df.columns), style_no=10)
col1 = 'FUND1' 
col2 = 'FUND2'
col3 = 'PCT'
for ct in ['column', 'bar', 'line','area', 'radar', 'scatter','pie', 'doughnut']:
    ex.chart(columns_list=[col1, col2], title=f"Graph {ct}", chart_type=ct)
    ex.chart(columns_list=[col1, col2], title=f"Graph {ct} stacked", chart_type=ct, subtype='stacked')
    ex.chart(columns_list=[col1, col2], title=f"Graph {ct} percent_stacked", chart_type=ct, subtype='percent_stacked')

ex.chart_combined(col_left=col1, col_right=col3, title=f"{col1} and {col3}")
ex.chart_scatter(col_x='PY', col_y=col1, col_name='PY', title=f"년도별 투자액", col_size=col2, fixed_node_size=5)
ex.close()

```
