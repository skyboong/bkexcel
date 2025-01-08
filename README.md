# bkexcel

- This program helps enhance user convenience when saving a pandas DataFrame to an Excel file by adding Excel sheets and Excel
- BK Choi
- stsboongkee@gmail.com 
- Oct 10th, 2024 


## install
- pip install git+https://github.com/skyboong/bkexcel.git

## tutorial


# The annual government investment amount, private investment amount, and the ratio of government investment to GDP are presented.

##  import 
```
import pandas as pd
from pandas import DataFrame, Series

from bkexcel import bkexcel as be 

```

## Example 1  
```
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
### Example 1-2
```
df=DataFrame(data, columns=['PY','FUND1','FUND2', 'PCT'])

ex = beg.BKExcelWriter(save_file_name="test.xlsx")     
t1 = df['PCT'].mean()
t2 = df['FUND1'].quantile(0.99)
ex.to_sheet(df=df, sheet_name="Sheet1", col_condition_list=[('FUND1',t2), ('PCT',t1)],
            dic_precision={'PCT':1, 'FUND1':0, 'FUND2':0}, font_size=9)
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
## Example 2 

```
dict1 = {'PY': {1: 2014.0, 2: 2015.0, 3: 2016.0, 4: 2017.0, 5: 2018.0, 6: 2019.0, 7: 2020.0, 8: 2021.0, 9: 2022.0}, 'COUNT': {1: 489, 2: 508, 3: 605, 4: 684, 5: 676, 6: 780, 7: 856, 8: 912, 9: 962}, 'FUND': {1: 774440024978, 2: 1060520824400, 3: 1251239811504, 4: 1448690184900, 5: 1487452351175, 6: 1762167598403, 7: 2146072944108, 8: 2049794327334, 9: 2062999362558}, 'FUND2': {1: 0.774440024978, 2: 1.0605208244, 3: 1.251239811504, 4: 1.4486901849, 5: 1.487452351175, 6: 1.762167598403, 7: 2.146072944108, 8: 2.049794327334, 9: 2.062999362558}, 'FUND_MEAN': {1: 1.5837219324703475, 2: 2.087639418110236, 3: 2.068164977692562, 4: 2.1179681065789473, 5: 2.200373300554734, 6: 2.259189228721795, 7: 2.5070945608738318, 8: 2.247581499269737, 9: 2.144489981869023}}

df_g = pd.DataFrame.from_dict(dict1)

style_no = 26 
col_x = 'PY'
col_y1 = 'FUND2' 
col_y2 = 'FUND_MEAN'
col_y3 = 'COUNT'

ex = be.BKExcelWriter(save_file_name= f"Table_01.xlsx")
ex.set_settings(x_column=col_x, w=2, left_gap=len(df_g.columns), style_no=style_no)
ex.to_sheet(df=df_g, sheet_name=f"Sheet1", dic_precision={col_y1:1, col_y2:2})

# combined chart : left, right 
ex.chart_combined(col_left=col_y1, col_right=col_y2, title=f"{col_y1} vs {col_y2}")

# scatter chart 
ex.chart_scatter(col_x=col_y1, col_y=col_y2, col_size=col_y3, col_name=col_x,
                title=f"{col_y1} vs {col_y2}, size={col_y3}" )
ex.close()

```

```
## Example 3 (2025.1..8)
```
dict1 = {'category': {1: 'FUND1', 2: 'FUND2', 3: 'PCT'},
 2012: {1: 13822078.0, 2: 41628038.0, 3: 3.85},
 2013: {1: 14241744.0, 2: 45059205.0, 3: 3.95},
 2014: {1: 15275007.0, 2: 48459119.0, 3: 4.08},
 2015: {1: 16293518.0, 2: 49665854.0, 3: 3.98},
 2016: {1: 16410047.0, 2: 52995483.0, 3: 3.99},
 2017: {1: 17737134.0, 2: 61052054.0, 3: 4.29},
 2018: {1: 18363011.0, 2: 67365704.0, 3: 4.52},
 2019: {1: 19095480.0, 2: 69951597.0, 3: 4.63},
 2020: {1: 21581228.0, 2: 71490459.0, 3: 4.8},
 2021: {1: 24094954.0, 2: 78040289.0, 3: 4.91},
 2022: {1: 26328329.0, 2: 86317679.0, 3: 5.21},
 'sum': {1: 203242530.0, 2: 672025481.0, 3: 48.21},
 'r_new': {1: 0.35427875750218224,
  2: 0.35095161369335043,
  3: 0.30947936112839664}}
df = DataFrame.from_dict(dict1)

style_no = 12 # 15 # 17 #18 # 16 

ex = be.BKExcelWriter(save_file_name=f"Test_style_no_{style_no}.xlsx")
ex.set_settings(style_no=style_no, df=df)
#ex.chart_total()

ex.set_data(   col_pie_list=['sum'], 
               col_doughnut_list=['sum'],
               col_column_list=['sum'],
               col_bar_list=['sum'],

               rows_list=['FUND1','FUND2'],   
               col_begin=2012, col_end=2022, 
               left_axis_title_line='FUND',
               bottom_axis_title_line='Year')

ex.chart_total(tf_graph1=True,
               col_pie_list=['sum'], 
               col_doughnut_list=['sum'],
               col_column_list=['sum'],
               col_bar_list=['sum']
    )

ex.chart_total(tf_trend_graph=True, sheet_name='sheet2')
ex.chart_total(
               #tf_set1 = False, 
               # scatter graph (Long)
               tf_scatter2=True,
               col_x='sum', col_y='r_new', col_size='sum',
               rows_list_sc = ['FUND1', 'FUND2'],
               fixed_node_size_sc=20,

               # combined graph 
               tf_combined=True,
               row_left_list=['FUND1', 'FUND2'],  row_right_list=['PCT'],
               left_name_list=['FUND1', 'FUND2'], right_name_list=['PCT'],
               left_axis_title_cb='FUND(M)', right_axis_title_cb='PCT(%)', 
               bottom_axis_title_cb='Bottom : Year', 

               # evolution track graph
               tf_evolution=True,
               row_x='FUND1', row_y='FUND2', row_size='PCT',
               left_axis_title_ep='Fund1(M)', bottom_axis_title_ep='Fund2(M)',
               fixed_node_size_ep=16,
               sheet_name='sheet3')
ex.close()

```



