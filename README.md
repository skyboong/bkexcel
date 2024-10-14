# bkexcel

- This program helps enhance user convenience when saving a pandas DataFrame to an Excel file by adding Excel sheets and Excel
- BK Choi
- stsboongkee@gmail.com 
- Oct 10th, 2024 


## install
- pip install git+https://github.com/skyboong/bkexcel.git

## tutorial

ex = be.BKExcelWriter(save_file_name=add_timestamp_to_filename("test.xlsx"))     
ex.to_sheet(df=df, sheet_name="Sheet1")
ex.set_settings(x_column='PY', w=3, left_gap=len(df.columns), style_no=10)
col1 = 'F1'
col2 = 'F2'
for ct in ['column', 'bar', 'line','area', 'radar', 'scatter','pie', 'doughnut']:
    ex.chart(columns_list=[col1, col2], title=f"Graph {ct}", chart_type=ct)
    ex.chart(columns_list=[col1, col2], title=f"Graph {ct} stacked", chart_type=ct, subtype='stacked')
    ex.chart(columns_list=[col1, col2], title=f"Graph {ct} percent_stacked", chart_type=ct, subtype='percent_stacked')

ex.chart_combined(col_left=col1, col_right=col2, title=f"{col1} and {col2}")
ex.chart_scatter(col_x='PY', col_y='F1', col_name='PY', title=f"년도별 투자액", col_size=None, fixed_node_size=5)
ex.close()


