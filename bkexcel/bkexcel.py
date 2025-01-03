# bkexcel.py
# 2024.10.09
# 2024.10.23
# 2024.11.19
# 2024.12.5
# 2024.12.10
# 2024.12.31
# 2025.1.3 update at chart_scatter_beta0

import os
import datetime
import random
from functools import partial

import pandas as pd
from pandas import DataFrame, Series
import numpy as np

from icecream import ic
from xlsxwriter.utility import xl_rowcol_to_cell, xl_range, xl_range_abs



class BKExcelWriter:
    def __init__(self, writer=None, save_file_name=None, engine='xlsxwriter', sheet_name=None, add_prefix=True):

        if add_prefix:
            self.save_file_name = generate_filename_with_timestamp(save_file_name)

        self.writer = writer or pd.ExcelWriter(self.save_file_name, engine=engine)
        self.sheet_name = 'Sheet1' if sheet_name is None else sheet_name
        self.df = None
        self.x_column = None
        self.graph_no = 0
        self.pos_row = 1
        self.pos_col = 1
        self.pos_row_delta = 15
        self.pos_col_delta = 8
        self.pos_row_initial = 0
        self.pos_col_initial = 0
        self.w = 1 # 횡축 차트 개수
        self.style_no = 11

        self.col_pie_list = None
        self.rows_list = None
        self.col_begin = None
        self.col_end = None

    def to_sheet(self, df:DataFrame=None, sheet_name:str='Sheet1',
                 dic_width:dict={'논문수': 8, '총피인용수': 10},
                 dic_color:dict={}, dic_precision:dict={}, freeze_row:int=1,
                 col_con1:str=None, col_con2:str=None, threshold:int=0, condition_color:str=None,
                 col_condition_list:list=None,
                 fixed_width:int=10,
                 font_size:int=10,):
        """Save dataframe to Excel with formatting"""

        self.df = df
        self.sheet_name = sheet_name

        df.replace(np.nan, None, inplace=True)
        df.index = pd.RangeIndex(1, len(df.index) + 1)
        df.to_excel(self.writer, sheet_name=sheet_name, index=False)

        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        header_format = workbook.add_format({
            "font_size":font_size,
            "bold": True,
            "text_wrap": True,
            "valign": "center",
            "align": "center",
            "fg_color": "#D7E4BC",
            "border": 1,
        })

        user_format = workbook.add_format({
            "font_size":font_size,
            "bold": False,
            "text_wrap": True,
            "valign": "center",
            "align": "center",
            "fg_color": "#eeeeee",
            "border": 1,
        })
        user_format1 = workbook.add_format({
            "font_size": font_size,
            "bold": False,
            "valign": "center",
            "align": "center",
        })

        columns = df.columns.tolist()

        for j in range(len(columns)):
            try:
                worksheet.write(0, j, columns[j], header_format)
            except Exception as e:
                print(f"Error writing header: {e}")

        #worksheet.set_column(0, 0, 8)
        #(설명) worksheet.set_column(first_col, last_col, width=None, cell_format=None, options=None)

        worksheet.set_column(0, len(df.columns)-1, fixed_width, user_format1) # 첫번재 열 부터 마지막 열까지 폭 지정, 0 부터 시작함

        for k, v in dic_width.items():
            if k in columns:
                pi = columns.index(k)
                worksheet.set_column(pi, pi, v)

        for k, v in dic_color.items():
            if k in columns:
                pi = columns.index(k)
                color_format1 = workbook.add_format({'fg_color': v, 'text_wrap': True, 'font_size': font_size})
                for i in range(len(df.index)):
                    worksheet.write(i+1, pi, df.iloc[i, columns.index(k)], color_format1)

        if col_con1 and col_con2:
            #print("col_con1 and col_con2")
            start_row = 1 # 0은 header
            end_row = len(df.index) # 0은 헤더가 차지하기에, 1 부터 시작함
            start_col = columns.index(col_con1)
            end_col = columns.index(col_con2)
            color1 = condition_color if condition_color is not None else "#D5F5E3" #C6EFCE"
            format_condition = workbook.add_format({"bg_color": color1, "font_color": "#006100", "font_size": font_size})
            worksheet.conditional_format(start_row,start_col,end_row,end_col,
                                         {"type": "cell", "criteria": ">=", "value": threshold, "format": format_condition})

        if col_condition_list is not None:
           for col_con, each_threshold in col_condition_list:
               #print("col_con1 and col_con2")
               start_row = 1 #
               end_row = len(df.index)  # 0은 헤더가 차지하기에, 1 부터 시작함
               start_col = columns.index(col_con)
               end_col = columns.index(col_con)
               color1 = condition_color if condition_color is not None else "#D5F5E3"  # C6EFCE"
               format_condition = workbook.add_format({"bg_color": color1, "font_color": "#006100", 'font_size': font_size})
               worksheet.conditional_format(start_row, start_col, end_row, end_col,
                                            {"type": "cell", "criteria": ">=", "value": each_threshold,
                                             "format": format_condition})

        for k, v in dic_precision.items():
            if k in columns:
                pi = columns.index(k)
                if v == 2:
                    precision_format = workbook.add_format({'num_format': '#,##0.00', 'font_size': font_size})
                elif v == 1:
                    precision_format = workbook.add_format({'num_format': '#,##0.0', 'font_size': font_size})
                elif v == 0:
                    precision_format = workbook.add_format({'num_format': '#,##0', 'font_size': font_size})
                else:
                    precision_format = workbook.add_format({'num_format': '#,##0.00', 'font_size': font_size})
                for i in range(len(df.index)):
                    worksheet.write(i+1, pi, df.iloc[i, columns.index(k)], precision_format)

        worksheet.freeze_panes(freeze_row, 0)
        worksheet.freeze_panes(freeze_row, 1)

    def close(self):
        self.writer.close()
        print(f"* saved file : '{self.save_file_name}'")



    def chart_scatter(self, col_x=None, col_y=None, col_size=None, col_name=None,
                      title=None, pos_row=None, pos_col=None,
                      style_no=None,
                      fixed_node_size=None,
                      min_range=3, max_range=30,
                      title_font_size=10,
                      line=False, line_width=8, line_marker_size=5, alpha=0.8,
                      label_left=None, label_right=None, label_bottom=None):

        self.graph_no += 1
        df = self.df
        sheet_name = self.sheet_name

        workbook = self.writer.book
        worksheet = self.writer.sheets.get(sheet_name)
        if worksheet is None:
            raise ValueError(f"Sheet name '{sheet_name}' does not exist. Please check the sheet name.")

        columns = df.columns.tolist()
        (max_row, max_col) = df.shape

        # X, Y, name, size column indices with checks for existence
        xi = columns.index(col_x) if col_x in columns else 0
        yi = columns.index(col_y) if col_y in columns else 1
        namei = columns.index(col_name) if col_name and col_name in columns else None

        if style_no is None:
            style_no = self.style_no

        self.chart_position()
        pos_row = self.pos_row if pos_row is None else pos_row
        pos_col = self.pos_col if pos_col is None else pos_col

        # Create scatter chart
        chart_type = {'type': 'scatter', 'subtype': 'straight_with_markers'} if line else {'type': 'scatter'}
        chart = workbook.add_chart(chart_type)

        # Generate colors for each unique value in col_name
        if col_name and col_name in columns:
            unique_names = df[col_name].unique()
            color_map = {name: "#{:06x}".format(random.randint(0, 0xFFFFFF)) for name in unique_names}
        else:
            color_map = {}

        col_size_list = None
        if col_size and col_size in df:
            min_size = min(df[col_size])
            max_size = max(df[col_size])

            # Scale sizes between min_range and max_range
            col_size_list = [
                int(((size - min_size) / (max_size - min_size)) * (max_range - min_range) + min_range)
                for size in df[col_size]
            ]

        # Add scatter points with colors based on col_name
        for i in range(1, max_row + 1):
            marker_size = col_size_list[i - 1] if col_size_list else (fixed_node_size or 5)
            marker_color = color_map.get(df[col_name].iloc[i - 1],
                                         '#000000')  # default to black if col_name is not in columns

            series_options = {
                'name': [sheet_name, i, namei] if namei is not None else '',
                'categories': [sheet_name, i, xi, i, xi],
                'values': [sheet_name, i, yi, i, yi],
                'marker': {
                    'type': 'circle',
                    'size': marker_size,
                    'border': {'color': 'black'},
                    'fill': {'color': marker_color}
                },
            }
            chart.add_series(series_options)

        # If line option is enabled, add a connecting line as a single series for the entire range
        if line:
            chart.add_series({
                'name': 'Line Series',
                'categories': [sheet_name, 1, xi, max_row, xi],
                'values': [sheet_name, 1, yi, max_row, yi],
                'line': {
                    'width': line_width,
                    'color': 'blue',
                },
                'marker': {'type': 'circle', 'size': line_marker_size}
            })
        # Set chart title and axes
        chart.set_title(
            {'name': title if title else 'Scatter Chart', 'name_font': {'size': title_font_size}, 'bold': True})
        chart.set_x_axis({'name': col_x if label_bottom is None else label_bottom})
        chart.set_y_axis({'name': col_y if label_left is None else label_left})
        # Set chart style
        chart.set_style(style_no)
        # Insert chart into sheet
        worksheet.insert_chart(pos_row, pos_col, chart)

    def _get_scaled_sizes(self, df, col_index, row_size, col_begin, col_end, min_range, max_range):
        """Calculate scaled sizes for chart markers."""
        if not row_size:
            return None
        left_list = df[col_index].tolist()
        columns_top = df.columns.tolist()
        min_size = df.iloc[left_list.index(row_size), columns_top.index(col_begin): columns_top.index(col_end) +1].min()
        max_size = df.iloc[left_list.index(row_size),  columns_top.index(col_begin): columns_top.index(col_end)+1].max()
        values = df.iloc[left_list.index(row_size), columns_top.index(col_begin): columns_top.index(col_end) +1].tolist()
        scaled_sizes = [
            ((value - min_size) / (max_size - min_size)) * (max_range - min_range) + min_range
            for value in values
        ]
        return [int(size) for size in scaled_sizes]


    def chart_scatter_beta0_W(self,
                              col_index=None,
                              columns_list=[], col_begin=None, col_end=None,
                              row_x=None,  row_y=None, row_size=None, row_name=None,
                              name_of_evolution_track=None, title='',
                              auto_no_title=True, pos_row=None, pos_col=None, style_no=None, fixed_node_size=None,
                              min_range=10, max_range=50, title_font_size=10, line=False, line_width=8,line_color=None,
                              line_marker_size=5, alpha=0.8, dict_scale=None,
                              label_left=None, label_bottom=None,
                              left_axis_title=None, bottom_axis_title=None,
                              data_kind='W', verbose=False):
        """Insert scatter chart into Excel sheet."""
        #print(">>> chart_scatter_beta0_W()")
        #try:
        if True:
            ic.disable()
            self.graph_no += 1
            title = f'Fig {self.graph_no}. ' + title if auto_no_title else title
            df = self.df
            #print(f"col_index={col_index}")
            #print(f"df.columns={df.columns}")
            # if col_index in df.columns:
            #     left_list = df[col_index].tolist()
            # else:
            #     left_list = df.index.tolist()
            sheet_name = self.sheet_name
            workbook = self.writer.book
            worksheet = self.writer.sheets[sheet_name]

            columns1 = df.columns.tolist()
            if col_index not in columns1:
                #print(f"* col_index is {col_index}, not in columns!")
                #print(f"* columns={columns1}")
                col_index = columns1[0] # 제일 첫 컬럼을 자동으로 지정해 주기
                #print(f"* col_index was automatically replaced by '{col_index}'")
            rows_total = df[col_index].tolist()

            #print(f"* rows_total: {rows_total}")
            # 최대 값
            (max_row, max_col) = df.shape
            # X, Y, name, size 컬럼의 인덱스 결정
            xi = 1 + rows_total.index(row_x) if row_x else 0
            yi = 1 + rows_total.index(row_y) if row_y else 1
            #print(f"xi={xi}, yi={yi}")
            namei = rows_total.index(row_name) if row_name is not None else 0
            #print("* namei", namei)
            if style_no is None:
                style_no = self.style_no
            self.chart_position()
            pos_row = self.pos_row if pos_row is None else pos_row
            pos_col = self.pos_col if pos_col is None else pos_col

            # Scatter chart 생성
            chart = workbook.add_chart({'type': 'scatter'})
            #ic()
            col_size_list = None
            if row_size:
                if max_range > 72 :
                    max_range = 72
                scaled_sizes = self._get_scaled_sizes(df=df, col_index=col_index, row_size=row_size, col_begin=col_begin, col_end=col_end, min_range=min_range,max_range=max_range,)
                if verbose:
                    print(f"* scaled_sizes={scaled_sizes}")
            ic()
            # 각 컬럼에 대해 시리즈 추가
            top_list = df.columns.tolist()[1:]

            if col_begin is not None:
                col_begin_i = fn_find_col_i(df, col=col_begin)
            else:
                col_begin_i = 0 + 1

            if col_end is not None:
                col_end_i = fn_find_col_i(df, col=col_end)
            else:
                col_end_i = len(top_list)
            col_list_i = []

            if top_list is not None:
                for each_col in top_list:
                    col_list_i.append(fn_find_col_i(df, col=each_col))

            left_list_i = list(range(col_begin_i, col_end_i+1))
            top_selected_list_i = left_list_i if top_list is not None else left_list_i
            if verbose:
                print(f"* top_selected_list_i=", top_selected_list_i)

            i = 0
            for ci in range(1, len(top_list) + 1):
                if ci in top_selected_list_i :
                    #print(f"*i={i}, ci={ci}")
                    series_options = {
                        'name':       [sheet_name, 0, ci],
                        'categories': [sheet_name, xi, ci, xi, ci],  # X축 데이터
                        'values':     [sheet_name, yi, ci, yi, ci],      # Y축 데이터
                        #'line':{'color':'blue'},
                        'marker': {
                            'type': 'circle',
                            'size': scaled_sizes[i] if fixed_node_size is None else fixed_node_size,  # 마커 크기 설정
                            'border':{'color':'black'},
                            'line': {'color': 'black'},
                        },
                    }
                    # series 추가
                    chart.add_series(series_options)
                    i +=1

            # line 추가 부분
            #if col_x in columns and col_y in columns:
            #ic(f"{col_x}")
            if line:
                #row_xi = df.index.tolist().index(row_x)
                #row_yi = df.index.tolist().index(row_y)
                row_xi = xi
                row_yi = yi

                chart.add_series({
                    #'name': [sheet_name, 0, row_xi],
                    'name': 'Evolution Track' if name_of_evolution_track is None else name_of_evolution_track,
                    'categories': [sheet_name, row_xi, col_begin_i, row_xi, col_end_i],
                    'values': [sheet_name, row_yi, col_begin_i, row_yi, col_end_i],
                    'fill': {'transparency': alpha},
                    #'marker':{ 'size':line_marker_size,
                    #           'type':'circle'},
                    #'marker':{'type':None},
                    'line':{'width':line_width, 'color': line_color if line_color else 'black'},
                })
                #print(745)

            # 차트 제목과 축 설정
            chart.set_title({'name': title if title else 'Scatter Chart',
                             'name_font':{'size': title_font_size},

                             'bold':True})
            chart.set_x_axis({'name': row_x if bottom_axis_title is None else bottom_axis_title})
            chart.set_y_axis({'name': row_y if left_axis_title is None else left_axis_title})

            # 스타일 설정
            chart.set_style(style_no)

            # 차트를 시트에 삽입
            #worksheet.insert_chart(pos_row, pos_col, chart)
            dict_scale1 = {'x_scale': 1.0, 'y_scale': 1.0, 'width': 640, 'height': 480}
            dict_scale = dict_scale1 if dict_scale is None else dict_scale
            if dict_scale:
                worksheet.insert_chart(pos_row, pos_col, chart, dict_scale)
            else:
                worksheet.insert_chart(pos_row, pos_col, chart)
        #except Exception as err:
        #    raise Exception(err)


    def chart_scatter_beta0(self, col_x=None, col_y=None, col_size=None, col_name=None,
                            rows_list=None,
                            name_of_evolution_track = None,
                            title=None, auto_no_title=True,
                            pos_row=None, pos_col=None,
                       style_no=None,
                       fixed_node_size=None,
                       min_range=3, max_range=30,
                       title_font_size=10,
                       line = False, line_width=8, line_marker_size=5,
                       alpha=0.8,
                        dict_scale=None,
                       label_left=None, label_bottom=None):
        """Insert scatter chart into Excel sheet."""
        print(f">>> chart_scatter_beta0()")
        try:
            self.graph_no += 1
            title = f'Fig {self.graph_no}. ' + title if auto_no_title else title

            df = self.df
            sheet_name = self.sheet_name
            workbook = self.writer.book
            worksheet = self.writer.sheets[sheet_name]
            columns = df.columns.tolist()
            # 최대 값
            (max_row, max_col) = df.shape
            # X, Y, name, size 컬럼의 인덱스 결정
            xi = columns.index(col_x) if col_x else 0
            yi = columns.index(col_y) if col_y else 1
            namei = columns.index(col_name) if col_name in columns else 0  # 없으면 0 번째 컬럼 지정
            if style_no is None:
                style_no = self.style_no
            self.chart_position()
            pos_row = self.pos_row if pos_row is None else pos_row
            pos_col = self.pos_col if pos_col is None else pos_col
            # Scatter chart 생성
            #if line :
            #    chart = workbook.add_chart({'type': 'scatter', 'subtype':'straight_with_markers'})
            #else:
            chart = workbook.add_chart({'type': 'scatter'})

            col_size_list = None
            if col_size:
                # col_size가 있으면, 해당 값을 리스트로 저장 (정수로 변환)
                # 최소 및 최대 크기 (xlsxwriter의 범위)
                #min_range = 2
                #max_range = 72
                # 너무 크게 되는 것을 제외시킴
                if max_range > 72 :
                    max_range = 72

                # 데이터에서 최소, 최대 값을 추출
                min_size = min(df[col_size])
                max_size = max(df[col_size])

                # 크기 데이터를 2 ~ 72 사이로 스케일링
                scaled_sizes = [
                    ((each_size - min_size) / (max_size - min_size)) * (max_range - min_range) + min_range for each_size in
                    df[col_size]
                ]
                try:
                    # 정수로 변환 가능한지 확인하고 변환
                    col_size_list = [int(size) for size in scaled_sizes]
                except ValueError:
                    raise ValueError(f"Column {col_size} contains non-integer values. Ensure all values are integers.")

            # selected rows list  2025.1.2
            rows_list_i = []
            col_name = col_name if col_name is not None else columns[0]
            index_list = df[col_name].tolist()
            for each in rows_list:
                row_i = fn_find_col_index(index_list, each) + 1
                rows_list_i.append(row_i)

            # 각 행에 대해 시리즈 추가
            #for i in range(1, len(df.index) + 1):
            no_i = 1
            for i in rows_list_i:
                series_options = {
                    'name': [sheet_name, i, namei] if namei is not None else '',
                    'categories': [sheet_name, i, xi, i, xi],  # X축 데이터
                    'values': [sheet_name, i, yi, i, yi],  # Y축 데이터
                    #'line':{'color':'blue'},
                    'marker': {
                        'type': 'circle',
                        'size': col_size_list[i - 1] if fixed_node_size is None else fixed_node_size,  # 마커 크기 설정
                        'border':{'color':'black'},
                        'line': {'color': 'black'},
                    },

                    'y2_axis':False if no_i == 1 else True, # 좀 이상해 보여도 적어도 1개는  False 로 지정해 줘야 함
                }
                # series 추가
                chart.add_series(series_options)
                no_i += 1

            # line 추가 부분
            #if col_x in columns and col_y in columns:
            #ic(f"{col_x}")
            if line:
                col_xi = df.columns.tolist().index(col_x)
                col_yi = df.columns.tolist().index(col_y)

                chart.add_series({
                    #'name': [sheet_name, 0, col_xi],
                    'name': 'Evolution Track' if name_of_evolution_track is None else name_of_evolution_track,
                    'categories': [sheet_name, 1, col_xi, max_row, col_xi],
                    'values': [sheet_name, 1, col_yi, max_row, col_yi],
                    'fill': {'transparency': alpha},
                    #'marker':{ 'size':line_marker_size,
                    #           'type':'circle'},
                    #'marker':{'type':None},
                    'line':{'width':line_width}
                })
                #print(745)

            # 차트 제목과 축 설정
            chart.set_title({'name': title if title else 'Scatter Chart',
                             'name_font':{'size': title_font_size},
                             'bold':True})
            chart.set_x_axis({'name': col_x if label_bottom is None else label_bottom,
                             'line': {'color':'black'} })
            chart.set_y_axis({'name': col_y if label_left is None else label_left,
                              'line': {'color': 'black'} ,
                              'major_gridlines': {'visible': True},
                              })
            chart.set_y2_axis({'line': {'color':'black'} ,
                               'major_gridlines': {'visible': False},
                               })

            # 스타일 설정
            chart.set_style(style_no)

            # 차트를 시트에 삽입
            #worksheet.insert_chart(pos_row, pos_col, chart)
            dict_scale1 = {'x_scale': 1.0, 'y_scale': 1.0, 'width': 640, 'height': 480}
            dict_scale = dict_scale1 if dict_scale is None else dict_scale
            if dict_scale:
                worksheet.insert_chart(pos_row, pos_col, chart, dict_scale)
            else:
                worksheet.insert_chart(pos_row, pos_col, chart)
        except Exception as err:
            raise Exception(err)



    def chart_combined_v3(self,
                          col_bottom=None,
                          col_left_list=None,
                          col_right_list=None,
                          title='', auto_no_title=True, title_font_size=10,
                          left_name_list=None, right_name_list=None,
                          left_axis_title=None, right_axis_title=None, bottom_axis_title=None,
                          pos_row=None, pos_col=None,
                          style_no=None, line_width=2, dict_scale={},
                          right_y_axis_range=None,

                          left_chart_style='column',
                          right_chart_style='line',

                          data_kind="L",
                          col_select=None,

                          col_begin=None,
                          col_end=None,
                          row_begin_left=None,
                          row_begin_right=None,
                          row_end_left=None,
                          row_end_right=None,
                          row_left_list=None,
                          row_right_list=None,
                          row_bottom=None,

                          data_labels_left_tf=False,
                          data_labels_right_tf=False,
                          data_labels_left_num_format=None,
                          data_labels_right_num_format=None,
                          left_labels_font_size=6,
                          right_labels_font_size=6,
                          line_marker={},

                          verbose=False
                          ):

        """Insert combined chart into Excel sheet"""
        # chart1, chat2 2개 chart 를 만들어서 결합시켜줌
        # bottom : Category_both 공통되게 지정해줌
        # name1, value1 : 왼쪽 y 값
        # name2, value2 : 오른쪽 y값

        ic.disable()

        self.graph_no += 1
        title = f'Fig {self.graph_no}. ' + title if auto_no_title else title
        df = self.df

        sheet_name = self.sheet_name
        col_bottom = self.x_column if col_bottom is None else col_bottom
        style_no = self.style_no if style_no is None else style_no

        left_name_list = col_left_list if left_name_list is None else left_name_list
        right_name_list = col_right_list if right_name_list is None else right_name_list

        left_axis_title = left_name_list[0] if left_axis_title is None else left_axis_title
        right_axis_title = right_name_list[0] if right_axis_title is None else right_axis_title

        # pos_row 자동 할당하기
        self.chart_position()
        pos_row = self.pos_row
        pos_col = self.pos_col
        #ic()
        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        (max_row, max_col) = df.shape

        columns = df.columns.tolist()



        match(data_kind):
            case 'L'|'Long':

                col_bottom_i = columns.index(col_bottom) if col_bottom in columns else 0
                col_y1i_list = [columns.index(each) for each in col_left_list]
                col_y2i_list = [columns.index(each) for each in col_right_list]

                category_both = f"={sheet_name}!{xl_range_abs(1, col_bottom_i, max_row, col_bottom_i)}"
                # chart1
                chart1 = workbook.add_chart({'type': left_chart_style})
                for i, col_yi in enumerate(col_y1i_list):
                    values = f"={sheet_name}!{xl_range_abs(1, col_yi, max_row, col_yi)}"
                    ic()
                    chart1.add_series({
                        'name': left_name_list[i],
                        'categories': category_both,
                        'values': values,
                        'line': {'width': line_width}
                    })
                # chart2
                chart2 = workbook.add_chart({'type': right_chart_style})
                for i, col_yi in enumerate(col_y2i_list):
                    ic(col_yi)
                    values = f"={sheet_name}!{xl_range_abs(1, col_yi, max_row, col_yi)}"
                    ic()
                    chart2.add_series({
                        'name': right_name_list[i],
                        'categories': category_both,
                        'values': values,
                        'y2_axis': True,
                        'line': {'width': line_width}
                    })
            case 'W'|'WIDE':
                # 컬럼 인덱스 번호 정리
                columns = df.columns.tolist()
                columns_list_no = []
                # 선택된 값을 가지는 row 추출 하기 위하여
                col_select = columns[0] if col_select is None else col_select

                col_begin_i = columns.index(
                    col_begin) if col_begin is not None and col_begin in columns else 1  # columns[0]
                col_end_i = columns.index(col_end) if col_end is not None and col_end in columns else len(
                    columns) - 1  # columns[0]

                # row 1 순위
                row_left_list_no = []
                row_right_list_no = []

                rows = df[col_select].tolist()
                if len(row_left_list) > 0:
                    for each in row_left_list:
                        row_left_list_no.append(rows.index(each) + 1)
                if len(row_right_list) > 0:
                    for each in row_right_list:
                        row_right_list_no.append(rows.index(each) + 1)

                #category_both = f"={sheet_name}!{xl_range_abs(0, col_begin_i, 0, col_end_i)}"
                category_both = f"={sheet_name}!{xl_range_abs(0, col_begin_i, 0, col_end_i)}"
                if verbose:
                    print(f"category_both={category_both}")

                data_labels_left = {
                    'value': True if data_labels_left_tf else False,
                    'num_format': '#,##0.00' if data_labels_left_num_format is None else data_labels_left_num_format,
                    'font':{'size':left_labels_font_size}
                }
                data_labels_right = {
                    'value': True if data_labels_right_tf else False,
                    'num_format': '#,##0.00' if data_labels_right_num_format is None else data_labels_right_num_format,
                    'font':{'size':right_labels_font_size}
                }
                # chart1
                chart1 = workbook.add_chart({'type': left_chart_style})
                for i, row_i in enumerate(row_left_list_no):
                    #category_both = f"={sheet_name}!{xl_range_abs(0, col_begin_i, 0, col_end_i)}"
                    values = f"={sheet_name}!{xl_range_abs(row_i, col_begin_i, row_i, col_end_i)}"
                    ic()
                    chart1.add_series({
                        'name': left_name_list[i],
                        'categories': category_both,
                        'values': values,
                        'line': {'width': line_width},
                        'marker': line_marker if line_marker else {},
                        'data_labels':data_labels_left
                    })
                # chart2
                chart2 = workbook.add_chart({'type': right_chart_style})
                for i, row_i in enumerate(row_right_list_no):
                    #category_both = f"={sheet_name}!{xl_range_abs(0, col_begin_i, 0, col_end_i)}"
                    values = f"={sheet_name}!{xl_range_abs(row_i, col_begin_i, row_i, col_end_i)}"
                    ic()
                    chart2.add_series({
                        'name': right_name_list[i],
                      #  'categories': category_both,
                        'values': values,
                        'y2_axis': True,
                        'line': {'width': line_width},
                        'marker': line_marker if line_marker else {},
                        'data_labels':data_labels_right
                    })
            case _ :
                raise ValueError(f'Unsupported data type: {data_kind}, supported data types are L or W')


        # Chart 합하기
        chart1.combine(chart2)
        chart1.set_title({'name': title,
                          'name_font': {'size': title_font_size  # 원하는 폰트 크기 설정
                                        }
                          })
        #print(f"**** col_bottom=", col_bottom)
        bottom_axis_title = columns[0] if bottom_axis_title is None else bottom_axis_title
        chart1.set_x_axis({'name': bottom_axis_title,
                           'categories': category_both})
        chart1.set_y_axis({'name': left_axis_title,
                           'major_gridlines': {'visible': True}
                           })
        if right_y_axis_range is not None:
            chart2.set_y2_axis({
                'name': right_axis_title,  # 오른쪽 축 레이블
                'major_gridlines': {'visible': False},  # 오른쪽 축에 그리드라인 표시하지 않음
                'min': right_y_axis_range[0],
                'max': right_y_axis_range[1],
            })
        else:
            chart2.set_y2_axis({
                'name': right_axis_title,  # 오른쪽 축 레이블
                'major_gridlines': {'visible': False},  # 오른쪽 축에 그리드라인 표시하지 않음
            })
        chart1.set_legend({'none': False,
                           'position': 'top'})
        chart1.set_style(style_no)
        dict_scale1 = {'x_scale': 1.0, 'y_scale': 1.0, 'width': 640, 'height': 480}
        dict_scale = dict_scale1 if dict_scale is None else dict_scale
        if dict_scale:
            worksheet.insert_chart(pos_row, pos_col, chart1, dict_scale)
        else:
            worksheet.insert_chart(pos_row, pos_col, chart1)


    def chart_combined_v2(self,
                          col_bottom=None, col_left=None, col_right_list=None, title='', auto_no_title=True,
                          title_font_size=10, left_name=None, right_name=None, left_axis_title=None,
                          right_axis_title=None, pos_row=None, pos_col=None, style_no=None, line_width=2, dict_scale={},
                          right_y_axis_range=None,
                          data_kind="L"):
        """Insert combined chart into Excel sheet"""
        ic.disable()
        self.graph_no += 1
        title = f'Fig {self.graph_no}. ' + title if auto_no_title else title
        df = self.df
        sheet_name = self.sheet_name

        col_bottom = self.x_column if col_bottom is None else col_bottom
        style_no = self.style_no if style_no is None else style_no
        left_name = col_left if pd.isna(left_name) else left_name
        right_name = col_right_list if pd.isna(right_name) else right_name
        left_axis_title = left_axis_title if left_axis_title is None else left_axis_title
        right_axis_title = right_axis_title if right_axis_title is None else right_axis_title

        # pos_row 자동 할당하기

        self.chart_position()
        pos_row = self.pos_row
        pos_col = self.pos_col

        ic()
        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]


        (max_row, max_col) = df.shape

        columns = df.columns.tolist()
        col_xi = columns.index(col_bottom)
        col_y1i = columns.index(col_left)
        col_y2i_list = [columns.index(each) for each in col_right_list]
        ic()

        category1  = f"={sheet_name}!{xl_range_abs(1, col_xi, max_row, col_xi)}"
        values1 = f"={sheet_name}!{xl_range_abs(1, col_y1i, max_row, col_y1i)}"

        chart1= workbook.add_chart({'type': 'column'})
        chart1.add_series({
            'name':left_name,
            'categories': category1,
            'values': values1,

        })
        ic()

        chart2 = workbook.add_chart({'type': 'line'})
        for i, col_y2i in enumerate(col_y2i_list):
            ic(col_y2i)
            ic(right_name)
            values2 = f"={sheet_name}!{xl_range_abs(1, col_y2i, max_row, col_y2i)}"

            ic()
            chart2.add_series({
                'name': right_name[i],
                'categories': category1,
                'values': values2,
                'y2_axis': True,
                'line':{'width':line_width}
            })
            ic()
        ic()

        chart1.combine(chart2)

        chart1.set_title({'name': title,
                          'name_font': {'size': title_font_size  # 원하는 폰트 크기 설정
                                        }
                          })
        chart1.set_x_axis({'name': col_bottom})
        chart1.set_y_axis({'name': left_axis_title,
                           'major_gridlines': {'visible': True}
                          } )

        ic()

        if right_y_axis_range is not None:
            chart2.set_y2_axis({
                'name': right_axis_title,  # 오른쪽 축 레이블
                'major_gridlines': {'visible': False},  # 오른쪽 축에 그리드라인 표시하지 않음
                'min':right_y_axis_range[0],
                'max':right_y_axis_range[1],

            })

        else:
            chart2.set_y2_axis({
                'name': right_axis_title,  # 오른쪽 축 레이블
                'major_gridlines': {'visible': False},  # 오른쪽 축에 그리드라인 표시하지 않음

            })

        chart1.set_legend({'none': False,
                            'position': 'top'})
        chart1.set_style(style_no)

        dict_scale1 = {'x_scale': 1.0, 'y_scale': 1.0, 'width': 640, 'height': 480}
        dict_scale = dict_scale1 if dict_scale is None else dict_scale
        if dict_scale:
            worksheet.insert_chart(pos_row, pos_col, chart1, dict_scale)
        else:
            worksheet.insert_chart(pos_row, pos_col, chart1)




    def chart_combined(self, col_x=None, col_left=None, col_right=None, title='',
                       pos_row=None, label_left=None,
                       label_right=None, pos_col=None, style_no=None,
                       line_width=2):
        """Insert combined chart into Excel sheet"""
        self.graph_no += 1
        df = self.df
        sheet_name = self.sheet_name
        if col_x is None:
            col_x = self.x_column

        if style_no is None:
            style_no = self.style_no

        # pos_row 자동 할당하기

        self.chart_position()
        pos_row = self.pos_row
        pos_col = self.pos_col


        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        label_left  = col_left if pd.isna(label_left) else label_left
        label_right = col_right if pd.isna(label_right) else label_right

        (max_row, max_col) = df.shape

        columns = df.columns.tolist()
        col_xi = columns.index(col_x)
        col_y1i = columns.index(col_left)
        col_y2i = columns.index(col_right)

        category1  = f"={sheet_name}!{xl_range_abs(1, col_xi, max_row, col_xi)}"
        values1 = f"={sheet_name}!{xl_range_abs(1, col_y1i, max_row, col_y1i)}"
        values2 = f"={sheet_name}!{xl_range_abs(1, col_y2i, max_row, col_y2i)}"

        chart1= workbook.add_chart({'type': 'column'})
        chart1.add_series({
            'name':label_left,
            'categories': category1,
            'values': values1,

        })

        chart2 = workbook.add_chart({'type': 'line'})
        chart2.add_series({
            'name': label_right,
            'categories': category1,
            'values': values2,
            'y2_axis': True,
            'line':{'width':line_width}

        })

        chart1.combine(chart2)

        chart1.set_title({'name': title})
        chart1.set_x_axis({'name': col_x})
        chart1.set_y_axis({'name': label_left,
                           'major_gridlines': {'visible': True}
                          } )


        chart2.set_y2_axis({
            'name': label_right,  # 오른쪽 축 레이블
            'major_gridlines': {'visible': False},  # 오른쪽 축에 그리드라인 표시하지 않음
        })

        chart1.set_legend({'none': False,
                                'position': 'top'})
        chart1.set_style(style_no)
        worksheet.insert_chart(pos_row, pos_col, chart1)


    def chart(self, col_x=None,

                  columns_list=[], col_begin=None, col_end=None, col_value_list=None,
                  rows_list=[], row_begin=None, row_end=None,
                  title='', title_font_size=10, auto_no_title=True,
                  label_left=None, label_right=None, label_bottom=None,
                  font_size=10, font_color=None,
                  name='',
                  pos_row=None, pos_col=None, chart_type='column',
                  subtype=None, style_no=None, precision=1, alpha=0,
                  legend_none=False,legend_font_size=None, legend_position='top',
                  col_error_bar = None,
                  dict_scale=None,
                  value_hide_value=True,
                  value_hide_category=True,
                  value_hide_percent=True,
                  value_hide_leader=True,
                  x_axis_range=None,
                  y_axis_range=None,
                  user_max_row:int=None,
                  data_kind = "L",
                  line_marker={},
              data_labels_tf=False,
              labels_font_size=8,
              num_format1='0.0%',
              num_format2='#,##0.0',
              verbose=False
              ):

        """Add a bar chart to the sheet
        col_x 에서 col_y 까지 컬럼 추출하여 그려줌
        """
        self.graph_no += 1
        #ic()
        title = f'Fig {self.graph_no}. ' + title if auto_no_title else title

        df = self.df
        sheet_name =  self.sheet_name
        if col_x is None:
            if data_kind == "L":
                col_x = self.x_column
            else:
                col_x = df.columns.tolist()[0]

        if verbose:
            print(f"col_x or col_bottom={col_x}")

        style_no = self.style_no if style_no is None else style_no

        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        #추가 : Graph 그릴 때 위에서 몇 개만 그릴 수 있게 헤주기 .(2024/11/16)
        # 사업명 그릴 때는 그릴 수가 없음
        (max_row, max_col) = df.shape
        max_row = max_row if user_max_row is None else user_max_row

        # col_end 입력하지 않으면 단일 컬럼 차트 그리기
        if col_end is None:
            col_end = col_begin

        self.chart_position()
        pos_row = self.pos_row
        pos_col = self.pos_col
        # subtype : 'clustered', 'stacked', 'percent_stacked'
        chart = workbook.add_chart({'type': chart_type, 'subtype': subtype})
        # 컬럼 인덱스 번호 정리
        columns = df.columns.tolist()
        columns_list_no  = []
        col_bottom = col_x
        col_bottom_i = columns.index(col_bottom) if col_bottom in columns else 0 # columns[0]
        # Check if col_begin is valid and determine its index
        if (col_begin is not None) and (col_begin in columns):
            col_begin_i = columns.index(col_begin)
        else:
            raise ValueError(f"Invalid column: {col_begin}")
        # Check if col_end is valid and determine its index
        if (col_end is not None) and (col_end in columns):
            col_end_i = columns.index(col_end)
        else:
            raise ValueError(f"Invalid column: {col_end}")
        # column 1 순위
        if len(columns_list) > 0:
            for each in columns_list:
                columns_list_no.append(columns.index(each))
            col_begin_i=columns_list_no[0]
            col_end_i=columns_list_no[-1]
        else: # 2순위, col_begin, col_end
            columns_list_no = list(range(col_begin_i, col_end_i + 1))

        # row 1 순위
        rows_list_no = []
        rows = df[col_x].tolist()
        if len(rows_list) > 0:
            for each in rows_list:
                rows_list_no.append(rows.index(each)+1)
            row_begin_i = rows_list_no[0]
            row_end_i = rows_list_no[-1]
        else:
            #2순위, col_begin, col_end
            row_begin_i = rows.index(row_begin) + 1 if row_begin is not None and row_begin in rows else 1
            row_end_i = rows.index(row_end) + 1 if row_end is not None and row_end in rows else 1
            rows_list_no = list(range(row_begin_i, row_end_i + 1))

        if verbose == True:
            print(f">>> verbose={verbose}, data_kind={data_kind}")
            print(f">>> columns_list_no = {columns_list_no}")
            print(f">>> rows_list_no = {rows_list_no}")
            print(f">>> *** ")

        col_error_bari = columns.index(col_error_bar) if col_error_bar is not None and col_error_bar in columns else None
        #print(f"* col_error_bari={col_error_bari}, {col_error_bar}")
        # 소수점 표현
        if precision == 1:
            number_format = '0.0%'
        elif precision == 2:
            number_format = '0.00%'
        else:
            number_format = '0%'
        # pie, doughnut에 사용할 내용임
        data_labels_pie_chart={ 'category':  False if value_hide_category else True ,  # 카테고리 이름 표시
                      'value': False if value_hide_value else True ,  # 값 표시
                      'font':{'size':font_size, 'color':font_color},
                      'percentage':  False if value_hide_percent else True,  # 퍼센트 표시
                      'num_format': '0.0%',
                      'separator': '\n',  # 레이블 사이에 구분자 설정
                     'leader_lines': False if value_hide_leader else True  # 리더 라인 (화살표 대체)
            }

        def fn_labels(kind='line', num_format1='0.0%', num_format2= '#,##0.00', labels_font_size=10):
            """레이블 리턴 함수
            :param kind:
            :return:
            """
            data_labels_pie_chart = {
                'category': not value_hide_category,  # 카테고리 이름 표시 여부
                'value': not value_hide_value,  # 값 표시 여부
                'font': {
                    'size': font_size or 10,  # 폰트 크기, 기본값 설정
                    'color': font_color or 'black',  # 폰트 색상, 기본값 설정
                },
                'percentage': not value_hide_percent,  # 퍼센트 표시 여부
                'num_format': num_format1, #'0.0%',  # 퍼센트 형식
                'separator': '\n',  # 레이블 간 구분자
                'leader_lines': not value_hide_leader,  # 리더 라인 표시 여부
            }
            data_labels_bar_chart = {'value': data_labels_tf, 'num_format': num_format2,'font': {'size': labels_font_size}}
            match kind:
                case 'line'|'column'|'bar':
                    #return {'value': data_labels_tf, 'num_format': '#,##0.0','font': {'size': labels_font_size}}
                    return data_labels_bar_chart
                case 'pie'|'doughnut':
                    return data_labels_pie_chart

        line_marker1 = {
            'type': 'square',
            'size': 8,
            'border': {'color': 'black'},
            'fill': {'color': 'red'},
        }

        if verbose:
            print(f"* col_begin_i={col_begin_i}, col_end_i={col_end_i}, row_begin_i={row_begin_i}, row_end_i={row_end_i}")

        match chart_type.lower():
            case 'bar'|'column'|'line'|'area'|'radar'|'scatter'|'pie'|'doughnut':
                if data_kind == 'L': # Long Type
                    for col_i in columns_list_no:
                        # error bar 설치
                        if col_error_bar is not None:
                            #print(
                            #    f"col_bottom_i={col_bottom_i}({columns[col_bottom_i]}), col_i={col_i}({columns[col_i]}), col_error_bari={col_error_bari}({columns[col_error_bari]})")
                            chart.add_series({
                                'name': [sheet_name, 0, col_bottom_i],
                                #'name': [sheet_name, 0, 0],
                                'categories': [sheet_name, row_begin_i, col_bottom_i, row_end_i, col_bottom_i],
                                'values': [sheet_name, row_begin_i, col_i, row_end_i, col_i],
                                #'num_format':num_format,
                                'y_error_bars': {
                                    'type': 'custom',
                                    'plus_values': [sheet_name, row_begin_i, col_error_bari, row_end_i, col_error_bari],
                                    'minus_values': [sheet_name, row_begin_i, col_error_bari, row_end_i, col_error_bari],
                                }
                            })
                        else:
                            #print(f">>> col_bottom={col_bottom}, col_bottom_i={col_bottom_i}, col_i={col_i}")
                            chart.add_series({
                                'name':[sheet_name, 0, col_i],
                                'categories': [sheet_name, row_begin_i, col_bottom_i, row_end_i, col_bottom_i], # 시트이름, 시작행번호, 시작열번호, 끝행번호, 끝열번호
                                'values': [sheet_name, row_begin_i, col_i, row_end_i, col_i],       # 시트이름, 시작행번호, 시작열번호, 끝행번호, 끝열번호
                                'fill': {'transparency': alpha},
                                'data_labels': fn_labels(chart_type, num_format1=num_format1, num_format2=num_format2, labels_font_size=labels_font_size),
                            })
                elif data_kind == 'W': # Wide Type
                    #for row_i in range(1, max_row+1):
                    for row_i in rows_list_no:
                        if verbose:
                            print(f">>> row_i = {row_i}, col_begin_i={col_begin_i}, col_end_i={col_end_i}")
                        # error bar 설치
                        if col_error_bar is not None:
                            pass
                            # print(
                            #    f"col_bottom_i={col_bottom_i}({columns[col_bottom_i]}), col_i={col_i}({columns[col_i]}), col_error_bari={col_error_bari}({columns[col_error_bari]})")
                        else:
                            chart.add_series({
                                'name': [sheet_name, row_i, 0],
                                'categories': [sheet_name, 0, col_begin_i, 0, col_end_i],
                                # 시트이름, 시작행번호, 시작열번호, 끝행번호, 끝열번호
                                'values': [sheet_name, row_i, col_begin_i, row_i, col_end_i],  # 시트이름, 시작행번호, 시작열번호, 끝행번호, 끝열번호
                                'fill': {'transparency': alpha},
                                'data_labels': fn_labels(chart_type, num_format1=num_format1, num_format2=num_format2, labels_font_size=labels_font_size),
                                'marker':line_marker if line_marker else {}
                            })


            case _ :
                print('* waring ! : chart_type is not defined correctly !')
                return

        chart.set_style(style_no)
        chart.set_title({'name': title,  'name_font':{'size': title_font_size}})

        if x_axis_range is not None:
            chart.set_x_axis({'name': col_x if label_bottom is None else label_bottom, 'min':x_axis_range[0], 'max':x_axis_range[1] })
        else:
            chart.set_x_axis({'name': col_x if label_bottom is None else label_bottom })

        if y_axis_range is not None:
            chart.set_y_axis({'name': "" if label_left is None else label_left, 'min':y_axis_range[0], 'max':y_axis_range[1] })
        else:
            chart.set_y_axis({'name': "" if label_left is None else label_left, })






        if subtype == 'percent_stacked' and chart_type=='column':
            chart.set_y_axis({
                'min': 0,
                'max': 1,  # 1 = 100%
                'num_format': '0%',  # 퍼센트 형식
                'major_gridlines': {'visible': True},  # 주요 그리드라인 표시
            })

        chart.set_legend({'none': legend_none,
                           'font':{'size': legend_font_size},
                          'position': legend_position})

        dict_scale1 = {'x_scale': 1.0, 'y_scale': 1.0, 'width': 640, 'height': 480}
        dict_scale = dict_scale1 if dict_scale is None else dict_scale
        if dict_scale:
            worksheet.insert_chart(pos_row, pos_col, chart, dict_scale)
        else:
            worksheet.insert_chart(pos_row, pos_col, chart)



    def set_x(self, name=None):
        if name is not None:
            self.x_column = name
            #print(f"* x column is set to be {name}")

    def set_width(self, w=3):
        self.w = w

    def set_settings(self, x_column=None, w=2, left_gap=8, style_no=None, graph_no=0):
        if x_column is not None:
            self.set_x(name=x_column)
        self.w=w # 횡축 차트수
        self.pos_col_initial=left_gap
        self.style_no=style_no if style_no is not None else self.style_no
        self.graph_no=graph_no
        print(f"*settings : style_no={self.style_no}")

    def chart_position(self):
        # pos_row 자동 할당하기
        self.pos_row = 1 + (self.graph_no-1) // self.w * self.pos_row_delta + self.pos_row_initial
        self.pos_col = 1 + (self.graph_no-1) % self.w  * self.pos_col_delta + self.pos_col_initial
        #print(self.pos_row, self.pos_col)

    def chart_total(self, df:DataFrame=None,
                    w:int=3,
                    style_no:int=None,
                    title_font_size:int=12,
                    dic_precison={},
                    sheet_no=0,
                    col_bottom:str=None,
                    rows_list:list=None,
                    col_begin=None,
                    col_end=None,
                    row_left_list=None,
                    row_right_list = None,

                    col_pie_list = None,
                    col_x = None,
                    col_y = None,
                    col_size = None,
                    fixed_node_size_sc=10,
                    rows_list_sc = None,

                    left_name_list = None,
                    right_name_list = None,

                    # line etc (line)
                    left_axis_title_line=None,
                    bottom_axis_title_line=None,

                    # combined graph (cb)
                    left_axis_title_cb = None,
                    right_axis_title_cb = None,
                    bottom_axis_title_cb=None,
                    left_chart_style='column',
                    right_chart_style='line',

                    # evolution pathway (ep)
                    max_range=None,
                    min_range=None,
                    col_index=None,

                    evolution_pathway_lw=4,

                    row_x=None,
                    row_y=None,
                    row_size=None,
                    left_axis_title_ep=None,
                    bottom_axis_title_ep=None,
                    fixed_node_size_ep=None,

                    tf_pie=True,
                    tf_doughnut = True,
                    tf_column = True,
                    tf_line = True,
                    tf_bar = True,
                    tf_area = True,
                    tf_radar = True,
                    tf_scatter1 = True,
                    tf_scatter2 = False,
                    tf_set1 = True,
                    tf_combined = False,
                    tf_evolution = False,


                    sheet_name = None,
    ):
        """전체 그래프 그리기"""
        col_pie_list = self.col_pie_list if col_pie_list is None else col_pie_list
        rows_list = self.rows_list if rows_list is None else rows_list
        col_begin = self.col_begin if col_begin is None else col_begin
        col_end = self.col_end if col_end is None else col_end

        if style_no is not None:
            self.style_no = style_no


        if tf_set1 :
            pass
        else:
            tf_pie=False
            tf_doughnut=False
            tf_column=False
            tf_line=False
            tf_bar=False
            tf_area=False
            tf_radar=False
            tf_scatter1=False
            #tf_scatter2=False


        col_bottom = col_bottom if col_bottom is not None else df.columns.tolist()[0]

        columns_all = df.columns.tolist()
        rows_list1 = df[col_bottom].tolist()
        rows_list = rows_list if rows_list is not None else rows_list1
        col_begin = col_begin if col_begin is not None else columns_all[1]
        col_end = col_end if col_end is not None else columns_all[-1]

        row_left_list = row_left_list if row_left_list is not None else [rows_list[0]]
        row_right_list = row_right_list if row_right_list is not None else [rows_list[1]]
        row_x = row_x if row_x is not None else rows_list[0]
        row_y = row_y if row_y is not None else rows_list[1]
        row_size = row_size if row_size is not None else rows_list[0]

        col_pie_list = col_pie_list if col_pie_list is not None else 'sum'

        # scatter2 : Long type
        col_x = col_x if col_x is not None else columns_all[-2]
        col_y = col_y if col_y is not None else columns_all[-1]
        col_size = col_size if col_size is not None else columns_all[-1]

        bottom_axis_title_cb =bottom_axis_title_cb if bottom_axis_title_cb is not None else 'bottom_axis_title'
        left_name_list = left_name_list if left_name_list is not None else row_left_list
        right_name_list =right_name_list if right_name_list is not None else row_right_list

        left_axis_title_line = 'left_axis_title_line' if left_axis_title_line is None else left_axis_title_line
        bottom_axis_title_line = 'bottom_axis_title_line' if bottom_axis_title_line is None else bottom_axis_title_line
        max_range = max_range if max_range is not None else 30
        min_range = min_range if min_range is not None else 10

        if len(df.index)>0:
            if True:
                df_each = df
                dic_precision = fn_dictionary_precision(df_each)
                dic_width = {'name': 40}
                sheet_no += 1
                sheet_name = sheet_name if sheet_name is not None else f"Table{sheet_no}"
                self.set_settings(x_column=col_bottom, w=w, left_gap=find_left_gap(df_each), style_no=style_no)
                self.to_sheet(df=df_each, sheet_name=sheet_name, dic_precision=dic_precision, dic_width=dic_width,
                            fixed_width=8)

                chart_func = partial(self.chart,
                                     col_x=col_bottom, col_begin=col_begin, col_end=col_end,
                                     rows_list=rows_list,
                                     title=f"Test ",
                                     legend_none=False, title_font_size=title_font_size,
                                     label_bottom=bottom_axis_title_line, label_left=left_axis_title_line,
                                     data_kind='W',
                                     data_labels_tf=True,
                                     verbose=False)


                if len(col_pie_list) > 0 :
                    for col_sum in col_pie_list:
                        if col_sum in df_each.columns:
                            if tf_pie:
                                chart_func(title=f"pie '{col_sum}'-L", chart_type='pie', col_begin=col_sum, col_end=col_sum,
                                           value_hide_value=False, value_hide_percent=False, data_kind='L')
                                chart_func(title=f"column '{col_sum}'-L", chart_type='column', col_begin=col_sum, col_end=col_sum, label_bottom='Category', label_left=left_axis_title_line,
                                           value_hide_value=False, value_hide_percent=False,data_kind='L')
                                chart_func(title=f"bar '{col_sum}'-L", chart_type='bar', col_begin=col_sum, col_end=col_sum, label_left='Category', label_bottom='Category',
                                           data_kind='L')

                                chart_func(title=f"pie '{col_sum}'-L", chart_type='pie', col_begin=col_sum, col_end=col_sum,
                                           value_hide_value=True, value_hide_percent=False, data_kind='L')

                                chart_func(title=f"column '{col_sum}'-W", chart_type='column', col_begin=col_sum, col_end=col_sum,
                                           label_bottom='Category', label_left=left_axis_title_line,
                                           value_hide_value=False, value_hide_percent=False, data_kind='W')
                                chart_func(title=f"bar '{col_sum}'-W", chart_type='bar', col_begin=col_sum, col_end=col_sum,
                                           label_left='Category', label_bottom='Category',
                                           data_kind='W')

                                # chart_func(title=f"pie3", chart_type='pie', col_begin=col_sum, col_end=col_sum,
                                #            value_hide_value=True, value_hide_percent=False, data_kind='L')
                            if tf_doughnut:
                                #if (col_sum is not None) and (col_sum in df_each.columns):
                                #print(f">>> tf_doughnut={tf_doughnut}")
                                chart_func(title=f"doughnut '{col_sum}'-L", chart_type='doughnut', col_begin=col_sum, col_end=col_sum,
                                               value_hide_value=False, value_hide_percent=False, data_kind='L')
                                chart_func(title=f"doughnut '{col_sum}'-W", chart_type='doughnut', data_kind='W')
                                chart_func(title=f"doughnut '{col_sum}'-L", chart_type='doughnut', data_kind='L')

                graph_type_list = []
                if tf_column:
                    graph_type_list.append('column')
                if tf_bar:
                    graph_type_list.append('bar')
                if tf_area:
                    graph_type_list.append('area')
                if tf_line:
                    graph_type_list.append('line')

                if len(graph_type_list) > 0:
                    for each_kind in graph_type_list:
                        if each_kind in ['bar']:
                            bottom_axis_title_temp = left_axis_title_line
                            left_axis_title_temp = bottom_axis_title_line
                            chart_func(title=f"{each_kind}-W", chart_type=each_kind, label_bottom=bottom_axis_title_temp, label_left=left_axis_title_temp)
                            for each_sub in ['stacked', 'percent_stacked']:
                                chart_func(title=f"{each_kind}_{each_sub}-W", chart_type=each_kind, subtype=each_sub, label_bottom=bottom_axis_title_temp, label_left=left_axis_title_temp)
                        else:
                            chart_func(title=f"{each_kind}-W", chart_type=each_kind)
                            for each_sub in ['stacked', 'percent_stacked']:
                                chart_func(title=f"{each_kind}_{each_sub}-W", chart_type=each_kind, subtype=each_sub)
                        if each_kind == 'line':
                            line_marker = {'type': 'circle', 'size': 10, 'fill': {'color': '#ffffff'}}
                            chart_func(title=f"line2 -W", chart_type="line", line_marker=line_marker, data_labels_tf=True)
                            for each_sub2 in ['stacked', 'percent_stacked']:
                                chart_func(title=f"{each_kind}_{each_sub}-W",
                                           chart_type=each_kind,
                                           subtype=each_sub2,
                                           line_marker=line_marker,
                                           data_labels_tf=True,
                                           value_hide_value=False)


                if tf_scatter1:
                    # chart_func(title=f"scatter1", chart_type='scatter',)
                    chart_func(title=f"scatterI-1 'straight_with_markers", chart_type='scatter', subtype='straight_with_markers')
                    # chart_func(title=f"scatter3", chart_type='scatter', subtype='straight')
                    chart_func(title=f"scatterI-2 'smooth_with_markers", chart_type='scatter', subtype='smooth_with_markers')
                    chart_func(title=f"scatterI-3 'smooth'", chart_type='scatter', subtype='smooth')

                if tf_scatter2:
                    #print(f"*1 tf_scatter2={tf_scatter2} : col_x={col_x}, col_y={col_y}, col_size={col_size}")
                    condition1 = col_x in df_each.columns and col_y in df_each.columns and col_size in df_each.columns
                    #print(f"condition={condition1}")
                    if condition1:
                        rows_list_sc = rows_list_sc if rows_list_sc is not None else rows_list
                        graph_ep2 = partial(self.chart_scatter_beta0,
                                            # col_name='항목',
                                            style_no=style_no,
                                            min_range=min_range, max_range=max_range,
                                            rows_list=rows_list_sc)
                        graph_ep2(title=f"Graph scatterII - 1 style_no={style_no}, size=f\'{col_size}\'", col_x=col_x, col_y=col_y, col_size=col_size)
                        graph_ep2(title=f"Graph scatterII - 2 style_no={style_no}, size={fixed_node_size_sc}", col_x=col_x, col_y=col_y, fixed_node_size=fixed_node_size_sc)
                        graph_ep2(title=f"Graph scatterII - 3 style_no=26, size=f\'{col_size}'",         col_x=col_x, col_y=col_y, col_size=col_size, style_no=26)

                        if False:
                            print(f"*2* col_x={col_x}, col_y={col_y}, col_size={col_size}")

                if tf_radar:
                    chart_func(title=f"radar1", chart_type='radar')
                    chart_func(title=f"radar2", chart_type='radar', subtype='with_markers')
                    chart_func(title=f"radar3", chart_type='radar', subtype='filled')

                if tf_combined :  # combined graph
                    line_marker = {'type': 'circle', 'size': 10, 'fill': {'color': '#ffffff'}}
                    #print(f"* row_left_list={row_left_list}, row_right_list={row_right_list}")
                    graph_combined = partial(self.chart_combined_v3,
                                             # col_bottom='항목',
                                             row_left_list=row_left_list,
                                             row_right_list=row_right_list,
                                             col_begin=col_begin,
                                             col_end=col_end,
                                             title='Combined (W)',
                                             title_font_size=title_font_size,
                                             left_name_list=left_name_list,  # ['항목1 이름'],
                                             right_name_list=right_name_list,  # ['항목 2 이름'],
                                             left_axis_title=left_axis_title_cb,  # "항목 1 왼쪽 축 타이틀",
                                             right_axis_title=right_axis_title_cb,  # "항목2 오른쪽 축 타이틀",
                                             bottom_axis_title=bottom_axis_title_cb,
                                             left_chart_style=left_chart_style,
                                             right_chart_style=right_chart_style,
                                             data_kind='W',
                                             style_no=style_no)
                    # graph_combined(title='Combined (W)', data_labels_right_tf=False,data_labels_left_tf=True, line_marker=line_marker)
                    graph_combined(title='Combined (W1)', data_labels_right_tf=True, data_labels_left_tf=True,
                                   line_marker=line_marker)
                    graph_combined(title='Combined (W2)', data_labels_right_tf=False, data_labels_left_tf=False,
                                   line_marker=line_marker)
                    graph_combined(title='Combined (W3)', data_labels_right_tf=True, data_labels_left_tf=False, left_chart_style='line',
                                   line_marker=line_marker)

                if tf_evolution:  # Evolution Pathway
                    lw = 3
                    graph_ep = partial(self.chart_scatter_beta0_W,
                                       col_index=col_index,
                                       col_begin=col_begin, col_end=col_end,
                                       max_range=max_range,
                                       min_range=min_range,
                                       # title=f"Sample, lw={lw}",
                                       left_axis_title=left_axis_title_ep,
                                       bottom_axis_title=bottom_axis_title_ep,
                                       line=True,
                                       line_width=lw,
                                       line_color="#749040",
                                       style_no=style_no)
                    graph_ep(title=f"Evolution Pathway, lw={lw}, style no={self.style_no} node_size={fixed_node_size_ep}", row_x=row_x, row_y=row_y,
                             row_size=row_size, fixed_node_size=fixed_node_size_ep)
                    graph_ep(title=f"Evolution Pathway, lw={lw}, style no={self.style_no} node_size=12)", row_x=row_x, row_y=row_y,
                             row_size=row_size, fixed_node_size=12)
                    graph_ep(title=f"Evolution Pathway, lw={lw}, style no={self.style_no} node_size={col_size}, ({min_range}, {max_range})", row_x=row_x,
                             row_y=row_y,                row_size=row_size)
                    lw = evolution_pathway_lw
                    for each_style_no, each_min, each_max in zip([27, 28, 29],[10,20,30], [30,30,30]):
                        graph_ep(title=f"Evolutonn Pathway, lw={lw}, style_no={each_style_no}, ({each_min}-{each_max})", row_x=row_y,
                                 row_y=row_x, row_size=row_size, style_no=each_style_no, max_range=each_max, min_range=each_min)
        # 내부 변수 지정
        self.col_pie_list = col_pie_list
        self.rows_list = rows_list
        self.col_begin = col_begin
        self.col_end = col_end

# 2024/11/18 added
def make_top_years_table(df:DataFrame, col_category:str, col_PY:str, col_FUND:str, agg_method='sum'):
    dfb = df.groupby([col_category, col_PY])[col_FUND].agg(agg_method).unstack(fill_value=0)
    dfb['sum'] = dfb.sum(axis=1)

    # 상위 10개만 추출
    dfb['rank'] = dfb['sum'].rank(ascending=False, method='dense')
    dfb['percentile'] = 100 * dfb['sum'].rank(pct=True, ascending=False)
    dfb['share'] = 100 * dfb['sum'] / dfb['sum'].sum()
    dfb.sort_values(by='sum', ascending=False, inplace=True)
    dfb.reset_index(inplace=True)
    return dfb


# 2024/11/11 added
def make_table_no1(df:DataFrame=None, col_left:str='PY', col_target:str=None, data_kind='L')->DataFrame:
    print(f">>> make_table_no1(df, col_left, col_target)")
    if df is not None:
        df = df.copy()
        df_t1 = df.groupby([col_left], observed=True)[col_target].agg(['max',
                                                                  'mean', 'median', 'min',
                                                                  ('1/4q', lambda x: x.quantile(0.25)),
                                                                  ('2/4q', lambda x: x.quantile(0.5)),
                                                                  ('3/4q', lambda x: x.quantile(0.75)),
                                                                  ('99%', lambda x: x.quantile(0.99)),
                                                                  ('95%', lambda x: x.quantile(0.95)),
                                                                  ('90%', lambda x: x.quantile(0.90)),
                                                                  ('85%', lambda x: x.quantile(0.85)),
                                                                  ('80%', lambda x: x.quantile(0.80)),
                                                                  ('75%', lambda x: x.quantile(0.75)),
                                                                  ('70%', lambda x: x.quantile(0.70)),
                                                                  ('65%', lambda x: x.quantile(0.65)),
                                                                  ('60%', lambda x: x.quantile(0.60)),
                                                                  ('55%', lambda x: x.quantile(0.55)),
                                                                  ('50%', lambda x: x.quantile(0.50)),
                                                                  ('45%', lambda x: x.quantile(0.45)),
                                                                  ('40%', lambda x: x.quantile(0.40)),
                                                                  ('35%', lambda x: x.quantile(0.35)),
                                                                  ('30%', lambda x: x.quantile(0.30)),
                                                                  ('25%', lambda x: x.quantile(0.25)),
                                                                  ('20%', lambda x: x.quantile(0.20)),
                                                                  ('15%', lambda x: x.quantile(0.15)),
                                                                  ('10%', lambda x: x.quantile(0.10)),
                                                                  ('5%', lambda x: x.quantile(0.05)),
                                                                  'sum',
                                                                  'std',
                                                                  'count'])
    df_t1.reset_index(inplace=True)
    df_t1[col_left] = df_t1[col_left].astype(int)
    df_t1['ratio_fund'] = 100 * df_t1['sum'] / df_t1['sum'].sum()
    df_t1['ratio_count'] = 100 * df_t1['count'] / df_t1['count'].sum()
    # ranking, percentile added
    df_t1['rank'] = df_t1['sum'].rank(ascending=False, method='dense')
    df_t1['percentile'] = 100 * df_t1['sum'].rank(pct=True, ascending=False)

    if data_kind == 'W':
        df_t1.set_index(col_left, inplace=True)
        df_t1 = df_t1.T.reset_index()
        df_t1.rename(columns={'index':col_left}, inplace=True)

    return df_t1








def make_other(df: DataFrame = None,
                     col_threshold: str = 'ratio_fund',
                     threshold: float = 5,
                     col_value:dict=None, by:str='upper', sorting:bool=False) -> DataFrame:
    # 2024/11/11
    # col_value = {'PY':'Other'}

    df = df.copy()
    columns = df.columns.tolist()
    if col_value is None:
        col_value = {columns[0]:'Other'}

    if col_threshold not in df.columns:
        raise Exception("col_threshold must be your columns")
    # 임계값 이상과 미만으로 구분
    match by:
        case 'upper':
            above_threshold = df[df[col_threshold] >= threshold]
            below_threshold = df[df[col_threshold] < threshold]
        case 'lower':
            above_threshold = df[df[col_threshold] <= threshold]
            below_threshold = df[df[col_threshold] > threshold]
        case _ :
            raise ValueError("'by'는 'upper' 또는 'lower'만 가능합니다.")
    # 임계값 미만의 값들을 'Others'로 묶기
    if not below_threshold.empty:
        others_sum_temp = []
        for col in df.columns:
            if col in col_value.keys():
                value = col_value[col]
            else:
                 value = below_threshold[col].sum()
            others_sum_temp.append(value)
        others_row = DataFrame([others_sum_temp], columns=df.columns, index=['Others'])
        # 임계값 이상인 값들을 큰 값 순으로 정렬
        if sorting:
            above_threshold_sorted = above_threshold.sort_values(by=[col_threshold], ascending=False)
        else:
            above_threshold_sorted = above_threshold
        df_final = pd.concat([above_threshold_sorted, others_row], ignore_index=False)
    else:
        df_final = above_threshold.sort_values(by=[col_threshold], ascending=False)

    df_final.index = pd.RangeIndex(len(df_final.index))
    # 순위 수정
    #df_final.loc['other','rank']=len(df_final.index)

    return df_final







# 2024/10/30 added


def make_graph_fund(df, filename='temp1.xlsx', w=4,
                    col_bottom1=None, col_left1='', col_right_list1=[], label_left1=None, label_right1=None, left_axis_title1=None, right_axis_title1='', title1=None, figure1=True,
                    col_bottom2=None, col_left2='', col_right_list2=[], label_left2=None, label_right2=None, left_axis_title2=None, right_axis_title2='', title2=None, figure2=True,
                    col_x=None, col_y=None, col_y2=None, col_size=None, col_name=None, label_right=None, title3=None, title4=None, figure3=True,
                    style_no=26, title_font_size=10,
                    dic_precision={}):
    # settings
    try:
        print(f"col_x={col_x}")
        filename1 = generate_filename_with_timestamp(filename)

        col_name = col_x if col_name is None else col_name
        col_size = col_y if col_size is None else col_size
        title1 = f"{col_left1} vs {col_right_list1}" if title1 is None else title1
        title2 = f"{col_left2} vs {col_right_list2}" if title2 is None else title2
        title3 = f"{col_x} vs {col_y}, size={col_size}" if title3 is None else title3
        title4 = f"{col_x} vs {col_y2}, size={col_size}" if title4 is None else title4

        left_axis_title1 = col_left1 if left_axis_title1 is None else left_axis_title1
        left_axis_title2 = col_left2 if left_axis_title2 is None else left_axis_title2



        ex = BKExcelWriter(save_file_name=filename1)

        ex.set_settings(x_column=col_bottom1, w=w, left_gap=len(df.columns), style_no=style_no)
        ex.to_sheet(df=df, sheet_name=f"Sheet1", dic_precision=dic_precision)

        if figure1 == True:
            ex.chart_combined_v2(col_left=col_left1, col_right_list=col_right_list1, title=title1,
                                 title_font_size=title_font_size, left_name=label_left1, right_name=label_right1,
                                 left_axis_title=left_axis_title1, right_axis_title=right_axis_title1)

        if figure2 == True:
            col_bottom2 = col_bottom1 if pd.isna(col_bottom2) else col_bottom2
            col_left2 = col_left1 if pd.isna(col_left2) else col_left2
            ex.set_settings(x_column=col_bottom2, w=w, left_gap=len(df.columns), style_no=style_no)
            ex.chart_combined_v2(col_left=col_left2, col_right_list=col_right_list2, title=title2,
                                 title_font_size=title_font_size, left_name=label_left2, right_name=label_right2,
                                 left_axis_title=left_axis_title2, right_axis_title=right_axis_title2)

        # scatter chart
        if figure3 == True:
            if col_x is not None:
                col_y2 = col_y if col_y2 is None else col_y2

                ex.chart_scatter(col_x=col_x, col_y=col_y, col_size=col_size, col_name=col_name, title=title3, title_font_size=title_font_size)
                ex.chart_scatter(col_x=col_y2, col_y=col_y, col_size=col_size, col_name=col_name, title=title4, title_font_size=title_font_size)

        ex.close()

        print(f"* file saved successfully : {filename1}")

    except Exception as e:
        print(f"*** error = {e}")


def fn_find_col_i(df, col):
    # Check if col is valid and determine its index
    columns = df.columns.tolist()
    if (col is not None) and (col in columns):
        col_i = columns.index(col)
    else:
        raise ValueError(f"Invalid column: {col}")
    return col_i

def fn_find_col_index(columns:list, col:str)-> int:
    # Check if col is valid and determine its index
    if (col is not None) and (col in columns):
        col_i = columns.index(col)
    else:
        raise ValueError(f"Invalid column: {col}")
    return col_i


def fn_dictionary_precision(df:DataFrame, stop_columns=['PY','count']):
    """정밀도 사전"""
    dic_precision = {}
    for col in df.columns:
        if col in stop_columns:
            continue
        dic_precision[col] = 0

    # 사전 업데이트를 한 번에 처리
    dic_precision.update({
        'sum': 0,
        'count': 0,
        'mean': 1,
        'median': 1,
        'std': 1,
        'ratio_fund': 2,
        'ratio_count': 2,
        'ratio1':2,
        'ratio2':2,
        'r_new':3
    })
    return dic_precision

def find_left_gap(df:DataFrame=None, no=None) -> int:
    if no is None:
        left_gap=len(df.columns)
    else:
        left_gap = len(df.columns) if len(df.columns) < no else no
    return left_gap

def generate_filename_with_timestamp(original_filename):
    # 파일명과 확장자를 분리
    base_name, extension = os.path.splitext(original_filename)
    # 현재 시간을 "YYYYMMDD_HHMMSS" 형식으로 가져오기
    current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    # 파일 이름에 시간 정보를 추가하여 반환
    return f"{base_name}_{current_time}{extension}"

if __name__ == '__main__':
    ex = help(BKExcelWriter)
    print(ex)
    print(__name__)
