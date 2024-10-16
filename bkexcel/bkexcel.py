# bkexcel.py
# 2024.10.09

import pandas as pd
import numpy as np
from pandas.core.interchange.dataframe_protocol import DataFrame
from unicodedata import category

from xlsxwriter.utility import xl_rowcol_to_cell, xl_range, xl_range_abs

class BKExcelWriter:
    def __init__(self, writer=None, save_file_name=None, engine='xlsxwriter', sheet_name=None):
        self.writer = writer or pd.ExcelWriter(save_file_name, engine=engine)
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

    def to_sheet(self, df:DataFrame=None, sheet_name='Sheet1', dic_width={'논문수': 8, '총피인용수': 10},
                 dic_color={}, dic_precision={}, freeze_row=1, col_con1=None, col_con2=None):
        """Save dataframe to Excel with formatting"""

        self.df = df
        self.sheet_name = sheet_name

        df.replace(np.nan, None, inplace=True)
        df.index = pd.RangeIndex(1, len(df.index) + 1)
        df.to_excel(self.writer, sheet_name=sheet_name, index=False)

        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "center",
            "align": "center",
            "fg_color": "#D7E4BC",
            "border": 1,
        })

        color_format = workbook.add_format({
            "bold": False,
            "text_wrap": True,
            "valign": "center",
            "align": "center",
            "fg_color": "#eeeeee",
            "border": 1,
        })

        columns = df.columns.tolist()

        for j in range(len(columns)):
            try:
                worksheet.write(0, j, columns[j], header_format)
            except Exception as e:
                print(f"Error writing header: {e}")

        #worksheet.set_column(0, 0, 8)
        worksheet.set_column(0, len(df.columns)-1, 15) # 첫번재 열 부터 마지막 열까지 폭 지정, 0 부터 시작함

        for k, v in dic_width.items():
            if k in columns:
                pi = columns.index(k)
                worksheet.set_column(pi, pi, v)

        for k, v in dic_color.items():
            if k in columns:
                pi = columns.index(k)
                color_format1 = workbook.add_format({'fg_color': v, 'text_wrap': True})
                for i in range(len(df.index)):
                    worksheet.write(i+1, pi, df.iloc[i, columns.index(k)], color_format1)

        if col_con1 and col_con2:
            format_condition = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
            length_1 = len(df.index) + 1
            worksheet.conditional_format(f"{col_con1}2:{col_con2}{length_1}",
                                         {"type": "cell", "criteria": ">", "value": 0, "format": format_condition})

        for k, v in dic_precision.items():
            if k in columns:
                pi = columns.index(k)
                if v == 2:
                    precision_format = workbook.add_format({'num_format': '#,##0.00'})
                elif v == 1:
                    precision_format = workbook.add_format({'num_format': '#,##0.0'})
                elif v == 0:
                    precision_format = workbook.add_format({'num_format': '#,##0'})
                else:
                    precision_format = workbook.add_format({'num_format': '#,##0.00'})
                for i in range(len(df.index)):
                    worksheet.write(i+1, pi, df.iloc[i, columns.index(k)], precision_format)

        worksheet.freeze_panes(freeze_row, 0)

    def close(self):
        self.writer.close()

    def chart_scatter(self, col_x=None, col_y=None, col_size=None,
                      col_name=None, title=None, pos_row=None, pos_col=None, style_no=None,
                      fixed_node_size=10):
        """Insert scatter chart into Excel sheet"""
        self.graph_no += 1

        df = self.df
        sheet_name = self.sheet_name

        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        columns = df.columns.tolist()
        if col_x is not None:
            xi = columns.index(col_x)
        else:
            xi = 0
        if col_y is not None:
            yi = columns.index(col_y)
        else:
            yi = 1
        namei = columns.index(col_name)

        if style_no is None:
            style_no = self.style_no

        self.chart_position()
        pos_row = self.pos_row
        pos_col = self.pos_col

        chart = workbook.add_chart({'type': 'scatter'})

        if col_size is None:
            pass
        else:
            col_size_list = df[col_size].tolist()

        for i in range(1, len(df.index)+1): # 엑셀 행은 1부터 시작하기에.
            chart.add_series({
                'name': [sheet_name, i, namei], #
                'categories': [sheet_name, i, xi, i, xi], # x 축
                'values': [sheet_name, i, yi, i, yi], # y 축
                'marker': {
                    'type': 'circle',
                    'size': fixed_node_size,  # 사이즈를 크기 값으로 설정
                },
            })

        chart.set_title({'name': title})
        chart.set_x_axis({'name': col_x})
        chart.set_y_axis({'name': col_y})
        chart.set_style(style_no)

        worksheet.insert_chart(pos_row, pos_col, chart)

    def chart_combined(self, col_x=None, col_left=None, col_right=None, title='',
                       pos_row=None, label_left=None,
                       label_right=None, pos_col=None, style_no=None):
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

    def chart_bar_old(self, col_x=None, col_y=None, title='', name='', pos_row=1, pos_col=3, style_no=11):
        """Add a bar chart to the sheet"""

        df = self.df
        sheet_name =  self.sheet_name

        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        (max_row, max_col) = df.shape

        columns = df.columns.tolist()
        col_xi = columns.index(col_x)
        col_yi = columns.index(col_y)

        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({
            'name':[sheet_name, 1, col_xi, 1, col_xi],
            'categories': [sheet_name, 1, col_xi, max_row, col_xi],
            'values': [sheet_name, 1, col_yi, max_row, col_yi],
        })
        chart.set_title({'name': title})
        chart.set_legend({'none': True, 'position': 'top'})
        chart.set_style(style_no)

        worksheet.insert_chart(pos_row, pos_col, chart)

    def chart(self, col_x=None, columns_list=[], col_begin=None, col_end=None, col_value_list=None, title='', name='',
                  pos_row=None, pos_col=None, chart_type='column',
                  subtype=None, style_no=None, precision=1, alpha=0):
        """Add a bar chart to the sheet
        col_x 에서 col_y 까지 컬럼 추출하여 그려줌
        """
        self.graph_no += 1
        df = self.df
        sheet_name =  self.sheet_name
        if col_x is None:
            col_x = self.x_column
        style_no = self.style_no if style_no is None else style_no

        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        (max_row, max_col) = df.shape

        # col_end 입력하지 않으면 단일 컬럼 차트 그리기
        if col_end is None:
            col_end = col_begin



        self.chart_position()
        pos_row = self.pos_row
        pos_col = self.pos_col

        chart = workbook.add_chart({'type': chart_type, 'subtype': subtype})
        columns = df.columns.tolist()

        columns_list_no  = []
        col_xi = columns.index(col_x)

        # 소수점 표현
        if precision == 1:
            number_format = '0.0%'
        elif precision == 2:
            number_format = '0.00%'
        else:
            number_format = '0%'

        match chart_type.lower():
            case 'bar'|'column'|'line'|'b'|'c'|'l'|'area'|'a'|'radar'|'r'|'scatter'|'s'|'pie'|'p'|'doughnut'|'d':
                if len(columns_list)>0:
                    for each in columns_list:
                        columns_list_no.append(columns.index(each))
                    # for col_i in columns_list_no:
                    #     chart.add_series({
                    #         'name':[sheet_name, 0, col_i],
                    #         'categories': [sheet_name, 1, col_xi, max_row, col_xi],
                    #         'values': [sheet_name, 1, col_i, max_row, col_i],
                    #     })
                else:
                    col_1 = columns.index(col_begin)
                    col_2 = columns.index(col_end)
                    columns_list_no=list(range(col_1, col_2+1))

                for col_i in columns_list_no:
                    chart.add_series({
                        'name':[sheet_name, 0, col_i],
                        'categories': [sheet_name, 1, col_xi, max_row, col_xi],
                        'values': [sheet_name, 1, col_i, max_row, col_i],
                        'fill': {'transparency': alpha}
                    })

            # case 'pie'|'p'|'doughnut'|'d':
            #
            #     for col_value in col_value_list:
            #         col_vi = columns.index(col_value)
            #         # 차트에 데이터 추가
            #         chart.add_series({
            #             'name': [sheet_name, 0, col_xi],
            #             'categories': [sheet_name, 1, col_xi, max_row, col_xi],
            #             'values': [sheet_name, 1, col_vi, max_row, col_vi],
            #             'data_labels': {'percentage': True,
            #                             'number_format': number_format,
            #                                 },  # 백분율 표시
            #             })
            case _ :
                print('* waring ! : chart_type is not defined correctly !')
                return







        chart.set_style(style_no)
        chart.set_title({'name': title})
        chart.set_legend({'none': False, 'position': 'top'})

        if subtype == 'percent_stacked' and chart_type=='column':
            chart.set_y_axis({
                'min': 0,
                'max': 1,  # 1 = 100%
                'num_format': '0%',  # 퍼센트 형식
                'major_gridlines': {'visible': True},  # 주요 그리드라인 표시
            })
        worksheet.insert_chart(pos_row, pos_col, chart)

    def chart_line_old(self, col_x=None, col_begin=None, col_end=None, title='', name='',
                   pos_row=None, pos_col=None, subtype=None, style_no=None):
        """Add a bar chart to the sheet"""
        #print(">>> chart_line_multi()")
        self.graph_no += 1
        df = self.df
        sheet_name =  self.sheet_name
        if col_x is None:
            col_x = self.x_column

        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        (max_row, max_col) = df.shape

        # col_end is None 일때
        if col_end is None:
            col_end = col_begin
        if style_no is None:
            style_no = self.style_no


        self.chart_position()
        pos_row = self.pos_row
        pos_col = self.pos_col


        columns = df.columns.tolist()
        col_xi = columns.index(col_x)
        col_1 = columns.index(col_begin)
        col_2 = columns.index(col_end)

        chart = workbook.add_chart({'type': 'line', 'subtype': subtype})


        for col_i in range(col_1,col_2+1):
            chart.add_series({
                'name':[sheet_name, 0, col_i],
                'categories': [sheet_name, 1, col_xi, max_row, col_xi],
                'values': [sheet_name, 1, col_i, max_row, col_i],
            })

        chart.set_style(style_no)
        chart.set_title({'name': title})
        chart.set_legend({'none': False, 'position': 'top'})

        if subtype == 'percent_stacked':
            chart.set_y_axis({
                'min': 0,
                'max': 1,  # 1 = 100%
                'num_format': '0%',  # 퍼센트 형식
                'major_gridlines': {'visible': True},  # 주요 그리드라인 표시
            })
        worksheet.insert_chart(pos_row, pos_col, chart)


    def chart_pie(self, col_name=None,
                  col_value=None,
                  title='',
                  pos_row=None, pos_col=None, style_no=None, precision=1):
        """Add a bar chart to the sheet"""
        #ic(">>> chart_pie()")
        self.graph_no += 1
        df = self.df
        sheet_name =  self.sheet_name

        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        (max_row, max_col) = df.shape

        # 놀기
        columns = df.columns.tolist()
        col_xi = columns.index(col_name)
        col_vi = columns.index(col_value)

        if style_no is None:
            style_no = self.style_no

        self.chart_position()
        pos_row = self.pos_row
        pos_col = self.pos_col

        chart = workbook.add_chart({'type': 'doughnut'})

        # 소수점 표현
        if precision == 1:
            number_foramt = '0.0%'
        elif precision == 2:
            number_foramt = '0.00%'
        else:
            number_foramt = '0%'
        # 차트에 데이터 추가
        chart.add_series({
            'name': [sheet_name, 0, col_xi],
            'categories': [sheet_name, 1, col_xi, max_row, col_xi],
            'values': [sheet_name, 1, col_vi, max_row, col_vi],
            'data_labels': {'percentage': True,
                            'number_format': number_foramt,
                            },  # 백분율 표시
        })

        chart.set_style(style_no)
        chart.set_title({'name': title})
        chart.set_legend({'none': False, 'position': 'top'})


        worksheet.insert_chart(pos_row, pos_col, chart)


    def set_x(self, name=None):
        if name is not None:
            self.x_column = name
            #print(f"* x column is set to be {name}")

    def set_settings(self, x_column, w=2, left_gap=8, style_no=11):
        self.set_x(name=x_column)
        self.w=w # 횡축 차트수
        self.pos_col_initial=left_gap
        self.style_no=style_no

    def chart_position(self):
        # pos_row 자동 할당하기
        self.pos_row = (self.graph_no-1) // self.w * self.pos_row_delta + self.pos_row_initial
        self.pos_col = (self.graph_no-1) % self.w  * self.pos_col_delta + self.pos_col_initial
        #print(self.pos_row, self.pos_col)




if __name__ == '__main__':
    ex = help(BKExcelWriter)
    print(ex)
    print(__name__)
