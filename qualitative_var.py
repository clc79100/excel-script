import pandas as pd
import xlsxwriter
from basic_tables import Descriptive_Statistics as ds
from basic_tables import Basic_Frequency_Table as bft
from xlsx_utlilites import Xlsx_Utilities as xlsx_util

class Qualitative_Variable:
    def __init__(self, wb, ws, props, starting_from):
        self.props = props
        self.data = props['data']
        self.wb = wb
        self.ws = ws
        #######################################
        #print(f'{self.props['title']} qualitative')
        xl_util = xlsx_util(starting_from)
        self.cv_mv = xl_util.row_col_converter_moved
        self.col_cv = xl_util.col_converter_moved

        #Columna de datos
        self.ws.write(self.cv_mv(0, 0), props['title'])
        self.ws.write_column(self.cv_mv(1, 0), self.data)

        self.frequency_table()

    def frequency_table(self):

        table_values = {
            'category':[],
            'value':[],
        }

        pd_serie = pd.Series(self.data)
        frequency = pd_serie.value_counts()
        data_len = len(frequency)
        percentage = self.wb.add_format({'num_format':'0.00%'})

        self.ws.write_row(self.cv_mv(0,2), ['titulo','F','FR'])
        #Categoria
        self.ws.write_column(self.cv_mv(1,2), frequency.index.to_list())
        #F
        self.ws.write_column(self.cv_mv(1,3), frequency.values.tolist())
        #total de F
        self.ws.write(self.cv_mv((data_len+1),3), frequency.sum())
        #FR
        self.ws.write_column(self.cv_mv(1,4), xlsx_util.loop_on_column(f'= {self.col_cv(3)}#/{self.cv_mv((data_len+1),3,True,True)}',2,data_len), percentage)

        return table_values

