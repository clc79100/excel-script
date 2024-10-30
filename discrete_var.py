import math
import xlsxwriter
from basic_tables import Descriptive_Statistics as ds
from basic_tables import Basic_Frequency_Table as bft
from xlsx_utlilites import Xlsx_Utilities as xlsx_util

class Discrete_varaible:
    def __init__(self, wb, ws, props, starting_from):
        self.wb = wb
        self.ws = ws
        self.props = props
        self.data = props['data']
        ###############################
        xl_util = xlsx_util(starting_from)
        self.cv_mv = xl_util.row_col_converter_moved
        self.col_cv = xl_util.col_converter_moved

        self.values = ds.create_table(self.wb, self.ws, self.data, self.props['title'], self.cv_mv)
        
        if('convenience_size_classes' in self.props):
            self.frequency_table_large_range()
        else:
            self.frequency_table_small_range()
    
    def frequency_table_small_range(self):
        percentage = self.wb.add_format({'num_format':'0.00%'})
        self.num_classes = round(self.values['max'])

        self.ws.write_row(self.cv_mv(4,5), ['No Clase','F','FA','FR','FRA',])
        #No. Clase
        self.ws.write_column(self.cv_mv(5,5), self.num_classes_to_list())
        #F
        self.ws.write_array_formula(f'{self.cv_mv(5,6)}:{self.cv_mv(self.num_classes+4,6)}', f'=FREQUENCY({self.cv_mv(1,0)}:{self.cv_mv((self.props['end_row']-1),0)}, {self.cv_mv(5,5)}:{self.cv_mv(self.num_classes+4,5)})')
        #FA
        self.ws.write(self.cv_mv(5,7), f'={self.cv_mv(5,6)}')
        self.ws.write_column(self.cv_mv(6,7), xlsx_util.loop_on_column(f'={self.col_cv(6)}#+{self.col_cv(7)}#',7,(self.num_classes-1), 6))
        #FR
        self.ws.write_column(self.cv_mv(5,8), xlsx_util.loop_on_column(f'={self.col_cv(6)}#/{self.cv_mv(14,3,True,True)}', 6, self.num_classes), percentage)
        #FRA
        self.ws.write(self.cv_mv(5,9),f'={self.cv_mv(5,8)}', percentage)
        self.ws.write_column(self.cv_mv(6,9), xlsx_util.loop_on_column(f'={self.col_cv(8)}#+{self.col_cv(9)}#', 7, (self.num_classes-1),6), percentage)


        #bft.create(self.wb, self.ws, self.data, self.num_classes, self.cv_mv, self.col_cv, self.props['end_row'],True)

    def frequency_table_large_range(self):
        self.num_classes = round(self.values['range']/self.props['convenience_size_classes'])
        self.ws.merge_range(f'{self.cv_mv(0,5)}:{self.cv_mv(0,6)}', 'Tabla de Frecuencias')
        
        self.ws.write(self.cv_mv(1,5), 'Numero de clase')
        self.ws.write(self.cv_mv(1,6), f'={self.cv_mv(10,3)}/{self.cv_mv(2,6)}')
        self.ws.write(self.cv_mv(1,7), f'=ROUND({self.cv_mv(1,6)}, 0)')

        self.ws.write(self.cv_mv(2,5), 'Tamaño de clase')
        self.ws.write(self.cv_mv(2,6), self.props['convenience_size_classes'])
        
        self.ws.write_row(self.cv_mv(4,5), ['No. Clase','LI','LS','MC',])
        #No. Clase
        self.ws.write_column(self.cv_mv(5,5), self.num_classes_to_list())
        #LI
        self.ws.write(self.cv_mv(5,6),f'={self.cv_mv(11,3)}')
        self.ws.write_column(self.cv_mv(6,6),xlsx_util.loop_on_column(f'={self.col_cv(7)}#', 6, (self.num_classes-1)))
        #LS
        self.ws.write_column(self.cv_mv(5,7), xlsx_util.loop_on_column(f'={self.col_cv(6)}#+{self.cv_mv(2,6, True,True)}',6,self.num_classes))
        #MC
        self.ws.write_column(self.cv_mv(5,8),xlsx_util.loop_on_column(f'=AVERAGE({self.col_cv(6)}#:{self.col_cv(7)}#)', 6, self.num_classes))

        bft.create(self.wb, self.ws, self.data, self.num_classes, self.cv_mv, self.col_cv, self.props['end_row'])

    def num_classes_to_list(self):
        column = list(range(1,self.num_classes+1))
        return column