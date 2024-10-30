import math
import xlsxwriter
from basic_tables import Descriptive_Statistics as ds
from basic_tables import Basic_Frequency_Table as bft
from xlsx_utlilites import Xlsx_Utilities as xlsx_util

class Continuous_Variable:
  def __init__(self, wb, ws, props, starting_from):
    self.props = props
    self.data = props['data']
    self.ws = ws
    ##################################
    xl_util = xlsx_util(starting_from)
    self.cv_mv = xl_util.row_col_converter_moved
    self.col_cv = xl_util.col_converter_moved

    self.values = ds.create_table(wb, self.ws, self.data, props['title'], self.cv_mv)
    self.num_classes = round(math.sqrt(self.values['count']))

    self.class_limits()
    self.frequency_table()

    bft.create(wb, self.ws, self.data, self.num_classes, self.cv_mv, self.col_cv, self.props['end_row'])
    #self.ws.autofit()

  def class_limits(self):
    self.ws.merge_range(f'{self.cv_mv(0,5)}:{self.cv_mv(0,6)}', 'Tabla de Frecuencias')
    #Numero de clase
    self.ws.write(self.cv_mv(1,5), 'Numero de clase')
    self.ws.write(self.cv_mv(1,6), f'=SQRT({self.cv_mv(14,3)})')
    self.ws.write(self.cv_mv(1,7), f'=ROUND({self.cv_mv(1,6)}, 0)')
    #Tamaño de clase
    self.ws.write(self.cv_mv(2,5),'Tamaño de clase')
    self.ws.write(self.cv_mv(2,6), f'={self.cv_mv(10,3)}/{self.cv_mv(1,7)}')

  def frequency_table(self):
    self.ws.write_row(self.cv_mv(4,5), ['No. Clase','LI','LS','MC',])
    #No. Clase
    self.ws.write_column(self.cv_mv(5,5), self.num_classes_to_list())
    #LI
    self.ws.write(self.cv_mv(5,6),f'={self.cv_mv(11,3)}')
    self.ws.write_column(self.cv_mv(6,6),xlsx_util.loop_on_column(f'={self.col_cv(7)}#', 6, (self.num_classes-1)))#TODO: enviar argumento num_classes
    #LS
    self.ws.write_column(self.cv_mv(5,7), xlsx_util.loop_on_column(f'={self.col_cv(6)}#+{self.cv_mv(2,6, True,True)}', 6, self.num_classes))
    #MC
    self.ws.write_column(self.cv_mv(5,8),xlsx_util.loop_on_column(f'=AVERAGE({self.col_cv(6)}#:{self.col_cv(7)}#)', 6, self.num_classes))

      

  def num_classes_to_list(self):
    column = list(range(1,self.num_classes+1))
    return column