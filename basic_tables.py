import pandas as pd
import xlsxwriter
from xlsx_utlilites import Xlsx_Utilities as xlsx_util

class Descriptive_Statistics:
    @staticmethod
    def create_table(wb, ws, data, title, cv_mv):

        title_style = wb.add_format({"italic" : True, "bottom" : True, 'align':'center_across'})
        bottom_style = wb.add_format({"bottom": True})

        pd_serie = pd.Series(data)

        #Columna de datos
        ws.write(cv_mv(0, 0), title)
        ws.write_column(cv_mv(1, 0), data)

        ws.write(cv_mv(0, 2), title, title_style)
        ws.write(cv_mv(0, 3), '', title_style)

        min = pd_serie.min()
        max = pd_serie.max()
        range = max-min

        table_values = {
            'mean': pd_serie.mean(),
            'sem': pd_serie.sem(),
            'median': pd_serie.median(),
            'mode': pd_serie.mode().iloc[0],
            'std': pd_serie.std(),
            'var': pd_serie.var(),
            'kurt': pd_serie.kurt(),
            'skew': pd_serie.skew(),
            'range': range,
            'min': min,
            'max': max,
            'sum': pd_serie.sum(),
            'count': pd_serie.count(),
        }

        ws.write(cv_mv(2, 2), 'Media')
        ws.write(cv_mv(2, 3), table_values['mean'])

        ws.write(cv_mv(3, 2), 'Error típico')
        ws.write(cv_mv(3, 3), table_values['sem'])

        ws.write(cv_mv(4, 2), 'Mediana')
        ws.write(cv_mv(4, 3), table_values['median'])

        ws.write(cv_mv(5, 2), 'Moda')
        ws.write(cv_mv(5, 3), table_values['mode'])

        ws.write(cv_mv(6, 2), 'Desviación estándar')
        ws.write(cv_mv(6, 3), table_values['std'])

        ws.write(cv_mv(7, 2), 'Varianza de la muestra')
        ws.write(cv_mv(7, 3), table_values['var'])

        ws.write(cv_mv(8, 2), 'Curtosis')
        ws.write(cv_mv(8, 3), table_values['kurt'])

        ws.write(cv_mv(9, 2), 'Coeficiente de asimetría')
        ws.write(cv_mv(9, 3), table_values['skew'])

        ws.write(cv_mv(10, 2), 'Rango')
        ws.write(cv_mv(10, 3), table_values['range'])

        ws.write(cv_mv(11, 2), 'Mínimo')
        ws.write(cv_mv(11, 3), table_values['min'])

        ws.write(cv_mv(12, 2), 'Máximo')
        ws.write(cv_mv(12, 3), table_values['max'])

        ws.write(cv_mv(13, 2), 'Suma')
        ws.write(cv_mv(13, 3), table_values['sum'])

        ws.write(cv_mv(14, 2), 'Cuenta', bottom_style)
        ws.write(cv_mv(14, 3), table_values['count'], bottom_style)
        return table_values

class Basic_Frequency_Table:
    @staticmethod
    def create(wb, ws, data, num_classes, cv_mv, col_cv, end_row, discrete_var_small_range=False):
        percentage = wb.add_format({'num_format':'0.00%'})
        ws.write_row(cv_mv(4,9), ['F','FA','FR','FRA',])

        if(discrete_var_small_range):
            ws.write_array_formula(f'{cv_mv(5,9)}:{cv_mv(num_classes+4,9)}', f'=FREQUENCY({cv_mv(1,0)}:{cv_mv((end_row-1),0)}, {cv_mv(5,8)}:{cv_mv(num_classes+4,8)})')
        else:#TODO: falta agregar condicional para diferentes Frecunecias de discrete_var
            #F
            ws.write_array_formula(f'{cv_mv(5,9)}:{cv_mv(num_classes+4,9)}', f'=FREQUENCY({cv_mv(1,0)}:{cv_mv((end_row-1),0)}, {cv_mv(5,7)}:{cv_mv(num_classes+4,7)})')
        
        #FA
        ws.write(cv_mv(5,10), f'={cv_mv(5,9)}')
        ws.write_column(cv_mv(6,10), xlsx_util.loop_on_column(f'={col_cv(9)}#+{col_cv(10)}#',7,(num_classes-1), 6))
        #FR
        ws.write_column(cv_mv(5,11), xlsx_util.loop_on_column(f'={col_cv(9)}#/{cv_mv(14,3,True,True)}', 6, num_classes), percentage)
        #FRA
        ws.write(cv_mv(5,12),f'={cv_mv(5,11)}', percentage)
        ws.write_column(cv_mv(6,12), xlsx_util.loop_on_column(f'={col_cv(11)}#+{col_cv(12)}#', 7, (num_classes-1),6), percentage)
        
