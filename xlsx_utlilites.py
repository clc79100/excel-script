from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol

class Xlsx_Utilities:

    def __init__(self, starting_from):
      self.starting_from = starting_from

    @staticmethod
    def only_column_cell(cell):
      row_col = xl_cell_to_rowcol(cell)
      cell = xl_rowcol_to_cell(row=0, col=row_col[1])
      return cell.replace('1','')

    @staticmethod
    def only_row_cell(cell):
      row_col = xl_cell_to_rowcol(cell)
      cell = xl_rowcol_to_cell(row=row_col[0], col=0)
      return cell.replace('A','')

    def column_to_only_column_cell(col):
      cell = xl_rowcol_to_cell(row=0,col=col)
      return cell.replace('1','')


    @staticmethod
    def loop_on_column(formula, start, limit, start2=None):
      list = []
      row, row2 = 0, 0

      if(start2):
        formula_formatted = lambda row, row2: formula.replace('#', str(row),1).replace('#', str(row2),1)#TODO: falta probar caso de 2 valores
      else:
        formula_formatted = lambda row, row2: formula.replace('#', str(row))
        start2 = 0

      for i in range(limit):
        row = i+start
        row2 = i+start2
        list.append(formula_formatted(row, row2))
      return list

    def move_cell(self, col_value):
      return col_value + self.starting_from

    def row_col_converter_moved(self, row, col, row_abs=False, col_abs=False):
      return xl_rowcol_to_cell(row, self.move_cell(col), row_abs, col_abs)

    #solo soporta de A a ZZ columnas (676 celdas)
    def col_converter_moved(self, col):
      cell = xl_rowcol_to_cell(0,self.move_cell(col))
      if(len(cell) > 2):
        return cell[:2]
      
      return cell[0]