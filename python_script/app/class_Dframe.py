from class_Data import Data


class Dframe(Data):
    def index_col(self):
        self.DataFrame_Sheet = self.DataFrame_Sheet.set_index('Наименование команды ТУ')
