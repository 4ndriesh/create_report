# -*- coding: utf-8 -*-
__author__ = 'BiziurAA'
from class_Sheets import Sheets
import pandas as pd
import shutil
import sys
from class_Data import Data
from logging_err import *
from ini_file import open_ini_file
import os




def total_main():
    ini = open_ini_file.inst()
    dt = Data()
    # path = os.getcwd()
    path_report=os.path.join(ini.BASE_DIR,'report')
    path_temp = os.path.join(ini.BASE_DIR, 'temp')
    for conv_file in dt.list_station:
        name_report = conv_file
        cr_report_TU = Sheets()
        cr_report_TU.workSheets_TC_TU = '{0}{1}'.format('К_', name_report)
        cr_report_TU.work_cols = 'A:C,J'
        try:
            file_name = '{0}{1}{2}'.format('Ведомость ', name_report, '.xlsx')
            shutil.copy(os.path.join(path_temp, 'Tamplate.xlsx'), os.path.join(path_report, file_name))
            file_name_maket = '{0}{1}{2}{3}'.format('Ведомость ', name_report, '  макет', '.xlsx')
            shutil.copy(os.path.join(path_temp, 'Tamplate_maket.xlsx'), os.path.join(path_report, file_name_maket))
        except Exception as e:
            exeption_print(e)
            sys.exit()

        if cr_report_TU.read_sheet_by_excel(True) is 0: break
        cr_report_TU.choose_unic_part()
        cr_report_TU.index_col()

        start_key = ''
        for end_key in cr_report_TU.Dict_in_param.keys():
            if start_key is not '':
                write_to_excel = cr_report_TU.DataFrame_Sheet.loc[start_key: end_key, ['Unnamed: 1', '', 'Код ПУ']]
                cr_report_TU.Dict_in_param[start_key] = write_to_excel[0:].reset_index()[
                    cr_report_TU.Format_out[start_key]]
            start_key = end_key

        cr_report_TU.format_sheet()

        DataFrame_dict_list = list()
        local_DataFrame_dict_list = list()
        for wp in cr_report_TU.choose_param:
            for ppp in wp:
                # if ppp == 'Ответственные команды' :
                # # if cr_report_TU.choose_param[-1] == wp:
                #     local_DataFrame_dict_list.append(cr_report_TU.Dict_in_param.get(ppp)[0:].fillna(''))
                # else:
                local_DataFrame_dict_list.append(cr_report_TU.Dict_in_param.get(ppp)[0:-1].fillna(''))
            if len(wp) == 0:
                DataFrame_dict_list.append(pd.DataFrame())
            else:
                DataFrame_dict_list.append(pd.concat(local_DataFrame_dict_list))
            local_DataFrame_dict_list.clear()

        cr_report_TU.Data_for_record = dict(zip(cr_report_TU.Sheet_dict, DataFrame_dict_list))
        # cr_report_TU.write_to_excel_win32com_pd(8, 1, path, file_name, conv_file)
        cr_report_TU.write_to_excel(7, 0, path_report, file_name, conv_file)

        # --------------------------------------------------------------------------------------------------------------
        cr_report_TC = Sheets()
        cr_report_TC.workSheets_TC_TU = name_report
        cr_report_TC.work_cols = 'C:D,F'

        cr_report_TC.read_sheet_by_excel(False)
        DataFrame_dict_list = list()
        DataFrame_dict_list = cr_report_TC.DataFrame_Sheet.dropna(axis=0, how='any').loc[:,
                              cr_report_TC.Format_out['Наименование']].fillna('')

        cr_report_TC.Data_for_record = {'№5': DataFrame_dict_list}
        # cr_report_TC.write_to_excel_win32com_pd(7, 2, path, file_name, conv_file)
        cr_report_TC.write_to_excel(6, 1, path_report, file_name, conv_file)

        # --------------------------------------------------------------------------------------------------------------
        cr_report_TU.workSheets_TC_TU = 'ТУ'
        cr_report_TU.work_cols = 'C:D,F'

        write_to_excel = cr_report_TU.DataFrame_Sheet.loc[:, ['Unnamed: 1', '', 'Код ПУ']].fillna('')
        a = write_to_excel[1:].reset_index()[cr_report_TU.Format_out['Управляющий приказ']]

        cr_report_TU.Data_for_record = {'ТУ': a}
        # cr_report_TU.write_to_excel_win32com_pd(6, 1, path, file_name_maket, conv_file)
        cr_report_TU.write_to_excel(5, 0, path_report, file_name_maket, conv_file)
        # --------------------------------------------------------------------------------------------------------------
        cr_report_TC.workSheets_TC_TU = 'ТC'
        cr_report_TC.work_cols = 'C:D,F'
        cr_report_TC.Data_for_record = {
            'TC': cr_report_TC.DataFrame_Sheet.loc[:, cr_report_TC.Format_out['Наименование1']].fillna('')}
        cr_report_TC.write_to_excel(6, 1, path_report, file_name_maket, conv_file)
        # cr_report_TC.write_to_excel_win32com_pd(7, 2, path, file_name_maket, conv_file)
        del cr_report_TU
        del cr_report_TC
