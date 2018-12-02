from dataHelper import *
from excelHelper import *

import xlwings
import shutil
import os
import datetime


# templates folders
# 数据文件夹和模板文件夹
template_folder = r'.\template'
data_folder = r'.\data'
cmm_data_folder = r'.\CMM'
program_data_folder = r'.\source_data'


# grr excel template files
p1_grr_template = 'p1 grr.xlsx'
# corr excel template files
p1_corr_template = 'p1 corr.xlsx'


# csv raw data heads
csv_data_heads = ['StartTime', 'FinishTime', 'Status', 'Error', 'CT', 'Barcode', 'ProductType', 'Description', 'SPC']

# module spec items
p1_specs = ['AB', 'J', 'M', 'O', 'N']
p1_spec_type = ['-', '=', '-', '+', '+']



# product dicts
grr_template_dict = {
            'p1': p1_grr_template,
        }
corr_data_dict = {
            'p1': 'CMM p1.xlsx',
        }
corr_template_dict = {
            'p1': p1_corr_template,
        }

product_spec_dict = {
            'p1': p1_specs,
        }
product_spec_type_dict = {
            'p1': p1_spec_type,
        }
product_csv_data_head_dict = {
            'p1': csv_data_heads + p1_specs,
        }


# grr excel formats
grr_cell_cols = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']
grr_cell_rows = [10, 11, 12, 15, 16, 17, 20, 21, 22]

grr_results = ['Repeatability', 'Reproducibility', 'Appraiser x Part', 'GRR']
grr_result_cells = ['N37', 'N38', 'N39', 'N40']

# grr_mode = 'normal' # 12345678910 12345678910
# grr_mode = 'repeat' # 111111111 222222222 ... 101010101010101010


# corr excel formats
corr_product_cols = ExcelHelper().get_cols('M', 30)
corr_type_col = 'E'
corr_spec_col = 'F'
corr_spec_start_row = 7

corr_results = ['slope', 'r2']
corr_result_cols = ['AW', 'AX']

# cmm raw excel format
cmm_product_cols = ExcelHelper().get_cols('C', 30)
cmm_spec_start_row = 2


if not os.path.exists('temp'):
    os.mkdir('temp')

if __name__ == '__main__':
    print('environment')
    pass
















