from environment import *


class CorrHelper(object):

    def __init__(self):
        pass

    def create(self, product, corr_type):
        self.product = product
        self.corr_type = corr_type

        # create corr dir
        self.corr_dir = "%s corr %s %s" % (datetime.datetime.now().strftime('%Y%m%d_%H%M%S'), self.corr_type, self.product)
        os.mkdir(self.corr_dir)

        # copy corr excel from template
        self.corr_template_file = os.path.join(template_folder, corr_template_dict[product])
        self.corr_xlsx_file = datetime.datetime.now().strftime('%Y_%m_%d_%H_%M_%S') + " " + self.product + " corr " + self.corr_type + ".xlsx"
        self.corr_xlsx_file = os.path.join(self.corr_dir, self.corr_xlsx_file)
        shutil.copy(self.corr_template_file, self.corr_xlsx_file)

        # init corr params
        self.corr_specs = product_spec_dict[product]
        self.corr_specs_type = product_spec_type_dict[product]
        self.corr_csv_data_heads = product_csv_data_head_dict[product]
        self.corr_xlsx = xlwings.Book(self.corr_xlsx_file)
        pass

    def run_corr_with_cmm(self, cmm_xlsx_file, corr_count=30):
        """fill excel with cmm data
        :param cmm_xlsx_file:
        :param corr_count:
        """
        if self.corr_type == 'lr':
            return

        cmm_xlsx = xlwings.Book(cmm_xlsx_file)
        type = "CMM"

        # fill corr excel with cmm raw data
        # spec row
        for r in range(0, len(self.corr_specs)):
            # product col
            for c in range(0, corr_count):
                # min value
                cur_cmm_value_min = cmm_xlsx.sheets['cmm'].range(cmm_product_cols[c] + str(cmm_spec_start_row + r * 2 + 1)).value
                # max value
                cur_cmm_value_max = cmm_xlsx.sheets['cmm'].range(cmm_product_cols[c] + str(cmm_spec_start_row + r * 2)).value

                # select cmm raw value to corr based on spec type
                if self.corr_specs_type[r] == '-':
                    cur_cmm_value = cur_cmm_value_min
                elif self.corr_specs_type[r] == '+' or self.corr_specs_type[r] == '=':
                    cur_cmm_value = cur_cmm_value_max
                else:
                    continue

                # fill cmm raw data to corr excel
                cur_corr_row = corr_spec_start_row + r * 2
                cur_corr_col = corr_product_cols[c]
                self.corr_xlsx.sheets['corr'].range(cur_corr_col + str(cur_corr_row)).value = cur_cmm_value
                self.corr_xlsx.sheets['corr'].range(corr_type_col + str(cur_corr_row)).value = type
                self.corr_xlsx.sheets['corr'].range(corr_spec_col + str(cur_corr_row + r * 2)).value = self.corr_specs[r]

        cmm_xlsx.close()
        pass

    def run_corr_with_data(self, datafile, corr='lc', corr_count=30):
        """fill excel with station data
        :param corr:
        :param datafile: csv raw data file
        :param corr_count:
        """
        # lr corr
        if corr == 'l' or corr == 'L':
            cur_corr_spec_start_row = corr_spec_start_row + 0
            type = 'L'
        # lr corr
        elif corr == 'r' or corr == 'R':
            cur_corr_spec_start_row = corr_spec_start_row + 1
            type = 'R'
        # lc corr
        elif corr == 'lc' or corr == 'LC':
            cur_corr_spec_start_row = corr_spec_start_row + 1
            type = 'L'
        # rc corr
        elif corr == 'rc' or corr == 'RC':
            cur_corr_spec_start_row = corr_spec_start_row + 1
            type = 'R'
        else:
            return

        if not os.path.exists(datafile):
            print('run_corr_with_data:', datafile, ' not exists')
            return

        with open(datafile, 'r', encoding='UTF-8') as df:
            lines = df.readlines()
            csv_corr_lines = lines[len(lines) - corr_count:len(lines)]

            # fill corr excel with raw csv data
            spec_update = [0 for r in self.corr_specs]

            # product col
            for r in range(0, corr_count):
                data = csv_corr_lines[r].split(',')
                # spec row
                for spec_index in range(0, len(self.corr_specs)):
                    csv_spec_index = self.corr_csv_data_heads.index(self.corr_specs[spec_index])
                    # fill cmm raw data to corr excel
                    cur_corr_row = str(cur_corr_spec_start_row + spec_index * 2)

                    if self.corr_specs_type[spec_index] == '=':
                        self.corr_xlsx.sheets['corr'].range(corr_product_cols[r] + cur_corr_row).value = data[csv_spec_index]
                    elif self.corr_specs_type[spec_index] == '-':
                        self.corr_xlsx.sheets['corr'].range(corr_product_cols[r] + cur_corr_row).value = data[csv_spec_index+1+len(self.corr_specs)]
                    elif self.corr_specs_type[spec_index] == '+':
                        self.corr_xlsx.sheets['corr'].range(corr_product_cols[r] + cur_corr_row).value = data[csv_spec_index+1+len(self.corr_specs)+1+len(self.corr_specs)]
                    else:
                        continue

                    if spec_update[spec_index] == 0:
                        self.corr_xlsx.sheets['corr'].range(corr_type_col + cur_corr_row).value = type
                        self.corr_xlsx.sheets['corr'].range(corr_spec_col + cur_corr_row).value = self.corr_specs[spec_index]
                        spec_update[spec_index] = 1

            # back up cmm raw data
            cmm_raw_data_file = os.path.join(self.corr_dir, "raw corr %s %s.csv" % (self.product, corr))
            with open(cmm_raw_data_file, 'w', encoding='UTF-8') as rf:
                rf.writelines(csv_corr_lines)
        pass

    def run_corr_with_lr(self, left_file, right_file, count=30):
        self.run_corr_with_data(left_file, 'l', count)
        self.run_corr_with_data(right_file, 'r', count)
        pass

    def display(self):
        res_row = corr_spec_start_row
        res_cols = corr_result_cols

        res_display = []
        for i in range(0, len(self.corr_specs)):
            res_row = corr_spec_start_row + i * 2
            corr_res = []
            for c in res_cols:
                corr_res.append(self.corr_xlsx.sheets['corr'].range(c + str(res_row)).value)
            res = "Spec: {}  slope: {:>4.2f}  r2: {:>4.2f}".format(self.corr_specs[i], float(corr_res[0]), float(corr_res[1]))
            print(res)
            res_display.append(res)

        with open(os.path.join(self.corr_dir, 'corr.dat'), 'w') as wf:
            for r in res_display:
                wf.write(r + '\n')
        return res_display
        pass


class CorrUtil(object):

    def __init__(self, product):
        self.product = product
        pass

    def run_corr_cmm2station(self, corr_type, cmm_file, cmm_count, raw_data_file, raw_count):
        if corr_type == 'lc':
            cmm_corr_type = 'lc'
            station = 'left'
        elif corr_type == 'rc':
            cmm_corr_type = 'rc'
            station = 'right'
        else:
            return

        print("-------------------\nrun %s cmm corr %s %s %s %s start\n" % (station, self.product, cmm_corr_type, cmm_file, raw_data_file))
        c = CorrHelper()
        c.create(self.product, cmm_corr_type)
        c.run_corr_with_cmm(cmm_file, cmm_count)
        c.run_corr_with_data(raw_data_file, cmm_corr_type, raw_count)
        corr_display = c.display()
        print("\nrun %s cmm corr %s %s %s %s finish\n-------------------------" % (station, self.product, cmm_corr_type, cmm_file, raw_data_file))
        return corr_display
        pass

    def run_corr_station2station(self, left_file, raw_count_left, right_file, raw_count_right):
        cmm_corr_type = 'lr'
        corr_type_L = 'l'
        corr_type_R = 'r'
        # left right corr
        print("-------------------\nrun left right corr %s %s %s %s start\n" % (self.product, cmm_corr_type, left_file, right_file))
        c = CorrHelper()
        c.create(self.product, cmm_corr_type)
        c.run_corr_with_data(left_file, corr_type_L, raw_count_left)
        c.run_corr_with_data(right_file, corr_type_R, raw_count_right)
        corr_display = c.display()
        print("\nrun left right corr %s %s %s %s finish\n-------------------------" % (self.product, cmm_corr_type, left_file, right_file))
        return corr_display


if __name__ == '__main__':
    pass