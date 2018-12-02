from environment import *


class GrrHelper(object):

    def create(self, product, station, grr_mode = 'normal'):
        self.grr_data_cells = []
        if grr_mode == 'normal':
            for r in grr_cell_rows:
                for c in grr_cell_cols:
                    self.grr_data_cells.append('%s%d' % (c, r))
        elif grr_mode == 'repeat':
            for c in grr_cell_cols:
                for r in grr_cell_rows:
                    self.grr_data_cells.append('%s%d' % (c, r))
        else:
            return

        self.product = product
        self.station = station

        # create grr dir
        self.grr_dir = "%s grr %s %s" %(datetime.datetime.now().strftime('%Y%m%d_%H%M%S'), self.station, self.product)
        os.mkdir(self.grr_dir)

        # copy grr excel from template
        self.grr_template_file = os.path.join(template_folder, grr_template_dict[self.product])
        self.grr_xlsx_file = datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + " grr " + self.station + ".xlsx"
        self.grr_xlsx_file = os.path.join(self.grr_dir, self.grr_xlsx_file)
        shutil.copy(self.grr_template_file, self.grr_xlsx_file)

        # init grr params
        self.grr_specs = product_spec_dict[self.product]
        self.grr_specs_type = product_spec_type_dict[self.product]
        self.grr_csv_data_heads = product_csv_data_head_dict[self.product]
        self.grr_xlsx = xlwings.Book(self.grr_xlsx_file)
        pass

    def run_with_data(self, datafile, count = 90):
        with open(datafile, 'r', encoding='UTF-8') as data:
            lines = data.readlines()
            raw_data_lines = lines[len(lines)-count:len(lines)]
            # fill grr excel
            for raw_data_index in range(0, len(raw_data_lines)):
                raw_data = raw_data_lines[raw_data_index].split(',')
                for spec in self.grr_specs:
                    spec_index = self.grr_specs.index(spec)
                    spec_type = self.grr_specs_type[spec_index]
                    if spec_type == '=':
                        # spc value
                        self.grr_xlsx.sheets[spec].range(self.grr_data_cells[raw_data_index]).value = raw_data[self.grr_csv_data_heads.index(spec)]
                    elif spec_type == '-':
                        # spc min
                        self.grr_xlsx.sheets[spec].range(self.grr_data_cells[raw_data_index]).value = raw_data[self.grr_csv_data_heads.index(spec) + 1 + len(self.grr_specs)]
                    elif spec_type == '+':
                        # spc max
                        self.grr_xlsx.sheets[spec].range(self.grr_data_cells[raw_data_index]).value = raw_data[self.grr_csv_data_heads.index(spec) + 1 + len(self.grr_specs) + 1 + len(self.grr_specs)]
                    else:
                        continue

            # backup grr raw data
            with open(os.path.join(self.grr_dir, "%s %s raw.csv" %(self.product, self.station)), 'w', encoding='UTF-8') as rf:
                    rf.writelines(raw_data_lines)
            pass

    def display(self):
        res_display = []
        for item in self.grr_specs:
            r1 = self.grr_xlsx.sheets[item].range(grr_result_cells[0]).value
            r2 = self.grr_xlsx.sheets[item].range(grr_result_cells[1]).value
            r3 = self.grr_xlsx.sheets[item].range(grr_result_cells[2]).value
            r4 = self.grr_xlsx.sheets[item].range(grr_result_cells[3]).value
            if r1 == '':
                r1 = '0'
            if r2 == '':
                r2 = '0'
            if r3 == '':
                r3 = '0'
            if r4 == '':
                r4 = '0'

            res = "Spec: {:>4} Repeatability: {:>6.2%} Reproducibility:{:>6.2%} Appraiser x Part: {:>6.2%} GRR: {:>6.2%}".format(
                item, float(r1), float(r2), float(r3), float(r4))
            print(res)
            res_display.append(res)

        with open(os.path.join(self.grr_dir, 'grr.dat'), 'w') as wf:
            for r in res_display:
                wf.write(r+'\n')
        return res_display
        pass


class GrrUtil(object):

    def __init__(self, product, station):
        self.product = product
        self.station = station

    def run_grr(self, raw_data_file, grr_mode = 'normal'):
        print("---------------------\nrun grr %s %s %s start\n" % (self.product, self.station, raw_data_file))
        grr = GrrHelper()
        grr.create(self.product, self.station, grr_mode)
        grr.run_with_data(raw_data_file)
        grr_display = grr.display()
        print("\nrun grr %s %s %s finish\n---------------------" % (self.product, self.station, raw_data_file))
        return grr_display
        pass


if __name__ == '__main__':

    pass