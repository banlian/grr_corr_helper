from environment import *

import time
import os


class CopyDataHelper(object):

    def __init__(self):
        self.src_data_folder = program_data_folder

    def run_copy(self):
        cur_file = get_today_data_file()
        data_file = os.path.join(self.src_data_folder, 'Data', cur_file)
       
        if not os.path.exists(self.src_data_folder):
            raise FileExistsError(self.src_data_folder)
        try:
            if os.path.exists(data_file):
                print('copy :' + data_file)
                shutil.copy(data_file, '.\\LeftData\\' + cur_file)
        except:
            print('copy data error:')
            raise FileExistsError('copy error')

        print("copy finish........")
        return 1


if __name__ == "__main__":
    today_datafile = get_today_data_file()
    print('start copy: ' + today_datafile)

    src_folder = r'data_foler'
    

    data_file = os.path.join(src_folder, 'Data', today_datafile)

    try:
        if os.path.exists(data_file):
            print('copy left:' + data_file)
            shutil.copy(data_file, '.\\Data\\' + today_datafile)
    except:
        print('copy left data error')



    print("copy finish........")

    time.sleep(0.8)
