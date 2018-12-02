import shutil
import datetime


def get_today_data_file():
    return datetime.datetime.now().strftime('%Y%m%d')+'.csv'


class DataHelper(object):

    def set_backup_folder(self, raw_folder):
        self.raw_folder = raw_folder
        pass

    def backup_csv(self, obj_folder, suffix):
        cur_file = datetime.datetime.now().strftime('%Y%m%d')
        shutil.copy(self.raw_folder + cur_file+".csv", obj_folder + cur_file + " "+suffix+ ".csv")


if __name__ == '__main__':
    pass
