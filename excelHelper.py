

class ExcelHelper(object):

    def __init__(self):
        self.ROW_RANGE = [0 ,1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9]
        self.COL_RANGE = ['A' ,'B' ,'C' ,'D' ,'E' ,'F' ,'G' ,'H' ,'I' ,'J' ,'K' ,'L' ,
                          'M' ,'N' ,'O' ,'P' ,'Q' ,'R' ,'S' ,'T' ,'U' ,'V' ,'W' ,'X' ,'Y' ,'Z']

    def convert_str2int(self, col):
        start_index = 0
        for i in range(0, len(col)):
            start_index = start_index + (self.COL_RANGE.index(col[len(col) - 1 - i]) + 1) * (26 ** i)
        return start_index

    def convert_int2str(self, index):
        col = ''
        while index > 0:
            first = int((index-1)%26)
            if first >= 0:
                col = col + self.COL_RANGE[int((index-1) % 26)]
                index = int((index- ((index-1) % 26)) / 26)
        return col[::-1]

    def get_cols(self, start_col, col_count):
        cols = []
        # convert col str to int
        start_index = self.convert_str2int(start_col)
        # convert col int to str
        for i in range(start_index, start_index + col_count):
            cols.append(self.convert_int2str(i))

        return cols


if __name__ == '__main__':
    pass