

import shutil
import xlwings


if __name__ == "__main__":

    shutil.copy(r'.\template\grr.xlsx', r'.\temp\test.xlsx')

    xlsx = xlwings.Book(r'.\temp\test.xlsx')
    xlsx.app.display_alerts = False

    template_sheet_name = 'SPC1'

    sht = xlsx.sheets[template_sheet_name]
    # sht.api.Copy(Before=sht.api)

    # template variables initialize
    fais = ['f1', 'f2', 'f3', 'f4']
    tols = [0.2,0.3,0.4,0.5]
    tol_cell = 'D7'
    sigma = 5.15
    sigma_cell = 'N35'

    # template generator
    newBook = xlwings.Book()
    for i in range(len(fais)):
        sht.api.Copy(Before = newBook.sheets[i].api)
        newBook.sheets[template_sheet_name].range(tol_cell).value = tols[i]
        newBook.sheets[template_sheet_name].range(sigma_cell).value = sigma
        newBook.sheets[template_sheet_name].name = fais[i]
    newBook.sheets['Sheet1'].delete()

    newBook.save(r'.\temp\testtemplate.xlsx')


