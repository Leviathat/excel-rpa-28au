# coding=utf-8
import os
import xlsxwriter

cwd = os.getcwd()
files_list = [["Папка", "Имя Файла", "Расширение"]]


def get_files(folder):

    # Elements list in cwd folder
    directory = os.listdir(folder)

    # Segregation folders from files
    for elem in directory:
        if os.path.isdir(elem):
            os.chdir(elem)
            get_files(os.getcwd())
        else:
            filename, file_extension = os.path.splitext(str(elem))
            file = [os.getcwd(), filename, file_extension]
            files_list.append(file)


# Calling function
if __name__ == '__main__':
    get_files(cwd)

# Changing directory to create xlsx file in cwd dir
os.chdir(cwd)

workbook = xlsxwriter.Workbook('list.xlsx')

worksheet = workbook.add_worksheet()
worksheet.set_column(0, 3, 25)
worksheet.set_column(0, 0, 30)

for row_num, data in enumerate(files_list):
    worksheet.write_row(row_num, 0, data)

workbook.close()
