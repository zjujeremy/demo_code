# #!/usr/bin/python
import xlrd
import xlsxwriter as xw
import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import os
import cv2
import argparse

dirname, filename = os.path.split(os.path.abspath( __file__))
IMG_COL = 'C'

image_path_dir = os.path.join(dirname, "images")
if not os.path.exists(image_path_dir):
    os.mkdir(image_path_dir)


def read_excel(args):
    xlsx = xlrd.open_workbook(args.input)
    sheet1 = xlsx.sheets()[0] 
    sheet1_nrows = sheet1.nrows
    excel_data = []
    for i in range(sheet1_nrows):
        excel_data.append(sheet1.row_values(i))
    return excel_data
 
def write_excel(excel_data, args):
    fileName = args.output
    workbook = xw.Workbook(fileName)
    worksheet1 = workbook.add_worksheet(args.sheet)
    worksheet1.activate()

    for j in range(len(excel_data)):
        insertData = excel_data[j]
        image_url = excel_data[j][ord(args.url) - ord('A')]
        image_path = save_img(image_url)
        row = 'A' + str(j + 1)
        worksheet1.write_row(row, insertData)
        worksheet1.set_row(int(j), int(args.size))
        img = cv2.imread(image_path)
        x = float(int(args.size) / img.shape[0])
        worksheet1.insert_image(args.image + str(j + 1), image_path, {'x_scale': x, 'y_scale': x})
    workbook.close()

def save_img(url):
    res = requests.get(url)
    file_name = os.path.join(image_path_dir, url.split('/')[-1])
    with open(file_name, 'wb') as f:
        for data in res.iter_content(128):
            f.write(data)
    return file_name

if __name__ == "__main__":
    print("*****************  welcome! Zhuzhengzheng! 请我们吃饭！  *****************")
    print('usage: ')
    print('"-i", "--input", help="please input src excel path"')
    print('"-o", "--output", help="please input result excel path"')
    print('"-sheet", "--output", help="please input sheet name", default="sheet1"')
    print('"-url", help="please input url column", default="B"')
    print('"-image", help="please input image column", default="C"')
    print('"-size", "--size", help="please input image size", default="140"')
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input", help="please input src excel path")
    parser.add_argument("-o", "--output", help="please input result excel path", default="dome.xlsx")
    parser.add_argument("-size", "--size", help="please input image size", default="140")
    parser.add_argument("-sheet", "--sheet", help="please input sheet name", default="sheet1")
    parser.add_argument("-url", help="please input url column", default="B")
    parser.add_argument("-image", help="please input image column", default="E")
    args = parser.parse_args()
    excel_data = read_excel(args)
    write_excel(excel_data, args)


