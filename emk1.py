import DocUtils
import os
from ExcelUtils import Excel
import shutil
import re
import time

tdr_data_list = []


def test():
    add_time = 10
    print("开始处理！")
    if not os.path.exists("excel"):
        os.mkdir("excel")
    if not os.path.exists("docx"):
        os.mkdir("docx")
    clear_dir("excel")
    clear_dir("docx")
    DocUtils.doc_2_docx(os.getcwd() + "/doc", os.getcwd() + "/docx")
    for parent, directory, files in os.walk(os.getcwd() + "/docx"):
        for f in files:
            if not re.match(".*格式件.*", f):
                continue
            try:
                deal_data_2_excel(parent + "/" + f)
            except Exception:
                print("文件：" + parent + "/" + f + "格式有误，读取失败，请人工读取！")
                add_time += 500
    excel = Excel()
    shutil.copyfile("temp/tdr_temp.xls", "excel/TDR.xls")
    excel.modify_excel("excel/TDR.xls", [(2, 1, tdr_data_list)], sheet_index=1)
    print("处理完毕！")
    time.sleep(add_time)


def deal_data_2_excel(file_path):
    global tdr_data_list
    table_list = DocUtils.get_table_text_list(file_path)
    textbox_list = DocUtils.get_textbox_text_list(file_path)
    a_6 = str(table_list[6])["Shipper's Name and Address".__len__():].strip()
    a_9 = str(table_list[9])["Reference #".__len__():].strip()  # BOOKING NUMBER
    a_11 = str(table_list[11])["Bill of Lading :".__len__():].strip()
    a_18 = str(table_list[18])["Consignee (Negotiable only if consigned 'To Order1, 'To Order of', To Order of Bearer*)".__len__():].strip()
    a_24 = str(table_list[24])["Notify Party (see clause 2)".__len__():].strip()
    a_48 = str(table_list[48])["Port of Loading".__len__():].strip()
    a_49 = str(table_list[49])["Port of Discharge".__len__():].strip()
    tdr_free_day = ""
    b_type_arr = []
    tdr_type_arr = []
    b_container = []
    start_index = 0
    end_index = 0
    hs_code = []
    b_16_other_1 = ""
    b_16_other_2 = ""
    for index, b_c in enumerate(textbox_list):
        if re.match(r"([A-Z]*[0-9]+/[A-Z]*[0-9]+)|(([A-Z]*[0-9]+/[A-Z]*[0-9]+)(/[A-Z]*[0-9]+){3})", b_c) is not None:
            b_container.append(str(b_c).strip())
        if re.match(r"\d+X\d+[HC|GP]", b_c) is not None:
            b_type_a = str(b_c).strip().split("+")
            for b_type in b_type_a:
                count = int(b_type[0: b_type.index("X")])
                for i in range(count):
                    b_type_arr.append("1" + b_type[b_type.index("X"):])
                    tdr_type_arr.append(re.findall("X(?P<tag_name>\d+)(HC|GP)", b_type)[0])
        if re.match(r".*Packages/General Description of Goods", b_c) is not None:
            start_index = index + 1
        if re.match(r".*DAYS FREE DETENTION AT DESTINATION", b_c) is not None:
            tdr_free_day = re.findall("(?P<tag_name>\d+).*DAYS FREE DETENTION AT DESTINATION", b_c)[0]
            end_index = index
        if re.match(r"H\D{0,2}S\D{0,2}CODE", b_c) is not None:
            hs_code.append(b_c)
        if re.match(r"Container No./Seal No./Marks & Numbers", b_c) is not None:
            b_16_other_1 = textbox_list[index + 1]
        if re.match(r"Gross Weight \(Kgs\) Measurement \(CBM\)", b_c) is not None:
            b_16_other_2 = textbox_list[index + 1]
            if re.match(r".*CBM.*", b_16_other_2) is None:
                b_16_other_2 = b_16_other_2 + textbox_list[index + 2]
    commodity = "\n".join(textbox_list[start_index: end_index])
    if b_type_arr.__len__() != b_container.__len__():
        raise Exception()
    data_list = []
    for i in range(b_container.__len__()):
        b_16_arr = str(b_container[i]).split("/")
        if b_16_arr.__len__() == 2:
            b_16_arr.append(b_16_other_1)
            b_16_arr.append(b_16_other_2[:b_16_other_2.index("KGS") + 3].strip())
            b_16_arr.append(b_16_other_2[b_16_other_2.index("KGS") + 3:].strip())
        data_list.append([a_11, b_16_arr[0], b_16_arr[1],  b_type_arr[i], a_6, a_18, a_24, a_48, a_48, a_49, a_49, commodity, "\n".join(hs_code), b_16_arr[2], b_16_arr[3], b_16_arr[4]])
        tdr_data_list.append([b_16_arr[0], b_16_arr[1], tdr_type_arr[i][0], tdr_type_arr[i][1], "F", a_11, a_9, b_16_arr[2], commodity, "", "", b_16_arr[3], a_49, tdr_free_day, "", "", "CY-CY", ""])
    modify_data_list = [(4, 1, data_list)]
    excel = Excel()
    excel_file_name = file_path[file_path.rindex("/") + 1: file_path.rindex("格式件")] + "-CARGO MANIFEST"
    shutil.copyfile("temp/temp.xls", "excel/{}.xls".format(excel_file_name))
    excel.modify_excel("excel/{}.xls".format(excel_file_name), modify_data_list, sheet_index=1)


def clear_dir(dir_path):
    for root, dirs, files in os.walk(dir_path, topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))


if __name__ == '__main__':
    test()
