from docxtpl import DocxTemplate


def test():
    data_dic = {
        't1': '燕子',
        't2': '杨柳',
        't3': '桃花',
        't4': '针尖',
    }
    doc = DocxTemplate('temp/temp.docx')  # 加载模板文件
    doc.render(data_dic)  # 填充数据
    doc.save('docx/temp.docx')  # 保存目标文件


if __name__ == '__main__':
    test()
