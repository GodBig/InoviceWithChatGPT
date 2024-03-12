import pprint
import time
import shutil
import xlwings
import pandas
from datetime import datetime
import os
from PIL import Image
from PIL import ImageGrab
import pytesseract
import fitz
import openai
import json
import webbrowser


def chatgpt_check(res_ocr):
    key = "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"
    # Load your API key from an environment variable or secret management service
    openai.api_key = key
    message_dict = {"role": "user",
                    "content": "Firstly, I'm from China State Construction Middle East Company. I've received many invoices and work completion confirmation forms. I scanned them all together into images and used Tesseract OCR to recognize the content in the images. However, the image content is somewhat chaotic.Please assist me in identifying, correcting, and even filling in crucial information." +
                               "Secondly, document type, purchasing department or project, invoice date, invoice number, LPO code, TRN code, and supplier name from the recognized content, and enter the results into the JSON format below." +
                               str({
                                   "type": "",
                                   "project": "",
                                   "date": "",
                                   "invoice": "",
                                   "lpo": "",
                                   "trn": "",
                                   "supplier": "",
                               }) +
                               "If the provided content does not contain the required information to fill in the JSON format.Just give me the raw Json without any update."
                               "This is the content from Tesseract OCR:" + res_ocr}
    chat_completion = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[message_dict]
    )
    gpt_answer = chat_completion["choices"][0]["message"]["content"]
    with open('gpt_train.txt', 'a') as file:
        lines = str(message_dict) + "\n" + gpt_answer
        file.writelines(lines)
    return gpt_answer


def pdf2png(input_pdf, file_num):
    output_path = ".\\pdf2img\\"
    compression = 'zip'  # "zip", "lzw", "group4" - need binarized image...

    zoom = 3  # to increase the resolution
    mat = fitz.Matrix(zoom, zoom)

    file_path = ".\\raw_pdf\\"
    doc = fitz.open(file_path + input_pdf)
    image_list = []
    for page in doc:
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        image_list.append(img)

    page_num = 1
    for final_image in image_list:
        blacked_image = final_image.convert('L')  # 灰度化
        # 自定义灰度界限，这里可以大于这个值为黑色，小于这个值为白色。threshold可根据实际情况进行调整(最大可为255)。
        threshold = 100
        table = []
        for i in range(256):
            if i < threshold:
                table.append(0)
            else:
                table.append(1)
        save_image = blacked_image.point(table, '1')
        # 根据部门编号码
        # 黑白化了，保存处理好的图片
        output_name = file_num + str(page_num) + ".png"
        page_num += 1
        save_image.save(
            output_path + output_name,
            save_all=True,
            append_images=image_list[1:],
            compression=compression,
            dpi=(300, 300),
        )
    return page_num


def readinvoice():
    file_path = "D:\ProjectOfFinance\AutoPaymentCertificate\\raw_pdf\\"
    file_list = os.listdir(file_path)
    img_name = "1"
    for lpo_file_name in file_list:
        page_num = pdf2png(lpo_file_name, img_name)
        img_name = str(int(img_name) + 1)
    img_path = "D:\ProjectOfFinance\AutoPaymentCertificate\\pdf2img\\"
    file_list = os.listdir(img_path)
    for lpo_file_name in file_list:
        # 正式识别内容
        image = Image.open(img_path + lpo_file_name)
        # cropped = image.crop((0, 0, 1785, 1780))
        # cropped.show()
        # 解析图片，lang='chi_sim'表示识别简体中文，默认为English
        # 如果是只识别数字，可再加上参数config='--psm 6 --oem 3 -c tessedit_char_whitelist=0123456789'
        content = pytesseract.image_to_string(image)
        # print("content:\n", content)
        chatgpt_res = chatgpt_check(res_ocr=content)
        print(chatgpt_res)
        time.sleep(1)


def xlwings_example(filename):
    supplier = "SPEED COMMUNICATION NETWORKS LLC"
    payment_log = xlwings.Book("Z:\\Shared\\HQ\\13 IT\\2-IT Procurement\\5.Payment\\Payment Log.xlsx")
    paylog_sheet1 = payment_log.sheets[0]
    last_row = paylog_sheet1.used_range.last_cell.row
    paylog_sheet1.used_range.api.AutoFilter(Field := 10, Criteria1 := supplier, VisibleDropDown=False)
    last_cell = paylog_sheet1.range('A' + last_row)
    last_cell.value = filename
    last_cell.Hyperlinks.Add(Anchor=last_cell.api,
                             Address="Z:\\Shared\\HQ\\13 IT\\2-IT Procurement\\4.LPO\\2023\\" + filename + "xlsx")
    print(last_row)


def get_scanner():
    img_path = ".\\pdf2img\\"
    shutil.rmtree(img_path)
    os.mkdir(img_path)
    file_path = ".\\raw_pdf\\"
    file_list = os.listdir(file_path)
    img_name = "0"
    for lpo_file_name in file_list:
        page_num = pdf2png(lpo_file_name, img_name)
        img_name = str(int(img_name) + 1)
    file_list = os.listdir(img_path)
    format_dict = {
        "type": "",
        "project": "",
        "date": "",
        "invoice": "",
        "lpo": "",
        "trn": "",
        "supplier": "",
    }
    for lpo_file_name in file_list:
        # 正式识别内容
        image = Image.open(img_path + lpo_file_name)
        # cropped = image.crop((0, 0, 1785, 1780))
        # cropped.show()
        # 解析图片，lang='chi_sim'表示识别简体中文，默认为English
        # 如果是只识别数字，可再加上参数config='--psm 6 --oem 3 -c tessedit_char_whitelist=0123456789'
        content = pytesseract.image_to_string(image)
        # print("content:\n", content)
        chatgpt_res = chatgpt_check(res_ocr=content)
        ex_split = chatgpt_res.split("{")[-1]
        parttern_result = "{" + ex_split.split("}")[0] + "}"
        # print("parttern_result:", parttern_result)
        if len(parttern_result) > 0:
            str_dict = parttern_result.replace("\n", "").replace("'", '"').replace("\\", '')
            # print("str_dict:", str_dict)
            try:
                chatgpt_res_dict = json.loads(str_dict)
            except Exception as err:
                print("line151:", str_dict, err)
                continue
            if format_dict["type"] == "":
                format_dict["type"] = chatgpt_res_dict["type"]
            if format_dict["project"] == "":
                format_dict["project"] = chatgpt_res_dict["project"]
            if format_dict["date"] == "":
                format_dict["date"] = chatgpt_res_dict["date"]
            if format_dict["invoice"] == "":
                format_dict["invoice"] = chatgpt_res_dict["invoice"]
            if format_dict["lpo"] == "":
                format_dict["lpo"] = chatgpt_res_dict["lpo"]
            if format_dict["trn"] == "":
                format_dict["trn"] = chatgpt_res_dict["trn"]
            if format_dict["supplier"] == "":
                format_dict["supplier"] = chatgpt_res_dict["supplier"]
        else:
            pass
        time.sleep(1)
    return format_dict


def get_LPO(res_scan):
    only_year = datetime.today().strftime("%Y")
    supplier = res_scan["supplier"]
    department = "HQ"
    project = res_scan["project"]
    invoice = res_scan["invoice"]
    lpo = res_scan["lpo"]
    due_date = res_scan["date"]
    trn = res_scan["trn"]
    per_time = "90 Days"
    lpo_file_name = lpo.replace("/", "")
    # 各类格式的日期生成
    name_date = datetime.today().strftime("%d%b")
    send_date = datetime.today().strftime("%d-%b-%Y")
    work_form_date = datetime.today().strftime("%b/%Y")
    lpo_file_name_date = datetime.today().strftime("%d%b")
    path = "Z:\\Shared\\HQ\\13 IT\\2-IT Procurement\\4.LPO\\" + only_year + "\\"
    # workbook_1指的是LPO文件
    workbook_1 = xlwings.Book(path + lpo_file_name + ".xlsx")
    sheet1_demo = workbook_1.sheets["Rebar"]
    workbook_2 = xlwings.Book("PaymentCertificate" + ".xlsx")
    sheet1 = workbook_2.sheets["Attachment 1"]
    sheet2 = workbook_2.sheets["2Conf. of M. Received)"]
    # supplier = sheet1_demo.range('B16').value
    print("supplier:", supplier)
    for i in range(16, 30):
        discription = sheet1_demo.range('B' + str(i)).value
        quantity = sheet1_demo.range('E' + str(i)).value
        unit = sheet1_demo.range('F' + str(i)).value
        rate = sheet1_demo.range('G' + str(i)).value
        total = sheet1_demo.range('H' + str(i)).value
        if total is None:
            break
        sheet1.range('B' + str(i - 2)).value = str(i - 15)
        sheet1.range('C' + str(i - 2)).value = discription
        sheet1.range('E' + str(i - 2)).value = quantity
        sheet1.range('F' + str(i - 2)).value = unit
        sheet1.range('G' + str(i - 2)).value = rate
        sheet1.range('H' + str(i - 2)).value = total
        sheet1.range('I' + str(i - 2)).value = total
        sheet1.range('J' + str(i - 2)).value = "100%"
        sheet1.range('K' + str(i - 2)).value = total
        # 下面是2Conf. of M. Received)
        sheet2.range('A' + str(i - 9)).value = str(i - 15)
        sheet2.range('E' + str(i - 9)).value = discription
        sheet2.range('F' + str(i - 9)).value = discription
        sheet2.range('G' + str(i - 9)).value = unit
        sheet2.range('H' + str(i - 9)).value = quantity
        sheet2.range('J' + str(i - 9)).value = total
        sheet2.range('I' + str(i - 9)).value = rate
        # 根据单价判断是什么类型的商品
        low_disc = discription.lower()
        if "program" in low_disc or "license" in low_disc:
            goods_type = "software"
        else:
            if int(rate) < 2000:
                goods_type = "office supplies"
            else:
                goods_type = "office equipment"
        sheet2.range('B' + str(i - 9)).value = goods_type
    sheet2.range('A' + str(4)).value = "Headoffice/Division:                     HQ                              " \
                                       "Department or Project:       " + project + "           " \
                                                                                   "Place of Receipt：                             Date:    11/04/2023"
    # 根据表单上的公式得出总价格
    no_tax_total = sheet1.range('K' + str(30)).value
    lpo_total = no_tax_total * 1.05
    input_tax = no_tax_total * 0.05
    # 修改Payment log，不改的话就注释掉
    payment_log = xlwings.Book("Z:\\Shared\\HQ\\13 IT\\2-IT Procurement\\5.Payment\\Payment Log.xlsx")
    paylog_sheet1 = payment_log.sheets[0]
    paylog_sheet1.used_range.api.AutoFilter(Field := 10, Criteria1 := supplier, VisibleDropDown=False)
    data_pd = paylog_sheet1.range('A1').options(pandas.DataFrame, header=1, index=False, expand='table').value
    last_pd = data_pd[data_pd['Supplier'] == supplier]
    # print("last_pd:\n", last_pd)
    payment_number = last_pd.iloc[-1].tolist()
    last_number = payment_number[1]
    # 如果写错了重来不能用这个
    new_number = int(last_number + 1)
    new_number = 237
    last_row = paylog_sheet1.used_range.last_cell.row
    try:
        while last_row > 0:
            if paylog_sheet1.range('B' + str(last_row)).value is not None:
                new_row = last_row + 1
                paylog_sheet1.range('A' + str(new_row)).value = [
                    "Finance Form " + only_year + "hq-" + name_date + "-" + project, new_number]
                paylog_sheet1.range('D' + str(new_row)).value = lpo_total
                paylog_sheet1.range('E' + str(new_row)).value = [project, send_date, send_date, send_date, send_date,
                                                                 supplier]
                paylog_sheet1.range('A' + str(new_row)).color = (124, 252, 0)
                paylog_sheet1.range('B' + str(new_row)).color = (124, 252, 0)
                paylog_sheet1.range('D' + str(new_row)).color = (124, 252, 0)
                cell_A = 'A' + str(new_row)
                cell_M = "M" + str(new_row)
                # 设置内边线
                paylog_sheet1.range(cell_A, cell_M).api.Borders(9).lineStyle = 1
                paylog_sheet1.range(cell_A, cell_M).api.Borders(10).lineStyle = 1
                paylog_sheet1.range(cell_A, cell_M).api.Borders(11).lineStyle = 1
                # payment_log.save()
                # payment_log.close()
                break
            else:
                last_row = last_row - 1
    except Exception as err:
        print("5.Payment\\Payment Log.xlsx 可能被占用")
        print("err:", err)
    # 此时再回到付款确认单第一页填写内容
    sheet0 = workbook_2.sheets["Payment certificate"]
    sheet0.range('F' + str(5)).value = "PO Ref.No:     " + str(lpo)
    sheet0.range('F' + str(9)).value = "Payment Due Date:     " + str(due_date)
    sheet0.range('F' + str(13)).value = ("Payment Certificate No:    " + str(new_number))
    sheet0.range('A' + str(9)).value = "Name of Subcontractor/Supplier:     " + str(supplier)
    sheet0.range('A' + str(10)).value = "Tax Registration Number:     " + str(trn)
    sheet0.range('F' + str(15)).value = "Period  of  work  from:    " + str(work_form_date)
    sheet0.range('A' + str(7)).value = "Project Name:     HQ / IT"
    sheet0.range('A' + str(8)).value = str(project)
    sheet1.range('B' + str(7)).value = "Department:     " + str(project)
    sheet1.range('E' + str(7)).value = "     Date   :     " + str(send_date)
    sheet1.range('J' + str(7)).value = "For Payment Certificate No :     " + str(new_number)
    # 下面是纯税务Attachment4
    sheet4 = workbook_2.sheets["Attachment4"]
    sheet4.range('A' + str(5)).value = "Department:     HQ / IT"
    sheet4.range('C' + str(5)).value = None
    sheet4.range('D' + str(5)).value = "Date        :     " + str(send_date)
    sheet4.range('E' + str(5)).value = "For Payment Certificate No :     " + str(int(new_number))
    sheet4.range('B' + str(9)).value = str(invoice)
    sheet4.range('D' + str(9)).value = str(input_tax)
    sheet4.range('E' + str(9)).value = str(input_tax)
    # 关闭之前先打印 1,2,5,6
    # sheet0.api.PrintOut(Copies=1, ActivePrinter="\\172.16.40.2\Kyocera TASKalfa 5053ci KX(Copy Room)", Collate=True)
    # sheet1.api.PrintOut(Copies=1, ActivePrinter="\\172.16.40.2\Kyocera TASKalfa 5053ci KX(Copy Room)", Collate=True)
    # sheet2.api.PrintOut(Copies=1, ActivePrinter="\\172.16.40.2\Kyocera TASKalfa 5053ci KX(Copy Room)", Collate=True)
    # sheet4.api.PrintOut(Copies=1, ActivePrinter="\\172.16.40.2\Kyocera TASKalfa 5053ci KX(Copy Room)", Collate=True)
    # 顺便直接打印LPO
    pdf_lpo_name = ""
    lpo_file_name_list = os.listdir("Z:\Shared\HQ\\13 IT\\2-IT Procurement\\4.LPO\\" + only_year)
    for raw_lpo_name in lpo_file_name_list:
        if raw_lpo_name[-4:] == ".pdf":
            deal_lpo_file_name = raw_lpo_name.replace("LPO-", "").replace(".pdf", "")
            split_lpo_file_name = deal_lpo_file_name.split("-")
            if split_lpo_file_name[0] == lpo_file_name:
                pdf_lpo_name = raw_lpo_name
                # 直接打印不多BB
                # commend = '"D:\AdobeDC\Acrobat DC\Acrobat\Acrobat.exe" /t "Z:\Shared\HQ\\13 IT\\2-IT Procurement\\4.LPO\\' + only_year + '\\' + pdf_lpo_name + '" "\\\\172.16.40.2\Kyocera TASKalfa 5053ci KX(Copy Room)"'
                # print(commend)
                # os.system(commend)
                break
    # 关闭所有项目
    # workbook_1.close()
    pc_path = "Z:\\Shared\\HQ\\13 IT\\2-IT Procurement\\5.Payment\\1.Payment Certificate\\HQ\\" + only_year + "\\"
    final_pc_name = "Finance Form " + only_year + "hq-" + lpo_file_name_date + "-" + supplier + ".xlsx"
    print("final_pc_name：\n", pc_path + final_pc_name)
    final_pc = open(pc_path + final_pc_name, "w")
    final_pc.close()
    workbook_2.save(pc_path + final_pc_name)
    # 删掉不打印的页面
    workbook_2.sheets[7].delete()
    workbook_2.sheets[0].delete()
    printing_path = ".\\printing_xlsx\\"
    workbook_2.save(printing_path + final_pc_name)
    workbook_printing = xlwings.Book(printing_path + final_pc_name)
    workbook_1.close()
    webbrowser.open("Z:\Shared\HQ\\13 IT\\2-IT Procurement\\4.LPO\\" + only_year + "\\" + pdf_lpo_name)


def main():
    res_scan = get_scanner()
    get_LPO(res_scan)

    pprint.pprint(res_scan)


if __name__ == '__main__':
    main()
