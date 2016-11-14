#-*- coding: utf-8 -*-

import codecs
import chardet
import time

# csv output elements
TITLE = {"name": u"收件人", "tel": u"电话", "address": u"地址",
         "price": u"价格", "num": u"数量", "time": u"下单时间",
         "id": u"订单号", "remark": u"备注", "title": u"产品名称"}
#TITLE_KEYS = ("name", "tel", "address", "price", "num", "time", "id", "remark", "title")
TITLE_KEYS = ("name", "tel", "address", "price", "num", 'title')

CSV_ROW = ""
for k in TITLE_KEYS:
    CSV_ROW = CSV_ROW + "%%(%s)s," % k
CSV_ROW = CSV_ROW[:-1] + '\n'

WORD_W = 3
END = u"发货"

TIME = u"201"
PROVINCE = (u'北京', u'天津', u'上海', u'重庆', u'河北', u'山西',
            u'辽宁', u'吉林', u'黑龙', u'江苏', u'浙江', u'安徽',
            u'福建', u'江西', u'山东', u'河南', u'湖北', u'湖南',
            u'广东', u'海南', u'四川', u'贵州', u'云南', u'陕西',
            u'甘肃', u'青海', u'台湾', u'内蒙', u'广西', u'西藏',
            u'宁夏', u'新疆', u'香港', u'澳门')
REMARK = u"备注"
ORDER_DETAIL = u"订单"
CONTACT = u"联系"
PRODUCT_MODEL = u"型号"
USELESS = (END, ORDER_DETAIL, CONTACT, PRODUCT_MODEL)

_ = lambda x: x.decode('utf-8')


def is_name_tel(order_line):
    name_tel = order_line.split(" ")
    if len(name_tel) == 2 \
        and is_tel(name_tel[1]):
        return True
    return False


def is_address(order_line):
    if order_line[:2] in PROVINCE:
        return True
    return False

def is_time(order_line):
    if order_line[:3] in TIME:
        return True
    return False

def is_remark(order_line):
    if order_line[:2] in REMARK:
        return True
    return False
    
def is_tel(order_line):
    if len(order_line) == 11 \
        and order_line.isdigit():
        return True
    return False


def order_line_to_dict(order):

    order_dict = {}
    for k in TITLE_KEYS:
        order_dict[k] = u""
    for line in order:
        if is_name_tel(line):
            name_tel = line.split(" ")
            order_dict["name"] = name_tel[0]
            order_dict["tel"] = name_tel[1]
        elif is_address(line):
            order_dict["address"] = line
        elif is_time(line):
            array = line.split(" ")
            order_dict["time"] = u"%s %s" % (array[0], array[1]) if "time" in TITLE_KEYS else u""
            order_dict["id"] = array[2][4:]
            order_dict["num"] = array[3][1:-3]
            order_dict["price"] = array[4][3:-10]
        elif is_remark(line):
            order_dict["remark"] = line[3:]
        elif line[:2] in USELESS or is_tel(line):
            continue
        else:
            order_dict["title"] = order_dict["title"] + line + ' '
    return order_dict


def detect_encoding(path):
    with open(path, 'rb') as f:
        contents = f.read()
        return chardet.detect(contents)['encoding']


def read_file(path, encoding):
    if encoding in ('UTF-16BE', 'UTF-16LE'):
        with codecs.open(path, 'r', encoding) as fp:
            lines = fp.readlines()
    else:
        with open(path, 'rb') as fp:
            lines = fp.readlines()
            lines = [l.decode(encoding, 'ignore') for l in lines]

    return lines


def parse_lines(lines):
    orders = []
    order = []
    start = False
    for line in lines:
        line = line.strip()
        if line[:3] in TIME:
            start = True
        if start:
            if line:
                order.append(line)
                l = line[:2]
                if l in END:
                    orders.append(order)
                    order = []
    return orders


def parse_orders(orders):
    # make the order line list to dict
    orders_list = []
    for item in orders:
        order_dict = order_line_to_dict(item)
        orders_list.append(order_dict)
    return orders_list


def covert_encode(data, encode="utf-8"):
    encode_data = {}
    for k, v in data.items():
        encode_data[k] = v.encode(encode, 'ignore')

    return encode_data


def write_csv(data, filename, title=None, encoding="gbk"):
    if not data:
        raise Exception("Error: No data to write file! Check whether the huiben.txt correct.")

    with open(filename, 'w') as outf:
        data_gbk = [covert_encode(row, encoding) for row in data]
        for row in data_gbk:
            outf.write(CSV_ROW % row)


def write_xlsx(data, filename='huiben.xlsx', headers=None):
    import xlsxwriter

    with xlsxwriter.Workbook(filename) as workbook:
        sheet1 = workbook.add_worksheet()
        thead = data[0]
        width = {'name': 7, 'tel': 12, 'address': 50, "price": 8,
                 "num": 4, "time": 20, "id": 16, "remark": 8, 'title': 8}

        # get table headers, if header is None, auto order
        if headers is None:
            headers = []
            for k, v in thead.items():
                headers.append(k)

        # set column width
        col = 0
        for col, k in enumerate(headers):
            sheet1.set_column(col, col, width[k])
        sheet1.autofilter(0, 0, 0, col)

        # write data
        for row, tdata in enumerate(data):
            for col, k in enumerate(headers):
                sheet1.write_string(row, col, tdata[k])

def check_output_opened(output):
    is_open = True
    while is_open:
        try:
            file_name = output+'xlsx'
            with open(file_name, 'a') as f:
                f.write("")
                is_open = False
        except IOError as e:
            is_open = True
            msg = "'%s' is Opened? close and retry" % file_name
            print(msg)
            wait()


def main(argv=None):

    if len(argv) == 2:
        path = argv[1]
    else:
        path = "huiben.txt"
        
    output = path[:-3]
    check_output_opened(output)

    encoding = detect_encoding(path)
    # separated the order file contents to list
    lines = read_file(path, encoding)
    orders = parse_lines(lines)
    orders_list = parse_orders(orders)

    # add order title
    orders_list.insert(0, TITLE)

    # write to file
    try:
        # write the orders to xlsx file
        file_name = output + 'xlsx'
        write_xlsx(orders_list, file_name, headers=TITLE_KEYS)
        print(u"export to %s." % file_name)
    except Exception as e:
        print(e)
        # write the orders to csv file
        file_name = output + 'csv'
        write_csv(orders_list, path[:-3]+'csv', encoding="gbk")
        print(u"export to %s." % file_name)


def wait():
    import msvcrt as m
    print("Press any key to continue...")
    m.getch()

import sys
if __name__ == "__main__":
    help = _("请将原始订单数据拷贝到huiben.txt文件里，\n"
    "从订单时间那一行开始拷贝，按住shift不放，拷贝到最后的发货结束。\n"
    "Enjoy it!\n"
    "")
    print(help)

    print(_("正在导出订单，请稍后。。。\n"))
    #time.sleep(1)
    try:
        main(sys.argv)
    except Exception as e:
        print(str(e))
        #time.sleep(5)
    else:
        print(_("导出订单成功。"))
        #time.sleep(5)

    wait()
