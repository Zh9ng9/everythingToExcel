#-*- coding: UTF-8 -*-
import argparse
import time
import json
import traceback

import xlwt

SET_DEAULT = "\033[0m"
SPECIAL_PURPLE = "\033[0;35m[*] "
INFO_BLUE = "\033[0;34m[i] "
TIME_GREEN = "\033[7;32m[+] "
ERROR_RED = "\033[0;31m[!] "

def EholeJsonData():
    global args
    result = []
    with open(args.inputFile, encoding="utf8") as f:
        origin = f.readlines()
    for line in origin:
        line_result = []
        line_json = json.loads(line.strip())
        for item in args.header:
            # 判断是否含key，没有的话放空字符串
            if line_json[item] == "None" or line_json[item] == None:
                line_result.append("")
            else:
                line_result.append(str(line_json[item]))
        result.append(line_result)
    return result


def txtData():
    global args
    result = []
    with open(args.inputFile, encoding="utf8") as f:
        origin = f.readlines()
    for line in origin:
        result.append(line.strip().split(args.split))
    return result


def csvData():
    global args
    result = []
    with open(args.inputFile, encoding="utf8") as f:
        origin = f.readlines()
    for line in origin:
        result.append(line.strip().split(args.split))
    return result


def main():
    # 程序的特殊信息
    print(SPECIAL_PURPLE + 'everythingToExcel')
    print(SPECIAL_PURPLE + '把多种格式的数据转为Excel')
    print(SPECIAL_PURPLE + 'Write by zh9ng9 <\033[5;35mhttps://github.com/Zh9ng9/everythingToExcel\033[0;35m> (202207)')
    print(SPECIAL_PURPLE + '''Example:\n    everythingToExcel.py -iF ./input.json -oF output.xls\n    everythingToExcel.py -t txt -iF ./input.txt -H "序号|姓名|xxx|yyy|zzz" -oF output.xls''')
    print(SET_DEAULT)
    # 参数配置
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--type', required=False, type=str, default='EholeJson',
                        choices=['EholeJson', 'txt', 'csv'],
                        help='string in {EholeJson,txt,csv} 文件类型(default:EholeJson)')
    parser.add_argument('-s', '--split', required=False, type=str, default=",", help='char 分隔符 (default:\',\')')
    parser.add_argument('-H', '--header', required=False, type=str, default=None,
                        help='string excel表头,以\'|\'分割 (default:空)')
    parser.add_argument('-iF', '--inputFile', required=True, type=str, default=None, help='string 输入的文件名')
    parser.add_argument('-oF', '--outputFile', required=True, type=str, default=None,
                        help='string 输出的文件名(后缀如果不是xls会自动添加xls)')
    # args作为全局变量
    global args, data
    args = parser.parse_args()
    # 开始时间
    start = time.time()
    # 打印输出获取到的参数
    print(INFO_BLUE + "文件类型：" + str(args.type))
    print(INFO_BLUE + "分隔符：" + str(args.split))
    print(INFO_BLUE + "excel表头：" + str(args.header))
    print(INFO_BLUE + "输入的文件名：" + str(args.inputFile))
    if not args.outputFile.endswith(".xls"):
        args.outputFile = args.outputFile + ".xls"
    print(INFO_BLUE + "输出的文件名：" + str(args.outputFile))
    print(SET_DEAULT)
    # 打印输出开始运行提示
    print(TIME_GREEN + "程序开始运行..." + SET_DEAULT)
    try:
        # xlwt-workbook 建立
        workbook = xlwt.Workbook(encoding='utf-8')
        # xlwt-sheet 建立
        sheet = workbook.add_sheet('GoodLuck')
        # 表头的修改
        if args.header:
            args.header = args.header.split('|')
        # 根据不同的type从不同函数获取数据
        if args.type == "txt":
            data = csvData()
        elif args.type == "csv":
            data = txtData()
        elif args.type == "EholeJson":
            # EholeJson的表头
            header = ["url", "cms", "server", "statuscode", "length", "title"]
            # EholeJson的表头写入header参数
            args.header = header
            data = EholeJsonData()
        # 若数据为空，报错
        if data == None or data == []:
            raise Exception("输入文件的数据不能为空")
        # 写入表头（如果表头存在）
        if args.header:
            for header_j in range(len(args.header)):
                sheet.write(0, header_j, args.header[header_j])
        # 最大列数
        cols_count = 0
        for i in data:
            cols_count = len(i) if len(i) > cols_count else cols_count
        # 记录每列最大宽度
        col_list = [len(args.header[x].encode('gb18030')) for x in range(len(args.header))]
        # 数据写入Excel
        for i in range(len(data)):
            for j in range(len(data[i])):
                # 写入
                sheet.write(i + 1, j, data[i][j])
                # 找更大的值
                col_list[j] = len(data[i][j].encode('gb18030')) if len(data[i][j].encode('gb18030')) > col_list[j] else col_list[j]
        for i in range(0, len(col_list)):
            # 256*字符数得到excel列宽
            sheet.col(i).width = 255 * (col_list[i] + 2)
        # 保存xls
        workbook.save(args.outputFile)
    except Exception as e:
        print(ERROR_RED + "程序出现错误,已停止")
        print(traceback.format_exc())

    end = time.time()
    print(TIME_GREEN + "程序用时: " + str(end - start) + SET_DEAULT)


if __name__ == '__main__':
    main()
