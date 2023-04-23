#!/usr/bin/python3.7
# -*- coding: utf-8 -*-
# @Author : Sma11New
# @Github : https://github.com/Sma11New

import os, re, sys, time, openpyxl
import traceback

from openpyxl.styles import Font

from colorama import init
init(autoreset=True)

class NewPrint:
    @staticmethod
    def info(data):
        print(f"[\033[36m*\033[0m] \033[36m{data}\033[0m")
    @staticmethod
    def success(data):
        print(f"[\033[32m+\033[0m] \033[32m{data}\033[0m")
    @staticmethod
    def error(data):
        print(f"[\033[31m!\033[0m] \033[31m{data}\033[0m")

# 读取result文件内容
def readFile(file):
    dataList = []
    with open(file, "r", encoding="utf8") as f:
        for i in f.readlines():
            dataList.append(i.strip())
    with open(file, "r", encoding="utf8") as f:
        dataStr = f.read()
    return dataList, dataStr

# 开放端口
def parsePortInfo(dataList):
    resultList = [["IP", "Port", "IP:Port"]]
    pattern = re.compile(r"^\d+\.\d+\.\d+\.\d+:\d+")
    for data in dataList:
        res = re.findall(pattern, data)
        if res:
            resultList.append([res[0].split(":")[0], res[0].split(":")[1], res[0]])

    writeCsvFile("开放端口", resultList)
    NewPrint.info(f"开放端口：{len(resultList) - 1}")

# web资产
def parseWebInfo(dataList):
    resultList = [["URL", "Code", "Length", "Title", "CmsInfo"]]

    # WebTitle
    patternList = [
        re.compile(r"http[^\s]+"),
        re.compile(r"code.\d+"),
        re.compile(r"len.\d+"),
        re.compile(r" title.*")
    ]

    delDataList = ["", "code:", "len:", "title:"]

    pattern = re.compile(r".*WebTitle.*")
    for data in dataList:
        res = re.findall(pattern, data)
        if res:
            tmp = []
            for patt, delData in zip(patternList, delDataList):
                tmp.append(re.findall(patt, data)[0].replace(delData, ""))
            resultList.append(tmp + [""])

    # InfoScan 匹配指纹
    pattern = re.compile(r".*InfoScan.*")
    for data in dataList:
        res = re.findall(pattern, data)
        if res:
            url = re.findall(patternList[0], data)[0]
            title = res[0].split(url)[1].strip()
            # 匹配并添加指纹(CmsInfo列)
            for i in range(1, len(resultList)):
                resultList[i][4] = title if resultList[i][0] == url else None

    writeCsvFile("Web资产", resultList)
    NewPrint.info(f"Web资产：{len(resultList) - 1}")

# 弱口令信息
def parsePasswordInfo(dataList):
    resultList = [["Service", "IP", "Port", "UserName", "PassWord"]]
    pattern = re.compile(r"((ftp|mysql|mssql|smb|rdp|Postgres|SSH|mongodb|oracle):.*)")
    for data in dataList:
        res = re.findall(pattern, data)
        if res:
            tmp = list(res[0][0].split(":"))
            if "//" in tmp[1]:
                tmp[1] = tmp[1].replace("//", "")
            try:
                userName, passWord = tmp[-1].split(" ")[0], tmp[-1].split(" ")[1]
            except:
                userName, passWord = tmp[-1], None
            tmp[-1] = userName
            tmp.append(passWord)
            resultList.append(tmp)

    writeCsvFile("弱口令", resultList)
    NewPrint.info(f"弱口令：{len(resultList) - 1}")

# 漏洞信息
def parseVulnInfo(dataList):
    try:
        resultList = [["Address", "Vuln"]]
        pattern = re.compile(r"(^\[\+].(?!(ftp|mysql|mssql|smb|rdp|Postgresql|SSH|mongodb|oracle|Info)).*$)")
        addressPatt = re.compile(r"((https?://)?\d+\.\d+\.\d+\.\d+(:\d+)?)")
        for data in dataList:
            res = re.findall(pattern, data)
            if res:
                address = re.findall(addressPatt, res[0][0])[0][0]
                vuln = res[0][0].replace(address, "").replace("[+]", "").replace("\t", "").strip()
                resultList.append([address, vuln])

        writeCsvFile("漏洞", resultList)
        NewPrint.info(f"漏洞：{len(resultList) - 1}")
    except Exception as e:
        print(traceback.format_exc())

# NetBios连接
def parseNetInfo(dataStr):
    resultList = [["IP", "NetBios"]]
    pattern = re.compile(r"(NetInfo.(\s+.*\n)+)")
    res = re.findall(pattern, dataStr)
    if res:
        patternIP = re.compile(r"\d+\.\d+\.\d+\.\d+")
        for i in res:
            ip = re.findall(patternIP, i[0])[0]
            resultList.append([ip, i[0].replace("NetInfo:\n", "")])

    writeCsvFile("NetBios连接", resultList)
    NewPrint.info(f"NetBios连接：{len(resultList) - 1}")

# 写csv文件
def writeCsvFile(sheetName, dataList):
    # 数据排序
    title = dataList.pop(0)
    dataList.sort(key=lambda x: x[0])
    dataList.insert(0, title)
    # 写数据
    sheet = resCsvFileObj.create_sheet(sheetName)
    for data in dataList:
        sheet.append(data)
    # 首行添加格式（字体大小、加粗）
    for row in sheet[f"A1:{chr(65 + len(dataList[0]) - 1)}1"]:
        for cell in row:
            cell.font = Font(size=12, bold=True)

def getInput():
    if len(sys.argv) != 2:
        print("\nfscan结果整理脚本，输出为.xlsx文件\n\nUsage: \n    python3 fscanAux.py [fscanResultFile]\n")
        exit()
    if not os.path.exists(sys.argv[1]):
        NewPrint.error(f"[{sys.argv[1]}] 文件不存在")
        exit()
    return sys.argv[1]

if __name__ == "__main__":
    dataList, dataStr = readFile(getInput())
    NewPrint.success("文件已读取，结果处理中……")
    resCsvFile = f"fscanAuxResult_{time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime())}.xlsx"
    resCsvFileObj = openpyxl.Workbook()
    parsePortInfo(dataList)
    parseWebInfo(dataList)
    parsePasswordInfo(dataList)
    parseVulnInfo(dataList)
    parseNetInfo(dataStr)
    del resCsvFileObj["Sheet"]
    resCsvFileObj.save(resCsvFile)
    NewPrint.success(f"处理结果已保存至：{resCsvFile}")