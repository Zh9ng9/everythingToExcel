# everythingToExcel
把多种格式的数据转为Excel。支持txt、csv、Ehole输出的json。支持自定义表头

## 参数

'-t', '--type', string in {EholeJson,txt,csv}, 文件类型(default:EholeJson)
'-s', '--split', char, 分隔符 (default:\',\')
'-H', '--header', string, excel表头,以\'|\'分割 (default:空)
'-iF', '--inputFile', string 输入的文件名
'-oF', '--outputFile', string, 输出的文件名(后缀如果不是xls会自动添加xls)')

## 使用方法

将EholeJson转为Excel

```
python3 everythingToExcel.py -iF ./input.json -oF output.xls
```

将txt转为Excel，自定义表头

```
python3 everythingToExcel.py -t txt -iF ./input.txt -oF output.xls -H "序号|姓名|xxx|yyy|zzz"
```

## 使用截图：Ehole的输出结果json转Excel

将bat放在Ehole输出目录，双击运行，默认为result.json转为result.xls。

![1.jpg](https://s2.loli.net/2022/07/14/rBTJ63Vmy8t2n1G.png)

![2.jpg](https://s2.loli.net/2022/07/14/E8L1IT5sYuPlpMN.png)
