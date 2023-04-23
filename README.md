# everythingToExcel
Ehole、Yasso、Fscan扫描输出结果转Excel，同时支持txt、csv自定义表头生成Excel。

Fscan的转换简单集成了Sma11New师傅的脚本[Sma11New/fscanAux: fscan结果处理辅助脚本，整理分类输出为Excel文件，方便查看 (github.com)](https://github.com/Sma11New/fscanAux)

## 参数

'-t', '--type', string in {EholeJson,YassoJson,FscanTxt,txt,csv}, 文件类型(default:EholeJson)

'-s', '--split', char, 分隔符 (default:\',\')

'-H', '--header', string, excel表头,以\'|\'分割 (default:空)

'-iF', '--inputFile', string 输入的文件名

'-oF', '--outputFile', string, 输出的文件名(后缀如果不是xls会自动添加xls)')

## 使用方法

将EholeJson转为Excel

```
python3 everythingToExcel.py -iF ./input.json -oF output.xls
```

将YassoJson转为Excel

```
python3 everythingToExcel.py -t YassoJson -iF ./input.json -oF output.xls
```

将Fscan扫描结果txt转为Excel

```
python3 everythingToExcel.py -t FscanTxt -iF ./input.json -oF output.xls
```

将txt转为Excel，自定义表头

```
python3 everythingToExcel.py -t txt -iF ./input.txt -oF output.xls -H "序号|姓名|xxx|yyy|zzz"
```

## 使用截图：Ehole的输出结果json转Excel

将bat放在Ehole输出目录，双击运行，默认为result.json转为result.xls。

![1.jpg](https://s2.loli.net/2022/07/14/rBTJ63Vmy8t2n1G.png)

![2.jpg](https://s2.loli.net/2022/07/14/E8L1IT5sYuPlpMN.png)
