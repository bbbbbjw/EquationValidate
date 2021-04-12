# Equation Validate

项目背景：学校作业需要检测word文档中的公式格式是否正确，主要包括1.公式是否居中，2.公式前只许有“解:”“假定“等且前置需两个空格，3.公式编号是否正确。由于word16版本以后自带的公式功能能很好的检测这些错误，本项目主要检测的是手动插入Mathtype公式且利用制表符实现”解“在最左边，公式居中，编号右对齐。

本项目只能检测.docx文件，.doc文件可以手动转存.docx或用脚本实现。

用到的第三方库：

org.w3c.dom

apache.poi

主要难点在于读懂word文档的content.xml文件

运行方式：

1.进入到readme所在的目录下 

2.

```
java -jar deal-equation-1.0-SNAPSHOT-shaded.jar sample1.docx
```

其中sample1.docx文件为各种文字错误，sample1_notatmid.docx为公式不居中，也可以使用自己的.docx文件进行检测。