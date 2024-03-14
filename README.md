# EPPlus

使用EPPlus .net3.5 库来解析xlsx文件；（注意：不支持xls）

EPPlus下载依赖库地址（www.nuget.org）：

https://www.nuget.org/packages/EPPlus#dependencies-body-tab

# 数据类型

支持一下基本的数据类型：

`string`、`string[]`、`string[][]`、`int`、`int[]`、`int[][]`、`int10`、`int10[]`、`int10[][]`、`int100`、`int100[]`、`int100[][]`、`int1000`、`int1000[]`、`int1000[][]`、`bool`、`bool[]`、`bool[][]`、`local`、`local[]`、`local[][]`、`id`、`id[]`、`id[][]`

`***`表示的基本数值，`***[]`表示一维数组，`***[][]`表示二维数组；

`int10`：表示保留一位小数的浮点数；

`int100`：表示保留两位小数的浮点数；

`int1000`：表示保留三位小数的浮点数；

`id`：表示其它表格中的id，可以相互引用；

`local`：表示`Localization`表中的id；

