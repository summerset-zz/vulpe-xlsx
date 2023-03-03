# vulpe-xlsx

parse xlsx worksheet into an object array. each column can be set to be parsed to following type of data:
* `string` - default
* `number`
* `boolean`
* `array` - array of string, splitted by customizable separator.

this module works in node environment, and thanks to `node-xlsx` for its fundemental contribution.

将excel文件转换为一个对象数组。可以设置表格中每一列的转换方式（string,number,boolean或array）。


# usage 使用方法

Suppose you have an xlsx file with a single sheet 'Sheet1', the content is like below:
| name     | age | titles | feats                 | isVulpe |
| -------- | --- | ------ | --------------------- | ------- |
| Tachanka | 24  | Lord   | bigBadGun,defender    | FALSE   |
| Ash      | 27  | Flash  | infiniteDash,attacker | TRUE    |


```javascript
import {parseXlsx} from 'vulpe-xlsx';

// book is an object indexed by sheetName, and valued by the WorksheetHandler class.
// book是一个以工作表名为key、工作表管理器为value的对象。
const book = parseXlsx('./someTestXlsx.xlsx');

// and this is a WorksheetHandler instance.
// sheet1是代表工作表Sheet1的工作表管理器
const sheet1 = book['Sheet1'];

// the 'age' column will be parsed as number.
// 将age列设置为数字
sheet1.setType('age','number');

// and the 'titles' and 'feats' column will be parsed as string array,and 'isVulpe' column as boolean
// notice the setType method always return the instance itself, so you can chain it.
// 将titles列和feats列设置为数组，并将isVulpe列设置为布尔型。
// setType方法总是返回管理器实例本身，所以可以使用链式语法。
sheet1.setType(['titles','feats'],'array')
    .setType('isVulpe','boolean')

// if the array shall be splitted by separator other than \n, use this:
// 数组默认会以回车为分隔方式，你可以使用setArraySeparator方法更换为别的分隔符。
sheet1.setArraySeparator(',');
// notice the separator is global in the sheet.
// 分隔符是整个工作表共用的。

// make an output
// 输出
const output = sheet1.toObject();

```
you shall get an object like this:

输出结果如下：
```javascript
{
    count:2,
    columns:{
        name:'string',
        age:'number',
        titles:'array',
        feats:'array',
        isVulpe:'boolean',
    },
    content: [
        {
            name:'Tachanka',
            age:24,
            titles:['Lord'],
            feats:['bigBadGun','defender'],
            isVulpe:false
        },
        {
            name:'Ash',
            age:27,
            titles:['Flash'],
            feats:['infiniteDash','attacker'],
            isVulpe:true
        }
    ]

}

```


# documentation 详细文档

## sheet columns 列的定义
The first line of sheet will be considered as 'header', a.k.a. columnNames. only headered column will be considered as valid column.

if column name happens to be duplicate, you may lose your data.

Column names will be force-parsed into string, because it will be used as indexes for output objects.

工作表的第一行会被视为标题行。只有有标题的列才会被视为数据。

如果标题出现重复，则可能会出现数据覆盖的情况。

标题会被用作输出对象的索引，因此会被强制转为string类型。


## sheet rows 行的定义
The data row will start from line 2 (the first line is considered header). Only valid row will be considered as data row. 

A valid row should be like:
* the first cell must exist (not null, not undefined)
* the first cell string length must greater than 0.
* the first cell must not start with `#comment`. this keyword is for your description text.

工作表管理器会从第二行开始（index为1）进行数据行判断。有效的数据行如下：
* 第一格不为null或undefined
* 第一格不为空字符串（强制转化为文本时不为空字符串）
* 第一格不以`#comment`开头。以该关键字开头时，整行会被视为注释。


## data types 数据类型
use `setType` to declare output type. The syntax is:

使用`setType`方法指定一列的输出类型。语法如下：
```javascript
// set one column to number 将一列设置为number
sheet.setType('colName','number')

// set multiple column to number 将多列设置为number
sheet.setType(['colName1','colName2'],'number')
```


### string
The default type. Cells in string column will be parsed with `String(cell)` function. If the cell happend to be undefined or null, an empty string will be returned.

默认类型。数据会使用`String(cell)`方法强制转化为字符串。空或异常的单元格会视为空字符串。

### number
Set a column type to `number` to parse the content with`Number(cell)` function. If the cell happened to be undefined or null, `0` will be returned.

数据会使用`Number(cell)`方法强制转化为数字。空或异常的单元格会视为0。

### boolean
Set a column type to `boolean` to parse the content to a boolean value. values like below will be parsed to true:
* string 'TRUE','true'
* number 1
* boolean true

other values (including undefined and null) will be considered false.

该列满足以下条件的单元格将视为true：
* 字符串型'TRUE'或'true'
* 数字型1
* 布尔型true (excel中居中大写TRUE会视为布尔型true)

其他情况均视为false（包括空或异常单元格）。


### array
Set a column type to `array` to parse content to a string array. The default separator is `\n`,which means each line in the cell is considered an element. 

If the original cell data is number, it will be force-parsed to string then split.

Use `setArraySeparator(sep)` to use another separator (like `|` or `,`).

Empty cell will return empty array.

该列数据会根据分隔符转换为字符串数组。例如初始状况下分隔符为回车（`\n`），则单元格内每一行均视为输出数组的一个元素。如果原始数据不是字符串型，则处理时会先强制转为字符串型。

空单元格或异常单元格会被转为空数组（`[]`）。

使用`setArraySeparator(sep)`方法可以将整个工作表的分隔符转为你需要的分隔符（例如`|`或`,`）。



## using other data sources 使用其他数据来源
the `parseXlsx()` function accepts xlsx file path or Buffer. If you have other types of data (like `csv`), you can try parse it into a bi-dimension array, then pack it into a `Sheet` object(which is also exported by this module). Then handle it with `WorkSheetHandler` class.

上文中提到的`parseXlsx()`是用来处理整个工作簿（单个xlsx文件）的，可接受文件路径或一个Buffer对象。如果你的数据来自非xlsx文件（例如`csv`），你可以将数据先处理为二维数组，然后将其包装成本库提供的`Sheet`类型，然后直接由`WorkSheetHandler`类进行处理。

```javascript
import { WorkSheetHandler } from 'vulpe-xlsx';
import myCustomCsvData from 'someOtherFile';

const forkThisSheet = {
    // the name doesn't matter. 名字无所谓
    name:'mySheetName',
    // and the data must be an array of array of any 这是导入的二维数组
    data:myCustomCsvData,
}
const sheetHandler = new WorksheetHander(forkThisSheet)

// then you can do anything with this sheet. 然后 sheetHandler可以使用上文提到的各种方法。
```

# compability with the many-excel-like-things 兼容性问题
As I know there is a lot of xlsx editor. The output file might have a little differences, and I tried my best to meet the compability. the known difference might be:
* empty cells might be treated as `null`, `undefined`, or empty string.
* the return mark inside cell might be `\n` or `\r\n`.
* some editor (like kingsoft wps) might wrap content in direction mark

And please pay extra attension to excel Date format, it will output a float number, which we commonly not using.

xlsx有很多编辑器（如微软的excel、openOffice、WPS等），这些编辑器输出的内容可能具有一定差异，本库已尽可能的做了一些兼容,但还是请注意：
* 编辑器里的空单元格有可能被输出为null，undefined或者空字符串（所以本库做了一些兼容）
* 单元格里的回车可能是`\n`，也可能是`\r\n`
* 有些编辑器还会强制给你的文本加一组unicode的书写方向标识（通常不可见但实际上数据里有）

另外请格外注意在excel里的日期类单元格，这类单元格始终会输出浮点数（而不是时间戳什么的）。

# finale 结语

Feel free to ask questions on github page https://github.com/summerset-zz/vulpe-xlsx/

如有问题可在github页中提出~