# vulpe-xlsx

parse xlsx worksheet into an object array. each column can be set to be parsed to following type of data:
* `string` - default
* `number`
* `boolean`
* `array` - array of string, splitted by customizable separator.

this module works in node environment, and thanks to `node-xlsx` for its fundemental contribution.

# usage

```javascript
import {parseXlsx} from 'vulpe-xlsx';

// book is an object indexed by sheetName, and valued by the WorksheetHandler class.
const book = parseXlsx('./someTestXlsx.xlsx');

// and this is a WorksheetHandler instance.
const sheet1 = book['Sheet1'];

// the 'age' column will be parsed as number.
sheet1.setType('age','number');

// and the 'titles' and 'feats' column will be parsed as string array.
//notice the setType method always return the instance itself, so you can chain it.
sheet1.setType(['titles','feats'],'array')
    .setType('isVulpe','boolean')

// if the array shall be splitted by separator other than \n, use this:
sheet1.setArraySeparator('|');
// notice the separator is global in the sheet.

// make an output
const output = sheet1.toObject();

```

# documentation

## sheet columns
The first line of sheet will be considered as 'header', a.k.a. columnNames. only headered column will be considered as valid column.

if column name happens to be duplicate, you may lose your data.

Column names will be force-parsed into string, because it will be used as indexes for output objects.

## sheet rows
The data row will start from line 2 (the first line is considered header). Only valid row will be considered as data row. 

A valid row should be like:
* the first cell must exist (not null, not undefined)
* the first cell string length must greater than 0.
* the first cell must not start with `#comment`. this keyword is for your description text.


## data types
use `setType` to declare output type.
### string
The default type. Cells in string column will be parsed with `String(cell)` function. If the cell happend to be undefined or null, an empty string will be returned.

### number
Set a column type to `number` to parse the content with`Number(cell)` function. If the cell happened to be undefined or null, `0` will be returned.

### boolean
Set a column type to `boolean` to parse the content to a boolean value. values like below will be parsed to true:
* string 'TRUE','true'
* number 1
* boolean true

other values (including undefined and null) will be considered false.

### array
Set a column type to `array` to parse content to a string array. The default separator is `\n`,which means each line in the cell is considered an element. 

If the original cell data is number, it will be force-parsed to string then split.

Use `setArraySeparator(sep)` to use another separator (like `|`,`,`).


## using other data sources
the `parseXlsx()` function accepts xlsx file path or Buffer. If you have other types of data (like `csv`), you can try parse it into a bi-dimension array, then pack it into a `Sheet` object(which is also exported by this module). Then handle it with `WorkSheetHandler` class.

```javascript
import {WorkSheetHandler} from 'vulpe-xlsx';
import myCustomCsvData from 'someOtherFile';

const forkThisSheet = {
    // the name doesn't matter.
    name:'mySheetName',
    // and the data must be an array of array of any
    data:myCustomCsvData,
}
const sheetHandler = new WorksheetHander(forkThisSheet)

// then you can do anything with this sheet.
```

# compability with the many-excel-like-things
As I know there is a lot of xlsx editor. The output file might have a little differences, and I tried my best to meet the compability. the known difference might be:
* empty cells might be treated as `null`, `undefined`, or empty string.
* the return mark inside cell might be `\n` or `\r\n`.
* some editor (like kingsoft wps) might wrap content in direction mark

And please pay extra attension to excel Date format, it will output a float number, which we commonly not using.