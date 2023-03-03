import xlsx from "node-xlsx";

export type CellType = string | number | boolean;
export type ContentRow = CellType[];
export type HeaderRow = string[];
export type Sheet = {
    name: string;
    data: ContentRow[];
};
export type OutputCellType = "string" | "number" | "boolean" | "array";
export type SheetOutputType<T extends any = any> = {
    count: number;
    columns: Record<string, OutputCellType>;
    content: T[];
};
export class WorkSheetHandler {
    // sheet name
    readonly name: string;
    // raw data from node-xlsx
    readonly rawData: Sheet;
    // first line defines column
    columnNames: HeaderRow;
    // data type for each column
    dataTypes: Record<string, OutputCellType>;
    // the contents.
    contents: ContentRow[];
    mainSeparator: string = "\n";
    /**
     * parse a Sheet into the handler
     * @param workSheet sheet from node-xlsx
     */
    constructor(workSheet: Sheet) {
        this.name = workSheet.name;
        this.rawData = Object.assign(workSheet, {});
        this.columnNames = [];
        this.dataTypes = {};
        this.rawData.data[0].forEach((cell: CellType) => {
            this.columnNames.push(String(cell));
            this.dataTypes[String(cell)] = "string";
        });
        this.contents = [];
        /**
        lines like below will be ignored:
        the first line (as header)
        the first column is empty, or empty string
        the first column starts with #comment
         */
        this.rawData.data.forEach((line, index) => {
            if (
                index > 0 &&
                line[0] !== null &&
                line[0] !== undefined &&
                line[0].toString().trim().length > 0 &&
                !line[0].toString().startsWith("#comment")
            ) {
                this.contents.push(line);
            }
        });
    }
    /**
     * set array separator. default is \n. if a column is set to "array", the output of each cell in the column will be converted to array by splitting the cell value with the separator.
     * @param separator
     * @returns
     */
    setArraySeparator(separator: string) {
        this.mainSeparator = separator;
        return this;
    }
    /**
     * set the data type for one or many columns.
     * @param colNames one or more column names.
     * @param type one of "string", "number", "boolean", "array"
     * @returns
     */
    setType(colNames: string | string[], type: OutputCellType) {
        if (Array.isArray(colNames)) {
            colNames.forEach((colName) => {
                if (this.columnNames.includes(colName)) {
                    this.dataTypes[colName] = type;
                } else {
                    throw new Error(
                        `'${colName}' is not a header for this sheet`
                    );
                }
            });
        } else {
            if (this.columnNames.includes(colNames)) {
                this.dataTypes[colNames] = type;
            } else {
                throw new Error(`'${colNames}' is not a header for this sheet`);
            }
        }
        return this;
    }
    /**
     * convert the instance to an object. each cell will be transformed by the column type.
     * @returns Object
     */
    toObject<T extends any = any>(): SheetOutputType<T> {
        let result: any[] = [];

        this.contents.forEach((line) => {
            let obj: Record<string, any> = {
                raw: {},
            };
            this.columnNames.forEach((key, colIndex) => {
                const headerType = this.dataTypes[key];
                const cell = line[colIndex];

                if (headerType == "boolean") {
                    if (cell === undefined || cell === null) {
                        obj[key] = false;
                    } else {
                        obj[key] = ["TRUE", "true", 1, true].includes(
                            cell as string | number
                        );
                    }
                } else if (headerType == "number") {
                    if (cell === undefined || cell === null) {
                        obj[key] = 0;
                    } else {
                        obj[key] = Number(cell);
                    }
                } else if (headerType == "array") {
                    if (cell === undefined || cell === null) {
                        obj[key] = [];
                    } else {
                        obj[key] = cell.toString().split(this.mainSeparator);
                    }
                } else if (headerType == "string") {
                    if (cell === undefined || cell === null) {
                        obj[key] = "";
                    } else {
                        obj[key] = String(cell);
                    }
                }
            });
            result.push(obj);
        });

        return {
            columns: this.dataTypes,
            count: result.length,
            content: result as T[],
        };
    }
}
/**
 * Parse an xlsx file and return an object with all sheet instance.
 * @param input a buffer or a file path
 * @returns
 */
export const parseXlsx = (input: string | Buffer) => {
    const workBook = xlsx.parse(input);
    const result: Record<string, WorkSheetHandler> = {};
    workBook.forEach((sheet) => {
        const handler = new WorkSheetHandler(sheet as Sheet);
        result[sheet.name] = handler;
    });
    return result;
};
