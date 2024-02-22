import {
  AugmentedFieldType,
  FIMType,
  FieldMap,
  MatrixArrayType,
  MatrixType,
  SheetFieldType,
  TimeItResponse,
} from '@core/types/addon';
import { type Dayjs } from 'dayjs';

/**
 * Collection of utility functions to use with the add-on
 *
 * @namespace
 */
namespace Utils {
  export const usd = (n: number): string => currencygs(n).format();

  export const shortDate = (d?: Date | number | string | Dayjs): string =>
    daygs(d).format('YYMMDD');

  export function toCamelCase(header: string): string {
    return header
      .replace('#', 'Number')
      .replace(/([/_])(.)/g, ' $2')
      .replace(/^\W|\W$/, '')
      .replace(/[^a-zA-Z0-9_\s]+/gi, '')
      .toLowerCase()
      .split(/\s+/)
      .filter((v) => v !== '')
      .map((v, i) => (i ? v.slice(0, 1).toUpperCase() + v.slice(1) : v))
      .join('');
  }

  export function toCapCase(header: string): string {
    return header
      .replace('#', 'Number')
      .replace(/([/_])(.)/g, ' $2')
      .replace(/^\W|\W$/, '')
      .replace(/[^a-zA-Z0-9_\s]+/gi, '')
      .toUpperCase()
      .split(/\s+/)
      .filter((v) => v !== '')
      .join('_');
  }

  export function toHeader(str: string): string {
    return str
      .replace(/([a-z])([A-Z0-9])/g, '$1 $2')
      .replace(/([A-Z])_([A-Z0-9])/g, '$1 $2')
      .replace(/number/gi, '#')
      .split(' ')
      .map((v) => v.slice(0, 1).toUpperCase() + v.slice(1).toLowerCase())
      .join(' ');
  }

  export function isJsonString(value: string): boolean {
    // possibly a json string
    return (
      (value.slice(0, 1) === '{' && value.slice(-1) === '}') ||
      (value.slice(0, 1) === '[' && value.slice(-1) === ']')
    );
  }

  export function isUrl(value: string): boolean {
    // possibly a url
    return value.slice(0, 7) === 'http://' || value.slice(0, 8) === 'https://';
  }

  /**
   * convert a data into a suitable format for API
   * @param {Date} dt the date
   * @return {string} converted data
   */
  export function gaDate(dt: Date): string {
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  /**
   * Get the letter associated with column number
   * @param {number} col the column number starting at 1
   * @returns {string}
   */
  export function colToABC(col: number): string {
    let temp: number,
      letter = '';
    while (col > 0) {
      temp = (col - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      col = (col - temp - 1) / 26;
    }
    return letter;
  }

  export function colFromABC(abc: string): number {
    const rubrik = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    return abc
      .split('')
      .map((x, i) => rubrik.indexOf(x) + (i ? 0 : 1))
      .reverse()
      .map((x, i) => x * (26 ^ i))
      .reduce((t, c) => t + c);
  }

  /**
   * Returns the value of the last element in the array where predicate is true, and undefined
   * otherwise. It's similar to the native find method, but searches in descending order.
   * @param list the array to search in.
   * @param predicate find calls predicate once for each element of the array, in descending
   * order, until it finds one where predicate returns true. If such an element is found, find
   * immediately returns that element value. Otherwise, find returns undefined.
   */
  export function findLast<T>(
    list: Array<T>,
    predicate: (value: T, index: number, obj: T[]) => unknown,
  ): T | undefined {
    for (let index = list.length - 1; index >= 0; index--) {
      let currentValue = list[index];
      let predicateResult = predicate(currentValue, index, list);
      if (predicateResult) {
        return currentValue;
      }
    }
    return undefined;
  }

  /**
   * Fill in blank cells with first mask that contains a value in that cell
   * @param {Array} params
   * @param {MatrixType} params.rows
   * @param {...MatrixType} params.masks
   * @returns {MatrixType}
   *
   * | 1,  2,  3 |   |'', '', '' |   |'', 10, '' |     | 1,  2,  3 |
   * |'',  5, '' |   |'', '',  6 |   | 4, '', '' |  =  | 4,  5,  6 |
   * | 7,  8,  9 |   |'', '', '' |   |'', '', '' |     | 7,  8,  9 |
   *     rows            mask1          mask2
   */
  export function maskFill([rows, ...masks]: MatrixArrayType): MatrixType {
    console.log('maskFill');
    // return this.reduce2d(rows, masks, (t, c) => t || c)
    return rows.map((row, i) =>
      row.map((cell, j) => masks.reduce((t, c) => t || c[i][j], cell)),
    );
  }

  /**
   * Overwrites cell values with value from last mask containing value for that cell
   * @param {Array} params
   * @param {MatrixType} params.rows
   * @param {...MatrixType} params.masks
   * @returns {MatrixType}
   *
   *
   * | 1,  2,  3 |   |'', '', '' |   |'', '', '' |     | 1,  2,  3 |
   * | 4,  5,  6 |   |'', '', 10 |   |'', '', 99 |  =  | 4,  5, 99 |
   * | 7,  8,  9 |   |'', '', '' |   |'', '', '' |     | 7,  8,  9 |
   *     rows            mask1          mask2
   */
  export function maskOverwrite([rows, ...masks]: MatrixArrayType): MatrixType {
    // return this.reduce2d(rows, masks, (t, c) => c || t)
    console.log('maskOverwrite');
    return rows.map((row, i) =>
      row.map((cell, j) => masks.reduce((t, c) => c[i][j] || t, cell)),
    );
  }

  /**
   * This method uses the masks to punch out holes in given arr when value exists in mask's cell
   * @param {Array} params
   * @param {MatrixType} params.rows
   * @param {...MatrixType} params.masks
   * @param {boolean=} useNull whether to use null or '' when creating holes in the matrix (default=true)
   * @returns {MatrixType}
   *
   * | 1,  2,  3 |   |'', 99, '' |   |'', '', '' |     | 1, '',  3 |
   * | 4,  5,  6 |   |'', '',  6 |   | 4, '', '' |  =  |'',  5, '' |
   * | 7,  8,  9 |   |'', '', '' |   |'', '', '' |     | 7,  8,  9 |
   *     rows            mask1          mask2
   */
  export function maskPunch(
    [rows, ...masks]: MatrixArrayType,
    useNull?: boolean,
  ): MatrixType {
    console.log('maskPunch');
    const blank = !useNull ? '' : null;
    // return this.reduce2d(rows, masks, (t, c) => (!!c ? blank : t))
    return rows.map((row, i) =>
      row.map((cell, j) =>
        masks.reduce((t, c) => (!!c[i][j] ? blank : t), cell),
      ),
    );
    //  return rows.map(function (row, i) { return row.map(function (cell, j) { return masks.reduce(function (t, c) { return (!!c[i][j] ? blank : t); }, cell); }); });
    // console.log({ matrix, ...masks, punchedMatrix })
    // return punchedMatrix
  }

  /**
   * This method cuts out a hole if the cell has a value but masks do not have value in that cell
   * @param {Array} params
   * @param {MatrixType} params.rows
   * @param {...MatrixType} params.masks
   * @param {boolean=} useNull whether to use null or '' when creating holes in the matrix (default=true)
   * @returns {MatrixType}
   *
   * | 1,  2,  3 |   | 1, '',  3 |   | 7,  8,  9 |     | 1,  2,  3 |
   * | 4,  5,  6 |   | 4,  5,  6 |   | 4,  5, '' |  =  | 4,  5,  6 |
   * | 7,  8,  9 |   | 7, '',  9 |   | 1, '',  3 |     | 7, '',  9 |
   *     rows            mask1          mask2
   */
  export function maskCut(
    [rows, ...masks]: MatrixArrayType,
    useNull?: boolean,
  ): MatrixType {
    console.log('maskCut');
    const blank = !useNull ? '' : null;
    return rows.map((row, i) =>
      row.map((cell, j) =>
        masks.reduce((t, c) => t && c[i][j] === blank, cell) ? blank : cell,
      ),
    );
  }

  export function forEach2d(matrix: MatrixType, fn): void {
    matrix.forEach((row, i) => row.forEach((cell, j) => fn(cell, i, j)));
  }

  export function map2d(matrix: MatrixType, fn): MatrixType {
    return matrix.map((row, i) =>
      row.map((cell, j) => (!fn ? cell : fn(cell, i, j))),
    );
  }

  export function reduce2d(
    initialMatrix: MatrixType,
    masks: MatrixArrayType,
    fn,
  ) {
    return initialMatrix.map((row, i) =>
      row.map((cell, j) => masks.reduce((t, c) => fn(t, c[i][j]), cell)),
    );
  }

  /**
   * @param {SheetFieldType[]} fields
   * @returns {AugmentedFieldType[]}
   */
  export function createFields(fields: SheetFieldType[]): AugmentedFieldType[] {
    return fields.map((field, index) => {
      const toField = (str: string): string =>
        str
          .replace('#', 'No')
          .replace(/([/_])(.)/g, ' $2')
          .replace(/^\W|\W$/, '')
          .replace(/[^a-zA-Z0-9_\s]+/gi, '')
          .toLowerCase()
          .split(/\s+/)
          .filter((v) => v !== '')
          .map((v, i) => (i ? v.slice(0, 1).toUpperCase() + v.slice(1) : v))
          .join('');

      const toFIELD = (str: string): string =>
        str
          .toString()
          .replace('#', 'No')
          .replace(/([/_])(.)/g, ' $2')
          .replace(/^\W|\W$/, '')
          .replace(/[^a-zA-Z0-9_\s]+/gi, '')
          .toLowerCase()
          .split(/\s+/)
          .filter((v) => v !== '')
          .map((v) => v.toUpperCase())
          .join('_');

      const result: AugmentedFieldType = {
        i: index,
        c: index + 2,
        a: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[index],
        l: typeof field === 'string' ? field : field.l,
        f:
          typeof field === 'string'
            ? toField(field)
            : field.f || toField(field.l),
        F:
          typeof field === 'string'
            ? toFIELD(field)
            : field.F || toFIELD(field.l),
      };

      if (typeof field !== 'string' && field.cF) {
        result.cF = field.cF;
      }

      if (typeof field !== 'string' && field.aF) {
        result.aF = field.aF;
      }

      return result;
    });
  }

  /**
   * @param {SheetFieldType[]} fields
   * @returns {FieldMap}
   */
  export function createFieldMap(fields: SheetFieldType[]): FieldMap {
    const augmentedFields = this.createFields(fields);
    return augmentedFields.reduce((tot, augmentedField, index) => {
      tot[index] = augmentedField;
      tot[augmentedField.a] = augmentedField;
      tot[augmentedField.F] = augmentedField;
      tot[augmentedField.f] = augmentedField;

      return tot;
    }, {});
  }

  /** Display's toast in lower right corner of spreadsheet
   * @param {string} message the main message to display
   * @param {string} [title='GV App'] the title of the toast. Default is 'GV App'
   * @param {number} [timeoutSeconds=15] how long to display the toast. (default: 15 sec)
   */
  export function toastIt(
    message: string,
    title: string = 'GV App',
    timeoutSeconds: number = 15,
  ) {
    g.ss.toast(message, title, timeoutSeconds);
  }

  /** Display alert box in center of screen
   * @param {string} message the message to display
   * @param {string} [title='Oops!'] the title of the alert box
   * @return {GoogleAppsScript.Base.Button}
   */
  export function alertIt(
    message: string,
    title: string = 'Oops!',
  ): GoogleAppsScript.Base.Button {
    const ui = SpreadsheetApp.getUi();
    return ui.alert(title, message, ui.ButtonSet.OK);
  }

  /** Prompt user with a question to answer
   * @param {string} question the question to ask
   * @param {string} [title] the title of the prompt box
   * @return {string} the answer to the question
   */
  export function askQuestion(question: string, title: string = ''): string {
    const ui = SpreadsheetApp.getUi();
    const result = ui.prompt(title, question, ui.ButtonSet.OK_CANCEL);
    return result.getResponseText();
  }

  /** Displays a dialog box
   * @param {GoogleAppsScript.HTML.HtmlOutput} html HtmlOutput to display
   * @param {string} title title of dialog box
   * @param {number} [width=900] width of dialog box
   * @param {number} [height=600] height of dialog box
   */
  export function displayIt(
    html: GoogleAppsScript.HTML.HtmlOutput,
    title: string,
    width: number = 900,
    height: number = 600,
  ): void {
    const ui = SpreadsheetApp.getUi();
    ui.showModelessDialog(html.setWidth(width).setHeight(height), title);
  }

  /** Display an error message and finalizes process logger
   * @param {string} message
   * @param {string} title
   */
  export function displayError(message: string, title?: string) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title || 'Oops!', message, ui.ButtonSet.OK);
  }

  /** execute a function and time how long it takes
   * @param {function} func the thing to execute
   * @return {TimeItResponse} the result and stats
   */
  export function timeIt(func: Function, ...args: any[]): TimeItResponse {
    const startedAt = new Date().getTime();
    let result = null;
    let error = null;
    try {
      result = func(...args);
    } catch (err) {
      error = err;
    }
    const finishedAt = new Date().getTime();
    const elapsed = finishedAt - startedAt;
    const summary = {
      startedAt,
      finishedAt,
      result,
      error,
      elapsed,
    };

    console.log('summary :>> ', summary);
    return summary;
  }

  /** Picks cells or columns from array or matrix, according to provided indexes.
   * @param {any[]|any[][]} arr
   * @param {number[]} indexes
   * @return {any[]|any[][]}
   */
  export const choose = (arr: any[] | any[][], indexes: number[]) => {
    if (arr[0].constructor === Array) {
      return arr.map((row) => indexes.map((i) => row[i]));
    }
    const results = indexes.map((i) => arr[i]);
    return results;
  };

  /** Verifies that sheet is the correct sheet by name, or displays error
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
   * @param {string} sheetName
   */
  export function verifySheet(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    sheetName: string,
  ) {
    if (sheet.getName() !== sheetName)
      Utils.displayError(`Must have ${sheetName} sheet open`);
  }

  export const header2Field = (str: string): string =>
    str
      .replace('#', 'No')
      .replace(/([/_])(.)/g, ' $2')
      .replace(/^\W|\W$/, '')
      .replace(/[^a-zA-Z0-9_\s]+/gi, '')
      .toLowerCase()
      .split(/\s+/)
      .filter((v) => v !== '')
      .map((v, i) => (i ? v.slice(0, 1).toUpperCase() + v.slice(1) : v))
      .join('');

  export const header2FIELD = (str: string): string =>
    str
      .toString()
      .replace('#', 'No')
      .replace(/([/_])(.)/g, ' $2')
      .replace(/^\W|\W$/, '')
      .replace(/[^a-zA-Z0-9_\s]+/gi, '')
      .toLowerCase()
      .split(/\s+/)
      .filter((v) => v !== '')
      .map((v) => v.toUpperCase())
      .join('_');

  export const createFieldIndexMapFromSheet = (
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
  ): FIMType => {
    const headers = sheet.getDataRange().offset(0, 0, 1).getValues()[0];
    return headers
      .map((h) => Utils.header2Field(h))
      .reduce((tot, field, i) => ({ ...tot, [field + 'I']: i }), {});
  };

  export const createFieldColumnMap = (fields: AugmentedFieldType[]): FIMType =>
    fields.reduce((tot, cur) => ({ ...tot, [cur.f + 'C']: cur.c }), {});

  /**
   * Converts provided date into 'YYYY-MM-DD' format
   * @param {Date | number | string | Dayjs} d the date to convert to string format
   * @param {string} [inFormat] the format of the provided date when providing as string
   * @param {string} [outFormat='YYYY-MM-DD'] the outgoing string format
   * @return {string}
   */
  export function dFull(
    d?: Date | number | string | Dayjs,
    outFormat: string = 'YYYY-MM-DD',
    inFormat?: string,
  ): string {
    return daygs(d || undefined, inFormat).format(outFormat);
  }

  /**
   * Returns the given date's end-of-month
   * @date 3/4/2023 - 4:23:51 PM
   *
   * @param {?(Date | number | string | Dayjs)} [d]
   * @returns {Dayjs}
   */
  export const dEom = (d?: Date | number | string | Dayjs) =>
    daygs(d || undefined).endOf('month');

  /**
   * Returns either the 15th or end of month, the one upper bounding the given date
   * @date 3/4/2023 - 4:23:51 PM
   *
   * @param {?(Date | number | string | Dayjs)} [d]
   * @returns {Dayjs}
   */
  export const dEop = (d?: Date | number | string | Dayjs) =>
    daygs(d || undefined).date() <= 15
      ? daygs(d || undefined).set('date', 15)
      : daygs(d || undefined).endOf('month');
  export const round = (val: number, dec = 2) => Number(val.toFixed(dec));
  /**
   * Generates an ID using the date fields and current row index
   * @date 3/4/2023 - 4:22:39 PM
   *
   * @export
   * @param {(string | number | Date)} d
   * @param {number} rowNo
   * @returns {string}
   */
  export function getId(d: string | number | Date, rowNo: number): string {
    const dt: Dayjs = daygs(d);
    const yy = dt.year() - 2000;
    const mm = dt.month() + 1;
    const dd = dt.date();
    const dateBase = (
      yy.toString(36) +
      mm.toString(36) +
      dd.toString(36)
    ).toUpperCase();
    const rowBase = `00${rowNo.toString(36).toUpperCase()}`.slice(-3);
    return `${dateBase}${rowBase}`;
  }

  export function titleCaseAndRemoveSpaces(str?: string): string {
    if (!str) return '';

    // Convert the string to lowercase and split by spaces
    let words = str.replace('#', 'Number').toLowerCase().split(' ');

    // Capitalize the first letter of each word
    words = words.map((word) => word.charAt(0).toUpperCase() + word.slice(1));

    // Join the words together without spaces
    let result = words.join('');

    return result;
  }

  export function countItemsInMatrix(matrix: any[][]): number {
    return (
      matrix.length * matrix[0].length -
      matrix.flat().filter((el) => el != null).length
    );
  }

  export function groupText(str: string): string[] {
    let result = [];
    let stack = [];
    let current = '';

    for (let i = 0; i < str.length; i++) {
      let char = str[i];

      if (char === '(' || char === '[' || char === '{') {
        if (stack.length === 0) {
          current += char;
          result.push(current);
          current = '';
        } else {
          current += char;
        }
        stack.push(char);
      } else if (char === ')' || char === ']' || char === '}') {
        if (stack.length === 1 && current.length) {
          result.push(current);
          current = char;
        } else {
          current += char;
        }
        stack.pop();
      } else if (char === ',') {
        if (stack.length === 1) {
          result.push(current);
          result.push(char);
          current = '';
        } else {
          current += char;
        }
      } else {
        current += char;
      }
    }

    result.push(current);
    return result;
  }
}

export { Utils };
