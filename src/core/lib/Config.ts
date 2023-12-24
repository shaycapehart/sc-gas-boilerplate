import { Settings } from '@core/environment/Settings';
import { SettingsOptions } from '../types/addon';
import { AppLib } from '../types/globals';

// console.log('Reading CONFIG')
globalThis.g = {};
globalThis.daygs = AppLib.dayjs;
globalThis.beautifygs = AppLib.beautify;
globalThis.currencygs = AppLib.currency;

const addGetter_ = <T>(propName: string, value: () => T, target = g) => {
  Object.defineProperty(target, propName, {
    enumerable: true,
    configurable: true,
    get() {
      delete this[propName];
      return (this[propName] = value());
    },
  });
  return target;
};

//MY GLOBAL VARIABLES in g
var myGlobalConfig: [string, () => unknown][] = [
  ['UserSettings', (): SettingsOptions => Settings.getSettingsForUser()],
  [
    'ss',
    (): GoogleAppsScript.Spreadsheet.Spreadsheet => SpreadsheetApp.getActive(),
  ],
  [
    'ActiveSheet',
    (): GoogleAppsScript.Spreadsheet.Sheet => g.ss.getActiveSheet(),
  ],
  [
    'ActiveRange',
    (): GoogleAppsScript.Spreadsheet.Range => g.ActiveSheet.getActiveRange(),
  ],
  [
    'ActiveRowsRange',
    (): GoogleAppsScript.Spreadsheet.Range =>
      g.ActiveSheet.getRange(
        g.ActiveRange.getRow(),
        1,
        g.ActiveRange.getHeight(),
        g.ActiveSheet.getLastColumn(),
      ),
  ],
  ['ActiveStartRow', (): number => g.ActiveRange.getRow()],
  [
    'ABCs',
    () =>
      ['', 'A', 'B', 'C', 'D']
        .map((f) => 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('').map((s) => f + s))
        .flat(),
  ],
];

myGlobalConfig.forEach(([propName, value]) => {
  // console.log('Creating getter for:', propName)
  return addGetter_(propName, value);
});
