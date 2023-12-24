import { default as dayjs } from 'dayjs';
import 'dayjs/plugin/duration';
import currency from 'currency.js';
import beautify from 'json-beautify';
import { SettingsOptions } from './addon';

declare var AppLib: {
  dayjs: typeof dayjs;
  currency: typeof currency;
  beautify: typeof beautify;
};

declare global {
  var g: {
    UserSettings?: SettingsOptions;
    ss?: GoogleAppsScript.Spreadsheet.Spreadsheet;
    ActiveSheet?: GoogleAppsScript.Spreadsheet.Sheet;
    ActiveRange?: GoogleAppsScript.Spreadsheet.Range;
    ActiveRowsRange?: GoogleAppsScript.Spreadsheet.Range;
    ActiveStartRow?: number;
    ABCs?: string[];
    LOG_SHEETNAME?: string;
  };
  var daygs: typeof dayjs;
  var beautifygs: typeof beautify;
  var currencygs: typeof currency;
}
