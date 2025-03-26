import { default as dayjs } from 'dayjs';
import 'dayjs/plugin/duration';
import { ImageCollectionType, SettingsOptions, ThemeType } from './addon';

// declare var AppLib: {
//   dayjs: typeof dayjs;
// };
declare var Dayjs: { dayjs: typeof dayjs };

declare global {
  var g: {
    UserSettings?: SettingsOptions;
    ss?: GoogleAppsScript.Spreadsheet.Spreadsheet;
    ActiveSheet?: GoogleAppsScript.Spreadsheet.Sheet;
    ActiveRange?: GoogleAppsScript.Spreadsheet.Range;
    ActiveRowsRange?: GoogleAppsScript.Spreadsheet.Range;
    ActiveStartRow?: number;
    ABCs?: string[];
    Theme?: ThemeType;
    LOG_SHEETNAME?: string;
    SheetIcons?: ImageCollectionType;
  };
  var daygs: typeof dayjs;
}
