import { default as dayjs } from 'dayjs';
import 'dayjs/plugin/duration';
import { ImageCollectionType, SettingsOptions } from './addon';

declare var AppLib: {
  dayjs: typeof dayjs;
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
    SheetIcons?: ImageCollectionType;
  };
  var daygs: typeof dayjs;
}
