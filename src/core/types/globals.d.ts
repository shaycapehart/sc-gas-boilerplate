import { ImageCollectionType, SettingsOptions, ThemeType } from './addon';

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
}
