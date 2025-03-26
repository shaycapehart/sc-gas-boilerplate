import { Settings } from '@core/environment/Settings';
import { ImageCollectionType, SettingsOptions } from '@core/types/addon';

globalThis.g = {};
globalThis.daygs = Dayjs.dayjs;

function getRgb_(themeColorType) {
  return SpreadsheetApp.getActive()
    .getSpreadsheetTheme()
    .getConcreteColor(themeColorType)
    .asRgbColor()
    .asHexString();
}

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
  [
    'Theme',
    () => ({
      fontFamily: SpreadsheetApp.getActive()
        .getSpreadsheetTheme()
        .getFontFamily(),
      textColor: getRgb_(SpreadsheetApp.ThemeColorType.TEXT),
      chartBackground: getRgb_(SpreadsheetApp.ThemeColorType.BACKGROUND),
      accent1: getRgb_(SpreadsheetApp.ThemeColorType.ACCENT1),
      accent2: getRgb_(SpreadsheetApp.ThemeColorType.ACCENT2),
      accent3: getRgb_(SpreadsheetApp.ThemeColorType.ACCENT3),
      accent4: getRgb_(SpreadsheetApp.ThemeColorType.ACCENT4),
      accent5: getRgb_(SpreadsheetApp.ThemeColorType.ACCENT5),
      accent6: getRgb_(SpreadsheetApp.ThemeColorType.ACCENT6),
      hyperlinkColor: getRgb_(SpreadsheetApp.ThemeColorType.HYPERLINK),
    }),
  ],
  [
    'SheetIcons',
    (): ImageCollectionType => {
      var folder = DriveApp.getFolderById(Settings.SPREADSHEET_ICON_FOLDER_ID);
      var filecontents = folder.getFiles();

      var icons = {};
      var file: GoogleAppsScript.Drive.File;

      while (filecontents.hasNext()) {
        var file = filecontents.next();
        const name = file.getName().replace(/\..*/, '');
        icons[name] =
          '=image("https://docs.google.com/uc?id=' + file.getId() + '")';
      }
      return icons;
    },
  ],
];

myGlobalConfig.forEach(([propName, value]) => {
  return addGetter_(propName, value);
});
