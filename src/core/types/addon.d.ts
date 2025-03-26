export interface OnEditEvent {
  authMode: GoogleAppsScript.Script.AuthMode;
  value: string;
  oldValue: string;
  range: GoogleAppsScript.Spreadsheet.Range;
  source: GoogleAppsScript.Spreadsheet.Spreadsheet;
  triggerUid: string;
  user: GoogleAppsScript.Base.User;
}

export type AddonResponse =
  | GoogleAppsScript.Card_Service.Card
  | GoogleAppsScript.Card_Service.Card[]
  | GoogleAppsScript.Card_Service.ActionResponse
  | GoogleAppsScript.Card_Service.UniversalActionResponse;

export type ErrorHandler = (exception: Error) => AddonResponse;

export interface ErrorCardOptions {
  exception?: Error;
  errorText?: string;
  showStackTrace?: boolean;
}

export interface RGBType {
  r: number;
  g: number;
  b: number;
}

export interface HSLType {
  h: number;
  s: number;
  l: number;
}

export interface SettingsOptions {
  hasTitle?: boolean;
  hasHeaders?: boolean;
  hasFooter?: boolean;
  leaveTop?: boolean;
  leaveLeft?: boolean;
  leaveBottom?: boolean;
  noBottom?: boolean;
  centerAll?: boolean;
  alternating?: boolean;
  backgroundTitle: number;
  backgroundHeaders: number;
  backgroundDataFirst: number;
  backgroundDataSecond: number;
  backgroundFooter: number;
  bordersAll: number;
  bordersHorizontal: number;
  bordersThickness: number;
  bordersVertical: number;
  bordersTitleBottom: number;
  bordersHeadersBottom: number;
  bordersHeadersVertical: number;
  debugControl: string;
  helpControl: string;
  themeControl: string;
  printEventObject?: boolean;
  colorPicker?: string;
  previousEvent?: GoogleAppsScript.Addons.EventObject;
  customThemes: {
    theme1: ThemeType | null;
    theme2: ThemeType | null;
    theme3: ThemeType | null;
    theme4: ThemeType | null;
    theme5: ThemeType | null;
    theme6: ThemeType | null;
  };
}

export interface TableFormatOptionsType {
  color?: string;
  hasTitle: boolean;
  hasHeaders: boolean;
  hasFooter?: boolean;
  leaveTop?: boolean;
  leaveLeft?: boolean;
  leaveBottom?: boolean;
  noBottom?: boolean;
  centerAll: boolean;
  alternating: boolean;
}

export interface ThemeType {
  title?: string;
  description?: string;
  fontFamily: string;
  chartBackground: string;
  textColor: string;
  hyperlinkColor: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
}

export interface SortType {
  column: number;
  ascending: boolean;
}

export type MatrixType = any[][];
export type MatrixArrayType = MatrixType[];

export interface AugmentedFieldType {
  l: string;
  f?: string;
  F?: string;
  cF?: string;
  aF?: string;
  i?: number;
  c?: number;
  a?: string;
}

export type SheetFieldType = string | AugmentedFieldType;

export type FieldMap = {
  [key: string]: AugmentedFieldType;
};

export interface FIMType {
  [key: string]: number;
}

export interface TimeItResponse {
  startedAt: number;
  finishedAt: number;
  result: any;
  error: string;
  elapsed: number;
}

export type ImageCollectionType = {
  [key: string]: string;
};
