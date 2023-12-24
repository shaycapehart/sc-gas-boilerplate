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
