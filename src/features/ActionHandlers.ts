import { Settings } from '@core/environment/Settings';
import { MP, COLORS } from '@core/lib/COLORS';
import { SortType, TableFormatOptionsType } from '@core/types/addon';
import { Utils } from '@core/utils/Utils';
import { Views } from './Views';
import { Shlog } from '@core/logging/Shlog';

/**
 * Collection of functions to handle user interactions with the add-on.
 *
 * @namespace
 */
namespace ActionHandlers {
  /* -------------------------------------------------- Empty Logs -------------------------------------------------- */
  export function emptyLogRecords() {
    const newLogsCount = Shlog.outputTictocs();

    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          newLogsCount
            ? `Successfully added ${newLogsCount} log rows`
            : 'No logs to add at this time',
        ),
      )
      .build();
  }

  /* ------------------------------------------------- Format Tables ------------------------------------------------ */
  export function tableFormat(
    rng: GoogleAppsScript.Spreadsheet.Range,
    options: TableFormatOptionsType,
  ) {
    const {
      color,
      hasTitle,
      hasHeaders,
      leaveTop,
      leaveLeft,
      leaveBottom,
      noBottom,
      centerAll,
      alternating,
      hasFooter,
    } = options;

    const OFF_WHITE = '#FCFCFC';
    const WHITE = '#FFFFFF';
    const CENTER = 'center';
    const BOLD = 'bold';

    const { BASE, LIGHT, LIGHTER, LIGHTEST } = COLORS[color];

    // define the color shades
    const PALLETE = {
      BACKGROUND_TITLE: BASE,
      BACKGROUND_HEADERS: LIGHT,
      BACKGROUND_DATA_FIRST: LIGHTEST,
      BACKGROUND_DATA_SECOND: LIGHTER,
      BACKGROUND_FOOTER: LIGHT,
      BORDERS_ALL: BASE,
      BORDERS_HORIZONTAL: BASE,
      BORDERS_VERTICAL: BASE,
      BORDERS_TITLE_BOTTOM: null,
      BORDERS_HEADERS_BOTTOM: null,
      BORDERS_HEADERS_VERTICAL: null,
    };

    const { SOLID, SOLID_MEDIUM, DASHED, SOLID_THICK } =
      SpreadsheetApp.BorderStyle;

    const tableRange = (rng || g.ActiveRange).offset(
      +hasTitle,
      0,
      (rng || g.ActiveRange).getHeight() - +hasTitle,
    );
    const dataRange = (rng || g.ActiveRange).offset(
      +hasTitle + +hasHeaders,
      0,
      (rng || g.ActiveRange).getHeight() - +hasTitle - +hasHeaders,
    );
    const headersRange = hasHeaders ? tableRange.offset(0, 0, 1) : null;
    const titleRange = hasTitle ? tableRange.offset(-1, 0, 1) : null;

    const fontSizes = (rng || g.ActiveRange).getFontSizes();

    (rng || g.ActiveRange).setBorder(
      leaveTop ? null : true,
      leaveLeft ? null : true,
      leaveBottom ? null : noBottom ? false : true,
      true,
      null,
      null,
      PALLETE.BORDERS_ALL,
      hasTitle ? SOLID_MEDIUM : SOLID,
    );

    if (alternating) {
      const banding =
        tableRange.getBandings()[0] ||
        tableRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
      banding
        .setHeaderRowColor(hasHeaders ? PALLETE.BACKGROUND_HEADERS : null)
        .setFirstRowColor(PALLETE.BACKGROUND_DATA_FIRST)
        .setSecondRowColor(PALLETE.BACKGROUND_DATA_SECOND)
        .setFooterRowColor(hasFooter ? PALLETE.BACKGROUND_FOOTER : null);
    } else {
      if (headersRange) headersRange.setBackground(PALLETE.BACKGROUND_HEADERS);

      dataRange
        .setBorder(
          null,
          null,
          null,
          null,
          null,
          true,
          PALLETE.BORDERS_HORIZONTAL,
          DASHED,
        )
        .setBackground(PALLETE.BACKGROUND_DATA_FIRST);
    }

    if (titleRange)
      titleRange
        .merge()
        .setHorizontalAlignment(CENTER)
        .setFontWeight(BOLD)
        .setBackground(PALLETE.BACKGROUND_TITLE);

    if (headersRange) headersRange.setHorizontalAlignment(CENTER);

    if (centerAll) dataRange.setHorizontalAlignment(CENTER);

    (rng || g.ActiveRange).setFontSizes(fontSizes);

    dataRange.setBorder(
      null,
      null,
      null,
      null,
      true,
      null,
      PALLETE.BORDERS_VERTICAL,
      SOLID,
    );
  }
  export function setTableFormatDefaults(
    e: GoogleAppsScript.Addons.EventObject,
  ) {
    const tableOptions: string[] =
      e.commonEventObject.formInputs?.tableOptions?.stringInputs?.value || [];
    var settings = {
      hasTitle: tableOptions.includes('hasTitle'),
      hasHeaders: tableOptions.includes('hasHeaders'),
      hasFooter: tableOptions.includes('hasFooter'),
      leaveTop: tableOptions.includes('leaveTop'),
      leaveLeft: tableOptions.includes('leaveLeft'),
      leaveBottom: tableOptions.includes('leaveBottom'),
      noBottom: tableOptions.includes('noBottom'),
      centerAll: tableOptions.includes('centerAll'),
      alternating: tableOptions.includes('alternating'),
    };

    Settings.updateSettingsForUser(settings);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().popCard())
      .setNotification(
        CardService.newNotification().setText('Defaults updated'),
      )
      .build();
  }
  export function formatRangeAsTable(
    rng: GoogleAppsScript.Spreadsheet.Range,
    options: TableFormatOptionsType,
  ) {
    const {
      color,
      hasTitle,
      hasHeaders,
      leaveTop,
      leaveLeft,
      leaveBottom,
      noBottom,
      centerAll,
      alternating,
      hasFooter,
    } = options;
    const title = Number(hasTitle);
    const headers = Number(hasHeaders);
    const OFF_WHITE = '#FCFCFC';
    const CENTER = 'center';
    const BOLD = 'bold';

    const settings = Settings.getSettingsForUser();

    // define the color shades
    const PALLETE = {
      BACKGROUND_TITLE: MP[color][settings.backgroundTitle],
      BACKGROUND_HEADERS: MP[color][settings.backgroundHeaders],
      BACKGROUND_DATA_FIRST: MP[color][settings.backgroundDataFirst],
      BACKGROUND_DATA_SECOND: MP[color][settings.backgroundDataSecond],
      BACKGROUND_FOOTER: MP[color][settings.backgroundFooter],
      BORDERS_ALL: MP[color][settings.bordersAll],
      BORDERS_HORIZONTAL: MP[color][settings.bordersHorizontal],
      BORDERS_VERTICAL: MP[color][settings.bordersVertical],
      BORDERS_TITLE_BOTTOM: MP[color][settings.bordersTitleBottom],
      BORDERS_HEADERS_BOTTOM: MP[color][settings.bordersHeadersBottom],
      BORDERS_HEADERS_VERTICAL: MP[color][settings.bordersHeadersVertical],
    };

    const { SOLID, SOLID_MEDIUM, DASHED, SOLID_THICK } =
      SpreadsheetApp.BorderStyle;

    const tableRange = rng.offset(title, 0, rng.getHeight() - title);
    const dataRange = rng.offset(
      title + headers,
      0,
      rng.getHeight() - title - headers,
    );
    const headersRange = hasHeaders ? tableRange.offset(0, 0, 1) : null;
    const titleRange = hasTitle ? tableRange.offset(-1, 0, 1) : null;

    const fontSizes = rng.getFontSizes();

    rng.setBorder(
      leaveTop ? null : true,
      leaveLeft ? null : true,
      leaveBottom ? null : noBottom ? false : true,
      true,
      null,
      null,
      PALLETE.BORDERS_ALL,
      settings.bordersThickness === 1
        ? SOLID
        : settings.bordersThickness === 2
        ? SOLID_MEDIUM
        : SOLID_THICK,
    );

    if (alternating) {
      const banding =
        tableRange.getBandings()[0] ||
        tableRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
      banding
        .setHeaderRowColor(hasHeaders ? PALLETE.BACKGROUND_HEADERS : null)
        .setFirstRowColor(PALLETE.BACKGROUND_DATA_FIRST)
        .setSecondRowColor(PALLETE.BACKGROUND_DATA_SECOND)
        .setFooterRowColor(hasFooter ? PALLETE.BACKGROUND_FOOTER : null);
    } else {
      if (headersRange) headersRange.setBackground(PALLETE.BACKGROUND_HEADERS);

      dataRange
        .setBorder(
          null,
          null,
          null,
          null,
          null,
          true,
          PALLETE.BORDERS_HORIZONTAL,
          DASHED,
        )
        .setBackground(PALLETE.BACKGROUND_DATA_FIRST);
    }

    if (titleRange)
      titleRange
        .merge()
        .setHorizontalAlignment(CENTER)
        .setFontColor(OFF_WHITE)
        .setBackground(PALLETE.BACKGROUND_TITLE)
        .setBorder(
          null,
          null,
          true,
          null,
          null,
          null,
          PALLETE.BORDERS_TITLE_BOTTOM,
          hasTitle ? SOLID_MEDIUM : SOLID,
        );

    if (headersRange)
      headersRange
        .setFontWeight(BOLD)
        .setHorizontalAlignment(CENTER)
        .setBorder(
          null,
          null,
          null,
          null,
          true,
          null,
          PALLETE.BORDERS_HEADERS_VERTICAL,
          SOLID,
        )
        .setBorder(
          null,
          null,
          true,
          null,
          null,
          null,
          PALLETE.BORDERS_HEADERS_BOTTOM,
          SOLID,
        );

    if (centerAll) dataRange.setHorizontalAlignment(CENTER);

    // rng.offset(title + headers, 0, rng.getHeight() - title - headers).setHorizontalAlignment('center')

    rng.setFontSizes(fontSizes);

    // .offset(title + headers, 0, rng.getHeight() - title - headers)

    dataRange.setBorder(
      null,
      null,
      null,
      null,
      true,
      null,
      PALLETE.BORDERS_VERTICAL,
      SOLID,
    );
  }

  export function squareSelectedCells(e: GoogleAppsScript.Addons.EventObject) {
    console.log('e :>> ', e);
    console.log(
      'e.commonEventObject.formInputs :>> ',
      e.commonEventObject.formInputs.pixels.stringInputs,
    );
    const pixels = parseInt(
      e.commonEventObject.formInputs.pixels.stringInputs.value[0],
    );
    const nRows = g.ActiveRange.getHeight();
    const nCols = g.ActiveRange.getWidth();
    const startRow = g.ActiveRange.getRow();
    const startCol = g.ActiveRange.getColumn();
    g.ActiveSheet.setRowHeightsForced(startRow, nRows, pixels).setColumnWidths(
      startCol,
      nCols,
      pixels,
    );
  }

  export function formatTable(e: GoogleAppsScript.Addons.EventObject) {
    const tableOptions: string[] =
      e.commonEventObject.formInputs?.tableOptions?.stringInputs?.value || [];

    tableFormat(g.ActiveRange, {
      color: e.commonEventObject.parameters.color,
      hasTitle: tableOptions.includes('hasTitle'),
      hasHeaders: tableOptions.includes('hasHeaders'),
      hasFooter: tableOptions.includes('hasFooter'),
      leaveTop: tableOptions.includes('leaveTop'),
      leaveLeft: tableOptions.includes('leaveLeft'),
      leaveBottom: tableOptions.includes('leaveBottom'),
      noBottom: tableOptions.includes('noBottom'),
      centerAll: tableOptions.includes('centerAll'),
      alternating: tableOptions.includes('alternating'),
    });
  }

  /**
   * Crops the current sheet to the user's selection.
   */
  export function cropToSelection() {
    console.log('Cropping to selection');
    var range = g.ActiveRange;
    cropToRange_(range);
  }
  /* ------------------------------------------------- Crop to Data ------------------------------------------------- */
  /**
   * Crops the current sheet to the data.
   */
  export function cropToData() {
    console.log('Cropping to data');
    var range = g.ActiveSheet.getDataRange();
    cropToRange_(range);
  }
  /* ------------------------------------------------- Crop to Range ------------------------------------------------ */
  /**
   * Crops the sheet such that it only contains the given range.
   * @param {SpreadsheetApp.Range} range The range to crop to.
   */
  export function cropToRange_(range) {
    var sheet = range.getSheet();
    var spreadsheet = sheet.getParent();
    var firstRow = range.getRow();
    var lastRow = firstRow + range.getNumRows() - 1;
    var firstColumn = range.getColumn();
    var lastColumn = firstColumn + range.getNumColumns() - 1;
    var maxRows = sheet.getMaxRows();
    var maxColumns = sheet.getMaxColumns();

    if (lastRow < maxRows) {
      sheet.deleteRows(lastRow + 1, maxRows - lastRow);
    }
    if (firstRow > 1) {
      sheet.deleteRows(1, firstRow - 1);
    }
    if (lastColumn < maxColumns) {
      sheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn);
    }
    if (firstColumn > 1) {
      sheet.deleteColumns(1, firstColumn - 1);
    }
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  }

  function getRangeType_(range: GoogleAppsScript.Spreadsheet.Range): string {
    const typeRegex = /([a-zA-Z]*)(\d*)(:?)([a-zA-Z]*)(\d*)/;
    const a1Notation = range.getA1Notation();
    const matches = typeRegex.exec(a1Notation);
    const [all, a1, n1, sep, a2, n2] = matches;
    const type = !sep
      ? 'cell'
      : !n1 && !n2
      ? a1 === a2
        ? 'column'
        : 'table'
      : !n1 || !n2
      ? 'range'
      : n1 === n2
      ? 'row'
      : 'range';
    return type;
  }
  /* ------------------------------------------------- Import Range ------------------------------------------------- */
  /**
   * Imports range data from one Google Sheet to another.
   * @param {string} sourceID - The id of the source Google Sheet.
   * @param {string} sourceRange - The Sheet tab and range to copy.
   * @param {string} destinationID - The id of the destination Google Sheet.
   * @param {string} destinationRangeStart - The destintation location start cell as a sheet name and cell.
   */
  export function importRange(
    sourceID: string,
    sourceRange: string,
    destinationID: string,
    destinationRangeStart: string,
  ) {
    // Gather the source range values
    const sourceSS = SpreadsheetApp.openById(sourceID);
    const sourceRng = sourceSS.getRange(sourceRange);
    const sourceVals = sourceRng.getValues();
    const sourceType = getRangeType_(sourceRng);

    // Get the destiation sheet and cell location.
    const destinationSS = SpreadsheetApp.openById(destinationID);
    const destStartRange = destinationSS.getRange(destinationRangeStart);
    const destSheet = destStartRange.getSheet();
    const destType = getRangeType_(destStartRange);

    // Get the full data range to paste from start range.
    const destRange = destSheet.getRange(
      destStartRange.getRow(),
      destStartRange.getColumn(),
      sourceVals.length,
      sourceVals[0].length,
    );

    destRange.setValues(sourceVals);

    if (
      (/table|column/.test(sourceType) || /table|row|column/.test(destType)) &&
      destStartRange.getRow() + destRange.getHeight() - 1 <
        destSheet.getMaxRows()
    )
      destSheet.deleteRows(
        destStartRange.getRow() + destRange.getHeight(),
        destSheet.getMaxRows() -
          destStartRange.getRow() -
          destRange.getHeight() +
          1,
      );

    SpreadsheetApp.flush();
  }

  export function sortRange(
    ssid: string,
    range: string,
    sortParams: SortType[],
  ) {
    if (sortParams.length % 2 !== 0) return;
    const ss = SpreadsheetApp.openById(ssid);
    const rng = ss.getRange(range);
    rng.sort(sortParams);
  }

  export function getNamedFunctions() {
    const url = `https://docs.google.com/spreadsheets/export?exportFormat=xlsx&id=${g.ss.getId()}`;
    const resHttp = UrlFetchApp.fetch(url, {
      headers: { authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    });

    // Retrieve the data from XLSX data.
    const blobs = Utilities.unzip(
      resHttp.getBlob().setContentType('application/zip'),
    );
    const workbook = blobs.find((b) => b.getName() == 'xl/workbook.xml');
    if (!workbook) {
      throw new Error('No file.');
    }

    // Parse XLSX data and retrieve the named functions.
    const root = XmlService.parse(workbook.getDataAsString()).getRootElement();
    const definedNames = root
      .getChild('definedNames', root.getNamespace())
      .getChildren();
    const res = definedNames.map((e) => ({
      definedName: e.getAttribute('name').getValue(),
      definedFunction: e.getValue(),
    }));
    let sheet = g.ss.getSheetByName('__named_functions__');
    if (!sheet) {
      sheet = g.ss.insertSheet('__named_functions__');
    }
    sheet.getRange('A:C').clearContent();
    sheet
      .getRange('C1')
      .setFormula(
        `=BYROW(B:B,LAMBDA(fn,IF(ISBLANK(fn),,IF(ROW(fn)=ROW(B:B),HSTACK("Function",INDEX("Argument"&SEQUENCE(1,8))),LET(str,SUBSTITUTE(fn,CHAR(10),""),IFERROR(HSTACK(REGEXEXTRACT(str,"(?i)^LAMBDA\\((?:[a-z0-9_,\\s]+)(?:,\\s?)(.+)\\)$"),SPLIT(REGEXEXTRACT(str,"(?i)^LAMBDA\\(([a-z0-9_,\\s]+)(?:,\\s?)(?:.+)\\)$"),", ",0,0)),IFERROR(REGEXEXTRACT(str,"(?i)LAMBDA\\((.*)\\)$"))))))))`,
      );
    sheet
      .getRange(1, 1, res.length + 1, 2)
      .setValues([
        ['Defined Name', 'Defined Function'],
        ...res
          .sort(
            (a, b) =>
              (a.definedFunction.startsWith('LAMBDA') &&
              !b.definedFunction.startsWith('LAMBDA')
                ? -1
                : !a.definedFunction.startsWith('LAMBDA') &&
                  b.definedFunction.startsWith('LAMBDA')
                ? 1
                : 0) ||
              a.definedName.localeCompare(b.definedName) ||
              a.definedFunction.localeCompare(b.definedFunction),
          )
          .map(({ definedName, definedFunction }) => [
            definedName,
            definedFunction,
          ]),
      ]);

    // DriveApp.getFiles(); // This comment line is used for automatically detecting the scope of Drive API.
  }

  export function createNamedRangesDashboard() {
    let namedRanges = g.ss.getNamedRanges();
    let sheet: GoogleAppsScript.Spreadsheet.Sheet,
      range: GoogleAppsScript.Spreadsheet.Range;

    sheet = g.ss.getSheetByName('__named_ranges__');

    if (!sheet) {
      sheet = g.ss.insertSheet('__named_ranges__');

      const headers = [
        'NAMED RANGE',
        'RANGE',
        'SHEET',
        'FULL RANGE',
        'TYPE',
        'DELETE',
        'EDIT',
        'NEW NAME',
        'NEW RANGE',
        'NEW SHEET',
      ];

      const notes = [
        `Optional: 
        Name must start with a letter and may only contain letters, numbers, underscore, or period. Cannot look like a1notation (i.e. up to three letters followed by numbers. A1 - ZZZ1000)`,
        `Optional:
        Range needs to be in A1 Notation. Open-ended ranges are acceptable (i.e. A:A).`,
        `Optional`,
      ];

      sheet
        .setHiddenGridlines(true)
        .setColumnWidths(1, 4, 200)
        .setColumnWidths(8, 3, 200);

      sheet
        .getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .offset(0, 7, 1, 3)
        .setNotes([notes]);

      if (sheet.getMaxColumns() > 10)
        sheet.deleteColumns(11, sheet.getMaxColumns() - 10);
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(1);
    } else {
      if (sheet.getLastRow() > 1)
        sheet
          .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
          .clearContent();
    }

    const ssId = g.ss.getId();
    const baseUrl = `https://docs.google.com/spreadsheets/d/${ssId}/edit#gid=`;
    const richTextValues = [];

    const rows = [];
    namedRanges.forEach((nr) => {
      const name = nr.getName();
      const range = nr.getRange();
      const sheetName = range.getSheet().getName();
      const sheetId = range.getSheet().getSheetId();
      const a1Notation = range.getA1Notation();
      const url = `${baseUrl}${sheetId}&range=${a1Notation}`;
      richTextValues.push([
        SpreadsheetApp.newRichTextValue().setText(name).setLinkUrl(url).build(),
      ]);
      const typeRegex = /([a-zA-Z]*)(\d*)(:?)([a-zA-Z]*)(\d*)/;
      const matches = typeRegex.exec(a1Notation);
      const [all, a1, n1, sep, a2, n2] = matches;
      const type = !sep
        ? 'cell'
        : !n1 && !n2
        ? a1 === a2
          ? 'column'
          : 'table'
        : !n1 || !n2
        ? 'range'
        : n1 === n2
        ? 'row'
        : 'range';
      rows.push([a1Notation, sheetName, `'${sheetName}'!${a1Notation}`, type]);
    });

    if (sheet.getMaxRows() > rows.length + 1)
      sheet.deleteRows(rows.length + 2, sheet.getMaxRows() - (rows.length + 1));
    sheet.getRange(`A2:A${rows.length + 1}`).setRichTextValues(richTextValues);
    sheet.getRange(`B2:E${rows.length + 1}`).setValues(rows);

    sheet.getRange(`A2:E${rows.length + 1}`).sort([
      { column: 3, ascending: true },
      { column: 5, ascending: false },
      { column: 2, ascending: true },
    ]);

    sheet
      .getRange(`F2:G${rows.length + 1}`)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .setAllowInvalid(true)
          .requireCheckbox()
          .build(),
      );

    tableFormat(sheet.getDataRange(), {
      color: 'BLUE',
      hasTitle: false,
      hasHeaders: true,
      centerAll: false,
      alternating: false,
    });

    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(`Built fresh Named Ranges sheet`),
      )
      .build();
  }

  export function updateNamedRangeSheet() {
    let namedRanges = g.ss.getNamedRanges();
    let rows: any[][];
    let nrS = g.ss.getSheetByName('__named_ranges__');
    if (!nrS) return createNamedRangesDashboard();

    rows = nrS.getDataRange().getValues().slice(1);
    rows
      .filter((row) => !!row[5]) // Filter rows to delete
      .forEach((row) => {
        try {
          const nr = namedRanges.find(
            (namedRange) => namedRange.getName() === row[0],
          );
          nr.remove();
        } catch (error) {
          Utils.displayError(
            `Problem deleting ${row[0]} named range: ${error}`,
          );
        }
      });
    rows
      .filter((row) => !!row[6]) // Filter rows to update
      .forEach((row) => {
        const [
          nrName,
          nrA1,
          nrSheetName,
          nrFull,
          nrType,
          nrDelete,
          nrUpdate,
          newName,
          newRangeA1,
          newSheetName,
        ] = row;
        console.log('Edit row :>> ', row);

        try {
          const existingNr = namedRanges.find(
            (namedRange) => namedRange.getName() === nrName,
          );
          const namedRangeRangeA1 = `'${newSheetName || nrSheetName}'!${
            newRangeA1 || nrA1
          }`;
          const namedRangeName = `${newName || nrName}`;
          existingNr.setName(namedRangeName);
          existingNr.setRange(g.ss.getRange(namedRangeRangeA1));

          g.ss
            .getNamedRanges()
            .find((nr) => nr.getName() === namedRangeName)
            .setRange(g.ss.getRange(namedRangeRangeA1));
        } catch (error) {
          Utils.displayError(
            `Problem modifying ${newName || nrName} named range.`,
          );
        }
      });

    return createNamedRangesDashboard();
  }

  export function createNamedFunctionsDashboard(): GoogleAppsScript.Spreadsheet.Sheet {
    const headers = [
      [
        'Defined Name',
        'Defined Function',
        '=BYROW(B:B,LAMBDA(fn,IF(ISBLANK(fn),,IF(ROW(fn)=ROW(B:B),HSTACK("Function",INDEX("Argument"&SEQUENCE(1,9))),LET(str,SUBSTITUTE(fn,CHAR(10),""),IFERROR(HSTACK(REGEXEXTRACT(str,"(?i)^LAMBDA\\((?:[a-z0-9_,\\s]+)(?:,\\s?)(.+)\\)$"),SPLIT(REGEXEXTRACT(str,"(?i)^LAMBDA\\(([a-z0-9_,\\s]+)(?:,\\s?)(?:.+)\\)$"),", ",0,0)),IFERROR(REGEXEXTRACT(str,"(?i)LAMBDA\\((.*)\\)$"))))))))',
      ],
    ];

    let SHEET: GoogleAppsScript.Spreadsheet.Sheet,
      range: GoogleAppsScript.Spreadsheet.Range;

    SHEET =
      SpreadsheetApp.getActive().getSheetByName('__named_functions__') ||
      SpreadsheetApp.getActive().insertSheet('__named_functions__');

    range = SHEET.clear()
      .setHiddenGridlines(true)
      .setColumnWidth(1, 250)
      .setColumnWidth(2, 600)
      .setColumnWidth(3, 600)
      .setColumnWidths(4, 9, 150)
      .getRange(1, 1, SHEET.getMaxRows(), 12)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    range.offset(0, 0, 1, 3).setValues(headers);

    SHEET.deleteColumns(13, SHEET.getMaxColumns() - 12);
    SHEET.setFrozenRows(1);
    SHEET.setFrozenColumns(1);

    return SHEET;
  }
  /* ---------------------------------------- Show Named Functions Dashboard ---------------------------------------- */
  export function showNamedFunctionsDashboard() {
    const url = `https://docs.google.com/spreadsheets/export?exportFormat=xlsx&id=${SpreadsheetApp.getActive().getId()}`;
    const resHttp = UrlFetchApp.fetch(url, {
      headers: { authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    });

    // Retrieve the data from XLSX data.
    const blobs = Utilities.unzip(
      resHttp.getBlob().setContentType('application/zip'),
    );
    const workbook = blobs.find((b) => b.getName() == 'xl/workbook.xml');
    if (!workbook) {
      throw new Error('No file.');
    }

    // Parse XLSX data and retrieve the named functions.
    const root = XmlService.parse(workbook.getDataAsString()).getRootElement();
    const definedNames = root
      .getChild('definedNames', root.getNamespace())
      .getChildren();
    const res = definedNames.reduce((tot, cur) => {
      if (cur.getValue().startsWith('LAMBDA'))
        tot.push({
          definedName: cur.getAttribute('name').getValue(),
          definedFunction: cur.getValue(),
        });
      return tot;
    }, []);

    const headers = [
      [
        'Defined Name',
        'Defined Function',
        '=BYROW(B:B,LAMBDA(fn,IF(ISBLANK(fn),,IF(ROW(fn)=ROW(B:B),HSTACK("Function",INDEX("Argument"&SEQUENCE(1,9))),LET(str,SUBSTITUTE(fn,CHAR(10),""),IFERROR(HSTACK(REGEXEXTRACT(str,"(?i)^LAMBDA\\((?:[a-z0-9_,\\s]+)(?:,\\s?)(.+)\\)$"),SPLIT(REGEXEXTRACT(str,"(?i)^LAMBDA\\(([a-z0-9_,\\s]+)(?:,\\s?)(?:.+)\\)$"),", ",0,0)),IFERROR(REGEXEXTRACT(str,"(?i)LAMBDA\\((.*)\\)$"))))))))',
      ],
    ];

    let sheet: GoogleAppsScript.Spreadsheet.Sheet,
      range: GoogleAppsScript.Spreadsheet.Range;

    sheet = (
      g.ss.getSheetByName('__named_functions__') ||
      g.ss.insertSheet('__named_functions__')
    ).clear();

    range = sheet
      .clear()
      .setHiddenGridlines(true)
      .setColumnWidth(1, 250)
      .setColumnWidth(2, 600)
      .setColumnWidth(3, 600)
      .setColumnWidths(4, 9, 150)
      .getRange(1, 1, sheet.getMaxRows(), 12)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    range.offset(0, 0, 1, 3).setValues(headers);

    if (sheet.getMaxColumns() > 12)
      sheet.deleteColumns(13, sheet.getMaxColumns() - 12);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
    console.log('created new named functions sheet');

    sheet
      .getRange(2, 1, res.length, 2)
      .setValues([
        ...res
          .sort(
            (a, b) =>
              (a.definedFunction.startsWith('LAMBDA') &&
              !b.definedFunction.startsWith('LAMBDA')
                ? -1
                : !a.definedFunction.startsWith('LAMBDA') &&
                  b.definedFunction.startsWith('LAMBDA')
                ? 1
                : 0) ||
              a.definedName.localeCompare(b.definedName) ||
              a.definedFunction.localeCompare(b.definedFunction),
          )
          .map(({ definedName, definedFunction }) => [
            definedName,
            definedFunction,
          ]),
      ]);

    if (sheet.getMaxRows() > res.length + 1)
      sheet.deleteRows(res.length + 2, sheet.getMaxRows() - res.length - 1);
    tableFormat(sheet.getDataRange(), {
      color: 'GREEN',
      hasTitle: false,
      hasHeaders: true,
      centerAll: false,
      alternating: false,
    });

    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          `Successful refresh of named functions sheet`,
        ),
      )
      .build();
  }
  export function createSheetWithToggleButtons() {
    const tictoc = Shlog.tic('createSheetWithToggleButtons');
    const sheet = (
      g.ss.getSheetByName('Color Toggles') || g.ss.insertSheet('Color Toggles')
    ).clear();

    sheet.clearConditionalFormatRules();
    const activeRules = [];
    const pallete = Object.values(MP);
    console.log('pallete :>> ', pallete);
    const start = 1;
    const gap = 1;
    const range = sheet.getRange(5, 5, 18, 20 - gap - start + 1);
    range
      .insertCheckboxes()
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        '#0c343d',
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
      );

    for (let i = 0; i < 18; i += 1) {
      for (let j = 1; j <= 20 - gap; j += 1) {
        const cell = sheet.getRange(5 + i, 5 - start + j);
        const activeHex = pallete[i][2];
        const passiveHex = pallete[i][j + gap];
        cell
          .setBackground(passiveHex)
          .setFontColor(
            passiveHex.slice(-2) === '00'
              ? passiveHex.slice(0, -1) + '1'
              : passiveHex.slice(0, -2) +
                  (parseInt(passiveHex.slice(-2), 16) - 1)
                    .toString(16)
                    .padStart(2, '0'),
          );
        activeRules.push(
          SpreadsheetApp.newConditionalFormatRule()
            .setRanges([cell])
            .whenFormulaSatisfied(`=${cell.getA1Notation()}=true`)
            .setBackground(activeHex)
            .setFontColor(
              activeHex.slice(-2) === '00'
                ? activeHex.slice(0, -1) + '1'
                : activeHex.slice(0, -2) +
                    (parseInt(activeHex.slice(-2), 16) - 1)
                      .toString(16)
                      .padStart(2, '0'),
            )
            .build(),
        );
      }
    }

    sheet
      .setColumnWidths(5, 20 - gap - start + 1, 20)
      .setConditionalFormatRules(activeRules);

    return (
      Shlog.toc(tictoc) &&
      CardService.newActionResponseBuilder()
        .setNotification(
          CardService.newNotification().setText(
            `Finished creating the toggle sheet. Delete it when finished.`,
          ),
        )
        .build()
    );
  }
  export function getAllFormulas() {
    const tempS = g.ss.insertSheet('temp');
    const sheets = g.ss
      .getSheets()
      .filter((s) => s.getName() !== '__pages__' && s.getName() !== 'temp');
    let values = [
      sheets.map(
        (s, i) => `=LET(name,"${s.getName()}",
     idx,${i + 1},
   IF(ISERROR(INDIRECT(name&"!A1")),
      HSTACK(idx,name,"DELETED"),
      LET(rng,INDIRECT(name&"!A:ZZZ"),
          maxR,ROWS(rng),
          maxC,COLUMNS(rng),
          nCells,maxR*maxC,
          lastR,MAX(ARRAYFORMULA(SEQUENCE(maxR)*N(rng<>""))),
          lastC,MAX(ARRAYFORMULA(SEQUENCE(1,maxC)*N(rng<>""))),
          nData,lastR*lastC,
          nActive,COUNTA(OFFSET(rng,0,0,lastR-1,lastC-1)),
          formulas,TOCOL(FLATTEN(MAP(OFFSET(rng,0,0,lastR,lastC),LAMBDA(x,IFERROR(FORMULATEXT(x))))),1),
          nFormulas,ROWS(formulas),
          result,VSTACK(idx,name,,maxR,maxC,nCells,nCells,lastR,lastC,nData,nData/nCells,nData/nCells,nActive,nActive/nCells,nActive/nCells,nFormulas,IF(nFormulas,formulas,"No Formulas")),
        result)))`,
      ),
    ];
    tempS.getRange(1, 1, 1, sheets.length).setValues(values);
    SpreadsheetApp.flush();
    values = tempS.getDataRange().getValues();
    g.ss.deleteSheet(tempS);
    const pagesData = [];
    for (let j = 0; j < values[0].length; j += 1) {
      const pageRow = [];
      for (let i = 0; i < 16; i += 1) {
        pageRow.push(values[i][j]);
      }
      pagesData.push(pageRow);
      for (let i = 16; i < values.length; i += 1) {
        if (values[i][j]) pagesData.push([null, `'${values[i][j]}`]);
      }
    }
    return pagesData;
  }

  export function createStatsDashboard() {
    // leave these two variables as is
    const TEMPLATE_SSID = '1vmgmyaphx9dcIbcnbK0mNwVWE7qHOflf975mEOgJm14';
    const PLAIN_SHEETNAME = '__stats__';
    const RAINBOW_SHEETNAME = '__rainbow_stats__';

    // change this if you want to,
    const DESTINATION_SHEETNAME = '__stats__';

    const ui = SpreadsheetApp.getUi();
    const ssidResponse = ui.prompt(
      'Enter the destination SSID',
      ui.ButtonSet.OK_CANCEL,
    );

    if (ssidResponse.getSelectedButton() == ui.Button.CANCEL) return;

    const colorResponse = ui.alert(
      'Which color theme?',
      'Click YES for the rainbow or No for plain.',
      ui.ButtonSet.YES_NO,
    );

    const destSS = ssidResponse.getResponseText()
      ? SpreadsheetApp.openById(ssidResponse.getResponseText())
      : SpreadsheetApp.getActive();

    if (destSS.getSheetByName(DESTINATION_SHEETNAME))
      destSS.deleteSheet(destSS.getSheetByName(DESTINATION_SHEETNAME));

    const sheetNames = destSS
      .getSheets()
      .map((s) => s.getName())
      .filter((sn) => !sn.startsWith('__') && !(sn == DESTINATION_SHEETNAME));

    const statsSheet = SpreadsheetApp.openById(TEMPLATE_SSID)
      .getSheetByName(
        colorResponse == ui.Button.YES ? RAINBOW_SHEETNAME : PLAIN_SHEETNAME,
      )
      .copyTo(destSS)
      .setName(DESTINATION_SHEETNAME)
      .getRange('A1')
      .check()
      .getSheet();

    sheetNames.forEach((sn, i) =>
      statsSheet
        .getRange(3 + 100 * i, colorResponse == ui.Button.YES ? 4 : 2)
        .setValue(sn),
    );

    SpreadsheetApp.flush();
  }

  /* --------------------------------------------- Get Background Colors -------------------------------------------- */
  export function getBackgroundColorsToNotes() {
    const backgrounds = g.ActiveRange.getBackgrounds();
    const notes = backgrounds.map((r) =>
      r.map((v) => {
        if (v.toUpperCase() === '#FFFFFF' || !v) return;
        return [
          v,
          'R: ' + parseInt(v.slice(1, 3), 16),
          'G: ' + parseInt(v.slice(3, 5), 16),
          'B: ' + parseInt(v.slice(-2), 16),
        ].join('\n');
      }),
    );
    g.ActiveRange.setNotes(notes);
  }

  export function getBackgroundColorsToValues() {
    g.ActiveRange.setValues(g.ActiveRange.getBackgrounds());
  }
  /* -------------------------------------- Set Background Colors Using Values -------------------------------------- */
  export function setBackgroundsColorsUsingValues() {
    const formulas = g.ActiveRange.getFormulas();
    const values = g.ActiveRange.getValues();
    g.ActiveRange.setBackgrounds(
      values.map((r) =>
        r.map((c) =>
          c.toString()[0] === '#' ? c.toString() : '#' + c.toString(),
        ),
      ),
    );
    for (let i = 0; i < values.length; i += 1) {
      for (let j = 0; j < values[0].length; j += 1) {
        if (formulas[i][j] || /#[a-zA-Z0-9]{6}/.test(values[i][j].toString()))
          continue;

        g.ActiveRange.offset(i, j, 1, 1).setValue(
          '#' + values[i][j].toString(),
        );
      }
    }
  }

  export function showFileScopeRequestCard() {
    let card = Views.buildFileScopeRequestCard();
    return [card];
  }

  export function shiftFormulaDown() {
    const headerCell = g.ActiveSheet.getRange(
      1,
      g.ActiveRange.getColumn(),
      1,
      1,
    );
    const formula = headerCell.getFormula();
    if (!formula) return;

    headerCell.setNote(`Continue${formula}`);
    const staticRange = headerCell.offset(0, 0, g.ActiveRange.getRow() - 1);
    staticRange.setValues(staticRange.getDisplayValues());
    g.ActiveRange.setFormula(formula);

    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(`Finished shifting down formula`),
      )
      .build();
  }

  /**
   * Shows the user settings card.
   * @return {UniversalActionResponse}
   */
  export function showSettings(
    e: GoogleAppsScript.Addons.EventObject,
  ): GoogleAppsScript.Card_Service.UniversalActionResponse {
    // var settings = Settings.getSettingsForUser();
    var card = Views.buildSettingsCard();

    return CardService.newUniversalActionResponseBuilder()
      .displayAddOnCards([card])
      .build();
  }

  /**
   * Saves the user's settings.
   *
   * @param {Event} e - Event from Gmail
   * @return {ActionResponse}
   */
  export function saveSettings(
    e: GoogleAppsScript.Addons.EventObject,
  ): GoogleAppsScript.Card_Service.ActionResponse {
    var {
      backgroundTitle,
      backgroundHeaders,
      backgroundDataFirst,
      backgroundDataSecond,
      backgroundFooter,
      bordersAll,
      bordersHorizontal,
      bordersVertical,
      bordersTitleBottom,
      bordersHeadersBottom,
      bordersHeadersVertical,
      bordersThickness,
      debugControl,
    } = e.commonEventObject.formInputs;
    const tableOptions: string[] =
      e.commonEventObject.formInputs?.tableOptions?.stringInputs?.value || [];
    var settings = {
      hasTitle: tableOptions.includes('hasTitle'),
      hasHeaders: tableOptions.includes('hasHeaders'),
      hasFooter: tableOptions.includes('hasFooter'),
      leaveTop: tableOptions.includes('leaveTop'),
      leaveLeft: tableOptions.includes('leaveLeft'),
      leaveBottom: tableOptions.includes('leaveBottom'),
      noBottom: tableOptions.includes('noBottom'),
      centerAll: tableOptions.includes('centerAll'),
      alternating: tableOptions.includes('alternating'),
      backgroundTitle: parseInt(backgroundTitle.stringInputs.value[0]),
      backgroundHeaders: parseInt(backgroundHeaders.stringInputs.value[0]),
      backgroundDataFirst: parseInt(backgroundDataFirst.stringInputs.value[0]),
      backgroundDataSecond: parseInt(
        backgroundDataSecond.stringInputs.value[0],
      ),
      backgroundFooter: parseInt(backgroundFooter.stringInputs.value[0]),
      bordersAll: parseInt(bordersAll.stringInputs.value[0]),
      bordersHorizontal: parseInt(bordersHorizontal.stringInputs.value[0]),
      bordersVertical: parseInt(bordersVertical.stringInputs.value[0]),
      bordersTitleBottom: parseInt(bordersTitleBottom.stringInputs.value[0]),
      bordersHeadersBottom: parseInt(
        bordersHeadersBottom.stringInputs.value[0],
      ),
      bordersHeadersVertical: parseInt(
        bordersHeadersVertical.stringInputs.value[0],
      ),
      bordersThickness: parseInt(bordersThickness.stringInputs.value[0]),
      debugControl: debugControl ? 'ON' : 'OFF',
    };

    Settings.updateSettingsForUser(settings);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().popCard())
      .setNotification(CardService.newNotification().setText('Settings saved.'))
      .build();
  }

  /**
   * Resets the user settings to the defaults.
   * @return {ActionResponse}
   */
  export function resetSettings(): GoogleAppsScript.Card_Service.ActionResponse {
    Settings.resetSettingsForUser();
    var card = Views.buildSettingsCard();
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(card))
      .setNotification(CardService.newNotification().setText('Settings reset.'))
      .build();
  }

  export function prepareForHelp() {
    Settings.updateSettingsForUser({ ...g.UserSettings, helpControl: 'on' });
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          `Now click on a command button to see more information about that command.`,
        ),
      )
      .build();
  }

  export function getHelp(actionName: string) {
    Settings.updateSettingsForUser({ ...g.UserSettings, helpControl: 'off' });
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          `Here's your help card on ${actionName}`,
        ),
      )
      .build();
  }

  /**
   * Returns collection of homepage cards
   * @return {GoogleAppsScript.Card_Service.Card|GoogleAppsScript.Card_Service.Card[]}
   */
  export function showSheetsHomepage():
    | GoogleAppsScript.Card_Service.Card
    | GoogleAppsScript.Card_Service.Card[] {
    var toolsCard = Views.buildToolsCard();
    return [toolsCard];
  }

  export function formatAsCaption(e) {
    const location = e.parameters.location;
    const [vertical, horizontal] = location.split('-');
    if (vertical === 'clear') {
      g.ss.getActiveRangeList().clearFormat();
    } else {
      g.ss
        .getActiveRangeList()
        .setFontStyle('italic')
        .setVerticalAlignment(vertical)
        .setFontSize(8)
        .setFontFamily('Roboto')
        .setHorizontalAlignment(horizontal);
    }
  }

  export function showFormula() {
    let formula = "'" + g.ActiveRange.offset(1, 0, 1, 1).getFormula();
    if (formula.length === 1) {
      formula = "'" + g.ActiveRange.offset(2, 0, 1, 1).getFormula();
    }
    const parts = Utils.groupText(formula);
    let cum = -1;

    g.ActiveRange.setFontColor(COLORS.GREEN.DARKER)
      .setFontWeight('bold')
      .setFontStyle('italic')
      .setFontFamily('Inconsolata');

    const richText = SpreadsheetApp.newRichTextValue().setText(formula);

    if (parts.length >= 3) {
      richText.setTextStyle(
        (cum += parts[0].length),
        (cum += parts[1].length),
        SpreadsheetApp.newTextStyle()
          .setForegroundColor(COLORS.ORANGE.BOLD)
          .build(),
      );
    }

    if (parts.length >= 5) {
      richText.setTextStyle(
        (cum += 1),
        (cum += parts[3].length),
        SpreadsheetApp.newTextStyle()
          .setForegroundColor(COLORS.PURPLE.BOLD)
          .build(),
      );
    }

    if (parts.length >= 7) {
      richText.setTextStyle(
        (cum += 1),
        (cum += parts[5].length),
        SpreadsheetApp.newTextStyle()
          .setForegroundColor(COLORS.BLUE.BOLD)
          .build(),
      );
    }

    g.ActiveRange.setRichTextValue(richText.build());
  }

  export function makeRangeStatic() {
    makeRangeStatic_(g.ActiveRange);
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(`Successfully made range static`),
      )
      .build();
  }

  export function makeRangeDynamic() {
    makeRangeDynamic_(g.ActiveRange);
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          `Successfully restored range formulas`,
        ),
      )
      .build();
  }

  export function makeSheetStatic() {
    makeSheetStatic_(g.ActiveSheet);
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(`Successfully made sheet static`),
      )
      .build();
  }

  export function makeSheetDynamic() {
    makeSheetDynamic_(g.ActiveSheet);
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          `Successfully restored sheet formulas`,
        ),
      )
      .build();
  }

  /** Makes sheet static by placing content of either hard-coded values or functions into notes.
   * @param {GoogleAppsScript.Spreadsheet.Sheet|string} [sheet] - Optional sheet object or name of sheet to make static (defaults to active sheet)
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   */
  function makeSheetStatic_(
    sheet?: GoogleAppsScript.Spreadsheet.Sheet | string,
  ): GoogleAppsScript.Spreadsheet.Sheet {
    const s = sheet
      ? typeof sheet === 'string'
        ? g.ss.getSheetByName(sheet)
        : sheet
      : g.ActiveSheet;

    const range = s.getDataRange();

    const response = makeRangeStatic_(range);

    if (response !== 'failed') s.setTabColor('#E92D18');

    return s;
  }

  /** Places all notes into cells, whether they be formulas or values
   * @param {GoogleAppsScript.Spreadsheet.Sheet|string=} [sheet] - Optional sheet object or name of sheet to make dynamic (defaults to active sheet)
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   */
  function makeSheetDynamic_(
    sheet?: GoogleAppsScript.Spreadsheet.Sheet | string,
  ): GoogleAppsScript.Spreadsheet.Sheet {
    const s = sheet
      ? typeof sheet === 'string'
        ? g.ss.getSheetByName(sheet)
        : sheet
      : g.ActiveSheet;

    const range = s.getDataRange();

    const response = makeRangeDynamic_(range);

    // change tab color to indicate sheet is dynamic
    if (response !== 'failed') s.setTabColor('#249A41');

    return s;
  }

  /** Make range notes the user entered value/formula, but avoid shadow values
   * @param {GoogleAppsScript.Spreadsheet.Range} [range] - range to make static, otherwise uses active range
   * @returns {GoogleAppsScript.Spreadsheet.Range|string} returns range for chaining
   */
  function makeRangeStatic_(
    range?: GoogleAppsScript.Spreadsheet.Range,
  ): GoogleAppsScript.Spreadsheet.Range | string {
    const r = range || g.ActiveRange;

    // save original values for the end
    const originalValues = r.getValues();

    // get range formulas, where no formula means ''
    const formulas = r.getFormulas();

    if (new Set(r.getNotes().flat()).size > 1) {
      const confirmation = Utils.askQuestion(
        'Warning - appears to already be static',
        'Are you sure you want to continue?',
      );

      if (!confirmation) {
        Utils.displayError('Process cancelled!');
        return 'failed';
      }
    }

    // clear cell content for each formula
    formulas.forEach((row, i) => {
      row.forEach((cell, j) => {
        if (cell) r.offset(i, j, 1, 1).clearContent();
      });
    });

    const combinedNotes = r
      .getValues()
      .map((row, i) => row.map((cell, j) => (cell ? cell : formulas[i][j])));
    // set notes with updated values plus formulas to remove shadow values
    r.setNotes(combinedNotes);

    r.setValues(originalValues); // finally fill range with original values

    return r;
  }

  /** Replaces range values with value/formula stored in cell notes.
   * @param {GoogleAppsScript.Spreadsheet.Range} [range] - range to make dynamic, otherwise uses active range
   * @returns {GoogleAppsScript.Spreadsheet.Range|string} returns range for chaining
   */
  function makeRangeDynamic_(
    range?: GoogleAppsScript.Spreadsheet.Range,
  ): GoogleAppsScript.Spreadsheet.Range | string {
    const r = range || g.ActiveRange;
    if (new Set(r.getFormulas().flat()).size > 1) {
      const confirmation = Utils.askQuestion(
        'Warning - appears to already be dynamic',
        'Are you sure you want to continue?',
      );

      if (!confirmation) {
        Utils.displayError('Process cancelled!');
        return 'failed';
      }
    }

    r.clearContent().setValues(r.getNotes()).setNote(null);

    return r;
  }

  export function createNamedRangesFromSheet() {
    const ss = g.ss;
    const sheetName = g.ActiveSheet.getName();
    const dataRange = g.ActiveSheet.getDataRange();
    const startColIndex = dataRange.getColumn() - 1;
    const width = dataRange.getWidth();
    const abc = g.ABCs.slice(startColIndex, startColIndex + width);
    const fields = dataRange
      .offset(0, 0, 1)
      .getValues()[0]
      .map((h) => Utils.titleCaseAndRemoveSpaces(h));

    const tableName = Utils.titleCaseAndRemoveSpaces(sheetName);

    const fieldList = [];
    const rangeListA1 = [];

    fields.forEach((field, i) => {
      if (field) {
        fieldList.push(field);
        rangeListA1.push(`${abc[i]}:${abc[i]}`);
      }
    });

    const rangeList = ss.getRangeList(rangeListA1);
    const colRanges = rangeList.getRanges();

    fieldList.forEach((field, i) => {
      if (field) {
        const namedRangeName = `${tableName}.${field}`;
        ss.setNamedRange(namedRangeName, colRanges[i]);

        ss.getNamedRanges()
          .find((nr) => nr.getName() === namedRangeName)
          .setRange(colRanges[i]);
      }
    });
    ss.setNamedRange(
      tableName,
      g.ActiveSheet.getRange(`${abc[0]}:${abc[width - 1]}`),
    );
    ss.getNamedRanges()
      .find((nr) => nr.getName() === tableName)
      .setRange(ss.getRange(`'${sheetName}'!${abc[0]}:${abc[width - 1]}`));
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          `Successfully created named ranges.`,
        ),
      )
      .build();
  }

  export function playground1() {
    const tictoc = Shlog.tic('playground1');

    return (
      Shlog.toc(tictoc) &&
      CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(`Finished:`))
        .build()
    );
  }

  export function playground2() {
    const tictoc = Shlog.tic('playground2');
    return (
      Shlog.toc(tictoc) &&
      CardService.newActionResponseBuilder()
        .setNotification(
          CardService.newNotification().setText(
            `Finished running Playground 2`,
          ),
        )
        .build()
    );
  }

  export function playground3() {
    const tictoc = Shlog.tic('playground3');
    return (
      Shlog.toc(tictoc) &&
      CardService.newActionResponseBuilder()
        .setNotification(
          CardService.newNotification().setText(`process complete`),
        )
        .build()
    );
  }
}

export { ActionHandlers };
