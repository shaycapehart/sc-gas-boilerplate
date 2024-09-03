import { Settings } from '@core/environment/Settings';
import { MP, COLORS } from '@core/lib/COLORS';
import { SortType, TableFormatOptionsType } from '@core/types/addon';
import { Utils } from '@core/utils/Utils';
import { Views } from './Views';
import { Shlog, TicToc } from '@core/logging/Shlog';
import { Help } from '@core/lib/Help';
import { ImportJSON } from '@core/lib/ImportJSON';

/**
 * Collection of functions to handle user interactions with the add-on.
 *
 * @namespace
 */
namespace ActionHandlers {
  /* -------------------------------------------------- Empty Logs -------------------------------------------------- */
  export function emptyLogRecords(): GoogleAppsScript.Card_Service.ActionResponse {
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
    const tictoc = Shlog.tic('tableFormat');
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

    return Shlog.toc(tictoc);
  }
  export function setTableFormatDefaults(
    e: GoogleAppsScript.Addons.EventObject,
  ) {
    const tictoc = Shlog.tic('setTableFormatDefaults');
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
    return finished(`Table format options updated`, tictoc);
  }
  export function formatRangeAsTable(
    e: GoogleAppsScript.Addons.EventObject,
    range?: GoogleAppsScript.Spreadsheet.Range,
  ) {
    const tictoc = Shlog.tic('formatRangeAsTable');
    const rng = range || g.ActiveRange;
    const tableOptions: string[] =
      e.commonEventObject.formInputs?.tableOptions?.stringInputs?.value || [];

    const color = e.commonEventObject.parameters.color,
      hasTitle = tableOptions.includes('hasTitle'),
      hasHeaders = tableOptions.includes('hasHeaders'),
      leaveTop = tableOptions.includes('leaveTop'),
      leaveLeft = tableOptions.includes('leaveLeft'),
      leaveBottom = tableOptions.includes('leaveBottom'),
      noBottom = tableOptions.includes('noBottom'),
      centerAll = tableOptions.includes('centerAll'),
      alternating = tableOptions.includes('alternating'),
      hasFooter = tableOptions.includes('hasFooter');

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

    return finished(`Successful formatted the table`, tictoc);
  }

  export function squareSelectedCells(e?: GoogleAppsScript.Addons.EventObject) {
    const tictoc = Shlog.tic('squareSelectedCells');

    const pixels =
      parseInt(
        e.commonEventObject.formInputs?.pixels?.stringInputs?.value[0],
      ) || 20;
    const nRows = g.ActiveRange.getHeight();
    const nCols = g.ActiveRange.getWidth();
    const startRow = g.ActiveRange.getRow();
    const startCol = g.ActiveRange.getColumn();
    g.ActiveSheet.setRowHeightsForced(startRow, nRows, pixels).setColumnWidths(
      startCol,
      nCols,
      pixels,
    );

    return finished(`Successful squared off the selected cells`, tictoc);
  }

  export function formatTable(e: GoogleAppsScript.Addons.EventObject) {
    const tictoc = Shlog.tic('formatTable');
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

    return finished(`Successful completed formatTable`, tictoc);
  }

  /**
   * Crops the current sheet to the user's selection.
   */
  export function cropToSelection() {
    const tictoc = Shlog.tic('cropToSelection');
    console.log('Cropping to selection');
    var range = g.ActiveRange;
    cropToRange_(range);

    return finished(`Successful cropped to selection`, tictoc);
  }
  /* ------------------------------------------------- Crop to Data ------------------------------------------------- */
  /**
   * Crops the current sheet to the data.
   */
  export function cropToData() {
    const tictoc = Shlog.tic('cropToData');
    console.log('Cropping to data');
    var range = g.ActiveSheet.getDataRange();
    cropToRange_(range);

    return finished(`Successful cropped to data`, tictoc);
  }
  /* ------------------------------------------------- Crop to Range ------------------------------------------------ */
  /**
   * Crops the sheet such that it only contains the given range.
   * @param {GoogleAppsScript.Spreadsheet.Range} range The range to crop to.
   */
  export function cropToRange_(range) {
    const tictoc = Shlog.tic('cropToRange_');
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

    return Shlog.toc(tictoc);
  }

  function getRangeType_(range: GoogleAppsScript.Spreadsheet.Range): string {
    const tictoc = Shlog.tic('getRangeType_');
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
    return Shlog.toc(tictoc) && type;
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
    const tictoc = Shlog.tic('importRange');
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

    return Shlog.toc(tictoc);
  }

  export function sortRange(
    ssid: string,
    range: string,
    sortParams: SortType[],
  ) {
    const tictoc = Shlog.tic('sortRange');
    if (sortParams.length % 2 !== 0) return;
    const ss = SpreadsheetApp.openById(ssid);
    const rng = ss.getRange(range);
    rng.sort(sortParams);
    return Shlog.toc(tictoc);
  }

  export function outputEventObjectToSheet(
    e: GoogleAppsScript.Addons.EventObject,
  ) {
    let sheetName = 'Event Object';
    let sheet = g.ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = g.ss.insertSheet(sheetName);
    } else {
      // sheet.getDataRange().clear({ contentsOnly: true });
    }
    const data = ImportJSON.parseJSONObject(e);

    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }

  export function getNamedFunctions() {
    const tictoc = Shlog.tic('getNamedFunctions');
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
    return finished(`Successful retrieved named functions`, tictoc);
    // DriveApp.getFiles(); // This comment line is used for automatically detecting the scope of Drive API.
  }

  export function createNamedRangesDashboard(
    e?: GoogleAppsScript.Addons.EventObject,
  ) {
    const tictoc = Shlog.tic('createNamedRangesDashboard');
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

      const sheetIcons = ['range', 'cell', 'table', 'row', 'col'].map(
        (n) =>
          `=image("https://shaycapehart.github.io/web-host-files/data${n}.jpg")`,
      );

      sheet.getRange('K1:O1').setValues([sheetIcons]);
      sheet.hideColumns(11, 5);

      sheet
        .setHiddenGridlines(true)
        .setColumnWidths(1, 1, 200)
        .setColumnWidths(2, 1, 100)
        .setColumnWidths(3, 1, 200)
        .setColumnWidths(4, 1, 400)
        .setColumnWidths(5, 1, 65)
        .setColumnWidths(8, 3, 200);

      sheet
        .getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .offset(0, 7, 1, 3)
        .setNotes([notes]);

      if (sheet.getMaxColumns() > 15)
        sheet.deleteColumns(16, sheet.getMaxColumns() - 15);
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
        SpreadsheetApp.newRichTextValue()
          .setText(name.replaceAll(`'`, ''))
          .setLinkUrl(url)
          .build(),
      ]);
      const typeRegex = /([a-zA-Z]*)(\d*)(:?)([a-zA-Z]*)(\d*)/;
      const matches = typeRegex.exec(a1Notation);
      const [all, a1, n1, sep, a2, n2] = matches;
      const type = !sep
        ? 2
        : !n1 && !n2
          ? a1 === a2
            ? 5
            : 3
          : !n1 || !n2
            ? 1
            : n1 === n2
              ? 4
              : 1;

      rows.push([
        a1Notation,
        sheetName,
        `="'${sheetName}'!${a1Notation}"`,
        `=index($K$1:$O$1,${type})`,
      ]);
    });

    if (rows.length) {
      if (sheet.getMaxRows() > rows.length + 1) {
        sheet.deleteRows(
          rows.length + 2,
          sheet.getMaxRows() - (rows.length + 1),
        );
      }
      sheet
        .getRange(`A2:A${rows.length + 1}`)
        .setRichTextValues(richTextValues);
      sheet.getRange(`B2:E${rows.length + 1}`).setValues(rows);
      sheet.getRange(`E2:E${rows.length + 1}`).setHorizontalAlignment('center');

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
      sheet.setRowHeights(2, rows.length, 60);
      tableFormat(sheet.getDataRange(), {
        color: 'BLUE',
        hasTitle: false,
        hasHeaders: true,
        centerAll: false,
        alternating: false,
      });
    }

    return finished(`Build fresh named ranges dashboard`, tictoc, e);
  }

  export function updateNamedRangeSheet(
    e?: GoogleAppsScript.Addons.EventObject,
  ) {
    const tictoc = Shlog.tic('updateNamedRangeSheet');
    let namedRanges = g.ss.getNamedRanges();
    let rows: any[][];
    let nrS = g.ss.getSheetByName('__named_ranges__');
    if (!nrS) return createNamedRangesDashboard();

    rows = nrS.getDataRange().getValues().slice(1);
    // check for delete items
    rows.forEach((row, i, arr) => {
      if (!!arr[arr.length - 1 - i][5]) {
        try {
          const nr = namedRanges.find(
            (namedRange) =>
              namedRange.getName().replaceAll(`'`, '') ===
              arr[arr.length - 1 - i][0],
          );
          nr.remove();
        } catch (error) {
          Utils.displayError(
            `Problem deleting ${
              arr[arr.length - 1 - i][0]
            } named range: ${error}`,
          );
        }
      }
    });
    // check for edit items
    rows.forEach((row, i, arr) => {
      if (!!arr[arr.length - 1 - i][6]) {
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
        ] = arr[arr.length - 1 - i];
        console.log('Edit row :>> ', arr[arr.length - 1 - i]);
        console.log(
          'nrName >> ',
          nrName,
          'nrA1 >> ',
          nrA1,
          'nrSheetName >> ',
          nrSheetName,
          'nrFull >> ',
          nrFull,

          'nrType >> ',
          nrType,

          'nrDelete >> ',
          nrDelete,

          'nrUpdate >> ',
          nrUpdate,

          'newName >> ',
          newName,

          'newRangeA1 >> ',
          newRangeA1,

          'newSheetName >> ',
          newSheetName,
        );
        const cleanNr = nrName.replace(/^.*!/, '');
        try {
          const existingNr = namedRanges.find(
            (namedRange) => namedRange.getName().replaceAll(`'`, '') === nrName,
          );
          const namedRangeRangeA1 = `'${newSheetName || nrSheetName}'!${
            newRangeA1 || nrA1
          }`;
          const namedRangeName = `${newName || nrName}`;
          console.log('namedRangeName :>> ', namedRangeName);
          existingNr.setName(namedRangeName);
          existingNr.setRange(g.ss.getRange(namedRangeRangeA1));

          g.ss
            .getNamedRanges()
            .find(
              (nr) =>
                nr.getName().replaceAll(`'`, '') ===
                namedRangeName.replaceAll(`'`, ''),
            )
            .setRange(g.ss.getRange(namedRangeRangeA1));
        } catch (error) {
          Utils.displayError(
            `Problem modifying ${newName || nrName} named range.`,
          );
        }
      }
    });

    return Shlog.toc(tictoc) && createNamedRangesDashboard(e);
  }

  export function createNamedFunctionsDashboard(): GoogleAppsScript.Spreadsheet.Sheet {
    const tictoc = Shlog.tic('createNamedFunctionsDashboard');
    const headers = [
      [
        'Defined Name',
        'Defined Function',
        '=BYROW(B:B,LAMBDA(fn,IF(ISBLANK(fn),,IF(ROW(fn)=ROW(B:B),HSTACK("Function",INDEX("Argument"&SEQUENCE(1,9))),LET(str,SUBSTITUTE(fn,CHAR(10),""),IFERROR(HSTACK(REGEXEXTRACT(str,"(?i)^LAMBDA\\((?:[a-z0-9_,\\s]+)(?:,\\s?)(.+)\\)$"),SPLIT(REGEXEXTRACT(str,"(?i)^LAMBDA\\(([a-z0-9_,\\s]+)(?:,\\s?)(?:.+)\\)$"),", ",0,0)),IFERROR(REGEXEXTRACT(str,"(?i)LAMBDA\\((.*)\\)$"))))))))',
      ],
    ];

    let sheet: GoogleAppsScript.Spreadsheet.Sheet,
      range: GoogleAppsScript.Spreadsheet.Range;

    sheet = g.ss.getSheetByName('__named_functions__');

    if (!sheet) {
      sheet = g.ss.insertSheet('__named_functions__');

      range = sheet
        .clear()
        .setHiddenGridlines(true)
        .setColumnWidth(1, 250)
        .setColumnWidth(2, 600)
        .setColumnWidth(3, 600)
        .setColumnWidths(4, 9, 150)
        .getRange(1, 1, sheet.getMaxRows(), 12)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

      // sheet.deleteColumns(13, sheet.getMaxColumns() - 12);
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(1);
    }

    sheet.clear().getRange(1, 1, 1, headers[0].length).setValues(headers);

    return Shlog.toc(tictoc) && sheet;
  }
  /* ---------------------------------------- Show Named Functions Dashboard ---------------------------------------- */
  export function showNamedFunctionsDashboard() {
    const tictoc = Shlog.tic('showNamedFunctionsDashboard');
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

    const sheet = createNamedFunctionsDashboard();

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
    if (sheet.getMaxRows() > sheet.getLastRow())
      sheet.deleteRows(
        sheet.getLastRow() + 1,
        sheet.getMaxRows() - sheet.getLastRow(),
      );
    if (sheet.getMaxColumns() > sheet.getLastColumn())
      sheet.deleteColumns(
        sheet.getLastColumn() + 1,
        sheet.getMaxColumns() - sheet.getLastColumn(),
      );
    tableFormat(sheet.getDataRange(), {
      color: 'GREEN',
      hasTitle: false,
      hasHeaders: true,
      centerAll: false,
      alternating: false,
    });

    return finished(`Successful refreshed named functions`, tictoc);
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
    const range = sheet.getRange(5, 5, 10, 18 - gap - start + 1);
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

    for (let i = 0; i < 10; i += 1) {
      for (let j = 1; j <= 18 - gap; j += 1) {
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
      .setColumnWidths(5, 18 - gap - start + 1, 20)
      .setConditionalFormatRules(activeRules);

    return finished(
      `Finished creating the toggle sheet. Delete it when finished.`,
      tictoc,
    );
  }
  export function getAllFormulas() {
    const tictoc = Shlog.tic('getAllFormulas');
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
    return Shlog.toc(tictoc) && pagesData;
  }

  export function createStatsDashboard() {
    const tictoc = Shlog.tic('createStatsDashboard');
    // leave these two variables as is
    const TEMPLATE_SSID = '1vmgmyaphx9dcIbcnbK0mNwVWE7qHOflf975mEOgJm14';
    const SHEETNAME = '__stats__';

    // change this if you want to,
    const DESTINATION_SHEETNAME = '__stats__';

    const ui = SpreadsheetApp.getUi();
    const ssidResponse = ui.prompt(
      'Enter the destination SSID (leave blank and hit ok to use current sheet)',
      ui.ButtonSet.OK_CANCEL,
    );

    if (ssidResponse.getSelectedButton() == ui.Button.CANCEL) return;

    // const colorResponse = ui.alert(
    //   'Which color theme?',
    //   'Click YES for the rainbow or No for plain.',
    //   ui.ButtonSet.YES_NO,
    // );

    const destSS = ssidResponse.getResponseText()
      ? SpreadsheetApp.openById(ssidResponse.getResponseText())
      : g.ss;

    if (destSS.getSheetByName(DESTINATION_SHEETNAME))
      destSS.deleteSheet(destSS.getSheetByName(DESTINATION_SHEETNAME));

    const sheetNames = destSS
      .getSheets()
      .map((s) => s.getName())
      .filter((sn) => !sn.startsWith('__') && !(sn == DESTINATION_SHEETNAME));

    const statsSheet = SpreadsheetApp.openById(TEMPLATE_SSID)
      .getSheetByName(SHEETNAME)
      .copyTo(destSS)
      .setName(DESTINATION_SHEETNAME)
      .getRange('A1')
      .check()
      .getSheet();

    sheetNames.forEach((sn, i) =>
      statsSheet.getRange(3 + 100 * i, 4).setValue(sn),
    );

    SpreadsheetApp.flush();

    return finished(`Successful created stats dashboard`, tictoc);
  }

  /* --------------------------------------------- Get Background Colors -------------------------------------------- */
  export function getBackgroundColorsToNotes() {
    const tictoc = Shlog.tic('getBackgroundColorsToNotes');
    const backgrounds = g.ActiveRange.getBackgrounds();
    const data = { values: [], fonts: [] };

    backgrounds.forEach((r) => {
      const dataRow = [];
      const fontsRow = [];

      r.forEach((c) => {
        const hex = c.toUpperCase();
        const brightness = Utils.brightnessByColor(hex);
        const niceFontColor = Utils.niceColor(hex);
        const fontHex = brightness > 130 ? '#000000' : '#ffffff';
        const red = parseInt(hex.slice(-6, -4), 16);
        const green = parseInt(hex.slice(-4, -2), 16);
        const blue = parseInt(hex.slice(-2), 16);
        dataRow.push(
          `Brightness: ${brightness}\nR: ${red}\nG: ${green}\nB: ${blue}\nHex: ${hex}\nFont: ${fontHex}\nNice Font: ${niceFontColor}`,
        );
        fontsRow.push(fontHex);
      });
      data.values.push(dataRow);
      data.fonts.push(fontsRow);
    });
    console.log('data :>> ', data);

    g.ActiveRange.setNotes(data.values).setFontColors(data.fonts);

    return finished(`Successful backgrounds to notes`, tictoc);
  }

  export function getBackgroundColorsToValues() {
    const tictoc = Shlog.tic('getBackgroundColorsToValues');
    g.ActiveRange.setValues(g.ActiveRange.getBackgrounds());
    return finished(`Successfully retrieved background colors`, tictoc);
  }

  export function matchFontToBackground() {
    const tictoc = Shlog.tic('matchFontToBackground');
    const backgroundColors = g.ActiveRange.getBackgrounds();
    const fontColors = backgroundColors.map((r) =>
      r.map((c) =>
        c.endsWith('f')
          ? c.slice(0, -1) + 'e'
          : c.slice(0, -2) +
            ('0' + (parseInt(c.slice(-2), 16) + 1).toString(16)).slice(-2),
      ),
    );
    g.ActiveRange.setFontColors(fontColors);
    return finished(
      `Successfully matched font color to the background`,
      tictoc,
    );
  }

  export function makeFontNiceColor() {
    const tictoc = Shlog.tic('makeFontNiceColor');
    const backgroundColors = g.ActiveRange.getBackgrounds();
    const fontColors = backgroundColors.map((r) => r.map(Utils.niceColor));
    g.ActiveRange.setFontColors(fontColors);
    return finished(
      `Successfully matched font color to the background`,
      tictoc,
    );
  }
  /* -------------------------------------- Set Background Colors Using Values -------------------------------------- */
  export function setBackgroundsColorsUsingValues() {
    const tictoc = Shlog.tic('setBackgroundsColorsUsingValues');
    const values = g.ActiveRange.getValues();

    g.ActiveRange.setBackgrounds(
      values.map((r) =>
        r.map((c) =>
          c === 0
            ? '#000000'
            : c.toString()[0] === '#'
              ? c.toString()
              : '#' + c.toString(),
        ),
      ),
    );

    ActionHandlers.matchFontToBackground();

    return finished(`Successfully set backgrounds using vals`, tictoc);
  }

  export function showFileScopeRequestCard(): GoogleAppsScript.Card_Service.Card[] {
    let card = Views.buildFileScopeRequestCard();
    return [card];
  }

  export function shiftFormulaDown() {
    const tictoc = Shlog.tic('shiftFormulaDown');
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

    return finished(`Successfully shifted formula down`, tictoc);
  }

  /**
   * Shows the user settings card.
   * @param {Event} e - Event from Gmail
   * @return {UniversalActionResponse}
   */
  export function showSettings(e) {
    const tictoc = Shlog.tic('showSettings');
    // console.log('e :>> ', e);
    // var settings = Settings.getSettingsForUser();
    // console.log('settings :>> ', settings);
    if (!e) {
      var card = Views.buildSettingsCard({
        backgroundTitle: g.UserSettings.backgroundTitle,
        backgroundHeaders: g.UserSettings.backgroundHeaders,
        backgroundDataFirst: g.UserSettings.backgroundDataFirst,
        backgroundDataSecond: g.UserSettings.backgroundDataSecond,
        backgroundFooter: g.UserSettings.backgroundFooter,
        bordersAll: g.UserSettings.bordersAll,
        bordersHorizontal: g.UserSettings.bordersHorizontal,
        bordersVertical: g.UserSettings.bordersVertical,
        bordersTitleBottom: g.UserSettings.bordersTitleBottom,
        bordersHeadersBottom: g.UserSettings.bordersHeadersBottom,
        bordersHeadersVertical: g.UserSettings.bordersHeadersVertical,
        debugControl: g.UserSettings.debugControl,
        helpControl: g.UserSettings.helpControl,
        bordersThickness: g.UserSettings.bordersThickness,
      });
      return finished('Settings shown', tictoc, e);
    }
    return showResults(JSON.stringify(g.UserSettings), tictoc, e);
  }

  /**
   * Saves the user's settings.
   *
   * @param {GoogleAppsScript.Addons.EventObject} e - Event from Gmail
   * @return {GoogleAppsScript.Card_Service.ActionResponse}
   */
  export function saveSettings(
    e: GoogleAppsScript.Addons.EventObject,
  ): GoogleAppsScript.Card_Service.ActionResponse {
    console.log(
      'e.commonEventObject :>> ',
      JSON.stringify(e.commonEventObject),
    );
    const previousSettings = Settings.getSettingsForUser();
    var formInputs = e.commonEventObject.formInputs;
    var settings = {
      bordersTitleBottom: parseInt(
        formInputs.bordersTitleBottom.stringInputs.value[0],
      ),
      backgroundDataSecond: parseInt(
        formInputs.backgroundDataSecond.stringInputs.value[0],
      ),
      backgroundHeaders: parseInt(
        formInputs.backgroundHeaders.stringInputs.value[0],
      ),
      bordersHeadersVertical: parseInt(
        formInputs.bordersHeadersVertical.stringInputs.value[0],
      ),
      bordersAll: parseInt(formInputs.bordersAll.stringInputs.value[0]),
      backgroundDataFirst: parseInt(
        formInputs.backgroundDataFirst.stringInputs.value[0],
      ),
      bordersThickness: parseInt(
        formInputs.bordersThickness.stringInputs.value[0],
      ),
      backgroundTitle: parseInt(
        formInputs.backgroundTitle.stringInputs.value[0],
      ),
      debugControl: formInputs.debugControl ? 'ON' : 'OFF',
      bordersVertical: parseInt(
        formInputs.bordersVertical.stringInputs.value[0],
      ),
      bordersHorizontal: parseInt(
        formInputs.bordersHorizontal.stringInputs.value[0],
      ),
      backgroundFooter: parseInt(
        formInputs.backgroundFooter.stringInputs.value[0],
      ),
      bordersHeadersBottom: parseInt(
        formInputs.bordersHeadersBottom.stringInputs.value[0],
      ),
      helpControl: 'off',
    };

    Settings.updateSettingsForUser(settings);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().popCard())
      .setNotification(CardService.newNotification().setText('Settings saved.'))
      .build();
  }

  /**
   * Resets the user settings to the defaults.
   * @return {GoogleAppsScript.Card_Service.ActionResponse}
   */
  export function resetSettings(): GoogleAppsScript.Card_Service.ActionResponse {
    Settings.resetSettingsForUser();
    var settings = Settings.getSettingsForUser();
    var card = Views.buildSettingsCard({
      backgroundTitle: settings.backgroundTitle,
      backgroundHeaders: settings.backgroundHeaders,
      backgroundDataFirst: settings.backgroundDataFirst,
      backgroundDataSecond: settings.backgroundDataSecond,
      backgroundFooter: settings.backgroundFooter,
      bordersAll: settings.bordersAll,
      bordersHorizontal: settings.bordersHorizontal,
      bordersVertical: settings.bordersVertical,
      bordersTitleBottom: settings.bordersTitleBottom,
      bordersHeadersBottom: settings.bordersHeadersBottom,
      bordersHeadersVertical: settings.bordersHeadersVertical,
      bordersThickness: settings.bordersThickness,
      debugControl: settings.debugControl,
      helpControl: settings.helpControl,
    });
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(card))
      .setNotification(CardService.newNotification().setText('Settings reset.'))
      .build();
  }

  export function createToggles(): GoogleAppsScript.Card_Service.ActionResponse {
    const tictoc = Shlog.tic('createToggles');

    selectColor();
    const activeRules = g.ActiveSheet.getConditionalFormatRules;

    g.ActiveRange.insertCheckboxes().setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      '#0c343d',
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
    );

    return (
      Shlog.toc(tictoc) &&
      CardService.newActionResponseBuilder()
        .setNotification(
          CardService.newNotification().setText(
            `Now click on one of the top row of color icons to use for the toggles.`,
          ),
        )
        .build()
    );
  }

  export function selectColor(e?: GoogleAppsScript.Addons.EventObject) {
    const tictoc = Shlog.tic('selectColor');
    if (g.UserSettings.colorPicker === 'off') {
      Settings.updateSettingsForUser({ ...g.UserSettings, colorPicker: 'on' });
      return finished(
        `Now click on one of the top row of color icons to use for the toggles.`,
        tictoc,
      );
    }
    Settings.updateSettingsForUser({ ...g.UserSettings, colorPicker: 'off' });
    const color = e.commonEventObject.parameters.color;
    return finished(`Here's your the color you picked ${color}`, tictoc);
  }

  export function prepareForHelp() {
    const tictoc = Shlog.tic('prepareForHelp');
    Settings.updateSettingsForUser({ ...g.UserSettings, helpControl: 'on' });

    return finished(
      `Now click on a command button to see more information about that command.`,
      tictoc,
    );
  }

  export function getHelp(event: GoogleAppsScript.Addons.EventObject) {
    const tictoc = Shlog.tic('getHelp');

    if (g.UserSettings.helpControl === 'off') {
      Settings.updateSettingsForUser({ ...g.UserSettings, helpControl: 'on' });
      return finished(
        `Now click on a command button to see more information about that command.`,
        tictoc,
      );
    }
    var actionName = event.commonEventObject.parameters.action;
    Settings.updateSettingsForUser({ ...g.UserSettings, helpControl: 'off' });
    const info = Help[actionName] || '';
    const line = !info
      ? `There is no information on ${actionName}`
      : info.help || info.description;
    Utils.alertIt(line, actionName);
    return finished(`Here's your help card on ${actionName}`, tictoc);
  }

  export function createToggleButtons(actionName: string) {
    const tictoc = Shlog.tic('getHelp');
    Settings.updateSettingsForUser({ ...g.UserSettings, helpControl: 'off' });
    const info = Help[actionName] || '';
    const line = !info
      ? `There is no information on ${actionName}`
      : info.help || info.description;
    Utils.alertIt(line, actionName);
    return finished(`Here's your help card on ${actionName}`, tictoc);
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

  /**
   *
   * @param {GoogleAppsScript.Addons.EventObject} e
   */
  export function formatAsCaption(e: { parameters: { location: any } }) {
    const tictoc = Shlog.tic('formatAsCaption');
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

    return finished(`Successfully formatted caption`, tictoc);
  }

  export function showFormula() {
    const tictoc = Shlog.tic('showFormula');
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

    return finished(`Successfully show formula`, tictoc);
  }

  /**
   *
   * @param {GoogleAppsScript.Addons.EventObject} e
   */
  export function makeRangeStatic(e: { parameters: { noNotes: any } }) {
    const tictoc = Shlog.tic('makeRangeStatic');
    const { noNotes } = e.parameters;
    makeRangeStatic_(g.ActiveRange, noNotes);
    return finished(`Successfully made range static`, tictoc);
  }

  export function makeRangeDynamic() {
    const tictoc = Shlog.tic('makeRangeDynamic');
    makeRangeDynamic_(g.ActiveRange);
    return (
      Shlog.toc(tictoc) &&
      CardService.newActionResponseBuilder()
        .setNotification(
          CardService.newNotification().setText(
            `Successfully restored range formulas`,
          ),
        )
        .build()
    );
  }

  function finished(
    line: string,
    tt?: number,
    e?: GoogleAppsScript.Addons.EventObject,
  ) {
    return Shlog.toc(tt) && e
      ? CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification().setText(line))
          .build()
      : g.ss.toast(line);
  }

  /**
   *
   * @param {GoogleAppsScript.Addons.EventObject} e
   */
  export function makeSheetsStatic(e: {
    parameters: { use: any; noNotes: any };
  }) {
    const tictoc = Shlog.tic('makeSheetsStatic');
    const { use, noNotes } = e.parameters;
    const sheets =
      use === 'all'
        ? g.ss.getSheets()
        : use === 'list'
          ? g.ActiveRange.getValues().flat()
          : use === 'current'
            ? [g.ActiveSheet]
            : [];
    sheets.forEach((sheet) => makeSheetStatic_(sheet, noNotes));
    return finished(`Successfully made formulas static`, tictoc);
  }

  /**
   *
   * @param {GoogleAppsScript.Addons.EventObject} e
   */
  export function makeSheetsDynamic(e: { parameters: { use: any } }) {
    const tictoc = Shlog.tic('makeSheetsDynamic');
    const { use } = e.parameters;
    const sheets =
      use === 'all'
        ? g.ss.getSheets()
        : use === 'list'
          ? g.ActiveRange.getValues().flat()
          : use === 'current'
            ? [g.ActiveSheet]
            : [];
    sheets.forEach((sheet) => makeSheetDynamic_(sheet));
    return finished(`Successfully restored formulas`, tictoc);
  }

  /** Makes sheet static by placing content of either hard-coded values or functions into notes.
   * @param {GoogleAppsScript.Spreadsheet.Sheet|string} [sheet] - Optional sheet object or name of sheet to make static (defaults to active sheet)
   * @param {boolean} noNotes - When true, function will not place values in notes
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   */
  function makeSheetStatic_(
    sheet?: GoogleAppsScript.Spreadsheet.Sheet | string,
    noNotes: boolean = false,
  ): GoogleAppsScript.Spreadsheet.Sheet {
    const tictoc = Shlog.tic('makeSheetStatic_');
    const s = sheet
      ? typeof sheet === 'string'
        ? g.ss.getSheetByName(sheet)
        : sheet
      : g.ActiveSheet;

    const range = s.getDataRange();

    const response = makeRangeStatic_(range, noNotes);

    if (response !== 'failed' && !noNotes) s.setTabColor('#E92D18');

    return Shlog.toc(tictoc) && s;
  }

  /** Places all notes into cells, whether they be formulas or values
   * @param {GoogleAppsScript.Spreadsheet.Sheet|string=} [sheet] - Optional sheet object or name of sheet to make dynamic (defaults to active sheet)
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   */
  function makeSheetDynamic_(
    sheet?: GoogleAppsScript.Spreadsheet.Sheet | string,
  ): GoogleAppsScript.Spreadsheet.Sheet {
    const tictoc = Shlog.tic('makeSheetDynamic_');
    const s = sheet
      ? typeof sheet === 'string'
        ? g.ss.getSheetByName(sheet)
        : sheet
      : g.ActiveSheet;

    const range = s.getRange(1, 1, s.getLastRow(), s.getLastColumn());

    const response = makeRangeDynamic_(range);

    // change tab color to indicate sheet is dynamic
    if (response !== 'failed') s.setTabColor('#249A41');

    return Shlog.toc(tictoc) && s;
  }

  /** Make range notes the user entered value/formula, but avoid shadow values
   * @param {GoogleAppsScript.Spreadsheet.Range} [range] - range to make static, otherwise uses active range
   * @returns {GoogleAppsScript.Spreadsheet.Range|string} returns range for chaining
   */
  function makeRangeStatic_(
    range?: GoogleAppsScript.Spreadsheet.Range,
    noNotes: boolean = false,
  ): GoogleAppsScript.Spreadsheet.Range | string {
    const tictoc = Shlog.tic('makeRangeStatic_');
    const r = range || g.ActiveRange;
    let status;

    if (noNotes) {
      return r.setValues(r.getValues());
    }

    // save original values for the end
    const originalValues = r
      .getValues()
      .map((r) => r.map((c) => (c.toString().startsWith('=') ? `'${c}` : c)));

    // get range formulas, where no formula means ''
    const formulas = r.getFormulas();

    if (new Set(r.getNotes().flat()).size > 1) {
      const confirmation = Utils.askQuestion(
        'Warning - appears to already be static',
        'Are you sure you want to continue?',
      );

      if (!confirmation) {
        Utils.displayError('Process cancelled!');
        status = 'failed';
      }
    }

    if (status !== 'failed') {
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
    }
    return (Shlog.toc(tictoc) && status) || r;
  }

  /** Replaces range values with value/formula stored in cell notes.
   * @param {GoogleAppsScript.Spreadsheet.Range} [range] - range to make dynamic, otherwise uses active range
   * @returns {GoogleAppsScript.Spreadsheet.Range|string} returns range for chaining
   */
  function makeRangeDynamic_(
    range?: GoogleAppsScript.Spreadsheet.Range,
  ): GoogleAppsScript.Spreadsheet.Range | string {
    const tictoc = Shlog.tic('makeRangeDynamic_');
    const r = range || g.ActiveRange;
    let status;

    if (new Set(r.getFormulas().flat()).size > 1) {
      const confirmation = Utils.askQuestion(
        'Warning - appears to already be dynamic',
        'Are you sure you want to continue?',
      );

      if (!confirmation) {
        Utils.displayError('Process cancelled!');
        status = 'failed';
      }
    }

    if (status !== 'failed')
      r.clearContent().setValues(r.getNotes()).setNote(null);

    return (Shlog.toc(tictoc) && status) || r;
  }

  export function createNamedRangesFromSheet(): GoogleAppsScript.Card_Service.ActionResponse {
    const tictoc = Shlog.tic('createNamedRangesFromSheet');
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
    return (
      Shlog.toc(tictoc) &&
      CardService.newActionResponseBuilder()
        .setNotification(
          CardService.newNotification().setText(
            `Successfully created named ranges.`,
          ),
        )
        .build()
    );
  }

  export function createNamedRangesFromSelection(): GoogleAppsScript.Card_Service.ActionResponse {
    const tictoc = Shlog.tic('createNamedRangesFromSheet');
    const ss = g.ss;
    const sheetName = g.ActiveSheet.getName();
    const dataRange = g.ActiveRange;
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
    return (
      Shlog.toc(tictoc) &&
      CardService.newActionResponseBuilder()
        .setNotification(
          CardService.newNotification().setText(
            `Successfully created named ranges.`,
          ),
        )
        .build()
    );
  }

  /**
   * Displays the GraphQL query editor for a cell
   */
  export function showFunctionBuilder() {
    const tictoc = Shlog.tic('showFunctionBuilder');
    const ui = HtmlService.createHtmlOutputFromFile('query-editor')
      .setWidth(600)
      .setHeight(400)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(ui, 'Query Editor');
    return Shlog.toc(tictoc);
  }

  export function insertCommentBlock() {
    g.ActiveRange.merge()
      .setBackground('#FFF2CC')
      .setBorder(
        true,
        true,
        true,
        true,
        null,
        null,
        '#0c343d',
        SpreadsheetApp.BorderStyle.SOLID,
      )
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontColor('#134f5c');
  }

  export function requestRandomData(e) {
    const tictoc = Shlog.tic('requestRandomData');
    const randomOptions: string[] =
      e.commonEventObject.formInputs?.randomOptions?.stringInputs?.value || [];

    const randomNumberOptions: string[] =
      e.commonEventObject.formInputs?.randomOptions?.stringInputs?.value || [];

    const nRandom = parseInt(
      e.commonEventObject.formInputs.nRandomData.stringInputs.value[0],
    );
    let data;
    const fields = ['id', 'name', 'dob', 'location', 'phone', 'email'];

    const fieldBool = [
      randomOptions.includes('randomId'),
      randomOptions.includes('randomName'),
      randomOptions.includes('randomDate'),
      randomOptions.includes('randomAddress'),
      randomOptions.includes('randomPhone'),
      randomOptions.includes('randomEmail'),
    ];

    const requestedFields = fields.filter((f, i) => fieldBool[i]);

    const response = UrlFetchApp.fetch(
      `https://randomuser.me/api/?results=${nRandom}&nat=us&inc=${requestedFields.join(',')}`,
    );
    const res = JSON.parse(response.getContentText());
    data = [requestedFields];
    res.results.forEach((r) => {
      data.push(
        requestedFields.map((f) => {
          let value;
          switch (f) {
            case 'id':
              value = r[f].value;
              break;
            case 'name':
              value = r[f].first + ' ' + r[f].last;
              break;
            case 'dob':
              value = new Date(r[f].date);
              break;
            case 'location':
              value = `${r[f].city}, ${r[f].state}`;
              break;
            case 'phone':
              value = r[f];
              break;
            case 'email':
              value = r[f];
              break;

            default:
              break;
          }
          return value;
        }),
      );
    });

    g.ActiveRange.offset(0, 0, data.length, data[0].length).setValues(data);

    return finished(`Successful completed formatTable`, tictoc);
  }

  export function requestRandomNumbers(e) {
    const tictoc = Shlog.tic('requestRandomNumbers');

    const randomNumberOptions: string[] =
      e.commonEventObject.formInputs?.randomNumberOptions?.stringInputs
        ?.value || [];

    const nRandom = parseInt(
      e.commonEventObject.formInputs.nRandomNumbers.stringInputs.value[0],
    );
    let data;
    if (randomNumberOptions.length) {
      data = Array.from({ length: nRandom }, () =>
        randomNumberOptions.includes('randomNumbers')
          ? [Math.random() * 100]
          : [Math.floor(Math.random() * 100) + 1],
      );
    }

    g.ActiveRange.offset(0, 0, data.length, data[0].length).setValues(data);

    return finished(`Successful completed formatTable`, tictoc);
  }

  export function playground1(e) {
    const tictoc = Shlog.tic('playground1');
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

    console.log(workbook.getDataAsString());
    // Parse XLSX data and retrieve the named functions.
    const root = XmlService.parse(workbook.getDataAsString()).getRootElement();
    console.log('root :>> ', XmlService.getPrettyFormat().format(root));
    Logger.log('root :>> ', XmlService.getPrettyFormat().format(root));
    let sheetName = 'Event Object';
    let sheet = g.ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = g.ss.insertSheet(sheetName);
    } else {
      // sheet.getDataRange().clear({ contentsOnly: true });
    }
    const data = ImportJSON.parseJSONObject(root);

    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    return showResults(JSON.stringify(e), tictoc, e);
  }

  export function playground2() {
    const tictoc = Shlog.tic('playground2');
    g.ss.getNamedRanges().forEach((namedRange) => {
      const name = namedRange.getName();

      if (name[0] == "'") {
        const rangeA1n = namedRange.getRange().getA1Notation();
        const sheetName = namedRange.getRange().getSheet().getName();
        const cleanName = name.replace(/^.+!/, '');
        namedRange.remove();
        g.ss.setNamedRange(
          cleanName,
          g.ss.getRange(`'${sheetName}'!${rangeA1n}`),
        );
      }
    });

    return finished(`Playground2 complete`, tictoc);
  }

  export function playground3(e) {
    const tictoc = Shlog.tic('playground3');
    showSettings(e);
    return showResults(JSON.stringify(e), tictoc, e);
  }

  export function parseJSONFromSheet() {
    const tictoc = Shlog.tic('parseJSONFromSheet');
    const sourceSheetName = g.ActiveSheet.getName();
    let sheetName = 'JSON_Output';
    let sheet = g.ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = g.ss.insertSheet(sheetName);
    } else {
      // sheet.getDataRange().clear({ contentsOnly: true });
    }
    const data = ImportJSON.ImportJSONFromSheet(
      sourceSheetName,
      '',
      'noTruncate',
    );

    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    return finished(`parseJSONFromSheet complete`, tictoc);
  }

  export function parseJSONFromRange() {
    const tictoc = Shlog.tic('parseJSONFromRange');
    let sheetName = 'JSON_Output';
    let sheet = g.ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = g.ss.insertSheet(sheetName);
    } else {
      // sheet.getDataRange().clear({ contentsOnly: true });
    }
    const data = ImportJSON.ImportJSONFromSheet('', '', 'noTruncate');

    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    return finished(`parseJSONFromRange complete`, tictoc);
  }

  function showResults(
    line: string,
    tt?: number,
    e?: GoogleAppsScript.Addons.EventObject,
  ) {
    if (!e) {
      const ui = HtmlService.createHtmlOutput(`<p>${line}</p>`)
        .setWidth(600)
        .setHeight(400);

      return Shlog.toc(tt) && g.ss.show(ui);
    }
    return (
      Shlog.toc(tt) &&
      CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(line))
        .build()
    );
  }
}

export { ActionHandlers };
