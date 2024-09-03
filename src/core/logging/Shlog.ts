// @ts-ignore
import { Settings } from '@core/environment/Settings';
import { type Dayjs } from 'dayjs';

export interface TicToc {
  id?: number;
  start?: Dayjs;
  stop?: Dayjs;
  depth?: number;
  name?: string;
  key?: string;
  parameters?: string;
  result?: string;
  message?: string;
  error?: string;
}

namespace Shlog {
  const LIMIT_WIDTH = 80;
  const LIMIT_LINES = 20;
  const LIMIT_ROWS = 1000;
  let SHEET_NAME = 'app_log';
  let LOG: GoogleAppsScript.Spreadsheet.Sheet;
  let ID: number;
  let DEPTH: number = 0;
  const TICTOCS: TicToc[] = [];

  export function init(sheetName: string = SHEET_NAME) {
    const start = daygs();
    ID = start.unix();
    TICTOCS.push({
      id: start.valueOf(),
      start,
      name: 'Pre-execution',
      depth: 0,
    });

    SHEET_NAME = sheetName;
  }

  /** Crops string output into rectangular shape using height and width option parameters
   *
   * @param str {string}
   * @returns {string}
   */
  function cropString_(str: string): string {
    return str
      .split(/\r?\n|\r|\n/g)
      .slice(0, LIMIT_LINES)
      .map((line) => line.slice(0, LIMIT_WIDTH))
      .join('\n');
  }

  export function tic(
    name: string,
    options: {
      key?: string;
      message?: string;
      parameters?: any;
      result?: any;
      error?: string;
    } = {},
  ): number {
    const settings = Settings.getSettingsForUser();
    // console.log('settings :>> ', settings);
    DEPTH++;
    const start: Dayjs = daygs();
    const tictoc = {
      id: start.valueOf(),
      start,
      depth: DEPTH,
      name,
      key: options.key || '',
      parameters: options.parameters
        ? Object.entries(options.parameters)
            .map(([key, value]) => {
              return `${key}: ${JSON.stringify(value, null, 2)}`;
            })
            .join('\n')
        : '',
      message: options.message,
      result: '',
      error: '',
    };
    if (TICTOCS.length === 1) TICTOCS[0].stop = start;

    TICTOCS.push(tictoc);

    printTictocs_(`tic - ${tictoc.name}`);
    return tictoc.id;
  }

  function getSheet_() {
    LOG = g.ss.getSheetByName(SHEET_NAME);

    if (!LOG) {
      LOG = g.ss.insertSheet(SHEET_NAME).setHiddenGridlines(true);
      const headers = [
        'ID',
        'Started',
        'Duration',
        'Function',
        'Arguments',
        'Message',
        'Result',
        'Error',
        `=MAP(A:A,B:B,C:C,LAMBDA(id,start,duration,LET(i,ROW(id),starttimes,filter(B:B,A:A=id),nColor,MOD(COUNTUNIQUE(OFFSET($A$1,0,0,i)),12)+1,IF(id="",,IF(i=1,"Duration Chart",SPARKLINE(({86400*(start-VLOOKUP(id,$A:$B,2,0)),duration}),({"charttype","bar";"max",86400*(MAX(starttimes)-MIN(starttimes))+FILTER(C:C,A:A=id,D:D="Post Execution");"color1",INDEX({"#DEDBF0";"#EDF0DB";"#E2F0DB";"#E9DBF0";"#F0DBED";"#DBF0DE";"#DBF0E9";"#F0DBE2";"#F0DEDB";"#DBEDF0";"#DBE2F0";"#F0E9DB"},nColor);"color2",INDEX({"#4D3F9D";"#8F9D3F";"#609D3F";"#7C3F9D";"#9D3F8F";"#3F9D4D";"#3F9D7C";"#9D3F60";"#9D4D3F";"#3F8F9D";"#3F609D";"#9D7C3F"},nColor)})))))))`,
      ];

      LOG.setColumnWidth(1, 80)
        .setColumnWidth(2, 115)
        .setColumnWidth(3, 70)
        .setColumnWidth(4, 285)
        .setColumnWidths(5, 4, 200)
        .setColumnWidth(9, 250)
        .getRange(1, 1, LIMIT_ROWS, headers.length)
        .setFontFamily('Lato')
        .offset(0, 0, 1)
        .setBackground('#38761d')
        .setFontColor('#ffffff')
        .setHorizontalAlignment('center')
        .setValues([headers]);

      LOG.deleteColumns(
        headers.length + 1,
        LOG.getMaxColumns() - headers.length,
      );

      LOG.setFrozenRows(1);
      LOG.getRange('A:H').setFontSize(10);
      LOG.getRange('A2:C').setFontSize(8);
      var sheetRange = LOG.getRange(
        1,
        1,
        LOG.getMaxRows(),
        LOG.getMaxColumns(),
      );

      sheetRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

      LOG.getRange('B:B').setNumberFormat('YYYY-MM-DD HH:mm:ss.SSS');
      LOG.getRange('C:C').setNumberFormat('0.00 "sec"');

      var conditionalFormatRules = LOG.getConditionalFormatRules();

      conditionalFormatRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .setRanges([LOG.getRange('A1:H1000')])
          .whenFormulaSatisfied('=$D1="Pre-execution"')
          .setBackground('#B6D7A8')
          .build(),
      );

      conditionalFormatRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .setRanges([LOG.getRange('A1:H1000')])
          .whenFormulaSatisfied('=$D1="Post Execution"')
          .setBackground('#E2EFD9')
          .build(),
      );

      LOG.getRange('1:1').setBackground('#38761d').setFontColor('#ffffff');
      LOG.setConditionalFormatRules(conditionalFormatRules);
      LOG.getRange('A:I')
        .setBorder(
          null,
          null,
          null,
          true,
          null,
          null,
          '#38761d',
          SpreadsheetApp.BorderStyle.SOLID,
        )
        .setBorder(
          null,
          true,
          null,
          null,
          null,
          null,
          '#38761d',
          SpreadsheetApp.BorderStyle.SOLID,
        )
        .setBorder(
          null,
          null,
          null,
          null,
          true,
          null,
          '#6aa84f',
          SpreadsheetApp.BorderStyle.SOLID,
        );
      LOG.getRange('A2:H1000').setBorder(
        null,
        null,
        null,
        null,
        null,
        true,
        '#6aa84f',
        SpreadsheetApp.BorderStyle.DASHED,
      );

      LOG.setRowHeightsForced(1, 1000, 21);
      LOG.hideSheet();
    }
  }

  function diffMs_(start: Dayjs, finish?: Dayjs) {
    return daygs(finish || undefined).diff(start, 'millisecond') / 1000;
  }

  function saveTicTocs_(tt: TicToc) {
    const runStop = tt.stop;
    const newRows = [];
    TICTOCS.forEach((tictoc) => {
      const { start, stop, name, parameters, depth, result, message, error } =
        tictoc;
      const methodName = ' ⇨ '.repeat(depth || 0) + name;

      newRows.push([
        ID,
        start.format('YYYY-MM-DD HH:mm:ss.SSS'),
        stop ? diffMs_(start, stop) : null,
        methodName,
        parameters,
        message,
        result,
        error,
      ]);
    });

    let store = PropertiesService.getDocumentProperties();
    const tt_rows = store.getKeys().includes('tt_rows')
      ? (JSON.parse(store.getProperty('tt_rows')) as any[][])
      : [];

    const willExceedLimit = newRows.length + tt_rows.length + 1 > LIMIT_ROWS;
    if (willExceedLimit) outputTictocs();

    store.setProperty(
      'tt_rows',
      JSON.stringify(
        [
          ...newRows,
          [
            ID,
            runStop.format('YYYY-MM-DD HH:mm:ss.SSS'),
            diffMs_(runStop),
            'Post Execution',
            ...Array(4).fill(null),
          ],
        ].concat(willExceedLimit ? [] : tt_rows),
      ),
    );

    if (tt.error) throw new Error(tt.error);
  }

  export function outputTictocs() {
    let store = PropertiesService.getDocumentProperties();
    const tt_rows = JSON.parse(store.getProperty('tt_rows')) as any[][];

    if (!tt_rows || tt_rows.length === 0) return;

    getSheet_();

    const newLogsCount = tt_rows.length;
    const [headers, ...rows] = LOG.getRange(1, 1, LIMIT_ROWS, 8).getValues();
    const new_rows = tt_rows.concat(rows);

    LOG.getRange(1, 1, LIMIT_ROWS, 8).setValues([
      headers,
      ...new_rows.slice(0, LIMIT_ROWS - 1),
    ]);

    store.deleteProperty('tt_rows');

    return newLogsCount;
  }

  export function toc(
    id: number | boolean,
    options: {
      label?: string;
      message?: string;
      parameters?: any;
      result?: any;
      error?: string;
    } = {},
  ) {
    if (!id) return true;
    DEPTH--;
    const tictocIndex = TICTOCS.findIndex((tt) => tt.id && tt.id === id);
    const tictoc = TICTOCS[tictocIndex];
    tictoc.stop = daygs();
    tictoc.result = options.result
      ? cropString_(JSON.stringify(options.result, null, 2))
      : '';
    tictoc.error = options.error;
    tictoc.message = [tictoc.message, options.message].join('\n').trim();
    const message = [
      ...Array(tictoc.depth).fill('⇨'),
      tictoc.name.split('.').slice(-1)[0],
    ].join(' ');

    printTictocs_(`toc - ${tictoc.name}`);
    if (DEPTH === 0 || !!options.error) saveTicTocs_(tictoc);

    return options.result || true;
  }

  function secStr_(startD: Dayjs, stopD: Dayjs) {
    // @ts-ignore
    return daygs.duration(stopD.diff(startD)).format('s.SSS');
  }

  function printTictocs_(title: string) {
    const sep = `|${[7, 36, 14, 14, 13].map((w) => '-'.repeat(w)).join('|')}|`;

    const lines = [
      [`OVERALL DEPTH: ${DEPTH}`],
      [
        'Depth',
        '              Method              ',
        '    Start   ',
        '    Stop    ',
        ' Duration  ',
      ],
    ]
      .concat(
        TICTOCS.map(({ name, depth, start, stop }) => {
          const depthStr = `     ${depth} `.slice(-5);
          const nameStr =
            name.length > 31
              ? `${name.slice(0, 31)}...`
              : `${name}${' '.repeat(34)}`.slice(0, 34);
          const startStr = start.format('HH:mm:ss.SSS');
          const stopStr = stop ? stop.format('HH:mm:ss.SSS') : '            ';
          const duration = stop
            ? `      ${secStr_(start, stop)} sec`.slice(-11)
            : ' running...';
          return [depthStr, nameStr, startStr, stopStr, duration];
        }),
      )
      .map((row) => `| ${row.join(' | ')} |`);

    lines.splice(2, 0, sep);

    console.log(lines.join('\n'));
  }

  export function snap(message: string) {
    console.log('snap :>> ', message);
    TICTOCS.push({
      name: 'snap',
      depth: DEPTH,
      start: daygs(),
      stop: daygs(),
      message,
    });
  }

  export function tap(
    title: string,
    options: {
      message?: string;
      parameters?: any;
      result?: any;
      error?: string;
    } = {},
  ) {
    const stop = daygs();
    const tictoc: TicToc = {
      id: stop.valueOf(),
      start: stop,
      stop,
      depth: DEPTH,
      name: `TAP: ${title}`,
      message: options.message,
      parameters: options.parameters
        ? JSON.stringify(options.parameters, null, 2)
        : '',
      result: options.result ? JSON.stringify(options.result, null, 2) : '',
      error: options.error,
    };
    TICTOCS.push(tictoc);

    return true;
  }
}

export { Shlog };
