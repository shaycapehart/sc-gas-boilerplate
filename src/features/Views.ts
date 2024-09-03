import { COLORS } from '@core/lib/COLORS';
import { Shlog } from '@core/logging/Shlog';
import { ErrorCardOptions, SettingsOptions } from '@core/types/addon';

namespace Views {
  const DIVIDER = CardService.newDivider();

  /**
   * A widgets object containing versatile action widgets for every action needed for Workspace Addon.
   * Eqch widget will generally have at a minimum: title, description, and action. Most will also have an iconUrl
   * for the icon that can either be used for icon buttons or as a begger icon decorated text. Other properties
   * include parameters for the action and details about switch toggle.
   */
  const widgets = {
    emptyLogRecordsWidget: createActionWidget_({
      title: 'Empty logs',
      description: 'Empty cached log records',
      iconUrl: buildIconUrl_('bug_report'),
      action: 'emptyLogRecords',
    }),
    playground1Widget: createActionWidget_({
      title: 'Playground 1',
      description: 'Run playground 1 function',
      iconUrl: buildIconUrl_('looks_one'),
      action: 'playground1',
    }),
    playground2Widget: createActionWidget_({
      title: 'Playground 2',
      description: 'Run playground 2 function',
      iconUrl: buildIconUrl_('looks_two'),
      action: 'playground2',
    }),
    playground3Widget: createActionWidget_({
      title: 'Playground 3',
      description: 'Run playground 3 function',
      iconUrl: buildIconUrl_('looks_3'),
      action: 'playground3',
    }),
    helpWidget: createActionWidget_({
      title: 'Click here and then any command',
      description:
        'Get help on any button by first clicking this icon and then click the other button to see a description.',
      iconUrl: buildIconUrl_('help_outline'),
      action: 'getHelp',
    }),
    cropToSelectionWidget: createActionWidget_({
      title: 'Crop to Selection',
      description: 'Crop sheet rows and columns down to selected range',
      iconUrl: buildIconUrl_('crop_square'),
      action: 'cropToSelection',
    }),
    cropToDataWidget: createActionWidget_({
      title: 'Crop to Data',
      description: 'Crop sheet rows and columns down to data range',
      iconUrl: buildIconUrl_('crop'),
      action: 'cropToData',
    }),
    setBackgroundsUsingCellValuesWidget: createActionWidget_({
      title: 'Set Backgrounds Using Values',
      description: 'Set background color using cell value',
      iconUrl: buildIconUrl_('brush'),
      action: 'setBackgroundsColorsUsingValues',
    }),
    getBackgroundColorToValuesWidget: createActionWidget_({
      title: 'Get Background Colors To Values',
      description: 'Put HEX value in cells',
      iconUrl: buildIconUrl_('colorize'),
      action: 'getBackgroundColorsToValues',
    }),
    getBackgroundColorToNotesWidget: createActionWidget_({
      title: 'Get Background Colors To Notes',
      description: 'Put HEX and RGB info in notes',
      iconUrl: buildIconUrl_('palette'),
      action: 'getBackgroundColorsToNotes',
    }),
    matchFontToBackgroundWidget: createActionWidget_({
      title: 'Match font to background',
      description: 'Make font color nearly the same as background color',
      iconUrl: buildIconUrl_('font_download'),
      action: 'matchFontToBackground',
    }),
    makeFontNiceColorWidget: createActionWidget_({
      title: 'Make font a nice color',
      description: 'Make font color nice compared to background',
      iconUrl: buildIconUrl_('font_download'),
      action: 'makeFontNiceColor',
    }),
    squareSelectedCellsWidget: createActionWidget_({
      title: 'Square the selected cells',
      description: 'Give same width/height to selected cells',
      iconUrl: buildIconUrl_('view_module'),
      action: 'squareSelectedCells',
    }),
    createTogglesWidget: createActionWidget_({
      title: 'Create toggles',
      description: 'Choose a color and turn selected range into toggles',
      iconUrl: buildIconUrl_('check_box'),
      action: 'createToggles',
    }),
    createSheetOfTogglesWidget: createActionWidget_({
      title: 'Create toggles sheet',
      description: 'Create a sheet of colored toggle checkboxes',
      iconUrl: buildIconUrl_('apps'),
      action: 'createSheetWithToggleButtons',
    }),
    createStatsDashboardWidget: createActionWidget_({
      title: 'Create Stats Dashboard',
      description: 'Get sheet utilization and formulas',
      iconUrl: buildIconUrl_('poll'),
      action: 'createStatsDashboard',
    }),
    createNamedRangesDashboardWidget: createActionWidget_({
      title: 'Create Named Ranges Dashboard',
      description: 'Refresh/build named ranges sheet',
      iconUrl: buildIconUrl_('list_alt'),
      action: 'createNamedRangesDashboard',
    }),
    updateNamedRangesDashboardWidget: createActionWidget_({
      title: 'Update Named Ranges Dashboard',
      description: 'Update named ranges',
      iconUrl: buildIconUrl_('update'),
      action: 'updateNamedRangeSheet',
    }),
    showNamedFunctionsDashboardWidget: createActionWidget_({
      title: 'Create Named Functions Dashboard',
      description: 'Show/create named functions dashboard',
      iconUrl: buildIconUrl_('loyalty'),
      action: 'showNamedFunctionsDashboard',
    }),
    createNamedRangesFromSheetWidget: createActionWidget_({
      title: 'Create NR Using Current Sheet',
      description: 'sheetName.columnName + sheetName',
      iconUrl: buildIconUrl_('burst_mode'),
      action: 'createNamedRangesFromSheet',
    }),
    createNamedRangesFromSelectionWidget: createActionWidget_({
      title: 'Create NR Using Selection',
      description: 'sheetName.columnName + sheetName',
      iconUrl: buildIconUrl_('burst_mode'),
      action: 'createNamedRangesFromSelection',
    }),
    showFormulaWidget: createActionWidget_({
      title: 'Show Formula Down Below',
      altText: '',
      description: 'Shows formula in cell directly below or two below',
      iconUrl: buildIconUrl_('functions'),
      action: 'showFormula',
    }),
    makeRangeStaticWidget: createActionWidget_({
      title: 'Make range static',
      description: 'Convert selected formulas to values, save to notes',
      iconUrl:
        'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjAgMjAiPjxwYXRoIGZpbGw9ImN1cnJlbnRDb2xvciIgZD0iTTE1LjUgOWMuNTQ2IDAgMS4wNTkuMTQ2IDEuNS40MDFWOGgtNHYyLjM0MUEzIDMgMCAwIDEgMTUuNSA5TTExIDE0YzAtLjM2NC4wOTctLjcwNi4yNjgtMUg4djRoM3ptMS0ySDhWOGg0em0tNSAwVjhIM3Y0em0tNCAxaDR2NEg1LjVBMi41IDIuNSAwIDAgMSAzIDE0LjV6bTEwLTZoNFY1LjVBMi41IDIuNSAwIDAgMCAxNC41IDNIMTN6bS0xLTR2NEg4VjN6TTcgM3Y0SDNWNS41QTIuNSAyLjUgMCAwIDEgNS41IDN6bTYuNSA5djFIMTNhMSAxIDAgMCAwLTEgMXY0YTEgMSAwIDAgMCAxIDFoNWExIDEgMCAwIDAgMS0xdi00YTEgMSAwIDAgMC0xLTFoLS41di0xYTIgMiAwIDEgMC00IDBtMSAxdi0xYTEgMSAwIDEgMSAyIDB2MXptMSAyLjI1YS43NS43NSAwIDEgMSAwIDEuNWEuNzUuNzUgMCAwIDEgMC0xLjUiLz48L3N2Zz4=',
      action: 'makeRangeStatic',
    }),
    makeRangeDynamicWidget: createActionWidget_({
      title: 'Make range dynamic',
      iconUrl: buildIconUrl_('lock_open'),
      description: 'Restore selected note formulas',
      action: 'makeRangeDynamic',
    }),
    makeSheetStaticWidget: createActionWidget_({
      title: 'Make sheet static',
      description: 'Convert all formulas to values, save to notes',
      iconUrl:
        'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMzIgMzIiPjxwYXRoIGZpbGw9ImN1cnJlbnRDb2xvciIgZD0iTTE3LjA4NiA0LjQwMWMtMS41ODgtMS4yNTYtMy45MzUtLjIyOC00LjA4NyAxLjc5MWwtLjIxIDIuODA0SDE3YTEgMSAwIDEgMSAwIDJoLTQuMzYyTDExLjQ4NiAyNi4zMWMtLjI4OCAzLjgzMS00LjkxOCA1LjU4My03LjY3IDIuOTAzbC0uNTE0LS41YTEgMSAwIDAgMSAxLjM5Ni0xLjQzNGwuNTEzLjVjMS41MzYgMS40OTcgNC4xMi41MiA0LjI4MS0xLjYybDEuMTQtMTUuMTYzSDhhMSAxIDAgMSAxIDAtMmgyLjc4M2wuMjIyLTIuOTU0Yy4yNzItMy42MTggNC40NzctNS40NiA3LjMyMi0zLjIwOWwuNzk0LjYyOWExIDEgMCAxIDEtMS4yNDIgMS41Njh6bS0uODggMTEuODdhMS4xMTUgMS4xMTUgMCAwIDEgMS43NS4zOGwxLjg2MiA0LjExM2wtNS41MjUgNS41MjVhMSAxIDAgMCAwIDEuNDE0IDEuNDE0bDQuOTkyLTQuOTkybDEuNTc0IDMuNDc2Yy45MDUgMS45OTggMy41MzUgMi40NiA1LjA2Ny44OTJsLjM3NS0uMzg1YTEgMSAwIDEgMC0xLjQzLTEuMzk3bC0uMzc2LjM4NGExLjExNSAxLjExNSAwIDAgMS0xLjgxNC0uMzE5TDIyLjIxIDIxLjJsNS40OTctNS40OTdhMSAxIDAgMCAwLTEuNDE0LTEuNDE0bC00Ljk2NCA0Ljk2NGwtMS41NTItMy40MjdhMy4xMTUgMy4xMTUgMCAwIDAtNC44ODctMS4wNjJsLS41NDguNDc5YTEgMSAwIDAgMCAxLjMxNiAxLjUwNnoiLz48L3N2Zz4=',
      action: 'makeSheetsStatic',
      parameters: { use: 'current' },
    }),
    makeSheetDynamicWidget: createActionWidget_({
      title: 'Make sheet dynamic',
      description: 'Restore all note formulas',
      iconUrl: buildIconUrl_('lock_open'),
      action: 'makeSheetsDynamic',
      parameters: { use: 'current' },
    }),
    makeAllSheetsStaticWidget: createActionWidget_({
      title: 'Make all sheets static',
      description:
        'Convert all formulas on all sheets to values, save in notes',
      iconUrl:
        'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMTYgMTYiPjxnIGZpbGw9ImN1cnJlbnRDb2xvciI+PHBhdGggZD0iTTEzIDUuNjk4YTUgNSAwIDAgMS0uOTA0LjUyNUMxMS4wMjIgNi43MTEgOS41NzMgNyA4IDdzLTMuMDIyLS4yODktNC4wOTYtLjc3N0E1IDUgMCAwIDEgMyA1LjY5OFY3YzAgLjM3NC4zNTYuODc1IDEuMzE4IDEuMzEzQzUuMjM0IDguNzI5IDYuNTM2IDkgOCA5Yy42NjYgMCAxLjI5OC0uMDU2IDEuODc2LS4xNTZjLS40My4zMS0uODA0LjY5My0xLjEwMiAxLjEzMkExMiAxMiAwIDAgMSA4IDEwYy0xLjU3MyAwLTMuMDIyLS4yODktNC4wOTYtLjc3N0E1IDUgMCAwIDEgMyA4LjY5OFYxMGMwIC4zNzQuMzU2Ljg3NSAxLjMxOCAxLjMxM0M1LjIzNCAxMS43MjkgNi41MzYgMTIgOCAxMmguMDI3YTQuNiA0LjYgMCAwIDAtLjAxNy44QTIgMiAwIDAgMCA4IDEzYy0xLjU3MyAwLTMuMDIyLS4yODktNC4wOTYtLjc3N0E1IDUgMCAwIDEgMyAxMS42OThWMTNjMCAuMzc0LjM1Ni44NzUgMS4zMTggMS4zMTNDNS4yMzQgMTQuNzI5IDYuNTM2IDE1IDggMTVjMCAuMzYzLjA5Ny43MDQuMjY2Ljk5N1E4LjEzNCAxNi4wMDEgOCAxNmMtMS41NzMgMC0zLjAyMi0uMjg5LTQuMDk2LS43NzdDMi44NzUgMTQuNzU1IDIgMTQuMDA3IDIgMTNWNGMwLTEuMDA3Ljg3NS0xLjc1NSAxLjkwNC0yLjIyM0M0Ljk3OCAxLjI4OSA2LjQyNyAxIDggMXMzLjAyMi4yODkgNC4wOTYuNzc3QzEzLjEyNSAyLjI0NSAxNCAyLjk5MyAxNCA0djQuMjU2YTQuNSA0LjUgMCAwIDAtMS43NTMtLjI0OUMxMi43ODcgNy42NTQgMTMgNy4yODkgMTMgN3ptLTguNjgyLTMuMDFDMy4zNTYgMy4xMjQgMyAzLjYyNSAzIDRjMCAuMzc0LjM1Ni44NzUgMS4zMTggMS4zMTNDNS4yMzQgNS43MjkgNi41MzYgNiA4IDZzMi43NjYtLjI3IDMuNjgyLS42ODdDMTIuNjQ0IDQuODc1IDEzIDQuMzczIDEzIDRjMC0uMzc0LS4zNTYtLjg3NS0xLjMxOC0xLjMxM0MxMC43NjYgMi4yNzEgOS40NjQgMiA4IDJzLTIuNzY2LjI3LTMuNjgyLjY4N1oiLz48cGF0aCBkPSJNOSAxM2ExIDEgMCAwIDEgMS0xdi0xYTIgMiAwIDEgMSA0IDB2MWExIDEgMCAwIDEgMSAxdjJhMSAxIDAgMCAxLTEgMWgtNGExIDEgMCAwIDEtMS0xem0zLTNhMSAxIDAgMCAwLTEgMXYxaDJ2LTFhMSAxIDAgMCAwLTEtMSIvPjwvZz48L3N2Zz4=',
      action: 'makeSheetsStatic',
      parameters: { use: 'all' },
    }),
    makeAllSheetsDynamicWidget: createActionWidget_({
      title: 'Restore formulas for all sheets',
      description: 'Restore all note formulas on all sheets',
      iconUrl: buildIconUrl_('lock_open'),
      action: 'makeSheetsDynamic',
      parameters: { use: 'all' },
    }),
    makeSheetListStaticWidget: createActionWidget_({
      title: 'Make sheet name list static',
      description:
        'Convert all formulas on selected sheet names to values, save in notes',
      iconUrl:
        'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxwYXRoIGZpbGw9IiNmNjA5MDkiIGQ9Ik0yMS44IDE2di0xLjVjMC0xLjQtMS40LTIuNS0yLjgtMi41cy0yLjggMS4xLTIuOCAyLjVWMTZjLS42IDAtMS4yLjYtMS4yIDEuMnYzLjVjMCAuNy42IDEuMyAxLjIgMS4zaDUuNWMuNyAwIDEuMy0uNiAxLjMtMS4ydi0zLjVjMC0uNy0uNi0xLjMtMS4yLTEuM20tMS4zIDBoLTN2LTEuNWMwLS44LjctMS4zIDEuNS0xLjNzMS41LjUgMS41IDEuM3pNMyAxM3YtMmgxMnYyem0wLTdoMTh2Mkgzem0wIDEydi0yaDZ2MnoiLz48L3N2Zz4=',
      action: 'makeSheetsStatic',
      parameters: { use: 'list' },
    }),
    makeSheetListDynamicWidget: createActionWidget_({
      title: 'Make sheet name list dynamic',
      description: 'Restore all note formulas on selected sheet names',
      iconUrl: buildIconUrl_('lock_open'),
      action: 'makeSheetsDynamic',
      parameters: { use: 'list' },
    }),
    makeRangeStaticNoNotesWidget: createActionWidget_({
      title: 'Make range static, no notes',
      description:
        'Convert all formulas in selected range to values. May affect array formulas above range',
      iconUrl: buildIconUrl_('lock'),
      action: 'makeRangeStatic',
      parameters: { noNotes: 'true' },
    }),
    makeSheetStaticNoNotesWidget: createActionWidget_({
      title: 'Make sheet static, no notes',
      description: 'Convert all formulas to values',
      iconUrl: buildIconUrl_('lock'),
      action: 'makeSheetsStatic',
      parameters: { use: 'current', noNotes: 'true' },
    }),
    makeAllSheetsStaticNoNotesWidget: createActionWidget_({
      title: 'Make all sheets static, no notes',
      description: 'Convert all formulas on all sheets to values',
      iconUrl: buildIconUrl_('lock'),
      action: 'makeSheetsStatic',
      parameters: { use: 'all', noNotes: 'true' },
    }),
    makeSheetListStaticNoNotesWidget: createActionWidget_({
      title: 'Make sheet list static, no notes',
      description: 'Convert all formulas on selected sheet names to values',
      iconUrl: buildIconUrl_('lock'),
      action: 'makeSheetsStatic',
      parameters: { use: 'list', noNotes: 'true' },
    }),
    insertCommentBlockWidget: createActionWidget_({
      title: 'Insert comment block',
      description: 'Merge selected range into a single yellow comment cell',
      iconUrl: buildIconUrl_('comment'),
      action: 'insertCommentBlock',
    }),
    captionTopLeftWidget: createActionWidget_({
      title: '⭶',
      action: 'formatAsCaption',
      parameters: { location: 'top-left' },
    }),
    captionTopRightWidget: createActionWidget_({
      title: '⭷',
      action: 'formatAsCaption',
      parameters: { location: 'top-right' },
    }),
    captionBottomLeftWidget: createActionWidget_({
      title: '⭹',
      action: 'formatAsCaption',
      parameters: { location: 'bottom-left' },
    }),
    captionBottomRightWidget: createActionWidget_({
      title: '⭸',
      action: 'formatAsCaption',
      parameters: { location: 'bottom-right' },
    }),
    captionBottomWidget: createActionWidget_({
      title: '⭣',
      action: 'formatAsCaption',
      parameters: { location: 'bottom-center' },
    }),
    captionBottomBottomWidget: createActionWidget_({
      title: '⮇',
      action: 'formatAsCaption',
      parameters: { location: 'bottombottom' },
    }),
    captionRightWidget: createActionWidget_({
      title: '⭢',
      action: 'formatAsCaption',
      parameters: { location: 'middle-right' },
    }),
    captionLeftWidget: createActionWidget_({
      title: '⭠',
      action: 'formatAsCaption',
      parameters: { location: 'middle-left' },
    }),
    captionTopWidget: createActionWidget_({
      title: '⭡',
      action: 'formatAsCaption',
      parameters: { location: 'top-center' },
    }),
    captionTopTopWidget: createActionWidget_({
      title: '⮅',
      action: 'formatAsCaption',
      parameters: { location: 'toptop' },
    }),
    captionClearWidget: createActionWidget_({
      title: '⭙',
      altText: 'Clear formatting',
      action: 'formatAsCaption',
      parameters: { location: 'clear' },
    }),
    setTableFormatDefaultsWidget: createActionWidget_({
      title: 'Set as default',
      description: 'Sets current options as default.',
      iconUrl: buildIconUrl_('lock'),
      action: 'setTableFormatDefaults',
    }),
    squareSideLengthWidget: createFormElementWidget_({
      title: 'Length of side (pixels)',
      id: 'pixels',
      text: {
        multiLine: false,
      },
    }),
    requestRandomDataWidget: createActionWidget_({
      title: 'Get random data',
      description:
        'Enter the number of rows and what data to insert at selected cell',
      iconUrl: buildIconUrl_('shuffle'),
      action: 'requestRandomData',
    }),
    randomDataRequestedWidget: createFormElementWidget_({
      title: 'How many do you want?',
      id: 'nRandomData',
      text: {
        multiLine: false,
      },
    }),
    requestRandomNumbersWidget: createActionWidget_({
      title: 'Get random numbers',
      description:
        'Enter the number of rows and what data to insert at selected cell',
      iconUrl: buildIconUrl_('shuffle'),
      action: 'requestRandomNumbers',
    }),
    randomNumbersRequestedWidget: createFormElementWidget_({
      title: 'How many do you want?',
      id: 'nRandomNumbers',
      text: {
        multiLine: false,
      },
    }),
    parseJSONFromSheetWidget: createActionWidget_({
      title: 'Parse JSON text',
      description: 'Parse the JSON formatted text in the active sheet',
      iconUrl: `data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIgc3Ryb2tlLXdpZHRoPSIyIj48cGF0aCBkPSJNMTUgMkg2YTIgMiAwIDAgMC0yIDJ2MTZhMiAyIDAgMCAwIDIgMmgxMmEyIDIgMCAwIDAgMi0yVjdaIi8+PHBhdGggZD0iTTE0IDJ2NGEyIDIgMCAwIDAgMiAyaDRtLTEwIDRhMSAxIDAgMCAwLTEgMXYxYTEgMSAwIDAgMS0xIDFhMSAxIDAgMCAxIDEgMXYxYTEgMSAwIDAgMCAxIDFtNCAwYTEgMSAwIDAgMCAxLTF2LTFhMSAxIDAgMCAxIDEtMWExIDEgMCAwIDEtMS0xdi0xYTEgMSAwIDAgMC0xLTEiLz48L2c+PC9zdmc+`,
      action: 'parseJSONFromSheet',
    }),
    parseJSONFromRangeWidget: createActionWidget_({
      title: 'Parse JSON text',
      description: 'Parse the JSON formatted text in the selected range',
      iconUrl: `data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMTYgMTYiPjxwYXRoIGZpbGw9ImJsYWNrIiBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik02IDIuOTg0VjJoLS4wOXEtLjQ3IDAtLjkwOS4xODVhMi4zIDIuMyAwIDAgMC0uNzc1LjUzYTIuMiAyLjIgMCAwIDAtLjQ5My43NTN2LjAwMWEzLjUgMy41IDAgMCAwLS4xOTguODN2LjAwMmE2IDYgMCAwIDAtLjAyNC44NjNxLjAxOC40MzUuMDE4Ljg2OXEwIC4zMDQtLjExNy41NzJ2LjAwMWExLjUgMS41IDAgMCAxLS43NjUuNzg3YTEuNCAxLjQgMCAwIDEtLjU1OC4xMTVIMnYuOTg0aC4wOXEuMjkyIDAgLjU1Ni4xMjFsLjAwMS4wMDFxLjI2Ny4xMTcuNDU1LjMxOGwuMDAyLjAwMnEuMTk2LjE5NS4zMDcuNDY1bC4wMDEuMDAycS4xMTcuMjcuMTE3LjU2NnEwIC40MzUtLjAxOC44NjlxLS4wMTguNDQzLjAyNC44N3YuMDAxcS4wNS40MjUuMTk3LjgyNHYuMDAxcS4xNi40MS40OTQuNzUzcS4zMzUuMzQ1Ljc3NS41M3QuOTEuMTg1SDZ2LS45ODRoLS4wOXEtLjMgMC0uNTYzLS4xMTVhMS42IDEuNiAwIDAgMS0uNDU3LS4zMmExLjcgMS43IDAgMCAxLS4zMDktLjQ2N3EtLjExLS4yNy0uMTEtLjU3M3EwLS4zNDMuMDExLS42NzJxLjAxMi0uMzQyIDAtLjY2NWE1IDUgMCAwIDAtLjA1NS0uNjRhMi43IDIuNyAwIDAgMC0uMTY4LS42MDlBMi4zIDIuMyAwIDAgMCAzLjUyMiA4YTIuMyAyLjMgMCAwIDAgLjczOC0uOTU1cS4xMi0uMjg4LjE2OC0uNjAycS4wNS0uMzE1LjA1NS0uNjRxLjAxMi0uMzMgMC0uNjY2dC0uMDEyLS42NzhhMS40NyAxLjQ3IDAgMCAxIC44NzctMS4zNTRhMS4zIDEuMyAwIDAgMSAuNTYzLS4xMjF6bTQgMTAuMDMyVjE0aC4wOXEuNDcgMCAuOTA5LS4xODV0Ljc3NS0uNTN0LjQ5My0uNzUzdi0uMDAxcS4xNS0uNC4xOTgtLjgzdi0uMDAycS4wNDItLjQyLjAyNC0uODYzcS0uMDE4LS40MzUtLjAxOC0uODY5cTAtLjMwNC4xMTctLjU3MnYtLjAwMWExLjUgMS41IDAgMCAxIC43NjUtLjc4N2ExLjQgMS40IDAgMCAxIC41NTgtLjExNUgxNHYtLjk4NGgtLjA5cS0uMjkzIDAtLjU1Ny0uMTIxbC0uMDAxLS4wMDFhMS40IDEuNCAwIDAgMS0uNDU1LS4zMThsLS4wMDItLjAwMmExLjQgMS40IDAgMCAxLS4zMDctLjQ2NXYtLjAwMmExLjQgMS40IDAgMCAxLS4xMTgtLjU2NnEwLS40MzUuMDE4LS44NjlhNiA2IDAgMCAwLS4wMjQtLjg3di0uMDAxYTMuNSAzLjUgMCAwIDAtLjE5Ny0uODI0di0uMDAxYTIuMiAyLjIgMCAwIDAtLjQ5NC0uNzUzYTIuMyAyLjMgMCAwIDAtLjc3NS0uNTNhMi4zIDIuMyAwIDAgMC0uOTEtLjE4NUgxMHYuOTg0aC4wOXEuMyAwIC41NjIuMTE1cS4yNi4xMjMuNDU3LjMycS4xOS4yMDEuMzA5LjQ2N3EuMTEuMjcuMTEuNTczcTAgLjM0Mi0uMDExLjY3MnEtLjAxMi4zNDIgMCAuNjY1cS4wMDYuMzMzLjA1NS42NHEuMDUuMzIuMTY4LjYwOWEyLjMgMi4zIDAgMCAwIC43MzguOTU1YTIuMyAyLjMgMCAwIDAtLjczOC45NTVhMi43IDIuNyAwIDAgMC0uMTY4LjYwMnEtLjA1LjMxNS0uMDU1LjY0YTkgOSAwIDAgMCAwIC42NjZxLjAxMi4zMzYuMDEyLjY3OGExLjQ3IDEuNDcgMCAwIDEtLjg3NyAxLjM1NGExLjMgMS4zIDAgMCAxLS41NjMuMTIxeiIgY2xpcC1ydWxlPSJldmVub2RkIi8+PC9zdmc+`,
      action: 'parseJSONFromRange',
    }),
  };
  const temp =
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMTYgMTYiPjxwYXRoIGZpbGw9ImJsYWNrIiBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik02IDIuOTg0VjJoLS4wOXEtLjQ3IDAtLjkwOS4xODVhMi4zIDIuMyAwIDAgMC0uNzc1LjUzYTIuMiAyLjIgMCAwIDAtLjQ5My43NTN2LjAwMWEzLjUgMy41IDAgMCAwLS4xOTguODN2LjAwMmE2IDYgMCAwIDAtLjAyNC44NjNxLjAxOC40MzUuMDE4Ljg2OXEwIC4zMDQtLjExNy41NzJ2LjAwMWExLjUgMS41IDAgMCAxLS43NjUuNzg3YTEuNCAxLjQgMCAwIDEtLjU1OC4xMTVIMnYuOTg0aC4wOXEuMjkyIDAgLjU1Ni4xMjFsLjAwMS4wMDFxLjI2Ny4xMTcuNDU1LjMxOGwuMDAyLjAwMnEuMTk2LjE5NS4zMDcuNDY1bC4wMDEuMDAycS4xMTcuMjcuMTE3LjU2NnEwIC40MzUtLjAxOC44NjlxLS4wMTguNDQzLjAyNC44N3YuMDAxcS4wNS40MjUuMTk3LjgyNHYuMDAxcS4xNi40MS40OTQuNzUzcS4zMzUuMzQ1Ljc3NS41M3QuOTEuMTg1SDZ2LS45ODRoLS4wOXEtLjMgMC0uNTYzLS4xMTVhMS42IDEuNiAwIDAgMS0uNDU3LS4zMmExLjcgMS43IDAgMCAxLS4zMDktLjQ2N3EtLjExLS4yNy0uMTEtLjU3M3EwLS4zNDMuMDExLS42NzJxLjAxMi0uMzQyIDAtLjY2NWE1IDUgMCAwIDAtLjA1NS0uNjRhMi43IDIuNyAwIDAgMC0uMTY4LS42MDlBMi4zIDIuMyAwIDAgMCAzLjUyMiA4YTIuMyAyLjMgMCAwIDAgLjczOC0uOTU1cS4xMi0uMjg4LjE2OC0uNjAycS4wNS0uMzE1LjA1NS0uNjRxLjAxMi0uMzMgMC0uNjY2dC0uMDEyLS42NzhhMS40NyAxLjQ3IDAgMCAxIC44NzctMS4zNTRhMS4zIDEuMyAwIDAgMSAuNTYzLS4xMjF6bTQgMTAuMDMyVjE0aC4wOXEuNDcgMCAuOTA5LS4xODV0Ljc3NS0uNTN0LjQ5My0uNzUzdi0uMDAxcS4xNS0uNC4xOTgtLjgzdi0uMDAycS4wNDItLjQyLjAyNC0uODYzcS0uMDE4LS40MzUtLjAxOC0uODY5cTAtLjMwNC4xMTctLjU3MnYtLjAwMWExLjUgMS41IDAgMCAxIC43NjUtLjc4N2ExLjQgMS40IDAgMCAxIC41NTgtLjExNUgxNHYtLjk4NGgtLjA5cS0uMjkzIDAtLjU1Ny0uMTIxbC0uMDAxLS4wMDFhMS40IDEuNCAwIDAgMS0uNDU1LS4zMThsLS4wMDItLjAwMmExLjQgMS40IDAgMCAxLS4zMDctLjQ2NXYtLjAwMmExLjQgMS40IDAgMCAxLS4xMTgtLjU2NnEwLS40MzUuMDE4LS44NjlhNiA2IDAgMCAwLS4wMjQtLjg3di0uMDAxYTMuNSAzLjUgMCAwIDAtLjE5Ny0uODI0di0uMDAxYTIuMiAyLjIgMCAwIDAtLjQ5NC0uNzUzYTIuMyAyLjMgMCAwIDAtLjc3NS0uNTNhMi4zIDIuMyAwIDAgMC0uOTEtLjE4NUgxMHYuOTg0aC4wOXEuMyAwIC41NjIuMTE1cS4yNi4xMjMuNDU3LjMycS4xOS4yMDEuMzA5LjQ2N3EuMTEuMjcuMTEuNTczcTAgLjM0Mi0uMDExLjY3MnEtLjAxMi4zNDIgMCAuNjY1cS4wMDYuMzMzLjA1NS42NHEuMDUuMzIuMTY4LjYwOWEyLjMgMi4zIDAgMCAwIC43MzguOTU1YTIuMyAyLjMgMCAwIDAtLjczOC45NTVhMi43IDIuNyAwIDAgMC0uMTY4LjYwMnEtLjA1LjMxNS0uMDU1LjY0YTkgOSAwIDAgMCAwIC42NjZxLjAxMi4zMzYuMDEyLjY3OGExLjQ3IDEuNDcgMCAwIDEtLjg3NyAxLjM1NGExLjMgMS4zIDAgMCAxLS41NjMuMTIxeiIgY2xpcC1ydWxlPSJldmVub2RkIi8+PC9zdmc+';
  const tableIconButtonUris = [
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiPjxwYXRoIHN0cm9rZT0iI0UwNjY2NiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBzdHJva2Utd2lkdGg9IjIiIGQ9Ik0zIDhWNmEyIDIgMCAwIDEgMi0yaDE0YTIgMiAwIDAgMSAyIDJ2Mk0zIDh2Nm0wLTZoNm0xMiAwdjZtMC02SDltMTIgNnY0YTIgMiAwIDAgMS0yIDJIOW0xMi02SDltLTYgMHY0YTIgMiAwIDAgMCAyIDJoNG0tNi02aDZtMC02djZtMCAwdjZtNi0xMnYxMiIvPjxwYXRoIGZpbGw9IiNFMDY2NjYiIGQ9Ik0zIDZhMiAyIDAgMCAxIDItMmgxNGEyIDIgMCAwIDEgMiAydjJIM3oiLz48L2c+PC9zdmc+',
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiPjxwYXRoIHN0cm9rZT0iI0Y2QjI2QiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBzdHJva2Utd2lkdGg9IjIiIGQ9Ik0zIDhWNmEyIDIgMCAwIDEgMi0yaDE0YTIgMiAwIDAgMSAyIDJ2Mk0zIDh2Nm0wLTZoNm0xMiAwdjZtMC02SDltMTIgNnY0YTIgMiAwIDAgMS0yIDJIOW0xMi02SDltLTYgMHY0YTIgMiAwIDAgMCAyIDJoNG0tNi02aDZtMC02djZtMCAwdjZtNi0xMnYxMiIvPjxwYXRoIGZpbGw9IiNGNkIyNkIiIGQ9Ik0zIDZhMiAyIDAgMCAxIDItMmgxNGEyIDIgMCAwIDEgMiAydjJIM3oiLz48L2c+PC9zdmc+',
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiPjxwYXRoIHN0cm9rZT0iIzkzQzQ3RCIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBzdHJva2Utd2lkdGg9IjIiIGQ9Ik0zIDhWNmEyIDIgMCAwIDEgMi0yaDE0YTIgMiAwIDAgMSAyIDJ2Mk0zIDh2Nm0wLTZoNm0xMiAwdjZtMC02SDltMTIgNnY0YTIgMiAwIDAgMS0yIDJIOW0xMi02SDltLTYgMHY0YTIgMiAwIDAgMCAyIDJoNG0tNi02aDZtMC02djZtMCAwdjZtNi0xMnYxMiIvPjxwYXRoIGZpbGw9IiM5M0M0N0QiIGQ9Ik0zIDZhMiAyIDAgMCAxIDItMmgxNGEyIDIgMCAwIDEgMiAydjJIM3oiLz48L2c+PC9zdmc+',
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiPjxwYXRoIHN0cm9rZT0iIzZGQThEQyIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBzdHJva2Utd2lkdGg9IjIiIGQ9Ik0zIDhWNmEyIDIgMCAwIDEgMi0yaDE0YTIgMiAwIDAgMSAyIDJ2Mk0zIDh2Nm0wLTZoNm0xMiAwdjZtMC02SDltMTIgNnY0YTIgMiAwIDAgMS0yIDJIOW0xMi02SDltLTYgMHY0YTIgMiAwIDAgMCAyIDJoNG0tNi02aDZtMC02djZtMCAwdjZtNi0xMnYxMiIvPjxwYXRoIGZpbGw9IiM2RkE4REMiIGQ9Ik0zIDZhMiAyIDAgMCAxIDItMmgxNGEyIDIgMCAwIDEgMiAydjJIM3oiLz48L2c+PC9zdmc+',
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiPjxwYXRoIHN0cm9rZT0iIzhFN0NDMyIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBzdHJva2Utd2lkdGg9IjIiIGQ9Ik0zIDhWNmEyIDIgMCAwIDEgMi0yaDE0YTIgMiAwIDAgMSAyIDJ2Mk0zIDh2Nm0wLTZoNm0xMiAwdjZtMC02SDltMTIgNnY0YTIgMiAwIDAgMS0yIDJIOW0xMi02SDltLTYgMHY0YTIgMiAwIDAgMCAyIDJoNG0tNi02aDZtMC02djZtMCAwdjZtNi0xMnYxMiIvPjxwYXRoIGZpbGw9IiM4RTdDQzMiIGQ9Ik0zIDZhMiAyIDAgMCAxIDItMmgxNGEyIDIgMCAwIDEgMiAydjJIM3oiLz48L2c+PC9zdmc+',
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiPjxwYXRoIHN0cm9rZT0iI0MyN0JBMCIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBzdHJva2Utd2lkdGg9IjIiIGQ9Ik0zIDhWNmEyIDIgMCAwIDEgMi0yaDE0YTIgMiAwIDAgMSAyIDJ2Mk0zIDh2Nm0wLTZoNm0xMiAwdjZtMC02SDltMTIgNnY0YTIgMiAwIDAgMS0yIDJIOW0xMi02SDltLTYgMHY0YTIgMiAwIDAgMCAyIDJoNG0tNi02aDZtMC02djZtMCAwdjZtNi0xMnYxMiIvPjxwYXRoIGZpbGw9IiNDMjdCQTAiIGQ9Ik0zIDZhMiAyIDAgMCAxIDItMmgxNGEyIDIgMCAwIDEgMiAydjJIM3oiLz48L2c+PC9zdmc+',
  ];

  const cpButtonSet = () => {
    let buttonSet = CardService.newButtonSet();

    ['RED', 'ORANGE', 'GREEN', 'BLUE', 'PURPLE', 'PINK'].forEach((color, i) => {
      let hex = COLORS[color].BASE;
      buttonSet.addButton(
        createActionIconButtonWidget_({
          title: color,
          // iconUrl: `https://ui-avatars.com/api/?size=20&background=${hex.slice(
          //   1,
          // )}&color=${hex.slice(1)}&rounded=true`,
          iconUrl: tableIconButtonUris[i],
          action: 'formatTable',
          parameters: { color },
        }),
      );
    });
    return buttonSet;
  };

  const mpButtonSet = () => {
    let buttonSet = CardService.newButtonSet();

    ['#B30000', '#B36B00', '#00B300', '#0000B3', '#6B00B3', '#B300B3'].forEach(
      (color, i) => {
        buttonSet.addButton(
          createActionIconButtonWidget_({
            title: color,
            iconUrl: `https://ui-avatars.com/api/?size=20&background=${color.slice(
              1,
            )}&color=${color.slice(1)}&rounded=true`,
            // iconUrl: tableIconButtonUris[i],
            action: 'formatRangeAsTable',
            parameters: { color },
          }),
        );
      },
    );
    return buttonSet;
  };

  export function buildFileScopeRequestCard() {
    return CardService.newCardBuilder()
      .addSection(
        CardService.newCardSection()
          .addWidget(
            CardService.newTextParagraph().setText(
              'The add-on needs permission to access this spreadsheet.',
            ),
          )
          .addWidget(
            CardService.newTextButton()
              .setText('Request permission')
              .setOnClickAction(
                CardService.newAction().setFunctionName(
                  'onRequestFileScopeButtonClicked',
                ),
              ),
          ),
      )
      .build();
  }

  export function buildToolsCard() {
    const logPlaygroundSection = CardService.newCardSection()
      .setHeader('LOGS AND PLAYGROUNDS')
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(1)
      .addWidget(
        CardService.newButtonSet()
          .addButton(widgets.emptyLogRecordsWidget.asButtonIcon())
          .addButton(widgets.playground1Widget.asButtonIcon())
          .addButton(widgets.playground2Widget.asButtonIcon())
          .addButton(widgets.parseJSONFromSheetWidget.asButtonIcon())
          .addButton(widgets.parseJSONFromRangeWidget.asButtonIcon())
          .addButton(
            CardService.newImageButton()
              .setAltText('Run onOpen')
              .setOnClickAction(
                CardService.newAction().setFunctionName('onOpen'),
              )
              .setIconUrl(buildIconUrl_('menu_open')),
          )
          .addButton(widgets.helpWidget.asButtonIcon()),
      )
      .addWidget(DIVIDER)
      .addWidget(widgets.emptyLogRecordsWidget.asDecoratedText())
      .addWidget(widgets.playground1Widget.asDecoratedText())
      .addWidget(widgets.playground2Widget.asDecoratedText())
      .addWidget(widgets.parseJSONFromSheetWidget.asDecoratedText())
      .addWidget(widgets.parseJSONFromRangeWidget.asDecoratedText());

    const cropsAndColorsSection = CardService.newCardSection()
      .setHeader('CROPS AND COLOR TOOLS')
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(2)
      .addWidget(
        CardService.newButtonSet()
          .addButton(widgets.cropToSelectionWidget.asButtonIcon())
          .addButton(widgets.cropToDataWidget.asButtonIcon())
          .addButton(widgets.setBackgroundsUsingCellValuesWidget.asButtonIcon())
          .addButton(widgets.getBackgroundColorToValuesWidget.asButtonIcon())
          .addButton(widgets.getBackgroundColorToNotesWidget.asButtonIcon())
          .addButton(widgets.createSheetOfTogglesWidget.asButtonIcon()),
      )
      .addWidget(
        CardService.newButtonSet()
          .addButton(widgets.insertCommentBlockWidget.asButtonIcon())
          .addButton(widgets.createTogglesWidget.asButtonIcon()),
      )
      .addWidget(DIVIDER)
      .addWidget(widgets.cropToSelectionWidget.asDecoratedText())
      .addWidget(widgets.cropToDataWidget.asDecoratedText())
      .addWidget(DIVIDER)
      .addWidget(widgets.setBackgroundsUsingCellValuesWidget.asDecoratedText())
      .addWidget(widgets.getBackgroundColorToValuesWidget.asDecoratedText())
      .addWidget(widgets.getBackgroundColorToNotesWidget.asDecoratedText())
      .addWidget(widgets.matchFontToBackgroundWidget.asDecoratedText())
      .addWidget(widgets.makeFontNiceColorWidget.asDecoratedText())
      .addWidget(DIVIDER)
      .addWidget(widgets.squareSelectedCellsWidget.asDecoratedText())
      .addWidget(widgets.squareSideLengthWidget);

    /* ------------------------------------------------- Color Table ------------------------------------------------ */

    const tableFormatSection = CardService.newCardSection()
      .setHeader('TABLE FORMAT')
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(1)
      .addWidget(cpButtonSet())
      .addWidget(mpButtonSet())
      .addWidget(
        createFormElementWidget_({
          title:
            'Table format options. Shading preferences for the second row of colors can be set in the Addon settings (top-right menu)',
          id: 'tableOptions',
          checkboxGroup: {
            items: [
              {
                label: 'Title',
                value: 'hasTitle',
                selected: g.UserSettings.hasTitle,
              },
              {
                label: 'Headers',
                value: 'hasHeaders',
                selected: g.UserSettings.hasHeaders,
              },
              {
                label: 'Footer',
                value: 'hasFooter',
                selected: g.UserSettings.hasFooter,
              },
              {
                label: 'Leave top unchanged',
                value: 'leaveTop',
                selected: g.UserSettings.leaveTop,
              },
              {
                label: 'Leave left unchanged',
                value: 'leaveLeft',
                selected: g.UserSettings.leaveLeft,
              },
              {
                label: 'Leave bottom unchanged',
                value: 'leaveBottom',
                selected: g.UserSettings.leaveBottom,
              },
              {
                label: 'No bottom',
                value: 'noBottom',
                selected: g.UserSettings.noBottom,
              },
              {
                label: 'Center all',
                value: 'centerAll',
                selected: g.UserSettings.centerAll,
              },
              {
                label: 'Alternating',
                value: 'alternating',
                selected: g.UserSettings.alternating,
              },
            ],
          },
        }),
      )
      .addWidget(widgets.setTableFormatDefaultsWidget.asDecoratedText());

    const randomSection = CardService.newCardSection()
      .setHeader('RANDOM DATA')
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(0)
      .addWidget(widgets.requestRandomDataWidget.asDecoratedText())
      .addWidget(widgets.randomDataRequestedWidget)
      .addWidget(
        createFormElementWidget_({
          title: 'Fields to include',
          id: 'randomOptions',
          checkboxGroup: {
            items: [
              {
                label: 'Date',
                value: 'randomDate',
                selected: false,
              },
              {
                label: 'ID',
                value: 'randomId',
                selected: false,
              },
              {
                label: 'Name',
                value: 'randomName',
                selected: false,
              },
              {
                label: 'Address',
                value: 'randomAddress',
                selected: false,
              },
              {
                label: 'Phone',
                value: 'randomPhone',
                selected: false,
              },
              {
                label: 'Email',
                value: 'randomEmail',
                selected: false,
              },
            ],
          },
        }),
      )
      .addWidget(DIVIDER)
      .addWidget(widgets.requestRandomNumbersWidget.asDecoratedText())
      .addWidget(widgets.randomNumbersRequestedWidget)
      .addWidget(
        createFormElementWidget_({
          title: `It's either the data above or random numbers. Can't return both at the same time.`,
          id: 'randomNumberOptions',
          checkboxGroup: {
            items: [
              {
                label: 'Integers',
                value: 'randomIntegers',
                selected: false,
              },
              {
                label: 'Real Numbers',
                value: 'randomNumbers',
                selected: false,
              },
            ],
          },
        }),
      );

    const dashboardsSection = CardService.newCardSection()
      .setHeader('DASHBOARDS')
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(1)
      .addWidget(
        CardService.newButtonSet()
          .addButton(widgets.createStatsDashboardWidget.asButtonIcon())
          .addButton(widgets.createNamedRangesDashboardWidget.asButtonIcon())
          .addButton(widgets.showNamedFunctionsDashboardWidget.asButtonIcon())
          .addButton(widgets.showFormulaWidget.asButtonIcon()),
      )
      .addWidget(widgets.createStatsDashboardWidget.asDecoratedText())
      .addWidget(widgets.createNamedRangesDashboardWidget.asDecoratedText())
      .addWidget(widgets.updateNamedRangesDashboardWidget.asDecoratedText())
      .addWidget(widgets.createNamedRangesFromSheetWidget.asDecoratedText())
      .addWidget(widgets.createNamedRangesFromSelectionWidget.asDecoratedText())
      .addWidget(widgets.showFormulaWidget.asDecoratedText())
      .addWidget(DIVIDER);

    const staticSection = CardService.newCardSection()
      .setHeader('STATIC/DYNAMIC')
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(1)
      .addWidget(
        CardService.newButtonSet()
          .addButton(widgets.makeRangeStaticWidget.asButtonIcon())
          .addButton(widgets.makeRangeDynamicWidget.asButtonIcon())
          .addButton(widgets.makeSheetStaticWidget.asButtonIcon())
          .addButton(widgets.makeSheetDynamicWidget.asButtonIcon())
          .addButton(widgets.makeSheetStaticNoNotesWidget.asButtonIcon()),
      )
      .addWidget(DIVIDER)
      .addWidget(widgets.makeRangeStaticWidget.asDecoratedText())
      .addWidget(widgets.makeSheetStaticWidget.asDecoratedText())
      .addWidget(widgets.makeAllSheetsStaticWidget.asDecoratedText())
      .addWidget(widgets.makeSheetListStaticWidget.asDecoratedText())
      .addWidget(DIVIDER)
      .addWidget(widgets.makeRangeDynamicWidget.asDecoratedText())
      .addWidget(widgets.makeSheetDynamicWidget.asDecoratedText())
      .addWidget(widgets.makeAllSheetsDynamicWidget.asDecoratedText())
      .addWidget(widgets.makeSheetListDynamicWidget.asDecoratedText())
      .addWidget(DIVIDER)
      .addWidget(widgets.makeRangeStaticNoNotesWidget.asDecoratedText())
      .addWidget(widgets.makeSheetStaticNoNotesWidget.asDecoratedText());

    const captionFormatSection = CardService.newCardSection()
      .setHeader('CAPTIONS')
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(1)
      .addWidget(
        CardService.newTextParagraph().setText(
          'Format selected cells as captions',
        ),
      )
      .addWidget(
        CardService.newButtonSet()
          .addButton(widgets.captionTopLeftWidget.asTextButton())
          .addButton(widgets.captionTopWidget.asTextButton())
          .addButton(widgets.captionTopRightWidget.asTextButton()),
      )
      .addWidget(
        CardService.newButtonSet()
          .addButton(widgets.captionLeftWidget.asTextButton())
          .addButton(widgets.captionClearWidget.asTextButton())
          .addButton(widgets.captionRightWidget.asTextButton()),
      )
      .addWidget(
        CardService.newButtonSet()
          .addButton(widgets.captionBottomLeftWidget.asTextButton())
          .addButton(widgets.captionBottomWidget.asTextButton())
          .addButton(widgets.captionBottomRightWidget.asTextButton()),
      );

    /* --------------------------------------------- Build Final Card --------------------------------------------- */
    const toolsCard = CardService.newCardBuilder()
      .setHeader(
        CardService.newCardHeader()
          .setTitle(SpreadsheetApp.getActive().getName())
          .setSubtitle(getActiveInfoStr()),
      )
      .addSection(cropsAndColorsSection)
      .addSection(tableFormatSection)
      .addSection(staticSection)
      .addSection(dashboardsSection)
      .addSection(randomSection)
      .addSection(captionFormatSection)
      .addSection(logPlaygroundSection)
      .build();

    return toolsCard;
  }

  function getActiveInfoStr() {
    const activeSheetName = SpreadsheetApp.getActiveSheet().getName();
    const activeRangeA1Notation =
      SpreadsheetApp.getActiveRange().getA1Notation();
    return `activeSheet: ${activeSheetName}, activeRange: ${activeRangeA1Notation}`;
  }

  /**
   * Builds a card that displays details of an error.
   *
   * @param opts Parameters for building the card
   * @param opts.exception Exception that caused the error
   * @param opts.errorText Error message to show
   * @param opts.showStackTrace True if full stack trace should be displayed
   * @return The error card
   */
  export function buildErrorCard(
    opts: ErrorCardOptions,
  ): GoogleAppsScript.Card_Service.Card {
    var errorText = opts.errorText;

    if (opts.exception && !errorText) {
      errorText = opts.exception.toString();
    }

    if (!errorText) {
      errorText = 'No additional information is available.';
    }

    var card = CardService.newCardBuilder();
    card.setHeader(
      CardService.newCardHeader().setTitle('An unexpected error occurred'),
    );
    card.addSection(
      CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText(errorText),
      ),
    );

    if (opts.showStackTrace && opts.exception && opts.exception.stack) {
      var stack = opts.exception.stack.replace(/\n/g, '<br/>');
      card.addSection(
        CardService.newCardSection()
          .setHeader('Stack trace')
          .addWidget(CardService.newTextParagraph().setText(stack)),
      );
    }

    return card.build();
  }

  /**
   * Builds a card that displays the user settings.
   *
   * @param {SettingsOptions} opts Parameters for building the card
   * @param {number} opts.backgroundTitle
   * @param {number} opts.backgroundHeaders
   * @param {number} opts.backgroundDataFirst
   * @param {number} opts.backgroundDataSecond
   * @param {number} opts.backgroundFooter
   * @param {number} opts.bordersAll
   * @param {number} opts.bordersHorizontal
   * @param {number} opts.bordersVertical
   * @param {number} opts.bordersTitleBottom
   * @param {number} opts.bordersHeadersBottom
   * @param {number} opts.bordersHeadersVertical
   * @param {number} opts.bordersThickness
   * @param {string} opts.debugControl
   * @param {string} opts.helpControl
   * @return {Card}
   */
  export function buildSettingsCard(opts: SettingsOptions) {
    console.log('opts :>> ', opts);
    const builder = CardService.newCardBuilder();

    builder.setHeader(
      CardService.newCardHeader().setTitle(
        SpreadsheetApp.getActive().getName(),
      ),
    );

    const bordersThicknessWidget = CardService.newSelectionInput()
      .setTitle('Border thickness')
      .setFieldName('bordersThickness')
      .setType(CardService.SelectionInputType.DROPDOWN);

    [
      ['Solid', 1],
      ['Medium', 2],
      ['Thick', 3],
    ].forEach(([text, value]) => {
      bordersThicknessWidget.addItem(
        text,
        value,
        value == opts.bordersThickness,
      );
    });

    const tableFormatSection = CardService.newCardSection()
      .setHeader('Table format preferences')
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(0);

    [
      ['Title background', 'backgroundTitle'],
      ['Headers background', 'backgroundHeaders'],
      ['Data first background', 'backgroundDataFirst'],
      ['Data second background', 'backgroundDataSecond'],
      ['Footer background', 'backgroundFooter'],
      ['Borders all', 'bordersAll'],
      ['Horizontal borders', 'bordersHorizontal'],
      ['Vertical borders', 'bordersVertical'],
      ['Title bottom border', 'bordersTitleBottom'],
      ['Headers bottom border', 'bordersHeadersBottom'],
      ['Headers vertical border', 'bordersHeadersVertical'],
    ].forEach(([title, id]) => {
      const widget = CardService.newSelectionInput()
        .setTitle(title)
        .setFieldName(id)
        .setType(CardService.SelectionInputType.DROPDOWN);

      [...new Array(25)]
        .map((_, i) => [`${i + 1}`, i + 1])
        .forEach(([text, value]) => {
          widget.addItem(text, value, value == opts[id]);
        });
      tableFormatSection.addWidget(widget);
    });

    tableFormatSection.addWidget(bordersThicknessWidget);

    const debugSwitchWidget = CardService.newDecoratedText()
      .setText('Debug')
      .setBottomLabel('Toggle on to record actions')
      .setSwitchControl(
        CardService.newSwitch()
          .setControlType(CardService.SwitchControlType.SWITCH)
          .setFieldName('debugControl')
          .setValue('ON')
          .setSelected(opts.debugControl === 'ON'),
      );

    const actionButtonsWidget = CardService.newButtonSet()
      .addButton(
        CardService.newTextButton()
          .setText('Save')
          .setOnClickAction(
            CardService.newAction().setFunctionName(
              'ActionHandlers.saveSettings',
            ),
          ),
      )
      .addButton(
        CardService.newTextButton()
          .setText('Reset to defaults')
          .setOnClickAction(
            CardService.newAction().setFunctionName(
              'ActionHandlers.resetSettings',
            ),
          ),
      );

    const actionSection = CardService.newCardSection()
      .addWidget(actionButtonsWidget)
      .addWidget(debugSwitchWidget)
      .addWidget(CardService.newDivider());

    builder.addSection(tableFormatSection).addSection(actionSection);

    // const actionSection = CardService.newCardSection()
    //   .addWidget(actionButtonsWidget)
    //   .addWidget(debugSwitchWidget);

    // builder.addSection(actionSection);
    return builder.build();
  }

  /**
   * Constructs the full URL for an icon.
   * @param name Base file name of the icon
   * @return The full URL to the hosted icon.
   */
  function buildIconUrl_(name: string): string {
    return `https://fonts.gstatic.com/s/i/googlematerialicons/${name}/v8/gm_grey-24dp/2x/gm_${name}_gm_grey_24dp.png`;
  }

  /**
   * Creates an action that routes through the `dispatchAction` entry point.
   *
   * @param name Action handler name
   * @param optParams Additional parameters to pass through
   * @return The action
   * @private
   */
  function createAction_(
    name: string,
    optParams: object = {},
  ): GoogleAppsScript.Card_Service.Action {
    var params: { action?: string } = Object.assign({}, optParams);
    params.action = name;
    return (
      CardService.newAction()
        // .setFunctionName(`ActionHandlers.${name}`)
        .setFunctionName(`dispatchAction`)
        .setParameters(params)
    );
  }

  function createActionDecoratedTextWidget_(opts: {
    title?: string;
    description?: string;
    iconUrl?: string;
    action: string;
    parameters?: object;
    switch?: {
      id: string;
      value: string;
      type?: 'switch' | 'checkbox';
    };
  }): GoogleAppsScript.Card_Service.DecoratedText {
    let decoratedText = CardService.newDecoratedText()
      .setText(opts.title)
      .setBottomLabel(opts.description)
      .setOnClickAction(createAction_(opts.action, opts.parameters));

    if (false && opts.iconUrl) {
      let iconImage = CardService.newIconImage()
        .setIconUrl(opts.iconUrl)
        .setAltText(opts.title)
        .setImageCropType(CardService.ImageCropType.CIRCLE);

      decoratedText.setStartIcon(iconImage);
    }

    if (opts.switch) {
      decoratedText.setSwitchControl(
        CardService.newSwitch()
          .setControlType(
            opts.switch.type === 'checkbox'
              ? CardService.SwitchControlType.CHECK_BOX
              : CardService.SwitchControlType.SWITCH,
          )
          .setFieldName(opts.switch.id)
          .setValue(opts.switch.value),
      );
    }

    return decoratedText;
  }

  function createActionTextButtonWidget_(opts: {
    title?: string;
    action: string;
    parameters?: object;
  }): GoogleAppsScript.Card_Service.TextButton {
    return CardService.newTextButton()
      .setText(opts.title)
      .setTextButtonStyle(CardService.TextButtonStyle.OUTLINED)
      .setOnClickAction(createAction_(opts.action, opts.parameters));
  }

  function createActionIconButtonWidget_(opts: {
    title?: string;
    iconUrl?: string;
    action: string;
    parameters?: object;
  }): GoogleAppsScript.Card_Service.ImageButton {
    let button = CardService.newImageButton()
      .setAltText(opts.title)
      .setOnClickAction(createAction_(opts.action, opts.parameters));

    if (opts.iconUrl) {
      button.setIconUrl(opts.iconUrl);
    }

    return button;
  }

  function createFormElementWidget_(opts: {
    title?: string;
    id: string;
    text?: {
      hint?: string;
      default?: string;
      multiLine?: boolean;
    };
    dropdown?: {
      choices?: any[][];
      defaultChoice?: any;
    };
    checkboxGroup?: {
      items: Array<{
        label: Object;
        value: Object;
        selected: Boolean;
      }>;
    };
  }):
    | GoogleAppsScript.Card_Service.SelectionInput
    | GoogleAppsScript.Card_Service.TextInput {
    let formElementWidget;
    if (opts.text) {
      formElementWidget = CardService.newTextInput()
        .setTitle(opts.title)
        .setFieldName(opts.id);
      opts.text.hint && formElementWidget.setHint(opts.text.hint);
      opts.text.default && formElementWidget.setValue(opts.text.default);
      opts.text.multiLine && formElementWidget.setMultiline(true);
    }
    if (opts.dropdown) {
      formElementWidget = CardService.newSelectionInput()
        .setTitle(opts.title)
        .setFieldName(opts.id)
        .setType(CardService.SelectionInputType.DROPDOWN);
      for (var i = 0; i < opts.dropdown.choices.length; ++i) {
        var text = opts.dropdown.choices[i][0];
        var value = opts.dropdown.choices[i][1] || i;
        formElementWidget.addItem(
          text,
          value,
          text == opts.dropdown.defaultChoice ||
            value == opts.dropdown.defaultChoice,
        );
      }
    }
    if (opts.checkboxGroup) {
      formElementWidget = CardService.newSelectionInput()
        .setTitle(opts.title)
        .setFieldName(opts.id)
        .setType(CardService.SelectionInputType.CHECK_BOX);
      for (var i = 0; i < opts.checkboxGroup.items.length; ++i) {
        var label = opts.checkboxGroup.items[i].label;
        var itemValue = opts.checkboxGroup.items[i].value;
        var selected = opts.checkboxGroup.items[i].selected;
        formElementWidget.addItem(label, itemValue, selected);
      }
    }

    return formElementWidget;
  }

  /**
   * Creates an action that routes through the `dispatchAction` entry point.
   *
   * @param opts Parameters to pass through
   * @return Object that can be either a text widget, icon button, or form element
   * @private
   */
  function createActionWidget_(opts: {
    title?: string;
    description?: string;
    iconUrl?: string;
    altText?: string;
    action: string;
    parameters?: object;
    switch?: {
      id: string;
      value: string;
    };
    input?: {
      id: string;
      text?: {
        hint?: string;
        default?: string;
        multiLine?: boolean;
      };
      dropdown?: {
        choices?: any[][];
        defaultChoice?: any;
      };
      checkboxGroup?: {
        items: Array<{
          label: Object;
          value: Object;
          selected: Boolean;
        }>;
      };
    };
  }) {
    return {
      asDecoratedText: () => createActionDecoratedTextWidget_(opts),
      asButtonIcon: () => createActionIconButtonWidget_(opts),
      asTextButton: () => createActionTextButtonWidget_(opts),
      asFormElement: () =>
        createFormElementWidget_({ title: opts.title, ...opts.input }),
    };
  }
}

export { Views };
