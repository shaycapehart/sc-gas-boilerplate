const Help = {
  emptyLogRecords: {
    title: 'Empty logs',
    description: 'Empty cached log records',
    iconUrl: 'bug_report',
    action: 'emptyLogRecords',
  },
  playground1: {
    title: 'Playground 1',
    description: 'Run playground 1 function',
    iconUrl: 'looks_one',
    action: 'playground1',
  },
  playground2: {
    title: 'Playground 2',
    description: 'Run playground 2 function',
    iconUrl: 'looks_two',
    action: 'playground2',
  },
  playground3: {
    title: 'Playground 3',
    description: 'Run playground 3 function',
    iconUrl: 'looks_3',
    action: 'playground3',
  },
  help: {
    title: 'Click here and then any command',
    description:
      'Get help on any button by first clicking this icon and then click the other button to see a description.',
    iconUrl: 'help_outline',
    action: 'getHelp',
  },
  cropToSelection: {
    title: 'Crop to Selection',
    description: 'Crop sheet rows and columns down to selected range',
    iconUrl: 'crop_square',
    action: 'cropToSelection',
  },
  cropToData: {
    title: 'Crop to Data',
    description: 'Crop sheet rows and columns down to data range',
    iconUrl: 'crop',
    action: 'cropToData',
  },
  setBackgroundsUsingCellValues: {
    title: 'Set Backgrounds Using Values',
    description: 'Set background color using cell value',
    iconUrl: 'brush',
    action: 'setBackgroundsColorsUsingValues',
  },
  getBackgroundColorToValues: {
    title: 'Get Background Colors To Values',
    description: 'Put HEX value in cells',
    iconUrl: 'colorize',
    action: 'getBackgroundColorsToValues',
  },
  getBackgroundColorToNotes: {
    title: 'Get Background Colors To Notes',
    description: 'Put HEX and RGB info in notes',
    iconUrl: 'palette',
    action: 'getBackgroundColorsToNotes',
  },
  matchFontToBackground: {
    title: 'Match font to background',
    description: 'Make font color nearly the same as background color',
    iconUrl: 'font_download',
    action: 'matchFontToBackground',
  },
  makeFontNiceColor: {
    title: 'Make font a nice color',
    description: 'Make font color nice compared to background',
    iconUrl: 'font_download',
    action: 'makeFontNiceColor',
  },
  squareSelectedCells: {
    title: 'Square the selected cells',
    description: 'Give same width/height to selected cells',
    iconUrl: 'view_module',
    action: 'squareSelectedCells',
  },
  createToggles: {
    title: 'Create toggles',
    description: 'Choose a color and turn selected range into toggles',
    iconUrl: 'check_box',
    action: 'createToggles',
  },
  createSheetOfToggles: {
    title: 'Create toggles sheet',
    description: 'Create a sheet of colored toggle checkboxes',
    iconUrl: 'apps',
    action: 'createSheetWithToggleButtons',
  },
  createStatsDashboard: {
    title: 'Create Stats Dashboard',
    description: 'Get sheet utilization and formulas',
    iconUrl: 'poll',
    action: 'createStatsDashboard',
  },
  createNamedRangesDashboard: {
    title: 'Create Named Ranges Dashboard',
    description: 'Refresh/build named ranges sheet',
    iconUrl: 'list_alt',
    action: 'createNamedRangesDashboard',
    help: `
  Creates a dashboard sheet that lists all your named ranges. The dashboard allows you to edit or delete the named ranges in bulk or individually.`,
  },
  updateNamedRangesDashboard: {
    title: 'Update Named Ranges Dashboard',
    description: 'Update named ranges',
    iconUrl: 'update',
    action: 'updateNamedRangeSheet',
  },
  showNamedFunctionsDashboard: {
    title: 'Create Named Functions Dashboard',
    description: 'Show/create named functions dashboard',
    iconUrl: 'loyalty',
    action: 'showNamedFunctionsDashboard',
  },
  createNamedRangesFromSheet: {
    title: 'Create Named Ranges Using Current Sheet',
    description: 'sheetName.columnName + sheetName',
    iconUrl: 'burst_mode',
    action: 'createNamedRangesFromSheet',
  },
  createNamedRangesFromSelection: {
    title: 'Create Named Ranges Using Selection',
    description: 'sheetName.columnName + sheetName',
    iconUrl: 'burst_mode',
    action: 'createNamedRangesFromSelection',
  },
  showFormula: {
    title: 'Show Formula Down Below',
    altText: '',
    description: 'Shows formula in cell directly below or two below',
    iconUrl: 'functions',
    action: 'showFormula',
  },
  makeRangeStatic: {
    title: 'Make range static',
    description: 'Convert selected formulas to values, save to notes',
    iconUrl:
      'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjAgMjAiPjxwYXRoIGZpbGw9ImN1cnJlbnRDb2xvciIgZD0iTTE1LjUgOWMuNTQ2IDAgMS4wNTkuMTQ2IDEuNS40MDFWOGgtNHYyLjM0MUEzIDMgMCAwIDEgMTUuNSA5TTExIDE0YzAtLjM2NC4wOTctLjcwNi4yNjgtMUg4djRoM3ptMS0ySDhWOGg0em0tNSAwVjhIM3Y0em0tNCAxaDR2NEg1LjVBMi41IDIuNSAwIDAgMSAzIDE0LjV6bTEwLTZoNFY1LjVBMi41IDIuNSAwIDAgMCAxNC41IDNIMTN6bS0xLTR2NEg4VjN6TTcgM3Y0SDNWNS41QTIuNSAyLjUgMCAwIDEgNS41IDN6bTYuNSA5djFIMTNhMSAxIDAgMCAwLTEgMXY0YTEgMSAwIDAgMCAxIDFoNWExIDEgMCAwIDAgMS0xdi00YTEgMSAwIDAgMC0xLTFoLS41di0xYTIgMiAwIDEgMC00IDBtMSAxdi0xYTEgMSAwIDEgMSAyIDB2MXptMSAyLjI1YS43NS43NSAwIDEgMSAwIDEuNWEuNzUuNzUgMCAwIDEgMC0xLjUiLz48L3N2Zz4=',
    action: 'makeRangeStatic',
  },
  makeRangeDynamic: {
    title: 'Make range dynamic',
    iconUrl: 'lock_open',
    description: 'Restore selected note formulas',
    action: 'makeRangeDynamic',
  },
  makeSheetStatic: {
    title: 'Make sheet static',
    description: 'Convert all formulas to values, save to notes',
    iconUrl:
      'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMzIgMzIiPjxwYXRoIGZpbGw9ImN1cnJlbnRDb2xvciIgZD0iTTE3LjA4NiA0LjQwMWMtMS41ODgtMS4yNTYtMy45MzUtLjIyOC00LjA4NyAxLjc5MWwtLjIxIDIuODA0SDE3YTEgMSAwIDEgMSAwIDJoLTQuMzYyTDExLjQ4NiAyNi4zMWMtLjI4OCAzLjgzMS00LjkxOCA1LjU4My03LjY3IDIuOTAzbC0uNTE0LS41YTEgMSAwIDAgMSAxLjM5Ni0xLjQzNGwuNTEzLjVjMS41MzYgMS40OTcgNC4xMi41MiA0LjI4MS0xLjYybDEuMTQtMTUuMTYzSDhhMSAxIDAgMSAxIDAtMmgyLjc4M2wuMjIyLTIuOTU0Yy4yNzItMy42MTggNC40NzctNS40NiA3LjMyMi0zLjIwOWwuNzk0LjYyOWExIDEgMCAxIDEtMS4yNDIgMS41Njh6bS0uODggMTEuODdhMS4xMTUgMS4xMTUgMCAwIDEgMS43NS4zOGwxLjg2MiA0LjExM2wtNS41MjUgNS41MjVhMSAxIDAgMCAwIDEuNDE0IDEuNDE0bDQuOTkyLTQuOTkybDEuNTc0IDMuNDc2Yy45MDUgMS45OTggMy41MzUgMi40NiA1LjA2Ny44OTJsLjM3NS0uMzg1YTEgMSAwIDEgMC0xLjQzLTEuMzk3bC0uMzc2LjM4NGExLjExNSAxLjExNSAwIDAgMS0xLjgxNC0uMzE5TDIyLjIxIDIxLjJsNS40OTctNS40OTdhMSAxIDAgMCAwLTEuNDE0LTEuNDE0bC00Ljk2NCA0Ljk2NGwtMS41NTItMy40MjdhMy4xMTUgMy4xMTUgMCAwIDAtNC44ODctMS4wNjJsLS41NDguNDc5YTEgMSAwIDAgMCAxLjMxNiAxLjUwNnoiLz48L3N2Zz4=',
    action: 'makeSheetsStatic',
    parameters: { use: 'current' },
  },
  makeSheetDynamic: {
    title: 'Make sheet dynamic',
    description: 'Restore all note formulas',
    iconUrl: 'lock_open',
    action: 'makeSheetsDynamic',
    parameters: { use: 'current' },
  },
  makeAllSheetsStatic: {
    title: 'Make all sheets static',
    description: 'Convert all formulas on all sheets to values, save in notes',
    iconUrl:
      'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMTYgMTYiPjxnIGZpbGw9ImN1cnJlbnRDb2xvciI+PHBhdGggZD0iTTEzIDUuNjk4YTUgNSAwIDAgMS0uOTA0LjUyNUMxMS4wMjIgNi43MTEgOS41NzMgNyA4IDdzLTMuMDIyLS4yODktNC4wOTYtLjc3N0E1IDUgMCAwIDEgMyA1LjY5OFY3YzAgLjM3NC4zNTYuODc1IDEuMzE4IDEuMzEzQzUuMjM0IDguNzI5IDYuNTM2IDkgOCA5Yy42NjYgMCAxLjI5OC0uMDU2IDEuODc2LS4xNTZjLS40My4zMS0uODA0LjY5My0xLjEwMiAxLjEzMkExMiAxMiAwIDAgMSA4IDEwYy0xLjU3MyAwLTMuMDIyLS4yODktNC4wOTYtLjc3N0E1IDUgMCAwIDEgMyA4LjY5OFYxMGMwIC4zNzQuMzU2Ljg3NSAxLjMxOCAxLjMxM0M1LjIzNCAxMS43MjkgNi41MzYgMTIgOCAxMmguMDI3YTQuNiA0LjYgMCAwIDAtLjAxNy44QTIgMiAwIDAgMCA4IDEzYy0xLjU3MyAwLTMuMDIyLS4yODktNC4wOTYtLjc3N0E1IDUgMCAwIDEgMyAxMS42OThWMTNjMCAuMzc0LjM1Ni44NzUgMS4zMTggMS4zMTNDNS4yMzQgMTQuNzI5IDYuNTM2IDE1IDggMTVjMCAuMzYzLjA5Ny43MDQuMjY2Ljk5N1E4LjEzNCAxNi4wMDEgOCAxNmMtMS41NzMgMC0zLjAyMi0uMjg5LTQuMDk2LS43NzdDMi44NzUgMTQuNzU1IDIgMTQuMDA3IDIgMTNWNGMwLTEuMDA3Ljg3NS0xLjc1NSAxLjkwNC0yLjIyM0M0Ljk3OCAxLjI4OSA2LjQyNyAxIDggMXMzLjAyMi4yODkgNC4wOTYuNzc3QzEzLjEyNSAyLjI0NSAxNCAyLjk5MyAxNCA0djQuMjU2YTQuNSA0LjUgMCAwIDAtMS43NTMtLjI0OUMxMi43ODcgNy42NTQgMTMgNy4yODkgMTMgN3ptLTguNjgyLTMuMDFDMy4zNTYgMy4xMjQgMyAzLjYyNSAzIDRjMCAuMzc0LjM1Ni44NzUgMS4zMTggMS4zMTNDNS4yMzQgNS43MjkgNi41MzYgNiA4IDZzMi43NjYtLjI3IDMuNjgyLS42ODdDMTIuNjQ0IDQuODc1IDEzIDQuMzczIDEzIDRjMC0uMzc0LS4zNTYtLjg3NS0xLjMxOC0xLjMxM0MxMC43NjYgMi4yNzEgOS40NjQgMiA4IDJzLTIuNzY2LjI3LTMuNjgyLjY4N1oiLz48cGF0aCBkPSJNOSAxM2ExIDEgMCAwIDEgMS0xdi0xYTIgMiAwIDEgMSA0IDB2MWExIDEgMCAwIDEgMSAxdjJhMSAxIDAgMCAxLTEgMWgtNGExIDEgMCAwIDEtMS0xem0zLTNhMSAxIDAgMCAwLTEgMXYxaDJ2LTFhMSAxIDAgMCAwLTEtMSIvPjwvZz48L3N2Zz4=',
    action: 'makeSheetsStatic',
    parameters: { use: 'all' },
  },
  makeAllSheetsDynamic: {
    title: 'Restore formulas for all sheets',
    description: 'Restore all note formulas on all sheets',
    iconUrl: 'lock_open',
    action: 'makeSheetsDynamic',
    parameters: { use: 'all' },
  },
  makeSheetListStatic: {
    title: 'Make sheet name list static',
    description:
      'Convert all formulas on selected sheet names to values, save in notes',
    iconUrl:
      'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxwYXRoIGZpbGw9IiNmNjA5MDkiIGQ9Ik0yMS44IDE2di0xLjVjMC0xLjQtMS40LTIuNS0yLjgtMi41cy0yLjggMS4xLTIuOCAyLjVWMTZjLS42IDAtMS4yLjYtMS4yIDEuMnYzLjVjMCAuNy42IDEuMyAxLjIgMS4zaDUuNWMuNyAwIDEuMy0uNiAxLjMtMS4ydi0zLjVjMC0uNy0uNi0xLjMtMS4yLTEuM20tMS4zIDBoLTN2LTEuNWMwLS44LjctMS4zIDEuNS0xLjNzMS41LjUgMS41IDEuM3pNMyAxM3YtMmgxMnYyem0wLTdoMTh2Mkgzem0wIDEydi0yaDZ2MnoiLz48L3N2Zz4=',
    action: 'makeSheetsStatic',
    parameters: { use: 'list' },
  },
  makeSheetListDynamic: {
    title: 'Make sheet name list dynamic',
    description: 'Restore all note formulas on selected sheet names',
    iconUrl: 'lock_open',
    action: 'makeSheetsDynamic',
    parameters: { use: 'list' },
  },
  makeRangeStaticNoNotes: {
    title: 'Make range static, no notes',
    description:
      'Convert all formulas in selected range to values. May affect array formulas above range',
    iconUrl: 'lock',
    action: 'makeRangeStatic',
    parameters: { noNotes: 'true' },
  },
  makeSheetStaticNoNotes: {
    title: 'Make sheet static, no notes',
    description: 'Convert all formulas to values',
    iconUrl: 'lock',
    action: 'makeSheetsStatic',
    parameters: { use: 'current', noNotes: 'true' },
  },
  makeAllSheetsStaticNoNotes: {
    title: 'Make all sheets static, no notes',
    description: 'Convert all formulas on all sheets to values',
    iconUrl: 'lock',
    action: 'makeSheetsStatic',
    parameters: { use: 'all', noNotes: 'true' },
  },
  makeSheetListStaticNoNotes: {
    title: 'Make sheet list static, no notes',
    description: 'Convert all formulas on selected sheet names to values',
    iconUrl: 'lock',
    action: 'makeSheetsStatic',
    parameters: { use: 'list', noNotes: 'true' },
  },
  insertCommentBlock: {
    title: 'Insert comment block',
    description: 'Merge selected range into a single yellow comment cell',
    iconUrl: 'comment',
    action: 'insertCommentBlock',
  },
  captionTopLeft: {
    title: '⭶',
    action: 'formatAsCaption',
    parameters: { location: 'top-left' },
  },
  captionTopRight: {
    title: '⭷',
    action: 'formatAsCaption',
    parameters: { location: 'top-right' },
  },
  captionBottomLeft: {
    title: '⭹',
    action: 'formatAsCaption',
    parameters: { location: 'bottom-left' },
  },
  captionBottomRight: {
    title: '⭸',
    action: 'formatAsCaption',
    parameters: { location: 'bottom-right' },
  },
  captionBottom: {
    title: '⭣',
    action: 'formatAsCaption',
    parameters: { location: 'bottom-center' },
  },
  captionBottomBottom: {
    title: '⮇',
    action: 'formatAsCaption',
    parameters: { location: 'bottombottom' },
  },
  captionRight: {
    title: '⭢',
    action: 'formatAsCaption',
    parameters: { location: 'middle-right' },
  },
  captionLeft: {
    title: '⭠',
    action: 'formatAsCaption',
    parameters: { location: 'middle-left' },
  },
  captionTop: {
    title: '⭡',
    action: 'formatAsCaption',
    parameters: { location: 'top-center' },
  },
  captionTopTop: {
    title: '⮅',
    action: 'formatAsCaption',
    parameters: { location: 'toptop' },
  },
  captionClear: {
    title: '⭙',
    altText: 'Clear formatting',
    action: 'formatAsCaption',
    parameters: { location: 'clear' },
  },
  setTableFormatDefaults: {
    title: 'Set as default',
    description: 'Sets current options as default.',
    iconUrl: 'lock',
    action: 'setTableFormatDefaults',
  },
  squareSideLength: {
    title: 'Length of side (pixels)',
    id: 'pixels',
    text: {
      multiLine: false,
    },
  },
  requestRandomData: {
    title: 'Get random data',
    description:
      'Enter the number of rows and what data to insert at selected cell',
    iconUrl: 'shuffle',
    action: 'requestRandomData',
  },
  randomDataRequested: {
    title: 'How many do you want?',
    id: 'nRandomData',
    text: {
      multiLine: false,
    },
  },
  requestRandomNumbers: {
    title: 'Get random numbers',
    description:
      'Enter the number of rows and what data to insert at selected cell',
    iconUrl: 'shuffle',
    action: 'requestRandomNumbers',
  },
  randomNumbersRequested: {
    title: 'How many do you want?',
    id: 'nRandomNumbers',
    text: {
      multiLine: false,
    },
  },
  tableIconButtonUris: [
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiPjxwYXRoIHN0cm9rZT0iI0UwNjY2NiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBzdHJva2Utd2lkdGg9IjIiIGQ9Ik0zIDhWNmEyIDIgMCAwIDEgMi0yaDE0YTIgMiAwIDAgMSAyIDJ2Mk0zIDh2Nm0wLTZoNm0xMiAwdjZtMC02SDltMTIgNnY0YTIgMiAwIDAgMS0yIDJIOW0xMi02SDltLTYgMHY0YTIgMiAwIDAgMCAyIDJoNG0tNi02aDZtMC02djZtMCAwdjZtNi0xMnYxMiIvPjxwYXRoIGZpbGw9IiNFMDY2NjYiIGQ9Ik0zIDZhMiAyIDAgMCAxIDItMmgxNGEyIDIgMCAwIDEgMiAydjJIM3oiLz48L2c+PC9zdmc+',
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiPjxwYXRoIHN0cm9rZT0iI0Y2QjI2QiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBzdHJva2Utd2lkdGg9IjIiIGQ9Ik0zIDhWNmEyIDIgMCAwIDEgMi0yaDE0YTIgMiAwIDAgMSAyIDJ2Mk0zIDh2Nm0wLTZoNm0xMiAwdjZtMC02SDltMTIgNnY0YTIgMiAwIDAgMS0yIDJIOW0xMi02SDltLTYgMHY0YTIgMiAwIDAgMCAyIDJoNG0tNi02aDZtMC02djZtMCAwdjZtNi0xMnYxMiIvPjxwYXRoIGZpbGw9IiNGNkIyNkIiIGQ9Ik0zIDZhMiAyIDAgMCAxIDItMmgxNGEyIDIgMCAwIDEgMiAydjJIM3oiLz48L2c+PC9zdmc+',
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiPjxwYXRoIHN0cm9rZT0iIzkzQzQ3RCIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBzdHJva2Utd2lkdGg9IjIiIGQ9Ik0zIDhWNmEyIDIgMCAwIDEgMi0yaDE0YTIgMiAwIDAgMSAyIDJ2Mk0zIDh2Nm0wLTZoNm0xMiAwdjZtMC02SDltMTIgNnY0YTIgMiAwIDAgMS0yIDJIOW0xMi02SDltLTYgMHY0YTIgMiAwIDAgMCAyIDJoNG0tNi02aDZtMC02djZtMCAwdjZtNi0xMnYxMiIvPjxwYXRoIGZpbGw9IiM5M0M0N0QiIGQ9Ik0zIDZhMiAyIDAgMCAxIDItMmgxNGEyIDIgMCAwIDEgMiAydjJIM3oiLz48L2c+PC9zdmc+',
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiPjxwYXRoIHN0cm9rZT0iIzZGQThEQyIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBzdHJva2Utd2lkdGg9IjIiIGQ9Ik0zIDhWNmEyIDIgMCAwIDEgMi0yaDE0YTIgMiAwIDAgMSAyIDJ2Mk0zIDh2Nm0wLTZoNm0xMiAwdjZtMC02SDltMTIgNnY0YTIgMiAwIDAgMS0yIDJIOW0xMi02SDltLTYgMHY0YTIgMiAwIDAgMCAyIDJoNG0tNi02aDZtMC02djZtMCAwdjZtNi0xMnYxMiIvPjxwYXRoIGZpbGw9IiM2RkE4REMiIGQ9Ik0zIDZhMiAyIDAgMCAxIDItMmgxNGEyIDIgMCAwIDEgMiAydjJIM3oiLz48L2c+PC9zdmc+',
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiPjxwYXRoIHN0cm9rZT0iIzhFN0NDMyIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBzdHJva2Utd2lkdGg9IjIiIGQ9Ik0zIDhWNmEyIDIgMCAwIDEgMi0yaDE0YTIgMiAwIDAgMSAyIDJ2Mk0zIDh2Nm0wLTZoNm0xMiAwdjZtMC02SDltMTIgNnY0YTIgMiAwIDAgMS0yIDJIOW0xMi02SDltLTYgMHY0YTIgMiAwIDAgMCAyIDJoNG0tNi02aDZtMC02djZtMCAwdjZtNi0xMnYxMiIvPjxwYXRoIGZpbGw9IiM4RTdDQzMiIGQ9Ik0zIDZhMiAyIDAgMCAxIDItMmgxNGEyIDIgMCAwIDEgMiAydjJIM3oiLz48L2c+PC9zdmc+',
    'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxZW0iIGhlaWdodD0iMWVtIiB2aWV3Qm94PSIwIDAgMjQgMjQiPjxnIGZpbGw9Im5vbmUiPjxwYXRoIHN0cm9rZT0iI0MyN0JBMCIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBzdHJva2Utd2lkdGg9IjIiIGQ9Ik0zIDhWNmEyIDIgMCAwIDEgMi0yaDE0YTIgMiAwIDAgMSAyIDJ2Mk0zIDh2Nm0wLTZoNm0xMiAwdjZtMC02SDltMTIgNnY0YTIgMiAwIDAgMS0yIDJIOW0xMi02SDltLTYgMHY0YTIgMiAwIDAgMCAyIDJoNG0tNi02aDZtMC02djZtMCAwdjZtNi0xMnYxMiIvPjxwYXRoIGZpbGw9IiNDMjdCQTAiIGQ9Ik0zIDZhMiAyIDAgMCAxIDItMmgxNGEyIDIgMCAwIDEgMiAydjJIM3oiLz48L2c+PC9zdmc+',
  ],
};

export { Help };
