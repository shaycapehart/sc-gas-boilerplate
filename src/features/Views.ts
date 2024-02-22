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
      action: 'prepareForHelp',
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
    squareSelectedCellsWidget: createActionWidget_({
      title: 'Square the selected cells',
      description: 'Give same width/height to selected cells',
      iconUrl: buildIconUrl_('view_module'),
      action: 'squareSelectedCells',
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
      title: 'Create Named Ranges Using Current Sheet',
      description: 'sheetName.columnName + sheetName',
      iconUrl: buildIconUrl_('burst_mode'),
      action: 'createNamedRangesFromSheet',
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
      iconUrl: buildIconUrl_('lock'),
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
      iconUrl: buildIconUrl_('lock'),
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
      iconUrl: buildIconUrl_('lock'),
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
      iconUrl: buildIconUrl_('lock'),
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
  };

  const cpButtonSet = () => {
    let buttonSet = CardService.newButtonSet();

    ['RED', 'ORANGE', 'GREEN', 'BLUE', 'PURPLE', 'PINK'].forEach((color) => {
      let hex = COLORS[color].BASE;
      buttonSet.addButton(
        createActionIconButtonWidget_({
          title: color,
          iconUrl: `https://ui-avatars.com/api/?size=20&background=${hex.slice(
            1,
          )}&color=${hex.slice(1)}&rounded=true`,
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
      (color) => {
        buttonSet.addButton(
          createActionIconButtonWidget_({
            title: color,
            iconUrl: `https://ui-avatars.com/api/?size=20&background=${color.slice(
              1,
            )}&color=${color.slice(1)}&rounded=true`,
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
          .addButton(widgets.playground3Widget.asButtonIcon())
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
      .addWidget(widgets.playground3Widget.asDecoratedText());

    const cropsAndColorsSection = CardService.newCardSection()
      .setHeader('CROPS AND COLOR TOOLS')
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(1)
      .addWidget(
        CardService.newButtonSet()
          .addButton(widgets.cropToSelectionWidget.asButtonIcon())
          .addButton(widgets.cropToDataWidget.asButtonIcon())
          .addButton(widgets.setBackgroundsUsingCellValuesWidget.asButtonIcon())
          .addButton(widgets.getBackgroundColorToValuesWidget.asButtonIcon())
          .addButton(widgets.getBackgroundColorToNotesWidget.asButtonIcon())
          .addButton(widgets.createSheetOfTogglesWidget.asButtonIcon()),
      )
      .addWidget(DIVIDER)
      .addWidget(widgets.cropToSelectionWidget.asDecoratedText())
      .addWidget(widgets.cropToDataWidget.asDecoratedText())
      .addWidget(DIVIDER)
      .addWidget(widgets.setBackgroundsUsingCellValuesWidget.asDecoratedText())
      .addWidget(widgets.getBackgroundColorToValuesWidget.asDecoratedText())
      .addWidget(widgets.getBackgroundColorToNotesWidget.asDecoratedText())
      .addWidget(DIVIDER)
      .addWidget(widgets.squareSelectedCellsWidget.asDecoratedText())
      .addWidget(
        createFormElementWidget_({
          title: 'Length of side (pixels)',
          id: 'pixels',
          text: {
            multiLine: false,
          },
        }),
      );

    /* ------------------------------------------------- Color Table ------------------------------------------------ */

    const tableFormatSection = CardService.newCardSection()
      .setHeader('TABLE FORMAT')
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(1)
      .addWidget(cpButtonSet())
      .addWidget(mpButtonSet())
      .addWidget(
        createFormElementWidget_({
          title: 'Table format options',
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

    builder
      .addSection(tableFormatSection)
      .addSection(
        CardService.newCardSection()
          .setHeader('Settings')
          .setCollapsible(false)
          .addWidget(bordersThicknessWidget),
      );

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
      .addWidget(debugSwitchWidget);

    builder.addSection(actionSection);
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
    return CardService.newAction()
      .setFunctionName(`ActionHandlers.${name}`)
      .setParameters(params);
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
      .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
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
