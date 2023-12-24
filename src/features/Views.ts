import { COLORS } from '@core/lib/COLORS';
import { ErrorCardOptions } from '@core/types/addon';

namespace Views {
  const DIVIDER = CardService.newDivider();
  /**
   * A widgets object containing versatile action widgets for every action needed for Workspace Addon.
   * Eqch widget will generally have at a minimum: title, description, and action. Most will also have an iconUrl
   * for the icon that can either be used for icon buttons or as a begger icon decorated text. Other properties
   * include parameters for the action and details about switch toggle.
   */
  const widgets = {
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
    makeRangeStaticWidget: createActionWidget_({
      title: 'Make range static',
      description: 'Convert selected formulas to notes',
      action: 'makeRangeStatic',
    }),
    makeRangeDynamicWidget: createActionWidget_({
      title: 'Make range dynamic',
      description: 'Restore selected note formulas',
      action: 'makeRangeDynamic',
    }),
    makeSheetStaticWidget: createActionWidget_({
      title: 'Make sheet static',
      description: 'Convert all formulas to notes',
      iconUrl: buildIconUrl_('lock'),
      action: 'makeSheetStatic',
    }),
    makeSheetDynamicWidget: createActionWidget_({
      title: 'Make sheet dynamic',
      description: 'Restore all note formulas',
      iconUrl: buildIconUrl_('lock_open'),
      action: 'makeSheetDynamic',
    }),
    setBackgroundsUsingCellValuesWidget: createActionWidget_({
      title: 'Set Background',
      description: 'Set background color using cell value',
      iconUrl: buildIconUrl_('brush'),
      action: 'setBackgroundsColorsUsingValues',
    }),
    getBackgroundColorToNotesWidget: createActionWidget_({
      title: 'Get Background Colors To Notes',
      description: 'Put HEX and RGB info in notes',
      iconUrl: buildIconUrl_('palette'),
      action: 'getBackgroundColorsToNotes',
    }),
    getBackgroundColorToValuesWidget: createActionWidget_({
      title: 'Get Background Colors To Values',
      description: 'Put HEX value in cells',
      iconUrl: buildIconUrl_('colorize'),
      action: 'getBackgroundColorsToValues',
    }),
    createSheetOfTogglesWidget: createActionWidget_({
      title: 'Create sheet of toggles',
      description: 'Create a sheet of colored toggle checkboxes',
      iconUrl: buildIconUrl_('apps'),
      action: 'createSheetWithToggleButtons',
    }),
    squareSelectedCellsWidget: createActionWidget_({
      title: 'Square cells',
      description: 'Give same width/height to selected cells',
      iconUrl: buildIconUrl_('view_module'),
      action: 'squareSelectedCells',
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
    createNamedRangesFromSheetWidget: createActionWidget_({
      title: 'Create Sheet Named Ranges',
      description: 'sheetName.columnName + sheetName',
      iconUrl: buildIconUrl_('burst_mode'),
      action: 'createNamedRangesFromSheet',
    }),
    createNamedRangesDashboardWidget: createActionWidget_({
      title: 'Create Named Ranges Dashboard',
      description: 'Refresh/build named ranges sheet',
      iconUrl: buildIconUrl_('list_alt'),
      action: 'createNamedRangesDashboard',
    }),
    updateNamedRangesDashboardWidget: createActionWidget_({
      title: 'Update list',
      description: 'Update named ranges',
      iconUrl: buildIconUrl_('update'),
      action: 'updateNamedRangeSheet',
    }),
    createStatsDashboardWidget: createActionWidget_({
      title: 'Create Stats Dashboard',
      description: 'Get sheet utilization and formulas',
      iconUrl: buildIconUrl_('poll'),
      action: 'createStatsDashboard',
    }),
    emptyLogRecordsWidget: createActionWidget_({
      title: 'Empty logs',
      description: 'Empty cached log records',
      iconUrl: buildIconUrl_('bug_report'),
      action: 'emptyLogRecords',
    }),
    showNamedFunctionsDashboardWidget: createActionWidget_({
      title: 'Named Functions Dashboard',
      description: 'Show/create named functions dashboard',
      iconUrl: buildIconUrl_('loyalty'),
      action: 'showNamedFunctionsDashboard',
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
    showFormulaWidget: createActionWidget_({
      title: 'Show Formula',
      altText: '',
      description: 'Shows formula in cell directly below or two below',
      iconUrl: buildIconUrl_('functions'),
      action: 'showFormula',
    }),
    setTableFormatDefaultsWidget: createActionWidget_({
      title: 'Set as default',
      description: 'Sets current options as default.',
      iconUrl: buildIconUrl_('lock'),
      action: 'setTableFormatDefaults',
    }),
    cellSizeWidget: createFormElementWidget_({
      title: 'Length of side (pixels)',
      id: 'pixels',
      text: {
        multiLine: false,
      },
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
      .addWidget(widgets.cellSizeWidget);

    /* ------------------------------------------------- Color Table ------------------------------------------------ */

    // var settings = Settings.getSettingsForUser();

    const tableFormatSection = CardService.newCardSection()
      .setHeader('TABLE FORMAT')
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(1)
      .addWidget(cpButtonSet())
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
      .addWidget(widgets.makeRangeStaticWidget.asDecoratedText())
      .addWidget(widgets.makeRangeDynamicWidget.asDecoratedText())
      .addWidget(widgets.makeSheetStaticWidget.asDecoratedText())
      .addWidget(widgets.makeSheetDynamicWidget.asDecoratedText());

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
      .setHeader(CardService.newCardHeader().setTitle('GAS Tools'))
      .addSection(logPlaygroundSection)
      .addSection(cropsAndColorsSection)
      .addSection(tableFormatSection)
      .addSection(dashboardsSection)
      .addSection(captionFormatSection)
      .build();

    return toolsCard;
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
  export function buildSettingsCard() {
    const bordersThicknessWidget = createFormElementWidget_({
      title: 'Border thickness',
      id: 'bordersThickness',
      dropdown: {
        choices: [
          ['Solid', 1],
          ['Medium', 2],
          ['Thick', 3],
        ],
        defaultChoice: g.UserSettings.bordersThickness || 1,
      },
    });
    const debugSwitchWidget = CardService.newDecoratedText()
      .setText('Debug')
      .setBottomLabel('Toggle on to record actions')
      .setSwitchControl(
        CardService.newSwitch()
          .setControlType(CardService.SwitchControlType.SWITCH)
          .setFieldName('debugControl')
          .setValue('ON')
          .setSelected(g.UserSettings.debugControl === 'ON'),
      );

    const actionButtonsWidget = CardService.newButtonSet()
      .addButton(
        CardService.newTextButton()
          .setText('Save')
          .setOnClickAction(createAction_('saveSettings')),
      )
      .addButton(
        CardService.newTextButton()
          .setText('Reset to defaults')
          .setOnClickAction(createAction_('resetSettings')),
      );

    var tableFormatSection = CardService.newCardSection()
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
    ].forEach(([title, id]) =>
      tableFormatSection.addWidget(
        createFormElementWidget_({
          title,
          id,
          dropdown: {
            choices: [...new Array(25)].map((_, i) => [`${i + 1}`, i + 1]),
            defaultChoice: g.UserSettings[id],
          },
        }),
      ),
    );

    tableFormatSection.addWidget(bordersThicknessWidget);

    var actionSection = CardService.newCardSection()
      .addWidget(debugSwitchWidget)
      .addWidget(actionButtonsWidget);

    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Settings'))
      .addSection(tableFormatSection)
      .addSection(actionSection)
      .build();
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
      .setFunctionName('dispatchAction')
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

    if (opts.iconUrl) {
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

  /**
   * Creates a drop down for selecting from a list of choices.
   *
   * @param label Top label of widget
   * @param name Key used in form submits
   * @param choices Array of choices to display in dropdown
   * @param defaultChoice Default choice
   * @return The selected choice
   * @private
   */
  function createSelectDropdown_(
    label: string,
    name: string,
    choices: any[][],
    defaultChoice: any,
  ): GoogleAppsScript.Card_Service.SelectionInput {
    var widget = CardService.newSelectionInput()
      .setTitle(label)
      .setFieldName(name)
      .setType(CardService.SelectionInputType.DROPDOWN);
    for (var i = 0; i < choices.length; ++i) {
      var text = choices[i][0];
      var value = choices[i][1] || i;
      widget.addItem(text, value, text == defaultChoice);
    }
    return widget;
  }

  /**
   * Creates a drop down for selecting a time of day (hours only).
   *
   * @param label Top label of widget
   * @param name Key used in form submits
   * @param defaultValue Default duration to select (0-23)
   * @return The selected time of day
   * @private
   */
  function createIntegerSelectDropdown_(
    label: string,
    name: string,
    start: number,
    end: number,
    defaultValue: number,
  ): GoogleAppsScript.Card_Service.SelectionInput {
    var widget = CardService.newSelectionInput()
      .setTitle(label)
      .setFieldName(name)
      .setType(CardService.SelectionInputType.DROPDOWN);
    for (var i = start; i < end; ++i) {
      var text = `${i}`;
      widget.addItem(text, i, i == defaultValue);
    }
    return widget;
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
   * @param name Action handler name
   * @param optParams Additional parameters to pass through
   * @return The action
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
