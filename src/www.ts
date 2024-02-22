import { AddonResponse, ErrorHandler } from '@core/types/addon';
import { Shlog } from '@core/logging/Shlog';
import { Settings } from '@core/environment/Settings';
import { Views } from '@features/Views';
import { ActionHandlers } from '@features/ActionHandlers';
import { COLORS } from '@core/lib/COLORS';

Shlog.init();

const onFileScopeGranted = (e) => handleShowSheetsHomepage(e);

/** Callback function for when requesting permissions.
 * Instructs Docs to display a permissions dialog to the user, requesting `drive.file`
 * scope for the current file on behalf of this add-on.
 *
 * @return The dialog requesting scope permission
 */
const onRequestFileScopeButtonClicked =
  (): GoogleAppsScript.Card_Service.EditorFileScopeActionResponse => {
    return CardService.newEditorFileScopeActionResponseBuilder()
      .requestFileScopeForActiveDocument()
      .build();
  };

/**
 * Entry point for the add-on. Handles an user event and
 * invokes the corresponding action
 *
 * @return The settings card
 */
function handleShowSettings(
  event: GoogleAppsScript.Addons.EventObject,
): AddonResponse {
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
    debugControl: settings.debugControl,
    helpControl: settings.helpControl,
    bordersThickness: settings.bordersThickness,
  });
  return [card];
}

/**
 * Entry point for the add-on. Handles an user event and
 * invokes the corresponding action
 *
 * @param event - user event to process
 * @return Collection of home page cards
 */
function handleShowSheetsHomepage(
  event: GoogleAppsScript.Addons.EventObject,
): AddonResponse {
  event.commonEventObject.parameters = { action: 'showSheetsHomepage' };
  return dispatchActionInternal(event, addOnErrorHandler);
}

/**
 * Entry point for secondary actions. Handles an user event and
 * invokes the corresponding action
 *
 * @param event - user event to process
 * @return Card or form action
 */
function dispatchAction(
  event: GoogleAppsScript.Addons.EventObject,
): AddonResponse {
  return dispatchActionInternal(event, actionErrorHandler);
}

/**
 * Validates and dispatches an action.
 *
 * @param event - user event to process
 * @param optErrorHandler - Handles errors, optionally
 *        returning a card or action response.
 * @return Optional card or action response to render
 */
function dispatchActionInternal(
  event: GoogleAppsScript.Addons.EventObject,
  optErrorHandler: ErrorHandler,
): AddonResponse {
  let tictoc: number;
  const settings = Settings.getSettingsForUser();
  if (settings.debugControl === 'ON') {
    tictoc = Shlog.tic(event.commonEventObject.parameters.action, {
      parameters: event,
    });
    console.time('dispatchActionInternal');
  }
  try {
    var actionName = event.commonEventObject.parameters.action;
    if (!actionName) {
      throw new Error('Missing action name.');
    }

    if (settings.helpControl === 'on') {
      return ActionHandlers.getHelp(actionName);
    }
    var actionFn = ActionHandlers[actionName];
    if (!actionFn) {
      throw new Error('Action not found: ' + actionName);
    }
    return actionFn(event);
  } catch (err) {
    if (optErrorHandler) {
      return optErrorHandler(err);
    } else {
      throw err;
    }
  } finally {
    if (settings.debugControl === 'ON') {
      Shlog.toc(tictoc);
      console.timeEnd('dispatchActionInternal');
    }
  }
}

/**
 * Handle unexpected errors for the main entry points.
 *
 * @param err - Exception to handle
 * @return Optional card or action response to render
 */
function addOnErrorHandler(err: Error): AddonResponse {
  var card = Views.buildErrorCard({
    exception: err,
    showStackTrace: Settings.DEBUG,
  });
  return [card];
}

/**
 * Handle unexpected errors for the main universal action entry points.
 *
 * @param err - Exception to handle
 * @Return Card or action response to render
 */
function universalActionErrorHandler(
  err: Error,
): GoogleAppsScript.Card_Service.UniversalActionResponse {
  var card = Views.buildErrorCard({
    exception: err,
    showStackTrace: Settings.DEBUG,
  });
  return CardService.newUniversalActionResponseBuilder()
    .displayAddOnCards([card])
    .build();
}

/**
 * Handle unexpected errors for secondary actions.
 *
 * @param err - Exception to handle
 * @Return card or action response to render
 */
function actionErrorHandler(
  err: Error,
): GoogleAppsScript.Card_Service.ActionResponse {
  var card = Views.buildErrorCard({
    exception: err,
    showStackTrace: Settings.DEBUG,
  });
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card))
    .build();
}

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('My Sample React Project') // edit me!
    .addItem('Test Menu Action', 'ActionHandlers.createNamedRangesDashboard')
    .addItem('PLayground1', 'ActionHandlers.playground1')
    .addItem('PLayground3', 'ActionHandlers.playground3')
    .addItem('Show Settings', 'ActionHandlers.showSettings')
    .addItem('Show Event Object', 'showEventObject')
    .addItem('Run onOpen', 'onOpen')
    .addToUi();
}

function testMenuAction(e) {
  ActionHandlers.showSettings(e);
}
function showEventObject(e) {
  console.log('e:', JSON.stringify(e));
}

function dummy() {
  const values = [
    [
      'Color',
      'Bold',
      'Lightest',
      'Lighter',
      'Light',
      'Dark',
      'Darker',
      'Darkest',
    ],
    ...Object.keys(COLORS).map((key) => {
      return [
        key,
        COLORS[key].BOLD,
        COLORS[key].LIGHTEST,
        COLORS[key].LIGHTER,
        COLORS[key].LIGHT,
        COLORS[key].DARK,
        COLORS[key].DARKER,
        COLORS[key].DARKEST,
      ];
    }),
  ];
  g.ActiveSheet.getRange(1, 1, values.length, values[0].length).setValues(
    values,
  );
}
