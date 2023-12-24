import { SettingsOptions } from '@core/types/addon';

namespace Settings {
  export const DEBUG = false;
  const SETTINGS_KEY = 'settings';
  export const APP_TITLE = 'GV-Tools';
  export const DEFAULT = {
    TABLE: {
      HAS_TITLE: false,
      HAS_HEADERS: true,
      HAS_FOOTER: false,
      LEAVE_TOP: false,
      LEAVE_LEFT: false,
      LEAVE_BOTTOM: false,
      NO_BOTTOM: false,
      CENTER_ALL: true,
      ALTERNATING: false,
    },
    BACKGROUND_TITLE: 6,
    BACKGROUND_HEADERS: 14,
    BACKGROUND_DATA_FIRST: 19,
    BACKGROUND_DATA_SECOND: 16,
    BACKGROUND_FOOTER: 13,
    BORDERS_ALL: 6,
    BORDERS_HORIZONTAL: 10,
    BORDERS_VERTICAL: 10,
    BORDERS_TITLE_BOTTOM: 6,
    BORDERS_HEADERS_BOTTOM: 9,
    BORDERS_HEADERS_VERTICAL: 10,
    BORDERS_THICKNESS: 1,
    DEBUG_CONTROL: 'ON',
    HELP_CONTROL: 'off',
    LOG_SHEETNAME: 'app_log',
  };

  /**
   * Attempts to determine the user's timezone. Defaults to the script's
   * timezone if unable to do so.
   *
   * @return {string}
   */
  export function getUserTimezone() {
    var result = Calendar.Settings.get('timezone');
    return result.value ? result.value : Session.getScriptTimeZone();
  }

  /**
   * Get the effective settings for the current user.
   *
   * @return {Object}
   */
  export function getSettingsForUser(): SettingsOptions {
    var savedSettings = cachedPropertiesForUser_().get(SETTINGS_KEY, {});
    return Object.assign(
      {},
      {
        hasTitle: Settings.DEFAULT.TABLE.HAS_TITLE,
        hasHeaders: Settings.DEFAULT.TABLE.HAS_HEADERS,
        hasFooter: Settings.DEFAULT.TABLE.HAS_FOOTER,
        leaveTop: Settings.DEFAULT.TABLE.LEAVE_TOP,
        leaveLeft: Settings.DEFAULT.TABLE.LEAVE_LEFT,
        leaveBottom: Settings.DEFAULT.TABLE.LEAVE_BOTTOM,
        noBottom: Settings.DEFAULT.TABLE.NO_BOTTOM,
        centerAll: Settings.DEFAULT.TABLE.CENTER_ALL,
        alternating: Settings.DEFAULT.TABLE.ALTERNATING,
        backgroundTitle: Settings.DEFAULT.BACKGROUND_TITLE,
        backgroundHeaders: Settings.DEFAULT.BACKGROUND_HEADERS,
        backgroundDataFirst: Settings.DEFAULT.BACKGROUND_DATA_FIRST,
        backgroundDataSecond: Settings.DEFAULT.BACKGROUND_DATA_SECOND,
        backgroundFooter: Settings.DEFAULT.BACKGROUND_FOOTER,
        bordersAll: Settings.DEFAULT.BORDERS_ALL,
        bordersHorizontal: Settings.DEFAULT.BORDERS_HORIZONTAL,
        bordersVertical: Settings.DEFAULT.BORDERS_VERTICAL,
        bordersTitleBottom: Settings.DEFAULT.BORDERS_TITLE_BOTTOM,
        bordersHeadersBottom: Settings.DEFAULT.BORDERS_HEADERS_BOTTOM,
        bordersThickness: Settings.DEFAULT.BORDERS_THICKNESS,
        bordersHeadersVertical: Settings.DEFAULT.BORDERS_HEADERS_VERTICAL,
        helpControl: Settings.DEFAULT.HELP_CONTROL,
        debugControl: Settings.DEFAULT.DEBUG_CONTROL,
      },
      savedSettings,
    );
  }

  /**
   * Save the user's settings.
   *
   * @param {Object} settings - User settings to save.
   */
  export function updateSettingsForUser(settings) {
    cachedPropertiesForUser_().put(SETTINGS_KEY, settings);
  }

  /**
   * Deletes saved settings.
   */
  export function resetSettingsForUser() {
    cachedPropertiesForUser_().clear(SETTINGS_KEY);
  }
}

/**
 * Prototype object for cached access to script/user properties.
 */
var cachedPropertiesPrototype = {
  /**
   * Retrieve a saved property.
   *
   * @param {string} key - Key to lookup
   * @param {Object} defaultValue - Value to return if no value found in storage
   * @return {Object} retrieved value
   */
  get: function (key: string, defaultValue: object): object {
    var value = this.cache.get(key);
    if (!value) {
      value = this.properties.getProperty(key);
      if (value) {
        this.cache.put(key, value);
      }
    }
    if (value) {
      return JSON.parse(value);
    }
    return defaultValue;
  },

  /**
   * Saves a value to storage.
   *
   * @param key - Key to identify value
   * @param value - Value to save, will be serialized to JSON.
   */
  put: function (key: string, value: object) {
    var serializedValue = JSON.stringify(value);
    this.cache.remove(key);
    this.properties.setProperty(key, serializedValue);
  },

  /**
   * Deletes any saved settings.
   *
   * @param key - Key to identify value
   */
  clear: function (key: string) {
    this.cache.remove(key);
    this.properties.deleteProperty(key);
  },
};

/**
 * Gets a cached property instance for the current user.
 *
 * @return {CachedProperties}
 */
function cachedPropertiesForUser_() {
  return Object.assign(Object.create(cachedPropertiesPrototype), {
    properties: PropertiesService.getUserProperties(),
    cache: CacheService.getUserCache(),
  });
}

export { Settings };
