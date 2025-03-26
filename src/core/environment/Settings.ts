import { SettingsOptions } from '@core/types/addon';

namespace Settings {
  export const DEBUG = true;
  const SETTINGS_KEY = 'settings';
  export const APP_TITLE = 'SC GAS BOILERPLATE';
  export const SPREADSHEET_ICON_FOLDER_ID = '1pC5a4ehrnLL5_McnQPN5UXm1gQ2hVf9R';
  export const DEFAULT = {
    hasTitle: false,
    hasHeaders: true,
    hasFooter: false,
    leaveTop: false,
    leaveLeft: false,
    leaveBottom: false,
    noBottom: false,
    centerAll: true,
    alternating: false,
    backgroundTitle: 6,
    backgroundHeaders: 14,
    backgroundDataFirst: 19,
    backgroundDataSecond: 16,
    backgroundFooter: 13,
    bordersAll: 6,
    bordersHorizontal: 10,
    bordersVertical: 10,
    bordersTitleBottom: 6,
    bordersHeadersBottom: 9,
    bordersThickness: 1,
    bordersHeadersVertical: 10,
    helpControl: 'off',
    themeControl: 'off',
    debugControl: 'on',
    customThemes: {
      theme1: null,
      theme2: null,
      theme3: null,
      theme4: null,
      theme5: null,
      theme6: null,
    },
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
    const savedSettings = cachedPropertiesForUser_().get(SETTINGS_KEY, {});
    const settings = Object.assign({}, Settings.DEFAULT, savedSettings);
    return settings;
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
