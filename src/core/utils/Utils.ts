import { HSLType, RGBType, TimeItResponse } from '@core/types/addon';

/**
 * Collection of utility functions to use with the add-on
 *
 * @namespace
 */
namespace Utils {
  /** Display's toast in lower right corner of spreadsheet
   * @param {string} message the main message to display
   * @param {string} [title='GV App'] the title of the toast. Default is 'GV App'
   * @param {number} [timeoutSeconds=15] how long to display the toast. (default: 15 sec)
   */
  export function toastIt(
    message: string,
    title: string = 'GV App',
    timeoutSeconds: number = 15,
  ) {
    g.ss.toast(message, title, timeoutSeconds);
  }

  /** Display alert box in center of screen
   * @param {string} message the message to display
   * @param {string} [title='Oops!'] the title of the alert box
   * @return {GoogleAppsScript.Base.Button}
   */
  export function alertIt(
    message: string,
    title: string = 'Oops!',
  ): GoogleAppsScript.Base.Button {
    const ui = SpreadsheetApp.getUi();
    return ui.alert(title, message, ui.ButtonSet.OK);
  }

  /** Prompt user with a question to answer
   * @param {string} question the question to ask
   * @param {string} [title] the title of the prompt box
   * @return {string} the answer to the question
   */
  export function askQuestion(question: string, title: string = ''): string {
    const ui = SpreadsheetApp.getUi();
    const result = ui.prompt(title, question, ui.ButtonSet.OK_CANCEL);
    return result.getResponseText();
  }

  /** Displays a dialog box
   * @param {GoogleAppsScript.HTML.HtmlOutput} html HtmlOutput to display
   * @param {string} title title of dialog box
   * @param {number} [width=900] width of dialog box
   * @param {number} [height=600] height of dialog box
   */
  export function displayIt(
    html: GoogleAppsScript.HTML.HtmlOutput,
    title: string,
    width: number = 900,
    height: number = 600,
  ): void {
    const ui = SpreadsheetApp.getUi();
    ui.showModelessDialog(html.setWidth(width).setHeight(height), title);
  }

  /** Display an error message and finalizes process logger
   * @param {string} message
   * @param {string} title
   */
  export function displayError(message: string, title?: string) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title || 'Oops!', message, ui.ButtonSet.OK);
  }

  /** execute a function and time how long it takes
   * @param {function} func the thing to execute
   * @return {TimeItResponse} the result and stats
   */
  export function timeIt(func: Function, ...args: any[]): TimeItResponse {
    const startedAt = new Date().getTime();
    let result = null;
    let error = null;
    try {
      result = func(...args);
    } catch (err) {
      error = err;
    }
    const finishedAt = new Date().getTime();
    const elapsed = finishedAt - startedAt;
    const summary = {
      startedAt,
      finishedAt,
      result,
      error,
      elapsed,
    };

    console.log('summary :>> ', summary);
    return summary;
  }

  export const round = (val: number, dec = 2) => Number(val.toFixed(dec));

  export function titleCaseAndRemoveSpaces(str?: string): string {
    if (!str) return '';

    // Convert the string to lowercase and split by spaces
    let words = str.replace('#', 'Number').toLowerCase().split(' ');

    // Capitalize the first letter of each word
    words = words.map((word) => word.charAt(0).toUpperCase() + word.slice(1));

    // Join the words together without spaces
    let result = words.join('');

    return result;
  }

  export function hexToRgb(hex: string): RGBType {
    var m = hex.slice(1).match(hex.length == 7 ? /(\S{2})/g : /(\S{1})/g);
    if (m)
      return {
        r: parseInt(m[0], 16),
        g: parseInt(m[1], 16),
        b: parseInt(m[2], 16),
      };
  }

  export function hslToHex(hsl: HSLType): string {
    const { h, s, l } = hsl;

    const hDecimal = l / 100;
    const a = (s * Math.min(hDecimal, 1 - hDecimal)) / 100;
    const f = (n: number) => {
      const k = (n + h / 30) % 12;
      const color = hDecimal - a * Math.max(Math.min(k - 3, 9 - k, 1), -1);

      // Convert to Hex and prefix with "0" if required
      return Math.round(255 * color)
        .toString(16)
        .padStart(2, '0');
    };
    return `#${f(0)}${f(8)}${f(4)}`;
  }

  /**
   *
   * @param {RGBType} rgb The color components
   * @returns {HSLType}
   */
  export function rgbToHsl(rgb: RGBType): HSLType {
    let { r, g, b } = rgb;
    r /= 255;
    g /= 255;
    b /= 255;

    var max = Math.max(r, g, b),
      min = Math.min(r, g, b);
    var h,
      s,
      l = (max + min) / 2;

    if (max == min) {
      h = s = 0; // achromatic
    } else {
      var d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

      switch (max) {
        case r:
          h = (g - b) / d + (g < b ? 6 : 0);
          break;
        case g:
          h = (b - r) / d + 2;
          break;
        case b:
          h = (r - g) / d + 4;
          break;
      }

      h /= 6;
    }

    return { h: h * 360, s: s * 100, l: l * 100 };
  }

  export function niceColor(hex: string) {
    // assumes "rgb(R,G,B)" string
    const rgb = hexToRgb(hex);
    console.log('rgb :>> ', rgb);
    let hsl = rgbToHsl(rgb);
    console.log('hsl :>> ', hsl);
    let { h, s, l } = hsl;
    h = (h / 360 + 0.5) % 1; // Hue
    s = (s / 100 + 0.5) % 1; // Saturation
    l = (l / 100 + 0.5) % 1; // Luminocity
    return hslToHex({ h: h * 360, s: s * 100, l: l * 100 });
  }

  /**
   * Calculate brightness value by RGB or HEX color.
   * @param color (String) The color value in RGB or HEX (for example: #000000 || #000 || rgb(0,0,0) || rgba(0,0,0,0))
   * @returns (Number) The brightness value (dark) 0 ... 255 (light)
   */
  export function brightnessByColor(color: string) {
    var color = '' + color,
      isHEX = color.indexOf('#') == 0,
      isRGB = color.indexOf('rgb') == 0;
    let r, g, b;
    if (isHEX) {
      var m = color.slice(1).match(color.length == 7 ? /(\S{2})/g : /(\S{1})/g);
      if (m) {
        r = parseInt(m[0], 16);
        g = parseInt(m[1], 16);
        b = parseInt(m[2], 16);
      }
    }
    if (isRGB) {
      var m = color.match(/(\d+){3}/g);
      if (m) {
        r = m[0];
        g = m[1];
        b = m[2];
      }
    }
    if (typeof r != 'undefined') return (r * 299 + g * 587 + b * 114) / 1000;
  }

  export function groupText(str: string): string[] {
    let result = [];
    let stack = [];
    let current = '';

    for (let i = 0; i < str.length; i++) {
      let char = str[i];

      if (char === '(' || char === '[' || char === '{') {
        if (stack.length === 0) {
          current += char;
          result.push(current);
          current = '';
        } else {
          current += char;
        }
        stack.push(char);
      } else if (char === ')' || char === ']' || char === '}') {
        if (stack.length === 1 && current.length) {
          result.push(current);
          current = char;
        } else {
          current += char;
        }
        stack.pop();
      } else if (char === ',') {
        if (stack.length === 1) {
          result.push(current);
          result.push(char);
          current = '';
        } else {
          current += char;
        }
      } else {
        current += char;
      }
    }

    result.push(current);
    return result;
  }
}

export { Utils };
