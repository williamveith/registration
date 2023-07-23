/**
 * Calculates the next business day from the provided date or the current date if not provided.
 *
 * @param {Date} [today=new Date()] - The date to start calculating from. Defaults to the current date.
 * @returns {string} The next business day in string format (formatted as "Day Month Date Year HH:mm:ss GMT±nnnn (Timezone)").
 *
 * @example
 * // Calculate the next business day from the current date.
 * const nextBusinessDay = nextBusinessDay();
 * Logger.log(nextBusinessDay);
 *
 * @example
 * // Calculate the next business day from a specific date (e.g., "2023-07-25").
 * const specificDate = new Date("2023-07-25");
 * const nextBusinessDay = nextBusinessDay(specificDate);
 * Logger.log(nextBusinessDay);
 */
function nextBusinessDay(today = new Date()) {
  const daysLeftInWeek = 7 - today.getDay();
  const activeDay = daysLeftInWeek > 2 ? today.setDate(today.getDate() + 1) : today.setDate(today.getDate() + (daysLeftInWeek + 1));
  const nextBusinessDate = new Date(new Date(activeDay).setHours(13, 0, 0));
  return nextBusinessDate.toString();
}

/**
 * Formats a phone number into a standardized format: (XXX) XXX-XXXX.
 *
 * @param {string|number} number - The phone number to be formatted.
 * @returns {string} The formatted phone number.
 *
 * @example
 * const phoneNumber = "1234567890";
 * const formattedPhoneNumber = formatPhoneNumber(phoneNumber);
 * Logger.log(formattedPhoneNumber); // Outputs: (123) 456-7890
 */
function formatPhoneNumber(number) {
  const phoneNumberString = number.toString();
  const match = phoneNumberString.replace(/\D/g, '').match(/^(1|)?(\d{3})(\d{3})(\d{4})$/);
  const formattedPhoneNumber = `${match[1] ? '+1 ' : ''}(${match[2]}) ${match[3]}-${match[4]}`;
  return formattedPhoneNumber;
}

/**
 * Transforms a string into title case, capitalizing the first letter of each word and making the rest lowercase.
 *
 * @param {string} str - The input string to be transformed.
 * @returns {string} The transformed string in title case.
 *
 * @example
 * const inputText = "hello woRLD";
 * const transformedText = textTransformCapitalize(inputText);
 * Logger.log(transformedText); // Outputs: "Hello World"
 */
function textTransformCapitalize(str) {
  const nameParts = str.split(" ");
  const transformedText = nameParts.map(name => `${name.charAt(0).toUpperCase()}${name.slice(1).toLowerCase()}`).join(" ");
  return transformedText;
}

/**
 * Creates and returns a default text style for Google Sheets.
 *
 * @returns {GoogleAppsScript.Spreadsheet.TextStyle} The default text style.
 *
 * @example
 * const defaultStyle = defaultTextStyle();
 * Logger.log(defaultStyle.getFontFamily()); // Outputs: "Times New Roman"
 */
function defaultTextStyle() {
  return SpreadsheetApp.newTextStyle()
    .setUnderline(false)
    .setFontFamily("Times New Roman")
    .setForegroundColor("#000000")
    .setFontSize(12)
    .build();
}