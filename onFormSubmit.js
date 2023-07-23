/**
 * Throws an error indicating that the current row has already been processed.
 * This function is used to prevent duplicate processing of rows in the sheet.
 *
 * @param {any[]} values - The values of the row that has already been processed.
 * @throws {string} - A JSON stringified error object with a reason and rowInfo properties.
 * @example
 * // Example usage:
 * const rowValues = [1, "John Doe", "john@example.com"];
 * rowAlreadyDone(rowValues);
 */
function rowAlreadyDone(values) {
  const error = {
    reason: "This row has already been done. Code should not have selected that row.",
    rowInfo: `Row Info: ${values}`,
  };
  throw JSON.stringify(error);
}

function onFormSubmission() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn());
  switch (sheet.getName()) {
    case "New User Registration":
      const userInfoObj = getUserObj(range);
      saveToCalendar(userInfoObj);
      sendText(userInfoObj);
      sendConfirmationEmail(userInfoObj);
      addRichTextUserRegistration(sheet, userInfoObj);
      formatUserRegistrationSheet(sheet);
      break;
    case "Basket Assignment":
      const [basketObj, qrCodeData] = getBasketObj(range);
      const qrCodeFile = basketObj.status === "Assign" ? generateBasketQRCode(qrCodeData) : undefined;
      sendBasketAssignmentEmail(basketObj, qrCodeFile);
      addRichTextBasketRegistration(sheet, basketObj);
      formatBasketAssignmentSheet(sheet);
      break;
  }
  FormApp.openByUrl(sheet.getFormUrl()).deleteAllResponses();
}

/**
 * Creates a user object from the given range data and stores it in a Google Sheets range.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range - A range of cells in Google Sheets.
 * @returns {Object} The user object with the following properties:
 *   - timestamp {string} - The timestamp from the range data (userInfoArray[0]).
 *   - firstName {string} - The capitalized first name from the range data (userInfoArray[1]).
 *   - lastName {string} - The capitalized last name from the range data (userInfoArray[2]).
 *   - password {string} - The password from the range data (userInfoArray[3]).
 *   - phone {string} - The formatted phone number from the range data (userInfoArray[4]).
 *   - email {string} - The lowercase email address from the range data (userInfoArray[5]).
 *   - professor {string} - The capitalized professor name from the range data (userInfoArray[6]).
 *   - eid {string} - The lowercase EID (Employee ID) from the range data (userInfoArray[7]).
 *   - name {string} - The capitalized full name derived from firstName and lastName.
 *   - username {string} - The username generated from the first and last name, with non-alphabetic characters removed.
 *   - activation {Date} - The next business day's date based on the timestamp (userInfoArray[0]).
 *   - record {function} - A function that returns the JSON string representation of the user object.
 *
 * @example
 * // Assuming the input range contains user information in columns A to I,
 * // and the output will be written to columns J to R.
 * const inputRange = sheet.getRange("A2:I2");
 * const userObject = getUserObj(inputRange);
 * // The userObject is created and stored in columns J to R based on the information in the input range.
 * console.log(userObject.firstName); // Output: "John"
 * console.log(userObject.lastName); // Output: "Doe"
 * console.log(userObject.username); // Output: "johndoe"
 */
function getUserObj(range) {
  const userInfoArray = range.getValues().flat();
  if (userInfoArray[8] !== "") {
    rowAlreadyDone(userInfoArray);
  }
  const userObj = {
    timestamp: userInfoArray[0],
    firstName: textTransformCapitalize(userInfoArray[1]),
    lastName: textTransformCapitalize(userInfoArray[2]),
    password: userInfoArray[3],
    phone: formatPhoneNumber(userInfoArray[4]),
    email: userInfoArray[5].toLowerCase(),
    professor: textTransformCapitalize(userInfoArray[6]),
    eid: userInfoArray[7].toLowerCase(),
    name: textTransformCapitalize(`${userInfoArray[1]} ${userInfoArray[2]}`),
    username: `${userInfoArray[1].replace(/[^a-zA-Z]/gm, "").toLowerCase()}_${userInfoArray[2].replace(/[^a-zA-Z]/g, "").toLowerCase()}`,
    activation: nextBusinessDay(userInfoArray[0]),
    record: (userObj) => { return JSON.stringify(userObj) }
  };
  const formattedValues = [[userObj.timestamp, userObj.firstName, userObj.lastName, userObj.password, userObj.phone, userObj.email, userObj.professor, userObj.eid, userObj.record(userObj)]];
  range.setValues(formattedValues);
  return userObj;
}

/**
 * Saves user information to a Google Calendar as an event.
 *
 * @param {Object} userInfoObj - The user information object.
 * @param {string} userInfoObj.name - The full name of the user.
 * @param {Date} userInfoObj.activation - The activation date for the user's account.
 *
 * @example
 * // Example user information object:
 * const userInformation = {
 *   name: "John Doe",
 *   activation: new Date("2023-07-25T12:00:00"),
 * };
 *
 * // Save the user information to the calendar.
 * saveToCalendar(userInformation);
 *
 * // This will create a new event on the calendar:
 * // Title: "Lab Access: Create Account | User: John Doe"
 * // Start: July 25, 2023 12:00:00 PM
 * // End: July 25, 2023 12:10:00 PM (10 minutes duration)
 * // Description: The event description generated from the "text message template" file
 * // Color: Gray
 * // Popup reminder: 15 minutes before the event.
 */
function saveToCalendar(userInfoObj) {
  const start = new Date(userInfoObj.activation);
  const end = new Date(start.getTime() + 10 * 60 * 1000);
  const title = `Lab Access: Create Account | User: ${userInfoObj.name}`;
  const description = (() => {
    let template = HtmlService.createTemplateFromFile("text message template");
    template.dynamicData = userInfoObj;
    return template.evaluate().getContent();
  })();
  CalendarApp.getCalendarById("williamveith@utexas.edu").createEvent(title, start, end, {
    description: description
  }).setColor(CalendarApp.EventColor.GRAY)
    .addPopupReminder(15);
}

/**
 * Adds a rich-text link to the user registration EID (Employee ID) on the specified sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Google Sheets sheet to add the rich-text link.
 * @param {Object} userInfoObj - The user information object returned from getUserObj function.
 * @param {string} userInfoObj.eid - The lowercase Employee ID (EID) of the user.
 *
 * @example
 * // Assuming the input range contains user information in columns A to I,
 * // and the output will be written to columns J to R in Google Sheets.
 * const inputRange = sheet.getRange("A2:I2");
 * const userObject = getUserObj(inputRange);
 * // The userObject is created and stored in columns J to R based on the information in the input range.
 * // The user information object can now be used to add a rich-text link to the EID on the specified sheet.
 * addRichTextUserRegistration(sheet, userObject);
 *
 * // This will add a rich-text link to the EID on the specified sheet:
 * // For example, if the EID is "abcd123":
 * // The cell in column H (last row of the sheet) will contain a hyperlink to
 * // "https://utdirect.utexas.edu/webapps/eidlisting/eid_details?eid=abcd123".
 */
function addRichTextUserRegistration(sheet, userInfoObj) {
  const hyperlinkEid = SpreadsheetApp.newRichTextValue()
    .setText(userInfoObj.eid)
    .setLinkUrl(`https://utdirect.utexas.edu/webapps/eidlisting/eid_details?eid=${userInfoObj.eid}`)
    .setTextStyle(defaultTextStyle())
    .build();
  sheet.getRange(`H${sheet.getLastRow()}`).setRichTextValue(hyperlinkEid);
}

/**
 * Formats the user registration sheet with specified font, font size, font color,
 * column widths, and date number format.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Google Sheets sheet to format.
 *
 * @example
 * // Assuming the user registration sheet is already defined as "sheet".
 * // Format the user registration sheet with the specified settings.
 * formatUserRegistrationSheet(sheet);
 */
function formatUserRegistrationSheet(sheet) {
  // Function implementation...
}

/**
 * Extracts user information from the given range data and creates a basket object.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range - A range of cells in Google Sheets containing user information.
 * @returns {Object} The basket object with the following properties:
 *   - timestamp {string} - The timestamp from the range data (values[0]).
 *   - eid {string} - The lowercase Employee ID (EID) from the range data (values[1]).
 *   - phone {string} - The formatted phone number from the range data (values[2]).
 *   - email {string} - The lowercase email address from the range data (values[3]).
 *   - firstName {string} - The capitalized first name from the range data (values[4]).
 *   - lastName {string} - The capitalized last name from the range data (values[5]).
 *   - status {string} - The status from the range data (values[6]).
 *   - basket {string} - The basket information from the range data (values[7]).
 *
 * @example
 * // Assuming the input range contains user information in columns A to H,
 * // and the output will be written to columns I to P in Google Sheets.
 * const inputRange = sheet.getRange("A2:H2");
 * const basketObject = getBasketObj(inputRange);
 * // The basketObject is created based on the information in the input range.
 * console.log(basketObject.firstName); // Output: "John"
 * console.log(basketObject.lastName); // Output: "Doe"
 */
function formatUserRegistrationSheet(sheet) {
  const minColumnWidths = [100, 100, 100, 100, 130, 130, 160, 100];
  sheet.getDataRange()
    .setFontFamily("Times New Roman")
    .setFontSize(12)
    .setFontColor("#000000");
  sheet.autoResizeColumns(1, 8);
  minColumnWidths.forEach((minWidth, index) => {
    if (sheet.getColumnWidth(index + 1) < minWidth) {
      sheet.setColumnWidth(index + 1, minWidth);
    }
  });
  ["A"].forEach(column => {
    sheet.getRange(`${column}${sheet.getLastRow()}`)
      .setNumberFormat("YYYY-MM-DD hh:mm:ss");
  });
}

/**
 * Extracts user information from the given range data and creates a basket object with associated QR code data.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range - A range of cells in Google Sheets containing user information.
 * @returns {[Object, Object]} An array containing the basket object and the QR code data object.
 *   - basketObj {Object} - The basket object with the following properties:
 *     - timestamp {string} - The timestamp from the range data (values[0]).
 *     - eid {string} - The lowercase Employee ID (EID) from the range data (values[1]).
 *     - phone {string} - The formatted phone number from the range data (values[2]).
 *     - email {string} - The lowercase email address from the range data (values[3]).
 *     - firstName {string} - The capitalized first name from the range data (values[4]).
 *     - lastName {string} - The capitalized last name from the range data (values[5]).
 *     - status {string} - The status from the range data (values[6]).
 *     - basket {string} - The basket information from the range data (values[7]).
 *   - qrCodeData {Object} - The QR code data object with the following properties:
 *     - eid {string} - The lowercase Employee ID (EID) from the basketObj.
 *     - name {string} - The full name (first name and last name) from the basketObj.
 *     - phone {string} - The phone number from the basketObj.
 *     - email {string} - The email address from the basketObj.
 *     - basket {string} - The basket information as a string from the basketObj.
 *     - assigned {string} - The formatted timestamp (YYYY-MM-DD) from the basketObj.
 *     - hash {function} - A function that returns the SHA-256 hash of the QR code data as a hexadecimal string.
 *
 * @example
 * // Assuming the input range contains user information in columns A to H,
 * // and the output will be written to columns I to P in Google Sheets.
 * const inputRange = sheet.getRange("A2:H2");
 * const [basketObject, qrCodeData] = getBasketObj(inputRange);
 * // The basketObject and qrCodeData are created based on the information in the input range.
 * console.log(basketObject.firstName); // Output: "John"
 * console.log(qrCodeData.email); // Output: "john.doe@example.com"
 */
function getBasketObj(range) {
  const values = range.getValues().flat();
  if (values[8] !== "") {
    rowAlreadyDone(values);
  }

  const basketObj = {
    timestamp: values[0],
    eid: values[1].toLowerCase(),
    phone: formatPhoneNumber(values[2]),
    email: values[3].toLowerCase(),
    firstName: textTransformCapitalize(values[4]),
    lastName: textTransformCapitalize(values[5]),
    status: values[6],
    basket: values[7]
  };

  const qrCodeData = {
    eid: basketObj.eid,
    name: `${basketObj.firstName} ${basketObj.lastName}`,
    phone: basketObj.phone,
    email: basketObj.email,
    basket: basketObj.basket.toString(),
    assigned: `${basketObj.timestamp.getFullYear()}-${(basketObj.timestamp.getMonth() + 1).toString().padStart(2, "0")}-${(basketObj.timestamp.getDate()).toString().padStart(2, "0")}`,
    hash: (qrCodeDate) => {
      const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(qrCodeDate));
      return bytes.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('').toUpperCase();
    }
  }

  range.setValues([[basketObj.timestamp, basketObj.eid, basketObj.phone, basketObj.email, basketObj.firstName, basketObj.lastName, basketObj.status, basketObj.basket, qrCodeData.hash(qrCodeData)]]);
  return [basketObj, qrCodeData];
}

/**
 * Generates a QR code for the given QR code data and creates a PDF file containing the QR code image.
 *
 * @param {Object} qrCodeData - The QR code data object containing information for generating the QR code.
 * @param {string} [size="255"] - The size of the QR code image (width and height) in pixels. Defaults to "255".
 * @returns {GoogleAppsScript.Base.Blob} The QR code as a PDF file in the form of a Blob.
 *
 * @example
 * // Example QR code data object:
 * const qrCodeData = {
 *   eid: "abcd123",
 *   name: "John Doe",
 *   phone: "123-456-7890",
 *   email: "john.doe@example.com",
 *   basket: "Basket1",
 *   assigned: "2023-07-25",
 *   hash: "abcdef0123456789", // The hash should be already calculated.
 * };
 *
 * // Generate a QR code PDF file using the provided QR code data.
 * const qrCodePdf = generateBasketQRCode(qrCodeData, "300");
 *
 * // Now you can use qrCodePdf to do further processing, such as sending it as an email attachment or saving it to Drive.
 */
function generateBasketQRCode(qrCodeData, size = "255") {
  const encodedData = encodeURIComponent(JSON.stringify(qrCodeData));
  const fileName = `${qrCodeData.basket.padEnd(5, " ")}${qrCodeData.assigned.padEnd(12, " ")}${qrCodeData.name}`;
  const qrCode = UrlFetchApp.fetch(`https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&data=${encodedData}`).getAs('image/png');
  const template = HtmlService.createTemplateFromFile("basket qr code");
  template.bytes = Utilities.base64Encode(qrCode.getBytes());
  template.basketNumber = qrCodeData.basket;
  const htmlContent = template.evaluate();
  const pdfFile = DriveApp.getFolderById("1Mc6T6Wlqs1BC300EM5uGpeSpLvihzSU3")
    .createFile(htmlContent.getBlob().getAs(MimeType.PDF))
    .setName(`${fileName}.pdf`)
    .setDescription(JSON.stringify(qrCodeData));
  return pdfFile.getAs(MimeType.PDF);
}

/**
 * Adds a rich-text link to the basket registration EID (Employee ID) on the specified sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Google Sheets sheet to add the rich-text link.
 * @param {Object} userInfoObj - The user information object returned from getBasketObj function.
 * @param {string} userInfoObj.eid - The lowercase Employee ID (EID) of the user.
 *
 * @example
 * // Assuming the user information object with the basket data is already defined as "userInfo".
 * // Add a rich-text link to the basket registration EID on the specified sheet.
 * addRichTextBasketRegistration(sheet, userInfo);
 *
 * // This will add a rich-text link to the EID on the specified sheet:
 * // For example, if the EID is "abcd123":
 * // The cell in column B (last row of the sheet) will contain a hyperlink to
 * // "https://utdirect.utexas.edu/webapps/eidlisting/eid_details?eid=abcd123".
 */
function addRichTextBasketRegistration(sheet, userInfoObj) {
  const hyperlinkEid = SpreadsheetApp.newRichTextValue()
    .setText(userInfoObj.eid)
    .setLinkUrl(`https://utdirect.utexas.edu/webapps/eidlisting/eid_details?eid=${userInfoObj.eid}`)
    .setTextStyle(defaultTextStyle())
    .build();
  sheet.getRange(`B${sheet.getLastRow()}`).setRichTextValue(hyperlinkEid);
}

/**
 * Formats the basket assignment sheet with specified font, font size, font color,
 * column widths, and date/number formats.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Google Sheets sheet to format.
 *
 * @example
 * // Assuming the basket assignment sheet is already defined as "sheet".
 * // Format the basket assignment sheet with the specified settings.
 * formatBasketAssignmentSheet(sheet);
 */
function formatBasketAssignmentSheet(sheet) {
  const minColumnWidths = [100, 100, 130, 100, 100, 150, 150];
  sheet.getDataRange()
    .setFontFamily("Times New Roman")
    .setFontSize(12)
    .setFontColor("#000000");
  sheet.getRange("A2:A").setNumberFormat("YYYY-MM-DD hh:mm:ss");
  sheet.getRange("B:I").setNumberFormat("@");
  sheet.autoResizeColumns(1, 9);
  minColumnWidths.forEach((minWidth, index) => {
    if (sheet.getColumnWidth(index + 1) < minWidth) {
      sheet.setColumnWidth(index + 1, minWidth);
    }
  });
}