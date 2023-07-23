function unitTest_FullTest(sheetsToTest = undefined) {
  const options = {

    "New User Registration": {
      method: 'post',
      payload: {
        'entry.231462656': 'WiLliam',
        'entry.1494373819': 'veith',
        'entry.1061124162': 'Organics@1',
        'entry.1139611484': '978 7980710',
        'entry.1776831010': 'willIam.veith@gmail.cOm',
        'entry.1385700785': 'DR. BanerJee',
        'entry.1067143744': 'weV222'
      }
    },

    "Basket Assignment": {
      method: 'post',
      payload: {
        'entry.1665609710': 'weV222',
        'entry.1403771101': '9787980710',
        'entry.168334817': 'willIam.veith@gmail.cOm',
        'entry.1688326647': 'WiLliam',
        'entry.1438217424': 'veith',
        'entry.2097641953': 'Assign',
        'entry.828885758': 'S30'
      }
    }
  }

  const testParts = sheetsToTest === undefined ? Object.keys(options) : Array.isArray(sheetsToTest) ? sheetsToTest : [sheetsToTest];
  const sheetForms = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(sheet => {
      return {
        name: sheet.getName(),
        formUrl: sheet.getFormUrl(),
        postUrl: undefined
      };
    })
    .filter(sheetForm => sheetForm.formUrl !== null && testParts.includes(sheetForm.name))
    .map(sheetForm => {
      sheetForm.postUrl = FormApp.openByUrl(sheetForm.formUrl).getPublishedUrl().replace("viewform", "formResponse");
      return sheetForm;
    })

  sheetForms.forEach(form => {
    try {
      const response = UrlFetchApp.fetch(form.postUrl, options[form.name]);
      Logger.log(JSON.stringify({ "Status Code": response.getResponseCode() }));
    } catch (error) {
      console.log("Form:", form.name);
      console.log("Name:", error.name);
      console.log("Stack:", error.stack);
      console.log("message:", error.message);
    }
  })
}

function unitTest_basketRegistration() {
unitTest_FullTest("Basket Assignment")
//unitTest_FullTest("New User Registration")
}

/**
 * Tests the saveRecord function by optionally specifying the entry row.
 * @param {number} [entryRow=80] - The row number of the entry to test (default: 80).
 */
function testSaveRecord(entryRow = 80) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("New User Registration");
  const row = entryRow === undefined ? sheet.getLastRow() : entryRow;
  const range = sheet.getRange(`A${row}:I${row}`);
  const userInfoArray = range.getValues().flat();
  if (userInfoArray[8] !== "") {
    const error = {
      reason: `This row has already been done. Code should not have selected that row.`,
      rowNumber: `Row Number: ${row}`,
      rowInfo: `Row Info: ${userInfoArray}`
    };
    throw JSON.stringify(error);
  }
  const userInfoObj = {
    professor: userInfoArray[6],
    eid: userInfoArray[7].toLowerCase(),
    name: textTransformCapitalize(`${userInfoArray[1]} ${userInfoArray[2]}`),
    username: `${userInfoArray[1].replace(/[^a-zA-Z]/gm, "").toLowerCase()}_${userInfoArray[2].replace(/[^a-zA-Z]/, "").toLowerCase()}`,
    password: userInfoArray[3],
    phone: formatPhoneNumber(sheet, userInfoArray[4]),
    email: userInfoArray[5].toLowerCase(),
    activation: nextBusinessDay(userInfoArray[0]),
  };
  const record = JSON.stringify(userInfoObj);
  sheet.getRange(`I${row}`).setValue(record);
}

/**
 * Moves the records from the active sheet to a specified database file.
 */
function databasemover() {
  const file = DriveApp.getFileById("19DDG4AVg9AO1e-nyiYUYFKhcx4DK1xgU");
  const sheet = SpreadsheetApp.getActiveSheet();
  const records = sheet.getRange(`I2:I${sheet.getLastRow()}`).getValues().flat().map(row => JSON.parse(row));
  records.forEach(record => record.activation = new Date(record.activation).toISOString().slice(0, 19).replace('T', ' '));
  file.setContent(JSON.stringify(records));
}

function unitTestBasketAssignment() {
  // https://docs.google.com/forms/d/e/1FAIpQLSee_iUSVQwgWtCFxhe31tULJtHmI7qn7mC7UaC9T3j24aCcag/viewform?usp=pp_url&entry.1665609710=wev222&entry.168334817=williamveith@gmail.com&entry.1688326647=William&entry.1438217424=Veith&entry.2097641953=Return
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Basket Assignment");
  switch (sheet.getName()) {
    case "Basket Assignment":
      const values = getBasketValues(sheet);
      const qrCodeFile = generateBasketQRCode(values);
      sendBasketAssignmentEmail(values, qrCodeFile);
      formatBasketAssignmentSheet(sheet);
      sheet.sort(1, true);
      FormApp.openById(`1uHdYi31nuO0bd5e_CYIkrXz9gLqoUHorRvnt2Tzh9BM`).deleteAllResponses();
      break;
    case "New User Registration":
      sheet.sort(1, true);
      const userInfoObj = getUserObj(sheet);
      sendToCalendar(userInfoObj);
      sendText(userInfoObj);
      sendConfirmationEmail(userInfoObj);
      saveRecord(sheet, userInfoObj);
      formatUserRegistrationSheet(sheet);
      break;
  }
}

function unitTestImportBasketList() {
  const formData = {
    publishedUrl: "https://docs.google.com/forms/d/e/1FAIpQLSee_iUSVQwgWtCFxhe31tULJtHmI7qn7mC7UaC9T3j24aCcag/viewform",
    options: {
      method: 'post',
      payload: {
        'entry.1665609710': 'wev222',
        'entry.168334817': 'williamveith@utexas.edu',
        'entry.1688326647': 'William',
        'entry.1438217424': 'Veith',
        'entry.828885758': 'S30',
        'entry.2097641953': 'Assign'
      }
    }
  }
  const formUrl = formData.publishedUrl.replace("viewform", "formResponse");
  try {
    const response = UrlFetchApp.fetch(formUrl, formData.options);
    Logger.log(JSON.stringify({ "Status Code": response.getResponseCode() }));
  } catch (error) {
    console.log("Name:", error.name)
    console.log("Stack:", error.stack)
    console.log("message:", error.message)
  }
}

function unitTestSendEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Basket Assignment");
  sendBasketAssignmentEmail(sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn()).getValues().flat());
}

function unitTestQRCode() {
  const values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Basket Assignment").getRange("A3:G3").getValues().flat();
  const basketObj = {
    timestamp: values[0].getTime(),
    eid: values[1].toLowerCase(),
    email: values[2].toLowerCase(),
    firstName: textTransformCapitalize(values[3]),
    lastName: textTransformCapitalize(values[4]),
    basket: values[6],
    status: values[5]
  };
  basketObj.record = JSON.stringify(basketObj);
  generateBasketQRCode(basketObj)
}

function qrPDF() {
  const values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Basket Assignment").getRange("A4:I4").getValues().flat();

  const basketObj = {
    timestamp: values[0],
    eid: values[1].toLowerCase(),
    phone: formatPhoneNumber(values[2]),
    email: values[3].toLowerCase(),
    firstName: textTransformCapitalize(values[4]),
    lastName: textTransformCapitalize(values[5]),
    status: values[6],
    basket: values[7],
    qrCodeHash: (qrCodeDate) => {
      const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(qrCodeDate));
      return bytes.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('').toUpperCase();
    }
  };

  const qrCodeData = {
    eid: basketObj.eid,
    name: `${basketObj.firstName} ${basketObj.lastName}`,
    phone: basketObj.phone,
    email: basketObj.email,
    basket: basketObj.basket,
    assigned: `${basketObj.timestamp.getFullYear()}-${(basketObj.timestamp.getMonth() + 1).toString().padStart(2, "0")}-${(basketObj.timestamp.getDate()).toString().padStart(2, "0")}`,
    hash: (qrCodeDate) => {
      const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(qrCodeDate));
      return bytes.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('').toUpperCase();
    }
  }

  const test = qrCodeData.hash(qrCodeData)
  range.setValues([[basketObj.timestamp, basketObj.eid, basketObj.phone, basketObj.email, basketObj.firstName, basketObj.lastName, basketObj.status, basketObj.basket, basketObj.qrCodeHash(qrCodeData)]]);
  return [basketObj, qrCodeData];
}

function unitTestaddRichTextNewUserRegistration() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New User Registration");
  const range = sheet.getRange(`H2:H${sheet.getLastRow()}`);
  const values = range.getValues().flat()
  const richTexValues = values.map(value => {
    const richTextBuilder = SpreadsheetApp.newRichTextValue();
    richTextBuilder.setText(value);
    richTextBuilder.setLinkUrl(`https://utdirect.utexas.edu/webapps/eidlisting/eid_details?eid=${value}`);
    return [richTextBuilder.build()]
  });
  range.setRichTextValues(richTexValues)
}

function unitTestaddRichBasketRegistration() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Basket Assignment");
  const range = sheet.getRange(`B2:B${sheet.getLastRow()}`);
  const values = range.getValues().flat()
  const richTexValues = values.map(value => {
    const richTextBuilder = SpreadsheetApp.newRichTextValue();
    richTextBuilder.setText(value);
    richTextBuilder.setLinkUrl(`https://utdirect.utexas.edu/webapps/eidlisting/eid_details?eid=${value}`);
    return [richTextBuilder.build()]
  });
  range.setRichTextValues(richTexValues)
}