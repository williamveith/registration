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

function addRichTextUserRegistration(sheet, userInfoObj) {
  const hyperlinkEid = SpreadsheetApp.newRichTextValue()
    .setText(userInfoObj.eid)
    .setLinkUrl(`https://utdirect.utexas.edu/webapps/eidlisting/eid_details?eid=${userInfoObj.eid}`)
    .setTextStyle(defaultTextStyle())
    .build();
  sheet.getRange(`H${sheet.getLastRow()}`).setRichTextValue(hyperlinkEid);
}

function formatUserRegistrationSheet(sheet, userInfoObj) {
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

function addRichTextBasketRegistration(sheet, userInfoObj) {
  const hyperlinkEid = SpreadsheetApp.newRichTextValue()
    .setText(userInfoObj.eid)
    .setLinkUrl(`https://utdirect.utexas.edu/webapps/eidlisting/eid_details?eid=${userInfoObj.eid}`)
    .setTextStyle(defaultTextStyle())
    .build();
  sheet.getRange(`B${sheet.getLastRow()}`).setRichTextValue(hyperlinkEid);
}


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