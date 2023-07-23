const messageConfigs = {
  "Lab Access Account Text": {
    displayName: "New User Setup",
    sendAs: "williamveith@utexas.edu",
    sendTo: "9787980710@mms.att.net",
    subject: "New User Setup",
    bodyTemplateName: "text message template"
  },
  "Lab Access Account Email": {
    displayName: "Lab Access Automated Message",
    sendAs: "williamveith@utexas.edu",
    subject: "Confirmation: UT MRC Equipment Access User Authorization Form Received",
    bodyTemplateName: "email account creation"
  },
  "Basket Assignment Email": {
    displayName: "Automated Basket Assignment",
    sendAs: "williamveith@utexas.edu",
    bodyTemplateName: "email basket assignment"
  },
  getConfigs(configKey) {
    return this[configKey];
  },
  getBody(configs, dynamicData) {
    let template = HtmlService.createTemplateFromFile(configs.bodyTemplateName);
    template.dynamicData = dynamicData;
    return template.evaluate().getContent();
  }
}

function sendText(userInfoObj) {
  const configs = messageConfigs.getConfigs("Lab Access Account Text");
  GmailApp.sendEmail(configs.sendTo, configs.subject, "", {
    from: configs.sendAs,
    htmlBody: messageConfigs.getBody(configs, userInfoObj),
    name: configs.displayName
  });
}

function sendConfirmationEmail(userInfoObj) {
  const configs = messageConfigs.getConfigs("Lab Access Account Email");
  GmailApp.sendEmail(userInfoObj.email, configs.subject, "", {
    from: configs.sendAs,
    htmlBody: messageConfigs.getBody(configs, userInfoObj),
    name: configs.displayName
  });
}

function sendBasketAssignmentEmail(basketObj, qrCodeFile = undefined) {
  const configs = messageConfigs.getConfigs("Basket Assignment Email");
  const dynamicData = {
    basket: basketObj.basket,
    status: basketObj.status === "Assign" ? "Assigned To You" : "Returned",
    name: `${basketObj.firstName} ${basketObj.lastName}`,
    message: basketObj.status === "Assign" ? `You have been assigned cleanroom basket ${basketObj.basket}. The QR code for your basket is attached below. Print it and place it inside the basket tag holder` : `The cleanroom basket you were assigned, ${basketObj.basket}, has been returned.`
  };
  GmailApp.sendEmail(basketObj.email, `Cleanroom Basket ${basketObj.status}ed`, "", {
    from: configs.sendAs,
    htmlBody: messageConfigs.getBody(configs, dynamicData),
    attachments: qrCodeFile === undefined ? [] : [qrCodeFile],
    name: configs.displayName
  });
}
