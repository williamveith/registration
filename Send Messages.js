/**
 * The message configuration object that defines settings for different types of messages.
 *
 * @typedef {Object} MessageConfigs
 *
 * @property {Object} Lab Access Account Text - Configuration for Lab Access Account Text messages.
 * @property {string} Lab Access Account Text.displayName - The display name for the Lab Access Account Text message.
 * @property {string} Lab Access Account Text.sendAs - The sender email address for the Lab Access Account Text message.
 * @property {string} Lab Access Account Text.sendTo - The recipient phone number for the Lab Access Account Text message.
 * @property {string} Lab Access Account Text.subject - The subject for the Lab Access Account Text message.
 * @property {string} Lab Access Account Text.bodyTemplateName - The template name for the body of the Lab Access Account Text message.
 *
 * @property {Object} Lab Access Account Email - Configuration for Lab Access Account Email messages.
 * @property {string} Lab Access Account Email.displayName - The display name for the Lab Access Account Email message.
 * @property {string} Lab Access Account Email.sendAs - The sender email address for the Lab Access Account Email message.
 * @property {string} Lab Access Account Email.subject - The subject for the Lab Access Account Email message.
 * @property {string} Lab Access Account Email.bodyTemplateName - The template name for the body of the Lab Access Account Email message.
 *
 * @property {Object} Basket Assignment Email - Configuration for Basket Assignment Email messages.
 * @property {string} Basket Assignment Email.displayName - The display name for the Basket Assignment Email message.
 * @property {string} Basket Assignment Email.sendAs - The sender email address for the Basket Assignment Email message.
 * @property {string} Basket Assignment Email.bodyTemplateName - The template name for the body of the Basket Assignment Email message.
 *
 * @property {function} getConfigs - A function to get the message configuration by providing a configuration key.
 * @param {string} configKey - The configuration key to retrieve the specific message configuration.
 * @returns {Object} The configuration object for the given configuration key.
 *
 * @property {function} getBody - A function to generate the message body by providing the configuration and dynamic data.
 * @param {Object} configs - The message configuration object (e.g., Lab Access Account Text or Lab Access Account Email).
 * @param {Object} dynamicData - The dynamic data to be used for populating the message template.
 * @returns {string} The message body content generated from the provided template and dynamic data.
 */
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

/**
 * Sends a text message to a recipient based on the provided user information object.
 *
 * @param {Object} userInfoObj - The user information object containing the necessary data for sending the text message.
 * @param {string} userInfoObj.eid - The lowercase Employee ID (EID) of the user.
 * @param {string} userInfoObj.firstName - The first name of the user.
 * @param {string} userInfoObj.lastName - The last name of the user.
 * @param {string} userInfoObj.activation - The activation date for the user's account.
 *
 * @example
 * // Assuming the user information object is already defined as "userInfo".
 * // Send a text message based on the provided user information.
 * sendText(userInfo);
 *
 * // The text message will be sent to the recipient specified in the configuration.
 */
function sendText(userInfoObj) {
  const configs = messageConfigs.getConfigs("Lab Access Account Text");
  GmailApp.sendEmail(configs.sendTo, configs.subject, "", {
    from: configs.sendAs,
    htmlBody: messageConfigs.getBody(configs, userInfoObj),
    name: configs.displayName
  });
}

/**
 * Sends a confirmation email to the user's email address based on the provided user information object.
 *
 * @param {Object} userInfoObj - The user information object containing the necessary data for sending the confirmation email.
 * @param {string} userInfoObj.eid - The lowercase Employee ID (EID) of the user.
 * @param {string} userInfoObj.firstName - The first name of the user.
 * @param {string} userInfoObj.lastName - The last name of the user.
 * @param {string} userInfoObj.email - The email address of the user to which the confirmation email will be sent.
 * @param {string} userInfoObj.activation - The activation date for the user's account.
 *
 * @example
 * // Assuming the user information object is already defined as "userInfo".
 * // Send a confirmation email to the user's email address.
 * sendConfirmationEmail(userInfo);
 *
 * // The confirmation email will be sent to the user's email address specified in the configuration.
 */
function sendConfirmationEmail(userInfoObj) {
  const configs = messageConfigs.getConfigs("Lab Access Account Email");
  GmailApp.sendEmail(userInfoObj.email, configs.subject, "", {
    from: configs.sendAs,
    htmlBody: messageConfigs.getBody(configs, userInfoObj),
    name: configs.displayName
  });
}

/**
 * Sends a basket assignment email to the user's email address based on the provided basket object.
 *
 * @param {Object} basketObj - The basket object containing the necessary data for sending the assignment email.
 * @param {string} basketObj.basket - The name or ID of the basket.
 * @param {string} basketObj.status - The status of the basket assignment ("Assign" or "Return").
 * @param {string} basketObj.firstName - The first name of the user to whom the basket is assigned/returned.
 * @param {string} basketObj.lastName - The last name of the user to whom the basket is assigned/returned.
 * @param {string} basketObj.email - The email address of the user to whom the basket is assigned/returned.
 *
 * @param {GoogleAppsScript.Base.Blob|undefined} [qrCodeFile=undefined] - The QR code file (Blob) to attach to the email.
 * If not provided, no attachment will be added.
 *
 * @example
 * // Assuming the basket object and QR code file (Blob) are already defined as "basket" and "qrCodeFile".
 * // Send a basket assignment email to the user's email address.
 * sendBasketAssignmentEmail(basket, qrCodeFile);
 *
 * // The basket assignment email will be sent to the user's email address specified in the configuration.
 * // If a QR code file (Blob) is provided, it will be attached to the email.
 */
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
