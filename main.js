/*
 * Email merging tools developed for use at Icograda.
 *
 * @author F. Gabriel Gosselin <fggosselin@icograda.org>, <gabriel@evidens.ca>
 * @website https://github.com/beyondwords/gamerge
 */

var STATUS_SENT = 'Sent',
    STATUS_IGNORE = 'Ignore',
    STATUS_ERROR = 'Error',
    IGNORE_STATUSES = [STATUS_SENT, STATUS_ERROR, STATUS_IGNORE],
    CONTACTS_SHEET      = 'Contact List',
    TEST_CONTACTS_SHEET = 'Test Contact List',
    // Attachments types for appending files to email
    FILE_ATTACH         = 'attached',
    FILE_INLINE         = 'inline';

function onOpen() {
  var mySheet = SpreadsheetApp.getActiveSpreadsheet(),
      menuEntries = [
        // {name: 'Import HTML email', functionName: 'menuImportHTML' },
        // null, // Separator
        {name: "Step 1: Test mail template", functionName: "menuTestTemplate"},
        {name: "Step 2: Send Test Mail", functionName: "menuSendTestEmail"},
        {name: "Step 3: Start Mail Merge", functionName: "menuSendEmail"},
        null, // Separator
        {name: "Check remaining emails", functionName: "menuShowRemainingEmails"},
        {name: "Help / About", functionName: "menuShowHelp"}];
  mySheet.addMenu("Mail Merge", menuEntries);
}

/* ****************************************************************************
 * Internal Functions
 *****************************************************************************/

/*
 * Checks the remaining emails available, and offsets for the numbers of emails
 * you want reserve for the rest of the day.
 * Google will LOCK YOUR ACCOUNT for 24 hours if you exceed this quota!
 * Assumes 'EmailQuotaReserve' data range is defined and correct
 *
 * Ex. 500 emails lefts, reserve of 50, Mail Merge will send a maximum of 450 emails
 *     allowing you to send up to 50 more emails without being locked out of GMail.
 */
function getRemainingQuota(aSpread) {
  var remainingQuota = MailApp.getRemainingDailyQuota(),
      reserveAmount = aSpread.getRangeByName("EmailQuotaReserve").getValue();

  return (remainingQuota - reserveAmount > 0) ? remainingQuota - reserveAmount : 0;
}

/*
 * Grabs the body of the email from the template tab.
 * Assumes 'EmailBody' data range is defined and correct
 */
function getMailBody(activeSpread, newLine) {
  var mailBody = "",
      value, values, i,
      range = activeSpread.getRangeByName("EmailBody");

  if (range !== null) {
    Logger.log("Num rows found " + range.getNumRows());
    values = range.getValues();
    for (i=0; i < values.length; i += 1) {
      value = values[i][0];
      if (value.length > 0) {
        mailBody += value + newLine;
      }
    }
  } else {
    Logger.log("Nothing returned");
  }
  return mailBody;
}

/*
 * Replaces <br> HTML tags with line returns, strips all other tags
 */
function convertToPlainText (html) {
  var plain = html.replace(/<br ?\/?>|<\/p>/g, "\n");
  plain = plain.replace(/<.*?>/g, '');
  return plain;
}

/*
 * Grabs the elements of the email template
 * Assumes 'EmailBody', 'EmailSenderName', 'EmailReplyTo', 'EmailSubject'
 * data ranges are defined and correct
 */
function getEmailData(activeSpread) {
  var newLine = "<br><br>",
      plainText = activeSpread.getRangeByName('EmailPlaintext').getValue(),
      email = {};

  email.senderName = activeSpread.getRangeByName("EmailSenderName").getValue();
  email.replyTo    = activeSpread.getRangeByName("EmailReplyTo").getValue();
  email.subject    = activeSpread.getRangeByName("EmailSubject").getValue();
  email.body       = getMailBody(activeSpread, newLine);
  // Allow plaintext override
  if (plainText && plainText.length > 0) {
    email.plainText = plainText;
  } else {
    email.plainText  = convertToPlainText(email.body);
  }

  return email;
}

/*
 * Makes a hash table of Column headers from the 'Contact list'
 * Ex. {...'Email': 4,...}
 */
function columnHeadersKeys (headers) {
  var keys = {}, i;
  for (i=0; i < headers.length; i += 1) {
    keys[headers[i]] = i;
  }
  return keys;
}

/*
 * Makes a hash table of tags found in the email template body
 * against the columns where the data will be found
 * Ex. {...'First Name': '{{First Name}}',...}
 *
 * See mergeTemplate
 */
function makeTagSearchList (mailBody) {
  var allTags = mailBody.match(/\{\{([\w ]*?)\}\}/g),
      searchList = {},
      tag, colName, i;

  if (allTags === null) {
    Logger.log("No tags found in mail body.");
    return null;
  }

  Logger.log("Found " + allTags.length + " tags in mail body.");
  // Filter down to one occurence of each
  for (i = allTags.length - 1; i >= 0; i -= 1) {
    tag = allTags[i];
    colName = tag.substr(2, tag.length - 4); // Remove tags

    if (!searchList.hasOwnProperty(colName)) {
      searchList[colName] = tag;
    }
  }

  return searchList;
}

/*
 * Merges the given row of the 'Contact List' sheet into the
 * email body template according to the tags.
 */
function mergeTemplate (bodyTemplate, searchList, data, keys) {
  var name, key, pattern, field,
      output = bodyTemplate;

  for (name in searchList) {
    key = keys[name];
    field = data[key];
    pattern = searchList[name];

    output = output.replace(pattern, field, 'gm');
  }

  return output;
}

/*
 * Find any tags in the email template that don't match
 * the data headings available (Case sensitive)
 */
function findWrongTags(searchList, colKeys) {
  var wrongTags = '';
  for (searchName in searchList) {
    if (!colKeys.hasOwnProperty(searchName)) {
      wrongTags += searchList[searchName] + ", \n";
    }
  }

  return wrongTags;
}

/*
 * Retrieves attachments specified in the template sheet.
 *
 * @return object properties formatted for Mail advancedArgs
 */
function prepareAttachments(aSpread) {
  var attachList = aSpread.getRangeByName("EmailAttachments").getValues(),
      attachItem, i, attachName, attachMode, attachURL, fileType, attachContent,
      inlineImages = {}, attachments = [], advancedArgs = {};

  for (i=0; i < attachList.length; i+=1) {
    attachItem = attachList[i];
    attachName = attachItem[0];
    attachMode = attachItem[1];
    attachURL = attachItem[2];
    fileType = attachItem[3];
    if (attachURL && attachURL.length > 0) {
      attachContent = UrlFetchApp.fetch(attachURL).getContent();
      Logger.log(attachName + ': ' + attachMode + ', ' + attachURL + ', ' + fileType);
    } else {
      Logger.log('Skipping ' + attachName + '. No URL.');
      continue;
    }
    // Add as either a regular attachment or an inline image
    if (attachMode.indexOf(FILE_ATTACH) !== -1) {
      attachments.push({fileName: attachName, mimeType: fileType,content: attachContent});
    } else if (attachMode.indexOf(FILE_INLINE) !== -1) {
      inlineImages[attachName] = attachContent;
    }
  }

  if (inlineImages != {}) {
    advancedArgs.inlineImages = inlineImages;
  }
  if (attachments.length > 0) {
    advancedArgs.attachments = attachments;
  }
  return advancedArgs;
}

/*
 * Retrieves contacts and email template
 * Cycles through all contacts and sends an individualised email to
 * each recipient according to the content and merge tags specified
 *
 */
function sendEmailToList (listName) {

  var aSpread = SpreadsheetApp.getActiveSpreadsheet(),
      contactSheet = aSpread.getSheetByName(listName),
      template = getEmailData(aSpread),
      mergeData = contactSheet.getDataRange().getValues(),
      searchList = makeTagSearchList(template.body),
      colKeys = columnHeadersKeys(mergeData[0]),
      statusCol = colKeys['Status'],
      emailCol = colKeys['Email'],
      wrongTags = findWrongTags(searchList, colKeys),
      remainingQuota = getRemainingQuota(aSpread),
      advancedArgs = prepareAttachments(aSpread),
      htmlMsg, plainTextMsg, row, status, address, i, numSent = 0, errLog = '';

  if (wrongTags.length > 0) {
    Browser.msgBox("These tags don't have matching data columns: " + wrongTags);
    return null;
  }

  for (i=1; i < mergeData.length; i+=1) {
    if (remainingQuota - numSent <= 0) {
      break;
    }
    row = mergeData[i];
    status = row[statusCol];
    address = row[emailCol];
    if (address !== "" && IGNORE_STATUSES.indexOf(status) === -1) {
      // Create customised email
      htmlMsg = mergeTemplate(template.body, searchList, row, colKeys);
      plainTextMsg = mergeTemplate(template.plainText, searchList, row, colKeys);
      // Prep email
      advancedArgs.htmlBody = htmlMsg;
      advancedArgs.name     = template.senderName;
      advancedArgs.replyTo  = template.replyTo;

      // Send email, error otherwise
      try {
        MailApp.sendEmail(address, template.subject, plainTextMsg , advancedArgs);
        // Set sent status
        contactSheet.getRange(i+1,statusCol+1).setValue(SENT_STATUS);
        numSent += 1;
      } catch (e) {
        errLog += e + "\n";
        Logger.log(e);
        contactSheet.getRange(i+1,statusCol+1).setValue('Error');
      }
    }
  }

  // Log errors
  var sentMessage = "Sent " + numSent + " emails.\n";
  Logger.log(sentMessage);
  if (errLog) {
    sentMessage += "---Errors---\n\n" + errLog;
  }
  Browser.msgBox( sentMessage );
  // Clean up
  SpreadsheetApp.flush();
}

function captureURL () {
  return '';
}

/* ****************************************************************************
 * Menu functions
 *****************************************************************************/

function menuImportHTML () {
  var importURL = Browser.inputBox("Import HTML Email", "Please specify the URL of the HTML email to import. (The URL must be public. Warning: This will overwrite the Email Body of the template.)", Browser.Buttons.OK_CANCEL);
      HTMLBody = null;

  // Convert
}

function menuTestTemplate() {
  var aSpread = SpreadsheetApp.getActiveSpreadsheet(),
      contactSheet = aSpread.getSheetByName(CONTACTS_SHEET),
      template = getEmailData(aSpread),
      mergeData = contactSheet.getDataRange().getValues(),
      searchList = makeTagSearchList(template.body),
      colKeys = columnHeadersKeys(mergeData[0]),
      contactRow = mergeData[1],
      mergeTest = mergeTemplate(template.plainText, searchList, contactRow, colKeys),
      wrongTags = findWrongTags(searchList, colKeys);


  if (wrongTags.length > 0) {
    Browser.msgBox("These tags don't have matching data columns: " + wrongTags);
  } else {
    Browser.msgBox(mergeTest);
  }
}

function menuSendEmail() {
  var response = Browser.msgBox("Confirm Send", "Start the mail merge?", Browser.Buttons.YES_NO);
  if (response === 'yes') {
    sendEmailToList(CONTACTS_SHEET);
  }
}

function menuSendTestEmail() {
  sendEmailToList(TEST_CONTACTS_SHEET);
}

function menuShowRemainingEmails() {
  var aSpread = SpreadsheetApp.getActiveSpreadsheet(),
      numEmails = MailApp.getRemainingDailyQuota(),
      reserveAmount = aSpread.getRangeByName("EmailQuotaReserve").getValue();
  Browser.msgBox("You can send " + (numEmails - reserveAmount) + " (" + numEmails + ") more emails." );
}

function menuShowHelp() {
  Browser.msgBox("Originally based on Mail Merge by Labnol <http://labnol.org/?p=13289> @labnol on Twitter.", Browser.Buttons.OK_CANCEL);
}


/* ****************************************************************************
 * Tests
 *****************************************************************************/

function testMergeTags() {
  var aSpread = SpreadsheetApp.getActiveSpreadsheet(),
      newLine = "<br><br>",

      emailBody = getMailBody(aSpread, newLine),
      tags = makeTagSearchList(emailBody),
      display = '', name;

  for (name in tags) {
    display += name + ': ' + tags[name] + "\n";
  }

  Logger.log(display);
}

function testColumnHeaders() {
  var aSpread = SpreadsheetApp.getActiveSpreadsheet(),
      contactSheet = aSpread.getSheetByName(CONTACTS_SHEET),
      template = getEmailData(aSpread),
      mergeData = contactSheet.getDataRange().getValues(),
      tagCols = columnHeadersKeys(mergeData[0]),
      display = '', name;

  for (name in tagCols) {
    display += name + ': ' + tagCols[name] + "\n";
  }

  Logger.log(display);
}

function testPlaintextEmail() {
  var aSpread = SpreadsheetApp.getActiveSpreadsheet(),
      template = getEmailData(aSpread);

  Logger.log(template.body);
  Logger.log(template.plainText);
}

function testPrepareAttachments() {
  var aSpread = SpreadsheetApp.getActiveSpreadsheet(),
      advArgs = prepareAttachments(aSpread);;

  Logger.log(advArgs);
}
