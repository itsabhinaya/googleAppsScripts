//Sends mail merge from draft email

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createAddonMenu()
      .addItem('sendEmails', 'sendEmails')
      .addToUi();
}

function sendEmails(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("email");  //Change the sheetname to what ever want it to be
  
  var senderName = sheet.getRange(2, 4).getValue();
  var draftMessageSubject = sheet.getRange(1, 4).getValue();
    Logger.log(senderName);
        Logger.log(draftMessageSubject);

  var index = 4;
  
  var data2 = sheet.getDataRange().getValues();
  for (var i = 0, len = data2.length; i < len; i++) {
    if (data2[i][0] =='') break; 
  }
  var lastRow = i;
               Logger.log(i);

  for(;index <= lastRow; index++){
    var emailAddress = sheet.getRange(index, 1).getValue();
    var ccEmail = sheet.getRange(index, 2).getValue();
    var bccEmail = sheet.getRange(index, 3).getValue();
    var subject = sheet.getRange(index, 4).getValue();
    var body = sheet.getRange(index, 5).getValue();
    var attachment = sheet.getRange(index, 6).getValue();
               Logger.log(emailAddress);

    if(attachment != null && attachment != ""){
      var file = DriveApp.getFileById(attachment);
      var options = {
        attachments: file,
        name: senderName
      };
     }else{
       var options = {
          name: senderName,
          cc: ccEmail,
          bcc: bccEmail
        };
     }
           Logger.log("1");

    sendGmailTemplate(emailAddress, subject, body, options,draftMessageSubject)
    var currentDate = new Date();
    sheet.getRange(index, 7).setValue("Sent at:"+ currentDate); 
  }
}

/**
 * Signature code from here with some modifications: https://stackoverflow.com/questions/18493808/gmail-sending-emails-from-spreadsheet-how-to-add-signature-with-image
 * Insert the given email body text into an email template, and send
 * it to the indicated recipient. The template is a draft message with
 * the subject "TEMPLATE"; if the template message is not found, an
 * exception will be thrown. The template must contain text indicating
 * where email content should be placed: {BODY}.
 *
 * @param {String} recipient  Email address to send message to.
 * @param {String} subject    Subject line for email.
 * @param {String} body       Email content, may be plain text or HTML.
 * @param {Object} options    (optional) Options as supported by GmailApp.
 *
 * @returns        GmailApp   the Gmail service, useful for chaining
 */
function sendGmailTemplate(recipient, subject, body, options,draftMessageSubject) {
  options = options || {};  // default is no options
      Logger.log(options);

  var drafts = GmailApp.getDraftMessages();
  var found = false;
  for (var i=0; i<drafts.length && !found; i++) {
    if (drafts[i].getSubject() == draftMessageSubject) {
      found = true;
      var template = drafts[i];
    }
  }
  if (!found) throw new Error( "TEMPLATE not found in drafts folder" );

  // Generate htmlBody from template, with provided text body
  var imgUpdates = updateInlineImages(template);
  options.htmlBody = imgUpdates.templateBody.replace('{BODY}', body);
    if(imgUpdates.attachments.length > 0){
      options.attachments = imgUpdates.attachments;
    }
  options.inlineImages = imgUpdates.inlineImages;
  Logger.log(options);
  return GmailApp.sendEmail(recipient, subject, body, options);
}


/**
 * This function was adapted from YetAnotherMailMerge by Romain Vaillard.
 * Given a template email message, identify any attachments that are used
 * as inline images in the message, and move them from the attachments list
 * to the inlineImages list, updating the body of the message accordingly.
 *
 * @param   {GmailMessage} template  Message to use as template
 * @returns {Object}                 An object containing the updated 
 *                                   templateBody, attachments and inlineImages.
 */
function updateInlineImages(template) {
  //////////////////////////////////////////////////////////////////////////////
  // Get inline images and make sure they stay as inline images
  //////////////////////////////////////////////////////////////////////////////
  var templateBody = template.getBody();
  var rawContent = template.getRawContent();
  var attachments = template.getAttachments();

  var regMessageId = new RegExp(template.getId(), "g");
  if (templateBody.match(regMessageId) != null) {
    var inlineImages = {};
    var nbrOfImg = templateBody.match(regMessageId).length;
    var imgVars = templateBody.match(/<img[^>]+>/g);
    var imgToReplace = [];
    if(imgVars != null){
      for (var i = 0; i < imgVars.length; i++) {
        if (imgVars[i].search(regMessageId) != -1) {
          var id = imgVars[i].match(/realattid=([^&]+)&/);
          if (id != null) {
            var temp = rawContent.split(id[1])[1];
            temp = temp.substr(temp.lastIndexOf('Content-Type'));
            var imgTitle = temp.match(/name="([^"]+)"/);
            if (imgTitle != null) imgToReplace.push([imgTitle[1], imgVars[i], id[1]]);
          }
        }
      }
    }
    for (var i = 0; i < imgToReplace.length; i++) {
      for (var j = 0; j < attachments.length; j++) {
        if(attachments[j].getName() == imgToReplace[i][0]) {
          inlineImages[imgToReplace[i][2]] = attachments[j].copyBlob();
          attachments.splice(j, 1);
          var newImg = imgToReplace[i][1].replace(/src="[^\"]+\"/, "src=\"cid:" + imgToReplace[i][2] + "\"");
          templateBody = templateBody.replace(imgToReplace[i][1], newImg);
        }
      }
    }
  }
  var updatedTemplate = {
    templateBody: templateBody,
    attachments: attachments,
    inlineImages: inlineImages
  }
  return updatedTemplate;
}
