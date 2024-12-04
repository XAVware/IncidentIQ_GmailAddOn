
const baseURL = 'https://<YOUR_DOMAIN>.incidentiq.com'
const configSheetURL = 'https://docs.google.com/spreadsheets/d/<CONFIG_SHEET_ID>';
const configSheetName = 'CONFIG'; // Or whatever you named it

// Only users that have an Incident IQ API key tied to their email in the configSheet will be able to click this button.
//    - See getApiKey() and buildAddOn()
async function submitClicked(e) {
  const data = e.formInput;
  const requesterEmail = data.requester_email;
  const subject = data.ticket_subject;
  const body = data.ticket_body;

  // Should add additional error handling
  if (requesterEmail.length == 0 || subject.length == 0 || body.length == 0) { 
    return 
  }

  try {
    const requester = getUser(requesterEmail);
    const ticketData = await createTicket(requester, subject, body);
    const newTicketId = String(ticketData["Uid"]);

    // Fetch the currently logged-in user
    const currentUserEmail = Session.getEffectiveUser().getEmail();
    const currentUser = getUser(currentUserEmail);
    const currentUserId = currentUser["UserId"];

    const response = assignTicket(newTicketId, currentUserId)
  } catch (error) {
    console.log(error)
  }  
}

// Lookup the user's Incident IQ API Key
function getApiKey(email) {
  const configSheet = SpreadsheetApp.openByUrl(configSheetURL);
  const config = configSheet.getSheetByName(configSheetName);
  const emailIndex = 0;
  const apiKeyIndex = 2;
  var data = config.getDataRange().getValues();

  for (row in data) {
    const rowData = data[row];
    const rowEmail = rowData[emailIndex];

    if (rowEmail == email) {
      return rowData[apiKeyIndex];
    }
  }
}

// Create ticket in IIQ via POST method with ticket details. Return JSON representation of response.
async function createTicket(requester, subject, details) {
  const url = `${baseURL}/services/tickets/new`;
  const requesterId = requester["UserId"];

  let config = {
    "method": 'POST',
    "headers": getHeaders(),
    "payload" : JSON.stringify({
      "Assets": [
        {
          "AssetId": null
        }
      ],
      "IsUrgent": false,
      "TicketFollowers": null,
      "Users": null,
      "HasSensitiveInformation": true,
      "IssueDescription": details,
      "SourceId": 1,
      "Subject": subject,
      // Owner Id and ForId should be set to the same ID, otherwise it will look like "{OwnerId} created the ticket on behalf of {ForId}."
      "OwnerId": requesterId,
      "ForId": requesterId
    })
  };
  
  var ticketResponse = UrlFetchApp.fetch(url, config);
  return JSON.parse(ticketResponse);
}

function assignTicket(ticketId, assignToUserId) {
  const url = `${baseURL}/api/v1.0/tickets/${ticketId}/assign`;

  let config = {
    "method": 'POST',
    "headers": getHeaders(),
    "payload": JSON.stringify({
      "TicketId": ticketId,
      "AssignToUserId": assignToUserId
      })
  };
  
  return UrlFetchApp.fetch(url, config);
}

// Generate network call headers
function getHeaders() {
  const currentUserEmail = Session.getEffectiveUser().getEmail();
  const API_KEY = getApiKey(currentUserEmail);
  return {
      "Client": 'ApiClient', 
      "Accept": 'application/json, text/plain, */*', 
      "Authorization": `Bearer ${API_KEY}`,  
      "Pragma": 'no-cache', 
      "Accept-Encoding": 'gzip, deflate', 
      "Content-Type": 'application/json'
    };
}

// Fetch a user object from IncidentIQ with their email
function getUser(email) {
  var url = `${baseURL}/api/v1.0/search`;
  var options = {
    "method": "POST",
    "headers": getHeaders(),
    "payload": JSON.stringify({
      "Query": email,
      "Facets": 4,
      "Limit": 1
    })
  };

  var response = UrlFetchApp.fetch(url, options);
  const responseString = JSON.parse(response.getContentText());
  const firstUser = responseString["Item"]["Users"][0];
  return firstUser;
}

// Extract the email address from the data returned from GmailApp.getMessageById(id).getFrom()
function extractEmail(str) {
  var matches = str.match(/<([^>]+)>/);
  return (matches && matches.length > 1) ? matches[1] : null;
}

// ~~~~~~~~~~ UI Related Functions ~~~~~~~~~~
function buildAddOn(e) {
  const currentUserEmail = Session.getEffectiveUser().getEmail();
  const userHasAccess = getApiKey(currentUserEmail) != '';
  const header = CardService.newCardHeader().setTitle('Quick Ticket').setImageUrl('https://www.gstatic.com/images/icons/material/system/1x/label_googblue_24dp.png');

  if (userHasAccess) {
    // Access Granted
    var accessToken = e.messageMetadata.accessToken;
    GmailApp.setCurrentMessageAccessToken(accessToken);
    var messageId = e.messageMetadata.messageId;

    var message = GmailApp.getMessageById(messageId);

    var msgFrom = extractEmail(message.getFrom());
    var msgSubject = message.getSubject();
    var msgBody = message.getBody();

    var requesterSection = CardService.newCardSection()
      .setHeader(getFormattedHeader("Requester Email"));

    var requesterField = CardService.newTextInput()
      .setFieldName("requester_email")
      .setValue(msgFrom);

    requesterSection.addWidget(requesterField);

    var ticketSubjectSection = CardService.newCardSection()
      .setHeader(getFormattedHeader("Ticket Subject"));

    var subjectField = CardService.newTextInput()
      .setFieldName("ticket_subject")
      .setValue(msgSubject);

    ticketSubjectSection.addWidget(subjectField);

    var ticketBodySection = CardService.newCardSection()
      .setHeader(getFormattedHeader("Ticket Details"));

    var bodyField = CardService.newTextInput()
      .setFieldName("ticket_body")
      .setMultiline(true)
      .setValue(msgBody);

    ticketBodySection.addWidget(bodyField);

    var action = CardService.newAction()
      .setFunctionName('submitClicked');

    var submitButton = CardService.newTextButton()
      .setText("Create Ticket")
      .setOnClickAction(action);
    
    ticketBodySection.addWidget(submitButton);

    var card = CardService.newCardBuilder()
      .setHeader(header)
      .addSection(requesterSection)
      .addSection(ticketSubjectSection)
      .addSection(ticketBodySection)
      .build();

    return [card];

  } else {

    var noAccessSection = CardService.newCardSection()
      .setHeader(getFormattedHeader("No Access"));

    var textParagraph = CardService.newTextParagraph()
      .setText("You do not have access to this add on. Please contact your system administrator.");

    noAccessSection.addWidget(textParagraph);

    var card = CardService.newCardBuilder()
      .setHeader(header)
      .addSection(noAccessSection)
      .build();

    return [card];

  }
}

function getFormattedHeader(text) {
  return `<font color=\"#000000\"><b>${text}</b></font>`;
}