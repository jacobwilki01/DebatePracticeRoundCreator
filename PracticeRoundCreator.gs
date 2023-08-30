// Script created for the purpose of creating a Google Form and Sheet for the KU Debate Team to facilitate practice debates.
// Probably not the cleanest code on the planet but it does what it needs to.

// Function called when the button is pressed.
function buttonPressed()
{
  // Grab the creation spreadsheet.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create Form & Sheet");

  // Grab the values in the user-filled boxes.
  var dateRange = sheet.getRange("B5").getValue();
  var allowPartners = sheet.getRange("B6").getValue();
  var allowStrikes = sheet.getRange("B7").getValue();
  var timeslots = sheet.getRange("B8").getValue();

  // Create the name string and timeslots array.
  var name = "KU Practice Rounds " + dateRange;
  var timeArray = timeslots.split(", ");

  // Create the form and spreadsheet.
  var form = FormApp.create(name);
  var formSheet = SpreadsheetApp.create(name);

  // Set the information for the form.
  form.setTitle(name)
    .setAllowResponseEdits(true)
    .setAcceptingResponses(true)
    .setDestination(FormApp.DestinationType.SPREADSHEET, formSheet.getId());

  // Set the question to obtain the debater's name.
  var getName = form.addTextItem();
  getName.setTitle("Enter your name. (First & Last)");
  getName.isRequired = true;

  // Set the question to obtain the debater's availability.
  var dateChoice = form.addCheckboxItem();
  dateChoice.setTitle("Select which times you are available: ");
  dateChoice.isRequired = true;

  // Set the choice values based on the timeslots defined by the user.
  var choiceArray = [];
  for (var i = 0; i < timeArray.length; i += 1)
  {
    choiceArray.push(dateChoice.createChoice(timeArray[i]));
  }

  dateChoice.setChoices(choiceArray);

  // Set the question to obtain debater requests.
  var getRequests = form.addParagraphTextItem();
  getRequests.setTitle("Do you have any specific requests?")
    .setHelpText("(e.g., NEG vs. Policy, AFF with Bricker, etc.)");
  
  // If requested, create the question to allow for partner requests.
  if (allowPartners == "Yes")
  {
    var getPartners = form.addParagraphTextItem();
    getPartners.setTitle("Anyone you would like to debate with?")
      .setHelpText("No guarantees");
  }

  // If requested, create the question to allow for partner strikes.
  if (allowStrikes == "Yes")
  {
    var getStrikes = form.addParagraphTextItem();
    getStrikes.setTitle("Anyone you would would like to NOT debate with?")
      .setHelpText("Only coaches will see this information.");
  }

  // Format the spreadsheet and create the Availability Matrix page.
  formSheet.deleteSheet(formSheet.getSheetByName("Sheet1"));
  var availabilityMatrix = formSheet.insertSheet();
  availabilityMatrix.setName("Availability Matrix");

  // Set the top row of values for the page.
  availabilityMatrix.getRange("A1").setValue("Name");

  var columns = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"];
  for (var i = 0; i < timeArray.length; i += 1)
  {
    var range = availabilityMatrix.getRange(columns[i] + "1");
    range.setValue(timeArray[i]);
    availabilityMatrix.setColumnWidth(i + 2, 130);
  }

  availabilityMatrix.setColumnWidth(1, 130);

  // Create the debater requests page.
  var debaterRequests = formSheet.insertSheet();
  debaterRequests.setName("Debater Requests");
  debaterRequests.getRange("A1").setValue("Name");
  debaterRequests.getRange("B1").setValue("Requests");
  if (allowPartners == "Yes")
  {
    debaterRequests.getRange("C1").setValue("Partners");
  }
  if (allowStrikes == "Yes")
  {
    debaterRequests.getRange("D1").setValue("Strikes");
  }
  debaterRequests.getRange("A:A").setWrap(true);
  debaterRequests.getRange("B:B").setWrap(true);
  debaterRequests.getRange("C:C").setWrap(true);
  debaterRequests.getRange("D:D").setWrap(true);

  debaterRequests.setColumnWidth(1, 130);

  for (var i = 2; i < 5; i += 1)
  {
    debaterRequests.setColumnWidth(i, 150);
  }

  //Create the schedule page.
  var schedule = formSheet.insertSheet();
  schedule.setName("Schedule");
  schedule.setColumnWidth(1, 130);

  var availColumns = ["F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"];

  for (var i = 0; i < timeArray.length; i += 1)
  {
    var row = (i * 5) + 1;
    schedule.getRange("A" + row.toString()).setValue(timeArray[i]);
    schedule.getRange("B" + row.toString()).setValue("AFF").setFontWeight("bold");
    schedule.getRange("C" + row.toString()).setValue("NEG").setFontWeight("bold");
    schedule.getRange("D" + row.toString()).setValue("Coach").setFontWeight("bold");

    schedule.getRange(availColumns[i] + "1").setValue(timeArray[i]);
    schedule.setColumnWidth(i + 6, 130);
  }

  schedule.setColumnWidth(5, 20);

  var availColumns = ["F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"];

  // Tie the form response trigger to the below function.
  ScriptApp.newTrigger("handleFormResponse")
    .forSpreadsheet(formSheet)
    .onFormSubmit()
    .create();

  // Fetch the URLs of the sheet and form, and email them to the user.
  var sheetUrl = formSheet.getUrl();
  var formUrl = form.getEditUrl();

  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: name,
    htmlBody: "Sheet URL: " + sheetUrl + "    Form URL: " + formUrl
  });
}

// Function that handles all form responses.
function handleFormResponse(e)
{
  // Fetch the spreadhseet pages.
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var availabilityMatrix = sheet.getSheetByName("Availability Matrix");
  var debaterRequests = sheet.getSheetByName("Debater Requests");
  var schedule = sheet.getSheetByName("Schedule");

  // Fetch the current row and number of time slots.
  var curRow = getCurRow(availabilityMatrix.getRange("A:A").getValues()) + 1;
  var numCols = getNumCols(availabilityMatrix.getRange("1:1").getValues()) - 1;

  // Fill out the availability matrix page.
  availabilityMatrix.getRange("A" + curRow.toString()).setValue(e.values[1]);

  var columns = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"];
  var avails = [];
  for (var i = 0; i < numCols; i += 1)
  {
    var avail = e.values[2].includes(availabilityMatrix.getRange(columns[i] + "1").getValue());
    var cell = availabilityMatrix.getRange(columns[i] + curRow.toString());

    if (avail)
    {
      cell.setValue("Yes");
      cell.setBackgroundRGB(255, 50, 36)
    }
    else
    {
      cell.setValue("No");
      cell.setBackgroundRGB(45, 186, 52)
    }

    cell.setHorizontalAlignment("center");

    avails.push(avail);
  }

  // Set the values on the debater requests page.
  debaterRequests.getRange("A" + curRow.toString()).setValue(e.values[1]);
  debaterRequests.getRange("B" + curRow.toString()).setValue(e.values[3]);
  if (debaterRequests.getRange("C1").getValue() == "Partners")
  {
    debaterRequests.getRange("C" + curRow.toString()).setValue(e.values[4]);
  }
  if (debaterRequests.getRange("D1").getValue() == "Strikes")
  {
    debaterRequests.getRange("D" + curRow.toString()).setValue(e.values[5]);
  }

  // Set the values in the schedule tab.
  var availColumns = ["F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"];
  for (var i = 0; i < avails.length; i += 1)
  {
    var curCol = schedule.getRange(availColumns[i] + ":" + availColumns[i]).getValues();
    var curRow = getCurRow(curCol) + 1;

    if (avails[i])
    {
      schedule.getRange(availColumns[i] + curRow.toString()).setValue(e.values[1]);
    }
  }
}

// Function to obtain the next empty row in a column.
function getCurRow(column)
{
  for (var i = 0; i <= column.length; i += 1)
  {
    if (column[i][0] == "")
    {
      return i;
    }
  }
}

// Function to grab the number of columns with values.
function getNumCols(row)
{
  var i = 0;
  while (row[0][i] != "")
  {
    i += 1;
  }

  return i;
}