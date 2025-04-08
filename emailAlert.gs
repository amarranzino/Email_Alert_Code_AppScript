//function to pull in data from S&T Metrics spreadsheet and calculate duedates then send email alerts when due dates approach
/// Next steps: write code for metrics additions
 
function emailAlert(){
  
  Logger.log("Function started");
  
  //Initialize blank array to store upcoming due dates
  var upcomingDueDates = [];
  //Get Today's Date
  var today =  new Date();
  
  Logger.log("Today's date: " + today);
  
  // getting data from spreadsheet S&T Metrics using the Google Sheet ID
  var spreadsheet = SpreadsheetApp.openById("17WNH_DGlP9pvsFj4v4EkQnrRv2ygJhR7kxUKliq0qYI");
  // set which of the tabs in the spreadsheet to pull data from
  var sheet = spreadsheet.getSheetByName("Grants");
  
  if (!sheet) {
    Logger.log("Grants not found within sheet!");
    SpreadsheetApp.getUi().alert('Error: Sheet Grants not found!');
    return;
  }

  // Lookup Table for Column Headers
  const PROJECT_NAME_HEADER = "LT Assigned Project Name (from LT - do NOT type here)"; 
  const POC_NAME_HEADER = "OER POC"; 
  const POC_EMAIL_HEADER = "OER POC Email"; 
  const PROJECT_FY_HEADER = "Project FY (project funded)"; 
  const GRANT_NUMBER_HEADER = "Grant Award Number"; 
  const START_DATE_HEADER = "Grant Award Date"; 
  const END_DATE_HEADER = "Grant End Date"; 
  const CRUISE_START_HEADER = "Start (UTC)"; 
  const CRUISE_END_HEADER = "End (UTC)"; 
  const PI_NAME_HEADER = "PI Last Name"; 
  const DATA_DUE_HEADER = "Data Due Date"; 
  const UNIQUE_ID_HEADER = "Unique ID";


  // Read Headers (Row 1) and produce error message if cannot read in the headers
  var headerRowValues;
    try {
      headerRowValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      Logger.log("Read in header rows.");
    } catch(e) {
      Logger.log("Error reading header row: " + e);
      return;
    }

    // Create Header Map to Store Column Header Names 
    const headerMap = {};
    headerRowValues.forEach((header, index) => {
      if (header && typeof header === 'string') { // only map non-empty string headers
        const normalizedHeader = header.replace(/\s+/g, ' ').trim(); // collapse all whitespace (incl. newlines) to a single space
        headerMap[normalizedHeader] = index;
      }
    })

    // Check to make sure we've read in all of our headers correctly
    const requiredHeaders = [
      PROJECT_NAME_HEADER, POC_NAME_HEADER, POC_EMAIL_HEADER, PROJECT_FY_HEADER,
      GRANT_NUMBER_HEADER, START_DATE_HEADER, END_DATE_HEADER, CRUISE_START_HEADER,
      CRUISE_END_HEADER, PI_NAME_HEADER, DATA_DUE_HEADER, UNIQUE_ID_HEADER
    ];
    let missingHeaders = [];
    requiredHeaders.forEach(header => {
      if (headerMap[header] === undefined) {
        missingHeaders.push(header);
      }
    });

    if (missingHeaders.length > 0) {
      const errorMessage = 'Error: The following required columns were not found in the "Grants" sheet header row: ' + missingHeaders.join(',');
      Logger.log(errorMessage);
      return;
    }

    // Turn on logger function to check data has pulled in correctly if needed
    //Logger.log(sheet)
    
    //Get the data for the specific range within the spreadsheet. 
    //Note: in script, first column (A) counts as [0] so one is subtracted from all subsequent columns to call appropriate column

    //Find first empty row from the bottom of the data based on values in column A
    var columnA = sheet.getRange("A:A").getValues();
    var firstEmptyRow = -1;
    for (var i = sheet.getMaxRows() -1; i >= 1; i--) {
      if (columnA[i] && columnA[i][0] !== "") {
        firstEmptyRow = i + 1;
        Logger.log("Last data row found in A is row: " + (i + 1) + ". Reading up until this row!");
        break;
      }
    }

    if (firstEmptyRow <= 1) {
      Logger.log("No data found below the header row in 'Grants' sheet.");
      return;
    }

    //Use the firstEmptyRow variable to determine the data range
    var dataRange = sheet.getRange (2,1, firstEmptyRow - 1, Object.keys(headerMap).length);
    var data = dataRange.getValues();
    Logger.log("Successfully read " + data.length + " rows of data.");
    //Get the values from the data range
    for(var row=0; row < data.length; row++){
      var currentRowInSheet = row + 2
        try{
          //~GENERAL PROJECT INFO~
          var pocName = data[row][headerMap[POC_NAME_HEADER]];
          var pocEmail = data [row][headerMap[POC_EMAIL_HEADER]];
          var projectFY = data [row][headerMap[PROJECT_FY_HEADER]];
          var project = data [row][headerMap[PROJECT_NAME_HEADER]];
          var grantNumber = data[row][headerMap[GRANT_NUMBER_HEADER]];
          var grantStart = data[row][headerMap[START_DATE_HEADER]]; 
          var cruiseStart = data[row][headerMap[CRUISE_START_HEADER]];
          var cruiseEnd = data[row][headerMap[CRUISE_END_HEADER]];
          var piName = data[row][headerMap[PI_NAME_HEADER]];
          var uniqueID = data[row][headerMap[UNIQUE_ID_HEADER]];
          
          // Define 'today' once for the row processing
          var today = new Date();

          //~DATA DUE DATE ~ Handling Start
          var rawDataDueDateValue = data[row][headerMap[DATA_DUE_HEADER]];
          Logger.log("Row " + currentRowInSheet + ": Raw value from '" + DATA_DUE_HEADER + "': '" + rawDataDueDateValue + "'");
          var dataDue = new Date(rawDataDueDateValue);

          if (!dataDue || isNaN(dataDue.getTime())) {
              Logger.log("Row " + currentRowInSheet + ": Invalid or non-date value in Data Due Date column. Skipping Data Due calculations for this row.");
          } else {
              Logger.log("Row " + currentRowInSheet + ": Valid Data Due Date found: " + dataDue);
              var daystodataDueDate = Math.floor((dataDue - today) / (1000 * 60 * 60 * 24));
              var formattedDataDueDate = Utilities.formatDate(dataDue, "GMT", "MM/dd/yyyy");
              //IF data is due in less than 90 days, add project to upcomingDueDates array
              if (daystodataDueDate >= 0 && daystodataDueDate <= 90) {
                  upcomingDueDates.push({
                    type:"Data Submission",
                    piName: piName,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystodataDueDate,
                    formattedDueDate: formattedDataDueDate,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project,
                    uniqueID: uniqueID
                  });
                  Logger.log("Row " + currentRowInSheet + ": Added 'Data Due' reminder.");
                  }
                }
          


          //~FINAL REPORT~ 
          var endDate = new Date(data[row][headerMap[END_DATE_HEADER]]); // Get endDate here
          if (endDate && !isNaN(endDate.getTime())) {
              var finalReportDue = new Date(endDate);
              finalReportDue.setDate(endDate.getDate()+120); //Adds 120 days to end of grant to calculate the final report due date
              var daystoFinalReportDue = Math.floor((finalReportDue - today) / (1000 * 60 * 60 * 24));
              var formattedFinalReport = Utilities.formatDate(finalReportDue, "GMT", "MM/dd/yyyy");
              //IF final report is due in less than 60 days, add project to upcomingDueDates array
              if (daystoFinalReportDue >= 0 && daystoFinalReportDue <= 60) {
                  upcomingDueDates.push({
                    type:"Final Report",
                    piName: piName,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystoFinalReportDue,
                    formattedDueDate: formattedFinalReport,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project,
                    uniqueID: uniqueID
                  });
              } else {
              Logger.log("Row " + currentRowInSheet + ": Invalid Grant End Date. Skipping Final Report & NCE calculations.");
              }
          }

          //~CRUISE REPORT~ 
          var cruiseEnd = new Date(data[row][headerMap[CRUISE_END_HEADER]]); // Get cruiseEnd here
          if (cruiseEnd && !isNaN(cruiseEnd.getTime())) {
            var cruiseReportDue = new Date(cruiseEnd);
            cruiseReportDue.setDate (cruiseEnd.getDate()+60); //Calculate the cruise report due date - 60 days after cruise ends
            var daystoCruiseReportDue = Math.floor((cruiseReportDue - today) / (1000*60*60*24));
            var formattedCruiseReportDue = Utilities.formatDate(cruiseReportDue, "GMT", "MM/dd/yyyy");
            // If Cruise report is due in less than 45 days, add project to upcomingDueDates array
            if (daystoCruiseReportDue >=0 && daystoCruiseReportDue <=45){
              upcomingDueDates.push({
                    type:"Cruise Report",
                    piName: piName,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystoCruiseReportDue,
                    formattedDueDate: formattedCruiseReportDue,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project,
                    uniqueID: uniqueID
                  });
              }
          } else {
              Logger.log("Row " + currentRowInSheet + ": Invalid Cruise End Date. Skipping Cruise Report calculations.");
          }

          //~CRUISE PLAN~ 
          var cruiseStart = new Date(data[row][headerMap[CRUISE_START_HEADER]]); // Get cruiseStart here
          if (cruiseStart && !isNaN(cruiseStart.getTime())) {
            var cruisePlanDue = new Date (cruiseStart); 
           //Set cruise plan due date to 30 days before cruise for any project funded in FY22 or earlier and to 60 days prior to cruise for any project funded in or after FY23
            if(projectFY >= 23){
              cruisePlanDue.setDate(cruiseStart.getDate() - 60); 
            }
            else{
              cruisePlanDue.setDate(cruiseStart.getDate() - 30);
            }

            var daystoCruisePlanDue = Math.floor((cruisePlanDue - today) / (1000*60*60*24));
            var formattedCruisePlanDue = Utilities.formatDate(cruisePlanDue, "GMT", "MM/dd/yyyy");
            // If Cruise plan is due in less than 30 days, add project to upcomingDueDates array
            if (daystoCruisePlanDue >=0 && daystoCruisePlanDue <=45){
              upcomingDueDates.push({
                    type:"Cruise Plan",
                    piName: piName,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystoCruisePlanDue,
                    formattedDueDate: formattedCruisePlanDue,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project,
                    uniqueID: uniqueID
                  });
              }
          } else {
              Logger.log("Row " + currentRowInSheet + ": Invalid Cruise Start Date. Skipping Cruise Plan calculations.");
          }

          //~NO COST EXTENSION~ (Depends on validated endDate)
          if (endDate && !isNaN(endDate.getTime())) {
              var noCostExtension = new Date(endDate);
              noCostExtension.setDate(endDate.getDate()-30); //No Cost Extension calculated as 30 days before grant ends (NOTE: PI's should submit NCEs 30 - 60 days before grant end)
              var daystoNoCostExtension = Math.floor((noCostExtension - today) / (1000 * 60 * 60 * 24));
              var formattedNoCostExtension = Utilities.formatDate(noCostExtension, "GMT", "MM/dd/yyyy");
            // If No Cost Extension is due in less than 30 days, add project to upcomingDueDates array
            if (daystoNoCostExtension >=0 && daystoNoCostExtension <=45){
              upcomingDueDates.push({
                    type:"No Cost Extension Submission Upcoming",
                    piName: piName,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystoNoCostExtension,
                    formattedDueDate: formattedNoCostExtension,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project,
                    uniqueID: uniqueID
                  });
              } 
          } // Skipped if endDate is invalid

         
          Logger.log ("Row"+ currentRowInSheet + " processed");

        } catch (error) // End of try block
        {
          Logger.log("Error processing row " + currentRowInSheet + ": " + error + (error.stack ? "\nStack: " + error.stack : ""));
        }
      } // End of for loop
      
      //Sort upcomingDueDates array so that the values are in an order you want
      Logger.log( "Number of upcoming due dates:  " + upcomingDueDates.length);

      
      upcomingDueDates.sort(function(a,b){
        Logger.log ("Sorting Due Dates");
        //First sort by report type
        if(a.type < b.type) return -1;
        if (a.type > b.type) return 1;
        //Then sort by POC name
        if(a.pocName < b.pocName) return -1;
        if (a.pocName > b.pocName) return 1;
        //Then sort by due date
        return a.daystoDue - b.daystoDue; 
        });
            
      //print out the array to check if it is in the appropriate order
        Logger.log ("Sorted due dates:  " + JSON.stringify(upcomingDueDates));
          
      // Send an email with the upcoming due dates 
      // call the sendUpcomingDueDates function defined below to send a single email with all upcoming due dates to individuals specified in the function
      // update the text of the messages or the recipients within the respective functions 
      sendUpcomingDueDatesEmail(upcomingDueDates);
      sendPOCemails(upcomingDueDates);
  }
