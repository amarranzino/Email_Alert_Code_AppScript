//function saved from https://mccarthydanielle.medium.com/trigger-email-reminders-based-on-dates-in-google-sheets-9aa2060d7aa2
/// Next steps: UPDATE RPPR DUE DATES SO FEDS ARE JUST EVERY 6 MONTHS (pull in PI Affiliation col, then filter for All fields starting with "Federal" and if those, then reports are due 30 days after end of reporting window) ; write code for metrics additions
 
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
  const GRANT_STATUS_HEADER = "Grant Status";
  const PI_AFFILIATION_HEADER = "PI Affiliation"


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
      CRUISE_END_HEADER, PI_NAME_HEADER, DATA_DUE_HEADER, UNIQUE_ID_HEADER, GRANT_STATUS_HEADER,
      PI_AFFILIATION_HEADER
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

     // FOR DEBUGGING - turn on the following if loop and select the row you would like to test- this will skip any rows besides the one you input.
     /* if(currentRowInSheet !==614){ //skip over all rows besides one - turned on for debugging
        continue;
        }
     */
        try{
          //~GENERAL PROJECT INFO~
          var pocName = data[row][headerMap[POC_NAME_HEADER]];
          var pocEmail = data [row][headerMap[POC_EMAIL_HEADER]];
          var projectFY = data [row][headerMap[PROJECT_FY_HEADER]];
          var project = data [row][headerMap[PROJECT_NAME_HEADER]];
          var grantNumber = data[row][headerMap[GRANT_NUMBER_HEADER]];
          var cruiseStart = data[row][headerMap[CRUISE_START_HEADER]];
          var cruiseEnd = data[row][headerMap[CRUISE_END_HEADER]];
          var piName = data[row][headerMap[PI_NAME_HEADER]];
          var uniqueID = data[row][headerMap[UNIQUE_ID_HEADER]];
          var grantStatus = data[row][headerMap[GRANT_STATUS_HEADER]];
          var affiliation = data[row][headerMap[PI_AFFILIATION_HEADER]];
          
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
          /*var rawEndDateValue = data[row][headerMap[END_DATE_HEADER]];
          Logger.log("Row " + currentRowInSheet + ": Raw value from '" + END_DATE_HEADER + "': '" + rawEndDateValue + "'");
         */
          var endDate = new Date(data[row][headerMap[END_DATE_HEADER]]);
          
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
              Logger.log("Row " + currentRowInSheet + ": Final Report due date > 60 days.");
              }
          } else {
            Logger.log ("Row " + currentRowInSheet + ": Invalid Grant End Date. Skipping Final Report & NCE calculations.");
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
              }else {
              Logger.log("Row " + currentRowInSheet + ": Cruise Report not due within 45 days.");
              }
          } else {
              Logger.log("Row " + currentRowInSheet + ": Invalid Cruise End Date. Skipping Cruise Report calculations.");
          }

          //~CRUISE PLAN~ 
          var cruiseStart = new Date(data[row][headerMap[CRUISE_START_HEADER]]); // Get cruiseStart here
          if (cruiseStart && !isNaN(cruiseStart.getTime())) {
            var cruisePlanDue = new Date (cruiseStart); 
            cruisePlanDue.setDate(cruiseStart.getDate() - 60); // Modified to just set this for 60 days prior to start date so that 2nd cruises are captured (currently not captured because the FY roject Funded field is not filled in for 2nd fieldwork rows). Will note that POCs need to verify grant funding date to determine actual due date. 
           /*
           //Set cruise plan due date to 30 days before cruise for any project funded in FY22 or earlier and to 60 days prior to cruise for any project funded in or after FY23
            if(projectFY >= 23){
              cruisePlanDue.setDate(cruiseStart.getDate() - 60); 
            }
            else{
              cruisePlanDue.setDate(cruiseStart.getDate() - 30);
            }
            */
            var daystoCruisePlanDue = Math.floor((cruisePlanDue - today) / (1000*60*60*24));
            var formattedCruisePlanDue = Utilities.formatDate(cruisePlanDue, "GMT", "MM/dd/yyyy");
            // If Cruise plan is due in less than 45 days, add project to upcomingDueDates array
            if (daystoCruisePlanDue >=0 && daystoCruisePlanDue <=45){
              upcomingDueDates.push({
                    type:"Cruise Plan*",
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
              }else {
              Logger.log("Row " + currentRowInSheet + ": Cruise Plan not due within 45 days.");
              }
          } else {
              Logger.log("Row " + currentRowInSheet + ": Invalid Cruise Start Date. Skipping Cruise Plan calculations.");
          }

          //~NO COST EXTENSION~ (Depends on validated endDate)
          if (endDate && !isNaN(endDate.getTime())) {
              var noCostExtension = new Date(endDate);
              noCostExtension.setDate(endDate.getDate()-60); //No Cost Extension calculated as 30 days before grant ends (NOTE: PI's should submit NCEs 30 - 60 days before grant end)
              var daystoNoCostExtension = Math.floor((noCostExtension - today) / (1000 * 60 * 60 * 24));
              var formattedNoCostExtension = Utilities.formatDate(noCostExtension, "GMT", "MM/dd/yyyy");
            // If No Cost Extension is due in less than 30 days, add project to upcomingDueDates array
            if (daystoNoCostExtension >=0 && daystoNoCostExtension <=30){
              upcomingDueDates.push({
                    type:"No Cost Extension Submission**",
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
              } else {
                Logger.log ("Row " + currentRowInSheet + ": NCE not due within 45 days.");
              }
          } else {
            Logger.log ("Row " + currentRowInSheet + ": Invalid Grant End Date. Skipping Final Report & NCE calculations.");
          }
        

       //~FIRST semiannual report (6 month RPPR)~ 
          //Starting 2025 - the first semiannual RPPR covers the first 6 months of the grant and is due in eRA 1 month after the 6 month reporting period ends. 
          var startDate = new Date(data[row][headerMap[START_DATE_HEADER]]); //get grant start date
          Logger.log("StartDate: " + startDate);
          if (startDate && !isNaN(startDate.getTime())) {
            //Calculate the due date ofthe first semiannual RPPR to be 7 months after start date 
            var semiannual6moDue =new Date (startDate);
            semiannual6moDue.setMonth (semiannual6moDue.getMonth()+7); //use .sentMonth and .getMonth instead of the .setDate and .getDate commands used above
            Logger.log ("Start date: " + startDate);
            var daystosemiannual6moDue = Math.floor((semiannual6moDue - today) / (1000*60*60*24));
            var formattedsemiannual6moDue = Utilities.formatDate(semiannual6moDue, "GMT", "MM/dd/yyyy");
            // If  report is due in less than 45 days, add project to upcomingDueDates array
            if (daystosemiannual6moDue >=0 && daystosemiannual6moDue <=45){
              upcomingDueDates.push({
                    type:"6 month RPPR",
                    piName: piName,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystosemiannual6moDue,
                    formattedDueDate: formattedsemiannual6moDue,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project,
                    uniqueID: uniqueID
                  });
              } else {
              Logger.log("Row " + currentRowInSheet + ": First RPPR not due in 45 days.");
              }
          } else {
              Logger.log("Row " + currentRowInSheet + ": Invalid Grant Start Date. Skipping RPPR calculations.");
          }
          

      // ~RPPRs after the first 6 month RPPR submission !
        //~January and July Semiannual RPPRs~ 
        // Starting 2025, after the first 6 month report, the next 6 month RPPR will be due on either 30 Jan or 30 July following the close of the reporting period (see https://docs.google.com/spreadsheets/d/1Xds4itU7zr5cR9MYUwDiMEHtHGkqlFcG/edit?gid=911150554#gid=911150554 for examples)        

                
         
          if (grantStatus === "Open" && !affiliation.toString().toLowerCase().includes("federal")) {          
            let reportStart = new Date (startDate); // clones the startDate and cleans up values to avoid odd Java errors 
            Logger.log (Utilities.formatDate(reportStart, "GMT", "yyyy-MM-dd")); //ensure date is calculating correctly (otherwise dates may be off by a month)
            let currentYear = new Date().getFullYear();
            let semiannualJan = new Date (currentYear, 0, 30); // sets the due date for Juanuary semiannual reports for the 30th of January of the current year. Note format is (year, month, day) and Java calculates month as -1 from the calendar month so January = 0. 
            let semiannualJuly = new Date (currentYear, 6, 30); //sets the due date for July semiannual reports for the 30th of July of the current year 
            let daystosemiannualJan = Math.floor((semiannualJan - today) / (1000*60*60*24)); 
            let daystosemiannualJuly = Math.floor((semiannualJuly - today) / (1000*60*60*24));
            let formattedsemiannualJan = Utilities.formatDate(semiannualJan, "GMT", "MM/dd/yyyy"); 
            let formattedsemiannualJuly = Utilities.formatDate(semiannualJuly, "GMT", "MM/dd/yyyy"); 
            let july1 = new Date ((currentYear-1),6,1); //sets a date for 1 July of previous year
            let dec31 = new Date ((currentYear-1),11,31); //sets a date for 31 Dec of previous year
            let jan1 = new Date(currentYear, 0,1); //sets a date for 1 Jan of current year
            let june30 = new Date (currentYear, 5,30); //sets a date for 30 June of current year
            let sevenmonthsago = new Date (today);
            sevenmonthsago.setMonth(sevenmonthsago.getMonth()-7);
            Logger.log("7 months ago:" + sevenmonthsago);

            Logger.log("Starting RPPR loop"); 
          
            while (reportStart <= endDate) {
              let reportEnd =new Date (reportStart);
              reportEnd.setMonth(reportEnd.getMonth()+6); // calculate the end of the reporting period as every 6 months from startDate for the duration of the grant cycle
              
              Logger.log ("Report period: " + reportStart + " - " + reportEnd);

              if ((reportEnd >= july1 && reportEnd <= dec31) && (daystosemiannualJan >=0 && daystosemiannualJan <=45)){
                upcomingDueDates.push({
                  type: "Semi-annual RPPR",
                    piName: piName,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystosemiannualJan,
                    formattedDueDate: formattedsemiannualJan,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project,
                    uniqueID: uniqueID
                });
                Logger.log("Next RPPR due in " + daystosemiannualJan + " days on " + formattedsemiannualJan); 
              } else if((reportEnd >= jan1 && reportEnd <= june30) && 
               (startDate <= sevenmonthsago) && 
               (daystosemiannualJuly >=0 && daystosemiannualJuly <=45)){
                upcomingDueDates.push({
                  type: "Semi-annual RPPR",
                    piName: piName,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystosemiannualJuly,
                    formattedDueDate: formattedsemiannualJuly,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project,
                    uniqueID: uniqueID
                });
                Logger.log("Next Rppr due in " + daystosemiannualJuly + " days on " + formattedsemiannualJuly);
              } else{
                Logger.log("Row " + currentRowInSheet + " RPPR not due within 45 days.");
              }
              reportStart.setMonth(reportStart.getMonth()+6); //Adds 6 months to the end of the last start Month for the next loop
            }
           
          } else if (grantStatus === "Open" && affiliation.toString().toLowerCase().includes("federal")){
            let reportStart = new Date (startDate); // clones the startDate and cleans up values to avoid odd Java errors 
            Logger.log ("PI affiliation: " + affiliation); //ensure date is calculating correctly (otherwise dates may be off by a month)
            let currentYear = new Date().getFullYear();
            
            while (reportStart <= endDate) {
              let reportEnd =new Date (reportStart);
              reportEnd.setMonth(reportEnd.getMonth()+6); // calculate the end of the reporting period as every 6 months from startDate for the duration of the grant cycle
              Logger.log ("Report period: " + reportStart + " - " + reportEnd);
              let semiAnnualRPPRdue = new Date(reportEnd);
              semiAnnualRPPRdue.setMonth(semiAnnualRPPRdue.getMonth()+1);
              let daystosemiAnnualRPPRdue = Math.floor((semiAnnualRPPRdue - today) / (1000*60*60*24));
              let formattedsemiAnnualRPPRdue = Utilities.formatDate(semiAnnualRPPRdue, "GMT", "MM/dd/yyyy");

              Logger.log("Next RPPR due: " + formattedsemiAnnualRPPRdue);

              if (daystosemiAnnualRPPRdue >=0 && daystosemiAnnualRPPRdue <=45){
              upcomingDueDates.push({
                  type: "Semi-annual Report - Federal",
                    piName: piName,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystosemiAnnualRPPRdue,
                    formattedDueDate: formattedsemiAnnualRPPRdue,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project,
                    uniqueID: uniqueID
                });
                
              }
              reportStart.setMonth(reportStart.getMonth()+6); //Adds 6 months to theend of the last start month for the next Loop
            }
          } else {
            Logger.log ("Row " + currentRowInSheet + "skipped: " + grantStatus);
            continue; //move to next row in sheet if grant is closed
          }
                        
          Logger.log ("Row "+ currentRowInSheet + " processed");

        } catch (error) // End of try block
        {
          Logger.log("Error processing row " + currentRowInSheet + ": " + error + (error.stack ? "\nStack: " + error.stack : ""));
        }
      } // End of for loop
      
      //Sort upcomingDueDates array so that the values are in an order you want
      Logger.log( "Number of upcoming due dates:  " + upcomingDueDates.length);

      Logger.log ("Sorting Due Dates");
      upcomingDueDates.sort(function(a,b){
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
