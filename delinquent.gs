//Create a function that sends a notification about delinquent reports / data submissions based on due dates in upcomingDueDates array (see emailAlert.gs)

function delinquent(){
  
  Logger.log("Function started");
  
  //Initialize blank array to store upcoming due dates
  var delinquent = [];
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
  // Add in any additional column header names if needed for additional metrics reporting and add below to the section checking if headers have been read correctly
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
  const FINAL_REPORT_HEADER ="Final Grant Report";
  const CRUISE_PLAN_HEADER = "Cruise Plan";
  const CRUISE_REPORT_HEADER = "Cruise Report";
  const DATA_STATUS_HEADER = "Data Status";
  const FIELDWORK_COMPLETED_HEADER = "Fieldwork Completed";
  const UNIQUE_ID_HEADER = "Unique ID";
  const SEMI_ANNUAL_6M0_HEADER = "Semi Annual 1 6 months";
  const GRANT_STATUS_HEADER = "Grant Status"
  const FY_HEADER = "FY"



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
    // UPDATE with new headers if added in Lookup Column Headers section above
    const requiredHeaders = [
      PROJECT_NAME_HEADER, POC_NAME_HEADER, POC_EMAIL_HEADER, PROJECT_FY_HEADER,
      GRANT_NUMBER_HEADER, START_DATE_HEADER, END_DATE_HEADER, CRUISE_START_HEADER,
      CRUISE_END_HEADER, PI_NAME_HEADER, DATA_DUE_HEADER, FINAL_REPORT_HEADER, 
      CRUISE_PLAN_HEADER, CRUISE_REPORT_HEADER, DATA_STATUS_HEADER, FIELDWORK_COMPLETED_HEADER, 
      UNIQUE_ID_HEADER, SEMI_ANNUAL_6M0_HEADER, GRANT_STATUS_HEADER, FY_HEADER
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

    // Debugging: Turn on logger function to check data has pulled in correctly if needed
    //Logger.log(sheet)
    
    //Get the data for the specific range within the spreadsheet. 
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
          //~GENERAL PROJECT INFO~ Add more columns if needed (make sure they have already been pulled in above)
          var pocName = data[row][headerMap[POC_NAME_HEADER]];
          var pocEmail = data [row][headerMap[POC_EMAIL_HEADER]];
          var projectFY = data [row][headerMap[PROJECT_FY_HEADER]];
          var project = data [row][headerMap[PROJECT_NAME_HEADER]];
          var grantNumber = data[row][headerMap[GRANT_NUMBER_HEADER]];
          var grantStart = data[row][headerMap[START_DATE_HEADER]]; 
          var cruiseStart = data[row][headerMap[CRUISE_START_HEADER]];
          var cruiseEnd = data[row][headerMap[CRUISE_END_HEADER]];
          var piName = data[row][headerMap[PI_NAME_HEADER]];
          var fieldwork = data[row][headerMap[FIELDWORK_COMPLETED_HEADER]];
          var uniqueID = data[row][headerMap[UNIQUE_ID_HEADER]];
          var fy = data[row][headerMap[FY_HEADER]];
          var grantStatus = data[row][headerMap[GRANT_STATUS_HEADER]];
          
          // Define 'today' once for the row processing
          var today = new Date();

          if(((projectFY ==="" && fy >=18)|| (projectFY !=="" && projectFY>=18) && grantStatus == "Open")){ //checks for projects funded after FY18 since those projects fall under PAAR requirments. 
          //projectFY var will only be filled in for project's first row, so code checks for fy (fieldwork year) if projectFY is empty to serve as a proxy for funding year and catch reports for cruises subsequent to first round of fieldwork. 
          
          //~DATA DELINQUENT (wait until talking to Anna/ Adrienne to determine when data is delinquent)
         /*
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
              if (daystodataDueDate < 0) {
                  upcomingDueDates.push({
                    type:"Data Delinquent",
                    piName: piName,
                    uniqueID: uniqueID,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystodataDueDate,
                    formattedDueDate: formattedDataDueDate,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project
                    
                  });
                  Logger.log("Row " + currentRowInSheet + ": Added 'Data Due' reminder.");
                  }
                }
                */ 


          //~FINAL REPORT~ 
          var endDate = new Date(data[row][headerMap[END_DATE_HEADER]]); // Get endDate here
          var finalReportSubmitted = data[row][headerMap[FINAL_REPORT_HEADER]]; 
          if (endDate && !isNaN(endDate.getTime())) {
              var finalReportDue = new Date(endDate);
              finalReportDue.setDate(endDate.getDate()+120); //Adds 120 days to end of grant to calculate the final report due date
              var daystoFinalReportDue = Math.floor((finalReportDue - today) / (1000 * 60 * 60 * 24));
              var formattedFinalReport = Utilities.formatDate(finalReportDue, "GMT", "MM/dd/yyyy");
              //IF final report is overdue and has not been added to S&T Metrics, add to the delinquent array
              if (daystoFinalReportDue < 0 && (!finalReportSubmitted || finalReportSubmitted === "")) {
                  delinquent.push({
                    type:"Final Report",
                    piName: piName,
                    uniqueID: uniqueID,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystoFinalReportDue,
                    formattedDueDate: formattedFinalReport,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project
                  });
              } else {
              Logger.log("Row " + currentRowInSheet + ": Invalid Grant End Date. Skipping Final Report & NCE calculations.");
              }
          }

          //~CRUISE REPORT~  
          var cruiseEnd = new Date(data[row][headerMap[CRUISE_END_HEADER]]); // Get cruiseEnd here
          var cruiseReportSubmitted = (data[row][headerMap[CRUISE_REPORT_HEADER]]);
          if (cruiseEnd && !isNaN(cruiseEnd.getTime())) {
            var cruiseReportDue = new Date(cruiseEnd);
            cruiseReportDue.setDate (cruiseEnd.getDate()+60); //Calculate the cruise report due date - 60 days after cruise ends
            var daystoCruiseReportDue = Math.floor((cruiseReportDue - today) / (1000*60*60*24));
            var formattedCruiseReportDue = Utilities.formatDate(cruiseReportDue, "GMT", "MM/dd/yyyy");
            // If Cruise report is overdue, has not been submitted, and fieldwork was completed for the project, add the project to the delinquent array
            if (daystoCruiseReportDue <0 && (!cruiseReportSubmitted || cruiseReportSubmitted === "") && !fieldwork === true){
              delinquent.push({
                    type:"Cruise Report",
                    piName: piName,
                    uniqueID: uniqueID,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystoCruiseReportDue,
                    formattedDueDate: formattedCruiseReportDue,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project
                  });
              }
          } else {
              Logger.log("Row " + currentRowInSheet + ": Invalid Cruise End Date. Skipping Cruise Report calculations.");
          }

          //~CRUISE PLAN~ UPDATE HERE
          var cruiseStart = new Date(data[row][headerMap[CRUISE_START_HEADER]]); // Get cruiseStart here
          var cruisePlanSubmitted = data[row][headerMap[CRUISE_PLAN_HEADER]];
          //original curise plan due date calculation removed because projectFY is left blank after the first fieldwork instead, the 60 day report period imposed for all projects beginning in FY23 is calculated and viewers must check if that report is delinquent or not based on the actual project funding year. 
          if (cruiseStart && !isNaN(cruiseStart.getTime())) {
          var cruisePlanDue = new Date (cruiseStart);
          cruisePlanDue.setDate(cruiseStart.getDate() - 60);
          var daystoCruisePlanDue = Math.floor((cruisePlanDue - today) / (1000*60*60*24));
          var formattedCruisePlanDue = Utilities.formatDate(cruisePlanDue, "GMT", "MM/dd/yyyy");
            // If Cruise plan is is overdue and has not been added to S&T Metrics, add to the delinquent array
            if (daystoCruisePlanDue <0 && (!cruisePlanSubmitted || cruisePlanSubmitted === "")){
              delinquent.push({
                    type:"Cruise Plan",
                    piName: piName,
                    uniqueID: uniqueID,
                    projectFY: projectFY,
                    grantNumber: grantNumber,
                    daystoDue: daystoCruisePlanDue,
                    formattedDueDate: formattedCruisePlanDue,
                    pocName: pocName,
                    pocEmail: pocEmail,
                    project: project
                  });
              }
          } else {
              Logger.log("Row " + currentRowInSheet + ": Invalid Cruise Start Date. Skipping Cruise Plan calculations.");
          }
        
          } else{
            Logger.log ("No delinquent reports for projects funded after FY18");
          } 

          Logger.log ("Row"+ currentRowInSheet + " processed");

        } catch (error) // End of try block
        {
          Logger.log("Error processing row " + currentRowInSheet + ": " + error + (error.stack ? "\nStack: " + error.stack : ""));
        }
      } // End of for loop
      
      //Sort delinquent array so that the values are in an order you want
      Logger.log( "Number of delinquent reports:  " + delinquent.length);

      Logger.log ("Sorting Due Dates");
      delinquent.sort(function(a,b){
        /*//First sort by due date
        if(a.daystoDue - b.daystoDue)
        if(a.type < b.type) return -1;
        if (a.type > b.type) return 1;
        //Then sort by POC name
        if(a.pocName < b.pocName) return -1;
        if (a.pocName > b.pocName) return 1;
        */
        //Sort by due date so most recently delinquent are first
        return b.daystoDue - a.daystoDue; 
        });
            
      //print out the array to check if it is in the appropriate order
        Logger.log ("Sorted due dates:  " + JSON.stringify(delinquent));
          
      // Send an email with the upcoming due dates 
      // call the sendUpcomingDueDates function defined below to send a single email with all upcoming due dates to individuals specified in the function
      sendDelinquentEmails(delinquent); //update email addresses or text to email within the sendDelinquentEmails.gs file
        }
