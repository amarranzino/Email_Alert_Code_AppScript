//~~FUNCTION TO SEND EMAIL TO POCS~~
function sendPOCemails (upcomingDueDates){
  //Print number of due dates in the array and note that the function has started
 Logger.log("Function sendPOCEmails called. Number of due dates:  " + upcomingDueDates.length);
  
  //Create an array of emails sent to each POC
  var pocEmails = {};
  
  if (upcomingDueDates.length >0){
    upcomingDueDates.forEach(function(dueDate){
      var pocName = dueDate.pocName;
      var pocEmail = dueDate.pocEmail; // Send email to each POC's email account
      
      // If the email for the POC is not already in the object, initialize it
      if (!pocEmails[pocEmail]){
        pocEmails[pocEmail]= {
          name: pocName,
          content: "Hello " + pocName + ", "+ "\n\n"+
          "This is a reminder of the following upcoming NOFO report due dates. Please email  the PI to remind them of the upcoming due date. Please send any relevant templates (i.e. cruise report template) and refer to the NOFO POC Manual (https://drive.google.com/file/d/1wGF36iINNm-TWx-Nv5Z3VXVdzbod7wc9/view?usp=sharing) for additional information on the upcoming deadline." + "\n\n"
        };
      }

      //customize content for each specific POC
      pocEmails[pocEmail].content += "A <b>" + dueDate.type + "</b> for " + dueDate.project + " is due in <b>" + dueDate.daystoDue + " days</b>.<br>"+
      "Project Details: <br>"+
      "PI Name: " + dueDate.piName +  "<br>"+
      "Unique ID: " + dueDate.uniqueID + "<br>"+
      "Project FY: " + dueDate.projectFY + "<br>"+
      "Grant Number: " + dueDate.grantNumber + "<br>"+
      "Due Date: " + dueDate.formattedDueDate + "<br><br>";
      });
      
    emailContent += "<br><i>*Projects funded prior to FY23 are only required to submit Cruise Plans is only due 30 days before the start of fieldwork. This due date is correct for projects funded in FY23 and beyond but is 30 days earlier than the due date for projects funded prior to FY23. Check project funding year to verify Cruise Plan due date. <br> No Cost Extension (NCE) deadline is calculated as 60 days before grant ends. PIs should submit their request for a NCE 60 days prior to the grant end and no later than 30 days before the grant ends. </i>";
      
      //Send email to each POC Name in the group
      for (var pocEmail in pocEmails){
        var emailContent = pocEmails[pocEmail].content;
        MailApp.sendEmail({
              to: pocEmail,
              subject: "POC Upcoming Due Date Notification",
              htmlBody: emailContent
            });
            
      Logger.log("Email sent to " + pocEmail);
      }
      } else {
    Logger.log("No upcoming due dates to send.");
  }
}

