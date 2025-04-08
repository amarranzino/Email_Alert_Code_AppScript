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
          "This is a reminder of the following upcoming NOFO report due dates. Please email  the PI to remind them of the upocoming due date. Please send any relevant templates (i.e. cruise report template) and refer to the NOFO POC Manual (https://drive.google.com/file/d/1_FfysKXe3A7hz6m_h4ALJP_9T9wjqT2t/view?usp=drive_link) for additional information on the upcoming deadline." + "\n\n"
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

