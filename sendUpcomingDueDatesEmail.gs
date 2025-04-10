//~~FUNCTION TO SEND EMAIL WITH ALL DUE DATES~~
// Send one email out with all upcoming due dates
function sendUpcomingDueDatesEmail(upcomingDueDates){
  //Print that the function has been called and the number of due dates in the array
  Logger.log( "Function sendUpcomingDueDatesEmail called. Number of due dates:  " + upcomingDueDates.length);
  Logger.log (upcomingDueDates);
  
  //If there are any due dates listed in the array (upcomingDueDates), then create an email listing the information for each upcoming due date
  if (upcomingDueDates.length >0){
    var emailContent =  "Due dates are approaching for <b>" + upcomingDueDates.length + " projects </b>in S&T Metrics.</b><br>"+
    "The following projects have upcoming due dates. POCs for each project will be notified as well. <br><br>";
    upcomingDueDates.forEach(function(dueDate){ 
      emailContent += "A <b>" + dueDate.type + "</b> for " + dueDate.project + " is due in <b>" + dueDate.daystoDue + " days</b>.<br>"+
      "Project Details: <br>"+
      "PI Name: " + dueDate.piName +  "<br>"+
      "Unique ID: " + dueDate.uniqueID + "<br>"+
      "Project FY: " + dueDate.projectFY + "<br>"+
      "Grant Number: " + dueDate.grantNumber + "<br>"+
      "Due Date: " + dueDate.formattedDueDate + "<br>"+
      "POC Name: " + dueDate.pocName + "<br>"+
      "POC Email: " + dueDate.pocEmail+ "<br><br>";  
    });

    emailContent += "<br><i>*Projects funded prior to FY23 are only required to submit Cruise Plans is only due 30 days before the start of fieldwork. This due date is correct for projects funded in FY23 and beyond but is 30 days earlier than the due date for projects funded prior to FY23. Check project funding year to verify Cruise Plan due date.</i>";
         
    //check if the email contains any due dates prior to sending
    //var emailAddress = ["ashley.marranzino@noaa.gov", "christina.ortiz@noaa.gov", "adrienne.copeland@noaa.gov", "anna.s.lienesch@noaa.gov"].join(", "); //change this to anyone who should receive full email update
    //**FOR DEBUGGING** turn off var emailAddress above and turn on the one below. (remove .join(",") and send to a single recipient instead)
    var emailAddress = "ashley.marranzino@noaa.gov";
    
    //send email
    MailApp.sendEmail({
        to: emailAddress,
        subject: "Upcoming NOFO Due Date Notification",
        htmlBody: emailContent
      });
      
    //Print the email address the email was sent to if it sends. Otherwise, print "No upcoming due dates to send"
    Logger.log ("Email sent to " + emailAddress);
  } else {
    Logger.log ("No upcoming due dates to send.");
  };
} 
