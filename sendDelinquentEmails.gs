//~~FUNCTION TO SEND EMAIL WITH ALL DELINQUENT REPORTS~~
// Send one email out with all upcoming due dates


function sendDelinquentEmails(delinquent){
  //Print that the function has been called and the number of due dates in the array
  Logger.log( "Function sendDelinquentEmails called. Number of due delinquent reports:  " + delinquent.length);
   
  //If there are any due dates listed in the array (delinquent), then create an email listing the information for each delinquent report
  //Use html format for bolding <b> (start of bolding) and </b> (end of bolding) and then use html for spaceing - <br> for hard return (\n in java)
   if (delinquent.length >0){
    var emailContent =  "Total number of delinquent reports:<b>" + delinquent.length + "<br><br>"
    delinquent.forEach(function(dueDate){ 
      emailContent += "The " + dueDate.type + " for " + "<b>" + dueDate.uniqueID + "</b>" + " is overdue by " + "<b>" + Math.abs(dueDate.daystoDue) + " days</b>.<br>"
      + "\n";  
    });
         
    //check if the email contains any due dates prior to sending
    var emailAddress = ["ashley.marranzino@noaa.gov", "christina.ortiz@noaa.gov", "adrienne.copeland@noaa.gov"].join(", "); //change this to anyone who should receive full email update
    //**FOR DEBUGGING** turn off var emailAddress above and turn on the one below. (remove .join(",") and send to a single recipient instead)
    //var emailAddress = "ashley.marranzino@noaa.gov";
    
    //send email
    MailApp.sendEmail({
        to: emailAddress,
        subject: "Delinquent Reports S&T Metrics",
        htmlBody: emailContent //htmlBody (instead of just "body") allows for html coding to be read (bold, italics, etc.) in the above email code
      });
      
    //Print the email address the email was sent to if it sends. Otherwise, print "No upcoming due dates to send"
    Logger.log ("Email sent to " + emailAddress);
  } else {
    Logger.log ("No delinquent reports to send.");
  };
} 
