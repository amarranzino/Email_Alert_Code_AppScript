/**
 * Logs the delinquent reports from array created in delinquent.gs file
 */
function logDelinquentReports(delinquent) {
if (delinquent.length === 0) {
    Logger.log('No delinquent reports found.');
    return;
  }

// Create a temporary object to group the reports by POC name.
  const groupedReports = delinquent.reduce((acc, item) => {
    // If the POC name doesn't exist as a key, create an empty array for it.
    if (!acc[item.pocName]) {
      acc[item.pocName] = [];
    }
    // Add the current report to the array for that POC.
    acc[item.pocName].push(item);
    return acc;
  }, {});

Logger.log('--- Delinquent Reports ---');
  
  // Iterate through each POC group and log a formatted section for them.
  for (const pocName in groupedReports) {
    if (Object.prototype.hasOwnProperty.call(groupedReports, pocName)) {
      Logger.log(`POC: ${pocName}`);
      
      // Sort the reports for the current POC by due date.
      groupedReports[pocName].sort((a, b) => b.daystoDue - a.daystoDue);
      
      // Map the reports to a readable string and join them with new lines.
      const reports = groupedReports[pocName].map(item => {
   return `The ${item.type} for ${item.piName} (${item.uniqueID}) is overdue by ${Math.abs(item.daystoDue)} days.`;
  }).join('\n');

   Logger.log(reports);
    }
  }

  Logger.log('-------------------------------');

}
