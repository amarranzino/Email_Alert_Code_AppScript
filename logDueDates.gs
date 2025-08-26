
/**
 * Logs the upcoming due dates array to the Apps Script console in a readable format.
 * This is useful for debugging and checking the array contents without a popup.
 *
 * @param {Array<Object>} upcomingDueDates The array of due date objects.
 */

function logDueDates(upcomingDueDates) {
  // Check if the array is empty before logging.
  if (upcomingDueDates.length === 0) {
    Logger.log('No upcoming due dates found.');
    return;
  }

  // Use JSON.stringify for a clean, full-object view, or map to a custom string for readability.
  // The map() approach is often more helpful for quick debugging.
  const loggableOutput = upcomingDueDates.map(item => {
   return `Project: ${item.project}\n PI: ${item.piName}\n Type: ${item.type}\n Due in: ${item.daystoDue} days on ${item.formattedDueDate}\n POC: ${item.pocName}`;
  }).join('\n\n');


  // Log the formatted message to the console.
  Logger.log('--- Upcoming Due Dates Report ---');
  Logger.log(loggableOutput);
  Logger.log('-------------------------------');

}
