function myFunction() {
  // Read data from the spreadsheet
  var data = SpreadsheetApp.getActiveSpreadsheet().getDataRange();
  var values = data.getValues();
  var row = data.getLastRow() - 1;
  var startTime = new Date(values[row][7]);
  var endTime = new Date(values[row][8]);

  var name = '🚴‍ ' + values[row][1] + ' (bike)';
  var distance = values[row][2];
  var duration = values[row][3];
  var linkToActivity = values[row][5];

  var goals = {
    2018: 1900,
    2019: 5400
  }

  // Calculate statistical info
  var workoutsTotal = 0.0;
  var durationTotal = 0.0;
  var distanceTotal = 0.0;
  var workoutsThisYear = 0.0;
  var durationThisYear = 0.0;
  var distanceThisYear = 0.0;
  var tripsAroundTheWorld;
  var tripsToTheMoon;
  for (var i = 1; i <= row; i++) {
    workoutsTotal++;
    durationTotal += parseInt(values[i][4]);
    distanceTotal += parseFloat(values[i][2]);
    if (new Date(values[i][7]).getFullYear() == new Date().getYear()) {
      workoutsThisYear++;
      durationThisYear += parseFloat(values[i][4]);
      distanceThisYear += parseFloat(values[i][2]);
    }
  }
  var distanceTotalInKilometers = distanceTotal / 1000;
  var distanceThisYearInKilometers = distanceThisYear / 1000;

  // Calculate fun facts
  var durationThisYearFormatted = getDuration(durationThisYear);
  var durationTotalFormatted = getDuration(durationTotal);
  tripsAroundTheWorld = distanceTotalInKilometers / 40075;
  tripsToTheMoon = distanceTotalInKilometers / 384400;

  // Create a description
  var currentYear = new Date().getFullYear();
  var currentDate = new Date().getTime() - new Date(currentYear, 0, 1, 0, 0, 0, 0).getTime();
  var endOfYear = new Date(currentYear+1, 0, 1, 0, 0, 0, 0).getTime() - new Date(currentYear, 0, 1, 0, 0, 0, 0).getTime();
  var goalProgressPercentage = (distanceThisYearInKilometers / goals[currentYear]) * 100;
  var yearProgressPercentage = ((endTime.getTime() - new Date(currentYear, 0, 1, 0, 0, 0, 0).getTime())/ endOfYear) * 100;
  var progressSymbol;
  if (goalProgressPercentage > yearProgressPercentage) progressSymbol = '✅'
  else progressSymbol = '❌';
  var eventDescription =
    '<b>Activity</b><br>' +
    'Distance: ' + distance + ' m <br>' +
    'Duration: ' + duration + '<br>' +
    'Link to activity: <a href=' + linkToActivity + '>Strava</a><br>' +
    '<br>' +
    '<b>This year</b><br>' +
    'Workouts: ' + workoutsThisYear + '<br>' +
    'Duration: ' + durationThisYearFormatted + '<br>' +
    'Distance: ' + distanceThisYearInKilometers.toFixed(1) + ' km<br>' +
    'Goal progress (' + goals[currentYear] + ' km) ' + progressSymbol + ':<br>' +
    getProgressBar(goalProgressPercentage) + '<br>' +
    'Year progress:<br>' +
    getProgressBar(yearProgressPercentage) + '<br>' +
    '<br>' +
    '<b>Total</b><br>' +
    'Workouts: ' + workoutsTotal + '<br>' +
    'Duration: ' + durationTotalFormatted + '<br>' +
    'Distance: ' + distanceTotalInKilometers.toFixed(1) + ' km<br>' +
    '🌍 Trips around the world: ' + tripsAroundTheWorld.toFixed(3) + '<br>' +
    '🚀 Trips to the Moon: ' + tripsToTheMoon.toFixed(3);

  // Create the event
  var event = CalendarApp.getCalendarById('3o84kndcvvt9f0qmik17j510fg@group.calendar.google.com').createEvent(
    name,
    new Date(startTime),
    new Date(endTime),
    { description: eventDescription });
  Logger.log('Event ID: ' + event.getId());
}

function getDuration(input) {
  var seconds = 0;
  var minutes = 0;
  var hours = 0;
  var days = 0;
  var years = 0;
  var finalString = '';
  seconds = input % 60;
  input = (input - seconds) / 60;
  minutes = input % 60;
  input = (input - minutes) / 60;
  hours = input % 24;
  input = (input - hours) / 24;
  days = input % 365;
  input = (input - days) / 365;
  years = input;
  if (years > 0) {
    finalString += years + 'y:';
  }
  if (days > 0) {
    finalString += days + 'd:';
  }
  if (hours > 0) {
    finalString += hours + 'h:';
  }
  if (minutes > 0) {
    finalString += minutes + 'm:';
  }
  if (seconds > 0) {
    finalString += seconds + 's';
  } else {
    finalString += '0s';
  }
  return finalString;
}

function getProgressBar(percentage) {
  if (percentage > 100) {
    percentage = 100;
  }
  const symbolArray = [
    '⣀',
    '⣄',
    '⣆',
    '⣇',
    '⣧',
    '⣷',
    '⣿'
  ]
  var progressBar = '';
  var progressBarLength = 20;
  var numberOfBarsDrawn = 0;
  var numberOfFullBars = Math.floor((percentage * progressBarLength) / 100);
  var remainder = ((percentage * progressBarLength) / 100) % 1;

  var whichSymbol = Math.floor(remainder * 7);

  while (numberOfFullBars > 0) {
    progressBar += symbolArray[6];
    numberOfFullBars--;
    numberOfBarsDrawn++;
  }
  if (numberOfBarsDrawn < progressBarLength) {
    progressBar += symbolArray[whichSymbol];
    numberOfBarsDrawn++;
  }

  while (numberOfBarsDrawn < progressBarLength) {
    progressBar += symbolArray[0];
    numberOfBarsDrawn++;

  }
  return progressBar + ' ' + percentage.toFixed(2) + '%';
}