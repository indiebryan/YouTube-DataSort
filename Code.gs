/**

  Created by Bryan Curran - indieprogrammer.com
  OAuth 2.0 assistance from Martin Hawksey - mashe.hawksey.info
  
  || This Apps Script project is used to import, interpret, sort, and organize YouTube analytics data from a channel you own or manage


 */

/**
 * @OnlyCurrentDoc
 */
 
var service = getYouTubeService();

YouTube.setTokenService(function(){ return service.getAccessToken(); } );
YouTubeAnalytics.setTokenService(function(){ return service.getAccessToken(); });

var ss = SpreadsheetApp.getActiveSpreadsheet();

//Save our current sheet to a var
var sheet = SpreadsheetApp.getActiveSheet();

//Save our UI to a var, initialized in onOpen()
var ui = SpreadsheetApp.getUi();

//Run when sheet is loaded; setup custom menus
function onOpen() {
  
  ui.createMenu("YouTube")
  .addItem("New Report", "createReportSheet")
  .addItem("Process Report", "processReportSheet")
  .addToUi();
}

//Create a new sheet and fill with instructions to be used to run a report on custom date / metric
function createReportSheet() {
  ss.insertSheet();
  var reportSheet = ss.getActiveSheet();
  
  //If a sheet named 'New Report' exists, name new sheet differently
  try {
  reportSheet.setName("New Report");
  } catch (e) {
    reportSheet.setName("New Report (2)");
  }
  
  //Set Row Headers
  reportSheet.getRange(1, 1, 5).setValues([['New Report'], ['Start Date (YYYY-MM-DD)'], ['End Date (YYYY-MM-DD)'], ['Metric'], ['Dimension']]);
  
  // Set the data validation for start date
  var startDateCell = reportSheet.getRange(2, 2);
  var startDateRule = SpreadsheetApp.newDataValidation().requireDate().build();
  startDateCell.setDataValidation(startDateRule);
  
  print(startDateCell);
  print(startDateCell.getValue());
  print(startDateCell.getDisplayValue());
  
  // Set the data validation for end date
  var endDateCell = reportSheet.getRange(3, 2);
  var endDateRule = SpreadsheetApp.newDataValidation().requireDateAfter(new Date(startDateCell.getDisplayValue())).build();
  endDateCell.setDataValidation(endDateRule);
  
  //Set dropdown menu for Metric cell
  var metricCell = reportSheet.getRange(4, 2);
  var metricRule = SpreadsheetApp.newDataValidation().requireValueInList(['views', 'estimatedRevenue']).build();
  metricCell.setDataValidation(metricRule);
  
  //Set dropdown menu for Dimension cell
  var dimensionCell = reportSheet.getRange(5, 2);
  var dimensionRule = SpreadsheetApp.newDataValidation().requireValueInList(['video', 'day', 'month']).build();
  dimensionCell.setDataValidation(dimensionRule);
}

//Take current report sheet and process it into a report
function processReportSheet() {
  
  var curSheet = ss.getActiveSheet();
  
  var startDate = curSheet.getRange(2, 2).getDisplayValue();
  var endDate = curSheet.getRange(3, 2).getDisplayValue();
  var metric = curSheet.getRange(4, 2).getValue();
  var dimension = curSheet.getRange(5, 2).getValue();
  
  //If dimension == 'month', make sure conditions are met.  Otherwise, abort
  if (dimension == 'month' && startDate.substring(8, 10) != '01') {
      var ui = SpreadsheetApp.getUi();
      ui.alert('When using the \'month\' dimension, the start and end dates must be equal to the first day of a month.');
    } else {
  
      //Set the value of the 2 'date' cells to a character, otherwise it will try to format new data as a date
      //Use opportunity to set cell as error message, as user should never actually see the following data in the cell
      for (var i = 2; i <= 5; i++) {
        var cell = curSheet.getRange(i, 2);
        cell.clearDataValidations();
        cell.setValue('ERROR in processReportSheet()');
      }
      
      curSheet.getRange(3, 2).setValue('ERROR in processReportSheet()');
      
      curSheet.clear();
      
      //If a sheet named same already exists, name new sheet differently
      try {
        curSheet.setName(metric + " | " + startDate + ' - ' + endDate);
      } catch (e) {
        curSheet.setName(startDate + ' - ' + endDate + "(2)");
      }
      
      spreadsheetAnalytics(startDate, endDate, metric, dimension, curSheet);
      
      }
  } 

function spreadsheetAnalytics(startDate, endDate, metric, dimension, reportSheet) {
  
  // Get the channel ID
  var myChannels = YouTube.channelsList('id', {mine: true});
  var channel = myChannels[0];
  var channelId = channel.id;
  
  //Query YouTube Analytics API for data based on selected metrics
  var analyticsResponse = YouTubeAnalytics.reportsQuery('channel==MINE', 
                                                        startDate, 
                                                        endDate, 
                                                        metric, 
                                                        {
                                                        dimensions: dimension,
                                                        'max-results': '200',
                                                        sort: '-' + metric
                                                        });

  //Equal to the number of videos that earned revenue for the given date range, less than maxResults
  var numRows = analyticsResponse.rows.length;
  
  //Returns [videoID,earnings] - analyticsResponse.rows[0]
  //print(JSON.stringify(analyticsResponse));
  //print(analyticsResponse.rows[0]);
  
  if (dimension == 'video') {
  
    /**
    For VIDEO dimension
    */
    
    //Create reference to sheet which holds video percentage ownership data
    var pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Percentages");
    
    
    
    //Set headers
    if (metric == "views")
      reportSheet.getRange(1, 1, 1, 5).setValues([['Video Title', 'Views', 'Bryan %', 'Liam %', 'Misc']]).setFontWeight('bold');
    else
      reportSheet.getRange(1, 1, 1, 8).setValues([['Video Title', 'Revenue', 'Bryan %', 'Liam %', 'Bryan Revenue', 'Liam Revenue', 'Bryan Total', 'Liam Total']]).setFontWeight('bold');
    
    //Loop through the number of videos returned
    for (var i = 1; i <= numRows; i++) {
      
      //Extract video name using the following two lines (instead of just listing video id)
      var vList = YouTube.videosList('snippet', {'id': analyticsResponse.rows[i - 1][0]});
      var vidName = vList[0].snippet.title;
      
      //Set the first column, current row, to the video name
      reportSheet.getRange(i + 1, 1).setValue(vidName);
      
      //Set the second column, current row, to the video metric
      reportSheet.getRange(i + 1, 2).setValue(analyticsResponse.rows[i - 1][1]);
    }
    
    
    // IMPORT PERCENTAGE OWNERSHIP
    var pNames = pSheet.getRange(2, 1, pSheet.getLastRow()).getDisplayValues();
    
    //Gather all names in reportSheet into an array (of arrays, each with 1 element) to sort through without having to always call getRange().getDisplayValue() for each name
    var reportSheetNames = reportSheet.getRange(1, 1, reportSheet.getLastRow(), 1).getValues();
    
    //Add empty element at the beginning of the array to line up pNames and reportSheetNames
    reportSheetNames.unshift([0]);
    
    //Gather all percentage values in pSheet into an array (of arrays, each with 1 element) to insert into reportSheet without having to call getRange() for every row
    var pSheetValues = pSheet.getRange(2, 2, pSheet.getLastRow(), 2).getDisplayValues();
    
    //Gather sorted video revenue into an array
    var reportSheetValues = reportSheet.getRange(1, 2, reportSheet.getLastRow(), 1).getValues();
    
    //For each video that generated revenue
    for (var i = 2; i < numRows + 2; i++) {
      
      //Loop through pNames until matching video title is found
      for (var j = 1; j < pNames.length; j++) {
        
        //If the current video name from reportSheetNames matches the row we're searching in pNames
        if (reportSheetNames[i][0] == pNames[j - 1][0]) {
          
          //Insert the percentages from pSheetValues next to the video in reportSheet
          reportSheet.getRange(i, 3, 1, 2).setValues([pSheetValues[j - 1]]);
          
          //Put the total revenue earned by user in their columns
          reportSheet.getRange(i, 5, 1, 2).setValues([[pSheetValues[j - 1][0] * 0.01 * reportSheetValues[i - 1][0], //User 1 revenue
                                                      pSheetValues[j - 1][1] * 0.01 * reportSheetValues[i - 1][0]]]); //User 2 revenue
          
          //Break out of for loop
          break;
        } else {
          if (j == pNames.length - 1) {
            ui.alert("Could not find match for " + reportSheet.getRange(i, 1).getDisplayValue());
          }
        }
      }
    }
    
    //Put totals for each user above
    reportSheet.getRange(2, reportSheet.getLastColumn() - 1, 1, 2).setFormulas([["=SUM(E2:E)", "=SUM(F2:F)"]]);
    
    
    
    
    // -- PERCENTAGES
    
    
  } else {
  
    /**
    For TIME dimensions
    */
    
    //Loop through the number of days returned
    for (var i = 1; i < numRows; i++) {
      
      //Set the first column, current row, to the TIMEPERIOD (day, month, etc.)
      reportSheet.getRange(i, 1).setValue(analyticsResponse.rows[i-1][0]);
      
      //Set the second column, current row, to the video metric
      reportSheet.getRange(i, 2).setValue(analyticsResponse.rows[i-1][1]);
    }
    
  }

} 

function print(string) {
  Logger.log(string);
}



  // The YouTubeAnalytics.Reports.query() function has four required parameters and one optional
  // parameter. The first parameter identifies the channel or content owner for which you are
  // retrieving data. The second and third parameters specify the start and end dates for the
  // report, respectively. The fourth parameter identifies the metrics that you are retrieving.
  // The fifth parameter is an object that contains any additional optional parameters
  // (dimensions, filters, sort, etc.) that you want to set
