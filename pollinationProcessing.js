function processNewSubmissions(e) { // Function that does the processing. Runs automatically when new data is submitted from a form

//Logger.log(JSON.stringify(e));

  // Extract parts of timestamp
  var timestamp = e.namedValues['Timestamp'];
  var parts = timestamp.toString().split(" ");
  var date = parts[0];
  var time = parts[1];
  var pollination_ids_text;

  if (e.namedValues.hasOwnProperty('Pollination IDs')) { // Handle pollination submissions
    var name = e.namedValues['Name'];
    pollination_ids_text = e.namedValues['Pollination IDs'];
    pollination_ids = pollination_ids_text.toString().split("\n");

    // Remove duplicates of unique ids, ex. 4x1.001 (not of non-unique cross codes (ex. 4x1) from dedicated crosses).
    var cleaned_pollination_ids = removeDuplicateIds(pollination_ids);

    var design_sheet = e.source.getSheetByName('Design');
    var design_data = design_sheet.getDataRange().getValues();
    var design_header = design_data.shift();
    var cross_name_index = design_header.indexOf('cross_name');
    var cross_code_index = design_header.indexOf('cross_code');
    var submission_sheet = e.source.getSheetByName('Submissions');

    for (i = 0; i < cleaned_pollination_ids.length; i++) {
      if (!cleaned_pollination_ids[i]) {
        // No ids supplied, do nothing
      } else {
        // Match tag prefix to cross_code, get and saved matched code and name
        var pollination_id = cleaned_pollination_ids[i];
        var id_parts = cleaned_pollination_ids[i].split(".");
        var cross_code = id_parts[0];

        for(var j = 0; j<design_data.length; j++){
          if( design_data[j][cross_code_index] == cross_code ){ // Found the matching design row
            var cross_name = design_data[j][cross_name_index];
            submission_sheet.appendRow([cross_name, pollination_id, date, time, roundTimeQuarterHour(timestamp), name.toString(), 1]);
          }
        }

      }
    }

  } else if (e.namedValues.hasOwnProperty('Cross Code')) { // Handle seed counts
    
    var cross_code = e.namedValues['Cross Code'];
    var new_seeds = e.namedValues['Number of Seeds'];

    var design_sheet = e.source.getSheetByName('Design');
    var design_data = design_sheet.getDataRange().getValues();
    var design_header = design_data.shift();
    var cross_code_index = design_header.indexOf('cross_code');
    var seed_count_index = design_header.indexOf('seed_count');
    var count_date_index = design_header.indexOf('count_date');

     for(var j = 0; j<design_data.length; j++){
          if( design_data[j][cross_code_index] == cross_code ){ // Found the matching design row
            var previous_count = design_sheet.getRange(j+2,seed_count_index+1).getValue() || 0;
            new_count = parseInt(previous_count) + parseInt(new_seeds);
            design_sheet.getRange(j+2,seed_count_index+1).setValue(new_count);
            design_sheet.getRange(j+2,count_date_index+1).setValue(date);
          }
     }

  } else { // Handle remaining data types

    if (e.namedValues.hasOwnProperty('Successful Pollination IDs')) { // Handle capsule collections
      pollination_ids_text = e.namedValues['Successful Pollination IDs'];
    } else if (e.namedValues.hasOwnProperty('Failed Pollination IDs')) { // Handle failed pollinations
      pollination_ids_text = e.namedValues['Failed Pollination IDs'];
    }
    pollination_ids = pollination_ids_text.toString().split("\n");
    
    var data_sheet = e.source.getSheetByName('Submissions');
    var data = data_sheet.getDataRange().getValues();
    var header = data.shift();
    var pollination_id_index = header.indexOf('pollination_id');
    var failure_date_index = header.indexOf('failure_date');
    var failure_count_index = header.indexOf('failure_count');
    var collection_date_index = header.indexOf('collection_date');
    var collection_count_index = header.indexOf('collection_count');

    for (i = 0; i < pollination_ids.length; i++) {
      if (!pollination_ids[i]) {
        // No ids supplied, do nothing
      } else {
        // Write date, time, name and count to data sheet row where pollination_id matches
        for(var j = 0; j<=data.length; j++){
          if ((data[j][pollination_id_index] == pollination_ids[i]) && e.namedValues.hasOwnProperty('Failed Pollination IDs')){
            // Found the right row, and data submitted refers to failed pollination
            data_sheet.getRange(j+2,failure_date_index+1).setValue(date);
            data_sheet.getRange(j+2,failure_count_index+1).setValue(1);
            break;
          } else if ((data[j][pollination_id_index] == pollination_ids[i]) && e.namedValues.hasOwnProperty('Successful Pollination IDs')){
            // Found the right row, and data submitted refers to successful pollination
            data_sheet.getRange(j+2,collection_date_index+1).setValue(date);
            data_sheet.getRange(j+2,collection_count_index+1).setValue(1);
            break;
          }
        }
      }
    }
  }
}


function roundTimeQuarterHour(time) {
  var timeToReturn = new Date(time);
  timeToReturn.setMilliseconds(Math.round(timeToReturn.getMilliseconds() / 1000) * 1000);
  timeToReturn.setSeconds(Math.round(timeToReturn.getSeconds() / 60) * 60);
  timeToReturn.setMinutes(Math.round(timeToReturn.getMinutes() / 15) * 15);
  return timeToReturn.getHours() + ":" + timeToReturn.getMinutes();
}

function removeDuplicateIds(ids) {
  //Logger.log("Removing duplicates");
    var seen = {};
    return ids.filter(function(item) {
        if(item.indexOf('.') !== -1) {
          return seen.hasOwnProperty(item) ? false : (seen[item] = true);
        } else {
          return true;
        }
    });
 }
