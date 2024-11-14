function getDataFromAPI() {
    Logger.log("Starting getDataFromAPI");

    // Clear the entire sheet except for the first row
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent(); // Clears all the content of the sheet except the first row
    }

    var apiURL = 'https://api.learningstream.com/api/api_event_registration_data.aspx?aid=NCSOE'; 
    var options = {
        'method' : 'post',
        'payload' : {
            'security_key': '', // Configured in Learning Stream
            'event_code': '', // ENROLL is the event code configured in Learning Stream
            'start_date': '', // Currently these are arbitrary but should get them from Sherry
            'end_date': ''
        }
    };

    var response = UrlFetchApp.fetch(apiURL, options);
    Logger.log("API response received");

    var xml = response.getContentText();
    Logger.log("XML from response: " + xml.substring(0, 200));

    var document = XmlService.parse(xml);
    var json = xmlToJson(document.getRootElement());

    Logger.log("Converted JSON: " + JSON.stringify(json).substring(0, 200));

    if (json) {
        Logger.log("Processing data");
        var processedData = processData(json);
        insertDataIntoSheet(processedData);
    } else {
        Logger.log('No data found');
    }
}

function xmlToJson(xml) {
    var obj = {};

    if (xml.getChildren().length === 0) {
        return xml.getText();
    }

    xml.getChildren().forEach(function (child) {
        var name = child.getName();
        if (!obj[name]) {
            obj[name] = xmlToJson(child);
        } else {
            if (!Array.isArray(obj[name])) {
                var temp = obj[name];
                obj[name] = [temp];
            }
            obj[name].push(xmlToJson(child));
        }
    });

    return obj;
}

function processData(data) {
    Logger.log("Starting processData");
    var processedData = [];
    var fieldMapping = {
        'Date': 'registration_date',
        'First Name': 'first_name',
        'Last Name': 'last_name',
        'Email': 'email_address',
        'Phone': 'phone_number',
        'School': 'School',
        'NCSOE Program': 'event_title',
        'Year in Program': 'Which_year_of_Teacher_Induction_are_you_enrolling_in_If_you_are_a_Mentor_this_will_be_your_Mentees_program_year_If_supporting_more_than_one_mentee_select_all_that_apply',
        'District': 'District',
        'Enrolling_as': 'Enrolling_as'
    };
    
    var headers = Object.keys(fieldMapping);
    processedData.push(headers); // Add headers

    var records = data.registration_record;
    if (!records) {
        Logger.log("No registration records found");
        return [];
    }

    records.forEach(function (record, index) {
        if (index < 5) {
            Logger.log("Record " + index + " structure: " + JSON.stringify(record, null, 2));
        }

        // Check if this record is for a mentor before processing
        var enrollingAs = findFieldValue(record, 'Enrolling_as');
        if (enrollingAs && enrollingAs.toLowerCase().includes('mentor')) {
            Logger.log("Skipping mentor record: " + JSON.stringify(record));
            return; // Skip this record
        }

        var row = [];
        for (var i = 0; i < headers.length; i++) {
            var header = headers[i];
            var field = fieldMapping[header];
            var value = findFieldValue(record, field);
            row.push(value || '');
        }

        processedData.push(row);
        Logger.log("Added row " + index + ": " + JSON.stringify(row));
    });

    Logger.log("Completed processData");
    return processedData;
}

function findFieldValue(record, fieldName) {
    if (record.hasOwnProperty(fieldName)) {
        return record[fieldName];
    }

    if (record.registration_questions && record.registration_questions.question) {
        var questions = Array.isArray(record.registration_questions.question) 
            ? record.registration_questions.question 
            : [record.registration_questions.question];
        
        for (var i = 0; i < questions.length; i++) {
            if (questions[i].text === fieldName) {
                return questions[i].responses.response;
            }
        }
    }

    for (var prop in record) {
        if (typeof record[prop] === 'object' && record[prop] !== null) {
            var value = findFieldValue(record[prop], fieldName);
            if (value) return value;
        }
    }

    return '';
}

function insertDataIntoSheet(processedData) {
    Logger.log("Starting insertDataIntoSheet");

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var existingData = sheet.getDataRange().getValues();
    var existingIds = existingData.map(function (row) { return row[0]; });

    var headers = processedData.shift();

    if (existingData.length === 0) {
        var headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setValues([headers]);
        Logger.log("Inserted headers: " + JSON.stringify(headers));
    }

    var newData = processedData.filter(function (row) {
        return existingIds.indexOf(row[0]) === -1;
    });

    if (newData.length > 0) {
        var startRow = existingData.length + 1;
        var range = sheet.getRange(startRow, 1, newData.length, headers.length);
        range.setValues(newData);
        Logger.log("Inserted new data into sheet: " + JSON.stringify(newData));
    } else {
        Logger.log("No new data to insert");
    }

    Logger.log("Completed insertDataIntoSheet");
}
