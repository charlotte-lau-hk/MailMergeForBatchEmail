//////////////////////////////////////////////////////////////////////////////////////////
// The code below is from the 'Reading Spreadsheet data using JavaScript Objects'
// tutorial (2016-08-16) with modifications later on.
// Ref: https://developers.google.com/apps-script/articles/mail_merge
//////////////////////////////////////////////////////////////////////////////////////////

// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
// [2025-03-28] Use Base64 embeddng if opts is not null
function fillInTemplateFromObject_(template, data, opts = null) {
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);

  if (templateVars != null) { // [2017-01-13] bug fix to avoid the case of no merge field
    // Replace variables from the template with the actual values from the data object.
    // If no value is available, replace with the empty string.
    for (var i = 0; i < templateVars.length; ++i) {
      // normalizeHeader ignores ${""} so we can call it directly here.
      var variableHead = normalizeHeader_(templateVars[i]);
      var variableData = data[variableHead];
      // [2019-04-23] updated to avoid writing zero as empty string
      variableData = (variableData === undefined)? "" : variableData;  
      // [2022-04-21] Adding to embed image or qr code
      // [2025-03-28] Replace deprecated substr() with substring()
      if ( (variableHead.substring(0,7)=="imglink" ||
            variableHead.substring(0,7)=="imgfile" ||
            variableHead.substring(0,6)=="qrdata"
           ) && variableData !="") {
        if (!opts) {
          variableData = "<img src='cid:"+variableHead+"'>";  
        } else {
          let qrApiUrl = opts["qrApiUrl"];
          let folder = opts["folder"];
          if (variableHead.substring(0,7)=="imglink") {
            variableData = "<img src='"+variableData+"'>";  
          } else {
            let imgBlob = null;
            if ((variableHead.substring(0,7)=="imgfile") && (folder != null)) {
              let files = folder.getFilesByName(variableData);
              if (files.hasNext()) {
                imgBlob = files.next().getBlob().setName(variableHead);
              }
            } else {
              // this case: variableHead.substring(0,6)=="qrdata"
              var qrImgLnk = qrApiUrl + encodeURI(variableData);
              try {
                imgBlob = UrlFetchApp.fetch(qrImgLnk).getBlob().setName(variableHead);
              } catch (e) {
                Logger.log("QR code fetch failed for " + variableHead + ": " + e.message);
              }
            }
            if (imgBlob) {
              let contentType = imgBlob.getContentType();
              let contentData = Utilities.base64Encode(imgBlob.getBytes());
              variableData = "<img src='data:"+contentType+";base64,"+contentData+"'>";  
            }
          }
        }
      }
      email = email.replace(templateVars[i], variableData);
    }
  }

  return email;
}

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData_(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects_(range.getValues(), normalizeHeaders_(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects_(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty_(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders_(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader_(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader_(header) {
  var key = "";
  var upperCase = false;

  for (var i=0; i <header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum_(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit_(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  // [2020.03.11] If key is empty, use a dummy random string to avoice problems
  if (key.length==0) key = "_dummy"+Math.floor(Math.random()*10000);
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty_(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum_(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit_(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit_(char) {
  return char >= '0' && char <= '9';
}
