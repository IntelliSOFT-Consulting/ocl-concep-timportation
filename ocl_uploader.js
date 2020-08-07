function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('OCL')
    .addItem('Upload to OCL', 'sendConcepts')
    .addToUi();
}

var owner = "moshonk";
var ownerType = "User";
var source = "carepay-ciel"

var CONCEPT_ID = "concept_id";
var CONCEPT_CLASS = "concept_class";
var DATATYPE = "datatype";
var CONCEPT_NAME = "concept_name";
var DOSAGE_FORM = "dosage_form";
var DOSAGE_STRENGTH = "dosage_strength";
var CIEL_ID = "ciel_id";
var ICD10_ID = "icd10_id";
var SNOMED_ID = "snomed_id";

var columnToProperty = {
  "CAREPAY Map": CONCEPT_ID,
  "CIEL Concept Class": CONCEPT_CLASS,
  "DataType": DATATYPE,
  "CAREPAY Name": CONCEPT_NAME,
  "Dosage Form": DOSAGE_FORM,
  "Dosage Strength": DOSAGE_STRENGTH,
  "CIEL Concept ID": CIEL_ID,
  "ICD-10 Map": ICD10_ID,
  "SNOMED Map": SNOMED_ID
};

var OCL_API_ENDPOINT = "https://api.openconceptlab.org/";
var ORIGIN_SOURCE_URL = "/users/moshonk/sources/carepay-ciel/concepts/";
var CEIL_SOURCE_URL = "/orgs/CIEL/sources/CIEL/concepts/";
var ICD10_SOURCE_URL = "/orgs/WHO/sources/ICD-10-WHO/concepts/";
var SNOMED_SOURCE_URL = "/orgs/CIEL/sources/SNOMED-MVP/";

/**
 * Sets up properties required for sending concepts.
 */
function setup() {
  var range = SpreadsheetApp.getActiveSheet().getDataRange();
  var width = range.getWidth();

  //create map between concept properties and sheet rows 
  var findProperty = function (header) {
    for (var column in columnToProperty) {
      if (columnToProperty.hasOwnProperty(column) && header.startsWith(column)) {
        return columnToProperty[column];
      }
    }
  }

  //a map between concept properties and sheet rows
  var propertyToPosition = {};
  for (var col = 1; col <= width; col++) {
    var property = findProperty(range.getCell(1, col).getValue());
    if (property) {
      propertyToPosition[property] = col;
    }
  }
  return {
    totalConcepts: range.getHeight() - 1,
    propertyToPosition: propertyToPosition
  }
}

/**
 * Initiates the send process
 */
function sendConcepts() {
  var html = HtmlService
    .createTemplateFromFile('Progress')
    .evaluate()
    .setHeight(100)
    .setWidth(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sending');
}

function displayError() {
  SpreadsheetApp.getUi().alert("Please check that you are uploading the correct sheet!");
}

//helper to load Javascript file
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

/**
 * Parses documents and sends concepts in batches.
 * Assumptions:
 * - The user will only click the 'Upload to OCL' button while viewing an uploadable sheet
 * - Only one name will be provided for concepts
 * - All names are of type 'Fully Specified' so as to pass OpenMRS concept validation
 * - All mappings are of type 'Same As'
 * - All mappings are internal mappings
 */
function sendBatch(start, end, propertyToPosition) {
  Logger.log(propertyToPosition);
  Logger.log("Start: %s, End: %s", start, end);

  var range = SpreadsheetApp.getActiveSheet().getDataRange();

  var skipRow = function (rowId) {
    var skip = false;
    // skip the header row
    if (rowId === 1) {
      skip = true;
    }
    // concept id should be provided
    if (!range.getCell(rowId, propertyToPosition[CONCEPT_ID]).getValue()) {
      skip = true;
    }
    // concept class should be provided
    if (!range.getCell(rowId, propertyToPosition[CONCEPT_CLASS]).getValue()) {
      skip = true;
    }
    // datatype should be provided
    if (!range.getCell(rowId, propertyToPosition[DATATYPE]).getValue()) {
      skip = true;
    }
    // concept name should be provided
    if (!range.getCell(rowId, propertyToPosition[CONCEPT_NAME]).getValue()) {
      skip = true;
    }
    return skip;
  }

  //create request body
  var requestBody = '';
  for (var row = start + 1; row <= end; row++) {
    if (skipRow(row)) {
      continue; // data in row is incomplete, skip
    }
    // concepts resource
    requestBody += '{ ';
    requestBody += '"type": "Concept", '
    requestBody += '"id": "' + range.getCell(row, propertyToPosition[CONCEPT_ID]).getValue() + '", ';
    requestBody += '"concept_class": "' + range.getCell(row, propertyToPosition[CONCEPT_CLASS]).getValue() + '", ';
    requestBody += '"datatype": "' + range.getCell(row, propertyToPosition[DATATYPE]).getValue() + '", ';
    requestBody += '"names": [{ ';
    requestBody += '"name": "' + range.getCell(row, propertyToPosition[CONCEPT_NAME]).getValue() + '", ';
    requestBody += '"locale": "en", ';
    requestBody += '"name_type": "Fully Specified" ';
    requestBody += '}], ';
    if (range.getCell(row, propertyToPosition[DOSAGE_FORM]).getValue()) { // assume that if dosage form is given then dosage strength is also given
      requestBody += '"extras": { ';
      requestBody += '"dosage_form": "' + range.getCell(row, propertyToPosition[DOSAGE_FORM]).getValue() + '", ';
      requestBody += '"dosage_strength": "' + range.getCell(row, propertyToPosition[DOSAGE_STRENGTH]).getValue() + '" ';
      requestBody += '}, ';
    }
    requestBody += '"owner": "' + owner + '", ';
    requestBody += '"owner_type": "' + ownerType + '", ';
    requestBody += '"source": "' + source + '" ';
    requestBody += '}\n';

    // mappings resource
    if (range.getCell(row, propertyToPosition[CIEL_ID]).getValue()) {
      // CEIL
      requestBody += '{ ';
      requestBody += '"type": "Mapping", ';
      requestBody += '"map_type": "Same As", ';
      requestBody += '"from_concept_url": "' + ORIGIN_SOURCE_URL + range.getCell(row, propertyToPosition[CONCEPT_ID]).getValue() + '/", ';
      requestBody += '"to_concept_url": "' + CEIL_SOURCE_URL + range.getCell(row, propertyToPosition[CIEL_ID]).getValue() + '/", ';
      requestBody += '"owner": "' + owner + '", ';
      requestBody += '"owner_type": "' + ownerType + '", ';
      requestBody += '"source": "' + source + '"';
      requestBody += '}\n';
    }

    if (range.getCell(row, propertyToPosition[ICD10_ID]).getValue()) {
      //ICD 10
      requestBody += '{ ';
      requestBody += '"type": "Mapping", ';
      requestBody += '"map_type": "Same As", ';
      requestBody += '"from_concept_url": "' + ORIGIN_SOURCE_URL + range.getCell(row, propertyToPosition[CONCEPT_ID]).getValue() + '/", ';
      requestBody += '"to_concept_url": "' + ICD10_SOURCE_URL + range.getCell(row, propertyToPosition[ICD10_ID]).getValue() + '/", ';
      requestBody += '"owner": "' + owner + '", ';
      requestBody += '"owner_type": "' + ownerType + '", ';
      requestBody += '"source": "' + source + '" ';
      requestBody += '}\n';
    }

    if (range.getCell(row, propertyToPosition[SNOMED_ID]).getValue()) {
      //SNOMED
      requestBody += '{ ';
      requestBody += '"type": "Mapping", ';
      requestBody += '"map_type": "Same As", ';
      requestBody += '"from_concept_url": "' + ORIGIN_SOURCE_URL + range.getCell(row, propertyToPosition[CONCEPT_ID]).getValue() + '/", ';
      requestBody += '"to_concept_url": "' + SNOMED_SOURCE_URL + range.getCell(row, propertyToPosition[SNOMED_ID]).getValue() + '/", ';
      requestBody += '"owner": "' + owner + '", ';
      requestBody += '"owner_type": "' + ownerType + '", ';
      requestBody += '"source": "' + source + '" ';
      requestBody += '}\n';
    }
  }

  //Logger.log('Request body: \n %s', requestBody);

  var options = {
    'method': 'post',
    'contentType': 'application/json+jsonl',
    'headers': { 'Authorization': 'Token 48837c7228cf8e3f619360ca523c696bf50f1e96' },
    'payload': requestBody
  };
  var response = UrlFetchApp.fetch(OCL_API_ENDPOINT + '/manage/bulkimport/', options);
  Logger.log(response.getContentText("UTF-8"));

  if (JSON.parse(response.getContentText("UTF-8")).state === "PENDING") {
    return { success: true };
  } else {
    return { success: false };
  }

}
