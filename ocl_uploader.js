function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('OCL')
    .addItem('Upload to OCL', 'uploadToOCL')
    .addToUi();
}

var CONCEPT_ID = "concept_id";
var CONCEPT_CLASS = "concept_class";
var DATATYPE = "datatype";
var CONCEPT_NAME = "concept_name";
var DOSAGE_FORM = "dosage_form";
var DOSAGE_STRENGTH = "dosage_strength";
var CIEL_ID = "ciel_id";
var ICD10_ID = "icd10_id";
var SNOMED_CT_ID = "snomed_id";

var columnToProperty = {
  "CAREPAY Map": CONCEPT_ID,
  "CIEL Concept Class": CONCEPT_CLASS,
  "DataType": DATATYPE,
  "CAREPAY Name": CONCEPT_NAME,
  "Dosage Form": DOSAGE_FORM,
  "Dosage Strength": DOSAGE_STRENGTH,
  "CIEL Concept ID": CIEL_ID,
  "ICD-10 Map": ICD10_ID,
  "SNOMED Map": SNOMED_CT_ID
};

var OCL_API_ENDPOINT = "https://api.openconceptlab.org/";
var ORIGIN_SOURCE_URL = "/users/moshonk/collections/carepay-ciel/concepts/";
var CEIL_SOURCE_URL = "/orgs/CIEL/sources/CIEL/concepts/";
var ICD10_SOURCE_URL = "/orgs/WHO/sources/ICD-10-WHO/concepts/";
var SNOMED_CT_SOURCE_URL = "/orgs/IHTSDO/sources/SNOMED-CT/concepts/";

/**
 * Parses current sheet and performs a bulk export of concepts and mappings to ocl.
 * Assumptions:
 * - The user will click the 'Upload to OCL' button while viewing an uploadable sheet
 * - Only one name will be provided for concepts
 * - All mappings are of type 'Same As'
 * - All mappings are internal mappings
 */
function uploadToOCL() {
  var range = SpreadsheetApp.getActiveSheet().getDataRange();
  var height = range.getHeight();
  var width = range.getWidth();
  var ui = SpreadsheetApp.getUi();

  var findProperty = function (header) {
    for (var column in columnToProperty) {
      if (columnToProperty.hasOwnProperty(column) && header.startsWith(column)) {
        return columnToProperty[column];
      }
    }
  }

  //create a map between resource properties and sheet rows
  var propertyToPosition = {};
  for (var col = 1; col <= width; col++) {
    var property = findProperty(range.getCell(1, col).getValue());
    if (property) {
      propertyToPosition[property] = col;
    }
  }

  if (Object.keys(propertyToPosition).length === 0 && propertyToPosition.constructor === Object) {
    //no column were mapped implying that the wrong sheet is active
    ui.alert("Please check that you are uploading the correct sheet!");
    return;
  }

  var skipRow = function(rowId) {
    var skip = false;
    // concept id should be provided
    if (!range.getCell(row, propertyToPosition[CONCEPT_ID]).getValue()) {
      skip = true;
    }
    // concept class should be provided
    if (!range.getCell(row, propertyToPosition[CONCEPT_CLASS]).getValue()) {
      skip = true;
    }
    // datatype should be provided
    if (!range.getCell(row, propertyToPosition[DATATYPE]).getValue()) {
      skip = true;
    }
    // concept name should be provided
    if (!range.getCell(row, propertyToPosition[CONCEPT_NAME]).getValue()) {
      skip = true;
    }
    return skip;
  }

  //create request body
  var requestBody = '';
  for (var row = 2; row <= height; row++) {
    if (skipRow(row)) {
      continue; // row is incomplete, skip
    }
    // concepts resource
    requestBody += '{ ';
    requestBody += 'type: "Concept", '
    requestBody += 'id: "' + range.getCell(row, propertyToPosition[CONCEPT_ID]).getValue() + '", ';
    requestBody += 'concept_class: "' + range.getCell(row, propertyToPosition[CONCEPT_CLASS]).getValue() + '", ';
    requestBody += 'datatype: "' + range.getCell(row, propertyToPosition[DATATYPE]).getValue() + '", ';
    requestBody += 'names: [{';
    requestBody += 'name: "' + range.getCell(row, propertyToPosition[CONCEPT_NAME]).getValue() + '", ';
    requestBody += 'locale: "en" ';
    requestBody += '}]';
    if (range.getCell(row, propertyToPosition[DOSAGE_FORM]).getValue()) { // assume that if dosage form is given then dosage strength is also given
      requestBody += ', extras: { ';
      requestBody += 'dosage_form: "' + range.getCell(row, propertyToPosition[DOSAGE_FORM]).getValue() + '", ';
      requestBody += 'dosage_strength: "' + range.getCell(row, propertyToPosition[DOSAGE_STRENGTH]).getValue() + '" ';
      requestBody += '}';
    }
    requestBody += '}\n';

    // mappings resource
    if (range.getCell(row, propertyToPosition[CIEL_ID]).getValue()) {
      // CEIL
      requestBody += '{ ';
      requestBody += 'type: "Mapping", ';
      requestBody += 'map_type: "Same As", ';
      requestBody += 'from_concept_url: "' + ORIGIN_SOURCE_URL + range.getCell(row, propertyToPosition[CONCEPT_ID]).getValue() + '/", ';
      requestBody += 'to_concept_url: "' + CEIL_SOURCE_URL + range.getCell(row, propertyToPosition[CIEL_ID]).getValue() + '/"';
      requestBody += '}\n';
    }

    if (range.getCell(row, propertyToPosition[ICD10_ID]).getValue()) {
      //ICD 10
      requestBody += '{ ';
      requestBody += 'type: "Mapping", ';
      requestBody += 'map_type: "Same As", ';
      requestBody += 'from_concept_url: "' + ORIGIN_SOURCE_URL + range.getCell(row, propertyToPosition[CONCEPT_ID]).getValue() + '/", ';
      requestBody += 'to_concept_url: "' + ICD10_SOURCE_URL + range.getCell(row,propertyToPosition[ICD10_ID]).getValue() + '/"';
      requestBody += '}\n';

    }

    if (range.getCell(row, propertyToPosition[SNOMED_CT_ID]).getValue()) {
      //SNOMED
      requestBody += '{ ';
      requestBody += 'type: "Mapping", ';
      requestBody += 'map_type: "Same As", ';
      requestBody += 'from_concept_url: "' + ORIGIN_SOURCE_URL + range.getCell(row, propertyToPosition[CONCEPT_ID]).getValue() + '/", ';
      requestBody += 'to_concept_url: "' + SNOMED_CT_SOURCE_URL + range.getCell(row, propertyToPosition[SNOMED_CT_ID]).getValue() + '/"';
      requestBody += '}\n';

    }
  }

  Logger.log('Request body: \n %s', requestBody);

  var options = {
    'method' : 'post',
    'contentType': 'application/json',
    'payload' : requestBody
  };
  var response = UrlFetchApp.fetch(OCL_API_ENDPOINT + '/manage/bulkimport/', options);

  ui.alert(response.getContentText("UTF-8"));

}