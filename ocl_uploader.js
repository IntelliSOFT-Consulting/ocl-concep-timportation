function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('OCL')
    .addItem('Upload to OCL', 'uploadToOCL')
    .addToUi();
}

var DATA_ELEMENT = 0;
var CEIL_ENGLISH_NAME = 1;
var CEIL_CONCEPT_CLASS = 2;
var CEIL_CONCEPT_ID = 3;

var DATATYPE = 5;
var CAREPAY_MAP = 6;
var DOSAGE_FORM = 7;
var DOSAGE_STRENGTH = 8;
var DATA_ELEMENT_DESCRIPTION = 9;
var REPRESENTATION = 10;
var TRANSLATION_SPECIFIC_LOCALE = 11;
var CAREPAY_NAME = 12;
var ICD_MAP = 13;
var ICD_10_MAP = 14;
var ICD_10_NAME = 15;
var CPT_MAP = 16;
var SNOMED_MAP = 17;
var SNOMED_NAME = 18;


var STATUS = 21;
var LOINC_CODE = 22;

var OCL_API_ENDPOINT = "https://api.openconceptlab.org/";
var ORIGIN_SOURCE_URL = "https://www.openconceptlab.org/users/moshonk/collections/carepay-ciel/concepts/";
var CEIL_SOURCE_URL = "https://www.openconceptlab.org/orgs/CIEL/sources/CIEL/concepts/";
var ICD_10_SOURCE_URL = "/orgs/WHO/sources/ICD-10-WHO/";
var SNOMED_CT_SOURCE_URL = "/orgs/IHTSDO/sources/SNOMED-CT/";

/**
 * Parses current sheet and performs a bulk export of concepts and mappings to ocl.
 * Assumptions:
 * - The user will only click the 'Upload to OCL' button while viewing an uploadable sheet
 * - Only one name will be provided for concepts
 * - All mappings are of type 'Same As'
 * - All mappings are internal mappings
 */
function uploadToOCL() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var height = sheet.getHeight();

  //create request body
  var requestBody = '';
  for (var row = 1; row < height; row++) {
    // concepts resource
    requestBody += '{ ';
    requestBody += 'type: "Concept", '
    requestBody += 'id: "' + sheet[row][CAREPAY_MAP] + '", '; // check if provided fail otherwise
    requestBody += 'concept_class: "' + sheet[row][CEIL_CONCEPT_CLASS] + '", '; // check if provided fail otherwise
    requestBody += 'datatype: "' + sheet[row][DATATYPE] + '", ';
    requestBody += 'names: [{';
    requestBody += 'name: "' + sheet[row][CAREPAY_NAME] + '", ';
    requestBody += 'locale: "en"';
    requestBody += '}]';
    requestBody += '}\n';

    // mappings resource
    if (sheet[row][CEIL_CONCEPT_ID]) {
      // CEIL
      requestBody += ', { ';
      requestBody += 'type: "Mapping", ';
      requestBody += 'map_type: "Same As", ';
      requestBody += 'from_concept_url: "' + ORIGIN_SOURCE_URL + sheet[row][CAREPAY_MAP] + '/", ';
      requestBody += 'to_concept_url: "' + CEIL_SOURCE_URL + sheet[row][CEIL_CONCEPT_ID] + '/", ';
      requestBody += '}\n';
    }

    if (sheet[row][ICD_10_MAP]) {
      //ICD 10
      requestBody += ', { ';
      requestBody += 'type: "Mapping", ';
      requestBody += 'map_type: "Same As", ';
      requestBody += 'from_concept_url: "' + ORIGIN_SOURCE_URL + sheet[row][CAREPAY_MAP] + '/", ';
      requestBody += 'to_concept_url: "' + ICD_10_SOURCE_URL + sheet[row][ICD_10_MAP] + '/", ';
      requestBody += '}\n';

    }

    if (sheet[row][SNOMED_MAP]) {
      //SNOMED
      requestBody += ', { ';
      requestBody += 'type: "Mapping", ';
      requestBody += 'map_type: "Same As", ';
      requestBody += 'from_concept_url: "' + ORIGIN_SOURCE_URL + sheet[row][CAREPAY_MAP] + '/", ';
      requestBody += 'to_concept_url: "' + SNOMED_CT_SOURCE_URL + sheet[row][SNOMED_MAP] + '/", ';
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

  SpreadsheetApp.getUi().alert(response.getContentText("UTF-8"));
}