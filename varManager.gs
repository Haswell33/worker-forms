//Global sheet positions
const timestampCol = 'B'
const workerNameCol = 'C'

// Worker sheet positions
const familyStateColWorkerSheet = 'P'
const coursesStateColWorkerSheet = 'W'
const jobStateColWorkerSheet = 'X'
const emailColWorkerSheet = 'F'

// Employment sheet positions
//const contactStateColEmploymentSheet = 'F'
const schoolStateColEmploymentSheet = 'J'
const coursesStateColEmploymentSheet = 'Q'
const jobStateColEmploymentSheet = 'R'
const emailColEmploymentSheet = 'I'

// Markers, if someone will sent empty field (one of these below) 
// js script from wordpress will send to google sheet raw markers, 
// so this script need to recognize this, beside of left marker as a value
const familyStateMarker = '{{family-state}}'
const coursesStateMarker = '{{courses-state}}'
const jobStateMarker = '{{job-state}}'
const schoolStateMarker = '{{school-state}}'

// delimeters which are used in output from wp forms
const newLineDelimeter = '{||}'
const delimeter = '|||'

const employmentMarkers = {
  'C': '{{worker-name}}',
  'D': '{{worker-born-date}}',
  'E': '{{worker-born-place}}',
  'F': '{{worker-pesel}}',
  'G': '{{contact-data-address}}',
  'H': '{{contact-data-phone}}',
  'I': '{{contact-data-email}}',
  'J': '{{study-state-education}}',
  'K': '{{study-state-name}}',
  'L': '{{study-state-start-date}}',
  'M': '{{study-state-end-date}}',
  'N': '{{study-state-degree}}',
  'O': '{{study-state-count}}',
  'P': '{{study-state-speciality}}',
  'Q': '{{courses-state}}',
  'R': '{{job-state-company-name}}',
  'S': '{{job-state-position}}',
  'T': '{{job-state-place}}',
  'U': '{{job-state-start-date}}',
  'V': '{{job-state-end-date}}',
}
const workerMarkers = {
  'C': '{{worker-name}}',
  'D': '{{worker-pesel}}',
  'E': '{{worker-phone}}',
  'F': '{{worker-email}}',
  'G': '{{address-province}}',
  'H': '{{address-county}}',
  'I': '{{address-community}}',
  'J': '{{address-street}}',
  'K': '{{address-flat-number}}',
  'L': '{{address-place-number}}',
  'M': '{{address-city}}',
  'N': '{{adress-postal-code}}',
  'O': '{{address-post}}',
  'P': '{{family-state-relationship}}',
  'Q': '{{family-state-name}}',
  'R': '{{family-state-born-date}}',
  'S': '{{family-state-study}}',
  'T': '{{family-state-family-benefits}}',
  'U': '{{family-state-care-benefits}}',
  'V': '{{personal-data-alt}}',
  'W': '{{courses-state}}',
  'X': '{{job-state-company-name}}',
  'Y': '{{job-state-position}}',
  'Z': '{{job-state-place}}',
  'AA': '{{job-state-start-date}}',
  'AB': '{{job-state-end-date}}',
  'AC': '{{emergency-person-name}}',
  'AD': '{{emergency-person-street}}',
  'AE': '{{emergency-person-flat-number}}',
  'AF': '{{emergency-person-place-number}}',
  'AG': '{{emergency-person-cty}}',
  'AH': '{{emergency-person-postal-code}}',
  'AI': '{{emergency-person-post}}',
  'AJ': '{{emergency-person-phone}}',
  'AK': '{{emergency-person-email}}',
}

const styleTitle = {
  [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
  [DocumentApp.Attribute.FONT_SIZE]: 11,
  [DocumentApp.Attribute.BOLD]: true,
  [DocumentApp.Attribute.SPACING_AFTER]: 10,
  [DocumentApp.Attribute.SPACING_BEFORE]: 10}
const styleParagraph = {
  [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
  [DocumentApp.Attribute.BOLD]: false,
  [DocumentApp.Attribute.FONT_SIZE]: 8}
const signatureParagraph = {
  [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
  [DocumentApp.Attribute.BOLD]: false,
  [DocumentApp.Attribute.SPACING_BEFORE]: 30,
  [DocumentApp.Attribute.FONT_SIZE]: 8}
const subTextSignatureParagraph = {
  [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
  [DocumentApp.Attribute.BOLD]: false,
  [DocumentApp.Attribute.SPACING_BEFORE]: 0,
  [DocumentApp.Attribute.FONT_SIZE]: 7}
