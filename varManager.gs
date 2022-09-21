// Worker sheet positions
const familyStateColWorkerSheet = 'P'
const coursesStateColWorkerSheet = 'W'
const jobStateColWorkerSheet = 'X'

// Employment sheet positions
const contactStateColEmploymentSheet = 'F'
const schoolStateColEmploymentSheet = 'I'
const coursesStateColEmploymentSheet = 'P'
const jobStateColEmploymentSheet = 'Q'

// Markers, if someone will sent empty field (one of these below) 
// js script from wordpress will send to google sheet raw markers, 
// so this script need to recognize this, beside of left marker as a value
const familyStateMarker = '{{family-state}}'
const coursesStateMarker = '{{courses-state}}'
const jobStateMarker = '{{job-state}}'
const schoolStateMarker = '{{school-state}}'
const contactStateMarker = '{{contact-state}}'

// delimeters which are used in output from wp redge.com forms
const newLineDelimeter = '{||}'
const delimeter = '|||'

const employmentMarkers = {
  'C': '{{worker-name}}',
  'D': '{{worker-born-date}}',
  'E': '{{worker-pesel}}',
  'F': '{{contact-data-address}}',
  'G': '{{contact-data-phone}}',
  'H': '{{contact-data-email}}',
  'I': '{{study-state-education}}',
  'J': '{{study-state-name}}',
  'K': '{{study-state-start-date}}',
  'L': '{{study-state-end-date}}',
  'M': '{{study-state-degree}}',
  'N': '{{study-state-count}}',
  'O': '{{study-state-speciality}}',
  'P': '{{courses-state}}',
  'Q': '{{job-state-company-name}}',
  'R': '{{job-state-position}}',
  'S': '{{job-state-place}}',
  'T': '{{job-state-start-date}}',
  'U': '{{job-state-end-date}}',
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
