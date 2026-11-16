/**
 * ============================================================================
 * NHS University Hospitals Birmingham - Seminar Room Booking Backend API
 * Google Apps Script Web App for Google Sheet ID:
 * https://docs.google.com/spreadsheets/d/11mlTa-Jkt_lSfrJKXBFGD6NVEsKnJ42_5SerYNFy1xE/edit
 * ============================================================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏥 NHS Room Booking')
    .addItem('🚀 Setup / Initialize Sheet Tabs', 'setupSpreadsheet')
    .addItem('📋 Test Backend Status', 'checkBackendStatus')
    .addToUi();
}

/**
 * Super-fast setup (< 2 seconds):
 * - No blocking UI popups (prevents 6-minute timeout)
 * - Fixed column widths (replaces slow autoResize)
 * - Overwrites/populates if sheet is empty or header-only
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup "Bookings" Tab
  let bookingsSheet = ss.getSheetByName('Bookings');
  if (!bookingsSheet) {
    bookingsSheet = ss.insertSheet('Bookings', 0);
  }
  const bookingHeaders = [
    'ID', 'Date', 'StartTime', 'EndTime', 'DurationMinutes',
    'BookerName', 'Email', 'Bleep', 'Department', 'Title',
    'Attendees', 'Equipment', 'PIN', 'RoomId', 'IsBlocked',
    'IsRecurring', 'RecurrenceGroupId', 'CreatedAt'
  ];
  if (bookingsSheet.getLastRow() === 0) {
    bookingsSheet.appendRow(bookingHeaders);
  }
  formatHeaderRow(bookingsSheet, '#005EB8', bookingHeaders.length);
  setStandardColumnWidths(bookingsSheet, [140, 100, 90, 90, 110, 150, 200, 100, 140, 220, 90, 180, 80, 140, 90, 90, 140, 180]);

  // 2. Setup "Rooms" Tab
  let roomsSheet = ss.getSheetByName('Rooms');
  if (!roomsSheet) {
    roomsSheet = ss.insertSheet('Rooms', 1);
  }
  const roomHeaders = ['ID', 'Name', 'Site', 'Location', 'Capacity', 'Equipment'];
  if (roomsSheet.getLastRow() <= 1) {
    roomsSheet.clearContents();
    roomsSheet.appendRow(roomHeaders);
    const defaultRooms = [
      ['qeh-seminar-room', 'Pharmacy Seminar Room', 'QEH Heritage', 'Level 1, Pharmacy Department (Opposite Dispensary)', 16, '["wall_display", "owl_camera", "whiteboard", "flipchart", "polycom_phone"]'],
      ['qeh-conf-a', 'Clinical Trials Conf Room A', 'QEH Heritage', 'Level 2, Clinical Research Facility', 10, '["wall_display", "whiteboard"]'],
      ['ghh-seminar', 'Education Centre Seminar 2', 'Good Hope', 'Lower Ground Floor, Education Centre', 14, '["wall_display", "owl_camera", "flipchart"]'],
      ['bh-conf-1', 'Heartlands Pharmacy Conf Room', 'Heartlands', 'Pharmacy Office Suite, Block 23', 12, '["wall_display", "whiteboard", "polycom_phone"]'],
      ['sh-meeting', 'Solihull Ambulatory Meeting Room', 'Solihull', 'Ground Floor, Outpatients Wing', 8, '["wall_display"]']
    ];
    roomsSheet.getRange(2, 1, defaultRooms.length, roomHeaders.length).setValues(defaultRooms);
  }
  formatHeaderRow(roomsSheet, '#005EB8', roomHeaders.length);
  setStandardColumnWidths(roomsSheet, [140, 220, 120, 300, 90, 350]);

  // 3. Setup "Users" Tab (270 Staff Directory)
  let usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    usersSheet = ss.insertSheet('Users', 2);
  }
  const userHeaders = ['ID', 'Name', 'Email', 'Bleep', 'Department', 'Site', 'Role', 'PIN', 'IsActive'];
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.clearContents();
    usersSheet.appendRow(userHeaders);
    const initialStaff = [
      ['usr-001', 'Abdallah Alkhalaf', 'Abdallah.Alkhalaf@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-002', 'Abideen Bishi', 'Abideen.Bishi@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-003', 'Abigail Jenkins', 'Abigail.Jenkins@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-004', 'Abigail Parkes', 'Abigail.Parkes@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-005', 'Abigail Whorton', 'Abigail.Whorton@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-006', 'Abimbola Osuntade', 'Abimbola.Osuntade@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-007', 'Adam Mahroof', 'Adam.Mahroof@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-008', 'Ajith Vayyattu Manivarnan', 'AjithVayyattu.Manivarnan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-009', 'Alex Wong', 'Alex.Wong@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-010', 'Alexander Smith', 'Alexander.Smith@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-011', 'Alice Mallinder', 'Alice.Mallinder@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-012', 'Alice Tew', 'Alice.Tew@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-013', 'Alisha Patel', 'Alisha.Patel@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-014', 'Alison Maidment', 'Alison.Maidment@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-015', 'Amanda Li', 'Amanda.Li@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-016', 'Amandeep Kaur Goyal', 'Amandeep.KaurGoyal@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-017', 'Ameera Basheerkutty', 'Ameera.Basheerkutty@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-018', 'Ameera Gulshan', 'Ameera.Gulshan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-019', 'Amina Khan', 'Amina.Khan5@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-020', 'Amisha Desai', 'Amisha.Desai@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-021', 'Amreen Luhar', 'Amreen.Luhar@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-022', 'Amrit Atwal', 'Amritpal.Atwal@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-023', 'Anet Paul', 'Anet.Paul@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-024', 'Angela Betts', 'Angela.Betts@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-025', 'Angelika Khafizi', 'Angelika.Khafizi@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-026', 'Anil Badhan', 'Anil.Badhan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-027', 'Anisah Kassim', 'Anisah.Kassim2@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-028', 'Anita Sangera', 'Anita.Sangera@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-029', 'Aniya Mohiuddin', 'Aniya.Mohiuddin@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-030', 'Antonio Ochoa-Ferraro', 'Antonio.Ochoa-Ferraro@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-031', 'Anuj Mendhule', 'Anuj.Mendhule@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-032', 'Anum Kashif', 'Anum.Kashif@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-033', 'Anum Rahman', 'Anum.Rahman@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-034', 'Aqsa Ahmed', 'Aqsa.Ahmed@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-035', 'Ariba Afzal', 'Ariba.Afzal@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-036', 'Arti Dasoar', 'Arti.Dasoar@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-037', 'Ashutosh Goel', 'Ashutosh.Goel@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-038', 'Asif Sarwar', 'Asif.Sarwar@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-039', 'Ayesha Mehmood', 'Ayesha.Mehmood@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-040', 'Azka Naveed', 'Azka.Naveed@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-041', 'Baljit Ahitan', 'Baljit.Ahitan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-042', 'Baljit Kaur', 'Baljit.Kaur3@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-043', 'Baljodh Singh', 'Baljodh.Singh@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-044', 'Bethan Zanovello', 'Bethan.Zanovello@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-045', 'Callum Holt', 'Callum.Holt@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-046', 'Carly Haytree', 'Carly.Haytree@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-047', 'Carmen Statter-Tsang', 'Carmen.Statter-Tsang@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-048', 'Charlotte Alsop', 'Charlotte.Alsop@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-049', 'Charlotte James', 'Charlotte.James@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-050', 'Christy Chung', 'Christy.Chung@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-051', 'Claire Bourne', 'Claire.Bourne@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-052', 'Claire Draycott', 'Claire.Draycott@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-053', 'Claire Russon', 'Claire.Russon@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-054', 'Daniel Welch', 'Daniel.Welch@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-055', 'David Carter', 'David.Carter@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-056', 'David Langley', 'David.Langley@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-057', 'David Stoneystreet', 'David.Stoneystreet@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-058', 'Debby Edwards', 'Debby.Edwards@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-059', 'Deepak Chahal', 'Deepak.Chahal@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-060', 'Derek Robertson', 'Derek.Robertson2@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-061', 'Diana Kawonga', 'Diana.Kawonga@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-062', 'Dilem Cantay', 'Dilem.Cantay@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-063', 'Dominika Mazur', 'Dominika.Mazur@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-064', 'Ejiroghene Okoro', 'Ejiroghene.Okoro@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-065', 'Ekta Punj', 'Ekta.Punj@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-066', 'Eleanor Jones', 'Eleanor.Jones@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-067', 'Eleanor Lines', 'Eleanor.Lines@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-068', 'Ellie Jones', 'Ellie.Jones2@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-069', 'Emma Coleman', 'Emma.Coleman@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-070', 'Emman Abdullah', 'Emman.Abdullah@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-071', 'Erika Maculevica', 'Erika.Maculevica@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-072', 'Erika Welbourne', 'Erika.Welbourne@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-073', 'Faiz Alam', 'Faiz.Alam@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-074', 'Faraeen Avestaei', 'Faraeen.Avestaei@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-075', 'Farhana Haque', 'Farhana.Haque@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-076', 'Fathima Begum', 'Fathima.Begum@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-077', 'Ferin Koshy', 'Ferin.Koshy@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-078', 'Fesani Mahmood', 'Fesani.Mahmood@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-079', 'Fiona Pellowe', 'Fiona.Pellowe@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-080', 'Gabriela Szozda', 'Gabriela.Szozda@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-081', 'Gayatri Parekh', 'Gayatri.Parekh@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-082', 'Gobind Chaggar', 'Gobind.Chaggar@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-083', 'Gopika Suresh', 'Gopika.Suresh@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-084', 'Gurjinder Mudhar', 'Gurjinder.Mudhar@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-085', 'Gurjit Sohal', 'Gurjit.Sohal@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-086', 'Gursharon Takhar', 'Gursharon.Takhar@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-087', 'Hadeeqa Sajjad', 'Hadeeqa.Sajjad@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-088', 'Hafiza Khanum', 'Hafiza.Khanum@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-089', 'Halima Sadia Rashid', 'Halima.SadiaRashid@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-090', 'Hamza Hussain', 'Hamza.Hussain4@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-091', 'Hamzah Ghiwala', 'Hamzah.Ghiwala@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-092', 'Hanand Badesha', 'Hanand.Badesha@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-093', 'Harpreet Khangooda', 'Harpreet.Khangooda@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-094', 'Harun Saleh', 'Harun.Saleh@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-095', 'Harvinder Sandhu', 'Harvinder.Sandhu@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-096', 'Hasan Varachhia', 'Hasan.Varachhia@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-097', 'Haseeb Ahmed', 'Haseeb.Ahmed@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-098', 'Helen Laws', 'Helen.Laws@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-099', 'Holli Brown', 'Holli.Brown2@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-100', 'Hyacinth Wong', 'Hyacinth.Wong@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-101', 'Hyador Ali', 'Hyador.Ali@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-102', 'Ian Reading', 'Ian.Reading@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-103', 'Ieuan Morgan', 'Ieuan.Morgan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-104', 'Ihsaan Anwar', 'Ihsaan.Anwar@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-105', 'Inderjit Singh', 'Inderjit.Singh@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-106', 'Iqbal Singh', 'Iqbal.Singh@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-107', 'Iqra Habib', 'Iqra.Habib@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-108', 'Isaque Khan', 'Isaque.Khan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-109', 'Jack Stone', 'Jack.Stone@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-110', 'Jackie Shale', 'Jackie.Shale@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-111', 'Jacob Day', 'Jacob.Day@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-112', 'Jagtej Rye', 'Jagtej.Rye@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-113', 'James Laws', 'James.Laws@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-114', 'Jane Dawes', 'Jane.Dawes2@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-115', 'Jas Sall', 'Jas.Sall@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-116', 'Jasharon Sandhu', 'Jasharon.Sandhu@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-117', 'Jasmine Skinner', 'Jasmine.Skinner@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-118', 'Jasraj Singh', 'Jasraj.Singh@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-119', 'Jaswant Kaur', 'Jaswant.Kaur@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-120', 'Javeria Maskin', 'Javeria.Maskin@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-121', 'Jaymisha Mistry', 'Jaymisha.Mistry@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-122', 'Jeff Aston', 'Jeff.Aston@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-123', 'Jenni Hatton', 'Jenni.Hatton@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-124', 'Jennifer de Val', 'Jennifer.deVal@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-125', 'Jennifer Young', 'Jennifer.Young@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-126', 'Jesna Varghese', 'Jesna.Varghese@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-127', 'Jiaqian Xie', 'Jiaqian.Xie@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-128', 'Jonalyn Marzan', 'Jonalyn.Marzan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-129', 'Joshua Roberts', 'Joshua.Roberts@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-130', 'Justin Freeman', 'Justin.Freeman@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-131', 'Kadeer Akhtar', 'Kadeer.Akhtar@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-132', 'Kai-Yee Chan', 'Kai-Yee.Chan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-133', 'Kamaljeet Kaur', 'Kamaljeet.Kaur@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-134', 'Karandeep Singh', 'Karandeep.Singh@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-135', 'Karum Kaur', 'Karum.Kaur@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-136', 'Kaye Garratt', 'Kaye.Garratt@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-137', 'Kiera Corless', 'Kiera.Corless@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-138', 'Kimberley Steele', 'Kimberley.Steele@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-139', 'Kiran Kadara', 'Kiran.Kadara@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-140', 'Kuldip Kaur', 'Kuldip.Kaur2@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-141', 'Kuldip Kaur', 'Kuldip.Kaur@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-142', 'Laura Clarkson', 'Laura.Clarkson@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-143', 'Laura Kettlewell', 'Laura.Kettlewell@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-144', 'Laura-Jane Blendell', 'Laura-Jane.Blendell@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-145', 'Leah Efuna', 'Leah.Efuna@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-146', 'Leila Fares', 'Leila.Fares@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-147', 'Lisa Havard', 'Lisa.Havard@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-148', 'Liz Russell', 'Liz.Russell@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-149', 'Liz Willmott', 'Liz.Willmott@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-150', 'Lorraine Watson', 'Lorraine.Watson@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-151', 'Louisa Knowles', 'Louisa.Knowles@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-152', 'Lubna Khan', 'Lubna.Khan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-153', 'Lynn Ansell', 'Lynn.Ansell@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-154', 'Lynn Do', 'LynnNguyen.Do@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-155', 'Manisha Suman', 'Manisha.Suman@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-156', 'Maria Jackson', 'Maria.Jackson@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-157', 'Mariam Hussain', 'Mariam.Hussain@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-158', 'Marie Malange', 'Marie.Malange@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-159', 'Mark McGuirk', 'Mark.McGuirk@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-160', 'Maryam Shahzadi', 'Maryam.Shahzadi@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-161', 'Maryam Sharifi', 'Maryam.Sharifi@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-162', 'Matthew ODell', 'Matthew.ODell@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-163', 'Matthew Underhill', 'Matthew.Underhill@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-164', 'Maymoona Ibraheem', 'Maymoona.Ibraheem@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-165', 'Mayur Mistry', 'Mayur.Mistry@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-166', 'Mohammed Ibrahim', 'Mohammed.Ibrahim@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-167', 'Mohammed Shlaka', 'Mohammed.Shlaka@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-168', 'Mollie Harper', 'Mollie.Harper@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-169', 'Muhammed Sajid', 'Muhammed.Sajid@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-170', 'Muntazah Mockbill', 'Muntazah.Mockbill@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-171', 'Murtaza Qadir', 'Murtaza.Qadir@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-172', 'Muzafar Iqbal', 'Muzafar.Iqbal@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-173', 'Nafees Rafiq', 'Nafees.Rafiq@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-174', 'Nagina Koser', 'Nagina.Koser@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-175', 'Natasha Brown', 'Natasha.Brown@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-176', 'Nick Duncan', 'Nick.Duncan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-177', 'Nicola Holdsworth', 'Nicola.Holdsworth@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-178', 'Nicola Marchant', 'Nicola.Marchant@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-179', 'Nicola Turner', 'Nicola.Turner@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-180', 'Nicole Bagdi', 'Nicole.Bagdi@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-181', 'Nida Kazmi', 'Nida.Kazmi@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-182', 'Nida Manan', 'Nida.Manan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-183', 'Nina Hasso', 'Nina.Hasso@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-184', 'Nisa Varghese', 'Nisa.Varghese@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-185', 'Nita Shah', 'Nita.Shah@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-186', 'Nuha Hussain', 'Nuha.Hussain@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-187', 'Odunola Popoola', 'Odunola.Popoola@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-188', 'oi ming ivy chow', 'oimingivy.chow@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-189', 'Oluebube Favour Okoli', 'OluebubeFavour.Okoli@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-190', 'Oluwole Akinpelu', 'Oluwole.Akinpelu@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-191', 'Owen Turner', 'Owen.Turner@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-192', 'Parvinder Dalay', 'Parvinder.Dalay@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-193', 'Patricia Bhebhe', 'Patricia.Bhebhe@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-194', 'Patrick Molloy', 'Patrick.Molloy@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-195', 'Paul Smillie', 'Paul.Smillie@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-196', 'Pavindeep Gill', 'Pavindeep.Gill@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-197', 'Penny Mervyn', 'Penny.Mervyn@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-198', 'Perminder Sunner', 'Perminder.Sunner@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-199', 'Piush Champaneri', 'Piush.Champaneri@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-200', 'Pooja Sandhu', 'Pooja.Sandhu@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-201', 'Pulvinder Kandola', 'Pulvinder.Kandola@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-202', 'Rachel Johnson', 'Rachel.Johnson@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-203', 'Rahul Kumar', 'Rahul.Kumar@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-204', 'Rameesa Qureshi', 'Rameesa.Qureshi@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-205', 'Ranu Ali', 'Ranu.Ali@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-206', 'Rawan Ahmed', 'Rawan.Ahmed@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-207', 'Rebecca Tongue', 'Rebecca.Tongue@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-208', 'Rehan Ashraf', 'Rehan.Ashraf@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-209', 'Richard Cole', 'Richard.Cole@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-210', 'Robert Crook', 'Robert.Crook@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-211', 'Ruqayyah Yaqub', 'Ruqayyah.Yaqub@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-212', 'Saba Ashraf', 'Saba.Ashraf3@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-213', 'Saba Yafaie', 'Saba.Yafaie@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-214', 'Sabah Farooque', 'Sabah.Farooque@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-215', 'Sabareesh Dinesh', 'Sabareesh.Dinesh@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-216', 'Sabeeah Javed', 'Sabeeah.Javed@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-217', 'Sabena Begum', 'Sabena.Begum@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-218', 'Safia Asad', 'Safia.Asad@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-219', 'Saghir Hussain', 'Saghir.Hussain@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-220', 'Saika LLoyd', 'Saika.LLoyd@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-221', 'Saju John', 'Saju.John@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-222', 'Salma Hassan', 'Salma.Hassan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-223', 'Sam Stainforth', 'Sam.Stainforth@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-224', 'Samantha Murray', 'Samantha.Murray@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-225', 'Sami Chowdhury', 'Sami.Chowdhury@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-226', 'Samira Bi', 'Samira.Bi@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-227', 'Samiyah Rafiq', 'Samiyah.Rafiq@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-228', 'Sanah Sheraz', 'Sanah.Sheraz@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-229', 'Sarah Harper', 'Sarah.Harper2@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-230', 'Sawan Singh', 'Sawan.Singh@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-231', 'Shakeel Duberia', 'Shakeel.Duberia@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-232', 'Shinaz Hashem', 'Shinaz.Hashem@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-233', 'Shuaib Elahi', 'Shuaib.Elahi@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-234', 'Simendeep Kaur', 'Simendeep.Kaur@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-235', 'Simran Dhadda', 'Simran.Dhadda@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-236', 'Simran Kaur', 'Simran.Kaur4@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-237', 'Sonia Iqbal', 'Sonia.Iqbal@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-238', 'Sooraj Thomas Raju', 'Sooraj.ThomasRaju@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-239', 'Sophie Edwards', 'Sophie.Edwards@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-240', 'Sophie Rattigan', 'Sophie.Rattigan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-241', 'Sradha Sasi', 'Sradha.Sasi@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-242', 'Sreelakshmi Gopakumar', 'Sreelakshmi.Gopakumar@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-243', 'Stephanie Murray', 'Stephanie.Murray@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-244', 'Stephen Hill', 'Stephen.Hill@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-245', 'Steven Cho', 'Steven.Cho@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-246', 'Sumayyah Khan', 'Sumayyah.Khan@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-247', 'Sunny Mahmood', 'Sunny.Mahmood@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-248', 'Susan Dhopatkar', 'Susan.Dhopatkar@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-249', 'Sutej Sivia', 'Sutej.Sivia@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-250', 'Tajrin Ahmed', 'Tajrin.Ahmed@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-251', 'Tamar David', 'Tamar.David@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-252', 'Tasnim Rezwana', 'Tasnim.Rezwana@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-253', 'Teresa Chu', 'Teresa.Chu@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-254', 'Thomas Bicknell', 'Thomas.Bicknell@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-255', 'Timothy Redway', 'Timothy.Redway@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-256', 'Titilope Familusi', 'Titilope.Familusi@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-257', 'Uduakobong Bassey', 'Uduakobong.Bassey@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-258', 'Vaishali Dulobdas', 'Vaishali.Dulobdas@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-259', 'Vongai Mucheri', 'Vongai.Mucheri@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-260', 'Waqar Raza', 'Waqar.Raza@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-261', 'Yanyan Ng', 'Yanyan.Ng@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-262', 'Yin May Chin', 'YinMay.Chin@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-263', 'Yumnaa Mughal', 'Yumnaa.Mughal@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-264', 'Zaheda Begum', 'Zaheda.Begum@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-265', 'Zahra Irshad', 'Zahra.Irshad@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-266', 'Zahraa Karim', 'Zahraa.Karim@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-267', 'Zain Hussain', 'Zain.Hussain2@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-268', 'Zainab Sami Faris', 'Zainab.SamiFaris@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-269', 'Zainab Zara', 'Zainab.Zara@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true],
      ['usr-270', 'Zakir Hussain', 'Zakir.Hussain@uhb.nhs.uk', '', 'Pharmacy', 'QEH Heritage', 'staff', '1234', true]
    ];
    if (initialStaff.length > 0) {
      usersSheet.getRange(2, 1, initialStaff.length, userHeaders.length).setValues(initialStaff);
    }
  }
  formatHeaderRow(usersSheet, '#005EB8', userHeaders.length);
  setStandardColumnWidths(usersSheet, [100, 200, 240, 100, 140, 120, 80, 80, 80]);

  // Clean default Sheet1 if empty
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1 && sheet1.getLastRow() <= 1) {
    try { ss.deleteSheet(sheet1); } catch (e) {}
  }

  Logger.log('Setup finished in < 2 seconds.');
  return { status: 'success', message: 'All 3 tabs successfully configured with 270 staff and rooms.' };
}

function setStandardColumnWidths(sheet, widths) {
  for (let i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}

function formatHeaderRow(sheet, colorHex, colCount) {
  const range = sheet.getRange(1, 1, 1, colCount);
  range.setBackground(colorHex);
  range.setFontColor('#FFFFFF');
  range.setFontWeight('bold');
  range.setFontSize(10);
  range.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 30);
}

function checkBackendStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const b = ss.getSheetByName('Bookings') ? Math.max(0, ss.getSheetByName('Bookings').getLastRow() - 1) : 0;
  const u = ss.getSheetByName('Users') ? Math.max(0, ss.getSheetByName('Users').getLastRow() - 1) : 0;
  const r = ss.getSheetByName('Rooms') ? Math.max(0, ss.getSheetByName('Rooms').getLastRow() - 1) : 0;
  Logger.log('Status: Bookings=' + b + ', Users=' + u + ', Rooms=' + r);
  try {
    SpreadsheetApp.getUi().alert('📊 Backend Status:\n\n• Bookings: ' + b + '\n• Staff Users: ' + u + '\n• Rooms: ' + r);
  } catch (e) {}
}

/**
 * GET Handler (REST API for PWA / Web Client)
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'ping';
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Direct Web Setup Trigger
    if (action === 'setup' || action === 'init') {
      const res = setupSpreadsheet();
      return jsonResponse(res);
    }

    if (action === 'ping') {
      const bCount = ss.getSheetByName('Bookings') ? Math.max(0, ss.getSheetByName('Bookings').getLastRow() - 1) : 0;
      const uCount = ss.getSheetByName('Users') ? Math.max(0, ss.getSheetByName('Users').getLastRow() - 1) : 0;
      const rCount = ss.getSheetByName('Rooms') ? Math.max(0, ss.getSheetByName('Rooms').getLastRow() - 1) : 0;
      return jsonResponse({
        status: 'ok',
        sheetId: ss.getId(),
        sheetTitle: ss.getName(),
        bookingsCount: bCount,
        usersCount: uCount,
        roomsCount: rCount,
        serverTime: new Date().toISOString()
      });
    }

    if (action === 'getBookings') {
      const sheet = ss.getSheetByName('Bookings');
      if (!sheet || sheet.getLastRow() < 2) return jsonResponse({ status: 'success', data: [] });
      const rows = getSheetDataAsObjects(sheet);
      return jsonResponse({ status: 'success', data: rows });
    }

    if (action === 'getUsers') {
      const sheet = ss.getSheetByName('Users');
      if (!sheet || sheet.getLastRow() < 2) return jsonResponse({ status: 'success', data: [] });
      const rows = getSheetDataAsObjects(sheet);
      return jsonResponse({ status: 'success', data: rows });
    }

    if (action === 'getRooms') {
      const sheet = ss.getSheetByName('Rooms');
      if (!sheet || sheet.getLastRow() < 2) return jsonResponse({ status: 'success', data: [] });
      const rows = getSheetDataAsObjects(sheet);
      return jsonResponse({ status: 'success', data: rows });
    }

    if (action === 'getAll') {
      const bSheet = ss.getSheetByName('Bookings');
      const uSheet = ss.getSheetByName('Users');
      const rSheet = ss.getSheetByName('Rooms');
      return jsonResponse({
        status: 'success',
        bookings: bSheet ? getSheetDataAsObjects(bSheet) : [],
        users: uSheet ? getSheetDataAsObjects(uSheet) : [],
        rooms: rSheet ? getSheetDataAsObjects(rSheet) : []
      });
    }

    return jsonResponse({ status: 'error', message: 'Unknown GET action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/**
 * POST Handler (Create, Update, Delete for PWA Client)
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter && e.parameter.payload) {
      payload = JSON.parse(e.parameter.payload);
    }

    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Add or Update Booking (Upsert)
    if (action === 'addBooking' || action === 'saveBooking') {
      const sheet = ss.getSheetByName('Bookings') || ss.insertSheet('Bookings');
      const b = payload.booking || {};
      const newId = b.id || ('QE-SR-' + new Date().getTime());
      const data = sheet.getDataRange().getValues();
      let foundRow = -1;
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][0]) === String(newId)) {
          foundRow = r + 1;
          break;
        }
      }
      const rowValues = [
        newId,
        b.date || '',
        "'" + (b.startTime || '09:00'),
        "'" + (b.endTime || '09:30'),
        b.durationMinutes || 30,
        b.bookerName || '',
        b.email || '',
        b.bleep || '',
        b.department || 'Pharmacy',
        b.title || 'Seminar Meeting',
        b.attendees || 1,
        Array.isArray(b.equipment) ? JSON.stringify(b.equipment) : '[]',
        b.pin || '',
        b.roomId || 'qeh-seminar-room',
        Boolean(b.isBlocked),
        Boolean(b.isRecurring),
        b.recurrenceGroupId || '',
        b.createdAt || new Date().toISOString()
      ];
      if (foundRow > 0) {
        sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
        let emailRes = null;
        if (payload.sendEmailInvite === true) {
          try { emailRes = sendCalendarInviteEmail(b); } catch (e) { emailRes = { sent: false, error: e.toString() }; Logger.log('Email error: ' + e); }
        }
        return jsonResponse({ status: 'success', action: 'updated', id: newId, emailResult: emailRes });
      } else {
        sheet.appendRow(rowValues);
        let emailRes = null;
        if (payload.sendEmailInvite !== false && b.sendEmailInvite !== false) {
          try { emailRes = sendCalendarInviteEmail(b); } catch (e) { emailRes = { sent: false, error: e.toString() }; Logger.log('Email error: ' + e); }
        }
        return jsonResponse({ status: 'success', action: 'inserted', id: newId, emailResult: emailRes });
      }
    }

    // 2. Update Booking
    if (action === 'updateBooking') {
      const sheet = ss.getSheetByName('Bookings');
      if (!sheet) return jsonResponse({ status: 'error', message: 'Bookings sheet not found' });
      const id = payload.id;
      const b = payload.booking || {};
      const data = sheet.getDataRange().getValues();
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][0]) === String(id)) {
          const rowNum = r + 1;
          if (b.date !== undefined) sheet.getRange(rowNum, 2).setValue(b.date);
          if (b.startTime !== undefined) sheet.getRange(rowNum, 3).setValue(b.startTime);
          if (b.endTime !== undefined) sheet.getRange(rowNum, 4).setValue(b.endTime);
          if (b.durationMinutes !== undefined) sheet.getRange(rowNum, 5).setValue(b.durationMinutes);
          if (b.bookerName !== undefined) sheet.getRange(rowNum, 6).setValue(b.bookerName);
          if (b.email !== undefined) sheet.getRange(rowNum, 7).setValue(b.email);
          if (b.bleep !== undefined) sheet.getRange(rowNum, 8).setValue(b.bleep);
          if (b.department !== undefined) sheet.getRange(rowNum, 9).setValue(b.department);
          if (b.title !== undefined) sheet.getRange(rowNum, 10).setValue(b.title);
          if (b.attendees !== undefined) sheet.getRange(rowNum, 11).setValue(b.attendees);
          if (b.equipment !== undefined) sheet.getRange(rowNum, 12).setValue(JSON.stringify(b.equipment));
          if (b.pin !== undefined) sheet.getRange(rowNum, 13).setValue(b.pin);
          if (b.roomId !== undefined) sheet.getRange(rowNum, 14).setValue(b.roomId);
          return jsonResponse({ status: 'success', id: id });
        }
      }
      return jsonResponse({ status: 'error', message: 'Booking ID not found: ' + id });
    }

    // 3. Delete Booking
    if (action === 'deleteBooking') {
      const sheet = ss.getSheetByName('Bookings');
      if (!sheet) return jsonResponse({ status: 'error', message: 'Bookings sheet not found' });
      const id = payload.id;
      const data = sheet.getDataRange().getValues();
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][0]) === String(id)) {
          sheet.deleteRow(r + 1);
          return jsonResponse({ status: 'success', deletedId: id });
        }
      }
      return jsonResponse({ status: 'success', note: 'Row already removed' });
    }

    // 4. Add or Update User
    if (action === 'saveUser') {
      const sheet = ss.getSheetByName('Users') || ss.insertSheet('Users');
      const u = payload.user || {};
      const data = sheet.getDataRange().getValues();
      let foundRow = -1;
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][0]) === String(u.id) || (u.email && String(data[r][2]).toLowerCase() === String(u.email).toLowerCase())) {
          foundRow = r + 1;
          break;
        }
      }
      if (foundRow > 0) {
        sheet.getRange(foundRow, 1, 1, 9).setValues([[
          u.id || data[foundRow - 1][0],
          u.name || data[foundRow - 1][1],
          u.email || data[foundRow - 1][2],
          u.bleep || '',
          u.department || 'Pharmacy',
          u.site || 'QEH Heritage',
          u.role || 'staff',
          u.pin || '1234',
          u.isActive !== false
        ]]);
        return jsonResponse({ status: 'success', action: 'updated', id: u.id });
      } else {
        const newId = u.id || ('usr-' + new Date().getTime());
        sheet.appendRow([
          newId,
          u.name || '',
          u.email || '',
          u.bleep || '',
          u.department || 'Pharmacy',
          u.site || 'QEH Heritage',
          u.role || 'staff',
          u.pin || '1234',
          u.isActive !== false
        ]);
        return jsonResponse({ status: 'success', action: 'created', id: newId });
      }
    }

    // 5. Delete User
    if (action === 'deleteUser') {
      const sheet = ss.getSheetByName('Users');
      if (!sheet) return jsonResponse({ status: 'error', message: 'Users sheet not found' });
      const id = payload.id;
      const data = sheet.getDataRange().getValues();
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][0]) === String(id)) {
          sheet.deleteRow(r + 1);
          return jsonResponse({ status: 'success', deletedId: id });
        }
      }
      return jsonResponse({ status: 'success', note: 'User row already removed' });
    }

        // 6. Send or Resend Calendar Invite Email
    if (action === 'sendInviteEmail') {
      const b = payload.booking || {};
      const res = sendCalendarInviteEmail(b);
      return jsonResponse({ status: 'success', emailResult: res });
    }

    return jsonResponse({ status: 'error', message: 'Unknown POST action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/**
 * Helper: Converts Sheet Rows to JSON Object Array based on Headers
 */
function getSheetDataAsObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const rows = [];

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (row.every(cell => cell === '' || cell === null)) continue;
    let obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        if (h === 'starttime' || h === 'endtime') {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
        } else if (h === 'createdat') {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
        } else {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
      }
      if (h === 'equipment' && typeof val === 'string' && val.startsWith('[')) {
        try { val = JSON.parse(val); } catch (e) {}
      }
      if (h === 'id') obj.id = String(val);
      else if (h === 'date') obj.date = String(val);
      else if (h === 'starttime') {
        let st = String(val);
        if (!st || st === '1899-12-30' || !st.includes(':')) {
          const m = String(obj.id || '').match(/-(\d{2})(\d{2})$/);
          if (m) st = m[1] + ':' + m[2];
        }
        obj.startTime = st;
      }
      else if (h === 'endtime') {
        let et = String(val);
        if (!et || et === '1899-12-30' || !et.includes(':')) {
          if (obj.startTime && obj.startTime.includes(':')) {
            const dur = Number(obj.durationMinutes) || 30;
            const parts = obj.startTime.split(':');
            const totalMins = parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10) + dur;
            et = String(Math.floor(totalMins / 60)).padStart(2, '0') + ':' + String(totalMins % 60).padStart(2, '0');
          }
        }
        obj.endTime = et;
      }
      else if (h === 'durationminutes') obj.durationMinutes = Number(val) || 30;
      else if (h === 'bookername') obj.bookerName = String(val);
      else if (h === 'email') obj.email = String(val);
      else if (h === 'bleep') obj.bleep = String(val);
      else if (h === 'department') obj.department = String(val);
      else if (h === 'title') obj.title = String(val);
      else if (h === 'attendees') obj.attendees = Number(val) || 1;
      else if (h === 'equipment') obj.equipment = Array.isArray(val) ? val : [];
      else if (h === 'pin') obj.pin = String(val);
      else if (h === 'roomid') obj.roomId = String(val);
      else if (h === 'isblocked') obj.isBlocked = Boolean(val);
      else if (h === 'isrecurring') obj.isRecurring = Boolean(val);
      else if (h === 'recurrencegroupid') obj.recurrenceGroupId = String(val);
      else if (h === 'createdat') obj.createdAt = String(val);
      else if (h === 'name') obj.name = String(val);
      else if (h === 'site') obj.site = String(val);
      else if (h === 'role') obj.role = String(val);
      else if (h === 'isactive') obj.isActive = Boolean(val);
      else if (h === 'location') obj.location = String(val);
      else if (h === 'capacity') obj.capacity = Number(val) || 16;
      else obj[h] = val;
    });
    rows.push(obj);
  }
  return rows;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ============================================================================
 * Calendar Invitation Email Delivery via MailApp (RFC 5545 iCalendar format)
 * ============================================================================
 */
function sendCalendarInviteEmail(b) {
  if (!b || !b.email || typeof b.email !== 'string' || !b.email.includes('@')) {
    return { sent: false, reason: 'Invalid or missing email' };
  }

  // 1. Build RFC 5545 iCalendar payload
  const cleanDate = String(b.date || '').replace(/[^0-9]/g, '');
  const startParts = String(b.startTime || '09:00').split(':');
  const endParts = String(b.endTime || '09:30').split(':');
  const pad = function(n) { return String(n).padStart(2, '0'); };

  const dtStart = cleanDate + 'T' + pad(startParts[0]) + pad(startParts[1]) + '00';
  const dtEnd = cleanDate + 'T' + pad(endParts[0]) + pad(endParts[1]) + '00';
  const dtStamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const uid = (b.id || ('QE-SR-' + new Date().getTime())) + '@uhb.nhs.uk';
  const roomName = b.roomName || 'Pharmacy Seminar Room';
  const location = roomName + ', Level 1 Pharmacy Department, Queen Elizabeth Hospital Birmingham (UHB), Mindelsohn Way, B15 2GW';
  const bookerName = b.bookerName || 'Staff Member';
  const title = b.title || 'Pharmacy Seminar Room Booking';

  const icsLines = [
    'BEGIN:VCALENDAR',
    'PRODID:-//University Hospitals Birmingham NHS Foundation Trust//Seminar Room Booking//EN',
    'VERSION:2.0',
    'CALSCALE:GREGORIAN',
    'METHOD:REQUEST',
    'BEGIN:VEVENT',
    'UID:' + uid,
    'DTSTAMP:' + dtStamp,
    'DTSTART:' + dtStart,
    'DTEND:' + dtEnd,
    'SUMMARY:' + title,
    'DESCRIPTION:NHS Pharmacy Seminar Room Booking\\n\\nSession: ' + title + '\\nBooker: ' + bookerName + '\\nDepartment: ' + (b.department || 'Pharmacy') + '\\nBleep: ' + (b.bleep || 'N/A') + '\\nLocation: ' + location + '\\nBooking Ref: ' + b.id,
    'LOCATION:' + location,
    'ORGANIZER;CN="QE Pharmacy Seminar Room":mailto:pharmacy.admin@uhb.nhs.uk',
    'ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;CN="' + bookerName + '":mailto:' + b.email,
    'STATUS:CONFIRMED',
    'SEQUENCE:0',
    'TRANSP:OPAQUE',
    'END:VEVENT',
    'END:VCALENDAR'
  ];

  const icsString = icsLines.join('\r\n');
  const icsBlob = Utilities.newBlob(icsString, 'text/calendar; charset=UTF-8; method=REQUEST', 'invite.ics');

  // 2. Build NHS-branded HTML email body
  const htmlBody = ''
    + '<div style="font-family: Arial, Helvetica, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #D8DDE0; border-radius: 8px; overflow: hidden; background: #FFFFFF;">'
    + '  <div style="background-color: #005EB8; color: #FFFFFF; padding: 20px 24px;">'
    + '    <div style="font-size: 13px; font-weight: bold; letter-spacing: 0.5px; opacity: 0.9;">NHS • UNIVERSITY HOSPITALS BIRMINGHAM</div>'
    + '    <h1 style="margin: 6px 0 0 0; font-size: 20px; font-weight: bold; color: #FFFFFF;">Pharmacy Seminar Room Booking Confirmation</h1>'
    + '  </div>'
    + '  <div style="padding: 24px; color: #212B32; line-height: 1.6;">'
    + '    <p style="font-size: 15px; margin-top: 0;">Hello <strong>' + escapeHtmlGas(bookerName) + '</strong>,</p>'
    + '    <p>Your room booking has been successfully confirmed. An interactive calendar invitation is attached to this email and will automatically sync with your <strong>NHS Outlook</strong> or mobile calendar.</p>'
    + '    <div style="background: #F0F4F8; border-left: 4px solid #005EB8; padding: 14px 18px; border-radius: 4px; margin: 20px 0;">'
    + '      <table style="width: 100%; border-collapse: collapse; font-size: 14px;">'
    + '        <tr><td style="padding: 6px 0; color: #4C6272; width: 32%;"><strong>Session:</strong></td><td style="padding: 6px 0; color: #005EB8; font-weight: bold;">' + escapeHtmlGas(title) + '</td></tr>'
    + '        <tr><td style="padding: 6px 0; color: #4C6272;"><strong>Date:</strong></td><td style="padding: 6px 0;">📅 ' + b.date + '</td></tr>'
    + '        <tr><td style="padding: 6px 0; color: #4C6272;"><strong>Time:</strong></td><td style="padding: 6px 0;">🕒 ' + b.startTime + ' – ' + b.endTime + ' (' + (b.durationMinutes || 30) + ' mins)</td></tr>'
    + '        <tr><td style="padding: 6px 0; color: #4C6272;"><strong>Location:</strong></td><td style="padding: 6px 0;">📍 ' + escapeHtmlGas(location) + '</td></tr>'
    + '        <tr><td style="padding: 6px 0; color: #4C6272;"><strong>Booker:</strong></td><td style="padding: 6px 0;">👤 ' + escapeHtmlGas(bookerName) + ' (' + escapeHtmlGas(b.department || 'Pharmacy') + ')</td></tr>'
    + '        <tr><td style="padding: 6px 0; color: #4C6272;"><strong>Booking Ref:</strong></td><td style="padding: 6px 0;"><code style="background: #E8EDEE; padding: 2px 6px; border-radius: 3px;">' + b.id + '</code></td></tr>'
    + '      </table>'
    + '    </div>'
    + '    <p style="font-size: 13px; color: #4C6272; border-top: 1px solid #E8EDEE; padding-top: 16px; margin-bottom: 0;">'
    + '      💡 <strong>Calendar Sync:</strong> Open the attached <code>invite.ics</code> file or click <strong>Accept</strong> in Outlook to add this slot to your work calendar.<br>'
    + '      🔒 <strong>Edit or Cancel:</strong> To release this room, visit the booking app and enter the 4-digit PIN you chose when booking.'
    + '    </p>'
    + '  </div>'
    + '</div>';

  try {
    MailApp.sendEmail({
      to: b.email,
      subject: '🏥 Booking Confirmation: ' + title + ' (' + b.date + ' ' + b.startTime + ')',
      htmlBody: htmlBody,
      attachments: [icsBlob]
    });
    return { sent: true };
  } catch (err) {
    Logger.log('MailApp error: ' + err.toString());
    return { sent: false, error: err.toString() };
  }
}

function escapeHtmlGas(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Run this function ONCE from the Apps Script editor (Run > authorizeEmailPermission)
 * to grant Google's OAuth permission for sending automated booking emails.
 */
function authorizeEmailPermission() {
  const remaining = MailApp.getRemainingDailyQuota();
  Logger.log('MailApp authorized! Remaining daily email quota: ' + remaining);
  return remaining;
}