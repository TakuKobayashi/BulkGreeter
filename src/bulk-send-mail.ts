const contactSheerUrl = 'https://docs.google.com/spreadsheets/d/15u-Un6DA-2BT7gqdt7Aox-893B5PhciRMQBAn6ivWjs/edit';

function bulkSendMail() {
  const contactDataObjs = loadContactsFromSheet();
  for (const contactData of contactDataObjs) {
    sendMail(contactData);
  }
}

function loadContactsFromSheet(): { [key: string]: any }[] {
  const targetSpreadSheet = SpreadsheetApp.openByUrl(contactSheerUrl);
  const targetSheets = targetSpreadSheet.getSheets();
  const targetSheet = targetSheets[0];
  const contactDataObjs: { [key: string]: any }[] = [];
  if (targetSheet) {
    const lastColumnNumber = targetSheet.getLastColumn();
    const lastRowNumber = targetSheet.getLastRow();
    const headerDataRange = targetSheet.getRange(1, 1, 1, lastColumnNumber);
    const headerData = headerDataRange.getValues();
    if (headerData.length > 0) {
      // headerを除いた2行目以降のCellデータ全部
      const contactDataRange = targetSheet.getRange(2, 1, lastRowNumber - 1, lastColumnNumber);
      const contactData = contactDataRange.getValues();
      for (const rowData of contactData) {
        const rowContactDataObj: { [key: string]: any } = {};
        headerData[0].forEach((header, columnIndex) => {
          rowContactDataObj[header.toString()] = rowData[columnIndex];
        });
        contactDataObjs.push(rowContactDataObj);
      }
    }
  }
  return contactDataObjs;
}

function sendMail(contactObj: { [key: string]: any }) {
  const subject = '~~件名~~';
  const options = {};

  const headerGreetLine = [contactObj['会社名'], contactObj['名前'], '様'].join(' ');
  const body = [headerGreetLine, 'メール本文'].join('\n');

  GmailApp.sendEmail(contactObj['電子メール'], subject, body, options);
}
