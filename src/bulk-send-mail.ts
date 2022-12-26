const contactSheerUrl = 'https://docs.google.com/spreadsheets/d/15u-Un6DA-2BT7gqdt7Aox-893B5PhciRMQBAn6ivWjs/edit';

// ファイルの読み込みはあらかじめ行なっておく
const attachmentFileBlob = DriveApp.getFileById('1ObS1pgmhIcqnIxvqjdVvsa4MKNp0D2OF').getBlob();

function bulkSendMail() {
  const contactDataObjs = loadContactsFromSheet();
  for (const contactData of contactDataObjs) {
    sendMail(contactData);
  }
}

function loadContactsFromSheet(): { [key: string]: any }[] {
  const targetSpreadSheet = SpreadsheetApp.openByUrl(contactSheerUrl);
  const targetSheets = targetSpreadSheet.getSheets();
  // シートは一つしかないので、一つのシートを指定する
  const targetSheet = targetSheets[0];
  const contactDataObjs: { [key: string]: any }[] = [];
  if (targetSheet) {
    const lastColumnNumber = targetSheet.getLastColumn();
    const lastRowNumber = targetSheet.getLastRow();
    // 1行目に記載されているheader情報だけをまずは取得する
    const headerDataRange = targetSheet.getRange(1, 1, 1, lastColumnNumber);
    const headerData = headerDataRange.getValues();
    if (headerData.length > 0) {
      // headerを除いた2行目以降のCellデータ全部を取得する
      const contactDataRange = targetSheet.getRange(2, 1, lastRowNumber - 1, lastColumnNumber);
      const contactData = contactDataRange.getValues();
      for (const rowData of contactData) {
        // 1行分のデータ
        const rowContactDataObj: { [key: string]: any } = {};
        // header情報をkey、そのkeyに対する値を保持する
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
  const subject = '謹賀新年'; // 件名
  const headerGreetLine = [contactObj['会社名'], contactObj['名前'], '様'].join(' ');
  const body = [headerGreetLine, 'あけましておめでとうございます', '昨年は大変お世話になりました', '今年もよろしくお願いします'].join('\n'); // メールの本文
  // const options = {}; // ファイル添付などを行わない場合

  // ファイル添付を行う
  const options = { attachments: attachmentFileBlob }; // ファイルを添付する

  // HTMLメールで本文に画像を挿入する場合
  // const body = [headerGreetLine, 'あけましておめでとうございます', '昨年は大変お世話になりました','本年もよろしくお願いいたします', '<img src="cid:inlineImg">'].join('<br>'); // メールの本文
  /*
  const options = {
    htmlBody: body,
    inlineImages: {
      inlineImg: attachmentFileBlob,
    },
  }; // ファイルを添付しつつ本文の中にインライン画像を埋め込む
  */

  GmailApp.sendEmail(contactObj['電子メール'], subject, body, options);
}
