
function bulkSendMail() {
  sendMail("keep_slimbody@yahoo.co.jp")
  Logger.log("hello world");
}

function sendMail(address: string) {
  const subject = '~~件名~~'; // メールの件名
  const options = { };

  let body = "メール本文";

  GmailApp.sendEmail(address, subject, body, options);

}
