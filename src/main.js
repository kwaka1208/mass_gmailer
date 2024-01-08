const col = {
    insert1: 0,
    insert2: 1,
    insert3: 2,
    mail_address: 3,
  }
  
  function main(){
    const panel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
    const send_list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('送信先');
  
    const doc = DocumentApp.openById(GASLibrary.getDocumentIdByURL(panel.getRange('ドキュメントURL').getValue()));
    const docText = doc.getBody().getText();
    const subject = panel.getRange('件名').getValue(); // Subject
    const options = { cc: panel.getRange('BCC送信先').getValue(), // BCC送信先メールアドレス
                      name: panel.getRange('送信者名').getValue()} // 送信者名
  
    let startRow = send_list.getRange('送信先リスト').getRow();
    let lastRow = send_list.getLastRow()
    const values = send_list.getRange(startRow, 1, lastRow, col.mail_address + 1).getValues();
    Logger.log(values);
    Logger.log(`${startRow}から${lastRow}まで`);
    Logger.log('件名： ' + subject);
  
    for(let i = 0; i < (lastRow - startRow + 1); i++){
      let insert1 = values[i][col.insert1]; //宛先
      let insert2 = values[i][col.insert2]; //宛先
      let insert3 = values[i][col.insert3]; //宛先
      let mailAddress = values[i][col.mail_address]; // メールアドレス
      Logger.log("メールアドレス" + mailAddress)
      if (mailAddress == "")
        break;
      let body = docText
      Logger.log(body);
      body = body.replace(/{埋込1}/g, insert1)
      body = body.replace(/{埋込2}/g, insert2)
      body = body.replace(/{埋込3}/g, insert3)
      GmailApp.sendEmail(mailAddress, subject, body, options);
      Logger.log('送信完了: ' + mailAddress)
    }
    GASLibrary.showMsg('送信完了しました')
  }
  