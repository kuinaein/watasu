'use strict';

var SALT = '7v37BuSWJAkHvwp1ySXSzUwWdbv5GyqM';
var ADMIN_EMAIL = '+xxx@gmail.com';
var STORAGE_FOLDER_ID = 'xxxxxx';

function publish() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('トップ');
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  for (var i = 0; i < data.length; ++i) {
    var toName = data[i][0];
    var filename = data[i][1];
    var to = data[i][2];
    var timeLimit = data[i][4];
    var status = data[i][5];
    if (status === '公開中') {
      if (new Date().getTime() > timeLimit.getTime()) {
        sheet.getRange(2 + i, 6).setValue('公開終了');
      }
    }
    if (status === '公開準備完了') {
      var passcode = generatePassword(1 + i, filename, to, timeLimit);
      sheet.getRange(2 + i, 4, 1, 3).setValues([[passcode, timeLimit, '公開中']]);
      GmailApp.sendEmail(to, filename + '送信のお知らせ', toName + '様\n\
\n\
ファイル名: ' + filename + '\n\
ダウンロード期限: ' + Utilities.formatDate(timeLimit, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') + '\n\
\n\
下記URLよりダウンロードしてください: \n\
' + ScriptApp.getService().getUrl() + '?fileId=' + (2 + i) + '&pwd=' + passcode,
        { from: ADMIN_EMAIL, bcc: ADMIN_EMAIL });
    }
  }
}

function generatePassword(fileId, filename, to, timeLimit) {
  var base = SALT + Math.random() + new Date() + fileId + filename + to + timeLimit;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, base);
  return Utilities.base64EncodeWebSafe(digest);
}

function doGet(req) {
  var tmpl = HtmlService.createTemplateFromFile('index.html');
  tmpl.fileId = req.parameter['fileId'] || '';
  tmpl.pwd = req.parameter['pwd'] || '';
  return tmpl.evaluate();
}

function doPost(req) {
  var reqArgs = { req: req };
  try {
    var fileId = req.parameter.fileId;
    if (!fileId.match(/^\d+$/)) {
      reportDownloadFailed('ファイルID不正', reqArgs);
      return ContentService.createTextOutput('不正なファイルIDです');
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('トップ');
    var row = sheet.getRange('A' + fileId + ':F' + fileId);
    var vals = row.getValues()[0];
    var passcode = vals[3];
    var fileDesc = {
      filename: filename = vals[1],
      timeLimit: vals[4],
      publishStatus: vals[5]
    }
    reqArgs.fileDesc = fileDesc;

    if (passcode !== req.parameter.pwd) {
      reportDownloadFailed('パスワード不一致', reqArgs);
      return ContentService.createTextOutput('パスワードが違います');
    }
    if (new Date().getTime() > fileDesc.timeLimit.getTime() ||
        '公開中' !== fileDesc.publishStatus) {
      reportDownloadFailed('公開期間終了', reqArgs);
      return ContentService.createTextOutput('公開期間を過ぎています');
    }

    var folder = DriveApp.getFolderById(STORAGE_FOLDER_ID);
    var iter = folder.getFilesByName(filename);
    var f;
    do {
      if (!iter.hasNext()) {
        reportDownloadFailed('ファイルが存在しない', reqArgs);
        return ContentService.createTextOutput('ファイルが見つかりません');
      }
      f = iter.next();
    } while(f.isTrashed());
    if (iter.hasNext()) {
      while (iter.hasNext()) {
        var g = iter.next();
        if (!g.isTrashed()) {
          reportDownloadFailed('同名のファイルが複数存在', reqArgs);
          return ContentService.createTextOutput('内部エラー');
        }
      }
    }

    var tmpl = HtmlService.createTemplateFromFile('download.html');
    tmpl.browser = req.parameter.browser;
    tmpl.filename = filename;
    tmpl.mime = f.getMimeType();
    tmpl.data = Utilities.base64Encode(f.getBlob().getBytes());
    reportDownloadsucceeded(reqArgs);
    return tmpl.evaluate();
  } catch (err) {
    reqArgs.err = err;
    reportDownloadFailed('不明なエラー', reqArgs);
    return ContentService.createTextOutput('不明なエラー');
  }
}

function reportDownloadsucceeded(reqArgs) {
  _report('【ワタス君】ダウンロード対象データ送信', 'ダウンロード対象データを送信', reqArgs);
}

function reportDownloadFailed(msg, reqArgs) {
  _report('【ワタス君】ダウンロード失敗レポート', msg, reqArgs);
}

function _report(title, msg, reqArgs) {
  Logger.log(title);
  Logger.log(reqArgs);
  var body = '【メッセージ】\n' + msg;
  if (reqArgs.fileDesc) {
    body += '\n\n【ファイルの状況】\n' + JSON.stringify(reqArgs.fileDesc, null, 2);
  }
  body += '\n\n【リクエスト】\n' + JSON.stringify(reqArgs.req, null, 2);
  if (reqArgs.err) {
    body += '\n\n【例外】\n' + JSON.stringify(reqArgs.err, null, 2);
  }
  GmailApp.sendEmail(ADMIN_EMAIL, title, body, {
    from: ADMIN_EMAIL
  });
}
