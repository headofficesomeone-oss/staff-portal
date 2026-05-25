const CONFIG = {
  USERS_SS_ID: "1CiFgLLXQr_2GdMFbytNjQ-FBtdAngO3AL8UwQ0Y9q8c",
  USERS_SHEET: "WEB_USERS",

  MASTER_SS_ID: "1yU7ltwvVy7UgcS87hxrzlC1Nh6Z9IJIziDTRgMlPZik",
  MASTER_SHEET: "WEB_FILE_MASTER",

  FILES_SS_ID: "1JTc_5FdZY9xhhwaZVC7JSx9kElBYkwdSp8PzQeS6P_8",
  FILES_SHEET: "WEB_FILES"
};

function doGet(e) {
  return HtmlService
    .createHtmlOutput("<h2>SAMUWAN API は起動しています</h2><p>ログイン画面からアクセスしてください。</p>")
    .setTitle("SAMUWAN");
}

function doPost(e) {
  try {
    const action = String(e.parameter.action || "").trim();

    if (action !== "login") {
      return showErrorPage_("不明な処理です。");
    }

    return login_(e.parameter);

  } catch (err) {
    return showErrorPage_("エラーが発生しました。<br>" + escapeHtml_(err.message));
  }
}

function login_(data) {
  const staffId = String(data.staffId || "").trim();
  const password = String(data.password || "");

  if (!staffId || !password) {
    return showErrorPage_("職員IDとパスワードを入力してください。");
  }

  const user = findUser_(staffId, password);

  if (!user) {
    return showErrorPage_("職員IDまたはパスワードが違います。");
  }

  const masterFiles = getMasterFiles_(user);
  const individualFiles = getIndividualFiles_(user);
  const files = masterFiles.concat(individualFiles);

  return showPortalPage_(user, files);
}

function findUser_(staffId, password) {
  const ss = SpreadsheetApp.openById(CONFIG.USERS_SS_ID);
  const sheet = ss.getSheetByName(CONFIG.USERS_SHEET);

  if (!sheet) throw new Error("WEB_USERSシートがありません");

  const values = sheet.getDataRange().getValues();
  values.shift();

  for (const row of values) {
    const rowStaffId = String(row[0] || "").trim();
    const name = String(row[1] || "").trim();
    const rowPassword = String(row[2] || "");
    const active = String(row[3] || "").toUpperCase();

    if (
      rowStaffId === staffId &&
      rowPassword === password &&
      active === "TRUE"
    ) {
      return {
        staffId: rowStaffId,
        name: name
      };
    }
  }

  return null;
}

function getMasterFiles_(user) {
  const ss = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
  const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET);

  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  values.shift();

  const result = [];
  const prefix = "pr" + user.name + "_";

  for (const row of values) {
    const fileLabel = String(row[0] || "").trim();
    const folderId = String(row[1] || "").trim();
    const fixedFileId = String(row[2] || "").trim();
    const active = String(row[3] || "").toUpperCase();

    if (active !== "TRUE") continue;

    if (fixedFileId) {
      try {
        const file = DriveApp.getFileById(fixedFileId);
        result.push({
          title: fileLabel || file.getName(),
          url: makeDriveViewUrl_(file.getId())
        });
      } catch (err) {}
      continue;
    }

    if (!folderId) continue;

    try {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFilesByType(MimeType.PDF);

      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();

        if (fileName.startsWith(prefix)) {
          result.push({
            title: fileLabel || fileName,
            url: makeDriveViewUrl_(file.getId())
          });
        }
      }
    } catch (err) {}
  }

  return result;
}

function getIndividualFiles_(user) {
  const ss = SpreadsheetApp.openById(CONFIG.FILES_SS_ID);
  const sheet = ss.getSheetByName(CONFIG.FILES_SHEET);

  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  values.shift();

  const result = [];

  for (const row of values) {
    const staffId = String(row[0] || "").trim();
    const title = String(row[1] || "").trim();
    const fileId = String(row[2] || "").trim();
    const active = String(row[3] || "").toUpperCase();

    if (staffId === user.staffId && fileId && active === "TRUE") {
      result.push({
        title: title || "個別ファイル",
        url: makeDriveViewUrl_(fileId)
      });
    }
  }

  return result;
}

function showPortalPage_(user, files) {
  let fileHtml = "";

  if (!files || files.length === 0) {
    fileHtml = "<div class='empty'>現在、閲覧可能なファイルはありません。</div>";
  } else {
    files.forEach(file => {
      fileHtml += ""
        + "<a class='file' href='" + escapeHtml_(file.url) + "' target='_blank' rel='noopener'>"
        + "<span>🧾 " + escapeHtml_(file.title) + "</span>"
        + "<em>開く</em>"
        + "</a>";
    });
  }

  const html = `
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SAMUWAN</title>
  <style>
    *{box-sizing:border-box}
    body{margin:0;font-family:"Hiragino Sans","Noto Sans JP",sans-serif;color:#4a3b3b;background:linear-gradient(135deg,#fff7fb,#f2fbff);min-height:100vh}
    .page{min-height:100vh;display:grid;place-items:center;padding:28px}
    .card{width:100%;max-width:480px;background:rgba(255,255,255,.9);border:3px solid #fff;border-radius:34px;padding:30px;box-shadow:0 20px 60px rgba(255,128,171,.22)}
    h1{text-align:center;color:#6d4c41;font-size:26px;margin:0 0 10px}
    .message{text-align:center;color:#78666d;line-height:1.8;font-size:14px;margin-bottom:22px}
    .file-list{display:grid;gap:10px}
    .file{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px;background:#fff;border:2px solid #ffe0ec;border-radius:18px;color:#5d4037;text-decoration:none;font-weight:800}
    .file em{font-style:normal;color:#f06292;font-size:13px}
    .empty{padding:16px;background:#fff8fb;border-radius:18px;text-align:center;color:#8d7b82}
    .back{display:block;margin-top:18px;text-align:center;color:#f06292;text-decoration:none;font-weight:bold}
  </style>
</head>
<body>
  <div class="page">
    <div class="card">
      <h1>${escapeHtml_(user.name)} さん</h1>
      <p class="message">閲覧可能なファイル一覧です。</p>
      <div class="file-list">
        ${fileHtml}
      </div>
      <a class="back" href="javascript:history.back()">ログイン画面へ戻る</a>
    </div>
  </div>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html).setTitle("SAMUWAN");
}

function showErrorPage_(message) {
  const html = `
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SAMUWAN エラー</title>
<style>
body{font-family:sans-serif;background:#fff7fb;padding:24px;color:#4a3b3b}
.box{max-width:420px;margin:40px auto;background:#fff;padding:28px;border-radius:24px;box-shadow:0 12px 30px rgba(0,0,0,.1)}
h1{text-align:center;color:#c2185b}
p{line-height:1.8;text-align:center}
a{display:block;text-align:center;margin-top:20px;color:#f06292;font-weight:bold}
</style>
</head>
<body>
<div class="box">
<h1>ログインできません</h1>
<p>${message}</p>
<a href="javascript:history.back()">戻る</a>
</div>
</body>
</html>`;
  return HtmlService.createHtmlOutput(html).setTitle("SAMUWAN エラー");
}

function makeDriveViewUrl_(fileId) {
  return "https://drive.google.com/file/d/" + fileId + "/view";
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
