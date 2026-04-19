const SHEET_ROSTER = '名簿';
const SHEET_ATTENDANCE = '出欠記録';
const SHEET_MEETINGS = '会議マスタ';

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.token = e.parameter.token || ''; 
  template.mode = e.parameter.sample || 'user'; // カスタム設定を引き継ぎ
  
  return template.evaluate()
    .setTitle(template.mode === 'value' ? '【スタッフ用】受付システム' : '二次元コード発行')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getMeetings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MEETINGS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => row[0]).filter(val => val !== "");
}

// --- ★追加：サーバー側で現在の会議を記憶・取得する ---
function getActiveMeeting() {
  return PropertiesService.getScriptProperties().getProperty('ACTIVE_MEETING') || '';
}

function setActiveMeeting(meetingName) {
  PropertiesService.getScriptProperties().setProperty('ACTIVE_MEETING', meetingName);
  return meetingName;
}
// ---------------------------------------------------

// 参加者用：未発行の団体のみを取得
function getRoster() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  let availableOrgs = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][5]) { 
      availableOrgs.push({ id: data[i][0], org: data[i][1] });
    }
  }
  return availableOrgs;
}

// スタッフ用：全団体のリストを取得
function getStaffRoster() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  let orgs = [];
  for (let i = 1; i < data.length; i++) {
    orgs.push({
      id: data[i][0],
      org: data[i][1],
      rep: data[i][2],
      isIssued: data[i][5] ? true : false
    });
  }
  return orgs;
}

// 参加者用：初回発行処理
function issueQRCode(orgId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  const appUrl = ScriptApp.getService().getUrl(); 
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(orgId)) {
      if (data[i][5]) return { success: false, message: 'この団体は既に別の端末で発行されています。\n\n※ご自身の団体であるにも関わらず発行できない場合や、誤って別の団体を選択してしまった場合は、受付スタッフまたは mepo.jimukyoku@gmail.com までご連絡ください。' };
      
      const email = data[i][3];
      if (!email) return { success: false, message: 'メールアドレスが登録されていません。受付スタッフまたは mepo.jimukyoku@gmail.com までご連絡ください。' };

      const token = Utilities.getUuid();
      const timestamp = new Date();
      
      sheet.getRange(i + 1, 5).setValue(token);
      sheet.getRange(i + 1, 6).setValue(timestamp);
      
      const scanUrl = appUrl + "?token=" + token;
      const qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=' + encodeURIComponent(scanUrl);
      
      try {
        const body = `
          <p>${data[i][1]} ${data[i][2]} 様</p>
          <p>会議の参加用二次元コードが発行されました。<br>
          当日の受付にて、このメールの画面をスタッフにご提示ください。</p>
          <p><img src="${qrApiUrl}" alt="二次元コード" style="border: 1px solid #ccc;"></p>
          <hr>
          <p>※上記に二次元コードの画像が表示されない場合は、以下のリンクをタップして表示してください。<br>
          <a href="${qrApiUrl}">👉 二次元コードを表示する</a></p>
        `;
        MailApp.sendEmail({ to: email, subject: "【重要】会議参加用二次元コードのご案内", htmlBody: body });
      } catch (e) {
        sheet.getRange(i + 1, 5).clearContent();
        sheet.getRange(i + 1, 6).clearContent();
        return { success: false, message: 'メール送信エラーが発生しました。受付スタッフまたは mepo.jimukyoku@gmail.com までご連絡ください。' };
      }
      return { success: true, qrUrl: qrApiUrl };
    }
  }
  return { success: false, message: '団体が見つかりません。' };
}

// スタッフ用：再発行処理
function reissueQRCodeByStaff(orgId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  const appUrl = ScriptApp.getService().getUrl(); 
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(orgId)) {
      const email = data[i][3];
      if (!email) return { success: false, message: 'メールアドレスが登録されていません。' };

      const token = Utilities.getUuid();
      const timestamp = new Date();
      
      sheet.getRange(i + 1, 5).setValue(token);
      sheet.getRange(i + 1, 6).setValue(timestamp);
      
      const scanUrl = appUrl + "?token=" + token;
      const qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=' + encodeURIComponent(scanUrl);
      
      try {
        const body = `
          <p>${data[i][1]} ${data[i][2]} 様</p>
          <p style="color:red;"><b>受付スタッフにより、参加用二次元コードが【再発行】されました。</b><br>
          以前の二次元コードは無効になりますのでご注意ください。</p>
          <p>当日の受付にて、このメールの画面をスタッフにご提示ください。</p>
          <p><img src="${qrApiUrl}" alt="二次元コード" style="border: 1px solid #ccc;"></p>
          <hr>
          <p><a href="${qrApiUrl}">👉 二次元コードを表示する</a></p>
        `;
        MailApp.sendEmail({ to: email, subject: "【再発行】会議参加用二次元コードのご案内", htmlBody: body });
      } catch (e) {
        return { success: false, message: 'メール送信エラーが発生しました。' };
      }
      return { success: true, qrUrl: qrApiUrl };
    }
  }
  return { success: false, message: '団体が見つかりません。' };
}

// スタッフ用：QRスキャン時の出欠処理
// ★引数から meetingName を消し、サーバーの記憶から読み出すように変更しました
function recordAttendance(scannedToken) {
  const meetingName = getActiveMeeting();
  
  // サーバー側で受付が停止されている場合はエラーを返す
  if (!meetingName) {
    return { success: false, message: '現在、受付中の会議がありません。\nスタッフ画面から受付を開始してください。' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(SHEET_ROSTER);
  const attendanceSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  const rosterData = rosterSheet.getDataRange().getValues();
  
  let targetOrg = null;

  for (let i = 1; i < rosterData.length; i++) {
    if (String(rosterData[i][4]) === String(scannedToken) && scannedToken !== "") {
      targetOrg = { org: rosterData[i][1], rep: rosterData[i][2] };
      break;
    }
  }

  if (!targetOrg) return { success: false, message: '無効な二次元コードです。\n(再発行された古い二次元コードである可能性があります)' };

  attendanceSheet.appendRow([new Date(), meetingName, targetOrg.org, targetOrg.rep, '出席済']);
  return { success: true, org: targetOrg.org, rep: targetOrg.rep, meeting: meetingName };
}
