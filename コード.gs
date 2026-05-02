const SHEET_ROSTER = '名簿';
const SHEET_ATTENDANCE = '出欠記録';
const SHEET_MEETINGS = '会議マスタ';

// =========================================================
// ▼ 列の設定（0始まり：A列=0, B列=1, C列=2, D列=3...）
// ※画像の構成に合わせて数値を変更してください。
// 例：団体名がC列なら「2」、メールアドレスがE列なら「4」
// =========================================================
const COL_ID     = 0;  // 団体ID
const COL_ORG    = 1;  // 団体名
const COL_REP    = 2;  // 代表者名
const COL_EMAIL  = 3;  // メールアドレス
const COL_TOKEN  = 4;  // トークンを記録する空き列
const COL_ISSUED = 5;  // 発行日時を記録する空き列
// =========================================================

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

// サーバー側で現在の会議を記憶・取得する
function getActiveMeeting() {
  return PropertiesService.getScriptProperties().getProperty('ACTIVE_MEETING') || '';
}

function setActiveMeeting(meetingName) {
  PropertiesService.getScriptProperties().setProperty('ACTIVE_MEETING', meetingName);
  return meetingName;
}

// 参加者用：未発行の団体のみを取得
function getRoster() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  let orgMap = {}; // 団体IDでグループ化するための箱
  
  for (let i = 1; i < data.length; i++) {
    let id = data[i][COL_ID];
    if (!id) continue;
    
    // まだ発行されていない場合
    if (!data[i][COL_ISSUED]) { 
      if (!orgMap[id]) {
        orgMap[id] = { id: id, org: data[i][COL_ORG] };
      }
    }
  }
  return Object.values(orgMap);
}

// スタッフ用：全団体のリストを取得
function getStaffRoster() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  let orgMap = {};
  
  for (let i = 1; i < data.length; i++) {
    let id = data[i][COL_ID];
    if (!id) continue;
    
    if (!orgMap[id]) {
      orgMap[id] = {
        id: id,
        org: data[i][COL_ORG],
        reps: [],
        isIssued: data[i][COL_ISSUED] ? true : false
      };
    }
    // 複数名いる場合は配列に名前を追加していく
    if (data[i][COL_REP]) {
      orgMap[id].reps.push(data[i][COL_REP]);
    }
  }
  
  // 複数名を「鈴木大喜 / 矢口遙香」のように繋げて返す
  return Object.values(orgMap).map(org => ({
    id: org.id,
    org: org.org,
    rep: org.reps.join(' / '),
    isIssued: org.isIssued
  }));
}

// 参加者用：初回発行処理
function issueQRCode(orgId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  const appUrl = ScriptApp.getService().getUrl(); 
  
  let targetRows = [];
  let emails = [];
  let repNames = [];
  let orgName = "";
  let isIssued = false;

  // 対象の団体IDを持つ【すべて】の行を探す
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_ID]) === String(orgId)) {
      if (data[i][COL_ISSUED]) isIssued = true;
      
      targetRows.push(i);
      orgName = data[i][COL_ORG];
      if (data[i][COL_REP]) repNames.push(data[i][COL_REP]);
      
      const email = data[i][COL_EMAIL];
      if (email && !emails.includes(email)) {
        emails.push(email); // 重複しないようにアドレスを収集
      }
    }
  }

  if (targetRows.length === 0) return { success: false, message: '団体が見つかりません。' };
  if (isIssued) return { success: false, message: 'この団体は既に発行されています。\n\n※誤って別の団体を選択してしまった場合は、受付スタッフまたは mepo.jimukyoku@gmail.com までご連絡ください。' };
  if (emails.length === 0) return { success: false, message: 'メールアドレスが登録されていません。受付スタッフまたは mepo.jimukyoku@gmail.com までご連絡ください。' };

  const token = Utilities.getUuid();
  const timestamp = new Date();
  
  // 該当する全員の行に同じトークンを書き込む
  targetRows.forEach(rowIdx => {
    sheet.getRange(rowIdx + 1, COL_TOKEN + 1).setValue(token);
    sheet.getRange(rowIdx + 1, COL_ISSUED + 1).setValue(timestamp);
  });
  
  const scanUrl = appUrl + "?token=" + token;
  const qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=' + encodeURIComponent(scanUrl);
  const repNameStr = repNames.join(' / ');
  
  try {
    const body = `
      <p>${orgName}<br>${repNameStr} 様</p>
      <p>会議の参加用二次元コードが発行されました。<br>
      当日の受付にて、このメールの画面をスタッフにご提示ください。</p>
      <p><img src="${qrApiUrl}" alt="二次元コード" style="border: 1px solid #ccc;"></p>
      <hr>
      <p>※複数名ご登録いただいている場合、代表者皆様にBCCで一斉送信されています。<br>
      ※受付の際は、皆様お揃いの上でこの二次元コードを一度だけご提示ください。</p>
      <p>※上記に画像が表示されない場合は、以下のリンクをタップしてください。<br>
      <a href="${qrApiUrl}">👉 二次元コードを表示する</a></p>
    `;
    
    // ★変更箇所：Toを事務局に、全員をBCCに設定
    let mailOptions = { 
      to: 'mepo.jimukyoku@gmail.com',  // Toには事務局アドレスを指定
      bcc: emails.join(','),           // 収集した全員のアドレスをBCCに指定
      subject: "【重要】会議参加用二次元コードのご案内", 
      htmlBody: body 
    };
    
    MailApp.sendEmail(mailOptions);
    
  } catch (e) {
    targetRows.forEach(rowIdx => {
      sheet.getRange(rowIdx + 1, COL_TOKEN + 1).clearContent();
      sheet.getRange(rowIdx + 1, COL_ISSUED + 1).clearContent();
    });
    return { success: false, message: 'メール送信エラーが発生しました。受付スタッフまたは mepo.jimukyoku@gmail.com までご連絡ください。' };
  }
  return { success: true, qrUrl: qrApiUrl };
}

// スタッフ用：再発行処理
function reissueQRCodeByStaff(orgId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  const appUrl = ScriptApp.getService().getUrl(); 
  
  let targetRows = [];
  let emails = [];
  let repNames = [];
  let orgName = "";

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_ID]) === String(orgId)) {
      targetRows.push(i);
      orgName = data[i][COL_ORG];
      if (data[i][COL_REP]) repNames.push(data[i][COL_REP]);
      
      const email = data[i][COL_EMAIL];
      if (email && !emails.includes(email)) emails.push(email);
    }
  }

  if (targetRows.length === 0) return { success: false, message: '団体が見つかりません。' };
  if (emails.length === 0) return { success: false, message: 'メールアドレスが登録されていません。' };

  const token = Utilities.getUuid();
  const timestamp = new Date();
  
  targetRows.forEach(rowIdx => {
    sheet.getRange(rowIdx + 1, COL_TOKEN + 1).setValue(token);
    sheet.getRange(rowIdx + 1, COL_ISSUED + 1).setValue(timestamp);
  });
  
  const scanUrl = appUrl + "?token=" + token;
  const qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=' + encodeURIComponent(scanUrl);
  const repNameStr = repNames.join(' / ');
  
  try {
    const body = `
      <p>${orgName}<br>${repNameStr} 様</p>
      <p style="color:red;"><b>受付スタッフにより、参加用二次元コードが【再発行】されました。</b><br>
      以前のコードは無効になりますのでご注意ください。</p>
      <p>当日の受付にて、このメールの画面をスタッフにご提示ください。</p>
      <p><img src="${qrApiUrl}" alt="二次元コード" style="border: 1px solid #ccc;"></p>
      <hr>
      <p><a href="${qrApiUrl}">👉 二次元コードを表示する</a></p>
    `;
    
    // ★変更箇所：Toを事務局に、全員をBCCに設定
    let mailOptions = { 
      to: 'mepo.jimukyoku@gmail.com',  // Toには事務局アドレスを指定
      bcc: emails.join(','),           // 収集した全員のアドレスをBCCに指定
      subject: "【再発行】会議参加用二次元コードのご案内", 
      htmlBody: body 
    };
    
    MailApp.sendEmail(mailOptions);
    
  } catch (e) {
    return { success: false, message: 'メール送信エラーが発生しました。' };
  }
  return { success: true, qrUrl: qrApiUrl };
}

// スタッフ用：QRスキャン時の出欠処理
function recordAttendance(scannedToken) {
  const meetingName = getActiveMeeting();
  if (!meetingName) {
    return { success: false, message: '現在、受付中の会議がありません。\nスタッフ画面から受付を開始してください。' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(SHEET_ROSTER);
  const attendanceSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  const proxySheet = ss.getSheetByName('代理出席'); // ★追加：代理出席シートを読み込む
  const rosterData = rosterSheet.getDataRange().getValues();
  
  let targetOrgId = null;
  let targetOrgName = null;
  let repNames = [];

  // スキャンしたトークンと一致する団体IDと全員の名前を収集
  for (let i = 1; i < rosterData.length; i++) {
    if (String(rosterData[i][COL_TOKEN]) === String(scannedToken) && scannedToken !== "") {
      if (!targetOrgName) {
        targetOrgId = rosterData[i][COL_ID];   // ★追加：団体IDを取得
        targetOrgName = rosterData[i][COL_ORG];
      }
      if (rosterData[i][COL_REP]) {
        repNames.push(rosterData[i][COL_REP]);
      }
    }
  }

  if (!targetOrgName) return { success: false, message: '無効な二次元コードです。\n(再発行された古いコードである可能性があります)' };

  let repNameStr = repNames.join(' / ');
  let displayRepStr = repNameStr; // 画面表示用の文字列
  let recordRepStr = repNameStr;  // スプレッドシート記録用の文字列

  // =========================================================
  // ★追加：代理出席のリストと照合する処理
  // =========================================================
  if (proxySheet) {
    const proxyData = proxySheet.getDataRange().getValues();
    for (let p = 1; p < proxyData.length; p++) {
      // A列(0):会議名, B列(1):団体ID, C列(2):代理出席者名 が一致するかチェック
      if (proxyData[p][0] === meetingName && String(proxyData[p][1]) === String(targetOrgId)) {
        const proxyName = proxyData[p][2];
        
        // 画面には赤文字で代理出席者を強調表示する
        displayRepStr = `${repNameStr}<br><span style="color:#d93025; font-size:16px;">（本日は代理出席：${proxyName}）</span>`;
        // スプレッドシートにはカッコ書きで記録する
        recordRepStr = `${repNameStr}（代理：${proxyName}）`;
        break;
      }
    }
  }
  // =========================================================
  
  // 出席記録には1行で打刻
  attendanceSheet.appendRow([new Date(), meetingName, targetOrgName, recordRepStr, '出席済']);
  
  return { success: true, org: targetOrgName, rep: displayRepStr, meeting: meetingName };
}
