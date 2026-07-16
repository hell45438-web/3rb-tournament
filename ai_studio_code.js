// دالة لجلب ورقة الإعدادات أو إنشائها تلقائياً إذا لم تكن موجودة
function getConfigSheet(ss) {
  let configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    configSheet = ss.insertSheet("Config");
    configSheet.appendRow(["SettingName", "Value"]);
    configSheet.appendRow(["background_url", "https://images.unsplash.com/photo-1542751371-adc38448a05e?q=80&w=2070"]);
  }
  return configSheet;
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const players = [];
  
  // جلب رابط الخلفية الحية من ورقة الإعدادات
  const configSheet = getConfigSheet(ss);
  const configData = configSheet.getDataRange().getValues();
  let backgroundUrl = "https://images.unsplash.com/photo-1542751371-adc38448a05e?q=80&w=2070"; // القيمة الافتراضية
  
  for (let i = 1; i < configData.length; i++) {
    if (String(configData[i][0]) === "background_url") {
      backgroundUrl = String(configData[i][1]);
      break;
    }
  }
  
  // تخطي الصف الأول (العناوين) للاعبين
  for (let i = 1; i < data.length; i++) {
    players.push({
      id: String(data[i][0]),
      playerId: String(data[i][1]),
      inGameName: String(data[i][2]),
      discordTag: String(data[i][3]),
      avatarUrl: String(data[i][4]),
      score: Number(data[i][5] || 0),
      createdAt: Number(data[i][6])
    });
  }
  
  // ترتيب اللاعبين تنازلياً حسب النقاط
  players.sort((a, b) => b.score - a.score);
  
  // إرجاع قائمة اللاعبين ورابط الخلفية الحالي
  return ContentService.createTextOutput(JSON.stringify({ players: players, backgroundUrl: backgroundUrl }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*');
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const postData = JSON.parse(e.postData.contents);
  const action = postData.action;
  
  // إجراء حفظ وتحديث رابط الخلفية سحابياً
  if (action === 'updateBackground') {
    const configSheet = getConfigSheet(ss);
    const newBgUrl = String(postData.backgroundUrl).trim();
    
    // تحديث القيمة في الشيت
    configSheet.getRange(2, 2).setValue(newBgUrl);
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');
  }
  
  if (action === 'register') {
    const data = sheet.getDataRange().getValues();
    const newPlayerId = String(postData.playerId).trim().toLowerCase();
    const newIgn = String(postData.inGameName).trim().toLowerCase();
    
    // التحقق من التكرار
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === newPlayerId || 
          String(data[i][2]).trim().toLowerCase() === newIgn) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'duplicate' }))
          .setMimeType(ContentService.MimeType.JSON)
          .setHeader('Access-Control-Allow-Origin', '*');
      }
    }
    
    // إضافة اللاعب
    const uuid = Utilities.getUuid();
    sheet.appendRow([
      uuid,
      postData.playerId,
      postData.inGameName,
      postData.discordTag,
      postData.avatarUrl, // Base64 image
      0, // النقاط الابتدائية
      Date.now()
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');
  }
  
  if (action === 'updateScores') {
    const scoresMap = postData.scores; // { "player-uuid": score }
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]);
      if (scoresMap[rowId] !== undefined) {
        sheet.getRange(i + 1, 6).setValue(Number(scoresMap[rowId]));
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');
  }
  
  if (action === 'deletePlayer') {
    const targetId = String(postData.id);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === targetId) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');
  }
  
  if (action === 'resetAll') {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*');
}
