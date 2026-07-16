function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const players = [];
  
  // Skip row 1 (headers)
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
  
  // Sort players by score descending
  players.sort((a, b) => b.score - a.score);
  
  return ContentService.createTextOutput(JSON.stringify({ players: players }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*');
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const postData = JSON.parse(e.postData.contents);
  const action = postData.action;
  
  if (action === 'register') {
    const data = sheet.getDataRange().getValues();
    const newPlayerId = String(postData.playerId).trim().toLowerCase();
    const newIgn = String(postData.inGameName).trim().toLowerCase();
    
    // Check duplicates
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === newPlayerId || 
          String(data[i][2]).trim().toLowerCase() === newIgn) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'duplicate' }))
          .setMimeType(ContentService.MimeType.JSON)
          .setHeader('Access-Control-Allow-Origin', '*');
      }
    }
    
    // Append player
    const uuid = Utilities.getUuid();
    sheet.appendRow([
      uuid,
      postData.playerId,
      postData.inGameName,
      postData.discordTag,
      postData.avatarUrl, // Base64 image
      0, // initial score
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