function doGet(e) {
  try {
    const data = getTaskDataForAnalysis();
    // データをJSON形式に変換して返す
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * スプレッドシートから技のデータを取得し、画像か動画かを判定して返す
 */
function getTaskDataForAnalysis() {
  const ssId = '1Ug2fLzPFMg37Cz8Pu5sWe4ww8tr7LAMn5XO1dq8jjXk';
  const sheetName = 'rubric_tasks';
  
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
  if (!sheet) throw new Error("シートが見つかりません");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const taskNameIdx = headers.indexOf('task_name');
  const imageUrlIdx = headers.indexOf('image_url');

  if (taskNameIdx === -1 || imageUrlIdx === -1) {
    throw new Error("必要な列が見つかりません");
  }

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const taskName = data[i][taskNameIdx];
    const url = data[i][imageUrlIdx];

    if (taskName && url) {
      let type = 'video';
      
      const driveMatch = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (driveMatch && driveMatch[1]) {
        try {
          const file = DriveApp.getFileById(driveMatch[1]);
          if (file.getMimeType().indexOf('image/') !== -1) {
            type = 'image';
          }
        } catch (e) {
          console.warn("MIME type fetch error for ID: " + driveMatch[1]);
        }
      } else if (url.match(/\.(jpeg|jpg|gif|png|webp)$/i)) {
        type = 'image';
      }

      result.push({ name: taskName, url: url, type: type });
    }
  }
  return result;
}
