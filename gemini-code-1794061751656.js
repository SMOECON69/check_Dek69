function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var targetStudentId = String(data.studentId || "").trim();
    var targetName = String(data.name || "").trim();
    var statusText = data.status || "รับแล้ว";
    var isReceived = (statusText === "รับแล้ว");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timestamp = new Date();
    
    // --- ส่วนที่ 1: ไปติ๊กถูก (TRUE/FALSE) ในชีท "รับของ" ---
    var sheetReceiving = ss.getSheetByName("รับของ");
    if (sheetReceiving) {
      var dataRange = sheetReceiving.getDataRange().getValues();
      for (var i = 1; i < dataRange.length; i++) {
        var rowId = String(dataRange[i][0] || "").trim();     // คอลัมน์ A: รหัสนักศึกษา
        var rowName = String(dataRange[i][1] || "").trim();   // คอลัมน์ B: ชื่อ-สกุล
        
        var isMatch = false;
        if (targetStudentId && targetStudentId !== "-" && rowId === targetStudentId) {
          isMatch = true;
        } else if (targetName && (rowName.indexOf(targetName) !== -1 || targetName.indexOf(rowName) !== -1)) {
          isMatch = true;
        }
        
        if (isMatch) {
          // คอลัมน์ C คือคอลัมน์ที่ 3 (ช่อง Checkbox สถานะ)
          sheetReceiving.getRange(i + 1, 3).setValue(isReceived);
        }
      }
    }
    
    // --- ส่วนที่ 2: บันทึกประวัติการรับลงในชีท "รับผ้าคาด" ---
    var sheetLog = ss.getSheetByName("รับผ้าคาด");
    if (!sheetLog) {
      sheetLog = ss.insertSheet("รับผ้าคาด");
      sheetLog.appendRow(["วันที่-เวลา", "รหัสนักศึกษา", "ชื่อ-นามสกุล", "สถานะ"]);
    }
    
    // บันทึกตามหัวตารางในรูป: A=วันที่, B=รหัสนักศึกษา, C=ชื่อ-นามสกุล, D=สถานะ
    sheetLog.appendRow([timestamp, targetStudentId, targetName, statusText]);
    
    return ContentService.createTextOutput(JSON.stringify({
      result: "success"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      result: "error",
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}