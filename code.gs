// doPost يستقبل البيانات من تطبيق الهاتف ويحفظها في Google Sheets
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  
  // فحص إذا كانت البيانات عبارة عن مصفوفة (للمزامنة المتعددة) أو سجل واحد
  var records = Array.isArray(data) ? data : [data];
  
  records.forEach(function(record) {
    sheet.appendRow([
      new Date(), // Timestamp
      record.studentName,
      record.gradeLevel,
      record.trackType, // تأسيس أو منهج عادي
      record.paymentAmount
    ]);
  });
  
  return ContentService.createTextOutput(JSON.stringify({"status": "success", "message": "تم حفظ البيانات بنجاح"}))
    .setMimeType(ContentService.MimeType.JSON);
}

// doGet لجلب بيانات الطلاب لعرضها للإدارة
function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var students = [];
  
  // تخطي صف العناوين الأول
  for (var i = 1; i < data.length; i++) {
    students.push({
      date: data[i][0],
      name: data[i][1],
      grade: data[i][2],
      track: data[i][3],
      payment: data[i][4]
    });
  }
  
  return ContentService.createTextOutput(JSON.stringify(students))
    .setMimeType(ContentService.MimeType.JSON);
}