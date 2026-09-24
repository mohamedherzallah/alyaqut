
// إعدادات الورقة والعناوين
const SHEET_NAME = "سجلات الطلاب والمدفوعات";
const HEADERS = ["المعرف", "التاريخ والوقت", "اسم الطالب", "المرحلة الدراسية", "نوع المسار", "المبلغ المدفوع (شيكل)"];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    // تنسيق صف العناوين باللون الأزرق الداكن والخط الذهبي/الأبيض
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground("#003366")
               .setFontColor("#FFFFFF")
               .setFontWeight("bold")
               .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// استقبال البيانات (سواء سجل واحد أو دفعة أوفلاين)
function doPost(e) {
  try {
    const sheet = getSheet();
    const data = JSON.parse(e.postData.contents);
    const records = Array.isArray(data) ? data : [data];
    const newRows = [];

    records.forEach(r => {
      const id = "YQ-" + Math.floor(100000 + Math.random() * 900000);
      const timestamp = Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm");
      newRows.push([
        id,
        timestamp,
        r.studentName,
        r.gradeLevel,
        r.trackType, // تأسيس / منهج عادي
        Number(r.paymentAmount) || 0
      ]);
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, HEADERS.length).setValues(newRows);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "تم حفظ البيانات بنجاح",
      count: newRows.length
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// جلب البيانات للإدارة
function doGet(e) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const students = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { // التأكد من وجود سجل
        students.push({
          id: data[i][0],
          date: data[i][1],
          studentName: data[i][2],
          gradeLevel: data[i][3],
          trackType: data[i][4],
          paymentAmount: data[i][5]
        });
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      data: students
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
