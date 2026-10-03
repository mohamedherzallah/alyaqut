// ====== إعدادات ======
const SHEET_NAME = "سجلات الطلاب والمدفوعات";
const HEADERS = ["المعرف", "التاريخ والوقت", "اسم الطالب", "المرحلة الدراسية", "نوع المسار", "المبلغ المدفوع (شيكل)"];
// كلمة السر: يجب أن تطابق SECRET_TOKEN الموجودة في index.html
const SECRET_TOKEN = "ChangeMe-1234";

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground("#003366")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// يمنع تنفيذ أي معادلة إذا بدأ النص بـ = أو + أو - أو @
function safeText(v) {
  const t = String(v == null ? "" : v);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// استقبال البيانات (سجل واحد أو دفعة أوفلاين)
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.token !== SECRET_TOKEN) throw new Error("غير مصرّح");

    const sheet = getSheet();
    const records = Array.isArray(payload.records) ? payload.records : [];
    const newRows = records.map(r => [
      "YQ-" + Math.floor(100000 + Math.random() * 900000),
      Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm"),
      safeText(r.studentName),
      safeText(r.gradeLevel),
      safeText(r.trackType),
      Number(r.paymentAmount) || 0
    ]);

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, HEADERS.length).setValues(newRows);
    }
    return reply({ status: "success", message: "تم حفظ البيانات بنجاح", count: newRows.length });
  } catch (err) {
    return reply({ status: "error", message: err.toString() });
  }
}

// جلب البيانات للإدارة
function doGet(e) {
  try {
    if (!e || !e.parameter || e.parameter.token !== SECRET_TOKEN) throw new Error("غير مصرّح");

    const data = getSheet().getDataRange().getValues();
    const students = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        students.push({
          id: data[i][0],
          date: String(data[i][1]),
          studentName: data[i][2],
          gradeLevel: data[i][3],
          trackType: data[i][4],
          paymentAmount: data[i][5]
        });
      }
    }
    return reply({ status: "success", data: students });
  } catch (err) {
    return reply({ status: "error", message: err.toString() });
  }
}
