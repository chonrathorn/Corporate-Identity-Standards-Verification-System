const SHEET_NAME = "Sheet1";
const PAGE_SIZE = 10;
const COLS = 15;


/* ================= ADMIN LOGIN ================= */
const ADMIN_EMAIL = "mwa.cisystem@gmail.com";
const ADMIN_PASSWORD = "1268Psk";

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : "index";
  const t = HtmlService.createTemplateFromFile(page);

  // ✅ URL ของ deployment ตัวจริง (exec)
  t.WEBAPP_URL = ScriptApp.getService().getUrl();

  return t.evaluate();
}

// ================= SAVE DATA =================
function saveData(form){
  try{
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    const now = new Date();

    // 🔥 insert row หลัง header (แถว 17)
    sheet.insertRowBefore(17);

    sheet.getRange(17,1,1,12).setValues([[
      "",
      form.firstName,
      form.lastName,
      now,
      form.department,
      form.phone,
      form.email,
      form.mediaType,
      form.workUrl,
      form.description,
      form.round,
      "รอพิจารณา"
    ]]);

    return {status:"success"};

  }catch(e){
    Logger.log(e);
    return {status:"error", message:e.message};
  }
}

// ================= LOAD ADMIN =================
function loadAdmin(page){
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

  const lastRow = sheet.getLastRow();
  const startRow = Math.max(17, lastRow - PAGE_SIZE + 1); // เอาแค่ท้ายสุด

  const values = sheet.getRange(startRow, 2, PAGE_SIZE, 10).getValues().reverse();

  return {
    values,
    fromRow: startRow
  };
}

// ================= LOGIN ADMIN =================
function checkAdminAccess(){
  const email = Session.getActiveUser().getEmail();

  return {
    access: email === ADMIN_EMAIL,
    email: email
  };
}

function checkEmailManual(inputEmail, inputPassword){
  const okEmail = String(inputEmail || "").trim().toLowerCase() === ADMIN_EMAIL.toLowerCase();
  const okPassword = String(inputPassword || "") === ADMIN_PASSWORD;

  return {
    access: okEmail && okPassword,
    email: inputEmail
  };
}

// ================= UPDATE STATUS BY LINK=================
function updateStatusByLink(workUrl, status, comment, by){
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for(let i=0;i<data.length;i++){
    if(data[i][8] === workUrl){ // col workUrl = index 8
      sheet.getRange(i+1, 12, 1, 4).setValues([[status, new Date(), comment, by]]);
      return {status:"success"};
    }
  }

  return {status:"not found"};
}

// ================= UPDATE STATUS =================
function sendEmailUpdate(email, name, status, comment, link){
  if(!email) return {status:"no email"};

  MailApp.sendEmail({
    to: email,
    subject: "📌[ผล] ระบบตรวจเช็กความถูกต้องของเอกลักษณ์องค์กร Corporate Identity Standards Verification System : ผลการตรวจสอบชิ้นงาน",
    body: `สวัสดีค่ะ คุณ ${name}

เนื่องจากทาง ฝสอ. ได้มีการตรวจสอบชิ้นงาน/สื่อที่ท่านได้ทำการส่งมาเรียบร้อยแล้ว และมีความเห็นตามรายละเอียดแนบด้านล่างค่ะ

สถานะ: ${status}
ความเห็น: ${comment}

🔗ลิงก์งาน:
${link}

หากท่านมีข้อสงสัยเพิ่มเติม สามารถติดต่อสอบถามได้ที่ ฝสอ. 📞โทร 1268
ขอบคุณค่ะ🙏`
  });

  return {status:"sent"};
}
