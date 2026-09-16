// ==========================================
// Configuration
// ==========================================
const SPREADSHEET_ID = '1jO_iOJnj2xW3TSP71xi3b4M0qSzy1dEsNwI2LjaeQj8';
const ADMIN_PASSWORD = '1205101';

function doGet(e) {
  const html = HtmlService.createTemplateFromFile('index.html').evaluate();
  html.setTitle('MSU Projector Booking');
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

// ==========================================
// 1. ตรวจสอบรหัสนิสิต
// ==========================================
function verifyStudent(studentId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('ฐานข้อมูล');
    if (!sheet) return { success: false, message: 'ไม่พบแผ่นงาน "ฐานข้อมูล"' };
    const data = sheet.getDataRange().getValues();
    const searchId = String(studentId).trim();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === searchId) {
        return {
          success: true,
          data: {
            studentId: String(data[i][0]).trim(),
            name: String(data[i][1]).trim(),
            faculty: "วิทยาการสารสนเทศ"
          }
        };
      }
    }
    return { success: false, message: 'ไม่พบรหัสนิสิตในฐานข้อมูล' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.message };
  }
}

// ==========================================
// 2. ดึงข้อมูลการจองทั้งหมด (แปลงวันที่เป็น YYYY-MM-DD ชัดเจน)
// ==========================================
function getBookings() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('การจอง');
    if (!sheet) {
      sheet = setupBookingSheet(ss);
      return [];
    }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const bookings = [];
    for (let i = 1; i < data.length; i++) {
      let rawDate = data[i][5];
      let formattedDate = "";
      if (rawDate instanceof Date) {
        formattedDate = Utilities.formatDate(rawDate, "Asia/Bangkok", "yyyy-MM-dd");
      } else {
        formattedDate = String(rawDate).trim();
      }
      bookings.push({
        id: String(data[i][0]),
        studentId: String(data[i][1]),
        studentName: String(data[i][2]),
        groupName: String(data[i][3]),
        repName: String(data[i][4]),
        date: formattedDate,
        slot: String(data[i][6]),
        machine: Number(data[i][7]),
        status: String(data[i][8]).toUpperCase(),
        createdAt: String(data[i][9])
        // หมายเหตุ: ไม่ส่ง CancelCode (คอลัมน์ K) กลับไปที่ client โดยเจตนา เพื่อความปลอดภัย
      });
    }
    return bookings;
  } catch (error) {
    console.error("getBookings Error:", error);
    return [];
  }
}

// ==========================================
// 3. บันทึกการจอง + ป้องกันคิวซ้อน + ส่งอีเมล (+ สร้างรหัสยกเลิก 4 หลัก)
// ==========================================
function saveBooking(bookingData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // ล็อคคิว 10 วินาที
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('การจอง');
    if (!sheet) sheet = setupBookingSheet(ss);
    const data = sheet.getDataRange().getValues();
    const targetDate = String(bookingData.date).trim();
    const targetSlot = String(bookingData.slot).trim();

    // กรองหาเครื่องที่โดนจองแล้วในรอบนี้
    const bookedMachines = [];
    for (let i = 1; i < data.length; i++) {
      let rDate = data[i][5] instanceof Date ?
        Utilities.formatDate(data[i][5], "Asia/Bangkok", "yyyy-MM-dd") :
        String(data[i][5]).trim();
      let rSlot = String(data[i][6]).trim();
      let rStatus = String(data[i][8]).toUpperCase();
      if (rDate === targetDate && rSlot === targetSlot && rStatus !== 'CANCELLED') {
        bookedMachines.push(Number(data[i][7]));
      }
    }

    const machines = [1, 2, 3, 4, 5];
    const availableMachines = machines.filter(m => !bookedMachines.includes(m));
    if (availableMachines.length === 0) {
      return { success: false, message: 'ช่วงเวลานี้เต็มแล้ว กรุณาเลือกรอบอื่น' };
    }

    const assignedMachine = Math.min(...availableMachines);
    const bookingId = 'BK-' + Utilities.formatDate(new Date(), "Asia/Bangkok", "yyMMddHHmmss");
    const timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd'T'HH:mm:ss");
    const cancelCode = generateCancelCode();

    // บันทึกลง Sheet (คอลัมน์ K = รหัสยกเลิก)
    sheet.appendRow([
      bookingId,
      String(bookingData.studentId).trim(),
      bookingData.studentName,
      bookingData.groupName,
      bookingData.repName,
      targetDate,
      targetSlot,
      assignedMachine,
      'ACTIVE',
      timestamp,
      cancelCode
    ]);

    // ส่งอีเมลยืนยันการจอง (แนบรหัสยกเลิกไปด้วย)
    const emailSent = sendConfirmationEmail(bookingData, bookingId, assignedMachine, cancelCode);

    return {
      success: true,
      bookingId: bookingId,
      machine: assignedMachine,
      cancelCode: cancelCode,
      emailSent: emailSent
    };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.message };
  } finally {
    lock.releaseLock();
  }
}

// สุ่มรหัสยกเลิก 4 หลัก (1000-9999)
function generateCancelCode() {
  return String(Math.floor(1000 + Math.random() * 9000));
}

// ==========================================
// 4. ฟังก์ชันส่งอีเมลยืนยันการจอง (แสดงรหัสยกเลิก)
// ==========================================
function sendConfirmationEmail(bookingData, bookingId, machineNumber, cancelCode) {
  try {
    const studentEmail = String(bookingData.studentId).trim() + "@msu.ac.th";
    const subject = `[MSU ProjBook] ยืนยันการจองโปรเจคเตอร์ - รหัสการจอง: ${bookingId}`;
    const htmlMessage = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e5e7eb; border-radius: 12px; background-color: #ffffff;">
        <h2 style="color: #2563eb; text-align: center; margin-bottom: 20px;">ยืนยันการจองโปรเจคเตอร์สำเร็จ</h2>
        <p>เรียนคุณ <strong>${bookingData.studentName}</strong>,</p>
        <p>ระบบได้รับการจองโปรเจคเตอร์เรียบร้อยแล้ว โดยมีรายละเอียดดังนี้:</p>
        <div style="background-color: #f3f4f6; padding: 15px; border-radius: 8px; margin: 20px 0;">
          <p style="margin: 5px 0;"><strong>รหัสการจอง:</strong> <span style="color: #2563eb; font-weight: bold;">${bookingId}</span></p>
          <p style="margin: 5px 0;"><strong>รหัสนิสิต:</strong> ${bookingData.studentId}</p>
          <p style="margin: 5px 0;"><strong>ชื่อกลุ่ม/โครงการ:</strong> ${bookingData.groupName}</p>
          <p style="margin: 5px 0;"><strong>วันที่ใช้งาน:</strong> ${bookingData.date}</p>
          <p style="margin: 5px 0;"><strong>รอบเวลา:</strong> ${bookingData.slot}</p>
          <p style="margin: 5px 0;"><strong>อุปกรณ์ที่ได้รับ:</strong> <span style="color: #059669; font-weight: bold;">โปรเจคเตอร์ เครื่องที่ ${machineNumber}</span></p>
        </div>
        <div style="background-color: #fef3c7; border: 1px dashed #f59e0b; padding: 15px; border-radius: 8px; margin: 20px 0; text-align: center;">
          <p style="margin: 0 0 5px 0; color: #92400e; font-size: 13px;">รหัสยกเลิกการจอง (เก็บไว้เป็นความลับ)</p>
          <p style="margin: 0; font-size: 28px; font-weight: bold; letter-spacing: 6px; color: #92400e;">${cancelCode}</p>
          <p style="margin: 8px 0 0 0; color: #92400e; font-size: 12px;">ใช้รหัสนี้เพื่อยกเลิกการจองด้วยตนเองที่หน้า "การจองของฉัน"</p>
        </div>
        <p style="color: #4b5563; font-size: 14px;">* กรุณาติดต่อขอรับอุปกรณ์และยื่นแสดงอีเมลนี้แก่เจ้าหน้าที่ตามวันและเวลาที่กำหนด</p>
        <hr style="border: none; border-top: 1px solid #e5e7eb; margin: 20px 0;" />
        <p style="text-align: center; color: #9ca3af; font-size: 12px;">MSU Projector Booking System</p>
      </div>
    `;
    MailApp.sendEmail({
      to: studentEmail,
      subject: subject,
      htmlBody: htmlMessage
    });
    return true;
  } catch (err) {
    console.error("sendConfirmationEmail error:", err);
    return false;
  }
}

// ==========================================
// 5. ยกเลิกการจอง (โดย Admin ในหน้า Admin)
// ==========================================
function cancelBooking(bookingId, adminPass) {
  if (adminPass !== ADMIN_PASSWORD) {
    return { success: false, message: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('การจอง');
    if (!sheet) return { success: false, message: 'ไม่พบชีตข้อมูลการจอง' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(bookingId).trim()) {
        sheet.getRange(i + 1, 9).setValue('CANCELLED');
        return { success: true, message: 'ยกเลิกสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบรายการจองนี้' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.message };
  }
}

// ==========================================
// 6. ยกเลิกการจองด้วยตนเอง (นิสิตกดที่รายการของตัวเอง + กรอกรหัสยกเลิก 4 หลัก)
// ==========================================
function cancelBookingByCode(bookingId, cancelCode) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('การจอง');
    if (!sheet) return { success: false, message: 'ไม่พบชีตข้อมูลการจอง' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(bookingId).trim()) {
        const currentStatus = String(data[i][8]).toUpperCase();
        if (currentStatus === 'CANCELLED') {
          return { success: false, message: 'รายการนี้ถูกยกเลิกไปแล้ว' };
        }
        const storedCode = String(data[i][10] || '').trim(); // คอลัมน์ K (index 10)
        if (storedCode === '' || storedCode !== String(cancelCode).trim()) {
          return { success: false, message: 'รหัสยกเลิกไม่ถูกต้อง' };
        }
        sheet.getRange(i + 1, 9).setValue('CANCELLED');
        return { success: true, message: 'ยกเลิกการจองสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบรายการจองนี้' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.message };
  }
}

// ==========================================
// 7. ยืนยันว่ามาใช้งานแล้ว (โดย Admin ในหน้า Admin)
// ==========================================
function markAsUsed(bookingId, adminPass) {
  if (adminPass !== ADMIN_PASSWORD) {
    return { success: false, message: 'รหัสผ่าน Admin ไม่ถูกต้อง' };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('การจอง');
    if (!sheet) return { success: false, message: 'ไม่พบชีตข้อมูลการจอง' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(bookingId).trim()) {
        const currentStatus = String(data[i][8]).toUpperCase();
        if (currentStatus === 'CANCELLED') {
          return { success: false, message: 'รายการนี้ถูกยกเลิกไปแล้ว ไม่สามารถยืนยันการเข้าใช้งานได้' };
        }
        sheet.getRange(i + 1, 9).setValue('USED');
        return { success: true, message: 'ยืนยันการเข้าใช้งานสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบรายการจองนี้' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.message };
  }
}

function setupBookingSheet(ss) {
  const sheet = ss.insertSheet('การจอง');
  const headers = ['Booking ID', 'Student ID', 'Student Name', 'Group Name',
    'Rep Name', 'Date', 'Slot', 'Machine', 'Status', 'CreatedAt', 'CancelCode'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f4f6');
  sheet.setFrozenRows(1);
  return sheet;
}
