function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.setSpreadsheetTimeZone('Asia/Vientiane');
  ss.setSpreadsheetLocale('th_TH');

  const tables = {
    'หน้าหลัก': [
      'รายการสรุป', 'จำนวน'
    ],

    'พุธ-พฤหัส': [
      'วันที่เดินทาง', 'รอบรถ', 'เส้นทาง', 'รถ',
      'คนขับ', 'จำนวนผู้จอง', 'รายชื่อผู้โดยสาร', 'สถานะ'
    ],

    'ศุกร์-เสาร์': [
      'วันที่เดินทาง', 'รอบรถ', 'เส้นทาง', 'รถ',
      'คนขับ', 'จำนวนผู้จอง', 'รายชื่อผู้โดยสาร', 'สถานะ'
    ],

    'ข้อมูลการจอง': [
      'เวลาจอง', 'รหัสการจอง', 'ประเภทผู้จอง',
      'ชื่อ-นามสกุล', 'เบอร์โทรศัพท์', 'วันที่เดินทาง',
      'รหัสรอบ', 'ชื่อรอบ', 'จุดขึ้นรถ', 'จุดลงรถ',
      'จำนวนผู้โดยสาร', 'สถานะ', 'หมายเหตุ'
    ],

    'รายชื่อพนักงาน': [
      'รหัสพนักงาน', 'ชื่อ-นามสกุล', 'แผนก',
      'สถานที่ทำงาน', 'สถานะ'
    ],

    'ข้อมูลรถ': [
      'รหัสรถ', 'ทะเบียนรถ', 'ชื่อคนขับ',
      'จำนวนที่นั่งอ้างอิง', 'สถานะรถ'
    ],

    'จุดขึ้น-ลง': [
      'รหัสจุด', 'ชื่อจุด', 'จังหวัด/พื้นที่',
      'รายละเอียดเพิ่มเติม', 'สถานะ'
    ],

    'รอบเดินรถ': [
      'รหัสรอบ', 'วันประจำสัปดาห์', 'ชื่อวัน',
      'ชื่อรอบ', 'ทิศทาง', 'จุดเริ่มต้น',
      'จุดหมาย', 'เวลาออก', 'เวลาถึง',
      'รหัสรถ', 'คนขับ', 'สถานะ', 'หมายเหตุ'
    ],

    'บัญชีผู้ดูแล': [
      'ชื่อผู้ใช้', 'ชื่อผู้ดูแล', 'Salt',
      'Password Hash', 'สิทธิ์', 'สถานะ',
      'วันที่สร้าง', 'เข้าสู่ระบบล่าสุด'
    ],

    'บัญชีคนขับ': [
      'ชื่อผู้ใช้', 'รหัสผ่าน Hash', 'Salt',
      'ชื่อคนขับ', 'รหัสรถ', 'สถานะ',
      'วันที่สร้าง', 'เข้าสู่ระบบล่าสุด'
    ],

    'ยอดผู้โดยสาร': [
      'รหัสรายการ', 'วันที่เดินทาง', 'รหัสรอบ',
      'ชื่อรอบ', 'จำนวนผู้โดยสารเริ่มต้น',
      'จำนวนผู้โดยสารคงเหลือ', 'สถานะรอบ',
      'อัปเดตล่าสุด', 'ผู้แก้ไขล่าสุด'
    ],

    'ประวัติยอดผู้โดยสาร': [
      'เวลาที่แก้ไข', 'รหัสรายการ', 'วันที่เดินทาง',
      'รหัสรอบ', 'จำนวนก่อนแก้ไข',
      'จำนวนหลังแก้ไข', 'ผู้แก้ไข', 'รายละเอียด'
    ],

    'มอบหมายรถ': [
      'รหัสรายการ', 'วันที่เดินทาง', 'รหัสรอบ',
      'ชื่อรอบ', 'รหัสรถ', 'ชื่อคนขับ',
      'เวลาออก', 'เวลาถึง', 'สถานะ', 'หมายเหตุ'
    ],

    'รอบพิเศษ': [
      'รหัสรอบพิเศษ', 'วันที่เดินทาง', 'ชื่อรอบ',
      'ทิศทาง', 'จุดเริ่มต้น', 'จุดหมาย',
      'เวลาออก', 'เวลาถึง', 'รหัสรถ',
      'คนขับ', 'สถานะ', 'หมายเหตุ'
    ],
    'การแจ้งเตือน': [
  'รหัสแจ้งเตือน', 'วันที่เดินทาง', 'รหัสรอบ',
  'ชื่อรอบ', 'ประเภทการแจ้งเตือน', 'รายละเอียด',
  'จำนวนผู้โดยสาร', 'ความจุรถ', 'สถานะ'
],
  };

  Object.entries(tables).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    // เพิ่มหัวตารางเฉพาะชีตที่ยังว่างเท่านั้น
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers]);
    }

    sheet.setFrozenRows(1);

    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#D9EAD3')
      .setWrap(true);
  });

  // เพิ่มรอบประจำเริ่มต้นเฉพาะกรณีที่ยังไม่มีข้อมูลรอบ
  const tripSheet = ss.getSheetByName('รอบเดินรถ');

  // หากชีตยังไม่มีข้อมูล ให้เพิ่มรอบเดิม R001-R004 ก่อน
  if (tripSheet.getLastRow() === 1) {
    const trips = [
      ['R001', 3, 'วันพุธ', 'พุธขาไป', 'ขาไป',
        'ลำปาง', 'ลาว', '', '', '', '', 'เปิดใช้งาน', ''],

      ['R002', 4, 'วันพฤหัสบดี', 'พฤหัสขากลับ', 'ขากลับ',
        'ลาว', 'ลำปาง', '', '', '', '', 'เปิดใช้งาน', ''],

      ['R003', 5, 'วันศุกร์', 'ศุกร์ขาไป', 'ขาไป',
        'ลำปาง', 'ลาว', '', '', '', '', 'เปิดใช้งาน', ''],

      ['R004', 6, 'วันเสาร์', 'เสาร์ขากลับ', 'ขากลับ',
        'ลาว', 'ลำปาง', '', '', '', '', 'เปิดใช้งาน', '']
    ];

    tripSheet.getRange(2, 1, trips.length, 13)
      .setValues(trips);
  }

  // เพิ่ม R005 และ R006 เฉพาะรหัสที่ยังไม่มีในชีต
  const existingTripIds = tripSheet.getLastRow() >= 2
    ? tripSheet.getRange(
        2, 1, tripSheet.getLastRow() - 1, 1
      ).getValues().flat().map(id => String(id).trim())
    : [];

  const newTrips = [
    ['R005', 1, 'วันจันทร์', 'จันทร์ขากลับ', 'ขากลับ',
      'ลาว', 'ลำปาง', '', '', '', '', 'เปิดใช้งาน', ''],

    ['R006', 0, 'วันอาทิตย์', 'อาทิตย์ขาไป', 'ขาไป',
      'ลำปาง', 'ลาว', '', '', '', '', 'เปิดใช้งาน', '']
  ];

  const tripsToAdd = newTrips.filter(
    trip => !existingTripIds.includes(trip[0])
  );

  if (tripsToAdd.length > 0) {
    const startRow = tripSheet.getLastRow() + 1;

    tripSheet.getRange(
      startRow, 1, tripsToAdd.length, 13
    ).setValues(tripsToAdd);
  }


  SpreadsheetApp.getUi().alert(
    'ตรวจสอบและเตรียมฐานข้อมูล Employee Shuttle เรียบร้อยแล้ว'
  );
}
// เปิดหน้าเว็บไซต์จองรถ
function doGet(e) {
  const page = e && e.parameter ? e.parameter.page : '';

  if (page === 'adminSetup') {
    return HtmlService.createHtmlOutputFromFile('AdminSetup')
      .setTitle('ตั้งค่าผู้ดูแลระบบ');
  }

  if (page === 'adminLogin') {
    return HtmlService.createHtmlOutputFromFile('AdminLogin')
      .setTitle('เข้าสู่ระบบผู้ดูแล');
  }

  if (page === 'adminDashboard') {
    return HtmlService.createHtmlOutputFromFile('AdminDashboard')
      .setTitle('ระบบจัดการรถตู้บริษัท');
  }

  if (page === 'driver') {
  return HtmlService.createHtmlOutputFromFile('Driver')
    .setTitle('ระบบสำหรับคนขับรถ');
}

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Employee Shuttle');
}

// บันทึกข้อมูลการจอง

function saveBooking(data) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const bookingSheet = ss.getSheetByName('ข้อมูลการจอง');
    const tripSheet = ss.getSheetByName('รอบเดินรถ');
    const specialSheet = ss.getSheetByName('รอบพิเศษ');
    const assignmentSheet = ss.getSheetByName('มอบหมายรถ');
    const vehicleSheet = ss.getSheetByName('ข้อมูลรถ');
    const alertSheet = ss.getSheetByName('การแจ้งเตือน');

    if (!bookingSheet || !tripSheet || !specialSheet ||
    !assignmentSheet || !vehicleSheet || !alertSheet) {
  throw new Error('ไม่พบชีตฐานข้อมูลที่จำเป็น');
}
    // 1. รับข้อมูลจากหน้าจอง
    const bookingType = String(data.bookingType || '').trim();
    const fullName = String(data.fullName || '').trim();
    const phone = String(data.phone || '').trim();
    const passengerCount = Number(data.passengerCount);
    const travelDate = String(data.travelDate || '').trim();
    const tripId = String(data.trip || data.tripId || '').trim();
    const pickup = String(data.pickup || '').trim();
    const dropoff = String(data.dropoff || '').trim();
    const note = String(data.note || '').trim();

    if (!bookingType || !fullName || !phone ||
        !travelDate || !tripId || !pickup || !dropoff) {
      throw new Error('กรุณากรอกข้อมูลให้ครบทุกช่อง');
    }

    if (!['พนักงาน', 'บุคคลภายนอก'].includes(bookingType)) {
      throw new Error('ประเภทผู้จองไม่ถูกต้อง');
    }

    if (fullName.length > 150 || phone.length > 30) {
      throw new Error('ชื่อหรือเบอร์โทรศัพท์ยาวเกินกำหนด');
    }

    if (!Number.isInteger(passengerCount) ||
        passengerCount < 1 || passengerCount > 100) {
      throw new Error('จำนวนผู้โดยสารต้องเป็นจำนวนเต็มตั้งแต่ 1-100 คน');
    }

    if (pickup.toLowerCase() === dropoff.toLowerCase()) {
      throw new Error('จุดขึ้นรถและจุดลงรถต้องไม่เหมือนกัน');
    }

    // 2. ตรวจสอบวันที่
    const match = travelDate.match(/^(\d{4})-(\d{2})-(\d{2})$/);

    if (!match) {
      throw new Error('รูปแบบวันที่ไม่ถูกต้อง');
    }

    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);

    const date = new Date(year, month - 1, day);

    if (
      date.getFullYear() !== year ||
      date.getMonth() !== month - 1 ||
      date.getDate() !== day
    ) {
      throw new Error('วันที่เดินทางไม่ถูกต้อง');
    }

    // 3. ค้นหารอบประจำก่อน
    const regularRows = tripSheet.getDataRange().getValues();

    let trip = regularRows.find((row, index) =>
      index > 0 && String(row[0]).trim() === tripId
    );

    let tripName = '';
    let isSpecial = false;
    let specialTrip = null;

    if (trip) {
      if (String(trip[11]).trim() !== 'เปิดใช้งาน') {
        throw new Error('รอบรถนี้ยังไม่เปิดให้จอง');
      }

      if (date.getDay() !== Number(trip[1])) {
        throw new Error('วันที่เลือกไม่ตรงกับวันเดินรถของรอบนี้');
      }

      tripName = String(trip[3]).trim();

    } else {
      // 4. ถ้าไม่พบรอบประจำ ให้ค้นหารอบพิเศษ
      const specialRows = specialSheet.getDataRange().getValues();

      specialTrip = specialRows.find((row, index) => {
        if (index === 0) return false;

        const savedDate = row[1] instanceof Date
          ? Utilities.formatDate(
              row[1], 'Asia/Vientiane', 'yyyy-MM-dd'
            )
          : String(row[1]).trim();

        return (
          String(row[0]).trim() === tripId &&
          savedDate === travelDate
        );
      });

      if (!specialTrip) {
        throw new Error('ไม่พบรอบรถในวันที่เลือก');
      }

      if (String(specialTrip[10]).trim() !== 'เปิดใช้งาน') {
        throw new Error('รอบพิเศษนี้ยังไม่เปิดให้จอง');
      }

      isSpecial = true;
      tripName = String(specialTrip[2]).trim();
    }

    // 5. ตรวจสอบการจองซ้ำด้วยเบอร์โทรศัพท์
    const lastRow = bookingSheet.getLastRow();

    if (lastRow >= 2) {
      const bookings = bookingSheet
        .getRange(2, 1, lastRow - 1, 13)
        .getValues();

      const duplicate = bookings.some(row => {
        const savedDate = row[5] instanceof Date
          ? Utilities.formatDate(
              row[5], 'Asia/Vientiane', 'yyyy-MM-dd'
            )
          : String(row[5]).trim();

        return (
          String(row[4]).trim() === phone &&
          savedDate === travelDate &&
          String(row[6]).trim() === tripId &&
          String(row[11]).trim() !== 'ยกเลิก'
        );
      });

      if (duplicate) {
        throw new Error('เบอร์โทรศัพท์นี้มีรายการจองรอบนี้แล้ว');
      }
    }

    // 6. ค้นหารถที่ได้รับมอบหมายในวันเดินทาง
    const assignmentRows = assignmentSheet.getDataRange().getValues();

    let assignment = assignmentRows.find((row, index) => {
      if (index === 0) return false;

      const savedDate = row[1] instanceof Date
        ? Utilities.formatDate(
            row[1], 'Asia/Vientiane', 'yyyy-MM-dd'
          )
        : String(row[1]).trim();

      return (
        savedDate === travelDate &&
        String(row[2]).trim() === tripId &&
        String(row[8]).trim() === 'เปิดใช้งาน'
      );
    });

    let vehicleId = '';
    let driverName = '';

    if (assignment) {
      vehicleId = String(assignment[4]).trim();
      driverName = String(assignment[5]).trim();

    } else if (isSpecial) {
      vehicleId = String(specialTrip[8]).trim();
      driverName = String(specialTrip[9]).trim();

    } else {
      vehicleId = String(trip[9]).trim();
      driverName = String(trip[10]).trim();
    }

    // 7. ตรวจสอบความจุรถ
    let capacity = 0;

    if (vehicleId) {
      const vehicles = vehicleSheet.getDataRange().getValues();

      const vehicle = vehicles.find((row, index) =>
        index > 0 &&
        String(row[0]).trim() === vehicleId &&
        String(row[4]).trim() === 'เปิดใช้งาน'
      );

      if (vehicle) {
        capacity = Number(vehicle[3]) || 0;
      }
    }

    // รวมยอดผู้โดยสารเดิมของรอบเดียวกัน
    let alreadyBooked = 0;

    if (lastRow >= 2) {
      const bookings = bookingSheet
        .getRange(2, 1, lastRow - 1, 13)
        .getValues();

      bookings.forEach(row => {
        const savedDate = row[5] instanceof Date
          ? Utilities.formatDate(
              row[5], 'Asia/Vientiane', 'yyyy-MM-dd'
            )
          : String(row[5]).trim();

        if (
          savedDate === travelDate &&
          String(row[6]).trim() === tripId &&
          String(row[11]).trim() !== 'ยกเลิก'
        ) {
          alreadyBooked += Number(row[10]) || 0;
        }
      });
    }

    const totalPassengers = alreadyBooked + passengerCount;

    const overCapacity =
      capacity > 0 && totalPassengers > capacity;

    // 8. สร้างรหัสการจอง
    const nextRow = bookingSheet.getLastRow() + 1;

    const bookingId = 'BK-' +
      Utilities.formatDate(
        new Date(),
        'Asia/Vientiane',
        'yyyyMMdd'
      ) + '-' +
      String(nextRow - 1).padStart(4, '0');

    // 9. บันทึกข้อมูล 13 คอลัมน์
    bookingSheet.getRange(nextRow, 1, 1, 13).setValues([[
      new Date(),
      bookingId,
      bookingType,
      fullName,
      phone,
      date,
      tripId,
      tripName,
      pickup,
      dropoff,
      passengerCount,
      'จองสำเร็จ',
      note
    ]]);
    // บันทึกคำเตือนเมื่อผู้โดยสารเกินความจุรถ
let alertError = '';

if (overCapacity) {
  try {
    upsertOverCapacityAlert_(
      alertSheet,
      travelDate,
      tripId,
      tripName,
      totalPassengers,
      capacity
    );
  } catch (error) {
    alertError = error.message;
  }
}

    // 10. ส่งผลกลับไปยังหน้าจอง
    return {
      success: true,
      bookingId: bookingId,
      fullName: fullName,
      travelDate: travelDate,
      tripName: tripName,
      pickup: pickup,
      dropoff: dropoff,
      passengerCount: passengerCount,
      totalPassengers: totalPassengers,
      capacity: capacity,
      vehicleId: vehicleId,
      driverName: driverName,
      overCapacity: overCapacity,
      warning: overCapacity
        ? 'คำเตือน: จำนวนผู้โดยสาร ' + totalPassengers +
          ' คน เกินความจุรถ ' + capacity + ' ที่นั่ง'
        : '',
      alertError: alertError
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// ดึงรายชื่อผู้จองในวันและรอบรถเดียวกัน
function getTripPassengers(travelDate, tripId) {
  try {
    if (!travelDate || !tripId) {
      throw new Error('กรุณาเลือกวันที่และรอบรถ');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ข้อมูลการจอง');

    if (!sheet) {
      throw new Error('ไม่พบชีตข้อมูลการจอง');
    }

    const values = sheet.getDataRange().getValues();
    const timezone = 'Asia/Vientiane';
    const passengers = [];
    let totalPassengers = 0;

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      // วันที่เดินทาง
      let savedDate = '';

      if (row[5] instanceof Date) {
        savedDate = Utilities.formatDate(
          row[5], timezone, 'yyyy-MM-dd'
        );
      } else if (row[5]) {
        savedDate = String(row[5]).trim();
      }

      const savedTripId = String(row[6] || '').trim();
      const status = String(row[11] || '').trim();

      // เอาเฉพาะวันและรอบที่ตรงกัน และจองสำเร็จ
      if (
        savedDate !== travelDate ||
        savedTripId !== String(tripId).trim() ||
        status !== 'จองสำเร็จ'
      ) {
        continue;
      }

      const count = Number(row[10]) || 0;

      passengers.push({
        bookingType: String(row[2] || ''),
        fullName: String(row[3] || ''),
        pickup: String(row[8] || ''),
        dropoff: String(row[9] || ''),
        passengerCount: count
      });

      totalPassengers += count;
    }

    return {
      success: true,
      passengers: passengers,
      totalPassengers: totalPassengers,
      bookingCount: passengers.length
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}


function generateAdminSetupKey() {
  const props = PropertiesService.getScriptProperties();

  if (props.getProperty('ADMIN_SETUP_KEY')) {
    throw new Error('มีรหัสสมัครผู้ดูแลอยู่แล้ว');
  }

  const key =
    Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '');

  props.setProperty('ADMIN_SETUP_KEY', key);

  return {
    success: true,
    setupKey: key,
    message: 'สร้างรหัสสมัครผู้ดูแลแล้ว กรุณาเก็บรักษาไว้เป็นความลับ'
  };
}

function createFirstAdmin(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const props = PropertiesService.getScriptProperties();
    const setupKey = props.getProperty('ADMIN_SETUP_KEY');

    if (!setupKey || data.setupKey !== setupKey) {
      throw new Error('รหัสอนุญาตไม่ถูกต้องหรือหมดอายุ');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('บัญชีผู้ดูแล');

    if (!sheet) {
      throw new Error('ไม่พบชีตบัญชีผู้ดูแล');
    }

    if (sheet.getLastRow() > 1) {
      throw new Error('มีบัญชีผู้ดูแลอยู่แล้ว');
    }

    const username = String(data.username || '').trim();
    const adminName = String(data.adminName || '').trim();
    const password = String(data.password || '');

    if (!/^[a-zA-Z0-9_]{4,30}$/.test(username)) {
      throw new Error('ชื่อผู้ใช้ต้องมี 4-30 ตัว ใช้ภาษาอังกฤษ ตัวเลข หรือ _');
    }

    if (!adminName || adminName.length > 100) {
      throw new Error('กรุณากรอกชื่อผู้ดูแลให้ถูกต้อง');
    }

    if (password.length < 12 || password.length > 128) {
      throw new Error('รหัสผ่านต้องมีความยาว 12-128 ตัวอักษร');
    }

    const salt = Utilities.getUuid() + Utilities.getUuid();
    const pepper = Utilities.getUuid() + Utilities.getUuid();

    const hash = hashAdminPassword_(password, salt, pepper);

    props.setProperty('ADMIN_PASSWORD_PEPPER', pepper);

    sheet.appendRow([
      username,
      adminName,
      salt,
      hash,
      'SUPER_ADMIN',
      'เปิดใช้งาน',
      new Date(),
      ''
    ]);

    props.deleteProperty('ADMIN_SETUP_KEY');

    return {
      success: true,
      message: 'สร้างบัญชีผู้ดูแลหลักสำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };

  } finally {
    lock.releaseLock();
  }
}
function hashAdminPassword_(password, salt, pepper) {
  let hash = salt + password;

  for (let i = 0; i < 10000; i++) {
    hash = Utilities.base64Encode(
      Utilities.computeHmacSha256Signature(
        hash,
        pepper,
        Utilities.Charset.UTF_8
      )
    );
  }

  return hash;
}

function loginAdmin(data) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const username = String(data.username || '').trim().toLowerCase();
    const password = String(data.password || '');

    if (!username || !password) {
      throw new Error('กรุณากรอกชื่อผู้ใช้และรหัสผ่าน');
    }

    const props = PropertiesService.getScriptProperties();
    const pepper = props.getProperty('ADMIN_PASSWORD_PEPPER');

    if (!pepper) {
      throw new Error('ระบบตรวจสอบรหัสผ่านยังไม่พร้อม');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('บัญชีผู้ดูแล');

    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error('ไม่พบบัญชีผู้ดูแล');
    }

    const rows = sheet.getRange(
      2, 1, sheet.getLastRow() - 1, 8
    ).getValues();

    const index = rows.findIndex(row =>
      String(row[0]).trim().toLowerCase() === username
    );

    if (index === -1) {
      throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
    }

    const row = rows[index];

    if (
      String(row[5]).trim() !== 'เปิดใช้งาน' ||
      String(row[4]).trim() !== 'SUPER_ADMIN'
    ) {
      throw new Error('บัญชีนี้ไม่มีสิทธิ์เข้าสู่ระบบ');
    }

    const salt = String(row[2]);
    const savedHash = String(row[3]);
    const enteredHash = hashAdminPassword_(
      password, salt, pepper
    );

    if (savedHash !== enteredHash) {
      throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
    }

    const token =
      Utilities.getUuid().replace(/-/g, '') +
      Utilities.getUuid().replace(/-/g, '');

    const session = {
      username: String(row[0]),
      adminName: String(row[1]),
      role: String(row[4])
    };

    CacheService.getScriptCache().put(
      'ADMIN_SESSION_' + token,
      JSON.stringify(session),
      21600
    );

    sheet.getRange(index + 2, 8).setValue(new Date());

    return {
      success: true,
      token: token,
      adminName: session.adminName,
      message: 'เข้าสู่ระบบสำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}



function checkAdminSession(token) {
  if (!token || typeof token !== 'string') {
    return { success: false };
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'ADMIN_SESSION_' + token;
  const session = cache.get(cacheKey);

  if (!session) {
    return { success: false };
  }

  try {
    const data = JSON.parse(session);

    if (data.role !== 'SUPER_ADMIN' || !data.username) {
      cache.remove(cacheKey);
      return { success: false };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('บัญชีผู้ดูแล');

    if (!sheet || sheet.getLastRow() < 2) {
      cache.remove(cacheKey);
      return { success: false };
    }

    const rows = sheet.getRange(
      2, 1, sheet.getLastRow() - 1, 8
    ).getValues();

    const account = rows.find(row =>
      String(row[0]).trim().toLowerCase() ===
      String(data.username).trim().toLowerCase()
    );

    if (
      !account ||
      String(account[5]).trim() !== 'เปิดใช้งาน' ||
      String(account[4]).trim() !== 'SUPER_ADMIN'
    ) {
      cache.remove(cacheKey);
      return { success: false };
    }

    return {
      success: true,
      adminName: String(account[1]),
      username: String(account[0]),
      role: String(account[4])
    };

  } catch (error) {
    cache.remove(cacheKey);
    return { success: false };
  }
}



function logoutAdmin(token) {
  if (!token || typeof token !== 'string') {
    return { success: false };
  }

  CacheService.getScriptCache().remove(
    'ADMIN_SESSION_' + token
  );

  return { success: true };
}


function getAvailableTrips(dateString) {
  try {
    // 1. ตรวจสอบรูปแบบวันที่
    if (
      !dateString ||
      !/^\d{4}-\d{2}-\d{2}$/.test(dateString)
    ) {
      return {
        success: false,
        message: 'กรุณาเลือกวันที่เดินทาง'
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timezone = 'Asia/Vientiane';

    const parts = dateString.split('-').map(Number);

    const travelDate = new Date(
      parts[0], parts[1] - 1, parts[2]
    );

    if (
      travelDate.getFullYear() !== parts[0] ||
      travelDate.getMonth() !== parts[1] - 1 ||
      travelDate.getDate() !== parts[2]
    ) {
      throw new Error('วันที่เดินทางไม่ถูกต้อง');
    }

    const weekday = travelDate.getDay();
    const trips = [];

    // 2. ฟังก์ชันแปลงวันที่ให้เป็น yyyy-MM-dd
    function formatDate(value) {
      if (!value) return '';

      if (value instanceof Date) {
        return Utilities.formatDate(
          value, timezone, 'yyyy-MM-dd'
        );
      }

      const text = String(value).trim();

      if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
        return text;
      }

      return text;
    }

    // 3. ฟังก์ชันแปลงเวลา
    function formatTime(value) {
  if (value === '' || value === null || value === undefined) {
    return '';
  }

  // กรณี Google Sheets ส่งค่าเวลาเป็น Date
  if (value instanceof Date) {
    return Utilities.formatDate(
      value, timezone, 'HH:mm'
    );
  }

  // กรณี Google Sheets ส่งค่าเวลาเป็นตัวเลข
  if (typeof value === 'number') {
    const totalMinutes = Math.round(value * 1440) % 1440;

    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;

    return String(hours).padStart(2, '0') + ':' +
      String(minutes).padStart(2, '0');
  }

  return String(value).trim();
}

    // 4. อ่านข้อมูลการมอบหมายรถ
    const assignmentSheet =
      ss.getSheetByName('มอบหมายรถ');

    if (!assignmentSheet) {
      throw new Error('ไม่พบชีตมอบหมายรถ');
    }

    const assignmentData =
      assignmentSheet.getDataRange().getValues();

    // 5. ค้นหาการมอบหมายรถของแต่ละรอบ
    function findAssignment(tripId) {
      return assignmentData.find((row, index) => {
        if (index === 0) return false;

        return (
          formatDate(row[1]) === dateString &&
          String(row[2]).trim() === tripId &&
          String(row[8]).trim() === 'เปิดใช้งาน'
        );
      });
    }

    // 6. โหลดรอบประจำ
    const regularSheet =
      ss.getSheetByName('รอบเดินรถ');

    if (!regularSheet) {
      throw new Error('ไม่พบชีตรอบเดินรถ');
    }

    const regularData =
      regularSheet.getDataRange().getValues();

    for (let i = 1; i < regularData.length; i++) {
      const row = regularData[i];

      if (
        Number(row[1]) !== weekday ||
        String(row[11]).trim() !== 'เปิดใช้งาน'
      ) {
        continue;
      }

      const tripId = String(row[0]).trim();
      const assignment = findAssignment(tripId);

      // ใช้ข้อมูลรถและคนขับจากการมอบหมายก่อน
      // หากไม่มีข้อมูล ให้ใช้ค่าตั้งต้นจากรอบประจำ
      const vehicleId = assignment &&
        String(assignment[4]).trim()
          ? String(assignment[4]).trim()
          : String(row[9] || '').trim();

      const driverName = assignment &&
        String(assignment[5]).trim()
          ? String(assignment[5]).trim()
          : String(row[10] || '').trim();

      trips.push({
        id: tripId,
        name: String(row[3]).trim(),
        direction: String(row[4]).trim(),
        start: String(row[5]).trim(),
        destination: String(row[6]).trim(),
        departure: assignment &&
  assignment[6] !== '' &&
  assignment[6] !== null
    ? formatTime(assignment[6])
    : formatTime(row[7]),

arrival: assignment &&
  assignment[7] !== '' &&
  assignment[7] !== null
    ? formatTime(assignment[7])
    : formatTime(row[8]),
        vehicleId: vehicleId,
        driverName: driverName,
        type: 'regular'
      });
    }

    // 7. โหลดรอบพิเศษ
    const specialSheet =
      ss.getSheetByName('รอบพิเศษ');

    if (!specialSheet) {
      throw new Error('ไม่พบชีตรอบพิเศษ');
    }

    const specialData =
      specialSheet.getDataRange().getValues();

    for (let i = 1; i < specialData.length; i++) {
      const row = specialData[i];

      if (
        formatDate(row[1]) !== dateString ||
        String(row[10]).trim() !== 'เปิดใช้งาน'
      ) {
        continue;
      }

      const tripId = String(row[0]).trim();
      const assignment = findAssignment(tripId);

      // ใช้ข้อมูลรถและคนขับจากการมอบหมายก่อน
      // หากไม่มีข้อมูล ให้ใช้ค่าตั้งต้นจากรอบพิเศษ
      const vehicleId = assignment &&
        String(assignment[4]).trim()
          ? String(assignment[4]).trim()
          : String(row[8] || '').trim();

      const driverName = assignment &&
        String(assignment[5]).trim()
          ? String(assignment[5]).trim()
          : String(row[9] || '').trim();

      trips.push({
        id: tripId,
        name: String(row[2]).trim(),
        direction: String(row[3]).trim(),
        start: String(row[4]).trim(),
        destination: String(row[5]).trim(),
        departure: assignment &&
  assignment[6] !== '' &&
  assignment[6] !== null
    ? formatTime(assignment[6])
    : formatTime(row[6]),

arrival: assignment &&
  assignment[7] !== '' &&
  assignment[7] !== null
    ? formatTime(assignment[7])
    : formatTime(row[7]),
        vehicleId: vehicleId,
        driverName: driverName,
        type: 'special'
      });
    }

    // 8. ป้องกันรหัสรอบซ้ำ
    const ids = trips.map(trip => trip.id);

    if (new Set(ids).size !== ids.length) {
      throw new Error(
        'พบรหัสรอบซ้ำ กรุณาตรวจสอบชีตรอบเดินรถและรอบพิเศษ'
      );
    }

    // 9. ส่งข้อมูลกลับไปยังหน้าเว็บ
    return {
      success: true,
      trips: trips
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

// ตรวจสอบรอบที่ยังไม่ได้กำหนดรถหรือคนขับ
function getUnassignedTripWarnings(token) {
  try {
    const session = checkAdminSession(token);

    if (!session.success) {
      throw new Error('กรุณาเข้าสู่ระบบผู้ดูแลก่อน');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timezone = 'Asia/Vientiane';

    const tripSheet = ss.getSheetByName('รอบเดินรถ');
    const specialSheet = ss.getSheetByName('รอบพิเศษ');
    const assignmentSheet = ss.getSheetByName('มอบหมายรถ');

    if (!tripSheet || !specialSheet || !assignmentSheet) {
      throw new Error('ไม่พบชีตข้อมูลรอบรถที่จำเป็น');
    }

    const warnings = [];

    const assignmentRows = assignmentSheet.getDataRange().getValues();
    const regularRows = tripSheet.getDataRange().getValues();
    const specialRows = specialSheet.getDataRange().getValues();

    const todayString = Utilities.formatDate(
      new Date(), timezone, 'yyyy-MM-dd'
    );

    const todayParts = todayString.split('-').map(Number);

    const today = Date.UTC(
      todayParts[0],
      todayParts[1] - 1,
      todayParts[2]
    );

    function formatDate(value) {
      if (value instanceof Date) {
        return Utilities.formatDate(
          value, timezone, 'yyyy-MM-dd'
        );
      }

      return String(value || '').trim();
    }

    function isFutureOrToday(dateString) {
      const parts = dateString.split('-').map(Number);

      if (parts.length !== 3 || parts.some(isNaN)) {
        return false;
      }

      const dateValue = Date.UTC(
        parts[0], parts[1] - 1, parts[2]
      );

      return dateValue >= today;
    }

    function findAssignment(date, tripId) {
      return assignmentRows.find((row, index) => {
        if (index === 0) return false;

        return (
          formatDate(row[1]) === date &&
          String(row[2]).trim() === tripId &&
          String(row[8]).trim() === 'เปิดใช้งาน'
        );
      });
    }

    function checkWarning(
      date, tripId, tripName, defaultVehicle, defaultDriver
    ) {
      const assignment = findAssignment(date, tripId);

      const vehicleId = assignment
        ? String(assignment[4] || '').trim()
        : String(defaultVehicle || '').trim();

      const driverName = assignment
        ? String(assignment[5] || '').trim()
        : String(defaultDriver || '').trim();

      if (!vehicleId || !driverName) {
        warnings.push({
          date: date,
          tripId: tripId,
          tripName: tripName,
          message: !vehicleId && !driverName
            ? 'ยังไม่ได้กำหนดรถและคนขับ'
            : !vehicleId
              ? 'ยังไม่ได้กำหนดรถ'
              : 'ยังไม่ได้กำหนดคนขับ'
        });
      }
    }

    // ตรวจรอบประจำ 30 วัน
    for (let i = 1; i < regularRows.length; i++) {
      const row = regularRows[i];

      if (String(row[11]).trim() !== 'เปิดใช้งาน') {
        continue;
      }

      for (let d = 0; d < 30; d++) {
        const dateValue = new Date(
          today + d * 86400000
        );

        const weekday = dateValue.getUTCDay();

        if (weekday !== Number(row[1])) {
          continue;
        }

        const dateString =
          dateValue.getUTCFullYear() + '-' +
          String(dateValue.getUTCMonth() + 1).padStart(2, '0') + '-' +
          String(dateValue.getUTCDate()).padStart(2, '0');

        checkWarning(
          dateString,
          String(row[0]).trim(),
          String(row[3]).trim(),
          row[9],
          row[10]
        );
      }
    }

    // ตรวจรอบพิเศษ เฉพาะวันนี้และวันข้างหน้า
    for (let i = 1; i < specialRows.length; i++) {
      const row = specialRows[i];

      if (String(row[10]).trim() !== 'เปิดใช้งาน') {
        continue;
      }

      const dateString = formatDate(row[1]);

      if (!isFutureOrToday(dateString)) {
        continue;
      }

      checkWarning(
        dateString,
        String(row[0]).trim(),
        String(row[2]).trim(),
        row[8],
        row[9]
      );
    }

    return {
      success: true,
      warnings: warnings,
      count: warnings.length
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}
// บันทึกหรือปรับปรุงการแจ้งเตือนเมื่อผู้โดยสารเกินความจุรถ
function upsertOverCapacityAlert_(
  sheet,
  travelDate,
  tripId,
  tripName,
  totalPassengers,
  capacity
) {
  const alertId =
    'OVER_CAPACITY|' + travelDate + '|' + tripId;

  const detail =
    'จำนวนผู้โดยสาร ' + totalPassengers +
    ' คน เกินความจุรถ ' + capacity + ' ที่นั่ง';

  const lastRow = sheet.getLastRow();
  let targetRow = 0;
  let currentStatus = 'รอดำเนินการ';

  if (lastRow >= 2) {
    const ids = sheet.getRange(
      2, 1, lastRow - 1, 1
    ).getValues();

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === alertId) {
        targetRow = i + 2;

        currentStatus = String(
          sheet.getRange(targetRow, 9).getValue() ||
          'รอดำเนินการ'
        ).trim();

        break;
      }
    }
  }

  const alertData = [[
    alertId,
    travelDate,
    tripId,
    tripName,
    'เกินความจุรถ',
    detail,
    totalPassengers,
    capacity,
    currentStatus
  ]];

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, 9)
      .setValues(alertData);
  } else {
    sheet.appendRow(alertData[0]);
  }
}
function getAdminAlerts(token) {
  try {
    const session = checkAdminSession(token);

    if (!session.success) {
      throw new Error('กรุณาเข้าสู่ระบบผู้ดูแลก่อน');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('การแจ้งเตือน');

    if (!sheet) {
      throw new Error('ไม่พบชีตการแจ้งเตือน');
    }

    const lastRow = sheet.getLastRow();
    const timezone = 'Asia/Vientiane';

    let alerts = [];

    if (lastRow >= 2) {
      const values = sheet
        .getRange(2, 1, lastRow - 1, 9)
        .getValues();

      alerts = values.map(function(row) {

        let date = '';

        if (row[1] instanceof Date) {
          date = Utilities.formatDate(
            row[1],
            timezone,
            'dd/MM/yyyy'
          );
        } else if (row[1]) {
          const dateText = String(row[1]).trim();

          const match = dateText.match(
            /^(\d{4})-(\d{2})-(\d{2})$/
          );

          date = match
            ? match[3] + '/' + match[2] + '/' + match[1]
            : dateText;
        }

        return {
          id: String(row[0] || ''),
          date: date,
          tripId: String(row[2] || ''),
          tripName: String(row[3] || ''),
          type: String(row[4] || ''),
          detail: String(row[5] || ''),
          totalPassengers: Number(row[6]) || 0,
          capacity: Number(row[7]) || 0,
          status: String(row[8] || '')
        };
      });
    }

    return {
      success: true,
      count: alerts.length,
      alerts: alerts
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

// ==========================================
// DRIVER SYSTEM : ระบบคนขับรถ
// ==========================================

// เข้าสู่ระบบคนขับ

function loginDriver(data) {
  try {
    const username = String(data.username || '').trim().toLowerCase();
    const password = String(data.password || '');

    if (!username || !password) {
      throw new Error('กรุณากรอกชื่อผู้ใช้และรหัสผ่าน');
    }

    const props = PropertiesService.getScriptProperties();
    const pepper = props.getProperty('ADMIN_PASSWORD_PEPPER');

    if (!pepper) {
      throw new Error('ระบบตรวจสอบรหัสผ่านยังไม่พร้อม');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('บัญชีคนขับ');

    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error('ยังไม่มีบัญชีคนขับในระบบ');
    }

    const rows = sheet.getRange(
      2, 1, sheet.getLastRow() - 1, 8
    ).getValues();

    const index = rows.findIndex(row =>
      String(row[0]).trim().toLowerCase() === username
    );

    if (index === -1) {
      throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
    }

    const row = rows[index];

    if (String(row[5]).trim() !== 'เปิดใช้งาน') {
      throw new Error('บัญชีคนขับนี้ถูกปิดใช้งาน');
    }

    const enteredHash = hashAdminPassword_(
      password,
      String(row[2]),
      pepper
    );

    if (String(row[1]) !== enteredHash) {
      throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
    }

    const token =
      Utilities.getUuid().replace(/-/g, '') +
      Utilities.getUuid().replace(/-/g, '');

    const session = {
      username: String(row[0]),
      driverName: String(row[3]),
      vehicleId: String(row[4]),
      role: 'DRIVER'
    };

    CacheService.getScriptCache().put(
      'DRIVER_SESSION_' + token,
      JSON.stringify(session),
      21600
    );

    sheet.getRange(index + 2, 8).setValue(new Date());

    return {
      success: true,
      token: token,
      driverName: session.driverName,
      vehicleId: session.vehicleId,
      message: 'เข้าสู่ระบบสำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}


// ตรวจสอบเซสชันคนขับ
function checkDriverSession(token) {
  if (!token || typeof token !== 'string') {
    return { success: false };
  }

  const cache = CacheService.getScriptCache();
  const key = 'DRIVER_SESSION_' + token;
  const saved = cache.get(key);

  if (!saved) {
    return { success: false };
  }

  try {
    const session = JSON.parse(saved);

    if (session.role !== 'DRIVER' || !session.username) {
      cache.remove(key);
      return { success: false };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('บัญชีคนขับ');

    if (!sheet || sheet.getLastRow() < 2) {
      cache.remove(key);
      return { success: false };
    }

    const rows = sheet.getRange(
      2, 1, sheet.getLastRow() - 1, 8
    ).getValues();

    const account = rows.find(row =>
      String(row[0]).trim().toLowerCase() ===
      session.username.toLowerCase()
    );

    if (
      !account ||
      String(account[5]).trim() !== 'เปิดใช้งาน'
    ) {
      cache.remove(key);
      return { success: false };
    }

    return {
      success: true,
      driverName: session.driverName,
      vehicleId: session.vehicleId
    };

  } catch (error) {
    cache.remove(key);
    return { success: false };
  }
}


// ดึงข้อมูลรอบรถและผู้โดยสารทั้งหมดของบริษัท

function getDriverDashboardData(token) {
  try {
    const session = checkDriverSession(token);

    if (!session.success) {
      throw new Error('กรุณาเข้าสู่ระบบคนขับก่อน');
    }
    updateCompletedDriverAssignments();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timezone = 'Asia/Vientiane';

    const bookingSheet = ss.getSheetByName('ข้อมูลการจอง');
    const tripSheet = ss.getSheetByName('รอบเดินรถ');
    const specialSheet = ss.getSheetByName('รอบพิเศษ');
    const assignmentSheet = ss.getSheetByName('มอบหมายรถ');
    const vehicleSheet = ss.getSheetByName('ข้อมูลรถ');

    if (
      !bookingSheet ||
      !tripSheet ||
      !specialSheet ||
      !assignmentSheet ||
      !vehicleSheet
    ) {
      throw new Error('ไม่พบชีตข้อมูลที่จำเป็น');
    }

    function formatDate(value) {
      if (value instanceof Date) {
        return Utilities.formatDate(
          value, timezone, 'yyyy-MM-dd'
        );
      }

      const text = String(value || '').trim();

      if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
        return text;
      }

      return '';
    }

    function formatTime(value) {
      if (value instanceof Date) {
        return Utilities.formatDate(
          value, timezone, 'HH:mm'
        );
      }

      return String(value || '').trim();
    }

    const regularRows = tripSheet.getDataRange().getValues();
    const specialRows = specialSheet.getDataRange().getValues();
    const assignmentRows = assignmentSheet.getDataRange().getValues();
    const vehicleRows = vehicleSheet.getDataRange().getValues();

    const bookings = bookingSheet.getLastRow() >= 2
      ? bookingSheet.getRange(
          2, 1, bookingSheet.getLastRow() - 1, 13
        ).getValues()
      : [];

    // ==================================
    // 1. สร้างข้อมูลรอบรถต้นทาง
    // ==================================

    const tripMap = {};

    regularRows.slice(1).forEach(row => {
      const id = String(row[0] || '').trim();

      if (!id) return;

      tripMap[id] = {
        name: String(row[3] || ''),
        direction: String(row[4] || ''),
        start: String(row[5] || ''),
        destination: String(row[6] || ''),
        departure: formatTime(row[7]),
        arrival: formatTime(row[8]),
        vehicleId: String(row[9] || '').trim(),
        driverName: String(row[10] || '').trim()
      };
    });

    specialRows.slice(1).forEach(row => {
      const id = String(row[0] || '').trim();

      if (!id) return;

      tripMap[id] = {
        name: String(row[2] || ''),
        direction: String(row[3] || ''),
        start: String(row[4] || ''),
        destination: String(row[5] || ''),
        departure: formatTime(row[6]),
        arrival: formatTime(row[7]),
        vehicleId: String(row[8] || '').trim(),
        driverName: String(row[9] || '').trim()
      };
    });

    // ==================================
    // 2. สร้างข้อมูลรถ
    // ==================================

    const vehicleMap = {};

    vehicleRows.slice(1).forEach(row => {
      const vehicleId = String(row[0] || '').trim();

      if (!vehicleId) return;

      vehicleMap[vehicleId] = {
        plate: String(row[1] || ''),
        driver: String(row[2] || ''),
        capacity: Number(row[3]) || 0
      };
    });

    // ==================================
    // 3. สร้างข้อมูลการมอบหมายรถ
    // ==================================

    const assignmentMap = {};

    assignmentRows.slice(1).forEach(row => {
      if (String(row[8]).trim() !== 'เปิดใช้งาน') return;

      const date = formatDate(row[1]);
      const tripId = String(row[2] || '').trim();

      if (!date || !tripId) return;

      const key = date + '|' + tripId;

      assignmentMap[key] = {
        vehicleId: String(row[4] || '').trim(),
        driverName: String(row[5] || '').trim(),
        departure: formatTime(row[6]),
        arrival: formatTime(row[7])
      };
    });

    // ==================================
    // 4. สร้างกลุ่มรอบรถทั้งหมด
    // ==================================

    const tripGroups = {};

    function addTrip(date, tripId, base) {
      const key = date + '|' + tripId;

      const assignment = assignmentMap[key] || {};

      const vehicleId =
        assignment.vehicleId || base.vehicleId || '';

      const driverName =
        assignment.driverName || base.driverName || '';

      const vehicle = vehicleMap[vehicleId] || {};

      tripGroups[key] = {
        date: date,
        tripId: tripId,
        tripName: base.name || '',
        direction: base.direction || '',
        start: base.start || '',
        destination: base.destination || '',
        departure:
          assignment.departure || base.departure || '',
        arrival:
          assignment.arrival || base.arrival || '',
        vehicleId: vehicleId,
        plate: vehicle.plate || '',
        driverName: driverName || vehicle.driver || '',
        capacity: vehicle.capacity || 0,
        passengerCount: 0,
        passengers: []
      };
    }

    // ==================================
    // 5. สร้างรอบประจำล่วงหน้า 30 วัน
    // ==================================

    for (let i = 0; i < 30; i++) {
      const dateObject = new Date();
      dateObject.setDate(dateObject.getDate() + i);

      const date = Utilities.formatDate(
        dateObject, timezone, 'yyyy-MM-dd'
      );

      // วันประจำสัปดาห์ตาม JavaScript:
      // 0 = อาทิตย์, 1 = จันทร์, ... 6 = เสาร์
      const isoDay = Number(
        Utilities.formatDate(dateObject, timezone, 'u')
      );

      const weekday = isoDay % 7;

      regularRows.slice(1).forEach(row => {
        const tripId = String(row[0] || '').trim();

        if (!tripId) return;

        if (String(row[11]).trim() !== 'เปิดใช้งาน') {
          return;
        }

        if (Number(row[1]) !== weekday) {
          return;
        }

        const base = {
          name: String(row[3] || ''),
          direction: String(row[4] || ''),
          start: String(row[5] || ''),
          destination: String(row[6] || ''),
          departure: formatTime(row[7]),
          arrival: formatTime(row[8]),
          vehicleId: String(row[9] || '').trim(),
          driverName: String(row[10] || '').trim()
        };

        addTrip(date, tripId, base);
      });
    }

    // ==================================
    // 6. เพิ่มรอบพิเศษตามวันที่กำหนด
    // ==================================

    specialRows.slice(1).forEach(row => {
      const tripId = String(row[0] || '').trim();
      const date = formatDate(row[1]);

      if (!tripId || !date) return;

      if (String(row[10]).trim() !== 'เปิดใช้งาน') {
        return;
      }

      const base = {
        name: String(row[2] || ''),
        direction: String(row[3] || ''),
        start: String(row[4] || ''),
        destination: String(row[5] || ''),
        departure: formatTime(row[6]),
        arrival: formatTime(row[7]),
        vehicleId: String(row[8] || '').trim(),
        driverName: String(row[9] || '').trim()
      };

      const key = date + '|' + tripId;

      // ถ้ามีรอบประจำในวันเดียวกันและรหัสเดียวกัน
      // ให้ข้อมูลรอบพิเศษเป็นข้อมูลล่าสุด
      addTrip(date, tripId, base);
    });

    // ==================================
    // 7. นำข้อมูลการจองมาใส่ในแต่ละรอบ
    // ==================================

    bookings.forEach(row => {
      if (String(row[11]).trim() === 'ยกเลิก') return;

      const date = formatDate(row[5]);
      const tripId = String(row[6] || '').trim();

      if (!date || !tripId) return;

      const key = date + '|' + tripId;

      // รองรับการจองที่อยู่นอกช่วง 30 วัน
      if (!tripGroups[key]) {
        const base = tripMap[tripId] || {};

        addTrip(date, tripId, base);
      }

      const group = tripGroups[key];

      const count = Number(row[10]) || 0;

      group.passengerCount += count;

      group.passengers.push({
        bookingId: String(row[1] || ''),
        bookingType: String(row[2] || ''),
        fullName: String(row[3] || ''),
        phone: String(row[4] || ''),
        pickup: String(row[8] || ''),
        dropoff: String(row[9] || ''),
        passengerCount: count,
        status: String(row[11] || ''),
        note: String(row[12] || '')
      });
    });

    // ==================================
    // 8. เรียงลำดับรอบรถตามวันที่และเวลา
    // ==================================

    const trips = Object.values(tripGroups).sort((a, b) => {
      const dateCompare = a.date.localeCompare(b.date);

      if (dateCompare !== 0) return dateCompare;

      return a.departure.localeCompare(b.departure);
    });

    return {
      success: true,
      driverName: session.driverName,
      vehicleId: session.vehicleId,
      totalTrips: trips.length,
      totalPassengers: trips.reduce(
        (sum, trip) => sum + trip.passengerCount, 0
      ),
      trips: trips
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}



// ออกจากระบบคนขับ
function logoutDriver(token) {
  if (!token || typeof token !== 'string') {
    return { success: false };
  }

  CacheService.getScriptCache().remove(
    'DRIVER_SESSION_' + token
  );

  return { success: true };
}

// ==========================================
// สร้างบัญชีคนขับโดยผู้ดูแลระบบ
// ==========================================

function createDriverAccount(token, data) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    // ตรวจสอบสิทธิ์ผู้ดูแล
    const admin = checkAdminSession(token);

    if (!admin.success) {
      throw new Error('กรุณาเข้าสู่ระบบผู้ดูแลก่อน');
    }

    const username = String(data.username || '').trim();
    const password = String(data.password || '');
    const driverName = String(data.driverName || '').trim();
    const vehicleId = String(data.vehicleId || '').trim();

    if (!username || !password || !driverName || !vehicleId) {
      throw new Error('กรุณากรอกข้อมูลให้ครบทุกช่อง');
    }

    if (!/^[a-zA-Z0-9_]{4,30}$/.test(username)) {
      throw new Error('ชื่อผู้ใช้ต้องมี 4-30 ตัว ใช้ภาษาอังกฤษ ตัวเลข หรือ _');
    }

    if (password.length < 12 || password.length > 128) {
      throw new Error('รหัสผ่านต้องมีความยาว 12-128 ตัวอักษร');
    }

    if (driverName.length > 100) {
      throw new Error('ชื่อคนขับยาวเกินกำหนด');
    }

    const props = PropertiesService.getScriptProperties();
    const pepper = props.getProperty('ADMIN_PASSWORD_PEPPER');

    if (!pepper) {
      throw new Error('ระบบรหัสผ่านยังไม่พร้อม กรุณาตรวจสอบบัญชีผู้ดูแล');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const accountSheet = ss.getSheetByName('บัญชีคนขับ');
    const vehicleSheet = ss.getSheetByName('ข้อมูลรถ');

    if (!accountSheet || !vehicleSheet) {
      throw new Error('ไม่พบชีตบัญชีคนขับหรือข้อมูลรถ');
    }

    // ตรวจสอบชื่อผู้ใช้ซ้ำ
    const accounts = accountSheet.getDataRange().getValues();

    const duplicate = accounts.slice(1).some(row =>
      String(row[0]).trim().toLowerCase() === username.toLowerCase()
    );

    if (duplicate) {
      throw new Error('ชื่อผู้ใช้นี้มีอยู่ในระบบแล้ว');
    }

    // ตรวจสอบรหัสรถ
    const vehicles = vehicleSheet.getDataRange().getValues();

    const vehicle = vehicles.slice(1).find(row =>
      String(row[0]).trim() === vehicleId &&
      String(row[4]).trim() === 'เปิดใช้งาน'
    );

    if (!vehicle) {
      throw new Error('ไม่พบรถที่เปิดใช้งาน กรุณาตรวจสอบรหัสรถ');
    }

    // สร้าง Salt และ Hash
    const salt = Utilities.getUuid() + Utilities.getUuid();

    const passwordHash = hashAdminPassword_(
      password,
      salt,
      pepper
    );

    // บันทึกบัญชีคนขับ
    accountSheet.appendRow([
      username,
      passwordHash,
      salt,
      driverName,
      vehicleId,
      'เปิดใช้งาน',
      new Date(),
      ''
    ]);

    return {
      success: true,
      message: 'สร้างบัญชีคนขับสำเร็จ',
      username: username,
      driverName: driverName,
      vehicleId: vehicleId
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// ==========================================
// ดึงข้อมูลบัญชีคนขับและรถ สำหรับผู้ดูแล
// ==========================================

function getDriverManagementData(token) {
  try {
    const admin = checkAdminSession(token);

    if (!admin || !admin.success) {
      throw new Error('ไม่มีสิทธิ์เข้าถึง กรุณาเข้าสู่ระบบใหม่');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const accountSheet = ss.getSheetByName('บัญชีคนขับ');
    const vehicleSheet = ss.getSheetByName('ข้อมูลรถ');

    if (!accountSheet || !vehicleSheet) {
      throw new Error('ไม่พบชีตบัญชีคนขับหรือข้อมูลรถ');
    }

    const accountRows = accountSheet.getDataRange().getValues();
    const vehicleRows = vehicleSheet.getDataRange().getValues();

    const accounts = accountRows.slice(1)
      .filter(row => row[0])
      .map(row => ({
        username: String(row[0]),
        driverName: String(row[3]),
        vehicleId: String(row[4]),
        status: String(row[5])
      }));

    const vehicles = vehicleRows.slice(1)
      .filter(row =>
        row[0] &&
        String(row[4]).trim() === 'เปิดใช้งาน'
      )
      .map(row => ({
        vehicleId: String(row[0]),
        plate: String(row[1])
      }));

    return {
      success: true,
      accounts: accounts,
      vehicles: vehicles
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}
// คืนค่า URL ของ Web App
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}


// ==========================================
// จัดการข้อมูลรถ
// ==========================================

// โหลดรายการรถทั้งหมด
function getVehicleManagementData(token) {
  try {
    const admin = checkAdminSession(token);

    if (!admin || !admin.success) {
      throw new Error('ไม่มีสิทธิ์เข้าถึง กรุณาเข้าสู่ระบบใหม่');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ข้อมูลรถ');

    if (!sheet) {
      throw new Error('ไม่พบชีตข้อมูลรถ');
    }

    const rows = sheet.getDataRange().getValues();

    const vehicles = rows.slice(1)
      .filter(row => row[0])
      .map(row => ({
        vehicleId: String(row[0]),
        plate: String(row[1]),
        driverName: String(row[2]),
        seats: Number(row[3]),
        status: String(row[4])
      }));

    return {
      success: true,
      vehicles: vehicles
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}


// เพิ่มรถใหม่
function createVehicle(token, data) {
  try {
    const admin = checkAdminSession(token);

    if (!admin || !admin.success) {
      throw new Error('ไม่มีสิทธิ์เข้าถึง กรุณาเข้าสู่ระบบใหม่');
    }

    if (!data) {
      throw new Error('ไม่พบข้อมูลรถ');
    }

    const vehicleId = String(data.vehicleId || '').trim();
    const plate = String(data.plate || '').trim();
    const driverName = String(data.driverName || '').trim();
    const seats = Number(data.seats);

    if (!vehicleId || !plate || !driverName) {
      throw new Error('กรุณากรอกข้อมูลรถให้ครบ');
    }

    if (!Number.isInteger(seats) || seats < 1) {
      throw new Error('จำนวนที่นั่งต้องเป็นจำนวนเต็มตั้งแต่ 1 ขึ้นไป');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ข้อมูลรถ');

    if (!sheet) {
      throw new Error('ไม่พบชีตข้อมูลรถ');
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const rows = sheet.getDataRange().getValues();

      const duplicateId = rows.slice(1).some(row =>
        String(row[0]).trim().toLowerCase() === vehicleId.toLowerCase()
      );

      if (duplicateId) {
        throw new Error('รหัสรถนี้มีอยู่ในระบบแล้ว');
      }

      const duplicatePlate = rows.slice(1).some(row =>
        String(row[1]).trim().toLowerCase() === plate.toLowerCase()
      );

      if (duplicatePlate) {
        throw new Error('ทะเบียนรถนี้มีอยู่ในระบบแล้ว');
      }

      sheet.appendRow([
        vehicleId,
        plate,
        driverName,
        seats,
        'เปิดใช้งาน'
      ]);

    } finally {
      lock.releaseLock();
    }

    return {
      success: true,
      message: 'เพิ่มข้อมูลรถเรียบร้อยแล้ว'
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

function archiveBookingRows_(rows) {
  if (!rows || rows.length === 0) {
    return 0;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookingSheet = ss.getSheetByName('ข้อมูลการจอง');
  const archiveSheet = ss.getSheetByName('ประวัติการจอง');

  if (!bookingSheet || !archiveSheet) {
    throw new Error('ไม่พบชีตข้อมูลการจองหรือประวัติการจอง');
  }

  const existingIds = new Set();

  if (archiveSheet.getLastRow() >= 2) {
    const ids = archiveSheet
      .getRange(2, 2, archiveSheet.getLastRow() - 1, 1)
      .getValues();

    ids.forEach(row => {
      if (row[0]) {
        existingIds.add(String(row[0]));
      }
    });
  }

  const newRows = [];

  rows.forEach(row => {
    if (!row || row.length !== 13) {
      throw new Error('รูปแบบข้อมูลการจองไม่ครบ 13 คอลัมน์');
    }

    const bookingId = String(row[1] || '').trim();

    if (!bookingId) {
      throw new Error('พบรายการที่ไม่มีรหัสการจอง');
    }

    if (!existingIds.has(bookingId)) {
      newRows.push(row);
      existingIds.add(bookingId);
    }
  });

  if (newRows.length > 0) {
    const startRow = archiveSheet.getLastRow() + 1;

    archiveSheet
      .getRange(startRow, 1, newRows.length, 13)
      .setValues(newRows);
  }

  return newRows.length;
}

function cleanupOldBookings() {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bookingSheet = ss.getSheetByName('ข้อมูลการจอง');
    const archiveSheet = ss.getSheetByName('ประวัติการจอง');

    if (!bookingSheet || !archiveSheet) {
      throw new Error('ไม่พบชีตข้อมูลการจองหรือประวัติการจอง');
    }

    if (bookingSheet.getLastRow() < 2) {
      return 'ไม่มีข้อมูลการจองให้ตรวจสอบ';
    }

    // คำนวณวันที่ย้อนหลัง 7 วัน
    const timeZone = ss.getSpreadsheetTimeZone();
    const todayText = Utilities.formatDate(
      new Date(), timeZone, 'yyyy-MM-dd'
    );

    const parts = todayText.split('-').map(Number);
    const cutoffDate = new Date(
      Date.UTC(parts[0], parts[1] - 1, parts[2] - 7)
    );

    const cutoffText = Utilities.formatDate(
      cutoffDate, 'UTC', 'yyyy-MM-dd'
    );

    const lastRow = bookingSheet.getLastRow();

    const rows = bookingSheet.getRange(
      2, 1, lastRow - 1, 13
    ).getValues();

    const rowsToArchive = [];
    const sheetRowsToDelete = [];

    rows.forEach((row, index) => {
      const travelDate = row[5];

      if (!(travelDate instanceof Date) || isNaN(travelDate.getTime())) {
        return;
      }

      const travelDateText = Utilities.formatDate(
        travelDate, timeZone, 'yyyy-MM-dd'
      );

      if (travelDateText <= cutoffText) {
        rowsToArchive.push(row);
        sheetRowsToDelete.push(index + 2);
      }
    });

    if (rowsToArchive.length === 0) {
      return 'ยังไม่มีข้อมูลที่ครบกำหนดลบ 7 วัน';
    }

    // สำรองข้อมูลก่อนลบ
    archiveBookingRows_(rowsToArchive);

    // ลบจากล่างขึ้นบน เพื่อไม่ให้เลขแถวคลาดเคลื่อน
    sheetRowsToDelete.reverse().forEach(rowNumber => {
      bookingSheet.deleteRow(rowNumber);
    });

    return 'สำรองและลบข้อมูลเก่าเรียบร้อย จำนวน ' +
      rowsToArchive.length + ' รายการ';

  } catch (error) {
    throw new Error('ล้างข้อมูลไม่สำเร็จ: ' + error.message);

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function installBookingCleanupTrigger() {
  // ลบ Trigger เก่าของฟังก์ชันนี้ก่อน ป้องกันการทำงานซ้ำ
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (
      trigger.getHandlerFunction() === 'cleanupOldBookings'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // ตั้งให้ตรวจสอบทุกวัน เวลาประมาณ 02:00 น.
  ScriptApp.newTrigger('cleanupOldBookings')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  return 'ตั้งเวลาตรวจสอบข้อมูลเก่าสำเร็จ';
}

function deleteOldBookingsManually(token, travelDate, tripId) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    // ตรวจสอบสิทธิ์ผู้ใช้งาน
    const admin = checkAdminSession(token);
    const driver = admin.success
      ? null
      : checkDriverSession(token);

    if (!admin.success && (!driver || !driver.success)) {
      throw new Error('หมดเวลาเข้าสู่ระบบ กรุณาเข้าสู่ระบบใหม่');
    }

    // ตรวจสอบรูปแบบวันที่
    const dateText = String(travelDate || '').trim();

    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText)) {
      throw new Error('รูปแบบวันที่ไม่ถูกต้อง');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timeZone = ss.getSpreadsheetTimeZone();

    const selectedDate = new Date(dateText + 'T12:00:00');

    if (
      isNaN(selectedDate.getTime()) ||
      Utilities.formatDate(selectedDate, timeZone, 'yyyy-MM-dd') !== dateText
    ) {
      throw new Error('วันที่ที่เลือกไม่ถูกต้อง');
    }

    const todayText = Utilities.formatDate(
      new Date(), timeZone, 'yyyy-MM-dd'
    );

    // ห้ามลบข้อมูลของวันนี้หรือวันในอนาคต
    if (dateText >= todayText) {
      throw new Error('ลบได้เฉพาะข้อมูลที่วันเดินทางผ่านไปแล้วอย่างน้อย 1 วัน');
    }

    const selectedTripId = String(tripId || '').trim();

    // คนขับต้องระบุรอบ และต้องเป็นรอบที่ได้รับมอบหมาย
    if (!admin.success) {
      if (!selectedTripId) {
        throw new Error('คนขับต้องเลือกรอบรถที่ตนเองได้รับมอบหมาย');
      }

      const assignmentSheet = ss.getSheetByName('มอบหมายรถ');

      if (!assignmentSheet || assignmentSheet.getLastRow() < 2) {
        throw new Error('ไม่พบข้อมูลการมอบหมายรถ');
      }

      const assignments = assignmentSheet.getRange(
        2, 1, assignmentSheet.getLastRow() - 1, 10
      ).getValues();

      const isAssigned = assignments.some(row => {
        if (!(row[1] instanceof Date)) return false;

        const assignmentDate = Utilities.formatDate(
          row[1], timeZone, 'yyyy-MM-dd'
        );

        return (
          assignmentDate === dateText &&
          String(row[2]).trim() === selectedTripId &&
          String(row[4]).trim() === String(driver.vehicleId).trim() &&
          String(row[5]).trim() === String(driver.driverName).trim() &&
          String(row[8]).trim() === 'เปิดใช้งาน'
        );
      });

      if (!isAssigned) {
        throw new Error('คุณไม่มีสิทธิ์ลบข้อมูลของรอบรถนี้');
      }
    }

    const bookingSheet = ss.getSheetByName('ข้อมูลการจอง');
    const archiveSheet = ss.getSheetByName('ประวัติการจอง');

    if (!bookingSheet || !archiveSheet) {
      throw new Error('ไม่พบชีตข้อมูลการจองหรือประวัติการจอง');
    }

    if (bookingSheet.getLastRow() < 2) {
      return {
        success: true,
        count: 0,
        message: 'ไม่มีข้อมูลการจองในวันที่เลือก'
      };
    }

    const rows = bookingSheet.getRange(
      2, 1, bookingSheet.getLastRow() - 1, 13
    ).getValues();

    const rowsToArchive = [];
    const rowsToDelete = [];

    rows.forEach((row, index) => {
      if (!(row[5] instanceof Date)) return;

      const bookingDate = Utilities.formatDate(
        row[5], timeZone, 'yyyy-MM-dd'
      );

      const sameDate = bookingDate === dateText;
      const sameTrip = !selectedTripId ||
        String(row[6]).trim() === selectedTripId;

      if (sameDate && sameTrip) {
        rowsToArchive.push(row);
        rowsToDelete.push(index + 2);
      }
    });

    if (rowsToDelete.length === 0) {
      return {
        success: true,
        count: 0,
        message: 'ไม่พบข้อมูลการจองที่ตรงกับเงื่อนไข'
      };
    }

    // สำรองข้อมูลก่อนลบ
    archiveBookingRows_(rowsToArchive);

    // ลบจากล่างขึ้นบน
    rowsToDelete.reverse().forEach(rowNumber => {
      bookingSheet.deleteRow(rowNumber);
    });

    return {
      success: true,
      count: rowsToDelete.length,
      message: 'สำรองและลบข้อมูลเรียบร้อย'
    };

  } catch (error) {
    return {
      success: false,
      count: 0,
      message: error.message
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
// ==========================================
// DRIVER ASSIGNMENT : ดึงข้อมูลการมอบหมายรถ
// ==========================================

function getDriverAssignmentData(token) {
  try {
    const session = checkDriverSession(token);

    if (!session.success) {
      throw new Error('กรุณาเข้าสู่ระบบคนขับก่อน');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timezone = 'Asia/Vientiane';

    const assignmentSheet = ss.getSheetByName('มอบหมายรถ');
    const vehicleSheet = ss.getSheetByName('ข้อมูลรถ');

    if (!assignmentSheet || !vehicleSheet) {
      throw new Error('ไม่พบชีตข้อมูลการมอบหมายหรือข้อมูลรถ');
    }

    function formatDate(value) {
      if (value instanceof Date) {
        return Utilities.formatDate(
          value, timezone, 'yyyy-MM-dd'
        );
      }

      return String(value || '').trim();
    }

    function formatTime(value) {
      if (value instanceof Date) {
        return Utilities.formatDate(
          value, timezone, 'HH:mm'
        );
      }

      return String(value || '').trim();
    }

    // วันที่วันนี้ถึงอีก 7 วัน
    const todayText = Utilities.formatDate(
      new Date(), timezone, 'yyyy-MM-dd'
    );

    const parts = todayText.split('-').map(Number);

    const today = new Date(Date.UTC(
      parts[0], parts[1] - 1, parts[2]
    ));

    const endDate = new Date(today);
    endDate.setUTCDate(endDate.getUTCDate() + 7);

    function toDateNumber(dateText) {
      const p = dateText.split('-').map(Number);

      if (p.length !== 3 || p.some(isNaN)) {
        return null;
      }

      return Date.UTC(p[0], p[1] - 1, p[2]);
    }

    const startNumber = today.getTime();
    const endNumber = endDate.getTime();

    // ==================================
    // 1. ดึงรถที่เปิดใช้งาน
    // ==================================

    const vehicleRows = vehicleSheet.getDataRange().getValues();

    const vehicles = vehicleRows.slice(1)
      .filter(row =>
        row[0] &&
        String(row[4]).trim() === 'เปิดใช้งาน'
      )
      .map(row => ({
        vehicleId: String(row[0]).trim(),
        plate: String(row[1] || '').trim(),
        capacity: Number(row[3]) || 0
      }));

    // ==================================
    // 2. ดึงรายการมอบหมายรถ
    // ==================================

    const assignmentRows = assignmentSheet.getDataRange().getValues();

    const assignments = assignmentRows.slice(1)
      .map(row => {
        const date = formatDate(row[1]);

        return {
          assignmentId: String(row[0] || '').trim(),
          date: date,
          tripId: String(row[2] || '').trim(),
          tripName: String(row[3] || '').trim(),
          vehicleId: String(row[4] || '').trim(),
          driverName: String(row[5] || '').trim(),
          departure: formatTime(row[6]),
          arrival: formatTime(row[7]),
          status: String(row[8] || '').trim(),
          note: String(row[9] || '').trim()
        };
      })
      .filter(item => {
        const dateNumber = toDateNumber(item.date);

        return (
          item.assignmentId &&
          dateNumber !== null &&
          dateNumber >= startNumber &&
          dateNumber <= endNumber &&
          item.status !== 'ยกเลิก'
        );
      });

    // ==================================
    // 3. แยกรอบที่คนขับรับผิดชอบ
    // ==================================

    const myAssignments = assignments.filter(item =>
  String(item.driverName).trim() ===
  String(session.driverName).trim()
);

    return {
      success: true,
      driverName: session.driverName,
      today: todayText,
      vehicles: vehicles,
      assignments: assignments,
      myAssignments: myAssignments
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}
// ==========================================
// บันทึกการมอบหมายรถโดยคนขับ
// ==========================================

function saveDriverAssignment(token, data) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const session = checkDriverSession(token);

    if (!session.success) {
      throw new Error('กรุณาเข้าสู่ระบบคนขับก่อน');
    }

    if (!data) {
      throw new Error('ไม่พบข้อมูลการมอบหมาย');
    }

    const date = String(data.date || '').trim();
    const tripId = String(data.tripId || '').trim();
    const vehicleId = String(data.vehicleId || '').trim();
    const departure = String(data.departure || '').trim();
    const arrival = String(data.arrival || '').trim();
    const note = String(data.note || '').trim();

    if (!date || !tripId || !vehicleId ||
        !departure || !arrival) {
      throw new Error('กรุณากรอกข้อมูลให้ครบทุกช่อง');
    }

    // ตรวจสอบวันที่
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      throw new Error('รูปแบบวันที่ไม่ถูกต้อง');
    }

    const parts = date.split('-').map(Number);
    const selectedDate = new Date(
      parts[0], parts[1] - 1, parts[2]
    );

    if (
      selectedDate.getFullYear() !== parts[0] ||
      selectedDate.getMonth() !== parts[1] - 1 ||
      selectedDate.getDate() !== parts[2]
    ) {
      throw new Error('วันที่เดินทางไม่ถูกต้อง');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timezone = 'Asia/Vientiane';

    const todayText = Utilities.formatDate(
      new Date(), timezone, 'yyyy-MM-dd'
    );

    const todayParts = todayText.split('-').map(Number);
    const today = new Date(Date.UTC(
      todayParts[0], todayParts[1] - 1, todayParts[2]
    ));

    const selected = new Date(Date.UTC(
      parts[0], parts[1] - 1, parts[2]
    ));

    const dayDifference =
      Math.round((selected.getTime() - today.getTime()) / 86400000);

    if (dayDifference < 0 || dayDifference > 7) {
      throw new Error('เลือกวันที่ได้ตั้งแต่วันนี้ถึง 7 วันข้างหน้าเท่านั้น');
    }

    // ตรวจสอบเวลา
    function timeToMinutes(value) {
      const match = value.match(/^([01]\d|2[0-3]):([0-5]\d)$/);

      if (!match) return null;

      return Number(match[1]) * 60 + Number(match[2]);
    }

    const startMinutes = timeToMinutes(departure);
    const endMinutes = timeToMinutes(arrival);

    if (startMinutes === null || endMinutes === null) {
      throw new Error('รูปแบบเวลาไม่ถูกต้อง กรุณาใช้ HH:mm');
    }

    if (endMinutes <= startMinutes) {
      throw new Error('เวลาถึงต้องอยู่หลังเวลาออกเดินทาง');
    }

    if (note.length > 500) {
      throw new Error('หมายเหตุยาวเกิน 500 ตัวอักษร');
    }

    const tripSheet = ss.getSheetByName('รอบเดินรถ');
    const specialSheet = ss.getSheetByName('รอบพิเศษ');
    const vehicleSheet = ss.getSheetByName('ข้อมูลรถ');
    const assignmentSheet = ss.getSheetByName('มอบหมายรถ');

    if (!tripSheet || !specialSheet ||
        !vehicleSheet || !assignmentSheet) {
      throw new Error('ไม่พบชีตข้อมูลที่จำเป็น');
    }

    // ตรวจสอบรอบรถ
    const availableResult = getAvailableTrips(date);

    if (!availableResult.success) {
      throw new Error(availableResult.message);
    }

    const trip = availableResult.trips.find(item =>
      item.id === tripId
    );

    if (!trip) {
      throw new Error('ไม่พบรอบรถที่เลือก หรือรอบนี้ไม่ได้เปิดใช้งาน');
    }

    // ตรวจสอบรถ
    const vehicleRows = vehicleSheet.getDataRange().getValues();

    const vehicle = vehicleRows.slice(1).find(row =>
      String(row[0]).trim() === vehicleId &&
      String(row[4]).trim() === 'เปิดใช้งาน'
    );

    if (!vehicle) {
      throw new Error('รถคันนี้ไม่ได้เปิดใช้งาน');
    }

    // อ่านรายการมอบหมายเดิม
    const lastRow = assignmentSheet.getLastRow();

    const rows = lastRow >= 2
      ? assignmentSheet.getRange(
          2, 1, lastRow - 1, 10
        ).getValues()
      : [];

    function formatSavedDate(value) {
      if (value instanceof Date) {
        return Utilities.formatDate(
          value, timezone, 'yyyy-MM-dd'
        );
      }

      return String(value || '').trim();
    }

    function savedTimeToMinutes(value) {
      if (value instanceof Date) {
        return Number(
          Utilities.formatDate(value, timezone, 'H')
        ) * 60 + Number(
          Utilities.formatDate(value, timezone, 'm')
        );
      }

      return timeToMinutes(String(value || '').trim());
    }

    // ตรวจสอบรอบเดียวกันซ้ำ
    const duplicateTrip = rows.some(row =>
      formatSavedDate(row[1]) === date &&
      String(row[2]).trim() === tripId &&
      String(row[8]).trim() === 'เปิดใช้งาน'
    );

    if (duplicateTrip) {
      throw new Error('รอบนี้มีคนมอบหมายรถไปแล้ว');
    }

    // ตรวจสอบเวลาทับซ้อน
    rows.forEach(row => {
      if (
        formatSavedDate(row[1]) !== date ||
        String(row[8]).trim() !== 'เปิดใช้งาน'
      ) {
        return;
      }

      const oldVehicle = String(row[4]).trim();
      const oldDriver = String(row[5]).trim();

      const oldStart = savedTimeToMinutes(row[6]);
      const oldEnd = savedTimeToMinutes(row[7]);

      if (oldStart === null || oldEnd === null) {
        throw new Error(
          'พบข้อมูลเวลาเดิมไม่ถูกต้องในชีตมอบหมายรถ กรุณาให้ผู้ดูแลตรวจสอบ'
        );
      }

      const overlap =
        startMinutes < oldEnd &&
        oldStart < endMinutes;

      if (!overlap) return;

      if (oldVehicle === vehicleId) {
        throw new Error('รถคันนี้มีรอบเดินรถในช่วงเวลาดังกล่าวแล้ว');
      }

      if (oldDriver === session.driverName) {
        throw new Error('คุณมีรอบเดินรถอื่นในช่วงเวลานี้แล้ว');
      }
    });

    // สร้างรหัสรายการ
    const assignmentId =
      'AS-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();

    // บันทึกข้อมูล
    assignmentSheet.appendRow([
      assignmentId,
      selectedDate,
      tripId,
      trip.name,
      vehicleId,
      session.driverName,
      departure,
      arrival,
      'เปิดใช้งาน',
      note
    ]);

    return {
      success: true,
      message: 'มอบหมายรถสำเร็จ',
      assignmentId: assignmentId
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// ==========================================
// แก้ไขการมอบหมายรถโดยคนขับ
// ==========================================

function updateDriverAssignment(token, data) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const session = checkDriverSession(token);

    if (!session.success) {
      throw new Error('กรุณาเข้าสู่ระบบคนขับก่อน');
    }

    if (!data || !data.assignmentId) {
      throw new Error('ไม่พบรหัสรายการมอบหมาย');
    }

    const assignmentId = String(data.assignmentId).trim();
    const date = String(data.date || '').trim();
    const tripId = String(data.tripId || '').trim();
    const vehicleId = String(data.vehicleId || '').trim();
    const departure = String(data.departure || '').trim();
    const arrival = String(data.arrival || '').trim();
    const note = String(data.note || '').trim();

    if (!date || !tripId || !vehicleId ||
        !departure || !arrival) {
      throw new Error('กรุณากรอกข้อมูลให้ครบทุกช่อง');
    }

    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      throw new Error('รูปแบบวันที่ไม่ถูกต้อง');
    }

    const parts = date.split('-').map(Number);
    const selectedDate = new Date(
      parts[0], parts[1] - 1, parts[2]
    );

    if (
      selectedDate.getFullYear() !== parts[0] ||
      selectedDate.getMonth() !== parts[1] - 1 ||
      selectedDate.getDate() !== parts[2]
    ) {
      throw new Error('วันที่เดินทางไม่ถูกต้อง');
    }

    const timezone = 'Asia/Vientiane';

    const todayText = Utilities.formatDate(
      new Date(), timezone, 'yyyy-MM-dd'
    );

    const todayParts = todayText.split('-').map(Number);
    const today = new Date(Date.UTC(
      todayParts[0], todayParts[1] - 1, todayParts[2]
    ));

    const selected = new Date(Date.UTC(
      parts[0], parts[1] - 1, parts[2]
    ));

    const dayDifference = Math.round(
      (selected.getTime() - today.getTime()) / 86400000
    );

    if (dayDifference < 0 || dayDifference > 7) {
      throw new Error('เลือกวันที่ได้ตั้งแต่วันนี้ถึง 7 วันข้างหน้าเท่านั้น');
    }

    function timeToMinutes(value) {
      const match = value.match(/^([01]\d|2[0-3]):([0-5]\d)$/);

      if (!match) return null;

      return Number(match[1]) * 60 + Number(match[2]);
    }

    const startMinutes = timeToMinutes(departure);
    const endMinutes = timeToMinutes(arrival);

    if (startMinutes === null || endMinutes === null) {
      throw new Error('รูปแบบเวลาไม่ถูกต้อง กรุณาใช้ HH:mm');
    }

    if (endMinutes <= startMinutes) {
      throw new Error('เวลาถึงต้องอยู่หลังเวลาออกเดินทาง');
    }

    if (note.length > 500) {
      throw new Error('หมายเหตุยาวเกิน 500 ตัวอักษร');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const assignmentSheet = ss.getSheetByName('มอบหมายรถ');
    const vehicleSheet = ss.getSheetByName('ข้อมูลรถ');

    if (!assignmentSheet || !vehicleSheet) {
      throw new Error('ไม่พบชีตข้อมูลการมอบหมายหรือข้อมูลรถ');
    }

    const lastRow = assignmentSheet.getLastRow();

    if (lastRow < 2) {
      throw new Error('ไม่พบรายการมอบหมายรถ');
    }

    const rows = assignmentSheet.getRange(
      2, 1, lastRow - 1, 10
    ).getValues();

    function formatDate(value) {
      if (value instanceof Date) {
        return Utilities.formatDate(
          value, timezone, 'yyyy-MM-dd'
        );
      }

      return String(value || '').trim();
    }

    function savedTimeToMinutes(value) {
      if (value instanceof Date) {
        return Number(
          Utilities.formatDate(value, timezone, 'H')
        ) * 60 + Number(
          Utilities.formatDate(value, timezone, 'm')
        );
      }

      return timeToMinutes(String(value || '').trim());
    }

    // ค้นหารายการที่ต้องการแก้ไข
    const index = rows.findIndex(row =>
      String(row[0]).trim() === assignmentId
    );

    if (index === -1) {
      throw new Error('ไม่พบรายการมอบหมายรถนี้');
    }

    const oldRow = rows[index];

    // คนขับแก้ไขได้เฉพาะรายการของตัวเอง
    if (String(oldRow[5]).trim() !== session.driverName) {
      throw new Error('คุณไม่มีสิทธิ์แก้ไขรายการของคนขับคนอื่น');
    }

    if (String(oldRow[8]).trim() !== 'เปิดใช้งาน') {
      throw new Error('รายการนี้ไม่ได้อยู่ในสถานะเปิดใช้งาน');
    }

    // ห้ามแก้ไขเมื่อถึงเวลาออกเดินทาง
    const oldDate = formatDate(oldRow[1]);
    const oldDeparture = String(oldRow[6]).trim();

    const departureDateTime = new Date(
      oldDate + 'T' + oldDeparture + ':00+07:00'
    );

    if (
      isNaN(departureDateTime.getTime()) ||
      new Date().getTime() >= departureDateTime.getTime()
    ) {
      throw new Error('เลยเวลาออกเดินทางแล้ว ไม่สามารถแก้ไขได้');
    }

    // ตรวจสอบรอบรถที่เลือก
    const availableResult = getAvailableTrips(date);

    if (!availableResult.success) {
      throw new Error(availableResult.message);
    }

    const trip = availableResult.trips.find(item =>
      item.id === tripId
    );

    if (!trip) {
      throw new Error('ไม่พบรอบรถที่เลือก หรือรอบนี้ไม่ได้เปิดใช้งาน');
    }

    // ตรวจสอบรถที่เลือก
    const vehicleRows = vehicleSheet.getDataRange().getValues();

    const vehicle = vehicleRows.slice(1).find(row =>
      String(row[0]).trim() === vehicleId &&
      String(row[4]).trim() === 'เปิดใช้งาน'
    );

    if (!vehicle) {
      throw new Error('รถคันนี้ไม่ได้เปิดใช้งาน');
    }

    // ตรวจสอบรายการอื่น โดยไม่นับรายการที่กำลังแก้ไข
    rows.forEach((row, i) => {
      if (i === index) return;

      if (
        formatDate(row[1]) !== date ||
        String(row[8]).trim() !== 'เปิดใช้งาน'
      ) {
        return;
      }

      const oldVehicle = String(row[4]).trim();
      const oldDriver = String(row[5]).trim();

      // ตรวจสอบรอบซ้ำ
      if (String(row[2]).trim() === tripId) {
        throw new Error('รอบนี้มีคนมอบหมายรถไปแล้ว');
      }

      const oldStart = savedTimeToMinutes(row[6]);
      const oldEnd = savedTimeToMinutes(row[7]);

      if (oldStart === null || oldEnd === null) {
        throw new Error('พบข้อมูลเวลาเดิมไม่ถูกต้อง กรุณาให้ผู้ดูแลตรวจสอบ');
      }

      const overlap =
        startMinutes < oldEnd &&
        oldStart < endMinutes;

      if (!overlap) return;

      if (oldVehicle === vehicleId) {
        throw new Error('รถคันนี้มีรอบเดินรถในช่วงเวลาดังกล่าวแล้ว');
      }

      if (oldDriver === session.driverName) {
        throw new Error('คุณมีรอบเดินรถอื่นในช่วงเวลานี้แล้ว');
      }
    });

    // อัปเดตข้อมูลเดิม โดยคงรหัสรายการไว้
    assignmentSheet.getRange(index + 2, 1, 1, 10).setValues([[
      assignmentId,
      selectedDate,
      tripId,
      trip.name,
      vehicleId,
      session.driverName,
      departure,
      arrival,
      'เปิดใช้งาน',
      note
    ]]);

    return {
      success: true,
      message: 'แก้ไขการมอบหมายรถสำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function cancelDriverAssignment(token, assignmentId) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const session = checkDriverSession(token);

    if (!session.success) {
      throw new Error('กรุณาเข้าสู่ระบบคนขับก่อน');
    }

    assignmentId = String(assignmentId || '').trim();

    if (!assignmentId) {
      throw new Error('ไม่พบรหัสรายการมอบหมาย');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('มอบหมายรถ');

    if (!sheet) {
      throw new Error('ไม่พบชีตมอบหมายรถ');
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error('ไม่พบรายการมอบหมายรถ');
    }

    const rows = sheet.getRange(
      2, 1, lastRow - 1, 10
    ).getValues();

    const index = rows.findIndex(row =>
      String(row[0]).trim() === assignmentId
    );

    if (index === -1) {
      throw new Error('ไม่พบรายการมอบหมายรถนี้');
    }

    const row = rows[index];

    // ตรวจสอบว่าเป็นรายการของคนขับที่เข้าสู่ระบบ
    if (String(row[5]).trim() !== session.driverName) {
      throw new Error('คุณไม่มีสิทธิ์ยกเลิกรายการของคนขับคนอื่น');
    }

    if (String(row[8]).trim() !== 'เปิดใช้งาน') {
      throw new Error('รายการนี้ไม่ได้อยู่ในสถานะเปิดใช้งาน');
    }

    
    // อนุญาตให้คนขับยกเลิกรอบได้ แม้เลยเวลาออกเดินทางแล้ว
    // ยังคงตรวจสอบสิทธิ์คนขับและสถานะรอบก่อนยกเลิก


    // เปลี่ยนสถานะเป็นยกเลิก โดยไม่ลบข้อมูล
    sheet.getRange(index + 2, 9).setValue('ยกเลิก');

    return {
      success: true,
      message: 'ยกเลิกการมอบหมายรถสำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// ==========================================
// เปลี่ยนสถานะรอบที่เดินทางเสร็จสิ้นอัตโนมัติ
// ==========================================

function updateCompletedDriverAssignments() {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('มอบหมายรถ');

    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }

    const timezone = 'Asia/Vientiane';
    const now = new Date();

    const rows = sheet.getRange(
      2, 1, sheet.getLastRow() - 1, 10
    ).getValues();

    rows.forEach((row, index) => {
      const status = String(row[8] || '').trim();

      // ตรวจเฉพาะรายการที่เปิดใช้งาน
      if (status !== 'เปิดใช้งาน') return;

      let dateText = '';

      if (row[1] instanceof Date) {
        dateText = Utilities.formatDate(
          row[1], timezone, 'yyyy-MM-dd'
        );
      } else {
        dateText = String(row[1] || '').trim();
      }

      let arrivalText = '';

      if (row[7] instanceof Date) {
        arrivalText = Utilities.formatDate(
          row[7], timezone, 'HH:mm'
        );
      } else {
        arrivalText = String(row[7] || '').trim();
      }

      if (
        !/^\d{4}-\d{2}-\d{2}$/.test(dateText) ||
        !/^([01]\d|2[0-3]):([0-5]\d)$/.test(arrivalText)
      ) {
        return;
      }

      const arrivalDateTime = new Date(
        dateText + 'T' + arrivalText + ':00+07:00'
      );

      if (
        !isNaN(arrivalDateTime.getTime()) &&
        now.getTime() >= arrivalDateTime.getTime()
      ) {
        // คอลัมน์ I คือสถานะ
        sheet.getRange(index + 2, 9).setValue('เสร็จสิ้น');
      }
    });

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
// ==================================
// ค้นหาข้อมูลมอบหมายรถตาม ID
// ==================================

function getDriverAssignmentById(token, assignmentId) {

  const session = checkDriverSession(token);

  if (!session || !session.success) {
    return {
      success: false,
      message: "กรุณาเข้าสู่ระบบคนขับก่อน"
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("มอบหมายรถ");

  if (!sheet) {
    return {
      success: false,
      message: "ไม่พบชีตมอบหมายรถ"
    };
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      success: false,
      message: "ไม่พบข้อมูลการมอบหมายรถ"
    };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  for (let i = 0; i < data.length; i++) {

    const row = data[i];

    if (String(row[0]) === String(assignmentId)) {

      // ตรวจสอบว่าเป็นรายการของคนขับที่เข้าสู่ระบบ
      if (String(row[5]) !== String(session.driverName)) {
        return {
          success: false,
          message: "คุณไม่มีสิทธิ์แก้ไขรายการนี้"
        };
      }

      if (row[8] !== "เปิดใช้งาน") {
        return {
          success: false,
          message: "รายการนี้ไม่ได้เปิดใช้งานแล้ว"
        };
      }

      return {
        success: true,
        assignment: {
          assignmentId: row[0],
          date: Utilities.formatDate(
            new Date(row[1]),
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
          ),
          tripId: row[2],
          tripName: row[3],
          vehicleId: row[4],
          driverName: row[5],
          departure: row[6],
          arrival: row[7],
          status: row[8],
          note: row[9]
        }
      };
    }
  }

  return {
    success: false,
    message: "ไม่พบรายการมอบหมายรถที่ต้องการ"
  };
}