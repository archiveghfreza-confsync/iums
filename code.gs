/*******************************************************
 * ConfSync.ir
 * Authentication API
 * Google Apps Script
 *
 * Sheet: Login
 *
 * Row 1 = Headers
 *******************************************************/


/* =====================================================
   CRYPTO WORKER
   ===================================================== */

const CRYPTO_WORKER_URL =
  'https://hooshteb-sec.ghfreza.workers.dev';

/* =====================================================
   ENCRYPT
   ===================================================== */

function encryptdata(text) {

  const response =
    UrlFetchApp.fetch(
      CRYPTO_WORKER_URL + '/encrypt',
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          text: String(text || '')
        })
      }
    );

  return response
    .getContentText()
    .trim();

}


/* =====================================================
   DECRYPT
   ===================================================== */

function decryptdata(cipher) {

  const response =
    UrlFetchApp.fetch(
      CRYPTO_WORKER_URL + '/decrypt',
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          cipher: String(cipher || '').trim()
        })
      }
    );

  return response
    .getContentText()
    .trim();

}


/* =====================================================
   DECRYPT IF NEEDED
   ===================================================== */

function decryptIfNeeded(value) {

  value =
    String(value || '').trim();

  if (value === '') {

    return '';

  }

  return decryptdata(value);

}


/* =====================================================
   HASH DATA
   ===================================================== */

function hashdata(text) {

  const response =
    UrlFetchApp.fetch(
      CRYPTO_WORKER_URL + '/hash',
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          text: String(text || '').trim()
        })
      }
    );

  return response
    .getContentText()
    .trim();

}


/* =====================================================
   PERSIAN DATE
   ===================================================== */

function toPersianDate(date) {

  return new Intl.DateTimeFormat(
    'fa-IR-u-nu-latn',
    {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false,
      timeZone: 'Asia/Tehran'
    }
  ).format(date);

}


/* =====================================================
   LOG
   ===================================================== */

function logToSheet(message) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName('Logs');


  if (!sheet) {

    sheet =
      ss.insertSheet('Logs');

    sheet.appendRow([
      'Date',
      'Message'
    ]);

  }


  sheet.appendRow([
    new Date(),
    String(message || '')
  ]);

}


/* =====================================================
   GET
   ===================================================== */

function doGet(e) {

  return jsonResponse_({

    success: true,

    message:
      'ConfSync Authentication API is running.'

  });

}


/* =====================================================
   POST
   ===================================================== */

function doPost(e) {
  try {
  // ============================================================
    //  لاگ برای دیباگ
    // ============================================================
    logToSheet('=== doPost called ===');
    logToSheet('e.postData.contents:', e.postData.contents);
    logToSheet('Content-Type:', e.postData.type);

    // ============================================================
    //  بررسی درخواست
    // ============================================================
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({
        success: false,
        message: 'Invalid request. No data received.'
      });
    }

    // ============================================================
    //  تشخیص Content-Type و parse کردن داده
    // ============================================================
    const contentType = e.postData?.type || '';
    let data = {};

    // ✅ اگر JSON باشد (حتی با text/plain)
    try {
      data = JSON.parse(e.postData.contents);
      logToSheet('Parsed as JSON successfully:', data);
    } catch (parseError) {
      // ✅ اگر JSON نبود، از e.parameter استفاده کن
      logToSheet('Not JSON, using e.parameter');
      data = e.parameter || {};
    }

    // ============================================================
    //  لاگ action برای دیباگ
    // ============================================================
    logToSheet('data.action:', data.action);
    logToSheet('data.email:', data.email);

    switch (data.action) {

      case 'register':
        return jsonResponse_(registerUser_(data));

      case 'login':
        return jsonResponse_(loginUser_(data));

      case 'changePassword':
        return jsonResponse_(changePassword(data));

      case 'verifyEmail':
        return jsonResponse_(verifyEmail(data.email));

      case 'resendverifyEmail':
        return jsonResponse_(
                    sendVerificationEmail(data.email)
                );

      case 'forgotPassword':
        return jsonResponse_(
                    sendForgotPasswordEmail(data.email)
                );

      default:
        return jsonResponse_({
          success: false,
          message: 'Unknown action.' +data.action
        });
    }

  } catch (error) {
    console.error(error);

    return jsonResponse_({
      success: false,
      message: 'Server error: ' + error.message,
      error: error.message
    });
  }
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =====================================================
   APPEND ROW BY COLUMNS
   ===================================================== */

/**
 * textColumns:
 * آرایه‌ای از شماره ستون‌ها به صورت 0-based
 *
 * مثال:
 *
 * A = 0
 * B = 1
 * C = 2
 */

function appendRowByColumns(
  sheet,
  rowData,
  textColumns
) {

  const row =
    sheet.getLastRow() + 1;


  /*
   * تنظیم Text Format
   */

  textColumns.forEach(function(col) {

    if (
      typeof col === 'number' &&
      col >= 0
    ) {

      sheet
        .getRange(
          row,
          col + 1
        )
        .setNumberFormat('@');

    }

  });


  /*
   * درج کل ردیف
   */

  sheet
    .getRange(
      row,
      1,
      1,
      rowData.length
    )
    .setValues([
      rowData
    ]);


  return row;

}


/* =====================================================
   GET LOGIN SHEET
   ===================================================== */

function getSheet_(sheetname) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet =
    ss.getSheetByName(sheetname);
  if (!sheet) {
    throw new Error(`Sheet ${sheetname} not found.`);
  }
  return sheet;
}


/* =====================================================
   GET COLUMNS
   ===================================================== */

/**
 * خروجی این تابع 0-based است.
 *
 * مثال:
 *
 * A → 0
 * B → 1
 * C → 2
 *
 * برای rowData:
 *
 * rowData[COL.phone]
 *
 * برای getRange:
 *
 * COL.phone + 1
 */

function getLoginColumns_(sheet) {

  const lastColumn =
    sheet.getLastColumn();

  if (lastColumn < 1) {

    throw new Error(
      'Sheet has no columns.'
    );

  }

  const headers =
    sheet
      .getRange(1, 1, 1, lastColumn)
      .getValues()[0];

  const COL = {};


  headers.forEach(function(
    header,
    index
  ) {

    const name =
      String(header || '').trim();


    if (!name) {

      return;

    }


    /*
     * 0-based
     */

    COL[name] =
      index;

  });


  COL.ID =
    COL.ID;

  COL.phone =
    COL.phone;

  COL.email =
    COL.email;

  COL.password =
    COL.password;

  COL.password_hash =
    COL.password_hash;

  COL.fullName =
    COL.fullName;

  COL.university =
    COL.university;

  COL.major =
    COL.major;

  COL.degree =
    COL.degree;

  COL.researchInterest =
    COL.researchInterest;

  COL.registeredAt =
    COL.registeredAt;

  COL.lastLogin =
    COL.lastLogin;

  COL.isActive =
    COL.isActive;


  return COL;

}


/* =====================================================
   COLUMN EXISTS
   ===================================================== */

function columnExists_(COL, name) {

  return (
    typeof COL[name] === 'number' &&
    COL[name] >= 0
  );

}


/* =====================================================
   VALIDATE COLUMNS
   ===================================================== */

function validateLoginColumns_(COL) {

  const required = [

    'ID',
    'phone',
    'email',
    'password_hash',
    'fullName',
    'university',
    'major',
    'degree',
    'researchInterest',
    'registeredAt',
    'lastLogin',
    'isActive'

  ];


  const missing = [];


  required.forEach(function(name) {

    if (
      !columnExists_(
        COL,
        name
      )
    ) {

      missing.push(name);

    }

  });


  if (missing.length > 0) {

    throw new Error(
      'Missing columns: ' +
      missing.join(', ')
    );

  }

}




/* =====================================================
   REGISTER USER
   ===================================================== */

function registerUser_(data) {

  const sheet = getSheet_("Login");

  const COL = getLoginColumns_(sheet);

  validateLoginColumns_(COL);


  /* ---------------------------------------------------
     RECEIVE DATA
  --------------------------------------------------- */
  const phone =
    normalizePhone_(
      data.phone
    );

  const password =
    String(
      data.password || ''
    ).trim();

  const password_Hash = 
    hashdata(password);

  const email =
    String(
      data.email || ''
    )
      .trim()
      .toLowerCase();

  const fullName =
    String(
      data.fullName || ''
    ).trim();


  const university =
    String(
      data.university || ''
    ).trim();


  const major =
    String(
      data.major || ''
    ).trim();


  const degree =
    String(
      data.degree || ''
    ).trim();


  const researchInterest =
    String(
      data.researchInterest || ''
    ).trim();


  /* ---------------------------------------------------
     VALIDATION
  --------------------------------------------------- */

  /*
  if (!phone) {
    return {
      success: false,
      message:
        'شماره موبایل الزامی است.'
    };
  }
*/
/*
  if (!isValidPhone_(phone)) {
    return {
      success: false,
      message:
        'شماره موبایل معتبر نیست.'
    };
  }
*/

  if (!email) {
    return {
      success: false,
      message:
        'Email address is required..'
    };
  }


  if (!isValidEmail_(email)) {
    return {
      success: false,
      message:
        'Invalid email address.'
    };
  }


  if (!password) {
    return {
      success: false,
      message:
        'Password is required.'
    };

  }
  if (password.length < 6) {
    return {
      success: false,
      message:
        'Password must be at least 6 characters.'

    };

  }


  if (!fullName) {

    return {

      success: false,

      message:
        'Full name is required.'

    };

  }


  /* ---------------------------------------------------
     CHECK DUPLICATES
  --------------------------------------------------- */

  const lastRow =
    sheet.getLastRow();


  if (
    lastRow > 1 ) {

    const numberOfRows =
      lastRow - 1;

    /* -----------------------------------------------
       EMAIL
    ----------------------------------------------- */

    const emailValues =
      sheet
        .getRange(
          1 + 1,
          COL.email + 1,
          numberOfRows,
          1
        )
        .getValues();


    for (
      let i = 0;
      i < emailValues.length;
      i++
    ) {

      const existingEmail =
        String(
          emailValues[i][0] || ''
        )
          .trim()
          .toLowerCase();


      if (
        existingEmail &&
        existingEmail === email
      ) {

        return {

          success: false,

          message:
            'This email address is already registered.'

        };

      }

    }

  }


  /* ---------------------------------------------------
     PASSWORD HASH
  --------------------------------------------------- */

  const salt =
    generateSalt_();


  /* ---------------------------------------------------
     USER ID
  --------------------------------------------------- */

  
  const now =
    new Date();


  /* ---------------------------------------------------
     CREATE ROW
  --------------------------------------------------- */

  const rowData =
    new Array(
      sheet.getLastColumn()
    ).fill('');


  /* ---------------------------------------------------
     TEXT VALUES
  --------------------------------------------------- */

  rowData[COL.ID] =
    String(Utilities.getUuid());


  rowData[COL.phone] =
    String(phone);


  rowData[COL.email] =
    String(email);


  /*
   * خام ذخیره نمی‌شود
   */

  if (
    columnExists_(
      COL,
      'password'
    )
  ) {

    rowData[COL.password] = encryptdata(password);

  }


  rowData[COL.password_hash] = String(password_Hash);


  rowData[COL.fullName] =
    String(fullName);


  rowData[COL.university] =
    String(university);


  rowData[COL.major] =
    String(major);


  rowData[COL.degree] =
    String(degree);


  rowData[COL.researchInterest] =
    String(researchInterest);


  /* ---------------------------------------------------
     DATE
  --------------------------------------------------- */

  rowData[COL.registeredAt] =
    now;


  /* ---------------------------------------------------
     LAST LOGIN
  --------------------------------------------------- */

  rowData[COL.lastLogin] =
    '';


  /* ---------------------------------------------------
     ACTIVE
  --------------------------------------------------- */

  rowData[COL.isActive] = "false";


  /* ---------------------------------------------------
     TEXT COLUMNS
  --------------------------------------------------- */

  const textColumns = [

    COL.ID,

    COL.phone,

    COL.email,

    COL.password,

    COL.password_hash,

    COL.fullName,

    COL.university,

    COL.major,

    COL.degree,

    COL.researchInterest,

    COL.lastLogin,
    
    COL.isActive

  ].filter(function(col) {

    return (
      typeof col === 'number' &&
      col >= 0
    );

  });


  /* ---------------------------------------------------
     APPEND
  --------------------------------------------------- */

  const newRow =
    appendRowByColumns(
      sheet,
      rowData,
      textColumns
    );


  /* ---------------------------------------------------
     RESPONSE
  --------------------------------------------------- */
   sendVerificationEmail(email, fullName)

  return {

    success: true,

    message:
      '✅ Registration successful! Verification email has been sent.',

    row:
      newRow,

    user: {

      ID: String(Utilities.getUuid()),

      phone:
        String(phone),

      email:
        String(email),

      fullName:
        String(fullName),

      university:
        String(university),

      major:
        String(major),

      degree:
        String(degree),

      researchInterest:
        String(researchInterest)

    }

  };

}


/* =====================================================
   VALID EMAIL
   ===================================================== */

function isValidEmail_(email) {

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/
    .test(
      String(email || '')
        .trim()
        .toLowerCase()
    );

}


/* =====================================================
   LOGIN USER
   ===================================================== */

function loginUser_(data) {

  const sheet =
    getSheet_("Login");


  const COL =
    getLoginColumns_(sheet);

  validateLoginColumns_(COL);


  /* ---------------------------------------------------
     RECEIVE DATA
  --------------------------------------------------- */

  const email = data.email
  
  const password =
    String(
      data.password || ''
    ).trim();


  /* ---------------------------------------------------
     VALIDATION
  --------------------------------------------------- */

  if (
    !email ||
    !password
  ) {

    return {

      success: false,

      message:
        'Email and password are required.'

    };

  }


  /* ---------------------------------------------------
     FIND USER
  --------------------------------------------------- */

  const lastRow =
    sheet.getLastRow();


  if (
    lastRow <= 1) {

    return {

      success: false,

      message:
        'Invalid email or password(100).'

    };

  }

  const password_hash = hashdata(password)

  const numberOfRows =
    lastRow - 1;


  const emailValues =
    sheet
      .getRange(1 + 1,
        COL.email + 1,
        numberOfRows,
        1
      )
      .getValues();

  const password_hashlValues =
    sheet
      .getRange(1 + 1,
        COL.password_hash + 1,
        numberOfRows,
        1
      )
      .getValues();

  let userRow =
    -1;


  for (
    let i = 0;
    i < emailValues.length;
    i++
  ) {

    const rowEmai = String(emailValues[i][0])
    const rowpassword_hash = String(password_hashlValues[i][0])

    if (
      rowEmai === email && rowpassword_hash === password_hash
    ) {

      userRow =1 + 1 + i;

      break;

    }

  }


  if (userRow === -1) {

    return {

      success: false,

      message:
        'Invalid email or password.(200)'

    };

  }


  /* ---------------------------------------------------
     ACTIVE
  --------------------------------------------------- */

  const isActive =
    sheet
      .getRange(
        userRow,
        COL.isActive + 1
      )
      .getValue();


  if (
    isActive === false ||
    String(isActive).toLowerCase() === 'false'
  ) {

    return {

      success: false,

      message:
        'your account is inactive. An activation email has been sent to you.'

    };

  }


 
  /* ---------------------------------------------------
     UPDATE LAST LOGIN
  --------------------------------------------------- */

  const now =
    new Date();


  sheet
    .getRange(
      userRow,
      COL.lastLogin + 1
    )
    .setNumberFormat('@');


  sheet
    .getRange(
      userRow,
      COL.lastLogin + 1
    )
    .setValue(
      String(now)
    );


  /* ---------------------------------------------------
     READ USER
  --------------------------------------------------- */

  const user = {

    ID:
      String(
        sheet
          .getRange(
            userRow,
            COL.ID + 1
          )
          .getValue()
      ),

    phone:
      String(
        sheet
          .getRange(
            userRow,
            COL.phone + 1
          )
          .getValue()
      ),

    email:
      String(
        sheet
          .getRange(
            userRow,
            COL.email + 1
          )
          .getValue()
      ),

    fullName:
      String(
        sheet
          .getRange(
            userRow,
            COL.fullName + 1
          )
          .getValue()
      ),

    university:
      String(
        sheet
          .getRange(
            userRow,
            COL.university + 1
          )
          .getValue()
      ),

    major:
      String(
        sheet
          .getRange(
            userRow,
            COL.major + 1
          )
          .getValue()
      ),

    degree:
      String(
        sheet
          .getRange(
            userRow,
            COL.degree + 1
          )
          .getValue()
      ),

    researchInterest:
      String(
        sheet
          .getRange(
            userRow,
            COL.researchInterest + 1
          )
          .getValue()
      ),

    registeredAt:
      sheet
        .getRange(
          userRow,
          COL.registeredAt + 1
        )
        .getValue(),

    lastLogin:
      now,

    isActive:
      true

  };


  /* ---------------------------------------------------
     RESPONSE
  --------------------------------------------------- */

  return {

    success: true,

    message:
      ' Login successful.',

    user:
      user

  };

}


/* =====================================================
   PASSWORD HASH
   ===================================================== */


/* =====================================================
   VERIFY PASSWORD
   ===================================================== */


/* =====================================================
   GENERATE SALT
   ===================================================== */

function generateSalt_() {

  const bytes = [];


  for (
    let i = 0;
    i < 16;
    i++
  ) {

    bytes.push(
      Math.floor(
        Math.random() * 256
      )
    );

  }


  return bytesToHex_(bytes);

}


/* =====================================================
   NORMALIZE PHONE
   ===================================================== */

function normalizePhone_(phone) {

  let value =
    String(
      phone || ''
    )
      .trim()
      .replace(
        /\s+/g,
        ''
      )
      .replace(
        /-/g,
        ''
      );


  /* Persian → English */

  value =
    value.replace(
      /[۰-۹]/g,
      function(d) {

        return '۰۱۲۳۴۵۶۷۸۹'
          .indexOf(d);

      }
    );


  /* Arabic → English */

  value =
    value.replace(
      /[٠-٩]/g,
      function(d) {

        return '٠١٢٣٤٥٦٧٨٩'
          .indexOf(d);

      }
    );


  /* Remove + */

  if (
    value.startsWith('+')
  ) {

    value =
      value.substring(1);

  }

 
  return value;

}


/* =====================================================
   VALID PHONE
   ===================================================== */

function isValidPhone_(phone) {

  return /^989\d{9}$/
    .test(
      String(phone || '')
    );

}


/* =====================================================
   BYTES → HEX
   ===================================================== */

function bytesToHex_(bytes) {

  return bytes
    .map(function(byte) {

      const value =
        byte < 0
          ? byte + 256
          : byte;


      return value
        .toString(16)
        .padStart(
          2,
          '0'
        );

    })
    .join('');

}


/* =====================================================
   CONSTANT TIME COMPARE
   ===================================================== */

function constantTimeEqual_(
  a,
  b
) {

  a =
    String(a || '');

  b =
    String(b || '');


  if (
    a.length !== b.length
  ) {

    return false;

  }


  let result = 0;


  for (
    let i = 0;
    i < a.length;
    i++
  ) {

    result |=
      a.charCodeAt(i) ^
      b.charCodeAt(i);

  }


  return result === 0;

}


/* =====================================================
   JSON RESPONSE
   ===================================================== */
function authorizeExternalRequest() {

  UrlFetchApp.fetch(
    CRYPTO_WORKER_URL + '/hash',
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        text: 'authorization-test'
      })
    }
  );

}



function changePassword(e) {
  
  try {
    const { email, currentPassword, newPassword } = e;
     logToSheet(`email=${email}`)
    // 1. اعتبارسنجی ورودی‌ها
    if (!email ) {
      return {
        success: false,
        message: 'Email or phone number is required'
      };
    }
    
    if (!currentPassword || !newPassword) {
      return {
        success: false,
        message: 'Current password and new password are required'
      };
    }
    currentPassword_hash = hashdata(currentPassword)

    if (newPassword.length < 6) {
      return {
        success: false,
        message: 'New password must be at least 6 characters'
      };
    }
    
    // 2. دسترسی به شیت login
        const sheet = getSheet_("Login");



    // 3. دریافت داده‌ها
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // 4. پیدا کردن ایندکس ستون‌ها
    const emailIndex = headers.indexOf('email');
    //const phoneIndex = headers.indexOf('phone') !== -1 ? headers.indexOf('phone') : headers.indexOf('mobilenumber');
    const passwordIndex = headers.indexOf('password');
    const password_hashIndex = headers.indexOf('password_hash');
    const fullnameIndex = headers.indexOf('fullname') !== -1 ? headers.indexOf('fullname') : headers.indexOf('fullName');
    const isActiveIndex = headers.indexOf('isActive');
    
    if (emailIndex === -1 || passwordIndex === -1) {
      return {
        success: false,
        message: 'Required columns not found: email, phone, password'
      };
    }
    
    // 5. جستجوی کاربر
    let userRow = -1;
    let userData = null;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = String(row[emailIndex] || '').trim();
      //const rowPhone = String(row[phoneIndex] || '').trim();
      
      if ((email && rowEmail === email) /*|| (phone && rowPhone === phone)*/) {
        userRow = i;
        userData = row;
        break;
      }
    }
    
    if (userRow === -1) {
      return {
        success: false,
        message: 'User not found'
      };
    }
    
    // 6. بررسی رمز عبور فعلی
    const storedPassword_hash = String(userData[password_hashIndex] || '').trim();
    
    if (storedPassword_hash !== currentPassword_hash) {
      return {
        success: false,
        message: 'Current password is incorrect'
      };
    }
    
    // 7. به‌روزرسانی رمز عبور جدید
    const rowNumber = userRow + 1; // تبدیل به شماره ردیف (1-based)
    sheet.getRange(rowNumber, passwordIndex + 1).setValue(encryptdata(newPassword));
    sheet.getRange(rowNumber, password_hashIndex + 1).setValue(hashdata(newPassword));
    
    // 8. به‌روزرسانی وضعیت isTemporary به FALSE
    if (isActiveIndex !== -1) {
      sheet.getRange(rowNumber, isActiveIndex + 1).setValue('true');
    }
    
    // 9. ذخیره زمان تغییر رمز (اختیاری)
    const updatedAtIndex = headers.indexOf('updatedAt');
    if (updatedAtIndex !== -1) {
      sheet.getRange(rowNumber, updatedAtIndex + 1).setValue(new Date().toISOString());
    }
    
    return {
      success: true,
      message: 'Password changed successfully',
      user: {
        fullname: userData[fullnameIndex] || 'Unknown',
        email: userData[emailIndex],
        //phone: userData[phoneIndex],
        passwordChanged: true,
        changedAt: new Date().toISOString()
      }
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

/* =====================================================
   SEND EMAIL
   ===================================================== */
/**
 * تأیید ایمیل کاربر - فقط ستون isActive را true می‌کند
  * @param {string} email - ایمیل کاربر
 * @returns {object} - نتیجه عملیات
 */
function verifyEmail(email) {
  
  try {
    // 1. اعتبارسنجی ورودی
    if (!email) {
      return {
        success: false,
        message: 'email are required'
      };
    }
    
    // 2. دسترسی به شیت login
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('login');
    
    if (!sheet) {
      return {
        success: false,
        message: 'Sheet "login" not found'
      };
    }
    
    // 3. دریافت داده‌ها
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // پیدا کردن ایندکس ستون‌ها
    const emailIndex = headers.indexOf('email');
    const isActiveIndex = headers.indexOf('isActive');
    
    // بررسی وجود ستون‌های مورد نیاز
    if (emailIndex === -1) {
      return {
        success: false,
        message: 'Column "email" not found in login sheet'
      };
    }
    
    if (isActiveIndex === -1) {
      return {
        success: false,
        message: 'Column "isActive" not found in login sheet'
      };
    }
    
    // 4. جستجوی کاربر با email و
    let userRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = String(row[emailIndex] || '').trim();
      
      if (rowEmail === email) {
        userRow = i;
        break;
      }
    }
    
    if (userRow === -1) {
      return {
        success: false,
        message: 'Invalid verification link. User not found.'
      };
    }
    
    // 5. بررسی اینکه قبلاً فعال نشده باشد
    const rowNumber = userRow + 1;
    const isActive = String(sheet.getRange(rowNumber, isActiveIndex + 1).getValue() || '').trim().toUpperCase() === 'TRUE';
    
    if (isActive) {
      return {
        success: false,
        status:true,
        message: 'This account is already active. You can login to your account.'
      };
    }
    
    // 6. ✅ به‌روزرسانی ستون isActive به TRUE
    sheet.getRange(rowNumber, isActiveIndex + 1).setValue('true');
    
    return {
      success: true,
      message: 'Your email has been successfully verified! Your account is now active.',
      email: email,
      verifiedAt: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log('Error verifying email: ' + error.message);
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}


function getFullnameByEmail(email) {
  
  try {
    // 1. اعتبارسنجی ورودی
    if (!email) {
      return {
        success: false,
        message: 'Email is required'
      };
    }
    
    // 2. دسترسی به شیت login
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('login');
    
    if (!sheet) {
      return {
        success: false,
        message: 'Sheet "login" not found'
      };
    }
    
    // 3. دریافت داده‌ها
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // پیدا کردن ایندکس ستون‌ها
    const emailIndex = headers.indexOf('email');
    const fullnameIndex = headers.indexOf('fullname');
    
    if (emailIndex === -1) {
      return {
        success: false,
        message: 'Column "email" not found in login sheet'
      };
    }
    
    if (fullnameIndex === -1) {
      return {
        success: false,
        message: 'Column "fullname" not found in login sheet'
      };
    }
    
    // 4. جستجوی کاربر
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = String(row[emailIndex] || '').trim();
      
      if (rowEmail === email) {
        const fullname = String(row[fullnameIndex] || '').trim();
        
        return {
          success: true,
          fullname: fullname,
          email: email
        };
      }
    }
    
    // 5. کاربر پیدا نشد
    return {
      success: false,
      message: 'User not found with this email address'
    };
    
  } catch (error) {
    Logger.log('Error in getFullnameByEmail: ' + error.message);
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

/**
 * ارسال ایمیل خوش‌آمدگویی با لینک تأیید ایمیل
 * @param {string} fullname - نام کامل کاربر
 * @param {string} email - ایمیل کاربر
 * @returns {object} - نتیجه ارسال ایمیل
 */
function sendVerificationEmail(email, fullname='') {
  
  try {
    // 1. اعتبارسنجی
    if (!email) {
      return {
        success: false,
        message: 'Email is required'
      };
    }
    if (!fullname){
      const result = getFullnameByEmail(email);
          if (result.success) {
              fullname = result.fullname;
          } else {
              // مدیریت خطا
              fullname = 'User';
          }
    }
    
       
    // 3. تاریخ فعلی
    const now = new Date();
    const formattedDate = now.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
    const formattedTime = now.toLocaleTimeString('en-US', {
      hour: '2-digit',
      minute: '2-digit'
    });
    const fullDateTime = `${formattedDate} at ${formattedTime}`;
    
    // 4. تاریخ انقضا (7 روز بعد)
    const expiryDate = new Date(now);
    expiryDate.setDate(expiryDate.getDate() + 7);
    const expiryFormatted = expiryDate.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
    
    // 5. ساخت لینک تأیید
    const baseUrl = 'https://confsync.ir';
    //const baseUrl = 'http://192.168.20.80:8083';
    const verificationLink = `${baseUrl}/linksresponse.html?action=verifyEmail&email=${encodeURIComponent(email)}&expiry=${expiryFormatted}`;
    
    const subject = "🎉 Welcome to Confsync - Smart Research Assistant";

    // 6. ساخت HTML ایمیل
    const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="margin:0;padding:0;font-family:Arial,Helvetica,sans-serif;background:#f4f7fb;">
      <div style="background:#f4f7fb;padding:30px;">
        <table width="100%" cellpadding="0" cellspacing="0" style="max-width:650px;margin:auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e5e7eb;">
          
          <!-- Header -->
          <tr>
            <td style="background:#0A7B83;padding:24px;text-align:center;">
              <h1 style="color:#ffffff;margin:0;font-size:26px;">Confsync.Team</h1>
              <p style="color:#d7f4f6;margin:8px 0 0;">Smart Research Assistant for Conferences</p>
            </td>
          </tr>
          
          <!-- Body -->
          <tr>
            <td style="padding:30px;">
              <h2 style="color:#111827;margin-top:0;">Hello Dear ${fullname} 🌹</h2>
              
              <p style="color:#374151;line-height:2;">
                Welcome to <strong>Confsync</strong>! Your account has been successfully created.
              </p>
              
              <!-- User Info -->
              <table width="100%" cellpadding="10" cellspacing="0" style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;margin:16px 0;">
                <tr>
                  <td width="35%" style="color:#6b7280;"><strong>Full Name</strong></td>
                  <td style="color:#111827;">${fullname}</td>
                </tr>
                <tr>
                  <td style="color:#6b7280;"><strong>Email</strong></td>
                  <td style="color:#111827;">${email}</td>
                </tr>
                <!-- ✅ تاریخ ثبت‌نام -->
                <tr>
                  <td style="color:#6b7280;"><strong>Registered On</strong></td>
                  <td style="color:#111827;">${fullDateTime}</td>
                </tr>
                <!-- ✅ تاریخ انقضای لینک -->
                <tr>
                  <td style="color:#6b7280;"><strong>Link Expires</strong></td>
                  <td style="color:#DC2626;font-weight:bold;">${expiryFormatted}</td>
                </tr>
              </table>
              
              <!-- Verification Notice -->
              <div style="margin:24px 0;padding:16px;background:#FEF3C7;border-left:4px solid #F59E0B;border-radius:6px;">
                <strong style="color:#92400E;">⚠️ Email Verification Required</strong><br>
                <span style="color:#78350F;">
                  Please verify your email address by clicking the button below.
                </span>
              </div>
              
              <!-- Verify Button -->
              <div style="text-align:center;margin:30px 0;">
                <a href="${verificationLink}"
                   style="background:#0A7B83;color:#ffffff;text-decoration:none;padding:16px 32px;border-radius:8px;font-weight:bold;font-size:16px;display:inline-block;border:1px solid #0A7B83;">
                  ✅ Verify My Email
                </a>
              </div>
              
              <!-- Alternative Link -->
              <p style="color:#6b7280;font-size:12px;text-align:center;">
                Or copy this link into your browser:<br>
                <span style="color:#0A7B83;word-break:break-all;font-size:11px;">${verificationLink}</span>
              </p>
              
              <!-- Expiry -->
              <p style="color:#6b7280;font-size:12px;text-align:center;margin-top:16px;">
                ⏰ This link expires in 7 days (until ${expiryFormatted})
              </p>
              
              <p style="color:#6b7280;line-height:1.9;">
                If you have any questions, contact us at archive.ghfreza@gmail.com
              </p>
            </td>
          </tr>
          
          <!-- Footer -->
          <tr>
            <td style="background:#f3f4f6;padding:18px;text-align:center;color:#6b7280;font-size:12px;">
              © 2026 Confsync.ir — All Rights Reserved.
            </td>
          </tr>
          
        </table>
      </div>
    </body>
    </html>
    `;
    
    // 7. ارسال ایمیل
    MailApp.sendEmail({
      to: String(email).trim(),
      subject: subject,
      htmlBody: html,
      name: "ConfSync.Team"
    });
        
    return {
      success: true,
      message: 'Verification email sent successfully',
      email: email,
      registeredAt: fullDateTime,
      expiresAt: expiryFormatted,
      sentAt: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log('Error sending email: ' + error.message);
    return {
      success: false,
      message: 'Failed to send email: ' + error.message
    };
  }
}


/**
 * تولید رمز عبور تصادفی 6 رقمی
 * @returns {string} - رمز عبور 6 رقمی
 */
function generateRandomPassword() {
  // روش 1: عدد تصادفی 6 رقمی
  const randomNumber = Math.floor(100000 + Math.random() * 900000);
  return String(randomNumber);
  
  // روش 2: ترکیب حروف و اعداد (اختیاری)
  // const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  // let result = '';
  // for (let i = 0; i < 6; i++) {
  //   result += chars.charAt(Math.floor(Math.random() * chars.length));
  // }
  // return result;
}

/**
 * ارسال ایمیل فراموشی رمز عبور
 * @param {string} email - ایمیل کاربر
 * @param {string} fullname - نام کامل کاربر
 * @param {string} newPassword - رمز عبور جدید
 * @returns {object} - نتیجه عملیات
 */
/**
 * تابع فراموشی رمز عبور - با یک ورودی email
 * @param {string} email - ایمیل کاربر
 * @returns {object} - نتیجه عملیات
 */
function sendForgotPasswordEmail(email) {
  
  try {
    // 1. اعتبارسنجی
    if (!email) {
      return {
        success: false,
        message: 'Email address is required.'
      };
    }
    
    // 2. دسترسی به شیت login
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('login');
    
    if (!sheet) {
      return {
        success: false,
        message: 'System error. Please try again later.'
      };
    }
    
    // 3. دریافت داده‌ها
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // پیدا کردن ایندکس ستون‌ها
    const emailIndex = headers.indexOf('email');
    const passwordIndex = headers.indexOf('password');
    const password_hashIndex = headers.indexOf('password_hash');
    const fullnameIndex = headers.indexOf('fullName');
    const isActiveIndex = headers.indexOf('isActive');
    
    if (emailIndex === -1 || passwordIndex === -1 || password_hashIndex === -1) {
      return {
        success: false,
        message: 'System error. Please try again later.'
      };
    }
    
    // 4. جستجوی کاربر
    let userRow = -1;
    let userData = null;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = String(row[emailIndex] || '').trim();
      
      if (rowEmail === email) {
        userRow = i;
        userData = row;
        break;
      }
    }
    
    if (userRow === -1) {
      return {
        success: false,
        message: 'User not found with this email address.'
      };
    }
    
    // 5. بررسی فعال بودن حساب
    if (isActiveIndex !== -1) {
      const isActive = String(userData[isActiveIndex] || '').trim().toUpperCase() === 'TRUE';
      if (!isActive) {
        return {
          success: false,
          message: 'Your account is inactive. Please activate your account first.'
        };
      }
    }
    
    // 6. دریافت fullname از شیت
    const fullname = fullnameIndex !== -1 ? String(userData[fullnameIndex] || '').trim() : 'User';
    
    // 7. تولید رمز عبور تصادفی 6 رقمی
    const newPassword = String(Math.floor(100000 + Math.random() * 900000));
    
    // 8. به‌روزرسانی رمز عبور در شیت
    const rowNumber = userRow + 1;
    sheet.getRange(rowNumber, passwordIndex + 1).setValue(encryptdata(newPassword));
    sheet.getRange(rowNumber, password_hashIndex + 1).setValue(hashdata(newPassword));
    
    // 9. ارسال ایمیل
    const subject = "🔐 Confsync - Password Reset";
    
    const now = new Date();
    const formattedDate = now.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
    const formattedTime = now.toLocaleTimeString('en-US', {
      hour: '2-digit',
      minute: '2-digit'
    });
    const fullDateTime = `${formattedDate} at ${formattedTime}`;
    
    const loginLink = 'https://confsync.ir/?action=login';
    //const loginLink = 'http://192.168.20.80:8083/?action=login';
    
    const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="margin:0;padding:0;font-family:Arial,Helvetica,sans-serif;background:#f4f7fb;">
      <div style="background:#f4f7fb;padding:30px;">
        <table width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;margin:auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e5e7eb;">
          
          <tr>
            <td style="background:#0A7B83;padding:24px;text-align:center;">
              <img src="https://confsync.ir/assets/centericon.png" 
                   alt="Confsync Logo" 
                   style="width:60px;height:60px;display:block;margin:0 auto 10px;border-radius:50%;" />
              <h1 style="color:#ffffff;margin:0;font-size:24px;">Confsync.Team</h1>
              <p style="color:#d7f4f6;margin:8px 0 0;">Smart Research Assistant for Conferences</p>
            </td>
          </tr>
          
          <tr>
            <td style="padding:30px;">
              <h2 style="color:#111827;margin-top:0;">Hello Dear ${fullname} </h2>
              
              <p style="color:#374151;line-height:1.8;">
                You have requested to reset your password for your <strong>Confsync</strong> account.
              </p>
              
              <div style="margin:24px 0;padding:20px;background:#FEF3C7;border-left:4px solid #F59E0B;border-radius:8px;text-align:center;">
                <p style="color:#78350F;font-size:14px;margin:0 0 8px 0;">
                  <strong>Your new temporary password is:</strong>
                </p>
                <p style="color:#DC2626;font-size:28px;font-weight:bold;letter-spacing:4px;margin:0;">
                  ${newPassword}
                </p>
              </div>
              
              <div style="margin:24px 0;padding:14px 16px;background:#FEF3C7;border-left:4px solid #F59E0B;border-radius:6px;">
                <strong style="color:#92400E;">⚠️ Important Security Note:</strong><br>
                <span style="color:#78350F;font-size:14px;line-height:1.6;">
                  For security reasons, please <strong>change your password</strong> immediately after logging in.
                </span>
              </div>
              
              <div style="text-align:center;margin:30px 0;">
                <a href="${loginLink}"
                   style="background:#0A7B83;color:#ffffff;text-decoration:none;padding:14px 32px;border-radius:8px;font-weight:bold;font-size:16px;display:inline-block;">
                  🔐 Login to System
                </a>
              </div>
              
              <p style="color:#6b7280;font-size:12px;text-align:center;margin-top:16px;">
                Requested on: ${fullDateTime}
              </p>
              
              <p style="color:#6b7280;line-height:1.8;font-size:14px;">
                If you did not request this password reset, please ignore this email or contact support.
                <br>
                <span style="color:#0A7B83;">📧 archive.ghfreza@gmail.com</span>
              </p>
            </td>
          </tr>
          
          <tr>
            <td style="background:#f3f4f6;padding:18px;text-align:center;color:#6b7280;font-size:12px;">
              © 2026 Confsync.ir — All Rights Reserved.
            </td>
          </tr>
          
        </table>
      </div>
    </body>
    </html>
    `;
    
    const plainText = `
    Hello Dear ${fullname},
    
    You have requested to reset your password for your Confsync account.
    
    Your new temporary password is: ${newPassword}
    
    ⚠️ Important: Please change your password immediately after logging in.
    
    Login to your account: ${loginLink}
    
    Requested on: ${fullDateTime}
    
    If you did not request this password reset, please ignore this email.
    
    © 2026 Confsync.ir — All Rights Reserved.
    `;
    
    MailApp.sendEmail({
      to: String(email).trim(),
      subject: subject,
      htmlBody: html,
      body: plainText,
      name: "ConfSync.Team"
    });
    
    return {
      success: true,
      message: 'Password reset email sent successfully.',
      email: email,
      fullname: fullname,
      newPassword: newPassword,
      sentAt: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log('Error in sendForgotPasswordEmail: ' + error.message);
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

function testsendVerificationEmail() {

  sendVerificationEmail(
    "ghfreza@gmail.com",
    "رضا غفاری"
  );

}

