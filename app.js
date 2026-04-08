function AUTH_TEST() {
  return AUTH_TEST_();
}

function AUTH_TEST_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("OK spreadsheet: " + (ss ? ss.getName() : "NULL"));
}

function TEST_STUDENTS_COURSE() {
  const ss = getStudentsSS_();
  Logger.log("Hojas: " + ss.getSheets().length);

  const course = "6°A"; // pon uno que veas en tu combo
  const sh = findCourseSheet_(ss, course);
  Logger.log("Hoja encontrada: " + (sh ? sh.getName() : "NO"));

  const students = listStudentsInCourse_(course);
  Logger.log("Estudiantes: " + students.length);
  Logger.log(students.slice(0, 10).join(" | "));
}


/************ CONFIG ************/
const API_KEY = "deimerDh2191docentesRegistros2026appandoidios"; // <- pon una clave larga (ej: 30+ caracteres)
const TIMEZONE = "America/Bogota";
const LOCALE = "es-CO";
// ✅ WhatsApp fijo (Coordinadora de convivencia) - formato: 57XXXXXXXXXX
const WHATSAPP_COORD_PHONE = "573163021721";
// ✅ QR esperado (el QR debe contener EXACTAMENTE este texto)
const QR_EXPECTED_VALUE = "ASISTENCIA-DEIMERH-2026";


// ==============================
// 📚 LIBRO DE LISTAS DE ESTUDIANTES (EXTERNO)
// ==============================
const STUDENTS_SS_ID = "1ikkebcq3hlWWzwPPG-ularR6TxgjqLg5fJRjY_TuFNo";

// ==============================
// 📕 LIBRO DE FALTAS (EXTERNO)  ✅ NUEVO
// ==============================
const FALTAS_SS_ID = "1IuIn3xAntGpapsPp_Lt8cv6crmy0JzC9TLaukcraM8Q";


// Hora límite para "A TIEMPO" (Entrada):
// Ejemplo: 06:35:00 => si llega después es tarde
const ON_TIME_HHMM = "06:30";
// Hora mínima para "SALIDA OK" (Salida):
// Ejemplo: 14:00 => si sale antes es rojo, si sale desde esa hora en adelante es verde
const EXIT_OK_HHMM = "14:00";


/************ UTILIDADES ************/
function getStudentsSS_() {
  return SpreadsheetApp.openById(STUDENTS_SS_ID);
}

function getFaltasSS_() {
  return SpreadsheetApp.openById(FALTAS_SS_ID);
}



function monthNameEs(m) {
  const names = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  return names[m];
}

function sheetNameForDate(d) {
  const m = d.getMonth();
  const y = d.getFullYear();
  return `${monthNameEs(m)} ${y}`;
}

function requireKey_(e) {
  const key = (e.parameter && e.parameter.key) || "";
  const headerKey = (e.postData && e.postData.contents) ? "" : "";
  // En Apps Script no es fácil leer headers directo en doPost simple,
  // así que usamos key por query param y por body también.
  if (key === API_KEY) return true;

  if (e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      if (body.key === API_KEY) return true;
    } catch (err) {}
  }
  return false;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function validateQrOrThrow_(qrText) {
  const t = String(qrText || "").trim();
  if (!t) throw new Error("missing_qr");
const upper = t.toUpperCase();
const expected = QR_EXPECTED_VALUE.toUpperCase();
if (!(t === QR_EXPECTED_VALUE || upper.indexOf(expected) >= 0)) {
  throw new Error("invalid_qr");
}
return QR_EXPECTED_VALUE;
}


function normalizePhoneToWa_(raw) {
  let s = String(raw || "").trim();
  if (!s) return "";
  s = s.replace(/[^\d+]/g, ""); // deja dígitos y +
  // Si viene sin país y tiene 10 dígitos (CO), antepone 57
  const digits = s.replace(/\D/g, "");
  if (digits.length === 10) return "57" + digits;
  // Si viene con +57..., quita el +
  if (s.startsWith("+")) s = s.slice(1);
  return s.replace(/\D/g, "");
}


function parseWhatsAppRecipients_(raw) {
  const s = String(raw || "").trim();
  if (!s) return [];

  // separadores: coma, punto y coma, saltos de línea, espacios múltiples
  const parts = s.split(/[\s,;]+/g).map(x => x.trim()).filter(Boolean);

  // normaliza y elimina duplicados
  const uniq = [];
  const seen = {};
  for (const p of parts) {
    const norm = normalizePhoneToWa_(p);
    if (!norm) continue;
    if (!seen[norm]) {
      seen[norm] = true;
      uniq.push(norm);
    }
  }
  return uniq;
}



function buildWaLink_(phone, text) {
  const p = normalizePhoneToWa_(phone);
  if (!p) return "";
  const msg = encodeURIComponent(String(text || ""));
  return `https://wa.me/${p}?text=${msg}`;
}

function getWhatsAppToByDevice_(email, deviceId) {
  const info = isDeviceAuthorized_(email, deviceId);
  return String(info.whatsappTo || "").trim();
}

function setWhatsAppToByDevice_(email, deviceId, whatsappTo) {
  const sh = ensureDevicesSheet_();
  const lastRow = sh.getLastRow();
  const values = lastRow >= 2 ? sh.getRange(2, 1, lastRow - 1, 5).getValues() : [];

  const eKey = String(email || "").trim().toLowerCase();
  const dKey = String(deviceId || "").trim();
  const wa = normalizePhoneToWa_(whatsappTo);

  // busca fila existente
  for (let i = 0; i < values.length; i++) {
    const e = String(values[i][0] || "").trim().toLowerCase();
    const d = String(values[i][1] || "").trim();
    if (e === eKey && d === dKey) {
      sh.getRange(i + 2, 5).setValue(wa);
      return { updated:true, row: i + 2, whatsappTo: wa };
    }
  }

  // si no existe, crea fila (NO autorizado por defecto)
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 5).setValues([[eKey, dKey, "NO", "Auto-creado al guardar WhatsAppTo", wa]]);
  return { updated:true, row, whatsappTo: wa };
}

function fmtTsBogota_(ts) {
  const d = safeDate_(ts);
  if (!d) return { date:"", time:"", dt:"" };

  // strings listos para UI, SIN que el front haga new Date()
  const date = Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd");
  const time = Utilities.formatDate(d, TIMEZONE, "HH:mm:ss");
  const dt   = Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd HH:mm:ss");
  return { date, time, dt };
}


function safeDate_(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;

  const d = new Date(val);
  if (d instanceof Date && !isNaN(d.getTime())) return d;

  return null; // inválido (ej: "NTC")
}


function ensureMonthSheet_(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = sheetNameForDate(d);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange("A1").setValue("DOCENTES");
    sh.getRange("A2").setValue("TIPO");
    sh.setFrozenRows(2);
    sh.setFrozenColumns(0);
    sh.getRange("A1:A2").setFontWeight("bold");
  }
  return sh;
}


// ✅ NUEVA: obtener hoja de un mes "YYYY-MM" SIN crearla
function getMonthSheetByKey_(monthKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mk = String(monthKey || "").trim();

  // Si no mandan mes, usa el actual
  const d = mk ? new Date(`${mk}-01T00:00:00`) : new Date();
  const name = sheetNameForDate(d);

  // IMPORTANTE: aquí NO creamos hoja si no existe
  return ss.getSheetByName(name);
}

// ✅ NUEVA: listar meses disponibles en el spreadsheet (hojas "enero 2026", etc.)
function listAvailableMonths_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(s => s.getName());

  // Filtra hojas con formato: "<mes> <año>" (ej: "enero 2026")
  const monthsEs = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  const re = new RegExp(`^(${monthsEs.join("|")})\\s+(\\d{4})$`, "i");

  const items = [];

  for (const name of sheets) {
    const m = name.match(re);
    if (!m) continue;

    const mesTxt = m[1].toLowerCase();
    const year = Number(m[2]);
    const monthIndex = monthsEs.indexOf(mesTxt); // 0..11
    if (monthIndex < 0) continue;

    const key = `${year}-${String(monthIndex + 1).padStart(2, "0")}`; // YYYY-MM
    items.push({ key, label: name });
  }

  // Orden desc por key (más reciente primero)
  items.sort((a, b) => (a.key < b.key ? 1 : a.key > b.key ? -1 : 0));
  return items;
}



function ensureDevicesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("DISPOSITIVOS");
  if (!sh) {
    sh = ss.insertSheet("DISPOSITIVOS");
    sh.getRange("A1:E1").setValues([["EMAIL","DEVICE_ID","AUTORIZADO","NOTAS","WHATSAPP_TO"]]);
    sh.getRange("A1:E1").setFontWeight("bold");
    sh.setFrozenRows(1);
  } else {
    // Si la hoja ya existe con 4 columnas, agrega la 5ta (sin romper nada)
    const lastCol = sh.getLastColumn();
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x||"").trim().toUpperCase());
    if (!header.includes("WHATSAPP_TO")) {
      sh.getRange(1, lastCol + 1).setValue("WHATSAPP_TO").setFontWeight("bold");
    }
  }
  return sh;
}






function ensureConfigSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("CONFIG");
  if (!sh) {
    sh = ss.insertSheet("CONFIG");
    sh.getRange("A1:B1").setValues([["CLAVE","VALOR"]]);
    sh.getRange("A1:B1").setFontWeight("bold");
    sh.setFrozenRows(1);

    const defaults = [
      ["ON_TIME_HHMM", ON_TIME_HHMM],
      ["EXIT_OK_HHMM", EXIT_OK_HHMM],
      ["SCHOOL_RADIUS_METERS", "120"],
      ["REQUIRED_ACCURACY_METERS", "50"],
      ["SCHOOL_LAT", "4.55445"],
      ["SCHOOL_LNG", "-74.11165"],
    ];
    sh.getRange(2, 1, defaults.length, 2).setValues(defaults);
    sh.autoResizeColumns(1, 2);
  } else {
    const lastRow = sh.getLastRow();
    const values = lastRow >= 2 ? sh.getRange(2, 1, lastRow - 1, 2).getValues() : [];
    const keys = {};
    values.forEach(r => {
      const k = String(r[0] || "").trim();
      if (k) keys[k] = true;
    });

    const missing = [];
    if (!keys["ON_TIME_HHMM"]) missing.push(["ON_TIME_HHMM", ON_TIME_HHMM]);
    if (!keys["EXIT_OK_HHMM"]) missing.push(["EXIT_OK_HHMM", EXIT_OK_HHMM]);
    if (!keys["SCHOOL_RADIUS_METERS"]) missing.push(["SCHOOL_RADIUS_METERS", "120"]);
    if (!keys["REQUIRED_ACCURACY_METERS"]) missing.push(["REQUIRED_ACCURACY_METERS", "50"]);
    if (!keys["SCHOOL_LAT"]) missing.push(["SCHOOL_LAT", "4.55445"]);
    if (!keys["SCHOOL_LNG"]) missing.push(["SCHOOL_LNG", "-74.11165"]);

    if (missing.length) {
      sh.getRange(sh.getLastRow() + 1, 1, missing.length, 2).setValues(missing);
      sh.autoResizeColumns(1, 2);
    }
  }
  return sh;
}





function getConfig_() {
  const sh = ensureConfigSheet_();

  const defaults = {
    ON_TIME_HHMM: ON_TIME_HHMM,
    EXIT_OK_HHMM: EXIT_OK_HHMM,
    SCHOOL_RADIUS_METERS: 30,
    REQUIRED_ACCURACY_METERS: 15,
    SCHOOL_LAT: 4.55445,
    SCHOOL_LNG: -74.11165,
  };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return defaults;

  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();

  function toHHMM_(v) {
    if (v instanceof Date) {
      return Utilities.formatDate(v, TIMEZONE, "HH:mm");
    }

    if (typeof v === "number" && isFinite(v)) {
      const totalMinutes = Math.round(v * 24 * 60);
      const hh = String(Math.floor(totalMinutes / 60)).padStart(2, "0");
      const mm = String(totalMinutes % 60).padStart(2, "0");
      return `${hh}:${mm}`;
    }

    const s = String(v ?? "").trim();
    const m = s.match(/^(\d{1,2}):(\d{2})$/);
    if (m) return String(m[1]).padStart(2, "0") + ":" + m[2];

    return s;
  }

  const map = {};
  values.forEach(([k, v]) => {
    const key = String(k || "").trim();
    if (!key) return;

    if (key === "ON_TIME_HHMM" || key === "EXIT_OK_HHMM") {
      map[key] = toHHMM_(v);
    } else {
      map[key] = String(v ?? "").trim();
    }
  });

  const cfg = { ...defaults };

  if (map.ON_TIME_HHMM) cfg.ON_TIME_HHMM = map.ON_TIME_HHMM;
  if (map.EXIT_OK_HHMM) cfg.EXIT_OK_HHMM = map.EXIT_OK_HHMM;

  if (map.SCHOOL_RADIUS_METERS) {
    const n = Number(map.SCHOOL_RADIUS_METERS);
    if (!isNaN(n) && n > 0) cfg.SCHOOL_RADIUS_METERS = n;
  }

  if (map.REQUIRED_ACCURACY_METERS) {
    const n = Number(map.REQUIRED_ACCURACY_METERS);
    if (!isNaN(n) && n > 0) cfg.REQUIRED_ACCURACY_METERS = n;
  }

  if (map.SCHOOL_LAT) {
    const n = Number(map.SCHOOL_LAT);
    if (!isNaN(n) && n >= -90 && n <= 90) cfg.SCHOOL_LAT = n;
  }

  if (map.SCHOOL_LNG) {
    const n = Number(map.SCHOOL_LNG);
    if (!isNaN(n) && n >= -180 && n <= 180) cfg.SCHOOL_LNG = n;
  }

  return cfg;
}


function setConfig_(newCfg) {
  const sh = ensureConfigSheet_();

  const allowedKeys = [
    "ON_TIME_HHMM",
    "EXIT_OK_HHMM",
    "SCHOOL_RADIUS_METERS",
    "REQUIRED_ACCURACY_METERS",
    "SCHOOL_LAT",
    "SCHOOL_LNG"
  ];

  const lastRow = sh.getLastRow();
  const values = (lastRow >= 2) ? sh.getRange(2, 1, lastRow - 1, 2).getValues() : [];
  const rowByKey = {};
  values.forEach((r, i) => {
    const k = String(r[0] || "").trim();
    if (k) rowByKey[k] = i + 2;
  });

  allowedKeys.forEach(k => {
    if (!(k in newCfg)) return;
    const val = String(newCfg[k]).trim();
    if (!val) return;

    const row = rowByKey[k] || (sh.getLastRow() + 1);
    sh.getRange(row, 1, 1, 2).setValues([[k, val]]);
  });

  sh.autoResizeColumns(1, 2);
}

/************ WHATSAPP (DESTINO) ************/

const WHATSAPP_TO_KEY = "WHATSAPP_TO";

// Lee valor WHATSAPP_TO desde hoja CONFIG (si no existe, devuelve "")
function getWhatsAppTo_() {
  const sh = ensureConfigSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return "";

  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  for (const [k, v] of values) {
    const key = String(k || "").trim();
    if (key === WHATSAPP_TO_KEY) {
      return String(v || "").trim();
    }
  }
  return "";
}

// Upsert en CONFIG: WHATSAPP_TO = value
function setWhatsAppTo_(value) {
  const sh = ensureConfigSheet_();
  const val = String(value || "").trim();

  // permitir vacío (por si quieres borrar)
  const lastRow = sh.getLastRow();
  const values = (lastRow >= 2) ? sh.getRange(2, 1, lastRow - 1, 2).getValues() : [];

  let rowFound = 0;
  for (let i = 0; i < values.length; i++) {
    const key = String(values[i][0] || "").trim();
    if (key === WHATSAPP_TO_KEY) {
      rowFound = i + 2;
      break;
    }
  }

  if (rowFound) {
    sh.getRange(rowFound, 2).setValue(val);
  } else {
    sh.appendRow([WHATSAPP_TO_KEY, val]);
  }

  sh.autoResizeColumns(1, 2);
  return val;
}

// Normaliza a SOLO dígitos (ej "+57 300-123" -> "57300123" o "300123")
function normalizeWhatsapp_(s) {
  return String(s || "").replace(/\D+/g, "");
}



/************ ADMINISTRADORES ************/
const DEFAULT_ADMIN_PIN = "2012"; // PIN inicial por defecto

function ensureAdminsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("ADMINISTRADORES");
  if (!sh) {
    sh = ss.insertSheet("ADMINISTRADORES");
    sh.getRange("A1:C1").setValues([["PIN","UPDATED_AT","NOTAS"]]);
    sh.getRange("A1:C1").setFontWeight("bold");
    sh.setFrozenRows(1);

    const stamp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    sh.getRange("A2:C2").setValues([[DEFAULT_ADMIN_PIN, stamp, "PIN por defecto (cambiar al primer ingreso)"]]);
    sh.autoResizeColumns(1, 3);
  }
  return sh;
}

function getAdminPin_() {
  const sh = ensureAdminsSheet_();
  const pin = String(sh.getRange("A2").getValue() || "").trim();
  return pin || DEFAULT_ADMIN_PIN;
}

function setAdminPin_(newPin, notes) {
  const sh = ensureAdminsSheet_();
  const stamp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");
  sh.getRange("A2:C2").setValues([[String(newPin).trim(), stamp, String(notes || "")]]);
  sh.autoResizeColumns(1, 3);
}

function verifyAdminPin_(pin) {
  const saved = getAdminPin_();
  return String(saved) === String(pin || "").trim();
}



function emailHasAuthorizedOtherDevice_(email, deviceId) {
  const sh = ensureDevicesSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { conflict:false, devices:[] };

  const values = sh.getRange(2, 1, lastRow - 1, 5).getValues(); // A-E

  email = String(email || "").trim().toLowerCase();
  deviceId = String(deviceId || "").trim();

  const devices = [];

  for (const row of values) {
    const e = String(row[0] || "").trim().toLowerCase();
    const d = String(row[1] || "").trim();
    const auth = row[2];

    const a = String(auth).trim().toLowerCase();
    const authorized = (a === "si" || a === "sí" || a === "true" || auth === true);

    if (e === email && authorized) devices.push(d);
  }

  const conflict = devices.length > 0 && !devices.includes(deviceId);
  return { conflict, devices };
}



function isDeviceAuthorized_(email, deviceId) {
  const sh = ensureDevicesSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { authorized:false, allowedName:"", whatsappTo:"" };

  // A:EMAIL B:DEVICE_ID C:AUTORIZADO D:NOTAS E:WHATSAPP_TO
  const values = sh.getRange(2, 1, lastRow - 1, 5).getValues();

  email = String(email || "").trim().toLowerCase();
  deviceId = String(deviceId || "").trim();

  for (const row of values) {
    const e = String(row[0] || "").trim().toLowerCase();
    const d = String(row[1] || "").trim();
    const auth = row[2];
    const notes = String(row[3] || "");
    const whatsappTo = String(row[4] || "").trim();

    if (e === email && d === deviceId) {
      const a = String(auth).trim().toLowerCase();
      const authorized = (a === "si" || a === "sí" || a === "true" || auth === true);

      let allowedName = "";
      const parts = notes.split("|").map(s => s.trim());
      if (parts.length >= 2) allowedName = parts[1] || "";

      return { authorized, allowedName, whatsappTo };
    }
  }

  return { authorized:false, allowedName:"", whatsappTo:"" };
}

function listAuthorizedWhatsappRecipients_(excludeEmail, excludeDeviceId) {
  const sh = ensureDevicesSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  // A:EMAIL B:DEVICE_ID C:AUTORIZADO D:NOTAS E:WHATSAPP_TO
  const values = sh.getRange(2, 1, lastRow - 1, 5).getValues();

  const exE = String(excludeEmail || "").trim().toLowerCase();
  const exD = String(excludeDeviceId || "").trim();

  const out = [];
  const seen = {};

  for (const row of values) {
    const email = String(row[0] || "").trim().toLowerCase();
    const deviceId = String(row[1] || "").trim();
    const auth = row[2];
    const waRaw = String(row[4] || "").trim();

    const a = String(auth).trim().toLowerCase();
    const authorized = (a === "si" || a === "sí" || a === "true" || auth === true);
    if (!authorized) continue;

    // excluir al que está marcando
    if (email === exE && deviceId === exD) continue;

    const wa = normalizePhoneToWa_(waRaw);
    if (!wa) continue;

    if (!seen[wa]) {
      seen[wa] = true;
      out.push(wa);
    }
  }

  return out;
}



// Busca la columna del docente (par Entrada/Salida). Si no existe, la crea.
function ensureTeacherColumns_(sh, teacher) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headerRow1 = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  // El formato que usaremos:
  // Row1: [ ... , Teacher (merged 2 cols) ]
  // Row2: [ ... , "Entrada", "Salida" ]
  // Entonces el nombre del docente estará en la primera de esas dos columnas (la merged)
  // y la siguiente columna estará vacía en row1.
  for (let c = 1; c <= lastCol; c++) {
    if (String(headerRow1[c - 1]).trim().toLowerCase() === teacher.trim().toLowerCase()) {
      // docente ya existe en columna c (Entrada) y c+1 (Salida)
      return { entradaCol: c, salidaCol: c + 1 };
    }
  }

  // No existe -> crear al final, en el siguiente par disponible.
  const entradaCol = lastCol + 1;
  const salidaCol = lastCol + 2;

  // Set headers
  sh.getRange(1, entradaCol).setValue(teacher);
  sh.getRange(2, entradaCol).setValue("Entrada");
  sh.getRange(2, salidaCol).setValue("Salida");

  // Merge nombre docente en 2 columnas
  sh.getRange(1, entradaCol, 1, 2).merge();

  // Estilos
  sh.getRange(1, entradaCol, 2, 2).setFontWeight("bold");
  sh.getRange(1, entradaCol, 2, 2).setHorizontalAlignment("center");
  sh.autoResizeColumns(entradaCol, 2);

  return { entradaCol, salidaCol };
}

// Encuentra la siguiente fila vacía de UNA columna específica (apilar hacia abajo)
function nextEmptyRowInColumn_(sh, col) {
  const lastRow = Math.max(sh.getLastRow(), 2);
  const values = sh.getRange(3, col, Math.max(lastRow - 2, 1), 1).getValues(); // desde fila 3
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "" && values[i][0] !== null) {
      return 3 + i + 1;
    }
  }
  return 3; // primera fila de datos
}

function hasRecordForDay_(sh, col, dateObj) {
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return false;

  const dayStr = Utilities.formatDate(dateObj, TIMEZONE, "yyyy-MM-dd");

  const values = sh.getRange(3, col, lastRow - 2, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const val = values[i][0];
    if (!val) continue;

    const d = safeDate_(val);     // ✅ usa tu helper robusto
    if (!d) continue;             // ✅ si no se puede parsear, ignora

    const dStr = Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd");
    if (dStr === dayStr) return true;
  }
  return false;
}



function hasEntradaForDay_(sh, entradaCol, dateObj) {
  // Reutiliza la misma lógica de hasRecordForDay_
  return hasRecordForDay_(sh, entradaCol, dateObj);
}



function autoFitCell_(sh, row, col) {
  // Ajusta el ancho de la columna al contenido
  sh.autoResizeColumn(col);

  // Ajusta la altura de la fila (por si el contenido o formato lo requiere)
  if (typeof sh.autoResizeRows === "function") {
    sh.autoResizeRows(row, 1);
  }
}


function isLate_(dateObj) {
  const cfg = getConfig_();
  const [hh, mm] = String(cfg.ON_TIME_HHMM || ON_TIME_HHMM).split(":").map(Number);
  const limit = new Date(dateObj);
  limit.setHours(hh, mm, 0, 0);
  return dateObj.getTime() > limit.getTime();
}

function isEarlyExit_(dateObj) {
  const cfg = getConfig_();
  const [hh, mm] = String(cfg.EXIT_OK_HHMM || EXIT_OK_HHMM).split(":").map(Number);
  const limit = new Date(dateObj);
  limit.setHours(hh, mm, 0, 0);
  return dateObj.getTime() < limit.getTime();
}


/************ DISCIPLINA (LLAMADO DE ATENCIÓN) ************/
function ensureDisciplineSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("DISCIPLINA");
  if (!sh) {
    sh = ss.insertSheet("DISCIPLINA");
    sh.getRange("A1:H1").setValues([[
      "month", "student", "maxLevel", "totalMarks", "who_json", "updatedAt", "reported", "reportedAt"
    ]]);
    sh.getRange("A1:H1").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 8);
  }
  return sh;
}

// ✅ NUEVO: normaliza month venga como Date / "YYYY-M" / "YYYY-MM-DD" / con apóstrofe
function normalizeMonthKey_(val) {
  if (!val) return "";

  // Si Sheets lo guardó como Date
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, TIMEZONE, "yyyy-MM");
  }

  // String
  let s = String(val).trim();

  // Quita apóstrofe inicial si quedó como "'2026-02"
  if (s.startsWith("'")) s = s.slice(1).trim();

  // Si viene como "YYYY-MM-DD..." recorta
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 7);

  // Si viene como "YYYY-M" => pad
  const m1 = s.match(/^(\d{4})-(\d{1,2})$/);
  if (m1) return `${m1[1]}-${String(m1[2]).padStart(2, "0")}`;

  // Si ya es "YYYY-MM"
  const m2 = s.match(/^(\d{4})-(\d{2})$/);
  if (m2) return s;

  return s; // fallback (mejor devolver algo que vacío)
}


function discFindRow_(sh, month, student) {
  const last = sh.getLastRow();
  if (last < 2) return 0;

  const mk = normalizeMonthKey_(month);
  const st = String(student || "").trim().toLowerCase();

  const values = sh.getRange(2, 1, last - 1, 2).getValues(); // month, student
  for (let i = 0; i < values.length; i++) {
    const m = normalizeMonthKey_(values[i][0]);
    const s = String(values[i][1] || "").trim().toLowerCase();
    if (m === mk && s === st) return i + 2;
  }
  return 0;
}


function discipline_list_(month) {
  const sh = ensureDisciplineSheet_();
  const mk = normalizeMonthKey_(month);
  const last = sh.getLastRow();
  if (last < 2) return [];

  const data = sh.getRange(2, 1, last - 1, 8).getValues();
  const out = [];

  data.forEach(r => {
    const m = normalizeMonthKey_(r[0]);
    if (mk && m !== mk) return;

    let who = [];
    try { who = JSON.parse(r[4] || "[]"); } catch (e) { who = []; }
    if (!Array.isArray(who)) who = [];

    const whoClean = who.map(x => ({
      teacher: String(x.teacher || "").trim(),
      email: String(x.email || "").trim(),
      level: Number(x.level || 0),
      ts: x.ts || ""
    })).filter(x => x.teacher);

    const maxLevel = Number(r[2] || 0);
    const totalMarks = Number(r[3] || 0);

    const repCell = r[6];
    const reported =
      (repCell === true) ||
      String(repCell || "").trim().toUpperCase() === "SI" ||
      String(repCell || "").trim().toUpperCase() === "YES";

    out.push({
      month: m,
      student: r[1],
      maxLevel,
      totalMarks,
      teachersCount: whoClean.length,
who: whoClean.map(x => {
  const f = x.ts ? fmtTsBogota_(x.ts) : {date:"", time:"", dt:""};
  return {
    teacher: x.teacher,
    level: x.level,
    date: f.date,
    time: f.time,
    dt: f.dt
  };
}),

      reported,
      reportedAt: r[7] || "",
      updatedAt: r[5] || ""
    });
  });

  // orden: reportados primero, luego maxLevel desc, luego updatedAt desc
  out.sort((a, b) => {
    const ra = a.reported ? 1 : 0;
    const rb = b.reported ? 1 : 0;
    if (ra !== rb) return rb - ra;
    if (a.maxLevel !== b.maxLevel) return b.maxLevel - a.maxLevel;
    return String(b.updatedAt || "").localeCompare(String(a.updatedAt || ""));
  });

  return out;
}


function discipline_mark_(payload) {
  const month = normalizeMonthKey_(payload.month);
  const student = String(payload.student || "").trim().replace(/\s+/g, " ");
  const level = Number(payload.level || 0);
  const teacher = String(payload.teacher || "").trim();
  const email = String(payload.email || "").trim().toLowerCase();

  // Para WhatsApp (quién está marcando)
  const deviceId = String(payload.device_id || "").trim();

  if (!month || !student) throw new Error("Faltan datos: month/student");
  if (!teacher) throw new Error("Falta teacher");
  if (!level || level < 1 || level > 4) throw new Error("level inválido (1..4)");

  const sh = ensureDisciplineSheet_();
  const row = discFindRow_(sh, month, student);
  const now = new Date();

  let who = [];
  let maxLevel = 0;
  let totalMarks = 0;
  let reported = "NO";
  let reportedAt = "";

  if (row) {
    const r = sh.getRange(row, 1, 1, 8).getValues()[0];

    maxLevel = Number(r[2] || 0);
    totalMarks = Number(r[3] || 0);
    reported = String(r[6] || "NO").toUpperCase();
    reportedAt = r[7] || "";

    try { who = JSON.parse(r[4] || "[]"); } catch (e) { who = []; }
    if (!Array.isArray(who)) who = [];
  }

  // idempotencia por docente (email o nombre)
  const idx = who.findIndex(x =>
    String(x.email || "").toLowerCase() === email ||
    String(x.teacher || "").toLowerCase() === teacher.toLowerCase()
  );

  const entry = { teacher, email, level, ts: now.toISOString() };

  if (idx >= 0) {
    const prevLevel = Number(who[idx].level || 0);
    if (level > prevLevel) {
      who[idx] = entry;
      totalMarks += 1;
    }
  } else {
    who.push(entry);
    totalMarks += 1;
  }

  maxLevel = who.reduce((m, x) => Math.max(m, Number(x.level || 0)), 0);

  let justReported = false;
  let wa_url = "";
  let wa_msg = "";

  if (maxLevel >= 4 && reported !== "SI") {
    reported = "SI";
    reportedAt = now.toISOString();
    justReported = true;

    // ✅ Mensaje (incluye estudiante)
    wa_msg =
      `Hola profe,\n\n` +
      `El estudiante "${student}" alcanzó el cuarto llamado de atención por el uso del celular.\n\n` +
      `Mes: ${month}\n\n` +
      `Docentes participantes:\n` +
      who.map(x => `- ${x.teacher} (nivel ${x.level})`).join("\n");

 // ✅ Ahora NO forzamos WhatsApp a un número fijo.
// La app abrirá el "Compartir" nativo con el texto (wa_msg).
wa_url = "";

  }

  const whoJson = JSON.stringify(who);
  const updatedAt = now.toISOString();

  if (row) {
    sh.getRange(row, 1).setNumberFormat("@");
    sh.getRange(row, 1, 1, 8).setValues([[
      "'" + month, student, maxLevel, totalMarks, whoJson, updatedAt, reported, reportedAt
    ]]);
  } else {
    sh.appendRow(["'" + month, student, maxLevel, totalMarks, whoJson, updatedAt, reported, reportedAt]);
    const newRow = sh.getLastRow();
    sh.getRange(newRow, 1).setNumberFormat("@");
  }

  return {
    report: {
      student,
      maxLevel,
      totalMarks,
      who: who.map(x => ({ teacher: x.teacher, level: x.level }))
    },
    list: discipline_list_(month),
    justReported,
    wa_msg,
    wa_url
  };
}

// ✅ NUEVO: normaliza estudiante para comparar (más robusto: tildes/espacios)
function normalizeStudentKey_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // quita tildes
    .replace(/\s+/g, " "); // colapsa espacios
}



// ✅ NUEVO: elimina todas las filas de un estudiante en un mes (DISCIPLINA)
function discipline_delete_student_(payload) {
  const month = normalizeMonthKey_(payload.month);
  const studentRaw = String(payload.student || "").trim().replace(/\s+/g, " ");
  if (!month || !studentRaw) throw new Error("Faltan datos: month/student");

  const sh = ensureDisciplineSheet_();
  const last = sh.getLastRow();
  if (last < 2) {
    return { deleted_rows: 0, msg: "No hay registros.", list: discipline_list_(month) };
  }

  // Leemos month+student (A:B) para decidir qué filas borrar
  const values = sh.getRange(2, 1, last - 1, 2).getValues(); // A:month B:student
  const targetMonth = normalizeMonthKey_(month);
  const targetStudent = normalizeStudentKey_(studentRaw);

  // Borrar de abajo hacia arriba para no dañar índices
  let deleted = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    const rowMonth = normalizeMonthKey_(values[i][0]);
    const rowStudent = normalizeStudentKey_(values[i][1]);

    if (rowMonth === targetMonth && rowStudent === targetStudent) {
      sh.deleteRow(i + 2); // +2 porque values inicia en fila 2
      deleted++;
    }
  }

  const list = discipline_list_(month);

  return {
    deleted_rows: deleted,
    msg: deleted
      ? `Estudiante eliminado ✅ (${deleted} fila(s) borrada(s) en ${targetMonth}).`
      : "No se encontraron filas para eliminar.",
    list
  };
}


/************ PIERCING Y PERFORACIONES (LLAMADO DE ATENCIÓN) ************/
function ensurePiercingSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("PIERCING");
  if (!sh) {
    sh = ss.insertSheet("PIERCING");
    sh.getRange("A1:H1").setValues([[
      "month", "student", "maxLevel", "totalMarks", "who_json", "updatedAt", "reported", "reportedAt"
    ]]);
    sh.getRange("A1:H1").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 8);
  }
  return sh;
}

function piercFindRow_(sh, month, student) {
  const last = sh.getLastRow();
  if (last < 2) return 0;

  const mk = normalizeMonthKey_(month);
  const st = String(student || "").trim().toLowerCase();

  const values = sh.getRange(2, 1, last - 1, 2).getValues(); // month, student
  for (let i = 0; i < values.length; i++) {
    const m = normalizeMonthKey_(values[i][0]);
    const s = String(values[i][1] || "").trim().toLowerCase();
    if (m === mk && s === st) return i + 2;
  }
  return 0;
}

function piercing_list_(month) {
  const sh = ensurePiercingSheet_();
  const mk = normalizeMonthKey_(month);
  const last = sh.getLastRow();
  if (last < 2) return [];

  const data = sh.getRange(2, 1, last - 1, 8).getValues();
  const out = [];

  data.forEach(r => {
    const m = normalizeMonthKey_(r[0]);
    if (mk && m !== mk) return;

    let who = [];
    try { who = JSON.parse(r[4] || "[]"); } catch (e) { who = []; }
    if (!Array.isArray(who)) who = [];

    const whoClean = who.map(x => ({
      teacher: String(x.teacher || "").trim(),
      email: String(x.email || "").trim(),
      level: Number(x.level || 0),
      ts: x.ts || ""
    })).filter(x => x.teacher);

    const maxLevel = Number(r[2] || 0);
    const totalMarks = Number(r[3] || 0);

    const repCell = r[6];
    const reported =
      (repCell === true) ||
      String(repCell || "").trim().toUpperCase() === "SI" ||
      String(repCell || "").trim().toUpperCase() === "YES";

    out.push({
      month: m,
      student: r[1],
      maxLevel,
      totalMarks,
      teachersCount: whoClean.length,
who: whoClean.map(x => {
  const f = x.ts ? fmtTsBogota_(x.ts) : {date:"", time:"", dt:""};
  return {
    teacher: x.teacher,
    level: x.level,
    date: f.date,
    time: f.time,
    dt: f.dt
  };
}),
      reported,
      reportedAt: r[7] || "",
      updatedAt: r[5] || ""
    });
  });

  // orden: reportados primero, luego maxLevel desc, luego updatedAt desc
  out.sort((a, b) => {
    const ra = a.reported ? 1 : 0;
    const rb = b.reported ? 1 : 0;
    if (ra !== rb) return rb - ra;
    if (a.maxLevel !== b.maxLevel) return b.maxLevel - a.maxLevel;
    return String(b.updatedAt || "").localeCompare(String(a.updatedAt || ""));
  });

  return out;
}

function piercing_mark_(payload) {
  const month = normalizeMonthKey_(payload.month);
  const student = String(payload.student || "").trim().replace(/\s+/g, " ");
  const level = Number(payload.level || 0);
  const teacher = String(payload.teacher || "").trim();
  const email = String(payload.email || "").trim().toLowerCase();

  if (!month || !student) throw new Error("Faltan datos: month/student");
  if (!teacher) throw new Error("Falta teacher");
  if (!level || level < 1 || level > 4) throw new Error("level inválido (1..4)");

  const sh = ensurePiercingSheet_();
  const row = piercFindRow_(sh, month, student);
  const now = new Date();

  let who = [];
  let maxLevel = 0;
  let totalMarks = 0;
  let reported = "NO";
  let reportedAt = "";

  if (row) {
    const r = sh.getRange(row, 1, 1, 8).getValues()[0];

    maxLevel = Number(r[2] || 0);
    totalMarks = Number(r[3] || 0);
    reported = String(r[6] || "NO").toUpperCase();
    reportedAt = r[7] || "";

    try { who = JSON.parse(r[4] || "[]"); } catch (e) { who = []; }
    if (!Array.isArray(who)) who = [];
  }

  // idempotencia por docente (email o nombre)
  const idx = who.findIndex(x =>
    String(x.email || "").toLowerCase() === email ||
    String(x.teacher || "").toLowerCase() === teacher.toLowerCase()
  );

  const entry = { teacher, email, level, ts: now.toISOString() };

  if (idx >= 0) {
    const prevLevel = Number(who[idx].level || 0);
    if (level > prevLevel) {
      who[idx] = entry;
      totalMarks += 1;
    }
  } else {
    who.push(entry);
    totalMarks += 1;
  }

  maxLevel = who.reduce((m, x) => Math.max(m, Number(x.level || 0)), 0);

  let justReported = false;
  let wa_url = "";
  let wa_msg = "";

  if (maxLevel >= 4 && reported !== "SI") {
    reported = "SI";
    reportedAt = now.toISOString();
    justReported = true;

    wa_msg =
      `Hola profe,\n\n` +
      `El estudiante "${student}" alcanzó el cuarto llamado de atención por PIERCING o PERFORACIONES.\n\n` +
      `Mes: ${month}\n\n` +
      `Docentes participantes:\n` +
      who.map(x => `- ${x.teacher} (nivel ${x.level})`).join("\n");

    // ✅ Compartir nativo desde el teléfono (no WhatsApp fijo)
wa_url = "";

  }

  const whoJson = JSON.stringify(who);
  const updatedAt = now.toISOString();

  if (row) {
    sh.getRange(row, 1).setNumberFormat("@");
    sh.getRange(row, 1, 1, 8).setValues([[
      "'" + month, student, maxLevel, totalMarks, whoJson, updatedAt, reported, reportedAt
    ]]);
  } else {
    sh.appendRow(["'" + month, student, maxLevel, totalMarks, whoJson, updatedAt, reported, reportedAt]);
    const newRow = sh.getLastRow();
    sh.getRange(newRow, 1).setNumberFormat("@");
  }

  return {
    report: {
      student,
      maxLevel,
      totalMarks,
      who: who.map(x => ({ teacher: x.teacher, level: x.level }))
    },
    list: piercing_list_(month),
    justReported,
    wa_msg,
    wa_url
  };
}

function piercing_delete_student_(payload) {
  const month = normalizeMonthKey_(payload.month);
  const studentRaw = String(payload.student || "").trim().replace(/\s+/g, " ");
  if (!month || !studentRaw) throw new Error("Faltan datos: month/student");

  const sh = ensurePiercingSheet_();
  const last = sh.getLastRow();
  if (last < 2) {
    return { deleted_rows: 0, msg: "No hay registros.", list: piercing_list_(month) };
  }

  const values = sh.getRange(2, 1, last - 1, 2).getValues(); // A:month B:student
  const targetMonth = normalizeMonthKey_(month);
  const targetStudent = normalizeStudentKey_(studentRaw);

  let deleted = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    const rowMonth = normalizeMonthKey_(values[i][0]);
    const rowStudent = normalizeStudentKey_(values[i][1]);

    if (rowMonth === targetMonth && rowStudent === targetStudent) {
      sh.deleteRow(i + 2);
      deleted++;
    }
  }

  const list = piercing_list_(month);

  return {
    deleted_rows: deleted,
    msg: deleted
      ? `Estudiante eliminado ✅ (${deleted} fila(s) borrada(s) en ${targetMonth}).`
      : "No se encontraron filas para eliminar.",
    list
  };
}


/************ PORTE DEL UNIFORME (LLAMADO DE ATENCIÓN) ************/
function ensureUniformeSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("UNIFORME");
  if (!sh) {
    sh = ss.insertSheet("UNIFORME");
    sh.getRange("A1:H1").setValues([[
      "month", "student", "maxLevel", "totalMarks", "who_json", "updatedAt", "reported", "reportedAt"
    ]]);
    sh.getRange("A1:H1").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 8);
  }
  return sh;
}

function uniFindRow_(sh, month, student) {
  const last = sh.getLastRow();
  if (last < 2) return 0;

  const mk = normalizeMonthKey_(month);
  const st = String(student || "").trim().toLowerCase();

  const values = sh.getRange(2, 1, last - 1, 2).getValues(); // month, student
  for (let i = 0; i < values.length; i++) {
    const m = normalizeMonthKey_(values[i][0]);
    const s = String(values[i][1] || "").trim().toLowerCase();
    if (m === mk && s === st) return i + 2;
  }
  return 0;
}

function uniforme_list_(month) {
  const sh = ensureUniformeSheet_();
  const mk = normalizeMonthKey_(month);
  const last = sh.getLastRow();
  if (last < 2) return [];

  const data = sh.getRange(2, 1, last - 1, 8).getValues();
  const out = [];

  data.forEach(r => {
    const m = normalizeMonthKey_(r[0]);
    if (mk && m !== mk) return;

    let who = [];
    try { who = JSON.parse(r[4] || "[]"); } catch (e) { who = []; }
    if (!Array.isArray(who)) who = [];

    const whoClean = who.map(x => ({
      teacher: String(x.teacher || "").trim(),
      email: String(x.email || "").trim(),
      level: Number(x.level || 0),
      ts: x.ts || ""
    })).filter(x => x.teacher);

    const maxLevel = Number(r[2] || 0);
    const totalMarks = Number(r[3] || 0);

    const repCell = r[6];
    const reported =
      (repCell === true) ||
      String(repCell || "").trim().toUpperCase() === "SI" ||
      String(repCell || "").trim().toUpperCase() === "YES";

    out.push({
      month: m,
      student: r[1],
      maxLevel,
      totalMarks,
      teachersCount: whoClean.length,
who: whoClean.map(x => {
  const f = x.ts ? fmtTsBogota_(x.ts) : {date:"", time:"", dt:""};
  return {
    teacher: x.teacher,
    level: x.level,
    date: f.date,
    time: f.time,
    dt: f.dt
  };
}),
      reported,
      reportedAt: r[7] || "",
      updatedAt: r[5] || ""
    });
  });

  // orden: reportados primero, luego maxLevel desc, luego updatedAt desc
  out.sort((a, b) => {
    const ra = a.reported ? 1 : 0;
    const rb = b.reported ? 1 : 0;
    if (ra !== rb) return rb - ra;
    if (a.maxLevel !== b.maxLevel) return b.maxLevel - a.maxLevel;
    return String(b.updatedAt || "").localeCompare(String(a.updatedAt || ""));
  });

  return out;
}

function uniforme_mark_(payload) {
  const month = normalizeMonthKey_(payload.month);
  const student = String(payload.student || "").trim().replace(/\s+/g, " ");
  const level = Number(payload.level || 0);
  const teacher = String(payload.teacher || "").trim();
  const email = String(payload.email || "").trim().toLowerCase();

  if (!month || !student) throw new Error("Faltan datos: month/student");
  if (!teacher) throw new Error("Falta teacher");
  if (!level || level < 1 || level > 4) throw new Error("level inválido (1..4)");

  const sh = ensureUniformeSheet_();
  const row = uniFindRow_(sh, month, student);
  const now = new Date();

  let who = [];
  let maxLevel = 0;
  let totalMarks = 0;
  let reported = "NO";
  let reportedAt = "";

  if (row) {
    const r = sh.getRange(row, 1, 1, 8).getValues()[0];

    maxLevel = Number(r[2] || 0);
    totalMarks = Number(r[3] || 0);
    reported = String(r[6] || "NO").toUpperCase();
    reportedAt = r[7] || "";

    try { who = JSON.parse(r[4] || "[]"); } catch (e) { who = []; }
    if (!Array.isArray(who)) who = [];
  }

  // idempotencia por docente (email o nombre)
  const idx = who.findIndex(x =>
    String(x.email || "").toLowerCase() === email ||
    String(x.teacher || "").toLowerCase() === teacher.toLowerCase()
  );

  const entry = { teacher, email, level, ts: now.toISOString() };

  if (idx >= 0) {
    const prevLevel = Number(who[idx].level || 0);
    if (level > prevLevel) {
      who[idx] = entry;
      totalMarks += 1;
    }
  } else {
    who.push(entry);
    totalMarks += 1;
  }

  maxLevel = who.reduce((m, x) => Math.max(m, Number(x.level || 0)), 0);

  let justReported = false;
  let wa_url = "";
  let wa_msg = "";

  if (maxLevel >= 4 && reported !== "SI") {
    reported = "SI";
    reportedAt = now.toISOString();
    justReported = true;

    wa_msg =
      `Hola profe,\n\n` +
      `El estudiante "${student}" alcanzó el cuarto llamado de atención por PORTE DEL UNIFORME.\n\n` +
      `Mes: ${month}\n\n` +
      `Docentes participantes:\n` +
      who.map(x => `- ${x.teacher} (nivel ${x.level})`).join("\n");

    // ✅ Compartir nativo desde el teléfono (no WhatsApp fijo)
wa_url = "";

  }

  const whoJson = JSON.stringify(who);
  const updatedAt = now.toISOString();

  if (row) {
    sh.getRange(row, 1).setNumberFormat("@");
    sh.getRange(row, 1, 1, 8).setValues([[
      "'" + month, student, maxLevel, totalMarks, whoJson, updatedAt, reported, reportedAt
    ]]);
  } else {
    sh.appendRow(["'" + month, student, maxLevel, totalMarks, whoJson, updatedAt, reported, reportedAt]);
    const newRow = sh.getLastRow();
    sh.getRange(newRow, 1).setNumberFormat("@");
  }

  return {
    report: {
      student,
      maxLevel,
      totalMarks,
who: who.map(x => {
  const f = x.ts ? fmtTsBogota_(x.ts) : {date:"", time:"", dt:""};
  return { teacher: x.teacher, level: x.level, date: f.date, time: f.time, dt: f.dt };
})
    },
    list: uniforme_list_(month),
    justReported,
    wa_msg,
    wa_url
  };
}

function uniforme_delete_student_(payload) {
  const month = normalizeMonthKey_(payload.month);
  const studentRaw = String(payload.student || "").trim().replace(/\s+/g, " ");
  if (!month || !studentRaw) throw new Error("Faltan datos: month/student");

  const sh = ensureUniformeSheet_();
  const last = sh.getLastRow();
  if (last < 2) {
    return { deleted_rows: 0, msg: "No hay registros.", list: uniforme_list_(month) };
  }

  const values = sh.getRange(2, 1, last - 1, 2).getValues(); // A:month B:student
  const targetMonth = normalizeMonthKey_(month);
  const targetStudent = normalizeStudentKey_(studentRaw);

  let deleted = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    const rowMonth = normalizeMonthKey_(values[i][0]);
    const rowStudent = normalizeStudentKey_(values[i][1]);

    if (rowMonth === targetMonth && rowStudent === targetStudent) {
      sh.deleteRow(i + 2);
      deleted++;
    }
  }

  const list = uniforme_list_(month);

  return {
    deleted_rows: deleted,
    msg: deleted
      ? `Estudiante eliminado ✅ (${deleted} fila(s) borrada(s) en ${targetMonth}).`
      : "No se encontraron filas para eliminar.",
    list
  };
}


/************ ENDPOINTS ************/

// GET: ?action=list&month=YYYY-MM&limit=50&key=...
function doGet(e) {
  if (!requireKey_(e)) return json_({ ok: false, error: "unauthorized" });

  const action = (e.parameter.action || "list").toLowerCase();

  // --- ACCIÓN: LISTAR ---
  if (action === "list") {
    const limit = Number(e.parameter.limit || 50);
    const month = e.parameter.month || ""; // "2026-01"
    const now = new Date();
    const d = month ? new Date(`${month}-01T00:00:00`) : now;

    const sh = ensureMonthSheet_(d);
    const records = getLatestRecords_(sh, limit);
    return json_({ ok: true, sheet: sh.getName(), records });
  }

  // --- ACCIÓN: EXPORTAR XLSX (DESCARGA DIRECTA) ---
  if (action === "export_xlsx") {
    const month = String(e.parameter.month || "").trim();
    const teacher = String(e.parameter.teacher || "__ALL__").trim();

    const sh = getMonthSheetByKey_(month);
    if (!sh) {
      return ContentService
        .createTextOutput("No existe la hoja de ese mes.")
        .setMimeType(ContentService.MimeType.TEXT);
    }

    let blob;
    if (!teacher || teacher === "__ALL__") {
      // Exporta hoja completa
      blob = exportSheetAsXlsxBlob_(sh, `asistencia_${sh.getName()}.xlsx`);
    } else {
      // Exporta filtrado por docente
      blob = exportTeacherAsXlsxBlob_(sh, teacher);
    }
    
    return forceDownloadXlsx_(blob, blob.getName());
  }

  // --- ACCIÓN: EXPORTAR XLSX (LINK DE DRIVE) ---
  if (action === "export_xlsx_link") {
    const month = String(e.parameter.month || "").trim();
    const teacher = String(e.parameter.teacher || "__ALL__").trim();

    const sh = getMonthSheetByKey_(month);
    if (!sh) return json_({ ok: false, msg: "No existe la hoja de ese mes." });

    let blob;
    if (!teacher || teacher === "__ALL__") {
      blob = exportSheetAsXlsxBlob_(sh, `asistencia_${sh.getName()}.xlsx`);
    } else {
      blob = exportTeacherAsXlsxBlob_(sh, teacher);
    }

    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const id = file.getId();
    const dl = `https://drive.google.com/uc?export=download&id=${id}`;

    return json_({ ok: true, url: dl, file_id: id });
  }

  // --- ACCIÓN: PING ---
  if (action === "ping") {
    return json_({ ok: true, msg: "pong" });
  }

  return json_({ ok: false, error: "unknown_action" });
}



/************ CURSOS (HOJAS) + ESTUDIANTES ************/

// ✅ Convierte cualquier nombre/curso a una clave estándar: "6°A"
function normalizeCourseKey_(s) {
  return String(s || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function normalizeSubjectKey_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

// ✅ Encuentra la hoja aunque tenga espacios o variaciones ("6° A", "6 °a", etc.)
function findCourseSheet_(ss, courseKey) {
  const target = normalizeCourseKey_(courseKey);
  if (!target) return null;

  const sheets = ss.getSheets();
  for (const sh of sheets) {
    const key = normalizeCourseKey_(sh.getName());
    if (key === target) return sh;
  }
  return null;
}


function isCourseSheetName_(name) {
  // ✅ AHORA: acepta cualquier hoja con nombre válido (no vacío)
  const s = String(name || "").trim();
  if (!s) return false;

  // opcional: ignora hojas “técnicas”
  const upper = s.toUpperCase();
  if (upper === "README" || upper === "CONFIG") return false;

  return true;
}

function listCourseSheets_() {
  const ss = getStudentsSS_();
  const out = ss.getSheets()
    .map(sh => String(sh.getName() || "").trim())
    .filter(isCourseSheetName_);

  const uniq = Array.from(new Set(out));
  uniq.sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base" }));
  return uniq;
}

function listStudentsInCourse_(sheetName) {
  const ss = getStudentsSS_();
  const name = String(sheetName || "").trim();
  if (!name) return [];

  const sh = ss.getSheetByName(name);
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const startRow = 2;
  if (lastRow < startRow) return [];

  // ✅ Si hay 2+ columnas: Nombre + Apellido (A y B)
  // ✅ Si solo hay 1 columna: usa A completa
  const numCols = Math.max(1, Math.min(2, lastCol));
  const values = sh.getRange(startRow, 1, lastRow - startRow + 1, numCols).getValues();

  const names = values.map(r => {
    const a = String(r[0] || "").trim();
    const b = (numCols >= 2) ? String(r[1] || "").trim() : "";
    return (a + " " + b).trim();
  }).filter(Boolean);

  const uniq = Array.from(new Set(names));
  uniq.sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base" }));
  return uniq;
}

/************ FALTAS (Tipo 1/2/3) - EXTERNO ************/
function ensureFaltasSheet_() {
  const ss = getFaltasSS_();
  let sh = ss.getSheetByName("FALTAS");
  if (!sh) {
    sh = ss.insertSheet("FALTAS");
    sh.getRange("A1:K1").setValues([[
      "month", "tipo", "course", "student", "subject",
      "teacher", "email", "device_id",
      "ts", "date", "time"
    ]]);
    sh.getRange("A1:K1").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 11);
  }
  return sh;
}

function faltas_report_(payload) {
  const month   = normalizeMonthKey_(payload.month);
  const tipo    = Number(payload.tipo || 1);
  const course  = String(payload.course || "").trim();
  const student = String(payload.student || "").trim().replace(/\s+/g, " ");
  const subject = String(payload.subject || "").trim();
  const teacher = String(payload.teacher || "").trim();
  const email   = String(payload.email || "").trim().toLowerCase();
  const device  = String(payload.device_id || "").trim();

  if (!month) return { ok:false, msg:"Falta month (YYYY-MM)." };
  if (![1,2,3].includes(tipo)) return { ok:false, msg:"tipo inválido (1/2/3)." };
  if (!course) return { ok:false, msg:"Falta course." };
  if (!student) return { ok:false, msg:"Falta student." };
  if (!subject) return { ok:false, msg:"Falta subject." };
  if (!teacher) return { ok:false, msg:"Falta teacher." };

  const sh = ensureFaltasSheet_();
  const now = new Date();
  const ts = now.toISOString();

  const date = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd");
  const time = Utilities.formatDate(now, TIMEZONE, "HH:mm:ss");

  // guardamos month como texto (igual que los otros)
  sh.appendRow([
    "'" + month, tipo, course, student, subject,
    teacher, email, device,
    ts, date, time
  ]);

  return { ok:true };
}

// Lista reportes del mes. Opcional: filtrar por tipo
// Lista reportes del mes (y opcional tipo) pero AGRUPADO por estudiante,
// devolviendo who[] como Disciplina/Piercing/Uniforme
function faltas_list_(payload) {
  const month = normalizeMonthKey_(payload.month);
  const tipo = payload.tipo ? Number(payload.tipo) : 0;

  const sh = ensureFaltasSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];

  const data = sh.getRange(2, 1, last - 1, 11).getValues();
  // cols:
  // 0 month,1 tipo,2 course,3 student,4 subject,5 teacher,6 email,7 device,8 ts,9 date,10 time

  const map = {};

  data.forEach(r => {
    const m = normalizeMonthKey_(r[0]);
    if (month && m !== month) return;

    const t = Number(r[1] || 0);
    if (tipo && t !== tipo) return;

    const course  = String(r[2] || "").trim();
    const student = String(r[3] || "").trim();
    const subject = String(r[4] || "").trim();
    const teacher = String(r[5] || "").trim();
    const ts      = String(r[8] || "").trim();

    if (!student) return;

    const date = String(r[9] || "").trim();
    const time = String(r[10] || "").trim();
    const dt = (date && time) ? `${date} ${time}` : (date || "");

    const key = `${m}__${t}__${normalizeStudentKey_(student)}`;

    if (!map[key]) {
      map[key] = {
        month: m,
        tipo: t,
        course,
        student,
        subject: "",
        subjects: [],
        subjectText: "",
        totalFaltas: 0,
        faltas: [],
        teachersCount: 0,
        who: [],
        updatedAt: ts || dt || ""
      };
    }

    if (course) map[key].course = course;

    // ✅ guardar TODAS las faltas
    if (subject) {
      map[key].faltas.push({
        subject,
        teacher,
        dt,
        ts: ts || ""
      });

      const alreadySubject = map[key].subjects.some(x =>
        String(x || "").trim().toLowerCase() === subject.toLowerCase()
      );
      if (!alreadySubject) {
        map[key].subjects.push(subject);
      }
    }

    map[key].totalFaltas = map[key].faltas.length;
    map[key].subject = map[key].subjects[0] || "";
    map[key].subjectText = map[key].subjects.join(" • ");

    // docentes únicos
    if (teacher) {
      const exists = map[key].who.some(x =>
        String(x.teacher || "").trim().toLowerCase() === teacher.toLowerCase()
      );

      if (!exists) {
        map[key].who.push({
          teacher,
          dt: dt || ""
        });
      }
    }

    // updatedAt más reciente
    const curr = Date.parse(map[key].updatedAt || "");
    const next = Date.parse(ts || "");
    if (!isNaN(next) && (isNaN(curr) || next > curr)) {
      map[key].updatedAt = ts;
    }
  });

  const out = Object.values(map);

  out.forEach(it => {
    it.teachersCount = Array.isArray(it.who) ? it.who.length : 0;

    if (Array.isArray(it.faltas)) {
      it.faltas.sort((a, b) => {
        const ta = Date.parse(a.ts || "");
        const tb = Date.parse(b.ts || "");
        if (!isNaN(ta) && !isNaN(tb)) return tb - ta;
        return String(b.dt || "").localeCompare(String(a.dt || ""));
      });
    }
  });

  out.sort((a, b) => {
    const ta = Date.parse(a.updatedAt || "");
    const tb = Date.parse(b.updatedAt || "");
    if (!isNaN(ta) && !isNaN(tb)) return tb - ta;
    return String(b.updatedAt || "").localeCompare(String(a.updatedAt || ""));
  });

  return out;
}


function faltas_delete_(payload) {

  const month   = normalizeMonthKey_(payload.month);
  const tipo    = Number(payload.tipo || 0);
  const student = String(payload.student || "").trim().replace(/\s+/g," ");

  if(!month) return { ok:false, msg:"Falta month." };
  if(!tipo) return { ok:false, msg:"Falta tipo." };
  if(!student) return { ok:false, msg:"Falta student." };

  const sh = ensureFaltasSheet_();
  const last = sh.getLastRow();
  if(last < 2) return { ok:true };

  const data = sh.getRange(2,1,last-1,11).getValues();

  for(let i=data.length-1;i>=0;i--){

    const r = data[i];

    const rMonth   = String(r[0] || "").replace("'","");
    const rTipo    = Number(r[1] || 0);
    const rStudent = String(r[3] || "").trim();

    if(
      rMonth === month &&
      rTipo === tipo &&
      rStudent === student
    ){
      sh.deleteRow(i+2);
    }

  }

  return { ok:true };

}

/************ REPORTES (Actividades pendientes) ************/
function ensureReportesSheet_() {
  const ss = getFaltasSS_();
  let sh = ss.getSheetByName("REPORTES");
  if (!sh) {
    sh = ss.insertSheet("REPORTES");
  }

  const headers = [
    "month", "tipo", "course", "subject", "student",
    "actividad", "nota",
    "teacher", "email", "device_id",
    "ts", "date", "time",
    "reported", "reportedAt"
  ];

  const lastCol = Math.max(sh.getLastColumn(), 1);
  const currentHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x || "").trim());

  if (!currentHeaders.length || !currentHeaders[0]) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);
    return sh;
  }

  if (currentHeaders.indexOf("subject") === -1) {
    sh.insertColumnAfter(3);
    sh.getRange(1, 4).setValue("subject");
  }

  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight("bold");
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, sh.getLastColumn());

  return sh;
}

function reportes_pendientes_guardar_(payload) {
const month   = normalizeMonthKey_(payload.month);
const tipo    = String(payload.tipo || "actividades_pendientes").trim();
const course  = String(payload.course || "").trim();
const subject = String(payload.subject || "").trim();
const student = String(payload.student || "").trim().replace(/\s+/g, " ");
const teacher = String(payload.teacher || "").trim();
const email   = String(payload.email || "").trim().toLowerCase();
const device  = String(payload.device_id || "").trim();

  const actividades = Array.isArray(payload.actividades) ? payload.actividades : [];

if (!month)   return { ok:false, msg:"Falta month (YYYY-MM)." };
if (!course)  return { ok:false, msg:"Falta course." };
if (!subject) return { ok:false, msg:"Falta subject." };
if (!student) return { ok:false, msg:"Falta student." };
if (!teacher) return { ok:false, msg:"Falta teacher." };
if (!actividades.length) return { ok:false, msg:"Debes enviar al menos una actividad." };

  const limpias = actividades
    .map(x => ({
      actividad: String(x && x.actividad || "").trim(),
      nota: String(x && x.nota || "").trim()
    }))
    .filter(x => x.actividad);

  if (!limpias.length) {
    return { ok:false, msg:"No hay actividades válidas para guardar." };
  }

  const sh = ensureReportesSheet_();
  const now = new Date();
  const ts = now.toISOString();
  const date = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd");
  const time = Utilities.formatDate(now, TIMEZONE, "HH:mm:ss");

const rows = limpias.map(x => ([
  "'" + month, tipo, course, subject, student,
  x.actividad, x.nota,
  teacher, email, device,
  ts, date, time,
  "NO", ""
]));

 sh.getRange(sh.getLastRow() + 1, 1, rows.length, 15).setValues(rows);

  return { ok:true };
}


function reportes_pendientes_list_(payload) {
  const month = normalizeMonthKey_(payload.month);
  const tipo = String(payload.tipo || "actividades_pendientes").trim();

  const sh = ensureReportesSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];

  const data = sh.getRange(2, 1, last - 1, 15).getValues();
  // 0 month,1 tipo,2 course,3 subject,4 student,5 actividad,6 nota,7 teacher,8 email,9 device_id,10 ts,11 date,12 time,13 reported,14 reportedAt

  const map = {};

  data.forEach(r => {
    const m = normalizeMonthKey_(r[0]);
    if (month && m !== month) return;

    const t = String(r[1] || "").trim();
    if (tipo && t !== tipo) return;

    const course    = String(r[2] || "").trim();
    const subject   = String(r[3] || "").trim();
    const student   = String(r[4] || "").trim();
    const actividad = String(r[5] || "").trim();
    const nota      = String(r[6] || "").trim();
    const teacher   = String(r[7] || "").trim();
    const ts        = String(r[10] || "").trim();
    const date      = String(r[11] || "").trim();
    const time      = String(r[12] || "").trim();
    const dt        = (date && time) ? `${date} ${time}` : (date || "");

    if (!student) return;

    const key = [
      m,
      t,
      normalizeCourseKey_(course),
      normalizeSubjectKey_(subject),
      normalizeStudentKey_(student)
    ].join("__");

    if (!map[key]) {
      map[key] = {
        month: m,
        tipo: t,
        course,
        subject,
        student,
        totalPendientes: 0,
        actividades: [],
        who: [],
        updatedAt: ts || dt || ""
      };
    }

    if (course) map[key].course = course;
    if (subject) map[key].subject = subject;

    if (actividad) {
      map[key].actividades.push({
        actividad,
        nota,
        dt
      });
    }

    if (teacher) {
      const exists = map[key].who.some(x =>
        String(x.teacher || "").trim().toLowerCase() === teacher.toLowerCase()
      );

      if (!exists) {
        map[key].who.push({
          teacher,
          dt
        });
      }
    }

    const curr = Date.parse(map[key].updatedAt || "");
    const next = Date.parse(ts || "");
    if (!isNaN(next) && (isNaN(curr) || next > curr)) {
      map[key].updatedAt = ts;
    }
  });

  const out = Object.values(map);

  out.forEach(it => {
    it.totalPendientes = Array.isArray(it.actividades) ? it.actividades.length : 0;
    it.teachersCount = Array.isArray(it.who) ? it.who.length : 0;
  });

  out.sort((a, b) => {
    const ta = Date.parse(a.updatedAt || "");
    const tb = Date.parse(b.updatedAt || "");
    if (!isNaN(ta) && !isNaN(tb)) return tb - ta;
    return String(b.updatedAt || "").localeCompare(String(a.updatedAt || ""));
  });

  return out;
}

function reportes_pendientes_delete_(payload) {
  const month   = normalizeMonthKey_(payload.month);
  const tipo    = String(payload.tipo || "actividades_pendientes").trim();
  const course  = String(payload.course || "").trim().replace(/\s+/g, " ");
  const subject = String(payload.subject || "").trim().replace(/\s+/g, " ");
  const student = String(payload.student || "").trim().replace(/\s+/g, " ");

  if (!month) return { ok:false, msg:"Falta month." };
  if (!tipo) return { ok:false, msg:"Falta tipo." };
  if (!course) return { ok:false, msg:"Falta course." };
  if (!subject) return { ok:false, msg:"Falta subject." };
  if (!student) return { ok:false, msg:"Falta student." };

  const sh = ensureReportesSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { ok:true };

  const data = sh.getRange(2, 1, last - 1, 15).getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];

    const rMonth   = normalizeMonthKey_(r[0]);
    const rTipo    = String(r[1] || "").trim();
    const rCourse  = String(r[2] || "").trim();
    const rSubject = String(r[3] || "").trim();
    const rStudent = String(r[4] || "").trim();

    if (
      rMonth === month &&
      rTipo === tipo &&
      normalizeCourseKey_(rCourse) === normalizeCourseKey_(course) &&
      normalizeSubjectKey_(rSubject) === normalizeSubjectKey_(subject) &&
      normalizeStudentKey_(rStudent) === normalizeStudentKey_(student)
    ) {
      sh.deleteRow(i + 2);
    }
  }

  return { ok:true };
}


/************ HORARIO DE PISOS ************/
function ensureHorarioPisosSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("HORARIO_PISOS");

  if (!sh) {
    sh = ss.insertSheet("HORARIO_PISOS");
    sh.getRange("A1:D1").setValues([[
      "PISO", "DESCANSO", "DIA", "DOCENTE"
    ]]);
    sh.getRange("A1:D1").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 4);
  }

  return sh;
}

function normalizePisosDay_(s) {
  const v = String(s || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

  const map = {
    "LUNES": "LUNES",
    "MARTES": "MARTES",
    "MIERCOLES": "MIERCOLES",
    "JUEVES": "JUEVES",
    "VIERNES": "VIERNES"
  };

  return map[v] || "";
}

function normalizeBreakLabelPisos_(s) {
  const v = String(s || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

  if (v === "1ER DESCANSO" || v === "PRIMER DESCANSO" || v === "1 DESCANSO") return "1ER DESCANSO";
  if (v === "2DO DESCANSO" || v === "SEGUNDO DESCANSO" || v === "2 DESCANSO") return "2DO DESCANSO";
  return v;
}

function floorOrderValuePisos_(floor) {
  const v = String(floor || "").toUpperCase();
  if (v.indexOf("1") > -1) return 1;
  if (v.indexOf("2") > -1) return 2;
  if (v.indexOf("3") > -1) return 3;
  if (v.indexOf("4") > -1) return 4;
  if (v.indexOf("5") > -1) return 5;
  return 99;
}

function breakOrderValuePisos_(label) {
  const v = normalizeBreakLabelPisos_(label);
  if (v === "1ER DESCANSO") return 1;
  if (v === "2DO DESCANSO") return 2;
  return 99;
}

function pisos_days_list_() {
  const sh = ensureHorarioPisosSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 3, lastRow - 1, 1).getValues().flat();

  const orderMap = {
    "LUNES": 1,
    "MARTES": 2,
    "MIERCOLES": 3,
    "JUEVES": 4,
    "VIERNES": 5
  };

  const days = Array.from(new Set(
    values
      .map(x => normalizePisosDay_(x))
      .filter(Boolean)
  )).sort((a, b) => (orderMap[a] || 99) - (orderMap[b] || 99));

  return days;
}

function pisos_schedule_(dayRaw) {
  const day = normalizePisosDay_(dayRaw);
  if (!day) return [];

  const sh = ensureHorarioPisosSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, 4).getValues();

  return data
    .map(r => ({
      floor: String(r[0] || "").trim(),
      break_label: normalizeBreakLabelPisos_(r[1]),
      day_label: normalizePisosDay_(r[2]),
      teacher: String(r[3] || "").trim()
    }))
    .filter(x =>
      x.floor &&
      x.break_label &&
      x.day_label &&
      x.teacher &&
      x.day_label === day
    )
    .sort((a, b) => {
      const floorDiff = floorOrderValuePisos_(a.floor) - floorOrderValuePisos_(b.floor);
      if (floorDiff !== 0) return floorDiff;
      return breakOrderValuePisos_(a.break_label) - breakOrderValuePisos_(b.break_label);
    });
}


/************ PARQUE ************/
function ensureHorarioParqueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("HORARIO_PARQUE");

  if (!sh) {
    sh = ss.insertSheet("HORARIO_PARQUE");
    sh.getRange("A1:C1").setValues([[
      "DESCANSO", "DIA", "CURSO"
    ]]);
    sh.getRange("A1:C1").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 3);
  }

  return sh;
}

function normalizeParqueDay_(s) {
  return String(s || "")
    .trim()
    .toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function normalizeBreakLabelParque_(s) {
  const v = String(s || "")
    .trim()
    .toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");

  if (v.indexOf("1") > -1) return "1ER DESCANSO";
  if (v.indexOf("2") > -1) return "2DO DESCANSO";
  return v;
}

function breakOrderValueParque_(label) {
  const v = normalizeBreakLabelParque_(label);
  if (v === "1ER DESCANSO") return 1;
  if (v === "2DO DESCANSO") return 2;
  return 99;
}

function parque_days_list_() {
  const sh = ensureHorarioParqueSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 2, lastRow - 1, 1).getValues().flat();

  const orderMap = {
    "LUNES": 1,
    "MARTES": 2,
    "MIERCOLES": 3,
    "JUEVES": 4,
    "VIERNES": 5
  };

  const days = Array.from(new Set(
    values
      .map(x => normalizeParqueDay_(x))
      .filter(Boolean)
  )).sort((a, b) => (orderMap[a] || 99) - (orderMap[b] || 99));

  return days;
}

function parque_schedule_(dayRaw) {
  const day = normalizeParqueDay_(dayRaw);
  if (!day) return [];

  const sh = ensureHorarioParqueSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, 3).getValues();

  return data
    .map(r => ({
      break_label: normalizeBreakLabelParque_(r[0]),
      day_label: normalizeParqueDay_(r[1]),
      course: String(r[2] || "").trim()
    }))
    .filter(x =>
      x.break_label &&
      x.day_label &&
      x.course &&
      x.day_label === day
    )
    .sort((a, b) => breakOrderValueParque_(a.break_label) - breakOrderValueParque_(b.break_label));
}

/************ UBICACION DOCENTES ************/
function ensureHorariosDocentesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("HORARIOS_DOCENTES");

  if (!sh) {
    sh = ss.insertSheet("HORARIOS_DOCENTES");
    sh.getRange("A1:G1").setValues([[
      "DOCENTE", "DIA", "HORA_INICIO", "HORA_FIN", "MATERIA", "CURSO", "SALON"
    ]]);
    sh.getRange("A1:G1").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 7);
  }

  return sh;
}

function normalizeTeacherName_(s) {
  return String(s || "")
    .trim()
    .toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function normalizeDayNameEs_(d) {
  const days = ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"];
  return days[d.getDay()] || "";
}

function normalizeDayCell_(s) {
  return String(s || "")
    .trim()
    .toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function normalizeHourText_(value) {
  if (value === null || value === undefined || value === "") return "";

  // Caso 1: viene como Date desde Google Sheets
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TIMEZONE, "HH:mm");
  }

  // Caso 2: viene como número decimal de Sheets (fracción del día)
  if (typeof value === "number" && isFinite(value)) {
    const totalMinutes = Math.round(value * 24 * 60);
    const hh = Math.floor(totalMinutes / 60) % 24;
    const mm = totalMinutes % 60;
    return `${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}`;
  }

  // Caso 3: viene como texto: 7:15, 07:15, 7:15:00...
  let s = String(value).trim();

  // si viene con segundos, quitarlos
  const m1 = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (m1) {
    return `${String(m1[1]).padStart(2, "0")}:${m1[2]}`;
  }

  return "";
}

function hourToMinutes_(value) {
  const hhmm = normalizeHourText_(value);
  const m = hhmm.match(/^(\d{2}):(\d{2})$/);
  if (!m) return -1;
  return Number(m[1]) * 60 + Number(m[2]);
}

function currentBogotaDate_() {
  const now = new Date();
  const txt = Utilities.formatDate(now, TIMEZONE, "yyyy/MM/dd HH:mm:ss");
  return new Date(txt);
}

function ubicacion_docentes_list_() {
  const sh = ensureHorariosDocentesSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  const teachers = Array.from(new Set(
    values
      .map(x => String(x || "").trim())
      .filter(Boolean)
  ));

  teachers.sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base" }));
  return teachers;
}

function ubicacion_docente_now_(teacherRaw) {

  const teacher = String(teacherRaw || "").trim();

  const sh = ensureHorariosDocentesSheet_();
  const lastRow = sh.getLastRow();

  const now = new Date();
  const days = ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"];
  const dayNow = days[now.getDay()];

  const timeNow = Utilities.formatDate(now, TIMEZONE, "HH:mm");
  const timeNowMin = hourToMinutes_(timeNow);

  const data = sh.getRange(2,1,lastRow-1,7).getValues();

  const teacherKey = normalizeTeacherName_(teacher);

  let match = null;

  data.forEach(r => {

    const docente = String(r[0] || "").trim();
    const dia = String(r[1] || "").trim().toUpperCase();

    const horaInicio = normalizeHourText_(r[2]);
    const horaFin = normalizeHourText_(r[3]);

    const materia = String(r[4] || "").trim();
    const curso = String(r[5] || "").trim();
    const salon = String(r[6] || "").trim();

    if (!docente || !dia || !horaInicio || !horaFin) return;

    if (normalizeTeacherName_(docente) !== teacherKey) return;

    if (dia !== dayNow) return;

    const ini = hourToMinutes_(horaInicio);
    const fin = hourToMinutes_(horaFin);

    if (timeNowMin >= ini && timeNowMin < fin) {

      match = {
        teacher: docente,
        found_now: true,
        day_label: dia,
        time_now: timeNow,
        start_time: horaInicio,
        end_time: horaFin,
        subject: materia,
        course: curso,
        salon: salon
      };

    }

  });

  if (match) return match;

  return {
    teacher,
    found_now:false,
    day_label:dayNow,
    time_now:timeNow
  };

}


function dayOrderMap_() {
  return {
    "LUNES": 1,
    "MARTES": 2,
    "MIERCOLES": 3,
    "JUEVES": 4,
    "VIERNES": 5,
    "SABADO": 6,
    "DOMINGO": 7
  };
}

function getTodayDayEs_() {
  const now = new Date();
  const days = ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"];
  return days[now.getDay()];
}

function getNowHHMM_() {
  return Utilities.formatDate(new Date(), TIMEZONE, "HH:mm");
}

function horariosDocenteRows_(teacherRaw) {
  const teacher = String(teacherRaw || "").trim();
  const sh = ensureHorariosDocentesSheet_();
  const lastRow = sh.getLastRow();

  if (!teacher || lastRow < 2) return [];

  const teacherKey = normalizeTeacherName_(teacher);
  const data = sh.getRange(2, 1, lastRow - 1, 7).getValues();

  return data
    .map(r => ({
      teacher: String(r[0] || "").trim(),
      day_label: String(r[1] || "").trim().toUpperCase(),
      start_time: normalizeHourText_(r[2]),
      end_time: normalizeHourText_(r[3]),
      subject: String(r[4] || "").trim(),
      course: String(r[5] || "").trim(),
      salon: String(r[6] || "").trim()
    }))
    .filter(x =>
      x.teacher &&
      x.day_label &&
      x.start_time &&
      x.end_time &&
      normalizeTeacherName_(x.teacher) === teacherKey
    )
    .sort((a, b) => {
      const dm = dayOrderMap_();
      const dayDiff = (dm[a.day_label] || 99) - (dm[b.day_label] || 99);
      if (dayDiff !== 0) return dayDiff;
      return hourToMinutes_(a.start_time) - hourToMinutes_(b.start_time);
    });
}

function ubicacion_docente_schedule_(teacherRaw) {
  const teacher = String(teacherRaw || "").trim();
  const items = horariosDocenteRows_(teacher);

  return {
    teacher,
    items
  };
}

function ubicacion_docente_next_class_(teacherRaw) {
  const teacher = String(teacherRaw || "").trim();
  const items = horariosDocenteRows_(teacher);

  if (!items.length) {
    return {
      teacher,
      found_next: false
    };
  }

  const today = getTodayDayEs_();
  const nowMin = hourToMinutes_(getNowHHMM_());
  const dm = dayOrderMap_();
  const todayOrder = dm[today] || 99;

  let nextItem = null;

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const itemDayOrder = dm[item.day_label] || 99;
    const itemStart = hourToMinutes_(item.start_time);

    if (itemDayOrder > todayOrder) {
      nextItem = item;
      break;
    }

    if (itemDayOrder === todayOrder && itemStart > nowMin) {
      nextItem = item;
      break;
    }
  }

  if (!nextItem) {
    return {
      teacher,
      found_next: false
    };
  }

  return Object.assign({
    teacher,
    found_next: true
  }, nextItem);
}

/************ COMPRAS (Horario de compra) ************/
function ensureComprasSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("COMPRAS");
  if (!sh) {
    sh = ss.insertSheet("COMPRAS");
    sh.getRange(1, 1, 1, 3).setValues([["DAY", "COURSE", "TIME"]]);
    sh.getRange(1, 1, 1, 3).setFontWeight("bold");
    sh.setFrozenRows(1);

    const seed = [
      ["1 DESCANSO", "6°A",  "09:25"],
      ["1 DESCANSO", "6°B",  "09:25"],
      ["1 DESCANSO", "7°",   "09:30"],
      ["1 DESCANSO", "8°",   "09:35"],
      ["1 DESCANSO", "9°A",  "09:40"],
      ["1 DESCANSO", "10°",  "09:45"],
      ["1 DESCANSO", "11°A", "09:50"],
      ["1 DESCANSO", "11°B", "09:50"],

      ["2 DESCANSO", "6°A",  "11:45"],
      ["2 DESCANSO", "6°B",  "11:45"],
      ["2 DESCANSO", "7°",   "11:50"],
      ["2 DESCANSO", "8°",   "11:55"],
      ["2 DESCANSO", "9°A",  "12:00"],
      ["2 DESCANSO", "10°",  "12:05"],
      ["2 DESCANSO", "11°A", "12:10"],
      ["2 DESCANSO", "11°B", "12:10"]
    ];

    sh.getRange(2, 1, seed.length, 3).setValues(seed);
    sh.autoResizeColumns(1, 3);
  }
  return sh;
}

function comprasNormalizeText_(v) {
  return String(v || "").trim().toUpperCase().replace(/\s+/g, " ");
}

function comprasNormalizeCourse_(v) {
  return String(v || "").trim().toUpperCase().replace(/\s+/g, "");
}

function comprasNormalizeTime_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, TIMEZONE, "HH:mm");
  }

  if (typeof v === "number" && isFinite(v)) {
    const totalMinutes = Math.round(v * 24 * 60);
    const hh = String(Math.floor(totalMinutes / 60)).padStart(2, "0");
    const mm = String(totalMinutes % 60).padStart(2, "0");
    return `${hh}:${mm}`;
  }

  let s = String(v || "").trim();
  if (!s) return "";

  s = s.replace(/\./g, ":").replace(/\s+/g, "");

  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return s;

  const hh = Number(m[1]);
  const mm = Number(m[2]);
  if (isNaN(hh) || isNaN(mm)) return s;

  return `${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}`;
}

function compras_days_list_() {
  const sh = ensureComprasSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];

  const values = sh.getRange(2, 1, last - 1, 1).getValues(); // col A = DAY
  const seen = {};
  const days = [];

  values.forEach(r => {
    const day = comprasNormalizeText_(r[0]);
    if (!day) return;
    if (seen[day]) return;
    seen[day] = true;
    days.push(day);
  });

  const order = { "1 DESCANSO": 1, "2 DESCANSO": 2 };
  days.sort((a, b) => (order[a] || 99) - (order[b] || 99));

  return days;
}

function compras_schedule_(day) {
  const sh = ensureComprasSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];

  const dayNorm = comprasNormalizeText_(day);
  const data = sh.getRange(2, 1, last - 1, 3).getValues(); // A:C => day, course, time

  const items = [];

  data.forEach(r => {
    const itemDay = comprasNormalizeText_(r[0]);
    const course = comprasNormalizeCourse_(r[1]);
    const time = comprasNormalizeTime_(r[2]);

    if (!itemDay || !course || !time) return;
    if (dayNorm && itemDay !== dayNorm) return;

    items.push({
      day: itemDay,
      course,
      time
    });
  });

  items.sort((a, b) => a.time.localeCompare(b.time, "es"));
  return items;
}


// POST JSON: {action:"register", teacher, type, iso, distance_m, accuracy_m, lat, lng, key}
function doPost(e) {
  if (!requireKey_(e)) return json_({ ok:false, error:"unauthorized" });

  let body = {};
  try { body = JSON.parse(e.postData.contents || "{}"); }
  catch(err) { return json_({ ok:false, error:"invalid_json" }); }

const action = String(body.action || "").toLowerCase();

// =========================
// CURSOS: LISTAR HOJAS (6°A, 6°B, ...)
// body: { action:"courses_list" }
// =========================
if (action === "courses_list") {
  try {
    const courses = listCourseSheets_();
    return json_({ ok:true, courses });
  } catch (err) {
    return json_({ ok:false, msg: String(err && err.message ? err.message : err) });
  }
}


if (action === "faltas_delete") {
  try {
    const resp = faltas_delete_(body);
    return json_(resp);
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}


// =========================
// REPORTES: GUARDAR ACTIVIDADES PENDIENTES
// body: { action:"reportes_pendientes_guardar", month, tipo, course, student, actividades[], teacher, email, device_id }
// =========================
if (action === "reportes_pendientes_guardar") {
  try {
    const resp = reportes_pendientes_guardar_(body);
    return json_(resp);
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}

// =========================
// REPORTES: LISTAR ACTIVIDADES PENDIENTES
// body: { action:"reportes_pendientes_list", month:"YYYY-MM", tipo?:"actividades_pendientes" }
// =========================
if (action === "reportes_pendientes_list") {
  try {
    const items = reportes_pendientes_list_(body);
    return json_({ ok:true, items });
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}

// =========================
// REPORTES: ELIMINAR POR ESTUDIANTE
// body: { action:"reportes_pendientes_delete", month:"YYYY-MM", tipo?:"actividades_pendientes", student:"Nombre" }
// =========================
if (action === "reportes_pendientes_delete") {
  try {
    const resp = reportes_pendientes_delete_(body);
    return json_(resp);
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}

// =========================
// CURSOS: LISTAR ESTUDIANTES DE UN CURSO
// body: { action:"course_students", course:"6°A" }
// =========================
if (action === "course_students") {
  try {
    const course = String(body.course || "").trim();
    const students = listStudentsInCourse_(course);
    return json_({ ok:true, course, students });
  } catch (err) {
    return json_({ ok:false, msg: String(err && err.message ? err.message : err) });
  }
}


if (action === "pisos_days_list") {
  try {
    const days = pisos_days_list_();
    return json_({ ok:true, days });
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}

if (action === "pisos_schedule") {
  try {
    const items = pisos_schedule_(body.day);
    return json_({ ok:true, items });
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}

if (action === "parque_days_list") {
  try {
    const days = parque_days_list_();
    return json_({ ok:true, days });
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}

if (action === "parque_schedule") {
  try {
    const items = parque_schedule_(body.day);
    return json_({ ok:true, items });
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}


if (action === "compras_days_list") {
  try {
    const days = compras_days_list_();
    return json_({ ok:true, days });
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}

if (action === "compras_schedule") {
  try {
    const items = compras_schedule_(body.day);
    return json_({ ok:true, items });
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}

// =========================
// ASIGNATURAS
// =========================

if (action === "asignaturas_get") {
  try {
    const items = asignaturas_get_(body.device_id);
    return json_({ ok:true, items });
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}

if (action === "asignaturas_save") {
  try {
    const ok = asignaturas_save_(body.device_id, body.items);
    return json_({ ok:true, saved: !!ok });
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}


if (action === "ubicacion_docente_next_class") {
  try {
    const resp = ubicacion_docente_next_class_(body.teacher);
    return json_(Object.assign({ ok:true }, resp));
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}

if (action === "ubicacion_docente_schedule") {
  try {
    const resp = ubicacion_docente_schedule_(body.teacher);
    return json_(Object.assign({ ok:true }, resp));
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}
// =========================
// UBICATE: LISTAR DOCENTES
// body: { action:"ubicacion_docentes_list" }
// =========================
if (action === "ubicacion_docentes_list") {
  try {
    const teachers = ubicacion_docentes_list_();
    return json_({ ok:true, teachers });
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}


// =========================
// UBICATE: UBICACION ACTUAL DE DOCENTE
// body: { action:"ubicacion_docente_now", teacher:"CAROLINA" }
// =========================
if (action === "ubicacion_docente_now") {
  try {
    const resp = ubicacion_docente_now_(body.teacher);
    return json_(Object.assign({ ok:true }, resp));
  } catch (err) {
    return json_({ ok:false, msg:String(err && err.message ? err.message : err) });
  }
}

// ✅ Asegura que exista la hoja DISPOSITIVOS al primer POST
ensureDevicesSheet_();
ensureAdminsSheet_(); // ✅ crea ADMINISTRADORES en el primer POST

// ✅ Asegura hoja DISCIPLINA
ensureDisciplineSheet_();
// ✅ Asegura hoja PIERCING
ensurePiercingSheet_();

ensureUniformeSheet_(); // ✅ Asegura hoja UNIFORME

ensureReportesSheet_(); // ✅ Asegura hoja REPORTES



// =========================
// DISCIPLINA: LISTAR (MES)
// body: { action:"discipline_list", month:"YYYY-MM" }
// =========================
if (action === "discipline_list") {
  const month = String(body.month || "").trim();
  const items = discipline_list_(month);
  return json_({ ok:true, items });
}


// =========================
// DISCIPLINA: ELIMINAR ESTUDIANTE (DEL MES)
// body: { action:"discipline_delete_student", month:"YYYY-MM", student:"Nombre" }
// =========================
if (action === "discipline_delete_student") {
  try {
    const resp = discipline_delete_student_(body);
    return json_({
      ok: true,
      deleted_rows: resp.deleted_rows,
      msg: resp.msg,
      list: resp.list
    });
  } catch (err) {
    return json_({ ok:false, msg: String(err && err.message ? err.message : err) });
  }
}

// =========================
// DISCIPLINA: MARCAR NIVEL (1..4)
// body: { action:"discipline_mark", month, student, level, teacher, email }
// =========================
if (action === "discipline_mark") {
  try {
    const resp = discipline_mark_(body);
return json_({
  ok: true,
  report: resp.report,
  list: resp.list,
  justReported: resp.justReported,
  wa_msg: resp.wa_msg || "",
  wa_url: resp.wa_url || ""
});



  } catch (err) {
    return json_({ ok:false, msg: String(err && err.message ? err.message : err) });
  }
}

// =========================
// PIERCING: LISTAR (MES)
// body: { action:"piercing_list", month:"YYYY-MM" }
// =========================
if (action === "piercing_list") {
  const month = String(body.month || "").trim();
  const items = piercing_list_(month);
  return json_({ ok:true, items });
}

// =========================
// PIERCING: ELIMINAR ESTUDIANTE (DEL MES)
// body: { action:"piercing_delete_student", month:"YYYY-MM", student:"Nombre" }
// =========================
if (action === "piercing_delete_student") {
  try {
    const resp = piercing_delete_student_(body);
    return json_({
      ok: true,
      deleted_rows: resp.deleted_rows,
      msg: resp.msg,
      list: resp.list
    });
  } catch (err) {
    return json_({ ok:false, msg: String(err && err.message ? err.message : err) });
  }
}

// =========================
// PIERCING: MARCAR NIVEL (1..4)
// body: { action:"piercing_mark", month, student, level, teacher, email }
// =========================
if (action === "piercing_mark") {
  try {
    const resp = piercing_mark_(body);
    return json_({
      ok: true,
      report: resp.report,
      list: resp.list,
      justReported: resp.justReported,
      wa_msg: resp.wa_msg || "",
      wa_url: resp.wa_url || ""
    });
  } catch (err) {
    return json_({ ok:false, msg: String(err && err.message ? err.message : err) });
  }
}

// =========================
// UNIFORME: LISTAR (MES)
// body: { action:"uniform_list", month:"YYYY-MM" }
// =========================
if (action === "uniform_list") {
  const month = String(body.month || "").trim();
  const items = uniforme_list_(month);
  return json_({ ok:true, items });
}

// =========================
// UNIFORME: ELIMINAR ESTUDIANTE (DEL MES)
// body: { action:"uniform_delete_student", month:"YYYY-MM", student:"Nombre" }
// =========================
if (action === "uniform_delete_student") {
  try {
    const resp = uniforme_delete_student_(body);
    return json_({
      ok: true,
      deleted_rows: resp.deleted_rows,
      msg: resp.msg,
      list: resp.list
    });
  } catch (err) {
    return json_({ ok:false, msg: String(err && err.message ? err.message : err) });
  }
}

// =========================
// UNIFORME: MARCAR NIVEL (1..4)
// body: { action:"uniform_mark", month, student, level, teacher, email }
// =========================
if (action === "uniform_mark") {
  try {
    const resp = uniforme_mark_(body);
    return json_({
      ok: true,
      report: resp.report,
      list: resp.list,
      justReported: resp.justReported,
      wa_msg: resp.wa_msg || "",
      wa_url: resp.wa_url || ""
    });
  } catch (err) {
    return json_({ ok:false, msg: String(err && err.message ? err.message : err) });
  }
}

// =========================
// FALTAS: GUARDAR
// body: { action:"faltas_report", month, tipo, course, student, subject, teacher, email, device_id }
// =========================
if (action === "faltas_report") {
  try {
    const resp = faltas_report_(body);
    return json_(resp);
  } catch (err) {
    return json_({ ok:false, msg: String(err && err.message ? err.message : err) });
  }
}

// =========================
// FALTAS: LISTAR (MES) (opcional tipo)
// body: { action:"faltas_list", month:"YYYY-MM", tipo?:1|2|3 }
// =========================
if (action === "faltas_list") {
  try {
    const items = faltas_list_(body);
    return json_({ ok:true, items });
  } catch (err) {
    return json_({ ok:false, msg: String(err && err.message ? err.message : err) });
  }
}


// ✅ Admin: leer configuración
if (action === "admin_get_config") {
  const cfg = getConfig_();
  return json_({ ok:true, config: cfg });
}

// ✅ Admin: guardar configuración
if (action === "admin_set_config") {
  const onTime = String(body.ON_TIME_HHMM || body.on_time_hhmm || "").trim();
  const exitOk = String(body.EXIT_OK_HHMM || body.exit_ok_hhmm || "").trim();
  const radius = String(body.SCHOOL_RADIUS_METERS || body.school_radius_m || "").trim();
  const acc = String(body.REQUIRED_ACCURACY_METERS || body.required_accuracy_m || "").trim();
  const latStr = String(body.SCHOOL_LAT || "").trim();
  const lngStr = String(body.SCHOOL_LNG || "").trim();

  const hhmm = /^\d{2}:\d{2}$/;
  if (!hhmm.test(onTime) || !hhmm.test(exitOk)) {
    return json_({ ok:false, msg:"Horas inválidas. Usa formato HH:MM (ej: 06:30)." });
  }

  const r = Number(radius);
  const a = Number(acc);
  const lat = Number(latStr);
  const lng = Number(lngStr);

  if (isNaN(r) || r <= 0 || isNaN(a) || a <= 0) {
    return json_({ ok:false, msg:"Radio/precisión inválidos. Deben ser números > 0." });
  }

  if (isNaN(lat) || lat < -90 || lat > 90) {
    return json_({ ok:false, msg:"Latitud inválida." });
  }

  if (isNaN(lng) || lng < -180 || lng > 180) {
    return json_({ ok:false, msg:"Longitud inválida." });
  }

  setConfig_({
    ON_TIME_HHMM: onTime,
    EXIT_OK_HHMM: exitOk,
    SCHOOL_RADIUS_METERS: String(Math.round(r)),
    REQUIRED_ACCURACY_METERS: String(Math.round(a)),
    SCHOOL_LAT: String(lat),
    SCHOOL_LNG: String(lng),
  });

  return json_({ ok:true, msg:"Configuración guardada ✅", config: getConfig_() });
}


    // ✅ Solicitud de autorización de dispositivo (crea fila en DISPOSITIVOS con AUTORIZADO=NO)
// ✅ Solicitud de autorización de dispositivo (crea fila en DISPOSITIVOS con AUTORIZADO=NO)
if (action === "request_device_auth") {
  const email = String(body.email || "").trim().toLowerCase();
  const name = String(body.name || "").trim();
  const course = String(body.course || "").trim();
  const deviceId = String(body.device_id || "").trim();

  if (!email || !name || !course || !deviceId) {
    return json_({ ok:false, error:"missing_fields", msg:"Faltan datos para solicitar autorización." });
  }

  const sh = ensureDevicesSheet_();
  const lastRow = sh.getLastRow();

  // ✅ Aviso si el correo ya tiene OTRO dispositivo autorizado
  const conflictInfo = emailHasAuthorizedOtherDevice_(email, deviceId);
  const warningMsg = conflictInfo.conflict
    ? "⚠️ OJO: este correo ya está autorizado en otro teléfono. Si cambiaste de celular, el admin debe autorizar este nuevo dispositivo."
    : "";

  // Verificar si ya existe (email + deviceId) (evita duplicados)
  if (lastRow >= 2) {
    const values = sh.getRange(2, 1, lastRow - 1, 4).getValues(); // A-D
    for (let i = 0; i < values.length; i++) {
      const e = String(values[i][0] || "").trim().toLowerCase();
      const d = String(values[i][1] || "").trim();
      if (e === email && d === deviceId) {
        // ya hay fila para este device
        return json_({
          ok:true,
          exists:true,
          msg: warningMsg
            ? `Ya existe una solicitud para este dispositivo. ${warningMsg}`
            : "Ya existe una solicitud para este dispositivo."
        });
      }
    }
  }

  const row = lastRow + 1;
  const stamp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");
  const notes = `Solicitud ${stamp} | ${name} | Curso: ${course}`;

  sh.getRange(row, 1, 1, 4).setValues([[email, deviceId, "NO", notes]]);
  sh.autoResizeColumns(1, 4);

  return json_({
    ok:true,
    row,
    msg: warningMsg
      ? `Solicitud registrada. Pendiente aprobación. ${warningMsg}`
      : "Solicitud registrada. Pendiente aprobación."
  });
}



  // ✅ Verificar si el dispositivo ya está autorizado para ese correo
if (action === "check_device_auth") {
  const email = String(body.email || "").trim().toLowerCase();
  const deviceId = String(body.device_id || "").trim();

  if (!email || !deviceId) {
    return json_({ ok:false, error:"missing_fields", msg:"Faltan email o device_id." });
  }

  // 1) si este device está autorizado para este correo, OK normal
  const info = isDeviceAuthorized_(email, deviceId);
  if (info.authorized) {
    return json_({
      ok: true,
      authorized: true,
      allowed_name: info.allowedName || "",
      msg: "Dispositivo autorizado ✅"
    });
  }

  // 2) si el correo YA tiene otro dispositivo autorizado, avisar (conflicto)
  const conflictInfo = emailHasAuthorizedOtherDevice_(email, deviceId);
  if (conflictInfo.conflict) {
    return json_({
      ok: true,
      authorized: false,
      conflict: true,
      devices_count: conflictInfo.devices.length,
      msg: "⚠️ Este correo ya está autorizado en otro teléfono. Si cambiaste de celular, solicita autorización para este nuevo dispositivo."
    });
  }

  // 3) si no hay ningún autorizado todavía (o solo hay solicitudes NO), mensaje normal
  return json_({
    ok: true,
    authorized: false,
    conflict: false,
    allowed_name: info.allowedName || "",
    msg: "Aún no autorizado. Pide al admin que lo apruebe."
  });
}


// ✅ Admin: verificar PIN
if (action === "admin_check_pin") {
  const pin = String(body.pin || "").trim();
  if (!pin) return json_({ ok:false, error:"missing_pin", msg:"Falta el PIN." });

  const ok = verifyAdminPin_(pin);
  const saved = getAdminPin_();

  return json_({
    ok: true,
    authorized: ok,
    must_change: ok && (saved === DEFAULT_ADMIN_PIN),
    msg: ok ? "Admin autorizado ✅" : "PIN incorrecto ❌"
  });
}

// ✅ Admin: cambiar PIN (requiere pin actual correcto)
if (action === "admin_change_pin") {
  const oldPin = String(body.old_pin || "").trim();
  const newPin = String(body.new_pin || "").trim();

  if (!oldPin || !newPin) {
    return json_({ ok:false, error:"missing_fields", msg:"Faltan old_pin o new_pin." });
  }
  if (!/^\d{4,8}$/.test(newPin)) {
    return json_({ ok:false, error:"invalid_pin", msg:"El nuevo PIN debe tener 4 a 8 dígitos." });
  }
  if (!verifyAdminPin_(oldPin)) {
    return json_({ ok:false, error:"wrong_pin", msg:"PIN actual incorrecto." });
  }

  setAdminPin_(newPin, "PIN actualizado desde la app");
  return json_({ ok:true, msg:"PIN actualizado ✅" });
}

// ✅ Admin: listar docentes del mes actual + traer registros de un docente
if (action === "admin_list_teachers") {
  const month = String(body.month || "").trim(); // "YYYY-MM"
  const sh = getMonthSheetByKey_(month);
  if (!sh) return json_({ ok: true, teachers: [] });

  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return json_({ ok: true, teachers: [] });

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  const teachers = [];
  for (let c = 0; c < header.length; c++) {
    const t = String(header[c] || "").trim();
    if (!t || t.toUpperCase() === "DOCENTES") continue;
    teachers.push(t);
    c++; // saltar pareja entrada/salida
  }

  // únicos
  const uniq = [...new Set(teachers)];
  return json_({ ok: true, teachers: uniq });
}


if (action === "admin_get_teacher_records") {
  const teacher = String(body.teacher || "").trim();
  const month = String(body.month || "").trim(); // "YYYY-MM"
  if (!teacher) return json_({ ok:false, error:"missing_teacher", msg:"Falta teacher." });

  const sh = getMonthSheetByKey_(month);
  if (!sh) return json_({ ok:true, records: [] });

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 3 || lastCol < 1) return json_({ ok:true, records: [] });

  // Buscar columnas del docente por header row1
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  let entradaCol = -1;
  let salidaCol = -1;

  for (let c = 1; c <= lastCol; c++) {
    const h = String(header[c - 1] || "").trim();
    if (h && h.toLowerCase() === teacher.toLowerCase()) {
      entradaCol = c;
      salidaCol = c + 1;
      break;
    }
  }

  if (entradaCol === -1) return json_({ ok:true, records: [] });

  const records = [];

  // ===== ENTRADAS (valores + notes)
  const entradaRange = sh.getRange(3, entradaCol, lastRow - 2, 1);
  const entradaVals = entradaRange.getValues();
  const entradaNotes = entradaRange.getNotes();
entradaVals.forEach((v, idx) => {
  const val = v[0];
  if (!val) return;

  const d = safeDate_(val);
  if (!d) return; // ✅ evita crashear si el valor es "NTC" u otra basura

  const note = String((entradaNotes[idx] && entradaNotes[idx][0]) || "").trim().toUpperCase();

  records.push({
    teacher,
    type: "ENTRADA",
    iso: d.toISOString(),
    date: Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd"),
    time: Utilities.formatDate(d, TIMEZONE, "HH:mm:ss"),
    late: isLate_(d),
    ntc: (note === "NTC")
  });
});


  // ===== SALIDAS (valores + notes)
  if (salidaCol <= lastCol) {
    const salidaRange = sh.getRange(3, salidaCol, lastRow - 2, 1);
    const salidaVals = salidaRange.getValues();
    const salidaNotes = salidaRange.getNotes();

   salidaVals.forEach((v, idx) => {
  const val = v[0];
  if (!val) return;

  const d = safeDate_(val);
  if (!d) return; // ✅ evita crashear si el valor es "NTC"

  const note = String((salidaNotes[idx] && salidaNotes[idx][0]) || "").trim().toUpperCase();

  records.push({
    teacher,
    type: "SALIDA",
    iso: d.toISOString(),
    date: Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd"),
    time: Utilities.formatDate(d, TIMEZONE, "HH:mm:ss"),
    late: false,
    ntc: (note === "NTC")
  });
});

  }

  // ordenar desc por fecha/hora
  records.sort((a,b) => new Date(b.iso) - new Date(a.iso));
  return json_({ ok:true, records });
}




// ✅ NUEVO: lista de meses disponibles (para admin)
if (action === "admin_list_months") {
  const months = listAvailableMonths_();
  return json_({ ok: true, months });
}


// ✅ Admin: Registro manual (marca NOTE = "NTC")
if (action === "admin_manual_register") {
  const teacher = String(body.teacher || "").trim();
  const type = String(body.type || "").trim().toUpperCase(); // ENTRADA / SALIDA
  const iso = String(body.iso || "").trim(); // opcional, si no llega usamos "now"

  if (!teacher || (type !== "ENTRADA" && type !== "SALIDA")) {
    return json_({ ok:false, error:"missing_fields", msg:"Faltan datos (teacher/type)." });
  }

  const dateObj = iso ? new Date(iso) : new Date();
  const sh = ensureMonthSheet_(dateObj);
  const cols = ensureTeacherColumns_(sh, teacher);

  const targetCol = (type === "ENTRADA") ? cols.entradaCol : cols.salidaCol;

  // ✅ Regla: SALIDA solo si hay ENTRADA hoy
  if (type === "SALIDA") {
    const hasEntrada = hasEntradaForDay_(sh, cols.entradaCol, dateObj);
    if (!hasEntrada) {
      return json_({ ok:false, error:"missing_entrada", msg:`No puedes registrar SALIDA: no hay ENTRADA hoy para ${teacher}.` });
    }
  }

  // ✅ Regla: solo 1 por día
  if (hasRecordForDay_(sh, targetCol, dateObj)) {
    return json_({ ok:false, error:"already_registered", msg:`Ya existe ${type} registrada hoy para ${teacher}.` });
  }

  const row = nextEmptyRowInColumn_(sh, targetCol);

  // Guardar fecha
  const cell = sh.getRange(row, targetCol);
  cell.setValue(dateObj);
  cell.setNumberFormat("dd/MM/yyyy HH:mm:ss");

  // Colores igual que normal
  if (type === "ENTRADA") {
    const late = isLate_(dateObj);
    cell.setBackground(late ? "#fca5a5" : "#86efac");
  } else {
    const early = isEarlyExit_(dateObj);
    cell.setBackground(early ? "#fca5a5" : "#86efac");
  }

  // ✅ Marcar NTC (sin teléfono)
  cell.setNote("NTC");

  autoFitCell_(sh, row, targetCol);

  return json_({
    ok:true,
    msg:`Registro manual ${type} guardado (NTC) para ${teacher}.`,
    sheet: sh.getName(),
    row,
    col: targetCol
  });
}

if (action === "register") {
  const teacher = String(body.teacher || "").trim();
  const type = String(body.type || "").trim(); // "ENTRADA" / "SALIDA"
  const iso = String(body.iso || "").trim();

  const email = String(body.email || "").trim().toLowerCase();
  const course = String(body.course || "").trim();
  const deviceId = String(body.device_id || "").trim();

  // ✅ NUEVO: QR + GEO
  const qrRaw = body.qr; // viene del frontend
  const distanceM = Number(body.distance_m || 0);
  const accuracyM = Number(body.accuracy_m || 0);
  const lat = Number(body.lat || 0);
  const lng = Number(body.lng || 0);

  if (!email || !course || !deviceId) return json_({ ok:false, error:"missing_profile" });

  // ✅ Bloqueo: solo dispositivos autorizados para ese correo
  const info = isDeviceAuthorized_(email, deviceId);
  if (!info.authorized) {
    return json_({
      ok:false,
      error:"device_not_authorized",
      msg:"Este teléfono no está autorizado para este docente. Pide al admin que lo apruebe."
    });
  }

  if (!teacher || !type || !iso) return json_({ ok:false, error:"missing_fields" });

  // ✅ NUEVO: validar QR en servidor
  let qr;
  try {
    qr = validateQrOrThrow_(qrRaw);
  } catch (err) {
    const code = String(err && err.message ? err.message : err);
    if (code === "missing_qr") {
      return json_({ ok:false, error:"missing_qr", msg:"Debes escanear el QR para registrar." });
    }
    return json_({ ok:false, error:"invalid_qr", msg:"QR inválido. Escanea el QR oficial del colegio." });
  }

  // ✅ Crear fecha desde el ISO del cliente
  const dateObj = new Date(iso);
  if (isNaN(dateObj.getTime())) {
    return json_({ ok:false, error:"invalid_iso", msg:"Fecha/hora inválida (iso)." });
  }

  const sh = ensureMonthSheet_(dateObj);
  const cols = ensureTeacherColumns_(sh, teacher);

  const targetCol = (type === "ENTRADA") ? cols.entradaCol : cols.salidaCol;

  // ✅ Regla: no permitir SALIDA si no hay ENTRADA hoy
  if (type === "SALIDA") {
    const hasEntrada = hasEntradaForDay_(sh, cols.entradaCol, dateObj);
    if (!hasEntrada) {
      return json_({
        ok: false,
        error: "missing_entrada",
        msg: `No puedes registrar SALIDA porque aún no hay ENTRADA registrada hoy para ${teacher}.`
      });
    }
  }

  // ✅ Validación: solo 1 entrada y 1 salida por día
  if (hasRecordForDay_(sh, targetCol, dateObj)) {
    return json_({
      ok: false,
      error: "already_registered",
      msg: `Ya existe ${type} registrada hoy para ${teacher}.`
    });
  }

  const row = nextEmptyRowInColumn_(sh, targetCol);

  // ✅ Guardar como Date pero TRUNCANDO milisegundos (evita que Sheets redondee +1 segundo)
  const ms = dateObj.getTime();
  const exactDate = new Date(Math.floor(ms / 1000) * 1000);

  const cell = sh.getRange(row, targetCol);
  cell.setValue(exactDate);
  cell.setNumberFormat("dd/MM/yyyy HH:mm:ss");

  // Colores: ENTRADA (tarde/tiempo) y SALIDA (antes/después de 2:00 pm)
  if (type === "ENTRADA") {
    const late = isLate_(dateObj);
    cell.setBackground(late ? "#fca5a5" : "#86efac");
  }

  if (type === "SALIDA") {
    const early = isEarlyExit_(dateObj);
    cell.setBackground(early ? "#fca5a5" : "#86efac");
  }

  // ✅ NUEVO: guardar auditoría en NOTA (QR + ubicación + email/device)
  // (Esto NO cambia tu tabla, solo deja evidencia en la celda.)
  const note =
    `QR:${qr}\n` +
    `EMAIL:${email}\n` +
    `DEVICE:${deviceId}\n` +
    `DIST_M:${isFinite(distanceM) ? Math.round(distanceM) : ""}\n` +
    `ACC_M:${isFinite(accuracyM) ? Math.round(accuracyM) : ""}\n` +
    `LAT:${isFinite(lat) ? lat : ""}\n` +
    `LNG:${isFinite(lng) ? lng : ""}`;

  cell.setNote(note);

  // ✅ Auto-ajuste
  autoFitCell_(sh, row, targetCol);

  return json_({ ok:true, sheet: sh.getName(), row, col: targetCol });
}





  // Opcional: limpiar mes actual
  if (action === "clear_month") {
    const month = String(body.month || "").trim(); // "YYYY-MM"
    const d = month ? new Date(`${month}-01T00:00:00`) : new Date();
    const sh = ensureMonthSheet_(d);

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow >= 3 && lastCol >= 1) {
      sh.getRange(3, 1, lastRow - 2, lastCol).clearContent().clearFormat();
    }
    return json_({ ok:true, sheet: sh.getName(), cleared:true });
  }

  // =========================
  // ✅ ADMIN: LISTAR SOLICITUDES DE DISPOSITIVOS (PENDIENTES)
  // =========================
  if (action === "admin_list_device_requests") {
    const sh = ensureDevicesSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return json_({ ok:true, requests: [] });

    const values = sh.getRange(2, 1, lastRow - 1, 4).getValues(); // A-D
    const requests = [];

    for (let i = 0; i < values.length; i++) {
      const email = String(values[i][0] || "").trim().toLowerCase();
      const deviceId = String(values[i][1] || "").trim();
      const auth = String(values[i][2] || "").trim().toLowerCase();
      const notes = String(values[i][3] || "");

      // pendientes = NO (o vacío)
      const isPending = (auth === "no" || auth === "" || auth === "false");
      if (email && deviceId && isPending) {
        requests.push({
          row: i + 2, // porque empezamos en fila 2
          email,
          device_id: deviceId,
          authorized: false,
          notes
        });
      }
    }

    return json_({ ok:true, requests });
  }


  // =========================
  // ✅ ADMIN: APROBAR / RECHAZAR DISPOSITIVO (POR FILA)
  // body: { row, authorized:"SI"/"NO" }
  // =========================
  if (action === "admin_set_device_auth") {
    const sh = ensureDevicesSheet_();
    const row = Number(body.row || 0);
    const authorized = String(body.authorized || "").trim().toUpperCase(); // "SI" o "NO"

    if (!row || row < 2) {
      return json_({ ok:false, error:"invalid_row", msg:"Fila inválida." });
    }
    if (!(authorized === "SI" || authorized === "NO")) {
      return json_({ ok:false, error:"invalid_value", msg:"authorized debe ser SI o NO." });
    }

    // leer actual
    const current = sh.getRange(row, 1, 1, 4).getValues()[0];
    const email = String(current[0] || "").trim();
    const deviceId = String(current[1] || "").trim();

    if (!email || !deviceId) {
      return json_({ ok:false, error:"missing_row_data", msg:"La fila no tiene email o device_id." });
    }

    // actualizar AUTORIZADO
    sh.getRange(row, 3).setValue(authorized);

    // append sello en NOTAS
    const stamp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    const prevNotes = String(current[3] || "");
    const tag = authorized === "SI" ? "APROBADO" : "RECHAZADO";
    const newNotes = prevNotes
      ? `${prevNotes}\n${tag} ${stamp}`
      : `${tag} ${stamp}`;

    sh.getRange(row, 4).setValue(newNotes);

    return json_({ ok:true, msg:`Dispositivo ${tag.toLowerCase()} ✅`, row, authorized });
  }

// =========================
// WHATSAPP: OBTENER DESTINO
// body: { action:"get_whatsapp_to", email, device_id }
// =========================
if (action === "get_whatsapp_to") {
  const email = String(body.email || "").trim().toLowerCase();
  const deviceId = String(body.device_id || "").trim();

  if (!email || !deviceId) {
    return json_({ ok:false, error:"missing_fields", msg:"Faltan email o device_id." });
  }

  // ✅ Solo dispositivos autorizados pueden leer
  const info = isDeviceAuthorized_(email, deviceId);
  if (!info.authorized) {
    return json_({ ok:false, error:"device_not_authorized", msg:"Dispositivo no autorizado." });
  }

  // ✅ Leer por dispositivo (columna E en DISPOSITIVOS)
  const wa_to = getWhatsAppToByDevice_(email, deviceId);
  return json_({ ok:true, wa_to: wa_to || "" });
}



// =========================
// WHATSAPP: GUARDAR DESTINO
// body: { action:"set_whatsapp_to", email, device_id, whatsapp_to }
// =========================
// =========================
// WHATSAPP_TO: Guardar número destino por (email + device_id)
// body: { action:"set_whatsapp_to", email, device_id, whatsapp_to }
// =========================
if (action === "set_whatsapp_to") {
  const email = String(body.email || "").trim().toLowerCase();
  const deviceId = String(body.device_id || "").trim();
  const whatsappTo = String(body.whatsapp_to || "").trim();

  if (!email || !deviceId || !whatsappTo) {
    return json_({ ok:false, msg:"Faltan email/device_id/whatsapp_to" });
  }

  const resp = setWhatsAppToByDevice_(email, deviceId, whatsappTo);
  return json_({ ok:true, ...resp });
}


  return json_({ ok:false, error:"unknown_action" });
}


function listTeachersFromSheet_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return [];
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  const teachers = [];
  for (let c = 1; c <= lastCol; c++) {
    const t = String(header[c - 1] || "").trim();
    if (!t) continue;
    if (t.toUpperCase() === "DOCENTES") continue;
    teachers.push(t);
    c++; // saltar par (entrada/salida)
  }

  // únicos + orden
  const uniq = Array.from(new Set(teachers.map(x => x.trim()))).filter(Boolean);
  uniq.sort((a,b) => a.localeCompare(b, "es"));
  return uniq;
}

function getTeacherRecords_(sh, teacher) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 3 || lastCol < 1) return [];

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const key = String(teacher || "").trim().toLowerCase();

  let entradaCol = null;
  let salidaCol = null;

  for (let c = 1; c <= lastCol; c++) {
    const t = String(header[c - 1] || "").trim().toLowerCase();
    if (t && t === key) {
      entradaCol = c;
      salidaCol = c + 1;
      break;
    }
  }

  if (!entradaCol) return [];

  const out = [];

  // ENTRADAS
  const entradaVals = sh.getRange(3, entradaCol, lastRow - 2, 1).getValues();
  entradaVals.forEach(v => {
    const val = v[0];
    if (!val) return;
    const d = (val instanceof Date) ? val : new Date(val);
    out.push({
      teacher: teacher,
      type: "ENTRADA",
      iso: d.toISOString(),
      date: Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd"),
      time: Utilities.formatDate(d, TIMEZONE, "HH:mm:ss"),
      late: isLate_(d)
    });
  });

  // SALIDAS
  if (salidaCol && salidaCol <= lastCol) {
    const salidaVals = sh.getRange(3, salidaCol, lastRow - 2, 1).getValues();
    salidaVals.forEach(v => {
      const val = v[0];
      if (!val) return;
      const d = (val instanceof Date) ? val : new Date(val);
      out.push({
        teacher: teacher,
        type: "SALIDA",
        iso: d.toISOString(),
        date: Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd"),
        time: Utilities.formatDate(d, TIMEZONE, "HH:mm:ss"),
        late: false
      });
    });
  }

  out.sort((a,b) => new Date(b.iso) - new Date(a.iso));
  return out;
}


/************ LECTURA PARA LA APP ************/
function getLatestRecords_(sh, limit) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 3 || lastCol < 1) return [];

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const records = [];

  for (let c = 1; c <= lastCol; c++) {
    const teacher = String(header[c - 1] || "").trim();
    if (!teacher || teacher === "DOCENTES") continue;

    const entradaCol = c;
    const salidaCol = c + 1;

    // ===== ENTRADAS (valores + notes)
    const entradaRange = sh.getRange(3, entradaCol, lastRow - 2, 1);
    const entradaVals = entradaRange.getValues();
    const entradaNotes = entradaRange.getNotes();

 entradaVals.forEach((v, idx) => {
  const val = v[0];
  if (!val) return;

  const d = safeDate_(val);
  if (!d) return; // ✅ evita crashear si el valor es "NTC"

  const note = String((entradaNotes[idx] && entradaNotes[idx][0]) || "").trim().toUpperCase();

  records.push({
    teacher,
    type: "ENTRADA",
    iso: d.toISOString(),
    date: Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd"),
    time: Utilities.formatDate(d, TIMEZONE, "HH:mm:ss"),
    late: isLate_(d),
    ntc: (note === "NTC")
  });
});


    // ===== SALIDAS
    if (salidaCol <= lastCol) {
      const salidaRange = sh.getRange(3, salidaCol, lastRow - 2, 1);
      const salidaVals = salidaRange.getValues();
      const salidaNotes = salidaRange.getNotes();

    salidaVals.forEach((v, idx) => {
  const val = v[0];
  if (!val) return;

  const d = safeDate_(val);
  if (!d) return; // ✅ evita crashear si el valor es "NTC"

  const note = String((salidaNotes[idx] && salidaNotes[idx][0]) || "").trim().toUpperCase();

  records.push({
    teacher,
    type: "SALIDA",
    iso: d.toISOString(),
    date: Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd"),
    time: Utilities.formatDate(d, TIMEZONE, "HH:mm:ss"),
    late: false,
    ntc: (note === "NTC")
  });
});

    }

    c++; // saltar pareja
  }

  records.sort((a,b) => new Date(b.iso) - new Date(a.iso));
  return records.slice(0, limit);
}


function listTeachersFromSheet_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return [];
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  const teachers = [];
  for (let c = 1; c <= lastCol; c++) {
    const t = String(header[c - 1] || "").trim();
    if (!t || t === "DOCENTES") continue;
    teachers.push(t);
    c++; // saltar pareja
  }
  // únicos + orden
  return [...new Set(teachers)].sort((a,b)=>a.localeCompare(b, "es"));
}


function exportSheetAsXlsxBlob_(sheet, filename) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  const gid = sheet.getSheetId();

  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx&gid=${gid}`;

  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true,
  });

  return resp.getBlob()
    .setName(filename)
    .setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

function forceDownloadXlsx_(blob, filename) {
  const file = DriveApp.createFile(blob.setName(filename));

  // ✅ Permite que cualquier teléfono lo descargue sin iniciar sesión
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const id = file.getId();
  const dl = `https://drive.google.com/uc?export=download&id=${id}`;

  // ✅ En PWA a veces el meta refresh falla, por eso dejo link + JS
  return HtmlService.createHtmlOutput(`
    <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>Descargando…</title>
    </head>
    <body style="font-family:system-ui;padding:16px">
      <h3>Descarga lista ✅</h3>
      <p>Si no inicia automáticamente, toca aquí:</p>
      <p>
        <a href="${dl}"
           style="display:inline-block;padding:12px 14px;background:#0b5fff;color:#fff;border-radius:10px;text-decoration:none">
          Descargar Excel
        </a>
      </p>

      <script>
        // intento automático
        location.href = "${dl}";
      </script>
    </body>
    </html>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}



// ✅ Exporta solo un docente (manteniendo formato) en una copia temporal
function exportTeacherAsXlsxBlob_(sourceSheet, teacher) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const temp = SpreadsheetApp.create(`TEMP_EXPORT_${sourceSheet.getName()}_${teacher}`);
  const tempId = temp.getId();

  // borrar hoja default
  const tempSheets = temp.getSheets();
  if (tempSheets && tempSheets.length) temp.deleteSheet(tempSheets[0]);

  // copiar hoja completa con formato
  const copied = sourceSheet.copyTo(temp).setName(sourceSheet.getName());

  // dejar SOLO columnas del docente + A
  keepOnlyTeacherColumns_(copied, teacher);

  // exportar como xlsx
  const gid = copied.getSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${tempId}/export?format=xlsx&gid=${gid}`;

  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true,
  });

  const blob = resp.getBlob()
    .setName(`asistencia_${sourceSheet.getName()}_${teacher}.xlsx`)
    .setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

  // NO lo borres todavía o se rompe el download en algunos teléfonos
  // (luego lo automatizamos a papelera con trigger)
  // DriveApp.getFileById(tempId).setTrashed(true);

  return blob;
}



// ✅ Conserva SOLO columnas A + (Entrada/Salida del docente) en la hoja copiada
function keepOnlyTeacherColumns_(sh, teacher) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 2) return;

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  let entradaCol = -1;
  let salidaCol = -1;

  for (let c = 1; c <= lastCol; c++) {
    const h = String(header[c - 1] || "").trim().toLowerCase();
    if (h === teacher.trim().toLowerCase()) {
      entradaCol = c;
      salidaCol = c + 1;
      break;
    }
  }

  if (entradaCol === -1) return;

  // borrar de derecha a izquierda para no dañar índices
  for (let c = lastCol; c >= 2; c--) {
    const keep = (c === entradaCol || c === salidaCol);
    if (!keep) sh.deleteColumn(c);
  }

  // ajustar tamaños automáticamente (opcional)
  sh.autoResizeColumns(1, sh.getLastColumn());
}


/************ ASIGNATURAS EN DISPOSITIVOS ************/
function ensureDispositivosHeadersAsignaturas_() {
  const sh = ensureDevicesSheet_(); // 👈 OJO: en tu script se llama así
  const headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];

  const required = [
    "ASIGNATURA_1",
    "ASIGNATURA_2",
    "ASIGNATURA_3",
    "ASIGNATURA_4",
    "ASIGNATURA_5"
  ];

  let changed = false;
  required.forEach(name => {
    if (headers.indexOf(name) === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(name);
      changed = true;
    }
  });

  if (changed) {
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight("bold");
  }

  return sh;
}

function dispositivosHeaderMap_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    map[String(h || "").trim().toUpperCase()] = i + 1;
  });
  return map;
}

function asignaturas_get_(deviceIdRaw) {
  const deviceId = String(deviceIdRaw || "").trim();
  if (!deviceId) return [];

  const sh = ensureDispositivosHeadersAsignaturas_();
  const map = dispositivosHeaderMap_(sh);

  const colDevice = map["DEVICE_ID"];
  if (!colDevice) throw new Error("No encontré DEVICE_ID.");

  const data = sh.getDataRange().getValues();

  const row = data.find((r, i) => i > 0 && String(r[colDevice - 1]).trim() === deviceId);
  if (!row) return [];

  return [
    row[map["ASIGNATURA_1"] - 1] || "",
    row[map["ASIGNATURA_2"] - 1] || "",
    row[map["ASIGNATURA_3"] - 1] || "",
    row[map["ASIGNATURA_4"] - 1] || "",
    row[map["ASIGNATURA_5"] - 1] || ""
  ];
}

function asignaturas_save_(deviceIdRaw, itemsRaw) {
  const deviceId = String(deviceIdRaw || "").trim();
  if (!deviceId) throw new Error("device_id requerido");

  const items = Array.isArray(itemsRaw) ? itemsRaw : [];

  const sh = ensureDispositivosHeadersAsignaturas_();
  const map = dispositivosHeaderMap_(sh);

  const colDevice = map["DEVICE_ID"];
  if (!colDevice) throw new Error("No encontré DEVICE_ID.");

  const data = sh.getDataRange().getValues();

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colDevice - 1]).trim() === deviceId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error("Dispositivo no encontrado.");

  sh.getRange(rowIndex, map["ASIGNATURA_1"]).setValue(items[0] || "");
  sh.getRange(rowIndex, map["ASIGNATURA_2"]).setValue(items[1] || "");
  sh.getRange(rowIndex, map["ASIGNATURA_3"]).setValue(items[2] || "");
  sh.getRange(rowIndex, map["ASIGNATURA_4"]).setValue(items[3] || "");
  sh.getRange(rowIndex, map["ASIGNATURA_5"]).setValue(items[4] || "");

  return true;
}





