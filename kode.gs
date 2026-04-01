// ====== KONFIG ======
const SPREADSHEET_ID = '17JWUSVWAUrFeL8_ANqb8kJXLAi0CUvPujmh6-jdV9Jg'; 
const SHEET_USERS = 'Users';
const SHEET_LOG   = 'PresensiLog';
const SHEET_SETTINGS = 'Settings'; 
const SHEET_CALENDAR = 'CalendarEvents'; 
const CENTER_LAT = -7.760861250435269;
const CENTER_LNG = 110.41217797657056;
const RADIUS_M = 500;
const MAX_ACCURACY_M = 80;
const SELFIE_FOLDER_ID = '1nnTRsJa4Rg5QZ8lRzg0ZEu1ckOwq-wyU'; 
const TZ = 'Asia/Jakarta';

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

function doGet(e) {
 const email = Session.getActiveUser().getEmail().toLowerCase().trim();
 const isUiiEmail = email.endsWith('@uii.ac.id') || email.endsWith('@students.uii.ac.id');
  if (!email || !isUiiEmail) {
   const template = HtmlService.createTemplateFromFile('error_auth');
   template.scriptUrl = ScriptApp.getService().getUrl(); template.userEmail = email;
   return template.evaluate().setTitle('Akses Terkunci - Metive 4.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
 }
 if (e.parameter.page == 'dashboard') { return HtmlService.createTemplateFromFile('dashboard').evaluate().setTitle('Admin Dashboard').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1'); }
 return HtmlService.createTemplateFromFile('index').evaluate().setTitle('Metive 4.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=no, viewport-fit=cover');
}

function getCurrentPiketWeek(ss) {
   let shiftWeek = 0; 
   try {
       const sh = ss.getSheets().find(s => s.getName().toLowerCase().trim() === 'settings');
       if (sh) {
           const data = sh.getDataRange().getValues();
           const manualRow = data.find(r => String(r[0]).toUpperCase().trim() === 'MINGGU_AKTIF');
           if (manualRow && manualRow[1] !== '') { const val = parseInt(manualRow[1]); if (!isNaN(val) && val > 0) return val; }
           const shiftRow = data.find(r => String(r[0]).toUpperCase().trim() === 'GESER_JADWAL');
           if (shiftRow && shiftRow[1] !== '') { const sVal = parseInt(shiftRow[1]); if (!isNaN(sVal)) shiftWeek = sVal; }
       }
   } catch(e) {}
   const d = new Date(); const dUTC = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())); const dayNum = dUTC.getUTCDay() || 7; dUTC.setUTCDate(dUTC.getUTCDate() + 4 - dayNum); const yearStart = new Date(Date.UTC(dUTC.getUTCFullYear(), 0, 1)); const isoWeek = Math.ceil((((dUTC - yearStart) / 86400000) + 1) / 7); const effectiveWeek = isoWeek - shiftWeek + 1000; return (effectiveWeek % 2 === 0) ? 2 : 1;
}

function getInitialData() {
 try {
   const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
   const sheetUsers = ss.getSheetByName(SHEET_USERS); const sheetLog = ss.getSheetByName(SHEET_LOG);
   if (!sheetUsers || !sheetLog) throw new Error("Sheet tidak ditemukan.");

   const lastRow = sheetLog.getLastRow(); const startRow = Math.max(2, lastRow - 499); const numRows = (lastRow >= 2) ? (lastRow - startRow + 1) : 0;
   const maxCol = sheetLog.getLastColumn(); let logValues = []; if (numRows > 0 && maxCol > 0) logValues = sheetLog.getRange(startRow, 1, numRows, maxCol).getValues();

   const profile = getProfile(sheetUsers);
   const today = getTodayStatusFromArray(logValues, profile.email);
   const history = getMonthlyHistory(sheetLog, profile.email);
   const todayKey = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
   const stories = getTodayStoriesFromData(logValues, todayKey);
  
   const currentWeek = getCurrentPiketWeek(ss);
   const usersData = sheetUsers.getDataRange().getValues(); 
   const leaderboard = getWeeklyLeaderboardFromData(logValues, usersData, currentWeek);

   // AMBIL DATA USER AKTIF (Week 1 & 2) UNTUK LIST INVITATION
   const activeUsers = [];
   if(usersData.length > 1) {
       const header = usersData[0].map(h => String(h).toLowerCase().trim());
       const idxEmail = header.indexOf('email'); const idxNama = header.indexOf('nama'); const idxWeek = header.findIndex(h => h === 'week' || h === 'kelompok' || h === 'jadwal');
       if(idxEmail !== -1 && idxNama !== -1 && idxWeek !== -1) {
           for(let i=1; i<usersData.length; i++) {
               let w = parseInt(usersData[i][idxWeek]);
               if(w === 1 || w === 2) {
                   activeUsers.push({ email: String(usersData[i][idxEmail]).toLowerCase().trim(), nama: String(usersData[i][idxNama]).split(' ')[0] });
               }
           }
       }
   }

   const calendarEvents = getCalendarEventsData(ss);

   return { ok: true, profile: profile, todayStatus: today, history: history, stories: stories, currentWeek: currentWeek, leaderboard: leaderboard, calendarEvents: calendarEvents, activeUsers: activeUsers };
 } catch (e) { return { ok: false, message: e.toString() }; }
}

// ================== CRUD KALENDER ==================
function getCalendarEventsData(ss) {
    try {
        const sheetCal = ss.getSheetByName(SHEET_CALENDAR); if (!sheetCal) return [];
        const lastRow = sheetCal.getLastRow(); if (lastRow < 2) return [];
        const data = sheetCal.getRange(2, 1, lastRow - 1, 7).getValues();
        const events = [];
        data.forEach(row => {
            if (row[0] && row[1]) {
                let tglStr = String(row[1]);
                if (row[1] instanceof Date) tglStr = Utilities.formatDate(row[1], TZ, 'yyyy-MM-dd');
                else if (tglStr.length > 10) { try { tglStr = Utilities.formatDate(new Date(row[1]), TZ, 'yyyy-MM-dd'); } catch(e) { tglStr = tglStr.substring(0, 10); } }
                
                let pesertaArr = [];
                try { pesertaArr = JSON.parse(String(row[6] || '[]')); } catch(e) { pesertaArr = []; }

                events.push({
                    id: String(row[0]), tanggal: tglStr, waktu: String(row[2] || '-'), judul: String(row[3] || '-'),
                    email: String(row[4] || '').toLowerCase(), nama: String(row[5] || 'Anonim'), peserta: pesertaArr
                });
            }
        });
        return events;
    } catch(e) { return []; }
}

function addCalendarEvent(payload) {
    const lock = LockService.getUserLock(); if (!lock.tryLock(5000)) throw new Error('Sistem sibuk.');
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID); const sheetUsers = ss.getSheetByName(SHEET_USERS); const sheetCal = ss.getSheetByName(SHEET_CALENDAR);
        if (!sheetCal) throw new Error('Sheet CalendarEvents belum dibuat.');
        const profile = getProfile(sheetUsers); if (profile.nama === 'Guest') throw new Error('Akses ditolak.');
        const eventId = 'EVT-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
        const pesertaStr = JSON.stringify(payload.peserta || []);
        sheetCal.appendRow([eventId, payload.tanggal, payload.waktu, payload.judul, profile.email, profile.nama.split(' ')[0], pesertaStr]);
        return { ok: true, message: 'Event ditambahkan!', events: getCalendarEventsData(ss) };
    } catch(e) { return { ok: false, message: e.message }; } finally { lock.releaseLock(); }
}

// FUNGSI BARU: MENGEDIT EVENT YANG SUDAH ADA
function editCalendarEvent(payload) {
    const lock = LockService.getUserLock(); if (!lock.tryLock(5000)) throw new Error('Sistem sibuk.');
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID); const sheetUsers = ss.getSheetByName(SHEET_USERS); const sheetCal = ss.getSheetByName(SHEET_CALENDAR);
        if (!sheetCal) throw new Error('Sheet CalendarEvents belum dibuat.');
        const profile = getProfile(sheetUsers); const data = sheetCal.getDataRange().getValues(); let updated = false;
        
        for (let i = data.length - 1; i >= 1; i--) {
            if (String(data[i][0]) === String(payload.id)) {
                if (String(data[i][4]).toLowerCase() !== profile.email) throw new Error('Anda tidak memiliki izin mengedit event orang lain.');
                const pesertaStr = JSON.stringify(payload.peserta || []);
                const rowNum = i + 1;
                // Update ke sheet (Kolom B, C, D, G)
                sheetCal.getRange(rowNum, 2).setValue(payload.tanggal);
                sheetCal.getRange(rowNum, 3).setValue(payload.waktu);
                sheetCal.getRange(rowNum, 4).setValue(payload.judul);
                sheetCal.getRange(rowNum, 7).setValue(pesertaStr);
                updated = true; break;
            }
        }
        if (!updated) throw new Error('Event tidak ditemukan.');
        return { ok: true, message: 'Event diperbarui!', events: getCalendarEventsData(ss) };
    } catch(e) { return { ok: false, message: e.message }; } finally { lock.releaseLock(); }
}

function deleteCalendarEvent(eventId) {
    const lock = LockService.getUserLock(); if (!lock.tryLock(5000)) throw new Error('Sistem sibuk.');
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID); const sheetUsers = ss.getSheetByName(SHEET_USERS); const sheetCal = ss.getSheetByName(SHEET_CALENDAR);
        if (!sheetCal) throw new Error('Sheet CalendarEvents belum dibuat.');
        const profile = getProfile(sheetUsers); const data = sheetCal.getDataRange().getValues(); let deleted = false;
        for (let i = data.length - 1; i >= 1; i--) {
            if (String(data[i][0]) === String(eventId)) {
                if (String(data[i][4]).toLowerCase() !== profile.email) throw new Error('Anda tidak memiliki izin menghapus event orang lain.');
                sheetCal.deleteRow(i + 1); deleted = true; break;
            }
        }
        if (!deleted) throw new Error('Event tidak ditemukan.');
        return { ok: true, message: 'Event dihapus.', events: getCalendarEventsData(ss) };
    } catch(e) { return { ok: false, message: e.message }; } finally { lock.releaseLock(); }
}

// ================== PRESENSI LAMA ==================
function getProfile(sheetObj) {
 try {
   const email = requireUiiEmail_().toLowerCase().trim(); let sh = sheetObj; if (!sh) { const ss = SpreadsheetApp.openById(SPREADSHEET_ID); sh = ss.getSheetByName(SHEET_USERS); }
   const def = { email: email, nama:'Guest', nim:'-', divisi:'-', foto_fileid:'', week: '' }; if (!sh) return def;
   const values = sh.getDataRange().getValues(); if (values.length < 2) return def;
   const header = values[0].map(h => String(h).toLowerCase().trim()); const idxEmail = header.indexOf('email'); const idxNama = header.indexOf('nama'); const idxNim = header.indexOf('nim'); const idxDiv = header.indexOf('divisi'); const idxFoto = header.indexOf('foto_fileid'); const idxWeek = header.findIndex(h => h === 'week' || h === 'kelompok' || h === 'jadwal');
   if (idxEmail === -1) return def;
   const userRow = values.slice(1).find(row => String(row[idxEmail] || '').toLowerCase().trim() === email);
   if (userRow) { return { email: email, nama: (idxNama !== -1 ? String(userRow[idxNama] || '-') : '-'), nim: (idxNim !== -1 ? String(userRow[idxNim] || '-') : '-'), divisi:(idxDiv !== -1 ? String(userRow[idxDiv] || '-') : '-'), foto_fileid: (idxFoto !== -1 ? String(userRow[idxFoto] || '').trim() : ''), week: (idxWeek !== -1 ? String(userRow[idxWeek] || '').trim() : '') }; } return def;
 } catch (e) { return { email: 'Error', nama: e.toString(), nim:'-', divisi:'-', foto_fileid:'', week:'' }; }
}

function saveSelfieToDrive(base64Data, profile) { try { if (!base64Data || base64Data.indexOf('base64,') === -1) return 'Error: Data Gambar Rusak'; const folder = DriveApp.getFolderById(SELFIE_FOLDER_ID); const parts = base64Data.split(','); const mime = parts[0].match(/data:(.*);base64/)[1]; const data = Utilities.base64Decode(parts[1]); const blob = Utilities.newBlob(data, mime, 'selfie.jpg'); const timestamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss'); const fileName = `${timestamp}_${profile.nim || 'NONIM'}.jpg`; const file = folder.createFile(blob).setName(fileName); return file.getUrl(); } catch (e) { return 'Gagal Upload: ' + e.toString(); } }

function submitAttendance(payload) {
 const lock = LockService.getUserLock(); if (!lock.tryLock(5000)) { throw new Error('Sistem sedang memproses presensi Anda. Silakan tunggu beberapa saat.'); }
 try {
   const ss = SpreadsheetApp.openById(SPREADSHEET_ID); const sheetUsers = ss.getSheetByName(SHEET_USERS); const sheetLog = ss.getSheetByName(SHEET_LOG); const profile = getProfile(sheetUsers);
   if (profile.nama === 'Guest' || profile.nim === '-') { throw new Error('Akses ditolak: Email Anda belum terdaftar di Sheet Users.'); }
   const jenis = (payload.jenis || '').toString().toUpperCase(); const aksi = (payload.aksi || '').toString().toUpperCase();
   if (jenis === 'PIKET' && aksi === 'CHECKIN') { const currentWeek = getCurrentPiketWeek(ss); const userWeek = parseInt(profile.week); if (isNaN(userWeek) || userWeek <= 0) { throw new Error('Akses Ditolak: Anda tidak memiliki jadwal piket (Anggota Pensiun/Non-Aktif).'); } if (userWeek !== currentWeek) { throw new Error(`Akses Ditolak: Minggu ini adalah jadwal Week ${currentWeek}, sedangkan Anda Week ${userWeek}.`); } }
   const lastRow = sheetLog.getLastRow(); const startRow = Math.max(2, lastRow - 499); const numRows = (lastRow >= 2) ? (lastRow - startRow + 1) : 0; const maxCol = sheetLog.getLastColumn(); let logValues = []; if (numRows > 0 && maxCol > 0) { logValues = sheetLog.getRange(startRow, 1, numRows, maxCol).getValues(); }
   const lat = Number(payload.lat); const lng = Number(payload.lng); const akurasi = Number(payload.akurasi || 0); const poseReq = (payload.pose || '-').toString().trim(); const moodReq = (payload.mood || '-').toString().trim();
   if (!Number.isFinite(lat)) throw new Error('Lokasi kosong.'); const jarak = haversineMeters(lat, lng, CENTER_LAT, CENTER_LNG); let statusJarak = ''; if (jarak > RADIUS_M) statusJarak = `LUAR RADIUS (${jarak}m)`;
   const catatan = (payload.catatan_jobdesk || '').toString().trim(); const selfieRaw = (payload.selfie_url || '').toString().trim(); const alasanTelat = (payload.alasan_telat || '').toString().trim();
   if (!selfieRaw) throw new Error('Selfie wajib.'); if (aksi === 'CHECKOUT' && !catatan) throw new Error('Isi catatan jobdesk.');
   let selfieLink = ''; if (selfieRaw.length > 500) { selfieLink = saveSelfieToDrive(selfieRaw, profile); } else { selfieLink = selfieRaw; }
   const todayKey = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
   if (aksi === 'CHECKOUT') { const open = findOpenSessionFromData(logValues, profile.email, jenis); if (!open) throw new Error(`Belum ada CHECKIN ${jenis}.`); const openDate = Utilities.formatDate(open[0], TZ, 'yyyy-MM-dd'); if (openDate !== todayKey) throw new Error(`Sesi beda hari. Checkin ulang.`); }
   if (aksi === 'CHECKIN') { if (hasCheckinOnDateFromData(logValues, profile.email, jenis, todayKey)) throw new Error(`Sudah Checkin hari ini.`); const open = findOpenSessionFromData(logValues, profile.email, jenis); if (open) { const openDate = Utilities.formatDate(open[0], TZ, 'yyyy-MM-dd'); if (openDate === todayKey) throw new Error(`Sesi ${jenis} belum Checkout.`); } }
   const now = new Date(); let skorPoin = 0; const jam = now.getHours(); const menit = now.getMinutes(); const totalMenit = (jam * 60) + menit;
   if (jenis === 'PIKET') { if (aksi === 'CHECKIN') { if (totalMenit <= 545) skorPoin = 6; else if (totalMenit <= 555) skorPoin = 4; else skorPoin = 2; } else if (aksi === 'CHECKOUT') { if (totalMenit < 1020) skorPoin = -5; else skorPoin = 4; } } else if (jenis === 'KEGIATAN') { if (aksi === 'CHECKIN') skorPoin = 5; if (aksi === 'CHECKOUT') skorPoin = 5; }
   sheetLog.appendRow([now, profile.email, profile.nama, profile.nim, profile.divisi, jenis, aksi, lat, lng, akurasi, catatan, selfieLink, 'VALID', alasanTelat, statusJarak, skorPoin, poseReq, moodReq]);
   return { ok: true, message: `${aksi} ${jenis} Sukses jam ${Utilities.formatDate(now, TZ, 'HH:mm:ss')}.` };
 } finally { lock.releaseLock(); }
}
function getUserEmail_() { const a = (Session.getActiveUser().getEmail() || '').toLowerCase().trim(); const e = (Session.getEffectiveUser().getEmail() || '').toLowerCase().trim(); return a || e || ''; }
function requireUiiEmail_() { const email = getUserEmail_(); if (!email) throw new Error('Email tidak terbaca. Pastikan login akun Google.'); return email; }
function findOpenSessionFromData(rows, email, jenis) { if (!rows || rows.length === 0) return null; for (let i = rows.length - 1; i >= 0; i--) { const r = rows[i]; if (String(r[1]).toLowerCase().trim() === email.toLowerCase().trim() && String(r[5]).toUpperCase().trim() === jenis && String(r[12]).toUpperCase().trim() === 'VALID') { return String(r[6]).toUpperCase().trim() === 'CHECKIN' ? r : null; } } return null; }
function hasCheckinOnDateFromData(rows, email, jenis, dateKey) { if (!rows || rows.length === 0) return false; for (const r of rows) { const ts = r[0]; if (!(ts instanceof Date)) continue; if (Utilities.formatDate(ts, TZ, 'yyyy-MM-dd') !== dateKey) continue; if (String(r[1]).toLowerCase().trim() === email.toLowerCase().trim() && String(r[5]).toUpperCase().trim() === jenis && String(r[6]).toUpperCase().trim() === 'CHECKIN' && String(r[12]).toUpperCase().trim() === 'VALID') { return true; } } return false; }
function haversineMeters(lat1, lng1, lat2, lng2) { const R = 6371000; const toRad = d => d * Math.PI / 180; const dLat = toRad(lat2 - lat1); const dLng = toRad(lng2 - lng1); const a = Math.sin(dLat/2)**2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng/2)**2; return Math.round(2 * R * Math.asin(Math.sqrt(a))); }
function getTodayStatusFromArray(values, email) { if (!values || values.length === 0) return { mode: 'PRE_CHECKIN' }; const todayKey = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); let lastRow = null; for (let i = values.length - 1; i >= 0; i--) { const row = values[i]; const rowEmail = String(row[1] || '').toLowerCase().trim(); const status = String(row[12] || '').toUpperCase().trim(); if (status === 'VALID' && rowEmail === email) { lastRow = row; break; } } if (!lastRow) return { mode: 'PRE_CHECKIN' }; const lastJenis = String(lastRow[5]).toUpperCase().trim(); const lastAksi = String(lastRow[6]).toUpperCase().trim(); const ts = lastRow[0]; if (!(ts instanceof Date)) return { mode: 'PRE_CHECKIN' }; if (lastAksi === 'CHECKIN' && Utilities.formatDate(ts, TZ, 'yyyy-MM-dd') === todayKey) { return { mode: 'POST_CHECKIN', jenis: lastJenis, timestamp: Utilities.formatDate(ts, TZ, 'yyyy-MM-dd HH:mm:ss') }; } return { mode: 'PRE_CHECKIN' }; }

// ====== MODIFIKASI FOTO STORY (AMBIL FOTO TERBARU SAAT CHECKOUT) ======
function getTodayStoriesFromData(logValues, todayKey) { 
    const stories = {}; 
    for (let i = 0; i < logValues.length; i++) { 
        const row = logValues[i]; const ts = row[0]; 
        if (!(ts instanceof Date)) continue; 
        if (Utilities.formatDate(ts, TZ, 'yyyy-MM-dd') !== todayKey) continue; 
        if (String(row[12]).toUpperCase().trim() !== 'VALID') continue; 
        const email = String(row[1]).toLowerCase().trim(); 
        const aksi = String(row[6]).toUpperCase().trim(); 
        
        if (!stories[email]) { 
            stories[email] = { nama: String(row[2] || '').split(' ')[0], foto: String(row[11] || ''), mood: row.length > 17 ? String(row[17] || '😊') : '😊', waktu: Utilities.formatDate(ts, TZ, 'HH:mm'), isCheckedOut: false, lastAksiTime: ts.getTime() }; 
        } 
        
        if (row.length > 17 && row[17]) stories[email].mood = String(row[17]); 
        
        if (aksi === 'CHECKOUT') { 
            stories[email].isCheckedOut = true; 
            stories[email].lastAksiTime = ts.getTime(); 
            const newFoto = String(row[11] || '').trim();
            if (newFoto && newFoto !== '-') {
                stories[email].foto = newFoto; // Update foto dengan foto Checkout terbaru
            }
        } else if (aksi === 'AUTO_CHECKOUT') { 
            stories[email].isCheckedOut = true; 
            stories[email].lastAksiTime = ts.getTime(); 
            // Kalau auto_checkout gak ada foto baru, jadi tetep pake foto lama
        } else if (aksi === 'CHECKIN') { 
            stories[email].isCheckedOut = false; 
            stories[email].foto = String(row[11] || ''); 
            stories[email].waktu = Utilities.formatDate(ts, TZ, 'HH:mm'); 
            stories[email].lastAksiTime = ts.getTime(); 
        } 
    } 
    
    const result = Object.values(stories).map(s => { 
        let finalFoto = s.foto; 
        if (finalFoto) { 
            const matchD = finalFoto.match(/\/d\/([a-zA-Z0-9_-]+)/); 
            const matchId = finalFoto.match(/id=([a-zA-Z0-9_-]+)/); 
            if (matchD && matchD[1]) finalFoto = "https://lh3.googleusercontent.com/d/" + matchD[1] + "=s200"; 
            else if (matchId && matchId[1]) finalFoto = "https://lh3.googleusercontent.com/d/" + matchId[1] + "=s200"; 
        } 
        s.foto = finalFoto; 
        let emoji = '😊'; 
        const m = String(s.mood).toLowerCase(); 
        if (m.includes('sangat senang') || m.includes('great')) emoji = '🤩'; 
        else if (m.includes('senang') || m.includes('good')) emoji = '😊'; 
        else if (m.includes('biasa') || m.includes('okay')) emoji = '😐'; 
        else if (m.includes('sedih') || m.includes('bad')) emoji = '😔'; 
        else if (m.includes('stress')) emoji = '🤯'; 
        s.moodEmoji = emoji; return s; 
    }); 
    
    result.sort((a, b) => { 
        if (a.isCheckedOut === b.isCheckedOut) return b.lastAksiTime - a.lastAksiTime; 
        return a.isCheckedOut ? 1 : -1; 
    }); 
    return result; 
}

function getWeeklyLeaderboardFromData(logValues, usersData, currentWeek) { const now = new Date(); const day = now.getDay() || 7; const monday = new Date(now); monday.setDate(now.getDate() - day + 1); monday.setHours(0,0,0,0); const scores = {}; if (usersData && usersData.length > 1) { const header = usersData[0].map(h => String(h).toLowerCase().trim()); const idxEmail = header.indexOf('email'); const idxNama = header.indexOf('nama'); const idxFoto = header.indexOf('foto_fileid'); const idxWeek = header.findIndex(h => h === 'week' || h === 'kelompok' || h === 'jadwal'); if (idxEmail !== -1 && idxWeek !== -1) { for (let i = 1; i < usersData.length; i++) { const row = usersData[i]; const uEmail = String(row[idxEmail] || '').toLowerCase().trim(); const uWeek = parseInt(row[idxWeek]); if (uEmail && uWeek === currentWeek) { scores[uEmail] = { nama: String(row[idxNama] || '').split(' ')[0], skor: 0, foto: String(row[idxFoto] || '').trim() }; } } } } for (let i = 0; i < logValues.length; i++) { const row = logValues[i]; const ts = row[0]; if (!(ts instanceof Date) || ts < monday) continue; if (String(row[12]).toUpperCase().trim() !== 'VALID' || String(row[5]).toUpperCase().trim() !== 'PIKET') continue; const email = String(row[1]).toLowerCase().trim(); const skor = Number(row[15] || 0); const selfieUrl = String(row[11] || '').trim(); if (scores[email]) { scores[email].skor += skor; if (selfieUrl) { scores[email].foto = selfieUrl; } } } const result = Object.values(scores).sort((a, b) => b.skor - a.skor); result.forEach(s => { let finalFoto = s.foto; if (finalFoto) { const matchD = finalFoto.match(/\/d\/([a-zA-Z0-9_-]+)/); const matchId = finalFoto.match(/id=([a-zA-Z0-9_-]+)/); if (matchD && matchD[1]) finalFoto = "https://lh3.googleusercontent.com/d/" + matchD[1] + "=s200"; else if (matchId && matchId[1]) finalFoto = "https://lh3.googleusercontent.com/d/" + matchId[1] + "=s200"; } s.foto = finalFoto; }); return result; }
function getMonthlyHistory(sheetLog, email) { const result = { ok: true, rows: [], summary: {piket:0, kegiatan:0} }; try { if (!sheetLog) { const ss = SpreadsheetApp.openById(SPREADSHEET_ID); sheetLog = ss.getSheetByName(SHEET_LOG); } const lastRow = sheetLog.getLastRow(); if (lastRow < 2) return result; const now = new Date(); const thisMonthKey = Utilities.formatDate(now, TZ, 'yyyy-MM'); const limit = 500; const startRow = Math.max(2, lastRow - limit + 1); const numRows = (lastRow >= 2) ? (lastRow - startRow + 1) : 0; if (numRows === 0) return result; const maxCol = sheetLog.getLastColumn(); let data = []; if (maxCol > 0) { data = sheetLog.getRange(startRow, 1, numRows, maxCol).getValues(); } const grouped = {}; for (let i = 0; i < data.length; i++) { const r = data[i]; const rowEmail = String(r[1]).toLowerCase().trim(); const status = String(r[12]).toUpperCase().trim(); if (rowEmail !== email || status !== 'VALID') continue; const time = r[0]; if (!(time instanceof Date)) continue; const rowMonth = Utilities.formatDate(time, TZ, 'yyyy-MM'); if (rowMonth !== thisMonthKey) continue; const dateKey = Utilities.formatDate(time, TZ, 'yyyy-MM-dd'); const jenis = String(r[5]).toUpperCase().trim(); const aksi = String(r[6]).toUpperCase().trim(); const skor = Number(r[15] || 0); const uniqueKey = dateKey + '_' + jenis; if (!grouped[uniqueKey]) { grouped[uniqueKey] = { rawDate: dateKey, jenis: jenis, checkinTime: null, checkoutTime: null, skorTotal: 0, isAuto: false }; } grouped[uniqueKey].skorTotal += skor; if (aksi === 'CHECKIN') { if (!grouped[uniqueKey].checkinTime) grouped[uniqueKey].checkinTime = time; } else if (aksi === 'CHECKOUT' || aksi === 'AUTO_CHECKOUT') { grouped[uniqueKey].checkoutTime = time; if (aksi === 'AUTO_CHECKOUT') { grouped[uniqueKey].isAuto = true; } } } const sortedKeys = Object.keys(grouped).sort().reverse(); sortedKeys.forEach(k => { const item = grouped[k]; let durasi = 0; if (item.checkinTime && item.checkoutTime) { durasi = Math.round((item.checkoutTime - item.checkinTime) / (1000 * 60)); } const jamMasuk = item.checkinTime ? Utilities.formatDate(item.checkinTime, TZ, 'HH:mm') : '-'; const jamPulang = item.checkoutTime ? Utilities.formatDate(item.checkoutTime, TZ, 'HH:mm') : '-'; result.rows.push({ date: item.rawDate, jenis: item.jenis, checkin: jamMasuk, checkout: jamPulang, isAutoCheckout: item.isAuto, durasiMenit: durasi, skor: item.skorTotal }); if (item.jenis === 'PIKET') result.summary.piket += item.skorTotal; else result.summary.kegiatan += item.skorTotal; }); return result; } catch (e) { return { ok: false, message: e.toString() }; } }
function prosesLupaCheckout() { const lock = LockService.getScriptLock(); if (!lock.tryLock(10000)) return; try { const ss = SpreadsheetApp.openById(SPREADSHEET_ID); const sheetLog = ss.getSheetByName(SHEET_LOG); const lastRow = sheetLog.getLastRow(); if (lastRow < 2) return; const todayKey = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); const limit = 1500; const startRow = Math.max(2, lastRow - limit + 1); const numRows = lastRow >= 2 ? (lastRow - startRow + 1) : 0; if (numRows === 0) return; const data = sheetLog.getRange(startRow, 1, numRows, 18).getValues(); const tracking = {}; for (let i = 0; i < data.length; i++) { const row = data[i]; const time = row[0]; if (!(time instanceof Date)) continue; const rowDate = Utilities.formatDate(time, TZ, 'yyyy-MM-dd'); if (rowDate !== todayKey) continue; const status = String(row[12]).toUpperCase().trim(); if (status !== 'VALID') continue; const email = String(row[1]).toLowerCase().trim(); const jenis = String(row[5]).toUpperCase().trim(); const aksi = String(row[6]).toUpperCase().trim(); const uniqueKey = email + '_' + jenis; if (!tracking[uniqueKey]) tracking[uniqueKey] = { email: email, nama: String(row[2]).trim(), nim: String(row[3]).trim(), divisi: String(row[4]).trim(), jenis: jenis, hasCheckin: false, hasCheckout: false }; if (aksi === 'CHECKIN') tracking[uniqueKey].hasCheckin = true; if (aksi === 'CHECKOUT' || aksi === 'AUTO_CHECKOUT') tracking[uniqueKey].hasCheckout = true; } const rowsToAppend = []; const now = new Date(); for (const key in tracking) { const t = tracking[key]; if (t.hasCheckin && !t.hasCheckout) rowsToAppend.push([now, t.email, t.nama, t.nim, t.divisi, t.jenis, 'AUTO_CHECKOUT', 0, 0, 0, 'SISTEM: Lupa Checkout', '-', 'VALID', '-', '-', -5, '-', '-']); } if (rowsToAppend.length > 0) sheetLog.getRange(sheetLog.getLastRow() + 1, 1, rowsToAppend.length, 18).setValues(rowsToAppend); } catch(e) {} finally { lock.releaseLock(); } }
