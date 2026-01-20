const FORM_ID = "xxxxxxxxxxxxxxx"; //ganti dengan id form
const CALENDAR_ID = "xxxxxxxxxxx@group.calendar.google.com"; //ganti dengan id google calendar group
const ADMIN_EMAIL = "xxxxxxxxx@gmail.com"; //ganti dengan email admin atau email yang digunakan saat ini

function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let data;
  let lastRow;

  // 1. Ambil Data
  if (e && e.values) {
    data = e.values;
    lastRow = e.range.getRow();
  } else {
    let detectedRow = sheet.getLastRow();
    lastRow = detectedRow - 100; 
    if (lastRow < 2) lastRow = 2; 
    data = sheet.getRange(lastRow, 1, 1, 14).getValues()[0];
  }

  Logger.log('--- START PROCESS BARIS: ' + lastRow + ' ---');
  
  const email = data[1];
  const editAcara = data[3];
  const jenisAcara = data[4];
  const namaAcara = data[5];
  const tanggalMulaiRaw = data[6]; 
  const jamMulaiRaw = data[7];
  const deskripsiAcara = data[8];
  const tanggalSelesaiRaw = data[9] || data[6];
  const jamSelesaiRaw = data[10];

  // 2. Parsing Waktu
  const tanggalMulai = parseDate(tanggalMulaiRaw);
  const jamMulai = parseTime(jamMulaiRaw);
  tanggalMulai.setHours(jamMulai.hours, jamMulai.minutes, 0, 0);

  const tanggalSelesai = parseDate(tanggalSelesaiRaw);
  let jamSelesai = parseTime(jamSelesaiRaw || "12:00 PM");
  tanggalSelesai.setHours(jamSelesai.hours, jamSelesai.minutes, 0, 0);

  if (tanggalSelesai <= tanggalMulai) {
    tanggalSelesai.setTime(tanggalMulai.getTime() + (60 * 60 * 1000)); 
  }

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const eventTitle = `${jenisAcara} - ${namaAcara}`;
  const eventId = Utilities.getUuid();

  // 3. LOGIKA EDIT (Cari dan Hapus yang Lama)
  if (editAcara && editAcara !== "") {
    Logger.log('Mencoba mencari acara lama untuk dihapus: ' + editAcara);
    const allData = sheet.getDataRange().getValues();
    
    for (let i = 0; i < allData.length; i++) {
      if (allData[i][12] === editAcara) { 
        let oldICalUID = allData[i][13]; // Ambil iCalUID dari kolom N
        
        Logger.log('Data lama ditemukan di baris ' + (i + 1) + '. iCalUID: ' + oldICalUID);
        
        // Hapus dari Google Calendar
        if (oldICalUID) {
          try {
            const oldEvent = calendar.getEventById(oldICalUID);
            if (oldEvent) {
              oldEvent.deleteEvent();
              Logger.log('Berhasil menghapus acara lama di Calendar.');
            }
          } catch (err) {
            Logger.log("Gagal hapus di Calendar (mungkin sudah terhapus): " + err.toString());
          }
        }
        
        // Kosongkan label lama di Kolom M & N baris tersebut agar tidak muncul lagi di dropdown
        sheet.getRange(i + 1, 13, 1, 2).clearContent(); 
        break;
      }
    }
  }

  // 4. Buat Acara Baru di Baris Baru
  sheet.getRange(lastRow, 12).setValue(eventId); 

  try {
    const guests = email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : ADMIN_EMAIL;
    const event = calendar.createEvent(eventTitle, tanggalMulai, tanggalSelesai, {
      description: deskripsiAcara,
      guests: guests,
      sendInvites: true,
    });

    sheet.getRange(lastRow, 14).setValue(event.getId()); 
    Logger.log('SUCCESS: Acara baru dibuat di baris ' + lastRow);
  } catch (err) {
    Logger.log("ERROR SAAT BUAT ACARA: " + err.toString());
  }

  // 5. Update Dropdown
  updateDropdownLabelColumn();
  updateFormDropdownOptions();
  Logger.log('--- SELESAI ---');
}

// Parsing tanggal format MM/DD/YYYY atau objek Date
function parseDate(dateStr) {
  if (!dateStr) return new Date();
  if (dateStr instanceof Date) return new Date(dateStr);

  const parts = String(dateStr).split('/');
  if (parts.length === 3) {
    return new Date(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1]));
  }
  const attempt = new Date(dateStr);
  return isNaN(attempt.getTime()) ? new Date() : attempt;
}

//Parsing waktu format AM/PM atau 24H
function parseTime(timeInput) {
  if (!timeInput) return { hours: 0, minutes: 0 };
  if (timeInput instanceof Date) return { hours: timeInput.getHours(), minutes: timeInput.getMinutes() };

  const timeString = String(timeInput).trim();
  const parts = timeString.match(/(\d{1,2})(?::(\d{2}))?\s*(AM|PM)?/i);
  if (!parts) return { hours: 0, minutes: 0 };

  let hours = parseInt(parts[1]) || 0;
  const minutes = parseInt(parts[2]) || 0;
  const period = (parts[3] || "").toUpperCase();

  if (period === "PM" && hours !== 12) hours += 12;
  else if (period === "AM" && hours === 12) hours = 0;
  
  return { hours, minutes };
}

function updateDropdownLabelColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const jenis = data[i][4];
    const nama = data[i][5];
    const tanggal = data[i][6];
    const jam = data[i][7];
    const uid = data[i][13];
    if (nama && uid) {
      const dateObj = (tanggal instanceof Date) ? tanggal : parseDate(tanggal);
      const tglFormatted = Utilities.formatDate(dateObj, "GMT+07:00", "dd/MM/yyyy");
      const label = `${jenis} - ${nama} | ${tglFormatted} | ${jam}`;
      sheet.getRange(i + 1, 13).setValue(label);
    }
  }
}

function updateFormDropdownOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const labels = sheet.getRange(2, 13, lastRow - 1).getValues()
    .map(row => row[0])
    .filter(l => l && String(l).trim() !== "");
  
  if (labels.length === 0) return;
  const form = FormApp.openById(FORM_ID);
  const items = form.getItems(FormApp.ItemType.LIST);
  for (const item of items) {
    if (item.getTitle().toLowerCase().includes("edit acara")) {
      item.asListItem().setChoiceValues([...new Set(labels)]);
    }
  }
}

function runOnOpen() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) ScriptApp.deleteTrigger(t);
  
  ScriptApp.newTrigger("onFormSubmit").forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onFormSubmit().create();
  ScriptApp.newTrigger("updateFormDropdownOptions").timeBased().everyHours(1).create();
  
  updateDropdownLabelColumn();
  updateFormDropdownOptions();
}
