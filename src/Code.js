/**
 * Konfigurasi ID Form dan ID Google Calendar
 */
const FORM_ID = "xxxxx"; // Ganti dengan ID Form
const CALENDAR_ID = "xxxxxx@group.calendar.google.com"; // Ganti dengan ID Kalender
const ADMIN_EMAIL = "xxxxxx@gmail.com"; // Ganti dengan email admin

/**
 * Menangani submit dari Google Form
 */
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  Logger.log('Executing onFormSubmit for row: ' + lastRow); // LOG 1

  const data = sheet.getRange(lastRow, 1, 1, 14).getValues()[0];
  Logger.log('Data from sheet for last row (first 14 columns): ' + JSON.stringify(data)); // LOG 2

  const email = data[1];
  const editAcara = data[3];
  const jenisAcara = data[4];
  const namaAcara = data[5];
  const tanggalMulai = new Date(data[6]);
  const jamMulaiRaw = data[7];
  const deskripsiAcara = data[8];
  const tanggalSelesai = data[9] ? new Date(data[9]) : new Date(tanggalMulai);
  const jamSelesaiRaw = data[10];

  Logger.log(`Email: ${email}, Edit Acara: ${editAcara}, Jenis Acara: ${jenisAcara}, Nama Acara: ${namaAcara}`); // LOG 3
  Logger.log(`Tanggal Mulai: ${tanggalMulai}, Jam Mulai Raw: ${jamMulaiRaw}`); // LOG 4

  const jamMulai = parseTime(jamMulaiRaw);
  tanggalMulai.setHours(jamMulai.hours, jamMulai.minutes, 0);
  Logger.log('Parsed Jam Mulai: ' + JSON.stringify(jamMulai)); // LOG 5

  let jamSelesai;
  if (!jamSelesaiRaw && ["Evaluasi", "Rapat"].includes(jenisAcara)) {
    jamSelesai = { hours: jamMulai.hours + 2, minutes: jamMulai.minutes };
    Logger.log('Jam Selesai (default +2h): ' + JSON.stringify(jamSelesai)); // LOG 6a
  } else {
    jamSelesai = parseTime(jamSelesaiRaw || "12:00 PM");
    Logger.log('Jam Selesai (parsed): ' + JSON.stringify(jamSelesai)); // LOG 6b
  }
  tanggalSelesai.setHours(jamSelesai.hours, jamSelesai.minutes, 0);
  Logger.log(`Tanggal Selesai Final: ${tanggalSelesai}`); // LOG 7


  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const eventTitle = `${jenisAcara} - ${namaAcara}`;
  const eventId = Utilities.getUuid();
  let iCalUID;
  Logger.log(`Generated Event ID: ${eventId}`); // LOG 8

  // Jika mengedit acara lama
  if (editAcara) {
    Logger.log('Edit Acara mode active. Looking for old event: ' + editAcara); // LOG 9
    const allData = sheet.getDataRange().getValues();
    let foundOldEvent = false;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][11] === editAcara) { // Kolom L menyimpan Event ID
        iCalUID = allData[i][13]; // Kolom N
        Logger.log(`Old event found at row ${i + 1}. Old iCalUID: ${iCalUID}`); // LOG 10

        sheet.getRange(i + 1, 13).clearContent();
        sheet.getRange(i + 1, 14).clearContent();
        Logger.log(`Cleared old Dropdown Label (M) and iCalUID (N) at row ${i + 1}`); // LOG 11

        try {
          const oldEvent = calendar.getEventById(iCalUID);
          if (oldEvent) {
            oldEvent.deleteEvent();
            Logger.log(`Successfully deleted old event from Google Calendar: ${iCalUID}`); // LOG 12
          } else {
            Logger.log(`Old event with iCalUID ${iCalUID} not found in calendar.`); // LOG 13
          }
        } catch (e) {
          Logger.log("Failed to delete old event from Google Calendar: " + e.toString()); // LOG 14
        }
        foundOldEvent = true;
        break;
      }
    }
    if (!foundOldEvent) {
      Logger.log("No old event found matching editAcara value: " + editAcara); // LOG 15
    }
  }

  // Simpan ID Acara Baru
  sheet.getRange(lastRow, 12).setValue(eventId); // Kolom L (ini untuk baris baru)
  Logger.log(`Event ID ${eventId} set in sheet at L${lastRow}`); // LOG 16

  try {
    const guests = email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : ADMIN_EMAIL;
    Logger.log('Guests for event: ' + guests); // LOG 17

    const event = calendar.createEvent(eventTitle, tanggalMulai, tanggalSelesai, {
      description: deskripsiAcara,
      guests: guests,
      sendInvites: true,
    });
    Logger.log('Event created in Calendar. Event object: ' + event.getId()); // LOG 18 (Cek apakah ada ID)

    event.removeAllReminders();
    event.addPopupReminder(1440);
    event.addPopupReminder(300);
    event.addPopupReminder(180);

    sheet.getRange(lastRow, 14).setValue(event.getId()); // Kolom N
    Logger.log(`iCalUID ${event.getId()} set in sheet at N${lastRow}`); // LOG 19

  } catch (e) {
    Logger.log("Caught exception during event creation: " + e.toString()); // LOG 20
  }

  updateDropdownLabelColumn();
  updateFormDropdownOptions();
  Logger.log('updateDropdownLabelColumn and updateFormDropdownOptions called.'); // LOG 21
  Logger.log('onFormSubmit function completed.'); // LOG 22
}

/**
 * Parsing waktu seperti "07:00 AM" atau "13:30"
 */
function parseTime(timeInput) {
  if (!timeInput) {
    return { hours: 0, minutes: 0 }; // Default jika tidak ada input
  }

  // Coba parse sebagai objek Date jika sudah dalam format Date/Time
  if (timeInput instanceof Date) {
    return { hours: timeInput.getHours(), minutes: timeInput.getMinutes() };
  }

  // Konversi ke string jika bukan Date object
  const timeString = String(timeInput).trim();

  // Regex untuk menangani berbagai format waktu (HH:MM, HH:MM AM/PM, dll.)
  const parts = timeString.match(/(\d{1,2})(?::(\d{2}))?\s*(AM|PM)?/i);
  if (!parts) {
    // Coba parse dengan Date object jika regex gagal (misal: "07:00:00 GMT+0700 (Western Indonesia Time)")
    try {
      const dateFromTime = new Date(`2000-01-01T${timeString}`); // Menggunakan tanggal dummy
      if (!isNaN(dateFromTime)) {
        return { hours: dateFromTime.getHours(), minutes: dateFromTime.getMinutes() };
      }
    } catch (error) {
      Logger.log("Failed to parse time with Date object: " + error);
    }
    return { hours: 0, minutes: 0 }; // Default jika tidak ada input yang valid
  }

  let hours = parseInt(parts[1]) || 0;
  const minutes = parseInt(parts[2]) || 0;
  const period = (parts[3] || "").toUpperCase(); // Bisa kosong jika format 24 jam

  if (period === "PM" && hours !== 12) {
    hours += 12;
  } else if (period === "AM" && hours === 12) {
    hours = 0; // 12 AM (tengah malam) adalah 00:00
  }
  return { hours, minutes };
}


/**
 * Update kolom Dropdown Label berdasarkan data valid
 */
function updateDropdownLabelColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues(); // Ambil semua data termasuk header

  // Mulai dari baris kedua (indeks 1) karena baris pertama adalah header
  for (let i = 1; i < data.length; i++) {
    const jenis = data[i][4]; // Kolom E
    const nama = data[i][5]; // Kolom F
    const tanggal = data[i][6]; // Kolom G
    const jam = data[i][7]; // Kolom H
    const uid = data[i][13]; // Kolom N (iCalUID)

    // Buat label hanya jika Nama Acara dan iCalUID ada
    if (nama && uid) {
      // Pastikan tanggal adalah objek Date yang valid
      const dateObject = new Date(tanggal);
      const tanggalFormatted = Utilities.formatDate(dateObject, "GMT+07:00", "dd/MM/yyyy");
      const label = `${jenis} - ${nama} | ${tanggalFormatted} | ${jam}`;
      sheet.getRange(i + 1, 13).setValue(label); // Kolom M
    } else {
      // Kosongkan kolom M jika data tidak valid (misal, acara lama yang sudah dihapus)
      sheet.getRange(i + 1, 13).clearContent(); // Kolom M
    }
  }
}

/**
 * Update pilihan dropdown di Google Form (untuk Edit Acara)
 */
function updateFormDropdownOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Ambil semua label dari kolom M (Dropdown Label), mulai dari baris ke-2
  const labels = sheet.getRange(2, 13, sheet.getLastRow() - 1).getValues()
    .map(row => row[0]) // Ambil nilai pertama dari setiap baris (yaitu kolom M)
    .filter(label => label && String(label).trim() !== ""); // Filter label yang kosong atau hanya spasi

  const form = FormApp.openById(FORM_ID);
  const items = form.getItems(FormApp.ItemType.LIST); // Cari item bertipe LIST (Dropdown)

  for (const item of items) {
    // Periksa apakah judul item dropdown mengandung "edit acara" (case-insensitive)
    if (item.getTitle().toLowerCase().includes("edit acara")) {
      item.asListItem().setChoiceValues(labels); // Set pilihan dropdown dengan label yang valid
      Logger.log("Opsi dropdown 'Edit Acara' telah diperbarui.");
    }
  }
}

/**
 * Cek acara yang akan dimulai & kirim pengingat
 * Catatan: Bagian ini akan diubah untuk TIDAK mengirim email manual,
 * karena Google Calendar sudah menangani pengingat via email/popup.
 */
function checkAndSendReminders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  for (let i = 1; i < data.length; i++) {
    const title = `${data[i][4]} - ${data[i][5]}`; // Jenis Acara - Nama Acara
    const tanggal = new Date(data[i][6]); // Tanggal Mulai
    const jam = data[i][7]; // Jam Mulai
    const id = data[i][13]; // iCalUID
    if (!id || !jam || !tanggal) continue; // Lewati jika data kunci tidak lengkap

    const parsed = parseTime(jam);
    tanggal.setHours(parsed.hours, parsed.minutes, 0); // Gabungkan tanggal dan jam mulai

    const diffMins = Math.floor((tanggal.getTime() - now.getTime()) / 60000); // Selisih dalam menit

    // Jika Anda ingin Google Apps Script mengirim pengingat email manual (berlawanan dengan ketentuan),
    // aktifkan kembali blok if ini. Namun, disarankan untuk mengandalkan pengingat Google Calendar.
    /*
    if ([1440, 300, 180].includes(diffMins)) { // 1 hari, 5 jam, 3 jam
      try {
        const event = calendar.getEventById(id);
        if (event) {
          const reminderText = `Pengingat: '${title}' akan dimulai pada:\n` +
            Utilities.formatDate(tanggal, "GMT+07:00", "EEEE, dd MMMM, hh:mm a");
          // Kirim email ke semua tamu acara
          event.getGuestList().forEach(g => {
            const guestEmail = g.getEmail();
            if (guestEmail && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(guestEmail)) {
              MailApp.sendEmail(guestEmail, `Reminder: ${title}`, reminderText);
              Logger.log(`Pengingat email terkirim ke ${guestEmail} untuk acara: ${title}`);
            }
          });
        }
      } catch (e) {
        Logger.log("Reminder error (checkAndSendReminders): " + e.toString());
      }
    }
    */
  }

  // Tetap perbarui opsi dropdown setelah memeriksa pengingat, untuk menjaga konsistensi
  updateFormDropdownOptions();
}

/**
 * Jalankan otomatis saat Spreadsheet dibuka
 * Mengatur triggers dan memperbarui dropdown options
 */
function runOnOpen() {
  const triggers = ScriptApp.getProjectTriggers();
  // Hapus trigger lama untuk menghindari duplikasi
  for (const t of triggers) {
    if (t.getHandlerFunction() === "onFormSubmit" || t.getHandlerFunction() === "checkAndSendReminders") {
      ScriptApp.deleteTrigger(t);
    }
  }

  // Buat trigger untuk onFormSubmit
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  Logger.log("Trigger 'onFormSubmit' telah dibuat ulang.");


  // Buat trigger berbasis waktu untuk checkAndSendReminders (setiap 1 menit)
  ScriptApp.newTrigger("checkAndSendReminders")
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log("Trigger 'checkAndSendReminders' (setiap 1 menit) telah dibuat ulang.");

  // Perbarui kolom Dropdown Label dan opsi dropdown di Form saat spreadsheet dibuka
  updateDropdownLabelColumn();
  updateFormDropdownOptions();
  Logger.log("Triggers dan dropdown diperbarui saat onOpen.");
}
