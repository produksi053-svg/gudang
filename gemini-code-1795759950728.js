/**
 * SISTEM MANAJEMEN GUDANG & PRODUKSI (WMS) - BACKEND (Code.gs)
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('WMS Pro — Sistem Manajemen Gudang & Produksi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function formatDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val).trim();
}

function nowTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function generateId(prefix, sheet) {
  const lastRow = sheet.getLastRow();
  const num = lastRow < 2 ? 1 : lastRow;
  return prefix + '-' + String(num).padStart(3, '0') + '-' + Math.floor(100 + Math.random() * 900);
}

/**
 * 1. AUTENTIKASI LOGIN (EDITOR / VIEWER)
 */
function loginUser(username, password) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const uInput = String(username || '').trim().toLowerCase();
  const pInput = String(password || '').trim();

  for (let i = 1; i < data.length; i++) {
    const u = String(data[i][0] || '').trim().toLowerCase();
    const p = String(data[i][1] || '').trim();
    const nama = String(data[i][2] || u);
    const role = String(data[i][3] || 'Viewer').trim();
    const status = String(data[i][4] || 'Aktif').trim();

    if (u === uInput && p === pInput) {
      if (status.toLowerCase() !== 'aktif') {
        return { success: false, message: 'Akun dinonaktifkan. Hubungi Administrator.' };
      }
      return {
        success: true,
        user: {
          username: data[i][0],
          nama: nama,
          role: role
        }
      };
    }
  }
  return { success: false, message: 'Username atau Password salah.' };
}

/**
 * 2. AMBIL SELURUH DATA REALTIME & URL SPREADSHEET
 */
function getAllWmsData() {
  const ss = getSS();
  const spreadsheetUrl = ss.getUrl();

  // Gudang (Rows 4 - 20 di Dashboard_Gudang: dinamis sesuai nama gudang yang diisi)
  const dashSheet = ss.getSheetByName('Dashboard_Gudang');
  const gudangRaw = dashSheet.getRange('A4:K20').getValues();
  const gudangList = [];
  for (let i = 0; i < gudangRaw.length; i++) {
    const nama = String(gudangRaw[i][0] || '').trim();
    if (!nama) continue;
    gudangList.push({
      row: 4 + i,
      nama: nama,
      kapasitas: Number(gudangRaw[i][1]) || 0,
      stokAwal: Number(gudangRaw[i][2]) || 0,
      stokMasuk: Number(gudangRaw[i][3]) || 0,
      stokTransisi: Number(gudangRaw[i][4]) || 0,
      stokRebagging: Number(gudangRaw[i][5]) || 0,
      stokFG: Number(gudangRaw[i][6]) || 0,
      terpakai: Number(gudangRaw[i][7]) || 0,
      sisa: Number(gudangRaw[i][8]) || 0,
      utilisasi: Number(gudangRaw[i][9]) || 0,
      status: gudangRaw[i][10] || 'AMAN'
    });
  }

  // Master Barang
  const masterSheet = ss.getSheetByName('Master_Barang');
  const masterData = masterSheet.getDataRange().getValues();
  const masterList = [];
  for (let i = 1; i < masterData.length; i++) {
    if (!masterData[i][1]) continue;
    masterList.push({
      row: i + 1,
      kode: String(masterData[i][0] || ''),
      nama: String(masterData[i][1] || ''),
      kategori: String(masterData[i][2] || 'Komoditas Utama'),
      satuan: String(masterData[i][3] || 'Kg'),
      gudangTujuan: String(masterData[i][4] || ''),
      stokAwal: Number(masterData[i][5]) || 0,
      totalMasukKg: Number(masterData[i][6]) || 0,
      totalPassKg: Number(masterData[i][7]) || 0,
      totalFailKg: Number(masterData[i][8]) || 0,
      stokFGKg: Number(masterData[i][9]) || 0,
      totalKeluarKg: Number(masterData[i][10]) || 0,
      stokAkhirKg: Number(masterData[i][11]) || 0,
      stokAkhirTon: Number(masterData[i][12]) || 0,
      stokAwalKg: Number(masterData[i][13]) || 0,
      oleh: String(masterData[i][14] || '')
    });
  }

  // Masuk
  const masukData = ss.getSheetByName('Masuk').getDataRange().getValues();
  const masukList = [];
  for (let i = 1; i < masukData.length; i++) {
    if (!masukData[i][0]) continue;
    masukList.push({
      row: i + 1,
      id: String(masukData[i][0]),
      tanggal: formatDateStr(masukData[i][1]),
      shiftLine: String(masukData[i][2]),
      namaBarang: String(masukData[i][3]),
      qty: Number(masukData[i][4]) || 0,
      satuan: String(masukData[i][5] || 'Kg'),
      keterangan: String(masukData[i][6]),
      gudang: String(masukData[i][7]),
      qtyKg: Number(masukData[i][8]) || 0,
      inputOleh: String(masukData[i][9] || ''),
      dieditOleh: String(masukData[i][10] || ''),
      waktuUpdate: String(masukData[i][11] || '')
    });
  }

  // Status
  const statusData = ss.getSheetByName('Status').getDataRange().getValues();
  const statusList = [];
  for (let i = 1; i < statusData.length; i++) {
    if (!statusData[i][0]) continue;
    statusList.push({
      row: i + 1,
      id: String(statusData[i][0]),
      tanggal: formatDateStr(statusData[i][1]),
      shiftLine: String(statusData[i][2]),
      namaBarang: String(statusData[i][3]),
      qty: Number(statusData[i][4]) || 0,
      statusQC: String(statusData[i][5]),
      tujuanOtomatis: String(statusData[i][6]),
      gudang: String(statusData[i][7]),
      qtyKg: Number(statusData[i][8]) || 0,
      satuan: String(statusData[i][9] || 'Kg'),
      catatan: String(statusData[i][10] || ''),
      inputOleh: String(statusData[i][11] || ''),
      dieditOleh: String(statusData[i][12] || '')
    });
  }

  // Transisi
  const transisiData = ss.getSheetByName('Transisi').getDataRange().getValues();
  const transisiList = [];
  for (let i = 1; i < transisiData.length; i++) {
    if (!transisiData[i][0]) continue;
    transisiList.push({
      row: i + 1,
      id: String(transisiData[i][0]),
      refStatus: String(transisiData[i][1]),
      tanggal: formatDateStr(transisiData[i][2]),
      shiftLine: String(transisiData[i][3]),
      namaBarang: String(transisiData[i][4]),
      qty: Number(transisiData[i][5]) || 0,
      asalGudang: String(transisiData[i][6]),
      tindakan: String(transisiData[i][7]),
      gudangTujuan: String(transisiData[i][8]),
      statusProses: String(transisiData[i][9]),
      qtyKg: Number(transisiData[i][10]) || 0,
      satuan: String(transisiData[i][11] || 'Kg'),
      inputOleh: String(transisiData[i][12] || ''),
      dieditOleh: String(transisiData[i][13] || '')
    });
  }

  // Rebagging
  const rebagData = ss.getSheetByName('Rebagging').getDataRange().getValues();
  const rebagList = [];
  for (let i = 1; i < rebagData.length; i++) {
    if (!rebagData[i][0]) continue;
    rebagList.push({
      row: i + 1,
      id: String(rebagData[i][0]),
      sumber: String(rebagData[i][1]),
      refId: String(rebagData[i][2]),
      tanggal: formatDateStr(rebagData[i][3]),
      shiftLine: String(rebagData[i][4]),
      namaBarang: String(rebagData[i][5]),
      qtyMasuk: Number(rebagData[i][6]) || 0,
      qtyFG: Number(rebagData[i][7]) || 0,
      gudangFG: String(rebagData[i][8]),
      status: String(rebagData[i][9]),
      satuan: String(rebagData[i][10] || 'Kg'),
      qtyMasukKg: Number(rebagData[i][11]) || 0,
      qtyFGKg: Number(rebagData[i][12]) || 0,
      inputOleh: String(rebagData[i][13] || ''),
      dieditOleh: String(rebagData[i][14] || '')
    });
  }

  // FG (Finished Goods per barang dan per gudang dinamis)
  const fgData = ss.getSheetByName('FG').getDataRange().getValues();
  const fgHeaders = fgData[0] || [];
  const fgList = [];
  for (let i = 1; i < fgData.length; i++) {
    if (!fgData[i][0]) continue;
    const perGudang = {};
    for (let c = 5; c < fgHeaders.length; c++) {
      const hName = String(fgHeaders[c] || '').trim();
      if (hName && hName !== '-') {
        perGudang[hName] = Number(fgData[i][c]) || 0;
      }
    }
    fgList.push({
      namaBarang: String(fgData[i][0]),
      totalMasukKg: Number(fgData[i][1]) || 0,
      totalKeluarKg: Number(fgData[i][2]) || 0,
      stokSiapKirimKg: Number(fgData[i][3]) || 0,
      stokSiapKirimTon: Number(fgData[i][4]) || 0,
      perGudang: perGudang
    });
  }

  // Keluar
  const keluarData = ss.getSheetByName('Keluar').getDataRange().getValues();
  const keluarList = [];
  for (let i = 1; i < keluarData.length; i++) {
    if (!keluarData[i][0]) continue;
    keluarList.push({
      row: i + 1,
      id: String(keluarData[i][0]),
      tanggal: formatDateStr(keluarData[i][1]),
      namaBarang: String(keluarData[i][2]),
      nopol: String(keluarData[i][3]),
      qty: Number(keluarData[i][4]) || 0,
      asalGudang: String(keluarData[i][5]),
      tujuan: String(keluarData[i][6]),
      satuan: String(keluarData[i][7] || 'Kg'),
      qtyKg: Number(keluarData[i][8]) || 0,
      inputOleh: String(keluarData[i][9] || ''),
      dieditOleh: String(keluarData[i][10] || ''),
      waktuUpdate: String(keluarData[i][11] || '')
    });
  }

  return {
    spreadsheetUrl,
    gudangList,
    masterList,
    masukList: masukList.reverse(),
    statusList: statusList.reverse(),
    transisiList: transisiList.reverse(),
    rebagList: rebagList.reverse(),
    fgList,
    keluarList: keluarList.reverse()
  };
}

/**
 * 3. UPDATE NAMA & KAPASITAS GUDANG + TAMBAH GUDANG BARU (EDITOR ONLY)
 */
function saveGudangSettings(items, newGudang) {
  const ss = getSS();
  const dashSheet = ss.getSheetByName('Dashboard_Gudang');

  // Update nama & kapasitas gudang yang sudah ada
  items.forEach(item => {
    const r = Number(item.row);
    if (r >= 4 && r <= 20) {
      dashSheet.getRange(r, 1).setValue(String(item.nama || '').trim());
      dashSheet.getRange(r, 2).setValue(Number(item.kapasitas) || 0);
    }
  });

  // Tambah gudang baru pada slot kosong pertama (baris 4 - 20)
  if (newGudang && String(newGudang.nama || '').trim()) {
    const vals = dashSheet.getRange('A4:A20').getValues();
    let placed = false;
    for (let i = 0; i < vals.length; i++) {
      if (!String(vals[i][0] || '').trim()) {
        const targetRow = 4 + i;
        dashSheet.getRange(targetRow, 1).setValue(String(newGudang.nama).trim());
        dashSheet.getRange(targetRow, 2).setValue(Number(newGudang.kapasitas) || 0);
        placed = true;
        break;
      }
    }
    if (!placed) {
      return { success: false, message: 'Slot gudang penuh (maksimal 17 gudang pada baris 4–20).' };
    }
  }

  SpreadsheetApp.flush();
  return { success: true, message: 'Pengaturan nama & kapasitas gudang berhasil disimpan!' };
}

/**
 * 4. TAMBAH / EDIT MASTER BARANG DENGAN KODE OTOMATIS & GUDANG TUJUAN
 */
function addMasterBarang(payload, username) {
  const ss = getSS();
  const masterSheet = ss.getSheetByName('Master_Barang');
  const fgSheet = ss.getSheetByName('FG');

  // Cari baris kosong pertama di Master_Barang
  const existing = masterSheet.getRange('B2:B200').getValues();
  let nextRow = 2;
  for (let i = 0; i < existing.length; i++) {
    if (!String(existing[i][0] || '').trim()) {
      nextRow = 2 + i;
      break;
    }
  }

  const autoKode = payload.kode || ('BRG-' + String(nextRow - 1).padStart(3, '0'));
  const satuan = payload.satuan || 'Kg';

  masterSheet.getRange(nextRow, 1, 1, 15).setValues([[
    autoKode,
    payload.nama,
    payload.kategori || 'Komoditas Utama',
    satuan,
    payload.gudangTujuan,
    Number(payload.stokAwal) || 0,
    `=SUMIFS(Masuk!$I$2:$I$500, Masuk!$D$2:$D$500, B${nextRow})`,
    `=SUMIFS(Status!$I$2:$I$500, Status!$D$2:$D$500, B${nextRow}, Status!$F$2:$F$500, "Pass")`,
    `=SUMIFS(Status!$I$2:$I$500, Status!$D$2:$D$500, B${nextRow}, Status!$F$2:$F$500, "Fail")`,
    `=SUMIFS(FG!$D$2:$D$500, FG!$A$2:$A$500, B${nextRow})`,
    `=SUMIFS(Keluar!$I$2:$I$500, Keluar!$C$2:$C$500, B${nextRow})`,
    `=N${nextRow}+G${nextRow}-H${nextRow}-I${nextRow}`,
    `=L${nextRow}/1000`,
    `=IF(D${nextRow}="Ton", F${nextRow}*1000, F${nextRow})`,
    `${username} (${nowTimestamp()})`
  ]]);

  // Pastikan baris FG terhubung ke Master_Barang
  const fgRow = nextRow;
  fgSheet.getRange(fgRow, 1, 1, 5).setValues([[
    `=IF(Master_Barang!B${nextRow}="","",Master_Barang!B${nextRow})`,
    `=IF(A${fgRow}="","",SUMIFS(Rebagging!$M$2:$M$500,Rebagging!$F$2:$F$500,$A${fgRow},Rebagging!$J$2:$J$500,"Selesai - Masuk FG"))`,
    `=IF(A${fgRow}="","",SUMIFS(Keluar!$I$2:$I$500,Keluar!$C$2:$C$500,$A${fgRow}))`,
    `=IF(A${fgRow}="","",MAX(0,B${fgRow}-C${fgRow}))`,
    `=IF(A${fgRow}="","",D${fgRow}/1000)`
  ]]);

  for (let col = 6; col <= 22; col++) {
    const colLetter = fgSheet.getRange(1, col).getA1Notation().replace('1', '');
    fgSheet.getRange(fgRow, col).setFormula(
      `=IF(OR($A${fgRow}="",${colLetter}$1="",${colLetter}$1="-"),"",MAX(0,SUMIFS(Rebagging!$M$2:$M$500,Rebagging!$F$2:$F$500,$A${fgRow},Rebagging!$I$2:$I$500,${colLetter}$1,Rebagging!$J$2:$J$500,"Selesai - Masuk FG")-SUMIFS(Keluar!$I$2:$I$500,Keluar!$C$2:$C$500,$A${fgRow},Keluar!$F$2:$F$500,${colLetter}$1)))`
    );
  }

  SpreadsheetApp.flush();
  return { success: true, message: `Master barang ${autoKode} - "${payload.nama}" berhasil ditambahkan!` };
}

function updateMasterBarang(payload, username) {
  const ss = getSS();
  const masterSheet = ss.getSheetByName('Master_Barang');
  const r = Number(payload.row);
  masterSheet.getRange(r, 2, 1, 5).setValues([[
    payload.nama,
    payload.kategori || 'Komoditas Utama',
    payload.satuan || 'Kg',
    payload.gudangTujuan,
    Number(payload.stokAwal) || 0
  ]]);
  masterSheet.getRange(r, 15).setValue(`Edit: ${username} (${nowTimestamp()})`);
  SpreadsheetApp.flush();
  return { success: true, message: `Master barang ${payload.kode} berhasil diperbarui!` };
}

/**
 * 5. INPUT BARANG MASUK MULTI-BARIS (SATUAN KG / TON)
 */
function saveBatchMasuk(rows, username) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Masuk');
  const validRows = rows.filter(r => r.namaBarang && Number(r.qty) > 0);
  if (validRows.length === 0) {
    return { success: false, message: 'Isi minimal 1 baris data barang masuk dengan Qty > 0.' };
  }

  const existing = sheet.getRange('A2:A500').getValues();
  let startRow = 2;
  for (let i = 0; i < existing.length; i++) {
    if (!String(existing[i][0] || '').trim()) {
      startRow = 2 + i;
      break;
    }
  }

  const ts = nowTimestamp();
  validRows.forEach((r, idx) => {
    const rowNum = startRow + idx;
    const satuan = r.satuan || 'Kg';
    sheet.getRange(rowNum, 1, 1, 12).setValues([[
      generateId('MSK', sheet) + '-' + (idx + 1),
      r.tanggal,
      r.shiftLine,
      r.namaBarang,
      Number(r.qty),
      satuan,
      r.keterangan,
      r.gudang,
      `=IF(E${rowNum}="","",IF(F${rowNum}="Ton",E${rowNum}*1000,E${rowNum}))`,
      username,
      '-',
      ts
    ]]);
  });

  SpreadsheetApp.flush();
  return { success: true, message: `${validRows.length} baris barang masuk berhasil disimpan!` };
}

/**
 * 6. INPUT STATUS QC (PASS -> REBAGGING OTOMATIS | FAIL -> TRANSISI OTOMATIS)
 */
function saveStatusQC(payload, username) {
  const ss = getSS();
  const statusSheet = ss.getSheetByName('Status');
  const rebagSheet = ss.getSheetByName('Rebagging');
  const transisiSheet = ss.getSheetByName('Transisi');

  const existing = statusSheet.getRange('A2:A500').getValues();
  let nextRow = 2;
  for (let i = 0; i < existing.length; i++) {
    if (!String(existing[i][0] || '').trim()) {
      nextRow = 2 + i;
      break;
    }
  }

  const statusId = generateId('QC', statusSheet);
  const qty = Number(payload.qty) || 0;
  const satuan = payload.satuan || 'Kg';
  const ts = nowTimestamp();

  statusSheet.getRange(nextRow, 1, 1, 13).setValues([[
    statusId,
    payload.tanggal,
    payload.shiftLine,
    payload.namaBarang,
    qty,
    payload.statusQC,
    `=IF(F${nextRow}="Pass","Rebagging",IF(F${nextRow}="Fail","Transisi",""))`,
    payload.gudang,
    `=IF(E${nextRow}="","",IF(J${nextRow}="Ton",E${nextRow}*1000,E${nextRow}))`,
    satuan,
    payload.catatan || '',
    username,
    '-'
  ]]);

  if (payload.statusQC === 'Pass') {
    const rbgExisting = rebagSheet.getRange('A2:A500').getValues();
    let rbgRow = 2;
    for (let i = 0; i < rbgExisting.length; i++) {
      if (!String(rbgExisting[i][0] || '').trim()) { rbgRow = 2 + i; break; }
    }
    rebagSheet.getRange(rbgRow, 1, 1, 15).setValues([[
      generateId('RBG', rebagSheet),
      'QC Pass',
      statusId,
      payload.tanggal,
      payload.shiftLine,
      payload.namaBarang,
      qty,
      qty,
      payload.gudang,
      'Proses Rebagging',
      satuan,
      `=IF(G${rbgRow}="","",IF(K${rbgRow}="Ton",G${rbgRow}*1000,G${rbgRow}))`,
      `=IF(H${rbgRow}="","",IF(K${rbgRow}="Ton",H${rbgRow}*1000,H${rbgRow}))`,
      username,
      '-'
    ]]);
  } else if (payload.statusQC === 'Fail') {
    const trnExisting = transisiSheet.getRange('A2:A500').getValues();
    let trnRow = 2;
    for (let i = 0; i < trnExisting.length; i++) {
      if (!String(trnExisting[i][0] || '').trim()) { trnRow = 2 + i; break; }
    }
    transisiSheet.getRange(trnRow, 1, 1, 14).setValues([[
      generateId('TRN', transisiSheet),
      statusId,
      payload.tanggal,
      payload.shiftLine,
      payload.namaBarang,
      qty,
      payload.gudang,
      'Redryer',
      payload.gudang,
      'Menunggu Proses',
      `=IF(F${trnRow}="","",IF(L${trnRow}="Ton",F${trnRow}*1000,F${trnRow}))`,
      satuan,
      username,
      '-'
    ]]);
  }

  SpreadsheetApp.flush();
  const tujuan = payload.statusQC === 'Pass' ? 'Rebagging' : 'Transisi';
  return { success: true, message: `Status ${payload.statusQC} (${qty} ${satuan}) disimpan! Otomatis masuk ke ${tujuan}.` };
}

/**
 * 7. PROSES LANJUTAN DI TRANSISI
 */
function processTransisi(payload, username) {
  const ss = getSS();
  const transisiSheet = ss.getSheetByName('Transisi');
  const rebagSheet = ss.getSheetByName('Rebagging');
  const row = Number(payload.row);

  const rowData = transisiSheet.getRange(row, 1, 1, 14).getValues()[0];
  const trnId = rowData[0];
  const shiftLine = rowData[3];
  const namaBarang = rowData[4];
  const qty = Number(rowData[5]) || 0;
  const satuan = rowData[11] || 'Kg';

  transisiSheet.getRange(row, 8).setValue(payload.tindakan);
  transisiSheet.getRange(row, 9).setValue(payload.gudangTujuan);
  transisiSheet.getRange(row, 14).setValue(`${username} (${nowTimestamp()})`);

  if (payload.tindakan === 'Rebagging' && payload.selesaikan) {
    transisiSheet.getRange(row, 10).setValue('Selesai ke Rebagging');
    const rbgExisting = rebagSheet.getRange('A2:A500').getValues();
    let rbgRow = 2;
    for (let i = 0; i < rbgExisting.length; i++) {
      if (!String(rbgExisting[i][0] || '').trim()) { rbgRow = 2 + i; break; }
    }
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    rebagSheet.getRange(rbgRow, 1, 1, 15).setValues([[
      generateId('RBG', rebagSheet),
      'Transisi - Rework',
      trnId,
      today,
      shiftLine,
      namaBarang,
      qty,
      qty,
      payload.gudangTujuan,
      'Proses Rebagging',
      satuan,
      `=IF(G${rbgRow}="","",IF(K${rbgRow}="Ton",G${rbgRow}*1000,G${rbgRow}))`,
      `=IF(H${rbgRow}="","",IF(K${rbgRow}="Ton",H${rbgRow}*1000,H${rbgRow}))`,
      username,
      '-'
    ]]);
    SpreadsheetApp.flush();
    return { success: true, message: `Item ${trnId} dipindahkan ke Rebagging (${payload.gudangTujuan})!` };
  }

  SpreadsheetApp.flush();
  return { success: true, message: `Tindakan ${payload.tindakan} diperbarui pada ${trnId}.` };
}

/**
 * 8. SELESAIKAN REBAGGING -> MASUK FG SIAP KIRIM
 */
function completeRebagging(payload, username) {
  const ss = getSS();
  const rebagSheet = ss.getSheetByName('Rebagging');
  const row = Number(payload.row);

  rebagSheet.getRange(row, 8).setValue(Number(payload.qtyFG) || 0);
  rebagSheet.getRange(row, 9).setValue(payload.gudangFG);
  rebagSheet.getRange(row, 10).setValue('Selesai - Masuk FG');
  rebagSheet.getRange(row, 15).setValue(`${username} (${nowTimestamp()})`);

  SpreadsheetApp.flush();
  return { success: true, message: `Rebagging selesai! Stok FG siap kirim bertambah di ${payload.gudangFG}.` };
}

/**
 * 9. INPUT BARANG KELUAR (FG SIAP KIRIM -> PENGIRIMAN)
 */
function saveBarangKeluar(payload, username) {
  const ss = getSS();
  const keluarSheet = ss.getSheetByName('Keluar');
  const existing = keluarSheet.getRange('A2:A500').getValues();
  let nextRow = 2;
  for (let i = 0; i < existing.length; i++) {
    if (!String(existing[i][0] || '').trim()) { nextRow = 2 + i; break; }
  }

  const satuan = payload.satuan || 'Kg';
  const ts = nowTimestamp();

  keluarSheet.getRange(nextRow, 1, 1, 12).setValues([[
    generateId('KLR', keluarSheet),
    payload.tanggal,
    payload.namaBarang,
    String(payload.nopol || '').toUpperCase(),
    Number(payload.qty) || 0,
    payload.asalGudang,
    payload.tujuan || '-',
    satuan,
    `=IF(E${nextRow}="","",IF(H${nextRow}="Ton",E${nextRow}*1000,E${nextRow}))`,
    username,
    '-',
    ts
  ]]);

  SpreadsheetApp.flush();
  return { success: true, message: `Pengiriman ${payload.namaBarang} (${payload.qty} ${satuan}) Nopol ${payload.nopol} dicatat!` };
}

/**
 * 10. EDIT & HAPUS TRANSAKSI DENGAN HISTORI USER
 */
function updateTransaksi(payload, username) {
  const ss = getSS();
  const sheetName = payload.sheetName;
  const row = Number(payload.row);
  const sheet = ss.getSheetByName(sheetName);
  const ts = nowTimestamp();
  const editorTag = `${username} (${ts})`;

  if (!sheet || row < 2 || row > sheet.getLastRow()) {
    return { success: false, message: 'Data transaksi tidak ditemukan.' };
  }

  if (sheetName === 'Masuk') {
    sheet.getRange(row, 2, 1, 7).setValues([[
      payload.tanggal,
      payload.shiftLine,
      payload.namaBarang,
      Number(payload.qty) || 0,
      payload.satuan || 'Kg',
      payload.keterangan,
      payload.gudang
    ]]);
    sheet.getRange(row, 11, 1, 2).setValues([[username, ts]]);
  } else if (sheetName === 'Status') {
    sheet.getRange(row, 2, 1, 5).setValues([[
      payload.tanggal,
      payload.shiftLine,
      payload.namaBarang,
      Number(payload.qty) || 0,
      payload.statusQC
    ]]);
    sheet.getRange(row, 8).setValue(payload.gudang);
    sheet.getRange(row, 10, 1, 2).setValues([[payload.satuan || 'Kg', payload.catatan || '']]);
    sheet.getRange(row, 13).setValue(editorTag);
  } else if (sheetName === 'Transisi') {
    sheet.getRange(row, 3, 1, 8).setValues([[
      payload.tanggal,
      payload.shiftLine,
      payload.namaBarang,
      Number(payload.qty) || 0,
      payload.asalGudang,
      payload.tindakan,
      payload.gudangTujuan,
      payload.statusProses
    ]]);
    sheet.getRange(row, 12).setValue(payload.satuan || 'Kg');
    sheet.getRange(row, 14).setValue(editorTag);
  } else if (sheetName === 'Rebagging') {
    sheet.getRange(row, 4, 1, 8).setValues([[
      payload.tanggal,
      payload.shiftLine,
      payload.namaBarang,
      Number(payload.qtyMasuk) || 0,
      Number(payload.qtyFG) || 0,
      payload.gudangFG,
      payload.status,
      payload.satuan || 'Kg'
    ]]);
    sheet.getRange(row, 15).setValue(editorTag);
  } else if (sheetName === 'Keluar') {
    sheet.getRange(row, 2, 1, 7).setValues([[
      payload.tanggal,
      payload.namaBarang,
      String(payload.nopol || '').toUpperCase(),
      Number(payload.qty) || 0,
      payload.asalGudang,
      payload.tujuan || '-',
      payload.satuan || 'Kg'
    ]]);
    sheet.getRange(row, 11, 1, 2).setValues([[username, ts]]);
  }

  SpreadsheetApp.flush();
  return { success: true, message: `Transaksi pada ${sheetName} berhasil diperbarui oleh ${username}!` };
}

function deleteTransaksi(sheetName, row) {
  const ss = getSS();
  const sheet = ss.getSheetByName(sheetName);
  const rowNum = Number(row);
  if (!sheet || rowNum < 2 || rowNum > sheet.getLastRow()) {
    return { success: false, message: 'Baris transaksi tidak valid.' };
  }
  sheet.deleteRow(rowNum);
  SpreadsheetApp.flush();
  return { success: true, message: `Transaksi pada ${sheetName} berhasil dihapus!` };
}