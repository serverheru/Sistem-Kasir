function simpanTransaksi(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const transaksiSheet = getSheet("Transaksi");
    const detailSheet = getSheet("DetailTransaksi");

    if (!transaksiSheet) {
      throw new Error("Sheet Transaksi tidak ditemukan.");
    }

    if (!detailSheet) {
      throw new Error("Sheet DetailTransaksi tidak ditemukan.");
    }

    const items = normalisasiItems(payload && payload.items);
    const bayar = Number(payload && payload.bayar);

    if (items.length === 0) {
      throw new Error("Keranjang masih kosong.");
    }

    if (!bayar || bayar <= 0) {
      throw new Error("Nominal bayar tidak valid.");
    }

    const idTransaksi = buatIdTransaksi();
    const tanggal = new Date();
    const detailRows = [];
    let total = 0;

    items.forEach(item => {
      const produk = getProdukById(item.id_produk);

      if (!produk) {
        throw new Error("Produk tidak ditemukan: " + item.id_produk);
      }

      if (produk.stok < item.qty) {
        throw new Error("Stok tidak cukup untuk produk: " + produk.nama_produk);
      }

      const subtotal = produk.harga * item.qty;
      total += subtotal;

      detailRows.push([
        idTransaksi,
        produk.id_produk,
        produk.nama_produk,
        produk.harga,
        item.qty,
        subtotal
      ]);
    });

    if (bayar < total) {
      throw new Error("Uang bayar kurang.");
    }

    const kembalian = bayar - total;

    transaksiSheet.appendRow([
      idTransaksi,
      tanggal,
      total,
      bayar,
      kembalian
    ]);

    if (detailRows.length > 0) {
      detailSheet
        .getRange(detailSheet.getLastRow() + 1, 1, detailRows.length, detailRows[0].length)
        .setValues(detailRows);
    }

    items.forEach(item => {
      updateStokProduk(item.id_produk, item.qty);
    });

    return {
      status: true,
      message: "Transaksi berhasil disimpan.",
      id_transaksi: idTransaksi,
      tanggal: Utilities.formatDate(tanggal, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss"),
      total: total,
      bayar: bayar,
      kembalian: kembalian
    };
  } finally {
    lock.releaseLock();
  }
}

function normalisasiItems(items) {
  if (!items || !Array.isArray(items)) {
    return [];
  }

  const map = {};

  items.forEach(item => {
    const idProduk = item && String(item.id_produk || "").trim();
    const qty = Number(item && item.qty);

    if (!idProduk || qty <= 0 || !Number.isInteger(qty)) {
      throw new Error("Item transaksi tidak valid.");
    }

    map[idProduk] = (map[idProduk] || 0) + qty;
  });

  return Object.keys(map).map(idProduk => {
    return {
      id_produk: idProduk,
      qty: map[idProduk]
    };
  });
}

function buatIdTransaksi() {
  const waktu = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyyMMdd-HHmmss");
  const acak = Math.floor(Math.random() * 900 + 100);
  return "TRX-" + waktu + "-" + acak;
}

function getRiwayatTransaksi() {
  const sheet = getSheet("Transaksi");

  if (!sheet) {
    throw new Error("Sheet Transaksi tidak ditemukan.");
  }

  const data = sheet.getDataRange().getValues();

  data.shift();

  const jumlahItemByTransaksi = getJumlahItemByTransaksi();

  return data
    .filter(row => row[0])
    .reverse()
    .slice(0, 100)
    .map(row => {
      const idTransaksi = row[0];

      return {
        id_transaksi: idTransaksi,
        tanggal: Utilities.formatDate(new Date(row[1]), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss"),
        total: Number(row[2]) || 0,
        bayar: Number(row[3]) || 0,
        kembalian: Number(row[4]) || 0,
        jumlah_item: jumlahItemByTransaksi[idTransaksi] || 0
      };
    });
}

function getRingkasanHariIni() {
  const transaksiSheet = getSheet("Transaksi");
  const detailSheet = getSheet("DetailTransaksi");
  const ringkasan = {
    jumlah_transaksi: 0,
    omzet: 0,
    jumlah_item: 0
  };

  if (!transaksiSheet) {
    throw new Error("Sheet Transaksi tidak ditemukan.");
  }

  const transaksiData = transaksiSheet.getDataRange().getValues();
  transaksiData.shift();

  const timeZone = "Asia/Jakarta";
  const todayKey = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
  const transaksiHariIni = {};

  transaksiData.forEach(row => {
    const idTransaksi = row[0];
    const tanggal = row[1];

    if (!idTransaksi || !tanggal) {
      return;
    }

    const tanggalKey = Utilities.formatDate(new Date(tanggal), timeZone, "yyyy-MM-dd");

    if (tanggalKey === todayKey) {
      transaksiHariIni[idTransaksi] = true;
      ringkasan.jumlah_transaksi += 1;
      ringkasan.omzet += Number(row[2]) || 0;
    }
  });

  if (!detailSheet || Object.keys(transaksiHariIni).length === 0) {
    return ringkasan;
  }

  const detailData = detailSheet.getDataRange().getValues();
  detailData.shift();

  detailData.forEach(row => {
    const idTransaksi = row[0];

    if (transaksiHariIni[idTransaksi]) {
      ringkasan.jumlah_item += Number(row[4]) || 0;
    }
  });

  return ringkasan;
}

function getJumlahItemByTransaksi() {
  const sheet = getSheet("DetailTransaksi");
  const hasil = {};

  if (!sheet) {
    return hasil;
  }

  const data = sheet.getDataRange().getValues();
  data.shift();

  data.forEach(row => {
    const idTransaksi = row[0];
    const qty = Number(row[4]) || 0;

    if (idTransaksi) {
      hasil[idTransaksi] = (hasil[idTransaksi] || 0) + qty;
    }
  });

  return hasil;
}
