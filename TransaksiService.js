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
    const namaPembeli = payload && payload.nama_pembeli ? String(payload.nama_pembeli).trim() : "-";
    const metodePembayaran = payload && payload.metode_pembayaran ? String(payload.metode_pembayaran).trim() : "Tunai";

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
    const allProducts = getProduk();

    items.forEach(item => {
      const produk = allProducts.find(p => String(p.id_produk) === String(item.id_produk));

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
        subtotal,
        namaPembeli,
        metodePembayaran
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
      kembalian,
      namaPembeli,
      metodePembayaran
    ]);

    if (detailRows.length > 0) {
      detailSheet
        .getRange(detailSheet.getLastRow() + 1, 1, detailRows.length, detailRows[0].length)
        .setValues(detailRows);
    }

    updateStokProdukBatch(items);

    return {
      status: true,
      message: "Transaksi berhasil disimpan.",
      id_transaksi: idTransaksi,
      nama_pembeli: namaPembeli,
      tanggal: Utilities.formatDate(tanggal, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss"),
      total: total,
      bayar: bayar,
      kembalian: kembalian,
      metode_pembayaran: metodePembayaran
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
        nama_pembeli: row[5] || "-",
        total: Number(row[2]) || 0,
        bayar: Number(row[3]) || 0,
        kembalian: Number(row[4]) || 0,
        jumlah_item: jumlahItemByTransaksi[idTransaksi] || 0,
        metode_pembayaran: row[6] || "Tunai"
      };
    });
}

function getRingkasanHariIni() {
  const transaksiSheet = getSheet("Transaksi");
  const detailSheet = getSheet("DetailTransaksi");
  const ringkasan = {
    jumlah_transaksi: 0,
    omzet: 0,
    omzet_tunai: 0,
    omzet_qris: 0,
    jumlah_item: 0,
    produk_terlaris: "-"
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
      const tTotal = Number(row[2]) || 0;
      ringkasan.omzet += tTotal;
      const metode = row[6] || "Tunai";
      if (metode === "QRIS") {
        ringkasan.omzet_qris += tTotal;
      } else {
        ringkasan.omzet_tunai += tTotal;
      }
    }
  });

  if (!detailSheet || Object.keys(transaksiHariIni).length === 0) {
    return ringkasan;
  }

  const detailData = detailSheet.getDataRange().getValues();
  detailData.shift();
  
  const qtyPerProduk = {};

  detailData.forEach(row => {
    const idTransaksi = row[0];

    if (transaksiHariIni[idTransaksi]) {
      const namaProduk = row[2] || "-";
      const qty = Number(row[4]) || 0;
      ringkasan.jumlah_item += qty;
      qtyPerProduk[namaProduk] = (qtyPerProduk[namaProduk] || 0) + qty;
    }
  });

  let maxQty = 0;
  for (const p in qtyPerProduk) {
    if (qtyPerProduk[p] > maxQty) {
      maxQty = qtyPerProduk[p];
      ringkasan.produk_terlaris = p;
    }
  }

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
