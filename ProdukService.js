function getProduk() {
  const sheet = getSheet("Produk");

  if (!sheet) {
    throw new Error("Sheet Produk tidak ditemukan.");
  }

  const range = sheet.getDataRange();
  const data = range.getValues();
  const richTextData = range.getRichTextValues();
  const header = data[0];
  const kolom = getKolomProduk(header);

  return data
    .slice(1)
    .filter(row => row[kolom.id_produk] && row[kolom.nama_produk])
    .map((row, index) => {
      const dataRowIndex = index + 1;
      const gambar = getNilaiGambarProduk(
        row[kolom.gambar],
        richTextData[dataRowIndex] ? richTextData[dataRowIndex][kolom.gambar] : null
      );

      return {
        id_produk: String(row[kolom.id_produk]),
        nama_produk: String(row[kolom.nama_produk]),
        harga: Number(row[kolom.harga]) || 0,
        stok: Number(row[kolom.stok]) || 0,
        gambar: getUrlGambarProduk(gambar),
        gambar_asli: gambar
      };
    });
}

function getNilaiGambarProduk(value, richTextValue) {
  const richTextLink = getRichTextLinkUrl(richTextValue);

  if (richTextLink) {
    return richTextLink;
  }

  return value ? String(value).trim() : "";
}

function getRichTextLinkUrl(richTextValue) {
  if (!richTextValue) {
    return "";
  }

  const fullLink = richTextValue.getLinkUrl();

  if (fullLink) {
    return fullLink;
  }

  const runs = richTextValue.getRuns ? richTextValue.getRuns() : [];

  for (let i = 0; i < runs.length; i++) {
    const runLink = runs[i].getLinkUrl();

    if (runLink) {
      return runLink;
    }
  }

  return "";
}

function getUrlGambarProduk(gambar) {
  if (!gambar) {
    return "";
  }

  if (gambar.indexOf("data:image/") === 0) {
    return gambar;
  }

  const fileId = getGoogleDriveFileId(gambar);

  if (fileId) {
    return getDriveImageAsDataUrl(fileId, gambar);
  }

  if (isNamaFileGambar(gambar)) {
    const fileByName = getDriveFileIdByName(gambar);

    if (fileByName) {
      return getDriveImageAsDataUrl(fileByName, gambar);
    }
  }

  return gambar;
}

function getDriveImageAsDataUrl(fileId, fallbackValue) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const contentType = blob.getContentType();

    if (contentType.indexOf("image/") !== 0) {
      return fallbackValue;
    }

    return "data:" + contentType + ";base64," + Utilities.base64Encode(blob.getBytes());
  } catch (error) {
    return fallbackValue;
  }
}

function getDriveFileIdByName(fileName) {
  try {
    const files = DriveApp.getFilesByName(fileName);

    if (files.hasNext()) {
      return files.next().getId();
    }
  } catch (error) {
    return "";
  }

  return "";
}

function isNamaFileGambar(value) {
  return /\.(png|jpe?g|webp|gif)$/i.test(String(value || "").trim());
}

function getGoogleDriveFileId(url) {
  const text = String(url || "");
  const filePathMatch = text.match(/drive\.google\.com\/file\/d\/([^/]+)/i);

  if (filePathMatch && filePathMatch[1]) {
    return filePathMatch[1];
  }

  const queryMatch = text.match(/[?&]id=([^&]+)/i);

  if (text.indexOf("drive.google.com") !== -1 && queryMatch && queryMatch[1]) {
    return queryMatch[1];
  }

  return "";
}

function getKolomProduk(header) {
  const indexByName = {};

  header.forEach((namaKolom, index) => {
    indexByName[String(namaKolom).trim().toLowerCase()] = index;
  });

  return {
    id_produk: getIndexKolom(indexByName, "id_produk", 0),
    nama_produk: getIndexKolom(indexByName, "nama_produk", 1),
    harga: getIndexKolom(indexByName, "harga", 2),
    stok: getIndexKolom(indexByName, "stok", 3),
    gambar: getIndexKolom(indexByName, "gambar", 4)
  };
}

function getIndexKolom(indexByName, namaKolom, fallbackIndex) {
  return Object.prototype.hasOwnProperty.call(indexByName, namaKolom)
    ? indexByName[namaKolom]
    : fallbackIndex;
}

function simpanProduk(payload) {
  const sheet = getSheet("Produk");

  if (!sheet) {
    throw new Error("Sheet Produk tidak ditemukan.");
  }

  const produk = normalisasiPayloadProduk(payload);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const kolom = getKolomProduk(header);
  const originalId = payload && payload.original_id ? String(payload.original_id).trim() : "";
  let targetRow = -1;

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][kolom.id_produk] || "").trim();

    if (id === produk.id_produk && (!originalId || originalId !== produk.id_produk)) {
      throw new Error("ID produk sudah digunakan.");
    }

    if (originalId && id === originalId) {
      targetRow = i + 1;
    }
  }

  if (originalId) {
    if (targetRow < 0) {
      throw new Error("Produk yang akan diubah tidak ditemukan.");
    }

    const rowValues = buatRowProduk(produk, kolom, data[targetRow - 1].slice(), header.length);
    sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    return {
      status: true,
      message: "Produk berhasil diperbarui."
    };
  }

  const rowValues = buatRowProduk(produk, kolom, [], header.length);
  sheet.appendRow(rowValues);
  return {
    status: true,
    message: "Produk berhasil ditambahkan."
  };
}

function hapusProduk(idProduk) {
  const sheet = getSheet("Produk");

  if (!sheet) {
    throw new Error("Sheet Produk tidak ditemukan.");
  }

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const kolom = getKolomProduk(header);
  const idTarget = String(idProduk || "").trim();

  if (!idTarget) {
    throw new Error("ID produk tidak valid.");
  }

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][kolom.id_produk] || "").trim();

    if (id === idTarget) {
      sheet.deleteRow(i + 1);
      return {
        status: true,
        message: "Produk berhasil dihapus."
      };
    }
  }

  throw new Error("Produk tidak ditemukan.");
}

function normalisasiPayloadProduk(payload) {
  const idProduk = payload && String(payload.id_produk || "").trim();
  const namaProduk = payload && String(payload.nama_produk || "").trim();
  const harga = Number(payload && payload.harga);
  const stok = Number(payload && payload.stok);
  const gambar = payload && payload.gambar ? String(payload.gambar).trim() : "";

  if (!idProduk) {
    throw new Error("ID produk wajib diisi.");
  }

  if (!namaProduk) {
    throw new Error("Nama produk wajib diisi.");
  }

  if (harga < 0 || !Number.isFinite(harga)) {
    throw new Error("Harga produk tidak valid.");
  }

  if (stok < 0 || !Number.isInteger(stok)) {
    throw new Error("Stok produk harus angka bulat minimal 0.");
  }

  return {
    id_produk: idProduk,
    nama_produk: namaProduk,
    harga: harga,
    stok: stok,
    gambar: gambar
  };
}

function buatRowProduk(produk, kolom, existingRow, columnCount) {
  const row = existingRow.length > 0
    ? existingRow
    : new Array(Math.max(columnCount, 5)).fill("");

  while (row.length < Math.max(columnCount, 5)) {
    row.push("");
  }

  row[kolom.id_produk] = produk.id_produk;
  row[kolom.nama_produk] = produk.nama_produk;
  row[kolom.harga] = produk.harga;
  row[kolom.stok] = produk.stok;
  row[kolom.gambar] = produk.gambar;
  return row;
}

function getProdukById(idProduk) {
  const produkList = getProduk();
  return produkList.find(produk => String(produk.id_produk) === String(idProduk));
}

function updateStokProduk(idProduk, qtyTerjual) {
  const sheet = getSheet("Produk");

  if (!sheet) {
    throw new Error("Sheet Produk tidak ditemukan.");
  }

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const kolom = getKolomProduk(header);

  for (let i = 1; i < data.length; i++) {
    const id = data[i][kolom.id_produk];

    if (String(id) === String(idProduk)) {
      const stokSekarang = Number(data[i][kolom.stok]);
      const stokBaru = stokSekarang - qtyTerjual;

      if (stokBaru < 0) {
        throw new Error("Stok produk tidak cukup.");
      }

      sheet.getRange(i + 1, kolom.stok + 1).setValue(stokBaru);
      return true;
    }
  }

  throw new Error("Produk tidak ditemukan.");
}
