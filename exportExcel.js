import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

export async function exportExcelProfessional(
  rows,
  sheetName,
  fileName
) {
  const workbook = new ExcelJS.Workbook();

  workbook.creator = "Prototype Kartu Stok";
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet(sheetName);

  worksheet.views = [
    {
      state: "frozen",
      ySplit: 4,
    },
  ];

  worksheet.columns = [
    { width: 15 }, // Tanggal
    { width: 20 }, // Referensi
    { width: 12 }, // Masuk Zak
    { width: 12 }, // Masuk Kg
    { width: 12 }, // Keluar Zak
    { width: 12 }, // Keluar Kg
    { width: 12 }, // Sisa
    { width: 12 }, // Total
    { width: 18 }, // Lokasi
    { width: 18 }, // Lot
    { width: 15 }, // Pallet
    { width: 35 }, // Keterangan
  ];

// =====================================================
// JUDUL
// =====================================================

worksheet.mergeCells("A1:L1");

const titleCell = worksheet.getCell("A1");

titleCell.value = "KARTU STOK PARAFORMALDEHYDE";

titleCell.font = {
  bold: true,
  size: 16,
  name: "Times New Roman",
};

titleCell.alignment = {
  horizontal: "center",
  vertical: "middle",
};

worksheet.getRow(1).height = 28;

// Nama Produk
worksheet.mergeCells("A2:L2");

const productCell = worksheet.getCell("A2");

// rows[0][0] berisi teks lengkap "Nama Produk - Periode" (mis. "PFA 86% - Semua Periode")
// dipakai langsung di sini supaya tidak perlu baris terpisah lagi di bawah
productCell.value = (rows[0] && rows[0][0]) ? rows[0][0] : sheetName;

productCell.font = {
  bold: true,
  size: 13,
  name: "Times New Roman",
};

productCell.alignment = {
  horizontal: "center",
};

worksheet.getRow(2).height = 22;

// Baris kosong (row 3)
worksheet.addRow([]);

// =====================================================
// MASUKKAN DATA DARI buildSheetRows()
// =====================================================
// rows[0] = [productName]  -> sudah dipakai di judul row 2, jangan ditambahkan lagi
// rows[1] = []             -> baris kosong duplikat, dilewati juga
// rows[2] = header, dst    -> baru ditambahkan mulai row 4

rows.slice(2).forEach((row) => {
  worksheet.addRow(row);
});

// =====================================================
// HEADER
// =====================================================

const headerRow = worksheet.getRow(4);

headerRow.font = {
  bold: true,
  color: {
    argb: "000000",
  },
};

headerRow.alignment = {
  horizontal: "center",
  vertical: "middle",
  wrapText: true,
};

const totalColumns = worksheet.columns.length; // 12 kolom (A-L)

for (let col = 1; col <= totalColumns; col++) {
  const cell = headerRow.getCell(col);

  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: {
      argb: "D9EAD3",
    },
  };

  cell.border = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };
}

// =====================================================
// BORDER SEMUA DATA
// =====================================================

worksheet.eachRow((row, rowNumber) => {

  if (rowNumber >= 5) {

    row.eachCell((cell) => {

      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      if (typeof cell.value === "number") {
        cell.alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      } else {
        cell.alignment = {
          vertical: "middle",
        };
      }

    });

  }

});

// =====================================================
// AUTO FILTER
// =====================================================

worksheet.autoFilter = {
  from: "A4",
  to: "L4",
};

// =====================================================
// SIMPAN FILE
// =====================================================

const buffer = await workbook.xlsx.writeBuffer();

saveAs(
  new Blob([buffer]),
  fileName
);

}