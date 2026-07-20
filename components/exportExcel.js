import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

// Lebar kolom default untuk Kartu Stok per-produk (13 kolom):
// Tanggal, Referensi, Masuk(unit), Masuk(Kg), Keluar(unit), Keluar(Kg),
// Sisa(unit), Total(Kg), Lokasi, Lot No, Pallet Eksport, Pallet Lokal, Keterangan
const KARTU_STOK_COLUMN_WIDTHS = [15, 20, 12, 12, 12, 12, 12, 12, 18, 18, 15, 15, 35];

// Lebar kolom untuk Resume Bulanan (7 kolom):
// Produk, Masuk(Satuan), Masuk(Kg), Keluar(Satuan), Keluar(Kg), Sisa(Satuan), Total(Kg)
const RESUME_COLUMN_WIDTHS = [26, 14, 12, 14, 12, 12, 12];

// =====================================================
// FUNGSI BERSAMA: styling 1 worksheet
// =====================================================

function styleWorksheet(worksheet, rows, titleText, sheetName, options = {}) {
  const columnWidths = options.columnWidths || KARTU_STOK_COLUMN_WIDTHS;
  const mainTitle = options.mainTitle || "KARTU STOK PARAFORMALDEHYDE";

  worksheet.columns = columnWidths.map((width) => ({ width }));

  const totalColumns = worksheet.columns.length;
  const lastColLetter = worksheet.getColumn(totalColumns).letter;

  worksheet.views = [
    {
      state: "frozen",
      ySplit: 4,
    },
  ];

  // =====================================================
  // JUDUL
  // =====================================================

  worksheet.mergeCells(`A1:${lastColLetter}1`);

  const titleCell = worksheet.getCell("A1");

  titleCell.value = mainTitle;

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

  // Sub-judul (nama produk + periode, atau label bulan resume)
  worksheet.mergeCells(`A2:${lastColLetter}2`);

  const productCell = worksheet.getCell("A2");

  productCell.value = titleText || sheetName;

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
  // MASUKKAN DATA
  // =====================================================
  // rows[0] = header tabel -> jadi row 4
  // rows[1..] = data -> mulai row 5

  rows.forEach((row) => {
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
          // numFmt masih ikut regional Windows/Excel si user (bisa jadi titik lagi),
          // jadi angkanya langsung diubah ke teks berformat koma yang fixed, tidak ikut regional.
          cell.value = cell.value.toLocaleString("en-US");
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
  // SUB TOTAL BULANAN & TOTAL AKHIR (styling khusus)
  // =====================================================

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber < 5) return;

    const firstCellValue = row.getCell(1).value;
    const isMonthlySubtotal =
      typeof firstCellValue === "string" && firstCellValue.startsWith("SUB TOTAL ");

    if (isMonthlySubtotal) {
      // Label "SUB TOTAL BULAN TAHUN" digabung di kolom A:B, seluruh baris diwarnai hijau
      worksheet.mergeCells(`A${rowNumber}:B${rowNumber}`);

      for (let col = 1; col <= totalColumns; col++) {
        const cell = row.getCell(col);

        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF34A853" },
        };

        cell.font = {
          bold: true,
          color: { argb: "FFFFFFFF" },
          name: "Times New Roman",
        };

        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      }

      row.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
    }
  });

  // =====================================================
  // AUTO FILTER
  // =====================================================

  worksheet.autoFilter = {
    from: "A4",
    to: `${lastColLetter}4`,
  };
}

// =====================================================
// EXPORT 1 PRODUK (satu sheet) - Kartu Stok
// =====================================================

export async function exportExcelProfessional(
  rows,
  sheetName,
  fileName,
  titleText,
  options
) {
  const workbook = new ExcelJS.Workbook();

  workbook.creator = "Prototype Kartu Stok";
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet(sheetName);

  styleWorksheet(worksheet, rows, titleText, sheetName, options);

  const buffer = await workbook.xlsx.writeBuffer();

  saveAs(new Blob([buffer]), fileName);
}

// =====================================================
// EXPORT SEMUA PRODUK (banyak sheet, 1 file) - Kartu Stok
// =====================================================
// sheetsData: array of { sheetName, rows, titleText }

export async function exportExcelAllProducts(sheetsData, fileName, options) {
  const workbook = new ExcelJS.Workbook();

  workbook.creator = "Prototype Kartu Stok";
  workbook.created = new Date();

  for (const { sheetName, rows, titleText } of sheetsData) {
    const worksheet = workbook.addWorksheet(sheetName);
    styleWorksheet(worksheet, rows, titleText, sheetName, options);
  }

  const buffer = await workbook.xlsx.writeBuffer();

  saveAs(new Blob([buffer]), fileName);
}

// =====================================================
// EXPORT RESUME BULANAN (satu sheet, 9 kolom, judul beda)
// =====================================================

export async function exportExcelResume(rows, titleText, fileName) {
  return exportExcelProfessional(rows, "Resume Bulanan", fileName, titleText, {
    columnWidths: RESUME_COLUMN_WIDTHS,
    mainTitle: "REKAP BULANAN STOK PARAFORMALDEHYDE",
  });
}