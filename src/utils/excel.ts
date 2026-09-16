import * as XLSX from 'xlsx';

export interface ScanSummary {
  matched: string[];
  missing: string[];
  extra: string[];
  damaged: string[];
}

// 1. Đọc mã vận đơn từ file Excel tải lên
export async function parseManifestExcel(file: File): Promise<string[]> {
  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const wsName = wb.SheetNames[0];
  const ws = wb.Sheets[wsName];

  const rawRows: any[] = XLSX.utils.sheet_to_json(ws, { defval: '' });
  if (rawRows.length === 0) return [];

  // Tìm cột mã vận đơn tự động
  const headers = Object.keys(rawRows[0]);
  const trackingKey =
    headers.find((k) =>
      /tracking|waybill|mã vận đơn|mã đơn|order|spx/i.test(k)
    ) || headers[0];

  const codes = rawRows
    .map((row) => String(row[trackingKey] || '').trim())
    .filter((code) => code.length >= 5);

  return Array.from(new Set(codes));
}

// 2. Xuất file Excel biên bản sự vụ (gồm nhiều Sheet)
export function exportDiscrepancyReport(summary: ScanSummary, batchName: string) {
  const wb = XLSX.utils.book_new();

  const createSheet = (data: string[], colName: string) =>
    XLSX.utils.json_to_sheet(data.map((c) => ({ [colName]: c })));

  // Sheet 1: Đơn thiếu (Missed)
  XLSX.utils.book_append_sheet(
    wb,
    createSheet(summary.missing, 'MÃ ĐƠN THIẾU (CHƯA NHẬN)'),
    'Đơn Thiếu'
  );

  // Sheet 2: Đơn thừa (Extra)
  XLSX.utils.book_append_sheet(
    wb,
    createSheet(summary.extra, 'MÃ ĐƠN THỪA (NGOÀI MANIFEST)'),
    'Đơn Thừa'
  );

  // Sheet 3: Đơn hư hỏng/bể vỡ (Damaged)
  XLSX.utils.book_append_sheet(
    wb,
    createSheet(summary.damaged, 'MÃ ĐƠN SỰ VỤ / HƯ HỎNG'),
    'Đơn Bể Vỡ'
  );

  // Sheet 4: Đơn đã nhận đủ
  XLSX.utils.book_append_sheet(
    wb,
    createSheet(summary.matched, 'MÃ ĐƠN NHẬN THÀNH CÔNG'),
    'Đơn Khớp'
  );

  XLSX.writeFile(wb, `Bien_Ban_Su_Vu_${batchName || 'SPX'}.xlsx`);
}