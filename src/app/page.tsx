'use client';

import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';

type Phase = 'IMPORT_MANIFEST' | 'SCAN_BAG' | 'SCAN_ITEMS' | 'REPORT';

interface BagDetail {
  toCode: string;
  expectedCodes: string[];
}

interface ScanRecord {
  code: string;
  type: 'VALID' | 'MISPLACED' | 'EXTRA' | 'DUPLICATE';
  targetBag?: string; // Bao chính xác của kiện này (nếu bị lạc bao)
  time: string;
}

export default function DebaggingApp() {
  const [phase, setPhase] = useState<Phase>('IMPORT_MANIFEST');

  // Quản lý đa bao
  const [bagsMap, setBagsMap] = useState<Map<string, string[]>>(new Map());
  // Bảng tra cứu ngược: Mã kiện -> Mã bao sở hữu
  const [packageToBagMap, setPackageToBagMap] = useState<Map<string, string>>(new Map());

  // Bao đang thao tác
  const [currentBagCode, setCurrentBagCode] = useState<string>('');
  const [expectedSet, setExpectedSet] = useState<Set<string>>(new Set());

  // Các tập hợp kết quả của bao hiện tại
  const [scannedSet, setScannedSet] = useState<Set<string>>(new Set());
  const [misplacedSet, setMisplacedSet] = useState<Map<string, string>>(new Map()); // Mã kiện -> Thuộc bao nào
  const [extraSet, setExtraSet] = useState<Set<string>>(new Set()); // Kiện lạ không thuộc bất kỳ bao nào

  // Input nhập tay
  const [manualInput, setManualInput] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);

  // Lịch sử quét gần nhất
  const [recentScans, setRecentScans] = useState<ScanRecord[]>([]);

  // 1. Âm thanh Web Audio API
  const playTone = (type: 'success' | 'warn' | 'error') => {
    try {
      const AudioCtx = window.AudioContext || (window as unknown as { webkitAudioContext: typeof AudioContext }).webkitAudioContext;
      const ctx = new AudioCtx();
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain);
      gain.connect(ctx.destination);

      if (type === 'success') {
        osc.frequency.setValueAtTime(880, ctx.currentTime);
        gain.gain.setValueAtTime(0.15, ctx.currentTime);
        osc.start();
        osc.stop(ctx.currentTime + 0.08);
      } else if (type === 'warn') {
        osc.frequency.setValueAtTime(440, ctx.currentTime);
        gain.gain.setValueAtTime(0.2, ctx.currentTime);
        osc.start();
        osc.stop(ctx.currentTime + 0.15);
      } else {
        osc.frequency.setValueAtTime(200, ctx.currentTime);
        gain.gain.setValueAtTime(0.35, ctx.currentTime);
        osc.start();
        osc.stop(ctx.currentTime + 0.3);
      }
    } catch {
      // Trình duyệt chặn autoplay thì bỏ qua
    }
  };

  // 2. Đọc nhiều file Excel FMS cùng một lúc (Multi-file upload)
  const handleMultipleFilesUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const newBagsMap = new Map<string, string[]>();
    const newPkgMap = new Map<string, string>();

    try {
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' });

        if (rows.length === 0) continue;

        let detectedTO = '';
        let trackingColIndex = -1;
        let headerRowIndex = -1;

        // Quét tìm TO Number và cột LM Tracking Number
        for (let r = 0; r < rows.length; r++) {
          const row = rows[r];
          const firstCell = String(row[0] || '').trim();

          if (/^TO Number$/i.test(firstCell) && !detectedTO) {
            detectedTO = String(row[1] || '').trim();
          }

          for (let c = 0; c < row.length; c++) {
            const cellVal = String(row[c] || '').trim();
            if (/LM Tracking Number|Tracking Number/i.test(cellVal)) {
              trackingColIndex = c;
              headerRowIndex = r;
              break;
            }
          }

          if (headerRowIndex !== -1 && detectedTO) break;
        }

        // Lấy mã bao (hoặc dùng tên file nếu không đọc được dòng TO Number)
        const bagKey = detectedTO || file.name.replace(/\.[^/.]+$/, '');

        // Trích xuất mã: Chấp nhận cả SPX... và VN...
        const codes: string[] = [];
        if (headerRowIndex !== -1 && trackingColIndex !== -1) {
          for (let r = headerRowIndex + 1; r < rows.length; r++) {
            const val = String(rows[r][trackingColIndex] || '').trim().toUpperCase();
            // Regex chấp nhận mã bắt đầu bằng SPX hoặc VN với tối thiểu 8 ký tự
            if (val && /^(SPX|VN)/i.test(val)) {
              codes.push(val);
              newPkgMap.set(val, bagKey);
            }
          }
        }

        if (codes.length > 0) {
          newBagsMap.set(bagKey, Array.from(new Set(codes)));
        }
      }

      if (newBagsMap.size === 0) {
        alert('Không tìm thấy dữ liệu kiện hợp lệ trong các file đã chọn!');
        return;
      }

      setBagsMap(newBagsMap);
      setPackageToBagMap(newPkgMap);
      setPhase('SCAN_BAG');
    } catch {
      alert('Đã xảy ra lỗi khi đọc các file Excel. Vui lòng thử lại!');
    }
  };

  // 3. Quét hoặc chọn mã bao để kích hoạt xả bao
  const handleBagScanConfirm = (bagCodeInput: string) => {
    const cleanBag = bagCodeInput.trim().toUpperCase();
    if (!cleanBag) return;

    // Tìm bao tương ứng trong danh sách đã nạp
    let matchedBagKey = '';
    for (const key of bagsMap.keys()) {
      if (key.toUpperCase() === cleanBag) {
        matchedBagKey = key;
        break;
      }
    }

    if (!matchedBagKey) {
      // Trường hợp quét mã bao lạ không nằm trong danh sách file đã nạp
      const proceed = confirm(`Bao "${cleanBag}" không có trong danh sách file nạp. Bạn có muốn tạo bao trống để kiểm đếm không?`);
      if (!proceed) return;
      matchedBagKey = cleanBag;
      setExpectedSet(new Set());
    } else {
      setExpectedSet(new Set(bagsMap.get(matchedBagKey) || []));
    }

    setCurrentBagCode(matchedBagKey);
    setScannedSet(new Set());
    setMisplacedSet(new Map());
    setExtraSet(new Set());
    setRecentScans([]);
    setPhase('SCAN_ITEMS');
    playTone('success');
  };

  // 4. Xử lý khi quét hoặc nhập từng mã kiện
  const handleItemScan = (rawCode: string) => {
    const code = rawCode.trim().toUpperCase();
    if (!code || code.length < 5) return;
    const time = new Date().toLocaleTimeString('vi-VN');

    // Trường hợp 1: Quét trùng (đã quét qua rồi)
    if (scannedSet.has(code) || misplacedSet.has(code) || extraSet.has(code)) {
      playTone('warn');
      setRecentScans((prev) => [{ code, type: 'DUPLICATE', time }, ...prev.slice(0, 7)]);
      return;
    }

    // Trường hợp 2: Kiện hợp lệ (nằm đúng trong bao hiện tại)
    if (expectedSet.has(code)) {
      playTone('success');
      setScannedSet((prev) => new Set(prev).add(code));
      setRecentScans((prev) => [{ code, type: 'VALID', time }, ...prev.slice(0, 7)]);
      return;
    }

    // Trường hợp 3: Kiện lạc bao (Thuộc một bao khác trong cùng lô nạp)
    const belongingBag = packageToBagMap.get(code);
    if (belongingBag && belongingBag !== currentBagCode) {
      playTone('warn');
      setMisplacedSet((prev) => new Map(prev).set(code, belongingBag));
      setRecentScans((prev) => [
        { code, type: 'MISPLACED', targetBag: belongingBag, time },
        ...prev.slice(0, 7),
      ]);
      return;
    }

    // Trường hợp 4: Kiện lạ hoàn toàn (không thuộc bất kỳ bao nào)
    playTone('error');
    setExtraSet((prev) => new Set(prev).add(code));
    setRecentScans((prev) => [{ code, type: 'EXTRA', time }, ...prev.slice(0, 7)]);
  };

  // Xử lý gửi mã khi gõ tay hoặc paste
  const handleManualSubmit = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!manualInput.trim()) return;

    handleItemScan(manualInput.trim());
    setManualInput('');
    setTimeout(() => inputRef.current?.focus(), 50);
  };

  // Global listener bắt sự kiện máy quét PDA
  useEffect(() => {
    let buffer = '';
    let lastTime = Date.now();

    const onKeyDown = (e: KeyboardEvent) => {
      if (document.activeElement?.tagName === 'INPUT') return;

      const now = Date.now();
      if (now - lastTime > 60) buffer = '';
      lastTime = now;

      if (e.key === 'Enter') {
        if (buffer.length > 0) {
          if (phase === 'SCAN_BAG') {
            handleBagScanConfirm(buffer);
          } else if (phase === 'SCAN_ITEMS') {
            handleItemScan(buffer);
          }
          buffer = '';
        }
      } else if (e.key.length === 1) {
        buffer += e.key;
      }
    };

    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [phase, scannedSet, extraSet, misplacedSet, expectedSet, currentBagCode]);

  // ==================== BƯỚC 1: NẠP NHIỀU FILE EXCEL TỪ FMS ====================
  if (phase === 'IMPORT_MANIFEST') {
    return (
      <main className="min-h-screen bg-slate-950 text-white p-4 flex flex-col justify-center max-w-md mx-auto">
        <h1 className="text-xl font-bold text-emerald-400 mb-1">Inbound FMS SPX</h1>
        <p className="text-xs text-slate-400 mb-4">
          Tải một hoặc <b>nhiều file Excel</b> xuất từ FMS để tự động nhận diện bao
        </p>

        <div className="border-2 border-dashed border-slate-700 rounded-xl p-8 text-center bg-slate-900">
          <label className="cursor-pointer block">
            <span className="text-sm font-semibold block mb-2">Chọn các file Excel (.xlsx)</span>
            <span className="text-[11px] text-slate-500 block mb-4">Giữ phím Shift hoặc Ctrl để chọn nhiều file</span>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              multiple
              onChange={handleMultipleFilesUpload}
              className="hidden"
            />
            <span className="inline-block bg-emerald-600 hover:bg-emerald-500 text-xs font-bold px-4 py-2.5 rounded-lg shadow-lg">
              Tải lên các file manifest
            </span>
          </label>
        </div>
      </main>
    );
  }

  // ==================== BƯỚC 2: QUÉT HOẶC CHỌN MÃ BAO ====================
  const bagList = Array.from(bagsMap.keys());
  const totalBagsCount = bagList.length;
  const totalPackagesAllBags = Array.from(bagsMap.values()).reduce((sum, list) => sum + list.length, 0);

  if (phase === 'SCAN_BAG') {
    return (
      <main className="min-h-screen bg-slate-950 text-white p-4 flex flex-col max-w-md mx-auto">
        <div className="text-center my-auto">
          <div className="text-emerald-400 text-xs font-bold mb-1">
            ĐÃ NẠP {totalBagsCount} BAO ({totalPackagesAllBags} KIỆN)
          </div>
          <h2 className="text-2xl font-black mb-2">Quét mã bao lớn</h2>
          <p className="text-xs text-slate-400 mb-4">
            Bắn mã QR hoặc mã <b>TO Number</b> (ví dụ: TO2026...)
          </p>

          <input
            id="bagInput"
            placeholder="Chờ máy quét hoặc nhập mã bao..."
            className="w-full bg-slate-900 border border-emerald-500/50 rounded-lg p-3 text-center text-sm font-mono tracking-wider outline-none mb-3"
            onKeyDown={(e) => {
              if (e.key === 'Enter') handleBagScanConfirm((e.target as HTMLInputElement).value);
            }}
          />

          {/* Danh sách các bao có sẵn để bấm nhanh */}
          <div className="mt-4 text-left">
            <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider block mb-2">
              Hoặc chọn nhanh bao ({totalBagsCount}):
            </span>
            <div className="max-h-52 overflow-y-auto space-y-1.5 pr-1">
              {bagList.map((to) => {
                const count = bagsMap.get(to)?.length || 0;
                return (
                  <button
                    key={to}
                    onClick={() => handleBagScanConfirm(to)}
                    className="w-full flex items-center justify-between p-2.5 bg-slate-900 hover:bg-slate-800 border border-slate-800 rounded-lg text-xs font-mono transition-colors"
                  >
                    <span className="font-bold text-slate-200">{to}</span>
                    <span className="text-emerald-400">{count} kiện</span>
                  </button>
                );
              })}
            </div>
          </div>
        </div>
      </main>
    );
  }

  // ==================== BƯỚC 3: QUÉT HOẶC NHẬP KIỆN CON ====================
  const totalExpected = expectedSet.size;
  const matchedCount = scannedSet.size;
  const missingCount = Math.max(0, totalExpected - matchedCount);

  if (phase === 'SCAN_ITEMS') {
    return (
      <div className="flex flex-col h-screen bg-slate-950 text-white select-none">
        {/* Header đếm số to rõ */}
        <header className="p-3 bg-slate-900 border-b border-slate-800">
          <div className="text-xs text-slate-400 font-mono flex justify-between">
            <span>Bao: {currentBagCode}</span>
            <span className="text-emerald-400 font-bold">
              {totalExpected > 0 ? Math.round((matchedCount / totalExpected) * 100) : 0}%
            </span>
          </div>

          <div className="text-5xl font-black font-mono tracking-tight text-emerald-400 my-1">
            {matchedCount}
            <span className="text-2xl text-slate-500 font-normal">/{totalExpected}</span>
          </div>

          <div className="grid grid-cols-4 gap-1.5 mt-2 text-center text-[11px] font-bold">
            <div className="bg-emerald-950/60 border border-emerald-500/30 text-emerald-400 py-1 rounded">
              Khớp: {matchedCount}
            </div>
            <div className="bg-rose-950/60 border border-rose-500/30 text-rose-400 py-1 rounded">
              Thiếu: {missingCount}
            </div>
            <div className="bg-amber-950/60 border border-amber-500/30 text-amber-400 py-1 rounded">
              Lạc bao: {misplacedSet.size}
            </div>
            <div className="bg-slate-800/80 border border-slate-700 text-slate-300 py-1 rounded">
              Dư lạ: {extraSet.size}
            </div>
          </div>
        </header>

        {/* Ô nhập tay test trên máy tính */}
        <form onSubmit={handleManualSubmit} className="p-2.5 bg-slate-900 border-b border-slate-800 flex gap-2">
          <input
            ref={inputRef}
            type="text"
            value={manualInput}
            onChange={(e) => setManualInput(e.target.value)}
            placeholder="Nhập hoặc dán mã SPX / VN... rồi Enter"
            className="flex-1 bg-slate-950 border border-emerald-500/50 rounded-lg px-3 py-2 text-xs font-mono text-white outline-none focus:border-emerald-400"
            autoFocus
          />
          <button
            type="submit"
            className="bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs px-4 py-2 rounded-lg transition-colors whitespace-nowrap"
          >
            Nhập
          </button>
        </form>

        {/* Lịch sử quét gần nhất */}
        <main className="flex-1 overflow-y-auto p-3 space-y-2">
          {recentScans.length === 0 && (
            <div className="text-center text-slate-600 text-xs py-10">Bắn mã kiện hoặc nhập tay để kiểm...</div>
          )}

          {recentScans.map((item, idx) => (
            <div
              key={idx}
              className={`flex items-center justify-between p-2 rounded border text-xs font-mono ${
                item.type === 'VALID'
                  ? 'bg-slate-900 border-slate-800 text-slate-200'
                  : item.type === 'MISPLACED'
                  ? 'bg-amber-950/30 border-amber-700/60 text-amber-300'
                  : item.type === 'EXTRA'
                  ? 'bg-rose-950/30 border-rose-800/50 text-rose-300'
                  : 'bg-slate-900 border-slate-800 text-slate-400'
              }`}
            >
              <div>
                <div className="font-bold">{item.code}</div>
                <div className="text-[10px] text-slate-500">
                  {item.time}
                  {item.type === 'MISPLACED' && (
                    <span className="text-amber-400 ml-1.5 font-semibold">
                      (Thuộc bao: {item.targetBag})
                    </span>
                  )}
                </div>
              </div>
              <span
                className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                  item.type === 'VALID'
                    ? 'bg-emerald-950 text-emerald-400 border border-emerald-800'
                    : item.type === 'MISPLACED'
                    ? 'bg-amber-950 text-amber-300 border border-amber-800'
                    : item.type === 'EXTRA'
                    ? 'bg-rose-950 text-rose-300 border border-rose-800'
                    : 'bg-slate-800 text-slate-400'
                }`}
              >
                {item.type === 'VALID'
                  ? 'Khớp'
                  : item.type === 'MISPLACED'
                  ? 'Lạc bao'
                  : item.type === 'EXTRA'
                  ? 'Dư lạ'
                  : 'Trùng'}
              </span>
            </div>
          ))}
        </main>

        {/* Nút chốt bao */}
        <footer className="p-3 bg-slate-900 border-t border-slate-800">
          <button
            onClick={() => setPhase('REPORT')}
            className="w-full py-3 bg-emerald-600 hover:bg-emerald-500 font-bold text-sm rounded-lg"
          >
            Chốt bao & Xem chi tiết lệch
          </button>
        </footer>
      </div>
    );
  }

  // ==================== BƯỚC 4: BÁO CÁO ĐỐI SOÁT ====================
  const missingList: string[] = [];
  expectedSet.forEach((c) => {
    if (!scannedSet.has(c)) missingList.push(c);
  });

  const misplacedList = Array.from(misplacedSet.entries()); // [mã kiện, thuộc bao nào]

  return (
    <div className="min-h-screen bg-slate-950 text-white p-4 max-w-md mx-auto flex flex-col justify-between">
      <div className="space-y-4">
        <h1 className="text-xl font-bold text-emerald-400">Kết quả xả bao: {currentBagCode}</h1>

        <div className="bg-slate-900 border border-slate-800 rounded-lg p-3 grid grid-cols-4 gap-1.5 text-center text-xs">
          <div>
            <div className="text-slate-400">Dự kiến</div>
            <div className="text-base font-bold font-mono">{totalExpected}</div>
          </div>
          <div>
            <div className="text-slate-400">Khớp</div>
            <div className="text-base font-bold font-mono text-emerald-400">{matchedCount}</div>
          </div>
          <div>
            <div className="text-slate-400">Thiếu</div>
            <div className="text-base font-bold font-mono text-rose-400">{missingList.length}</div>
          </div>
          <div>
            <div className="text-slate-400">Lạc / Dư</div>
            <div className="text-base font-bold font-mono text-amber-400">
              {misplacedList.length + extraSet.size}
            </div>
          </div>
        </div>

        {/* Danh sách đơn thiếu */}
        {missingList.length > 0 && (
          <div>
            <div className="flex justify-between items-center mb-1">
              <span className="text-xs text-rose-400 font-bold">Danh sách đơn thiếu ({missingList.length})</span>
              <button
                onClick={() => {
                  navigator.clipboard.writeText(missingList.join('\n'));
                  alert('Đã copy danh sách đơn thiếu');
                }}
                className="text-[11px] underline text-slate-400 hover:text-white"
              >
                Copy
              </button>
            </div>
            <div className="bg-slate-900 border border-slate-800 rounded p-2 max-h-36 overflow-y-auto text-xs font-mono text-slate-300">
              {missingList.map((c) => (
                <div key={c}>{c}</div>
              ))}
            </div>
          </div>
        )}

        {/* Danh sách đơn lạc bao */}
        {misplacedList.length > 0 && (
          <div>
            <span className="text-xs text-amber-400 font-bold block mb-1">
              Kiện quét nhầm từ bao khác ({misplacedList.length})
            </span>
            <div className="bg-slate-900 border border-slate-800 rounded p-2 max-h-32 overflow-y-auto text-xs font-mono text-slate-300 space-y-1">
              {misplacedList.map(([code, targetBag]) => (
                <div key={code} className="flex justify-between">
                  <span>{code}</span>
                  <span className="text-amber-400 font-semibold">Thuộc: {targetBag}</span>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      <button
        onClick={() => setPhase('SCAN_BAG')}
        className="w-full py-3 bg-emerald-600 hover:bg-emerald-500 font-bold text-sm rounded-lg mt-4"
      >
        Tiếp tục quét bao khác
      </button>
    </div>
  );
}