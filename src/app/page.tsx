'use client';

import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';

type Phase = 'IMPORT_MANIFEST' | 'SCAN_BAG' | 'SCAN_ITEMS' | 'REPORT';

interface BagData {
  bagCode: string;
  expectedCodes: string[];
}

export default function DebaggingApp() {
  const [phase, setPhase] = useState<Phase>('IMPORT_MANIFEST');
  
  // Dữ liệu bao và kiện
  const [allPackages, setAllPackages] = useState<string[]>([]);
  const [currentBagCode, setCurrentBagCode] = useState<string>('');
  
  // Sets kiểm soát kiện đã quét
  const expectedSetRef = useRef<Set<string>>(new Set());
  const [scannedSet, setScannedSet] = useState<Set<string>>(new Set());
  const [extraSet, setExtraSet] = useState<Set<string>>(new Set());
  
  // Lịch sử quét gần nhất
  const [recentScans, setRecentScans] = useState<
    Array<{ code: string; type: 'VALID' | 'EXTRA' | 'DUPLICATE'; time: string }>
  >([]);

  // 1. Âm thanh Web Audio API
  const playTone = (type: 'success' | 'warn' | 'error') => {
    try {
      const ctx = new (window.AudioContext || (window as any).webkitAudioContext)();
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
    } catch (_) {}
  };

  // 2. Đọc file Excel xuất từ nút Export List của FMS
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows: any[] = XLSX.utils.sheet_to_json(ws, { defval: '' });

    if (rows.length === 0) return alert('File không có dữ liệu');

    // Tìm cột SPX Tracking Number theo đúng giao diện FMS
    const headers = Object.keys(rows[0]);
    const trackingCol = headers.find((h) =>
      /spx tracking number|tracking number|waybill/i.test(h)
    ) || headers[0];

    const codes = rows
      .map((r) => String(r[trackingCol] || '').trim().toUpperCase())
      .filter((c) => c.startsWith('SPX') || c.length >= 8);

    if (codes.length === 0) return alert('Không tìm thấy mã SPX Tracking Number hợp lệ!');

    setAllPackages(codes);
    setPhase('SCAN_BAG');
  };

  // 3. Xử lý quét mã bao
  const handleBagScanConfirm = (bagCodeInput: string) => {
    const cleanBag = bagCodeInput.trim().toUpperCase();
    if (!cleanBag) return;

    setCurrentBagCode(cleanBag);
    expectedSetRef.current = new Set(allPackages);
    setScannedSet(new Set());
    setExtraSet(new Set());
    setRecentScans([]);
    setPhase('SCAN_ITEMS');
    playTone('success');
  };

  // 4. Xử lý khi bóp cò quét từng kiện nhỏ
  const handleItemScan = (rawCode: string) => {
    const code = rawCode.trim().toUpperCase();
    if (!code || code.length < 5) return;
    const time = new Date().toLocaleTimeString('vi-VN');

    // Quét trùng
    if (scannedSet.has(code) || extraSet.has(code)) {
      playTone('warn');
      setRecentScans((prev) => [{ code, type: 'DUPLICATE', time }, ...prev.slice(0, 5)]);
      return;
    }

    // Kiện nằm trong bao
    if (expectedSetRef.current.has(code)) {
      playTone('success');
      setScannedSet((prev) => new Set(prev).add(code));
      setRecentScans((prev) => [{ code, type: 'VALID', time }, ...prev.slice(0, 5)]);
      return;
    }

    // Kiện lạ ngoài bao
    playTone('error');
    setExtraSet((prev) => new Set(prev).add(code));
    setRecentScans((prev) => [{ code, type: 'EXTRA', time }, ...prev.slice(0, 5)]);
  };

  // Global listener bắt mã quét từ máy PDA
  useEffect(() => {
    let buffer = '';
    let lastTime = Date.now();

    const onKeyDown = (e: KeyboardEvent) => {
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
  }, [phase, scannedSet, extraSet, allPackages]);

  // ==================== BƯỚC 1: NẠP FILE EXCEL TỪ FMS ====================
  if (phase === 'IMPORT_MANIFEST') {
    return (
      <main className="min-h-screen bg-slate-950 text-white p-4 flex flex-col justify-center max-w-md mx-auto">
        <h1 className="text-xl font-bold text-emerald-400 mb-1">Inbound FMS SPX</h1>
        <p className="text-xs text-slate-400 mb-4">Tải file danh sách kiện xuất từ nút "Export List" trên FMS</p>

        <div className="border-2 border-dashed border-slate-700 rounded-xl p-8 text-center bg-slate-900">
          <label className="cursor-pointer block">
            <span className="text-sm font-semibold block mb-2">Chọn file Excel (.xlsx)</span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} className="hidden" />
            <span className="inline-block bg-emerald-600 hover:bg-emerald-500 text-xs font-bold px-4 py-2 rounded">
              Tải file manifest
            </span>
          </label>
        </div>
      </main>
    );
  }

  // ==================== BƯỚC 2: QUÉT MÃ BAO (TO NUMBER) ====================
  if (phase === 'SCAN_BAG') {
    return (
      <main className="min-h-screen bg-slate-950 text-white p-4 flex flex-col justify-center max-w-md mx-auto">
        <div className="text-center">
          <div className="text-emerald-400 text-xs font-bold mb-2">ĐÃ NẠP {allPackages.length} KIỆN</div>
          <h2 className="text-2xl font-black mb-3">Quét mã bao lớn</h2>
          <p className="text-xs text-slate-400 mb-6">
            Bắn mã QR trên seal hoặc mã <b>TO Number</b> (ví dụ: TO20260912...)
          </p>

          <input
            id="bagInput"
            placeholder="Chờ máy quét mã bao..."
            className="w-full bg-slate-900 border border-emerald-500/50 rounded-lg p-3 text-center text-sm font-mono tracking-wider outline-none mb-3"
            onKeyDown={(e) => {
              if (e.key === 'Enter') handleBagScanConfirm((e.target as HTMLInputElement).value);
            }}
          />

          <button
            onClick={() => handleBagScanConfirm('TO-DEFAULT')}
            className="text-xs text-slate-400 underline hover:text-white"
          >
            Bỏ qua bước quét bao, kiểm đếm ngay
          </button>
        </div>
      </main>
    );
  }

  // ==================== BƯỚC 3: QUÉT KIỆN CON (450 / 500) ====================
  const totalExpected = expectedSetRef.current.size;
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
              {Math.round((matchedCount / totalExpected) * 100) || 0}%
            </span>
          </div>

          <div className="text-5xl font-black font-mono tracking-tight text-emerald-400 my-1">
            {matchedCount}
            <span className="text-2xl text-slate-500 font-normal">/{totalExpected}</span>
          </div>

          <div className="grid grid-cols-3 gap-2 mt-2 text-center text-xs font-bold">
            <div className="bg-emerald-950/60 border border-emerald-500/30 text-emerald-400 py-1 rounded">
              Khớp: {matchedCount}
            </div>
            <div className="bg-rose-950/60 border border-rose-500/30 text-rose-400 py-1 rounded">
              Thiếu: {missingCount}
            </div>
            <div className="bg-amber-950/60 border border-amber-500/30 text-amber-400 py-1 rounded">
              Thừa: {extraSet.size}
            </div>
          </div>
        </header>

        {/* Lịch sử quét gần nhất */}
        <main className="flex-1 overflow-y-auto p-3 space-y-2">
          {recentScans.length === 0 && (
            <div className="text-center text-slate-600 text-xs py-10">Bắn mã kiện SPX để kiểm...</div>
          )}

          {recentScans.map((item, idx) => (
            <div
              key={idx}
              className={`flex items-center justify-between p-2 rounded border text-xs font-mono ${
                item.type === 'VALID'
                  ? 'bg-slate-900 border-slate-800 text-slate-200'
                  : item.type === 'EXTRA'
                  ? 'bg-amber-950/30 border-amber-800/50 text-amber-300'
                  : 'bg-slate-900 border-slate-800 text-slate-400'
              }`}
            >
              <div>
                <div className="font-bold">{item.code}</div>
                <div className="text-[10px] text-slate-500">{item.time}</div>
              </div>
              <span
                className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                  item.type === 'VALID'
                    ? 'bg-emerald-950 text-emerald-400 border border-emerald-800'
                    : item.type === 'EXTRA'
                    ? 'bg-amber-950 text-amber-400 border border-amber-800'
                    : 'bg-slate-800 text-slate-400'
                }`}
              >
                {item.type === 'VALID' ? 'Khớp' : item.type === 'EXTRA' ? 'Thừa' : 'Trùng'}
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
            Chốt bao & Xem lệch
          </button>
        </footer>
      </div>
    );
  }

  // ==================== BƯỚC 4: BÁO CÁO ĐỐI SOÁT ====================
  const missingList: string[] = [];
  expectedSetRef.current.forEach((c) => {
    if (!scannedSet.has(c)) missingList.push(c);
  });

  return (
    <div className="min-h-screen bg-slate-950 text-white p-4 max-w-md mx-auto flex flex-col justify-between">
      <div>
        <h1 className="text-xl font-bold text-emerald-400 mb-1">Kết quả xả bao: {currentBagCode}</h1>
        <div className="bg-slate-900 border border-slate-800 rounded-lg p-3 grid grid-cols-3 gap-2 text-center text-xs my-4">
          <div>
            <div className="text-slate-400">Total</div>
            <div className="text-lg font-bold font-mono">{totalExpected}</div>
          </div>
          <div>
            <div className="text-slate-400">Khớp</div>
            <div className="text-lg font-bold font-mono text-emerald-400">{matchedCount}</div>
          </div>
          <div>
            <div className="text-slate-400">Thiếu</div>
            <div className="text-lg font-bold font-mono text-rose-400">{missingList.length}</div>
          </div>
        </div>

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
            <div className="bg-slate-900 border border-slate-800 rounded p-2 max-h-48 overflow-y-auto text-xs font-mono text-slate-300">
              {missingList.map((c) => (
                <div key={c}>{c}</div>
              ))}
            </div>
          </div>
        )}
      </div>

      <button
        onClick={() => setPhase('SCAN_BAG')}
        className="w-full py-3 bg-emerald-600 hover:bg-emerald-500 font-bold text-sm rounded-lg mt-4"
      >
        Quét tiếp bao khác
      </button>
    </div>
  );
}