"use client";

import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';

const PROGRAMS = [
  { key: 'TOPUP', label: 'TopUp', patterns: ['topup', 'top up', 'top-up'], color: '#1f6feb', fileExample: 'TopUp - Report 6.xlsx', cols: ['No','Date & Time','Transaction No','Reference No.','Phone Number','User Name','Nominal','Method','Payment Channel','VA No.','Status'] },
  { key: 'TRANSFER_OUT', label: 'Transfer Out', patterns: ['transfer', 'transfer_out', 'transfer out'], color: '#f85149', fileExample: 'Transfer_Out - Report 3.xlsx', cols: ['Tanggal','No Transaksi','Id Referensi','No Ponsel/No Rek Pengirim','Nama Pengirim','No Ponsel/Rek Penerima','Nama Penerima','Kode Bank','Product','Nominal','Biaya Admin','Total Transaksi','Metode','Catatan','Status'] },
  { key: 'ALL', label: 'All Transaction', patterns: ['all_transaction', 'all transaction', 'alltransaction'], color: '#a371f7', fileExample: 'All_Transaction - Report 3.xlsx', cols: ['No','Date & Time','Transaction No.','Receipt Extern No.','User Name.','Phone No','Currency Type','Product Group','Product Name','Product Type','Transaction Type','Discount','Total Amount','Payment Methode','Note','Status'] },
  { key: 'PPOB', label: 'PPOB', patterns: ['ppob'], color: '#d29922', fileExample: 'PPOB - Report 5.xlsx', cols: ['Tanggal','Nama Customer','No Transaksi','RRN','No Pelanggan','Jenis Transaksi','Jenis Produk','Kode Produk','Nominal/Denom','Harga Rajabiller','Biaya Admin','Harga Jual Qoin','Fee Rajabiller','Metode Pembayaran','Area','Status'] },
  { key: 'QRIS', label: 'QRIS', patterns: ['qris'], color: '#3fb950', fileExample: 'QRIS - Report 5.xlsx', cols: ['No','Date & Time','Transaction No.','Switch','Reference No','Product','RRN','Acquirer','Merchant','Lokasi','Nominal','Tips Percentage','Tips Amount','Potongan','Total','Method','ID Voucher','Status'] },
  { key: 'VOUCHER', label: 'Voucher', patterns: ['voucher'], color: '#ff7b72', fileExample: 'Voucher - Report 5.xlsx', cols: ['No','Date & Time','Transaction No.','Customer.','Customer No','Voucher Id','Voucher Name','Voucher Type','Payment Amount','Payment','Voucher Amount','Voucher Validity','Status','Redeem Status'] },
];

const ALL_TABS = [{ key: 'SUMMARY', label: 'Summary', color: '#238636' }, ...PROGRAMS];

export default function SummaryReconPortal() {
  const [activeTab, setActiveTab] = useState("SUMMARY");
  const [fileEntries, setFileEntries] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);

  const detectProgram = (fileName, fallbackKey) => {
    const lower = fileName.toLowerCase();
    for (let p of PROGRAMS) {
      if (p.patterns.some(pat => lower.includes(pat))) return p.key;
    }
    if (fallbackKey === 'SUMMARY') return null;
    return fallbackKey;
  };

  const findDateKey = (headers) => {
    const cands = ['date & time','date','tanggal'];
    for (let c of cands) {
      const found = headers.find(h => String(h).toLowerCase().includes(c));
      if (found) return found;
    }
    return headers[0];
  };
  const findAmountKey = (headers) => {
    const prio = ['nominal','total amount','payment amount','voucher amount','total'];
    for (let c of prio) {
      const found = headers.find(h => String(h).toLowerCase() === c || String(h).toLowerCase().includes(c));
      if (found) return found;
    }
    return null;
  };
  const findStatusKey = (headers) => headers.find(h => String(h).toLowerCase().includes('status')) || null;
  const parseDateOnly = (val) => {
    if (!val) return 'Unknown';
    try {
      const d = new Date(val);
      if (!isNaN(d.getTime())) return d.toISOString().slice(0,10);
      return String(val).slice(0,10);
    } catch { return String(val).slice(0,10); }
  };
  const toNumber = (v) => {
    if (v === null || v === undefined || v === '') return 0;
    const num = Number(String(v).replace(/[^0-9.-]/g, ''));
    return isNaN(num) ? 0 : num;
  };

  const handleFileSelect = async (e) => {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    setIsProcessing(true);
    let newEntries = [];
    for (let file of files) {
      try {
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: "array" });
        // ambil sheet pertama aja biar kolom persis
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        // ambil header row persis urutan file
        const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        if (aoa.length === 0) continue;
        const originalHeaders = aoa[0].map(h=> String(h).trim()).filter(h=>h!=="");
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (json.length===0) continue;
        // deteksi program
        const detected = detectProgram(file.name, activeTab);
        if (!detected && activeTab==='SUMMARY'){
          // kalau di SUMMARY tapi nama file tidak jelas, coba auto lagi dari header signature
          const headerStr = originalHeaders.join(' ').toLowerCase();
          if (headerStr.includes('phone number') && headerStr.includes('va no')) newEntries.push({id: `${Date.now()}-${Math.random()}`, name: file.name, programKey: 'TOPUP', rows: json, headers: originalHeaders, addedAt: new Date().toISOString()});
          else if (headerStr.includes('no ponsel/no rek pengirim')) newEntries.push({id: `${Date.now()}-${Math.random()}`, name: file.name, programKey: 'TRANSFER_OUT', rows: json, headers: originalHeaders, addedAt: new Date().toISOString()});
          else if (headerStr.includes('product group')) newEntries.push({id: `${Date.now()}-${Math.random()}`, name: file.name, programKey: 'ALL', rows: json, headers: originalHeaders, addedAt: new Date().toISOString()});
          else if (headerStr.includes('jenis transaksi') && headerStr.includes('kode produk')) newEntries.push({id: `${Date.now()}-${Math.random()}`, name: file.name, programKey: 'PPOB', rows: json, headers: originalHeaders, addedAt: new Date().toISOString()});
          else if (headerStr.includes('acquirer') && headerStr.includes('merchant')) newEntries.push({id: `${Date.now()}-${Math.random()}`, name: file.name, programKey: 'QRIS', rows: json, headers: originalHeaders, addedAt: new Date().toISOString()});
          else if (headerStr.includes('voucher id')) newEntries.push({id: `${Date.now()}-${Math.random()}`, name: file.name, programKey: 'VOUCHER', rows: json, headers: originalHeaders, addedAt: new Date().toISOString()});
          else {
            alert(`File ${file.name} tidak terdeteksi programnya. Rename file mengandung TopUp / Transfer_Out / All_Transaction / PPOB / QRIS / Voucher`);
            continue;
          }
          continue;
        }
        const programKey = detected || activeTab;
        newEntries.push({
          id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
          name: file.name,
          programKey,
          rows: json,
          headers: originalHeaders,
          addedAt: new Date().toISOString()
        });
      } catch (err) {
        alert(`Gagal baca ${file.name}: ${err.message}`);
      }
    }
    setFileEntries(prev => [...prev, ...newEntries]);
    setIsProcessing(false);
    e.target.value = null;
  };

  const removeFile = (id) => setFileEntries(prev => prev.filter(f => f.id !== id));
  const clearProgram = (key) => {
    if (key==='SUMMARY'){
      if (confirm('Hapus semua data portal?')) setFileEntries([]);
      return;
    }
    if (confirm(`Hapus semua data ${key}?`)) setFileEntries(prev => prev.filter(f => f.programKey !== key));
  };

  const programsData = useMemo(() => {
    const grouped = {};
    PROGRAMS.forEach(p => grouped[p.key] = []);
    fileEntries.forEach(entry => {
      grouped[entry.programKey] = [...(grouped[entry.programKey]||[]), ...entry.rows];
    });
    return grouped;
  }, [fileEntries]);

  const programsHeaders = useMemo(() => {
    const map = {};
    PROGRAMS.forEach(p => {
      const entries = fileEntries.filter(f => f.programKey === p.key);
      if (entries.length===0){ map[p.key]=p.cols; return; }
      // pakai header dari file pertama yang diupload untuk program itu, biar persis
      // tapi union kalau beda file beda kolom
      const seen = new Set(); const union=[];
      entries.forEach(en=>{
        en.headers.forEach(h=>{
          if (!seen.has(h)){ seen.add(h); union.push(h); }
        });
      });
      map[p.key]=union.length>0?union:p.cols;
    });
    return map;
  }, [fileEntries]);

  const summaryStats = useMemo(() => {
    const perProgram = {};
    let grandTotalCount=0, grandTotalAmount=0;
    PROGRAMS.forEach(p=>{
      const rows = programsData[p.key]||[];
      const headers = programsHeaders[p.key]||[];
      const amountKey = findAmountKey(headers);
      const statusKey = findStatusKey(headers);
      const amountSum = rows.reduce((s,r)=> s+ toNumber(r[amountKey]),0);
      const success = statusKey ? rows.filter(r=> String(r[statusKey]).toLowerCase().includes('success')).length : 0;
      const cancel = statusKey ? rows.filter(r=> String(r[statusKey]).toLowerCase().includes('cancel')).length : 0;
      perProgram[p.key]={count: rows.length, amountSum, success, cancel};
      grandTotalCount+=rows.length;
      grandTotalAmount+=amountSum;
    });
    const dailyMap={};
    PROGRAMS.forEach(p=>{
      const rows = programsData[p.key]||[];
      const headers = programsHeaders[p.key]||[];
      const dateKey = findDateKey(headers);
      const amountKey = findAmountKey(headers);
      rows.forEach(r=>{
        const dateOnly = parseDateOnly(r[dateKey]);
        if (!dailyMap[dateOnly]) dailyMap[dateOnly]={};
        if (!dailyMap[dateOnly][p.key]) dailyMap[dateOnly][p.key]={count:0, amount:0};
        dailyMap[dateOnly][p.key].count+=1;
        dailyMap[dateOnly][p.key].amount+= toNumber(r[amountKey]);
      });
    });
    const sortedDates = Object.keys(dailyMap).filter(d=>d!=='Unknown').sort();
    if (dailyMap['Unknown']) sortedDates.push('Unknown');
    return { perProgram, grandTotalCount, grandTotalAmount, dailyMap, sortedDates };
  }, [programsData, programsHeaders]);

  const totalDataCount = useMemo(()=> fileEntries.reduce((s,f)=> s+f.rows.length,0), [fileEntries]);


  // helper bikin sheet dengan kolom persis + auto width rapi
  const makeSheetWithExactColumns = (rows, headers) => {
    const aoa = [headers];
    rows.forEach(r=>{
      aoa.push(headers.map(h=> r[h] !== undefined ? r[h] : ""));
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    // auto width - lebar disesuaikan dengan yang paling panjang
    const colWidths = headers.map((_, colIdx) => {
      let maxLen = String(headers[colIdx] || "").length;
      aoa.forEach(row => {
        const cell = row[colIdx];
        if (cell !== undefined && cell !== null) {
          const len = String(cell).length;
          if (len > maxLen) maxLen = len;
        }
      });
      // padding + limit biar gak terlalu lebar, tapi tetap keliatan full
      const wch = Math.min(Math.max(maxLen + 2, 10), 60);
      return { wch };
    });
    ws['!cols'] = colWidths;
    return ws;
  };

  const handleDownloadSummaryExcel = () => {
    if (totalDataCount===0){ alert('Belum ada data untuk Summary'); return; }
    const wb = XLSX.utils.book_new();

    // HAPUS SUMMARY & YTD - cuma 6 Sub dengan kolom rapi
    PROGRAMS.forEach(p=>{
      const rows = programsData[p.key]||[];
      if (rows.length===0) return;
      const headers = programsHeaders[p.key]||p.cols;
      const ws = makeSheetWithExactColumns(rows, headers);
      XLSX.utils.book_append_sheet(wb, ws, p.label.slice(0,31));
    });

    const timeStr = new Date().toISOString().slice(0,10).replace(/-/g,'');
    XLSX.writeFile(wb, `SUMMARY_DAILY_RECON_QOINAPPS_2026_${timeStr}.xlsx`);
  };

  const handleDownloadSingle = () => {
    if (activeTab==='SUMMARY'){ handleDownloadSummaryExcel(); return; }
    const rows = programsData[activeTab]||[];
    if (rows.length===0){ alert('Tidak ada data'); return; }
    const headers = programsHeaders[activeTab]||[];
    const ws = makeSheetWithExactColumns(rows, headers);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeTab.slice(0,31));
    XLSX.writeFile(wb, `${activeTab}_Recon_${new Date().toISOString().slice(0,10)}.xlsx`);
  };


  const handleCopyData = () => {
    if (activeTab==='SUMMARY'){
      const lines=[['Jenis Trx','Jumlah Trx','Amount Trx','SUCCESS','CANCEL'].join('\t')];
      PROGRAMS.forEach(p=>{
        const s=summaryStats.perProgram[p.key];
        lines.push([p.label, s.count, s.amountSum, s.success, s.cancel].join('\t'));
      });
      navigator.clipboard.writeText(lines.join('\n')).then(()=>alert('Summary copied'));
      return;
    }
    const rows=programsData[activeTab]||[];
    const headers=programsHeaders[activeTab]||[];
    if (!rows.length){alert('Tidak ada data');return;}
    const tsv=[["No",...headers].join('\t'), ...rows.map((r,i)=>[i+1, ...headers.map(h=> String(r[h]??'-').replace(/\s+/g,' ').trim())].join('\t'))].join('\n');
    navigator.clipboard.writeText(tsv).then(()=>alert(`Copy ${rows.length} rows`));
  };

  const renderSummaryTab = () => {
    if (totalDataCount===0){
      return (
        <div style={{ marginTop: '20px', borderRadius: '8px', border: '1px dashed #30363d', backgroundColor: '#0d1117', padding: '60px', textAlign: 'center', color: '#8b949e' }}>
          <div style={{ fontSize: '40px' }}>📊</div>
          <div style={{ fontWeight: 'bold', color: '#c9d1d9', marginTop: '10px' }}>Belum ada data untuk Summary</div>
          <div style={{ fontSize: '12px', marginTop: '6px' }}>Upload 6 file: TopUp, Transfer_Out, All_Transaction, PPOB, QRIS, Voucher. Nanti Summary auto-create Excel dengan 6 sub sheet + kolom persis file asli.</div>
          <div style={{ marginTop: '14px', fontSize: '11px', color: '#484f58', display: 'flex', gap: '8px', justifyContent: 'center', flexWrap: 'wrap' }}>
            {PROGRAMS.map(p=><span key={p.key} style={{ backgroundColor: '#21262d', border: '1px solid #30363d', padding: '3px 8px', borderRadius: '12px' }}>{p.label}: {p.cols.slice(0,3).join(', ')}...</span>)}
          </div>
        </div>
      );
    }
    return (
      <div style={{ marginTop: '20px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(170px, 1fr))', gap: '12px' }}>
          {PROGRAMS.map(p=>{
            const s=summaryStats.perProgram[p.key];
            const headers = programsHeaders[p.key];
            return (
              <div key={p.key} style={{ backgroundColor: '#0d1117', border: `1px solid #30363d`, borderTop: `3px solid ${p.color}`, borderRadius: '10px', padding: '14px' }}>
                <div style={{ fontSize: '10px', color: '#8b949e', fontWeight: 'bold' }}>{p.label.toUpperCase()}</div>
                <div style={{ fontSize: '20px', color: '#fff', fontWeight: '900', marginTop: '4px' }}>{s.count} trx</div>
                <div style={{ fontSize: '11px', color: '#8b949e' }}>Rp {s.amountSum.toLocaleString('id-ID')}</div>
                <div style={{ fontSize: '9px', color: '#484f58', marginTop: '8px', lineHeight: '1.3' }}>{headers.length} kolom: {headers.slice(0,2).join(', ')}...</div>
                <div style={{ display: 'flex', gap: '8px', marginTop: '6px', fontSize: '10px' }}>
                  <span style={{ color: '#3fb950' }}>{s.success} OK</span>
                  <span style={{ color: '#f85149' }}>{s.cancel} CANCEL</span>
                </div>
              </div>
            );
          })}
          <div style={{ backgroundColor: '#0d1117', border: '1px solid #238636', borderTop: '3px solid #238636', borderRadius: '10px', padding: '14px' }}>
            <div style={{ fontSize: '10px', color: '#8b949e', fontWeight: 'bold' }}>TOTAL PORTAL</div>
            <div style={{ fontSize: '20px', color: '#fff', fontWeight: '900', marginTop: '4px' }}>{summaryStats.grandTotalCount} trx</div>
            <div style={{ fontSize: '11px', color: '#3fb950' }}>Rp {summaryStats.grandTotalAmount.toLocaleString('id-ID')}</div>
            <div style={{ fontSize: '10px', color: '#8b949e', marginTop: '8px' }}>{fileEntries.length} file • {summaryStats.sortedDates.length} hari</div>
          </div>
        </div>

        <div style={{ backgroundColor: '#0d1117', border: '1px solid #30363d', borderRadius: '8px', overflow: 'hidden' }}>
          <div style={{ padding: '12px 16px', borderBottom: '1px solid #30363d', fontWeight: 'bold', color: '#c9d1d9', fontSize: '13px' }}>Preview - Excel yang akan tercreate (6 Sub + SUMMARY + YTD)</div>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11px', color: '#c9d1d9' }}>
              <thead style={{ backgroundColor: '#21262d' }}>
                <tr><th style={{ padding: '8px', textAlign: 'left' }}>Sheet</th><th style={{ padding: '8px' }}>Jumlah Trx</th><th style={{ padding: '8px' }}>Kolom Persis File Asli</th><th style={{ padding: '8px' }}>Contoh Kolom</th></tr>
              </thead>
              <tbody>
                {PROGRAMS.map(p=>{
                  const s=summaryStats.perProgram[p.key];
                  const headers=programsHeaders[p.key];
                  return (
                    <tr key={p.key} style={{ borderTop: '1px solid #21262d' }}>
                      <td style={{ padding: '8px', fontWeight: 'bold' }}><span style={{ display: 'inline-block', width: '8px', height: '8px', borderRadius: '50%', backgroundColor: p.color, marginRight: '6px' }}></span>{p.label}</td>
                      <td style={{ padding: '8px', textAlign: 'center' }}>{s.count}</td>
                      <td style={{ padding: '8px', textAlign: 'center' }}>{headers.length} kolom</td>
                      <td style={{ padding: '8px', fontSize: '10px', color: '#8b949e' }}>{headers.slice(0,4).join(' | ')}{headers.length>4?' ...':''}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  };

  const renderTable = (programKey) => {
    const data = programsData[programKey]||[];
    const headers = programsHeaders[programKey]||[];
    if (!data.length){
      const prog = PROGRAMS.find(p=>p.key===programKey);
      return (
        <div style={{ marginTop: '20px', borderRadius: '8px', border: '1px dashed #30363d', backgroundColor: '#0d1117', padding: '40px', textAlign: 'center', color: '#8b949e' }}>
          <div>Belum ada data {programKey}</div>
          <div style={{ fontSize: '11px', marginTop: '6px' }}>Upload file {prog?.fileExample} • Kolom harus: {prog?.cols.join(', ')}</div>
        </div>
      );
    }
    return (
      <div style={{ marginTop: '20px', borderRadius: '8px', border: '1px solid #30363d', backgroundColor: '#0d1117', overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', color: '#c9d1d9', fontSize: '9.5px', whiteSpace: 'nowrap' }}>
          <thead style={{ backgroundColor: '#21262d', position: 'sticky', top: 0 }}>
            <tr>
              <th style={{ padding: '10px 8px', position: 'sticky', left: 0, backgroundColor: '#21262d', zIndex: 2 }}>No</th>
              {headers.map((h,i)=><th key={i} style={{ padding: '10px 8px', minWidth: '120px' }}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {data.map((row, rIdx)=>(
              <tr key={rIdx} style={{ borderBottom: '1px solid #21262d', backgroundColor: rIdx%2===0?'#0d1117':'#161b22' }}>
                <td style={{ padding: '7px 8px', position: 'sticky', left: 0, backgroundColor: rIdx%2===0?'#0d1117':'#161b22', fontWeight: 'bold' }}>{rIdx+1}</td>
                {headers.map((h,cIdx)=>{
                  const val=row[h]; const disp= val===''||val===undefined||val===null?'-':String(val).trim();
                  const isStatus = h.toLowerCase().includes('status');
                  let color='#c9d1d9';
                  if(isStatus){ const s=String(val).toLowerCase(); if(s.includes('success')||s.includes('paid')) color='#3fb950'; else if(s.includes('cancel')||s.includes('fail')) color='#f85149'; }
                  return <td key={cIdx} style={{ padding: '7px 8px', color }}>{disp}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div style={{ maxWidth: '1400px', margin: '0 auto', padding: '20px', fontFamily: 'sans-serif', backgroundColor: '#010409', minHeight: '100vh' }}>
      <div style={{ textAlign: 'center', marginBottom: '24px' }}>
        <h1 style={{ fontSize: '22px', fontWeight: '800', color: '#fff', margin: 0 }}>📊 Summary Daily Recon Qoin apps 2026</h1>
        <div style={{ fontSize: '12px', color: '#8b949e', marginTop: '6px' }}>SUMMARY = menu pertama • Excel output = 6 Sub (TopUp, Transfer Out, All Transaction, PPOB, QRIS, Voucher) dengan kolom persis file upload</div>
      </div>

      <div style={{ backgroundColor: '#161b22', border: '1px solid #30363d', borderRadius: '16px', padding: '24px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', flexWrap: 'wrap', gap: '12px', marginBottom: '16px' }}>
          <h2 style={{ fontSize: '15px', color: '#c9d1d9', margin: 0 }}>Upload File Recon (.xlsx)</h2>
          <span style={{ fontSize: '11px', color: '#8b949e' }}>{isProcessing?'⏳ Memproses...':`${fileEntries.length} file • ${totalDataCount} rows`}</span>
        </div>

        <div style={{ backgroundColor: '#0d1117', border: '1px dashed #30363d', padding: '20px', borderRadius: '10px', marginBottom: '24px' }}>
          <label style={{ display: 'inline-block', backgroundColor: '#1f6feb', color: '#fff', padding: '10px 20px', borderRadius: '6px', fontSize: '13px', fontWeight: '600', cursor: 'pointer' }}>
            ➕ Pilih File Excel
            <input type="file" multiple accept=".xlsx,.xls" onChange={handleFileSelect} style={{ display: 'none' }} />
          </label>
          <span style={{ marginLeft: '12px', fontSize: '11px', color: '#8b949e' }}>Auto-detect: nama file harus mengandung TopUp / Transfer_Out / All_Transaction / PPOB / QRIS / Voucher biar kolom persis</span>
          {fileEntries.length>0 && (
            <div style={{ marginTop: '16px', display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '180px', overflowY: 'auto' }}>
              {fileEntries.map(entry=>{
                const prog = PROGRAMS.find(p=>p.key===entry.programKey);
                return (
                  <div key={entry.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#21262d', border: '1px solid #30363d', padding: '8px 12px', borderRadius: '6px' }}>
                    <span style={{ color: '#c9d1d9', fontSize: '11px', display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap' }}>
                      📄 {entry.name}
                      <span style={{ backgroundColor: prog?.color, color: '#fff', padding: '2px 8px', borderRadius: '10px', fontSize: '10px', fontWeight: 'bold' }}>{prog?.label}</span>
                      <span>• {entry.rows.length} rows • {entry.headers.length} kolom persis</span>
                      <span style={{ color: '#484f58', fontSize: '10px' }}>{entry.headers.slice(0,3).join(', ')}...</span>
                    </span>
                    <button onClick={()=>removeFile(entry.id)} style={{ background: 'transparent', border: 'none', color: '#f85149', fontWeight: 'bold', cursor: 'pointer' }}>×</button>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #30363d', paddingBottom: '16px', flexWrap: 'wrap', gap: '12px' }}>
          <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
            {ALL_TABS.map(p=>(
              <button key={p.key} onClick={()=>setActiveTab(p.key)} style={{ padding: p.key==='SUMMARY'?'9px 22px':'8px 16px', backgroundColor: activeTab===p.key?p.color:'#21262d', color: '#fff', border: `1px solid ${activeTab===p.key?p.color:'#30363d'}`, borderRadius: '6px', fontWeight: 'bold', fontSize: p.key==='SUMMARY'?'13px':'12px', cursor: 'pointer', display: 'flex', gap: '6px', alignItems: 'center' }}>
                {p.key==='SUMMARY'?'📊 ':''}{p.label} <span style={{ backgroundColor: 'rgba(0,0,0,0.3)', padding: '1px 6px', borderRadius: '8px', fontSize: '10px' }}>{p.key==='SUMMARY'?summaryStats.grandTotalCount:(programsData[p.key]||[]).length}</span>
              </button>
            ))}
          </div>
          <div style={{ display: 'flex', gap: '10px' }}>
            <button onClick={()=>clearProgram(activeTab)} style={{ padding: '8px 12px', backgroundColor: '#21262d', color: '#f85149', border: '1px solid #30363d', borderRadius: '6px', fontWeight: 'bold', fontSize: '12px', cursor: 'pointer' }}>🗑️ Clear</button>
            <button onClick={handleCopyData} style={{ padding: '8px 16px', backgroundColor: '#21262d', color: '#c9d1d9', border: '1px solid #30363d', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', cursor: 'pointer' }}>📋 Copy</button>
            <button onClick={handleDownloadSingle} style={{ padding: '8px 16px', backgroundColor: '#238636', color: '#fff', border: '1px solid #2ea043', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', cursor: 'pointer', minWidth: '240px' }}>
              {activeTab==='SUMMARY'?'📥 Download SUMMARY (6 Sub + kolom persis)':`📥 Download ${activeTab} (kolom persis)`}
            </button>
          </div>
        </div>

        {activeTab==='SUMMARY'?renderSummaryTab():renderTable(activeTab)}
      </div>
    </div>
  );
}
