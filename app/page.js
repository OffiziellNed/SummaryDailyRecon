"use client";

import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';

const PROGRAMS = [
  { key: 'TOPUP', label: 'TopUp', patterns: ['topup', 'top up', 'top-up'], color: '#1f6feb' },
  { key: 'TRANSFER_OUT', label: 'Transfer Out', patterns: ['transfer', 'transfer_out', 'transfer out'], color: '#f85149' },
  { key: 'ALL', label: 'All Transaction', patterns: ['all_transaction', 'all transaction', 'alltransaction', 'all'], color: '#a371f7' },
  { key: 'PPOB', label: 'PPOB', patterns: ['ppob'], color: '#d29922' },
  { key: 'QRIS', label: 'QRIS', patterns: ['qris'], color: '#3fb950' },
  { key: 'VOUCHER', label: 'Voucher', patterns: ['voucher'], color: '#ff7b72' },
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
    if (fallbackKey === 'SUMMARY') return 'ALL';
    return fallbackKey;
  };

  const findDateKey = (headers) => {
    const cands = ['date', 'tanggal', 'transaction date', 'report date', 'trx date'];
    for (let c of cands) {
      const found = headers.find(h => String(h).toLowerCase().includes(c));
      if (found) return found;
    }
    return headers.find(h => String(h).toLowerCase().includes('time')) || headers[0];
  };

  const findAmountKey = (headers) => {
    const prio = ['nominal', 'amount', 'total amount', 'payment amount', 'voucher amount', 'transaction amount'];
    for (let c of prio) {
      const found = headers.find(h => String(h).toLowerCase().includes(c));
      if (found) return found;
    }
    return null;
  };

  const findStatusKey = (headers) => headers.find(h => String(h).toLowerCase().includes('status')) || null;

  const parseDateOnly = (val) => {
    if (!val) return 'Unknown';
    try {
      const d = new Date(val);
      if (!isNaN(d.getTime())) {
        return d.toISOString().slice(0,10);
      }
      const s = String(val).slice(0,10);
      return s;
    } catch { return String(val).slice(0,10); }
  };

  const toNumber = (v) => {
    if (v === null || v === undefined || v === '') return 0;
    const num = Number(String(v).replace(/[^0-9.-]/g, ''));
    return isNaN(num) ? 0 : num;
  };

  const handleFileSelect = async (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;
    setIsProcessing(true);
    let newEntries = [];
    for (let file of files) {
      try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: "array" });
        let allParsed = [];
        let allHeadersSet = new Set();
        workbook.SheetNames.forEach((sheetName) => {
          const sheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          json.forEach(r => Object.keys(r).forEach(k=> allHeadersSet.add(k)));
          allParsed = [...allParsed, ...json];
        });
        if (allParsed.length === 0) continue;
        const programKey = detectProgram(file.name, activeTab);
        const headers = Array.from(allHeadersSet);
        if (headers.length === 0 && allParsed[0]) headers.push(...Object.keys(allParsed[0]));
        newEntries.push({
          id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
          name: file.name,
          programKey,
          rows: allParsed,
          headers,
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
      if (confirm('Hapus semua data portal? (Summary akan kosong)')) setFileEntries([]);
      return;
    }
    if (confirm(`Hapus semua data ${key}?`)) setFileEntries(prev => prev.filter(f => f.programKey !== key));
  };
  const clearAll = () => {
    if (confirm('Hapus semua data portal?')) setFileEntries([]);
  };

  const programsData = useMemo(() => {
    const grouped = {};
    PROGRAMS.forEach(p => grouped[p.key] = []);
    fileEntries.forEach(entry => {
      if (!grouped[entry.programKey]) grouped[entry.programKey] = [];
      grouped[entry.programKey] = [...grouped[entry.programKey], ...entry.rows];
    });
    return grouped;
  }, [fileEntries]);

  const programsHeaders = useMemo(() => {
    const map = {};
    PROGRAMS.forEach(p => {
      const entries = fileEntries.filter(f => f.programKey === p.key);
      if (entries.length === 0) { map[p.key]=[]; return; }
      const seen = new Set(); const union=[];
      entries.forEach(en=>{
        en.headers.forEach(h=>{ if(!seen.has(h)){seen.add(h); union.push(h);} });
        en.rows.slice(0,3).forEach(r=>{ Object.keys(r).forEach(k=>{ if(!seen.has(k)){seen.add(k); union.push(k);} }); });
      });
      map[p.key]=union;
    });
    return map;
  }, [fileEntries]);

  const summaryStats = useMemo(() => {
    const perProgram = {};
    let grandTotalCount = 0; let grandTotalAmount = 0;
    PROGRAMS.forEach(p=>{
      const rows = programsData[p.key] || [];
      const headers = programsHeaders[p.key] || [];
      const amountKey = findAmountKey(headers);
      const statusKey = findStatusKey(headers);
      const amountSum = rows.reduce((s,r)=> s + toNumber(r[amountKey]), 0);
      const success = statusKey ? rows.filter(r=> String(r[statusKey]).toLowerCase().includes('success')).length : 0;
      const cancel = statusKey ? rows.filter(r=> String(r[statusKey]).toLowerCase().includes('cancel')).length : 0;
      perProgram[p.key] = { count: rows.length, amountSum, success, cancel, amountKey, statusKey, headers };
      grandTotalCount += rows.length;
      grandTotalAmount += amountSum;
    });
    const dailyMap = {};
    PROGRAMS.forEach(p=>{
      const rows = programsData[p.key] || [];
      const headers = programsHeaders[p.key] || [];
      const dateKey = findDateKey(headers);
      const amountKey = findAmountKey(headers);
      rows.forEach(r=>{
        const dateOnly = parseDateOnly(r[dateKey]);
        if (!dailyMap[dateOnly]) dailyMap[dateOnly] = {};
        if (!dailyMap[dateOnly][p.key]) dailyMap[dateOnly][p.key] = { count:0, amount:0 };
        dailyMap[dateOnly][p.key].count += 1;
        dailyMap[dateOnly][p.key].amount += toNumber(r[amountKey]);
      });
    });
    const sortedDates = Object.keys(dailyMap).filter(d=>d!=='Unknown').sort();
    if (dailyMap['Unknown']) sortedDates.push('Unknown');
    return { perProgram, grandTotalCount, grandTotalAmount, dailyMap, sortedDates };
  }, [programsData, programsHeaders]);

  const totalDataCount = useMemo(()=> fileEntries.reduce((s,f)=> s+f.rows.length,0), [fileEntries]);

  const handleCopyData = () => {
    if (activeTab === 'SUMMARY') {
      const lines = [];
      lines.push(['Jenis Trx','Jumlah Trx','Amount Trx','SUCCESS','CANCEL'].join('\t'));
      PROGRAMS.forEach(p=>{
        const s = summaryStats.perProgram[p.key];
        lines.push([p.label, s.count, s.amountSum, s.success, s.cancel].join('\t'));
      });
      lines.push(['TOTAL', summaryStats.grandTotalCount, summaryStats.grandTotalAmount, '', ''].join('\t'));
      navigator.clipboard.writeText(lines.join('\n')).then(()=>alert('Summary YTD copied!'));
      return;
    }
    const dataToCopy = programsData[activeTab] || [];
    const headers = programsHeaders[activeTab] || [];
    if (dataToCopy.length===0){ alert('Tidak ada data'); return; }
    const headerRow = ["No", ...headers];
    const rows = dataToCopy.map((row, idx)=>{
      const vals = headers.map(h=>{
        const v=row[h]; if(v===undefined||v===null||v==='') return '-'; return String(v).replace(/\s+/g,' ').trim()||'-';
      });
      return [idx+1,...vals].join('\t');
    });
    navigator.clipboard.writeText([headerRow.join('\t'),...rows].join('\n')).then(()=>alert(`Copy ${dataToCopy.length} rows ${activeTab}`));
  };

  const handleDownloadExcel = () => {
    if (activeTab === 'SUMMARY') {
      handleDownloadSummaryExcel();
      return;
    }
    const dataToExport = programsData[activeTab] || [];
    if (dataToExport.length===0){ alert('Tidak ada data'); return; }
    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeTab.slice(0,31));
    const timeStr = new Date().toISOString().slice(0,10);
    XLSX.writeFile(wb, `${activeTab}_Recon_${timeStr}.xlsx`);
  };

  const handleDownloadSummaryExcel = () => {
    const wb = XLSX.utils.book_new();
    const summaryHeader = ['Tanggal Transaksi'];
    PROGRAMS.forEach(p=>{ summaryHeader.push(`${p.label} - Jumlah Trx`, `${p.label} - Nominal Trx`); });
    summaryHeader.push('TOTAL Jumlah','TOTAL Nominal');
    const summaryRows = [summaryHeader];
    const grandRow = ['TOTAL YTD', ...PROGRAMS.flatMap(p=>[summaryStats.perProgram[p.key].count, summaryStats.perProgram[p.key].amountSum]), summaryStats.grandTotalCount, summaryStats.grandTotalAmount];
    summaryRows.push(grandRow);
    summaryRows.push([]);
    summaryStats.sortedDates.forEach(date=>{
      const row = [date];
      let dayTotalCount=0, dayTotalAmount=0;
      PROGRAMS.forEach(p=>{
        const d = summaryStats.dailyMap[date]?.[p.key] || {count:0, amount:0};
        row.push(d.count, d.amount);
        dayTotalCount+=d.count; dayTotalAmount+=d.amount;
      });
      row.push(dayTotalCount, dayTotalAmount);
      summaryRows.push(row);
    });
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, wsSummary, 'SUMMARY');

    const ytdHeader = ['Jenis Trx','Jumlah Trx','Amount Trx','Fee Qoin','SUCCESS','CANCEL'];
    const ytdRows = [ ['Qoincash Year To Date - Auto from Portal'], ytdHeader ];
    PROGRAMS.forEach(p=>{
      const s = summaryStats.perProgram[p.key];
      ytdRows.push([p.label, s.count, s.amountSum, 0, s.success, s.cancel]);
    });
    ytdRows.push(['TOTAL', summaryStats.grandTotalCount, summaryStats.grandTotalAmount, 0, '', '']);
    const wsYTD = XLSX.utils.aoa_to_sheet(ytdRows);
    XLSX.utils.book_append_sheet(wb, wsYTD, 'YTD');

    PROGRAMS.forEach(p=>{
      const data = programsData[p.key] || [];
      if (data.length>0){
        const ws = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, p.label.slice(0,31));
      }
    });

    const allCombined = [];
    fileEntries.forEach(entry=>{
      entry.rows.forEach(r=>{
        allCombined.push({ _Program: entry.programKey, _File: entry.name, ...r });
      });
    });
    if (allCombined.length>0){
      const wsAll = XLSX.utils.json_to_sheet(allCombined);
      XLSX.utils.book_append_sheet(wb, wsAll, 'ALL_DATA_PORTAL');
    }

    const timeStr = new Date().toISOString().slice(0,10).replace(/-/g,'');
    XLSX.writeFile(wb, `SUMMARY_DAILY_RECON_QOINAPPS_2026_${timeStr}.xlsx`);
  };

  const renderSummaryTab = () => {
    if (totalDataCount===0){
      return (
        <div style={{ marginTop: '20px', borderRadius: '8px', border: '1px dashed #30363d', backgroundColor: '#0d1117', padding: '60px', textAlign: 'center', color: '#8b949e' }}>
          <div style={{ fontSize: '40px', marginBottom: '12px' }}>📊</div>
          <div style={{ fontSize: '14px', fontWeight: 'bold', color: '#c9d1d9' }}>Belum ada data untuk Summary</div>
          <div style={{ fontSize: '12px', marginTop: '6px' }}>Upload file di 6 program (TopUp, Transfer Out, All Transaction, PPOB, QRIS, Voucher). Summary akan auto-generate dari data portal, bukan dari Excel manual.</div>
        </div>
      );
    }
    return (
      <div style={{ marginTop: '20px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: '12px' }}>
          {PROGRAMS.map(p=>{
            const s = summaryStats.perProgram[p.key];
            return (
              <div key={p.key} style={{ backgroundColor: '#0d1117', border: `1px solid #30363d`, borderTop: `3px solid ${p.color}`, borderRadius: '10px', padding: '14px' }}>
                <div style={{ fontSize: '10px', color: '#8b949e', fontWeight: 'bold', letterSpacing: '1px' }}>{p.label.toUpperCase()}</div>
                <div style={{ fontSize: '22px', color: '#fff', fontWeight: '900', marginTop: '4px' }}>{s.count} trx</div>
                <div style={{ fontSize: '11px', color: '#8b949e', marginTop: '2px' }}>Rp {s.amountSum.toLocaleString('id-ID')}</div>
                <div style={{ display: 'flex', gap: '8px', marginTop: '8px', fontSize: '10px' }}>
                  <span style={{ color: '#3fb950' }}>{s.success} OK</span>
                  <span style={{ color: '#f85149' }}>{s.cancel} CANCEL</span>
                </div>
              </div>
            );
          })}
          <div style={{ backgroundColor: '#0d1117', border: '1px solid #238636', borderTop: '3px solid #238636', borderRadius: '10px', padding: '14px' }}>
            <div style={{ fontSize: '10px', color: '#8b949e', fontWeight: 'bold' }}>TOTAL</div>
            <div style={{ fontSize: '22px', color: '#fff', fontWeight: '900', marginTop: '4px' }}>{summaryStats.grandTotalCount} trx</div>
            <div style={{ fontSize: '11px', color: '#3fb950', marginTop: '2px' }}>Rp {summaryStats.grandTotalAmount.toLocaleString('id-ID')}</div>
            <div style={{ fontSize: '10px', color: '#8b949e', marginTop: '8px' }}>{summaryStats.sortedDates.length} hari unik • {fileEntries.length} file</div>
          </div>
        </div>

        <div style={{ backgroundColor: '#0d1117', border: '1px solid #30363d', borderRadius: '8px', overflow: 'hidden' }}>
          <div style={{ padding: '12px 16px', borderBottom: '1px solid #30363d', fontWeight: 'bold', color: '#c9d1d9', fontSize: '13px', display: 'flex', justifyContent: 'space-between' }}>
            <span>YTD - Qoincash Year To Date (auto dari portal)</span>
            <span style={{ fontSize: '10px', color: '#8b949e' }}>Logic = Sheet YTD di Excel lo</span>
          </div>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', color: '#c9d1d9' }}>
              <thead style={{ backgroundColor: '#21262d' }}>
                <tr>
                  <th style={{ padding: '8px', textAlign: 'left' }}>Jenis Trx</th>
                  <th style={{ padding: '8px' }}>Jumlah Trx</th>
                  <th style={{ padding: '8px' }}>Amount Trx</th>
                  <th style={{ padding: '8px' }}>SUCCESS</th>
                  <th style={{ padding: '8px' }}>CANCEL/FAIL</th>
                  <th style={{ padding: '8px' }}>% Success</th>
                </tr>
              </thead>
              <tbody>
                {PROGRAMS.map(p=>{
                  const s = summaryStats.perProgram[p.key];
                  const pct = s.count>0 ? ((s.success/s.count)*100).toFixed(1) : '0';
                  return (
                    <tr key={p.key} style={{ borderTop: '1px solid #21262d' }}>
                      <td style={{ padding: '8px', fontWeight: 'bold' }}><span style={{ display: 'inline-block', width: '8px', height: '8px', borderRadius: '50%', backgroundColor: p.color, marginRight: '6px' }}></span>{p.label}</td>
                      <td style={{ padding: '8px', textAlign: 'center' }}>{s.count}</td>
                      <td style={{ padding: '8px', textAlign: 'right' }}>Rp {s.amountSum.toLocaleString('id-ID')}</td>
                      <td style={{ padding: '8px', textAlign: 'center', color: '#3fb950' }}>{s.success}</td>
                      <td style={{ padding: '8px', textAlign: 'center', color: '#f85149' }}>{s.cancel}</td>
                      <td style={{ padding: '8px', textAlign: 'center' }}>{pct}%</td>
                    </tr>
                  );
                })}
                <tr style={{ borderTop: '2px solid #30363d', backgroundColor: '#21262d', fontWeight: 'bold' }}>
                  <td style={{ padding: '8px' }}>TOTAL</td>
                  <td style={{ padding: '8px', textAlign: 'center' }}>{summaryStats.grandTotalCount}</td>
                  <td style={{ padding: '8px', textAlign: 'right' }}>Rp {summaryStats.grandTotalAmount.toLocaleString('id-ID')}</td>
                  <td colSpan={3} style={{ padding: '8px', textAlign: 'center', color: '#8b949e', fontSize: '11px' }}>Data dari {fileEntries.length} file • {totalDataCount} rows portal</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        <div style={{ backgroundColor: '#0d1117', border: '1px solid #30363d', borderRadius: '8px', overflow: 'hidden' }}>
          <div style={{ padding: '12px 16px', borderBottom: '1px solid #30363d', fontWeight: 'bold', color: '#c9d1d9', fontSize: '13px' }}>
            SUMMARY - Daily Breakdown (format = Sheet SUMMARY Excel asli)
          </div>
          <div style={{ overflowX: 'auto', maxHeight: '400px', overflowY: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '10px', color: '#c9d1d9', whiteSpace: 'nowrap' }}>
              <thead style={{ backgroundColor: '#21262d', position: 'sticky', top: 0, zIndex: 1 }}>
                <tr>
                  <th style={{ padding: '8px', position: 'sticky', left: 0, backgroundColor: '#21262d', minWidth: '100px' }}>Tanggal Transaksi</th>
                  {PROGRAMS.map(p=>(
                    <React.Fragment key={p.key}>
                      <th style={{ padding: '8px', borderLeft: '1px solid #30363d' }}>{p.label} Jml</th>
                      <th style={{ padding: '8px' }}>{p.label} Nom</th>
                    </React.Fragment>
                  ))}
                  <th style={{ padding: '8px', borderLeft: '2px solid #238636', backgroundColor: '#1a2a1a' }}>TOTAL Jml</th>
                  <th style={{ padding: '8px', backgroundColor: '#1a2a1a' }}>TOTAL Nom</th>
                </tr>
              </thead>
              <tbody>
                {summaryStats.sortedDates.map(date=>{
                  let dayCount=0, dayAmount=0;
                  PROGRAMS.forEach(p=>{ const d=summaryStats.dailyMap[date]?.[p.key]; if(d){dayCount+=d.count; dayAmount+=d.amount;}});
                  return (
                    <tr key={date} style={{ borderTop: '1px solid #21262d' }}>
                      <td style={{ padding: '6px 8px', position: 'sticky', left: 0, backgroundColor: '#0d1117', fontWeight: 'bold' }}>{date}</td>
                      {PROGRAMS.map(p=>{
                        const d = summaryStats.dailyMap[date]?.[p.key] || {count:0, amount:0};
                        return (
                          <React.Fragment key={p.key}>
                            <td style={{ padding: '6px 8px', textAlign: 'center', borderLeft: '1px solid #21262d' }}>{d.count||0}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'right' }}>{d.amount? d.amount.toLocaleString('id-ID'): '0'}</td>
                          </React.Fragment>
                        );
                      })}
                      <td style={{ padding: '6px 8px', textAlign: 'center', fontWeight: 'bold', borderLeft: '2px solid #238636' }}>{dayCount}</td>
                      <td style={{ padding: '6px 8px', textAlign: 'right', fontWeight: 'bold' }}>{dayAmount.toLocaleString('id-ID')}</td>
                    </tr>
                  );
                })}
                {summaryStats.sortedDates.length===0 && (
                  <tr><td colSpan={20} style={{ padding: '20px', textAlign: 'center', color: '#8b949e' }}>Tidak ada tanggal terdeteksi</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  };

  const renderTable = (programKey) => {
    const data = programsData[programKey] || [];
    const headers = programsHeaders[programKey] || [];
    if (data.length===0){
      return (
        <div style={{ marginTop: '20px', borderRadius: '8px', border: '1px dashed #30363d', backgroundColor: '#0d1117', padding: '40px', textAlign: 'center', color: '#8b949e' }}>
          <div style={{ fontSize: '28px', marginBottom: '10px' }}>📭</div>
          <div>Belum ada data {programKey}. Upload file {PROGRAMS.find(p=>p.key===programKey)?.label}</div>
        </div>
      );
    }
    return (
      <div style={{ marginTop: '20px', borderRadius: '8px', border: '1px solid #30363d', backgroundColor: '#0d1117', width: '100%', overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', color: '#c9d1d9', fontSize: '9.5px', textAlign: 'left', whiteSpace: 'nowrap' }}>
          <thead style={{ backgroundColor: '#21262d', borderBottom: '1px solid #30363d', position: 'sticky', top: 0 }}>
            <tr>
              <th style={{ padding: '10px 8px', minWidth: '40px', position: 'sticky', left: 0, backgroundColor: '#21262d', zIndex: 2 }}>No</th>
              {headers.map((h,i)=>(<th key={i} style={{ padding: '10px 8px', minWidth: '120px', fontWeight: '700', color: '#e6edf3' }}>{h}</th>))}
            </tr>
          </thead>
          <tbody>
            {data.map((row, rIdx)=>(
              <tr key={rIdx} style={{ borderBottom: '1px solid #21262d', backgroundColor: rIdx%2===0?'#0d1117':'#161b22' }}>
                <td style={{ padding: '7px 8px', color: '#fff', fontWeight: 'bold', position: 'sticky', left: 0, backgroundColor: rIdx%2===0?'#0d1117':'#161b22' }}>{rIdx+1}</td>
                {headers.map((h,cIdx)=>{
                  const val=row[h]; const display = val===''||val===undefined||val===null?'-':String(val).replace(/\s+/g,' ').trim();
                  const isStatus = h.toLowerCase().includes('status');
                  let color='#c9d1d9';
                  if(isStatus){ const s=String(val).toLowerCase(); if(s.includes('success')||s.includes('paid')) color='#3fb950'; else if(s.includes('cancel')||s.includes('fail')) color='#f85149'; }
                  return <td key={cIdx} style={{ padding: '7px 8px', color, maxWidth: '250px', overflow: 'hidden', textOverflow: 'ellipsis' }}>{display}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div style={{ width: '100%', maxWidth: '1400px', margin: '0 auto', padding: '20px', fontFamily: 'sans-serif', backgroundColor: '#010409', minHeight: '100vh' }}>
      <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', marginBottom: '30px', flexDirection: 'column', gap: '6px' }}>
        <h1 style={{ fontSize: '24px', fontWeight: '800', letterSpacing: '1px', color: '#ffffff', margin: 0, textAlign: 'center' }}>
          📊 Summary Daily Recon Qoin apps 2026
        </h1>
        <div style={{ fontSize: '12px', color: '#8b949e' }}>Menu SUMMARY = pertama masuk web • Auto dari data portal • Download = format asli Excel 31 sheet</div>
      </div>

      <div style={{ backgroundColor: '#161b22', border: '1px solid #30363d', borderRadius: '16px', padding: '24px', boxShadow: '0 10px 25px rgba(0,0,0,0.5)' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px', marginBottom: '16px' }}>
          <h2 style={{ fontSize: '16px', fontWeight: '700', color: '#c9d1d9', margin: 0 }}>
            Upload File Recon (.xlsx)
          </h2>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <span style={{ fontSize: '11px', color: '#8b949e' }}>{isProcessing? '⏳ Memproses...' : `${fileEntries.length} file • ${totalDataCount} rows`}</span>
            {fileEntries.length>0 && <button onClick={clearAll} style={{ padding: '4px 10px', backgroundColor: '#21262d', color: '#f85149', border: '1px solid #30363d', borderRadius: '6px', fontSize: '11px', cursor: 'pointer' }}>Clear All</button>}
          </div>
        </div>

        <div style={{ backgroundColor: '#0d1117', border: '1px dashed #30363d', padding: '20px', borderRadius: '10px', marginBottom: '24px' }}>
          <label style={{ display: 'inline-block', backgroundColor: '#1f6feb', color: '#fff', padding: '10px 20px', borderRadius: '6px', fontSize: '13px', fontWeight: '600', cursor: 'pointer' }}>
            ➕ Pilih File Excel
            <input type="file" multiple accept=".xlsx,.xls" onChange={handleFileSelect} style={{ display: 'none' }} />
          </label>
          <span style={{ marginLeft: '12px', fontSize: '11px', color: '#8b949e' }}>
            Auto-detect program. Tab aktif: <b style={{color:'#58a6ff'}}>{ALL_TABS.find(t=>t.key===activeTab)?.label}</b>
          </span>
          {fileEntries.length>0 && (
            <div style={{ marginTop: '16px', display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '180px', overflowY: 'auto' }}>
              {fileEntries.map(entry=>{
                const prog = PROGRAMS.find(p=>p.key===entry.programKey);
                return (
                  <div key={entry.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#21262d', border: '1px solid #30363d', padding: '8px 12px', borderRadius: '6px' }}>
                    <span style={{ color: '#c9d1d9', fontSize: '11px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                      📄 {entry.name}
                      <span style={{ backgroundColor: prog?.color||'#30363d', color: '#fff', padding: '2px 8px', borderRadius: '10px', fontSize: '10px', fontWeight: 'bold' }}>{prog?.label||entry.programKey}</span>
                      <span style={{ color: '#8b949e' }}>• {entry.rows.length} rows</span>
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
              <button key={p.key} onClick={()=>setActiveTab(p.key)}
                style={{
                  padding: p.key==='SUMMARY' ? '9px 22px' : '8px 16px',
                  backgroundColor: activeTab===p.key ? p.color : '#21262d',
                  color: '#fff',
                  border: `1px solid ${activeTab===p.key ? p.color : '#30363d'}`,
                  borderRadius: '6px',
                  fontWeight: 'bold',
                  fontSize: p.key==='SUMMARY' ? '13px' : '12px',
                  cursor: 'pointer',
                  display: 'flex', alignItems: 'center', gap: '6px',
                  boxShadow: p.key==='SUMMARY' && activeTab==='SUMMARY' ? '0 0 0 2px rgba(35,134,54,0.3)' : 'none'
                }}
              >
                {p.key==='SUMMARY' ? '📊 ' : ''}{p.label}
                <span style={{ backgroundColor: 'rgba(0,0,0,0.3)', padding: '1px 6px', borderRadius: '8px', fontSize: '10px' }}>
                  {p.key==='SUMMARY' ? summaryStats.grandTotalCount : (programsData[p.key]||[]).length}
                </span>
              </button>
            ))}
          </div>
          <div style={{ display: 'flex', gap: '10px' }}>
            <button onClick={()=>clearProgram(activeTab)} style={{ padding: '8px 12px', backgroundColor: '#21262d', color: '#f85149', border: '1px solid #30363d', borderRadius: '6px', fontWeight: 'bold', fontSize: '12px', cursor: 'pointer' }}>🗑️ Clear</button>
            <button onClick={handleCopyData} style={{ padding: '8px 16px', backgroundColor: '#21262d', color: '#c9d1d9', border: '1px solid #30363d', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', cursor: 'pointer' }}>📋 Copy</button>
            <button onClick={handleDownloadExcel} style={{ padding: '8px 16px', backgroundColor: '#238636', color: '#fff', border: '1px solid #2ea043', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', cursor: 'pointer', minWidth: '210px', justifyContent: 'center' }}>
              {activeTab==='SUMMARY' ? '📥 Download SUMMARY (YTD+ALL)' : `📥 Download ${activeTab}`}
            </button>
          </div>
        </div>

        {activeTab==='SUMMARY' ? renderSummaryTab() : renderTable(activeTab)}
      </div>
      <div style={{ textAlign: 'center', marginTop: '20px', fontSize: '11px', color: '#484f58' }}>
        Summary Daily Recon Qoin apps 2026 • SUMMARY = first menu • Auto from portal data • Download = format Excel asli
      </div>
    </div>
  );
}
