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

export default function SummaryReconPortal() {
  const [activeTab, setActiveTab] = useState("TOPUP");
  const [fileEntries, setFileEntries] = useState([]); // {id, name, programKey, rows, headers, addedAt}
  const [isProcessing, setIsProcessing] = useState(false);

  const detectProgram = (fileName, fallbackKey) => {
    const lower = fileName.toLowerCase();
    for (let p of PROGRAMS) {
      if (p.patterns.some(pat => lower.includes(pat))) return p.key;
    }
    return fallbackKey; // kalau nama file gak jelas, masuk ke tab aktif
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
        workbook.SheetNames.forEach((sheetName) => {
          const sheet = workbook.Sheets[sheetName];
          // defval "" biar kolom kosong tetap kebaca
          const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          allParsed = [...allParsed, ...json];
        });

        if (allParsed.length === 0) continue;

        const programKey = detectProgram(file.name, activeTab);
        const headers = Object.keys(allParsed[0] || {});

        newEntries.push({
          id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
          name: file.name,
          programKey,
          rows: allParsed,
          headers,
          addedAt: new Date().toISOString()
        });
      } catch (err) {
        console.error("Gagal parse", file.name, err);
        alert(`Gagal baca ${file.name}: ${err.message}`);
      }
    }

    setFileEntries(prev => [...prev, ...newEntries]);
    setIsProcessing(false);
    e.target.value = null;
  };

  const removeFile = (idToRemove) => {
    setFileEntries(prev => prev.filter(f => f.id !== idToRemove));
  };

  const clearProgram = (programKey) => {
    if (confirm(`Hapus semua data ${programKey}?`)) {
      setFileEntries(prev => prev.filter(f => f.programKey !== programKey));
    }
  };

  // Derived data
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
    const headersMap = {};
    PROGRAMS.forEach(p => {
      const entriesForProg = fileEntries.filter(f => f.programKey === p.key);
      if (entriesForProg.length === 0) {
        headersMap[p.key] = [];
        return;
      }
      // union headers preserve order dari file pertama
      const seen = new Set();
      const union = [];
      entriesForProg.forEach(en => {
        en.headers.forEach(h => {
          if (!seen.has(h)) {
            seen.add(h);
            union.push(h);
          }
        });
        // juga cek keys tambahan yang mungkin beda file beda kolom
        en.rows.slice(0,5).forEach(r => {
          Object.keys(r).forEach(k => {
            if (!seen.has(k)) {
              seen.add(k);
              union.push(k);
            }
          });
        });
      });
      headersMap[p.key] = union;
    });
    return headersMap;
  }, [fileEntries]);

  const totalDataCount = useMemo(() => {
    return fileEntries.reduce((sum, f) => sum + f.rows.length, 0);
  }, [fileEntries]);

  const getStatusCount = (rows, keyword) => {
    if (!rows || rows.length === 0) return 0;
    const statusKey = Object.keys(rows[0] || {}).find(k => k.toLowerCase().includes('status'));
    if (!statusKey) return 0;
    return rows.filter(r => String(r[statusKey] || '').toLowerCase().includes(keyword.toLowerCase())).length;
  };

  const handleCopyData = () => {
    const dataToCopy = programsData[activeTab] || [];
    const headers = programsHeaders[activeTab] || [];
    if (dataToCopy.length === 0) {
      alert("Tidak ada data untuk di-copy.");
      return;
    }
    const headerRow = ["No", ...headers];
    const rows = dataToCopy.map((row, index) => {
      const vals = headers.map(h => {
        const v = row[h];
        if (v === undefined || v === null) return "-";
        // bersihin whitespace berlebihan
        return String(v).replace(/\s+/g, ' ').trim() || "-";
      });
      return [index + 1, ...vals].join("\t");
    });
    const tsv = [headerRow.join("\t"), ...rows].join("\n");
    navigator.clipboard.writeText(tsv)
      .then(() => alert(`Berhasil copy ${dataToCopy.length} data ${activeTab}!`))
      .catch(() => alert("Gagal copy"));
  };

  const handleDownloadExcel = () => {
    const dataToExport = programsData[activeTab] || [];
    if (dataToExport.length === 0) {
      alert("Tidak ada data untuk di-download.");
      return;
    }
    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeTab.slice(0,31));
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2,'0')}.${String(now.getMinutes()).padStart(2,'0')}`;
    XLSX.writeFile(wb, `${activeTab}_Recon_${timeStr}.xlsx`);
  };

  const renderTable = (programKey) => {
    const data = programsData[programKey] || [];
    const headers = programsHeaders[programKey] || [];

    if (data.length === 0) {
      return (
        <div style={{ marginTop: '20px', borderRadius: '8px', border: '1px dashed #30363d', backgroundColor: '#0d1117', padding: '40px', textAlign: 'center', color: '#8b949e' }}>
          <div style={{ fontSize: '32px', marginBottom: '10px' }}>📭</div>
          <div>Belum ada data {programKey}. Upload file {PROGRAMS.find(p=>p.key===programKey)?.label} (.xlsx)</div>
          <div style={{ fontSize: '11px', marginTop: '8px' }}>Kolom akan otomatis menyesuaikan isi Excel.</div>
        </div>
      );
    }

    return (
      <div style={{ marginTop: '20px', borderRadius: '8px', border: '1px solid #30363d', backgroundColor: '#0d1117', width: '100%', overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', color: '#c9d1d9', fontSize: '9.5px', textAlign: 'left', whiteSpace: 'nowrap' }}>
          <thead style={{ backgroundColor: '#21262d', borderBottom: '1px solid #30363d', position: 'sticky', top: 0 }}>
            <tr>
              <th style={{ padding: '10px 8px', minWidth: '40px', position: 'sticky', left: 0, backgroundColor: '#21262d', zIndex: 2 }}>No</th>
              {headers.map((h, idx) => (
                <th key={idx} style={{ padding: '10px 8px', minWidth: '120px', fontWeight: '700', color: '#e6edf3' }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.map((row, rowIdx) => (
              <tr key={rowIdx} style={{ borderBottom: '1px solid #21262d', backgroundColor: rowIdx % 2 === 0 ? '#0d1117' : '#161b22' }}>
                <td style={{ padding: '7px 8px', color: '#ffffff', fontWeight: 'bold', position: 'sticky', left: 0, backgroundColor: rowIdx % 2 === 0 ? '#0d1117' : '#161b22' }}>{rowIdx + 1}</td>
                {headers.map((h, colIdx) => {
                  const val = row[h];
                  const display = val === "" || val === undefined || val === null ? "-" : String(val).replace(/\s+/g, ' ').trim();
                  // highlight status
                  const isStatusCol = h.toLowerCase().includes('status');
                  let color = '#c9d1d9';
                  if (isStatusCol) {
                    const s = String(val).toLowerCase();
                    if (s.includes('success') || s.includes('paid')) color = '#3fb950';
                    else if (s.includes('cancel') || s.includes('fail')) color = '#f85149';
                  }
                  return <td key={colIdx} style={{ padding: '7px 8px', color, maxWidth: '250px', overflow: 'hidden', textOverflow: 'ellipsis' }}>{display}</td>;
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
        <div style={{ fontSize: '12px', color: '#8b949e' }}>Auto-detect program dari nama file • Kolom dinamis per program</div>
      </div>

      <div style={{ backgroundColor: '#161b22', border: '1px solid #30363d', borderRadius: '16px', padding: '24px', boxShadow: '0 10px 25px rgba(0,0,0,0.5)' }}>
        
        <h2 style={{ fontSize: '16px', fontWeight: '700', marginBottom: '16px', color: '#c9d1d9', borderBottom: '1px solid #30363d', paddingBottom: '8px', display: 'flex', justifyContent: 'space-between' }}>
          <span>Upload File Recon (.xlsx / .xls)</span>
          <span style={{ fontSize: '11px', fontWeight: '400', color: '#8b949e' }}>{isProcessing ? '⏳ Memproses...' : `${fileEntries.length} file ter-upload`}</span>
        </h2>
        
        <div style={{ backgroundColor: '#0d1117', border: '1px dashed #30363d', padding: '20px', borderRadius: '10px', marginBottom: '24px' }}>
          <label style={{ display: 'inline-block', backgroundColor: '#1f6feb', color: '#ffffff', padding: '10px 20px', borderRadius: '6px', fontSize: '13px', fontWeight: '600', cursor: 'pointer' }}>
            ➕ Pilih File Excel
            <input
              type="file"
              multiple
              accept=".xlsx, .xls"
              onChange={handleFileSelect}
              style={{ display: 'none' }}
            />
          </label>
          <span style={{ marginLeft: '12px', fontSize: '11px', color: '#8b949e' }}>
            Sistem auto-detect: TopUp, Transfer_Out, All_Transaction, PPOB, QRIS, Voucher. Jika nama file tidak jelas, masuk ke tab aktif: <b style={{color:'#58a6ff'}}>{activeTab}</b>
          </span>
          
          {fileEntries.length > 0 && (
            <div style={{ marginTop: '16px', display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '200px', overflowY: 'auto' }}>
              {fileEntries.map((entry) => {
                const prog = PROGRAMS.find(p => p.key === entry.programKey);
                return (
                  <div key={entry.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#21262d', border: '1px solid #30363d', padding: '10px 14px', borderRadius: '6px' }}>
                    <span style={{ color: '#c9d1d9', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                      📄 {entry.name} 
                      <span style={{ backgroundColor: prog?.color || '#30363d', color: '#fff', padding: '2px 8px', borderRadius: '10px', fontSize: '10px', fontWeight: 'bold' }}>{prog?.label || entry.programKey}</span>
                      <span style={{ color: '#8b949e' }}>• {entry.rows.length} rows</span>
                    </span>
                    <button 
                      onClick={() => removeFile(entry.id)}
                      style={{ backgroundColor: 'transparent', border: 'none', color: '#f85149', fontSize: '16px', fontWeight: 'bold', cursor: 'pointer', padding: '0 5px' }}
                      title="Hapus file"
                    >
                      ×
                    </button>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* STATS GRID - Mirip UI lama tapi jadi per program */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: '16px', marginBottom: '24px' }}>
          <div style={{ backgroundColor: '#0d1117', border: '1px solid #30363d', borderRadius: '10px', padding: '16px', textAlign: 'center' }}>
            <div style={{ fontSize: '11px', color: '#8b949e', fontWeight: 'bold', letterSpacing: '1px' }}>TOTAL ROW</div>
            <div style={{ fontSize: '28px', color: '#ffffff', fontWeight: '900', marginTop: '4px' }}>{totalDataCount}</div>
          </div>
          {PROGRAMS.map(prog => {
            const count = (programsData[prog.key] || []).length;
            const success = getStatusCount(programsData[prog.key], 'success');
            return (
              <div key={prog.key} style={{ backgroundColor: '#0d1117', border: '1px solid #30363d', borderTop: `3px solid ${prog.color}`, borderRadius: '10px', padding: '16px', textAlign: 'center' }}>
                <div style={{ fontSize: '11px', color: '#8b949e', fontWeight: 'bold', letterSpacing: '1px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{prog.label.toUpperCase()}</div>
                <div style={{ fontSize: '28px', color: '#ffffff', fontWeight: '900', marginTop: '4px' }}>{count}</div>
                {count > 0 && <div style={{ fontSize: '10px', color: '#3fb950', marginTop: '4px' }}>{success} SUCCESS</div>}
              </div>
            );
          })}
        </div>

        {/* TABS + ACTIONS - UI SAMA PERSIS */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #30363d', paddingBottom: '16px', flexWrap: 'wrap', gap: '12px' }}>
          <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
            {PROGRAMS.map(p => (
              <button 
                key={p.key}
                onClick={() => setActiveTab(p.key)}
                style={{ 
                  padding: '8px 18px', 
                  backgroundColor: activeTab === p.key ? p.color : '#21262d', 
                  color: '#fff', 
                  border: `1px solid ${activeTab === p.key ? p.color : '#30363d'}`, 
                  borderRadius: '6px', 
                  fontWeight: 'bold', 
                  fontSize: '12px', 
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px'
                }}
              >
                {p.label} <span style={{ backgroundColor: 'rgba(0,0,0,0.3)', padding: '1px 6px', borderRadius: '8px', fontSize: '10px' }}>{(programsData[p.key] || []).length}</span>
              </button>
            ))}
          </div>
          
          <div style={{ display: 'flex', gap: '10px' }}>
            <button 
              onClick={() => clearProgram(activeTab)}
              style={{ padding: '8px 12px', backgroundColor: '#21262d', color: '#f85149', border: '1px solid #30363d', borderRadius: '6px', fontWeight: 'bold', fontSize: '12px', cursor: 'pointer' }}
            >
              🗑️ Clear {activeTab}
            </button>
            <button 
              onClick={handleCopyData}
              style={{ padding: '8px 16px', backgroundColor: '#21262d', color: '#c9d1d9', border: '1px solid #30363d', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px' }}
            >
              📋 Copy
            </button>
            <button 
              onClick={handleDownloadExcel}
              style={{ padding: '8px 16px', backgroundColor: '#238636', color: '#fff', border: '1px solid #2ea043', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px', minWidth: '160px', justifyContent: 'center' }}
            >
              📥 Download {activeTab}
            </button>
          </div>
        </div>

        {/* INFO KOLOM DINAMIS */}
        <div style={{ marginTop: '14px', fontSize: '11px', color: '#8b949e', display: 'flex', gap: '16px', flexWrap: 'wrap' }}>
          <span>Active: <b style={{color:'#fff'}}>{PROGRAMS.find(p=>p.key===activeTab)?.label}</b></span>
          <span>Kolom: {(programsHeaders[activeTab] || []).length} kolom • {(programsHeaders[activeTab] || []).slice(0,4).join(', ')}{ (programsHeaders[activeTab]||[]).length>4 ? '...' : ''}</span>
          <span style={{ color: '#3fb950' }}>● Auto-update saat upload baru</span>
        </div>

        {renderTable(activeTab)}

      </div>
      <div style={{ textAlign: 'center', marginTop: '20px', fontSize: '11px', color: '#484f58' }}>
        Summary Daily Recon Qoin apps 2026 • UI preserved from Representment Portal
      </div>
    </div>
  );
}
