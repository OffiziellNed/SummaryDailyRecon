"use client";

import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { jsPDF } from "jspdf";
import JSZip from "jszip";

export default function RepresentmentPortal() {
  const [activeTab, setActiveTab] = useState("ALTO");
  const [uploadedFiles, setUploadedFiles] = useState([]);
  
  const [totalDataCount, setTotalDataCount] = useState(0);
  const [altoData, setAltoData] = useState([]);
  const [rintisData, setRintisData] = useState([]);
  const [rrnData, setRrnData] = useState([]);
  
  const [isDownloading, setIsDownloading] = useState(false);
  const [downloadProgress, setDownloadProgress] = useState("");

  useEffect(() => {
    processFiles(uploadedFiles);
  }, [uploadedFiles]);

  const handleFileSelect = (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;
    setUploadedFiles(prev => [...prev, ...files]);
    e.target.value = null;
  };

  const removeFile = (indexToRemove) => {
    setUploadedFiles(prev => prev.filter((_, index) => index !== indexToRemove));
  };

  const processFiles = async (fileList) => {
    if (fileList.length === 0) {
      setTotalDataCount(0);
      setAltoData([]);
      setRintisData([]);
      setRrnData([]);
      return;
    }

    let allParsedData = [];

    for (let file of fileList) {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });
      
      workbook.SheetNames.forEach((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        allParsedData = [...allParsedData, ...json];
      });
    }

    setTotalDataCount(allParsedData.length);

    const alto = [];
    const rintis = [];
    const rrn = [];

    allParsedData.forEach((row) => {
      const swtcKey = Object.keys(row).find(key => key.trim().toLowerCase().includes('swtc'));
      const statusKey = Object.keys(row).find(key => key.trim().toLowerCase().includes('status'));
      
      let isRrn = false;
      
      if (statusKey) {
        const statusVal = String(row[statusKey]).trim().toLowerCase();
        if (statusVal.includes('rrn no data')) {
          rrn.push(row);
          isRrn = true;
        }
      }

      if (!isRrn && swtcKey) {
        const val = String(row[swtcKey]).trim().toUpperCase();
        if (val === 'ALT') {
          alto.push(row);
        } else if (val === 'RTS') {
          rintis.push(row);
        }
      }
    });

    setAltoData(alto);
    setRintisData(rintis);
    setRrnData(rrn);
  };

  const getColVal = (row, colName) => {
    let key = Object.keys(row).find(k => k.trim().toLowerCase() === colName.toLowerCase());
    if (!key) {
      key = Object.keys(row).find(k => k.trim().toLowerCase().includes(colName.toLowerCase()));
    }
    return key && row[key] !== undefined && row[key] !== "" ? row[key] : "-";
  };

  const formatRupiah = (value) => {
    if (!value || value === "-") return "-";
    const num = Number(String(value).replace(/[^0-9.-]+/g, ""));
    if (isNaN(num)) return value; 
    return "Rp " + num.toLocaleString("id-ID");
  };

  const cleanSpacing = (text) => {
    if (!text || text === "-") return "-";
    return String(text).replace(/\s+/g, ' ').trim();
  };

  const handleCopyData = () => {
    let dataToCopy = [];
    if (activeTab === "ALTO") dataToCopy = altoData;
    else if (activeTab === "RINTIS") dataToCopy = rintisData;
    else if (activeTab === "RRN") dataToCopy = rrnData;

    if (dataToCopy.length === 0) {
      alert("Tidak ada data untuk di-copy.");
      return;
    }

    const headers = ["No", "SWTC", "Bank Issuer", "Transaction Date", "Reff Nr", "Id Trx", "Reference", "Transaction Amount", "Merchant Name", "Reason", "Parent Name", "Status", "Invoice"];
    
    const rows = dataToCopy.map((row, index) => {
      return [
        index + 1,
        getColVal(row, "swtc"),
        getColVal(row, "Bank Issuer"),
        getColVal(row, "Transaction Date"),
        getColVal(row, "Reff Nr"),
        getColVal(row, "Id Trx"),
        getColVal(row, "Reference"),
        formatRupiah(getColVal(row, "Transaction Amount")), 
        cleanSpacing(getColVal(row, "Merchant Name")),
        getColVal(row, "Reason"),
        getColVal(row, "Parent Name"),
        getColVal(row, "Status"),
        getColVal(row, "Invoice")
      ].join("\t");
    });

    const tsvData = [headers.join("\t"), ...rows].join("\n");

    navigator.clipboard.writeText(tsvData)
      .then(() => alert(`Berhasil meng-copy ${dataToCopy.length} data ${activeTab}!`))
      .catch(() => alert("Gagal meng-copy data."));
  };

  // ==========================================
  // LOGIKA PEMBUATAN PDF DAN ZIP
  // ==========================================
  
  const fetchImageAsDataURL = async (originalUrl) => {
    if (!originalUrl || originalUrl === "-") return null;

    let urlsToTry = [originalUrl];

    if (originalUrl.includes("drive.google.com") || originalUrl.includes("drive.usercontent.google.com")) {
      const matchId = originalUrl.match(/[-\w]{25,}/);
      if (matchId) {
        const id = matchId[0];
        urlsToTry = [
          `https://lh3.googleusercontent.com/d/${id}`, 
          `https://drive.google.com/uc?export=view&id=${id}`
        ];
      }
    }

    const getProxies = (url) => [
      `https://wsrv.nl/?url=${encodeURIComponent(url)}`,
      `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`,
      `https://corsproxy.io/?${encodeURIComponent(url)}`,
      url
    ];

    for (let url of urlsToTry) {
      for (let proxy of getProxies(url)) {
        try {
          const res = await fetch(proxy);
          if (!res.ok) continue;

          const blob = await res.blob();
          if (!blob.type.startsWith('image/')) continue; 

          return await new Promise((resolve) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.readAsDataURL(blob);
          });
        } catch (e) {
          continue;
        }
      }
    }
    return null;
  };

  // Fungsi kompresi gambar via Canvas untuk mengurangi ukuran file PDF (Jadi hitungan KB)
  const compressImage = (base64Str, maxWidth = 800, quality = 0.6) => {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        let width = img.width;
        let height = img.height;

        if (width > maxWidth) {
          height = Math.round((height * maxWidth) / width);
          width = maxWidth;
        }

        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        
        // Background putih untuk menghindari background hitam kalau gambar asli PNG transparan
        ctx.fillStyle = "#FFFFFF";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        
        ctx.drawImage(img, 0, 0, width, height);
        
        // Export paksa ke JPEG agar file lebih kecil (PNG murni bikin PDF bengkak MB-an)
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = () => resolve(base64Str); // Fallback kalau gagal render
      img.src = base64Str;
    });
  };

  const getImageDimensions = (base64) => {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => resolve({ w: img.width, h: img.height });
      img.onerror = () => resolve(null);
      img.src = base64;
    });
  };

  const formatDateName = (dateStr) => {
    if (!dateStr || dateStr === "-") return "UnknownDate";
    const parts = String(dateStr).trim().split(/\s+/);
    if (parts.length >= 2) {
      return `${parts[0]} ${parts[1]}`;
    }
    return String(dateStr).replace(/[\/\\]/g, '-');
  };

  const handleDownloadZip = async () => {
    let dataToProcess = [];
    if (activeTab === "ALTO") dataToProcess = altoData;
    else if (activeTab === "RINTIS") dataToProcess = rintisData;
    else if (activeTab === "RRN") dataToProcess = rrnData;

    if (dataToProcess.length === 0) {
      alert("Tidak ada data untuk di-download.");
      return;
    }

    setIsDownloading(true);
    const mainZip = new JSZip(); 

    for (let i = 0; i < dataToProcess.length; i++) {
      setDownloadProgress(`Memproses PDF ${i + 1} dari ${dataToProcess.length}...`);
      await new Promise(resolve => setTimeout(resolve, 10)); 

      const row = dataToProcess[i];
      const doc = new jsPDF();
      
      const bankIssuer = getColVal(row, "Bank Issuer");
      const trxDateRaw = getColVal(row, "Transaction Date");
      const reffNr = getColVal(row, "Reff Nr");
      const cleanMerchant = cleanSpacing(getColVal(row, "Merchant Name"));
      
      const dateName = formatDateName(trxDateRaw);
      const cleanReff = reffNr !== "-" ? reffNr : `UnknownReff_${i+1}`;
      const baseFileName = `trx ${dateName} ${cleanReff}`;
      
      // ===== STRUKTUR PDF =====
      doc.setFontSize(12);
      doc.setFont("helvetica", "normal");
      doc.text(String(bankIssuer).toUpperCase(), 15, 20);

      let startY = 30;
      doc.text(`- Transaction Date: ${trxDateRaw}`, 15, startY); startY += 8;
      doc.text(`- Reff Nr: ${reffNr}`, 15, startY); startY += 8;
      doc.text(`- Id Trx: ${getColVal(row, "Id Trx")}`, 15, startY); startY += 8;
      doc.text(`- Reference: ${getColVal(row, "Reference")}`, 15, startY); startY += 8;
      doc.text(`- Transaction Amount: ${formatRupiah(getColVal(row, "Transaction Amount"))}`, 15, startY); startY += 8;
      doc.text(`- Merchant Name: ${cleanMerchant}`, 15, startY); startY += 8;
      doc.text(`- Reason: ${getColVal(row, "Reason")}`, 15, startY); startY += 14;

      doc.text("Keterangan barang sudah dikirim oleh merchant ke user/customer, berikut bukti invoice dari merchant:", 15, startY, { maxWidth: 180 });
      startY += 10;

      // ===== PROSES INVOICE =====
      const invoiceUrl = getColVal(row, "Invoice");
      if (invoiceUrl && invoiceUrl.startsWith("http")) {
        try {
          const rawImgBase64 = await fetchImageAsDataURL(invoiceUrl);
          if (rawImgBase64) {
            // Kompresi gambar menjadi ukuran KB yang enteng
            const imgBase64 = await compressImage(rawImgBase64, 800, 0.6);
            
            const dims = await getImageDimensions(imgBase64);
            if (dims) {
              let finalW = dims.w;
              let finalH = dims.h;
              const maxW = 180;
              const maxH = 280 - startY; 
              
              if (finalW > maxW) {
                finalH = (maxW / finalW) * finalH;
                finalW = maxW;
              }
              if (finalH > maxH) {
                finalW = (maxH / finalH) * finalW;
                finalH = maxH;
              }
              
              // Masukkan sebagai format kompresi "JPEG" ke JS-PDF
              doc.addImage(imgBase64, "JPEG", 15, startY, finalW, finalH, undefined, "FAST");
            }
          }
        } catch (e) {
          console.error("Gagal memuat gambar invoice:", e);
        }
      }

      const pdfBlob = doc.output("blob");
      
      const subZip = new JSZip();
      subZip.file(`${baseFileName}.pdf`, pdfBlob);
      const subZipBlob = await subZip.generateAsync({ type: "blob" });

      mainZip.file(`${baseFileName}.zip`, subZipBlob);
    }

    setDownloadProgress("Membungkus Master ZIP...");
    
    // Generate Jam saat didownload
    const now = new Date();
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const timeString = `${hours}.${minutes}`; 
    
    mainZip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } }).then((content) => {
      const link = document.createElement("a");
      link.href = URL.createObjectURL(content);
      link.download = `DATA ${activeTab} ${timeString}.zip`;
      link.click();
      setIsDownloading(false);
      setDownloadProgress("");
    });
  };

  const rrnAltoCount = rrnData.filter(r => getColVal(r, 'swtc').toUpperCase() === 'ALT').length;
  const rrnRintisCount = rrnData.filter(r => getColVal(r, 'swtc').toUpperCase() === 'RTS').length;

  const renderTable = (data) => (
    <div style={{ marginTop: '20px', borderRadius: '8px', border: '1px solid #30363d', backgroundColor: '#0d1117', width: '100%' }}>
      <table style={{ width: '100%', tableLayout: 'fixed', borderCollapse: 'collapse', color: '#c9d1d9', fontSize: '9.5px', textAlign: 'left', wordWrap: 'break-word' }}>
        <thead style={{ backgroundColor: '#21262d', borderBottom: '1px solid #30363d' }}>
          <tr>
            <th style={{ padding: '8px 4px', width: '3%' }}>No</th>
            <th style={{ padding: '8px 4px', width: '4%' }}>SWTC</th>
            <th style={{ padding: '8px 4px', width: '8%' }}>Bank Issuer</th>
            <th style={{ padding: '8px 4px', width: '7%' }}>Trx Date</th>
            <th style={{ padding: '8px 4px', width: '8%' }}>Reff Nr</th>
            <th style={{ padding: '8px 4px', width: '8%' }}>Id Trx</th>
            <th style={{ padding: '8px 4px', width: '8%' }}>Reference</th>
            <th style={{ padding: '8px 4px', width: '8%' }}>Amount</th>
            <th style={{ padding: '8px 4px', width: '10%' }}>Merchant Name</th>
            <th style={{ padding: '8px 4px', width: '10%' }}>Reason</th>
            <th style={{ padding: '8px 4px', width: '10%' }}>Parent Name</th>
            <th style={{ padding: '8px 4px', width: '8%' }}>Status</th>
            <th style={{ padding: '8px 4px', width: '8%' }}>Invoice</th>
          </tr>
        </thead>
        <tbody>
          {data.length > 0 ? (
            data.map((row, index) => (
              <tr key={index} style={{ borderBottom: '1px solid #30363d' }}>
                <td style={{ padding: '6px 4px', color: '#ffffff', fontWeight: 'bold' }}>{index + 1}</td>
                <td style={{ padding: '6px 4px', color: '#3fb950', fontWeight: 'bold' }}>{getColVal(row, "swtc")}</td>
                <td style={{ padding: '6px 4px' }}>{getColVal(row, "Bank Issuer")}</td>
                <td style={{ padding: '6px 4px' }}>{getColVal(row, "Transaction Date")}</td>
                <td style={{ padding: '6px 4px' }}>{getColVal(row, "Reff Nr")}</td>
                <td style={{ padding: '6px 4px' }}>{getColVal(row, "Id Trx")}</td>
                <td style={{ padding: '6px 4px' }}>{getColVal(row, "Reference")}</td>
                <td style={{ padding: '6px 4px' }}>{formatRupiah(getColVal(row, "Transaction Amount"))}</td>
                <td style={{ padding: '6px 4px' }}>{cleanSpacing(getColVal(row, "Merchant Name"))}</td>
                <td style={{ padding: '6px 4px' }}>{getColVal(row, "Reason")}</td>
                <td style={{ padding: '6px 4px' }}>{getColVal(row, "Parent Name")}</td>
                <td style={{ padding: '6px 4px' }}>{getColVal(row, "Status")}</td>
                <td style={{ padding: '6px 4px' }}>{getColVal(row, "Invoice")}</td>
              </tr>
            ))
          ) : (
            <tr>
              <td colSpan="13" style={{ padding: '20px', textAlign: 'center', color: '#8b949e', fontSize: '12px' }}>
                Tidak ada data yang sesuai.
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );

  return (
    <div style={{ width: '100%', maxWidth: '1400px', margin: '0 auto', padding: '20px', fontFamily: 'sans-serif' }}>
      
      <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', marginBottom: '30px' }}>
        <h1 style={{ fontSize: '24px', fontWeight: '800', letterSpacing: '1px', color: '#ffffff', margin: 0 }}>
          📊 REPRESENTMENT
        </h1>
      </div>

      <div style={{ backgroundColor: '#161b22', border: '1px solid #30363d', borderRadius: '16px', padding: '24px', boxShadow: '0 10px 25px rgba(0,0,0,0.5)' }}>
        
        <h2 style={{ fontSize: '16px', fontWeight: '700', marginBottom: '16px', color: '#c9d1d9', borderBottom: '1px solid #30363d', paddingBottom: '8px' }}>
          Upload File Rekapitulasi (.xlsx / .xls)
        </h2>
        
        <div style={{ backgroundColor: '#0d1117', border: '1px dashed #30363d', padding: '20px', borderRadius: '10px', marginBottom: '24px' }}>
          <label style={{ display: 'inline-block', backgroundColor: '#1f6feb', color: '#ffffff', padding: '10px 20px', borderRadius: '6px', fontSize: '13px', fontWeight: '600', cursor: 'pointer', transition: 'background-color 0.2s' }}>
            ➕ Pilih File Excel
            <input
              type="file"
              multiple
              accept=".xlsx, .xls"
              onChange={handleFileSelect}
              style={{ display: 'none' }}
            />
          </label>
          
          {uploadedFiles.length > 0 && (
            <div style={{ marginTop: '16px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
              {uploadedFiles.map((file, idx) => (
                <div key={idx} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#21262d', border: '1px solid #30363d', padding: '10px 14px', borderRadius: '6px' }}>
                  <span style={{ color: '#c9d1d9', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    📄 {file.name}
                  </span>
                  <button 
                    onClick={() => removeFile(idx)}
                    style={{ backgroundColor: 'transparent', border: 'none', color: '#f85149', fontSize: '16px', fontWeight: 'bold', cursor: 'pointer', padding: '0 5px' }}
                    title="Hapus file"
                  >
                    ×
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: '16px', marginBottom: '24px' }}>
          <div style={{ backgroundColor: '#0d1117', border: '1px solid #30363d', borderRadius: '10px', padding: '16px', textAlign: 'center' }}>
            <div style={{ fontSize: '11px', color: '#8b949e', fontWeight: 'bold', letterSpacing: '1px' }}>TOTAL ROW</div>
            <div style={{ fontSize: '28px', color: '#ffffff', fontWeight: '900', marginTop: '4px' }}>{totalDataCount}</div>
          </div>
          <div style={{ backgroundColor: '#0d1117', border: '1px solid #30363d', borderTop: '3px solid #1f6feb', borderRadius: '10px', padding: '16px', textAlign: 'center', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
            <div style={{ fontSize: '11px', color: '#8b949e', fontWeight: 'bold', letterSpacing: '1px' }}>DATA ALTO</div>
            <div style={{ fontSize: '28px', color: '#58a6ff', fontWeight: '900', marginTop: '4px' }}>{altoData.length}</div>
          </div>
          <div style={{ backgroundColor: '#0d1117', border: '1px solid #30363d', borderTop: '3px solid #f85149', borderRadius: '10px', padding: '16px', textAlign: 'center', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
            <div style={{ fontSize: '11px', color: '#8b949e', fontWeight: 'bold', letterSpacing: '1px' }}>DATA RINTIS</div>
            <div style={{ fontSize: '28px', color: '#ff7b72', fontWeight: '900', marginTop: '4px' }}>{rintisData.length}</div>
          </div>
          <div style={{ backgroundColor: '#0d1117', border: '1px solid #30363d', borderTop: '3px solid #d29922', borderRadius: '10px', padding: '16px', textAlign: 'center' }}>
            <div style={{ fontSize: '11px', color: '#8b949e', fontWeight: 'bold', letterSpacing: '1px' }}>RRN NO DATA</div>
            <div style={{ fontSize: '28px', color: '#e3b341', fontWeight: '900', marginTop: '4px', marginBottom: '8px' }}>{rrnData.length}</div>
            <div style={{ fontSize: '11px', color: '#8b949e', fontWeight: '600' }}>
              <span style={{ color: '#58a6ff' }}>ALT: {rrnAltoCount}</span><span style={{ margin: '0 6px' }}>|</span><span style={{ color: '#ff7b72' }}>RTS: {rrnRintisCount}</span>
            </div>
          </div>
        </div>

        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #30363d', paddingBottom: '16px' }}>
          <div style={{ display: 'flex', gap: '10px' }}>
            <button 
              onClick={() => setActiveTab("ALTO")}
              style={{ padding: '8px 24px', backgroundColor: activeTab === "ALTO" ? '#1f6feb' : '#21262d', color: '#fff', border: '1px solid #30363d', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', cursor: 'pointer' }}
            >
              ALTO
            </button>
            <button 
              onClick={() => setActiveTab("RINTIS")}
              style={{ padding: '8px 24px', backgroundColor: activeTab === "RINTIS" ? '#da3633' : '#21262d', color: '#fff', border: '1px solid #30363d', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', cursor: 'pointer' }}
            >
              RINTIS
            </button>
            <button 
              onClick={() => setActiveTab("RRN")}
              style={{ padding: '8px 24px', backgroundColor: activeTab === "RRN" ? '#d29922' : '#21262d', color: '#fff', border: '1px solid #30363d', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', cursor: 'pointer' }}
            >
              RRN NO DATA
            </button>
          </div>
          
          <div style={{ display: 'flex', gap: '10px' }}>
            <button 
              onClick={handleCopyData}
              style={{ padding: '8px 16px', backgroundColor: '#21262d', color: '#c9d1d9', border: '1px solid #30363d', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px' }}
            >
              📋 Copy
            </button>
            <button 
              onClick={handleDownloadZip}
              disabled={isDownloading}
              style={{ padding: '8px 16px', backgroundColor: isDownloading ? '#2ea043' : '#238636', color: '#fff', border: '1px solid #2ea043', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', cursor: isDownloading ? 'not-allowed' : 'pointer', display: 'flex', alignItems: 'center', gap: '6px', opacity: isDownloading ? 0.7 : 1, minWidth: '180px', justifyContent: 'center' }}
            >
              {isDownloading ? `⏳ ${downloadProgress}` : `📦 Download ZIP ${activeTab}`}
            </button>
          </div>
        </div>

        {activeTab === "ALTO" && renderTable(altoData)}
        {activeTab === "RINTIS" && renderTable(rintisData)}
        {activeTab === "RRN" && renderTable(rrnData)}

      </div>
    </div>
  );
}
