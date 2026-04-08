import React, { useState, useEffect } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { UploadCloud, Trash2, CheckCircle, AlertCircle, FileSpreadsheet, Download } from 'lucide-react';
import { FACTORY_DB } from '../resource/FactoryList';
import CONSIGNEE_DB from '../resource/ConsigneeList.json';

//[#region] 欄位的列號

  let rowStart = 1;
  let rowFM = rowStart+2;
  let rowSHIPPER = rowFM+2;
  let rowCONSIGNEE = rowSHIPPER+2;
  let rowNOTIFY = rowCONSIGNEE+3;
  let rowFROM = rowNOTIFY+4;
  let rowMARKS = rowFROM+2;
  let rowORDERNO = rowMARKS+1;
  let rowGW = rowMARKS+2;

  let rowCARGOREADYDATE = rowMARKS+12;
  let rowSHIPPEDBY = rowCARGOREADYDATE+1;
  let rowForwarder=rowSHIPPEDBY +2;





//[#endregion]欄位的列號
const getCellText = (cell) => {
  if (!cell || cell.value === null) return '';
  if (typeof cell.value === 'object') {
    if (cell.value.richText) return cell.value.richText.map(rt => rt.text).join('');
    if (cell.value.result !== undefined) return String(cell.value.result);
    return String(cell.value);
  }
  return String(cell.value);
};

export default function App() {
  const [formData, setFormData] = useState({
    pi: 'TE-694',
    date: new Date().toISOString().split('T')[0],
    fm: '巨瑞/Michelle',
    to: 'K&A/Sandra',
    shipper: 'GROWTH UP LIMITED',
    consignee: '請選擇',
    notify: '',
    notifyTel: '',
    notifyFax: '',
    from: 'SHENZHEN, CHINA',
    toDestination: 'NEW YORK, USA',
    marks: '',
    orderNo: '',
    gw: '',
    cbm: '',
    qty: '',
    ctns: '',
    cargoReadyDate: '2025-12-15',
    shippedBy: '',
    shippingTerm: '',
    shippingTerm2: 'SZ',
    forwarder: '',
    needCO: '',
    documentOwner: '',
    shippingDoc: '',
    factories: []
  });

  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [fileStats, setFileStats] = useState([]);
  const[mergedPLData, setMergedPLData] = useState([]);
  const [totals, setTotals] = useState({ ctns: 0, prs: 0, gw: 0, cbm: 0 });
  const[originalWorkbooks, setOriginalWorkbooks] = useState([]);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [selectedConsignee, setSelectedConsignee] = useState('');
  const [isCustomConsignee, setIsCustomConsignee] = useState(false);
  const [orderNoList, setOrderNoList] = useState([]);
  const [factoryNames, setFactoryNames] = useState([]);
  const [factorySelections, setFactorySelections] = useState({
    fwdPayment: '',
    arrangeTransport: '',
    customsClearance: ''
  });
  const [validationErrors, setValidationErrors] = useState({
    factory: '',
    consignee: ''
  });
  const [generationState, setGenerationState] = useState('idle'); // 'idle' | 'generating' | 'success'
  const [generatedBlob, setGeneratedBlob] = useState(null);
  const [generatedFileName, setGeneratedFileName] = useState('');

  // 檢核表單
  const validateForm = () => {
    const errors = {};
    if (formData.factories.length === 0) {
      errors.factory = '請至少選擇一個工廠';
    }
    if (!formData.consignee || formData.consignee === '請選擇') {
      errors.consignee = '請選擇 Consignee';
    }
    setValidationErrors(errors);
    return Object.keys(errors).length === 0;
  };

  // 當上傳檔案變更時，即時解析預覽資料
  useEffect(() => {
    const processFiles = async () => {
      if (uploadedFiles.length === 0) {
        setMergedPLData([]);
        setTotals({ ctns: 0, prs: 0, gw: 0, cbm: 0 });
        setFileStats([]);
        setOriginalWorkbooks([]);
        return;
      }

      setIsProcessing(true);
      let tempMerged =[];
      let tCTNS = 0, tPRS = 0, tGW = 0, tCBM = 0;
      let parsedWbs = [];
      let stats =[];
      let allOrders = new Set();
      let factoryNameSet = new Set();

      for (const file of uploadedFiles) {
        try {
          const wb = new ExcelJS.Workbook();
          await wb.xlsx.load(await file.arrayBuffer());
          parsedWbs.push({ file, wb });

          let plSheet = wb.worksheets.find(s => 
            s.name.includes('PL') || s.name.includes('SDL') || s.name.includes('SY') || s.name.includes('聖達龍') || s.name.includes('森源')
          ) || wb.worksheets[0];
          
          factoryNameSet.add(file.name.replace('.xlsx', ''));

          let inDataSection = false;
          let fileCTNS = 0, filePRS = 0, fileGW = 0, fileCBM = 0;
          let orderSet = new Set();
          let totalRowFound = null;
          let isAfterTotal = false;

          plSheet.eachRow((row) => {
            const cell1 = getCellText(row.getCell(1)).trim().toUpperCase();
            const cell2 = getCellText(row.getCell(2)).trim().toUpperCase();

            if (!inDataSection) {
              if (cell1 === 'P.O.NO' || cell2 === 'STYLE NO.') {
                inDataSection = true;
              }
              return;
            }

            // 找到 TOTAL 行並記錄，直接從該行讀取數據
            if (cell1.includes('TOTAL') || cell1.includes("TOTAL:")) {
              totalRowFound = row;
              isAfterTotal = true;
              return;
            }

            let rowIsEmpty = true;
            const rowData =[];
            for (let i = 1; i <= 19; i++) {
              let val = row.getCell(i).value;
              if (val && typeof val === 'object' && val.result !== undefined) val = val.result;
              rowData.push(val);
              if (getCellText(row.getCell(i)).trim() !== '') rowIsEmpty = false;
            }

            if (rowIsEmpty) return;

            // 只收集 TOTAL 前的資料
            if (!isAfterTotal && rowData[0]) {
              tempMerged.push(rowData);
              orderSet.add(rowData[0]);
              allOrders.add(rowData[0]);
            }
          });

          // 從 TOTAL row 讀取數據
          if (totalRowFound) {
            fileCTNS = parseFloat(getCellText(totalRowFound.getCell(8))) || 0;
            filePRS = parseFloat(getCellText(totalRowFound.getCell(10))) || 0;
            fileGW = parseFloat(getCellText(totalRowFound.getCell(15))) || 0; // GROSS
            fileCBM = parseFloat(getCellText(totalRowFound.getCell(19))) || 0;
            
            tCTNS += fileCTNS;
            tPRS += filePRS;
            tGW += fileGW;
            tCBM += fileCBM;
          }

          let warning = false;
          if (totalRowFound) {
            const origCTN = parseFloat(getCellText(totalRowFound.getCell(8))) || 0;
            if (origCTN > 0 && Math.abs(origCTN - fileCTNS) > 0.1) warning = true;
          }

          stats.push({ name: file.name, orders: orderSet.size, ctns: fileCTNS, prs: filePRS, gw: fileGW, cbm: fileCBM, warning });
        } catch (err) {
          console.error("解析檔案失敗:", file.name, err);
        }
      }

      setMergedPLData(tempMerged);
      setTotals({ ctns: tCTNS, prs: tPRS, gw: tGW, cbm: tCBM });
      setFileStats(stats);
      setOriginalWorkbooks(parsedWbs);
      setOrderNoList(Array.from(allOrders));
      factoryNameSet.add('凱安');
      setFactoryNames(Array.from(factoryNameSet));
      setFormData(prev => ({ ...prev, orderNo: Array.from(allOrders).join(', ') }));
      setIsProcessing(false);
    };

    processFiles();
  },[uploadedFiles]);

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    const files = Array.from(e.dataTransfer.files).filter(f => f.name.endsWith('.xlsx'));
    setUploadedFiles(prev => [...prev, ...files]);
  };

  const generateExcel = async () => {
    if (!validateForm()) {
      return;
    }
    
    setGenerationState('generating');
    try {
      const newWb = new ExcelJS.Workbook();
    
    // ==========================================
    // 1. 產生 SI Sheet
    // ==========================================
    const siSheet = newWb.addWorksheet('SI', { 
      views:[{ showGridLines: true, zoomScale: 85 }] 
    });
    
    siSheet.pageSetup = {
      orientation: 'portrait',
      paperSize: 9,
      scale: 100,
      fitToWidth: 1,
      fitToHeight: 1,
      horizontalCentered: true,
      verticalCentered: false,
      margins: {
        left: 0,
        right: 0,
        top: 0,
        bottom: 0,
        header: 0,
        footer: 0
      },
      showGridLines: false,
      showRowColHeaders: false
    };
    
    siSheet.columns = [
      { width: 16.11 },   // A
      { width: 10.33 },   // B
      { width: 9 },       // C
      { width: 10.11 },   // D
      { width: 10.33 },   // E
      { width: 9.22 },    // F
      { width: 9 },       // G
      { width: 15.66 },   // H
      { width: 12 },      // I
    ];

    siSheet.getCell(`A${rowStart}`).value = 'SHIPMENT INSTRUCTION';
    siSheet.getCell(`A${rowStart}`).font = { size: 14, bold: true, underline: true };
    siSheet.getCell(`H${rowStart}`).value = 'PI#';
    siSheet.getCell(`I${rowStart}`).value = formData.pi;

    siSheet.getCell(`A${rowFM}`).value = `FM: ${formData.fm}`;
    siSheet.getCell(`C${rowFM}`).value = `TO: ${formData.to}`;
    siSheet.getCell(`H${rowFM}`).value = 'DATE:';
    siSheet.getCell(`I${rowFM}`).value = formData.date;

    siSheet.getCell(`A${rowSHIPPER}`).value = 'SHIPPER :';
    siSheet.getCell(`B${rowSHIPPER}`).value = formData.shipper || '';

    const consLines = formData.consignee.split('\n');
    siSheet.getCell(`A${rowCONSIGNEE}`).value = 'CONSIGNEE :';
    siSheet.getCell(`B${rowCONSIGNEE}`).value = consLines[0] || '';
    siSheet.getCell(`B${rowCONSIGNEE+1}`).value = consLines.slice(1).join('\n');

    const notifyLines = formData.notify.split('\n');
    siSheet.getCell(`A${rowNOTIFY}`).value = 'NOTIFY :';
    siSheet.getCell(`B${rowNOTIFY}`).value = notifyLines[0] || '';
    siSheet.getCell(`B${rowNOTIFY+1}`).value = notifyLines.slice(1).join('\n');

    siSheet.getCell(`A${rowFROM}`).value = `FROM :`;
    siSheet.getCell(`B${rowFROM}`).value = formData.from;
    siSheet.getCell(`D${rowFROM}`).value = `TO :`;
    siSheet.getCell(`E${rowFROM}`).value = formData.toDestination;

    siSheet.getCell(`A${rowMARKS}`).value = 'MARKS :';
    const marksLines = formData.marks.split('\n');
    marksLines.forEach((line, index) => {
      siSheet.getCell(`A${rowMARKS + 1 + index}`).value = line;
    });
    siSheet.getCell(`C${rowMARKS}`).value = `P/I #${formData.pi}`;
    siSheet.getCell(`C${rowORDERNO}`).value = 'PO#';
    orderNoList.forEach((po, i) => {
      const rowOffset = Math.floor(i / 4);
      const colOffset = i % 4;
      const col = String.fromCharCode(68 + colOffset);
      siSheet.getCell(`${col}${rowORDERNO + rowOffset}`).value = po;
    });

    const writeStats = (startRow, statsArray, isTotal = false) => {
      let currentRow = startRow;
      
      for (const stat of statsArray) {
        const factoryName = stat.name.replace('.xlsx', '');
        siSheet.getCell(`I${currentRow}`).value = factoryName;
        currentRow++;
        
        siSheet.getCell(`H${currentRow}`).value = 'G.W.(kgs) :';
        siSheet.getCell(`I${currentRow}`).value = stat.gw;
        siSheet.getCell(`I${currentRow}`).numFmt = '0.000';
        currentRow++;
        
        siSheet.getCell(`H${currentRow}`).value = 'CBM:';
        siSheet.getCell(`I${currentRow}`).value = stat.cbm;
        siSheet.getCell(`I${currentRow}`).numFmt = '0.000';
        currentRow++;
        
        siSheet.getCell(`H${currentRow}`).value = "Q'NTY(prs) :";
        siSheet.getCell(`I${currentRow}`).value = stat.prs;
        siSheet.getCell(`I${currentRow}`).numFmt = '0.000';
        currentRow++;
        
        siSheet.getCell(`H${currentRow}`).value = 'CTNS:';
        siSheet.getCell(`I${currentRow}`).value = stat.ctns;
        siSheet.getCell(`I${currentRow}`).numFmt = '0.000';
        currentRow++;
      }
      
      if (isTotal) {
        siSheet.getCell(`I${currentRow}`).value = '加總';
        currentRow++;
        
        siSheet.getCell(`H${currentRow}`).value = 'G.W.(kgs) :';
        siSheet.getCell(`I${currentRow}`).value = totals.gw;
        siSheet.getCell(`I${currentRow}`).numFmt = '0.000';
        currentRow++;
        
        siSheet.getCell(`H${currentRow}`).value = 'CBM:';
        siSheet.getCell(`I${currentRow}`).value = totals.cbm;
        siSheet.getCell(`I${currentRow}`).numFmt = '0.000';
        currentRow++;
        
        siSheet.getCell(`H${currentRow}`).value = "Q'NTY(prs) :";
        siSheet.getCell(`I${currentRow}`).value = totals.prs;
        siSheet.getCell(`I${currentRow}`).numFmt = '0.000';
        currentRow++;
        
        siSheet.getCell(`H${currentRow}`).value = 'CTNS:';
        siSheet.getCell(`I${currentRow}`).value = totals.ctns;
        siSheet.getCell(`I${currentRow}`).numFmt = '0.000';
        currentRow++;
      }
      
      return currentRow;
    };

    const fileStatsWithSheetName = fileStats.map((stat, idx) => {
      const wb = originalWorkbooks[idx];
      if (wb && wb.wb.worksheets.length > 0) {
        const plSheet = wb.wb.worksheets.find(s => 
          s.name.includes('PL') || s.name.includes('SDL') || s.name.includes('SY') || s.name.includes('聖達龍') || s.name.includes('森源')
        ) || wb.wb.worksheets[0];
        return { ...stat, sheetName: plSheet.name };
      }
      return { ...stat, sheetName: stat.name.replace('.xlsx', '') };
    });

    if (fileStatsWithSheetName.length === 1) {
      writeStats(rowGW, fileStatsWithSheetName, true);
    } else {
      writeStats(rowGW, fileStatsWithSheetName, true);
    }

    siSheet.getCell(`C${rowCARGOREADYDATE-5}`).value = '"FREIGHT COLLECT"';
    siSheet.getCell(`C${rowCARGOREADYDATE-3}`).value = '"WE HEREBY CERTIFY THAT THIS SHIPMENT';
    siSheet.getCell(`C${rowCARGOREADYDATE-2}`).value = 'CONTAINS NO SOLID WOOD PACKING MATERIALS"';


    siSheet.getCell(`A${rowCARGOREADYDATE}`).value = `Cargo Ready Date:`;
    siSheet.getCell(`B${rowCARGOREADYDATE}`).value = formData.cargoReadyDate;
    siSheet.getCell(`E${rowCARGOREADYDATE}`).value = `Shipping term:`;
    siSheet.getCell(`F${rowCARGOREADYDATE}`).value = formData.shippingTerm + ' ' +formData.shippingTerm2;

    // 填入 Forwarder 資訊
    let rowOffset = rowForwarder;
    if (formData.forwarder) {
      siSheet.getCell(`A${rowOffset}`).value = 'Forwarder :';
      const forwarderLines = formData.forwarder.split('\n');
      forwarderLines.forEach((line, index) => {
        siSheet.getCell(`B${rowOffset + index}`).value = line;
      });
      rowOffset += forwarderLines.length + 1;
    }

    // 填入工廠聯絡資訊
    formData.factories.forEach(fKey => {
      const fData = FACTORY_DB[fKey];
      if (fData) {
        siSheet.getCell(`A${rowOffset}`).value = `Factory :`;
        siSheet.getCell(`B${rowOffset}`).value = fData.name;
        siSheet.getCell(`B${rowOffset}`).font = { bold: true };
        
        const lines =[...fData.address.split('\n'), ...fData.contact.split('\n')];
        lines.forEach((line, idx) => {
          siSheet.getCell(`B${rowOffset + 1 + idx}`).value = line;
        });
        rowOffset += lines.length + 2;
      }
    });

    // 填入其他資訊
    siSheet.getCell(`A${rowOffset}`).value = `需否申請 CO :`;
    siSheet.getCell(`B${rowOffset}`).value = formData.needCO;
    rowOffset++;

    siSheet.getCell(`A${rowOffset}`).value = `文件負責方:`;
    siSheet.getCell(`B${rowOffset}`).value = formData.documentOwner;
    rowOffset++;

    siSheet.getCell(`A${rowOffset}`).value = `船運單:`;
    siSheet.getCell(`B${rowOffset}`).value = `正本 BL`;
    siSheet.getCell(`C${rowOffset}`).value = `電放 BL`;
    siSheet.getCell(`D${rowOffset}`).value = `正本 FCR`;
    siSheet.getCell(`E${rowOffset}`).value = `電放 FCR`;
    siSheet.getCell(`F${rowOffset}`).value = `Sea Willbill`;
    rowOffset++;

    const shippingOptions = {
      'BL_original': 'B',
      'BL_telex': 'C',
      'FCR_original': 'D',
      'FCR_telex': 'E',
      'Sea_Willbill': 'F'
    };
    const selectedCol = shippingOptions[formData.shippingDoc];
    if (selectedCol) {
      siSheet.getCell(`${selectedCol}${rowOffset}`).value = `V`;
    }

    // 輸出工廠勾選資料表 (Row 53-56)
    if (factoryNames.length > 0) {
      rowOffset++;
      rowOffset++; // 空行
      
      // Row 1: 表頭 (凱安, 工廠1, 工廠2...)
      siSheet.getCell(`D${rowOffset}`).value = factoryNames[0];
      for (let i = 1; i < factoryNames.length; i++) {
        const col = String.fromCharCode(68 + i); // D=68, E=69...
        siSheet.getCell(`${col}${rowOffset}`).value = factoryNames[i];
      }
      rowOffset++;
      
      // Row 2: FWD費用由哪方付款
      siSheet.getCell(`B${rowOffset}`).value = 'FWD 費用由哪方付款';
      if (factorySelections.fwdPayment) {
        const idx = factoryNames.indexOf(factorySelections.fwdPayment);
        if (idx >= 0) {
          const col = String.fromCharCode(68 + idx);
          siSheet.getCell(`${col}${rowOffset}`).value = 'V';
        }
      }
      rowOffset++;
      
      // Row 3: 負責安排運輸
      siSheet.getCell(`B${rowOffset}`).value = '負責安排運輸';
      if (factorySelections.arrangeTransport) {
        const idx = factoryNames.indexOf(factorySelections.arrangeTransport);
        if (idx >= 0) {
          const col = String.fromCharCode(68 + idx);
          siSheet.getCell(`${col}${rowOffset}`).value = 'V';
        }
      }
      rowOffset++;
      
      // Row 4: 負責報關
      siSheet.getCell(`B${rowOffset}`).value = '負責報關';
      if (factorySelections.customsClearance) {
        const idx = factoryNames.indexOf(factorySelections.customsClearance);
        if (idx >= 0) {
          const col = String.fromCharCode(68 + idx);
          siSheet.getCell(`${col}${rowOffset}`).value = 'V';
        }
      }
    }

    // SI Sheet 設為無框線、無填滿
    siSheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'none' },
          left: { style: 'none' },
          bottom: { style: 'none' },
          right: { style: 'none' }
        };
        cell.fill = { type: 'pattern', pattern: 'none' };
      });
    });

    // ==========================================
    // 2. 產生合併 PL Sheet
    // ==========================================
    const plSheet = newWb.addWorksheet('PL', { views: [{ showGridLines: false }] });
    plSheet.columns =[
      { width: 12 }, { width: 14 }, { width: 8 }, { width: 4 }, { width: 8 },
      { width: 18 }, { width: 10 }, { width: 8 }, { width: 10 }, { width: 8 },
      { width: 12 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 12 },
      { width: 8 }, { width: 8 }, { width: 8 }, { width: 12 }
    ];

    plSheet.getCell('A1').value = `NO.: ${formData.pi}`;
    plSheet.getCell('J1').value = `MAIN MARK : ${formData.marks}`;
    plSheet.getCell('A2').value = `DATE : ${formData.date}`;
    plSheet.getCell('J2').value = `NY USA`;
    plSheet.getCell('A3').value = `PACKING LIST OF ${totals.ctns} CTNS`;
    plSheet.getCell('J3').value = 'PO#';
    orderNoList.forEach((po, i) => {
      const rowOffset = Math.floor(i / 4);
      const colOffset = i % 4;
      const col = String.fromCharCode(75 + colOffset); // K=75
      plSheet.getCell(`${col}${3 + rowOffset}`).value = po;
    });
    plSheet.getCell('A4').value = `FROM : ${formData.from}`;
    plSheet.getCell('F4').value = `TO : ${formData.toDestination}`;
    plSheet.getCell('J4').value = `CASE#`;
    plSheet.getCell('A5').value = `VESSEL :`;
    plSheet.getCell('J5').value = `MADE IN CHINA`;

    // 雙層表頭
    const header1 =['P.O.NO', 'STYLE NO.', 'CTN #', '', '', 'DESCRIPTION', 'SIZE RUN', 'CTN', 'PRS/CTN', 'PRS', 'SUB TOTAL', 'WEIGHT (KGS)', '', '', '', 'DIMENS. (CM)', '', '', 'MEASUR.'];
    const header2 =['', '', '', '', '', '', '', '', '', '', '', 'NET', '', 'GROSS', '', 'L', 'W', 'H', '(CBM)'];
    plSheet.getRow(7).values = header1;
    plSheet.getRow(8).values = header2;

    const merges =['A7:A8', 'B7:B8', 'C7:E8', 'F7:F8', 'G7:G8', 'H7:H8', 'I7:I8', 'J7:J8', 'K7:K8', 'L7:O7', 'P7:R7', 'S7:S8'];
    merges.forEach(m => plSheet.mergeCells(m));

    // 表頭樣式 (藍底白字)
    for(let r=7; r<=8; r++) {
      plSheet.getRow(r).eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F81BD' } };
        cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
      });
    }

    // 寫入明細資料
    let currentRow = 9;
    mergedPLData.forEach((data, index) => {
      const row = plSheet.getRow(currentRow);
      row.values = data;
      const isZebra = index % 2 === 1;
      
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber > 19) return;
        if (isZebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        if (colNumber >= 8) cell.alignment = { horizontal: 'right' };
        if ([12, 13, 14, 15].includes(colNumber)) cell.numFmt = '0.00';
        if (colNumber === 19) cell.numFmt = '0.000000';
      });
      currentRow++;
    });

    // 總計列
    const tRow = plSheet.getRow(currentRow);
    tRow.getCell('A').value = 'TOTAL:';
    tRow.getCell('H').value = totals.ctns;
    tRow.getCell('J').value = totals.prs;
    tRow.getCell('O').value = totals.gw;
    tRow.getCell('S').value = totals.cbm;

    tRow.eachCell((cell, colNum) => {
      if (colNum > 19) return;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFE0' } };
      cell.font = { bold: true };
      cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
      if (colNum >= 8) cell.alignment = { horizontal: 'right' };
      if ([12, 13, 14, 15].includes(colNum)) cell.numFmt = '0.00';
      if (colNum === 19) cell.numFmt = '0.000000';
    });

    // ==========================================
    // 3. 複製原始 Sheets - 只複製 PL 工作表
    // ==========================================
    originalWorkbooks.forEach(({ file, wb }) => {
      if (!file || !wb) return;
      try {
        const prefix = file.name.split('.xlsx')[0];
        wb.eachSheet((s) => {
          try {
            if (s.name === 'PL' || s.name.includes('PL-')) {
              const finalName = `PL-${prefix}`;
              const clonedSheet = newWb.addWorksheet(finalName);
              clonedSheet.model = Object.assign({}, s.model, { name: finalName });
              clonedSheet.properties.tabColor = { argb: 'FF808080' };
            }
          } catch (sheetErr) {
            console.warn('複製 sheet 失敗:', s.name, sheetErr);
          }
        });
      } catch (err) {
        console.warn('處理原始檔案失敗:', file.name, err);
      }
    });

    // 下載檔案
    const buffer = await newWb.xlsx.writeBuffer();
    const factorySuffix = formData.factories.map(f => FACTORY_DB[f].shortName).join('_');
    const outName = `PI_${formData.pi}${factorySuffix ? '_' + factorySuffix : ''}.xlsx`;
    setGeneratedBlob(new Blob([buffer]));
    setGeneratedFileName(outName);
    setGenerationState('success');
    } catch (err) {
      console.error('產生檔案失敗:', err);
      alert('產生檔案失敗，請重試');
      setGenerationState('idle');
    }
  };

  const handleDownload = () => {
    if (generatedBlob && generatedFileName) {
      saveAs(generatedBlob, generatedFileName);
      setGenerationState('idle');
      setGeneratedBlob(null);
      setGeneratedFileName('');
    }
  };

  const isFormValid = formData.pi && formData.date && uploadedFiles.length > 0;

  return (
    <div className="min-h-screen bg-gray-50 p-4 md:p-8">
      <div className="max-w-[1600px] mx-auto space-y-6">
        
        {/* 區塊1: 出貨文件自動化系統 - 懸空至頂 */}
        <div className="sticky top-0 z-50 bg-gradient-to-r from-blue-600 to-blue-700 p-4 md:p-6 rounded-2xl shadow-lg flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="text-center md:text-left">
            <h1 className="text-3xl md:text-4xl font-bold text-white">文件自動化系統 v1.0</h1>
            <p className="text-blue-100 mt-2 text-lg">前端解析與合併 PI Excel</p>
          </div>
          <button 
            onClick={generateExcel}
            disabled={!isFormValid || isProcessing}
            className={`px-8 py-4 rounded-xl font-bold flex items-center gap-3 text-lg transition-all shadow-lg ${
              isFormValid && !isProcessing ? 'bg-white text-blue-600 hover:bg-blue-50' : 'bg-gray-300 text-gray-500 cursor-not-allowed'
            }`}
          >
            <FileSpreadsheet size={24} />
            {isProcessing ? '資料解析中...' : '產生 SI 檔案'}
          </button>
        </div>

        {/* 區塊2: 拖曳上傳區 - 改為橫向排列 */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div 
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
            className={`bg-white p-6 rounded-xl border-2 border-dashed flex flex-col items-center justify-center text-center transition-all min-h-[200px] ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300 hover:border-blue-400'}`}
          >
            <UploadCloud className="text-gray-400 mb-3" size={48} />
            <p className="text-gray-700 font-semibold">拖曳 PL (Excel) 檔案至此</p>
            <p className="text-gray-400 text-sm mt-1 mb-4">或點擊下方按鈕選擇檔案 (可多選)</p>
            <label className="bg-white border shadow-sm px-4 py-2 rounded cursor-pointer hover:bg-gray-50 text-sm font-semibold">
              瀏覽檔案
              <input type="file" multiple accept=".xlsx" className="hidden" onChange={(e) => setUploadedFiles(prev => [...prev, ...Array.from(e.target.files)])} />
            </label>
          </div>

          {/* 右側：上傳檔案預覽區塊 */}
          {uploadedFiles.length > 0 ? (
            <div className="lg:col-span-2 bg-white p-6 rounded-xl shadow-sm border border-gray-100">
              <div className="flex justify-between items-center mb-4">
                <h2 className="text-lg font-bold">已上傳檔案預覽</h2>
                <button onClick={() => setUploadedFiles([])} className="text-sm text-red-500 hover:text-red-700 flex items-center gap-1"><Trash2 size={14}/> 清空</button>
              </div>
              
              <div className="space-y-3 mb-6">
                {fileStats.map((stat, i) => (
                  <div key={i} className="flex justify-between items-center p-3 bg-gray-50 rounded border">
                    <div className="flex items-center gap-2">
                      <FileSpreadsheet size={16} className="text-green-600" />
                      <span className="font-medium text-sm text-gray-700 truncate max-w-[150px]">{stat.name}</span>
                    </div>
                    <div className="flex gap-4 text-sm text-gray-600">
                      <span>訂單: <b>{stat.orders}</b></span>
                      <span>CTNS: <b>{stat.ctns}</b></span>
                      <span>PRS: <b>{stat.prs}</b></span>
                    </div>
                    {stat.warning && <AlertCircle size={16} className="text-orange-500 ml-2" title="TOTAL 列與系統加總不符" />}
                  </div>
                ))}
              </div>

              <div className="bg-blue-50 rounded-lg p-4 border border-blue-100">
                <h3 className="font-bold text-blue-800 mb-3 text-sm">彙總數據計算 (將帶入 SI Sheet)</h3>
                <div className="grid grid-cols-4 gap-4 text-center">
                  <div>
                    <p className="text-xs text-blue-600 font-semibold mb-1">總箱數 (CTNS)</p>
                    <p className="text-xl font-bold text-blue-900">{totals.ctns}</p>
                  </div>
                  <div>
                    <p className="text-xs text-blue-600 font-semibold mb-1">總雙數 (PRS)</p>
                    <p className="text-xl font-bold text-blue-900">{totals.prs}</p>
                  </div>
                  <div>
                    <p className="text-xs text-blue-600 font-semibold mb-1">總毛重 (G.W.)</p>
                    <p className="text-xl font-bold text-blue-900">{totals.gw.toFixed(2)}</p>
                  </div>
                  <div>
                    <p className="text-xs text-blue-600 font-semibold mb-1">總體積 (CBM)</p>
                    <p className="text-xl font-bold text-blue-900">{totals.cbm.toFixed(6)}</p>
                  </div>
                </div>
              </div>
            </div>
          ) : (
            <div className="lg:col-span-2 bg-gray-100 rounded-xl border-2 border-dashed border-gray-200 flex items-center justify-center min-h-[200px]">
              <p className="text-gray-400">上傳檔案後顯示預覽</p>
            </div>
          )}
        </div>

        {/* 區塊3: PI基本資訊 - 全寬顯示 */}
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-5">
          <h2 className="text-lg font-bold border-b pb-2 mb-4">PI 基本資訊</h2>
          
          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4">
            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">PI#</label>
              <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.pi} onChange={e => setFormData({...formData, pi: e.target.value})} />
            </div>
            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">Date</label>
              <input type="date" className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.date} onChange={e => setFormData({...formData, date: e.target.value})} />
            </div>
            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">FM</label>
              <select className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.fm} onChange={e => setFormData({...formData, fm: e.target.value})}>
                <option value="巨瑞/Michelle">巨瑞/Michelle</option>
                <option value="巨瑞/Shirely">巨瑞/Shirely</option>
              </select>
            </div>
            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">TO</label>
              <select className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.to} onChange={e => setFormData({...formData, to: e.target.value})}>
                <option value="K&A/Sandra">K&A/Sandra</option>
                <option value=""></option>
              </select>
            </div>
          </div>

            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">SHIPPER</label>
              <select className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.shipper} onChange={e => setFormData({...formData, shipper: e.target.value})}>
                <option value="GROWTH UP LIMITED">GROWTH UP LIMITED</option>
                <option value="TRISTAR FOOTWEAR TRADE CO.,LTD.">TRISTAR FOOTWEAR TRADE CO.,LTD.</option>
                <option value="GREATE SUCCESS CO.,LTD.">GREATE SUCCESS CO.,LTD.</option>
              </select>
            </div>

            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">Consignee</label>
              <select 
                className={`w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none ${validationErrors.consignee ? 'border-red-500' : ''}`}
                value={isCustomConsignee ? 'custom' : selectedConsignee}
                onChange={(e) => {
                  if (e.target.value === 'custom') {
                    setIsCustomConsignee(true);
                    setSelectedConsignee('');
                  } else {
                    setIsCustomConsignee(false);
                    setSelectedConsignee(e.target.value);
                    const selected = CONSIGNEE_DB[e.target.value];
                    if (selected) {
                      const notifyText = [selected.notify_name, selected.notify_address, selected.notify_tel ? `TEL: ${selected.notify_tel}` : '', selected.notify_fax ? `FAX:${selected.notify_fax}` : ''].filter(Boolean).join('\n');
                      const forwarderText = [selected.forwarder_name, selected.forwarder_deputy ? selected.forwarder_deputy : '', selected.forwarder_tel ? `T: ${selected.forwarder_tel}` : '', selected.forwarder_email ? `E: ${selected.forwarder_email}` : '', selected.forwarder_address ? `A: ${selected.forwarder_address}` : ''].filter(Boolean).join('\n');
                      setFormData({
                        ...formData,
                        consignee: selected.address,
                        notify: notifyText,
                        marks: selected.marks,
                        forwarder: forwarderText
                      });
                    }
                  }
                }}
              >
                <option value="">請選擇</option>
                {Object.entries(CONSIGNEE_DB).map(([key, data]) => (
                  <option key={key} value={key}>{data.name}</option>
                ))}
                <option value="custom">自訂</option>
              </select>
              {validationErrors.consignee && <p className="text-red-500 text-sm mt-1">{validationErrors.consignee}</p>}
            </div>

            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">Notify</label>
              <textarea 
                rows="3" 
                className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" 
                value={formData.notify} 
                onChange={e => setFormData({...formData, notify: e.target.value})}
              />
            </div>

            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1">FROM</label>
                <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.from} onChange={e => setFormData({...formData, from: e.target.value})} />
              </div>
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1">TO</label>
                <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.toDestination} onChange={e => setFormData({...formData, toDestination: e.target.value})} />
              </div>
            </div>

            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">MARKS</label>
              <textarea 
                rows="5" 
                className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" 
                value={formData.marks} 
                onChange={e => setFormData({...formData, marks: e.target.value})}
              />
            </div>

            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">P/I {formData.pi}</label>
              <textarea rows="3" className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none bg-gray-50" value={`P/I #${formData.pi}`} readOnly />
            </div>

            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">PO#</label>
              <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none bg-gray-50" value={orderNoList.join(', ')} readOnly placeholder="上傳 PL 檔案後自動帶入" />
            </div>

            {/* 數量相關欄位 - 排列在一起 */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1">CTNS</label>
                <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.ctns} onChange={e => setFormData({...formData, ctns: e.target.value})} placeholder={totals.ctns.toString()} />
              </div>
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1">Q'NTY(prs)</label>
                <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.qty} onChange={e => setFormData({...formData, qty: e.target.value})} placeholder={totals.prs.toString()} />
              </div>
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1">G.W.(kgs)</label>
                <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.gw} onChange={e => setFormData({...formData, gw: e.target.value})} placeholder={totals.gw.toFixed(2)} />
              </div>
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1">CBM</label>
                <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.cbm} onChange={e => setFormData({...formData, cbm: e.target.value})} placeholder={totals.cbm.toFixed(6)} />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1">Cargo Ready Date</label>
                <input type="date" className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.cargoReadyDate} onChange={e => setFormData({...formData, cargoReadyDate: e.target.value})} />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">Shipping term</label>
                  <select className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.shippingTerm} onChange={e => setFormData({...formData, shippingTerm: e.target.value})}>
                    <option value="FOB">FOB</option>
                    <option value="CIF">CIF</option>
                    <option value="ETA">ETA</option>
                    <option value="ETD">ETD</option>
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">Shipping term2</label>
                  <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.shippingTerm2} onChange={e => setFormData({...formData, shippingTerm2: e.target.value})} />
                </div>
              </div>
            </div>

            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">Forwarder</label>
              <textarea rows="5" className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.forwarder} onChange={e => setFormData({...formData, forwarder: e.target.value})} />
            </div>

            <div className="pt-4 border-t">
              <label className="block text-sm font-semibold text-gray-700 mb-3">工廠資訊</label>
              <div className="grid grid-cols-2 gap-4">
                {Object.entries(FACTORY_DB).map(([key, f]) => {
                  const isSelected = formData.factories.includes(key);
                  return (
                    <label key={key} className={`border p-4 rounded-lg cursor-pointer transition-all flex flex-col items-start relative ${isSelected ? 'border-blue-500 bg-blue-50 shadow-sm' : 'border-gray-200 hover:border-gray-300'}`}>
                      <input type="checkbox" className="hidden" 
                        checked={isSelected}
                        onChange={(e) => {
                          const newFactories = e.target.checked 
                            ? [...formData.factories, key] 
                            : formData.factories.filter(k => k !== key);
                          setFormData({...formData, factories: newFactories});
                        }} 
                      />
                      <span className="font-bold text-gray-800">{f.name}</span>
                      <span className="text-xs text-gray-500 mt-1 line-clamp-2">{f.address.replace('\n', ' ')}</span>
                      {isSelected && <CheckCircle className="absolute top-4 right-4 text-blue-500" size={18} />}
                    </label>
                  );
                })}
              </div>
              {validationErrors.factory && <p className="text-red-500 text-sm mt-2">{validationErrors.factory}</p>}
            </div>

            <div className="pt-4 border-t">
              <h3 className="text-lg font-bold border-b pb-2 mb-4">其他</h3>
              
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">需否申請 CO</label>
                  <select className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.needCO} onChange={e => setFormData({...formData, needCO: e.target.value})}>
                    <option value="">請選擇</option>
                    <option value="不需">不需</option>
                    <option value="需要">需要</option>
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">文件負責方</label>
                  <select className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.documentOwner} onChange={e => setFormData({...formData, documentOwner: e.target.value})}>
                    <option value="">請選擇</option>
                    <option value="巨瑞">巨瑞</option>
                    <option value="其他">其他</option>
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">Shipped By</label>
                  <select className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.shippedBy} onChange={e => setFormData({...formData, shippedBy: e.target.value})}>
                    <option value="1x40'HQ">1x40'HQ</option>
                    <option value="1x40'GP">1x40'GP</option>
                    <option value="1x20'">1x20'</option>
                    <option value="LCL">LCL</option>
                  </select>
                </div>
              </div>

              <div className="mt-4">
                <label className="block text-sm font-semibold text-gray-700 mb-2">船運單</label>
                <div className="flex flex-wrap gap-4">
                  <label className="flex items-center gap-2 cursor-pointer text-gray-700">
                    <input type="radio" name="shippingDoc" value="BL_original" checked={formData.shippingDoc === 'BL_original'} onChange={e => setFormData({...formData, shippingDoc: e.target.value})} className="w-4 h-4 text-blue-600" />
                    <span>正本 BL</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer text-gray-700">
                    <input type="radio" name="shippingDoc" value="BL_telex" checked={formData.shippingDoc === 'BL_telex'} onChange={e => setFormData({...formData, shippingDoc: e.target.value})} className="w-4 h-4 text-blue-600" />
                    <span>電放 BL</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer text-gray-700">
                    <input type="radio" name="shippingDoc" value="FCR_original" checked={formData.shippingDoc === 'FCR_original'} onChange={e => setFormData({...formData, shippingDoc: e.target.value})} className="w-4 h-4 text-blue-600" />
                    <span>正本 FCR</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer text-gray-700">
                    <input type="radio" name="shippingDoc" value="FCR_telex" checked={formData.shippingDoc === 'FCR_telex'} onChange={e => setFormData({...formData, shippingDoc: e.target.value})} className="w-4 h-4 text-blue-600" />
                    <span>電放 FCR</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer text-gray-700">
                    <input type="radio" name="shippingDoc" value="Sea_Willbill" checked={formData.shippingDoc === 'Sea_Willbill'} onChange={e => setFormData({...formData, shippingDoc: e.target.value})} className="w-4 h-4 text-blue-600" />
                    <span>Sea Willbill</span>
                  </label>
                </div>
              </div>
            </div>

            {/* 工廠勾選表格 */}
            {factoryNames.length > 0 && (
              <div className="pt-4 border-t">
                <h3 className="text-lg font-bold border-b pb-2 mb-4">勾選資料</h3>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-sm">
                    <thead>
                      <tr>
                        <th className="border bg-gray-100 p-2 text-left">項目</th>
                        {factoryNames.map(name => (
                          <th key={name} className="border bg-gray-100 p-2 text-center">{name}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td className="border p-2 font-medium">FWD 費用由哪方付款</td>
                        {factoryNames.map(name => (
                          <td key={name} className="border p-2 text-center">
                            <input 
                              type="radio" 
                              name="fwdPayment"
                              checked={factorySelections.fwdPayment === name}
                              onChange={() => setFactorySelections(prev => ({ ...prev, fwdPayment: name }))}
                              className="w-4 h-4 text-blue-600"
                            />
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="border p-2 font-medium">負責安排運輸</td>
                        {factoryNames.map(name => (
                          <td key={name} className="border p-2 text-center">
                            <input 
                              type="radio" 
                              name="arrangeTransport"
                              checked={factorySelections.arrangeTransport === name}
                              onChange={() => setFactorySelections(prev => ({ ...prev, arrangeTransport: name }))}
                              className="w-4 h-4 text-blue-600"
                            />
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td className="border p-2 font-medium">負責報關</td>
                        {factoryNames.map(name => (
                          <td key={name} className="border p-2 text-center">
                            <input 
                              type="radio" 
                              name="customsClearance"
                              checked={factorySelections.customsClearance === name}
                              onChange={() => setFactorySelections(prev => ({ ...prev, customsClearance: name }))}
                              className="w-4 h-4 text-blue-600"
                            />
                          </td>
                        ))}
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            )}
        </div>
      </div>

      {/* Loading 遮罩 */}
      {generationState !== 'idle' && (
        <div className="fixed inset-0 bg-gray-200 bg-opacity-80 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-8 flex flex-col items-center shadow-xl min-w-[250px]">
            {generationState === 'generating' ? (
              <>
                <div className="animate-spin rounded-full h-12 w-12 border-4 border-blue-600 border-t-transparent mb-4"></div>
                <p className="text-lg font-semibold text-gray-800">檔案產製中...</p>
              </>
            ) : (
              <>
                <CheckCircle className="text-green-500 mb-4" size={48} />
                <p className="text-lg font-semibold text-gray-800 mb-4">產製成功</p>
                <button 
                  onClick={handleDownload}
                  className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-lg flex items-center gap-2 font-semibold"
                >
                  <Download size={20} />
                  下載檔案
                </button>
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}