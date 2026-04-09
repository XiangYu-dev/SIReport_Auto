/* eslint-disable no-unused-vars */
import React, { useState, useEffect } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { UploadCloud, Trash2, CheckCircle, AlertCircle, FileSpreadsheet, Download } from 'lucide-react';
import { getCellText } from '../Common/utility';

let rowStart = 1;

export default function Invoice() {
  const [formData, setFormData] = useState({
    buyer: 'GLOBE FOOTWEAR CORP.',
    attn: 'Chris',
    piNo: '',
    date: new Date().toISOString().split('T')[0],
    shipDate: '',
    paymentTerms: 'AFTER SHIPMENT BY T/T or BY CHECK.',
    beneficiary: 'GROWTH UP LIMITED',
    accountNo: '',
    bankName: '',
    bankAddress: '',
    swiftCode: ''
  });

  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [fileStats, setFileStats] = useState([]);
  const [orderData, setOrderData] = useState([]);
  const [totals, setTotals] = useState({ prs: 0, amount: 0, factoryAmount: 0 });
  const [originalWorkbooks, setOriginalWorkbooks] = useState([]);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  
  const [validationErrors, setValidationErrors] = useState({
    file: '',
    buyer: ''
  });
  const [generationState, setGenerationState] = useState('idle');
  const [generatedBlob, setGeneratedBlob] = useState(null);
  const [generatedFileName, setGeneratedFileName] = useState('');

  const validateForm = () => {
    const errors = {};
    
    if (uploadedFiles.length === 0) {
      errors.file = '請上傳 OrderReport 檔案';
    }
    
    if (!formData.buyer) {
      errors.buyer = '請輸入買家名稱';
    }
    
    setValidationErrors(errors);
    return Object.keys(errors).length === 0;
  };

  useEffect(() => {
    const processFiles = async () => {
      if (uploadedFiles.length === 0) {
        setOrderData([]);
        setTotals({ prs: 0, amount: 0, factoryAmount: 0 });
        setFileStats([]);
        setOriginalWorkbooks([]);
        return;
      }

      setIsProcessing(true);
      let tempData = [];
      let tPRS = 0, tAmount = 0, tFactoryAmount = 0;
      let parsedWbs = [];
      let stats = [];

      for (const file of uploadedFiles) {
        try {
          const wb = new ExcelJS.Workbook();
          await wb.xlsx.load(await file.arrayBuffer());
          parsedWbs.push({ file, wb });

          const orderSheet = wb.worksheets.find(s => 
            s.name.includes('U.order') || s.name.includes('order') || s.name.includes('Order')
          ) || wb.worksheets[0];

          let inDataSection = false;
          let headerRow = null;

          orderSheet.eachRow((row) => {
            const cell1 = getCellText(row.getCell(1)).trim().toUpperCase();
            const cell2 = getCellText(row.getCell(2)).trim().toUpperCase();
            const cell3 = getCellText(row.getCell(3)).trim().toUpperCase();

            if (!inDataSection) {
              if (cell1 === 'DATE' || cell2 === 'BUYER' || cell3 === 'FTY') {
                inDataSection = true;
                headerRow = row.row;
              }
              return;
            }

            if (headerRow && row.row === headerRow) {
              return;
            }

            let orderNo = getCellText(row.getCell(4)).trim();
            if (!orderNo) return;

            const prs = parseFloat(getCellText(row.getCell(8))) || 0;
            const buyerPrice = parseFloat(getCellText(row.getCell(15))) || 0;
            const factoryPrice = parseFloat(getCellText(row.getCell(16))) || 0;
            
            if (prs > 0) {
              tempData.push({
                orderNo,
                stockNo: getCellText(row.getCell(5)).trim(),
                styleNo: getCellText(row.getCell(6)).trim(),
                description: getCellText(row.getCell(7)).trim(),
                prs,
                buyerPrice,
                factoryPrice,
                amount: prs * buyerPrice,
                factoryAmount: prs * factoryPrice
              });
              
              tPRS += prs;
              tAmount += prs * buyerPrice;
              tFactoryAmount += prs * factoryPrice;
            }
          });

          stats.push({ 
            name: file.name, 
            orders: tempData.length,
            prs: tempData.reduce((sum, d) => sum + d.prs, 0)
          });
        } catch (err) {
          console.error("解析檔案失敗:", file.name, err);
        }
      }

      setOrderData(tempData);
      setTotals({ prs: tPRS, amount: tAmount, factoryAmount: tFactoryAmount });
      setFileStats(stats);
      setOriginalWorkbooks(parsedWbs);
      setIsProcessing(false);
    };

    processFiles();
  }, [uploadedFiles]);

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

      const createPISheet = (sheetName, priceType) => {
        const ws = newWb.addWorksheet(sheetName, {
          views: [{ showGridLines: true, zoomScale: 85 }]
        });

        ws.pageSetup = {
          orientation: 'portrait',
          paperSize: 9,
          scale: 100,
          fitToWidth: 1,
          fitToHeight: 1,
          horizontalCentered: true,
          verticalCentered: false,
          margins: { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 },
          showGridLines: false,
          showRowColHeaders: false
        };

        ws.columns = [
          { width: 15 }, { width: 12 }, { width: 15 }, { width: 25 }, 
          { width: 12 }, { width: 15 }, { width: 18 }
        ];

        ws.getCell('A1').value = formData.beneficiary;
        ws.getCell('A1').font = { size: 14, bold: true };

        ws.getCell('A4').value = formData.buyer;
        ws.getCell('A4').font = { size: 12, bold: true };

        ws.getCell('A5').value = `ATTN: ${formData.attn}`;

        ws.getCell('A8').value = 'PROFORMA INVOICE';
        ws.getCell('A8').font = { size: 16, bold: true, underline: true };

        ws.getCell('A10').value = `Date: ${formData.date}`;
        ws.getCell('A11').value = formData.paymentTerms;
        ws.getCell('A11').font = { italic: true };

        const headerRow = 13;
        ws.getRow(headerRow).values = ['ORDER NO.', 'STOCK', 'STYLE', 'DESCRIPTION', "Q'NTY", 'PRICE', 'AMOUNT/USD'];
        ws.getRow(headerRow).font = { bold: true };
        ws.getRow(headerRow).alignment = { horizontal: 'center' };

        ws.getCell(`D${headerRow + 1}`).value = "MEN'S SHOES";
        ws.getCell(`E${headerRow + 1}`).value = 'PRS';
        ws.getCell(`F${headerRow + 1}`).value = 'USD/PR';
        ws.getCell(`G${headerRow + 1}`).value = 'FOB SHENZHEN';

        let currentRow = headerRow + 2;
        let totalPrs = 0;
        let totalAmount = 0;

        orderData.forEach((item) => {
          const price = priceType === 'buyer' ? item.buyerPrice : item.factoryPrice;
          const amount = item.prs * price;
          
          ws.getCell(`A${currentRow}`).value = item.orderNo;
          ws.getCell(`B${currentRow}`).value = item.stockNo;
          ws.getCell(`C${currentRow}`).value = item.styleNo;
          ws.getCell(`D${currentRow}`).value = item.description;
          ws.getCell(`E${currentRow}`).value = item.prs;
          ws.getCell(`F${currentRow}`).value = price;
          ws.getCell(`F${currentRow}`).numFmt = '0.00';
          ws.getCell(`G${currentRow}`).value = amount;
          ws.getCell(`G${currentRow}`).numFmt = '0.00';

          totalPrs += item.prs;
          totalAmount += amount;
          currentRow++;
        });

        const totalRow = currentRow;
        ws.getCell(`A${totalRow}`).value = 'TOTAL';
        ws.getCell(`A${totalRow}`).font = { bold: true };
        ws.getCell(`E${totalRow}`).value = totalPrs;
        ws.getCell(`E${totalRow}`).font = { bold: true };
        ws.getCell(`G${totalRow}`).value = totalAmount;
        ws.getCell(`G${totalRow}`).numFmt = '0.00';
        ws.getCell(`G${totalRow}`).font = { bold: true };

        currentRow += 3;
        ws.getCell(`A${currentRow}`).value = 'BENEFICIARY:';
        ws.getCell(`A${currentRow}`).font = { bold: true };
        currentRow++;
        ws.getCell(`A${currentRow}`).value = formData.beneficiary;
        currentRow++;
        ws.getCell(`A${currentRow}`).value = formData.accountNo;
        currentRow += 2;
        ws.getCell(`A${currentRow}`).value = 'BANK:';
        ws.getCell(`A${currentRow}`).font = { bold: true };
        currentRow++;
        ws.getCell(`A${currentRow}`).value = formData.bankName;
        currentRow++;
        ws.getCell(`A${currentRow}`).value = formData.bankAddress;
        currentRow++;
        ws.getCell(`A${currentRow}`).value = `SWIFT: ${formData.swiftCode}`;

        return ws;
      };

      createPISheet('PI', 'buyer');
      createPISheet('PI_M', 'factory');

      const buffer = await newWb.xlsx.writeBuffer();
      const outName = `PI_${formData.piNo || 'Invoice'}_${formData.date}.xlsx`;
      setGeneratedBlob(new Blob([buffer]));
      setGeneratedFileName(outName);
      setGenerationState('success');
    } catch (err) {
      console.error('產生檔案失敗:', err);
      alert('產生檔���失���，請重試');
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

  const isFormValid = uploadedFiles.length > 0 && formData.buyer;

  return (
    <div className="flex flex-col bg-gray-50 p-4 md:p-8">
      <div className="max-w-[1600px] mx-auto w-full flex flex-col">

        {/* 區塊1: Invoice 標題區 - 置頂 */}
        <div className="sticky top-0 z-40 p-4 md:p-6 rounded-2xl shadow-lg flex flex-col md:flex-row items-center justify-between gap-4" style={{ backgroundColor: '#77DDFF' }}>
          <div className="text-center md:text-left">
            <h1 className="text-3xl md:text-4xl font-bold text-gray-800">Proforma Invoice</h1>
          </div>
          <button
            onClick={generateExcel}
            disabled={!isFormValid || isProcessing}
            className={`px-8 py-4 rounded-xl font-bold flex items-center gap-3 text-lg transition-all shadow-lg ${isFormValid && !isProcessing ? 'bg-gray-800 text-white hover:bg-gray-700' : 'bg-gray-300 text-gray-500 cursor-not-allowed'}`}
          >
            <FileSpreadsheet size={24} />
            {isProcessing ? '資料解析中...' : '產生 PI 檔案'}
          </button>
        </div>

        {/* 區塊2: 拖曳上傳區 */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
            className={`bg-white p-6 rounded-xl border-2 border-dashed flex flex-col items-center justify-center text-center transition-all min-h-[200px] ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300 hover:border-blue-400'}`}
          >
            <UploadCloud className="text-gray-400 mb-3" size={48} />
            <p className="text-gray-700 font-semibold">拖曳 OrderReport (Excel) 檔案至此</p>
            <p className="text-gray-400 text-sm mt-1 mb-4">或點擊下方按鈕選擇檔案</p>
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
                <button onClick={() => setUploadedFiles([])} className="text-sm text-red-500 hover:text-red-700 flex items-center gap-1"><Trash2 size={14} /> 清空</button>
              </div>

              <div className="space-y-3 mb-6">
                {fileStats.map((stat, i) => (
                  <div key={i} className="flex justify-between items-center p-3 bg-gray-50 rounded border">
                    <div className="flex items-center gap-2">
                      <FileSpreadsheet size={16} className="text-green-600" />
                      <span className="font-medium text-sm text-gray-700 truncate max-w-[200px]">{stat.name}</span>
                    </div>
                    <div className="flex gap-4 text-sm text-gray-600">
                      <span>訂單: <b>{stat.orders}</b></span>
                      <span>PRS: <b>{stat.prs}</b></span>
                    </div>
                  </div>
                ))}
              </div>

              <div className="bg-blue-50 rounded-lg p-4 border border-blue-100">
                <h3 className="font-bold text-blue-800 mb-3 text-sm">彙總數據計算 (將帶入 PI Sheet)</h3>
                <div className="grid grid-cols-3 gap-4 text-center">
                  <div>
                    <p className="text-xs text-blue-600 font-semibold mb-1">總數量 (PRS)</p>
                    <p className="text-xl font-bold text-blue-900">{totals.prs}</p>
                  </div>
                  <div>
                    <p className="text-xs text-blue-600 font-semibold mb-1">買家金額 (USD)</p>
                    <p className="text-xl font-bold text-blue-900">{totals.amount.toFixed(2)}</p>
                  </div>
                  <div>
                    <p className="text-xs text-blue-600 font-semibold mb-1">工廠金額 (USD)</p>
                    <p className="text-xl font-bold text-blue-900">{totals.factoryAmount.toFixed(2)}</p>
                  </div>
                </div>
              </div>
            </div>
          ) : (
            <div className="lg:col-span-2 bg-gray-100 rounded-xl border-2 border-dashed border-gray-200 flex items-center justify-center min-h-[200px]">
              <p className="text-gray-400">上傳 OrderReport 檔案後顯示預覽</p>
            </div>
          )}
        </div>

        {/* 區塊3: PI 基本資訊 */}
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-5">
          <h2 className="text-lg font-bold border-b pb-2 mb-4">PI 基本資訊</h2>

          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4">
            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">PI#</label>
              <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.piNo} onChange={e => setFormData({...formData, piNo: e.target.value})} />
            </div>
            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">Date</label>
              <input type="date" className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.date} onChange={e => setFormData({...formData, date: e.target.value})} />
            </div>
            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">Buyer</label>
              <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.buyer} onChange={e => setFormData({...formData, buyer: e.target.value})} />
            </div>
            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">ATTN</label>
              <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.attn} onChange={e => setFormData({...formData, attn: e.target.value})} />
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">Payment Terms</label>
              <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.paymentTerms} onChange={e => setFormData({...formData, paymentTerms: e.target.value})} />
            </div>
            <div>
              <label className="block text-sm font-semibold text-gray-700 mb-1">Beneficiary</label>
              <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.beneficiary} onChange={e => setFormData({...formData, beneficiary: e.target.value})} />
            </div>
          </div>

          <div className="pt-4 border-t">
            <h3 className="text-lg font-bold border-b pb-2 mb-4">銀行資訊</h3>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1">Account No.</label>
                <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.accountNo} onChange={e => setFormData({...formData, accountNo: e.target.value})} />
              </div>
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1">Bank Name</label>
                <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.bankName} onChange={e => setFormData({...formData, bankName: e.target.value})} />
              </div>
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1">Bank Address</label>
                <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.bankAddress} onChange={e => setFormData({...formData, bankAddress: e.target.value})} />
              </div>
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1">SWIFT Code</label>
                <input className="w-full border rounded p-2 focus:ring-2 focus:ring-blue-500 outline-none" value={formData.swiftCode} onChange={e => setFormData({...formData, swiftCode: e.target.value})} />
              </div>
            </div>
          </div>
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
                <button onClick={handleDownload} className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-lg flex items-center gap-2 font-semibold">
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