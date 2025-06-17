document.addEventListener('DOMContentLoaded', function() {
  const fileUpload = document.getElementById('file-upload');
  const uploadBtn = document.getElementById('upload-btn');
  const cancelBtn = document.getElementById('cancel-btn');
  const useDataBtn = document.getElementById('use-data-btn');
  const downloadTemplateBtn = document.getElementById('download-template');
  const fileTypeSelect = document.getElementById('file-type');
  const filePreview = document.getElementById('file-preview');
  const previewHeaders = document.getElementById('preview-headers');
  const previewData = document.getElementById('preview-data');
  const resultsDiv = document.getElementById('calculation-results');
  const resultsContainer = document.getElementById('results-container');
  const errorDiv = document.getElementById('error-message');

  // รีเซ็ตฟอร์ม
  function resetForm() {
    fileUpload.value = '';
    filePreview.style.display = 'none';
    resultsDiv.style.display = 'none';
    errorDiv.style.display = 'none';
  }

  // แสดงข้อผิดพลาด
  function showError(message) {
    errorDiv.textContent = message;
    errorDiv.style.display = 'block';
    setTimeout(() => errorDiv.style.display = 'none', 5000);
  }

  // ดาวน์โหลดไฟล์ตัวอย่าง
  downloadTemplateBtn.addEventListener('click', function() {
    const fileType = fileTypeSelect.value;
    let fileName, wb;
    
    switch(fileType) {
      case 'cashflow':
        const cashflowData = [
          { date: '2023-01-01', description: 'รายได้จากขายสินค้า', amount: 50000 },
          { date: '2023-01-05', description: 'ค่าจ้างพนักงาน', amount: -20000 },
          { date: '2023-01-10', description: 'ค่าวัสดุ', amount: -15000 }
        ];
        const wsCashflow = XLSX.utils.json_to_sheet(cashflowData);
        wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, wsCashflow, 'กระแสเงินสด');
        fileName = 'cashflow-template.xlsx';
        break;
      
      case 'income':
        const incomeData = [
          { type: 'income', description: 'รายได้จากขายสินค้า', amount: 50000, date: '2023-01-01' },
          { type: 'expense', description: 'ค่าจ้างพนักงาน', amount: 20000, date: '2023-01-05' },
          { type: 'expense', description: 'ค่าวัสดุ', amount: 15000, date: '2023-01-10' }
        ];
        const wsIncome = XLSX.utils.json_to_sheet(incomeData);
        wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, wsIncome, 'รายรับ-รายจ่าย');
        fileName = 'income-expense-template.xlsx';
        break;
      
      case 'investment':
        const investmentData = [
          { project: 'โครงการ A', investment: 100000, return: 120000, period: 12 },
          { project: 'โครงการ B', investment: 50000, return: 60000, period: 6 }
        ];
        const wsInvestment = XLSX.utils.json_to_sheet(investmentData);
        wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, wsInvestment, 'การลงทุน');
        fileName = 'investment-template.xlsx';
        break;
      
      case 'business-valuation':
  // สร้างข้อมูลสำหรับ Business Valuation ตามไฟล์ตัวอย่างใหม่
  wb = XLSX.utils.book_new();
  
  // Sheet 1: income_statement
  const incomeStatementData = [
    ["ผลตอบแทน", "", "", "", "", ""],
    ["ผลิตภัณฑ์ (1) ที่สามารถผลิตได้ต่อปี", 1000, 1200, 1400, 1600, 1800],
    ["ราคาต่อหน่วย (1)", 50, 55, 60, 65, 70],
    ["ผลิตภัณฑ์ (2) ที่สามารถผลิตได้ต่อปี", 500, 600, 700, 800, 900],
    ["ราคาต่อหน่วย (2)", 80, 85, 90, 95, 100],
    ["", "", "", "", "", ""],
    ["ต้นทุน (ลงทุน)", "", "", "", "", ""],
    ["การซื้อเทคโนโลยี", 50000, 0, 0, 0, 0],
    ["ค่าวัตถุดิบ", 20000, 22000, 24000, 26000, 28000],
    ["การปรับปรุงโรงงานเดิม", 30000, 0, 0, 0, 0],
    ["ยานพาหนะและขนส่ง", 10000, 0, 0, 0, 0],
    ["เครื่องจักร", 50000, 0, 0, 0, 0],
    ["เครื่องมือ", 10000, 0, 0, 0, 0],
    ["แบบต่างๆ", 5000, 0, 0, 0, 0],
    ["สิ่งปลูกสร้างด้านการผลิต", 100000, 0, 0, 0, 0],
    ["สิ่งปลูกสร้างที่ไม่เกี่ยวกับการผลิต", 50000, 0, 0, 0, 0],
    ["เฟอร์นิเจอร์และอุปกรณ์สำนักงาน", 20000, 0, 0, 0, 0],
    ["อื่น ๆ", 10000, 0, 0, 0, 0]
  ];
  const wsIncomeStatement = XLSX.utils.aoa_to_sheet(incomeStatementData);
  XLSX.utils.book_append_sheet(wb, wsIncomeStatement, "income_statement");
  
  // Sheet 2: งบลงทุน
  const investmentSheetData = [
    ["เงินลงทุน", "ปีที่ 1", "ปีที่ 2", "ปีที่ 3", "ปีที่ 4", "ปีที่ 5"],
    ["", "ปีที่ 1", "ปีที่ 2", "ปีที่ 3", "ปีที่ 4", "ปีที่ 5"],
    ["ค่าออกแบบปรับปรุงอาคาร", 100000, 0, 0, 0, 0],
    ["ค่าปรับปรุงอาคาร", 50000, 0, 0, 0, 0],
    ["ระบบ/ลิขสิทธิ์", 20000, 0, 0, 0, 0],
    ["", "", "", "", "", ""],
    ["รวมเงินลงทุน", "=SUM(B3:B5)", "=SUM(C3:C5)", "=SUM(D3:D5)", "=SUM(E3:E5)", "=SUM(F3:F5)"]
  ];
  const wsInvestmentSheet = XLSX.utils.aoa_to_sheet(investmentSheetData);
  XLSX.utils.book_append_sheet(wb, wsInvestmentSheet, "งบลงทุน");
  
  // Sheet 3: งบบุคลากร
  const personnelData = [
    ["รายจ่ายต่อเดือน", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["", "ปีที่ 1", "ปีที่ 2", "ปีที่ 3", "ปีที่ 4", "ปีที่ 5"],
    ["กรรมการผู้จัดการ", 50000, 52000, 54000, 56000, 58000],
    ["ผู้จัดการฝ่ายการตลาด", 40000, 42000, 44000, 46000, 48000],
    ["เจ้าหน้าที่การตลาด", 30000, 32000, 34000, 36000, 38000],
    ["ผู้จัดการฝ่ายการผลิต", 45000, 47000, 49000, 51000, 53000],
    ["วิศวกรฝ่ายผลิต", 35000, 37000, 39000, 41000, 43000],
    ["แรงงานผลิต", 20000, 22000, 24000, 26000, 28000],
    ["พนักงานสำนักงาน บัญชี การเงิน", 25000, 27000, 29000, 31000, 33000],
    ["แม่บ้าน/รปภ.", 15000, 16000, 17000, 18000, 19000],
    ["", "", "", "", "", ""],
    ["รวมค่าใช้จ่ายบุคลากรต่อปี", "=SUM(C4:C11)*12", "=SUM(D4:D11)*12", "=SUM(E4:E11)*12", "=SUM(F4:F11)*12", "=SUM(G4:G11)*12"]
  ];
  const wsPersonnel = XLSX.utils.aoa_to_sheet(personnelData);
  XLSX.utils.book_append_sheet(wb, wsPersonnel, "งบบุคลากร");
  
  // Sheet 4: ค่าใช้จ่ายในการบริหาร
  const adminCostData = [
    ["ค่าใช้จ่ายในการบริหาร", "", "", "", "", ""],
    ["", "ปีที่ 1", "ปีที่ 2", "ปีที่ 3", "ปีที่ 4", "ปีที่ 5"],
    ["ค่าเช่าตึก", 50000, 55000, 60000, 65000, 70000],
    ["ค่าใช้จ่ายบุคลากร", "=งบบุคลากร!C12", "=งบบุคลากร!D12", "=งบบุคลากร!E12", "=งบบุคลากร!F12", "=งบบุคลากร!G12"],
    ["ค่าสาธารณูปโภค", 20000, 22000, 24000, 26000, 28000],
    ["ค่าโทรศัพท์", 5000, 5500, 6000, 6500, 7000],
    ["ค่าใช้จ่าย รปภ", 10000, 11000, 12000, 13000, 14000],
    ["ค่าวัสดุสำนักงาน", 5000, 5500, 6000, 6500, 7000],
    ["ค่าซ่อมบำรุง", 10000, 11000, 12000, 13000, 14000],
    ["ค่าใช้จ่ายอื่นๆ", 5000, 5500, 6000, 6500, 7000],
    ["ค่าจ้างเหมาบริการรถยนต์", 15000, 16000, 17000, 18000, 19000],
    ["ค่าน้ำมันรถยนต์", 10000, 11000, 12000, 13000, 14000],
    ["ค่าอบรม", 20000, 22000, 24000, 26000, 28000],
    ["", "", "", "", "", ""],
    ["ค่าเสื่อมราคา", 0, "=งบลงทุน!C7", "=งบลงทุน!D7", "=งบลงทุน!E7", "=งบลงทุน!F7"],
    ["รวมค่าใช้จ่าย", "=SUM(B3:B15)", "=SUM(C3:C15)", "=SUM(D3:D15)", "=SUM(E3:E15)", "=SUM(F3:F15)"]
  ];
  const wsAdminCost = XLSX.utils.aoa_to_sheet(adminCostData);
  XLSX.utils.book_append_sheet(wb, wsAdminCost, "ค่าใช้จ่ายในการบริหาร");
  
  fileName = 'business-valuation-template.xlsx';
  break;
}

    // สร้างไฟล์ Excel
    XLSX.writeFile(wb, fileName);
  });

  // อ่านไฟล์ CSV
  function parseCSV(content) {
    try {
      const lines = content.split('\n');
      const headers = lines[0].split(',').map(h => h.trim());
      const result = [];

      for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;
        
        const obj = {};
        const currentline = lines[i].split(',');

        for (let j = 0; j < headers.length; j++) {
          obj[headers[j]] = currentline[j] ? currentline[j].trim() : '';
        }

        result.push(obj);
      }

      return { headers, data: result };
    } catch (e) {
      showError('รูปแบบไฟล์ CSV ไม่ถูกต้อง');
      return null;
    }
  }

  // อ่านไฟล์ JSON
  function parseJSON(content) {
    try {
      const data = JSON.parse(content);
      if (!Array.isArray(data) || data.length === 0) {
        showError('ไฟล์ JSON ต้องเป็นอาร์เรย์ของข้อมูล');
        return null;
      }

      const headers = Object.keys(data[0]);
      return { headers, data };
    } catch (e) {
      showError('รูปแบบไฟล์ JSON ไม่ถูกต้อง');
      return null;
    }
  }

  async function parseExcel(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetsData = [];
          
          workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            sheetsData.push({
              name: sheetName,
              data: jsonData
            });
          });
          
          resolve(sheetsData);
        } catch (e) {
          reject('รูปแบบไฟล์ Excel ไม่ถูกต้อง');
        }
      };
      
      reader.onerror = function() {
        reject('เกิดข้อผิดพลาดในการอ่านไฟล์ Excel');
      };
      
      reader.readAsArrayBuffer(file);
    });
  }

  // แสดงตัวอย่างข้อมูล
  function displayPreview(parsedData) {
    if (Array.isArray(parsedData)) {
      // กรณี Excel ที่มีหลาย sheets
      filePreview.innerHTML = '';
      
      parsedData.forEach(sheet => {
        const sheetDiv = document.createElement('div');
        sheetDiv.innerHTML = `<h4>${sheet.name}</h4>`;
        
        const table = document.createElement('table');
        table.className = 'table';
        
        // สร้าง header จากแถวแรก (ถ้ามี)
        if (sheet.data.length > 0) {
          const thead = document.createElement('thead');
          const headerRow = document.createElement('tr');
          
          sheet.data[0].forEach(cell => {
            const th = document.createElement('th');
            th.textContent = cell || '';
            headerRow.appendChild(th);
          });
          
          thead.appendChild(headerRow);
          table.appendChild(thead);
        }
        
        // สร้างข้อมูล (แสดงแค่ 5 แถวแรก)
        const tbody = document.createElement('tbody');
        const rowsToShow = Math.min(sheet.data.length, 6); // แสดง header + 5 แถว
        
        for (let i = 1; i < rowsToShow; i++) {
          const row = document.createElement('tr');
          
          sheet.data[i].forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell || '';
            row.appendChild(td);
          });
          
          tbody.appendChild(row);
        }
        
        table.appendChild(tbody);
        sheetDiv.appendChild(table);
        filePreview.appendChild(sheetDiv);
      });
    } else {
      // กรณี CSV/JSON
      previewHeaders.innerHTML = parsedData.headers.map(h => `<th>${h}</th>`).join('');
      previewData.innerHTML = '';

      // แสดงเพียง 5 แถวแรก
      const rowsToShow = Math.min(parsedData.data.length, 5);
      for (let i = 0; i < rowsToShow; i++) {
        const row = document.createElement('tr');
        row.innerHTML = parsedData.headers.map(h => `<td>${parsedData.data[i][h] || '-'}</td>`).join('');
        previewData.appendChild(row);
      }
    }

    filePreview.style.display = 'block';
  }

  // ประมวลผลไฟล์
  async function processFile(file, fileType) {
    try {
      let parsedData;
      
      if (file.name.endsWith('.csv')) {
        const content = await readFileAsText(file);
        parsedData = parseCSV(content);
      } else if (file.name.endsWith('.json')) {
        const content = await readFileAsText(file);
        parsedData = parseJSON(content);
      } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        parsedData = await parseExcel(file);
      } else {
        showError('รองรับเฉพาะไฟล์ CSV, JSON และ Excel');
        return;
      }

      if (parsedData) {
        displayPreview(parsedData);
        analyzeData(fileType, parsedData);
      }
    } catch (error) {
      showError(error);
    }
  }

  // ฟังก์ชันอ่านไฟล์เป็นข้อความ
  function readFileAsText(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject('เกิดข้อผิดพลาดในการอ่านไฟล์');
      reader.readAsText(file);
    });
  }

  // วิเคราะห์ข้อมูลตามประเภท
  function analyzeData(fileType, data) {
    let analysisResult = '';
    
    switch(fileType) {
      case 'cashflow':
        analysisResult = analyzeCashFlow(data);
        break;
      case 'income':
        analysisResult = analyzeIncome(data);
        break;
      case 'investment':
        analysisResult = analyzeInvestment(data);
        break;
      case 'business-valuation':
        analysisResult = analyzeBusinessValuation(data);
        break;
      default:
        analysisResult = '<p>ไม่สามารถวิเคราะห์ข้อมูลประเภทนี้ได้</p>';
    }

    resultsContainer.innerHTML = analysisResult;
    resultsDiv.style.display = 'block';
  }

  // วิเคราะห์กระแสเงินสด
  function analyzeCashFlow(data) {
    let totalCashFlow = 0;
    let positiveMonths = 0;
    let negativeMonths = 0;
    
    // ตรวจสอบว่าเป็นข้อมูลจาก Excel หรือไม่
    const cashFlowData = Array.isArray(data) ? 
      convertExcelToCashFlow(data[0].data) : 
      data.data;

    cashFlowData.forEach(item => {
      let amount = 0;
      
      if (typeof item === 'object') {
        amount = parseFloat(item.amount) || 0;
      } else if (Array.isArray(item) && item.length >= 3) {
        amount = parseFloat(item[2]) || 0;
      }
      
      totalCashFlow += amount;
      
      if (amount > 0) positiveMonths++;
      else if (amount < 0) negativeMonths++;
    });

    return `
      <div class="result-item">
        <h4>กระแสเงินสดรวม: <span class="result-value">${totalCashFlow.toLocaleString('th-TH')} บาท</span></h4>
      </div>
      <div class="result-item">
        <h4>เดือนที่กระแสเงินสดบวก: <span class="result-value">${positiveMonths} เดือน</span></h4>
      </div>
      <div class="result-item">
        <h4>เดือนที่กระแสเงินสดลบ: <span class="result-value">${negativeMonths} เดือน</span></h4>
      </div>
    `;
  }

  // แปลงข้อมูล Excel เป็นรูปแบบกระแสเงินสด
  function convertExcelToCashFlow(excelData) {
    const result = [];
    const headers = excelData[0];
    
    for (let i = 1; i < excelData.length; i++) {
      const row = excelData[i];
      if (!row || row.length === 0) continue;
      
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] || '';
      });
      result.push(obj);
    }
    
    return result;
  }

  // วิเคราะห์รายรับ-รายจ่าย
  function analyzeIncome(data) {
    let totalIncome = 0;
    let totalExpense = 0;
    
    // ตรวจสอบว่าเป็นข้อมูลจาก Excel หรือไม่
    const incomeData = Array.isArray(data) ? 
      convertExcelToIncome(data[0].data) : 
      data.data;

    incomeData.forEach(item => {
      let type = '';
      let amount = 0;
      
      if (typeof item === 'object') {
        type = item.type || '';
        amount = parseFloat(item.amount) || 0;
      } else if (Array.isArray(item) && item.length >= 3) {
        type = item[0] || '';
        amount = parseFloat(item[2]) || 0;
      }
      
      if (type === 'income' || type === 'รายรับ') {
        totalIncome += amount;
      } else if (type === 'expense' || type === 'รายจ่าย') {
        totalExpense += amount;
      }
    });

    const profit = totalIncome - totalExpense;
    const profitMargin = totalIncome > 0 ? (profit / totalIncome) * 100 : 0;

    return `
      <div class="result-item">
        <h4>รายรับรวม: <span class="result-value">${totalIncome.toLocaleString('th-TH')} บาท</span></h4>
      </div>
      <div class="result-item">
        <h4>รายจ่ายรวม: <span class="result-value">${totalExpense.toLocaleString('th-TH')} บาท</span></h4>
      </div>
      <div class="result-item">
        <h4>กำไรสุทธิ: <span class="result-value">${profit.toLocaleString('th-TH')} บาท</span></h4>
      </div>
      <div class="result-item">
        <h4>อัตรากำไร: <span class="result-value">${profitMargin.toFixed(2)}%</span></h4>
      </div>
    `;
  }

  // แปลงข้อมูล Excel เป็นรูปแบบรายรับ-รายจ่าย
  function convertExcelToIncome(excelData) {
    const result = [];
    const headers = excelData[0];
    
    for (let i = 1; i < excelData.length; i++) {
      const row = excelData[i];
      if (!row || row.length === 0) continue;
      
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] || '';
      });
      result.push(obj);
    }
    
    return result;
  }

  // วิเคราะห์การลงทุน
  function analyzeInvestment(data) {
    let totalInvestment = 0;
    let totalReturn = 0;
    
    // ตรวจสอบว่าเป็นข้อมูลจาก Excel หรือไม่
    const investmentData = Array.isArray(data) ? 
      convertExcelToInvestment(data[0].data) : 
      data.data;

    investmentData.forEach(item => {
      let investment = 0;
      let returnVal = 0;
      
      if (typeof item === 'object') {
        investment = parseFloat(item.investment) || 0;
        returnVal = parseFloat(item.return) || 0;
      } else if (Array.isArray(item) && item.length >= 3) {
        investment = parseFloat(item[1]) || 0;
        returnVal = parseFloat(item[2]) || 0;
      }
      
      totalInvestment += investment;
      totalReturn += returnVal;
    });

    const netReturn = totalReturn - totalInvestment;
    const roi = totalInvestment > 0 ? (netReturn / totalInvestment) * 100 : 0;

    return `
      <div class="result-item">
        <h4>เงินลงทุนรวม: <span class="result-value">${totalInvestment.toLocaleString('th-TH')} บาท</span></h4>
      </div>
      <div class="result-item">
        <h4>ผลตอบแทนรวม: <span class="result-value">${totalReturn.toLocaleString('th-TH')} บาท</span></h4>
      </div>
      <div class="result-item">
        <h4>ผลตอบแทนสุทธิ: <span class="result-value">${netReturn.toLocaleString('th-TH')} บาท</span></h4>
      </div>
      <div class="result-item">
        <h4>อัตราผลตอบแทน (ROI): <span class="result-value">${roi.toFixed(2)}%</span></h4>
      </div>
    `;
  }

  // แปลงข้อมูล Excel เป็นรูปแบบการลงทุน
  function convertExcelToInvestment(excelData) {
    const result = [];
    const headers = excelData[0];
    
    for (let i = 1; i < excelData.length; i++) {
      const row = excelData[i];
      if (!row || row.length === 0) continue;
      
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] || '';
      });
      result.push(obj);
    }
    
    return result;
  }

// แก้ไขฟังก์ชัน analyzeBusinessValuation ใน upload.js
function analyzeBusinessValuation(data) {
  // หา sheet ที่ต้องการ
  let incomeStatement, investmentSheet, personnelSheet, adminCostSheet;
  
  if (Array.isArray(data)) {
    incomeStatement = data.find(sheet => 
      sheet.name.toLowerCase().includes('income') || 
      sheet.name.includes('ผลตอบแทน') ||
      sheet.name.includes('statement')
    );
    investmentSheet = data.find(sheet => 
      sheet.name.includes('ลงทุน') || 
      sheet.name.includes('เงินลงทุน') ||
      sheet.name.toLowerCase().includes('investment')
    );
    personnelSheet = data.find(sheet => 
      sheet.name.includes('บุคลากร') || 
      sheet.name.includes('พนักงาน') ||
      sheet.name.toLowerCase().includes('personnel')
    );
    adminCostSheet = data.find(sheet => 
      sheet.name.includes('บริหาร') || 
      sheet.name.includes('ค่าใช้จ่าย') ||
      sheet.name.toLowerCase().includes('admin')
    );
  } else {
    // กรณีไม่ใช่ Excel
    incomeStatement = data;
  }
  
  // สร้างผลลัพธ์ HTML
  let resultHTML = `
    <div class="valuation-analysis">
      <h4>ผลการวิเคราะห์ธุรกิจ</h4>
  `;
  
  // ข้อมูลรายได้
  resultHTML += `<div class="valuation-section"><h5>ข้อมูลรายได้</h5>`;
  
  if (incomeStatement) {
    const incomeData = Array.isArray(incomeStatement) ? incomeStatement.data : incomeStatement;
    
    // หาข้อมูลผลิตภัณฑ์
    const products = [];
    for (let i = 0; i < incomeData.length; i++) {
      const row = incomeData[i];
      if (row && row[0] && typeof row[0] === 'string' && 
          row[0].toString().includes('ผลิตภัณฑ์') && 
          row[0].toString().includes('ที่สามารถผลิตได้ต่อปี')) {
        const priceRowIndex = findPriceRow(incomeData, i);
        const priceRow = priceRowIndex !== -1 ? incomeData[priceRowIndex] : null;
        const price = priceRow && priceRow[1] ? parseFloat(priceRow[1]) : 0;
        
        products.push({
          name: row[0],
          year1: row[1] ? parseFloat(row[1]) : 0,
          priceRow: priceRowIndex,
          price: price
        });
      }
    }
    
    // แสดงข้อมูลผลิตภัณฑ์
    if (products.length > 0) {
      products.forEach(product => {
        resultHTML += `
          <p>${product.name}: ${product.year1.toLocaleString('th-TH')} หน่วย (ราคาต่อหน่วย: ${product.price.toLocaleString('th-TH')} บาท)</p>
        `;
      });
    } else {
      resultHTML += `<p>ไม่พบข้อมูลผลิตภัณฑ์</p>`;
    }
    
    // หาข้อมูลต้นทุน
    const costItems = [];
    const costStart = incomeData.findIndex(row => 
      row && row[0] && typeof row[0] === 'string' && 
      row[0].toString().includes('ต้นทุน')
    );
    
    if (costStart > -1) {
      for (let i = costStart + 1; i < incomeData.length; i++) {
        const row = incomeData[i];
        if (row && row[0] && typeof row[0] === 'string' && 
            row[0].toString().trim() && 
            !row[0].toString().includes('ค่าใช้จ่าย')) {
          const amount = row[1] ? parseFloat(row[1]) : 0;
          if (amount > 0) {
            costItems.push({
              name: row[0],
              amount: amount
            });
          }
        } else {
          break;
        }
      }
    }
    
    if (costItems.length > 0) {
      resultHTML += `<h6>ต้นทุนการผลิต:</h6><ul>`;
      costItems.forEach(item => {
        resultHTML += `<li>${item.name}: ${item.amount.toLocaleString('th-TH')} บาท</li>`;
      });
      resultHTML += `</ul>`;
    } else {
      resultHTML += `<p>ไม่พบข้อมูลต้นทุนการผลิต</p>`;
    }
  } else {
    resultHTML += `<p>ไม่พบข้อมูลรายได้</p>`;
  }
  
  resultHTML += `</div>`;
  
  // ข้อมูลการลงทุน
  resultHTML += `<div class="valuation-section"><h5>ข้อมูลการลงทุน</h5>`;
  
  if (investmentSheet) {
    const investmentData = investmentSheet.data;
    
    // หาแถวที่มีผลรวม
    const totalInvestmentRow = investmentData.find(row => 
      row && row[0] && typeof row[0] === 'string' && 
      row[0].toString().includes('รวมเงินลงทุน')
    );
    
    if (totalInvestmentRow && totalInvestmentRow.length >= 2) {
      const totalInvestment = parseFloat(totalInvestmentRow[1]) || 0;
      resultHTML += `<p>รวมเงินลงทุนปีที่ 1: ${totalInvestment.toLocaleString('th-TH')} บาท</p>`;
    }
    
    // แสดงรายการลงทุน
    const investmentItems = [];
    for (let i = 2; i < investmentData.length; i++) {
      const row = investmentData[i];
      if (row && row[0] && typeof row[0] === 'string' && 
          row[0].toString().trim() && 
          !row[0].toString().includes('รวม') && 
          !row[0].toString().includes('ค่าเสื่อม')) {
        const amount = row[1] ? parseFloat(row[1]) : 0;
        if (amount > 0) {
          investmentItems.push({
            name: row[0],
            amount: amount
          });
        }
      }
    }
    
    if (investmentItems.length > 0) {
      resultHTML += `<h6>รายการลงทุนหลัก:</h6><ul>`;
      investmentItems.forEach(item => {
        resultHTML += `<li>${item.name}: ${item.amount.toLocaleString('th-TH')} บาท</li>`;
      });
      resultHTML += `</ul>`;
    } else {
      resultHTML += `<p>ไม่พบรายการลงทุน</p>`;
    }
  } else {
    resultHTML += `<p>ไม่พบข้อมูลการลงทุน</p>`;
  }
  
  resultHTML += `</div>`;
  
  // ข้อมูลบุคลากร
  resultHTML += `<div class="valuation-section"><h5>ข้อมูลบุคลากร</h5>`;
  
  if (personnelSheet) {
    const personnelData = personnelSheet.data;
    
    const totalPersonnelRow = personnelData.find(row => 
      row && row[0] && typeof row[0] === 'string' && 
      row[0].toString().includes('รวมค่าใช้จ่ายบุคลากร')
    );
    
    if (totalPersonnelRow && totalPersonnelRow.length >= 2) {
      const totalPersonnel = parseFloat(totalPersonnelRow[1]) || 0;
      resultHTML += `<p>ค่าใช้จ่ายบุคลากรต่อปี: ${totalPersonnel.toLocaleString('th-TH')} บาท</p>`;
    }
    
    // แสดงรายการบุคลากร
    const personnelItems = [];
    for (let i = 3; i < personnelData.length; i++) {
      const row = personnelData[i];
      if (row && row[0] && typeof row[0] === 'string' && 
          row[0].toString().trim() && 
          !row[0].toString().includes('รวม')) {
        const salary = row[1] ? parseFloat(row[1]) : 0;
        if (salary > 0) {
          personnelItems.push({
            name: row[0],
            salary: salary
          });
        }
      }
    }
    
    if (personnelItems.length > 0) {
      resultHTML += `<h6>รายการบุคลากรหลัก:</h6><ul>`;
      personnelItems.forEach(item => {
        resultHTML += `<li>${item.name}: ${item.salary.toLocaleString('th-TH')} บาท/เดือน</li>`;
      });
      resultHTML += `</ul>`;
    } else {
      resultHTML += `<p>ไม่พบรายการบุคลากร</p>`;
    }
  } else {
    resultHTML += `<p>ไม่พบข้อมูลบุคลากร</p>`;
  }
  
  resultHTML += `</div>`;
  
  // ค่าใช้จ่ายในการบริหาร
  resultHTML += `<div class="valuation-section"><h5>ค่าใช้จ่ายในการบริหาร</h5>`;
  
  if (adminCostSheet) {
    const adminData = adminCostSheet.data;
    
    const totalAdminRow = adminData.find(row => 
      row && row[0] && typeof row[0] === 'string' && 
      row[0].toString().includes('รวมค่าใช้จ่าย')
    );
    
    if (totalAdminRow && totalAdminRow.length >= 2) {
      const totalAdmin = parseFloat(totalAdminRow[1]) || 0;
      resultHTML += `<p>ค่าใช้จ่ายบริหารรวม: ${totalAdmin.toLocaleString('th-TH')} บาท</p>`;
    }
    
    // แสดงรายการค่าใช้จ่าย
    const costItems = [];
    for (let i = 2; i < adminData.length; i++) {
      const row = adminData[i];
      if (row && row[0] && typeof row[0] === 'string' && 
          row[0].toString().trim() && 
          !row[0].toString().includes('รวม') && 
          !row[0].toString().includes('ค่าเสื่อม')) {
        const amount = row[1] ? parseFloat(row[1]) : 0;
        if (amount > 0) {
          costItems.push({
            name: row[0],
            amount: amount
          });
        }
      }
    }
    
    if (costItems.length > 0) {
      resultHTML += `<h6>รายการค่าใช้จ่ายหลัก:</h6><ul>`;
      costItems.forEach(item => {
        resultHTML += `<li>${item.name}: ${item.amount.toLocaleString('th-TH')} บาท</li>`;
      });
      resultHTML += `</ul>`;
    } else {
      resultHTML += `<p>ไม่พบรายการค่าใช้จ่าย</p>`;
    }
  } else {
    resultHTML += `<p>ไม่พบข้อมูลค่าใช้จ่ายบริหาร</p>`;
  }
  
  resultHTML += `</div></div>`;
  
  return resultHTML;
}

// ฟังก์ชันช่วยเหลือ
function findPriceRow(data, productRowIndex) {
  if (!Array.isArray(data)) return -1;
  
  for (let i = productRowIndex + 1; i < data.length; i++) {
    const row = data[i];
    if (row && row[0] && typeof row[0] === 'string' && 
        row[0].toString().includes('ราคาต่อหน่วย')) {
      return i;
    }
  }
  return -1;
}

function formatNumber(value) {
  if (typeof value === 'string' && value.startsWith('=')) {
    return 'สูตรคำนวณ';
  }
  const num = parseFloat(value) || 0;
  return num.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

  // Event Listeners
  uploadBtn.addEventListener('click', async function() {
    const file = fileUpload.files[0];
    if (!file) {
      showError('กรุณาเลือกไฟล์');
      return;
    }

    const fileType = document.getElementById('file-type').value;
    await processFile(file, fileType);
  });

  cancelBtn.addEventListener('click', resetForm);

  useDataBtn.addEventListener('click', function() {
    // บันทึกข้อมูลไปใช้ในหน้าอื่น (ใช้ localStorage)
    localStorage.setItem('financialData', JSON.stringify({
      type: document.getElementById('file-type').value,
      file: fileUpload.files[0].name,
      timestamp: new Date().toISOString()
    }));
    
    alert('บันทึกข้อมูลเรียบร้อยแล้ว สามารถใช้ในการคำนวณในหน้าอื่นได้');
    window.location.href = 'index.html';
  });
});