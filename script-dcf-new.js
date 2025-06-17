document.addEventListener('DOMContentLoaded', function() {
  // ตัวแปรหลัก
  const initialInvestmentInput = document.getElementById('initial-investment');
  const baseCashFlowInput = document.getElementById('base-cash-flow');
  const waccInput = document.getElementById('wacc');
  const growthRateTable = document.getElementById('growth-rate-table').getElementsByTagName('tbody')[0];
  const riskFactorsTable = document.getElementById('risk-factors-table').getElementsByTagName('tbody')[0];
  const dcfNewTable = document.getElementById('dcf-new-tbody');
  const calculateBtn = document.getElementById('calculate-dcf-new');
  
  // ผลลัพธ์
  const npvResult = document.getElementById('dcf-new-npv');
  const irrResult = document.getElementById('dcf-new-irr');
  const piResult = document.getElementById('dcf-new-pi');
  
  // Sensitivity Analysis
  const sensitivityResults = {
    6: { npv: document.getElementById('sens-6-npv'), pi: document.getElementById('sens-6-pi') },
    7: { npv: document.getElementById('sens-7-npv'), pi: document.getElementById('sens-7-pi') },
    8: { npv: document.getElementById('sens-8-npv'), pi: document.getElementById('sens-8-pi') },
    9: { npv: document.getElementById('sens-9-npv'), pi: document.getElementById('sens-9-pi') },
    10: { npv: document.getElementById('sens-10-npv'), pi: document.getElementById('sens-10-pi') },
    11: { npv: document.getElementById('sens-11-npv'), pi: document.getElementById('sens-11-pi') },
    12: { npv: document.getElementById('sens-12-npv'), pi: document.getElementById('sens-12-pi') }
  };

  // ฟังก์ชันเพิ่มช่วงปีใหม่
  function addNewGrowthRate() {
    const lastRow = growthRateTable.rows[growthRateTable.rows.length - 1];
    let newStart = 1, newEnd = 1;
    
    if (lastRow) {
      newStart = parseInt(lastRow.cells[1].querySelector('input').value) + 1;
      newEnd = newStart;
    }
    
    const row = growthRateTable.insertRow();
    row.innerHTML = `
      <td><input type="number" class="start-year" value="${newStart}" min="1"></td>
      <td><input type="number" class="end-year" value="${newEnd}" min="1"></td>
      <td><input type="number" class="growth-rate" value="5.00" step="0.01"></td>
      <td><button class="btn btn-danger remove-growth-rate">ลบ</button></td>
    `;
    
    row.querySelector('.remove-growth-rate').addEventListener('click', function() {
      if (growthRateTable.rows.length > 1) {
        row.remove();
        calculateNewDCF();
      } else {
        alert('ต้องมีอย่างน้อย 1 ช่วงปี');
      }
    });

    row.querySelectorAll('input').forEach(input => {
      input.addEventListener('input', calculateNewDCF);
    });

    calculateNewDCF();
  }

  // ฟังก์ชันจัดเรียงปี
  function autoArrangeYears() {
    const rows = Array.from(growthRateTable.rows);
    rows.sort((a, b) => {
      const aStart = parseInt(a.cells[0].querySelector('input').value);
      const bStart = parseInt(b.cells[0].querySelector('input').value);
      return aStart - bStart;
    });
    
    growthRateTable.innerHTML = '';
    rows.forEach(row => growthRateTable.appendChild(row));
    calculateNewDCF();
  }

  // ฟังก์ชันเพิ่มเหตุการณ์ใหม่
  function addNewRiskFactor() {
    const newId = riskFactorsTable.rows.length + 1;
    const row = riskFactorsTable.insertRow();
    row.innerHTML = `
      <td>${newId}</td>
      <td><input type="text" class="risk-event" value="เหตุการณ์ใหม่"></td>
      <td><input type="number" class="risk-factor" value="1.0" step="0.1"></td>
      <td><button class="btn btn-danger remove-risk-factor">ลบ</button></td>
    `;
    
    row.querySelector('.remove-risk-factor').addEventListener('click', function() {
      row.remove();
      calculateNewDCF();
    });

    row.querySelectorAll('input').forEach(input => {
      input.addEventListener('input', calculateNewDCF);
    });

    calculateNewDCF();
  }

  // ฟังก์ชันคำนวณ DCF แบบใหม่
  function calculateNewDCF() {
    try {
      // รับค่าจากฟอร์ม
      const initialInvestment = parseFloat(initialInvestmentInput.value) || 0;
      const baseCashFlow = parseFloat(baseCashFlowInput.value) || 0;
      const wacc = parseFloat(waccInput.value) / 100 || 0;
      
      // สร้างอาร์เรย์ของเหตุการณ์และปัจจัยความเสี่ยง
      const riskFactors = [];
      for (const row of riskFactorsTable.rows) {
        const factor = parseFloat(row.cells[2].querySelector('input').value);
        if (!isNaN(factor)) {
          riskFactors.push({
            id: row.cells[0].textContent,
            event: row.cells[1].querySelector('input').value,
            factor: factor
          });
        }
      }
      
      // ตรวจสอบว่ามีเหตุการณ์ความเสี่ยงหรือไม่
      if (riskFactors.length === 0) {
        alert('กรุณาเพิ่มเหตุการณ์ความเสี่ยงอย่างน้อย 1 รายการ');
        return;
      }
      
      // สร้างอาร์เรย์ของช่วงการเติบโต
      const growthRates = [];
      for (const row of growthRateTable.rows) {
        const startYear = parseInt(row.cells[0].querySelector('input').value);
        const endYear = parseInt(row.cells[1].querySelector('input').value);
        const rate = parseFloat(row.cells[2].querySelector('input').value) / 100;
        
        if (!isNaN(startYear) && !isNaN(endYear) && !isNaN(rate) && startYear > 0 && endYear >= startYear) {
          growthRates.push({
            startYear,
            endYear,
            rate
          });
        }
      }
      
      // เรียงลำดับช่วงปีตาม startYear
      growthRates.sort((a, b) => a.startYear - b.startYear);
      
      // ตรวจสอบความถูกต้องของช่วงปี
      if (!checkYearOverlaps(growthRates)) {
        alert('มีช่วงปีที่ซ้อนทับกัน กรุณาตรวจสอบอีกครั้ง');
        return;
      }
      
      // หาปีสูงสุด
      const maxYear = growthRates.length > 0 ? Math.max(...growthRates.map(g => g.endYear)) : 0;
      
      if (maxYear === 0) {
        alert('กรุณาระบุช่วงปีการเติบโต');
        return;
      }
      
      // สุ่มเหตุการณ์สำหรับแต่ละปี
      const eventsByYear = {};
      const riskFactorsLength = riskFactors.length;
      
      for (let year = 1; year <= maxYear; year++) {
        const randomIndex = riskFactorsLength > 1 ? 
          Math.floor(Math.random() * riskFactorsLength) : 0;
        eventsByYear[year] = riskFactors[randomIndex];
      }
      
      // คำนวณกระแสเงินสด
      dcfNewTable.innerHTML = '';
      const cashFlows = [-initialInvestment];
      let prevCashFlow = baseCashFlow;
      let totalPV = 0;
      
      for (let year = 1; year <= maxYear; year++) {
        // หาอัตราการเติบโตสำหรับปีนี้
        const growthRateObj = growthRates.find(g => year >= g.startYear && year <= g.endYear);
        const growthRate = growthRateObj ? growthRateObj.rate : 0;
        
        // คำนวณกระแสเงินสดพื้นฐาน
        const currentCashFlow = prevCashFlow * (1 + growthRate);
        
        // นำปัจจัยความเสี่ยงมาปรับค่า
        const event = eventsByYear[year] || { id: '-', event: 'ไม่มีข้อมูล', factor: 1 };
        const adjustedCashFlow = currentCashFlow * event.factor;
        
        // คำนวณมูลค่าปัจจุบัน
        const presentValue = adjustedCashFlow / Math.pow(1 + wacc, year);
        totalPV += presentValue;
        
        // เพิ่มแถวในตาราง
        const row = dcfNewTable.insertRow();
        row.innerHTML = `
          <td>ปี ${year}</td>
          <td>${currentCashFlow.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
          <td>${(growthRate * 100).toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}%</td>
          <td>${event.id}</td>
          <td>${event.event}</td>
          <td>${event.factor.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
          <td>${adjustedCashFlow.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
          <td>${(wacc * 100).toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}%</td>
          <td>${presentValue.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
        `;
        
        cashFlows.push(adjustedCashFlow);
        prevCashFlow = currentCashFlow;
      }
      
      // คำนวณผลลัพธ์
      const npv = totalPV - initialInvestment;
      const irr = calculateIRR(cashFlows);
      const pi = totalPV / initialInvestment;
      
      // แสดงผลลัพธ์
      npvResult.textContent = npv.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
      irrResult.textContent = (irr * 100).toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
      piResult.textContent = pi.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
      
      // คำนวณ Sensitivity Analysis
      calculateSensitivityAnalysis(cashFlows, initialInvestment);
    } catch (error) {
      console.error('เกิดข้อผิดพลาดในการคำนวณ:', error);
      alert('เกิดข้อผิดพลาดในการคำนวณ: ' + error.message);
    }
  }
  
  // ฟังก์ชันตรวจสอบการซ้อนทับของช่วงปี
  function checkYearOverlaps(growthRates) {
    if (growthRates.length <= 1) return true;
    
    growthRates.sort((a, b) => a.startYear - b.startYear);
    
    for (let i = 1; i < growthRates.length; i++) {
      if (growthRates[i].startYear <= growthRates[i-1].endYear) {
        return false;
      }
    }
    return true;
  }
  
  // ฟังก์ชันคำนวณ IRR
  function calculateIRR(cashFlows, guess = 0.1) {
    const tolerance = 0.0001;
    const maxIter = 100;
    let x0 = guess;
    
    // ตรวจสอบว่ามีกระแสเงินสดบวกหรือไม่
    if (cashFlows.slice(1).every(cf => cf <= 0)) {
      return 0; // ไม่มีผลตอบแทน
    }

    for (let i = 0; i < maxIter; i++) {
      const npvVal = npv(x0, cashFlows);
      const npvDeriv = npvDerivative(x0, cashFlows);
      
      if (Math.abs(npvDeriv) < 0.0001) break;
      
      const x1 = x0 - npvVal / npvDeriv;
      
      if (Math.abs(x1 - x0) < tolerance) return x1;
      x0 = x1;
    }
    return x0;
  }
  
  // ฟังก์ชันคำนวณ NPV
  function npv(rate, cashFlows) {
    return cashFlows.reduce((sum, cf, t) => sum + cf / Math.pow(1 + rate, t), 0);
  }
  
  // ฟังก์ชันคำนวณอนุพันธ์ของ NPV
  function npvDerivative(rate, cashFlows) {
    return cashFlows.reduce((sum, cf, t) => sum - t * cf / Math.pow(1 + rate, t + 1), 0);
  }
  
  // ฟังก์ชันคำนวณ Sensitivity Analysis
  function calculateSensitivityAnalysis(cashFlows, initialInvestment) {
    const rates = [0.06, 0.07, 0.08, 0.09, 0.10, 0.11, 0.12];
    
    rates.forEach(rate => {
      const totalPV = cashFlows.slice(1).reduce((sum, cf, index) => {
        return sum + (cf / Math.pow(1 + rate, index + 1));
      }, 0);
      
      const npv = totalPV - initialInvestment;
      const pi = totalPV / initialInvestment;
      
      // แสดงผลลัพธ์
      const rateKey = Math.round(rate * 100);
      sensitivityResults[rateKey].npv.textContent = npv.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
      sensitivityResults[rateKey].pi.textContent = pi.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    });
  }

  // เพิ่ม event listeners สำหรับ input ต่างๆ
  initialInvestmentInput.addEventListener('input', calculateNewDCF);
  baseCashFlowInput.addEventListener('input', calculateNewDCF);
  waccInput.addEventListener('input', calculateNewDCF);
  
  // เพิ่ม event listeners สำหรับตารางที่มีอยู่แล้ว
  document.querySelectorAll('#growth-rate-table input, #risk-factors-table input').forEach(input => {
    input.addEventListener('input', calculateNewDCF);
  });

  // เพิ่ม event listeners สำหรับปุ่มต่างๆ
  document.getElementById('add-growth-rate').addEventListener('click', addNewGrowthRate);
  document.getElementById('auto-arrange').addEventListener('click', autoArrangeYears);
  document.getElementById('add-risk-factor').addEventListener('click', addNewRiskFactor);
  document.getElementById('calculate-dcf-new').addEventListener('click', calculateNewDCF);

  // คำนวณครั้งแรกเมื่อโหลดหน้า
  calculateNewDCF();
});