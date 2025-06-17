document.addEventListener('DOMContentLoaded', function () {
    // ตรวจสอบว่าหน้าปัจจุบันเป็นหน้า income.html หรือไม่
    const isIncomePage = window.location.pathname.includes('income.html');
    
    // ถ้าเป็นหน้า income.html ให้โหลดเฉพาะฟังก์ชันที่เกี่ยวข้องกับ FCF
    if (isIncomePage) {
        // === FCF ===
        const fcfTableBody = document.querySelector('#fcf-table tbody');
        const addFcfRowBtn = document.getElementById('add-fcf-row');
        const totalIncomeEl = document.getElementById('total-income');
        const totalExpenseEl = document.getElementById('total-expense');
        const fcfResultEl = document.getElementById('fcf-result');
        const fcfTotalAllEl = document.getElementById('fcf-total-all');

        addFcfRowBtn.addEventListener('click', () => addFcfRow());
        addFcfRow(); // เพิ่มแถวเริ่มต้น

        // นำเข้าฟังก์ชันที่เกี่ยวข้องกับ FCF จากโค้ดเดิม
        function addFcfRow(type = 'รายรับ', desc = '', amount = 0, year = new Date().getFullYear()) {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>
                    <select class="fcf-type">
                        <option value="รายรับ" ${type === 'รายรับ' ? 'selected' : ''}>รายรับ</option>
                        <option value="รายจ่าย" ${type === 'รายจ่าย' ? 'selected' : ''}>รายจ่าย</option>
                    </select>
                </td>
                <td><input type="text" class="fcf-desc" value="${desc}"></td>
                <td><input type="number" class="fcf-amount" value="${amount}" step="0.01"></td>
                <td><input type="number" class="fcf-year" value="${year}" min="2000" max="2100"></td>
                <td><button class="remove-fcf-row">ลบ</button></td>
            `;
            fcfTableBody.appendChild(row);
            row.querySelectorAll('input, select').forEach(el => el.addEventListener('input', updateFcf));
            row.querySelector('.remove-fcf-row').addEventListener('click', () => {
                row.remove();
                updateFcf();
            });
            updateFcf();
        }

        function updateFcf() {
            let totalIncome = 0, totalExpense = 0;
            const summary = {};
            fcfTableBody.querySelectorAll('tr').forEach(row => {
                const type = row.querySelector('.fcf-type').value;
                const desc = row.querySelector('.fcf-desc').value;
                const amount = parseFloat(row.querySelector('.fcf-amount').value) || 0;
                const year = row.querySelector('.fcf-year').value || 'ไม่ระบุ';
            
                if (!summary[year]) summary[year] = { income: 0, expense: 0, incomeList: [], expenseList: [] };
            
                if (type === 'รายรับ') {
                    totalIncome += amount;
                    summary[year].income += amount;
                    summary[year].incomeList.push(`${desc} (${amount.toFixed(2)})`);
                } else {
                    totalExpense += amount;
                    summary[year].expense += amount;
                    summary[year].expenseList.push(`${desc} (${amount.toFixed(2)})`);
                }
            });
            
            const fcf = totalIncome - totalExpense;
            totalIncomeEl.textContent = totalIncome.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
totalExpenseEl.textContent = totalExpense.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
fcfResultEl.textContent = fcf.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
fcfTotalAllEl.textContent = totalFCF.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
            
            const summaryTable = document.querySelector('#fcf-summary-table tbody');
            summaryTable.innerHTML = '';
            let totalFCF = 0;
            
            Object.keys(summary).sort().forEach(year => {
                const data = summary[year];
                const yearFcf = data.income - data.expense;
                totalFCF += yearFcf;
                const row = document.createElement('tr');
                row.innerHTML = `
  <td>${year}</td>
  <td>${data.incomeList.join('<br>')}</td>
  <td class="number-format">${data.income.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
  <td>${data.expenseList.join('<br>')}</td>
  <td class="number-format">${data.expense.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
  <td class="number-format">${yearFcf.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
`;
                summaryTable.appendChild(row);
            });
            
            fcfTotalAllEl.textContent = totalFCF.toFixed(2);
        }
    } else {
        // === CAPM ===
        const rfInput = document.getElementById('rf');
        const betaInput = document.getElementById('beta');
        const rmInput = document.getElementById('rm');
        const calculateCapmBtn = document.getElementById('calculate-capm');
        const capmResult = document.getElementById('capm-result');

        // === WACC ===
        const totalValueInput = document.getElementById('total-value');
        const equityInput = document.getElementById('equity');
        const debtInput = document.getElementById('debt');
        const rdInput = document.getElementById('rd');
        const taxRateInput = document.getElementById('tax-rate');
        const calculateWaccBtn = document.getElementById('calculate-wacc');
        const waccResult = document.getElementById('wacc-result');

        // === DCF ===
        const initialInvestmentInput = document.getElementById('initial-investment');
        const baseCashFlowInput = document.getElementById('base-cash-flow');
        const growthRateTable = document.getElementById('growth-rate-table').getElementsByTagName('tbody')[0];
        const addGrowthRateBtn = document.getElementById('add-growth-rate');
        const autoArrangeBtn = document.getElementById('auto-arrange');
        const dcfTable = document.getElementById('dcf-table').getElementsByTagName('tbody')[0];
        const calculateDcfBtn = document.getElementById('calculate-dcf');
        const dcfResult = document.getElementById('dcf-result');
        const irrResult = document.getElementById('irr-result');
        const piResult = document.getElementById('pi-result');

        let capmValue = 0;
        let waccValue = 0;

        // === Event Listeners ===
        calculateCapmBtn.addEventListener('click', calculateCAPM);
        calculateWaccBtn.addEventListener('click', calculateWACC);
        calculateDcfBtn.addEventListener('click', calculateDCF);
        addGrowthRateBtn.addEventListener('click', addNewGrowthRate);
        autoArrangeBtn.addEventListener('click', autoArrangeYears);

        // Initialize
        initializeGrowthRates();
        calculateCAPM();

        // === CAPM ===
        function calculateCAPM() {
            const rf = parseFloat(rfInput.value) || 0;
            const beta = parseFloat(betaInput.value) || 0;
            const rm = parseFloat(rmInput.value) || 0;
            capmValue = rf + beta * (rm - rf);
            capmResult.textContent = capmValue.toFixed(2);
            calculateWACC();
        }

        function calculateWACC() {
    const equity = parseFloat(equityInput.value) || 0;
    const debt = parseFloat(debtInput.value) || 0;
    const totalValue = parseFloat(totalValueInput.value) || 0;
    const rd = parseFloat(rdInput.value) || 0;
    const taxRate = parseFloat(taxRateInput.value) || 0;
    
    // ใช้ค่า V ที่ผู้ใช้ป้อน หรือคำนวณจาก E + D ถ้าไม่ได้ป้อนค่า V
    const v = totalValue > 0 ? totalValue : equity + debt;
    
    const re = capmValue / 100;
    const rdDecimal = rd / 100;
    const taxRateDecimal = taxRate / 100;
    
    waccValue = (equity / v) * re + (debt / v) * rdDecimal * (1 - taxRateDecimal);
    waccResult.textContent = (waccValue * 100).toFixed(2);
    
    // อัปเดตตาราง DCF เมื่อ WACC เปลี่ยน
    generateDcfTable();
}

        // === DCF ===
        function initializeGrowthRates() {
            addGrowthRateRow(1, 5, 10.00);
            addGrowthRateRow(6, 8, 8.00);
            addGrowthRateRow(9, 10, 5.00);
        }

        function addNewGrowthRate() {
            const lastRow = growthRateTable.rows[growthRateTable.rows.length - 1];
            let newStart = 1, newEnd = 1;
            if (lastRow) {
                newStart = parseInt(lastRow.cells[1].querySelector('input').value) + 1;
                newEnd = newStart;
            }
            addGrowthRateRow(newStart, newEnd, 5.00);
        }

        function addGrowthRateRow(startYear, endYear, rate) {
            const row = growthRateTable.insertRow();
            row.innerHTML = `
                <td><input type="number" class="start-year" value="${startYear}"></td>
                <td><input type="number" class="end-year" value="${endYear}"></td>
                <td><input type="number" class="growth-rate" value="${rate}" step="0.01"></td>
                <td><button class="remove-growth-rate">ลบ</button></td>
            `;
            row.querySelector('.remove-growth-rate').addEventListener('click', () => {
                if (growthRateTable.rows.length > 1) {
                    row.remove();
                    checkYearOverlaps();
                    generateDcfTable();
                } else alert('ต้องมีอย่างน้อย 1 ช่วงปี');
            });
            row.querySelectorAll('input').forEach(input => {
                input.addEventListener('change', () => {
                    checkYearOverlaps();
                    generateDcfTable();
                });
            });
            generateDcfTable();
        }

        function checkYearOverlaps() {
            const rows = growthRateTable.rows;
            const overlaps = [];
            for (let i = 0; i < rows.length; i++) {
                const start1 = parseInt(rows[i].cells[0].querySelector('input').value);
                const end1 = parseInt(rows[i].cells[1].querySelector('input').value);
                for (let j = i + 1; j < rows.length; j++) {
                    const start2 = parseInt(rows[j].cells[0].querySelector('input').value);
                    const end2 = parseInt(rows[j].cells[1].querySelector('input').value);
                    if ((start1 <= end2 && end1 >= start2)) {
                        overlaps.push(`${i + 1} และ ${j + 1}`);
                    }
                }
            }
            if (overlaps.length > 0) {
                alert('ช่วงปีซ้อนทับ: ' + overlaps.join(', '));
                return false;
            }
            return true;
        }

        function autoArrangeYears() {
            const rows = Array.from(growthRateTable.rows);
            rows.sort((a, b) => parseInt(a.cells[0].querySelector('input').value) - parseInt(b.cells[0].querySelector('input').value));
            growthRateTable.innerHTML = '';
            rows.forEach(row => growthRateTable.appendChild(row));
            generateDcfTable();
        }

        function getGrowthRateForYear(year) {
            for (const row of growthRateTable.rows) {
                const start = parseInt(row.cells[0].querySelector('input').value);
                const end = parseInt(row.cells[1].querySelector('input').value);
                if (year >= start && year <= end) {
                    return parseFloat(row.cells[2].querySelector('input').value) || 0;
                }
            }
            return 0;
        }

        function generateDcfTable() {
            if (!checkYearOverlaps()) return;
            dcfTable.innerHTML = '';
            let maxYear = 0;
            for (const row of growthRateTable.rows) {
                maxYear = Math.max(maxYear, parseInt(row.cells[1].querySelector('input').value));
            }
            let prevCashFlow = parseFloat(baseCashFlowInput.value) || 0;
            for (let year = 1; year <= maxYear; year++) {
                const growth = getGrowthRateForYear(year);
                const cashFlow = prevCashFlow * (1 + growth / 100);
                const pv = cashFlow / Math.pow(1 + waccValue, year);
                dcfTable.insertRow().innerHTML = `
  <td>ปี ${year}</td>
  <td class="number-format">${cashFlow.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
  <td class="number-format">${(waccValue * 100).toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
  <td class="number-format">${pv.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
`;
                prevCashFlow = cashFlow;
            }
        }

        // เพิ่มฟังก์ชัน getMaxYear() ก่อนฟังก์ชัน calculateDCF()
function getMaxYear() {
    let maxYear = 0;
    for (const row of growthRateTable.rows) {
        const endYear = parseInt(row.cells[1].querySelector('input').value);
        if (!isNaN(endYear) && endYear > maxYear) {
            maxYear = endYear;
        }
    }
    return maxYear;
}

function calculateDCF() {
    try {
        // รับค่าจากฟอร์ม
        const initial = parseFloat(initialInvestmentInput.value) || 0;
        const baseCashFlow = parseFloat(baseCashFlowInput.value) || 0;
        
        // ตรวจสอบว่า WACC มีค่าที่ถูกต้อง
        if (isNaN(waccValue) || waccValue <= 0) {
            alert('กรุณาคำนวณ WACC ให้ถูกต้องก่อน');
            return;
        }
        
        const discountRate = waccValue; // ใช้ WACC ที่คำนวณแล้ว
        
        // หาปีสูงสุดจากตาราง growth rate
        const maxYear = getMaxYear();
        if (maxYear === 0) {
            alert('กรุณาระบุช่วงปีการเติบโต');
            return;
        }
        
        // คำนวณกระแสเงินสดและมูลค่าปัจจุบัน
        let prevCashFlow = baseCashFlow;
        let totalPV = 0;
        const cashFlows = [-initial];
        
        // ล้างตารางและสร้างใหม่
        dcfTable.innerHTML = '';
        
        for (let year = 1; year <= maxYear; year++) {
            const growth = getGrowthRateForYear(year) / 100;
            const cashFlow = prevCashFlow * (1 + growth);
            const pv = cashFlow / Math.pow(1 + discountRate, year);
            
            // เพิ่มแถวในตาราง
            const row = dcfTable.insertRow();
            row.innerHTML = `
                <td>ปี ${year}</td>
                <td class="number-format">${cashFlow.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                <td class="number-format">${(discountRate * 100).toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                <td class="number-format">${pv.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
            `;
            
            totalPV += pv;
            cashFlows.push(cashFlow);
            prevCashFlow = cashFlow;
        }
        
        // คำนวณผลลัพธ์
        const npv = totalPV - initial;
        const irr = calculateIRR(cashFlows);
        const pi = totalPV / initial;
        
        // แสดงผล
        dcfResult.textContent = npv.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
        irrResult.textContent = (irr * 100).toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
        piResult.textContent = pi.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2});
        
    } catch (error) {
        console.error('เกิดข้อผิดพลาดในการคำนวณ DCF:', error);
        alert('เกิดข้อผิดพลาดในการคำนวณ: ' + error.message);
    }
}

        function calculateIRR(cashFlows, guess = 0.1) {
            const tolerance = 0.00001; // เพิ่มความแม่นยำ
            const maxIter = 100;
            let x0 = guess, x1;
            
            for (let i = 0; i < maxIter; i++) {
                const f = npv(x0, cashFlows);
                const fPrime = npvDerivative(x0, cashFlows);
                
                if (Math.abs(fPrime) < 0.00001) break;
                
                x1 = x0 - f / fPrime;
                
                if (Math.abs(x1 - x0) < tolerance) return x1;
                x0 = x1;
            }
            
            return x0;
        }

        function npv(rate, cashFlows) {
            return cashFlows.reduce((acc, cf, t) => 
                acc + cf / Math.pow(1 + rate, t), 0);
        }

        function npvDerivative(rate, cashFlows) {
            return cashFlows.reduce((acc, cf, t) => 
                acc - t * cf / Math.pow(1 + rate, t + 1), 0);
        }
    }
});