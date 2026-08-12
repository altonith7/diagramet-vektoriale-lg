// ============================================
// DIAGRAMET VEKTORIALE - FILTER EDITION V7
// Kendet e rrymave TE GJITHA kundrejt U1
// ============================================

const canvas = document.getElementById('vektorCanvas');
const ctx = canvas.getContext('2d');

function setupCanvas() {
    const dpr = window.devicePixelRatio || 1;
    const rect = canvas.getBoundingClientRect();
    canvas.width = rect.width * dpr;
    canvas.height = rect.height * dpr;
    ctx.scale(dpr, dpr);
}

// ============================================
// VARIABLA GLOBALE
// ============================================
let excelData = [];
let currentRowIndex = 0;
let headers = [];
let originalHeaders = [];
let rowStatuses = [];
let rowComments = [];
let rowDiagnostics = [];
let filteredIndices = [];
let currentFilter = 'all';
let commentColIdx = -1;
let currentRowData = {};
let pdfDiagnostics = [];

// ============================================
// KONFIGURIMET E INDEKSEVE
// ============================================
const FORMAT_CONFIGS = {
    'ME_LN': {
        SERIAL: 1,
        STATION: 2,
        TIME: 3,
        I1: 5,
        I2: 6,
        I3: 7,
        I_LN: 8,
        U1: 9,
        U2: 10,
        U3: 11,
        ANG_I1: 12,     // Këndi U1-I1 (kundrejt U1)
        ANG_I2: 13,     // Këndi U1-I2 (kundrejt U1)
        ANG_I3: 14,     // Këndi U1-I3 (kundrejt U1)
        ANG_U2: 15,
        ANG_U3: 16,
        anglesRelativeToU1: true
    },
    'PA_LN': {
        SERIAL: 1,
        STATION: 2,
        TIME: 3,
        I1: 5,
        I2: 6,
        I3: 7,
        U1: 8,
        U2: 9,
        U3: 10,
        ANG_I1: 11,     // Këndi U1-I1
        ANG_I2: 12,     // Këndi U2-I2 (RELATIV) - do konvertohet ne kundrejt U1
        ANG_I3: 13,     // Këndi U3-I3 (RELATIV) - do konvertohet ne kundrejt U1
        ANG_U2: 14,
        ANG_U3: 15,
        anglesRelativeToU1: false
    }
};

let currentFormat = 'PA_LN';
let currentIDX = FORMAT_CONFIGS['PA_LN'];

// ============================================
// DETEKTO FORMATIN
// ============================================
function detectExcelFormat(headerArr) {
    if (!headerArr || headerArr.length === 0) return 'PA_LN';
    
    const hasLN = headerArr.some(h => {
        if (!h) return false;
        const headerStr = String(h).toLowerCase();
        return headerStr.includes('91.7.0') || 
               (headerStr.includes('rryma') && headerStr.includes('ln'));
    });
    
    const format = hasLN ? 'ME_LN' : 'PA_LN';
    
    console.log('=== FORMATI I DETEKTUAR ===');
    console.log('Ka kolonen "Rryma ne LN":', hasLN);
    console.log('Formati:', format);
    
    return format;
}

// ============================================
// LEXIMI I EXCEL-IT
// ============================================
document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (jsonData.length < 2) {
            alert('Excel-i është bosh ose nuk ka të dhëna!');
            return;
        }

        originalHeaders = jsonData[0].map(h => String(h));
        headers = jsonData[0].map(h => String(h).toLowerCase().trim());
        excelData = jsonData.slice(1).filter(row => row.length > 0);

        currentFormat = detectExcelFormat(originalHeaders);
        currentIDX = FORMAT_CONFIGS[currentFormat];

        commentColIdx = findColumnIndex(['coment', 'comment', 'note', 'verejtje', 'vërejtje', 'koment']);

        document.getElementById('searchBox').style.display = 'flex';
        document.getElementById('navControls').style.display = 'flex';
        document.getElementById('rowCounterWrapper').style.display = 'block';
        document.getElementById('filterPanel').style.display = 'block';

        const slider = document.getElementById('rowSlider');
        slider.max = excelData.length - 1;
        slider.value = 0;
        document.getElementById('totalRows').textContent = excelData.length;
        document.getElementById('manualRowInput').max = excelData.length;

        analizoTeGjitheRreshtat();
        
        currentFilter = 'all';
        filteredIndices = excelData.map((_, i) => i);
        
        currentRowIndex = 0;
        loadRow(0);
    };
    reader.readAsArrayBuffer(file);
});

// ============================================
// GJEJ INDEKSIN E KOLONES
// ============================================
function findColumnIndex(possibleNames) {
    for (let name of possibleNames) {
        const idx = headers.findIndex(h => h.includes(name.toLowerCase()));
        if (idx !== -1) return idx;
    }
    return -1;
}

// ============================================
// FUNKSIONE NDIHMESE
// ============================================
function parseAbsValue(val, def = 0) {
    const n = parseFloat(val);
    return isNaN(n) ? def : Math.abs(n);
}

function parseAngle(val, def) {
    if (val === undefined || val === null || val === '') return def;
    const n = parseFloat(val);
    if (isNaN(n)) return def;
    return ((n % 360) + 360) % 360;
}

// ============================================
// KTHEJ KENDET - TE GJITHA KUNDREJT U1
// ============================================
function getAnglesFromU1(row) {
    const IDX = currentIDX;
    
    const angI1_raw = parseAngle(row[IDX.ANG_I1], 30);
    const angI2_raw = parseAngle(row[IDX.ANG_I2], 30);
    const angI3_raw = parseAngle(row[IDX.ANG_I3], 30);
    const angU2 = parseAngle(row[IDX.ANG_U2], 240);
    const angU3 = parseAngle(row[IDX.ANG_U3], 120);
    
    let angI1_fromU1, angI2_fromU1, angI3_fromU1;
    
    if (IDX.anglesRelativeToU1) {
        // FORMATI 1: Kendet jane tashme kundrejt U1
        angI1_fromU1 = angI1_raw;
        angI2_fromU1 = angI2_raw;
        angI3_fromU1 = angI3_raw;
    } else {
        // FORMATI 2: Kendet jane relative (U1-I1, U2-I2, U3-I3)
        // Duhet t'i konvertojme ne kundrejt U1
        angI1_fromU1 = angI1_raw; // I1 tashme kundrejt U1
        angI2_fromU1 = ((angI2_raw + angU2) % 360 + 360) % 360; // I2 + U2 = I2 kundrejt U1
        angI3_fromU1 = ((angI3_raw + angU3) % 360 + 360) % 360; // I3 + U3 = I3 kundrejt U1
    }
    
    return {
        angI1_fromU1, angI2_fromU1, angI3_fromU1,
        angU2, angU3
    };
}

// ============================================
// LLOGARIT KENDIN RELATIV (I ndaj U te vet) per cosφ
// ============================================
function getRelativeAngle(angI_fromU1, angU_fromU1) {
    // Kthen kendin e I-se ndaj U-se se saj
    let diff = angI_fromU1 - angU_fromU1;
    // Normalizo ne [-180, 180]
    while (diff > 180) diff -= 360;
    while (diff < -180) diff += 360;
    return diff;
}

// ============================================
// ANALIZO TE GJITHE RRESHTAT
// ============================================
function analizoTeGjitheRreshtat() {
    rowStatuses = [];
    rowComments = [];
    rowDiagnostics = [];

    const IDX = currentIDX;

    for (let i = 0; i < excelData.length; i++) {
        const row = excelData[i];
        const U1 = parseAbsValue(row[IDX.U1]);
        const U2 = parseAbsValue(row[IDX.U2]);
        const U3 = parseAbsValue(row[IDX.U3]);
        const I1 = parseAbsValue(row[IDX.I1]);
        const I2 = parseAbsValue(row[IDX.I2]);
        const I3 = parseAbsValue(row[IDX.I3]);
        
        const angles = getAnglesFromU1(row);
        
        // Llogarit kendet relative per cosφ
        const phi1 = getRelativeAngle(angles.angI1_fromU1, 0); // U1 = 0°
        const phi2 = getRelativeAngle(angles.angI2_fromU1, angles.angU2);
        const phi3 = getRelativeAngle(angles.angI3_fromU1, angles.angU3);

        const analiza = klasifikoDetajuar(U1, U2, U3, I1, I2, I3, phi1, phi2, phi3);
        rowStatuses.push(analiza.status);
        rowDiagnostics.push(analiza.diagnostikim);

        if (commentColIdx !== -1) {
            rowComments.push(row[commentColIdx] || '');
        } else {
            rowComments.push('');
        }
    }

    perditesoStatistikat();
}

// ============================================
// KLASIFIKO ME DETAJE
// ============================================
function klasifikoDetajuar(U1, U2, U3, I1, I2, I3, phi1, phi2, phi3) {
    let hasError = false;
    let hasWarning = false;
    let problems = [];

    // Tensionet - kontroll me DALLIM ABSOLUT
    const maxU = Math.max(U1, U2, U3);
    const minU = Math.min(U1, U2, U3);
    const dallimiU = maxU - minU;
    
    if (maxU > 0) {
        if (dallimiU > 15.0) {
            hasError = true;
            problems.push('High Voltage Asymmetry ' + dallimiU.toFixed(2) + 'V');
        } else if (dallimiU >= 5.0 && dallimiU <= 15.0) {
            hasWarning = true;
            problems.push('Voltage Asymmetry ' + dallimiU.toFixed(2) + 'V');
        }
    }

    // Fazat
    const fazat = [
        { nr: 1, U: U1, I: I1, phi: phi1 },
        { nr: 2, U: U2, I: I2, phi: phi2 },
        { nr: 3, U: U3, I: I3, phi: phi3 }
    ];

    for (let f of fazat) {
        const absPhi = Math.abs(f.phi);
        const cosPhi = Math.cos(f.phi * Math.PI / 180);

        if (f.I >= 0.1 && absPhi > 90) {
            hasError = true;
            problems.push('CT Reversed L' + f.nr);
        } else if (f.I < 0.1) {
            hasWarning = true;
            problems.push('No Current L' + f.nr);
        } else if (cosPhi < 0.7) {
            hasError = true;
            problems.push('Low cosφ L' + f.nr + '=' + cosPhi.toFixed(2));
        } else if (cosPhi < 0.9) {
            hasWarning = true;
            problems.push('Medium cosφ L' + f.nr + '=' + cosPhi.toFixed(2));
        }
    }

    // Rrymat
    const maxI = Math.max(I1, I2, I3);
    const minI = Math.min(I1, I2, I3);
    const dallimiI = maxI - minI;
    
    if (maxI > 0.1) {
        if (dallimiI > 5.0) {
            hasError = true;
            problems.push('High Current Disbalance ' + dallimiI.toFixed(2) + 'A');
        } else if (dallimiI >= 1.0 && dallimiI <= 5.0) {
            hasWarning = true;
            problems.push('Current Disbalance ' + dallimiI.toFixed(2) + 'A');
        }
    }

    let status = 'ok';
    if (hasError) status = 'error';
    else if (hasWarning) status = 'warning';

    let diagnostikim = problems.length > 0 ? problems.join('; ') : 'OK - Gjithçka normale';

    return { status: status, diagnostikim: diagnostikim };
}

// ============================================
// PERDITESO STATISTIKAT
// ============================================
function perditesoStatistikat() {
    const total = rowStatuses.length;
    const errCount = rowStatuses.filter(s => s === 'error').length;
    const warnCount = rowStatuses.filter(s => s === 'warning').length;
    const okCount = rowStatuses.filter(s => s === 'ok').length;

    document.getElementById('statTotal').textContent = total;
    document.getElementById('statErr').textContent = errCount;
    document.getElementById('statWarn').textContent = warnCount;
    document.getElementById('statOk').textContent = okCount;
}

// ============================================
// APLIKO FILTRIN
// ============================================
function applyFilter(filter) {
    currentFilter = filter;

    document.querySelectorAll('.btn-filter').forEach(btn => {
        btn.classList.remove('active');
        if (btn.dataset.filter === filter) btn.classList.add('active');
    });

    if (filter === 'all') {
        filteredIndices = excelData.map((_, i) => i);
    } else {
        filteredIndices = [];
        for (let i = 0; i < rowStatuses.length; i++) {
            if (rowStatuses[i] === filter) filteredIndices.push(i);
        }
    }

    const info = document.getElementById('filterInfo');
    let filterName = 'Të gjitha rastet';
    if (filter === 'error') filterName = '🔴 Vetëm Errors';
    if (filter === 'warning') filterName = '🟡 Vetëm Warnings';
    if (filter === 'ok') filterName = '🟢 Vetëm OK';
    info.innerHTML = `Duke shfaqur: <strong>${filterName}</strong> (${filteredIndices.length} raste)`;

    const slider = document.getElementById('rowSlider');
    slider.max = Math.max(0, filteredIndices.length - 1);
    slider.value = 0;
    
    const totalRowsEl = document.getElementById('totalRows');
    if (totalRowsEl) {
        totalRowsEl.textContent = filteredIndices.length;
    }

    if (filteredIndices.length > 0) {
        loadRow(filteredIndices[0]);
    } else {
        alert('Nuk ka raste per kete filter!');
    }
}

// ============================================
// NGARKO NJE RRESHT
// ============================================
function loadRow(index) {
    if (index < 0 || index >= excelData.length) return;
    currentRowIndex = index;
    const row = excelData[index];
    const IDX = currentIDX;

    document.getElementById('u1_mag').value = parseAbsValue(row[IDX.U1]);
    document.getElementById('u2_mag').value = parseAbsValue(row[IDX.U2]);
    document.getElementById('u3_mag').value = parseAbsValue(row[IDX.U3]);
    document.getElementById('i1_mag').value = parseAbsValue(row[IDX.I1]);
    document.getElementById('i2_mag').value = parseAbsValue(row[IDX.I2]);
    document.getElementById('i3_mag').value = parseAbsValue(row[IDX.I3]);
    
    const angles = getAnglesFromU1(row);
    
    // Kendet TE GJITHA kundrejt U1
    document.getElementById('ang_i1').value = angles.angI1_fromU1.toFixed(1);
    document.getElementById('ang_i2').value = angles.angI2_fromU1.toFixed(1);
    document.getElementById('ang_i3').value = angles.angI3_fromU1.toFixed(1);
    document.getElementById('ang_u2').value = angles.angU2.toFixed(1);
    document.getElementById('ang_u3').value = angles.angU3.toFixed(1);

    document.getElementById('serialDisplay').textContent = 'Njehsori: ' + (row[IDX.SERIAL] || '-');

    const commentBox = document.getElementById('commentBox');
    const commentText = document.getElementById('commentText');
    if (rowComments[index] && String(rowComments[index]).trim() !== '') {
        commentBox.style.display = 'flex';
        commentText.textContent = rowComments[index];
    } else {
        commentBox.style.display = 'none';
    }

    const posInFiltered = filteredIndices.indexOf(index);
    if (posInFiltered !== -1) {
        document.getElementById('rowSlider').value = posInFiltered;
        document.getElementById('manualRowInput').value = posInFiltered + 1;
        const totalRowsEl = document.getElementById('totalRows');
        if (totalRowsEl) {
            totalRowsEl.textContent = filteredIndices.length;
        }
    } else {
        document.getElementById('manualRowInput').value = index + 1;
    }

    currentRowData = {};
    for (let i = 0; i < originalHeaders.length; i++) {
        const key = originalHeaders[i];
        const value = row[i];
        if (key && value !== undefined && value !== null && value !== '') {
            currentRowData[key] = value;
        }
    }

    vizato();
}

// ============================================
// NAVIGIMI
// ============================================
function prevRow() {
    const posInFiltered = filteredIndices.indexOf(currentRowIndex);
    if (posInFiltered > 0) {
        loadRow(filteredIndices[posInFiltered - 1]);
    } else if (posInFiltered === -1 && filteredIndices.length > 0) {
        loadRow(filteredIndices[0]);
    }
}

function nextRow() {
    const posInFiltered = filteredIndices.indexOf(currentRowIndex);
    if (posInFiltered !== -1 && posInFiltered < filteredIndices.length - 1) {
        loadRow(filteredIndices[posInFiltered + 1]);
    } else if (posInFiltered === -1 && filteredIndices.length > 0) {
        loadRow(filteredIndices[0]);
    }
}

function sliderChange() {
    const val = parseInt(document.getElementById('rowSlider').value);
    if (val >= 0 && val < filteredIndices.length) {
        loadRow(filteredIndices[val]);
    }
}

function jumpToRow() {
    const val = parseInt(document.getElementById('manualRowInput').value) - 1;
    
    if (val >= 0 && val < filteredIndices.length) {
        loadRow(filteredIndices[val]);
    } else {
        alert('Numri duhet të jetë ndërmjet 1 dhe ' + filteredIndices.length);
    }
}

function searchMeter() {
    const query = document.getElementById('searchInput').value.trim().toLowerCase();
    if (!query || excelData.length === 0) return;

    const IDX = currentIDX;
    
    for (let i = 0; i < excelData.length; i++) {
        if (String(excelData[i][IDX.SERIAL]).toLowerCase().includes(query)) {
            if (!filteredIndices.includes(i)) {
                const confirmMsg = 'Njehsori u gjet por nuk është në filtrin aktual.\nDëshironi të shfaqni të gjitha rastet?';
                if (confirm(confirmMsg)) {
                    applyFilter('all');
                    setTimeout(() => loadRow(i), 100);
                }
            } else {
                loadRow(i);
            }
            return;
        }
    }
    alert('Njehsori nuk u gjet!');
}

// ============================================
// SWAP
// ============================================
function swapKendetTensionit() {
    const u2 = document.getElementById('ang_u2');
    const u3 = document.getElementById('ang_u3');
    const temp = u2.value;
    u2.value = u3.value;
    u3.value = temp;
    vizato();
}

// ============================================
// VIZATIMI - Kendet TE GJITHA kundrejt U1
// ============================================
function vizato() {
    setupCanvas();
    const w = canvas.getBoundingClientRect().width;
    const h = canvas.getBoundingClientRect().height;
    const cx = w / 2;
    const cy = h / 2;
    const R = Math.min(w, h) / 2 - 30;

    ctx.clearRect(0, 0, w, h);
    vizatoShkallen(cx, cy, R);

    const U1 = parseFloat(document.getElementById('u1_mag').value) || 0;
    const U2 = parseFloat(document.getElementById('u2_mag').value) || 0;
    const U3 = parseFloat(document.getElementById('u3_mag').value) || 0;
    const I1 = parseFloat(document.getElementById('i1_mag').value) || 0;
    const I2 = parseFloat(document.getElementById('i2_mag').value) || 0;
    const I3 = parseFloat(document.getElementById('i3_mag').value) || 0;

    const angU1 = 0;
    const angU2 = parseFloat(document.getElementById('ang_u2').value) || 240;
    const angU3 = parseFloat(document.getElementById('ang_u3').value) || 120;

    // Kendet e rrymave jane DIREKT kundrejt U1
    const angI1 = parseFloat(document.getElementById('ang_i1').value) || 0;
    const angI2 = parseFloat(document.getElementById('ang_i2').value) || 0;
    const angI3 = parseFloat(document.getElementById('ang_i3').value) || 0;

    // Llogarit kendet relative per cosφ (I ndaj U te vet)
    const phi1 = getRelativeAngle(angI1, angU1);
    const phi2 = getRelativeAngle(angI2, angU2);
    const phi3 = getRelativeAngle(angI3, angU3);

    const gjatesiTensionit = R * 0.75;
    const gjatesiRrymes = R * 0.60;

    const U1_abs = Math.abs(U1);
    const U2_abs = Math.abs(U2);
    const U3_abs = Math.abs(U3);
    const I1_abs = Math.abs(I1);
    const I2_abs = Math.abs(I2);
    const I3_abs = Math.abs(I3);

    if (U1_abs > 0) vizatoVektor(cx, cy, gjatesiTensionit, angU1, '#2563eb', 'U1', U1_abs.toFixed(2) + 'V', false);
    if (U2_abs > 0) vizatoVektor(cx, cy, gjatesiTensionit, angU2, '#2563eb', 'U2', U2_abs.toFixed(2) + 'V', false);
    if (U3_abs > 0) vizatoVektor(cx, cy, gjatesiTensionit, angU3, '#2563eb', 'U3', U3_abs.toFixed(2) + 'V', false);

    if (I1_abs > 0.01) vizatoVektor(cx, cy, gjatesiRrymes, angI1, '#dc2626', 'I1', I1_abs.toFixed(2) + 'A', true);
    if (I2_abs > 0.01) vizatoVektor(cx, cy, gjatesiRrymes, angI2, '#dc2626', 'I2', I2_abs.toFixed(2) + 'A', true);
    if (I3_abs > 0.01) vizatoVektor(cx, cy, gjatesiRrymes, angI3, '#dc2626', 'I3', I3_abs.toFixed(2) + 'A', true);

    kalkuloFuqinë(U1_abs, U2_abs, U3_abs, I1_abs, I2_abs, I3_abs, phi1, phi2, phi3);
    diagnostikoRrjetin(U1_abs, U2_abs, U3_abs, I1_abs, I2_abs, I3_abs, phi1, phi2, phi3);
}

function vizatoShkallen(cx, cy, R) {
    ctx.beginPath();
    ctx.arc(cx, cy, R, 0, 2 * Math.PI);
    ctx.strokeStyle = '#e2e8f0';
    ctx.lineWidth = 1;
    ctx.stroke();

    ctx.font = '10px Segoe UI';
    ctx.fillStyle = '#64748b';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    for (let deg = 0; deg < 360; deg += 10) {
        const rad = (deg - 90) * Math.PI / 180;
        const x1 = cx + Math.cos(rad) * R;
        const y1 = cy + Math.sin(rad) * R;
        const x2 = cx + Math.cos(rad) * (R - (deg % 30 === 0 ? 10 : 5));
        const y2 = cy + Math.sin(rad) * (R - (deg % 30 === 0 ? 10 : 5));

        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.strokeStyle = deg % 30 === 0 ? '#94a3b8' : '#cbd5e1';
        ctx.stroke();

        const tx = cx + Math.cos(rad) * (R + 15);
        const ty = cy + Math.sin(rad) * (R + 15);
        ctx.fillStyle = deg % 30 === 0 ? '#334155' : '#94a3b8';
        ctx.font = deg % 30 === 0 ? 'bold 11px Segoe UI' : '9px Segoe UI';
        ctx.fillText(deg, tx, ty);
    }
}

function vizatoVektor(cx, cy, length, angleDeg, color, label, valueText, isCurrent) {
    if (length <= 0) return;

    const rad = (angleDeg - 90) * Math.PI / 180;
    const x2 = cx + Math.cos(rad) * length;
    const y2 = cy + Math.sin(rad) * length;

    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.lineTo(x2, y2);
    ctx.strokeStyle = color;
    ctx.lineWidth = 2.5;
    ctx.stroke();

    const arrowSize = 10;
    const arrowAngle = Math.PI / 7;
    ctx.beginPath();
    ctx.moveTo(x2, y2);
    ctx.lineTo(
        x2 - arrowSize * Math.cos(rad - arrowAngle),
        y2 - arrowSize * Math.sin(rad - arrowAngle)
    );
    ctx.lineTo(
        x2 - arrowSize * Math.cos(rad + arrowAngle),
        y2 - arrowSize * Math.sin(rad + arrowAngle)
    );
    ctx.closePath();
    ctx.fillStyle = color;
    ctx.fill();

    const labelOffset = 18;
    const lx = cx + Math.cos(rad) * (length + labelOffset);
    const ly = cy + Math.sin(rad) * (length + labelOffset);

    ctx.fillStyle = color;
    ctx.font = 'bold 12px Segoe UI';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(label + '=' + valueText, lx, ly);
}

function kalkuloFuqinë(U1, U2, U3, I1, I2, I3, phi1, phi2, phi3) {
    const rad1 = phi1 * Math.PI / 180;
    const rad2 = phi2 * Math.PI / 180;
    const rad3 = phi3 * Math.PI / 180;

    const P1 = U1 * I1 * Math.cos(rad1);
    const P2 = U2 * I2 * Math.cos(rad2);
    const P3 = U3 * I3 * Math.cos(rad3);
    const P_total = (P1 + P2 + P3) / 1000;

    const Q1 = U1 * I1 * Math.sin(rad1);
    const Q2 = U2 * I2 * Math.sin(rad2);
    const Q3 = U3 * I3 * Math.sin(rad3);
    const Q_total = (Q1 + Q2 + Q3) / 1000;

    const S_total = Math.sqrt(P_total * P_total + Q_total * Q_total);
    const PF = S_total !== 0 ? P_total / S_total : 0;

    document.getElementById('val_P').textContent = P_total.toFixed(2) + ' kW';
    document.getElementById('val_Q').textContent = Q_total.toFixed(2) + ' kvar';
    document.getElementById('val_S').textContent = S_total.toFixed(2) + ' kVA';
    document.getElementById('val_PF').textContent = PF.toFixed(3);
}

function diagnostikoRrjetin(U1, U2, U3, I1, I2, I3, phi1, phi2, phi3) {
    const raport = document.getElementById('analizaRaport');
    raport.innerHTML = '';
    pdfDiagnostics = [];

    // Tensionet
    const maxU = Math.max(U1, U2, U3);
    const minU = Math.min(U1, U2, U3);
    const dallimiU = maxU - minU;

    if (maxU > 0) {
        if (dallimiU < 5.0) {
            addAlert(raport, 'ok', '✅', 'VOLTAGE SYMMETRY', 
                'Tensionet e balancuara (dallimi: ' + dallimiU.toFixed(2) + 'V).');
        } else if (dallimiU >= 5.0 && dallimiU <= 15.0) {
            addAlert(raport, 'warn', '⚠️', 'VOLTAGE ASYMMETRY', 
                'Dallimi: ' + dallimiU.toFixed(2) + 'V (Max: ' + maxU.toFixed(2) + 'V, Min: ' + minU.toFixed(2) + 'V)');
        } else {
            addAlert(raport, 'err', '❌', 'HIGH VOLTAGE ASYMMETRY', 
                'Dallim i lartë: ' + dallimiU.toFixed(2) + 'V (Max: ' + maxU.toFixed(2) + 'V, Min: ' + minU.toFixed(2) + 'V)');
        }
    }

    kontrolloFazen(raport, 1, U1, I1, phi1);
    kontrolloFazen(raport, 2, U2, I2, phi2);
    kontrolloFazen(raport, 3, U3, I3, phi3);

    // Rrymat
    const maxI = Math.max(I1, I2, I3);
    const minI = Math.min(I1, I2, I3);
    const dallimiI = maxI - minI;
    
    if (maxI > 0.1) {
        if (dallimiI < 1.0) {
            addAlert(raport, 'ok', '✅', 'CURRENT BALANCE', 
                'Rrymat e balancuara (dallimi: ' + dallimiI.toFixed(2) + 'A).');
        } else if (dallimiI >= 1.0 && dallimiI <= 5.0) {
            addAlert(raport, 'warn', '⚠️', 'DISBALANCED CURRENT', 
                'Dallimi: ' + dallimiI.toFixed(2) + 'A (Max: ' + maxI.toFixed(2) + 'A, Min: ' + minI.toFixed(2) + 'A)');
        } else {
            addAlert(raport, 'err', '❌', 'HIGH CURRENT DISBALANCE', 
                'Dallim i lartë: ' + dallimiI.toFixed(2) + 'A (Max: ' + maxI.toFixed(2) + 'A, Min: ' + minI.toFixed(2) + 'A)');
        }
    }
}

function kontrolloFazen(raport, nr, U, I, phi) {
    const absPhi = Math.abs(phi);
    const cosPhi = Math.cos(phi * Math.PI / 180);

    if (I < 0.1) {
        addAlert(raport, 'warn', '⚠️', 'NO CURRENT (L' + nr + ')', 
            'Nuk ka rrymë në fazën ' + nr + ' (I' + nr + ' = ' + I.toFixed(2) + 'A)');
        return;
    }

    if (absPhi > 90) {
        addAlert(raport, 'err', '❌', 'CT REVERSED (L' + nr + ')', 
            'Polaritet i kundërt! Cosφ = ' + cosPhi.toFixed(2));
        return;
    }

    let karakter = '';
    if (phi > 5) karakter = '(Induktiv)';
    else if (phi < -5) karakter = '(Kapacitiv)';
    else karakter = '(Rezistiv)';

    if (cosPhi >= 0.9) {
        addAlert(raport, 'ok', '✅', 'COSφ (L' + nr + ')', 
            'Fakt. fuqisë normal: ' + cosPhi.toFixed(2) + ' ' + karakter);
    } else if (cosPhi >= 0.7) {
        addAlert(raport, 'warn', '⚠️', 'COSφ (L' + nr + ')', 
            'Fakt. fuqisë mesatar: ' + cosPhi.toFixed(2) + ' ' + karakter);
    } else {
        addAlert(raport, 'err', '❌', 'LOW COSφ (L' + nr + ')', 
            'Faktor fuqie i dobët: ' + cosPhi.toFixed(2) + ' ' + karakter);
    }
}

function addAlert(container, type, icon, title, message) {
    const div = document.createElement('div');
    div.className = 'alert alert-' + type;
    div.innerHTML = `
        <span class="alert-icon">${icon}</span>
        <div class="alert-content">
            <span class="alert-title">${title}</span>
            <span>${message}</span>
        </div>
    `;
    container.appendChild(div);
    
    pdfDiagnostics.push({
        type: type,
        icon: icon,
        title: title,
        msg: message
    });
}

// ============================================
// GJENERIMI I PDF
// ============================================
function eksportoPDF() {
    if (!window.jspdf) {
        alert("Ju lutem prisni pak ose kontrolloni internetin për të ngarkuar librarinë e PDF-së.");
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4'
    });

    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const marginLeft = 20;
    const marginRight = 190;
    const contentWidth = marginRight - marginLeft;
    const maxY = 275;

    let currentY = 20;

    function checkPage(neededSpace) {
        if (currentY + neededSpace > maxY) {
            doc.addPage();
            currentY = 20;
        }
    }

    function getVal(id) {
        const el = document.getElementById(id);
        if (!el) return '—';
        return el.value !== undefined && el.value !== '' ? el.value : (el.innerText || '—');
    }

    function getText(id) {
        const el = document.getElementById(id);
        if (!el) return '—';
        return el.innerText || el.textContent || '—';
    }

    doc.setFont("helvetica", "bold");
    doc.setFontSize(18);
    doc.setTextColor(30, 41, 59);
    doc.text("RAPORTI I ANALIZËS", pageWidth / 2, currentY, { align: "center" });
    currentY += 8;

    doc.setDrawColor(37, 99, 235);
    doc.setLineWidth(0.8);
    doc.line(60, currentY, 150, currentY);
    currentY += 8;

    let serialTxt = getText('serialDisplay');
    let serialVal = serialTxt.includes(':') ? serialTxt.split(':')[1].trim() : "Manual";

    doc.setFontSize(10);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(80, 80, 80);
    doc.text(`Njehsori: ${serialVal}`, marginLeft, currentY);
    doc.text(`Data: ${new Date().toLocaleString('sq-AL')}`, marginRight, currentY, { align: "right" });
    currentY += 4;

    if (rowStatuses[currentRowIndex]) {
        const status = rowStatuses[currentRowIndex].toUpperCase();
        let statusColor = [5, 150, 105];
        if (status === 'ERROR') statusColor = [220, 38, 38];
        else if (status === 'WARNING') statusColor = [217, 119, 6];
        
        doc.setTextColor(...statusColor);
        doc.setFont("helvetica", "bold");
        doc.text(`Status: ${status}`, marginLeft, currentY + 4);
        currentY += 4;
    }

    currentY += 4;
    doc.setDrawColor(200, 200, 200);
    doc.setLineWidth(0.3);
    doc.line(marginLeft, currentY, marginRight, currentY);
    currentY += 10;

    if (typeof currentRowData !== 'undefined' && currentRowData && Object.keys(currentRowData).length > 0) {
        checkPage(15);

        doc.setFontSize(13);
        doc.setFont("helvetica", "bold");
        doc.setTextColor(37, 99, 235);
        doc.text("1. Të dhënat gjenerale të njehsorit:", marginLeft, currentY);
        currentY += 2;

        doc.setDrawColor(37, 99, 235);
        doc.setLineWidth(0.3);
        doc.line(marginLeft, currentY, 120, currentY);
        currentY += 6;

        doc.setFontSize(9);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(50, 50, 50);

        for (let key in currentRowData) {
            if (!currentRowData.hasOwnProperty(key)) continue;

            let val = currentRowData[key];
            if (val === undefined || val === null || val === '') val = "—";

            let keyText = `${key}: `;
            let valText = `${val}`;
            let fullLine = keyText + valText;

            let splitLines = doc.splitTextToSize(fullLine, contentWidth);
            let neededHeight = splitLines.length * 4.5;

            checkPage(neededHeight + 2);

            doc.setFont("helvetica", "bold");
            doc.text(keyText, marginLeft + 5, currentY);

            let keyWidth = doc.getTextWidth(keyText);
            doc.setFont("helvetica", "normal");

            if (splitLines.length <= 1) {
                doc.text(valText, marginLeft + 5 + keyWidth, currentY);
                currentY += 5;
            } else {
                let valLines = doc.splitTextToSize(valText, contentWidth - keyWidth - 5);
                valLines.forEach((line, idx) => {
                    if (idx === 0) {
                        doc.text(line, marginLeft + 5 + keyWidth, currentY);
                    } else {
                        currentY += 4.5;
                        checkPage(5);
                        doc.text(line, marginLeft + 5 + keyWidth, currentY);
                    }
                });
                currentY += 5;
            }
        }
        currentY += 5;
    }

    checkPage(60);

    doc.setFontSize(13);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(37, 99, 235);
    doc.text("2. Vlerat Elektrike & Fuqia:", marginLeft, currentY);
    currentY += 2;
    doc.setDrawColor(37, 99, 235);
    doc.setLineWidth(0.3);
    doc.line(marginLeft, currentY, 110, currentY);
    currentY += 8;

    let u1 = getVal('u1_mag'), i1 = getVal('i1_mag'), ai1 = getVal('ang_i1');
    let u2 = getVal('u2_mag'), i2 = getVal('i2_mag'), ai2 = getVal('ang_i2'), au2 = getVal('ang_u2');
    let u3 = getVal('u3_mag'), i3 = getVal('i3_mag'), ai3 = getVal('ang_i3'), au3 = getVal('ang_u3');

    doc.setFontSize(10);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(0, 0, 0);

    doc.setFillColor(240, 245, 255);
    doc.rect(marginLeft, currentY - 4, contentWidth, 7, 'F');
    doc.setFont("helvetica", "bold");
    doc.setTextColor(30, 41, 59);
    doc.text("Faza", marginLeft + 5, currentY);
    doc.text("Tensioni (V)", marginLeft + 30, currentY);
    doc.text("Këndi U (°)", marginLeft + 70, currentY);
    doc.text("Rryma (A)", marginLeft + 105, currentY);
    doc.text("Këndi I-U1 (°)", marginLeft + 140, currentY);
    currentY += 7;

    doc.setFont("helvetica", "normal");
    doc.setTextColor(50, 50, 50);

    doc.text("L1", marginLeft + 5, currentY);
    doc.text(`${u1}`, marginLeft + 30, currentY);
    doc.text("0", marginLeft + 70, currentY);
    doc.text(`${i1}`, marginLeft + 105, currentY);
    doc.text(`${ai1}`, marginLeft + 140, currentY);
    currentY += 6;

    doc.setFillColor(248, 250, 255);
    doc.rect(marginLeft, currentY - 4, contentWidth, 6, 'F');
    doc.text("L2", marginLeft + 5, currentY);
    doc.text(`${u2}`, marginLeft + 30, currentY);
    doc.text(`${au2}`, marginLeft + 70, currentY);
    doc.text(`${i2}`, marginLeft + 105, currentY);
    doc.text(`${ai2}`, marginLeft + 140, currentY);
    currentY += 6;

    doc.text("L3", marginLeft + 5, currentY);
    doc.text(`${u3}`, marginLeft + 30, currentY);
    doc.text(`${au3}`, marginLeft + 70, currentY);
    doc.text(`${i3}`, marginLeft + 105, currentY);
    doc.text(`${ai3}`, marginLeft + 140, currentY);
    currentY += 10;

    doc.setDrawColor(220, 220, 220);
    doc.setLineWidth(0.2);
    doc.line(marginLeft, currentY - 3, marginRight, currentY - 3);

    let valP = getText('val_P');
    let valQ = getText('val_Q');
    let valS = getText('val_S');
    let valPF = getText('val_PF');

    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.setTextColor(30, 41, 59);

    doc.text("Fuqia Aktive (P):", marginLeft + 5, currentY);
    doc.setFont("helvetica", "normal");
    doc.text(valP, marginLeft + 50, currentY);

    doc.setFont("helvetica", "bold");
    doc.text("Fuqia Reaktive (Q):", marginLeft + 90, currentY);
    doc.setFont("helvetica", "normal");
    doc.text(valQ, marginLeft + 140, currentY);
    currentY += 7;

    doc.setFont("helvetica", "bold");
    doc.text("Fuqia Dukshme (S):", marginLeft + 5, currentY);
    doc.setFont("helvetica", "normal");
    doc.text(valS, marginLeft + 50, currentY);

    doc.setFont("helvetica", "bold");
    doc.text("Cos Phi (PF):", marginLeft + 90, currentY);
    doc.setFont("helvetica", "normal");
    doc.text(valPF, marginLeft + 140, currentY);
    currentY += 15;

    checkPage(20);

    doc.setFontSize(13);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(37, 99, 235);
    doc.text("3. Diagnostikimi i Rrjetit:", marginLeft, currentY);
    currentY += 2;
    doc.setDrawColor(37, 99, 235);
    doc.setLineWidth(0.3);
    doc.line(marginLeft, currentY, 105, currentY);
    currentY += 8;

    doc.setFontSize(10);

    if (typeof pdfDiagnostics === 'undefined' || !pdfDiagnostics || pdfDiagnostics.length === 0) {
        doc.setFont("helvetica", "italic");
        doc.setTextColor(120, 120, 120);
        doc.text("Nuk ka të dhëna për t'u analizuar.", marginLeft + 5, currentY);
        currentY += 10;
    } else {
        pdfDiagnostics.forEach((al, index) => {
            checkPage(20);

            let labelColor, textColor;
            if (al.type === 'err') {
                labelColor = [220, 38, 38];
                textColor = [127, 29, 29];
            } else if (al.type === 'warn') {
                labelColor = [217, 119, 6];
                textColor = [120, 53, 15];
            } else {
                labelColor = [5, 150, 105];
                textColor = [6, 78, 59];
            }

            doc.setFont("helvetica", "bold");
            doc.setTextColor(...labelColor);
            let titleText = `[${al.title || 'Info'}]`;
            doc.text(titleText, marginLeft + 5, currentY);

            let titleWidth = doc.getTextWidth(titleText) + 3;

            doc.setFont("helvetica", "normal");
            doc.setTextColor(...textColor);

            let msgStartX = marginLeft + 5 + titleWidth;
            let availableWidth = contentWidth - titleWidth - 10;

            if (availableWidth < 50) {
                currentY += 5;
                msgStartX = marginLeft + 10;
                availableWidth = contentWidth - 15;
            }

            let splitMsg = doc.splitTextToSize(al.msg || '', availableWidth);

            splitMsg.forEach((line, lineIdx) => {
                if (lineIdx === 0) {
                    doc.text(line, msgStartX, currentY);
                } else {
                    currentY += 4.5;
                    checkPage(5);
                    doc.text(line, msgStartX, currentY);
                }
            });

            currentY += 8;
        });
    }

    currentY += 5;

    let sectionNr = 4;
    if (rowComments[currentRowIndex] && String(rowComments[currentRowIndex]).trim() !== '') {
        checkPage(20);
        
        doc.setFontSize(13);
        doc.setFont("helvetica", "bold");
        doc.setTextColor(37, 99, 235);
        doc.text(`${sectionNr}. Komenti:`, marginLeft, currentY);
        currentY += 2;
        doc.setDrawColor(37, 99, 235);
        doc.setLineWidth(0.3);
        doc.line(marginLeft, currentY, 60, currentY);
        currentY += 6;

        doc.setFontSize(10);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(50, 50, 50);
        
        const commentLines = doc.splitTextToSize(String(rowComments[currentRowIndex]), contentWidth - 10);
        commentLines.forEach(line => {
            checkPage(5);
            doc.text(line, marginLeft + 5, currentY);
            currentY += 5;
        });
        currentY += 5;
        sectionNr++;
    }

    let diagramHeight = 130;
    checkPage(diagramHeight + 20);

    doc.setFontSize(13);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(37, 99, 235);
    doc.text(`${sectionNr}. Diagrama Vektoriale:`, marginLeft, currentY);
    currentY += 2;
    doc.setDrawColor(37, 99, 235);
    doc.setLineWidth(0.3);
    doc.line(marginLeft, currentY, 105, currentY);
    currentY += 5;

    try {
        const canvasEl = document.getElementById('vektorCanvas') || document.querySelector('canvas');
        if (canvasEl) {
            let canvasImg = canvasEl.toDataURL("image/png", 1.0);

            if (currentY + 125 > maxY) {
                doc.addPage();
                currentY = 20;
            }

            doc.setDrawColor(200, 200, 200);
            doc.setLineWidth(0.3);
            doc.rect(35, currentY, 140, 120);

            doc.addImage(canvasImg, 'PNG', 40, currentY + 2, 130, 116);
            currentY += 125;
        } else {
            doc.setFont("helvetica", "italic");
            doc.setTextColor(150, 150, 150);
            doc.text("Diagrama nuk u gjet.", marginLeft + 5, currentY + 10);
            currentY += 15;
        }
    } catch (e) {
        console.warn("Gabim gjate eksportimit te diagrames:", e);
    }

    let totalPages = doc.internal.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(150, 150, 150);
        doc.text(
            `Faqja ${i} nga ${totalPages} | Gjeneruar automatikisht`,
            pageWidth / 2,
            pageHeight - 10,
            { align: "center" }
        );
        doc.setDrawColor(220, 220, 220);
        doc.setLineWidth(0.2);
        doc.line(marginLeft, pageHeight - 15, marginRight, pageHeight - 15);
    }

    let fileName = `Raport_${serialVal}_${new Date().toISOString().slice(0, 10)}.pdf`;
    doc.save(fileName);
}

// ============================================
// EKSPORT NE EXCEL
// ============================================
function eksportoExcel() {
    if (excelData.length === 0) {
        alert('Nuk ka të dhëna për eksportim!');
        return;
    }

    const newHeaders = [...originalHeaders, 'Status', 'Diagnostikimi', 'Analiza_Data'];
    const newData = [newHeaders];
    
    const indicesToExport = currentFilter === 'all' ? 
        excelData.map((_, i) => i) : filteredIndices;

    for (let idx of indicesToExport) {
        const row = [...excelData[idx]];
        
        while (row.length < originalHeaders.length) {
            row.push('');
        }

        const status = rowStatuses[idx] ? rowStatuses[idx].toUpperCase() : 'OK';
        const diagnostikim = rowDiagnostics[idx] || 'OK';
        const dataAnalizes = new Date().toLocaleString('sq-AL');
        
        row.push(status);
        row.push(diagnostikim);
        row.push(dataAnalizes);

        newData.push(row);
    }

    const ws = XLSX.utils.aoa_to_sheet(newData);
    const wb = XLSX.utils.book_new();
    
    let sheetName = 'Te_Gjitha';
    if (currentFilter === 'error') sheetName = 'Errors';
    else if (currentFilter === 'warning') sheetName = 'Warnings';
    else if (currentFilter === 'ok') sheetName = 'OK';
    
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    const fileName = 'Raport_' + sheetName + '_' + Date.now() + '.xlsx';
    XLSX.writeFile(wb, fileName);
    
    alert('Excel-i u eksportua me sukses!\n' + 
          'Numri i rreshtave: ' + (newData.length - 1));
}

// ============================================
// INICIALIZIMI
// ============================================
window.addEventListener('load', vizato);
window.addEventListener('resize', vizato);
