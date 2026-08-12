// ============================================
// DIAGRAMET VEKTORIALE - V7 FINAL
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

// VARIABLA GLOBALE
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

const FORMAT_CONFIGS = {
    'ME_LN': {
        SERIAL: 1, STATION: 2, TIME: 3,
        I1: 5, I2: 6, I3: 7, I_LN: 8,
        U1: 9, U2: 10, U3: 11,
        ANG_I1: 12, ANG_I2: 13, ANG_I3: 14,
        ANG_U2: 15, ANG_U3: 16,
        anglesRelativeToU1: true
    },
    'PA_LN': {
        SERIAL: 1, STATION: 2, TIME: 3,
        I1: 5, I2: 6, I3: 7,
        U1: 8, U2: 9, U3: 10,
        ANG_I1: 11, ANG_I2: 12, ANG_I3: 13,
        ANG_U2: 14, ANG_U3: 15,
        anglesRelativeToU1: false
    }
};

let currentFormat = 'PA_LN';
let currentIDX = FORMAT_CONFIGS['PA_LN'];

function detectExcelFormat(headerArr) {
    if (!headerArr || headerArr.length === 0) return 'PA_LN';
    const hasLN = headerArr.some(h => {
        if (!h) return false;
        const s = String(h).toLowerCase();
        return s.includes('91.7.0') || (s.includes('rryma') && s.includes('ln'));
    });
    return hasLN ? 'ME_LN' : 'PA_LN';
}

document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (jsonData.length < 2) {
            alert('Excel-i është bosh!');
            return;
        }

        originalHeaders = jsonData[0].map(h => String(h));
        headers = jsonData[0].map(h => String(h).toLowerCase().trim());
        excelData = jsonData.slice(1).filter(row => row.length > 0);

        currentFormat = detectExcelFormat(originalHeaders);
        currentIDX = FORMAT_CONFIGS[currentFormat];

        commentColIdx = findColumnIndex(['coment', 'comment', 'verejtje', 'koment']);

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

function findColumnIndex(names) {
    for (let name of names) {
        const idx = headers.findIndex(h => h.includes(name.toLowerCase()));
        if (idx !== -1) return idx;
    }
    return -1;
}

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

function getAnglesFromU1(row) {
    const IDX = currentIDX;
    const angI1_raw = parseAngle(row[IDX.ANG_I1], 30);
    const angI2_raw = parseAngle(row[IDX.ANG_I2], 30);
    const angI3_raw = parseAngle(row[IDX.ANG_I3], 30);
    const angU2 = parseAngle(row[IDX.ANG_U2], 240);
    const angU3 = parseAngle(row[IDX.ANG_U3], 120);
    
    let angI1_fromU1, angI2_fromU1, angI3_fromU1;
    
    if (IDX.anglesRelativeToU1) {
        angI1_fromU1 = angI1_raw;
        angI2_fromU1 = angI2_raw;
        angI3_fromU1 = angI3_raw;
    } else {
        angI1_fromU1 = angI1_raw;
        angI2_fromU1 = ((angI2_raw + angU2) % 360 + 360) % 360;
        angI3_fromU1 = ((angI3_raw + angU3) % 360 + 360) % 360;
    }
    
    return { angI1_fromU1, angI2_fromU1, angI3_fromU1, angU2, angU3 };
}

function getRelativeAngle(angI, angU) {
    let diff = angI - angU;
    while (diff > 180) diff -= 360;
    while (diff < -180) diff += 360;
    return diff;
}

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
        const phi1 = getRelativeAngle(angles.angI1_fromU1, 0);
        const phi2 = getRelativeAngle(angles.angI2_fromU1, angles.angU2);
        const phi3 = getRelativeAngle(angles.angI3_fromU1, angles.angU3);

        const analiza = klasifikoDetajuar(U1, U2, U3, I1, I2, I3, phi1, phi2, phi3);
        rowStatuses.push(analiza.status);
        rowDiagnostics.push(analiza.diagnostikim);
        rowComments.push(commentColIdx !== -1 ? (row[commentColIdx] || '') : '');
    }
    perditesoStatistikat();
}

function klasifikoDetajuar(U1, U2, U3, I1, I2, I3, phi1, phi2, phi3) {
    let hasError = false, hasWarning = false;
    let problems = [];

    const maxU = Math.max(U1, U2, U3);
    const minU = Math.min(U1, U2, U3);
    const dallimiU = maxU - minU;
    
    if (maxU > 0) {
        if (dallimiU > 15.0) {
            hasError = true;
            problems.push('High Voltage Asymmetry ' + dallimiU.toFixed(2) + 'V');
        } else if (dallimiU >= 5.0) {
            hasWarning = true;
            problems.push('Voltage Asymmetry ' + dallimiU.toFixed(2) + 'V');
        }
    }

    const fazat = [
        { nr: 1, I: I1, phi: phi1 },
        { nr: 2, I: I2, phi: phi2 },
        { nr: 3, I: I3, phi: phi3 }
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
            problems.push('Low cosφ L' + f.nr);
        } else if (cosPhi < 0.9) {
            hasWarning = true;
            problems.push('Medium cosφ L' + f.nr);
        }
    }

    const maxI = Math.max(I1, I2, I3);
    const minI = Math.min(I1, I2, I3);
    const dallimiI = maxI - minI;
    
    if (maxI > 0.1) {
        if (dallimiI > 5.0) {
            hasError = true;
            problems.push('High Current Disbalance ' + dallimiI.toFixed(2) + 'A');
        } else if (dallimiI >= 1.0) {
            hasWarning = true;
            problems.push('Current Disbalance ' + dallimiI.toFixed(2) + 'A');
        }
    }

    let status = 'ok';
    if (hasError) status = 'error';
    else if (hasWarning) status = 'warning';

    return { 
        status: status, 
        diagnostikim: problems.length > 0 ? problems.join('; ') : 'OK'
    };
}

function perditesoStatistikat() {
    document.getElementById('statTotal').textContent = rowStatuses.length;
    document.getElementById('statErr').textContent = rowStatuses.filter(s => s === 'error').length;
    document.getElementById('statWarn').textContent = rowStatuses.filter(s => s === 'warning').length;
    document.getElementById('statOk').textContent = rowStatuses.filter(s => s === 'ok').length;
}

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

    let filterName = 'Të gjitha rastet';
    if (filter === 'error') filterName = '🔴 Vetëm Errors';
    if (filter === 'warning') filterName = '🟡 Vetëm Warnings';
    if (filter === 'ok') filterName = '🟢 Vetëm OK';
    document.getElementById('filterInfo').innerHTML = 
        `Duke shfaqur: <strong>${filterName}</strong> (${filteredIndices.length} raste)`;

    const slider = document.getElementById('rowSlider');
    slider.max = Math.max(0, filteredIndices.length - 1);
    slider.value = 0;
    document.getElementById('totalRows').textContent = filteredIndices.length;

    if (filteredIndices.length > 0) {
        loadRow(filteredIndices[0]);
    } else {
        alert('Nuk ka raste per kete filter!');
    }
}

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
    document.getElementById('ang_i1').value = angles.angI1_fromU1.toFixed(1);
    document.getElementById('ang_i2').value = angles.angI2_fromU1.toFixed(1);
    document.getElementById('ang_i3').value = angles.angI3_fromU1.toFixed(1);
    document.getElementById('ang_u2').value = angles.angU2.toFixed(1);
    document.getElementById('ang_u3').value = angles.angU3.toFixed(1);

    document.getElementById('serialDisplay').textContent = 'Njehsori: ' + (row[IDX.SERIAL] || '-');

    const commentBox = document.getElementById('commentBox');
    if (rowComments[index] && String(rowComments[index]).trim() !== '') {
        commentBox.style.display = 'flex';
        document.getElementById('commentText').textContent = rowComments[index];
    } else {
        commentBox.style.display = 'none';
    }

    const posInFiltered = filteredIndices.indexOf(index);
    if (posInFiltered !== -1) {
        document.getElementById('rowSlider').value = posInFiltered;
        document.getElementById('manualRowInput').value = posInFiltered + 1;
        document.getElementById('totalRows').textContent = filteredIndices.length;
    } else {
        document.getElementById('manualRowInput').value = index + 1;
    }

    currentRowData = {};
    for (let i = 0; i < originalHeaders.length; i++) {
        if (originalHeaders[i] && row[i] !== undefined && row[i] !== null && row[i] !== '') {
            currentRowData[originalHeaders[i]] = row[i];
        }
    }

    vizato();
}

function prevRow() {
    const pos = filteredIndices.indexOf(currentRowIndex);
    if (pos > 0) loadRow(filteredIndices[pos - 1]);
    else if (pos === -1 && filteredIndices.length > 0) loadRow(filteredIndices[0]);
}

function nextRow() {
    const pos = filteredIndices.indexOf(currentRowIndex);
    if (pos !== -1 && pos < filteredIndices.length - 1) loadRow(filteredIndices[pos + 1]);
    else if (pos === -1 && filteredIndices.length > 0) loadRow(filteredIndices[0]);
}

function sliderChange() {
    const val = parseInt(document.getElementById('rowSlider').value);
    if (val >= 0 && val < filteredIndices.length) loadRow(filteredIndices[val]);
}

function jumpToRow() {
    const val = parseInt(document.getElementById('manualRowInput').value) - 1;
    if (val >= 0 && val < filteredIndices.length) loadRow(filteredIndices[val]);
    else alert('Numri duhet të jetë ndërmjet 1 dhe ' + filteredIndices.length);
}

function searchMeter() {
    const query = document.getElementById('searchInput').value.trim().toLowerCase();
    if (!query || excelData.length === 0) return;
    const IDX = currentIDX;
    
    for (let i = 0; i < excelData.length; i++) {
        if (String(excelData[i][IDX.SERIAL]).toLowerCase().includes(query)) {
            if (!filteredIndices.includes(i)) {
                if (confirm('Njehsori u gjet por nuk është në filtrin aktual.\nDëshironi të shfaqni të gjitha rastet?')) {
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

function swapKendetTensionit() {
    const u2 = document.getElementById('ang_u2');
    const u3 = document.getElementById('ang_u3');
    const temp = u2.value;
    u2.value = u3.value;
    u3.value = temp;
    vizato();
}

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

    const angI1 = parseFloat(document.getElementById('ang_i1').value) || 0;
    const angI2 = parseFloat(document.getElementById('ang_i2').value) || 0;
    const angI3 = parseFloat(document.getElementById('ang_i3').value) || 0;

    const phi1 = getRelativeAngle(angI1, angU1);
    const phi2 = getRelativeAngle(angI2, angU2);
    const phi3 = getRelativeAngle(angI3, angU3);

    const gjatesiTensionit = R * 0.75;
    const gjatesiRrymes = R * 0.60;

    const U1a = Math.abs(U1), U2a = Math.abs(U2), U3a = Math.abs(U3);
    const I1a = Math.abs(I1), I2a = Math.abs(I2), I3a = Math.abs(I3);

    if (U1a > 0) vizatoVektor(cx, cy, gjatesiTensionit, angU1, '#2563eb', 'U1', U1a.toFixed(2) + 'V');
    if (U2a > 0) vizatoVektor(cx, cy, gjatesiTensionit, angU2, '#2563eb', 'U2', U2a.toFixed(2) + 'V');
    if (U3a > 0) vizatoVektor(cx, cy, gjatesiTensionit, angU3, '#2563eb', 'U3', U3a.toFixed(2) + 'V');

    if (I1a > 0.01) vizatoVektor(cx, cy, gjatesiRrymes, angI1, '#dc2626', 'I1', I1a.toFixed(2) + 'A');
    if (I2a > 0.01) vizatoVektor(cx, cy, gjatesiRrymes, angI2, '#dc2626', 'I2', I2a.toFixed(2) + 'A');
    if (I3a > 0.01) vizatoVektor(cx, cy, gjatesiRrymes, angI3, '#dc2626', 'I3', I3a.toFixed(2) + 'A');

    kalkuloFuqinë(U1a, U2a, U3a, I1a, I2a, I3a, phi1, phi2, phi3);
    diagnostikoRrjetin(U1a, U2a, U3a, I1a, I2a, I3a, phi1, phi2, phi3);
}

function vizatoShkallen(cx, cy, R) {
    ctx.beginPath();
    ctx.arc(cx, cy, R, 0, 2 * Math.PI);
    ctx.strokeStyle = '#e2e8f0';
    ctx.lineWidth = 1;
    ctx.stroke();

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

function vizatoVektor(cx, cy, length, angleDeg, color, label, valueText) {
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
    ctx.lineTo(x2 - arrowSize * Math.cos(rad - arrowAngle), y2 - arrowSize * Math.sin(rad - arrowAngle));
    ctx.lineTo(x2 - arrowSize * Math.cos(rad + arrowAngle), y2 - arrowSize * Math.sin(rad + arrowAngle));
    ctx.closePath();
    ctx.fillStyle = color;
    ctx.fill();

    const lx = cx + Math.cos(rad) * (length + 18);
    const ly = cy + Math.sin(rad) * (length + 18);
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

    const P_total = (U1 * I1 * Math.cos(rad1) + U2 * I2 * Math.cos(rad2) + U3 * I3 * Math.cos(rad3)) / 1000;
    const Q_total = (U1 * I1 * Math.sin(rad1) + U2 * I2 * Math.sin(rad2) + U3 * I3 * Math.sin(rad3)) / 1000;
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

    const maxU = Math.max(U1, U2, U3);
    const minU = Math.min(U1, U2, U3);
    const dallimiU = maxU - minU;

    if (maxU > 0) {
        if (dallimiU < 5.0) {
            addAlert(raport, 'ok', '✅', 'VOLTAGE SYMMETRY', 'Tensionet e balancuara (dallimi: ' + dallimiU.toFixed(2) + 'V).');
        } else if (dallimiU <= 15.0) {
            addAlert(raport, 'warn', '⚠️', 'VOLTAGE ASYMMETRY', 'Dallimi: ' + dallimiU.toFixed(2) + 'V');
        } else {
            addAlert(raport, 'err', '❌', 'HIGH VOLTAGE ASYMMETRY', 'Dallim i lartë: ' + dallimiU.toFixed(2) + 'V');
        }
    }

    kontrolloFazen(raport, 1, I1, phi1);
    kontrolloFazen(raport, 2, I2, phi2);
    kontrolloFazen(raport, 3, I3, phi3);

    const maxI = Math.max(I1, I2, I3);
    const minI = Math.min(I1, I2, I3);
    const dallimiI = maxI - minI;
    
    if (maxI > 0.1) {
        if (dallimiI < 1.0) {
            addAlert(raport, 'ok', '✅', 'CURRENT BALANCE', 'Rrymat e balancuara (dallimi: ' + dallimiI.toFixed(2) + 'A).');
        } else if (dallimiI <= 5.0) {
            addAlert(raport, 'warn', '⚠️', 'DISBALANCED CURRENT', 'Dallimi: ' + dallimiI.toFixed(2) + 'A');
        } else {
            addAlert(raport, 'err', '❌', 'HIGH CURRENT DISBALANCE', 'Dallim i lartë: ' + dallimiI.toFixed(2) + 'A');
        }
    }
}

function kontrolloFazen(raport, nr, I, phi) {
    const absPhi = Math.abs(phi);
    const cosPhi = Math.cos(phi * Math.PI / 180);

    if (I < 0.1) {
        addAlert(raport, 'warn', '⚠️', 'NO CURRENT (L' + nr + ')', 'Nuk ka rrymë në fazën ' + nr);
        return;
    }

    if (absPhi > 90) {
        addAlert(raport, 'err', '❌', 'CT REVERSED (L' + nr + ')', 'Polaritet i kundërt! Cosφ = ' + cosPhi.toFixed(2));
        return;
    }

    let karakter = phi > 5 ? '(Induktiv)' : (phi < -5 ? '(Kapacitiv)' : '(Rezistiv)');

    if (cosPhi >= 0.9) {
        addAlert(raport, 'ok', '✅', 'COSφ (L' + nr + ')', 'Fakt. fuqisë normal: ' + cosPhi.toFixed(2) + ' ' + karakter);
    } else if (cosPhi >= 0.7) {
        addAlert(raport, 'warn', '⚠️', 'COSφ (L' + nr + ')', 'Fakt. fuqisë mesatar: ' + cosPhi.toFixed(2) + ' ' + karakter);
    } else {
        addAlert(raport, 'err', '❌', 'LOW COSφ (L' + nr + ')', 'Faktor fuqie i dobët: ' + cosPhi.toFixed(2) + ' ' + karakter);
    }
}

function addAlert(container, type, icon, title, message) {
    const div = document.createElement('div');
    div.className = 'alert alert-' + type;
    div.innerHTML = `<span class="alert-icon">${icon}</span><div class="alert-content"><span class="alert-title">${title}</span><span>${message}</span></div>`;
    container.appendChild(div);
    pdfDiagnostics.push({ type, icon, title, msg: message });
}

function eksportoPDF() {
    if (!window.jspdf) {
        alert("Ju lutem prisni pak!");
        return;
    }
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    let y = 20;
    
    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.text("RAPORTI I ANALIZËS", 105, y, { align: "center" });
    y += 10;

    let serialTxt = document.getElementById('serialDisplay').textContent;
    let serialVal = serialTxt.includes(':') ? serialTxt.split(':')[1].trim() : "Manual";
    
    doc.setFontSize(10);
    doc.setFont("helvetica", "normal");
    doc.text('Njehsori: ' + serialVal, 20, y);
    doc.text('Data: ' + new Date().toLocaleString('sq-AL'), 190, y, { align: "right" });
    y += 10;

    const canvasImg = canvas.toDataURL("image/png");
    doc.addImage(canvasImg, 'PNG', 40, y, 130, 120);
    y += 130;

    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.text('Diagnostikimi:', 20, y);
    y += 7;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    
    pdfDiagnostics.forEach(al => {
        if (y > 270) { doc.addPage(); y = 20; }
        doc.text('• ' + al.title + ': ' + al.msg, 20, y);
        y += 5;
    });

    doc.save('Raport_' + serialVal + '_' + Date.now() + '.pdf');
}

function eksportoExcel() {
    if (excelData.length === 0) {
        alert('Nuk ka të dhëna!');
        return;
    }
    const newHeaders = [...originalHeaders, 'Status', 'Diagnostikimi'];
    const newData = [newHeaders];
    const indicesToExport = currentFilter === 'all' ? excelData.map((_, i) => i) : filteredIndices;

    for (let idx of indicesToExport) {
        const row = [...excelData[idx]];
        while (row.length < originalHeaders.length) row.push('');
        row.push(rowStatuses[idx] ? rowStatuses[idx].toUpperCase() : 'OK');
        row.push(rowDiagnostics[idx] || 'OK');
        newData.push(row);
    }

    const ws = XLSX.utils.aoa_to_sheet(newData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Raport');
    XLSX.writeFile(wb, 'Raport_' + Date.now() + '.xlsx');
    alert('Excel u eksportua!');
}

window.addEventListener('load', vizato);
window.addEventListener('resize', vizato);
