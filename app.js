/**
 * MES/ERP Production Report Anomaly Detection Engine
 * Công cụ tự động phân tích & phát hiện lỗi dữ liệu báo công Excel
 * Thiết kế ĐỘNG 100% - Tự động học Routing quy trình công nghệ cho mọi đơn hàng / mã hàng khác nhau
 */

// State
let rawData = [];
let analyzedResults = {
  kpis: {
    totalRows: 0,
    totalDispatches: 0,
    criticalErrors: 0,
    warnings: 0,
    flawedDispatches: 0,
    integrityScore: 100
  },
  anomalies: [],
  dispatchMap: {},
  batchMap: {},
  workOrderMap: {},
  inferredRoutings: {},
  categoryCounts: {
    all: 0,
    bypass: 0,
    duplicate: 0,
    quantity: 0,
    ticket: 0
  }
};

let currentFilter = 'all';
let currentSearch = '';
let currentSeverity = 'all';

// DOM Elements
document.addEventListener('DOMContentLoaded', () => {
  initEventListeners();
});

function initEventListeners() {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const btnSelectFile = document.getElementById('btnSelectFile');
  const btnClear = document.getElementById('btnClear');
  const btnExportExcel = document.getElementById('btnExportExcel');
  const btnCopyErrors = document.getElementById('btnCopyErrors');
  const searchInput = document.getElementById('searchInput');
  const severityFilter = document.getElementById('severityFilter');
  const modalClose = document.getElementById('modalClose');
  const modalOverlay = document.getElementById('modalOverlay');

  // Drag and drop
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('drag-active');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('drag-active');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag-active');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFile(e.dataTransfer.files[0]);
    }
  });

  // File input
  btnSelectFile.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
  });

  dropzone.addEventListener('click', () => {
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFile(e.target.files[0]);
    }
  });

  // Global paste support (Ctrl+V file or text)
  document.addEventListener('paste', (e) => {
    if (e.clipboardData.files && e.clipboardData.files.length > 0) {
      handleFile(e.clipboardData.files[0]);
      showToast('Đã nhận file Excel từ Clipboard!', 'success');
    } else {
      const text = e.clipboardData.getData('text');
      if (text && text.includes('\t') && text.split('\n').length > 2) {
        handlePastedTSV(text);
      }
    }
  });

  // Clear
  if (btnClear) {
    btnClear.addEventListener('click', (e) => {
      e.stopPropagation();
      clearData();
    });
  }

  // Export
  if (btnExportExcel) {
    btnExportExcel.addEventListener('click', exportAnomaliesExcel);
  }

  if (btnCopyErrors) {
    btnCopyErrors.addEventListener('click', copyErrorReport);
  }

  // Filter tabs
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentFilter = btn.dataset.category;
      renderTable();
    });
  });

  // Search & severity
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      currentSearch = e.target.value.toLowerCase().trim();
      renderTable();
    });
  }

  if (severityFilter) {
    severityFilter.addEventListener('change', (e) => {
      currentSeverity = e.target.value;
      renderTable();
    });
  }

  // Modal
  if (modalClose) {
    modalClose.addEventListener('click', closeModal);
  }
  if (modalOverlay) {
    modalOverlay.addEventListener('click', (e) => {
      if (e.target === modalOverlay) closeModal();
    });
  }
}

// Handle File Upload
function handleFile(file) {
  if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
    showToast('Vui lòng chọn file Excel định dạng .xlsx, .xls hoặc .csv', 'danger');
    return;
  }

  showToast('Đang đọc file: ' + file.name + '...', 'info');

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const json = XLSX.utils.sheet_to_json(worksheet);

      if (!json || json.length === 0) {
        showToast('File không có dữ liệu hợp lệ!', 'danger');
        return;
      }

      updateFileInfo(file.name, json.length, (file.size / 1024).toFixed(1) + ' KB');
      processData(json);
      showToast(`Đã phân tích xong ${json.length} dòng dữ liệu!`, 'success');
    } catch (err) {
      console.error(err);
      showToast('Lỗi khi đọc file Excel: ' + err.message, 'danger');
    }
  };
  reader.readAsArrayBuffer(file);
}

// Handle Pasted TSV Text (from Excel Copy)
function handlePastedTSV(tsvText) {
  try {
    const lines = tsvText.trim().split('\n');
    const headers = lines[0].split('\t').map(h => h.trim());
    const json = [];

    for (let i = 1; i < lines.length; i++) {
      const parts = lines[i].split('\t');
      const row = {};
      headers.forEach((h, idx) => {
        row[h] = parts[idx] ? parts[idx].trim() : '';
      });
      json.push(row);
    }

    if (json.length > 0) {
      updateFileInfo('Dữ liệu dán từ Clipboard (TSV)', json.length, `${(tsvText.length / 1024).toFixed(1)} KB`);
      processData(json);
      showToast(`Đã phân tích xong ${json.length} dòng từ Clipboard!`, 'success');
    }
  } catch (err) {
    console.error(err);
    showToast('Lỗi khi phân tích dữ liệu dán: ' + err.message, 'danger');
  }
}

function updateFileInfo(name, rowCount, sizeStr) {
  document.getElementById('fileLoadedInfo').style.display = 'flex';
  document.getElementById('infoFileName').textContent = name;
  document.getElementById('infoRowCount').textContent = rowCount + ' dòng';
  document.getElementById('infoFileSize').textContent = sizeStr;
}

function clearData() {
  rawData = [];
  analyzedResults = {
    kpis: { totalRows: 0, totalDispatches: 0, criticalErrors: 0, warnings: 0, flawedDispatches: 0, integrityScore: 100 },
    anomalies: [],
    dispatchMap: {},
    batchMap: {},
    workOrderMap: {},
    inferredRoutings: {},
    categoryCounts: { all: 0, bypass: 0, duplicate: 0, quantity: 0, ticket: 0 }
  };
  document.getElementById('fileLoadedInfo').style.display = 'none';
  renderKPIs();
  renderFilterCounts();
  renderTable();
  showToast('Đã xóa dữ liệu thành công!', 'info');
}

// -------------------------------------------------------------
// DYNAMIC ROUTING INFERENCE ENGINE (HỌC QUY TRÌNH TỰ ĐỘNG)
// -------------------------------------------------------------
function buildDynamicRoutings(rows) {
  const workOrderMap = {};
  
  rows.forEach(r => {
    // Nhóm theo Đơn hàng (工单号), hoặc theo Mã loại vật tư (物料类别/物料编码) nếu không có mã GD
    const gd = (r['工单号'] || r['物料类别'] || r['物料编码'] || 'CHUNG').trim();
    if (!workOrderMap[gd]) workOrderMap[gd] = [];
    workOrderMap[gd].push(r);
  });

  const routings = {};

  Object.entries(workOrderMap).forEach(([gd, gdRows]) => {
    const nextOps = {};
    const inDegree = {};
    const allOps = new Set();
    const opFirstSeen = {};

    gdRows.forEach((r, idx) => {
      const op = (r['报工工序'] || '').trim();
      const next = (r['下一工序'] || '').trim();
      if (op) {
        allOps.add(op);
        if (opFirstSeen[op] === undefined) opFirstSeen[op] = idx;
      }
      if (next) {
        allOps.add(next);
        if (opFirstSeen[next] === undefined) opFirstSeen[next] = idx + 0.5;
      }
      if (op && next) {
        if (!nextOps[op]) nextOps[op] = new Set();
        nextOps[op].add(next);
        inDegree[next] = (inDegree[next] || 0) + 1;
        if (inDegree[op] === undefined) inDegree[op] = 0;
      }
    });

    // Các công đoạn khởi đầu (in-degree = 0), sắp xếp theo thời điểm xuất hiện đầu tiên
    const startNodes = Array.from(allOps)
      .filter(op => !inDegree[op] || inDegree[op] === 0)
      .sort((a, b) => opFirstSeen[a] - opFirstSeen[b]);

    const sequence = [];
    const visited = new Set();

    startNodes.forEach(start => {
      let curr = start;
      while (curr && !visited.has(curr)) {
        sequence.push(curr);
        visited.add(curr);
        const neighbors = nextOps[curr] ? Array.from(nextOps[curr]) : [];
        curr = neighbors.find(n => !visited.has(n));
      }
    });

    // Bổ sung các công đoạn còn lại chưa thăm
    Array.from(allOps)
      .filter(op => !visited.has(op))
      .sort((a, b) => opFirstSeen[a] - opFirstSeen[b])
      .forEach(op => {
        sequence.push(op);
        visited.add(op);
      });

    routings[gd] = sequence;
  });

  return { workOrderMap, routings };
}

// -------------------------------------------------------------
// CORE ANOMALY DETECTION ENGINE (KHÔNG HARDCODE)
// -------------------------------------------------------------
function processData(rows) {
  rawData = rows;
  const anomalies = [];
  const dispatchMap = {};
  const batchMap = {};

  // 1. Phân nhóm Lệnh phái công & Lô sản xuất
  rows.forEach((r, idx) => {
    const rowNum = idx + 2; // Header = dòng 1
    const pNo = (r['派工单号'] || '').trim();
    const itemCode = (r['物料类别'] || r['物料编码'] || 'Unknown').trim();
    const gd = (r['工单号'] || itemCode || 'CHUNG').trim();
    const batch = (r['生产批次'] || r['用料批次'] || r['上工序关联批次'] || '').trim();

    if (!dispatchMap[pNo]) {
      dispatchMap[pNo] = {
        pNo,
        itemCode,
        gd,
        rows: []
      };
    }
    const rowObj = { ...r, _rowNum: rowNum };
    dispatchMap[pNo].rows.push(rowObj);

    if (batch) {
      if (!batchMap[batch]) batchMap[batch] = [];
      batchMap[batch].push(rowObj);
    }
  });

  // 2. Học động chuỗi Routing công nghệ cho từng Đơn hàng
  const { workOrderMap, routings } = buildDynamicRoutings(rows);

  // --- RULE B: Duplicate Process Entry in Same Dispatch (Trùng lặp công đoạn - Hoàn toàn động) ---
  Object.entries(dispatchMap).forEach(([pNo, group]) => {
    if (!pNo || pNo === 'Trống') return;
    const opOccurrences = {};

    group.rows.forEach(r => {
      const op = (r['报工工序'] || '').trim();
      if (!op) return;
      if (!opOccurrences[op]) opOccurrences[op] = [];
      opOccurrences[op].push(r);
    });

    Object.entries(opOccurrences).forEach(([op, list]) => {
      if (list.length > 1) {
        const rowsStr = list.map(x => `Dòng #${x._rowNum} (${x['良品数']} cái)`).join(', ');
        anomalies.push({
          id: `D-${pNo}-${op}`,
          rowNum: list[0]._rowNum,
          pNo,
          op,
          category: 'duplicate',
          severity: 'critical',
          title: `Báo trùng công đoạn (${list.length} lần)`,
          detail: `Lệnh phái công ${pNo} bị ghi nhận công đoạn "${op}" tới ${list.length} lần: ${rowsStr}.`,
          recommendation: 'Hủy dòng báo công thừa hoặc gộp đúng sản lượng vào 1 phiếu duy nhất.'
        });
      }
    });
  });

  // --- RULE C: Dynamic Process Bypass & Truncation (Nhảy cóc / Bỏ qua công đoạn - ĐỘNG 100%) ---
  Object.entries(dispatchMap).forEach(([pNo, group]) => {
    if (!pNo || pNo === 'Trống') return;
    const gd = group.gd;
    const routing = routings[gd] || [];
    if (routing.length < 2) return;

    // Các công đoạn đã ghi nhận trong lệnh PGD này
    const recordedOps = new Set(group.rows.map(r => (r['报工工序'] || '').trim()));

    // Xác định vị trí công đoạn xa nhất đã hoàn thành trong chuỗi Routing chuẩn
    let maxStepIdx = -1;
    routing.forEach((step, idx) => {
      if (Array.from(recordedOps).some(o => o.includes(step) || step.includes(o))) {
        maxStepIdx = Math.max(maxStepIdx, idx);
      }
    });

    if (maxStepIdx > 0) {
      // Kiểm tra có bước trung gian nào trước maxStepIdx bị bỏ qua hoàn toàn không
      const skippedSteps = [];
      for (let i = 0; i < maxStepIdx; i++) {
        const expectedStep = routing[i];
        const isDone = Array.from(recordedOps).some(o => o.includes(expectedStep) || expectedStep.includes(o));
        if (!isDone) {
          skippedSteps.push(expectedStep);
        }
      }

      if (skippedSteps.length > 0) {
        anomalies.push({
          id: `BP-SKIP-${pNo}`,
          rowNum: group.rows[0]._rowNum,
          pNo,
          op: skippedSteps.join(' + '),
          category: 'bypass',
          severity: 'critical',
          title: `Nhảy cóc / Bỏ qua công đoạn (${skippedSteps.join(', ')})`,
          detail: `Lệnh ${pNo} đã làm đến bước "${routing[maxStepIdx]}" nhưng bỏ qua hoàn toàn các bước trung gian: ${skippedSteps.join(', ')}.`,
          recommendation: 'Kiểm tra xem sản phẩm có bỏ qua công đoạn trên thực tế không hoặc bổ sung dữ liệu báo công.'
        });
      }

      // Kiểm tra lệnh bị dừng dang dở giữa chừng mà không có công đoạn tiếp theo
      const isLastStep = maxStepIdx === routing.length - 1;
      const lastRecordedRow = group.rows[group.rows.length - 1];
      const hasNextOp = (lastRecordedRow['下一工序'] || '').trim();

      if (!isLastStep && !hasNextOp && maxStepIdx < routing.length - 2) {
        anomalies.push({
          id: `BP-TRUNC-${pNo}`,
          rowNum: group.rows[0]._rowNum,
          pNo,
          op: routing[maxStepIdx],
          category: 'bypass',
          severity: 'warning',
          title: `Lệnh phái công dừng lại ở bước ${maxStepIdx + 1}/${routing.length}`,
          detail: `Lệnh ${pNo} chỉ mới báo công đến bước "${routing[maxStepIdx]}", chưa hoàn tất các công đoạn còn lại (${routing.slice(maxStepIdx + 1).join(' -> ')}).`,
          recommendation: 'Kiểm tra tiến độ thực tế trên chuyền hoặc tiếp tục tạo các công đoạn tiếp theo.'
        });
      }
    }
  });

  // --- RULE E1: Step-to-Step Quantity Over-consumption / Defect Re-entry (Tiêu hao vượt mức đầu vào - ĐỘNG 100%) ---
  const overConsumedEvents = [];

  Object.entries(dispatchMap).forEach(([pNo, group]) => {
    if (!pNo || pNo === 'Trống') return;

    // Sắp xếp các dòng theo trình tự thời gian
    const sortedRows = [...group.rows].sort((a, b) => {
      const tA = new Date(a['录单日期'] || a['开工时间'] || 0).getTime();
      const tB = new Date(b['录单日期'] || b['开工时间'] || 0).getTime();
      return tA - tB;
    });

    for (let i = 0; i < sortedRows.length; i++) {
      const rPrev = sortedRows[i];
      const nextOpName = (rPrev['下一工序'] || '').trim();
      if (!nextOpName) continue;

      // Tìm dòng công đoạn kế tiếp tương ứng
      const rCurr = sortedRows.find((x, idx) => idx > i && (x['报工工序'] || '').trim().includes(nextOpName));
      if (!rCurr) continue;

      const prevGood = parseInt(rPrev['良品数'], 10) || 0;
      const currGood = parseInt(rCurr['良品数'], 10) || 0;
      const currDefect = parseInt(rCurr['不良待判数'], 10) || 0;
      const currConsumed = currGood + currDefect;

      if (currConsumed > prevGood) {
        const excess = currConsumed - prevGood;
        overConsumedEvents.push({
          pNo,
          rowNum: rCurr._rowNum,
          ticketNo: rCurr['报工单号'],
          opPrev: rPrev['报工工序'],
          opCurr: rCurr['报工工序'],
          prevGood,
          currGood,
          currDefect,
          excess,
          batch: (rCurr['生产批次'] || rPrev['生产批次'] || '').trim(),
          gd: rCurr['工单号']
        });

        anomalies.push({
          id: `Q-OVER-${pNo}-${rCurr._rowNum}`,
          rowNum: rCurr._rowNum,
          pNo,
          op: rCurr['报工工序'],
          category: 'quantity',
          severity: 'critical',
          title: `Tiêu hao vượt mức bàn giao từ công đoạn trước (+${excess} con)`,
          detail: `Công đoạn "${rCurr['报工工序']}" (Dòng #${rCurr._rowNum}, Phiếu ${rCurr['报工单号']}) báo tiêu hao ${currConsumed} con (${currGood} đạt + ${currDefect} phế), nhưng bước liền trước "${rPrev['报工工序']}" chỉ bàn giao ${prevGood} con đạt. Tiêu hao vượt quá ${excess} con khiến thâm hụt số lượng của cả lô/lệnh.`,
          recommendation: `Kiểm tra phiếu ${rCurr['报工单号']}: Nếu ${currDefect} con phế đã được báo ở công đoạn trước thì phải sửa cột Bất lương về 0, hoặc giảm Lương phẩm tương ứng để tổng tiêu hao không vượt quá ${prevGood}.`
        });
      } else if (currGood > prevGood) {
        anomalies.push({
          id: `Q-INFLATE-${pNo}-${rCurr._rowNum}`,
          rowNum: rCurr._rowNum,
          pNo,
          op: rCurr['报工工序'],
          category: 'quantity',
          severity: 'critical',
          title: `Lương phẩm tự nhiên tăng thêm so với công đoạn trước (+${currGood - prevGood} con)`,
          detail: `Bước "${rCurr['报工工序']}" báo ${currGood} con đạt, trong khi bước trước "${rPrev['报工工序']}" chỉ có ${prevGood} con đạt. Hàng hóa không thể tự tăng số lượng khi chưa có lệnh bổ sung.`,
          recommendation: `Kiểm tra lại số lượng thực tế kiểm đếm trên chuyền và điều chỉnh về đúng ${prevGood} con.`
        });
      }
    }
  });

  // --- RULE E2: Batch / Pool WIP Deficit on Next Step (Thâm hụt hạn ngạch Lô chung - ĐỘNG 100%) ---
  Object.entries(batchMap).forEach(([batch, bRows]) => {
    if (!batch || batch === 'Trống') return;

    // Tìm tất cả các cặp chuyển tiếp trong lô
    const opPairs = {};
    bRows.forEach(r => {
      const op = (r['报工工序'] || '').trim();
      const nextOp = (r['下一工序'] || '').trim();
      if (op && nextOp) {
        const key = `${op} -> ${nextOp}`;
        if (!opPairs[key]) opPairs[key] = { fromOp: op, toOp: nextOp, fromRows: [] };
        opPairs[key].fromRows.push(r);
      }
    });

    Object.values(opPairs).forEach(pair => {
      const toRows = bRows.filter(r => {
        const rOp = (r['报工工序'] || '').trim();
        return rOp.includes(pair.toOp) || pair.toOp.includes(rOp);
      });

      const sumFromGood = pair.fromRows.reduce((acc, r) => acc + (parseInt(r['良品数'], 10) || 0), 0);
      const sumToConsumed = toRows.reduce((acc, r) => acc + (parseInt(r['良品数'], 10) || 0) + (parseInt(r['不良待判数'], 10) || 0), 0);
      const available = sumFromGood - sumToConsumed;

      const doneFromPnos = new Set(pair.fromRows.map(r => (r['派工单号'] || '').trim()));
      const doneToPnos = new Set(toRows.map(r => (r['派工单号'] || '').trim()));
      const pendingPnos = Array.from(doneFromPnos).filter(p => p && !doneToPnos.has(p));

      pendingPnos.forEach(pNo => {
        const fromRow = pair.fromRows.find(r => (r['派工单号'] || '').trim() === pNo);
        const needed = parseInt(fromRow['良品数'], 10) || 0;

        if (available < needed) {
          const deficit = needed - available;
          const culprit = overConsumedEvents.find(e => e.batch === batch && (e.opCurr.includes(pair.toOp) || pair.toOp.includes(e.opCurr)));
          const culpritInfo = culprit 
            ? ` (do lệnh ${culprit.pNo} tại dòng #${culprit.rowNum} báo tiêu hao vượt ${culprit.excess} con)`
            : '';

          anomalies.push({
            id: `Q-DEFICIT-${pNo}-${pair.toOp}`,
            rowNum: fromRow._rowNum,
            pNo,
            op: `${pair.fromOp} ➔ ${pair.toOp}`,
            category: 'quantity',
            severity: 'critical',
            title: `Thiếu hụt số lượng khả dụng khi tạo bước tiếp theo (Chỉ còn ${available}/${needed} con)`,
            detail: `Lệnh ${pNo} đã hoàn thành bước "${pair.fromOp}" với ${needed} con, nhưng khi tạo bước tiếp theo "${pair.toOp}", lô chung ${batch} chỉ còn ${available} con khả dụng (thiếu ${deficit} con)${culpritInfo}.`,
            recommendation: culprit 
              ? `Điều chỉnh lại phiếu ${culprit.ticketNo} của lệnh ${culprit.pNo} (sửa Bất lương về 0 hoặc giảm Lương phẩm) để hoàn trả ${deficit} con cho lệnh ${pNo}.`
              : `Kiểm tra lại số lượng xuất/tiêu hao của các lệnh trước trong cùng lô ${batch}.`
          });
        }
      });
    });
  });

  // --- RULE F: Non-standard Ticket / Quick Dispatch Interception (Chứng từ chen ngang - ĐỘNG 100%) ---
  rows.forEach((r, idx) => {
    const rowNum = idx + 2;
    const ticketNo = (r['报工单号'] || '').trim();
    const ticketType = (r['单据类型'] || '').trim();
    const pNo = (r['派工单号'] || '').trim();
    const op = (r['报工工序'] || '').trim();

    if (!pNo || ticketType.includes('快速') || ticketNo.startsWith('KSBG')) {
      anomalies.push({
        id: `TK-${rowNum}`,
        rowNum,
        pNo: pNo || 'KHÔNG CÓ MÃ PGD',
        op,
        category: 'ticket',
        severity: 'critical',
        title: 'Chứng từ báo nhanh chen ngang (Thiếu mã phái công PGD)',
        detail: `Phiếu "${ticketNo}" (${ticketType}) ghi nhận ${r['良品数']} chiếc cho bước "${op}" nhưng không gắn mã Lệnh phái công (PGD).`,
        recommendation: 'Gán đúng mã phái công hoặc kiểm tra lý do phát sinh phiếu báo nhanh.'
      });
    }
  });

  // Compile KPIs & Category counts
  const categoryCounts = {
    all: anomalies.length,
    bypass: anomalies.filter(a => a.category === 'bypass').length,
    duplicate: anomalies.filter(a => a.category === 'duplicate').length,
    quantity: anomalies.filter(a => a.category === 'quantity').length,
    ticket: anomalies.filter(a => a.category === 'ticket').length
  };

  const criticalErrors = anomalies.filter(a => a.severity === 'critical').length;
  const warnings = anomalies.filter(a => a.severity === 'warning').length;
  const flawedDispatches = new Set(anomalies.map(a => a.pNo).filter(p => p && p !== 'Trống')).size;
  const totalDispatches = Object.keys(dispatchMap).filter(k => k && k !== 'Trống').length;

  let integrityScore = 100;
  if (rows.length > 0) {
    const penalty = (criticalErrors * 1.5) + (warnings * 0.5);
    integrityScore = Math.max(0, Math.round(100 - (penalty / rows.length * 100)));
  }

  analyzedResults = {
    kpis: {
      totalRows: rows.length,
      totalDispatches,
      criticalErrors,
      warnings,
      flawedDispatches,
      integrityScore
    },
    anomalies,
    dispatchMap,
    batchMap,
    workOrderMap,
    inferredRoutings: routings,
    categoryCounts
  };

  renderKPIs();
  renderFilterCounts();
  renderTable();
}

// -------------------------------------------------------------
// UI RENDERING
// -------------------------------------------------------------
function renderKPIs() {
  const { kpis } = analyzedResults;
  document.getElementById('kpiTotalRows').textContent = kpis.totalRows.toLocaleString();
  document.getElementById('kpiTotalDispatches').textContent = kpis.totalDispatches;
  document.getElementById('kpiCritical').textContent = kpis.criticalErrors;
  document.getElementById('kpiWarnings').textContent = kpis.warnings;
  document.getElementById('kpiFlawedDispatches').textContent = kpis.flawedDispatches;
  document.getElementById('kpiIntegrity').textContent = kpis.integrityScore + '%';

  const scoreEl = document.getElementById('kpiIntegrity');
  if (kpis.integrityScore < 70) {
    scoreEl.style.color = '#f43f5e';
  } else if (kpis.integrityScore < 90) {
    scoreEl.style.color = '#f59e0b';
  } else {
    scoreEl.style.color = '#10b981';
  }
}

function renderFilterCounts() {
  const { categoryCounts } = analyzedResults;
  Object.entries(categoryCounts).forEach(([cat, count]) => {
    const el = document.getElementById(`count-${cat}`);
    if (el) el.textContent = count;
  });
}

function renderTable() {
  const tbody = document.getElementById('anomalyTableBody');
  const emptyState = document.getElementById('emptyState');
  const tableCard = document.getElementById('tableCard');

  let filtered = analyzedResults.anomalies.filter(item => {
    // Category match
    if (currentFilter !== 'all' && item.category !== currentFilter) return false;
    // Severity match
    if (currentSeverity !== 'all' && item.severity !== currentSeverity) return false;
    // Search match
    if (currentSearch) {
      const matchPNo = (item.pNo || '').toLowerCase().includes(currentSearch);
      const matchOp = (item.op || '').toLowerCase().includes(currentSearch);
      const matchDetail = (item.detail || '').toLowerCase().includes(currentSearch);
      const matchRow = String(item.rowNum).includes(currentSearch);
      if (!matchPNo && !matchOp && !matchDetail && !matchRow) return false;
    }
    return true;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = '';
    emptyState.style.display = 'block';
    return;
  }

  emptyState.style.display = 'none';

  // Group by PGD so each PGD only appears once in the main list
  const grouped = [];
  const pMap = {};

  filtered.forEach(item => {
    const pNo = item.pNo || 'KHÔNG CÓ PGD';
    if (!pMap[pNo]) {
      pMap[pNo] = {
        pNo,
        items: [],
        severity: 'warning',
        categories: new Set(),
        rowNums: new Set(),
        ops: new Set()
      };
      grouped.push(pMap[pNo]);
    }
    const g = pMap[pNo];
    g.items.push(item);
    g.categories.add(item.category);
    if (item.rowNum) g.rowNums.add(item.rowNum);
    if (item.op) g.ops.add(item.op);
    if (item.severity === 'critical') g.severity = 'critical';
  });

  tbody.innerHTML = grouped.map((g, idx) => {
    const sevBadge = g.severity === 'critical' 
      ? '<span class="badge badge-danger">Lỗi Nghiêm trọng</span>' 
      : '<span class="badge badge-warning">Cảnh báo</span>';

    const catBadges = Array.from(g.categories).map(cat => getCategoryBadge(cat)).join(' ');
    const rowNumsList = Array.from(g.rowNums);
    const rowNumsStr = rowNumsList.length > 0 ? rowNumsList.map(r => '#' + r).join(', ') : '--';
    const opsStr = Array.from(g.ops).map(o => escapeHtml(o)).join('; ');

    let detailsHtml = '';
    if (g.items.length === 1) {
      const it = g.items[0];
      detailsHtml = `
        <div style="font-weight: 600; color: #e2e8f0; margin-bottom: 3px;">${escapeHtml(it.title)}</div>
        <div style="font-size: 12px; color: #94a3b8; line-height: 1.4;">${escapeHtml(it.detail)}</div>
        <div style="font-size: 11px; color: #38bdf8; margin-top: 4px;">💡 Khuyến nghị: ${escapeHtml(it.recommendation)}</div>
      `;
    } else {
      detailsHtml = `
        <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
          <span class="badge ${g.severity === 'critical' ? 'badge-danger' : 'badge-warning'}" style="font-size: 11px;">
            Tổng ${g.items.length} vi phạm
          </span>
        </div>
        <div style="display: flex; flex-direction: column; gap: 6px;">
          ${g.items.map((it, i) => `
            <div style="padding: 7px 10px; background: rgba(255, 255, 255, 0.03); border-radius: 6px; border-left: 3px solid ${it.severity === 'critical' ? '#f43f5e' : '#f59e0b'};">
              <div style="font-weight: 600; color: #e2e8f0; font-size: 12px; display: flex; justify-content: space-between;">
                <span>${i + 1}. ${escapeHtml(it.title)}</span>
                <span style="color: #64748b; font-weight: normal; font-size: 11px;">Dòng #${it.rowNum}</span>
              </div>
              <div style="font-size: 11.5px; color: #94a3b8; line-height: 1.35; margin-top: 2px;">${escapeHtml(it.detail)}</div>
              <div style="font-size: 11px; color: #38bdf8; margin-top: 2px;">💡 ${escapeHtml(it.recommendation)}</div>
            </div>
          `).join('')}
        </div>
      `;
    }

    return `
      <tr>
        <td style="text-align: center; font-weight: 600; color: #64748b; font-size: 13px;">${idx + 1}</td>
        <td>
          <span class="code-tag" onclick="openDispatchTimeline('${g.pNo}')" title="Bấm để xem toàn bộ routing của lệnh">${escapeHtml(g.pNo)}</span>
        </td>
        <td style="font-weight: 600; color: #94a3b8; font-size: 12px; line-height: 1.4;">${rowNumsStr}</td>
        <td>
          <div style="display: flex; flex-wrap: wrap; gap: 4px; margin-bottom: 5px;">${catBadges}</div>
          <div style="font-size: 11.5px; color: #cbd5e1; line-height: 1.3;">${opsStr}</div>
        </td>
        <td>${detailsHtml}</td>
        <td>${sevBadge}</td>
        <td style="text-align: right;">
          <button class="btn btn-secondary" style="padding: 6px 12px; font-size: 11px;" onclick="openDispatchTimeline('${g.pNo}')">
            Xem Routing
          </button>
        </td>
      </tr>
    `;
  }).join('');
}

function getCategoryBadge(cat) {
  switch (cat) {
    case 'weight': return '<span class="badge badge-danger">⚖️ Đơn trọng</span>';
    case 'bypass': return '<span class="badge badge-danger">🔀 Nhảy cóc</span>';
    case 'duplicate': return '<span class="badge badge-danger">🔁 Trùng lặp</span>';
    case 'time': return '<span class="badge badge-warning">⏱️ Ngược thời gian</span>';
    case 'quantity': return '<span class="badge badge-danger">🔢 Sai số lượng</span>';
    case 'ticket': return '<span class="badge badge-danger">⚠️ Sai chứng từ</span>';
    default: return '<span class="badge badge-info">Logic</span>';
  }
}

// -------------------------------------------------------------
// DISPATCH ROUTING TIMELINE MODAL (ĐỘNG THEO TỪNG ĐƠN HÀNG)
// -------------------------------------------------------------
function openDispatchTimeline(pNo) {
  if (!pNo || pNo === 'Trống' || pNo === 'KHÔNG CÓ MÃ PGD') {
    showToast('Lệnh này không có mã phái công hợp lệ để xem routing!', 'warning');
    return;
  }

  const group = analyzedResults.dispatchMap[pNo];
  if (!group) {
    showToast(`Không tìm thấy dữ liệu cho mã ${pNo}`, 'warning');
    return;
  }

  const modal = document.getElementById('modalOverlay');
  document.getElementById('modalTitle').textContent = `Chi tiết Routing Lệnh Phái Công: ${pNo} (Đơn hàng: ${group.gd || 'Chung'})`;

  // Lấy chuỗi Routing động đã được học riêng cho đơn hàng này
  const gd = group.gd;
  let standardSteps = analyzedResults.inferredRoutings[gd] ? [...analyzedResults.inferredRoutings[gd]] : [];

  // Nếu không có routing động, lấy tất cả các bước xuất hiện trong group
  if (standardSteps.length === 0) {
    standardSteps = Array.from(new Set(group.rows.map(r => (r['报工工序'] || '').trim()).filter(Boolean)));
  }

  // Đảm bảo không sót bất kỳ công đoạn nào của PGD này
  group.rows.forEach(r => {
    const op = (r['报工工序'] || '').trim();
    if (op && !standardSteps.some(s => s.includes(op) || op.includes(s))) {
      standardSteps.push(op);
    }
  });

  const recordedMap = {};
  group.rows.forEach(r => {
    const op = (r['报工工序'] || '').trim();
    if (!recordedMap[op]) recordedMap[op] = [];
    recordedMap[op].push(r);
  });

  // Kiểm tra lệnh này có đang bị lỗi kẹt hạn ngạch số lượng không
  const deficitAnomalies = analyzedResults.anomalies.filter(a => a.pNo === pNo && a.id.startsWith('Q-DEFICIT'));

  const timelineContainer = document.getElementById('timelineContainer');
  timelineContainer.innerHTML = standardSteps.map((stepName, stepIdx) => {
    // Tìm bản ghi tương ứng
    const matchKey = Object.keys(recordedMap).find(k => k.includes(stepName) || stepName.includes(k));
    const records = matchKey ? recordedMap[matchKey] : null;

    let nodeClass = 'node-ok';
    let nodeIcon = '✓';
    let statusText = '<span style="color: #34d399; font-weight: 600;">Đã hoàn thành</span>';
    let detailsHtml = '';

    if (!records || records.length === 0) {
      // Kiểm tra xem bước này có đang bị thâm hụt số lượng không
      const defAnomaly = deficitAnomalies.find(a => a.op.includes(stepName) || stepName.includes(a.op));

      if (defAnomaly) {
        nodeClass = 'node-warning';
        nodeIcon = '!';
        statusText = `<span style="color: #f43f5e; font-weight: 700;">⚠️ KẸT HẠN NGẠCH (Thiếu số lượng khả dụng)</span>`;
        detailsHtml = `
          <div style="padding: 8px 12px; background: rgba(244, 63, 94, 0.1); border-left: 3px solid #f43f5e; border-radius: 6px; margin-top: 5px;">
            <div style="font-weight: 600; color: #fca5a5; font-size: 12px;">${escapeHtml(defAnomaly.title)}</div>
            <div style="font-size: 11.5px; color: #cbd5e1; margin-top: 2px;">${escapeHtml(defAnomaly.detail)}</div>
            <div style="font-size: 11px; color: #38bdf8; margin-top: 4px;">💡 ${escapeHtml(defAnomaly.recommendation)}</div>
          </div>
        `;
      } else {
        nodeClass = 'node-bypass';
        nodeIcon = '✗';
        statusText = '<span style="color: #fb7185; font-weight: 700;">CHƯA THỰC HIỆN / BỎ QUA</span>';
        detailsHtml = `<span style="color: #fda4af;">Lệnh này chưa có bản ghi báo công cho công đoạn này.</span>`;
      }
    } else {
      const rec = records[0];
      const isDuplicate = records.length > 1;
      const weight = rec['单重'] ? rec['单重'] + ' g' : '--';
      const goodQty = rec['良品数'] !== undefined ? rec['良品数'] + ' con đạt' : '--';
      const defectQty = parseInt(rec['不良待判数'], 10) || 0;
      const dateStr = rec['录单日期'] ? formatDate(rec['录单日期']) : '--';

      if (isDuplicate) {
        nodeClass = 'node-warning';
        nodeIcon = '!';
        statusText = `<span style="color: #fbbf24; font-weight: 700;">Trùng lặp ${records.length} lần!</span>`;
      }

      detailsHtml = `
        <div class="timeline-details">
          <span>📦 Lương phẩm: <strong>${goodQty}</strong></span>
          ${defectQty > 0 ? `<span style="color: #f43f5e;">❌ Bất lương: <strong>${defectQty} con phế</strong></span>` : ''}
          <span>⚖️ Đơn trọng: <strong>${weight}</strong></span>
          <span>📅 Ngày nhập: <strong>${dateStr}</strong></span>
          <span>👤 Người nhập: <strong>${rec['创建人'] || '--'}</strong></span>
          <span>📄 Phiếu: <strong>${rec['报工单号'] || '--'}</strong> (Dòng #${records.map(x => x._rowNum).join(', #')})</span>
        </div>
      `;
    }

    return `
      <div class="timeline-item">
        <div class="timeline-node ${nodeClass}">${nodeIcon}</div>
        <div class="timeline-content">
          <div class="timeline-title-row">
            <span class="timeline-title">Bước ${stepIdx + 1}: ${escapeHtml(stepName)}</span>
            ${statusText}
          </div>
          ${detailsHtml}
        </div>
      </div>
    `;
  }).join('');

  modal.classList.add('active');
}

function closeModal() {
  document.getElementById('modalOverlay').classList.remove('active');
}

// -------------------------------------------------------------
// EXPORTING AND UTILITIES
// -------------------------------------------------------------
function exportAnomaliesExcel() {
  if (analyzedResults.anomalies.length === 0) {
    showToast('Không có dữ liệu lỗi để xuất!', 'warning');
    return;
  }

  // Group by PGD so each row in Excel is 1 PGD
  const pMap = {};
  const grouped = [];
  analyzedResults.anomalies.forEach(item => {
    const pNo = item.pNo || 'KHÔNG CÓ PGD';
    if (!pMap[pNo]) {
      pMap[pNo] = { pNo, items: [], severity: 'warning', rowNums: new Set(), ops: new Set() };
      grouped.push(pMap[pNo]);
    }
    pMap[pNo].items.push(item);
    if (item.rowNum) pMap[pNo].rowNums.add(item.rowNum);
    if (item.op) pMap[pNo].ops.add(item.op);
    if (item.severity === 'critical') pMap[pNo].severity = 'critical';
  });

  const exportData = grouped.map((g, idx) => ({
    'STT': idx + 1,
    'Mã lệnh phái công (PGD)': g.pNo,
    'Dòng Excel liên quan': Array.from(g.rowNums).map(r => '#' + r).join(', '),
    'Số lượng vi phạm': g.items.length,
    'Mức độ': g.severity === 'critical' ? 'LỖI NGHIÊM TRỌNG' : 'CẢNH BÁO',
    'Các công đoạn liên quan': Array.from(g.ops).join('; '),
    'Chi tiết các vi phạm': g.items.map((x, i) => `${i + 1}. [Dòng #${x.rowNum}] ${x.title}: ${x.detail}`).join(' | '),
    'Khuyến nghị khắc phục': g.items.map(x => x.recommendation).filter(Boolean).join(' | ')
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Loi_Theo_PGD');

  XLSX.writeFile(workbook, `Bao_Cao_Loi_Theo_PGD_${new Date().toISOString().slice(0,10)}.xlsx`);
  showToast('Đã xuất file Excel theo PGD thành công!', 'success');
}

function copyErrorReport() {
  if (analyzedResults.anomalies.length === 0) {
    showToast('Không có dữ liệu để sao chép!', 'warning');
    return;
  }

  const summary = [
    `=== BÁO CÁO TỔNG HỢP LỖI BÁO CÔNG MES/ERP ===`,
    `Tổng dòng kiểm tra: ${analyzedResults.kpis.totalRows}`,
    `Tổng số lệnh phái công: ${analyzedResults.kpis.totalDispatches}`,
    `Lỗi nghiêm trọng: ${analyzedResults.kpis.criticalErrors}`,
    `Cảnh báo logic: ${analyzedResults.kpis.warnings}`,
    `Điểm tin cậy dữ liệu: ${analyzedResults.kpis.integrityScore}%`,
    ``,
    `--- CÁC LỖI ĐIỂN HÌNH THEO MÃ ---`
  ];

  analyzedResults.anomalies.forEach((a, i) => {
    summary.push(`${i + 1}. [${a.pNo}] [Dòng #${a.rowNum}] ${a.title}: ${a.detail}`);
  });

  navigator.clipboard.writeText(summary.join('\n')).then(() => {
    showToast('Đã sao chép báo cáo vào Clipboard!', 'success');
  }).catch(() => {
    showToast('Không thể sao chép tự động, vui lòng thử lại!', 'danger');
  });
}

function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast`;
  if (type === 'danger') toast.style.borderLeftColor = '#f43f5e';
  if (type === 'success') toast.style.borderLeftColor = '#10b981';
  if (type === 'warning') toast.style.borderLeftColor = '#f59e0b';

  toast.innerHTML = `<span>${escapeHtml(message)}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

function formatDate(isoStr) {
  if (!isoStr) return '--';
  try {
    const d = new Date(isoStr);
    if (isNaN(d.getTime())) return isoStr;
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    const hours = String(d.getHours()).padStart(2, '0');
    const mins = String(d.getMinutes()).padStart(2, '0');
    return `${day}/${month}/${year} ${hours}:${mins}`;
  } catch (e) {
    return isoStr;
  }
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
