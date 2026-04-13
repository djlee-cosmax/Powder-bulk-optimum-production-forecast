// ============ 데이터 저장소 ============

// SAP 표준 대비 실적 데이터에서 추출
var sapRecords = [];       // 전체 레코드
var moldBulkMap = {};      // 성형물코드 → [{ bulkCode, bulkName, stdInputPerUnit, records: [...] }]
var moldNameIndex = {};    // 성형물코드 → 성형물명
var sapCount = 0;

// 환입/폐기 데이터
var returnIndex = {};      // 생산오더 → [{ bulkCode, qty, type('환입'|'폐기'), date }]
var returnCount = 0;

// ============ 유틸 ============
function parseNum(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  return parseFloat(String(val).replace(/,/g, '').replace(/%/g, '').trim()) || 0;
}

function getCol(row, names) {
  for (var i = 0; i < names.length; i++) {
    var key = Object.keys(row).find(function(k) {
      return k.replace(/\uFEFF/g, '').trim() === names[i];
    });
    if (key && row[key] !== undefined && row[key] !== '') return row[key];
  }
  return '';
}

// ============ SAP 표준 대비 실적 데이터 파싱 (CSV + XLSX 지원) ============
function processSapRow(r) {
  var prodDate = String(getCol(r, ['생산일자']) || '').trim();
  // 엑셀 시리얼 날짜 변환
  if (prodDate && !isNaN(prodDate) && Number(prodDate) > 40000) {
    var dd = new Date((Number(prodDate) - 25569) * 86400000);
    prodDate = dd.getFullYear() + '-' + ('0' + (dd.getMonth() + 1)).slice(-2) + '-' + ('0' + dd.getDate()).slice(-2);
  }
  var prodOrder = String(getCol(r, ['생산오더']) || '').trim();
  var moldCode = String(getCol(r, ['최상위자재']) || '').trim();
  var moldName = String(getCol(r, ['자재내역']) || '').trim();
  var orderQty = parseNum(getCol(r, ['오더수량']));
  var actualQty = parseNum(getCol(r, ['실적수량']));
  var bulkCode = String(getCol(r, ['구성부품']) || '').trim();
  var stdNeed = parseNum(getCol(r, ['표준소요량']));
  var actualInput = parseNum(getCol(r, ['투입소요량']));
  var damageQty = parseNum(getCol(r, ['사용파손수량']));
  var inputRateStr = String(getCol(r, ['표준대비투입율']) || '').trim();
  var inputRate = parseNum(inputRateStr);
  var workTeam = String(getCol(r, ['작업반명']) || '').trim();
  var machine = String(getCol(r, ['작업장내역']) || '').trim();
  var categoryName = String(getCol(r, ['관리유형내역']) || '').trim();

  if (!moldCode || !bulkCode || actualQty <= 0) return;

  // 벌크명: 중복 컬럼명 처리
  var bulkName = String(getCol(r, ['자재내역_1', '자재내역_2']) || '').trim();
  if (!bulkName) {
    var keys = Object.keys(r);
    var foundFirst = false;
    for (var ki = 0; ki < keys.length; ki++) {
      if (keys[ki].replace(/\uFEFF/g, '').trim() === '자재내역') {
        if (foundFirst) { bulkName = String(r[keys[ki]] || '').trim(); break; }
        foundFirst = true;
      }
    }
  }

  if (moldName && !moldNameIndex[moldCode]) {
    moldNameIndex[moldCode] = moldName;
  }

  var stdInputPerUnit = actualQty > 0 ? stdNeed / actualQty : 0;
  var actualInputPerUnit = actualQty > 0 ? actualInput / actualQty : 0;

  var record = {
    prodDate: prodDate,
    prodOrder: prodOrder,
    moldCode: moldCode,
    moldName: moldName,
    orderQty: orderQty,
    actualQty: actualQty,
    bulkCode: bulkCode,
    bulkName: bulkName,
    stdNeed: stdNeed,
    actualInput: actualInput,
    damageQty: damageQty,
    inputRate: inputRate,
    stdInputPerUnit: stdInputPerUnit,
    actualInputPerUnit: actualInputPerUnit,
    workTeam: workTeam,
    machine: machine,
    categoryName: categoryName
  };

  sapRecords.push(record);
  sapCount++;

  if (!moldBulkMap[moldCode]) moldBulkMap[moldCode] = {};
  if (!moldBulkMap[moldCode][bulkCode]) {
    moldBulkMap[moldCode][bulkCode] = {
      bulkCode: bulkCode,
      bulkName: bulkName,
      stdInputPerUnits: [],
      records: []
    };
  }
  var entry = moldBulkMap[moldCode][bulkCode];
  if (bulkName && !entry.bulkName) entry.bulkName = bulkName;
  entry.stdInputPerUnits.push(stdInputPerUnit);
  entry.records.push(record);
}

function sapUploadComplete(fileName) {
  document.getElementById('sapStatus').textContent = fileName + ' (' + sapCount.toLocaleString() + '건)';
  document.getElementById('sapStatus').classList.add('loaded');
  document.getElementById('loadingOverlay').style.display = 'none';
  updateAutocompleteData();
}

function setupSapUpload() {
  document.getElementById('sapFile').addEventListener('change', function(e) {
    var file = e.target.files[0];
    if (!file) return;
    document.getElementById('loadingOverlay').style.display = 'flex';
    document.querySelector('.loading-text').textContent = 'SAP 데이터 로딩 중...';

    sapRecords = [];
    moldBulkMap = {};
    moldNameIndex = {};
    sapCount = 0;

    var isXlsx = file.name.match(/\.xlsx?$/i);

    if (isXlsx && typeof XLSX !== 'undefined') {
      var reader = new FileReader();
      reader.onload = function(ev) {
        var wb = XLSX.read(ev.target.result, { type: 'array' });
        wb.SheetNames.forEach(function(sheetName) {
          var rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
          rows.forEach(function(r) { processSapRow(r); });
        });
        sapUploadComplete(file.name);
      };
      reader.readAsArrayBuffer(file);
    } else {
      setTimeout(function() {
        Papa.parse(file, {
          header: true,
          encoding: 'UTF-8',
          skipEmptyLines: true,
          step: function(row) { processSapRow(row.data); },
          complete: function() { sapUploadComplete(file.name); }
        });
      }, 50);
    }
  });
}

// ============ 환입/폐기 데이터 파싱 (CSV + XLSX 지원) ============
function processReturnRow(r, workTeamCode) {
  var prodOrder = String(getCol(r, ['생산오더']) || '').trim();
  var bulkCode = String(getCol(r, ['벌크코드']) || '').trim();
  var qty = parseNum(getCol(r, ['잔량(g)', '잔량', '환입량(g)', '환입량']));
  var type = String(getCol(r, ['처리']) || '').trim();
  var date = String(getCol(r, ['날짜']) || '').trim();

  // 오타 보정: "페기" → "폐기"
  if (type === '페기') type = '폐기';

  // 엑셀 시리얼 날짜 변환
  if (date && !isNaN(date) && Number(date) > 40000) {
    var d = new Date((Number(date) - 25569) * 86400000);
    date = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  }

  if (!prodOrder || qty <= 0) return;

  returnCount++;
  if (!returnIndex[prodOrder]) returnIndex[prodOrder] = [];
  returnIndex[prodOrder].push({
    bulkCode: bulkCode,
    qty: qty,
    type: type,
    date: date,
    workTeamCode: workTeamCode || ''
  });
}

function setupReturnUpload() {
  document.getElementById('returnFile').addEventListener('change', function(e) {
    var file = e.target.files[0];
    if (!file) return;
    document.getElementById('loadingOverlay').style.display = 'flex';
    document.querySelector('.loading-text').textContent = '환입/폐기 데이터 로딩 중...';

    returnIndex = {};
    returnCount = 0;

    var isXlsx = file.name.match(/\.xlsx?$/i);

    if (isXlsx && typeof XLSX !== 'undefined') {
      // XLSX 파싱
      var reader = new FileReader();
      reader.onload = function(ev) {
        var wb = XLSX.read(ev.target.result, { type: 'array' });
        // 모든 시트 순회 (화성, 평택 등)
        wb.SheetNames.forEach(function(sheetName) {
          // 시트 이름으로 작업반 코드 매핑 (화성=3002, 평택=7002)
          var workTeamCode = '';
          if (sheetName.indexOf('화성') !== -1) workTeamCode = '3002';
          else if (sheetName.indexOf('평택') !== -1) workTeamCode = '7002';
          var rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
          rows.forEach(function(r) { processReturnRow(r, workTeamCode); });
        });
        document.getElementById('returnStatus').textContent = file.name + ' (' + returnCount.toLocaleString() + '건)';
        document.getElementById('returnStatus').classList.add('loaded');
        document.getElementById('loadingOverlay').style.display = 'none';
      };
      reader.readAsArrayBuffer(file);
    } else {
      // CSV 파싱
      setTimeout(function() {
        Papa.parse(file, {
          header: true,
          encoding: 'UTF-8',
          skipEmptyLines: true,
          step: function(row) { processReturnRow(row.data); },
          complete: function() {
            document.getElementById('returnStatus').textContent = file.name + ' (' + returnCount.toLocaleString() + '건)';
            document.getElementById('returnStatus').classList.add('loaded');
            document.getElementById('loadingOverlay').style.display = 'none';
          }
        });
      }, 50);
    }
  });
}

// ============ 자동완성 ============
var acItems = []; // { code, name }

function updateAutocompleteData() {
  acItems = Object.keys(moldBulkMap).map(function(code) {
    return { code: code, name: moldNameIndex[code] || '' };
  });
}

function setupAutocomplete(input) {
  var list = input.parentElement.querySelector('.autocomplete-list');
  var activeIdx = -1;

  input.addEventListener('input', function() {
    var val = input.value.trim().toUpperCase();
    showList(val);
  });

  input.addEventListener('focus', function() {
    var val = input.value.trim().toUpperCase();
    showList(val);
  });

  input.addEventListener('keydown', function(e) {
    var items = list.querySelectorAll('.autocomplete-item');
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      activeIdx = Math.min(activeIdx + 1, items.length - 1);
      highlightItem(items);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      activeIdx = Math.max(activeIdx - 1, 0);
      highlightItem(items);
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (activeIdx >= 0 && items[activeIdx]) {
        input.value = items[activeIdx].dataset.code;
        list.classList.remove('show');
        activeIdx = -1;
      }
    } else if (e.key === 'Escape') {
      list.classList.remove('show');
    }
  });

  document.addEventListener('click', function(e) {
    if (!input.parentElement.contains(e.target)) {
      list.classList.remove('show');
    }
  });

  function showList(val) {
    activeIdx = -1;
    var filtered = acItems;
    if (val) {
      filtered = acItems.filter(function(item) {
        return item.code.toUpperCase().indexOf(val) !== -1 || item.name.toUpperCase().indexOf(val) !== -1;
      });
    }

    if (filtered.length === 0) {
      list.classList.remove('show');
      return;
    }

    var maxShow = 50;
    list.innerHTML = filtered.slice(0, maxShow).map(function(item) {
      return '<div class="autocomplete-item" data-code="' + item.code + '">' +
        '<span class="ac-code">' + item.code + '</span>' +
        '<span class="ac-name">' + item.name + '</span>' +
      '</div>';
    }).join('');

    list.querySelectorAll('.autocomplete-item').forEach(function(el) {
      el.addEventListener('mousedown', function(e) {
        e.preventDefault();
        input.value = el.dataset.code;
        list.classList.remove('show');
      });
    });

    list.classList.add('show');
  }

  function highlightItem(items) {
    items.forEach(function(el, i) {
      el.classList.toggle('active', i === activeIdx);
    });
    if (items[activeIdx]) items[activeIdx].scrollIntoView({ block: 'nearest' });
  }
}

// 초기 input에 자동완성 연결 (예측 탭 있을 때만)
var initialMoldInput = document.querySelector('.mold-code-input');
if (initialMoldInput) setupAutocomplete(initialMoldInput);

// ============ 파일 업로드 초기화 ============
setupSapUpload();
setupReturnUpload();

// ============ 숫자 콤마 포맷 ============
function formatQtyInput(e) {
  var input = e.target;
  var raw = input.value.replace(/[^0-9]/g, '');
  if (raw === '') { input.value = ''; return; }
  input.value = Number(raw).toLocaleString();
}

var initialOrderInput = document.querySelector('.order-qty-input');
if (initialOrderInput) initialOrderInput.addEventListener('input', formatQtyInput);

// ============ 품목 행 추가/삭제 ============
var addRowBtn = document.getElementById('addRowBtn');
if (addRowBtn) {
  addRowBtn.addEventListener('click', function() {
    var container = document.getElementById('inputRows');
    var index = container.children.length;
    var row = document.createElement('div');
    row.className = 'input-row';
    row.dataset.index = index;
    row.innerHTML =
      '<div class="input-group autocomplete-wrap">' +
        '<label>성형물 코드</label>' +
        '<input type="text" class="mold-code-input" placeholder="성형물 코드 입력 (2로 시작)" autocomplete="off">' +
        '<div class="autocomplete-list"></div>' +
      '</div>' +
      '<div class="input-group">' +
        '<label>성형 지시 수량</label>' +
        '<input type="text" class="order-qty-input" placeholder="수량 입력">' +
      '</div>' +
      '<button class="remove-row-btn" onclick="removeRow(this)" title="삭제">✕</button>';
    container.appendChild(row);
    row.querySelector('.order-qty-input').addEventListener('input', formatQtyInput);
    setupAutocomplete(row.querySelector('.mold-code-input'));
  });
}

function removeRow(btn) {
  var container = document.getElementById('inputRows');
  if (!container || container.children.length <= 1) return;
  btn.closest('.input-row').remove();
}

// ============ 단일 품목 예측 ============
function predictOne(moldCode, orderQty) {
  var bulkMap = moldBulkMap[moldCode];
  if (!bulkMap) {
    return [{ error: 'SAP 데이터에 해당 성형물 이력 없음', moldCode: moldCode }];
  }

  var bulkCodes = Object.keys(bulkMap);
  if (bulkCodes.length === 0) {
    return [{ error: '연결된 벌크 없음', moldCode: moldCode }];
  }

  return bulkCodes.map(function(bulkCode) {
    var entry = bulkMap[bulkCode];
    var records = entry.records;

    // 표준 투입용량: 전체 이력의 중앙값 사용 (이상치 제거)
    var stdPerUnits = entry.stdInputPerUnits.filter(function(v) { return v > 0; });
    stdPerUnits.sort(function(a, b) { return a - b; });
    var medianStdPerUnit = stdPerUnits.length > 0
      ? stdPerUnits[Math.floor(stdPerUnits.length / 2)]
      : 0;

    // 이론 필요량
    var theoryNeed = orderQty * medianStdPerUnit;

    // 최신 이력 1건으로 로스율 계산 (날짜 내림차순 정렬 후 유효한 첫 건)
    // 날짜 기준 정렬 (최신순)
    var sortedRecords = records.slice().sort(function(a, b) {
      return (b.prodDate || '').localeCompare(a.prodDate || '');
    });

    var avgLossRate = null;

    for (var i = 0; i < sortedRecords.length; i++) {
      var rec = sortedRecords[i];
      if (rec.stdNeed <= 3000) continue;

      var actualInput = rec.actualInput;
      var stdNeed = rec.stdNeed;

      // 환입/폐기 데이터로 보정
      var returnData = returnIndex[rec.prodOrder];
      if (returnData) {
        for (var j = 0; j < returnData.length; j++) {
          var rd = returnData[j];
          // 작업반 매칭 (화성=파우더성형실, 평택=파우더성형실(평택2))
          if (rd.workTeamCode && rec.workTeam) {
            var isHwaseong = rec.workTeam === '파우더성형실' || rec.workTeam.indexOf('3002') !== -1;
            var isPyeongtaek = rec.workTeam.indexOf('평택') !== -1 || rec.workTeam.indexOf('7002') !== -1;
            if (rd.workTeamCode === '3002' && !isHwaseong) continue;
            if (rd.workTeamCode === '7002' && !isPyeongtaek) continue;
          }
          if (!rd.bulkCode || rd.bulkCode === rec.bulkCode) {
            if (rd.type === '폐기') {
              // 폐기: 전산에 투입으로 잡혀있으나 실제 사용하지 않았으므로 차감
              actualInput = actualInput - rd.qty;
            }
          }
        }
      } else {
        actualInput = rec.actualInput + rec.damageQty;
      }

      var lossRate = ((actualInput - stdNeed) / stdNeed) * 100;

      // 극단적 이상치가 아닌 유효한 건이면 채택
      if (lossRate >= -50 && lossRate <= 200) {
        avgLossRate = lossRate;
        break;
      }
    }

    // 최적 제조량 = 이론필요량 × (1 + 로스율/100)
    var optimalQty = avgLossRate !== null ? Math.ceil(theoryNeed * (1 + avgLossRate / 100)) : null;

    return {
      moldCode: moldCode,
      moldName: moldNameIndex[moldCode] || '',
      bulkCode: bulkCode,
      bulkName: entry.bulkName || '',
      stdInputPerUnit: medianStdPerUnit,
      orderQty: orderQty,
      theoryNeed: theoryNeed,
      avgLossRate: avgLossRate,
      optimalQty: optimalQty,
      hasHistory: avgLossRate !== null,
      historyCount: records.length
    };
  });
}

// ============ 전체 예측 실행 ============
var predictBtn = document.getElementById('predictBtn');
if (predictBtn) predictBtn.addEventListener('click', function() {
  if (sapCount === 0) {
    alert('먼저 "표준 대비 실적 데이터"를 업로드해 주세요.');
    return;
  }

  var rows = document.querySelectorAll('#inputRows .input-row');
  var inputList = [];

  rows.forEach(function(row) {
    var moldCode = row.querySelector('.mold-code-input').value.trim();
    var orderQty = parseNum(row.querySelector('.order-qty-input').value);
    if (moldCode && orderQty) inputList.push({ moldCode: moldCode, orderQty: orderQty });
  });

  if (inputList.length === 0) {
    alert('성형물 코드와 지시 수량을 입력해 주세요.');
    return;
  }

  // 로딩 표시
  document.getElementById('loadingOverlay').style.display = 'flex';
  document.querySelector('.loading-text').textContent = '예측 중...';
  document.getElementById('resultSection').style.display = 'none';

  setTimeout(function() {
    var results = [];
    for (var i = 0; i < inputList.length; i++) {
      var predicted = predictOne(inputList[i].moldCode, inputList[i].orderQty);
      for (var j = 0; j < predicted.length; j++) {
        results.push(predicted[j]);
      }
    }

    // 로딩 숨김
    document.getElementById('loadingOverlay').style.display = 'none';

    // 결과 테이블 표시
    document.getElementById('resultSection').style.display = 'block';
    var tbody = document.getElementById('resultBody');

    var html = '';
    for (var i = 0; i < results.length; i++) {
      var r = results[i];
      if (r.error) {
        html += '<tr><td>' + (i + 1) + '</td><td>' + r.moldCode + '</td><td class="no-data" colspan="9">' + r.error + '</td></tr>';
      } else {
        html += '<tr>' +
          '<td>' + (i + 1) + '</td>' +
          '<td>' + r.moldCode + '</td>' +
          '<td>' + r.moldName + '</td>' +
          '<td>' + r.bulkCode + '</td>' +
          '<td>' + r.bulkName + '</td>' +
          '<td>' + r.stdInputPerUnit.toFixed(2) + '</td>' +
          '<td>' + r.orderQty.toLocaleString() + '</td>' +
          '<td>' + Math.round(r.theoryNeed).toLocaleString() + '</td>' +
          '<td>' + (r.avgLossRate !== null ? r.avgLossRate.toFixed(1) + '% (' + r.historyCount + '건)' : '-') + '</td>' +
          '<td>' + (function() {
            var parts = [];
            if (r.returnActualCount > 0) parts.push('환입 ' + r.returnActualCount + '건');
            if (r.returnDisposalCount > 0) parts.push('폐기 ' + r.returnDisposalCount + '건');
            return parts.length > 0 ? parts.join(' / ') : '-';
          })() + '</td>' +
          '<td class="optimal">' + (r.optimalQty !== null ? r.optimalQty.toLocaleString() : '-') + '</td>' +
        '</tr>';
      }
    }
    tbody.innerHTML = html;
  }, 100);
});
