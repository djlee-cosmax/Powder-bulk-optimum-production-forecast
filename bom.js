// ============ ML 예측 데이터 (ml_predictions.js에서 로드) ============
var mlPredictions = (typeof ML_PREDICTIONS !== 'undefined') ? ML_PREDICTIONS : null;

// ============ 탭 전환 ============
function switchTab(tabName) {
  document.querySelectorAll('.tab-btn').forEach(function(btn) {
    btn.classList.toggle('active', btn.textContent.indexOf('예측') !== -1);
  });
  document.querySelectorAll('.tab-content').forEach(function(el) {
    el.classList.remove('active');
  });
  document.getElementById('tab-' + tabName).classList.add('active');
}

// ============ BOM 파일 업로드 ============
var bomFileName = '';
document.getElementById('bomFile').addEventListener('change', function(e) {
  var file = e.target.files[0];
  if (!file) return;
  bomFileName = file.name;

  document.getElementById('loadingOverlay').style.display = 'flex';
  document.querySelector('.loading-text').textContent = 'BOM 데이터 로딩 중...';

  var reader = new FileReader();

  if (file.name.endsWith('.txt')) {
    // SAP TXT 파일: EUC-KR 인코딩 시도 → 실패 시 UTF-8
    reader.onload = function(ev) {
      var text = ev.target.result;
      parseBomTxt(text);
      document.getElementById('loadingOverlay').style.display = 'none';
    };
    reader.readAsText(file, 'euc-kr');
  } else if (file.name.endsWith('.csv')) {
    reader.onload = function(ev) {
      var text = ev.target.result;
      parseBomCsv(text);
      document.getElementById('loadingOverlay').style.display = 'none';
    };
    reader.readAsText(file, 'euc-kr');
  } else {
    // XLSX: SheetJS 필요 — 없으면 안내
    alert('XLSX 파일은 TXT 또는 CSV로 변환 후 업로드해 주세요.\nSAP_BOM조회.vbs를 사용하면 자동으로 TXT로 저장됩니다.');
    document.getElementById('loadingOverlay').style.display = 'none';
  }
});

// ============ SAP TXT 파싱 (탭 구분) ============
var prefilledOrderQtys = {}; // VBS에서 입력된 발주 수량

function parseBomTxt(text) {
  // ##ORDER_QTY## 섹션 파싱
  prefilledOrderQtys = {};
  var orderIdx = text.indexOf('##ORDER_QTY##');
  if (orderIdx !== -1) {
    var orderLines = text.substring(orderIdx).split('\n');
    for (var q = 1; q < orderLines.length; q++) {
      var parts = orderLines[q].trim().split('|');
      if (parts.length === 2 && parts[0] && parts[1]) {
        prefilledOrderQtys[parts[0].trim()] = parts[1].trim();
      }
    }
    text = text.substring(0, orderIdx);
  }

  var lines = text.split('\n');
  var bomData = [];
  var headerFound = false;
  var colMap = {};

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var cols = line.split('\t');

    // 빈 줄 스킵
    if (cols.length < 5) continue;

    // 헤더행 찾기: 'Lev' 포함
    if (!headerFound) {
      for (var h = 0; h < cols.length; h++) {
        var colName = cols[h].trim();
        if (colName === 'Lev') colMap.lev = h;
        if (colName === 'MTyp' || colName === '자재 유형') colMap.mtype = h;
        if (colName.indexOf('자재') !== -1 && !colMap.code && colName.indexOf('내역') === -1 && colName.indexOf('상태') === -1 && colName.indexOf('유형') === -1 && colName.indexOf('그룹') === -1 && colName.indexOf('코드') === -1) colMap.code = h;
        if (colName.indexOf('자재내역') !== -1 || colName.indexOf('자재 내역') !== -1) colMap.name = h;
        if (colName.indexOf('투입수량') !== -1) colMap.inputQty = h;
        if (colName.indexOf('소요수량') !== -1) colMap.needQty = h;
        if (colName === 'BUn' || colName.indexOf('단위') !== -1) colMap.unit = h;
        if (colName.indexOf('합계') !== -1 || colName.indexOf('재고합계') !== -1) colMap.stockTotal = h;
        if (colName.indexOf('가용') !== -1) colMap.available = h;
        if (colName.indexOf('품질') !== -1 && colName.indexOf('검사') !== -1) colMap.qualityInsp = h;
        if (colName === '설명1' || colName === '설명 1') colMap.desc1 = h;
        if (colName === '성명') colMap.fullName = h;
        else if (colName.indexOf('긴') !== -1 && colName.indexOf('이름') !== -1) colMap.fullName = h;
        else if (colName.indexOf('전체') !== -1 && colName.indexOf('이름') !== -1) colMap.fullName = h;
        if (colName.indexOf('로스율') !== -1 || colName.indexOf('로스') !== -1) colMap.lossRate = h;
        if (colName.indexOf('이름 1') !== -1 || colName.indexOf('공급업체') !== -1) colMap.supplier = h;
        // 추가 헤더 자동 감지
        if (colName === '구매담당자' || colName === '구매 담당자') colMap.purchaser = h;
        if (colName === '영업담당자' || colName === '영업 담당자') colMap.salesPerson = h;
        if (colName.indexOf('최종입고') !== -1 || colName.indexOf('최종 입고') !== -1) colMap.lastIn = h;
        if (colName.indexOf('최종출고') !== -1 || colName.indexOf('최종 출고') !== -1) colMap.lastOut = h;
        if (colName.indexOf('출고할당') !== -1 || colName.indexOf('출고 할당') !== -1) colMap.releaseAlloc = h;
        if (colName.indexOf('사용불가') !== -1 || colName.indexOf('사용 불가') !== -1) colMap.unusable = h;
        if (colName.indexOf('고객') !== -1 && colName.indexOf('이름') !== -1) colMap.customerName = h;
      }
      // 디버그: 감지된 컬럼 매핑 + 모든 헤더명 콘솔 출력
      console.log('[BOM 헤더 매핑]', colMap);
      console.log('[BOM 전체 헤더 목록]');
      for (var hh = 0; hh < cols.length; hh++) {
        if (cols[hh] && cols[hh].trim()) console.log('  [' + hh + '] ' + cols[hh].trim());
      }
      // 매핑된 컬럼 → 헤더명 확인
      console.log('[매핑 결과]');
      var mapKeys = Object.keys(colMap);
      for (var mk = 0; mk < mapKeys.length; mk++) {
        var key = mapKeys[mk];
        var idx = colMap[key];
        console.log('  ' + key + ' (' + idx + ') → "' + (cols[idx] ? cols[idx].trim() : '[fallback - 헤더 없음]') + '"');
      }
      if (colMap.lev !== undefined) {
        headerFound = true;

        // 자재코드: MTyp 다음 컬럼
        if (!colMap.code) colMap.code = (colMap.mtype || 1) + 1;
        if (!colMap.name) colMap.name = colMap.code + 1;
        // 투입수량 없으면 소요수량으로 대체
        if (colMap.inputQty === undefined && colMap.needQty !== undefined) {
          colMap.inputQty = colMap.needQty;
          console.log('[BOM] 투입수량 미감지 → 소요수량으로 대체');
        }

        // TXT 한글 깨짐 대비: BOM 조회 XLSX 기준 고정 인덱스 fallback
        // 헤더순서: Lev(1), MTyp(2), 자재(3), 자재내역(4), 자재상태(5), 사급여부(6), 투코드(7), 이름1(8), 투입수량(9), 소요수량(10), BUn(11), 재고합계(12), 가용(13), 품질검사(14), ...설명1(21), 전체이름(22)
        if (colMap.qualityInsp === undefined) colMap.qualityInsp = colMap.lev + 13;
        if (colMap.lastIn === undefined) colMap.lastIn = colMap.lev + 17;
        if (colMap.releaseAlloc === undefined) colMap.releaseAlloc = colMap.lev + 14;
        if (colMap.unusable === undefined) colMap.unusable = colMap.lev + 15;
        if (colMap.lastOut === undefined) colMap.lastOut = colMap.lev + 18;
        if (colMap.customerName === undefined) colMap.customerName = colMap.lev + 19;
        if (colMap.desc1 === undefined) colMap.desc1 = colMap.lev + 20;
        if (colMap.fullName === undefined) colMap.fullName = colMap.lev + 21;

        // 필수 컬럼 매핑 검증
        var requiredCols = [
          { key: 'lev', label: 'Lev (레벨)' },
          { key: 'mtype', label: 'MTyp (자재 유형)' },
          { key: 'code', label: '자재 (자재코드)' },
          { key: 'name', label: '자재내역 (제품명)' },
          { key: 'inputQty', label: '투입수량 또는 소요수량' }
        ];
        var warnCols = [
          { key: 'stockTotal', label: '재고합계' },
          { key: 'available', label: '가용 (가용재고)' },
          { key: 'unit', label: 'BUn (단위)' }
        ];
        var missingRequired = [];
        var missingWarn = [];
        for (var rq = 0; rq < requiredCols.length; rq++) {
          if (colMap[requiredCols[rq].key] === undefined) missingRequired.push(requiredCols[rq].label);
        }
        for (var wq = 0; wq < warnCols.length; wq++) {
          if (colMap[warnCols[wq].key] === undefined) missingWarn.push(warnCols[wq].label);
        }
        if (missingRequired.length > 0) {
          alert('⚠️ BOM 파일에서 필수 컬럼을 찾을 수 없습니다.\n\n' +
            '누락된 필수 컬럼:\n• ' + missingRequired.join('\n• ') +
            '\n\nSAP BOM 조회 레이아웃을 확인해 주세요.\n' +
            '필수 컬럼이 포함된 레이아웃으로 다시 조회해 주세요.');
        } else if (missingWarn.length > 0) {
          console.warn('[BOM 경고] 일부 컬럼 미감지 (fallback 적용됨): ' + missingWarn.join(', '));
        }
      }
      continue;
    }

    // 데이터행 파싱
    var lev = cols[colMap.lev] ? cols[colMap.lev].trim() : '';
    if (lev === '' || isNaN(parseInt(lev))) continue;

    var mtype = cols[colMap.mtype] ? cols[colMap.mtype].trim() : '';
    var code = cols[colMap.code] ? cols[colMap.code].trim() : '';
    var name = cols[colMap.name] ? cols[colMap.name].trim() : '';
    var inputQty = cols[colMap.inputQty] ? cols[colMap.inputQty].trim().replace(/,/g, '').replace(/\s/g, '') : '';
    var needQty = colMap.needQty !== undefined && cols[colMap.needQty] ? cols[colMap.needQty].trim().replace(/,/g, '').replace(/\s/g, '') : '';
    var unit = colMap.unit !== undefined && cols[colMap.unit] ? cols[colMap.unit].trim() : '';
    var stockTotal = colMap.stockTotal !== undefined && cols[colMap.stockTotal] ? cols[colMap.stockTotal].trim().replace(/,/g, '').replace(/\s/g, '') : '';
    var available = colMap.available !== undefined && cols[colMap.available] ? cols[colMap.available].trim().replace(/,/g, '').replace(/\s/g, '') : '';
    var qualityInsp = colMap.qualityInsp !== undefined && cols[colMap.qualityInsp] ? cols[colMap.qualityInsp].trim().replace(/,/g, '').replace(/\s/g, '') : '';
    var releaseAlloc = colMap.releaseAlloc !== undefined && cols[colMap.releaseAlloc] ? cols[colMap.releaseAlloc].trim().replace(/,/g, '').replace(/\s/g, '') : '';
    var unusable = colMap.unusable !== undefined && cols[colMap.unusable] ? cols[colMap.unusable].trim().replace(/,/g, '').replace(/\s/g, '') : '';
    var customerName = colMap.customerName !== undefined && cols[colMap.customerName] ? cols[colMap.customerName].trim() : '';
    var lossRate = colMap.lossRate !== undefined && cols[colMap.lossRate] ? cols[colMap.lossRate].trim() : '';
    var supplier = colMap.supplier !== undefined && cols[colMap.supplier] ? cols[colMap.supplier].trim() : '';

    // 설명1, 전체이름: 헤더 기반 매핑이 어려우면 인덱스 기반 (헤더행에서 20,21번째)
    var desc1 = '';
    var fullName = '';
    if (colMap.desc1 !== undefined) {
      desc1 = cols[colMap.desc1] ? cols[colMap.desc1].trim() : '';
    }
    if (colMap.fullName !== undefined) {
      fullName = cols[colMap.fullName] ? cols[colMap.fullName].trim() : '';
    }

    var lastIn = colMap.lastIn !== undefined && cols[colMap.lastIn] ? cols[colMap.lastIn].trim() : '';
    var lastOut = colMap.lastOut !== undefined && cols[colMap.lastOut] ? cols[colMap.lastOut].trim() : '';
    var purchaser = colMap.purchaser !== undefined && cols[colMap.purchaser] ? cols[colMap.purchaser].trim() : '';
    var salesPerson = colMap.salesPerson !== undefined && cols[colMap.salesPerson] ? cols[colMap.salesPerson].trim() : '';

    // 최종입고일, 최종출고일 둘 다 없으면 연구원에 PT 추가
    if (!lastIn && !lastOut) {
      fullName = fullName ? fullName + ' PT' : 'PT';
    }

    if (!code) continue;
    // 7자로 시작하는 원자재(ROH1) 제외
    if (code.charAt(0) === '7') continue;

    bomData.push({
      lev: parseInt(lev),
      mtype: mtype,
      code: code,
      name: name,
      inputQty: parseFloat(inputQty) || 0,
      needQty: parseFloat(needQty) || 0,
      unit: unit,
      stockTotal: parseFloat(stockTotal) || 0,
      available: parseFloat(available) || 0,
      qualityInsp: parseFloat(qualityInsp) || 0,
      releaseAlloc: parseFloat(releaseAlloc) || 0,
      unusable: parseFloat(unusable) || 0,
      customerName: customerName,
      lastIn: lastIn,
      lastOut: lastOut,
      desc1: desc1,
      fullName: fullName,
      lossRate: lossRate,
      supplier: supplier,
      purchaser: purchaser,
      salesPerson: salesPerson
    });
  }

  if (bomData.length === 0) {
    alert('BOM 데이터를 찾을 수 없습니다. 파일 형식을 확인해 주세요.');
    return;
  }

  document.getElementById('bomStatus').textContent = bomFileName + ' (' + bomData.length + '건)';
  document.getElementById('bomStatus').classList.add('loaded');
  document.getElementById('resetBomBtn').style.display = 'inline-block';
  if (typeof checkShowReset === 'function') checkShowReset();
  renderBomTree(bomData);
}

// ============ CSV 파싱 ============
function parseBomCsv(text) {
  var result = Papa.parse(text, { header: true, skipEmptyLines: true });
  var rows = result.data;
  var bomData = [];

  // CSV 필수 컬럼 검증
  if (rows.length > 0) {
    var csvHeaders = Object.keys(rows[0]);
    var csvRequired = [
      { names: ['Lev'], label: 'Lev (레벨)' },
      { names: ['MTyp', '자재 유형'], label: 'MTyp (자재 유형)' },
      { names: ['자재'], label: '자재 (자재코드)' },
      { names: ['자재내역'], label: '자재내역 (제품명)' },
      { names: ['투입수량', '구성부품소요수량', '소요수량'], label: '투입수량 또는 소요수량' }
    ];
    var csvWarn = [
      { names: ['재고합계'], label: '재고합계' },
      { names: ['가용'], label: '가용 (가용재고)' }
    ];
    var csvMissingReq = [];
    var csvMissingWarn = [];
    for (var cq = 0; cq < csvRequired.length; cq++) {
      var found = false;
      for (var cn = 0; cn < csvRequired[cq].names.length; cn++) {
        if (csvHeaders.indexOf(csvRequired[cq].names[cn]) !== -1) { found = true; break; }
      }
      if (!found) csvMissingReq.push(csvRequired[cq].label);
    }
    for (var cw = 0; cw < csvWarn.length; cw++) {
      var found2 = false;
      for (var cn2 = 0; cn2 < csvWarn[cw].names.length; cn2++) {
        if (csvHeaders.indexOf(csvWarn[cw].names[cn2]) !== -1) { found2 = true; break; }
      }
      if (!found2) csvMissingWarn.push(csvWarn[cw].label);
    }
    if (csvMissingReq.length > 0) {
      alert('⚠️ CSV 파일에서 필수 컬럼을 찾을 수 없습니다.\n\n' +
        '누락된 필수 컬럼:\n• ' + csvMissingReq.join('\n• ') +
        '\n\nSAP BOM 조회 레이아웃을 확인해 주세요.\n' +
        '필수 컬럼이 포함된 레이아웃으로 다시 조회해 주세요.');
    } else if (csvMissingWarn.length > 0) {
      console.warn('[BOM CSV 경고] 일부 컬럼 미감지: ' + csvMissingWarn.join(', '));
    }
  }

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var lev = r['Lev'] !== undefined ? r['Lev'] : '';
    if (lev === '' || isNaN(parseInt(lev))) continue;

    var code = (r['자재'] || '').trim();
    if (code.charAt(0) === '7') continue;
    var name = (r['자재내역'] || '').trim();
    var mtype = (r['자재 유형'] || r['MTyp'] || '').trim();

    bomData.push({
      lev: parseInt(lev),
      mtype: mtype,
      code: code,
      name: name,
      inputQty: parseFloat(String(r['투입수량'] || r['구성부품소요수량'] || r['소요수량'] || '0').replace(/,/g, '')) || 0,
      needQty: parseFloat(String(r['구성부품소요수량'] || r['소요수량'] || '0').replace(/,/g, '')) || 0,
      unit: (r['기본 단위'] || r['BUn'] || '').trim(),
      stockTotal: parseFloat(String(r['재고합계'] || '0').replace(/,/g, '')) || 0,
      qualityInsp: parseFloat(String(r['품질 검사'] || r['품질검사'] || '0').replace(/,/g, '')) || 0,
      available: parseFloat(String(r['가용'] || '0').replace(/,/g, '')) || 0,
      releaseAlloc: parseFloat(String(r['출고할당'] || '0').replace(/,/g, '')) || 0,
      unusable: parseFloat(String(r['사용불가'] || '0').replace(/,/g, '')) || 0,
      customerName: (r['고객이름'] || '').trim(),
      lastIn: (r['최종입고일'] || '').trim(),
      lastOut: (r['최종출고일'] || '').trim(),
      desc1: (r['설명1'] || r['설명 1'] || '').trim(),
      fullName: (function() {
        var fn = (r['전체 이름'] || r['전체이름'] || '').trim();
        var li = (r['최종입고일'] || '').trim();
        var lo = (r['최종출고일'] || '').trim();
        if (!li && !lo) return fn ? fn + ' PT' : 'PT';
        return fn;
      })(),
      lossRate: (r['로스율'] || '').trim(),
      supplier: (r['이름 1'] || '').trim(),
      purchaser: (r['구매담당자'] || r['구매 담당자'] || '').trim(),
      salesPerson: (r['영업담당자'] || r['영업 담당자'] || '').trim()
    });
  }

  if (bomData.length === 0) {
    alert('BOM 데이터를 찾을 수 없습니다.');
    return;
  }

  document.getElementById('bomStatus').textContent = bomFileName + ' (' + bomData.length + '건)';
  document.getElementById('bomStatus').classList.add('loaded');
  document.getElementById('resetBomBtn').style.display = 'inline-block';
  if (typeof checkShowReset === 'function') checkShowReset();
  renderBomTree(bomData);
}


// BOM 파싱 결과 전역 저장
var parsedBomData = [];
var bomGroups = []; // 완제품별 그룹
var bomCurrentPage = 0;

// ============ BOM 트리 렌더링 ============
function renderBomTree(data) {
  parsedBomData = data;

  // 완제품(FERT, Lev 0) 추출하여 발주수량 입력란 생성
  var fertItems = data.filter(function(d) { return d.lev === 0; });
  if (fertItems.length > 0) {
    renderOrderInputs(fertItems);
  }

  // 완제품별로 그룹 분리
  bomGroups = [];
  var currentGroup = null;
  for (var i = 0; i < data.length; i++) {
    if (data[i].lev === 0) {
      currentGroup = { fert: data[i], items: [data[i]] };
      bomGroups.push(currentGroup);
    } else if (currentGroup) {
      currentGroup.items.push(data[i]);
    }
  }

  // 그룹이 없으면 전체를 하나의 그룹으로
  if (bomGroups.length === 0) {
    bomGroups = [{ fert: data[0], items: data }];
  }

  bomCurrentPage = 0;
  document.getElementById('bomResultSection').style.display = 'block';
  renderBomPage();
}

function renderBomPage() {
  var group = bomGroups[bomCurrentPage];
  var pageData = group.items;

  // 제목 업데이트
  var title = document.getElementById('bomTitle');
  title.textContent = 'BOM 구조 (' + (bomCurrentPage + 1) + '/' + bomGroups.length + ')';

  // 네비게이션 버튼 상태
  document.getElementById('bomPrevBtn').disabled = bomCurrentPage === 0;
  document.getElementById('bomNextBtn').disabled = bomCurrentPage === bomGroups.length - 1;

  var container = document.getElementById('bomTree');

  var html = '<div class="bom-table-wrap"><table class="bom-table">' +
    '<thead><tr>' +
      '<th>구분</th>' +
      '<th>코드</th>' +
      '<th>제품명</th>' +
      '<th>소요수량</th>' +
      '<th>단위</th>' +
      '<th>재고합계</th>' +
      '<th>가용재고</th>' +
      '<th>품질검사</th>' +
      '<th>출고할당</th>' +
      '<th>사용불가</th>' +
      '<th>고객사</th>' +
      '<th>최종입고일</th>' +
      '<th>최종출고일</th>' +
      '<th>타입</th>' +
      '<th>연구원</th>' +
      '<th>구매담당자</th>' +
      '<th>영업담당자</th>' +
    '</tr></thead><tbody>';

  for (var i = 0; i < pageData.length; i++) {
    var d = pageData[i];
    var typeClass = d.mtype.toLowerCase().replace(/[0-9]/g, '');
    if (['fert', 'halb', 'hal', 'roh'].indexOf(typeClass) === -1) typeClass = 'other';
    if (d.mtype === 'HAL1') typeClass = 'hal1';
    if (d.mtype === 'HAL2') typeClass = 'hal2';
    if (d.mtype === 'ROH1') typeClass = 'roh1';

    var stockClass = d.available > 0 ? ' class="stock-pos"' : (d.stockTotal > 0 ? '' : ' class="stock-zero"');
    var availClass = d.available > 0 ? ' class="stock-pos"' : ' class="stock-zero"';

    var codeLabel = '';
    var firstChar = d.code.charAt(0);
    if (firstChar === '9') codeLabel = '완제품';
    else if (firstChar === '1') codeLabel = '반제품';
    else if (firstChar === '2') codeLabel = '성형물';
    else if (firstChar === '3') codeLabel = '벌크';
    else if (firstChar === '4') codeLabel = '베이스';
    else if (firstChar === '7') codeLabel = '자재';
    else codeLabel = d.mtype;

    html += '<tr>' +
      '<td>' + codeLabel + '</td>' +
      '<td class="bom-code">' + d.code + '</td>' +
      '<td class="bom-name" title="' + d.name + '">' + d.name + '</td>' +
      '<td class="num">' + (d.inputQty ? (Math.round(d.inputQty * 1000) / 1000) : '-') + '</td>' +
      '<td class="unit">' + d.unit + '</td>' +
      '<td' + stockClass + '>' + (d.stockTotal ? d.stockTotal.toLocaleString() : '-') + '</td>' +
      '<td' + availClass + '>' + (d.available ? d.available.toLocaleString() : '-') + '</td>' +
      '<td class="num">' + (d.qualityInsp ? d.qualityInsp.toLocaleString() : '-') + '</td>' +
      '<td class="num">' + (d.releaseAlloc ? d.releaseAlloc.toLocaleString() : '-') + '</td>' +
      '<td class="num">' + (d.unusable ? d.unusable.toLocaleString() : '-') + '</td>' +
      '<td class="bom-name">' + (d.customerName || '-') + '</td>' +
      '<td>' + (d.lastIn || '-') + '</td>' +
      '<td>' + (d.lastOut || '-') + '</td>' +
      '<td class="bom-name">' + (d.desc1 || '-') + '</td>' +
      '<td class="bom-name">' + (d.fullName || '-') + '</td>' +
      '<td class="bom-name">' + (d.purchaser || '-') + '</td>' +
      '<td class="bom-name">' + (d.salesPerson || '-') + '</td>' +
    '</tr>';
  }

  html += '</tbody></table></div>';
  container.innerHTML = html;
}

// ============ BOM 네비게이션 ============
function bomNavPrev() {
  if (bomCurrentPage > 0) {
    bomCurrentPage--;
    renderBomPage();
  }
}

function bomNavNext() {
  if (bomCurrentPage < bomGroups.length - 1) {
    bomCurrentPage++;
    renderBomPage();
  }
}

// ============ BOM 검색 필터 ============
function filterBomTree() {
  var keyword = document.getElementById('bomSearch').value.trim().toLowerCase();
  var rows = document.querySelectorAll('#bomTree .bom-table tbody tr');
  if (!keyword) {
    rows.forEach(function(tr) {
      tr.classList.remove('search-hidden');
      tr.classList.remove('search-highlight');
    });
    return;
  }
  rows.forEach(function(tr) {
    var text = tr.textContent.toLowerCase();
    if (text.indexOf(keyword) >= 0) {
      tr.classList.remove('search-hidden');
      tr.classList.add('search-highlight');
    } else {
      tr.classList.add('search-hidden');
      tr.classList.remove('search-highlight');
    }
  });
}

// ============ 발주수량 입력란 렌더링 ============
function renderOrderInputs(fertItems) {
  var section = document.getElementById('bomOrderSection');
  section.style.display = 'block';

  var container = document.getElementById('bomOrderInputs');
  var html = '<div class="bom-table-wrap"><table class="bom-table">' +
    '<thead><tr>' +
      '<th>코드</th>' +
      '<th>제품명</th>' +
      '<th>발주 수량</th>' +
    '</tr></thead><tbody>';

  for (var i = 0; i < fertItems.length; i++) {
    var f = fertItems[i];
    html += '<tr>' +
      '<td class="bom-code">' + f.code + '</td>' +
      '<td class="bom-name">' + f.name + '</td>' +
      '<td><input type="text" class="bom-order-qty" data-code="' + f.code + '" placeholder="수량 입력" oninput="formatBomQty(this)" onkeydown="bomQtyEnter(event)" style="padding:6px 10px;border:1px solid #ddd;border-radius:4px;font-size:14px;text-align:right;width:120px"></td>' +
    '</tr>';
  }

  html += '</tbody></table></div>';
  container.innerHTML = html;

  // VBS에서 입력된 발주 수량 자동 반영
  if (Object.keys(prefilledOrderQtys).length > 0) {
    var qtyInputs = document.querySelectorAll('.bom-order-qty');
    qtyInputs.forEach(function(inp) {
      var code = inp.dataset.code;
      if (prefilledOrderQtys[code]) {
        inp.value = Number(prefilledOrderQtys[code]).toLocaleString();
      }
    });
  }
}

function bomQtyEnter(e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    var inputs = document.querySelectorAll('.bom-order-qty');
    var arr = Array.prototype.slice.call(inputs);
    var idx = arr.indexOf(e.target);
    if (idx < arr.length - 1) {
      arr[idx + 1].focus();
    }
  }
}

function formatBomQty(input) {
  var raw = input.value.replace(/[^0-9]/g, '');
  if (raw === '') { input.value = ''; return; }
  input.value = Number(raw).toLocaleString();
}

// ============ BOM → 벌크 필요량 계산 ============
function calcBomNeeds() {
  var inputs = document.querySelectorAll('.bom-order-qty');
  var orderMap = {};
  var hasInput = false;

  inputs.forEach(function(inp) {
    var qty = parseFloat(inp.value.replace(/,/g, '')) || 0;
    if (qty > 0) {
      orderMap[inp.dataset.code] = qty;
      hasInput = true;
    }
  });

  if (!hasInput) {
    alert('발주 수량을 입력해 주세요.');
    return;
  }

  // === 성형물 가용재고 집계 (홋수별 그룹) ===
  var shadeGroups = {};
  var moldToShadeKey = {};
  var _parentStack = [];
  var _fertCode = null;

  for (var i = 0; i < parsedBomData.length; i++) {
    var d = parsedBomData[i];
    _parentStack.length = d.lev + 1;
    _parentStack[d.lev] = d;

    if (d.lev === 0 && orderMap[d.code]) {
      _fertCode = d.code;
    } else if (d.lev === 0) {
      _fertCode = null;
    }
    if (!_fertCode) continue;

    if (d.code.charAt(0) === '2' && d.mtype === 'HAL1') {
      var parent = d.lev > 0 ? _parentStack[d.lev - 1] : null;
      var groupKey;
      if (!parent || parent.code.charAt(0) === '9') {
        // 완제품 바로 아래 → 각 성형물이 개별 홋수
        groupKey = _fertCode + '_' + d.code;
      } else {
        // 중간 부모(반제품 등) 아래 → 같은 부모 = 같은 홋수
        groupKey = _fertCode + '_' + parent.code;
      }
      if (!shadeGroups[groupKey]) {
        shadeGroups[groupKey] = { totalAvailable: 0, moldCodes: [] };
      }
      shadeGroups[groupKey].totalAvailable += (d.available || 0);
      shadeGroups[groupKey].moldCodes.push(d.code);
      moldToShadeKey[_fertCode + '_' + d.code] = groupKey;
    }
  }

  // BOM 트리를 순회하여 완제품 → 성형물 → 벌크 구조 추출
  var results = [];
  var currentFert = null;
  var currentMold = null;

  for (var i = 0; i < parsedBomData.length; i++) {
    var d = parsedBomData[i];

    if (d.lev === 0 && orderMap[d.code]) {
      currentFert = { code: d.code, name: d.name, orderQty: orderMap[d.code] };
      currentMold = null;
    } else if (d.lev === 0) {
      currentFert = null;
      currentMold = null;
    }

    if (!currentFert) continue;

    // 성형물 (HAL1, 2코드) — Lev 1 또는 Lev 2
    if (d.code.charAt(0) === '2' && d.mtype === 'HAL1') {
      currentMold = { code: d.code, name: d.name, inputQty: d.inputQty };
    }

    // 벌크 (HAL2, 3코드)
    if (d.code.charAt(0) === '3' && d.mtype === 'HAL2' && currentMold) {
      var moldNeedQtyOriginal = currentFert.orderQty * currentMold.inputQty;
      // 성형물 가용재고 차감 (홋수별 합산)
      var shadeKey = moldToShadeKey[currentFert.code + '_' + currentMold.code];
      var shadeAvailable = (shadeKey && shadeGroups[shadeKey]) ? shadeGroups[shadeKey].totalAvailable : 0;
      var moldNeedQty = shadeAvailable > 0 ? Math.max(0, moldNeedQtyOriginal - shadeAvailable) : moldNeedQtyOriginal;
      var bulkTheoryNeed = moldNeedQty * d.inputQty;

      // 최적 제조량 예측 (SAP 실적 데이터 있으면)
      var optimalQty = null;
      var avgLossRate = null;
      var historyCount = 0;
      var historyRecords = [];
      var confidenceLevel = 'none';
      var confidenceScore = 0;
      var validHistoryCount = 0;
      var stdDev = 0;
      var recencyDays = null;
      if (typeof moldBulkMap !== 'undefined' && moldBulkMap[currentMold.code]) {
        var predicted = predictOne(currentMold.code, Math.round(moldNeedQty));
        for (var p = 0; p < predicted.length; p++) {
          if (!predicted[p].error && predicted[p].bulkCode === d.code) {
            avgLossRate = predicted[p].avgLossRate;
            historyCount = predicted[p].historyCount;
            confidenceLevel = predicted[p].confidenceLevel;
            confidenceScore = predicted[p].confidenceScore;
            validHistoryCount = predicted[p].validHistoryCount;
            stdDev = predicted[p].stdDev;
            recencyDays = predicted[p].recencyDays;
            // BOM 이론 필요량 기준으로 최적 제조량 재계산
            optimalQty = avgLossRate !== null ? Math.ceil(bulkTheoryNeed * (1 + avgLossRate / 100)) : null;
            break;
          }
        }
        // 실적 이력 데이터 수집
        if (moldBulkMap[currentMold.code][d.code]) {
          var recs = moldBulkMap[currentMold.code][d.code].records;
          for (var ri = 0; ri < recs.length; ri++) {
            var rec = recs[ri];
            // 환입/폐기 정보 확인
            var returnInfo = [];
            var adjustedInput = rec.actualInput;
            if (typeof returnIndex !== 'undefined' && returnIndex[rec.prodOrder]) {
              var rdList = returnIndex[rec.prodOrder];
              for (var rdi = 0; rdi < rdList.length; rdi++) {
                var rd = rdList[rdi];
                if (rd.workTeamCode && rec.workTeam) {
                  var isHwaseong = rec.workTeam === '파우더성형실' || rec.workTeam.indexOf('3002') !== -1;
                  var isPyeongtaek = rec.workTeam.indexOf('평택') !== -1 || rec.workTeam.indexOf('7002') !== -1;
                  if (rd.workTeamCode === '3002' && !isHwaseong) continue;
                  if (rd.workTeamCode === '7002' && !isPyeongtaek) continue;
                }
                if (!rd.bulkCode || rd.bulkCode === rec.bulkCode) {
                  returnInfo.push({ type: rd.type, qty: rd.qty });
                  if (rd.type === '폐기') {
                    // 폐기: 전산에 투입으로 잡혀있으나 실제 사용하지 않았으므로 차감
                    adjustedInput = adjustedInput - rd.qty;
                  }
                }
              }
            }
            historyRecords.push({
              prodDate: rec.prodDate,
              prodOrder: rec.prodOrder,
              orderQty: rec.orderQty,
              actualQty: rec.actualQty,
              stdNeed: rec.stdNeed,
              actualInput: rec.actualInput,
              adjustedInput: adjustedInput,
              inputRate: rec.inputRate,
              damageQty: rec.damageQty,
              machine: rec.machine,
              returnInfo: returnInfo
            });
          }
        }
      }

      results.push({
        fertCode: currentFert.code,
        fertName: currentFert.name,
        fertOrderQty: currentFert.orderQty,
        moldCode: currentMold.code,
        moldName: currentMold.name,
        moldNeedQty: moldNeedQty,
        moldNeedQtyOriginal: moldNeedQtyOriginal,
        shadeAvailable: shadeAvailable,
        bulkCode: d.code,
        bulkName: d.name,
        bulkInputPerUnit: d.inputQty,
        bulkTheoryNeed: bulkTheoryNeed,
        optimalQty: optimalQty,
        avgLossRate: avgLossRate,
        historyCount: historyCount,
        validHistoryCount: validHistoryCount,
        confidenceLevel: confidenceLevel,
        confidenceScore: confidenceScore,
        stdDev: stdDev,
        recencyDays: recencyDays,
        historyRecords: historyRecords
      });
    }
  }

  if (results.length === 0) {
    alert('벌크 데이터를 찾을 수 없습니다. BOM 구조를 확인해 주세요.');
    return;
  }

  // 결과 테이블 렌더링
  renderBomCalcResults(results);
}

// ============ 벌크 필요량 결과 → 새 창으로 표시 ============
function renderBomCalcResults(results) {
  window._bomCalcResults = results;

  // 완제품별로 그룹 분리
  var fertGroups = {};
  var fertOrder = [];
  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    if (!fertGroups[r.fertCode]) {
      fertGroups[r.fertCode] = { code: r.fertCode, name: r.fertName, items: [] };
      fertOrder.push(r.fertCode);
    }
    fertGroups[r.fertCode].items.push(r);
  }

  var w = Math.round(window.screen.width * 0.95);
  var h = Math.round(window.screen.height * 0.8);
  var left = Math.round((window.outerWidth - w) / 2 + window.screenX);
  var top = Math.round((window.outerHeight - h) / 2 + window.screenY);

  var popup = window.open('', 'bomCalcPopup', 'width=' + w + ',height=' + h + ',left=' + left + ',top=' + top + ',scrollbars=yes,resizable=yes');

  var sapNote = (typeof sapCount !== 'undefined' && sapCount > 0) ?
    '✅ 표준 대비 실적 데이터 반영됨 (' + sapCount.toLocaleString() + '건)' :
    '⚠️ 표준 대비 실적 데이터 미등록 — 로스율/최적 제조량은 실적 데이터 업로드 후 확인 가능';

  var html = '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">' +
    '<script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"><\/script>' +
    '<title>벌크별 필요 제조량</title>' +
    '<style>' +
      'body { font-family: "Segoe UI","Malgun Gothic",sans-serif; background:#f5f6fa; margin:0; padding:25px 30px; color:#333; }' +
      '.top-bar { display:flex; justify-content:space-between; align-items:center; margin-bottom:20px; }' +
      '.top-actions { display:flex; gap:8px; }' +
      '.nav-area { display:flex; align-items:center; gap:12px; }' +
      'h1 { font-size:20px; margin:0; display:flex; align-items:center; gap:8px; }' +
      'h1::before { content:""; display:inline-block; width:4px; height:16px; background:#c8102e; border-radius:2px; }' +
      '.nav-btn { background:#fff; border:1px solid #ddd; border-radius:6px; padding:6px 14px; font-size:16px; font-weight:700; cursor:pointer; color:#333; }' +
      '.nav-btn:hover { background:#f0f0f0; border-color:#999; }' +
      '.nav-btn:disabled { color:#ddd; cursor:default; background:#fafafa; border-color:#eee; }' +
      'table { width:100%; border-collapse:collapse; font-size:12px; background:#fff; border-radius:8px; overflow:hidden; box-shadow:0 2px 8px rgba(0,0,0,0.05); }' +
      'thead th { background:#f5f6fa; color:#666; padding:8px 12px; font-size:12px; font-weight:700; text-align:center; border-bottom:2px solid #e0e0e0; border-right:1px solid #ddd; white-space:nowrap; }' +
      'thead th:last-child { border-right:none; }' +
      'tbody td { padding:6px 12px; border-bottom:1px solid #f0f0f0; border-right:1px solid #eee; white-space:nowrap; text-align:center; }' +
      'tbody td:last-child { border-right:none; }' +
      'tbody tr:hover { background:#f8f9ff; }' +
      'td.name { text-align:left; color:#555; }' +
      'td.num { text-align:center; font-weight:600; }' +
      'td.highlight { color:#2962ff; font-weight:normal; text-align:center; }' +
      'td.optimal { color:#c8102e; font-weight:normal; text-align:center; }' +
      '.actions { margin-top:20px; }' +
      '.btn { padding:10px 24px; border:none; border-radius:6px; font-size:15px; font-weight:600; cursor:pointer; transition:background 0.2s; }' +
      '.btn-close { background:#e0e0e0; color:#333; }' +
      '.btn-close:hover { background:#ccc; }' +
      '.btn-copy { background:#59a14f; color:#fff; margin-right:8px; }' +
      '.btn-copy:hover { background:#468a3e; }' +
      '.btn-ml-copy { background:#7c4dff; color:#fff; margin-right:8px; }' +
      '.btn-ml-copy:hover { background:#6236d6; }' +
      '.btn-ml { background:#7c4dff; color:#fff; margin-right:8px; }' +
      '.btn-ml:hover { background:#6236d6; }' +
      '.btn-ml.active { background:#ff6d00; }' +
      '.btn-ml.active:hover { background:#e65100; }' +
      '.ml-col { background:#f3eaff; }' +
      'th.ml-col { background:#ede0ff; }' +
      '.btn-download { background:#4e79a7; color:#fff; margin-right:8px; }' +
      '.btn-download:hover { background:#3a6290; }' +
      '.toast { position:fixed; top:20px; right:20px; background:#333; color:#fff; padding:10px 20px; border-radius:6px; font-size:13px; display:none; z-index:999; }' +
      'tbody tr { cursor:pointer; }' +
      '.modal-overlay { display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:100; }' +
      '.modal { position:absolute; top:50%; left:50%; transform:translate(-50%,-50%); background:#fff; border-radius:10px; padding:24px; width:75%; height:75%; overflow-y:auto; box-shadow:0 8px 30px rgba(0,0,0,0.15); }' +
      '.modal h2 { font-size:16px; margin:0 0 4px; }' +
      '.modal .sub { font-size:12px; color:#999; margin-bottom:14px; }' +
      '.modal table { width:100%; }' +
      '.modal .close-btn { position:absolute; top:12px; right:16px; background:none; border:none; font-size:20px; cursor:pointer; color:#999; }' +
      '.modal .close-btn:hover { color:#333; }' +
      '.modal .summary { display:flex; gap:20px; margin-bottom:14px; font-size:13px; align-items:center; }' +
      '.modal .summary span { font-weight:700; }' +
      '.note { font-size:13px; color:#666; padding:8px 0; }' +
      '.help-icon { display:inline-block; width:16px; height:16px; line-height:14px; text-align:center; border:1px solid #c8102e; border-radius:50%; font-size:10px; color:#c8102e; cursor:help; font-weight:700; margin-left:4px; position:relative; }' +
      '.help-icon:hover .help-tip { display:block; }' +
      '.help-tip { display:none; position:absolute; top:24px; left:0; background:#1a1a1a; color:#fff; padding:14px 18px; border-radius:8px; font-size:12px; font-weight:400; white-space:pre; min-width:340px; text-align:left; line-height:1.7; box-shadow:0 4px 20px rgba(0,0,0,0.25); z-index:9999; }' +
      '.help-tip::before { content:""; position:absolute; bottom:100%; left:6px; border:6px solid transparent; border-bottom-color:#1a1a1a; }' +
      '.conf-badge { display:inline-block; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:700; }' +
      '.conf-high { background:#d4edda; color:#155724; }' +
      '.conf-medium { background:#fff3cd; color:#856404; }' +
      '.conf-low { background:#ffe0b2; color:#bf360c; }' +
      '.conf-verylow { background:#f8d7da; color:#721c24; }' +
      '.conf-none { background:#f0f0f0; color:#999; }' +
      '.sort-select { padding:7px 12px; border:1px solid #ddd; border-radius:6px; font-size:13px; background:#fff; cursor:pointer; font-family:inherit; }' +
      '.sort-select:focus { outline:none; border-color:#4e79a7; }' +
      'tr.row-low { background:#fff7ed; }' +
      'tr.row-low:hover { background:#ffedd5; }' +
      'tr.row-verylow { background:#fef2f2; }' +
      'tr.row-verylow:hover { background:#fee2e2; }' +
      'tr.row-none { background:#fafafa; color:#999; }' +
      'tr.row-none:hover { background:#f0f0f0; }' +
      'td.diff-warn { background:#fff3cd !important; color:#856404; font-weight:700; position:relative; }' +
      'td.diff-alert { background:#f8d7da !important; color:#721c24; font-weight:700; position:relative; }' +
      '.ml-sub { display:block; font-size:10px; color:#7c4dff; font-weight:600; margin-top:2px; }' +
    '</style></head><body>' +
    '<div class="top-bar">' +
      '<div class="nav-area">' +
        '<button class="nav-btn" id="prevBtn" onclick="navPrev()">&lt;</button>' +
        '<h1 id="pageTitle">벌크별 필요 제조량</h1>' +
        '<button class="nav-btn" id="nextBtn" onclick="navNext()">&gt;</button>' +
      '</div>' +
      '<div class="top-actions">' +
        '<select class="sort-select" id="sortSelect" onchange="changeSort(this.value)">' +
          '<option value="default">기본 순서</option>' +
          '<option value="confLow">신뢰도 낮은 순</option>' +
          '<option value="confHigh">신뢰도 높은 순</option>' +
          '<option value="qtyDesc">최적 제조량 큰 순</option>' +
          '<option value="qtyAsc">최적 제조량 작은 순</option>' +
          '<option value="lossDesc">로스율 높은 순</option>' +
        '</select>' +
        '<button class="btn btn-copy" onclick="copyOptimalQty()">최적 제조량 복사</button>' +
        '<button class="btn btn-ml-copy" onclick="copyMLQty()" id="mlCopyBtn" style="display:none">ML 제조량 복사</button>' +
        '<button class="btn btn-ml" onclick="toggleML()" id="mlBtn">ML 예측 비교</button>' +
        '<button class="btn btn-download" onclick="downloadExcel()">엑셀 다운로드</button>' +
        '<button class="btn btn-close" onclick="window.close()">닫기</button>' +
      '</div>' +
    '</div>' +
    '<div class="toast" id="toast">복사 완료!</div>' +
    '<div class="modal-overlay" id="modalOverlay" onclick="closeModal()">' +
      '<div class="modal" onclick="event.stopPropagation()">' +
        '<button class="close-btn" onclick="closeModal()">x</button>' +
        '<h2 id="modalTitle">실적 이력</h2>' +
        '<div class="sub" id="modalSub"></div>' +
        '<div class="summary" id="modalSummary"></div>' +
        '<div id="modalBody"></div>' +
      '</div>' +
    '</div>' +
    '<div id="tableArea"></div>' +
    '<div class="actions">' +
      '<div class="note">' + sapNote + '</div>' +
      '<div class="note ml-legend" id="mlLegend" style="display:none;margin-top:8px;padding:10px 14px;background:#fafbfc;border-radius:6px;border-left:3px solid #7c4dff">' +
        '<div style="font-weight:700;margin-bottom:6px;color:#333">📊 ML vs 통계 예측 차이 색상 안내</div>' +
        '<div style="display:flex;gap:18px;align-items:center;flex-wrap:wrap;font-size:12px">' +
          '<div style="display:flex;align-items:center;gap:6px"><span style="display:inline-block;width:16px;height:16px;background:#fff3cd;border:1px solid #856404;border-radius:3px"></span> 차이 5%p ~ 10%p (주의)</div>' +
          '<div style="display:flex;align-items:center;gap:6px"><span style="display:inline-block;width:16px;height:16px;background:#f8d7da;border:1px solid #721c24;border-radius:3px"></span> 차이 10%p 이상 (경고)</div>' +
          '<div style="color:#999">* 셀에 마우스를 올리면 정확한 차이 수치 표시</div>' +
        '</div>' +
      '</div>' +
    '</div>' +
    '<script>' +
      'var fertOrder = ' + JSON.stringify(fertOrder) + ';' +
      'var fertGroups = ' + JSON.stringify(fertGroups) + ';' +
      'var allResults = ' + JSON.stringify(results) + ';' +
      'var currentPage = 0;' +
      'var showML = false;' +
      'var currentSort = "default";' +
      // 원본 순서 백업 (정렬 후 복원용)
      'for (var fk in fertGroups) {' +
        'fertGroups[fk]._originalItems = fertGroups[fk].items.slice();' +
      '}' +
      'function changeSort(sortKey) {' +
        'currentSort = sortKey;' +
        'renderPage();' +
      '}' +
      'function getSortedItems(items) {' +
        'var sorted = items.slice();' +
        'var levelRank = { high: 4, medium: 3, low: 2, verylow: 1, none: 0 };' +
        'if (currentSort === "confLow") {' +
          'sorted.sort(function(a, b) { return (levelRank[a.confidenceLevel] || 0) - (levelRank[b.confidenceLevel] || 0); });' +
        '} else if (currentSort === "confHigh") {' +
          'sorted.sort(function(a, b) { return (levelRank[b.confidenceLevel] || 0) - (levelRank[a.confidenceLevel] || 0); });' +
        '} else if (currentSort === "qtyDesc") {' +
          'sorted.sort(function(a, b) { return (b.optimalQty || 0) - (a.optimalQty || 0); });' +
        '} else if (currentSort === "qtyAsc") {' +
          'sorted.sort(function(a, b) { return (a.optimalQty || Infinity) - (b.optimalQty || Infinity); });' +
        '} else if (currentSort === "lossDesc") {' +
          'sorted.sort(function(a, b) { return (b.avgLossRate !== null ? b.avgLossRate : -Infinity) - (a.avgLossRate !== null ? a.avgLossRate : -Infinity); });' +
        '}' +
        'return sorted;' +
      '}' +
      'var mlData = ' + (mlPredictions ? JSON.stringify(mlPredictions) : 'null') + ';' +
      'function getMLPrediction(bulkCode, stdNeed) {' +
        'if (!mlData || !mlData.bulkLookup || !mlData.bulkLookup[bulkCode]) return null;' +
        'var pred = mlData.bulkLookup[bulkCode].pred;' +
        'var ranges = mlData.modelInfo.qtyRanges;' +
        'var closest = ranges[0];' +
        'for (var i = 0; i < ranges.length; i++) {' +
          'if (Math.abs(ranges[i] - stdNeed) < Math.abs(closest - stdNeed)) closest = ranges[i];' +
        '}' +
        'return pred[String(closest)] !== undefined ? pred[String(closest)] : null;' +
      '}' +
      'function renderPage() {' +
        'var code = fertOrder[currentPage];' +
        'var group = fertGroups[code];' +
        'group.items = getSortedItems(group._originalItems);' +
        'document.getElementById("pageTitle").textContent = "벌크별 필요 제조량 (" + (currentPage+1) + "/" + fertOrder.length + ")";' +
        'document.getElementById("prevBtn").disabled = currentPage === 0;' +
        'document.getElementById("nextBtn").disabled = currentPage === fertOrder.length - 1;' +
        'var rows = "";' +
        'for (var i = 0; i < group.items.length; i++) {' +
          'var r = group.items[i];' +
          'var mlRate = showML ? getMLPrediction(r.bulkCode, r.bulkTheoryNeed) : null;' +
          'var mlQty = (mlRate !== null) ? Math.ceil(r.bulkTheoryNeed * (1 + mlRate / 100)) : null;' +
          // 신뢰도 배지
          'var confLabel = "-";' +
          'var confInfo = "";' +
          'var confExtra = "";' +
          'if (r.confidenceLevel !== "none") {' +
            'confInfo = "[현재 측정값]\\n• 표본 수: " + r.validHistoryCount + "건\\n• 편차: " + r.stdDev.toFixed(1) + "%" + (r.recencyDays !== null ? "\\n• 최신성: " + r.recencyDays + "일 전" : "") + "\\n• 점수: " + r.confidenceScore + "/100";' +
          '}' +
          'var confCriteria = "\\n\\n[점수 기준 (총 100점)]\\n• 표본 수 (40점)\\n  - 5건 이상: 40점\\n  - 3~4건: 25점\\n  - 1~2건: 10점\\n• 편차 (40점)\\n  - 3% 미만: 40점\\n  - 7% 미만: 25점\\n  - 15% 미만: 10점\\n• 최신성 (20점)\\n  - 30일 이내: 20점\\n  - 90일 이내: 10점\\n  - 180일 이내: 5점\\n\\n[등급]\\n• 높음: 80점+\\n• 보통: 50~79점\\n• 낮음: 30~49점 (참고용)\\n• 매우 낮음: 30점 미만 (실측 필수)";' +
          'if (r.confidenceLevel === "high") confLabel = "높음";' +
          'else if (r.confidenceLevel === "medium") confLabel = "보통";' +
          'else if (r.confidenceLevel === "low") { confLabel = "낮음"; confExtra = "\\n\\n참고용으로만 사용 권장"; }' +
          'else if (r.confidenceLevel === "verylow") { confLabel = "매우 낮음"; confExtra = "\\n\\n실측 확인 필수"; }' +
          'var confTipText = confInfo + confExtra + confCriteria;' +
          'var confBadge = r.confidenceLevel === "none" ? "<span class=\\"conf-badge conf-none\\">-</span>" : "<span class=\\"conf-badge conf-" + r.confidenceLevel + "\\" title=\\"" + confTipText + "\\">" + confLabel + "</span>";' +
          'var rowClass = "";' +
          'if (r.confidenceLevel === "low") rowClass = "row-low";' +
          'else if (r.confidenceLevel === "verylow") rowClass = "row-verylow";' +
          'else if (r.confidenceLevel === "none") rowClass = "row-none";' +
          'rows += "<tr class=\\"" + rowClass + "\\" ondblclick=\\"showHistory(" + i + ")\\">" +' +
            '"<td>" + r.moldCode + "</td>" +' +
            '"<td class=\\"name\\">" + r.moldName + "</td>" +' +
            '"<td>" + r.bulkCode + "</td>" +' +
            '"<td class=\\"name\\">" + r.bulkName + "</td>" +' +
            '"<td class=\\"num\\">" + (Math.round(r.bulkInputPerUnit * 1000) / 1000) + "</td>" +' +
            '"<td class=\\"num\\">" + Math.round(r.moldNeedQtyOriginal).toLocaleString() + "</td>" +' +
            '"<td class=\\"num\\">" + (r.shadeAvailable > 0 ? Math.round(r.shadeAvailable).toLocaleString() : "-") + "</td>" +' +
            '"<td class=\\"num\\">" + Math.round(r.moldNeedQty).toLocaleString() + "</td>" +' +
            '"<td class=\\"num highlight\\">" + Math.round(r.bulkTheoryNeed).toLocaleString() + "</td>" +' +
            // 로스율/제조량 셀 (ML 모드 시 같은 셀에 인라인 표시)
            '(function() {' +
              'var diff = null; var diffClass = "";' +
              'if (showML && mlRate !== null && r.avgLossRate !== null) {' +
                'diff = Math.abs(mlRate - r.avgLossRate);' +
                'if (diff >= 10) diffClass = "diff-alert";' +
                'else if (diff >= 5) diffClass = "diff-warn";' +
              '}' +
              'var diffTip = diff !== null ? " title=\\"통계 vs ML 차이: " + diff.toFixed(2) + "%p\\"" : "";' +
              'var html = "";' +
              // 로스율 셀: 통계 + ML 인라인
              'var lossText = r.avgLossRate !== null ? r.avgLossRate.toFixed(2) + "%" : "-";' +
              'if (showML && mlRate !== null) lossText += "<span class=\\"ml-sub\\">ML " + mlRate.toFixed(2) + "%</span>";' +
              'html += "<td class=\\"num " + diffClass + "\\"" + diffTip + ">" + lossText + "</td>";' +
              // 제조량 셀: 통계 + ML 인라인
              'var qtyText = r.optimalQty !== null ? Math.round(r.optimalQty).toLocaleString() : "-";' +
              'if (showML && mlQty !== null) qtyText += "<span class=\\"ml-sub\\">ML " + mlQty.toLocaleString() + "</span>";' +
              'html += "<td class=\\"num optimal\\">" + qtyText + "</td>";' +
              'html += "<td>" + confBadge + "</td>";' +
              'return html;' +
            '})() +' +
          '"</tr>";' +
        '}' +
        'document.getElementById("tableArea").innerHTML = "<table><thead><tr>" +' +
          '"<th>성형물 코드</th><th>성형물명</th><th>벌크 코드</th><th>벌크명</th><th>투입량(g)</th><th>필요 수량(ea)</th><th>가용 재고(ea)</th><th>최종 수량(ea)</th><th>이론 필요량(g)</th><th>로스율</th><th>최적 제조량(g)</th><th>신뢰도</th>" +' +
          '"</tr></thead><tbody>" + rows + "</tbody></table>";' +
      '}' +
      'function toggleML() {' +
        'var btn = document.getElementById("mlBtn");' +
        'if (!mlData) {' +
          'alert("ML 예측 데이터가 없습니다.\\n페이지를 새로고침 후 다시 시도해 주세요.");' +
          'return;' +
        '}' +
        'showML = !showML;' +
        'btn.textContent = showML ? "ML 숨기기" : "ML 예측 비교";' +
        'btn.classList.toggle("active", showML);' +
        'document.getElementById("mlCopyBtn").style.display = showML ? "inline-block" : "none";' +
        'document.getElementById("mlLegend").style.display = showML ? "block" : "none";' +
        'renderPage();' +
      '}' +
      'function navPrev() { if (currentPage > 0) { currentPage--; renderPage(); } }' +
      'function navNext() { if (currentPage < fertOrder.length - 1) { currentPage++; renderPage(); } }' +
      'renderPage();' +
      'function copyOptimalQty() {' +
        'var code = fertOrder[currentPage];' +
        'var group = fertGroups[code];' +
        'var lines = [];' +
        'for (var i = 0; i < group.items.length; i++) {' +
          'var r = group.items[i];' +
          'lines.push(r.optimalQty !== null ? Math.round(r.optimalQty).toLocaleString() : "-");' +
        '}' +
        'var text = lines.join("\\n");' +
        'if (navigator.clipboard) {' +
          'navigator.clipboard.writeText(text).then(function() { showToast("최적 제조량 복사 완료! (" + lines.length + "건)"); });' +
        '} else {' +
          'var ta = document.createElement("textarea");' +
          'ta.value = text; document.body.appendChild(ta); ta.select(); document.execCommand("copy"); document.body.removeChild(ta);' +
          'showToast("최적 제조량 복사 완료! (" + lines.length + "건)");' +
        '}' +
      '}' +
      'function copyMLQty() {' +
        'if (!mlData) { alert("ML 예측 비교 버튼을 먼저 클릭해 주세요."); return; }' +
        'var code = fertOrder[currentPage];' +
        'var group = fertGroups[code];' +
        'var lines = [];' +
        'for (var i = 0; i < group.items.length; i++) {' +
          'var r = group.items[i];' +
          'var mlRate = getMLPrediction(r.bulkCode, r.bulkTheoryNeed);' +
          'var mlQty = (mlRate !== null) ? Math.ceil(r.bulkTheoryNeed * (1 + mlRate / 100)) : null;' +
          'lines.push(mlQty !== null ? mlQty.toLocaleString() : "-");' +
        '}' +
        'var text = lines.join("\\n");' +
        'if (navigator.clipboard) {' +
          'navigator.clipboard.writeText(text).then(function() { showToast("ML 제조량 복사 완료! (" + lines.length + "건)"); });' +
        '} else {' +
          'var ta = document.createElement("textarea");' +
          'ta.value = text; document.body.appendChild(ta); ta.select(); document.execCommand("copy"); document.body.removeChild(ta);' +
          'showToast("ML 제조량 복사 완료! (" + lines.length + "건)");' +
        '}' +
      '}' +
      'function showToast(msg) {' +
        'var t = document.getElementById("toast"); t.textContent = msg; t.style.display = "block";' +
        'setTimeout(function() { t.style.display = "none"; }, 2000);' +
      '}' +
      'function showHistory(idx) {' +
        'var code = fertOrder[currentPage];' +
        'var group = fertGroups[code];' +
        'var r = group.items[idx];' +
        'if (!r.historyRecords || r.historyRecords.length === 0) {' +
          'alert("해당 항목의 실적 이력이 없습니다.");' +
          'return;' +
        '}' +
        'document.getElementById("modalTitle").textContent = r.bulkName + " 실적 이력";' +
        'document.getElementById("modalSub").textContent = "성형물: " + r.moldCode + " " + r.moldName + " / 벌크: " + r.bulkCode;' +
        // 가중평균 계산 과정 텍스트 생성
        'var calcText = "";' +
        'if (r.avgLossRate !== null && r.historyRecords) {' +
          'var sortedR = r.historyRecords.slice().sort(function(a,b){ return (b.prodDate||"").localeCompare(a.prodDate||""); });' +
          'var validList = [];' +
          'for (var c = 0; c < sortedR.length && validList.length < 5; c++) {' +
            'var rcc = sortedR[c];' +
            'if (rcc.stdNeed <= 3000) continue;' +
            'var uinn = rcc.adjustedInput || rcc.actualInput;' +
            'var lrr = rcc.stdNeed > 0 ? ((uinn - rcc.stdNeed) / rcc.stdNeed * 100) : 0;' +
            'if (lrr >= -50 && lrr <= 200) { validList.push({ date: rcc.prodDate, rate: lrr }); }' +
          '}' +
          'var lines = ["[가중평균 계산 과정]", ""];' +
          'var ws = 0; var tw = 0;' +
          'for (var c = 0; c < validList.length; c++) {' +
            'var w = validList.length - c;' +
            'lines.push((c+1) + ". " + (validList[c].date || "-") + "  " + validList[c].rate.toFixed(2) + "%  ×  가중치 " + w + "  =  " + (validList[c].rate * w).toFixed(2));' +
            'ws += validList[c].rate * w;' +
            'tw += w;' +
          '}' +
          'lines.push("");' +
          'lines.push("합계 " + ws.toFixed(2) + " ÷ 가중치합 " + tw + " = " + (ws/tw).toFixed(2) + "%");' +
          'calcText = lines.join("\\n");' +
        '}' +
        'var helpIcon = calcText ? "<span class=\\"help-icon\\">?<span class=\\"help-tip\\">" + calcText.replace(/</g, "&lt;").replace(/>/g, "&gt;") + "</span></span>" : "";' +
        'document.getElementById("modalSummary").innerHTML = "총 <span>" + r.historyRecords.length + "건</span> | 로스율 <span>" + (r.avgLossRate !== null ? r.avgLossRate.toFixed(2) + "%" : "-") + "</span> <span style=\\"color:#c8102e;font-size:11px;margin-left:8px\\">* 최근 유효 5건 가중평균 (최신=가중치 5, 가장 오래=가중치 1)</span>" + helpIcon;' +
        'var html = "<table><thead><tr><th>제조일</th><th>지시번호</th><th>지시수량(ea)</th><th>실적수량(ea)</th><th>표준소요량(g)</th><th>투입소요량(g)</th><th>환입/폐기</th><th>보정 후(g)</th><th>손실량(g)</th><th>가중치</th><th>설비</th></tr></thead><tbody>";' +
        'var sorted = r.historyRecords.slice().sort(function(a,b){ return (b.prodDate||"").localeCompare(a.prodDate||""); });' +
        // 1차 패스: 유효한 이력 건수 카운트 (최대 5)
        'var totalValid = 0;' +
        'for (var hh = 0; hh < sorted.length && totalValid < 5; hh++) {' +
          'var rrec = sorted[hh];' +
          'if (rrec.stdNeed <= 3000) continue;' +
          'var uin = rrec.adjustedInput || rrec.actualInput;' +
          'var lr = rrec.stdNeed > 0 ? ((uin - rrec.stdNeed) / rrec.stdNeed * 100) : 0;' +
          'if (lr >= -50 && lr <= 200) totalValid++;' +
        '}' +
        'var validCount = 0;' +
        'for (var h = 0; h < sorted.length; h++) {' +
          'var rec = sorted[h];' +
          'if (rec.stdNeed <= 3000) continue;' +
          'var returnText = "-";' +
          'var returnColor = "";' +
          'if (rec.returnInfo && rec.returnInfo.length > 0) {' +
            'var parts = [];' +
            'for (var ri = 0; ri < rec.returnInfo.length; ri++) {' +
              'var info = rec.returnInfo[ri];' +
              'if (info.type === "환입") { parts.push("환입 -" + Math.round(info.qty).toLocaleString() + "g"); }' +
              'else if (info.type === "폐기") { parts.push("폐기 " + Math.round(info.qty).toLocaleString() + "g"); }' +
            '}' +
            'returnText = parts.join(", ");' +
            'returnColor = "color:#4e79a7;font-weight:600";' +
          '}' +
          'var useInput = rec.adjustedInput || rec.actualInput;' +
          'var lossQty = useInput - rec.stdNeed;' +
          'var lossRate = rec.stdNeed > 0 ? ((useInput - rec.stdNeed) / rec.stdNeed * 100) : 0;' +
          'var lossColor = lossRate > 5 ? "color:#c8102e;font-weight:700" : "";' +
          'var adjustedDiff = (rec.adjustedInput && rec.adjustedInput !== rec.actualInput);' +
          'var adjustedStyle = adjustedDiff ? "color:#4e79a7;font-weight:600" : "";' +
          'var isUsed = false; var weightLabel = "-";' +
          'var isFirstUsed = false; var isLastUsed = false;' +
          'if (validCount < totalValid && lossRate >= -50 && lossRate <= 200) {' +
            'isUsed = true; validCount++; weightLabel = String(totalValid - validCount + 1);' +
            'if (validCount === 1) isFirstUsed = true;' +
            'if (validCount === totalValid) isLastUsed = true;' +
          '}' +
          // 그룹 외곽 테두리 스타일 (각 셀 단위로 적용)
          'var bg = isUsed ? "background:#fff5f5;" : "";' +
          'var bTop = isFirstUsed ? "border-top:2.5px solid #c8102e;" : "";' +
          'var bBot = isLastUsed ? "border-bottom:2.5px solid #c8102e;" : "";' +
          'var bLeft = isUsed ? "border-left:2.5px solid #c8102e;" : "";' +
          'var bRight = isUsed ? "border-right:2.5px solid #c8102e;" : "";' +
          'var firstCellStyle = bg + bTop + bBot + bLeft;' +
          'var lastCellStyle = bg + bTop + bBot + bRight;' +
          'var midCellStyle = bg + bTop + bBot;' +
          'var weightStyle = midCellStyle + (isUsed ? "color:#c8102e;font-weight:700" : "color:#bbb");' +
          'html += "<tr>" +' +
            '"<td style=\\"" + firstCellStyle + "\\">" + (rec.prodDate || "-") + "</td>" +' +
            '"<td style=\\"" + midCellStyle + "\\">" + (rec.prodOrder || "-") + "</td>" +' +
            '"<td class=\\"num\\" style=\\"" + midCellStyle + "\\">" + (rec.orderQty ? Math.round(rec.orderQty).toLocaleString() : "-") + "</td>" +' +
            '"<td class=\\"num\\" style=\\"" + midCellStyle + "\\">" + (rec.actualQty ? Math.round(rec.actualQty).toLocaleString() : "-") + "</td>" +' +
            '"<td class=\\"num\\" style=\\"" + midCellStyle + "\\">" + (rec.stdNeed ? Math.round(rec.stdNeed).toLocaleString() : "-") + "</td>" +' +
            '"<td class=\\"num\\" style=\\"" + midCellStyle + "\\">" + (rec.actualInput ? Math.round(rec.actualInput).toLocaleString() : "-") + "</td>" +' +
            '"<td class=\\"num\\" style=\\"" + midCellStyle + returnColor + "\\">" + returnText + "</td>" +' +
            '"<td class=\\"num\\" style=\\"" + midCellStyle + adjustedStyle + "\\">" + Math.round(useInput).toLocaleString() + "</td>" +' +
            '"<td class=\\"num\\" style=\\"" + midCellStyle + lossColor + "\\">" + Math.round(lossQty).toLocaleString() + " (" + lossRate.toFixed(1) + "%)</td>" +' +
            '"<td style=\\"" + weightStyle + "\\">" + weightLabel + "</td>" +' +
            '"<td style=\\"" + lastCellStyle + "\\">" + (rec.machine || "-") + "</td>" +' +
          '"</tr>";' +
        '}' +
        'html += "</tbody></table>";' +
        'document.getElementById("modalBody").innerHTML = html;' +
        'document.getElementById("modalOverlay").style.display = "block";' +
      '}' +
      'function closeModal() {' +
        'document.getElementById("modalOverlay").style.display = "none";' +
      '}' +
      'function downloadExcel() {' +
        'var code = fertOrder[currentPage];' +
        'var group = fertGroups[code];' +
        'var headers = ["성형물 코드","성형물명","벌크 코드","벌크명","투입량(g)","필요 수량(ea)","가용 재고(ea)","최종 수량(ea)","이론 필요량(g)","평균 로스율","최적 제조량(g)","신뢰도","신뢰도 점수","유효 표본수","편차(%)","최신 이력(일전)"];' +
        'var rows = [headers];' +
        'var confLabelMap = { high: "높음", medium: "보통", low: "낮음", verylow: "매우 낮음", none: "-" };' +
        'for (var i = 0; i < group.items.length; i++) {' +
          'var r = group.items[i];' +
          'var confLabel = confLabelMap[r.confidenceLevel] || "-";' +
          'rows.push([' +
            'r.moldCode, r.moldName, r.bulkCode, r.bulkName,' +
            'Math.round(r.bulkInputPerUnit * 1000) / 1000,' +
            'Math.round(r.moldNeedQtyOriginal), (r.shadeAvailable > 0 ? Math.round(r.shadeAvailable) : "-"), Math.round(r.moldNeedQty), Math.round(r.bulkTheoryNeed),' +
            '(r.avgLossRate !== null ? r.avgLossRate.toFixed(2) + "%" : "-"),' +
            '(r.optimalQty !== null ? Math.round(r.optimalQty) : "-"),' +
            'confLabel,' +
            '(r.confidenceLevel !== "none" ? r.confidenceScore : "-"),' +
            '(r.confidenceLevel !== "none" ? r.validHistoryCount : "-"),' +
            '(r.confidenceLevel !== "none" ? r.stdDev.toFixed(2) : "-"),' +
            '(r.recencyDays !== null ? r.recencyDays : "-")' +
          ']);' +
        '}' +
        'var ws = XLSX.utils.aoa_to_sheet(rows);' +
        'ws["!cols"] = [{wch:14},{wch:25},{wch:14},{wch:25},{wch:10},{wch:12},{wch:12},{wch:12},{wch:14},{wch:10},{wch:14},{wch:10},{wch:12},{wch:10},{wch:10},{wch:14}];' +
        'var wb = XLSX.utils.book_new();' +
        'XLSX.utils.book_append_sheet(wb, ws, "벌크별 필요 제조량");' +
        'XLSX.writeFile(wb, "벌크별_필요_제조량_" + new Date().toISOString().slice(0,10) + ".xlsx");' +
      '}' +
      'document.addEventListener("keydown", function(e) {' +
        'if (e.key === "Escape") {' +
          'var modal = document.getElementById("modalOverlay");' +
          'if (modal && modal.style.display !== "none") { modal.style.display = "none"; }' +
        '}' +
      '});' +
    '<\/script>' +
  '</body></html>';

  popup.document.open();
  popup.document.write(html);
  popup.document.close();
}

// ============ 예측 탭으로 보내기 ============
function sendToPredictTab() {
  var results = window._bomCalcResults;
  if (!results || results.length === 0) return;

  // 성형물별로 그룹핑 (같은 성형물에 여러 벌크가 있을 수 있음)
  var moldMap = {};
  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    if (!moldMap[r.moldCode]) {
      moldMap[r.moldCode] = Math.round(r.moldNeedQty);
    }
  }

  // 예측 탭의 입력행 세팅
  var container = document.getElementById('inputRows');
  container.innerHTML = '';

  var codes = Object.keys(moldMap);
  for (var i = 0; i < codes.length; i++) {
    var row = document.createElement('div');
    row.className = 'input-row';
    row.dataset.index = i;
    row.innerHTML =
      '<div class="input-group autocomplete-wrap">' +
        '<label>성형물 코드</label>' +
        '<input type="text" class="mold-code-input" value="' + codes[i] + '" autocomplete="off">' +
        '<div class="autocomplete-list"></div>' +
      '</div>' +
      '<div class="input-group">' +
        '<label>성형 지시 수량</label>' +
        '<input type="text" class="order-qty-input" value="' + moldMap[codes[i]].toLocaleString() + '">' +
      '</div>' +
      '<button class="remove-row-btn" onclick="removeRow(this)" title="삭제">✕</button>';
    container.appendChild(row);
    row.querySelector('.order-qty-input').addEventListener('input', formatQtyInput);
    setupAutocomplete(row.querySelector('.mold-code-input'));
  }

  // 예측 탭으로 전환
  switchTab('predict');

  // SAP 실적 데이터 있으면 자동 예측 실행
  if (typeof sapCount !== 'undefined' && sapCount > 0) {
    document.getElementById('predictBtn').click();
  }
}
