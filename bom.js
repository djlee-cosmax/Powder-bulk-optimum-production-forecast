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
document.getElementById('bomFile').addEventListener('change', function(e) {
  var file = e.target.files[0];
  if (!file) return;

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
        if (colName.indexOf('품질') !== -1) colMap.qualityInsp = h;
        if (colName === '설명1' || colName === '설명 1') colMap.desc1 = h;
        if (colName.indexOf('전체') !== -1 && colName.indexOf('이름') !== -1) colMap.fullName = h;
        if (colName.indexOf('로스율') !== -1 || colName.indexOf('로스') !== -1) colMap.lossRate = h;
        if (colName.indexOf('이름 1') !== -1 || colName.indexOf('공급업체') !== -1) colMap.supplier = h;
      }
      if (colMap.lev !== undefined) {
        headerFound = true;

        // 자재코드: MTyp 다음 컬럼
        if (!colMap.code) colMap.code = (colMap.mtype || 1) + 1;
        if (!colMap.name) colMap.name = colMap.code + 1;

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
      supplier: supplier
    });
  }

  if (bomData.length === 0) {
    alert('BOM 데이터를 찾을 수 없습니다. 파일 형식을 확인해 주세요.');
    return;
  }

  document.getElementById('bomStatus').textContent = bomData.length + '건 로드됨';
  document.getElementById('bomStatus').classList.add('loaded');
  renderBomTree(bomData);
}

// ============ CSV 파싱 ============
function parseBomCsv(text) {
  var result = Papa.parse(text, { header: true, skipEmptyLines: true });
  var rows = result.data;
  var bomData = [];

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
      inputQty: parseFloat(String(r['투입수량'] || '0').replace(/,/g, '')) || 0,
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
      supplier: (r['이름 1'] || '').trim()
    });
  }

  if (bomData.length === 0) {
    alert('BOM 데이터를 찾을 수 없습니다.');
    return;
  }

  document.getElementById('bomStatus').textContent = bomData.length + '건 로드됨';
  document.getElementById('bomStatus').classList.add('loaded');
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
      '<td class="num">' + (d.inputQty ? d.inputQty.toLocaleString() : '-') + '</td>' +
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
      var moldNeedQty = currentFert.orderQty * currentMold.inputQty;
      var bulkTheoryNeed = moldNeedQty * d.inputQty;

      // 최적 제조량 예측 (SAP 실적 데이터 있으면)
      var optimalQty = null;
      var avgLossRate = null;
      var historyCount = 0;
      var historyRecords = [];
      if (typeof moldBulkMap !== 'undefined' && moldBulkMap[currentMold.code]) {
        var predicted = predictOne(currentMold.code, Math.round(moldNeedQty));
        for (var p = 0; p < predicted.length; p++) {
          if (!predicted[p].error && predicted[p].bulkCode === d.code) {
            optimalQty = predicted[p].optimalQty;
            avgLossRate = predicted[p].avgLossRate;
            historyCount = predicted[p].historyCount;
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
        bulkCode: d.code,
        bulkName: d.name,
        bulkInputPerUnit: d.inputQty,
        bulkTheoryNeed: bulkTheoryNeed,
        optimalQty: optimalQty,
        avgLossRate: avgLossRate,
        historyCount: historyCount,
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
      '.modal { position:absolute; top:50%; left:50%; transform:translate(-50%,-50%); background:#fff; border-radius:10px; padding:24px; max-width:90%; max-height:80%; overflow-y:auto; box-shadow:0 8px 30px rgba(0,0,0,0.15); }' +
      '.modal h2 { font-size:16px; margin:0 0 4px; }' +
      '.modal .sub { font-size:12px; color:#999; margin-bottom:14px; }' +
      '.modal table { width:100%; }' +
      '.modal .close-btn { position:absolute; top:12px; right:16px; background:none; border:none; font-size:20px; cursor:pointer; color:#999; }' +
      '.modal .close-btn:hover { color:#333; }' +
      '.modal .summary { display:flex; gap:20px; margin-bottom:14px; font-size:13px; }' +
      '.modal .summary span { font-weight:700; }' +
      '.note { font-size:13px; color:#666; padding:8px 0; }' +
    '</style></head><body>' +
    '<div class="top-bar">' +
      '<div class="nav-area">' +
        '<button class="nav-btn" id="prevBtn" onclick="navPrev()">&lt;</button>' +
        '<h1 id="pageTitle">벌크별 필요 제조량</h1>' +
        '<button class="nav-btn" id="nextBtn" onclick="navNext()">&gt;</button>' +
      '</div>' +
      '<div class="top-actions">' +
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
    '</div>' +
    '<script>' +
      'var fertOrder = ' + JSON.stringify(fertOrder) + ';' +
      'var fertGroups = ' + JSON.stringify(fertGroups) + ';' +
      'var allResults = ' + JSON.stringify(results) + ';' +
      'var currentPage = 0;' +
      'var showML = false;' +
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
        'document.getElementById("pageTitle").textContent = "벌크별 필요 제조량 (" + (currentPage+1) + "/" + fertOrder.length + ")";' +
        'document.getElementById("prevBtn").disabled = currentPage === 0;' +
        'document.getElementById("nextBtn").disabled = currentPage === fertOrder.length - 1;' +
        'var rows = "";' +
        'for (var i = 0; i < group.items.length; i++) {' +
          'var r = group.items[i];' +
          'var mlRate = showML ? getMLPrediction(r.bulkCode, r.bulkTheoryNeed) : null;' +
          'var mlQty = (mlRate !== null) ? Math.ceil(r.bulkTheoryNeed * (1 + mlRate / 100)) : null;' +
          'rows += "<tr ondblclick=\\"showHistory(" + i + ")\\">" +' +
            '"<td>" + r.moldCode + "</td>" +' +
            '"<td class=\\"name\\">" + r.moldName + "</td>" +' +
            '"<td>" + r.bulkCode + "</td>" +' +
            '"<td class=\\"name\\">" + r.bulkName + "</td>" +' +
            '"<td class=\\"num\\">" + r.bulkInputPerUnit + "</td>" +' +
            '"<td class=\\"num\\">" + Math.round(r.moldNeedQty).toLocaleString() + "</td>" +' +
            '"<td class=\\"num highlight\\">" + Math.round(r.bulkTheoryNeed).toLocaleString() + "</td>" +' +
            '"<td class=\\"num\\">" + (r.avgLossRate !== null ? r.avgLossRate.toFixed(1) + "%" : "-") + "</td>" +' +
            '"<td class=\\"num optimal\\">" + (r.optimalQty !== null ? Math.round(r.optimalQty).toLocaleString() : "-") + "</td>" +' +
            '(showML ? "<td class=\\"num ml-col\\">" + (mlRate !== null ? mlRate.toFixed(1) + "%" : "-") + "</td>" +' +
            '"<td class=\\"num ml-col\\">" + (mlQty !== null ? mlQty.toLocaleString() : "-") + "</td>" : "") +' +
          '"</tr>";' +
        '}' +
        'var mlHeaders = showML ? "<th class=\\"ml-col\\">ML 로스율</th><th class=\\"ml-col\\">ML 제조량(g)</th>" : "";' +
        'document.getElementById("tableArea").innerHTML = "<table><thead><tr>" +' +
          '"<th>성형물 코드</th><th>성형물명</th><th>벌크 코드</th><th>벌크명</th><th>투입량(g)</th><th>필요 수량(ea)</th><th>이론 필요량(g)</th><th>평균 로스율</th><th>최적 제조량(g)</th>" + mlHeaders +' +
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
        'document.getElementById("modalSummary").innerHTML = "총 <span>" + r.historyRecords.length + "건</span> | 로스율 <span>" + (r.avgLossRate !== null ? r.avgLossRate.toFixed(1) + "%" : "-") + "</span> <span style=\\"color:#c8102e;font-size:11px;margin-left:8px\\">* 최신 이력 1건 기준</span>";' +
        'var html = "<table><thead><tr><th>제조일</th><th>지시번호</th><th>지시수량(ea)</th><th>실적수량(ea)</th><th>표준소요량(g)</th><th>투입소요량(g)</th><th>환입/폐기</th><th>보정 후(g)</th><th>손실량(g)</th><th>설비</th></tr></thead><tbody>";' +
        'var sorted = r.historyRecords.slice().sort(function(a,b){ return (b.prodDate||"").localeCompare(a.prodDate||""); });' +
        'var latestFound = false;' +
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
          'var isLatest = false;' +
          'if (!latestFound && lossRate >= -50 && lossRate <= 200) { isLatest = true; latestFound = true; }' +
          'var rowStyle = isLatest ? "outline:2.5px solid #c8102e;outline-offset:-1px;background:#fff5f5;" : "";' +
          'html += "<tr style=\\"" + rowStyle + "\\">" +' +
            '"<td>" + (rec.prodDate || "-") + "</td>" +' +
            '"<td>" + (rec.prodOrder || "-") + "</td>" +' +
            '"<td class=\\"num\\">" + (rec.orderQty ? Math.round(rec.orderQty).toLocaleString() : "-") + "</td>" +' +
            '"<td class=\\"num\\">" + (rec.actualQty ? Math.round(rec.actualQty).toLocaleString() : "-") + "</td>" +' +
            '"<td class=\\"num\\">" + (rec.stdNeed ? Math.round(rec.stdNeed).toLocaleString() : "-") + "</td>" +' +
            '"<td class=\\"num\\">" + (rec.actualInput ? Math.round(rec.actualInput).toLocaleString() : "-") + "</td>" +' +
            '"<td class=\\"num\\" style=\\"" + returnColor + "\\">" + returnText + "</td>" +' +
            '"<td class=\\"num\\" style=\\"" + adjustedStyle + "\\">" + Math.round(useInput).toLocaleString() + "</td>" +' +
            '"<td class=\\"num\\" style=\\"" + lossColor + "\\">" + Math.round(lossQty).toLocaleString() + " (" + lossRate.toFixed(1) + "%)</td>" +' +
            '"<td>" + (rec.machine || "-") + "</td>" +' +
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
        'var headers = ["성형물 코드","성형물명","벌크 코드","벌크명","투입량(g)","필요 수량(ea)","이론 필요량(g)","평균 로스율","최적 제조량(g)"];' +
        'var rows = [headers];' +
        'for (var i = 0; i < group.items.length; i++) {' +
          'var r = group.items[i];' +
          'rows.push([r.moldCode, r.moldName, r.bulkCode, r.bulkName, r.bulkInputPerUnit, Math.round(r.moldNeedQty), Math.round(r.bulkTheoryNeed), (r.avgLossRate !== null ? r.avgLossRate.toFixed(1) + "%" : "-"), (r.optimalQty !== null ? Math.round(r.optimalQty) : "-")]);' +
        '}' +
        'var ws = XLSX.utils.aoa_to_sheet(rows);' +
        'ws["!cols"] = [{wch:14},{wch:25},{wch:14},{wch:25},{wch:10},{wch:12},{wch:14},{wch:10},{wch:14}];' +
        'var wb = XLSX.utils.book_new();' +
        'XLSX.utils.book_append_sheet(wb, ws, "벌크별 필요 제조량");' +
        'XLSX.writeFile(wb, "벌크별_필요_제조량_" + new Date().toISOString().slice(0,10) + ".xlsx");' +
      '}' +
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
