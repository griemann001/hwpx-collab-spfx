/**
 * ============================================================
 * HwpxPatcher.js v3 — 스타일 인식 원본 XML 패치 (TipTap 호환)
 * ============================================================
 *
 * v3 변경사항:
 *   - extractHtmlCellStyledLines(): TipTap HTML 지원
 *     - <p> 태그를 줄 구분자로 인식 (기존 <br> 포함)
 *     - <strong> → bold (TipTap 기본 출력)
 *     - <em> → italic (향후 확장 대비, 현재 bold만 HWPX 매핑)
 *     - <u> → underline (향후 확장 대비)
 *     - <span style="color:..."> → color (TipTap Color 확장 출력)
 *     - <span style="font-size:..."> → fontSize (TipTap FontSize 확장 출력)
 *   - App.jsx의 extractStyledLinesFromHTML()과 연동
 *     (div 컨테이너에 TipTap HTML을 담아 전달)
 *
 * v2 유지:
 *   - patchParagraphStyled(): run별 charPrIDRef 매핑
 *   - charPr 매핑 테이블
 */

/**
 * HTML 셀에서 줄 배열 추출 (텍스트만, 하위 호환)
 */
export function extractHtmlCellLines(td) {
  const styled = extractHtmlCellStyledLines(td);
  return styled.map(line => line.runs.map(r => r.text).join(""));
}

/**
 * HTML 셀 또는 div 컨테이너에서 스타일 포함 줄 배열 추출
 *
 * 반환: [{ runs: [{ text, bold, italic, underline, color, fontSize }] }, ...]
 *
 * 입력 HTML 소스 (자동 감지):
 *   - TipTap 출력: <p><strong>...</strong></p> (줄 = <p>)
 *   - contentEditable 레거시: <br> 줄 구분, <b>/<font color>
 *
 * TipTap 스타일 매핑:
 *   <strong>           → bold
 *   <em>               → italic
 *   <u>                → underline
 *   <span style="color:...">     → color
 *   <span style="font-size:XXpt"> → fontSize
 */
export function extractHtmlCellStyledLines(td) {
  const clone = td.cloneNode(true);
  // TipTap collaboration cursor 등 제거
  clone.querySelectorAll(
    "svg, span[style*='color:#EC5F00'], .collab-badge-container, .collaboration-cursor__caret, .collaboration-cursor__label"
  ).forEach(el => el.remove());

  const lines = []; // [{ runs: [{ text, bold, color, fontSize }] }]
  let currentRuns = [];

  // 현재 스타일 스택
  function getStyle(node) {
    let bold      = false;
    let italic    = false;
    let underline = false;
    let color     = "";
    let fontSize  = 0;

    let el = node.nodeType === 1 ? node : node.parentElement;
    while (el && el !== clone) {
      const tag = el.tagName?.toLowerCase();
      // ── bold ──
      if (tag === "b" || tag === "strong") bold = true;
      // ── italic ──
      if (tag === "i" || tag === "em") italic = true;
      // ── underline ──
      if (tag === "u") underline = true;
      // ── inline style ──
      if (el.style) {
        if (el.style.fontWeight === "bold" || parseInt(el.style.fontWeight || 0) >= 700) bold = true;
        if (el.style.fontStyle === "italic") italic = true;
        if (el.style.textDecoration === "underline" || el.style.textDecorationLine === "underline") underline = true;
        if (el.style.color && !color) color = normalizeColor(el.style.color);
        if (el.style.fontSize && !fontSize) fontSize = parsePt(el.style.fontSize);
      }
      // <font color="..."> (execCommand 레거시)
      if (tag === "font" && el.getAttribute("color") && !color) {
        color = normalizeColor(el.getAttribute("color"));
      }
      el = el.parentElement;
    }
    return { bold, italic, underline, color, fontSize };
  }

  let currentAlign = 'left';  // 현재 블록의 정렬값

  function flush() {
    lines.push({
      align: currentAlign,
      runs: currentRuns.length > 0 ? currentRuns : [{ text: "", bold: false, color: "", fontSize: 0 }],
    });
    currentRuns = [];
  }

  function walk(node) {
    if (node.nodeType === 3) {
      const text = (node.textContent || "").replace(/\u00a0/g, " ");
      if (text) {
        const style = getStyle(node);
        const lastRun = currentRuns[currentRuns.length - 1];
        if (
          lastRun &&
          lastRun.bold      === style.bold      &&
          lastRun.italic    === style.italic     &&
          lastRun.underline === style.underline  &&
          lastRun.color     === style.color      &&
          lastRun.fontSize  === style.fontSize
        ) {
          lastRun.text += text;
        } else {
          currentRuns.push({ text, ...style });
        }
      }
    } else if (node.nodeType === 1) {
      const tag = node.tagName.toLowerCase();
      if (tag === "br") { flush(); return; }
      if (tag === "svg" || tag === "img") return;

      const isBlock = ["div", "p", "li"].includes(tag);
      if (isBlock) {
        if (currentRuns.length > 0) flush();
        // ★ <p style="text-align:center"> 또는 data-text-align 읽기
        const ta = node.style?.textAlign || node.getAttribute?.("data-text-align") || "left";
        currentAlign = ta || "left";
        for (let i = 0; i < node.childNodes.length; i++) walk(node.childNodes[i]);
        flush();
        currentAlign = "left";
      } else {
        for (let i = 0; i < node.childNodes.length; i++) walk(node.childNodes[i]);
      }
    }
  }

  for (let i = 0; i < clone.childNodes.length; i++) walk(clone.childNodes[i]);
  if (currentRuns.length > 0) flush();

  // 맨 뒤 빈 줄 정리
  while (lines.length > 1) {
    const last = lines[lines.length - 1];
    if (last.runs.length === 1 && last.runs[0].text === "") lines.pop();
    else break;
  }

  if (lines.length === 0) lines.push({ runs: [{ text: "", bold: false, color: "", fontSize: 0 }] });
  return lines;
}

/**
 * 색상 문자열 정규화 → #RRGGBB
 */
function normalizeColor(color) {
  if (!color) return "";
  if (color.startsWith("#")) {
    // #RGB → #RRGGBB
    if (color.length === 4) {
      return "#" + color[1]+color[1] + color[2]+color[2] + color[3]+color[3];
    }
    return color.toUpperCase();
  }
  const m = color.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
  if (m) return "#" + [m[1],m[2],m[3]].map(x => parseInt(x).toString(16).padStart(2,"0")).join("").toUpperCase();
  return color.toUpperCase();
}

/**
 * fontSize 문자열 → pt 숫자
 */
function parsePt(str) {
  if (!str) return 0;
  const n = parseFloat(str);
  if (str.endsWith("pt")) return Math.round(n);
  if (str.endsWith("px")) return Math.round(n * 0.75);
  return Math.round(n);
}


// ============================================================
// charPrIDRef 매핑
// ============================================================

/**
 * HTML 스타일을 원본 HWPX의 charPrIDRef에 매핑
 * 
 * 매핑 전략:
 *   1. 원본 header.xml에서 charPr 목록을 파싱하여 매핑 테이블 생성 (파싱 시)
 *   2. 저장 시 HTML의 bold/color/fontSize를 매핑 테이블에서 찾기
 *   3. 정확히 일치하는 charPr이 없으면 가장 가까운 것 선택
 *   4. 그래도 없으면 기본값(원본 셀의 charPrIDRef) 유지
 * 
 * @param {Array<{id, fontSize, color, bold}>} charPrList - header.xml에서 파싱한 charPr 목록
 * @param {Object} style - { bold, color, fontSize }
 * @param {string} defaultRef - 매칭 실패 시 기본값
 * @returns {string} charPrIDRef
 */
export function findCharPrId(charPrList, style, defaultRef) {
  if (!charPrList || charPrList.length === 0) return defaultRef;

  const { bold, color, fontSize } = style;
  const normColor = normalizeColor(color);

  // 스타일이 전혀 없으면 (기본 10pt 검정, 비굵게) → 기본값
  if (!bold && (!normColor || normColor === "#000000") && (!fontSize || fontSize === 10)) {
    return defaultRef;
  }

  // ★ exact match: color + fontSize + bold 3속성 모두 일치하는 charPr만 반환
  var targetColor = normColor || "#000000";
  var targetSize = (fontSize && fontSize > 0) ? fontSize : 10;
  var targetBold = !!bold;

  for (var i = 0; i < charPrList.length; i++) {
    var cp = charPrList[i];
    var cpColor = normalizeColor(cp.color) || "#000000";
    var cpSize = cp.fontSize || 10;
    var cpBold = !!cp.bold;

    if (cpColor === targetColor && cpSize === targetSize && cpBold === targetBold) {
      return cp.id;
    }
  }

  // 정확 일치 없으면 defaultRef 반환 (ensureCharPrIds에서 이미 새 charPr 추가됨)
  return defaultRef;
}


/**
 * 편집된 셀에서 사용된 스타일 중 원본 charPrList에 없는 것을 추가
 * charPrList를 변이(mutate)하고, 새로 추가된 charPr의 ID를 반환
 *
 * @param {Array} charPrList - 기존 charPr 목록 (mutate됨)
 * @param {Array} cellEdits - 편집된 셀 목록
 * @returns {Array<{id, bold, color, fontSize}>} 새로 추가된 charPr 목록
 */
export function ensureCharPrIds(charPrList, cellEdits) {
  if (!charPrList || !cellEdits) return [];

  const added = [];
  let maxId = 0;
  for (const cp of charPrList) {
    const n = parseInt(cp.id, 10);
    if (n > maxId) maxId = n;
  }

  // 편집에서 사용된 모든 스타일 수집
  const needed = new Map(); // key → style
  for (const edit of cellEdits) {
    if (!edit.styledLines) continue;
    for (const line of edit.styledLines) {
      for (const run of line.runs) {
        if (!run.text) continue;
        const color = normalizeColor(run.color);
        const bold = !!run.bold;
        const fontSize = run.fontSize || 10;
        const italic = !!run.italic;
        const underline = !!run.underline;

        // 기본 스타일이면 스킵
        if (!bold && !italic && !underline && (!color || color === '#000000') && fontSize === 10) continue;

        const key = [bold, italic, underline, color, fontSize].join('|');
        if (!needed.has(key)) {
          needed.set(key, { bold, italic, underline, color, fontSize });
        }
      }
    }
  }

  console.log('[ensureCharPrIds] needed:', needed.size, '개', [...needed.entries()].map(function(e) { return e[0] + ' → ' + JSON.stringify(e[1]); }));

  // 필요한 스타일 중 charPrList에 없는 것 찾기
  needed.forEach(function(style, key) {
    const existing = findCharPrId(charPrList, style, null);
    console.log('[ensureCharPrIds] key=' + key + ' existing=' + existing);
    if (existing !== null) {
      // 기존 charPr에서 정확히 일치하는지 한번 더 확인
      const cp = charPrList.find(function(c) { return c.id === existing; });
      if (cp) {
        const cpColor = normalizeColor(cp.color);
        const styleColor = normalizeColor(style.color);
        if (cp.bold === style.bold && cpColor === styleColor && cp.fontSize === style.fontSize) {
          return; // 정확히 일치 → 추가 불필요 (forEach에서는 return = continue)
        }
      }
    }

    // 새 charPr 추가
    maxId++;
    const newId = String(maxId);
    const newCp = {
      id: newId,
      bold: style.bold,
      italic: style.italic || false,
      underline: style.underline || false,
      color: style.color || '#000000',
      fontSize: style.fontSize || 10,
    };
    charPrList.push(newCp);
    added.push(newCp);
  });

  return added;
}

/**
 * 편집된 셀/블록에서 사용된 정렬 중 원본 paraPrList에 없는 것을 추가
 *
 * @param {Array} paraPrList - 기존 paraPr 목록 [{id, align}] (mutate됨)
 * @param {Array} allEdits - 편집 목록 (cellEdits + blockEdits)
 * @returns {Array<{id, align}>} 새로 추가된 paraPr 목록
 */
export function ensureParaPrIds(paraPrList, allEdits) {
  if (!paraPrList || !allEdits) return [];

  var alignMap = { left: "LEFT", center: "CENTER", right: "RIGHT", justify: "JUSTIFY" };

  // 기존 paraPrList에서 지원하는 정렬 확인
  var existingAligns = {};
  for (var i = 0; i < paraPrList.length; i++) {
    var key = Object.keys(alignMap).find(function(k) { return alignMap[k] === paraPrList[i].align; });
    if (key) existingAligns[key] = paraPrList[i].id;
  }

  // 편집에서 사용된 정렬 수집
  var neededAligns = {};
  for (var ei = 0; ei < allEdits.length; ei++) {
    var edit = allEdits[ei];
    if (!edit.styledLines) continue;
    for (var li = 0; li < edit.styledLines.length; li++) {
      var line = edit.styledLines[li];
      if (line.align && line.align !== 'left' && !existingAligns[line.align]) {
        neededAligns[line.align] = true;
      }
    }
  }

  // maxId 계산
  var maxId = 0;
  for (var pi = 0; pi < paraPrList.length; pi++) {
    var n = parseInt(paraPrList[pi].id, 10);
    if (n > maxId) maxId = n;
  }

  var added = [];
  var alignKeys = Object.keys(neededAligns);
  for (var ai = 0; ai < alignKeys.length; ai++) {
    var alignKey = alignKeys[ai];
    var hwpxAlign = alignMap[alignKey];
    if (!hwpxAlign) continue;

    maxId++;
    var newPp = { id: String(maxId), align: hwpxAlign };
    paraPrList.push(newPp);
    added.push(newPp);
  }

  return added;
}

/**
 * header.xml에 새 charPr 엘리먼트를 추가
 *
 * @param {string} headerXml - header.xml 원본 문자열
 * @param {Array<{id, bold, italic, underline, color, fontSize}>} newCharPrs
 * @returns {string} 패치된 header.xml
 */
export function patchHeaderXml(headerXml, newCharPrs) {
  if (!newCharPrs || newCharPrs.length === 0) return headerXml;

  const doc = new DOMParser().parseFromString(headerXml, "text/xml");
  const root = doc.documentElement;

  // charProperties 엘리먼트 찾기
  const all = root.getElementsByTagName("*");
  let charPropsEl = null;
  for (let i = 0; i < all.length; i++) {
    const tag = all[i].localName || all[i].tagName.split(":").pop();
    if (tag === "charProperties") { charPropsEl = all[i]; break; }
  }

  console.log('[patchHeaderXml] charPropsEl found:', !!charPropsEl, 'newCharPrs:', newCharPrs.length);
  if (!charPropsEl) {
    console.log('[patchHeaderXml] ★ charProperties를 찾지 못함! 원본 반환');
    return headerXml;
  }

  // 네임스페이스 확인
  const HH = "http://www.hancom.co.kr/hwpml/2011/head";

  // ★ 기존 charPr[0]을 템플릿으로 사용 — fontRef, ratio, spacing 등 필수 하위 요소 복사
  var templateCharPr = null;
  for (var ci = 0; ci < charPropsEl.childNodes.length; ci++) {
    var nd = charPropsEl.childNodes[ci];
    if (nd.nodeType === 1) {
      var ndTag = nd.localName || nd.tagName.split(":").pop();
      if (ndTag === "charPr") { templateCharPr = nd; break; }
    }
  }

  for (const cp of newCharPrs) {
    // 템플릿 charPr을 deep clone하여 시작 (fontRef, ratio, spacing 등 포함)
    var el;
    if (templateCharPr) {
      el = templateCharPr.cloneNode(true);
      // bold, italic, underline 등 스타일 태그는 제거 (새로 추가할 것이므로)
      var toRemove = [];
      for (var ri = 0; ri < el.childNodes.length; ri++) {
        var child = el.childNodes[ri];
        if (child.nodeType === 1) {
          var ctag = child.localName || child.tagName.split(":").pop();
          if (ctag === "bold" || ctag === "italic" || ctag === "underline" || ctag === "strikeout") {
            toRemove.push(child);
          }
        }
      }
      for (var ri = 0; ri < toRemove.length; ri++) {
        el.removeChild(toRemove[ri]);
      }
    } else {
      el = doc.createElementNS(HH, "hh:charPr");
    }

    // 속성 설정 (clone된 속성을 덮어씀)
    el.setAttribute("id", cp.id);
    el.setAttribute("height", String(Math.round(cp.fontSize * 100)));
    el.setAttribute("textColor", cp.color || "#000000");

    if (cp.bold) {
      var boldEl = doc.createElementNS(HH, "hh:bold");
      el.appendChild(boldEl);
    }
    if (cp.italic) {
      var italicEl = doc.createElementNS(HH, "hh:italic");
      el.appendChild(italicEl);
    }

    // ★ underline: 스타일 지정 시 BOTTOM, 아니면 NONE (원본 charPr과 동일 구조 유지)
    var hasUnderline = false;
    for (var ui = 0; ui < el.childNodes.length; ui++) {
      var uc = el.childNodes[ui];
      if (uc.nodeType === 1 && (uc.localName || uc.tagName.split(":").pop()) === "underline") {
        hasUnderline = true;
        if (cp.underline) {
          uc.setAttribute("type", "BOTTOM");
          uc.setAttribute("shape", "SOLID");
          uc.setAttribute("color", "#000000");
        } else {
          uc.setAttribute("type", "NONE");
          uc.setAttribute("shape", "SOLID");
          uc.setAttribute("color", "#000000");
        }
        break;
      }
    }
    if (!hasUnderline) {
      var ulEl = doc.createElementNS(HH, "hh:underline");
      ulEl.setAttribute("type", cp.underline ? "BOTTOM" : "NONE");
      ulEl.setAttribute("shape", "SOLID");
      ulEl.setAttribute("color", "#000000");
      el.appendChild(ulEl);
    }

    // ★ strikeout: 항상 NONE으로 추가 (원본 charPr 구조와 일치)
    var hasStrikeout = false;
    for (var si = 0; si < el.childNodes.length; si++) {
      var sc = el.childNodes[si];
      if (sc.nodeType === 1 && (sc.localName || sc.tagName.split(":").pop()) === "strikeout") {
        hasStrikeout = true;
        sc.setAttribute("shape", "NONE");
        sc.setAttribute("color", "#000000");
        break;
      }
    }
    if (!hasStrikeout) {
      var skEl = doc.createElementNS(HH, "hh:strikeout");
      skEl.setAttribute("shape", "NONE");
      skEl.setAttribute("color", "#000000");
      el.appendChild(skEl);
    }

    charPropsEl.appendChild(el);
  }

  // itemCnt 속성 업데이트
  var currentCnt = parseInt(charPropsEl.getAttribute("itemCnt") || "0", 10);
  charPropsEl.setAttribute("itemCnt", String(currentCnt + newCharPrs.length));

  const result = new XMLSerializer().serializeToString(doc);
  console.log('[patchHeaderXml] 원본 길이:', headerXml.length, '→ 결과 길이:', result.length);
  return result;
}

/**
 * header.xml에 새 paraPr 엘리먼트를 추가 (정렬 변경 시 필요)
 *
 * @param {string} headerXml - header.xml 원본 문자열
 * @param {Array<{id, align}>} newParaPrs - 추가할 paraPr 목록
 * @returns {string} 패치된 header.xml
 */
export function patchHeaderXmlParaPr(headerXml, newParaPrs) {
  if (!newParaPrs || newParaPrs.length === 0) return headerXml;

  var doc = new DOMParser().parseFromString(headerXml, "text/xml");
  var root = doc.documentElement;

  // paraProperties 엘리먼트 찾기
  var all = root.getElementsByTagName("*");
  var paraPropsEl = null;
  for (var i = 0; i < all.length; i++) {
    var tag = all[i].localName || all[i].tagName.split(":").pop();
    if (tag === "paraProperties") { paraPropsEl = all[i]; break; }
  }

  if (!paraPropsEl) return headerXml;

  var HH = "http://www.hancom.co.kr/hwpml/2011/head";
  var HC = "http://www.hancom.co.kr/hwpml/2011/core";

  // 기존 paraPr[0]을 템플릿으로 사용
  var templateParaPr = null;
  for (var ci = 0; ci < paraPropsEl.childNodes.length; ci++) {
    var nd = paraPropsEl.childNodes[ci];
    if (nd.nodeType === 1) {
      var ndTag = nd.localName || nd.tagName.split(":").pop();
      if (ndTag === "paraPr") { templateParaPr = nd; break; }
    }
  }

  for (var pi = 0; pi < newParaPrs.length; pi++) {
    var pp = newParaPrs[pi];
    var el;

    if (templateParaPr) {
      el = templateParaPr.cloneNode(true);
    } else {
      el = doc.createElementNS(HH, "hh:paraPr");
    }

    el.setAttribute("id", pp.id);

    // align 엘리먼트 찾아서 horizontal 변경
    var alignEl = null;
    for (var ai = 0; ai < el.childNodes.length; ai++) {
      var ac = el.childNodes[ai];
      if (ac.nodeType === 1 && (ac.localName || ac.tagName.split(":").pop()) === "align") {
        alignEl = ac;
        break;
      }
    }
    if (alignEl) {
      alignEl.setAttribute("horizontal", pp.align);
    } else {
      alignEl = doc.createElementNS(HH, "hh:align");
      alignEl.setAttribute("horizontal", pp.align);
      alignEl.setAttribute("vertical", "BASELINE");
      el.insertBefore(alignEl, el.firstChild);
    }

    paraPropsEl.appendChild(el);
  }

  // itemCnt 업데이트
  var currentCnt = parseInt(paraPropsEl.getAttribute("itemCnt") || "0", 10);
  paraPropsEl.setAttribute("itemCnt", String(currentCnt + newParaPrs.length));

  return new XMLSerializer().serializeToString(doc);
}

/**
 * section XML을 패치하여 셀 텍스트 + 스타일을 업데이트
 *
 * @param {string} xmlStr - section0.xml 원본
 * @param {Array<{tblIdx, rowAddr, colAddr, newLines, styledLines?}>} cellEdits
 * @param {Array} charPrList - header.xml에서 파싱한 charPr 목록
 * @param {Array} paraPrList - header.xml에서 파싱한 paraPr 목록 [{id, align}]
 * @param {Array<{blockIdx, newLines, styledLines?}>} blockEdits - 본문 텍스트 블록 편집
 * @returns {string}
 */
export function patchSectionXml(xmlStr, cellEdits, charPrList, paraPrList = [], blockEdits = [], headerSkip = 0, footerSkip = 0) {
  if (cellEdits.length === 0 && blockEdits.length === 0) return xmlStr;

  const HP = "http://www.hancom.co.kr/hwpml/2011/paragraph";
  const doc = new DOMParser().parseFromString(xmlStr, "text/xml");
  const root = doc.documentElement;

  // 정렬 → paraPrIDRef 매핑 (없으면 원본 id 유지)
  const alignToParaPrId = {};
  const alignMap = { left: "LEFT", center: "CENTER", right: "RIGHT", justify: "JUSTIFY" };
  for (const { id, align } of paraPrList) {
    const key = Object.keys(alignMap).find(k => alignMap[k] === align);
    if (key && !alignToParaPrId[key]) alignToParaPrId[key] = id;
  }

  const allTbls = [];
  function collectTbls(parent, depth) {
    for (let i = 0; i < parent.childNodes.length; i++) {
      const nd = parent.childNodes[i];
      if (nd.nodeType !== 1) continue;
      const tag = nd.localName || nd.tagName.split(":").pop();
      if (tag === "tbl") {
        allTbls.push({ el: nd, depth });
        collectTbls(nd, depth + 1);
      } else {
        collectTbls(nd, depth);
      }
    }
  }
  collectTbls(root, 0);

  let patchCount = 0;
  for (const edit of cellEdits) {
    const { tblIdx, rowAddr, colAddr, newLines, styledLines } = edit;
    if (tblIdx >= allTbls.length) continue;

    const tbl = allTbls[tblIdx].el;
    const tc = findCell(tbl, rowAddr, colAddr);
    if (!tc) continue;

    const subEls = [];
    for (let i = 0; i < tc.childNodes.length; i++) {
      const nd = tc.childNodes[i];
      if (nd.nodeType === 1) {
        const tag = nd.localName || nd.tagName.split(":").pop();
        if (tag === "subList") subEls.push(nd);
      }
    }
    const source = subEls.length > 0 ? subEls[0] : tc;

    const origPs = [];
    for (let i = 0; i < source.childNodes.length; i++) {
      const nd = source.childNodes[i];
      if (nd.nodeType === 1) {
        const tag = nd.localName || nd.tagName.split(":").pop();
        if (tag === "p") origPs.push(nd);
      }
    }

    if (origPs.length === 0) continue;

    const lineCount = newLines.length;
    const origCount = origPs.length;

    // 기존 줄 패치
    for (let li = 0; li < Math.min(lineCount, origCount); li++) {
      if (styledLines && styledLines[li]) {
        patchParagraphStyled(origPs[li], styledLines[li], charPrList, HP, alignToParaPrId);
      } else {
        patchParagraphText(origPs[li], newLines[li], HP);
      }
    }

    // 줄 추가
    if (lineCount > origCount) {
      const templateP = origPs[origPs.length - 1];
      for (let li = origCount; li < lineCount; li++) {
        const newP = templateP.cloneNode(true);
        if (styledLines && styledLines[li]) {
          patchParagraphStyled(newP, styledLines[li], charPrList, HP, alignToParaPrId);
        } else {
          patchParagraphText(newP, newLines[li], HP);
        }
        adjustVertpos(newP, li);
        source.appendChild(newP);
      }
    }

    // 줄 삭제 — ★ 마지막 빈줄은 보존 (TipTap이 끝 빈줄을 제거하는 문제 방지)
    if (lineCount < origCount) {
      for (let li = origCount - 1; li >= lineCount; li--) {
        // 원본 마지막 <p>가 빈줄이면 보존
        var delP = origPs[li];
        var delRuns = getRuns(delP);
        var delText = '';
        for (var dri = 0; dri < delRuns.length; dri++) {
          var dtEl = findChild(delRuns[dri], 't');
          if (dtEl && dtEl.textContent) delText += dtEl.textContent;
        }
        if (li === origCount - 1 && delText.trim() === '') {
          // 마지막 빈줄 보존 — 텍스트만 비우기
          if (delRuns.length > 0) setFirstRunText(delRuns[0], '', delP, HP);
          continue;
        }
        source.removeChild(delP);
      }
    }

    // ★ 셀 내 lineseg vertpos 재계산 — 글꼴 크기가 실제 변경된 경우만
    var hasStyleEdit = false;
    if (styledLines) {
      for (var sei = 0; sei < styledLines.length; sei++) {
        var seLine = styledLines[sei];
        if (!seLine || !seLine.runs) continue;
        for (var sri = 0; sri < seLine.runs.length; sri++) {
          var seRun = seLine.runs[sri];
          if (seRun.fontSize && seRun.fontSize !== 10) { hasStyleEdit = true; break; }
        }
        if (hasStyleEdit) break;
      }
    }
    if (hasStyleEdit) {
      var cellPs = [];
      for (var cpi = 0; cpi < source.childNodes.length; cpi++) {
        var cpnd = source.childNodes[cpi];
        if (cpnd.nodeType === 1 && (cpnd.localName || cpnd.tagName.split(":").pop()) === "p") {
          cellPs.push(cpnd);
        }
      }
      if (cellPs.length > 0) {
        var cumVertpos = 0;
        for (var cpj = 0; cpj < cellPs.length; cpj++) {
          var cpLsa = findChild(cellPs[cpj], "linesegarray");
          if (!cpLsa) continue;
          for (var cpk = 0; cpk < cpLsa.childNodes.length; cpk++) {
            var cpSeg = cpLsa.childNodes[cpk];
            if (cpSeg.nodeType !== 1) continue;
            var cpSegTag = cpSeg.localName || cpSeg.tagName.split(":").pop();
            if (cpSegTag !== "lineseg") continue;
            cpSeg.setAttribute("vertpos", String(cumVertpos));
            var vs = parseInt(cpSeg.getAttribute("vertsize") || "1000", 10);
            // lineSpacing 160% 기본
            cumVertpos += Math.round(vs * 1.6);
          }
        }
      }
    }

    patchCount++;
  }

  // ── 본문 텍스트 블록(data-block) 패치 ──
  if (blockEdits && blockEdits.length > 0) {
    // 전체 최상위 <p> 수집
    const allTopPs = [];
    for (let i = 0; i < root.childNodes.length; i++) {
      const nd = root.childNodes[i];
      if (nd.nodeType !== 1) continue;
      const tag = nd.localName || nd.tagName.split(":").pop();
      if (tag === "p") allTopPs.push(nd);
    }

    // blockIdx 순서로 정렬
    const sortedEdits = blockEdits.slice().sort(function(a, b) { return a.blockIdx - b.blockIdx; });

    for (const edit of sortedEdits) {
      const { styledLines, newLines, pCount } = edit;

      // pCount가 0이면 표 내부 텍스트 — 별도 XML <p>가 없으므로 건너뜀
      if (!pCount || pCount <= 0) continue;

      // xmlPStart 결정: Parser에서 전달받은 값 사용
      let xmlPStart = (typeof edit.xmlPStart === 'number') ? edit.xmlPStart : -1;

      // xmlPStart가 없으면(-1) 패치 불가 → 건너뜀
      if (xmlPStart < 0 || xmlPStart >= allTopPs.length) continue;

      // ★ 원본 <p>별 텍스트 추출
      var origTexts = [];
      for (var oi = 0; oi < pCount; oi++) {
        var oIdx = xmlPStart + oi;
        if (oIdx >= allTopPs.length) break;
        origTexts.push(getPlainText(allTopPs[oIdx]).trim());
      }

      // ★ styledLines에서 텍스트 추출
      var editTexts = [];
      var editLineCount = styledLines ? styledLines.length : (newLines ? newLines.length : 0);
      for (var ei = 0; ei < editLineCount; ei++) {
        if (styledLines && styledLines[ei]) {
          editTexts.push(styledLines[ei].runs.map(function(r) { return r.text || ''; }).join('').trim());
        } else if (newLines && newLines[ei]) {
          editTexts.push(newLines[ei].trim());
        } else {
          editTexts.push('');
        }
      }

      // ★ 각 원본 <p>에 대해 가장 적합한 styledLine 찾기 (순서 유지)
      // 전략: 원본 <p> 순서대로, editTexts에서 순차 검색하여 매칭
      var editCursor = 0; // 현재까지 소비한 styledLines 인덱스

      for (var li = 0; li < pCount; li++) {
        var pIdx = xmlPStart + li;
        if (pIdx >= allTopPs.length) break;
        var pEl = allTopPs[pIdx];
        if (!pEl) continue;

        var origText = origTexts[li] || '';

        // 빈줄 <p>는 편집할 게 없으면 원본 유지
        if (!origText && editCursor < editTexts.length && !editTexts[editCursor]) {
          // 양쪽 다 빈줄 → 커서만 전진
          editCursor++;
          continue;
        }
        if (!origText && (editCursor >= editTexts.length || editTexts[editCursor])) {
          // 원본은 빈줄인데 edit에는 빈줄이 없음 → 원본 유지, 커서 전진 안 함
          continue;
        }

        // 원본에 텍스트가 있는 경우 → styledLines에서 매칭
        if (editCursor < editLineCount) {
          try {
            if (styledLines && styledLines[editCursor]) {
              patchParagraphStyled(pEl, styledLines[editCursor], charPrList, HP, alignToParaPrId);
            } else if (newLines && newLines[editCursor]) {
              patchParagraphText(pEl, newLines[editCursor], HP);
            }
            patchCount++;
          } catch (e) {
            // 개별 패치 실패해도 계속
          }
          editCursor++;
        }
      }
    }
  }

  if (patchCount === 0) return xmlStr;
  return new XMLSerializer().serializeToString(doc);
}


// ── 내부 함수들 ──

function findCell(tbl, rowAddr, colAddr) {
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const tr = tbl.childNodes[i];
    if (tr.nodeType !== 1) continue;
    const trTag = tr.localName || tr.tagName?.split(":").pop();
    if (trTag !== "tr") continue;

    for (let j = 0; j < tr.childNodes.length; j++) {
      const tc = tr.childNodes[j];
      if (tc.nodeType !== 1) continue;
      const tcTag = tc.localName || tc.tagName?.split(":").pop();
      if (tcTag !== "tc") continue;

      for (let k = 0; k < tc.childNodes.length; k++) {
        const child = tc.childNodes[k];
        if (child.nodeType !== 1) continue;
        const ctag = child.localName || child.tagName?.split(":").pop();
        if (ctag === "cellAddr") {
          if (child.getAttribute("colAddr") === String(colAddr) &&
              child.getAttribute("rowAddr") === String(rowAddr)) return tc;
        }
      }
    }
  }
  return null;
}

/** 텍스트만 교체 (스타일 미변경 — 기존 방식) */
function patchParagraphText(pEl, newText, ns) {
  const runs = getRuns(pEl);
  if (runs.length === 0) return;

  setFirstRunText(runs[0], newText, pEl, ns);
  clearRemainingRuns(runs);
}

/** 스타일 포함 패치 — run별 charPrIDRef 매핑 */
function patchParagraphStyled(pEl, styledLine, charPrList, ns, alignToParaPrId = {}) {
  const origRuns = getRuns(pEl);
  if (origRuns.length === 0) return;

  const { runs: htmlRuns, align } = styledLine;

  // ★ 정렬 변경: 명시적으로 변경된 경우에만 paraPrIDRef 교체
  // 'left'는 TipTap 기본값이므로 원본 유지 (JUSTIFY 등을 덮어쓰지 않음)
  if (align && align !== 'left' && alignToParaPrId[align]) {
    pEl.setAttribute('paraPrIDRef', alignToParaPrId[align]);
  }
  // align === 'left'인 경우는 원본 paraPrIDRef 유지 (변경하지 않음)

  const defaultRef = origRuns[0].getAttribute("charPrIDRef") || "0";

  // ★ 빈 줄 감지: 텍스트가 없거나 공백만 있는 경우
  const allText = htmlRuns.map(function(r) { return r.text || ''; }).join('').trim();
  if (htmlRuns.length === 0 || !allText) {
    // 빈 줄 — 텍스트만 비우기 (원본 run/paraPrIDRef 유지)
    setFirstRunText(origRuns[0], "", pEl, ns);
    clearRemainingRuns(origRuns);
    return;
  }

  // 스타일이 모두 기본(검정, 10pt, 비굵게)이면 → 기존 방식으로 텍스트만 교체
  const allDefault = htmlRuns.every(r =>
    !r.bold && !r.italic && !r.underline &&
    (!r.color || r.color === "#000000") &&
    (!r.fontSize || r.fontSize === 10)
  );
  if (allDefault) {
    const fullText = htmlRuns.map(r => r.text).join("");
    setFirstRunText(origRuns[0], fullText, pEl, ns);
    clearRemainingRuns(origRuns);
    return;
  }

  // 스타일 있는 경우 — run별 처리
  // 전략: 첫 번째 원본 run을 템플릿으로 사용, htmlRuns 만큼 run 생성
  const templateRun = origRuns[0];

  // 기존 run 모두 제거 (linesegarray 등은 보존)
  for (const r of origRuns) {
    pEl.removeChild(r);
  }

  // ★ 빈 텍스트 run 필터링 (텍스트가 있는 run만 생성)
  var filteredRuns = [];
  for (var fi = 0; fi < htmlRuns.length; fi++) {
    if (htmlRuns[fi].text) filteredRuns.push(htmlRuns[fi]);
  }
  if (filteredRuns.length === 0) filteredRuns = htmlRuns; // 전부 빈 경우 원본 유지

  // htmlRuns 만큼 새 run 생성
  const linesegEl = findChild(pEl, "linesegarray");
  var maxFontHeight = 0; // 이 줄에서 가장 큰 글꼴 높이 추적
  for (var hri = 0; hri < filteredRuns.length; hri++) {
    var hr = filteredRuns[hri];
    const newRun = templateRun.cloneNode(true);

    // charPrIDRef 매핑
    if (charPrList && charPrList.length > 0) {
      const newRef = findCharPrId(charPrList, hr, defaultRef);
      newRun.setAttribute("charPrIDRef", newRef);

      // ★ 해당 charPr의 height를 추적하여 lineseg 업데이트에 사용
      var matchedCp = null;
      for (var ci = 0; ci < charPrList.length; ci++) {
        if (charPrList[ci].id === newRef) { matchedCp = charPrList[ci]; break; }
      }
      if (matchedCp) {
        var fontH = (matchedCp.fontSize || 10) * 100;
        if (fontH > maxFontHeight) maxFontHeight = fontH;
      }
    }

    // 텍스트 설정
    let tEl = findChild(newRun, "t");
    if (tEl) {
      while (tEl.firstChild) tEl.removeChild(tEl.firstChild);
      if (hr.text) tEl.appendChild(pEl.ownerDocument.createTextNode(hr.text));
    } else if (hr.text) {
      tEl = pEl.ownerDocument.createElementNS(ns, "hp:t");
      tEl.appendChild(pEl.ownerDocument.createTextNode(hr.text));
      newRun.insertBefore(tEl, newRun.firstChild);
    }

    // linesegarray 앞에 삽입
    if (linesegEl) {
      pEl.insertBefore(newRun, linesegEl);
    } else {
      pEl.appendChild(newRun);
    }
  }

  // ★ lineseg의 vertsize/textheight/baseline을 최대 글꼴 크기에 맞게 업데이트
  if (maxFontHeight > 0 && linesegEl) {
    for (var si = 0; si < linesegEl.childNodes.length; si++) {
      var seg = linesegEl.childNodes[si];
      if (seg.nodeType !== 1) continue;
      var segTag = seg.localName || seg.tagName.split(":").pop();
      if (segTag === "lineseg") {
        seg.setAttribute("vertsize", String(maxFontHeight));
        seg.setAttribute("textheight", String(maxFontHeight));
        seg.setAttribute("baseline", String(Math.round(maxFontHeight * 0.85)));
      }
    }
  }
}

function getRuns(pEl) {
  const runs = [];
  for (let i = 0; i < pEl.childNodes.length; i++) {
    const nd = pEl.childNodes[i];
    if (nd.nodeType === 1) {
      const tag = nd.localName || nd.tagName.split(":").pop();
      if (tag === "run") runs.push(nd);
    }
  }
  return runs;
}

function findChild(parent, name) {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const nd = parent.childNodes[i];
    if (nd.nodeType === 1) {
      const tag = nd.localName || nd.tagName.split(":").pop();
      if (tag === name) return nd;
    }
  }
  return null;
}

function setFirstRunText(run, text, pEl, ns) {
  let tEl = findChild(run, "t");
  if (tEl) {
    while (tEl.firstChild) tEl.removeChild(tEl.firstChild);
    if (text) tEl.appendChild(pEl.ownerDocument.createTextNode(text));
  } else if (text) {
    tEl = pEl.ownerDocument.createElementNS(ns, "hp:t");
    tEl.appendChild(pEl.ownerDocument.createTextNode(text));
    run.insertBefore(tEl, run.firstChild);
  }
}

function clearRemainingRuns(runs) {
  for (let ri = 1; ri < runs.length; ri++) {
    const tEl = findChild(runs[ri], "t");
    if (tEl) {
      while (tEl.firstChild) tEl.removeChild(tEl.firstChild);
    }
  }
}

function adjustVertpos(pEl, lineIndex) {
  const lsa = findChild(pEl, "linesegarray");
  if (!lsa) return;
  for (let j = 0; j < lsa.childNodes.length; j++) {
    const seg = lsa.childNodes[j];
    if (seg.nodeType === 1) {
      seg.setAttribute("vertpos", String(lineIndex * 1400));
    }
  }
}

/** XML <p> 요소에서 텍스트만 추출 (header 감지용) */
function getPlainText(pEl) {
  if (!pEl) return "";
  const texts = [];
  const all = pEl.getElementsByTagName("*");
  for (let i = 0; i < all.length; i++) {
    const tag = all[i].localName || all[i].tagName.split(":").pop();
    if (tag === "t" && all[i].textContent) texts.push(all[i].textContent);
  }
  return texts.join(" ");
}