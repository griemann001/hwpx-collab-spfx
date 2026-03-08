/**
 * ============================================================
 * HwpxParser.js — HWPX → HTML 변환 파서 (v13 기반)
 * ============================================================
 * 
 * teams-alimi의 HwpxParser v13을 모듈화.
 * 
 * 지원:
 *   - 중첩 표 (subList > p > run > tbl)
 *   - 인라인 소형 표 합침 (동의□ + 미동의□)
 *   - lineBreak, 이미지, charPr 스타일
 *   - 머리말/꼬리말 자동 감지 제외
 *   - 셀 원본 텍스트 기록 (변경 감지용)
 */

// ── 헬퍼 ──

function ln(el) { return el.localName || el.tagName.split(":").pop() || el.tagName; }

function kids(parent, name) {
  const r = [];
  for (let i = 0; i < parent.children.length; i++) {
    if (ln(parent.children[i]) === name) r.push(parent.children[i]);
  }
  return r;
}

function desc(parent, name) {
  const r = [];
  const all = parent.getElementsByTagName("*");
  for (let i = 0; i < all.length; i++) {
    if (ln(all[i]) === name) r.push(all[i]);
  }
  return r;
}

function hwpToPx(hwpUnit) { return Math.round(hwpUnit * 96 / 7200); }
function textLen(html) { return html.replace(/<[^>]*>/g, "").trim().length; }

// ── JSZip (npm 패키지) ──

const JSZip = require('jszip');

export async function getJSZip() {
  return JSZip;
}

// ── HwpxParser 클래스 ──

export class HwpxParser {
  imgMap = new Map();
  charStyles = new Map();
  paraStyles = new Map();
  borderFills = new Map();
  isLandscape = false;
  tableIndex = 0;
  bodyWidthHwp = 51026;
  cellOriginalTexts = new Map(); // "tbl-row-col" → "줄1\n줄2\n줄3"
  blockOriginalTexts = new Map(); // "block-N" → "줄1\n줄2\n줄3"
  blockParagraphCounts = new Map(); // "block-N" → 해당 block에 포함된 XML <p> 개수
  blockXmlPStartIndex = new Map(); // "block-N" → 전체 최상위 <p> 중 이 블록의 시작 인덱스
  headerParagraphCount = 0;
  footerParagraphCount = 0;
  blockIndex = 0;

  async parse(fileBuffer) {
    const result = { html: "", headerHtml: "", footerHtml: "", title: "", images: new Map(), errors: [] };
    this.tableIndex = 0;
    this.blockIndex = 0;
    this.cellOriginalTexts = new Map();
    this.blockOriginalTexts = new Map();
    this.blockParagraphCounts = new Map();
    this.blockXmlPStartIndex = new Map();
    this.headerParagraphCount = 0;
    this.footerParagraphCount = 0;
    try {
      const JSZip = await getJSZip();
      const zip = await JSZip.loadAsync(fileBuffer);
      await this.loadStyles(zip);
      await this.loadImages(zip, result);
      this.imgMap = result.images;

      const secFiles = [];
      zip.forEach((p) => { if (p.match(/Contents\/section\d+\.xml$/i)) secFiles.push(p); });
      secFiles.sort();
      if (secFiles.length === 0) { result.errors.push("section 파일 없음"); return result; }

      const parts = [];
      for (const sf of secFiles) {
        const zf = zip.file(sf);
        if (zf) parts.push(this.parseSection(await zf.async("string"), result));
      }
      result.html = parts.join('<hr style="page-break-after:always;"/>');
      result.isLandscape = this.isLandscape;
      result.pageSettings = this.pageSettings;
      result.cellOriginalTexts = this.cellOriginalTexts;
      result.blockOriginalTexts = this.blockOriginalTexts;
      result.blockParagraphCounts = this.blockParagraphCounts;
      result.blockXmlPStartIndex = this.blockXmlPStartIndex;
      result.headerParagraphCount = this.headerParagraphCount;
      result.footerParagraphCount = this.footerParagraphCount;
      result.charPrList = Array.from(this.charStyles.entries()).map(([id, s]) => ({
        id, fontSize: s.fontSize, color: s.color, bold: s.bold,
      }));
      // paraPr 목록 (정렬 매핑용)
      result.paraPrList = Array.from(this.paraStyles.entries()).map(([id, s]) => ({
        id, align: (s.align || 'left').toUpperCase(),
      }));
    } catch (e) { result.errors.push(`파싱 오류: ${e.message}`); }
    return result;
  }

  async loadStyles(zip) {
    const hf = zip.file("Contents/header.xml");
    if (!hf) return;
    const xml = await hf.async("string");
    const doc = new DOMParser().parseFromString(xml, "text/xml");

    for (const cp of desc(doc.documentElement, "charPr")) {
      const id = cp.getAttribute("id");
      if (!id) continue;
      const height = cp.getAttribute("height") || "0";
      const textColor = cp.getAttribute("textColor") || "#000000";
      let bold = false, italic = false, underline = false, strikeout = false;
      for (let i = 0; i < cp.children.length; i++) {
        const tag = ln(cp.children[i]);
        if (tag === "bold") bold = true;
        else if (tag === "italic") italic = true;
        else if (tag === "underline") underline = (cp.children[i].getAttribute("type") || "NONE") !== "NONE";
        else if (tag === "strikeout") {
          const shape = (cp.children[i].getAttribute("shape") || "NONE").toUpperCase();
          strikeout = shape !== "NONE" && shape !== "3D";
        }
      }
      this.charStyles.set(id, { bold, italic, underline, strikeout, fontSize: parseInt(height, 10) / 100, color: textColor });
    }
    for (const pp of desc(doc.documentElement, "paraPr")) {
      const id = pp.getAttribute("id");
      if (!id) continue;
      let align = "left";
      const als = kids(pp, "align");
      if (als.length > 0) {
        const h = (als[0].getAttribute("horizontal") || "").toUpperCase();
        const m = { LEFT: "left", CENTER: "center", RIGHT: "right", JUSTIFY: "justify" };
        align = m[h] || "left";
      }
      this.paraStyles.set(id, { align });
    }
    for (const bf of desc(doc.documentElement, "borderFill")) {
      const id = bf.getAttribute("id");
      if (!id) continue;
      let bg = "";
      const wb = desc(bf, "winBrush");
      if (wb.length > 0) {
        const fc = wb[0].getAttribute("faceColor");
        if (fc && fc !== "#000000" && fc.toLowerCase() !== "#ffffff") bg = fc;
      }
      const borders = { left: "SOLID", right: "SOLID", top: "SOLID", bottom: "SOLID" };
      for (const side of ["left", "right", "top", "bottom"]) {
        const bEls = kids(bf, `${side}Border`);
        if (bEls.length > 0) {
          const btype = (bEls[0].getAttribute("type") || "SOLID").toUpperCase();
          borders[side] = btype;
        }
      }
      let diagonal = "none";
      const slashEls = kids(bf, "slash");
      if (slashEls.length > 0 && (slashEls[0].getAttribute("type") || "NONE") !== "NONE") diagonal = "slash";
      const bsEls = kids(bf, "backSlash");
      if (bsEls.length > 0 && (bsEls[0].getAttribute("type") || "NONE") !== "NONE") diagonal = "backSlash";
      this.borderFills.set(id, { bg, borders, diagonal });
    }
  }

  async loadImages(zip, result) {
    const exts = [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"];
    const entries = Object.entries(zip.files);
    for (let i = 0; i < entries.length; i++) {
      const [path, file] = entries[i];
      if (file.dir) continue;
      const lp = path.toLowerCase();
      if (!lp.includes("bindata/")) continue;
      if (!exts.some(e => lp.endsWith(e))) continue;
      try {
        const b64 = await file.async("base64");
        const ext = path.split(".").pop()?.toLowerCase() || "png";
        const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "bmp" ? "image/bmp" : `image/${ext}`;
        const dataUrl = `data:${mime};base64,${b64}`;
        const fname = path.split("/").pop() || path;
        const noext = fname.replace(/\.[^.]+$/, "");
        result.images.set(fname, dataUrl);
        result.images.set(noext, dataUrl);
      } catch { result.errors.push(`이미지 실패: ${path}`); }
    }
  }

  parseSection(xml, result) {
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const root = doc.documentElement;
    this.calcBodyWidth(root);

    const top = [];
    for (let i = 0; i < root.children.length; i++) top.push(root.children[i]);
    if (top.length === 0) return "";

    const hc = this.detectHeader(top);
    const fc = this.detectFooter(top);
    this.headerParagraphCount = hc;
    this.footerParagraphCount = fc;
    if (hc >= 2) result.title = this.extractTitle(top[1]);

    const hp = [];
    for (let i = 0; i < hc; i++) hp.push(this.renderP(top[i], this.bodyWidthHwp));
    result.headerHtml = hp.join("\n");

    const fp = [];
    for (let i = top.length - fc; i < top.length; i++) fp.push(this.renderP(top[i], this.bodyWidthHwp));
    result.footerHtml = fp.join("\n");

    const bp = [];
    let prevVertpos = -1;
    let textBuffer = []; // 표 사이 텍스트를 모으는 버퍼
    let textBufferPCount = 0; // 버퍼에 포함된 원본 XML <p> 개수
    let textBufferStartXmlP = -1; // 버퍼의 첫 번째 XML <p>의 전체 인덱스
    let xmlPIndex = 0; // 전체 최상위 요소 중 현재 인덱스 (hc부터 시작)

    const flushTextBlock = () => {
      if (textBuffer.length === 0) return;
      const blockIdx = this.blockIndex++;
      const blockKey = 'block-' + blockIdx;
      const html = textBuffer.join("\n");
      // 원본 텍스트 저장 (편집 감지용)
      const tempDiv = typeof document !== 'undefined' ? document.createElement('div') : null;
      if (tempDiv) {
        tempDiv.innerHTML = html;
        this.blockOriginalTexts.set(blockKey, (tempDiv.textContent || '').trim());
      }
      // ★ 이 block에 포함된 원본 XML <p> 개수 및 시작 인덱스 기록
      this.blockParagraphCounts.set(blockKey, textBufferPCount);
      this.blockXmlPStartIndex.set(blockKey, textBufferStartXmlP);
      bp.push('<div data-block="' + blockIdx + '" style="min-height:1.2em;">' + html + '</div>');
      textBuffer = [];
      textBufferPCount = 0;
      textBufferStartXmlP = -1;
    };

    for (let i = hc; i < top.length - fc; i++) {
      const vertpos = this.getFirstVertpos(top[i]);
      if (vertpos === 0 && prevVertpos > 0 && i > hc) {
        flushTextBlock();
        bp.push('<div style="page-break-before:always;border-top:1px dashed #CBD5E1;margin:32px 0;"></div>');
      }
      if (vertpos >= 0) prevVertpos = vertpos;

      const rendered = this.renderP(top[i], this.bodyWidthHwp);

      // 표+텍스트 혼합 여부 판별
      if (rendered.indexOf('<table') === -1) {
        // 순수 텍스트 → 버퍼에 추가
        if (textBufferStartXmlP === -1) textBufferStartXmlP = i; // 최초 <p> 인덱스 기록
        textBuffer.push(rendered);
        textBufferPCount++;
      } else {
        // 표가 포함된 경우 → 표와 텍스트를 분리
        flushTextBlock();
        // rendered를 <table...>...</table> 기준으로 분리
        const parts = rendered.split(/(<table[\s\S]*?<\/table>)/g);
        for (let pi = 0; pi < parts.length; pi++) {
          const part = parts[pi].trim();
          if (!part) continue;
          if (part.indexOf('<table') === 0) {
            // 표 부분은 그대로
            flushTextBlock();
            bp.push(part);
          } else {
            // 표 뒤/앞 텍스트 → 블록으로
            // ★ 이 텍스트는 표가 포함된 <p> 안의 run 일부이므로
            //    별도 XML <p>가 아님 → pCount를 증가시키지 않고
            //    xmlPStart도 -1로 설정 (Patcher에서 패치 불가 표시)
            textBuffer.push(part);
            // textBufferPCount는 증가시키지 않음 (XML <p>가 아니므로)
            // textBufferStartXmlP도 설정하지 않음
          }
        }
        flushTextBlock();
      }
    }
    flushTextBlock(); // 마지막 텍스트 블록 flush
    return bp.join("\n");
  }

  getFirstVertpos(pEl) {
    const lsaList = kids(pEl, "linesegarray");
    if (lsaList.length === 0) return -1;
    const segs = kids(lsaList[0], "lineseg");
    if (segs.length === 0) return -1;
    const vp = segs[0].getAttribute("vertpos");
    return vp !== null ? parseInt(vp, 10) : -1;
  }

  calcBodyWidth(root) {
    const pagePrs = desc(root, "pagePr");
    if (pagePrs.length === 0) return;
    const pp = pagePrs[0];
    const landscape = (pp.getAttribute("landscape") || "WIDELY").toUpperCase();
    let pageWidth = parseInt(pp.getAttribute("width") || "0", 10);
    const pageHeight = parseInt(pp.getAttribute("height") || "0", 10);

    if (landscape === "NARROWLY" && pageHeight > pageWidth) {
      pageWidth = pageHeight;
      this.isLandscape = true;
    }

    this.pageSettings = {
      width: this.isLandscape ? pageHeight : pageWidth,
      height: this.isLandscape ? pageWidth : pageHeight,
      orientation: this.isLandscape ? "landscape" : "portrait",
      margin: { top: 5668, bottom: 4252, left: 8504, right: 8504, header: 4252, footer: 4252, gutter: 0 },
    };

    const margins = kids(pp, "margin");
    if (margins.length > 0 && pageWidth > 0) {
      let ml, mr;
      const mt = parseInt(margins[0].getAttribute("top") || "0", 10);
      const mb = parseInt(margins[0].getAttribute("bottom") || "0", 10);
      const mll = parseInt(margins[0].getAttribute("left") || "0", 10);
      const mrr = parseInt(margins[0].getAttribute("right") || "0", 10);
      const mh = parseInt(margins[0].getAttribute("header") || "0", 10);
      const mf = parseInt(margins[0].getAttribute("footer") || "0", 10);
      const mg = parseInt(margins[0].getAttribute("gutter") || "0", 10);

      if (this.isLandscape) { ml = mt; mr = mb; }
      else { ml = mll; mr = mrr; }
      this.bodyWidthHwp = pageWidth - ml - mr;

      this.pageSettings.margin = {
        top: this.isLandscape ? mll : mt,
        bottom: this.isLandscape ? mrr : mb,
        left: this.isLandscape ? mt : mll,
        right: this.isLandscape ? mb : mrr,
        header: mh, footer: mf, gutter: mg,
      };
    }
  }

  extractTitle(headerP) {
    const tbls = desc(headerP, "tbl");
    if (tbls.length === 0) return "";
    const trs = kids(tbls[0], "tr");
    if (trs.length < 2) return "";
    const tcs = kids(trs[1], "tc");
    return tcs.length > 0 ? this.plainText(tcs[0]).trim() : "";
  }

  detectHeader(els) {
    if (els.length < 2) return 0;
    const t0 = this.plainText(els[0]);
    if (!t0.match(/Tel|Fax|교무실|행정실|\d{2,3}-\d{3,4}-\d{4}/)) return 0;
    const t1 = els.length > 1 ? this.plainText(els[1]) : "";
    return t1.match(/교육통신|통신문|알림장|날\s*짜|대\s*상/) ? 2 : 1;
  }

  detectFooter(els) {
    let c = 0;
    for (let i = els.length - 1; i >= 0; i--) {
      if (desc(els[i], "tbl").length > 0) break;
      const t = this.plainText(els[i]).trim();
      if (t.length > 100) break;
      if (t === "") c++;
      else if (t.match(/초등학교장|중학교장|고등학교장|학교장/)) c++;
      else if (t.match(/^\d{4}\.\s*\d{1,2}\.\s*\d{1,2}/) || t.match(/^\d{4}년\s*\d{1,2}월/)) { c++; break; }
      else break;
    }
    if (c >= els.length) c = Math.max(0, els.length - 1);
    return c;
  }

  plainText(el) {
    const ts = [];
    const all = el.getElementsByTagName("*");
    for (let i = 0; i < all.length; i++) {
      if (ln(all[i]) === "t" && all[i].textContent) ts.push(all[i].textContent || "");
    }
    return ts.join(" ");
  }

  renderP(pEl, containerWidth) {
    if (ln(pEl) !== "p") return "";
    const runs = kids(pEl, "run");
    if (runs.length === 0) return '<p>&nbsp;</p>';
    let hasTable = false;
    for (const r of runs) { if (kids(r, "tbl").length > 0) { hasTable = true; break; } }
    if (hasTable) return this.renderPWithTables(pEl, runs, containerWidth);
    const align = this.getParaAlign(pEl);
    const content = runs.map(r => this.renderRun(r)).join("");
    if (!content) return '<p>&nbsp;</p>';
    const st = align && align !== "left" ? ` style="text-align:${align}"` : "";
    return `<p${st}>${content}</p>`;
  }

  renderPWithTables(pEl, runs, containerWidth) {
    const align = this.getParaAlign(pEl);
    const pieces = [];
    for (const r of runs) {
      const tbls = kids(r, "tbl");
      if (tbls.length > 0) {
        for (let ci = 0; ci < r.childNodes.length; ci++) {
          const nd = r.childNodes[ci];
          if (nd.nodeType !== 1) continue;
          const el = nd;
          const tag = ln(el);
          if (tag === "tbl") {
            const posEls = kids(el, "pos");
            const treatAsChar = posEls.length > 0 ? posEls[0].getAttribute("treatAsChar") : "1";
            if (treatAsChar === "0") {
              pieces.push({ type: "floatingTable", html: this.renderTable(el, containerWidth) });
            } else {
              const isSmall = this.isSmallTable(el, containerWidth);
              if (isSmall) pieces.push({ type: "inlineTable", html: "", tblEl: el });
              else pieces.push({ type: "blockTable", html: this.renderTable(el, containerWidth) });
            }
          } else if (tag === "t") {
            const html = this.renderT(el);
            if (html) {
              const style = this.getCharStyle(r);
              const wrapped = style ? `<span style="${style}">${html}</span>` : html;
              pieces.push({ type: textLen(html) <= 5 ? "shortText" : "longText", html: wrapped });
            }
          } else if (tag === "pic") {
            const html = this.renderPic(el);
            if (html) pieces.push({ type: "shortText", html });
          }
        }
      } else {
        const txt = this.renderRun(r);
        if (txt) pieces.push({ type: textLen(txt) <= 5 ? "shortText" : "longText", html: txt });
      }
    }

    const output = [];
    const floatingTables = [];
    let inlineTbls = [];
    let textBuf = [];
    const flushText = () => {
      if (textBuf.length === 0) return;
      const st = align && align !== "left" ? ` style="text-align:${align}"` : "";
      output.push(`<p${st}>${textBuf.join("")}</p>`);
      textBuf = [];
    };
    const flushInlineTables = () => {
      if (inlineTbls.length === 0) return;
      if (inlineTbls.length === 1) output.push(this.renderTable(inlineTbls[0], containerWidth));
      else output.push(this.mergeInlineTables(inlineTbls));
      inlineTbls = [];
    };

    for (const piece of pieces) {
      switch (piece.type) {
        case "blockTable": flushText(); flushInlineTables(); output.push(piece.html); break;
        case "floatingTable": floatingTables.push(piece.html); break;
        case "inlineTable": flushText(); if (piece.tblEl) inlineTbls.push(piece.tblEl); break;
        case "shortText": if (inlineTbls.length === 0) textBuf.push(piece.html); break;
        case "longText": flushInlineTables(); textBuf.push(piece.html); break;
      }
    }
    flushText(); flushInlineTables();
    output.push(...floatingTables);
    return output.join("\n");
  }

  mergeInlineTables(tbls) {
    const allCells = [];
    let totalWidthHwp = 0;
    for (let ti = 0; ti < tbls.length; ti++) {
      const tbl = tbls[ti];
      let tblWidthHwp = 0;
      const szEls = kids(tbl, "sz");
      if (szEls.length > 0) { const w = szEls[0].getAttribute("width"); if (w) tblWidthHwp = parseInt(w, 10); }
      totalWidthHwp += tblWidthHwp;
      const trs = kids(tbl, "tr");
      for (const tr of trs) {
        const tcs = kids(tr, "tc");
        for (const tc of tcs) {
          let cellWidthHwp = 0;
          const cellSzEls = kids(tc, "cellSz");
          if (cellSzEls.length > 0) { const cw = cellSzEls[0].getAttribute("width"); if (cw) cellWidthHwp = parseInt(cw, 10); }
          const px = hwpToPx(cellWidthHwp);
          const content = this.renderCellContent(tc, cellWidthHwp);
          allCells.push(`<td style="border:1px solid #999;padding:4px 6px;vertical-align:middle;width:${px}px;">${content || "&nbsp;"}</td>`);
        }
      }
      if (ti < tbls.length - 1) allCells.push(`<td style="border:none;padding:0;width:6px;">&nbsp;</td>`);
    }
    const totalPx = hwpToPx(totalWidthHwp) + (tbls.length - 1) * 6;
    return `<table style="border-collapse:collapse;width:${totalPx}px;margin:8px 0;"><tr>${allCells.join("")}</tr></table>`;
  }

  isSmallTable(tbl, containerWidth) {
    const szEls = kids(tbl, "sz");
    if (szEls.length === 0) return false;
    const w = szEls[0].getAttribute("width");
    if (!w) return false;
    return containerWidth > 0 && (parseInt(w, 10) / containerWidth) < 0.4;
  }

  renderRun(run, skipTbl) {
    const parts = [];
    for (let i = 0; i < run.childNodes.length; i++) {
      const nd = run.childNodes[i];
      if (nd.nodeType !== 1) continue;
      const el = nd;
      const tag = ln(el);
      if (tag === "t") parts.push(this.renderT(el));
      else if (tag === "pic") parts.push(this.renderPic(el));
      else if (tag === "tbl" && skipTbl) { /* skip */ }
    }
    const text = parts.join("");
    if (!text) return "";
    const style = this.getCharStyle(run);
    return style ? `<span style="${style}">${text}</span>` : text;
  }

  getCharStyle(run) {
    const ref = run.getAttribute("charPrIDRef");
    if (!ref) return "";
    const cs = this.charStyles.get(ref);
    if (!cs) return "";
    const s = [];
    if (cs.bold) s.push("font-weight:bold");
    if (cs.italic) s.push("font-style:italic");
    if (cs.underline) s.push("text-decoration:underline");
    if (cs.strikeout) s.push("text-decoration:line-through");
    if (cs.fontSize > 0 && cs.fontSize !== 10) s.push(`font-size:${cs.fontSize}pt`);
    if (cs.color && cs.color !== "#000000") s.push(`color:${cs.color}`);
    return s.join(";");
  }

  getParaAlign(pEl) {
    const ref = pEl.getAttribute("paraPrIDRef");
    if (!ref) return "";
    const ps = this.paraStyles.get(ref);
    return ps ? ps.align : "";
  }

  renderT(tEl) {
    const parts = [];
    for (let i = 0; i < tEl.childNodes.length; i++) {
      const nd = tEl.childNodes[i];
      if (nd.nodeType === 3) {
        const t = nd.textContent || "";
        if (t) parts.push(this.esc(this.replacePUA(t)));
      } else if (nd.nodeType === 1) {
        const tag = ln(nd);
        if (tag === "lineBreak") parts.push("<br/>");
        else if (tag === "fwSpace") parts.push(" ");
        else if (tag === "tab") parts.push("&emsp;");
      }
    }
    return parts.join("");
  }

  replacePUA(text) {
    let result = "";
    for (const ch of text) {
      const code = ch.codePointAt(0) || 0;
      if (code === 0xF02EF) { result += "\u2022"; }
      else if (code === 0xF007E) { result += "\u25B8"; }
      else if (code >= 0xF0000 && code <= 0xFFFFF) { result += "\u2022"; }
      else if (code === 0xF06F) { result += "\u2022"; }
      else if (code === 0xF06E) { result += "\u25AA"; }
      else if (code === 0xF020) { result += " "; }
      else { result += ch; }
    }
    return result;
  }

  renderPic(pic) {
    let width = 0, height = 0;
    const curSzEls = kids(pic, "curSz");
    if (curSzEls.length > 0) {
      const w = curSzEls[0].getAttribute("width");
      const h = curSzEls[0].getAttribute("height");
      if (w) width = hwpToPx(parseInt(w, 10));
      if (h) height = hwpToPx(parseInt(h, 10));
    }
    const imgs = desc(pic, "img");
    for (const img of imgs) {
      const ref = img.getAttribute("binaryItemIDRef");
      if (ref) {
        const src = this.imgMap.get(ref) || "";
        if (src) {
          const sizeStyle = width > 0 && height > 0 ? `width:${width}px;height:${height}px;` : "max-width:100%;height:auto;";
          return `<img src="${src}" style="${sizeStyle}"/>`;
        }
      }
    }
    return "";
  }

  renderTable(tbl, containerWidth) {
    const tblIdx = this.tableIndex++;
    const trs = kids(tbl, "tr");
    if (trs.length === 0) return "";
    let tblWidthHwp = 0;
    const szEls = kids(tbl, "sz");
    if (szEls.length > 0) { const w = szEls[0].getAttribute("width"); if (w) tblWidthHwp = parseInt(w, 10); }
    let tblWidthStyle = "width:100%;";
    if (tblWidthHwp > 0 && containerWidth > 0) {
      const ratio = tblWidthHwp / containerWidth;
      if (ratio >= 0.85) tblWidthStyle = "width:100%;";
      else if (ratio >= 0.4) tblWidthStyle = `width:${Math.round(ratio * 100)}%;`;
      else tblWidthStyle = `width:${hwpToPx(tblWidthHwp)}px;`;
    }
    const rows = [];
    for (let ri = 0; ri < trs.length; ri++) {
      const tr = trs[ri];
      const tcs = kids(tr, "tc");
      const cells = [];
      for (const tc of tcs) {
        const csEl = kids(tc, "cellSpan");
        let cs = "1", rs = "1";
        if (csEl.length > 0) { cs = csEl[0].getAttribute("colSpan") || "1"; rs = csEl[0].getAttribute("rowSpan") || "1"; }
        let colAddr = "0", rowAddr = "0";
        const caEls = kids(tc, "cellAddr");
        if (caEls.length > 0) {
          colAddr = caEls[0].getAttribute("colAddr") || "0";
          rowAddr = caEls[0].getAttribute("rowAddr") || "0";
        }
        let cellWidthStyle = "";
        if (tblWidthHwp > 0) {
          const cellSzEls = kids(tc, "cellSz");
          if (cellSzEls.length > 0) { const cw = cellSzEls[0].getAttribute("width"); if (cw) cellWidthStyle = `width:${Math.round(parseInt(cw, 10) / tblWidthHwp * 100)}%;`; }
        }
        let cellWidthHwp = tblWidthHwp;
        const cellSzEls2 = kids(tc, "cellSz");
        if (cellSzEls2.length > 0) { const cw2 = cellSzEls2[0].getAttribute("width"); if (cw2) cellWidthHwp = parseInt(cw2, 10); }
        let bgStyle = "";
        let borderStyle = "border:1px solid #999;";
        const bfRef = tc.getAttribute("borderFillIDRef");
        if (bfRef && this.borderFills.has(bfRef)) {
          const bf = this.borderFills.get(bfRef);
          if (bf.bg) bgStyle = `background-color:${bf.bg};`;
          const bs = [];
          bs.push(`border-left:${bf.borders.left === "NONE" ? "none" : "1px solid #999"};`);
          bs.push(`border-right:${bf.borders.right === "NONE" ? "none" : "1px solid #999"};`);
          bs.push(`border-top:${bf.borders.top === "NONE" ? "none" : "1px solid #999"};`);
          bs.push(`border-bottom:${bf.borders.bottom === "NONE" ? "none" : "1px solid #999"};`);
          borderStyle = bs.join("");
        }
        let vAlignStyle = "vertical-align:middle;";
        const subs = kids(tc, "subList");
        if (subs.length > 0) {
          const va = (subs[0].getAttribute("vertAlign") || "").toUpperCase();
          if (va === "TOP") vAlignStyle = "vertical-align:top;";
          else if (va === "BOTTOM") vAlignStyle = "vertical-align:bottom;";
        }
        let content = this.renderCellContent(tc, cellWidthHwp);
        if (!content || content === "&nbsp;") {
          const polys = desc(tc, "polygon");
          if (polys.length > 0) { content = '<span style="font-size:24px;color:#EC5F00;">➜</span>'; }
        }
        let diagonalHtml = "";
        if (bfRef && this.borderFills.has(bfRef)) {
          const bf2 = this.borderFills.get(bfRef);
          if (bf2.diagonal === "backSlash") {
            diagonalHtml = '<svg style="position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;"><line x1="0" y1="0" x2="100%" y2="100%" stroke="#000" stroke-width="0.5"/></svg>';
          } else if (bf2.diagonal === "slash") {
            diagonalHtml = '<svg style="position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;"><line x1="100%" y1="0" x2="0" y2="100%" stroke="#000" stroke-width="0.5"/></svg>';
          }
        }
        let attr = ` data-tbl="${tblIdx}" data-row="${rowAddr}" data-col="${colAddr}"`;
        if (parseInt(cs) > 1) attr += ` colspan="${cs}"`;
        if (parseInt(rs) > 1) attr += ` rowspan="${rs}"`;
        const posStyle = diagonalHtml ? "position:relative;" : "";
        cells.push(`<td${attr} style="${borderStyle}padding:8px;${vAlignStyle}${cellWidthStyle}${bgStyle}${posStyle}">${diagonalHtml}${content || "&nbsp;"}</td>`);

        // 원본 텍스트 기록 (변경 감지용)
        const cellKey = `${tblIdx}-${rowAddr}-${colAddr}`;
        const origText = this.extractCellPlainLines(tc);
        this.cellOriginalTexts.set(cellKey, origText);
      }
      rows.push(`<tr>${cells.join("")}</tr>`);
    }
    return `<table style="border-collapse:collapse;${tblWidthStyle}margin:8px 0;">${rows.join("")}</table>`;
  }

  extractCellPlainLines(tc) {
    const sub = kids(tc, "subList");
    const source = sub.length > 0 ? sub[0] : tc;
    const pEls = kids(source, "p");
    const lines = [];
    for (const p of pEls) {
      const parts = [];
      for (const t of desc(p, "t")) {
        for (let i = 0; i < t.childNodes.length; i++) {
          if (t.childNodes[i].nodeType === 3) {
            parts.push(t.childNodes[i].textContent || "");
          }
        }
      }
      lines.push(parts.join(""));
    }
    return lines.join("\n");
  }

  renderCellContent(tc, cellWidthHwp) {
    const parts = [];
    const subs = kids(tc, "subList");
    const source = subs.length > 0 ? subs[0] : tc;
    const pEls = kids(source, "p");
    for (const p of pEls) {
      const rendered = this.renderP(p, cellWidthHwp);
      if (rendered && rendered !== '<p>&nbsp;</p>') parts.push(rendered);
    }
    return parts.join("");
  }

  esc(t) {
    return t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
  }
}

export default HwpxParser;