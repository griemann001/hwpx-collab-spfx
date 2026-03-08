import * as React from 'react';
import { IHwpxCollabAppProps } from './IHwpxCollabAppProps';
import { HwpxParser } from './HwpxParser';
import { extractHtmlCellStyledLines, patchSectionXml } from './HwpxPatcher';
import { CollabManager } from './CollabManager';
import { SharePointService, IHwpxFileInfo } from './SharePointService';
import './tiptap-cell.css';

// JSZip은 require로 로드 (SPFx 환경 호환)
const JSZip: any = require('jszip');

/**
 * ============================================================
 * HwpxCollabApp — SPFx 웹파트 + SharePoint 문서 라이브러리 연동
 * ============================================================
 */

// ============================================================
// 툴바 명령
// ============================================================

function applyBold(collab: any): void {
  const editor: any = collab && collab.getFocusedEditor ? collab.getFocusedEditor() : null;
  if (!editor) return;
  editor.chain().focus().toggleBold().run();
}
function applyItalic(collab: any): void {
  const editor: any = collab && collab.getFocusedEditor ? collab.getFocusedEditor() : null;
  if (!editor) return;
  editor.chain().focus().toggleItalic().run();
}
function applyUnderline(collab: any): void {
  const editor: any = collab && collab.getFocusedEditor ? collab.getFocusedEditor() : null;
  if (!editor) return;
  editor.chain().focus().toggleUnderline().run();
}
function applyColor(collab: any, color: string): void {
  const editor: any = collab && collab.getFocusedEditor ? collab.getFocusedEditor() : null;
  if (!editor) return;
  editor.chain().focus().setColor(color).run();
}
function applyFontSize(collab: any, sizePt: number): void {
  const editor: any = collab && collab.getFocusedEditor ? collab.getFocusedEditor() : null;
  if (!editor) return;
  editor.chain().focus().setFontSize(sizePt + 'pt').run();
}
function applyAlign(collab: any, align: string): void {
  const editor: any = collab && collab.getFocusedEditor ? collab.getFocusedEditor() : null;
  if (!editor) return;
  editor.chain().focus().setTextAlign(align).run();
}

// ============================================================
// 타입
// ============================================================

interface IOpenedDoc {
  spFile: IHwpxFileInfo;
  html: string;
  isLandscape: boolean;
  originalBuffer: ArrayBuffer;
  pageSettings: any;
  cellOriginalTexts: Map<string, string>;
  cellOriginalHTMLs: Map<string, string>;
  charPrList: any[];
  paraPrList: any[];
  title: string;
}

// ============================================================
// 메인 컴포넌트
// ============================================================

export default class HwpxCollabApp extends React.Component<IHwpxCollabAppProps, {
  spFiles: IHwpxFileInfo[];
  selectedFileName: string | null;
  openedDoc: IOpenedDoc | null;
  showUpload: boolean;
  parsing: boolean;
  loading: boolean;
  parseLog: string[];
  collabStatus: string;
  collabUsers: any[];
}> {

  private _spService: SharePointService;
  private _fileRef: React.RefObject<HTMLInputElement>;
  private _editorRef: React.RefObject<HTMLDivElement>;
  private _collab: CollabManager | null = null;
  private _cancelled: boolean = false;

  constructor(props: IHwpxCollabAppProps) {
    super(props);
    this._spService = new SharePointService(props.spHttpClient, props.siteUrl);
    this._fileRef = React.createRef<HTMLInputElement>();
    this._editorRef = React.createRef<HTMLDivElement>();

    this.state = {
      spFiles: [],
      selectedFileName: null,
      openedDoc: null,
      showUpload: false,
      parsing: false,
      loading: false,
      parseLog: [],
      collabStatus: 'disconnected',
      collabUsers: [],
    };
  }

  public componentDidMount(): void {
    this._refreshFileList();
  }

  public componentWillUnmount(): void {
    this._cancelled = true;
    this._destroyCollab();
  }

  public componentDidUpdate(_prevProps: IHwpxCollabAppProps, prevState: any): void {
    if (prevState.selectedFileName !== this.state.selectedFileName) {
      this._openDocument();
    }
  }

  // ── SharePoint 파일 목록 ────────────────
  private _refreshFileList = async (): Promise<void> => {
    const files: IHwpxFileInfo[] = await this._spService.getFileList();
    this.setState({ spFiles: files });
  }

  // ── 파일 업로드 → SharePoint ──────────
  private _handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file: File | undefined = e.target.files ? e.target.files[0] : undefined;
    if (!file) return;

    this.setState({ parsing: true });
    this._addLog('📂 "' + file.name + '" 업로드 중...');

    try {
      const buffer: ArrayBuffer = await file.arrayBuffer();
      const result: any = await this._spService.uploadFile(file.name, buffer);
      if (!result) throw new Error('SharePoint 업로드 실패');

      this._addLog('✅ SharePoint 업로드 완료: "' + file.name + '"');
      // SharePoint 반영 대기 후 리스트 갱신
      await new Promise(function(r) { setTimeout(r, 1000); });
      await this._refreshFileList();
      this.setState({ selectedFileName: file.name });
    } catch (err: any) {
      this._addLog('❌ 업로드 실패: ' + (err.message || err));
    }

    this.setState({ parsing: false, showUpload: false });
    if (this._fileRef.current) this._fileRef.current.value = '';
  }

  // ── 문서 열기: SP 다운로드 → 파싱 → TipTap ──
  private _openDocument = async (): Promise<void> => {
    const { selectedFileName, spFiles } = this.state;

    this._destroyCollab();
    this.setState({ openedDoc: null, collabStatus: 'disconnected', collabUsers: [] });

    if (!selectedFileName) return;

    const spFile: IHwpxFileInfo | undefined = spFiles.find((f: IHwpxFileInfo) => f.name === selectedFileName);
    if (!spFile) return;

    this._cancelled = false;
    this.setState({ loading: true });

    this._addLog('📥 "' + spFile.name + '" 다운로드 중...');
    const buffer: ArrayBuffer | null = await this._spService.downloadFile(spFile.serverRelativeUrl);
    if (!buffer || this._cancelled) { this.setState({ loading: false }); return; }

    this._addLog('🔍 "' + spFile.name + '" 파싱 중...');
    const parser: any = new HwpxParser();
    const result: any = await parser.parse(buffer);
    if (this._cancelled) { this.setState({ loading: false }); return; }

    const doc: IOpenedDoc = {
      spFile: spFile,
      html: result.html,
      isLandscape: result.isLandscape || false,
      originalBuffer: buffer,
      pageSettings: result.pageSettings || null,
      cellOriginalTexts: result.cellOriginalTexts || new Map(),
      cellOriginalHTMLs: result.cellOriginalHTMLs || new Map(),
      charPrList: result.charPrList || [],
      paraPrList: result.paraPrList || [],
      title: result.title || spFile.name.replace(/\.hwpx$/i, ''),
    };

    const tableCount: number = (result.html.match(/<table/g) || []).length;
    this._addLog('✅ 파싱 완료: "' + doc.title + '" / 표 ' + tableCount + '개');

    this.setState({ openedDoc: doc, loading: false }, () => {
      // DOM 렌더 후 TipTap 마운트
      setTimeout(() => this._mountEditors(doc), 100);
    });
  }

  private _mountEditors = async (doc: IOpenedDoc): Promise<void> => {
    const container: HTMLDivElement | null = this._editorRef.current;
    if (!container || this._cancelled) return;

    container.innerHTML = doc.html;

    const fileSlug: string = doc.spFile.name.replace(/[^a-zA-Z0-9가-힣]/g, '_');
    const roomName: string = 'hwpx-' + fileSlug;
    const collab: CollabManager = new CollabManager(this.props.wsUrl, roomName);
    collab.setUser(this.props.userName, this.props.userColor);

    collab.onConnectionChange = (connected: boolean): void => {
      if (!this._cancelled) this.setState({ collabStatus: connected ? 'connected' : 'disconnected' });
    };
    collab.onUserUpdate = (users: any[]): void => {
      if (this._cancelled) return;
      this.setState({ collabUsers: users });
      this._updateBadges(container, collab);
    };

    try {
      await collab.connect();
      if (this._cancelled) { collab.destroy(); return; }

      collab.loadCells(doc.cellOriginalTexts);
      this._mountTiptapCells(container, collab, doc);
      this._collab = collab;
      this._addLog('🔗 Yjs 연결: ' + roomName + ' (' + collab.editorCount + '셀)');
    } catch (err: any) {
      console.error('Yjs 연결 실패:', err);
      this._addLog('⚠️ Yjs 연결 실패 — 오프라인 모드: ' + (err.message || err));
      collab.connectOffline();
      this._mountTiptapCells(container, collab, doc);
      this._collab = collab;
    }
  }

  private _mountTiptapCells(container: HTMLElement, collab: CollabManager, doc: IOpenedDoc): void {
    const tds: NodeListOf<Element> = container.querySelectorAll('td[data-tbl]');
    tds.forEach((td: Element) => {
      const tblIdx: string = td.getAttribute('data-tbl') || '0';
      const rowAddr: string = td.getAttribute('data-row') || '0';
      const colAddr: string = td.getAttribute('data-col') || '0';
      const cellKey: string = tblIdx + '-' + rowAddr + '-' + colAddr;
      const initialHTML: string = (doc.cellOriginalHTMLs && doc.cellOriginalHTMLs.get(cellKey)) || td.innerHTML || '';
      collab.bindCell(td as HTMLTableCellElement, cellKey, initialHTML);
    });
  }

  private _updateBadges(container: HTMLElement, collab: CollabManager): void {
    const tds: NodeListOf<Element> = container.querySelectorAll('td[data-tbl]');
    tds.forEach((td: Element) => {
      const cellKey: string = (td.getAttribute('data-tbl') || '') + '-' + (td.getAttribute('data-row') || '') + '-' + (td.getAttribute('data-col') || '');
      collab.updateCellBadges(td as HTMLTableCellElement, cellKey);
    });
  }

  // ── HWPX 저장 ────────────────────────
  private _handleSaveHwpx = async (): Promise<void> => {
    const { openedDoc } = this.state;
    if (!openedDoc || !this._editorRef.current) return;
    const collab: CollabManager | null = this._collab;

    try {
      const origZip: any = await JSZip.loadAsync(openedDoc.originalBuffer);
      const cellEdits: any[] = [];
      const origTexts: Map<string, string> = openedDoc.cellOriginalTexts;
      const charPrList: any[] = openedDoc.charPrList || [];
      const paraPrList: any[] = openedDoc.paraPrList || [];

      const tds: NodeListOf<Element> = this._editorRef.current.querySelectorAll('td[data-tbl]');
      for (let i: number = 0; i < tds.length; i++) {
        const td: Element = tds[i];
        const tblIdx: number = parseInt(td.getAttribute('data-tbl') || '0', 10);
        const rowAddr: number = parseInt(td.getAttribute('data-row') || '0', 10);
        const colAddr: number = parseInt(td.getAttribute('data-col') || '0', 10);
        const cellKey: string = tblIdx + '-' + rowAddr + '-' + colAddr;

        let htmlForExtract: string;
        if (collab) {
          htmlForExtract = collab.getCellHTML(cellKey);
        } else {
          const inner: Element | null = td.querySelector('.tiptap-cell-editor');
          htmlForExtract = inner ? inner.innerHTML : td.innerHTML;
        }

        const styledLines: any[] = extractStyledLinesFromHTML(htmlForExtract);
        const newLines: string[] = styledLines.map(function(l: any) { return l.runs.map(function(r: any) { return r.text; }).join(''); });
        const newText: string = newLines.join('\n');
        const origText: string = origTexts.get(cellKey) || '';

        const hasAlignChange: boolean = styledLines.some(function(l: any) { return l.align && l.align !== 'left'; });
        if (newText !== origText || hasAlignChange) {
          cellEdits.push({ tblIdx: tblIdx, rowAddr: rowAddr, colAddr: colAddr, newLines: newLines, styledLines: styledLines });
        }
      }

      this._addLog('🔍 변경된 셀: ' + cellEdits.length + '개');

      if (cellEdits.length > 0) {
        const secFiles: string[] = [];
        origZip.forEach(function(path: string) {
          if (path.match(/Contents\/section\d+\.xml$/i)) secFiles.push(path);
        });
        secFiles.sort();
        for (let si: number = 0; si < secFiles.length; si++) {
          const zf: any = origZip.file(secFiles[si]);
          if (!zf) continue;
          const origXml: string = await zf.async('string');
          const patchedXml: string = patchSectionXml(origXml, cellEdits, charPrList, paraPrList);
          origZip.file(secFiles[si], patchedXml);
        }
      }

      // SharePoint에 저장
      this._addLog('📤 SharePoint에 저장 중...');
      const arrayBuf: ArrayBuffer = await origZip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
      const saved: boolean = await this._spService.saveFile(openedDoc.spFile.name, arrayBuf);

      if (saved) {
        this._addLog('💾 SharePoint 저장 완료: ' + cellEdits.length + '셀 변경');
        await this._refreshFileList();
      } else {
        this._addLog('❌ SharePoint 저장 실패');
      }

    } catch (err: any) {
      this._addLog('❌ 저장 실패: ' + (err.message || err));
      console.error(err);
    }
  }

  // ── HWPX 다운로드 ────────────────────────
  private _handleDownloadHwpx = async (): Promise<void> => {
    const { openedDoc } = this.state;
    if (!openedDoc || !this._editorRef.current) return;
    const collab: CollabManager | null = this._collab;

    try {
      const origZip: any = await JSZip.loadAsync(openedDoc.originalBuffer);
      const cellEdits: any[] = [];
      const origTexts: Map<string, string> = openedDoc.cellOriginalTexts;
      const charPrList: any[] = openedDoc.charPrList || [];
      const paraPrList: any[] = openedDoc.paraPrList || [];

      const tds: NodeListOf<Element> = this._editorRef.current.querySelectorAll('td[data-tbl]');
      for (let i: number = 0; i < tds.length; i++) {
        const td: Element = tds[i];
        const tblIdx: number = parseInt(td.getAttribute('data-tbl') || '0', 10);
        const rowAddr: number = parseInt(td.getAttribute('data-row') || '0', 10);
        const colAddr: number = parseInt(td.getAttribute('data-col') || '0', 10);
        const cellKey: string = tblIdx + '-' + rowAddr + '-' + colAddr;

        let htmlForExtract: string;
        if (collab) {
          htmlForExtract = collab.getCellHTML(cellKey);
        } else {
          const inner: Element | null = td.querySelector('.tiptap-cell-editor');
          htmlForExtract = inner ? inner.innerHTML : td.innerHTML;
        }

        const styledLines: any[] = extractStyledLinesFromHTML(htmlForExtract);
        const newLines: string[] = styledLines.map(function(l: any) { return l.runs.map(function(r: any) { return r.text; }).join(''); });
        const newText: string = newLines.join('\n');
        const origText: string = origTexts.get(cellKey) || '';
        const hasAlignChange: boolean = styledLines.some(function(l: any) { return l.align && l.align !== 'left'; });
        if (newText !== origText || hasAlignChange) {
          cellEdits.push({ tblIdx: tblIdx, rowAddr: rowAddr, colAddr: colAddr, newLines: newLines, styledLines: styledLines });
        }
      }

      if (cellEdits.length > 0) {
        const secFiles: string[] = [];
        origZip.forEach(function(path: string) {
          if (path.match(/Contents\/section\d+\.xml$/i)) secFiles.push(path);
        });
        secFiles.sort();
        for (let si: number = 0; si < secFiles.length; si++) {
          const zf: any = origZip.file(secFiles[si]);
          if (!zf) continue;
          const origXml: string = await zf.async('string');
          const patchedXml: string = patchSectionXml(origXml, cellEdits, charPrList, paraPrList);
          origZip.file(secFiles[si], patchedXml);
        }
      }

      const base64: string = await origZip.generateAsync({ type: 'base64', compression: 'DEFLATE', compressionOptions: { level: 6 } });
      const dataUrl: string = 'data:application/octet-stream;base64,' + base64;
      const a: HTMLAnchorElement = document.createElement('a');
      a.href = dataUrl;
      a.download = openedDoc.spFile.name.replace(/\.hwpx$/i, '_편집.hwpx');
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      this._addLog('📥 다운로드 완료: ' + openedDoc.spFile.name);
    } catch (err: any) {
      this._addLog('❌ 다운로드 실패: ' + (err.message || err));
    }
  }

  // ── 문서 삭제 ────────────────────────
  private _handleDeleteFile = async (file: IHwpxFileInfo): Promise<void> => {
    const ok: boolean = await this._spService.deleteFile(file.serverRelativeUrl);
    if (ok) {
      this._addLog('🗑️ 삭제 완료: ' + file.name);
      if (this.state.selectedFileName === file.name) {
        this._destroyCollab();
        this.setState({ selectedFileName: null, openedDoc: null });
      }
      await this._refreshFileList();
    } else {
      this._addLog('❌ 삭제 실패: ' + file.name);
    }
  }

  private _destroyCollab(): void {
    if (this._collab) {
      this._collab.destroy();
      this._collab = null;
    }
  }

  private _addLog(msg: string): void {
    this.setState(function(prev) { return { parseLog: prev.parseLog.concat([msg]) }; });
  }

  // ── 렌더 ──────────────────────────────
  public render(): React.ReactElement<any> {
    const { userName } = this.props;
    const { spFiles, selectedFileName, openedDoc, showUpload, parsing, loading, parseLog, collabStatus, collabUsers } = this.state;
    const collab: any = this._collab;

    const S: any = {
      wrap:        { display:'flex', height:'100vh', fontFamily:"'Noto Sans KR','Malgun Gothic',sans-serif", background:'#F0F2F5', overflow:'hidden' },
      sidebar:     { width:280, minWidth:280, background:'#1B2A4A', display:'flex', flexDirection:'column', color:'#CBD5E1' },
      sideHeader:  { padding:'18px 14px 14px', borderBottom:'1px solid rgba(255,255,255,0.08)' },
      logo:        { display:'flex', alignItems:'center', gap:8, marginBottom:14 },
      logoIcon:    { width:32, height:32, background:'linear-gradient(135deg, #4F8EF7, #2563EB)', borderRadius:7, display:'flex', alignItems:'center', justifyContent:'center', fontSize:16 },
      userDot:     function(c: string) { return { width:28, height:28, borderRadius:'50%', background:c, border:'2px solid #1B2A4A', display:'flex', alignItems:'center', justifyContent:'center', fontSize:10, color:'#fff', fontWeight:700, marginRight:-6 }; },
      uploadBtn:   { width:'100%', padding:'10px 0', background:'linear-gradient(135deg, #4F8EF7, #2563EB)', color:'#fff', border:'none', borderRadius:7, fontSize:13, fontWeight:600, cursor:'pointer' },
      refreshBtn:  { width:'100%', padding:'8px 0', background:'transparent', color:'#7A8BA0', border:'1px solid rgba(255,255,255,0.12)', borderRadius:7, fontSize:12, cursor:'pointer', marginTop:6 },
      docList:     { flex:1, overflowY:'auto', padding:'6px 6px' },
      docItem:     function(sel: boolean) { return { padding:'12px 10px', borderRadius:7, cursor:'pointer', background: sel ? 'rgba(79,142,247,0.13)' : 'transparent', borderLeft: sel ? '3px solid #4F8EF7' : '3px solid transparent', marginBottom:3 }; },
      main:        { flex:1, display:'flex', flexDirection:'column' },
      toolbar:     { minHeight:52, background:'#fff', borderBottom:'1px solid #E2E8F0', display:'flex', alignItems:'center', justifyContent:'space-between', padding:'6px 18px', flexShrink:0, flexWrap:'wrap', gap:6 },
      editorWrap:  { flex:1, overflow:'auto', padding:'20px 24px', background:'#F8FAFC' },
      editorContent: function(landscape: boolean) { return { maxWidth: landscape ? 1200 : 800, margin:'0 auto', background:'#fff', borderRadius:8, boxShadow:'0 1px 3px rgba(0,0,0,0.08)', padding: landscape ? '24px 32px' : '32px 40px', minHeight:400, lineHeight:1.7, fontSize:13.5 }; },
      empty:       { flex:1, display:'flex', alignItems:'center', justifyContent:'center', flexDirection:'column', color:'#94A3B8' },
      modal:       { position:'fixed', inset:0, background:'rgba(0,0,0,0.45)', display:'flex', alignItems:'center', justifyContent:'center', zIndex:100 },
      modalBox:    { background:'#fff', borderRadius:12, padding:24, width:400, boxShadow:'0 16px 48px rgba(0,0,0,0.18)' },
      dropZone:    { border:'2px dashed #CBD5E1', borderRadius:10, padding:'36px 16px', textAlign:'center', cursor:'pointer', marginBottom:16 },
      btn1:        { padding:'7px 14px', background:'linear-gradient(135deg,#4F8EF7,#2563EB)', color:'#fff', border:'none', borderRadius:6, fontSize:12, fontWeight:600, cursor:'pointer' },
      statusDot:   function(ok: boolean) { return { width:8, height:8, borderRadius:'50%', background: ok ? '#10B981' : '#EF4444' }; },
      tbBtn:       { padding:'4px 8px', background:'#F1F5F9', color:'#475569', border:'1px solid #E2E8F0', borderRadius:5, fontSize:12, cursor:'pointer', lineHeight:1 },
    };

    return React.createElement('div', { style: S.wrap },
      // 사이드바
      React.createElement('div', { style: S.sidebar },
        React.createElement('div', { style: S.sideHeader },
          React.createElement('div', { style: S.logo },
            React.createElement('div', { style: S.logoIcon }, '📄'),
            React.createElement('div', null,
              React.createElement('div', { style: { color:'#E8ECF1', fontSize:14, fontWeight:700 } }, '한글 협업 편집기'),
              React.createElement('div', { style: { color:'#64748B', fontSize:10 } }, 'HWPX Collab (Teams) — ' + userName)
            )
          ),
          // 접속자 표시
          React.createElement('div', { style: { display:'flex', alignItems:'center', gap:6, marginBottom:12 } },
            React.createElement('div', { style: S.statusDot(collabStatus === 'connected') }),
            React.createElement('span', { style: { fontSize:10.5, color:'#7A8BA0' } },
              collabStatus === 'connected' ? collabUsers.length + '명 접속' : '대기 중'
            ),
            React.createElement('div', { style: { display:'flex', marginLeft:4 } },
              collabUsers.map(function(u: any, i: number) {
                return React.createElement('div', {
                  key: i,
                  style: S.userDot(u.color),
                  title: u.name + (u.editing ? ' — 셀 ' + u.editing + ' 편집 중' : '')
                }, u.name.charAt(0));
              })
            )
          ),
          React.createElement('button', { style: S.uploadBtn, onClick: () => this.setState({ showUpload: true }) }, '+ 문서 등록 (.hwpx)'),
          React.createElement('button', { style: S.refreshBtn, onClick: this._refreshFileList }, '🔄 목록 새로고침')
        ),
        // 문서 리스트
        React.createElement('div', { style: S.docList },
          spFiles.map((file: IHwpxFileInfo) =>
            React.createElement('div', {
              key: file.name,
              style: Object.assign({}, S.docItem(selectedFileName === file.name), { display:'flex', justifyContent:'space-between', alignItems:'center' }),
              onClick: () => this.setState({ selectedFileName: file.name })
            },
              React.createElement('div', { style: { minWidth:0, flex:1 } },
                React.createElement('div', { style: { color: selectedFileName === file.name ? '#93C5FD' : '#D1D5DB', fontSize:12.5, fontWeight: selectedFileName === file.name ? 600 : 400, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' } },
                  '📝 ' + file.name.replace(/\.hwpx$/i, '')
                ),
                React.createElement('div', { style: { color:'#5A6A80', fontSize:10, marginTop:2 } },
                  file.modifiedDate + ' · ' + file.modifiedBy
                )
              ),
              file.createdByEmail === this.props.userEmail && React.createElement('button', {
                onClick: (e: any) => { e.stopPropagation(); this._handleDeleteFile(file); },
                style: { background:'none', border:'none', color:'#5A6A80', cursor:'pointer', fontSize:13, padding:3, flexShrink:0 }
              }, '✕')
            )
          ),
          spFiles.length === 0 && React.createElement('div', { style: { textAlign:'center', padding:'36px 12px', color:'#5A6A80', fontSize:12 } },
            '등록된 문서 없음', React.createElement('br'), React.createElement('span', { style: { fontSize:11 } }, '.hwpx 파일을 업로드하세요')
          )
        ),
        // 로그
        parseLog.length > 0 && React.createElement('div', { style: { padding:'8px 10px', borderTop:'1px solid rgba(255,255,255,0.06)', maxHeight:120, overflowY:'auto' } },
          parseLog.slice(-8).map(function(log: string, i: number) {
            return React.createElement('div', { key: i, style: { fontSize:10.5, color:'#7A8BA0', lineHeight:1.5 } }, log);
          })
        ),
        React.createElement('div', { style: { padding:'10px 14px', borderTop:'1px solid rgba(255,255,255,0.06)', color:'#4A5568', fontSize:10.5 } },
          '문서 ' + spFiles.length + '개 · HWPX Collab (Teams+SP)'
        )
      ),
      // 메인
      React.createElement('div', { style: S.main },
        // 툴바
        openedDoc && React.createElement('div', { style: S.toolbar },
          React.createElement('div', { style: { display:'flex', alignItems:'center', gap:10 } },
            React.createElement('span', { style: { fontSize:15, fontWeight:700, color:'#1E293B' } }, openedDoc.title || openedDoc.spFile.name),
            React.createElement('span', { style: { fontSize:10.5, color:'#94A3B8', background:'#F1F5F9', padding:'2px 7px', borderRadius:4 } },
              '표 ' + (openedDoc.html.match(/<table/g) || []).length + '개'
            )
          ),
          React.createElement('div', { style: { display:'flex', gap:4, alignItems:'center', flexWrap:'wrap' }, onMouseDown: function(e: any) { e.preventDefault(); } },
            React.createElement('button', { onClick: function() { applyBold(collab); }, style: Object.assign({}, S.tbBtn, { fontWeight:700, fontSize:14 }), title:'굵게' }, 'B'),
            React.createElement('button', { onClick: function() { applyItalic(collab); }, style: Object.assign({}, S.tbBtn, { fontStyle:'italic', fontSize:13 }), title:'기울임' }, React.createElement('em', null, 'I')),
            React.createElement('button', { onClick: function() { applyUnderline(collab); }, style: Object.assign({}, S.tbBtn, { textDecoration:'underline' }), title:'밑줄' }, 'U'),
            React.createElement('div', { style: { width:1, height:22, background:'#E2E8F0', margin:'0 3px' } }),
            React.createElement('span', { style: { fontSize:10, color:'#94A3B8' } }, '색:'),
            [{ c:'#000000', l:'검정' }, { c:'#FF0000', l:'빨강' }, { c:'#0000FF', l:'파랑' }, { c:'#008000', l:'녹색' }, { c:'#FF6600', l:'주황' }].map(function(item) {
              return React.createElement('button', { key: item.c, onClick: function() { applyColor(collab, item.c); }, title: item.l, style: { width:20, height:20, background:item.c, border:'2px solid #E2E8F0', borderRadius:3, cursor:'pointer', padding:0 } });
            }),
            React.createElement('div', { style: { width:1, height:22, background:'#E2E8F0', margin:'0 3px' } }),
            React.createElement('span', { style: { fontSize:10, color:'#94A3B8' } }, '크기:'),
            [9, 10, 11, 12, 14, 18].map(function(sz: number) {
              return React.createElement('button', { key: sz, onClick: function() { applyFontSize(collab, sz); }, title: sz + 'pt', style: Object.assign({}, S.tbBtn, { fontSize:10, padding:'3px 5px', minWidth:24 }) }, String(sz));
            }),
            React.createElement('div', { style: { width:1, height:22, background:'#E2E8F0', margin:'0 3px' } }),
            [{ a:'left', i:'◀▬▬', t:'왼쪽' }, { a:'center', i:'▬◀▬', t:'가운데' }, { a:'right', i:'▬▬▶', t:'오른쪽' }, { a:'justify', i:'▬▬▬', t:'양쪽' }].map(function(item) {
              return React.createElement('button', { key: item.a, onClick: function() { applyAlign(collab, item.a); }, title: item.t, style: Object.assign({}, S.tbBtn, { fontSize:11, padding:'2px 5px', fontFamily:'monospace', letterSpacing:'-1px' }) }, item.i);
            }),
            React.createElement('div', { style: { width:1, height:22, background:'#E2E8F0', margin:'0 4px' } }),
            React.createElement('button', { style: S.btn1, onClick: this._handleSaveHwpx }, '💾 저장'),
            React.createElement('button', { style: Object.assign({}, S.tbBtn, { marginLeft:4 }), onClick: this._handleDownloadHwpx }, '📥 다운로드')
          )
        ),
        // 에디터 영역
        loading
          ? React.createElement('div', { style: S.empty },
              React.createElement('div', { style: { fontSize:32, marginBottom:12 } }, '⏳'),
              React.createElement('div', { style: { fontSize:14, color:'#64748B' } }, '문서 불러오는 중...')
            )
          : openedDoc
            ? React.createElement('div', { style: S.editorWrap },
                React.createElement('div', { style: S.editorContent(openedDoc.isLandscape), ref: this._editorRef })
              )
            : React.createElement('div', { style: S.empty },
                React.createElement('div', { style: { fontSize:48, opacity:0.3, marginBottom:12 } }, '📄'),
                React.createElement('div', { style: { fontSize:16, fontWeight:600, color:'#64748B', marginBottom:6 } }, '편집할 문서를 선택하세요'),
                React.createElement('div', { style: { fontSize:12.5, color:'#94A3B8' } }, '왼쪽 목록에서 .hwpx 파일을 클릭하세요')
              )
      ),
      // 업로드 모달
      showUpload && React.createElement('div', { style: S.modal, onClick: () => { if (!parsing) this.setState({ showUpload: false }); } },
        React.createElement('div', { style: S.modalBox, onClick: function(e: any) { e.stopPropagation(); } },
          React.createElement('div', { style: { fontSize:16, fontWeight:700, color:'#1E293B', marginBottom:3 } }, '문서 등록'),
          React.createElement('div', { style: { fontSize:12.5, color:'#64748B', marginBottom:16 } }, 'HWPX 파일을 업로드하면 SharePoint에 저장됩니다.'),
          React.createElement('div', { style: Object.assign({}, S.dropZone, { borderColor: parsing ? '#4F8EF7' : '#CBD5E1' }), onClick: () => { if (!parsing && this._fileRef.current) this._fileRef.current.click(); } },
            parsing
              ? React.createElement('div', { style: { color:'#4F8EF7', fontWeight:600, fontSize:13 } }, '⏳ 업로드 중...')
              : React.createElement(React.Fragment, null,
                  React.createElement('div', { style: { fontSize:32, marginBottom:6, opacity:0.4 } }, '📁'),
                  React.createElement('div', { style: { fontSize:13, fontWeight:600, color:'#475569' } }, '클릭하여 .hwpx 파일 선택'),
                  React.createElement('div', { style: { fontSize:11.5, color:'#94A3B8', marginTop:3 } }, 'SharePoint에 저장되어 모든 팀원이 볼 수 있습니다')
                )
          ),
          React.createElement('input', { ref: this._fileRef, type:'file', accept:'.hwpx', style:{ display:'none' }, onChange: this._handleFileUpload }),
          React.createElement('div', { style: { display:'flex', justifyContent:'flex-end' } },
            React.createElement('button', { onClick: () => this.setState({ showUpload: false }), disabled: parsing, style: { padding:'7px 16px', background:'#F1F5F9', color:'#475569', border:'none', borderRadius:7, fontSize:12.5, cursor:'pointer' } }, '닫기')
          )
        )
      )
    );
  }
}

// ── 헬퍼 ──
function extractStyledLinesFromHTML(htmlString: string): any[] {
  if (!htmlString) return [{ runs: [{ text: '', bold: false, color: '', fontSize: 0 }] }];
  const div: HTMLDivElement = document.createElement('div');
  div.innerHTML = htmlString;
  return extractHtmlCellStyledLines(div);
}
