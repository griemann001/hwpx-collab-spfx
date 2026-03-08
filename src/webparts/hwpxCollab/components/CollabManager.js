/**
 * ============================================================
 * CollabManager.js — Yjs 협업 바인딩 모듈 v3 (TipTap)
 * ============================================================
 *
 * v3 변경사항:
 *   - Y.Text → Y.XmlFragment  (셀마다 별도 field)
 *   - contentEditable 직접 바인딩 → TipTap Editor 인스턴스 생성
 *   - 스타일(bold/color/fontSize) 실시간 동기화 (Yjs XML 트리에 저장)
 *   - CollaborationCursor로 상대방 커서 표시
 *   - getEditorForCell() / getHTML() 으로 저장 시 HTML 추출
 *   - 오프라인 폴백: TipTap(Collaboration 없이)만 사용
 */

import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import { Editor } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import Underline from '@tiptap/extension-underline';
import TextStyle from '@tiptap/extension-text-style';
import { Color } from '@tiptap/extension-color';
import Collaboration from '@tiptap/extension-collaboration';
import CollaborationCursor from '@tiptap/extension-collaboration-cursor';
import TextAlign from '@tiptap/extension-text-align';
import { prosemirrorJSONToYXmlFragment } from 'y-prosemirror';
import { FontSize } from './TiptapFontSize';

// ── TipTap 공통 확장 (Collaboration 제외) ──
const BASE_EXTENSIONS = [
  StarterKit.configure({
    history: false,
    paragraph: {},
  }),
  Underline,
  TextStyle,
  Color,
  FontSize,
  TextAlign.configure({
    types: ['paragraph'],
    defaultAlignment: 'left',
  }),
];

export class CollabManager {
  constructor(wsUrl, roomName) {
    this.wsUrl    = wsUrl    || 'wss://collab.gfea-scratch.org';
    this.roomName = roomName || 'hwpx-default';

    this.doc      = null;
    this.provider = null;

    /** cellKey → Y.XmlFragment */
    this._xmlFragments = new Map();
    /** cellKey → TipTap Editor */
    this._editors = new Map();
    /** cellKey → cleanup 함수 */
    this._cleanups = new Map();
    /** cellKey → 원본 텍스트 (초기 로드용) */
    this._initialTexts = new Map();

    this._isConnected = false;
    this._isLoaded    = false;
    this._userName    = '';
    this._userColor   = '#4F8EF7';
    this._lastFocusedKey = null;   // ★ 마지막 포커스된 셀 키

    // 콜백
    this.onConnectionChange  = null;
    this.onUserUpdate        = null;
    this.onCellRemoteChange  = null;
  }

  // ────────────────────────────────────────────────
  // 연결
  // ────────────────────────────────────────────────

  async connect() {
    this.doc = new Y.Doc();

    this.provider = new WebsocketProvider(
      this.wsUrl, this.roomName, this.doc, { connect: true }
    );

    this.provider.on('status', (event) => {
      this._isConnected = event.status === 'connected';
      console.log(`[Collab] WebSocket: ${event.status}`);
      if (this.onConnectionChange) this.onConnectionChange(this._isConnected);
    });

    // 동기화 완료 대기
    await new Promise((resolve) => {
      if (this.provider.synced) { resolve(); return; }
      this.provider.once('synced', () => { console.log('[Collab] synced'); resolve(); });
      setTimeout(resolve, 3000);
    });

    this._applyAwareness();

    this.provider.awareness.on('change', () => {
      if (this.onUserUpdate) this.onUserUpdate(this.getUsers());
    });

    this._isLoaded = true;
  }

  /** 오프라인 모드 (WebSocket 없이 로컬 Yjs만 사용) */
  connectOffline() {
    this.doc = new Y.Doc();
    this._isConnected = false;
    this._isLoaded    = true;
    console.log('[Collab] 오프라인 모드');
  }

  setUser(name, color) {
    this._userName  = name;
    this._userColor = color || '#4F8EF7';
    this._applyAwareness();
  }

  _applyAwareness() {
    if (!this.provider?.awareness) return;
    this.provider.awareness.setLocalStateField('user', {
      name:    this._userName,
      color:   this._userColor,
      editing: null,
    });
  }

  // ────────────────────────────────────────────────
  // 셀 초기 데이터 로드
  // ────────────────────────────────────────────────

  /**
   * 초기 텍스트 저장 (bindCell 시 TipTap에 주입하기 위해 보관만 함)
   * ★ Y.XmlFragment를 직접 조작하지 않음 — TipTap이 자체 포맷으로 관리
   * @param {Map<string, string>} cellOriginalTexts  cellKey → 원본 텍스트
   */
  loadCells(cellOriginalTexts) {
    if (!this.doc) throw new Error('connect() 먼저');
    // 원본 텍스트를 보관 → bindCell에서 서버 데이터가 없을 때 초기값으로 사용
    this._initialTexts = cellOriginalTexts;
    console.log(`[Collab] ${cellOriginalTexts.size}개 셀 초기값 보관`);
  }

  _getOrCreateFragment(cellKey) {
    if (!this._xmlFragments.has(cellKey)) {
      const frag = this.doc.getXmlFragment(`cell-${cellKey}`);
      this._xmlFragments.set(cellKey, frag);
    }
    return this._xmlFragments.get(cellKey);
  }

  /**
   * Y.XmlFragment에서 텍스트 추출 (유효성 판단용)
   * 빈 paragraph만 있으면 "" 반환
   */
  _extractFragText(frag) {
    try {
      const str = frag.toString();
      return str.replace(/<[^>]*>/g, '').trim();
    } catch {
      return '';
    }
  }

  /**
   * Y.XmlFragment를 원본 텍스트로 초기화
   * prosemirrorJSONToYXmlFragment를 사용해 TipTap/y-prosemirror가
   * 인식할 수 있는 올바른 구조로 삽입
   *
   * @param {Y.XmlFragment} frag
   * @param {string} text  줄바꿈(\n) 포함 원본 텍스트
   */
  _initFragmentFromText(frag, text) {
    if (!text || frag.length > 0) return;
    try {
      const lines = text.split('\n');
      // ProseMirror doc JSON 구조 (TipTap StarterKit 기준)
      const pmJSON = {
        type: 'doc',
        content: lines.map(line => ({
          type: 'paragraph',
          content: line ? [{ type: 'text', text: line }] : [],
        })),
      };
      // y-prosemirror 공식 변환: pmJSON → Y.XmlFragment
      prosemirrorJSONToYXmlFragment(
        // schema는 편집기 생성 전이므로 null 전달 → y-prosemirror가 기본 스키마 사용
        null,
        pmJSON,
        frag
      );
      console.log(`[Collab] fragment 초기화 (pmJSON): ${lines.length}줄`);
    } catch (e) {
      console.warn('[Collab] fragment pmJSON 초기화 실패, 직접 삽입 시도:', e.message);
      // 폴백: 직접 XmlElement 삽입
      try {
        const lines = text.split('\n');
        this.doc.transact(() => {
          lines.forEach((line, li) => {
            const p = new Y.XmlElement('paragraph');
            if (line) {
              const t = new Y.XmlText();
              t.insert(0, line);
              p.insert(0, [t]);
            }
            frag.insert(li, [p]);
          });
        });
      } catch (e2) {
        console.error('[Collab] fragment 초기화 완전 실패:', e2);
      }
    }
  }

  // ────────────────────────────────────────────────
  // TipTap 에디터 생성 / 해제
  // ────────────────────────────────────────────────

  /**
   * td 엘리먼트에 TipTap 에디터를 마운트하고 Yjs에 연결
   *
   * @param {HTMLTableCellElement} td
   * @param {string} cellKey
   * @param {string} initialHtml  - HwpxParser가 생성한 초기 HTML
   */
  bindCell(td, cellKey, initialHtml = '') {
    if (!this.doc) throw new Error('connect() 먼저');
    this.unbindCell(cellKey); // 기존 바인딩 해제

    const frag = this._getOrCreateFragment(cellKey);
    const isOffline = !this.provider;

    // 원본 텍스트 → TipTap HTML
    const origText = this._initialTexts.get(cellKey) || '';
    const initialHTML = textToTiptapHTML(origText || initialHtml) || '<p></p>';

    // ★ fragment가 비어있거나 빈 paragraph만 있으면 원본 텍스트로 채움
    const fragText = this._extractFragText(frag);
    const isFragUsable = frag.length > 0 && fragText.trim() !== '';

    td.removeAttribute('contenteditable');
    td.innerHTML = '';
    td.style.padding = '2px 4px';
    td.style.verticalAlign = 'top';
    td.style.position = 'relative';

    const extensions = isOffline
      ? BASE_EXTENSIONS
      : [
          ...BASE_EXTENSIONS,
          Collaboration.configure({
            document: this.doc,
            field:    `cell-${cellKey}`,
          }),
          CollaborationCursor.configure({
            provider: this.provider,
            user: {
              name:  this._userName  || '편집자',
              color: this._userColor || '#4F8EF7',
            },
          }),
        ];

    const editor = new Editor({
      element:    td,
      extensions,
      // ★ Collaboration 모드에서도 fragment가 비어있으면 content 옵션이 Yjs에 저장됨
      // (y-prosemirror ySyncPlugin: fragment가 비면 editor.state에서 초기화)
      content: (isOffline || !isFragUsable) ? initialHTML : undefined,
      editorProps: {
        attributes: {
          class: 'tiptap-cell-editor',
          style: [
            'min-height:1.4em',
            'outline:none',
            'font-size:inherit',
            'font-family:inherit',
            'line-height:1.6',
            'padding:1px 2px',
          ].join(';'),
        },
      },
      onCreate: ({ editor: ed }) => {
        // ★ 보험: editor가 비어있는데 원본 텍스트가 있으면 강제로 채움
        if (origText.trim() && !ed.getText().trim()) {
          console.warn(`[Collab] ${cellKey}: 빈 editor 감지, setContent 적용`);
          // setContent는 Collaboration 모드에서 Yjs에 저장됨
          ed.commands.setContent(initialHTML, false);
        }
      },
      onFocus: () => {
        this._lastFocusedKey = cellKey;   // ★ 포커스 추적
        td.style.outline       = `2px solid ${this._userColor}`;
        td.style.outlineOffset = '-2px';
        td.style.background    = `${this._userColor}0d`;
        if (this.provider?.awareness) {
          this.provider.awareness.setLocalStateField('user', {
            name: this._userName, color: this._userColor, editing: cellKey,
          });
        }
      },
      onBlur: () => {
        // ★ _lastFocusedKey는 유지 (툴바 클릭 후에도 마지막 셀 기억)
        const remoteEditors = this.getCellEditors(cellKey);
        td.style.outline    = remoteEditors.length > 0 ? `2px solid ${remoteEditors[0].color}` : 'none';
        td.style.background = '';
        if (this.provider?.awareness) {
          this.provider.awareness.setLocalStateField('user', {
            name: this._userName, color: this._userColor, editing: null,
          });
        }
      },
      onUpdate: ({ editor: ed }) => {
        if (this.onCellRemoteChange) {
          // 변경 알림 (저장 버튼 활성화 등에 사용 가능)
          this.onCellRemoteChange(cellKey, ed.getHTML());
        }
      },
    });

    // 편집자 배지 컨테이너 (cursor 표시 위에 추가 배지)
    const badgeContainer = document.createElement('div');
    badgeContainer.className = 'collab-badge-container';
    badgeContainer.style.cssText = 'position:absolute;top:-14px;left:4px;display:flex;gap:3px;pointer-events:none;z-index:10;';
    td.appendChild(badgeContainer);

    this._editors.set(cellKey, editor);
    this._cleanups.set(cellKey, () => {
      badgeContainer.remove();
      editor.destroy();
    });
  }

  unbindCell(cellKey) {
    const cleanup = this._cleanups.get(cellKey);
    if (cleanup) cleanup();
    this._editors.delete(cellKey);
    this._cleanups.delete(cellKey);
  }

  unbindAll() {
    for (const k of [...this._cleanups.keys()]) this.unbindCell(k);
  }

  // ────────────────────────────────────────────────
  // 저장용 데이터 추출
  // ────────────────────────────────────────────────

  /**
   * 셀의 TipTap HTML 반환 (HwpxPatcher.extractHtmlCellStyledLines 입력용)
   * @returns {string} HTML string
   */
  getCellHTML(cellKey) {
    const editor = this._editors.get(cellKey);
    if (!editor) return '';
    return editor.getHTML();
  }

  /**
   * 셀의 plaintext 반환 (변경 감지용)
   */
  getCellText(cellKey) {
    const editor = this._editors.get(cellKey);
    if (!editor) return '';
    return editor.getText();
  }

  /**
   * 모든 셀의 HTML 반환
   * @returns {Map<string, string>}
   */
  getAllCellHTMLs() {
    const result = new Map();
    this._editors.forEach((editor, cellKey) => {
      result.set(cellKey, editor.getHTML());
    });
    return result;
  }

  /**
   * 특정 셀의 TipTap Editor 인스턴스 반환 (툴바 명령 실행용)
   */
  getEditorForCell(cellKey) {
    return this._editors.get(cellKey) || null;
  }

  /**
   * 현재 포커스된 셀의 Editor 반환
   * 툴바 버튼 클릭 시 blur가 발생해도 마지막 포커스 셀을 기억
   */
  getFocusedEditor() {
    // 1순위: 실제로 포커스된 에디터
    for (const [, editor] of this._editors) {
      if (editor.isFocused) return editor;
    }
    // 2순위: 마지막으로 포커스됐던 셀 (툴바 클릭 직후)
    if (this._lastFocusedKey) {
      return this._editors.get(this._lastFocusedKey) || null;
    }
    return null;
  }

  /**
   * 현재 포커스된 셀 키 반환
   */
  getFocusedCellKey() {
    for (const [key, editor] of this._editors) {
      if (editor.isFocused) return key;
    }
    return this._lastFocusedKey || null;
  }

  // ────────────────────────────────────────────────
  // Awareness
  // ────────────────────────────────────────────────

  getUsers() {
    if (!this.provider?.awareness) return [];
    const users = [];
    this.provider.awareness.getStates().forEach((state, clientId) => {
      if (state.user) {
        users.push({
          clientId,
          name:    state.user.name  || '익명',
          color:   state.user.color || '#94A3B8',
          editing: state.user.editing || null,
          isLocal: clientId === this.provider.awareness.clientID,
        });
      }
    });
    return users;
  }

  getCellEditors(cellKey) {
    return this.getUsers().filter(u => !u.isLocal && u.editing === cellKey);
  }

  // ────────────────────────────────────────────────
  // 편집자 배지 업데이트 (외부 호출)
  // ────────────────────────────────────────────────

  updateCellBadges(td, cellKey) {
    const container = td.querySelector('.collab-badge-container');
    if (!container) return;
    container.innerHTML = '';
    const editors = this.getCellEditors(cellKey);
    for (const ed of editors) {
      const badge = document.createElement('div');
      badge.textContent = ed.name;
      badge.style.cssText = [
        `background:${ed.color}`,
        'color:#fff',
        'font-size:9px',
        'font-weight:600',
        'padding:1px 6px',
        'border-radius:3px',
        'white-space:nowrap',
        'line-height:14px',
      ].join(';');
      container.appendChild(badge);
    }
    // 테두리
    if (editors.length > 0) {
      td.style.outline       = `2px solid ${editors[0].color}`;
      td.style.outlineOffset = '-2px';
    } else if (!this._editors.get(cellKey)?.isFocused) {
      td.style.outline = 'none';
    }
  }

  // ────────────────────────────────────────────────
  // 상태
  // ────────────────────────────────────────────────

  get connected()  { return this._isConnected; }
  get loaded()     { return this._isLoaded; }
  get editorCount(){ return this._editors.size; }

  destroy() {
    this.unbindAll();
    if (this.provider) {
      this.provider.disconnect();
      this.provider.destroy();
      this.provider = null;
    }
    if (this.doc) {
      this.doc.destroy();
      this.doc = null;
    }
    this._xmlFragments.clear();
    this._isConnected = false;
    this._isLoaded    = false;
    console.log('[Collab] 연결 해제됨');
  }
}

// ────────────────────────────────────────────────
// 하위 호환 헬퍼 (App.jsx의 renderCellEditors / highlightCellBorder 대체)
// ────────────────────────────────────────────────

/** @deprecated CollabManager.updateCellBadges() 사용 권장 */
export function renderCellEditors(td, editors) {
  let container = td.querySelector('.collab-badge-container');
  if (!container) {
    container = document.createElement('div');
    container.className  = 'collab-badge-container';
    container.style.cssText = 'position:absolute;top:-14px;left:4px;display:flex;gap:3px;pointer-events:none;z-index:10;';
    td.appendChild(container);
  }
  container.innerHTML = '';
  for (const ed of editors) {
    const badge = document.createElement('div');
    badge.textContent = ed.name;
    badge.style.cssText = `background:${ed.color};color:#fff;font-size:9px;font-weight:600;padding:1px 6px;border-radius:3px;white-space:nowrap;line-height:14px;`;
    container.appendChild(badge);
  }
}

/** @deprecated CollabManager.updateCellBadges() 사용 권장 */
export function highlightCellBorder(td, editors) {
  if (editors.length > 0) {
    td.style.outline       = `2px solid ${editors[0].color}`;
    td.style.outlineOffset = '-2px';
  } else if (document.activeElement !== td) {
    td.style.outline = 'none';
  }
}

export default CollabManager;

// ────────────────────────────────────────────────
// 내부 헬퍼
// ────────────────────────────────────────────────

/**
 * 원본 텍스트(줄바꿈 \n 포함) → TipTap content HTML
 * 각 줄을 <p>...</p>로 변환
 *
 * "줄1\n줄2\n줄3" → "<p>줄1</p><p>줄2</p><p>줄3</p>"
 * ""              → "<p></p>"
 * "단일줄"        → "<p>단일줄</p>"
 */
function textToTiptapHTML(text) {
  if (!text) return '<p></p>';
  // 이미 HTML 태그가 포함된 경우 (cellOriginalHTMLs 경로) 그대로 반환
  if (text.includes('<')) return text;
  const lines = text.split('\n');
  return lines
    .map(line => `<p>${escapeHtml(line)}</p>`)
    .join('');
}

function escapeHtml(s) {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
