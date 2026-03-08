/* eslint-disable @typescript-eslint/no-explicit-any */

// 외부 패키지
declare module 'jszip' {
  const JSZip: any;
  export = JSZip;
}
declare module 'yjs';
declare module 'y-websocket';
declare module 'y-prosemirror';
declare module '@tiptap/core';
declare module '@tiptap/starter-kit';
declare module '@tiptap/extension-underline';
declare module '@tiptap/extension-text-style';
declare module '@tiptap/extension-color';
declare module '@tiptap/extension-collaboration';
declare module '@tiptap/extension-collaboration-cursor';
declare module '@tiptap/extension-text-align';

// 프로젝트 내 JS 모듈
declare module './HwpxParser' {
  export class HwpxParser {
    parse(buffer: ArrayBuffer): Promise<any>;
  }
  export function getJSZip(): Promise<any>;
}

declare module './HwpxPatcher' {
  export function extractHtmlCellStyledLines(td: HTMLElement): any[];
  export function extractHtmlCellLines(td: HTMLElement): string[];
  export function patchSectionXml(xmlStr: string, cellEdits: any[], charPrList: any[], paraPrList?: any[]): string;
  export function findCharPrId(charPrList: any[], style: any, defaultRef: string): string;
  export function ensureCharPrIds(charPrList: any[], cellEdits: any[]): any[];
  export function patchHeaderXml(headerXml: string, newCharPrs: any[]): string;
}

declare module './CollabManager' {
  export class CollabManager {
    constructor(wsUrl: string, roomName: string);
    onConnectionChange: ((connected: boolean) => void) | null;
    onUserUpdate: ((users: any[]) => void) | null;
    onCellRemoteChange: ((cellKey: string, html: string) => void) | null;
    connect(): Promise<void>;
    connectOffline(): void;
    setUser(name: string, color: string): void;
    loadCells(cellOriginalTexts: Map<string, string>): void;
    bindCell(td: HTMLTableCellElement, cellKey: string, initialHtml?: string): void;
    unbindCell(cellKey: string): void;
    unbindAll(): void;
    getCellHTML(cellKey: string): string;
    getCellText(cellKey: string): string;
    getAllCellHTMLs(): Map<string, string>;
    getEditorForCell(cellKey: string): any;
    getFocusedEditor(): any;
    getFocusedCellKey(): string | null;
    getUsers(): any[];
    getCellEditors(cellKey: string): any[];
    updateCellBadges(td: HTMLTableCellElement, cellKey: string): void;
    readonly connected: boolean;
    readonly loaded: boolean;
    readonly editorCount: number;
    destroy(): void;
  }
  export function renderCellEditors(td: HTMLTableCellElement, editors: any[]): void;
  export function highlightCellBorder(td: HTMLTableCellElement, editors: any[]): void;
}

declare module './TiptapFontSize' {
  export const FontSize: any;
}

// CSS
declare module '*.css' {
  const content: any;
  export default content;
}
