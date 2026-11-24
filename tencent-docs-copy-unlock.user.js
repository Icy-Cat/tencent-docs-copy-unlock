// ==UserScript==
// @name         腾讯文档 · 解除复制限制（表格）
// @name:en      Tencent Docs Copy Unlock
// @namespace    https://github.com/Icy-Cat/tencent-docs-copy-unlock
// @version      1.1.0
// @description  解除腾讯文档表格的「禁止复制 / 禁止右键 / 禁止选择」限制；⌘C 直接复制选中区域，并带上底色、字体、边框、合并、对齐（text/plain + text/html）
// @description:en  Lift copy / right-click / text-selection restrictions on Tencent Docs spreadsheets. ⌘C copies the selection with styling (text/plain + text/html).
// @author       Icy-Cat
// @license      MIT
// @homepageURL  https://github.com/Icy-Cat/tencent-docs-copy-unlock
// @supportURL   https://github.com/Icy-Cat/tencent-docs-copy-unlock/issues
// @downloadURL  https://raw.githubusercontent.com/Icy-Cat/tencent-docs-copy-unlock/main/tencent-docs-copy-unlock.user.js
// @updateURL    https://raw.githubusercontent.com/Icy-Cat/tencent-docs-copy-unlock/main/tencent-docs-copy-unlock.user.js
// @match        https://docs.qq.com/sheet/*
// @match        https://docs.qq.com/*/sheet/*
// @match        https://docs.qq.com/*sheet/*
// @run-at       document-start
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const TAG = '[dcopy]';
  const log = (...a) => console.log('%c' + TAG, 'color:#0af;font-weight:bold', ...a);

  /* ------------------------------------------------------------------ *
   * 1. 让任何脚本都无法取消默认行为（复制 / 右键 / 选择 / 拖拽）
   *    做法：包装 addEventListener，把 copy/cut/contextmenu/selectstart
   *    的 preventDefault 变成空函数，事件本身照常派发（应用不会报错），
   *    但再也拦不住浏览器的默认动作。
   * ------------------------------------------------------------------ */
  const NOOP = function () {};
  const GUARDED = new Set(['copy', 'cut', 'contextmenu', 'selectstart', 'dragstart', 'beforecopy']);

  const origAdd = EventTarget.prototype.addEventListener;
  EventTarget.prototype.addEventListener = function (type, fn, opts) {
    if (GUARDED.has(type) && typeof fn === 'function' && !fn.__dcopy) {
      const wrapped = function (ev) {
        try { ev.preventDefault = NOOP; } catch (e) {}
        try { if (type === 'beforecopy') return; } catch (e) {}
        return fn.call(this, ev);
      };
      wrapped.__dcopy = true;
      return origAdd.call(this, type, wrapped, opts);
    }
    return origAdd.call(this, type, fn, opts);
  };

  // oncopy / oncut / oncontextmenu 这类 DOM0 属性也一起放过
  for (const prop of ['oncopy', 'oncut', 'oncontextmenu', 'onselectstart']) {
    try {
      const desc = Object.getOwnPropertyDescriptor(document, prop) ||
                   Object.getOwnPropertyDescriptor(HTMLElement.prototype, prop);
      if (desc && desc.set) {
        Object.defineProperty(document, prop, {
          configurable: true,
          get: () => null,
          set: () => {},
        });
      }
    } catch (e) {}
  }

  /* ------------------------------------------------------------------ *
   * 2. CSS：允许选择文本 / 放开水印层
   * ------------------------------------------------------------------ */
  const css = (sel, rules) => `${sel}{${rules}}`;
  function injectCSS() {
    const style = document.createElement('style');
    style.id = 'dcopy-style';
    style.textContent = [
      css('body', 'user-select: text !important; -webkit-user-select: text !important;'),
      css('.dui-modal-content, .dui-dialog', 'user-select: text !important;'),
      css('[class*="watermark"], [class*="WaterMark"]', 'opacity: 0 !important; pointer-events: none !important;'),
      css('*', '-webkit-touch-callout: default !important;'),
    ].join('\n');
    (document.head || document.documentElement).appendChild(style);
  }
  if (document.head) injectCSS();
  else origAdd.call(document, 'DOMContentLoaded', injectCSS, { once: true });

  /* ------------------------------------------------------------------ *
   * 3. 读取表格模型
   * ------------------------------------------------------------------ */
  const S = () => window.SpreadsheetApp;

  function ready() {
    const app = S();
    return !!(app && app.workbook && app.view && app.view.canvas);
  }

  function activeSheet() {
    const app = S();
    const wsm = app.workbook.worksheetManager;
    let id = null;
    try { id = app.view.canvas.spreadsheetView.getActiveSheetId(); } catch (e) {}
    if (!id) {
      try { id = wsm.sheetList[0]._KL; } catch (e) {}
    }
    const sheet = (wsm.getSheetBySheetId && wsm.getSheetBySheetId(id)) || wsm.sheetList[0];
    return { sheet, sheetId: id || (sheet && sheet._KL) };
  }

  function selectedRange(sheet) {
    const app = S();
    let sel = null;
    try { sel = app.view.canvas.normalSelectData._cpQ; } catch (e) {}
    if (!Array.isArray(sel) || !sel.length) sel = [];

    const blocks = [];
    for (const s of sel) {
      const y = s.yRange || {};   // 行
      const x = s.xRange || {};   // 列
      let r0 = Math.min(y[0] ?? 0, y[1] ?? 0);
      let r1 = Math.max(y[0] ?? 0, y[1] ?? 0);
      let c0 = Math.min(x[0] ?? 0, x[1] ?? 0);
      let c1 = Math.max(x[0] ?? 0, x[1] ?? 0);
      if (s._selectAll) {
        r0 = 0; c0 = 0;
        const ur = sheet.cellDataGrid && sheet.cellDataGrid.usedRange;
        r1 = ur && ur.endRowIndex >= 0 ? ur.endRowIndex : 0;
        c1 = ur && ur.endColIndex >= 0 ? ur.endColIndex : 0;
      }
      if (s._selectRow) { c0 = 0; c1 = Math.max(c1, lastCol(sheet)); }
      if (s._selectCol) { r0 = 0; r1 = Math.max(r1, lastRow(sheet)); }
      blocks.push({ r0, r1, c0, c1 });
    }
    if (!blocks.length) {
      const ur = sheet.cellDataGrid && sheet.cellDataGrid.usedRange;
      blocks.push({ r0: 0, r1: ur && ur.endRowIndex >= 0 ? ur.endRowIndex : 0, c0: 0, c1: ur && ur.endColIndex >= 0 ? ur.endColIndex : 0 });
    }
    const r0 = Math.min(...blocks.map(b => b.r0));
    const c0 = Math.min(...blocks.map(b => b.c0));
    const r1 = Math.max(...blocks.map(b => b.r1));
    const c1 = Math.max(...blocks.map(b => b.c1));
    return { r0, c0, r1, c1 };
  }

  const lastRow = (sheet) => {
    const ur = sheet.cellDataGrid && sheet.cellDataGrid.usedRange;
    return ur && ur.endRowIndex >= 0 ? ur.endRowIndex : 0;
  };
  const lastCol = (sheet) => {
    const ur = sheet.cellDataGrid && sheet.cellDataGrid.usedRange;
    return ur && ur.endColIndex >= 0 ? ur.endColIndex : 0;
  };

  function cellAt(sheet, r, c) {
    try { return sheet.getCellDataAtPosition(r, c) || null; } catch (e) { return null; }
  }

  function cellText(cell) {
    if (!cell) return '';
    const fmt = cell.style && cell.style.numberFormat;
    const shown = cell.formattedValue && cell.formattedValue.value;
    let raw;
    try { raw = cell.getValue(); } catch (e) { return ''; }
    if (raw === null || raw === undefined || raw === '') return '';
    if (typeof raw === 'number' || typeof raw === 'boolean') {
      return fmt && shown != null ? String(shown) : String(raw);
    }
    if (typeof raw === 'object') {
      if (raw.formulaResult) {
        const v = raw.formulaResult.value;
        if (v === null || v === undefined || v === '') return '';
        return fmt && shown != null ? String(shown) : String(v);
      }
      return '';
    }
    return String(raw);
  }

  const argb2hex = (rgb) => {
    if (!rgb || typeof rgb !== 'string' || rgb.length < 6) return null;
    const s = rgb.length === 8 ? rgb.slice(2) : rgb;
    return '#' + s.toUpperCase();
  };

  // 腾讯文档横向对齐：3 = 常规（跟随内容），不输出；1 左 / 2 居中 / 4 右 / 5 两端
  const halign = (h) => ({ 1: 'left', 2: 'center', 4: 'right', 5: 'justify' }[h] || null);
  const valign = (v) => ({ 1: 'top', 2: 'middle', 3: 'bottom' }[v] || null);

  // fill.fgColor 常常为空，颜色藏在 fill.key 里：{{{-FFFEFF00---}0{true----}224}0undefined2}
  function fillHex(fill) {
    if (!fill) return null;
    const direct = fill.fgColor && argb2hex(fill.fgColor.rgb);
    if (direct) return direct;
    const key = typeof fill.key === 'string' ? fill.key : '';
    const m = key.match(/-([0-9A-Fa-f]{8})/);
    if (!m) return null;
    return argb2hex(m[1]);
  }

  function cellCSS(cell) {
    const st = cell && cell.style;
    if (!st) return '';
    const out = [];
    const hex = fillHex(st.fill);
    if (hex && hex !== '#FFFFFF') out.push('background-color:' + hex);
    const f = st.font || {};
    if (f.name) out.push('font-family:' + f.name);
    if (f.sz) out.push('font-size:' + f.sz + 'pt');
    if (f.b) out.push('font-weight:bold');
    if (f.i) out.push('font-style:italic');
    if (f.u) out.push('text-decoration:underline');
    if (f.color && f.color.rgb) {
      const h = argb2hex(f.color.rgb);
      if (h && h !== '#000000') out.push('color:' + h);
    }
    const a = st.alignment || {};
    const ha = halign(a.horizontal);
    if (ha) out.push('text-align:' + ha);
    const va = valign(a.vertical);
    if (va) out.push('vertical-align:' + va);
    if (a.wrapText) out.push('white-space:pre-wrap');
    const b = st.border;
    if (b) {
      for (const side of ['top', 'right', 'bottom', 'left']) {
        if (b[side]) {
          const col = (b[side].color && argb2hex(b[side].color.rgb)) || '#000000';
          out.push(`border-${side}:1px solid ${col}`);
        }
      }
    }
    return out.join(';');
  }

  /* ------------------------------------------------------------------ *
   * 4. 生成 text/plain + text/html
   * ------------------------------------------------------------------ */
  function buildPayload(range, sheet) {
    const { r0, c0, r1, c1 } = range;
    const rows = [];
    const esc = (s) => String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

    // 合并单元格：记录跨度，避免重复输出
    const merges = [];
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const cell = cellAt(sheet, r, c);
        if (!cell || !cell.getMergeReference) continue;
        let m;
        try { m = cell.getMergeReference(); } catch (e) { continue; }
        if (!m) continue;
        if (m.startRowIndex === r && m.startColIndex === c) {
          merges.push({
            rs: Math.max(1, Math.min(m.endRowIndex, r1) - r + 1),
            cs: Math.max(1, Math.min(m.endColIndex, c1) - c + 1),
          });
        } else {
          merges.push(null); // 被覆盖的格子
        }
      }
    }
    const mergeAt = (r, c) => {
      const cell = cellAt(sheet, r, c);
      if (!cell || !cell.getMergeReference) return null;
      let m;
      try { m = cell.getMergeReference(); } catch (e) { return null; }
      if (!m) return null;
      if (m.startRowIndex === r && m.startColIndex === c) {
        return {
          rs: Math.max(1, Math.min(m.endRowIndex, r1) - r + 1),
          cs: Math.max(1, Math.min(m.endColIndex, c1) - c + 1),
        };
      }
      return { covered: true };
    };

    const plain = [];
    for (let r = r0; r <= r1; r++) {
      const cells = [];
      const line = [];
      for (let c = c0; c <= c1; c++) {
        const cell = cellAt(sheet, r, c);
        const text = cellText(cell);
        line.push(text.replace(/\t/g, ' ').replace(/\r?\n/g, ' '));
        const mg = mergeAt(r, c);
        if (mg && mg.covered) continue;
        const span = mg ? ` colspan="${mg.cs}" rowspan="${mg.rs}"` : '';
        const style = cellCSS(cell);
        cells.push(`<td${span}${style ? ` style="${style}"` : ''}>${esc(text).replace(/\n/g, '<br>')}</td>`);
      }
      plain.push(line.join('\t'));
      rows.push(`<tr>${cells.join('')}</tr>`);
    }

    const html =
      '<html><head><meta charset="utf-8"></head><body>' +
      '<table border="0" cellspacing="0" cellpadding="2" style="border-collapse:collapse">' +
      rows.join('') +
      '</table></body></html>';

    return { text: plain.join('\n'), html };
  }

  /* ------------------------------------------------------------------ *
   * 5. 接管 copy 事件
   * ------------------------------------------------------------------ */
  let lastPayload = null;

  function isEditable(el) {
    if (!el) return false;
    const tag = (el.tagName || '').toLowerCase();
    // 注意：表格容器本身是 contenteditable 的，不能拿它当「输入框」判断，
    // 否则 ⌘C 会被误判成“在输入框里复制”而直接放过（腾讯的提示就是这么弹出来的）
    return tag === 'input' || tag === 'textarea';
  }

  function editingCell() {
    const ae = document.activeElement;
    if (!ae) return false;
    const tag = (ae.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea') return true;
    // 单元格内联编辑器 / 公式栏
    if (ae.isContentEditable) {
      const cls = (ae.className || '').toString();
      if (/cellEdit|formula|editor/i.test(cls)) return true;
      const sel = window.getSelection && window.getSelection();
      if (sel && String(sel).trim().length > 0) return true;
    }
    return false;
  }

  function onCopy(ev) {
    if (!ready()) return;                       // 不是表格页，放过
    if (isEditable(ev.target) || editingCell()) return;  // 在输入框/单元格编辑器里复制，放过
    // 允许用户主动复制页面上的普通文本
    const sel = window.getSelection && window.getSelection();
    const hasTextSel = sel && String(sel).trim().length > 0;
    if (hasTextSel) return;

    try {
      const { sheet } = activeSheet();
      const range = selectedRange(sheet);
      const payload = buildPayload(range, sheet);
      lastPayload = payload;
      if (!payload.text && !payload.html) return;
      ev.clipboardData.setData('text/plain', payload.text);
      ev.clipboardData.setData('text/html', payload.html);
      ev.preventDefault();
      // 关键：腾讯的权限校验挂在 copy 监听里（弹「无法复制，管理员已禁止该权限」）
      // 必须同时掐断后续监听，否则提示依旧会弹
      if (ev.stopImmediatePropagation) ev.stopImmediatePropagation();
      const rows = payload.text.split('\n').length;
      const cols = payload.text.split('\n')[0].split('\t').length;
      log(`已复制 ${rows} 行 × ${cols} 列（含格式）`);
      toast(`已复制 ${rows} × ${cols}（含格式）`);
    } catch (e) {
      console.warn(TAG, 'copy failed', e);
    }
  }

  // 用 window 捕获阶段，保证早于页面自己的监听
  // 标记 __dcopy：别让自己也被 preventDefault 的补丁影响（否则默认行为会把我们的数据盖掉）
  onCopy.__dcopy = true;
  window.addEventListener('copy', onCopy, true);

  function fireCopy() {
    // 不能直接 execCommand('copy')：没有 DOM 选区时它会返回 false，剪贴板根本不会写
    // 统一走 copySelection()（隐藏容器 + 选区 + 自定义 copy 数据）
    copySelection();
  }

  function copySelection() {
    const { sheet } = activeSheet();
    writeClipboard(buildPayload(selectedRange(sheet), sheet), '选区');
  }

  function copyWholeSheet() {
    const { sheet } = activeSheet();
    const ur = sheet.cellDataGrid && sheet.cellDataGrid.usedRange;
    writeClipboard(buildPayload({
      r0: 0, c0: 0,
      r1: ur && ur.endRowIndex >= 0 ? ur.endRowIndex : 0,
      c1: ur && ur.endColIndex >= 0 ? ur.endColIndex : 0,
    }, sheet), '整表');
  }

  /* ------------------------------------------------------------------ *
   * 6. 键盘：页面把 ⌘C 吃掉时兜底
   * ------------------------------------------------------------------ */
  window.addEventListener('keydown', function (e) {
    const meta = e.metaKey || e.ctrlKey;
    if (!meta || !ready()) return;
    if (isEditable(e.target) || editingCell()) return;
    const k = (e.key || '').toLowerCase();
    if (k === 'c') {
      const sel = window.getSelection && window.getSelection();
      if (sel && String(sel).trim().length > 0) return;   // 普通文本选择
      e.stopImmediatePropagation();
      e.preventDefault();
      fireCopy();
    } else if (k === 'a' && e.shiftKey) {
      // ⌘⇧A：复制整张工作表
      e.stopImmediatePropagation();
      e.preventDefault();
      try { copyWholeSheet(); } catch (err) { console.warn(TAG, err); }
    }
  }, true);

  /* ------------------------------------------------------------------ *
   * 7. 整表复制 / 手动写剪贴板（给控制台和按钮用）
   * ------------------------------------------------------------------ */
  function writeClipboard(payload, label) {
    // 用真实 DOM 选区 + execCommand('copy')，这样 text/html 才会被写上剪贴板
    const holder = document.createElement('div');
    holder.style.cssText = 'position:fixed;left:-9999px;top:0;white-space:pre';
    holder.innerHTML = payload.html;
    document.body.appendChild(holder);

    const range = document.createRange();
    range.selectNodeContents(holder);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);

    const onCopyEv = (ev) => {
      ev.clipboardData.setData('text/plain', payload.text);
      ev.clipboardData.setData('text/html', payload.html);
      ev.preventDefault();
      if (ev.stopImmediatePropagation) ev.stopImmediatePropagation();
    };
    onCopyEv.__dcopy = true;
    document.addEventListener('copy', onCopyEv, true);
    let ok = false;
    try { ok = document.execCommand('copy'); } catch (e) {}
    document.removeEventListener('copy', onCopyEv, true);
    sel.removeAllRanges();
    holder.remove();

    if (!ok && navigator.clipboard) {
      navigator.clipboard.writeText(payload.text).catch(() => {});
    }
    lastPayload = payload;
    log('已写入剪贴板', payload.text.length, 'bytes', ok ? '' : '(降级为纯文本)');
    toast(`已复制${label || ''}（含格式）`);
    return ok;
  }

  window.TDocsCopy = {
    /** 复制当前选中区域（含格式） */
    selection() { copySelection(); },
    /** 复制整张工作表 */
    sheet() { copyWholeSheet(); },
    /** 只取数据不写剪贴板，方便调试 */
    dump(range) {
      const { sheet } = activeSheet();
      return buildPayload(range || selectedRange(sheet), sheet);
    },
    last() { return lastPayload; },
  };

  /* ------------------------------------------------------------------ *
   * 8. 小提示条
   * ------------------------------------------------------------------ */
  let toastEl = null;
  function toast(msg) {
    if (!document.body) return;
    if (!toastEl) {
      toastEl = document.createElement('div');
      toastEl.style.cssText = [
        'position:fixed', 'left:50%', 'bottom:48px', 'transform:translateX(-50%)',
        'background:rgba(23,92,235,.95)', 'color:#fff', 'font-size:13px',
        'padding:8px 14px', 'border-radius:8px', 'z-index:2147483647',
        'pointer-events:none', 'opacity:0', 'transition:opacity .18s',
        'font-family:-apple-system,"PingFang SC",sans-serif',
      ].join(';');
      document.body.appendChild(toastEl);
    }
    toastEl.textContent = msg;
    toastEl.style.opacity = '1';
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => { toastEl.style.opacity = '0'; }, 1600);
  }

  log('已加载：⌘C 复制选区（含格式）｜⌘⇧A 复制整表｜window.TDocsCopy');
})();
