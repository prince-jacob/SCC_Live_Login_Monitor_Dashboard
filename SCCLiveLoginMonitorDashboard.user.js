// ==UserScript==
// @name         SCC Live Login Monitor Dashboard - NCL1
// @namespace    prince-scc
// @version      1.8.7
// @description  Auto-detect wall, copy Pick/Pack logins with FCLM links, track SCC login changes, and show live changes on OneDrive Excel tab. Upgraded dashboard UI.
// @author       Prince Jacob ( Wprijaco )
// @match        https://staffingcommandcenter-eu.aka.amazon.com/NCL1/*
// @match        https://onedrive.live.com/edit*
// @match        https://excel.officeapps.live.com/*
// @match        file:///*
// @updateURL    https://raw.githubusercontent.com/prince-jacob/SCC_Live_Login_Monitor_Dashboard/main/SCCLiveLoginMonitorDashboard.user.js
// @downloadURL  https://raw.githubusercontent.com/prince-jacob/SCC_Live_Login_Monitor_Dashboard/main/SCCLiveLoginMonitorDashboard.user.js
// @grant        GM_setClipboard
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addValueChangeListener
// @noframes
// ==/UserScript==

(function () {
  'use strict';

  // Prevent duplicate panels inside embedded OneDrive/Excel frames.
  if (window.top !== window.self) return;

  const CREATOR = 'Prince Jacob ( Wprijaco )';
  const SCRIPT_VERSION = '1.8.7';
  const OFFICIAL_MARKER = 'OFFICIAL_SCC_LIVE_LOGIN_MONITOR_PRINCE_JACOB_V1_8_6';

  const SCC_CHANGE_KEY = 'pj_scc_live_login_changes';
  const SCC_LAYOUT_KEY = 'pj_scc_latest_layout';
  const SCC_EXCEL_LOG_KEY = 'pj_scc_excel_log';

  const STORAGE_MINIMISED = 'pj_scc_panel_minimised_v18';
  const STORAGE_EXCEL_MINIMISED = 'pj_scc_excel_panel_minimised_v18';

  const STORAGE_SCC_LEFT = 'pj_scc_panel_left_v18';
  const STORAGE_SCC_TOP = 'pj_scc_panel_top_v18';
  const STORAGE_SCC_WIDTH = 'pj_scc_panel_width_v18';
  const STORAGE_SCC_HEIGHT = 'pj_scc_panel_height_v18';

  const STORAGE_EXCEL_LEFT = 'pj_scc_excel_panel_left_v18';
  const STORAGE_EXCEL_TOP = 'pj_scc_excel_panel_top_v18';
  const STORAGE_EXCEL_WIDTH = 'pj_scc_excel_panel_width_v18';
  const STORAGE_EXCEL_HEIGHT = 'pj_scc_excel_panel_height_v18';

  const TRACK_INTERVAL_MS = 5000;

  const WALLS = {
    P1: { start: 101, end: 116 },
    P2: { start: 201, end: 216 },
    P3: { start: 301, end: 316 },
    P4: { start: 401, end: 416 }
  };

  const TABLE_SELECTOR = 'table[class*="AfeFloorPlanStationsTableV2-module__container"]';

  let lastSnapshot = null;
  let changeHistory = [];
  let trackingTimer = null;
  let trackingPaused = false;
  let checkingNow = false;

  function isSccPage() {
    return location.href.startsWith('https://staffingcommandcenter-eu.aka.amazon.com/NCL1/') ||
           location.protocol === 'file:';
  }

  function isExcelPage() {
    return location.href.includes('onedrive.live.com/edit') ||
           location.href.includes('excel.officeapps.live.com');
  }

  function clean(text) {
    return String(text || '')
      .replace(/\u00a0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function escapeHtml(text) {
    return String(text || '')
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#039;');
  }

  function makeEmployeeHref(login) {
    if (!login) return '';

    return 'https://fclm-portal.amazon.com/employee/activityDetails?employeeId='
      + encodeURIComponent(login)
      + '&warehouseId=NCL1';
  }

  function makeClipboardLink(login, href) {
    const safeLogin = escapeHtml(login || '');
    const safeHref = escapeHtml(href || makeEmployeeHref(login));

    if (!safeLogin) return '';

    return `<a href="${safeHref}">${safeLogin}</a>`;
  }

  function getDirectCells(row) {
    if (!row) return [];

    return [...row.children].filter(el =>
      el.tagName === 'TD' || el.tagName === 'TH'
    );
  }

  function parseStationHeader(text) {
    const match = clean(text).match(/^(\d+)\s*\|\s*Priority\s*(\d+)/i);

    if (!match) return null;

    return {
      station: Number(match[1]),
      priority: Number(match[2])
    };
  }

  function isBadLoginCandidate(login) {
    if (!login) return true;

    const value = clean(login).toLowerCase();

    if (!value) return true;
    if (value === '--') return true;
    if (value.length < 3 || value.length > 20) return true;
    if (/^ws/i.test(value)) return true;
    if (/^ws[-_]/i.test(value)) return true;
    if (/picktorebin/i.test(value)) return true;
    if (/rebin|station|priority|uph|pack\s*uph|pick\s*uph/i.test(value)) return true;
    if (value.includes('_')) return true;
    if (/^\d+$/.test(value)) return true;
    if (/^\d+\s*\|/.test(value)) return true;
    if (/^p[1-4]$/i.test(value)) return true;
    if (/^pack\d*$/i.test(value)) return true;
    if (/^pick$/i.test(value)) return true;
    if (/[^a-z0-9.-]/i.test(value)) return true;

    return false;
  }

  function parseAssociateCell(cell) {
    if (!cell) {
      return { login: '', href: '' };
    }

    const text = clean(cell.innerText);

    if (!text || text === '--') {
      return { login: '', href: '' };
    }

    let login = '';

    // Method 1: original SCC employee data-testid.
    const employeeNode = cell.querySelector(
      '[data-testid^="floor-plan-employee-"]'
    );

    if (employeeNode) {
      const testId = employeeNode.getAttribute('data-testid') || '';
      login = testId.replace(/^floor-plan-employee-/i, '').trim();
    }

    // Method 2: alternative SCC employee/associate/login data-testid.
    if (!login) {
      const nodes = [...cell.querySelectorAll('[data-testid]')];

      for (const node of nodes) {
        const testId = node.getAttribute('data-testid') || '';
        const match = testId.match(
          /(?:employee|associate|login)[-_]([a-z0-9.-]+)$/i
        );

        if (match) {
          login = match[1];
          break;
        }
      }
    }

    // Method 3: fallback to the first valid login-like visible token.
    if (!login) {
      const tokens = text
        .split(/\s+/)
        .map(token => token.replace(/[^a-z0-9.-]/gi, ''))
        .filter(Boolean);

      login = tokens.find(token => {
        const value = token.toLowerCase();

        return (
          /^[a-z][a-z0-9.-]{2,20}$/i.test(token) &&
          !isBadLoginCandidate(value)
        );
      }) || '';
    }

    login = clean(login)
      .replace(/[^a-z0-9.-]/gi, '')
      .trim();

    if (isBadLoginCandidate(login)) {
      login = '';
    }

    return {
      login,
      href: login ? makeEmployeeHref(login) : ''
    };
  }

  function injectSccStyles() {
    if (document.getElementById('pj-scc-style')) return;

    const style = document.createElement('style');
    style.id = 'pj-scc-style';

    style.textContent = `
      #pj-scc-panel, #pj-scc-panel * {
        box-sizing: border-box !important;
        font-family: Inter, "Segoe UI", Roboto, Arial, sans-serif !important;
        text-transform: none !important;
        letter-spacing: normal !important;
        line-height: 1.35 !important;
      }

      #pj-scc-panel {
        position: fixed !important;
        top: 88px !important;
        right: 18px !important;
        z-index: 999999 !important;
        width: 470px !important;
        color: #e5e7eb !important;
        background: rgba(15, 23, 42, 0.97) !important;
        border: 1px solid rgba(255,255,255,0.08) !important;
        border-radius: 18px !important;
        box-shadow: 0 18px 40px rgba(0,0,0,0.35), 0 6px 14px rgba(0,0,0,0.18) !important;
        padding: 14px !important;
      }

      #pj-scc-panel.pj-min {
        width: 142px !important;
        min-width: 0 !important;
        height: auto !important;
        min-height: 0 !important;
        padding: 0 !important;
        border-radius: 12px !important;
        resize: none !important;
      }

      #pj-scc-panel.pj-min .pj-scc-body {
        display: none !important;
      }

      .pj-scc-header {
        cursor: move !important;
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        gap: 10px !important;
      }

      .pj-scc-title {
        font-size: 18px !important;
        font-weight: 800 !important;
        color: #ffffff !important;
      }

      .pj-scc-subtitle {
        font-size: 11px !important;
        color: #94a3b8 !important;
        margin-top: 2px !important;
      }

      #pj-scc-panel.pj-min .pj-scc-title {
        font-size: 14px !important;
      }

      #pj-scc-panel.pj-min .pj-scc-subtitle {
        display: none !important;
      }

      .pj-scc-icon-btn {
        border: 0 !important;
        background: rgba(255,255,255,0.08) !important;
        color: #ffffff !important;
        width: 30px !important;
        height: 30px !important;
        border-radius: 10px !important;
        cursor: pointer !important;
        font-size: 18px !important;
        font-weight: 800 !important;
      }

      .pj-scc-body {
        margin-top: 12px !important;
      }

      .pj-scc-actions {
        display: flex !important;
        gap: 8px !important;
        flex-wrap: wrap !important;
        margin-bottom: 12px !important;
      }

      .pj-scc-btn {
        border: 0 !important;
        border-radius: 12px !important;
        padding: 9px 12px !important;
        cursor: pointer !important;
        font-size: 12px !important;
        font-weight: 800 !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15) !important;
      }

      .pj-copy-btn {
        background: linear-gradient(135deg, #10b981, #059669) !important;
        color: #ffffff !important;
      }

      .pj-recheck-btn {
        background: #e5e7eb !important;
        color: #111827 !important;
      }

      .pj-refresh-btn {
        background: linear-gradient(135deg, #facc15, #eab308) !important;
        color: #111827 !important;
      }

      .pj-scc-card {
        background: rgba(255,255,255,0.05) !important;
        border: 1px solid rgba(255,255,255,0.06) !important;
        border-radius: 14px !important;
        padding: 10px 12px !important;
        margin-top: 10px !important;
      }

      .pj-scc-note {
        font-size: 12px !important;
        color: #cbd5e1 !important;
        margin: 0 0 6px 0 !important;
      }

      .pj-scc-meta {
        font-size: 11px !important;
        color: #94a3b8 !important;
        margin: 0 !important;
      }

      .pj-scc-status {
        margin-top: 12px !important;
        padding: 10px 12px !important;
        border-radius: 12px !important;
        background: rgba(59,130,246,0.12) !important;
        border: 1px solid rgba(59,130,246,0.22) !important;
        color: #dbeafe !important;
        font-size: 12px !important;
        font-weight: 700 !important;
      }

      .pj-scc-changes {
        margin-top: 12px !important;
        background: rgba(255,255,255,0.05) !important;
        border: 1px solid rgba(255,255,255,0.07) !important;
        border-radius: 14px !important;
        padding: 10px !important;
        max-height: 150px !important;
        overflow: auto !important;
      }

      .pj-changes-title {
        font-size: 12px !important;
        font-weight: 800 !important;
        color: #ffffff !important;
        margin-bottom: 8px !important;
      }

      .pj-change {
        display: flex !important;
        gap: 8px !important;
        align-items: flex-start !important;
        padding: 7px 8px !important;
        margin-bottom: 6px !important;
        border-radius: 10px !important;
        font-size: 12px !important;
        color: #e5e7eb !important;
        background: rgba(15, 23, 42, 0.75) !important;
      }

      .pj-change-time {
        color: #94a3b8 !important;
        font-size: 11px !important;
        white-space: nowrap !important;
      }

      .pj-change-moved {
        border-left: 3px solid #facc15 !important;
      }

      .pj-change-added {
        border-left: 3px solid #22c55e !important;
      }

      .pj-change-removed {
        border-left: 3px solid #ef4444 !important;
      }

      .pj-no-changes {
        color: #94a3b8 !important;
        font-size: 12px !important;
      }

      .pj-scc-preview-wrap {
        margin-top: 12px !important;
        background: #ffffff !important;
        color: #111827 !important;
        border-radius: 14px !important;
        border: 1px solid #dbe3ee !important;
        overflow: hidden !important;
      }

      .pj-scc-preview-head {
        padding: 12px 14px !important;
        background: linear-gradient(180deg, #f8fafc, #f1f5f9) !important;
        border-bottom: 1px solid #e2e8f0 !important;
        font-size: 13px !important;
        font-weight: 800 !important;
        color: #0f172a !important;
      }

      .pj-scc-preview-scroll {
        max-height: 330px !important;
        overflow: auto !important;
      }

      .pj-scc-table {
        width: 100% !important;
        border-collapse: collapse !important;
        font-size: 13px !important;
      }

      .pj-scc-table th {
        position: sticky !important;
        top: 0 !important;
        z-index: 1 !important;
        background: #f8fafc !important;
        color: #334155 !important;
        text-align: left !important;
        font-size: 12px !important;
        font-weight: 800 !important;
        padding: 10px 12px !important;
        border-bottom: 1px solid #e2e8f0 !important;
      }

      .pj-scc-table td {
        padding: 9px 12px !important;
        border-bottom: 1px solid #eef2f7 !important;
        color: #111827 !important;
      }

      .pj-scc-table tr:nth-child(even) td {
        background: #fcfdff !important;
      }

      .pj-station-col {
        width: 76px !important;
        font-weight: 800 !important;
        color: #1e293b !important;
      }

      .pj-empty {
        color: #94a3b8 !important;
      }

      #pj-scc-panel {
        left: 20px !important;
        right: auto !important;
        width: 520px !important;
        height: 680px !important;
        max-height: calc(100vh - 75px) !important;
        min-width: 340px !important;
        min-height: 170px !important;
        max-width: 95vw !important;
        resize: both !important;
        overflow: hidden !important;
        background: #111827 !important;
        border-radius: 14px !important;
        padding: 0 !important;
      }

      #pj-scc-panel.pj-min {
        width: 142px !important;
        min-width: 0 !important;
        height: auto !important;
        min-height: 0 !important;
        resize: none !important;
        overflow: hidden !important;
      }

      #pj-scc-panel.pj-min .pj-scc-header {
        padding: 6px 7px !important;
        gap: 5px !important;
        cursor: move !important;
      }

      #pj-scc-panel.pj-min .pj-scc-title {
        max-width: 92px !important;
        font-size: 12px !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
      }

      #pj-scc-panel.pj-min .pj-scc-icon-btn {
        width: 22px !important;
        height: 22px !important;
        border-radius: 7px !important;
        font-size: 14px !important;
        line-height: 20px !important;
      }

      .pj-scc-header {
        padding: 12px 14px !important;
        background: linear-gradient(135deg, #232f3e, #111827) !important;
        border-bottom: 1px solid rgba(255,255,255,0.1) !important;
        cursor: move !important;
      }

      .pj-scc-body {
        padding: 12px !important;
        overflow: auto !important;
        height: calc(100% - 74px) !important;
        margin-top: 0 !important;
      }

      .pj-pause-btn { background: #b45309 !important; color: #fff !important; }
      .pj-copy-changes-btn { background: #2563eb !important; color: #fff !important; }

      .pj-scc-stats {
        display: grid !important;
        grid-template-columns: repeat(4, 1fr) !important;
        gap: 7px !important;
        margin-bottom: 10px !important;
      }

      .pj-scc-stat {
        background: #1f2937 !important;
        border-radius: 10px !important;
        padding: 8px !important;
        border: 1px solid rgba(255,255,255,0.08) !important;
        text-align: center !important;
      }

      .pj-scc-stat b { display: block !important; font-size: 18px !important; line-height: 18px !important; color: #fff !important; }
      .pj-scc-stat span { display: block !important; font-size: 10px !important; color: #d1d5db !important; margin-top: 3px !important; }

      #pj-scc-toast-host {
        position: fixed !important;
        right: 20px !important;
        bottom: 24px !important;
        z-index: 2147483647 !important;
        display: flex !important;
        flex-direction: column !important;
        gap: 8px !important;
        pointer-events: none !important;
      }
      .pj-scc-toast { min-width: 170px !important; max-width: 360px !important; opacity: 0 !important; transform: translateY(12px) !important; transition: opacity .25s ease, transform .25s ease !important; padding: 11px 14px !important; border-radius: 999px !important; color: #fff !important; font-size: 12px !important; font-weight: 800 !important; box-shadow: 0 12px 30px rgba(0,0,0,.35) !important; text-align: center !important; }
      .pj-scc-toast.show { opacity: 1 !important; transform: translateY(0) !important; }
      .pj-scc-toast.info { background: #374151 !important; }
      .pj-scc-toast.success { background: #047857 !important; }
      .pj-scc-toast.error { background: #b91c1c !important; }

    `;

    document.head.appendChild(style);
  }

  function injectExcelStyles() {
    if (document.getElementById('pj-scc-excel-style')) return;

    const style = document.createElement('style');
    style.id = 'pj-scc-excel-style';

    style.textContent = `
      #pj-scc-live-window, #pj-scc-live-window * {
        box-sizing: border-box !important;
        font-family: Inter, "Segoe UI", Roboto, Arial, sans-serif !important;
      }

      #pj-scc-live-window {
        position: fixed !important;
        top: 88px !important;
        right: 18px !important;
        z-index: 999999 !important;
        width: 390px !important;
        background: rgba(15, 23, 42, 0.97) !important;
        color: #e5e7eb !important;
        border: 1px solid rgba(255,255,255,0.08) !important;
        border-radius: 18px !important;
        box-shadow: 0 18px 40px rgba(0,0,0,0.35) !important;
        padding: 14px !important;
      }

      #pj-scc-live-window.min {
        width: 145px !important;
        padding: 6px 7px !important;
      }

      #pj-scc-live-window.min .pj-live-body {
        display: none !important;
      }

      .pj-live-header {
        cursor: move !important;
        display: flex !important;
        justify-content: space-between !important;
        align-items: center !important;
        gap: 10px !important;
      }

      .pj-live-title {
        font-size: 17px !important;
        font-weight: 800 !important;
        color: #fff !important;
      }

      .pj-live-subtitle {
        font-size: 11px !important;
        color: #94a3b8 !important;
        margin-top: 2px !important;
      }

      #pj-scc-live-window.min .pj-live-title {
        font-size: 13px !important;
      }

      #pj-scc-live-window.min .pj-live-subtitle {
        display: none !important;
      }

      .pj-live-min {
        border: 0 !important;
        width: 30px !important;
        height: 30px !important;
        border-radius: 10px !important;
        background: rgba(255,255,255,0.08) !important;
        color: #fff !important;
        font-weight: 800 !important;
        cursor: pointer !important;
      }

      .pj-live-body {
        margin-top: 12px !important;
      }

      .pj-live-status {
        padding: 10px 12px !important;
        border-radius: 12px !important;
        background: rgba(59,130,246,0.12) !important;
        border: 1px solid rgba(59,130,246,0.22) !important;
        color: #dbeafe !important;
        font-size: 12px !important;
        font-weight: 700 !important;
        margin-bottom: 10px !important;
      }

      .pj-live-actions {
        display: flex !important;
        gap: 8px !important;
        margin-bottom: 10px !important;
      }

      .pj-live-btn {
        border: 0 !important;
        border-radius: 11px !important;
        padding: 9px 12px !important;
        font-size: 12px !important;
        font-weight: 800 !important;
        cursor: pointer !important;
      }

      .pj-copy-latest {
        background: #10b981 !important;
        color: white !important;
      }

      .pj-clear-log {
        background: #334155 !important;
        color: white !important;
      }

      .pj-live-log {
        max-height: 330px !important;
        overflow: auto !important;
        display: flex !important;
        flex-direction: column !important;
        gap: 8px !important;
      }

      .pj-live-item {
        padding: 9px 10px !important;
        border-radius: 12px !important;
        background: rgba(255,255,255,0.06) !important;
        border-left: 4px solid #64748b !important;
        font-size: 12px !important;
      }

      .pj-live-item.moved {
        border-left-color: #facc15 !important;
      }

      .pj-live-item.added {
        border-left-color: #22c55e !important;
      }

      .pj-live-item.removed {
        border-left-color: #ef4444 !important;
      }

      .pj-live-time {
        color: #94a3b8 !important;
        font-size: 11px !important;
        margin-bottom: 3px !important;
      }

      .pj-live-empty {
        color: #94a3b8 !important;
        font-size: 12px !important;
        padding: 12px !important;
        text-align: center !important;
      }

      .pj-live-creator {
        margin-top: 10px !important;
        font-size: 11px !important;
        color: #94a3b8 !important;
      }

      #pj-scc-live-window {
        left: 20px !important;
        right: auto !important;
        width: 430px !important;
        height: 620px !important;
        max-height: calc(100vh - 75px) !important;
        min-width: 320px !important;
        min-height: 160px !important;
        resize: both !important;
        overflow: hidden !important;
        background: #111827 !important;
        border-radius: 14px !important;
        padding: 0 !important;
      }

      .pj-live-header {
        padding: 12px 14px !important;
        background: linear-gradient(135deg, #232f3e, #111827) !important;
        border-bottom: 1px solid rgba(255,255,255,0.1) !important;
        cursor: move !important;
      }

      .pj-live-body {
        padding: 12px !important;
        overflow: auto !important;
        height: calc(100% - 74px) !important;
        margin-top: 0 !important;
      }

      #pj-scc-live-window.min {
        width: 142px !important;
        min-width: 0 !important;
        height: auto !important;
        min-height: 0 !important;
        resize: none !important;
        overflow: hidden !important;
      }

      #pj-scc-live-window.min .pj-live-header {
        padding: 6px 7px !important;
        gap: 5px !important;
        cursor: move !important;
      }

      #pj-scc-live-window.min .pj-live-title {
        max-width: 92px !important;
        font-size: 12px !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
      }

      #pj-scc-live-window.min .pj-live-min {
        width: 22px !important;
        height: 22px !important;
        border-radius: 7px !important;
        font-size: 14px !important;
        line-height: 20px !important;
      }

    `;

    document.head.appendChild(style);
  }



  function showToast(message, type = 'info') {
    let host = document.getElementById('pj-scc-toast-host');

    if (!host) {
      host = document.createElement('div');
      host.id = 'pj-scc-toast-host';
      document.body.appendChild(host);
    }

    const item = document.createElement('div');
    item.className = `pj-scc-toast ${type}`;
    item.textContent = message;
    host.appendChild(item);

    setTimeout(() => item.classList.add('show'), 20);
    setTimeout(() => {
      item.classList.remove('show');
      setTimeout(() => item.remove(), 350);
    }, 2600);
  }

  function makePanelDraggable(panelId, headerSelector, leftKey, topKey) {
    const panel = document.getElementById(panelId);
    const header = panel?.querySelector(headerSelector);
    if (!panel || !header) return;

    const savedLeft = localStorage.getItem(leftKey);
    const savedTop = localStorage.getItem(topKey);
    if (savedLeft && savedTop) {
      panel.style.setProperty('left', savedLeft, 'important');
      panel.style.setProperty('top', savedTop, 'important');
      panel.style.setProperty('right', 'auto', 'important');
    }

    let dragging = false;
    let offsetX = 0;
    let offsetY = 0;

    header.addEventListener('mousedown', e => {
      if (e.target.closest('button')) return;
      dragging = true;
      const rect = panel.getBoundingClientRect();
      offsetX = e.clientX - rect.left;
      offsetY = e.clientY - rect.top;
      document.body.style.userSelect = 'none';
    });

    document.addEventListener('mousemove', e => {
      if (!dragging) return;
      let left = e.clientX - offsetX;
      let top = e.clientY - offsetY;
      left = Math.max(0, Math.min(left, window.innerWidth - panel.offsetWidth));
      top = Math.max(0, Math.min(top, window.innerHeight - 50));
      panel.style.setProperty('left', `${left}px`, 'important');
      panel.style.setProperty('top', `${top}px`, 'important');
      panel.style.setProperty('right', 'auto', 'important');
    });

    document.addEventListener('mouseup', () => {
      if (!dragging) return;
      dragging = false;
      document.body.style.userSelect = '';
      localStorage.setItem(leftKey, panel.style.left);
      localStorage.setItem(topKey, panel.style.top);
    });
  }

  function rememberPanelSize(panelId, widthKey, heightKey, minimizedCheck) {
    const panel = document.getElementById(panelId);
    if (!panel) return;

    const savedWidth = localStorage.getItem(widthKey);
    const savedHeight = localStorage.getItem(heightKey);
    if (savedWidth && !minimizedCheck()) panel.style.width = savedWidth;
    if (savedHeight && !minimizedCheck()) panel.style.height = savedHeight;

    let lastWidth = panel.offsetWidth;
    let lastHeight = panel.offsetHeight;

    setInterval(() => {
      if (minimizedCheck()) return;
      if (panel.offsetWidth !== lastWidth || panel.offsetHeight !== lastHeight) {
        lastWidth = panel.offsetWidth;
        lastHeight = panel.offsetHeight;
        localStorage.setItem(widthKey, `${lastWidth}px`);
        localStorage.setItem(heightKey, `${lastHeight}px`);
      }
    }, 1000);
  }

  function renderSccStats(rows, wallName) {
    const stats = document.getElementById('pj-scc-stats');
    if (!stats) return;

    const uniquePickers = new Set();
    const uniquePackers = new Set();

    rows.forEach(row => {
      const pickLogin = clean(row.pickLogin).toLowerCase();
      const pack1Login = clean(row.pack1Login).toLowerCase();
      const pack2Login = clean(row.pack2Login).toLowerCase();

      if (pickLogin && !isBadLoginCandidate(pickLogin)) {
        uniquePickers.add(pickLogin);
      }

      if (pack1Login && !isBadLoginCandidate(pack1Login)) {
        uniquePackers.add(pack1Login);
      }

      if (pack2Login && !isBadLoginCandidate(pack2Login)) {
        uniquePackers.add(pack2Login);
      }
    });

    const pickers = uniquePickers.size;
    const packers = uniquePackers.size;
    const staffed = pickers + packers;

    stats.innerHTML = `
      <div class="pj-scc-stat"><b>${escapeHtml(wallName || '-')}</b><span>Wall</span></div>
      <div class="pj-scc-stat"><b>${pickers}</b><span>Pickers</span></div>
      <div class="pj-scc-stat"><b>${packers}</b><span>Packers</span></div>
      <div class="pj-scc-stat"><b>${staffed}</b><span>Staffed</span></div>
    `;
  }

  function buildChangesText(items = changeHistory.slice(0, 30)) {
    if (!items.length) return 'No login changes detected.';

    return items.map(item => [
      item.time || '',
      item.type || '',
      item.login || '',
      item.text || ''
    ].join('\t')).join('\n');
  }

  function copyLatestChanges() {
    const text = buildChangesText();
    GM_setClipboard(text);
    showToast('Latest changes copied', 'success');
  }

  function toggleTrackingPause() {
    trackingPaused = !trackingPaused;
    const btn = document.getElementById('pj-pause-track-btn');
    const status = document.getElementById('pj-scc-status');
    if (btn) btn.textContent = trackingPaused ? 'Resume' : 'Pause';
    if (status) status.textContent = trackingPaused ? 'Tracking paused.' : 'Tracking resumed. Watching for login changes...';
    showToast(trackingPaused ? 'Tracking paused' : 'Tracking resumed', trackingPaused ? 'error' : 'success');
  }

  // Find station tables by visible content instead of Amazon's generated CSS class.
  // This is more reliable across different SCC builds and user accounts.
  function getStationTables() {
    return [...document.querySelectorAll('table')]
      .filter(table => {
        const headerText = clean(table.tHead?.innerText || table.innerText);
        return /\b\d{3}\s*\|\s*Priority\s*\d+/i.test(headerText);
      });
  }

  function captureSccFloorPlan() {
    const tables = getStationTables();
    const stations = [];

    tables.forEach(table => {
      const headerCells = getDirectCells(table.tHead?.rows?.[0]);

      const stationHeaders = headerCells
        .map(cell => parseStationHeader(cell.innerText))
        .filter(Boolean);

      if (!stationHeaders.length) return;

      const bodyRows = [];

      [...table.tBodies].forEach(tbody => {
        [...tbody.rows].forEach(row => {
          const cells = getDirectCells(row);
          if (cells.length) bodyRows.push(cells);
        });
      });

      const pickUphRowIndex = bodyRows.findIndex(row =>
        row.some(cell => /^Pick\s+UPH:/i.test(clean(cell.innerText)))
      );

      const packUphRowIndex = bodyRows.findIndex(row =>
        row.some(cell => /^Pack\s+UPH:/i.test(clean(cell.innerText)))
      );

      const pickAssociateRow =
        pickUphRowIndex >= 0 ? bodyRows[pickUphRowIndex + 1] || [] : [];

      const packAssociateRow1 =
        packUphRowIndex >= 0 ? bodyRows[packUphRowIndex + 1] || [] : [];

      const packAssociateRow2 =
        packUphRowIndex >= 0 ? bodyRows[packUphRowIndex + 2] || [] : [];

      stationHeaders.forEach((header, index) => {
        const picker = parseAssociateCell(pickAssociateRow[index]);
        const packer1 = parseAssociateCell(packAssociateRow1[index]);
        const packer2 = parseAssociateCell(packAssociateRow2[index]);

        stations.push({
          station: header.station,
          priority: header.priority,

          pickLogin: picker.login,
          pickHref: picker.href,

          pack1Login: packer1.login,
          pack1Href: packer1.href,

          pack2Login: packer2.login,
          pack2Href: packer2.href
        });
      });
    });

    return stations.sort((a, b) => a.station - b.station);
  }

  function detectWallName(stations) {
    const stationNumbers = stations.map(x => Number(x.station)).filter(Boolean);

    if (!stationNumbers.length) return '';

    const minStation = Math.min(...stationNumbers);

    if (minStation >= 101 && minStation <= 116) return 'P1';
    if (minStation >= 201 && minStation <= 216) return 'P2';
    if (minStation >= 301 && minStation <= 316) return 'P3';
    if (minStation >= 401 && minStation <= 416) return 'P4';

    return `P${String(minStation)[0]}`;
  }

  function getAutoWallRows() {
    const allStations = captureSccFloorPlan();
    const wallName = detectWallName(allStations);

    if (!wallName || !WALLS[wallName]) {
      return { wallName: '', rows: [] };
    }

    const wall = WALLS[wallName];
    const byStation = new Map();

    allStations.forEach(row => {
      byStation.set(Number(row.station), row);
    });

    const rows = [];

    for (let st = wall.start; st <= wall.end; st++) {
      const found = byStation.get(st);

      rows.push({
        station: st,

        pickLogin: found?.pickLogin || '',
        pickHref: found?.pickHref || '',

        pack1Login: found?.pack1Login || '',
        pack1Href: found?.pack1Href || '',

        pack2Login: found?.pack2Login || '',
        pack2Href: found?.pack2Href || ''
      });
    }

    return { wallName, rows };
  }

  function hasAnyLogin(rows) {
    return Array.isArray(rows) && rows.some(row =>
      row.pickLogin || row.pack1Login || row.pack2Login
    );
  }

  function formatCell(value) {
    return value ? escapeHtml(value) : '<span class="pj-empty">—</span>';
  }

  function buildPreviewHtml(rows, wallName) {
    const bodyRows = rows.map(row => {
      const pickerCount = row.pickLogin ? 1 : 0;
      const packerCount = (row.pack1Login ? 1 : 0) + (row.pack2Login ? 1 : 0);
      const pickerTitle = row.pickLogin ? ` title="${escapeHtml(row.pickLogin)}"` : '';
      const packerTitle = [row.pack1Login, row.pack2Login].filter(Boolean).join(' / ');
      const packerTitleAttr = packerTitle ? ` title="${escapeHtml(packerTitle)}"` : '';

      return `
        <tr>
          <td class="pj-station-col">${row.station}</td>
          <td${pickerTitle}><b>${pickerCount}</b></td>
          <td${packerTitleAttr}><b>${packerCount}</b></td>
        </tr>
      `;
    }).join('');

    return `
      <div class="pj-scc-preview-wrap">
        <div class="pj-scc-preview-head">Auto-detected wall: ${escapeHtml(wallName || 'Unknown')} • showing counts only</div>
        <div class="pj-scc-preview-scroll">
          <table class="pj-scc-table">
            <thead>
              <tr>
                <th>Station</th>
                <th>Pickers</th>
                <th>Packers</th>
              </tr>
            </thead>
            <tbody>
              ${bodyRows || `<tr><td colspan="3" style="padding:14px;">No station data found.</td></tr>`}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  function buildClipboardHtml(rows) {
    let html = `
      <table border="1" cellspacing="0" cellpadding="4"
        style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:12px;">
    `;

    rows.forEach(row => {
      html += `
        <tr>
          <td>${makeClipboardLink(row.pickLogin, row.pickHref)}</td>
          <td>${makeClipboardLink(row.pack1Login, row.pack1Href)}</td>
          <td>${makeClipboardLink(row.pack2Login, row.pack2Href)}</td>
        </tr>
      `;
    });

    html += '</table>';
    return html;
  }

  function buildClipboardText(rows) {
    return rows.map(row => [
      row.pickLogin || '',
      row.pack1Login || '',
      row.pack2Login || ''
    ].join('\t')).join('\n');
  }

  function getLoginPositions(rows) {
    const map = new Map();

    rows.forEach(row => {
      [
        { role: 'Pick', login: row.pickLogin },
        { role: 'Pack 1', login: row.pack1Login },
        { role: 'Pack 2', login: row.pack2Login }
      ].forEach(item => {
        const login = clean(item.login).toLowerCase();

        if (!login) return;

        map.set(login, {
          login: item.login,
          station: row.station,
          role: item.role
        });
      });
    });

    return map;
  }

  function compareSnapshots(oldRows, newRows) {
    const oldMap = getLoginPositions(oldRows || []);
    const newMap = getLoginPositions(newRows || []);
    const changes = [];

    newMap.forEach((newPos, loginKey) => {
      const oldPos = oldMap.get(loginKey);

      if (!oldPos) {
        changes.push({
          type: 'added',
          login: newPos.login,
          text: `${newPos.login} added to ${newPos.station} ${newPos.role}`
        });
        return;
      }

      if (oldPos.station !== newPos.station || oldPos.role !== newPos.role) {
        changes.push({
          type: 'moved',
          login: newPos.login,
          text: `${newPos.login} moved: ${oldPos.station} ${oldPos.role} → ${newPos.station} ${newPos.role}`
        });
      }
    });

    oldMap.forEach((oldPos, loginKey) => {
      if (!newMap.has(loginKey)) {
        changes.push({
          type: 'removed',
          login: oldPos.login,
          text: `${oldPos.login} removed from ${oldPos.station} ${oldPos.role}`
        });
      }
    });

    return changes;
  }

  function renderChangeHistory() {
    const box = document.getElementById('pj-scc-changes');
    if (!box) return;

    if (!changeHistory.length) {
      box.innerHTML = `<div class="pj-no-changes">No login changes detected yet.</div>`;
      return;
    }

    box.innerHTML = changeHistory
      .slice(0, 12)
      .map(item => `
        <div class="pj-change pj-change-${escapeHtml(item.type)}">
          <span class="pj-change-time">${escapeHtml(item.time)}</span>
          <span>${escapeHtml(item.text)}</span>
        </div>
      `)
      .join('');
  }

  function sendToExcelLiveWindow(wallName, rows, changes, time) {
    GM_setValue(SCC_LAYOUT_KEY, {
      time,
      wallName,
      rows
    });

    GM_setValue(SCC_CHANGE_KEY, {
      time,
      wallName,
      changes
    });
  }

  function checkForLoginChanges() {
    if (trackingPaused || checkingNow) return;
    checkingNow = true;

    const result = getAutoWallRows();

    if (!result.rows.length || !result.wallName || !hasAnyLogin(result.rows)) {
      checkingNow = false;
      return;
    }

    if (!lastSnapshot) {
      lastSnapshot = result.rows;
      GM_setValue(SCC_LAYOUT_KEY, {
        time: new Date().toLocaleTimeString(),
        wallName: result.wallName,
        rows: result.rows
      });
      renderChangeHistory();
      checkingNow = false;
      return;
    }

    const changes = compareSnapshots(lastSnapshot, result.rows);

    if (changes.length) {
      const time = new Date().toLocaleTimeString([], {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
      });

      changes.forEach(change => {
        changeHistory.unshift({
          ...change,
          time
        });
      });

      changeHistory = changeHistory.slice(0, 30);
      lastSnapshot = result.rows;

      sendToExcelLiveWindow(result.wallName, result.rows, changes, time);
      renderChangeHistory();
      showToast(`${changes.length} login change(s) detected`, 'success');
      refreshPreview();
    } else {
      GM_setValue(SCC_LAYOUT_KEY, {
        time: new Date().toLocaleTimeString(),
        wallName: result.wallName,
        rows: result.rows
      });
    }

    checkingNow = false;
  }

  function startChangeTracking() {
    if (trackingTimer) return;

    const initialiseWhenReady = () => {
      const result = getAutoWallRows();

      if (!result.wallName || !result.rows.length || !hasAnyLogin(result.rows)) {
        return false;
      }

      lastSnapshot = result.rows;

      GM_setValue(SCC_LAYOUT_KEY, {
        time: new Date().toLocaleTimeString(),
        wallName: result.wallName,
        rows: result.rows
      });

      trackingTimer = setInterval(checkForLoginChanges, TRACK_INTERVAL_MS);
      renderChangeHistory();
      return true;
    };

    if (!initialiseWhenReady()) {
      const readyTimer = setInterval(() => {
        if (initialiseWhenReady()) {
          clearInterval(readyTimer);
        }
      }, 1000);
    }
  }

  function resetChangeTracking() {
    const result = getAutoWallRows();

    if (!result.wallName || !result.rows.length || !hasAnyLogin(result.rows)) {
      const status = document.getElementById('pj-scc-status');
      if (status) {
        status.textContent = 'No login data available yet. Wait for SCC to finish loading.';
      }
      return;
    }

    lastSnapshot = result.rows;
    changeHistory = [];

    GM_setValue(SCC_LAYOUT_KEY, {
      time: new Date().toLocaleTimeString(),
      wallName: result.wallName,
      rows: result.rows
    });

    renderChangeHistory();

    const status = document.getElementById('pj-scc-status');
    if (status) {
      status.textContent = `${result.wallName || 'Wall'} tracking reset. Watching for login changes...`;
    }
  }

  function refreshPreview() {
    const result = getAutoWallRows();

    const previewBox = document.getElementById('pj-scc-preview');
    const status = document.getElementById('pj-scc-status');

    renderSccStats(result.rows, result.wallName);

    if (previewBox) {
      previewBox.innerHTML = buildPreviewHtml(result.rows, result.wallName);
    }

    if (status) {
      status.textContent = result.wallName && hasAnyLogin(result.rows)
        ? `${result.wallName} ready. Logins of Picker | Packer 1 | Packer 2.`
        : 'Waiting for SCC station login data...';
    }

    console.table(result.rows);
  }

  async function copyPreview() {
    const result = getAutoWallRows();

    if (!result.rows.length || !result.wallName || !hasAnyLogin(result.rows)) {
      alert('No floor plan login data found yet. Wait for SCC to finish loading, then click Recheck.');
      return;
    }

    const clipboardHtml = buildClipboardHtml(result.rows);
    const clipboardText = buildClipboardText(result.rows);

    try {
      if (navigator.clipboard && window.ClipboardItem) {
        const clipboardData = new ClipboardItem({
          'text/html': new Blob([clipboardHtml], { type: 'text/html' }),
          'text/plain': new Blob([clipboardText], { type: 'text/plain' })
        });

        await navigator.clipboard.write([clipboardData]);
      } else {
        GM_setClipboard(clipboardText);
      }

      showToast(`${result.wallName} copied with FCLM links`, 'success');
    } catch (err) {
      console.error('Clipboard copy failed:', err);
      GM_setClipboard(clipboardText);
      showToast('Copied as plain text', 'success');
    }
  }

  function refreshPage() {
    if (location.protocol === 'file:') {
      location.reload();
      return;
    }

    const url = new URL(location.href);
    url.searchParams.set('_tmRefresh', Date.now());
    location.href = url.toString();
  }

  function setSccPanelMinimised(isMinimised) {
    localStorage.setItem(STORAGE_MINIMISED, isMinimised ? '1' : '0');

    const panel = document.getElementById('pj-scc-panel');
    const btn = document.getElementById('pj-scc-min-btn');
    const title = document.getElementById('pj-scc-title');

    if (!panel || !btn || !title) return;

    if (isMinimised) {
      panel.classList.add('pj-min');
      title.textContent = 'SCC';
      btn.textContent = '+';
    } else {
      panel.classList.remove('pj-min');
      title.textContent = 'SCC Auto Login Detector';
      btn.textContent = '−';
      refreshPreview();
    }
  }

  function addSccPanel() {
    if (document.getElementById('pj-scc-panel')) return;

    injectSccStyles();

    const panel = document.createElement('div');
    panel.id = 'pj-scc-panel';

    panel.innerHTML = `
      <div class="pj-scc-header">
        <div>
          <div class="pj-scc-title" id="pj-scc-title">SCC Auto Login Detector</div>
          <div class="pj-scc-subtitle">Logins of Picker • Packer 1 • Packer 2</div>
        </div>
        <button class="pj-scc-icon-btn" id="pj-scc-min-btn" title="Minimise">−</button>
      </div>

      <div class="pj-scc-body">
        <div class="pj-scc-actions">
          <button class="pj-scc-btn pj-copy-btn" id="pj-copy-btn">Copy</button>
          <button class="pj-scc-btn pj-recheck-btn" id="pj-recheck-btn">Recheck</button>
          <button class="pj-scc-btn pj-recheck-btn" id="pj-reset-track-btn">Reset Track</button>
          <button class="pj-scc-btn pj-pause-btn" id="pj-pause-track-btn">Pause</button>
          <button class="pj-scc-btn pj-copy-changes-btn" id="pj-copy-changes-btn">Copy Changes</button>
          <button class="pj-scc-btn pj-refresh-btn" id="pj-refresh-btn">Refresh Page</button>
        </div>

        <div class="pj-scc-stats" id="pj-scc-stats"></div>

        <div class="pj-scc-card">
          <p class="pj-scc-note">Auto-detects the visible wall and watches Pick / Pack login changes.</p>
          <p class="pj-scc-meta">Creator: ${escapeHtml(CREATOR)} | v${escapeHtml(SCRIPT_VERSION)}</p>
        </div>

        <div class="pj-scc-status" id="pj-scc-status">Waiting for visible floor plan station cards...</div>

        <div class="pj-scc-changes">
          <div class="pj-changes-title">Login Changes</div>
          <div id="pj-scc-changes">
            <div class="pj-no-changes">No login changes detected yet.</div>
          </div>
        </div>

        <div id="pj-scc-preview"></div>
      </div>
    `;

    document.body.appendChild(panel);

    document.getElementById('pj-copy-btn').addEventListener('click', copyPreview);
    document.getElementById('pj-recheck-btn').addEventListener('click', refreshPreview);
    document.getElementById('pj-reset-track-btn').addEventListener('click', resetChangeTracking);
    document.getElementById('pj-pause-track-btn').addEventListener('click', toggleTrackingPause);
    document.getElementById('pj-copy-changes-btn').addEventListener('click', copyLatestChanges);
    document.getElementById('pj-refresh-btn').addEventListener('click', refreshPage);
    document.getElementById('pj-scc-min-btn').addEventListener('click', () => {
      const isMin = localStorage.getItem(STORAGE_MINIMISED) === '1';
      setSccPanelMinimised(!isMin);
    });

    const savedMinimised = localStorage.getItem(STORAGE_MINIMISED) === '1';
    setSccPanelMinimised(savedMinimised);

    if (!savedMinimised) {
      refreshPreview();
    }

    makePanelDraggable('pj-scc-panel', '.pj-scc-header', STORAGE_SCC_LEFT, STORAGE_SCC_TOP);
    rememberPanelSize('pj-scc-panel', STORAGE_SCC_WIDTH, STORAGE_SCC_HEIGHT, () => localStorage.getItem(STORAGE_MINIMISED) === '1');

    startChangeTracking();
  }

  function waitForFloorPlan() {
    const tables = getStationTables();

    if (tables.length) {
      addSccPanel();
    }
  }

  function runExcelLiveWindow() {
    if (!isExcelPage()) return;
    if (document.getElementById('pj-scc-live-window')) return;

    injectExcelStyles();

    const panel = document.createElement('div');
    panel.id = 'pj-scc-live-window';

    panel.innerHTML = `
      <div class="pj-live-header">
        <div>
          <div class="pj-live-title">SCC Live Changes</div>
          <div class="pj-live-subtitle">Login Change monitor | v${escapeHtml(SCRIPT_VERSION)}</div>
        </div>
        <button class="pj-live-min" id="pj-live-min-btn">−</button>
      </div>

      <div class="pj-live-body">
        <div class="pj-live-status" id="pj-live-status">
          Waiting for SCC login changes...
        </div>

        <div class="pj-live-actions">
          <button class="pj-live-btn pj-copy-latest" id="pj-live-copy-layout">Copy Latest Layout</button>
          <button class="pj-live-btn pj-copy-latest" id="pj-live-copy-changes">Copy Changes</button>
          <button class="pj-live-btn pj-clear-log" id="pj-live-clear-log">Clear</button>
        </div>

        <div class="pj-live-log" id="pj-live-log">
          <div class="pj-live-empty">No changes yet.</div>
        </div>

        <div class="pj-live-creator">Creator: ${escapeHtml(CREATOR)}</div>
      </div>
    `;

    document.body.appendChild(panel);

    const logBox = document.getElementById('pj-live-log');
    const statusBox = document.getElementById('pj-live-status');
    const minBtn = document.getElementById('pj-live-min-btn');
    const clearBtn = document.getElementById('pj-live-clear-log');
    const copyBtn = document.getElementById('pj-live-copy-layout');
    const copyChangesBtn = document.getElementById('pj-live-copy-changes');

    let logs = GM_getValue(SCC_EXCEL_LOG_KEY, []);

    function renderLogs() {
      if (!logs.length) {
        logBox.innerHTML = `<div class="pj-live-empty">No changes yet.</div>`;
        return;
      }

      logBox.innerHTML = logs.slice(0, 60).map(item => `
        <div class="pj-live-item ${escapeHtml(item.type || '')}">
          <div class="pj-live-time">${escapeHtml(item.time || '')} • ${escapeHtml(item.wallName || '')}</div>
          <div>${escapeHtml(item.text || '')}</div>
        </div>
      `).join('');
    }

    async function copyLatestLayout() {
      const data = GM_getValue(SCC_LAYOUT_KEY);

      if (!data || !data.rows || !data.rows.length) {
        alert('No SCC layout received yet.');
        return;
      }

      const text = buildClipboardText(data.rows);
      const html = buildClipboardHtml(data.rows);

      try {
        if (navigator.clipboard && window.ClipboardItem) {
          const clipboardData = new ClipboardItem({
            'text/html': new Blob([html], { type: 'text/html' }),
            'text/plain': new Blob([text], { type: 'text/plain' })
          });

          await navigator.clipboard.write([clipboardData]);
        } else {
          GM_setClipboard(text);
        }

        statusBox.textContent = `${data.wallName} latest layout copied with FCLM links. Paste into Excel.`;
      } catch (err) {
        console.error('Excel helper copy failed:', err);
        GM_setClipboard(text);
        statusBox.textContent = `${data.wallName} latest layout copied as plain text.`;
      }
    }

    function copyExcelChanges() {
      if (!logs.length) {
        statusBox.textContent = 'No changes to copy yet.';
        return;
      }

      const text = logs.slice(0, 100).map(item => [
        item.time || '',
        item.wallName || '',
        item.type || '',
        item.text || ''
      ].join('\t')).join('\n');

      GM_setClipboard(text);
      statusBox.textContent = `Copied ${logs.length} change log item(s).`;
    }

    GM_addValueChangeListener(SCC_CHANGE_KEY, (name, oldValue, newValue) => {
      if (!newValue || !newValue.changes) return;

      statusBox.textContent = `${newValue.wallName} changed at ${newValue.time}`;

      newValue.changes.forEach(change => {
        logs.unshift({
          time: newValue.time,
          wallName: newValue.wallName,
          type: change.type,
          text: change.text
        });
      });

      logs = logs.slice(0, 100);
      GM_setValue(SCC_EXCEL_LOG_KEY, logs);
      renderLogs();
    });

    GM_addValueChangeListener(SCC_LAYOUT_KEY, (name, oldValue, newValue) => {
      if (!newValue) return;

      if (!logs.length) {
        statusBox.textContent = `${newValue.wallName || 'Wall'} layout received at ${newValue.time}. Waiting for changes...`;
      }
    });

    minBtn.addEventListener('click', () => {
      const isMin = panel.classList.toggle('min');
      localStorage.setItem(STORAGE_EXCEL_MINIMISED, isMin ? '1' : '0');
      minBtn.textContent = isMin ? '+' : '−';
      const title = panel.querySelector('.pj-live-title');
      if (title) title.textContent = isMin ? 'SCC' : 'SCC Live Changes';
    });

    clearBtn.addEventListener('click', () => {
      logs = [];
      GM_setValue(SCC_EXCEL_LOG_KEY, logs);
      renderLogs();
      statusBox.textContent = 'Log cleared. Waiting for SCC login changes...';
    });

    copyBtn.addEventListener('click', copyLatestLayout);
    copyChangesBtn.addEventListener('click', copyExcelChanges);

    const savedMin = localStorage.getItem(STORAGE_EXCEL_MINIMISED) === '1';
    if (savedMin) {
      panel.classList.add('min');
      minBtn.textContent = '+';
      const title = panel.querySelector('.pj-live-title');
      if (title) title.textContent = 'SCC';
    }

    const latest = GM_getValue(SCC_LAYOUT_KEY);
    if (latest) {
      statusBox.textContent = `${latest.wallName || 'Wall'} layout received at ${latest.time}. Waiting for changes...`;
    }

    renderLogs();

    makePanelDraggable('pj-scc-live-window', '.pj-live-header', STORAGE_EXCEL_LEFT, STORAGE_EXCEL_TOP);
    rememberPanelSize('pj-scc-live-window', STORAGE_EXCEL_WIDTH, STORAGE_EXCEL_HEIGHT, () => panel.classList.contains('min'));
  }

  if (isExcelPage()) {
    runExcelLiveWindow();
    return;
  }

  if (!isSccPage()) {
    return;
  }

  setInterval(waitForFloorPlan, 1000);

})();