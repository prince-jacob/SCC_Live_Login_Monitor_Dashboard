// ==UserScript==
// @name         SCC Live Login Monitor Dashboard - NCL1
// @namespace    prince-scc
// @version      1.9.5
// @description  Auto-detect wall, copy Pick/Pack logins with FCLM links, track SCC login changes, and show live changes on OneDrive Excel and P2R Tracker pages, with optional P2R auto-apply, SCC trained-role capture, and longer hover role popup and strict P2R Pick/P2R Pack role filters on P2R Tracker. Upgraded dashboard UI.
// @author       Prince Jacob ( Wprijaco )
// @match        https://staffingcommandcenter-eu.aka.amazon.com/NCL1/*
// @match        https://onedrive.live.com/edit*
// @match        https://excel.officeapps.live.com/*
// @match        https://p2r-tracker.web.app/*
// @match        file:///*
// @updateURL    https://raw.githubusercontent.com/prince-jacob/SCC_Live_Login_Monitor_Dashboard/main/SCCLiveLoginMonitorDashboard.user.js
// @downloadURL  https://raw.githubusercontent.com/prince-jacob/SCC_Live_Login_Monitor_Dashboard/main/SCCLiveLoginMonitorDashboard.user.js
// @grant        GM_setClipboard
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addValueChangeListener
// @noframes
// @grant        GM_xmlhttpRequest
// @grant        GM_info
// @connect      raw.githubusercontent.com
// ==/UserScript==

(function () {
  'use strict';

  // Prevent duplicate execution inside embedded frames.
  if (window.top !== window.self) return;

  const CREATOR = 'Prince Jacob ( Wprijaco )';
  const SCRIPT_VERSION = '1.9.5';
  const OFFICIAL_MARKER = 'OFFICIAL_SCC_LIVE_LOGIN_MONITOR_PRINCE_JACOB_V1_9_5';

  const SCC_CHANGE_KEY = 'pj_scc_live_login_changes';
  const SCC_LAYOUT_KEY = 'pj_scc_latest_layout';
  const SCC_EXCEL_LOG_KEY = 'pj_scc_excel_log';
  const SCC_TRAINING_KEY = 'pj_scc_login_trained_roles_v192';

  const P2R_AUTO_APPLY_KEY = 'pj_scc_p2r_auto_apply_v190';
  const P2R_ROLE_HIGHLIGHT_KEY = 'pj_scc_p2r_role_highlight_v192';
  const P2R_ROLE_FILTER_KEY = 'pj_scc_p2r_role_filter_v194';
  const P2R_DB_BASE = 'https://p2r-tracker-default-rtdb.europe-west1.firebasedatabase.app';
  let p2rApplying = false;
  let p2rLastAppliedSignature = '';

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
  const TRAINING_REFRESH_INTERVAL_MS = 10 * 60 * 1000; // 10 minutes

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
  let sccTrainingByLogin = GM_getValue(SCC_TRAINING_KEY, {}) || {};
  let sccTrainingUpdatedAt = 0;
  let sccTrainingRefreshPromise = null;
  let sccTrainingTimer = null;
  let p2rRoleHighlightTimer = null;
  let p2rRoleTooltipHideTimer = null;

  function isSccPage() {
    return location.href.startsWith('https://staffingcommandcenter-eu.aka.amazon.com/NCL1/') ||
           location.protocol === 'file:';
  }

  function isExcelPage() {
    return location.href.includes('onedrive.live.com/edit') ||
           location.href.includes('excel.officeapps.live.com') ||
           location.href.includes('p2r-tracker.web.app');
  }

  function getHelperPageName() {
    if (location.href.includes('p2r-tracker.web.app')) return 'P2R Tracker';
    if (location.href.includes('onedrive.live.com') || location.href.includes('excel.officeapps.live.com')) return 'Excel';
    return 'Helper Page';
  }

  function isP2RTrackerPage() {
    return location.href.includes('p2r-tracker.web.app');
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

    // Best SCC source:
    // data-testid="floor-plan-employee-hadanieg"
    const employeeNode = cell.querySelector('[data-testid^="floor-plan-employee-"]');

    if (employeeNode) {
      const testId = employeeNode.getAttribute('data-testid') || '';
      login = testId.replace(/^floor-plan-employee-/i, '').trim();
    }

    // If no data-testid, use only the FIRST token.
    // Example:
    // "hadanieg 312 wsPickToRebin3_301_01" -> "hadanieg"
    if (!login) {
      const firstTokenMatch = text.match(/^([a-z0-9.-]+)/i);
      login = firstTokenMatch ? firstTokenMatch[1] : '';
    }

    // Final cleanup
    login = clean(login)
      .replace(/[^a-z0-9.-]/gi, '')
      .trim();

    // Never let workstation strings become logins
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

  function loginKey(login) {
    return clean(login).toLowerCase();
  }

  function trainedRoleLabel(role) {
    const value = String(role || '').trim();
    const upper = value.toUpperCase();

    const labels = {
      P2R_PICK: 'P2R Pick',
      P2R_PACK: 'P2R Pack',
      AFE_PACK: 'AFE Pack',
      AFE_REBIN: 'AFE Rebin',
      AFE_PACK_GW: 'Gift Wrap',
      PICK_AR: 'AR Pick',
      STOW: 'Stow',
      DECANT: 'Decant',
      ICQA_SIMPLE_RECORD_COUNT: 'ICQA',
      ICQA_SIMPLE_BIN_COUNT: 'ICQA',
      ICQA_CYCLE_COUNT: 'ICQA',
      PP_SINGLE_SMALL: 'Singles',
      PPSINGLESMALL: 'Singles',
      PPSINGLEMEDIUM: 'Singles',
      SHIP_DOCK: 'Ship Dock'
    };

    return labels[upper] || value.replace(/_/g, ' ').replace(/\b\w/g, char => char.toUpperCase());
  }

  function normalizeTrainedRoles(roles) {
    const list = Array.isArray(roles) ? roles : [];
    const unique = [];
    const seen = new Set();

    list.forEach(role => {
      const value = String(role || '').trim();
      const key = value.toUpperCase();
      if (!value || seen.has(key)) return;
      seen.add(key);
      unique.push(value);
    });

    const order = [
      'P2R_PICK', 'P2R_PACK', 'AFE_REBIN', 'AFE_PACK', 'AFE_PACK_GW',
      'PICK_AR', 'STOW', 'DECANT', 'PPSINGLESMALL', 'PPSINGLEMEDIUM',
      'ICQA_SIMPLE_RECORD_COUNT', 'ICQA_SIMPLE_BIN_COUNT', 'ICQA_CYCLE_COUNT',
      'SHIP_DOCK'
    ];

    unique.sort((a, b) => {
      const ia = order.indexOf(String(a).toUpperCase());
      const ib = order.indexOf(String(b).toUpperCase());
      return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
    });

    return unique;
  }

  function trainingInfoForLogin(login) {
    const key = loginKey(login);
    if (!key) return null;
    return sccTrainingByLogin[key] || null;
  }

  function canDoP2RRole(roles, boardRole) {
    const set = new Set((Array.isArray(roles) ? roles : []).map(role => String(role).toUpperCase()));
    const role = String(boardRole || '').toLowerCase();

    if (role === 'pick') return set.has('P2R_PICK');
    if (role === 'pack1' || role === 'pack2') return set.has('P2R_PACK');
    return false;
  }

  function shortTrainingLabel(roles, boardRole = '') {
    const set = new Set((Array.isArray(roles) ? roles : []).map(role => String(role).toUpperCase()));
    const parts = [];

    if (set.has('P2R_PICK')) parts.push('P2R Pick');
    if (set.has('P2R_PACK')) parts.push('P2R Pack');
    if (set.has('AFE_PACK')) parts.push('AFE Pack');
    if (set.has('AFE_REBIN')) parts.push('AFE Rebin');
    if (set.has('AFE_PACK_GW')) parts.push('Gift');
    if (set.has('PICK_AR')) parts.push('AR Pick');
    if (set.has('STOW')) parts.push('Stow');

    if (!parts.length) {
      return normalizeTrainedRoles(roles).slice(0, 2).map(trainedRoleLabel).join(' / ');
    }

    return parts.slice(0, 3).join(' / ');
  }

  function attachTrainingToRows(rows) {
    return (rows || []).map(row => {
      const pickTraining = trainingInfoForLogin(row.pickLogin);
      const pack1Training = trainingInfoForLogin(row.pack1Login);
      const pack2Training = trainingInfoForLogin(row.pack2Login);

      return {
        ...row,
        pickTrainingRoles: pickTraining?.trained || [],
        pack1TrainingRoles: pack1Training?.trained || [],
        pack2TrainingRoles: pack2Training?.trained || []
      };
    });
  }

  function buildSccLayoutPayload(wallName, rows, time = new Date().toLocaleTimeString()) {
    return {
      time,
      wallName,
      rows: attachTrainingToRows(rows || []),
      trainingByLogin: sccTrainingByLogin || {},
      trainingUpdatedAt: sccTrainingUpdatedAt || 0
    };
  }

  async function refreshSccTrainingCache(force = false, statusBox = null) {
    if (!location.hostname.includes('staffingcommandcenter-eu.aka.amazon.com')) {
      return sccTrainingByLogin;
    }

    const now = Date.now();
    if (!force && sccTrainingUpdatedAt && now - sccTrainingUpdatedAt < TRAINING_REFRESH_INTERVAL_MS) {
      return sccTrainingByLogin;
    }

    if (sccTrainingRefreshPromise) {
      return sccTrainingRefreshPromise;
    }

    if (statusBox) statusBox.textContent = 'Refreshing trained roles from SCC...';

    sccTrainingRefreshPromise = (async () => {
      const [profileResponse, trainingResponse] = await Promise.all([
        fetch('/getAssociateProfileDetails/NCL1', { credentials: 'include', cache: 'no-store' }),
        fetch('/getAssociateTrainedRoles/NCL1', { credentials: 'include', cache: 'no-store' })
      ]);

      if (!profileResponse.ok) throw new Error(`Profile details ${profileResponse.status}`);
      if (!trainingResponse.ok) throw new Error(`Trained roles ${trainingResponse.status}`);

      const profiles = await profileResponse.json();
      const training = await trainingResponse.json();
      const mapped = {};

      Object.entries(profiles || {}).forEach(([employeeId, profile]) => {
        const login = clean(profile?.employeeLogin || profile?.login || '');
        if (!login) return;

        const trained = normalizeTrainedRoles(training?.[employeeId]?.trained || []);
        mapped[login.toLowerCase()] = {
          login,
          trained,
          labels: trained.map(trainedRoleLabel),
          canP2RPick: canDoP2RRole(trained, 'pick'),
          canP2RPack: canDoP2RRole(trained, 'pack1')
        };
      });

      sccTrainingByLogin = mapped;
      sccTrainingUpdatedAt = Date.now();
      GM_setValue(SCC_TRAINING_KEY, mapped);

      const current = getAutoWallRows();
      if (current.wallName && current.rows.length) {
        GM_setValue(SCC_LAYOUT_KEY, buildSccLayoutPayload(current.wallName, current.rows));
      }

      if (statusBox) {
        statusBox.textContent = `Trained roles refreshed: ${Object.keys(mapped).length} login(s) cached.`;
      }

      return mapped;
    })().catch(error => {
      console.error('SCC trained role refresh failed:', error);
      if (statusBox) statusBox.textContent = `Trained role refresh failed: ${error.message || error}`;
      return sccTrainingByLogin;
    }).finally(() => {
      sccTrainingRefreshPromise = null;
    });

    return sccTrainingRefreshPromise;
  }

  function startSccTrainingRefreshLoop() {
    if (!location.hostname.includes('staffingcommandcenter-eu.aka.amazon.com')) return;
    if (sccTrainingTimer) return;

    refreshSccTrainingCache(false).catch(console.error);
    sccTrainingTimer = setInterval(() => {
      refreshSccTrainingCache(false).catch(console.error);
    }, TRAINING_REFRESH_INTERVAL_MS);
  }

  function getStationTables() {
    const classMatched = [...document.querySelectorAll(TABLE_SELECTOR)];

    if (classMatched.length) {
      return classMatched;
    }

    // Fallback for users receiving a different generated CSS-module class.
    return [...document.querySelectorAll('table')].filter(table =>
      /\b\d{3}\s*\|\s*Priority\s*\d+/i.test(clean(table.innerText))
    );
  }

  function extractEmployeeLogin(employeeCard) {
    if (!employeeCard) return '';

    const testId = employeeCard.getAttribute('data-testid') || '';
    let login = testId.replace(/^floor-plan-employee-/i, '').trim();

    if (!login || login.toLowerCase() === 'login') {
      const loginNode = employeeCard.querySelector('[data-testid="floor-plan-employee-login"]');
      login = clean(loginNode?.innerText || '');
    }

    if (!login) {
      const firstToken = clean(employeeCard.innerText).match(/^([a-z0-9.-]+)/i);
      login = firstToken ? firstToken[1] : '';
    }

    login = clean(login).replace(/[^a-z0-9.-]/gi, '');
    return isBadLoginCandidate(login) ? '' : login;
  }

  function extractEmployeeWorkstation(employeeCard) {
    const text = clean(employeeCard?.innerText || '');

    const packMatch = text.match(/\b(wsPickToRebin\d+_\d+_0[12])\b/i);
    if (packMatch) return packMatch[1];

    const pickMatch = text.match(/\b(ws-k-[a-z]-\d+-\d+)\b/i);
    return pickMatch ? pickMatch[1] : '';
  }

  function nearestStationByScreenPosition(employeeCard, headerCells, stationHeaders) {
    if (!employeeCard || !headerCells.length || !stationHeaders.length) return 0;

    const employeeRect = employeeCard.getBoundingClientRect();
    const employeeCenter = employeeRect.left + (employeeRect.width / 2);

    let nearestStation = 0;
    let nearestDistance = Infinity;

    headerCells.forEach((cell, index) => {
      const stationInfo = stationHeaders[index];
      if (!stationInfo) return;

      const rect = cell.getBoundingClientRect();
      const center = rect.left + (rect.width / 2);
      const distance = Math.abs(employeeCenter - center);

      if (distance < nearestDistance) {
        nearestDistance = distance;
        nearestStation = stationInfo.station;
      }
    });

    return nearestStation;
  }

  function captureSccFloorPlan() {
    const tables = getStationTables();
    const stationMap = new Map();

    tables.forEach(table => {
      const headerCells = getDirectCells(table.tHead?.rows?.[0]);
      const stationHeaders = headerCells.map(cell => parseStationHeader(cell.innerText));
      const validHeaders = stationHeaders.filter(Boolean);

      if (!validHeaders.length) return;

      validHeaders.forEach(header => {
        if (!stationMap.has(header.station)) {
          stationMap.set(header.station, {
            station: header.station,
            priority: header.priority,
            pickLogin: '',
            pickHref: '',
            pack1Login: '',
            pack1Href: '',
            pack2Login: '',
            pack2Href: ''
          });
        }
      });

      // Read employee cards directly. This works on both /approved/ and /plan/
      // pages, including pages where Pick UPH / Pack UPH rows are not rendered.
      const employeeCards = [...table.querySelectorAll('[data-testid^="floor-plan-employee-"]')]
        .filter(node => (node.getAttribute('data-testid') || '').toLowerCase() !== 'floor-plan-employee-login');

      employeeCards.forEach(card => {
        const login = extractEmployeeLogin(card);
        if (!login) return;

        const workstation = extractEmployeeWorkstation(card);
        let station = 0;
        let role = '';

        const packMatch = workstation.match(/wsPickToRebin\d+_(\d+)_0([12])/i);

        if (packMatch) {
          station = Number(packMatch[1]);
          role = packMatch[2] === '1' ? 'pack1' : 'pack2';
        } else {
          // Picker workstations do not include the wall station number.
          // Use the card's actual screen column and match it to the nearest
          // station header, which is stable on both SCC page variants.
          station = nearestStationByScreenPosition(card, headerCells, stationHeaders);
          role = 'pick';
        }

        if (!station || !stationMap.has(station)) return;

        const row = stationMap.get(station);
        const href = makeEmployeeHref(login);

        if (role === 'pick') {
          row.pickLogin = login;
          row.pickHref = href;
        } else if (role === 'pack1') {
          row.pack1Login = login;
          row.pack1Href = href;
        } else if (role === 'pack2') {
          row.pack2Login = login;
          row.pack2Href = href;
        }
      });
    });

    return [...stationMap.values()].sort((a, b) => a.station - b.station);
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

    return { wallName, rows: attachTrainingToRows(rows) };
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
    GM_setValue(SCC_LAYOUT_KEY, buildSccLayoutPayload(wallName, rows, time));

    GM_setValue(SCC_CHANGE_KEY, {
      time,
      wallName,
      changes
    });
  }

  function hasAnyLogin(rows) {
    return rows.some(row => row.pickLogin || row.pack1Login || row.pack2Login);
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
      GM_setValue(SCC_LAYOUT_KEY, buildSccLayoutPayload(result.wallName, result.rows));
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
      GM_setValue(SCC_LAYOUT_KEY, buildSccLayoutPayload(result.wallName, result.rows));
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

      GM_setValue(SCC_LAYOUT_KEY, buildSccLayoutPayload(result.wallName, result.rows));

      trackingTimer = setInterval(checkForLoginChanges, TRACK_INTERVAL_MS);
      renderChangeHistory();
      return true;
    };

    if (!initialiseWhenReady()) {
      const readyTimer = setInterval(() => {
        if (initialiseWhenReady()) clearInterval(readyTimer);
      }, 1000);
    }
  }

  function resetChangeTracking() {
    const result = getAutoWallRows();

    lastSnapshot = result.rows;
    changeHistory = [];

    GM_setValue(SCC_LAYOUT_KEY, buildSccLayoutPayload(result.wallName, result.rows));

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
      status.textContent = result.wallName
        ? `${result.wallName} ready. Logins of Picker | Packer 1 | Packer 2.`
        : 'Waiting for visible floor plan station cards...';
    }

    console.table(result.rows);
  }

  async function copyPreview() {
    const result = getAutoWallRows();

    if (!result.rows.length || !result.wallName) {
      alert('No floor plan station data found.');
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
          <button class="pj-scc-btn pj-recheck-btn" id="pj-refresh-roles-btn">Refresh Roles</button>
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
    document.getElementById('pj-refresh-roles-btn').addEventListener('click', () => refreshSccTrainingCache(true, document.getElementById('pj-scc-status')));
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
    startSccTrainingRefreshLoop();
  }

  function waitForFloorPlan() {
    if (getStationTables().length) {
      addSccPanel();
    }
  }


  function p2rRoleFromSccRole(role) {
    const value = String(role || '').toLowerCase().replace(/\s+/g, '');
    if (value === 'pick') return 'pick';
    if (value === 'pack1') return 'pack1';
    if (value === 'pack2') return 'pack2';
    return '';
  }

  function p2rSlotKey(floor, station, role) {
    return `${floor}|${Number(station)}|${role}`;
  }

  function p2rLayoutSignature(data) {
    if (!data || !data.rows) return '';
    return [
      data.wallName || '',
      ...data.rows.map(row => [
        row.station || '',
        row.pickLogin || '',
        row.pack1Login || '',
        row.pack2Login || ''
      ].join(':'))
    ].join('|').toLowerCase();
  }

  function p2rDefaultPerson(login) {
    return {
      login,
      name: '',
      manager: '',
      pick: false,
      pack: false,
      stow: false,
      decant: false,
      arsaw: false,
      singles: false,
      vendors: false,
      icqa: false,
      ws: false,
      ship: false,
      afeSort: false,
      restriction: 'Any',
      lastBoardChangeAt: 0,
      leave: 'full',
      floor: '',
      station: '',
      role: 'offplan',
      inDirectory: false
    };
  }

  function p2rNormalizePeople(data) {
    const source = Array.isArray(data) ? data : data ? Object.values(data) : [];

    return source
      .filter(item => item && item.login)
      .map(item => ({
        ...p2rDefaultPerson(String(item.login || '').trim()),
        ...item,
        login: String(item.login || '').trim(),
        station: item.station === '' || item.station == null ? '' : Number(item.station)
      }));
  }

  async function p2rFetchJson(path, options = {}) {
    const url = `${P2R_DB_BASE}/${path}.json`;
    const response = await fetch(url, {
      ...options,
      headers: {
        'Content-Type': 'application/json',
        ...(options.headers || {})
      }
    });

    if (!response.ok) {
      throw new Error(`P2R Firebase ${response.status}: ${await response.text()}`);
    }

    if (options.method === 'PUT' || options.method === 'PATCH') {
      try {
        return await response.json();
      } catch {
        return null;
      }
    }

    return response.json();
  }

  function p2rBuildTargetsFromLayout(data) {
    const targets = [];

    if (!data || !data.wallName || !Array.isArray(data.rows)) {
      return targets;
    }

    data.rows.forEach(row => {
      [
        { login: row.pickLogin, role: 'pick' },
        { login: row.pack1Login, role: 'pack1' },
        { login: row.pack2Login, role: 'pack2' }
      ].forEach(item => {
        const login = clean(item.login);
        if (!login || isBadLoginCandidate(login)) return;

        targets.push({
          login,
          loginKey: login.toLowerCase(),
          floor: data.wallName,
          station: Number(row.station),
          role: item.role
        });
      });
    });

    return targets;
  }

  async function p2rApplySccLayout(data, statusBox, sourceLabel = 'manual') {
    if (!isP2RTrackerPage()) return false;

    if (p2rApplying) {
      if (statusBox) statusBox.textContent = 'P2R auto apply already running...';
      return false;
    }

    if (!data || !data.wallName || !Array.isArray(data.rows) || !data.rows.length) {
      if (statusBox) statusBox.textContent = 'No SCC layout ready to apply to P2R Tracker.';
      return false;
    }

    const signature = p2rLayoutSignature(data);

    if (sourceLabel === 'auto' && signature && signature === p2rLastAppliedSignature) {
      return false;
    }

    const targets = p2rBuildTargetsFromLayout(data);

    if (!targets.length) {
      if (statusBox) statusBox.textContent = `${data.wallName || 'Wall'} has no logins to apply yet.`;
      return false;
    }

    p2rApplying = true;

    try {
      if (statusBox) statusBox.textContent = `Applying ${data.wallName} SCC layout to P2R Tracker...`;

      const remotePeople = await p2rFetchJson('people');
      const people = p2rNormalizePeople(remotePeople);

      const floor = data.wallName;
      const trackedRoles = new Set(['pick', 'pack1', 'pack2']);
      const targetByLogin = new Map();
      const targetSlotToLogin = new Map();

      targets.forEach(target => {
        targetByLogin.set(target.loginKey, target);
        targetSlotToLogin.set(p2rSlotKey(target.floor, target.station, target.role), target.loginKey);
      });

      // Clear old pick/pack slots on the same floor that are no longer matching the SCC layout.
      people.forEach(person => {
        const role = String(person.role || '');
        const loginKey = String(person.login || '').toLowerCase();
        const slotKey = p2rSlotKey(person.floor, person.station, role);

        if (person.floor === floor && trackedRoles.has(role)) {
          const wantedLoginForSlot = targetSlotToLogin.get(slotKey);

          if (wantedLoginForSlot !== loginKey) {
            person.floor = '';
            person.station = '';
            person.role = 'offplan';
            person.lastBoardChangeAt = 0;
          }
        }
      });

      // Apply every SCC slot to P2R. Existing occupants were cleared above.
      targets.forEach(target => {
        let person = people.find(item =>
          String(item.login || '').toLowerCase() === target.loginKey
        );

        if (!person) {
          person = p2rDefaultPerson(target.login);
          people.push(person);
        }

        person.floor = target.floor;
        person.station = target.station;
        person.role = target.role;

        if (!person.leave) person.leave = 'full';
        if (!person.restriction) person.restriction = 'Any';
        person.lastBoardChangeAt = Date.now();
      });

      localStorage.setItem('hc-planner-people', JSON.stringify(people));
      await p2rFetchJson('people', {
        method: 'PUT',
        body: JSON.stringify(people)
      });

      p2rLastAppliedSignature = signature;

      if (statusBox) {
        statusBox.textContent = `${data.wallName} applied to P2R Tracker: ${targets.length} login slot(s) updated.`;
      }

      return true;
    } catch (error) {
      console.error('P2R auto apply failed:', error);

      if (statusBox) {
        statusBox.textContent = `P2R auto apply failed: ${error.message || error}`;
      }

      return false;
    } finally {
      p2rApplying = false;
    }
  }


  function p2rRoleHighlightEnabled() {
    return GM_getValue(P2R_ROLE_HIGHLIGHT_KEY, false) === true;
  }

  function p2rGetTrainingMapFromStorage() {
    const latest = GM_getValue(SCC_LAYOUT_KEY, null);
    const fromLayout = latest?.trainingByLogin;
    const fromTrainingKey = GM_getValue(SCC_TRAINING_KEY, {});
    return (fromLayout && Object.keys(fromLayout).length) ? fromLayout : (fromTrainingKey || {});
  }

  function injectP2RRoleHighlightStyles() {
    if (!isP2RTrackerPage() || document.getElementById('pj-p2r-role-highlight-style')) return;

    const style = document.createElement('style');
    style.id = 'pj-p2r-role-highlight-style';
    style.textContent = `
      .aa-tile.pj-scc-role-trained,
      .aa-tile.pj-scc-role-other {
        position: relative !important;
        transition: box-shadow .16s ease, outline-color .16s ease, transform .16s ease !important;
      }

      /* Main colour code:
         Green = P2R Pick trained
         Blue = P2R Pack trained
         Purple = both P2R Pick + P2R Pack trained
         Orange = trained roles exist, but not P2R Pick/Pack
      */
      .aa-tile.pj-scc-role-pick {
        outline: 2px solid rgba(34,197,94,.98) !important;
        box-shadow: 0 0 0 2px rgba(34,197,94,.14) !important;
      }

      .aa-tile.pj-scc-role-pack {
        outline: 2px solid rgba(37,99,235,.98) !important;
        box-shadow: 0 0 0 2px rgba(37,99,235,.16) !important;
      }

      .aa-tile.pj-scc-role-both {
        outline: 2px solid rgba(124,58,237,.98) !important;
        box-shadow: 0 0 0 2px rgba(124,58,237,.17) !important;
      }

      .aa-tile.pj-scc-role-other-training {
        outline: 2px solid rgba(245,158,11,.95) !important;
        box-shadow: 0 0 0 2px rgba(245,158,11,.16) !important;
      }

      .aa-tile.pj-scc-role-slot-mismatch {
        box-shadow: inset 0 0 0 2px rgba(245,158,11,.75), 0 0 0 2px rgba(245,158,11,.12) !important;
      }

      .aa-tile.pj-scc-role-trained::after,
      .aa-tile.pj-scc-role-other::after {
        position: absolute !important;
        top: 3px !important;
        right: 4px !important;
        z-index: 5 !important;
        min-width: 17px !important;
        height: 16px !important;
        padding: 0 4px !important;
        border-radius: 999px !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        color: #fff !important;
        font-family: "Segoe UI", Arial, sans-serif !important;
        font-size: 9px !important;
        font-weight: 900 !important;
        line-height: 16px !important;
        pointer-events: none !important;
      }

      .aa-tile.pj-scc-role-pick::after {
        content: "P" !important;
        background: #16a34a !important;
      }

      .aa-tile.pj-scc-role-pack::after {
        content: "PK" !important;
        background: #2563eb !important;
      }

      .aa-tile.pj-scc-role-both::after {
        content: "B" !important;
        background: #7c3aed !important;
      }

      .aa-tile.pj-scc-role-other-training::after {
        content: "i" !important;
        background: #d97706 !important;
      }

      .aa-tile.pj-scc-role-hovered {
        transform: translateY(-1px) !important;
        z-index: 40 !important;
      }

      #pj-p2r-role-tooltip {
        position: fixed !important;
        z-index: 2147483647 !important;
        display: none !important;
        width: 285px !important;
        max-width: min(285px, calc(100vw - 24px)) !important;
        padding: 11px 12px !important;
        border-radius: 12px !important;
        background: rgba(15, 23, 42, .985) !important;
        border: 1px solid rgba(255,255,255,.16) !important;
        box-shadow: 0 14px 34px rgba(0,0,0,.38) !important;
        color: #e5e7eb !important;
        font-family: "Segoe UI", Arial, sans-serif !important;
        font-size: 13px !important;
        line-height: 1.4 !important;
        pointer-events: none !important;
      }

      #pj-p2r-role-tooltip.show {
        display: block !important;
      }

      .pj-p2r-role-tip-login {
        font-size: 16px !important;
        font-weight: 900 !important;
        color: #fff !important;
        margin-bottom: 4px !important;
        letter-spacing: .2px !important;
      }

      .pj-p2r-role-tip-row {
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        gap: 8px !important;
        margin: 6px 0 !important;
      }

      .pj-p2r-role-tip-label {
        color: #cbd5e1 !important;
        font-size: 12px !important;
        font-weight: 700 !important;
      }

      .pj-p2r-role-tip-status {
        display: inline-flex !important;
        align-items: center !important;
        gap: 6px !important;
        padding: 4px 8px !important;
        border-radius: 999px !important;
        color: #fff !important;
        font-size: 11px !important;
        font-weight: 900 !important;
        white-space: nowrap !important;
      }

      .pj-p2r-role-tip-status.ok {
        background: #15803d !important;
      }

      .pj-p2r-role-tip-status.warn {
        background: #b45309 !important;
      }

      .pj-p2r-role-tip-muted {
        color: #94a3b8 !important;
        font-size: 11px !important;
        margin-top: 7px !important;
        margin-bottom: 6px !important;
      }

      .pj-p2r-role-chip-wrap {
        display: flex !important;
        flex-wrap: wrap !important;
        gap: 5px !important;
        margin-top: 6px !important;
      }

      .pj-p2r-role-chip {
        display: inline-flex !important;
        align-items: center !important;
        border-radius: 999px !important;
        padding: 4px 8px !important;
        color: #fff !important;
        font-size: 11px !important;
        font-weight: 900 !important;
        font-family: "Segoe UI", Arial, sans-serif !important;
        white-space: nowrap !important;
      }

      .pj-p2r-role-chip.pick { background: #16a34a !important; }
      .pj-p2r-role-chip.pack { background: #2563eb !important; }
      .pj-p2r-role-chip.both { background: #7c3aed !important; }
      .pj-p2r-role-chip.rebin { background: #7c3aed !important; }
      .pj-p2r-role-chip.gift { background: #db2777 !important; }
      .pj-p2r-role-chip.icqa { background: #0891b2 !important; }
      .pj-p2r-role-chip.single { background: #475569 !important; }
      .pj-p2r-role-chip.other { background: #64748b !important; }
    `;
    document.head.appendChild(style);
  }

  function p2rRoleChipClass(role) {
    const value = String(role || '').toUpperCase();

    // Strict P2R-only colour coding.
    // ARSAW / AR Pick / Pick_AR must NOT count as P2R Pick.
    // AFE Pack / other Pack roles must NOT count as P2R Pack.
    if (value === 'P2R_PICK' || value.includes('P2R_PICK')) return 'pick';
    if (value === 'P2R_PACK' || value.includes('P2R_PACK')) return 'pack';
    if (value.includes('REBIN')) return 'rebin';
    if (value.includes('GW') || value.includes('GIFT')) return 'gift';
    if (value.includes('ICQA') || value.includes('COUNT')) return 'icqa';
    if (value.includes('SINGLE')) return 'single';

    return 'other';
  }

  function p2rHasPickTraining(roles) {
    const set = new Set((Array.isArray(roles) ? roles : []).map(role => String(role).toUpperCase()));

    // STRICT: only P2R Pick counts.
    // Do not count PICK_AR / ARSAW Pick / generic Pick.
    return set.has('P2R_PICK');
  }

  function p2rHasPackTraining(roles) {
    const set = new Set((Array.isArray(roles) ? roles : []).map(role => String(role).toUpperCase()));

    // STRICT: only P2R Pack counts.
    // Do not count AFE Pack / generic Pack.
    return set.has('P2R_PACK');
  }

  function p2rTrainingCategory(roles) {
    const hasPick = p2rHasPickTraining(roles);
    const hasPack = p2rHasPackTraining(roles);

    if (hasPick && hasPack) return 'both';
    if (hasPick) return 'pick';
    if (hasPack) return 'pack';

    return 'other';
  }

  function p2rRoleFilterMode() {
    const mode = GM_getValue(P2R_ROLE_FILTER_KEY, 'all');
    return ['all', 'pick', 'pack', 'both'].includes(mode) ? mode : 'all';
  }

  function p2rRoleFilterLabel(mode = p2rRoleFilterMode()) {
    const labels = {
      all: 'Show: All Roles',
      pick: 'Show: P2R Pick',
      pack: 'Show: P2R Pack',
      both: 'Show: Both P2R'
    };

    return labels[mode] || labels.all;
  }

  function p2rNextRoleFilterMode(mode = p2rRoleFilterMode()) {
    const order = ['all', 'pick', 'pack', 'both'];
    const index = order.indexOf(mode);
    return order[(index + 1) % order.length];
  }

  function p2rRolePassesFilter(roles, filterMode = p2rRoleFilterMode()) {
    const category = p2rTrainingCategory(roles);

    if (filterMode === 'all') return true;
    if (filterMode === 'pick') return category === 'pick' || category === 'both';
    if (filterMode === 'pack') return category === 'pack' || category === 'both';
    if (filterMode === 'both') return category === 'both';

    return true;
  }

  function p2rRoleCategoryLabel(category) {
    if (category === 'both') return 'P2R Pick + P2R Pack trained';
    if (category === 'pick') return 'P2R Pick trained';
    if (category === 'pack') return 'P2R Pack trained';
    return 'Other trained role';
  }

  function p2rCancelRoleTooltipHide() {
    if (p2rRoleTooltipHideTimer) {
      clearTimeout(p2rRoleTooltipHideTimer);
      p2rRoleTooltipHideTimer = null;
    }
  }

  function p2rScheduleRoleTooltipHide(tile = null, delay = 3000) {
    p2rCancelRoleTooltipHide();

    p2rRoleTooltipHideTimer = setTimeout(() => {
      p2rHideRoleTooltip(tile);
    }, delay);
  }

  function p2rRemoveRoleTooltip() {
    p2rCancelRoleTooltipHide();
    const tip = document.getElementById('pj-p2r-role-tooltip');
    if (tip) tip.remove();
  }

  function p2rEnsureRoleTooltip() {
    let tip = document.getElementById('pj-p2r-role-tooltip');

    if (!tip) {
      tip = document.createElement('div');
      tip.id = 'pj-p2r-role-tooltip';
      document.body.appendChild(tip);
    }

    return tip;
  }

  function p2rShowRoleTooltip(tile, event = null) {
    if (!tile || !p2rRoleHighlightEnabled()) return;

    p2rCancelRoleTooltipHide();

    const login = tile.getAttribute('data-pj-scc-login') || tile.dataset.login || '';
    const fullLabel = tile.getAttribute('data-pj-scc-trained-roles') || '';
    const boardRole = tile.getAttribute('data-pj-scc-board-role') || '';
    const matched = tile.getAttribute('data-pj-scc-role-match') === '1';
    const category = tile.getAttribute('data-pj-scc-role-category') || 'other';
    const trainedRaw = (tile.getAttribute('data-pj-scc-trained-raw') || '')
      .split('|')
      .map(x => x.trim())
      .filter(Boolean);

    const tip = p2rEnsureRoleTooltip();
    const statusText = matched ? 'Matches current slot' : 'Does not match current slot';

    tip.innerHTML = `
      <div class="pj-p2r-role-tip-login">${escapeHtml(login || 'Unknown login')}</div>
      <div class="pj-p2r-role-tip-row">
        <span class="pj-p2r-role-tip-label">Role type</span>
        <span class="pj-p2r-role-chip ${escapeHtml(category === 'both' ? 'both' : category)}">
          ${escapeHtml(p2rRoleCategoryLabel(category))}
        </span>
      </div>
      <div class="pj-p2r-role-tip-row">
        <span class="pj-p2r-role-tip-label">Current slot</span>
        <span class="pj-p2r-role-tip-status ${matched ? 'ok' : 'warn'}">${escapeHtml(boardRole || 'Unknown')} • ${escapeHtml(statusText)}</span>
      </div>
      <div class="pj-p2r-role-tip-muted">Trained roles</div>
      <div class="pj-p2r-role-chip-wrap">
        ${trainedRaw.map(role => `
          <span class="pj-p2r-role-chip ${escapeHtml(p2rRoleChipClass(role))}">
            ${escapeHtml(trainedRoleLabel(role))}
          </span>
        `).join('') || `<span class="pj-p2r-role-chip other">${escapeHtml(fullLabel || 'Role data')}</span>`}
      </div>
    `;

    const rect = tile.getBoundingClientRect();
    const preferredX = event?.clientX ? event.clientX + 14 : rect.left + rect.width + 10;
    const preferredY = event?.clientY ? event.clientY + 14 : rect.top;

    tip.classList.add('show');

    const tipRect = tip.getBoundingClientRect();
    let left = preferredX;
    let top = preferredY;

    if (left + tipRect.width + 10 > window.innerWidth) {
      left = Math.max(10, rect.left - tipRect.width - 10);
    }

    if (top + tipRect.height + 10 > window.innerHeight) {
      top = Math.max(10, window.innerHeight - tipRect.height - 10);
    }

    tip.style.left = `${Math.max(10, left)}px`;
    tip.style.top = `${Math.max(10, top)}px`;
  }

  function p2rHideRoleTooltip(tile = null) {
    if (tile) tile.classList.remove('pj-scc-role-hovered');

    const tip = document.getElementById('pj-p2r-role-tooltip');
    if (tip) tip.classList.remove('show');
  }

  function p2rClearRoleHighlights(removeTooltip = true) {
    document.querySelectorAll('.aa-tile[data-pj-scc-login], .aa-tile.pj-scc-role-trained, .aa-tile.pj-scc-role-other').forEach(tile => {
      tile.classList.remove(
        'pj-scc-role-trained',
        'pj-scc-role-other',
        'pj-scc-role-hovered',
        'pj-scc-role-pick',
        'pj-scc-role-pack',
        'pj-scc-role-both',
        'pj-scc-role-other-training',
        'pj-scc-role-slot-match',
        'pj-scc-role-slot-mismatch'
      );

      tile.removeAttribute('data-pj-scc-trained-roles');
      tile.removeAttribute('data-pj-scc-trained-raw');
      tile.removeAttribute('data-pj-scc-board-role');
      tile.removeAttribute('data-pj-scc-role-match');
      tile.removeAttribute('data-pj-scc-role-category');
      tile.removeAttribute('data-pj-scc-login');
      tile.querySelectorAll('.pj-scc-role-badge').forEach(badge => badge.remove());

      tile.onmouseenter = null;
      tile.onmousemove = null;
      tile.onmouseleave = null;
      tile.onclick = null;
    });

    if (removeTooltip) p2rRemoveRoleTooltip();
  }

  function p2rApplyRoleHighlights(statusBox = null) {
    if (!isP2RTrackerPage()) return;

    injectP2RRoleHighlightStyles();

    // Do not remove the tooltip during normal refresh; this prevents fast disappearing while hovering.
    p2rClearRoleHighlights(false);

    if (!p2rRoleHighlightEnabled()) {
      p2rClearRoleHighlights(true);
      return;
    }

    const trainingMap = p2rGetTrainingMapFromStorage();
    const keys = Object.keys(trainingMap || {});

    if (!keys.length) {
      if (statusBox) statusBox.textContent = 'Role hover ON, but no SCC trained-role data received yet. Keep SCC page open and click Refresh Roles.';
      return;
    }

    const filterMode = p2rRoleFilterMode();
    let highlighted = 0;
    let matched = 0;
    let pickCount = 0;
    let packCount = 0;
    let bothCount = 0;
    let otherCount = 0;

    document.querySelectorAll('.aa-tile[data-login]').forEach(tile => {
      const login = clean(tile.dataset.login || tile.querySelector('.login')?.textContent || '');
      const info = trainingMap[login.toLowerCase()];
      const trained = normalizeTrainedRoles(info?.trained || info?.roles || []);
      if (!login || !trained.length) return;

      if (!p2rRolePassesFilter(trained, filterMode)) return;

      const slot = tile.closest('.slot[data-role]');
      const boardRole = slot?.dataset?.role || '';
      const roleMatched = canDoP2RRole(trained, boardRole);
      const fullLabel = trained.map(trainedRoleLabel).join(', ');
      const category = p2rTrainingCategory(trained);

      tile.classList.add('pj-scc-role-trained');
      tile.classList.add(`pj-scc-role-${category === 'other' ? 'other-training' : category}`);
      tile.classList.add(roleMatched ? 'pj-scc-role-slot-match' : 'pj-scc-role-slot-mismatch');

      tile.setAttribute('data-pj-scc-login', login);
      tile.setAttribute('data-pj-scc-trained-roles', fullLabel);
      tile.setAttribute('data-pj-scc-trained-raw', trained.join('|'));
      tile.setAttribute('data-pj-scc-board-role', boardRole || 'Unknown');
      tile.setAttribute('data-pj-scc-role-match', roleMatched ? '1' : '0');
      tile.setAttribute('data-pj-scc-role-category', category);
      tile.title = `${login} • ${p2rRoleCategoryLabel(category)} • ${fullLabel}`;

      tile.onmouseenter = event => {
        tile.classList.add('pj-scc-role-hovered');
        p2rShowRoleTooltip(tile, event);
      };

      tile.onmousemove = event => {
        p2rShowRoleTooltip(tile, event);
      };

      tile.onmouseleave = () => {
        // Keep popup visible for a while after leaving the tile.
        p2rScheduleRoleTooltipHide(tile, 3200);
      };

      tile.onclick = event => {
        // Touch/tablet fallback: tap a highlighted tile to show the same role card longer.
        tile.classList.add('pj-scc-role-hovered');
        p2rShowRoleTooltip(tile, event);
        p2rScheduleRoleTooltipHide(tile, 9000);
      };

      highlighted += 1;
      if (roleMatched) matched += 1;

      if (category === 'both') bothCount += 1;
      else if (category === 'pick') pickCount += 1;
      else if (category === 'pack') packCount += 1;
      else otherCount += 1;
    });

    if (statusBox) {
      statusBox.textContent =
        `${p2rRoleFilterLabel(filterMode)} active: ${highlighted} marked | Pick ${pickCount}, Pack ${packCount}, Both ${bothCount}, Other ${otherCount}. Hover/tap marker for details.`;
    }
  }

  function startP2RRoleHighlightLoop(statusBox = null) {
    if (!isP2RTrackerPage() || p2rRoleHighlightTimer) return;

    p2rRoleHighlightTimer = setInterval(() => {
      if (p2rRoleHighlightEnabled()) p2rApplyRoleHighlights(statusBox);
    }, 5000);

    setTimeout(() => p2rApplyRoleHighlights(statusBox), 300);
  }


  function runExcelLiveWindow() {
    if (!isExcelPage()) return;
    if (document.getElementById('pj-scc-live-window')) return;

    const helperPageName = getHelperPageName();

    injectExcelStyles();

    const panel = document.createElement('div');
    panel.id = 'pj-scc-live-window';

    panel.innerHTML = `
      <div class="pj-live-header">
        <div>
          <div class="pj-live-title">SCC Live Changes</div>
          <div class="pj-live-subtitle">${escapeHtml(helperPageName)} helper | v${escapeHtml(SCRIPT_VERSION)}</div>
        </div>
        <button class="pj-live-min" id="pj-live-min-btn">−</button>
      </div>

      <div class="pj-live-body">
        <div class="pj-live-status" id="pj-live-status">
          Waiting for SCC login changes...
        </div>

        <div class="pj-live-status" style="background:rgba(16,185,129,0.10) !important;border-color:rgba(16,185,129,0.22) !important;color:#d1fae5 !important;">
          Use <b>Copy Latest Layout</b>, then paste directly into ${escapeHtml(helperPageName)}.
        </div>

        <div class="pj-live-actions">
          <button class="pj-live-btn pj-copy-latest" id="pj-live-copy-layout">Copy Latest Layout</button>
          ${isP2RTrackerPage() ? `<button class="pj-live-btn pj-copy-latest" id="pj-p2r-apply-now">Apply to P2R</button>` : ``}
          ${isP2RTrackerPage() ? `<button class="pj-live-btn pj-clear-log" id="pj-p2r-auto-toggle">Auto Apply: OFF</button>` : ``}
          ${isP2RTrackerPage() ? `<button class="pj-live-btn pj-copy-latest" id="pj-p2r-role-highlight-toggle">Role Hover: OFF</button>` : ``}
          ${isP2RTrackerPage() ? `<button class="pj-live-btn pj-clear-log" id="pj-p2r-role-filter-toggle">Show: All Roles</button>` : ``}
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
    const p2rApplyNowBtn = document.getElementById('pj-p2r-apply-now');
    const p2rAutoToggleBtn = document.getElementById('pj-p2r-auto-toggle');
    const p2rRoleHighlightBtn = document.getElementById('pj-p2r-role-highlight-toggle');
    const p2rRoleFilterBtn = document.getElementById('pj-p2r-role-filter-toggle');

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

        statusBox.textContent = `${data.wallName} latest layout copied with FCLM links. Paste into ${helperPageName}.`;
      } catch (err) {
        console.error('Excel helper copy failed:', err);
        GM_setClipboard(text);
        statusBox.textContent = `${data.wallName} latest layout copied as plain text. Paste into ${helperPageName}.`;
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

    function p2rAutoApplyEnabled() {
      return GM_getValue(P2R_AUTO_APPLY_KEY, false) === true;
    }

    function updateP2RAutoButton() {
      if (!p2rAutoToggleBtn) return;
      const enabled = p2rAutoApplyEnabled();
      p2rAutoToggleBtn.textContent = enabled ? 'Auto Apply: ON' : 'Auto Apply: OFF';
      p2rAutoToggleBtn.style.background = enabled ? '#047857' : '#334155';
    }

    function updateP2RRoleHighlightButton() {
      if (!p2rRoleHighlightBtn) return;
      const enabled = p2rRoleHighlightEnabled();
      p2rRoleHighlightBtn.textContent = enabled ? 'Role Hover: ON' : 'Role Hover: OFF';
      p2rRoleHighlightBtn.style.background = enabled ? '#7c3aed' : '#334155';
    }

    function updateP2RRoleFilterButton() {
      if (!p2rRoleFilterBtn) return;
      const mode = p2rRoleFilterMode();
      p2rRoleFilterBtn.textContent = p2rRoleFilterLabel(mode);

      const colours = {
        all: '#334155',
        pick: '#16a34a',
        pack: '#2563eb',
        both: '#7c3aed'
      };

      p2rRoleFilterBtn.style.background = colours[mode] || '#334155';
      p2rRoleFilterBtn.style.color = '#fff';
    }

    async function p2rApplyLatestFromHelper(sourceLabel = 'manual') {
      const latestLayout = GM_getValue(SCC_LAYOUT_KEY);
      return p2rApplySccLayout(latestLayout, statusBox, sourceLabel);
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

      if (isP2RTrackerPage() && p2rAutoApplyEnabled()) {
        setTimeout(() => p2rApplyLatestFromHelper('auto'), 250);
      }

      if (isP2RTrackerPage() && p2rRoleHighlightEnabled()) {
        setTimeout(() => p2rApplyRoleHighlights(statusBox), 350);
      }
    });

    GM_addValueChangeListener(SCC_LAYOUT_KEY, (name, oldValue, newValue) => {
      if (!newValue) return;

      if (!logs.length) {
        statusBox.textContent = `${newValue.wallName || 'Wall'} layout received at ${newValue.time}. Waiting for changes...`;
      }

      if (isP2RTrackerPage() && p2rAutoApplyEnabled()) {
        setTimeout(() => p2rApplySccLayout(newValue, statusBox, 'auto'), 250);
      }

      if (isP2RTrackerPage() && p2rRoleHighlightEnabled()) {
        setTimeout(() => p2rApplyRoleHighlights(statusBox), 350);
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

    if (p2rApplyNowBtn) {
      p2rApplyNowBtn.addEventListener('click', () => p2rApplyLatestFromHelper('manual'));
    }

    if (p2rAutoToggleBtn) {
      updateP2RAutoButton();
      p2rAutoToggleBtn.addEventListener('click', () => {
        const next = !p2rAutoApplyEnabled();
        GM_setValue(P2R_AUTO_APPLY_KEY, next);
        updateP2RAutoButton();
        statusBox.textContent = next
          ? 'P2R Auto Apply enabled. SCC moves will update this tracker page automatically.'
          : 'P2R Auto Apply disabled.';
      });
    }

    if (p2rRoleHighlightBtn) {
      updateP2RRoleHighlightButton();
      p2rRoleHighlightBtn.addEventListener('click', () => {
        const next = !p2rRoleHighlightEnabled();
        GM_setValue(P2R_ROLE_HIGHLIGHT_KEY, next);
        updateP2RRoleHighlightButton();

        if (next) {
          p2rApplyRoleHighlights(statusBox);
        } else {
          p2rClearRoleHighlights(true);
          statusBox.textContent = 'P2R Role Hover disabled.';
        }
      });
    }

    if (p2rRoleFilterBtn) {
      updateP2RRoleFilterButton();
      p2rRoleFilterBtn.addEventListener('click', () => {
        const next = p2rNextRoleFilterMode();
        GM_setValue(P2R_ROLE_FILTER_KEY, next);
        updateP2RRoleFilterButton();

        if (!p2rRoleHighlightEnabled()) {
          GM_setValue(P2R_ROLE_HIGHLIGHT_KEY, true);
          updateP2RRoleHighlightButton();
        }

        p2rApplyRoleHighlights(statusBox);
      });
    }

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

    if (isP2RTrackerPage() && p2rAutoApplyEnabled() && latest) {
      setTimeout(() => p2rApplySccLayout(latest, statusBox, 'auto'), 600);
    }

    if (isP2RTrackerPage()) {
      updateP2RRoleHighlightButton();
      startP2RRoleHighlightLoop(statusBox);
    }

    makePanelDraggable('pj-scc-live-window', '.pj-live-header', STORAGE_EXCEL_LEFT, STORAGE_EXCEL_TOP);
    rememberPanelSize('pj-scc-live-window', STORAGE_EXCEL_WIDTH, STORAGE_EXCEL_HEIGHT, () => panel.classList.contains('min'));
  }


  // ===== Prince Jacob Custom Update Checker - Every 10 Hours =====
  function princeUpdateChecker() {
    const UPDATE_URL = "https://raw.githubusercontent.com/prince-jacob/SCC_Live_Login_Monitor_Dashboard/main/SCCLiveLoginMonitorDashboard.user.js";
    const CHECK_KEY = "prince_last_update_check_" + GM_info.script.name;
    const CHECK_INTERVAL = 10 * 60 * 60 * 1000; // 10 hours

    const lastCheck = Number(GM_getValue(CHECK_KEY, 0));
    const now = Date.now();

    if (now - lastCheck < CHECK_INTERVAL) {
      return;
    }

    GM_setValue(CHECK_KEY, now);

    GM_xmlhttpRequest({
      method: "GET",
      url: UPDATE_URL,
      nocache: true,
      onload: function (res) {
        const remoteScript = res.responseText || "";
        const remoteMatch = remoteScript.match(/\/\/\s*@version\s+([0-9.]+)/i);

        if (!remoteMatch) {
          console.log("[Update Checker] Remote version not found.");
          return;
        }

        const remoteVersion = remoteMatch[1];
        const currentVersion = GM_info.script.version;

        if (isNewerVersion(remoteVersion, currentVersion)) {
          const openUpdate = confirm(
            "New script update available!\n\n" +
            "Script: " + GM_info.script.name + "\n" +
            "Current version: " + currentVersion + "\n" +
            "New version: " + remoteVersion + "\n\n" +
            "Open update page now?"
          );

          if (openUpdate) {
            window.open(UPDATE_URL, "_blank");
          }
        } else {
          console.log("[Update Checker] Up to date:", currentVersion);
        }
      },
      onerror: function () {
        console.log("[Update Checker] Failed to check update.");
      }
    });

    function isNewerVersion(remote, current) {
      const r = String(remote).split(".").map(Number);
      const c = String(current).split(".").map(Number);
      const len = Math.max(r.length, c.length);

      for (let i = 0; i < len; i++) {
        const rv = r[i] || 0;
        const cv = c[i] || 0;

        if (rv > cv) return true;
        if (rv < cv) return false;
      }

      return false;
    }
  }

  princeUpdateChecker();

  if (isExcelPage()) {
    runExcelLiveWindow();
    return;
  }

  if (!isSccPage()) {
    return;
  }

  setInterval(waitForFloorPlan, 1000);

})();