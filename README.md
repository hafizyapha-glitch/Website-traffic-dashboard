<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Aurora Diamond - Traffic & Conversion Dashboard</title>

  <!-- External Libraries (Switched to jsDelivr for better compression stability) -->
  <script src="https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Noto+Sans+Thai:wght@300;400;500;600;700&display=swap" rel="stylesheet">

  <style>
    :root {
      --primary: #2563EB;
      --primary-dark: #1D4ED8;
      --bg: #F3F4F6;
      --surface: #F9FAFB;
      --card: #FFFFFF;
      --border: #E5E7EB;
      --text: #111827;
      --text-secondary: #4B5563;
      --muted: #6B7280;
      --accent: #F59E0B;
      --green: #10B981;
      --blue: #3B82F6;
      --emerald: #10b981;
      --purple: #8b5cf6;
      --shadow-sm: 0 1px 2px 0 rgba(0,0,0,0.05);
      --shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06);
    }
    body.dark-mode {
      --primary: #3B82F6; --primary-dark: #60A5FA;
      --bg: #0F172A; --surface: #1E293B; --card: #1E293B; --border: #334155;
      --text: #F8FAFC; --text-secondary: #CBD5E1; --muted: #94A3B8;
      --shadow-sm: 0 1px 2px 0 rgba(0,0,0,0.3);
      --shadow: 0 4px 6px -1px rgba(0,0,0,0.4), 0 2px 4px -1px rgba(0,0,0,0.2);
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { background: var(--bg); color: var(--text); font-family: 'Inter','Noto Sans Thai',sans-serif; min-height: 100vh; line-height: 1.5; }

    /* Header & Navigation */
    .header { background: var(--card); border-bottom: 1px solid var(--border); padding: 16px 32px 0 32px; display: flex; flex-direction: column; position: sticky; top: 0; z-index: 100; box-shadow: var(--shadow-sm); }
    .header-top { display: flex; align-items: center; justify-content: space-between; padding-bottom: 16px; width: 100%; }
    .logo-area { display: flex; align-items: center; gap: 16px; }
    .logo-icon { width: 50px; height: 50px; border-radius: 12px; overflow: hidden; display: flex; align-items: center; justify-content: center; border: 1px solid var(--border); background: var(--surface); box-shadow: var(--shadow-sm); color: var(--primary); }
    .logo-icon svg { width: 28px; height: 28px; }
    .logo-title { font-size: 20px; font-weight: 700; color: var(--text); letter-spacing: -0.5px; }
    .logo-sub { font-size: 13px; color: var(--muted); font-weight: 500; }
    .header-right { display: flex; align-items: center; gap: 16px; }
    .date-badge { background: var(--surface); border: 1px solid var(--border); border-radius: 6px; padding: 6px 12px; font-size: 13px; font-weight: 500; color: var(--text-secondary); }
    
    .platform-tabs { display: flex; gap: 24px; border-top: 1px solid var(--border); padding-top: 8px; }
    .platform-tab { background: transparent; border: none; font-family: 'Inter', 'Noto Sans Thai', sans-serif; font-size: 14px; font-weight: 600; color: var(--text-secondary); cursor: pointer; padding: 10px 4px; border-bottom: 2px solid transparent; transition: all .2s; display: flex; align-items: center; gap: 6px; }
    .platform-tab svg { width: 16px; height: 16px; opacity: 0.7; }
    .platform-tab.active { color: var(--primary); border-bottom-color: var(--primary); }
    .platform-tab.active svg { opacity: 1; }
    .platform-tab:hover:not(.active) { color: var(--text); }

    .main { padding: 18px 64px; max-width: 1440px; margin: 0 auto; display: flex; flex-direction: column; gap: 24px; }

    .btn-outline { background: var(--surface); border: 1px solid var(--border); color: var(--text); cursor: pointer; transition: all .2s; }
    .btn-outline:hover { border-color: var(--primary); color: var(--primary); }
    .upload-label { display: flex; align-items: center; gap: 8px; background: var(--surface); border: 1px solid var(--border); border-radius: 20px; padding: 7px 16px; cursor: pointer; font-size: 12px; font-weight: 500; color: var(--text-secondary); transition: all .2s; margin: 0; }
    .upload-label:hover { border-color: var(--primary); color: var(--primary); }
    input[type="file"] { display: none; }

    .filter-row { display: flex; gap: 12px; flex-wrap: wrap; align-items: center; }
    .filter-row select { background: var(--card); border: 1px solid var(--border); border-radius: 8px; color: var(--text); padding: 8px 32px 8px 16px; font-size: 13px; font-weight: 500; font-family: 'Inter','Noto Sans Thai',sans-serif; outline: none; cursor: pointer; transition: border-color .2s; appearance: none; background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%2364748B' stroke-width='2' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E"); background-repeat: no-repeat; background-position: right 12px center; box-shadow: var(--shadow-sm); }
    body.dark-mode .filter-row select { background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%2394A3B8' stroke-width='2' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E"); }
    .filter-row select:focus { border-color: var(--primary); }
    .filter-label { font-size: 14px; font-weight: 500; color: var(--text-secondary); display: flex; align-items: center; gap: 6px; }
    .clear-btn { background: transparent; border: 1px solid transparent; border-radius: 8px; color: var(--muted); padding: 8px 12px; font-size: 13px; font-weight: 500; cursor: pointer; transition: all .2s; }
    .clear-btn:hover { color: var(--text); background: var(--surface); border-color: var(--border); }

    .kpi-group-label { font-size: 16px; font-weight: 600; color: var(--text); margin-bottom: 16px; margin-top: 16px; }
    .kpi-grid { display: grid; gap: 16px; margin-bottom: 8px; }
    .kpi-grid.row4 { grid-template-columns: repeat(4,1fr); }
    .kpi-grid.row5 { grid-template-columns: repeat(5,1fr); }
    .kpi-card { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 20px; box-shadow: var(--shadow-sm); display: grid; grid-template-areas: "label icon" "value icon" "sub sub" "bench bench"; grid-template-columns: 1fr auto; align-items: start; transition: transform .2s, box-shadow .2s; }
    .kpi-card:hover { transform: translateY(-2px); box-shadow: var(--shadow); }
    .kpi-icon { grid-area: icon; width: 44px; height: 44px; border-radius: 10px; display: flex; align-items: center; justify-content: center; }
    .kpi-icon svg { width: 22px; height: 22px; stroke: currentColor; stroke-width: 1.75; }
    .kpi-label { grid-area: label; font-size: 13px; font-weight: 500; color: var(--text-secondary); margin-bottom: 4px; }
    .kpi-value { grid-area: value; font-size: 24px; font-weight: 700; font-family: 'Inter',sans-serif; color: var(--text); line-height: 1.2; margin-bottom: 8px; }
    .kpi-sub { grid-area: sub; font-size: 12px; color: var(--muted); }
    .kpi-benchmark { grid-area: bench; font-size: 11px; color: var(--text-secondary); margin-top: 12px; padding-top: 12px; border-top: 1px solid var(--border); display: flex; align-items: center; justify-content: space-between; }
    .kpi-benchmark .bm-val { color: var(--text); font-weight: 600; font-family: 'Inter',sans-serif; }

    /* Layout */
    .dashboard-layout { display: flex; flex-direction: column; gap: 24px; margin-top: 24px; }
    .charts-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }
    .chart-card { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 24px; box-shadow: var(--shadow-sm); }
    .chart-card.wide { grid-column: span 2; }
    .chart-title { font-size: 16px; font-weight: 600; color: var(--text); margin-bottom: 4px; }
    .chart-sub { font-size: 13px; color: var(--text-secondary); margin-bottom: 24px; }
    .chart-wrap { position: relative; height: 280px; }
    .chart-summary { margin-top: 16px; padding: 12px 16px; background: rgba(59,130,246,0.05); border: 1px solid rgba(59,130,246,0.1); border-radius: 8px; font-size: 13px; color: var(--text-secondary); display: flex; align-items: flex-start; gap: 10px; line-height: 1.5; }
    body.dark-mode .chart-summary { background: rgba(59,130,246,0.1); border-color: rgba(59,130,246,0.2); color: var(--text-secondary); }
    .chart-summary strong { color: var(--text); }
    
    /* Pyramid Funnel Styles */
    .pyramid-funnel-container { display: flex; flex-direction: column; align-items: center; width: 100%; padding-top: 12px; }
    .pyramid-step { display: flex; align-items: center; justify-content: space-between; color: white; padding: 16px 20px; margin-bottom: 18px; border-radius: 8px; position: relative; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06); transition: transform 0.2s; }
    .pyramid-step:hover { transform: translateY(-2px); }
    .pyramid-step-left { display: flex; align-items: center; gap: 12px; min-width: 0; }
    .pyramid-icon { background: rgba(255,255,255,0.2); width: 36px; height: 36px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
    .pyramid-icon svg { width: 18px; height: 18px; stroke: white; }
    .pyramid-name { font-size: 13px; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .pyramid-val { font-size: 18px; font-weight: 700; font-family: 'Inter', sans-serif; flex-shrink: 0; margin-left: 12px; }
    .pyramid-badge { position: absolute; top: -14px; left: 50%; transform: translateX(-50%); background: var(--card); font-size: 11px; font-weight: 700; padding: 4px 12px; border-radius: 20px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.05); display: flex; align-items: center; gap: 4px; z-index: 10; white-space: nowrap; }
    .pyramid-badge svg { color: inherit; width: 12px; height: 12px; }

    /* Top Ads List Styles */
    .top-ad-list { display: flex; flex-direction: column; gap: 12px; height: 100%; justify-content: flex-start; }
    .top-ad-item { display: flex; align-items: center; background: var(--surface); border: 1px solid var(--border); border-radius: 8px; padding: 12px; gap: 16px; transition: transform 0.2s, box-shadow 0.2s; }
    .top-ad-item:hover { transform: translateY(-2px); box-shadow: var(--shadow-sm); border-color: var(--primary); }
    .top-ad-body { flex-grow: 1; min-width: 0; display: flex; flex-direction: column; justify-content: center; }
    .top-ad-name { font-size: 13px; font-weight: 600; color: var(--text); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .top-ad-metrics { display: flex; gap: 16px; align-items: center; flex-shrink: 0; }
    .top-ad-metric { display: flex; flex-direction: column; align-items: flex-end; min-width: 45px; }
    .top-ad-metric-val { font-size: 14px; font-weight: 700; color: var(--text); font-family: 'Inter', sans-serif; }
    .top-ad-metric-lbl { font-size: 11px; color: var(--muted); }
    .empty-state { display: flex; align-items: center; justify-content: center; height: 100%; color: var(--muted); font-size: 13px; }

    .table-card { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 0; box-shadow: var(--shadow-sm); overflow: hidden; margin-top: 24px; }
    .table-header { padding: 24px 24px 0; display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
    .table-responsive { width: 100%; overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; min-width: 900px; text-align: left; }
    th { font-size: 12px; font-weight: 600; color: var(--text-secondary); padding: 12px 24px; border-bottom: 1px solid var(--border); background: var(--surface); white-space: nowrap; }
    td { padding: 16px 24px; font-size: 13px; color: var(--text); border-bottom: 1px solid var(--border); white-space: nowrap; }
    tr:last-child td { border-bottom: none; }
    tr:hover td { background: var(--surface); }
    .sortable { cursor: pointer; user-select: none; transition: color 0.2s; }
    .sortable:hover { color: var(--primary); }
    .sort-icon { display: inline-block; width: 12px; margin-left: 4px; font-size: 10px; color: var(--primary); }
    .num { font-family: 'Inter',sans-serif; font-variant-numeric: tabular-nums; }
    .cost-cell { font-weight: 600; font-family: 'Inter',sans-serif; color: var(--text); }

    #statusBar { display: none; background: rgba(16,185,129,0.1); border: 1px solid rgba(16,185,129,0.2); border-radius: 8px; padding: 12px 16px; font-size: 14px; font-weight: 500; color: var(--green); margin-bottom: 16px; }

    /* Settings Drawer Styles */
    .drawer-overlay { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: rgba(0,0,0,0.5); z-index: 999; opacity: 0; visibility: hidden; transition: all 0.3s ease; }
    .drawer-overlay.show { opacity: 1; visibility: visible; }
    .drawer { position: fixed; top: 0; right: -400px; width: 100%; max-width: 400px; height: 100vh; background: var(--card); z-index: 1000; box-shadow: -4px 0 15px rgba(0,0,0,0.1); transition: right 0.3s ease; display: flex; flex-direction: column; border-left: 1px solid var(--border); }
    .drawer.open { right: 0; }
    .drawer-header { display: flex; align-items: center; justify-content: space-between; padding: 20px 24px; border-bottom: 1px solid var(--border); }
    .drawer-title { font-size: 18px; font-weight: 600; color: var(--text); display: flex; align-items: center; gap: 8px; }
    .drawer-close { background: transparent; border: none; cursor: pointer; color: var(--text-secondary); display: flex; align-items: center; justify-content: center; padding: 4px; border-radius: 6px; transition: all 0.2s; }
    .drawer-close:hover { background: var(--surface); color: var(--text); }
    .drawer-content { padding: 24px; overflow-y: auto; flex-grow: 1; }
    .setting-group { margin-bottom: 24px; }
    .setting-label { font-size: 14px; font-weight: 600; color: var(--text); margin-bottom: 8px; display: flex; align-items: center; gap: 6px;}
    .setting-desc { font-size: 12px; color: var(--text-secondary); margin-bottom: 16px; line-height: 1.5; }
    
    /* Theme Switch (Slide) */
    .theme-switch-wrapper { display: flex; align-items: center; justify-content: space-between; background: var(--surface); padding: 12px 16px; border-radius: 8px; border: 1px solid var(--border); }
    .theme-switch-text { font-size: 13px; font-weight: 500; color: var(--text); display: flex; align-items: center; gap: 8px; }
    .switch { position: relative; display: inline-block; width: 44px; height: 24px; flex-shrink: 0; }
    .switch input { opacity: 0; width: 0; height: 0; }
    .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #CBD5E1; transition: .3s; border-radius: 24px; }
    .slider:before { position: absolute; content: ""; height: 18px; width: 18px; left: 3px; bottom: 3px; background-color: white; transition: .3s; border-radius: 50%; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
    input:checked + .slider { background-color: var(--primary); }
    input:checked + .slider:before { transform: translateX(20px); }
    body.dark-mode .slider { background-color: #334155; }
    body.dark-mode input:checked + .slider { background-color: var(--primary); }

  @media (max-width: 900px) {
    .main { padding: 16px; gap: 16px; }
    .charts-grid { grid-template-columns: 1fr; }
    .chart-card.wide { grid-column: span 1; }
    .header { padding: 16px 16px 0 16px; flex-direction: column; align-items: flex-start; }
    .header-top { flex-direction: column; align-items: flex-start; gap: 16px; padding-bottom: 12px; }
    .platform-tabs { width: 100%; overflow-x: auto; gap: 16px; }
    .kpi-grid.row4, .kpi-grid.row5 { grid-template-columns: repeat(2,1fr); }
    td, th { padding: 12px 16px; }
  }
  @media (max-width: 600px) {
    .kpi-grid.row4, .kpi-grid.row5 { grid-template-columns: 1fr; }
    .pyramid-step { width: 100% !important; }
  }
</style>
</head>
<body>

<div class="header">
  <div class="header-top">
    <div class="logo-area">
      <div class="logo-icon">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3h12l4 6-10 13L2 9Z"/><path d="M11 3 8 9l4 13 4-13-3-6"/><path d="M2 9h20"/></svg>
      </div>
      <div class="logo-text">
        <div class="logo-title">Traffic & Conversion Dashboard</div>
        <div class="logo-sub">วิเคราะห์ประสิทธิภาพแคมเปญและเส้นทางของลูกค้า (Customer Journey)</div>
      </div>
    </div>
    <div class="header-right">
      <div class="date-badge" id="dataSourceLabel" style="display: flex; align-items: center; gap: 6px;"><span style="color:var(--accent)">●</span> แหล่งข้อมูล: ข้อมูลตัวอย่าง</div>
      <button id="settingsBtn" class="btn-outline" style="padding:8px;display:flex;align-items:center;justify-content:center;border-radius:50%;width:36px;height:36px;" title="ตั้งค่า (Settings)">
        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>
      </button>
    </div>
  </div>
  <!-- Primary Navigation Platform Tabs -->
  <div class="platform-tabs">
    <button id="tab-facebook" class="platform-tab active" onclick="switchPlatform('facebook')">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 2h-3a5 5 0 0 0-5 5v3H7v4h3v8h4v-8h3l1-4h-4V7a1 1 0 0 1 1-1h3z"/></svg>
      Facebook Ads
    </button>
    <button id="tab-google" class="platform-tab" onclick="switchPlatform('google')">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 8v8"/><path d="M8 12h8"/></svg>
      Google Ads
    </button>
  </div>
</div>

<div class="main">
  <div style="padding: 0 32px; margin-top: -8px;" id="innerMain">
    <div id="statusBar"></div>
    <div class="filter-row">
      <span class="filter-label">
        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="21" y1="4" x2="14" y2="4"/><line x1="10" y1="4" x2="3" y2="4"/><line x1="21" y1="12" x2="12" y2="12"/><line x1="8" y1="12" x2="3" y2="12"/><line x1="21" y1="20" x2="16" y2="20"/><line x1="12" y1="20" x2="3" y2="20"/><line x1="14" y1="2" x2="14" y2="6"/><line x1="8" y1="10" x2="8" y2="14"/><line x1="16" y1="18" x2="16" y2="22"/></svg>
        Filter:
      </span>
      <select id="filterCampaign" onchange="applyFilters()">
        <option value="">ทุก Campaign</option>
      </select>
      <button class="clear-btn" onclick="clearFilters()">✕ ล้างตัวกรอง</button>
    </div>

    <!-- KPIs Section -->
    <div id="kpiSection"></div>

    <!-- Dashboard Layout -->
    <div class="dashboard-layout">
      
      <!-- Row 1: Funnel & Sales Summary -->
      <div class="charts-grid">
        <!-- Customer Journey Funnel -->
        <div class="chart-card">
          <div class="chart-title" style="display: flex; align-items: center; gap: 8px;">
            Customer Journey Funnel
          </div>
          <div class="chart-sub">พฤติกรรมลูกค้าและการร่วงหล่น (Drop-off Rate)</div>
          <div id="funnelContainer">
            <!-- Funnel steps injected via JS -->
          </div>
        </div>

        <!-- Sales Performance Summary -->
        <div class="chart-card" style="display: flex; flex-direction: column;">
          <div class="chart-title">Sale Performance</div>
          <div class="chart-sub" style="margin-bottom: 16px;">ข้อมูลการสั่งซื้อและผลตอบแทนแบบรวมศูนย์</div>
          <div id="salesSummaryContainer" style="flex-grow: 1; display: grid; grid-template-columns: 1fr 1fr; gap: 16px; align-content: center;">
            <!-- Content injected via JS -->
          </div>
        </div>
      </div>

      <!-- Row 2: Conversion Charts -->
      <div class="charts-grid">
        <!-- Demographics -->
        <div class="chart-card wide">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; flex-wrap: wrap; gap: 8px;">
            <div class="chart-title" id="demoChartTitle">Demographics by Age & Gender</div>
            <select id="demoMetricSelect" onchange="updateDemoMetric()" style="padding: 4px 8px; border-radius: 6px; border: 1px solid var(--border); background: var(--surface); color: var(--text); font-size: 12px; font-family: 'Inter', sans-serif; cursor: pointer; outline: none;">
              <option value="spend">Spend (THB)</option>
              <option value="reach">Reach</option>
              <option value="impressions">Impressions</option>
              <option value="clicks">Link Clicks</option>
              <option value="lpViews">Landing Page Views</option>
              <option value="addCart">Adds to Cart</option>
              <option value="purchases">Purchases</option>
            </select>
          </div>
          <div class="chart-sub" id="demoChartSub">การวิเคราะห์กลุ่มเป้าหมายแยกตามอายุและเพศ (อิงตามค่าใช้จ่าย)</div>
          <div class="chart-wrap"><canvas id="demoChart"></canvas></div>
          <div id="demoSummary" class="chart-summary" style="display:none;"></div>
        </div>

        <!-- Ad Performance by Spend -->
        <div class="chart-card wide">
          <div class="chart-title">Ad Performance by Spend</div>
          <div class="chart-sub">สัดส่วนค่าใช้จ่าย (Spend) แยกตามชื่อโฆษณา</div>
          <div class="chart-wrap"><canvas id="adPieChart"></canvas></div>
        </div>
      </div>
    </div>

    <!-- Data Detail Table -->
    <div class="table-card" style="margin-top:24px;">
      <div class="table-header">
        <div>
          <div class="chart-title">Ad Performance</div>
          <div class="chart-sub" id="tableCount" style="margin-bottom:0;"></div>
        </div>
      </div>
      <div class="table-responsive">
        <table>
          <thead><tr>
            <th class="sortable" onclick="handleSort('ad')">Ad Name <span class="sort-icon" id="sortIcon-ad"></span></th>
            <th class="sortable" onclick="handleSort('spend')">Spend (฿) <span class="sort-icon" id="sortIcon-spend"></span></th>
            <th class="sortable" onclick="handleSort('roas')">ROAS <span class="sort-icon" id="sortIcon-roas"></span></th>
            <th class="sortable" onclick="handleSort('purchases')">Purchases <span class="sort-icon" id="sortIcon-purchases"></span></th>
            <th class="sortable" onclick="handleSort('cpa')">CPA (฿) <span class="sort-icon" id="sortIcon-cpa"></span></th>
            <th class="sortable" onclick="handleSort('convValue')">Conv. Value (฿) <span class="sort-icon" id="sortIcon-convValue"></span></th>
            <th class="sortable" onclick="handleSort('clicks')">Link Clicks <span class="sort-icon" id="sortIcon-clicks"></span></th>
            <th class="sortable" onclick="handleSort('lpViews')">LP Views <span class="sort-icon" id="sortIcon-lpViews"></span></th>
            <th class="sortable" onclick="handleSort('addCart')">Adds to Cart <span class="sort-icon" id="sortIcon-addCart"></span></th>
            <th class="sortable" onclick="handleSort('reach')">Reach <span class="sort-icon" id="sortIcon-reach"></span></th>
            <th class="sortable" onclick="handleSort('impressions')">Impressions <span class="sort-icon" id="sortIcon-impressions"></span></th>
          </tr></thead>
          <tbody id="tableBody"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<!-- Settings Drawer -->
<div id="settingsOverlay" class="drawer-overlay"></div>
<div id="settingsDrawer" class="drawer">
  <div class="drawer-header">
    <div class="drawer-title">
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>
      การตั้งค่า (Settings)
    </div>
    <button id="closeDrawerBtn" class="drawer-close">
      <svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    </button>
  </div>
  <div class="drawer-content">
    <div class="setting-group">
      <div class="setting-label">โหมดการแสดงผล (Theme)</div>
      <div class="setting-desc">สลับโหมดสีของหน้าจอระหว่างโหมดสว่างและโหมดมืด</div>
      <div class="theme-switch-wrapper">
        <div class="theme-switch-text" id="themeSwitchText"></div>
        <label class="switch">
          <input type="checkbox" id="themeToggleCheckbox">
          <span class="slider"></span>
        </label>
      </div>
    </div>

    <div style="border-top: 1px solid var(--border); margin: 24px 0;"></div>

    <div class="setting-group">
      <div class="setting-label"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="var(--primary)" stroke-width="2"><path d="M18 2h-3a5 5 0 0 0-5 5v3H7v4h3v8h4v-8h3l1-4h-4V7a1 1 0 0 1 1-1h3z"/></svg> ข้อมูล Facebook Ads</div>
      <div class="setting-desc">อัปโหลดไฟล์ (.csv หรือ .xlsx) สำหรับแพลตฟอร์ม Facebook</div>
      <div style="display:flex;flex-direction:column;align-items:flex-start;gap:8px;">
        <label class="upload-label" for="fbFileInput" style="width:100%; justify-content:center; padding: 10px 16px; font-size: 13px; transition: all 0.3s;">
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
          อัปโหลดไฟล์ Traffic-Report
        </label>
        <input type="file" id="fbFileInput" accept=".csv,.xlsx,.xls">
      </div>
    </div>

    <div style="border-top: 1px solid var(--border); margin: 24px 0;"></div>

    <div class="setting-group">
      <div class="setting-label"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="var(--accent)" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 8v8"/><path d="M8 12h8"/></svg> ข้อมูล Google Ads</div>
      <div class="setting-desc">อัปโหลดไฟล์ (.csv หรือ .xlsx) สำหรับแพลตฟอร์ม Google Ads</div>
      <div style="display:flex;flex-direction:column;align-items:flex-start;gap:8px;">
        <label class="upload-label" for="ggFileInput" style="width:100%; justify-content:center; padding: 10px 16px; font-size: 13px; transition: all 0.3s;">
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
          อัปโหลดไฟล์ Google Report
        </label>
        <input type="file" id="ggFileInput" accept=".csv,.xlsx,.xls">
      </div>
    </div>

  </div>
</div>

<script>
// ─── DATA & INIT ──────────────────────────────
// กำหนด Header ทิ้งไว้เพื่อให้ระบบอ่านค่าได้ แต่ไม่มีข้อมูล (Row) ข้างใน
const initialCSVData = `"Campaign name","Ad set name","Age","Gender","Amount spent (THB)","Reach","Impressions","Link clicks","Landing page views","Website adds to cart","Website purchases","Website purchases conversion value"`;

let platformData = {
  facebook: [],
  google: []
};
let uploadedFiles = {
  facebook: 'ไม่มีข้อมูล',
  google: 'ไม่มีข้อมูล'
};
let currentPlatform = 'facebook';
let allData = [];
let filteredData = [];
let charts = {};

let currentDemoMetric = 'spend';
window.updateDemoMetric = function() {
  currentDemoMetric = document.getElementById('demoMetricSelect').value;
  renderCharts();
};

window.addEventListener('DOMContentLoaded', () => {
    // กำหนด Path ของไฟล์ที่ต้องการดึงข้อมูลอัตโนมัติ (เปลี่ยนชื่อไฟล์ที่นี่ได้เลย)
    const facebookDataPath = './data/facebook-data.csv'; 
    const googleDataPath = './data/google-data.csv';

    // ฟังก์ชันดึงข้อมูลจาก Path
    async function fetchDataFromPath(path, platform) {
      try {
        const response = await fetch(path);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        
        const ext = path.split('.').pop().toLowerCase();
        
        if (ext === 'xlsx' || ext === 'xls') {
          const arrayBuffer = await response.arrayBuffer();
          const wb = XLSX.read(arrayBuffer, {type: 'array'});
          const ws = wb.Sheets[wb.SheetNames[0]];
          loadData(XLSX.utils.sheet_to_json(ws, {defval: 0}), platform, true);
        } else {
          // สันนิษฐานว่าเป็น CSV
          const text = await response.text();
          Papa.parse(text, {
            header: true,
            skipEmptyLines: true,
            complete: function(results) {
              loadData(results.data, platform, true);
            }
          });
        }
        
        uploadedFiles[platform] = `ดึงอัตโนมัติ (${path})`;
        if(platform === currentPlatform) updateSourceLabel();

      } catch (error) {
        console.warn(`ไม่พบไฟล์หรือโหลดไม่ได้จาก ${path}, กำลังกำหนดให้เป็นค่าว่าง`, error);
        // Fallback: ใช้ค่าว่าง (initialCSVData ที่ไม่มี row) หากดึงไฟล์ไม่สำเร็จ
        if(platform === 'facebook' || platform === 'google') {
          Papa.parse(initialCSVData, {
            header: true,
            complete: function(results) {
              loadData(results.data, platform, true); 
            }
          });
        }
      }
    }

    // เริ่มการดึงข้อมูลตอนโหลดหน้าเว็บ
    fetchDataFromPath(facebookDataPath, 'facebook');
    // fetchDataFromPath(googleDataPath, 'google'); // ยกเลิกคอมเมนต์บรรทัดนี้เมื่อมีไฟล์ของ Google
});

// Switch Platform Logic
window.switchPlatform = function(platform) {
  currentPlatform = platform;
  document.querySelectorAll('.platform-tab').forEach(el => el.classList.remove('active'));
  document.getElementById('tab-' + platform).classList.add('active');
  
  allData = platformData[platform] || [];
  clearFilters(); // This resets filters, copies allData to filteredData, and calls renderAll()
  updateSourceLabel();
};

function updateSourceLabel() {
  const sourceLabel = document.getElementById('dataSourceLabel');
  if (sourceLabel) {
    const isSampleOrEmpty = (uploadedFiles[currentPlatform] === 'ข้อมูลตัวอย่าง (FB)' || uploadedFiles[currentPlatform] === 'ไม่มีข้อมูล');
    const color = isSampleOrEmpty ? 'var(--accent)' : 'var(--emerald)';
    sourceLabel.innerHTML = `<span style="color:${color}">●</span> แหล่งข้อมูล: ${uploadedFiles[currentPlatform]}`;
  }
}

function destroyCharts() { Object.values(charts).forEach(c => { if(c) c.destroy(); }); charts = {}; }
function formatCurrency(n) { return new Intl.NumberFormat('th-TH', { style: 'currency', currency: 'THB' }).format(n); }
function formatNum(n) { return new Intl.NumberFormat('th-TH').format(Math.round(n)); }

function renderAll() { 
  renderKPIs(); 
  renderFunnel(); 
  renderSalesSummary();
  renderCharts(); 
  renderTable(); 
}

// ─── KPIS ──────────────────────────────────
function renderKPIs() {
  const d = filteredData;
  const spend = d.reduce((s,r)=>s+r.spend,0);
  const reach = d.reduce((s,r)=>s+r.reach,0);
  const impressions = d.reduce((s,r)=>s+r.impressions,0);
  const clicks = d.reduce((s,r)=>s+r.clicks,0);

  // Cost Efficiency Metrics
  const cpr = reach > 0 ? (spend / reach) * 1000 : 0;
  const cpm = impressions > 0 ? (spend / impressions) * 1000 : 0;
  const cpc = clicks > 0 ? spend / clicks : 0;
  const ctr = impressions > 0 ? (clicks / impressions) * 100 : 0;
  const freq = reach > 0 ? impressions / reach : 0;

  function card(icon, label, value, sub, color, benchmark = '') {
    const bmHtml = benchmark ? `<div class="kpi-benchmark"><span class="bm-label">เกณฑ์มาตรฐาน:</span><span class="bm-val">${benchmark}</span></div>` : '';
    return `<div class="kpi-card">
      <div class="kpi-label">${label}</div>
      <div class="kpi-icon" style="color:${color};background:color-mix(in srgb,${color} 12%,transparent);">${icon}</div>
      <div class="kpi-value">${value}</div>
      <div class="kpi-sub">${sub}</div>${bmHtml}</div>`;
  }

  // Icons
  const iconSpend = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>';
  const iconReach = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>';
  const iconImpr = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>';
  const iconClick = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 3h6v6"/><path d="M9 21H3v-6"/><path d="M21 3l-7 7"/><path d="M3 21l7-7"/></svg>';
  
  const iconCPR = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="8.5" cy="7" r="4"/><polyline points="17 11 19 13 23 9"/></svg>';
  const iconCPM = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg>';
  const iconCPC = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/><path d="m15 15 6 6"/></svg>';
  const iconCTR = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="m16 8-8 8"/><path d="M9 9h.01"/><path d="M15 15h.01"/></svg>';
  const iconFreq = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>';

  // Overview Category
  const overviewKPIs = [
    card(iconSpend, 'Spending', formatCurrency(spend), 'Amount spent (THB)', 'var(--blue)'),
    card(iconReach, 'Reach', formatNum(reach), 'Unique people reached', 'var(--emerald)'),
    card(iconImpr, 'Impressions', formatNum(impressions), 'Total ad impressions', 'var(--purple)'),
    card(iconClick, 'Link Clicks', formatNum(clicks), 'Total Link clicks', 'var(--accent)')
  ];

  // Cost Efficiency Category
  const costKPIs = [
    card(iconCPR, 'CPR', '฿' + cpr.toFixed(2), 'ต้นทุนต่อการเข้าถึงพันคน', 'var(--emerald)', '฿30 - ฿100'),
    card(iconCPM, 'CPM', '฿' + cpm.toFixed(2), 'ต้นทุนต่อการแสดงผลพันครั้ง', 'var(--blue)', '฿25 - ฿80'),
    card(iconCPC, 'CPC', '฿' + cpc.toFixed(2), 'ต้นทุนเฉลี่ยต่อการคลิก', 'var(--purple)', '฿3 - ฿15'),
    card(iconCTR, 'CTR', ctr.toFixed(2) + '%', 'อัตราการคลิกต่อการแสดงผล', '#f97316', '1.0% - 3.0%'),
    card(iconFreq, 'Frequency', freq.toFixed(2), 'ความถี่เฉลี่ยต่อคน', 'var(--accent)', '1.5 - 3.0')
  ];

  document.getElementById('kpiSection').innerHTML = `
    <div class="kpi-group-label">Overview</div>
    <div class="kpi-grid row4" style="margin-bottom: 24px;">${overviewKPIs.join('')}</div>
    <div class="kpi-group-label">Cost Efficiency</div>
    <div class="kpi-grid row5">${costKPIs.join('')}</div>
    <div style="text-align: right; font-size: 11px; color: var(--muted); margin-top: 8px;">
      Benchmark ref: Meta Business Suite / E-commerce Industry Average (Thailand)
    </div>
  `;
}

// ─── FUNNEL (CUSTOMER JOURNEY) ─────────────
function renderFunnel() {
  const d = filteredData;
  const impr = d.reduce((s,r)=>s+r.impressions,0);
  const clicks = d.reduce((s,r)=>s+r.clicks,0);
  const lpViews = d.reduce((s,r)=>s+r.lpViews,0);
  const addCart = d.reduce((s,r)=>s+r.addCart,0);
  const purchases = d.reduce((s,r)=>s+r.purchases,0);

  const steps = [
    { id: 1, name: 'การมองเห็น (Impressions)', val: impr, color: '#3B82F6', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>' },
    { id: 2, name: 'คลิกลิงก์ (Link Clicks)', val: clicks, color: '#06B6D4', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M15 3h6v6"/><path d="M9 21H3v-6"/><path d="M21 3l-7 7"/><path d="M3 21l7-7"/></svg>' },
    { id: 3, name: 'เข้าชมเว็บ (Landing Page Views)', val: lpViews, color: '#8B5CF6', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>' },
    { id: 4, name: 'หยิบใส่ตะกร้า (Add to Cart)', val: addCart, color: '#EC4899', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><circle cx="9" cy="21" r="1"/><circle cx="20" cy="21" r="1"/><path d="M1 1h4l2.68 13.39a2 2 0 0 0 2 1.61h9.72a2 2 0 0 0 2-1.61L23 6H6"/></svg>' },
    { id: 5, name: 'สั่งซื้อสำเร็จ (Purchases)', val: purchases, color: '#10B981', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><rect x="1" y="4" width="22" height="16" rx="2" ry="2"/><line x1="1" y1="10" x2="23" y2="10"/></svg>' }
  ];

  let html = '<div class="pyramid-funnel-container">';
  const widths = [100, 90, 80, 70, 60];

  steps.forEach((step, index) => {
    const prevVal = index > 0 ? steps[index - 1].val : step.val;
    const convRate = prevVal > 0 ? ((step.val / prevVal) * 100).toFixed(2) : '0.00';
    const dropRate = prevVal > 0 ? (100 - ((step.val / prevVal) * 100)).toFixed(2) : '0.00';
    const stepWidth = widths[index];

    let badgeHtml = '';
    if (index > 0) {
      badgeHtml = `
        <div class="pyramid-badge" style="color: ${step.color}">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/></svg>
          ${convRate}% | -${dropRate}%
        </div>
      `;
    }

    html += `
      <div class="pyramid-step" style="width: ${stepWidth}%; background-color: ${step.color};">
        ${badgeHtml}
        <div class="pyramid-step-left">
          <div class="pyramid-icon">${step.icon}</div>
          <div class="pyramid-name" title="${step.name}">${step.name}</div>
        </div>
        <div class="pyramid-val">${formatNum(step.val)}</div>
      </div>
    `;
  });

  html += '</div>';
  document.getElementById('funnelContainer').innerHTML = html;
}

// ─── SALES SUMMARY ─────────────────────────
function renderSalesSummary() {
  const d = filteredData;
  const spend = d.reduce((s,r)=>s+r.spend,0);
  const purchases = d.reduce((s,r)=>s+r.purchases,0);
  const convValue = d.reduce((s,r)=>s+r.convValue,0);
  const clicks = d.reduce((s,r)=>s+r.clicks,0);

  const aov = purchases > 0 ? convValue / purchases : 0;
  const roas = spend > 0 ? convValue / spend : 0;
  const convRate = clicks > 0 ? (purchases / clicks) * 100 : 0;

  // Dynamic Benchmark Analysis
  let benchmarkStatus = '';
  if (clicks === 0) {
    benchmarkStatus = `<span style="color: var(--muted); font-size: 11px;">⚠️ ยังไม่มีข้อมูลทราฟฟิกคลิก</span>`;
  } else if (convRate >= 3.00) {
    benchmarkStatus = `<span style="color: var(--emerald); font-weight: 600; font-size: 11px; display: flex; align-items: center; gap: 4px;">🔥 ดีเยี่ยม (สูงกว่าเกณฑ์ทั่วไป > 3.0%)</span>`;
  } else if (convRate >= 1.00) {
    benchmarkStatus = `<span style="color: var(--emerald); font-weight: 500; font-size: 11px; display: flex; align-items: center; gap: 4px;">✨ ปกติ (ตามเกณฑ์มาตรฐาน 1.0% - 3.0%)</span>`;
  } else {
    benchmarkStatus = `<span style="color: var(--accent); font-weight: 500; font-size: 11px; display: flex; align-items: center; gap: 4px;">⚠️ ต่ำกว่าเกณฑ์เฉลี่ย (ควรเช็กความเร็วเว็บหรือข้อเสนอพิเศษ)</span>`;
  }

  document.getElementById('salesSummaryContainer').innerHTML = `
    <div style="background: var(--surface); padding: 20px; border-radius: 12px; border: 1px solid var(--border); display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; transition: transform 0.2s; cursor: default;" onmouseover="this.style.transform='translateY(-2px)'" onmouseout="this.style.transform='translateY(0)'">
      <div style="color: var(--text-secondary); font-size: 13px; font-weight: 500; margin-bottom: 8px;">Website Purchases</div>
      <div style="color: var(--purple); font-size: 28px; font-weight: 700; font-family: 'Inter', sans-serif; line-height: 1;">${formatNum(purchases)}</div>
    </div>
    <div style="background: var(--surface); padding: 20px; border-radius: 12px; border: 1px solid var(--border); display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; transition: transform 0.2s; cursor: default;" onmouseover="this.style.transform='translateY(-2px)'" onmouseout="this.style.transform='translateY(0)'">
      <div style="color: var(--text-secondary); font-size: 13px; font-weight: 500; margin-bottom: 8px;">Conversion Value</div>
      <div style="color: var(--emerald); font-size: 28px; font-weight: 700; font-family: 'Inter', sans-serif; line-height: 1;">฿${formatNum(convValue)}</div>
    </div>
    <div style="background: var(--surface); padding: 20px; border-radius: 12px; border: 1px solid var(--border); display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; transition: transform 0.2s; cursor: default;" onmouseover="this.style.transform='translateY(-2px)'" onmouseout="this.style.transform='translateY(0)'">
      <div style="color: var(--text-secondary); font-size: 13px; font-weight: 500; margin-bottom: 8px;">AOV (ยอดซื้อเฉลี่ย)</div>
      <div style="color: var(--blue); font-size: 28px; font-weight: 700; font-family: 'Inter', sans-serif; line-height: 1;">฿${formatNum(aov)}</div>
    </div>
    <div style="background: var(--surface); padding: 20px; border-radius: 12px; border: 1px solid var(--border); display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; transition: transform 0.2s; cursor: default;" onmouseover="this.style.transform='translateY(-2px)'" onmouseout="this.style.transform='translateY(0)'">
      <div style="color: var(--text-secondary); font-size: 13px; font-weight: 500; margin-bottom: 8px;">ROAS</div>
      <div style="color: var(--accent); font-size: 28px; font-weight: 700; font-family: 'Inter', sans-serif; line-height: 1;">${roas.toFixed(2)}x</div>
    </div>
    <div style="grid-column: 1 / -1; background: var(--surface); padding: 20px; border-radius: 12px; border: 1px solid var(--border); display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; transition: transform 0.2s; cursor: default;" onmouseover="this.style.transform='translateY(-2px)'" onmouseout="this.style.transform='translateY(0)'">
      <div style="color: var(--text-secondary); font-size: 13px; font-weight: 500; margin-bottom: 4px;">Conversion Rate (Purchases / Clicks)</div>
      <div style="color: var(--primary); font-size: 28px; font-weight: 700; font-family: 'Inter', sans-serif; line-height: 1; margin-bottom: 6px;">${convRate.toFixed(2)}%</div>
      <div style="font-size: 12px; color: var(--text-secondary); display: flex; flex-direction: column; gap: 4px; align-items: center;">
        <div style="font-weight: 500; color: var(--muted)">เกณฑ์เฉลี่ย E-commerce มาตรฐาน: 1.00% - 3.00%</div>
        <div>สถานะ: ${benchmarkStatus}</div>
      </div>
    </div>
  `;
}

// ─── CHARTS ────────────────────────────────
const COLORS = ['#3B82F6','#10B981','#F59E0B','#8B5CF6','#F43F5E','#06B6D4','#64748B'];

function renderCharts() {
  destroyCharts();
  const isDark = document.body.classList.contains('dark-mode');
  const textColor = isDark ? '#F8FAFC' : '#111827';
  const d = filteredData;

  function axisStyle() {
    return { grid:{color:isDark?'rgba(255,255,255,0.05)':'rgba(0,0,0,0.04)'}, ticks:{color:isDark?'#94A3B8':'#6B7280',font:{family:'Inter,Noto Sans Thai',size:11}}, border:{color:'transparent'} };
  }
  function chartOpts(yLabel, stacked) {
    return { responsive:true, maintainAspectRatio:false, plugins:{legend:{labels:{color:textColor,font:{family:'Inter,Noto Sans Thai',size:12}}}}, scales:{x:{...axisStyle(),stacked}, y:{...axisStyle(),stacked,title:{display:true,text:yLabel,color:isDark?'#94A3B8':'#6B7280',font:{size:11}}}} };
  }

  // --- Aggregate By Ad Name ---
  const adStats = {};
  d.forEach(r => {
    if (r.ad) {
      if (!adStats[r.ad]) adStats[r.ad] = { spend: 0, purchases: 0, convValue: 0 };
      adStats[r.ad].spend += r.spend;
      adStats[r.ad].purchases += r.purchases;
      adStats[r.ad].convValue += r.convValue;
    }
  });

  // Demographics - Dynamic Metric by Age & Gender
  const agesDemo = [...new Set(d.map(r=>r.age).filter(a=>a))].sort();
  const gendersDemo = [...new Set(d.map(r=>r.gender).filter(g=>g))].sort();
  
  const metricLabels = {
    spend: 'Spend (THB)', reach: 'Reach', impressions: 'Impressions',
    clicks: 'Link Clicks', lpViews: 'Landing Page Views', addCart: 'Adds to Cart', purchases: 'Purchases'
  };
  const metricThLabels = {
    spend: 'ค่าใช้จ่าย', reach: 'การเข้าถึง (Reach)', impressions: 'การแสดงผล (Impressions)',
    clicks: 'การคลิกลิงก์ (Clicks)', lpViews: 'การเข้าชมเว็บ (LP Views)', addCart: 'หยิบใส่ตะกร้า (Add to Cart)', purchases: 'การสั่งซื้อ (Purchases)'
  };

  const selectedMetricLabel = metricLabels[currentDemoMetric];
  document.getElementById('demoChartSub').textContent = `การวิเคราะห์กลุ่มเป้าหมายแยกตามอายุและเพศ (อิงตาม${metricThLabels[currentDemoMetric]})`;

  const genderColors = {
    'male': '#3B82F6',   /* สีน้ำเงิน */
    'female': '#EC4899', /* สีชมพู */
    'unknown': '#94A3B8' /* สีเทา */
  };

  charts.demo = new Chart(document.getElementById('demoChart'), {
    type:'bar',
    data:{
      labels: agesDemo,
      datasets: gendersDemo.map((g, i) => ({
        label: g,
        data: agesDemo.map(a => d.filter(r=>r.age===a && r.gender===g).reduce((s,r)=>s+(r[currentDemoMetric]||0), 0)),
        backgroundColor: genderColors[g.toLowerCase()] || COLORS[i % COLORS.length],
        borderRadius: 4, borderSkipped: false
      }))
    },
    options: chartOpts(selectedMetricLabel, false)
  });

  // Demographic Summary
  let totalDemo = 0;
  d.forEach(r => totalDemo += (r[currentDemoMetric] || 0));

  if (totalDemo > 0) {
    const sortedPairs = [];
    agesDemo.forEach(age => {
      gendersDemo.forEach(gen => {
        const val = d.filter(r => r.age === age && r.gender === gen).reduce((s, r) => s + (r[currentDemoMetric] || 0), 0);
        if (val > 0) sortedPairs.push({ age, gender: gen, val });
      });
    });
    sortedPairs.sort((a, b) => b.val - a.val);
    const top3 = sortedPairs.slice(0, 3);
    
    let summaryHtml = `<svg style="flex-shrink:0;margin-top:2px;" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="var(--primary)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><path d="M12 8h.01"/></svg> <div style="display:flex; flex-direction:column; gap:4px;">`;
    summaryHtml += `<span>วิเคราะห์อันดับกลุ่มเป้าหมายเชิงลึก (${metricThLabels[currentDemoMetric]}) <strong>3 อันดับสูงสุด:</strong></span>`;
    
    top3.forEach((item, index) => {
      const genderTh = item.gender === 'male' ? 'ชาย' : item.gender === 'female' ? 'หญิง' : 'ไม่ระบุ';
      const pct = ((item.val / totalDemo) * 100).toFixed(1);
      summaryHtml += `<span><strong>อันดับที่ ${index + 1}:</strong> เพศ${genderTh} ช่วงอายุ ${item.age} ปี (สัดส่วน ${pct}% ของยอดรวม)</span>`;
    });
    summaryHtml += `</div>`;
    
    const demoSumEl = document.getElementById('demoSummary');
    if(demoSumEl) {
        demoSumEl.innerHTML = summaryHtml;
        demoSumEl.style.display = 'flex';
    }
  } else {
    const demoSumEl = document.getElementById('demoSummary');
    if(demoSumEl) demoSumEl.style.display = 'none';
  }

  // Ad Pie Chart - Spend by Ad
  const sortedAdSpend = Object.entries(adStats).sort((a,b)=>b[1].spend-a[1].spend);
  const adLabels = sortedAdSpend.map(item => item[0]);
  const adData = sortedAdSpend.map(item => item[1].spend);

  charts.adPie = new Chart(document.getElementById('adPieChart'), {
    type: 'doughnut',
    data: {
      labels: adLabels,
      datasets: [{
        data: adData,
        backgroundColor: adLabels.map((_, i) => COLORS[i % COLORS.length]),
        borderWidth: 2, borderColor: isDark ? '#1E293B' : '#FFF', hoverOffset: 8
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '50%',
      plugins: { legend: { position: 'right', labels: { color: textColor, font: { family: 'Inter,Noto Sans Thai', size: 11 } } } }
    }
  });
}

// ─── TABLE ─────────────────────────────────
let currentSortCol = 'roas';
let currentSortAsc = false;

window.handleSort = function(col) {
  if (currentSortCol === col) {
    currentSortAsc = !currentSortAsc;
  } else {
    currentSortCol = col;
    currentSortAsc = col === 'ad' ? true : false;
  }
  renderTable();
};

function renderTable() {
  let d = filteredData;
  const groupedObj = {};
  
  d.forEach(r => {
    if (!groupedObj[r.ad]) {
      groupedObj[r.ad] = {
        ad: r.ad, spend: 0, reach: 0, impressions: 0,
        clicks: 0, lpViews: 0, addCart: 0, purchases: 0, convValue: 0
      };
    }
    groupedObj[r.ad].spend += r.spend;
    groupedObj[r.ad].reach += r.reach;
    groupedObj[r.ad].impressions += r.impressions;
    groupedObj[r.ad].clicks += r.clicks;
    groupedObj[r.ad].lpViews += r.lpViews;
    groupedObj[r.ad].addCart += r.addCart;
    groupedObj[r.ad].purchases += r.purchases;
    groupedObj[r.ad].convValue += r.convValue;
  });

  // Derived metrics
  let tableData = Object.values(groupedObj).map(r => ({
    ...r,
    roas: r.spend > 0 ? r.convValue / r.spend : 0,
    cpa: r.purchases > 0 ? r.spend / r.purchases : 0
  }));

  // Sort
  tableData.sort((a, b) => {
    let valA = a[currentSortCol];
    let valB = b[currentSortCol];
    if (typeof valA === 'string') return currentSortAsc ? valA.localeCompare(valB) : valB.localeCompare(valA);
    return currentSortAsc ? valA - valB : valB - valA;
  });

  // Update Icons
  const cols = ['ad','spend','roas','purchases','cpa','convValue','clicks','lpViews','addCart','reach','impressions'];
  cols.forEach(c => {
    const el = document.getElementById(`sortIcon-${c}`);
    if (el) el.innerHTML = c === currentSortCol ? (currentSortAsc ? '▲' : '▼') : '';
  });

  document.getElementById('tableCount').textContent = `แสดง ${tableData.length} รายการ`;
  document.getElementById('tableBody').innerHTML = tableData.map(r => `
    <tr>
      <td title="${r.ad}" style="max-width:240px; overflow: hidden; text-overflow: ellipsis;">${r.ad}</td>
      <td class="cost-cell">฿${formatNum(r.spend)}</td>
      <td class="num" style="color:var(--emerald);font-weight:600;">${r.roas.toFixed(2)}x</td>
      <td class="num">${r.purchases.toLocaleString()}</td>
      <td class="num">฿${formatNum(r.cpa)}</td>
      <td class="num">฿${formatNum(r.convValue)}</td>
      <td class="num">${r.clicks.toLocaleString()}</td>
      <td class="num">${r.lpViews.toLocaleString()}</td>
      <td class="num">${r.addCart.toLocaleString()}</td>
      <td class="num">${r.reach.toLocaleString()}</td>
      <td class="num">${r.impressions.toLocaleString()}</td>
    </tr>
  `).join('');
}

// ─── FILTER & UPLOAD ───────────────────────
function populateFilters() {
  const campaigns = [...new Set(allData.map(r=>r.campaign).filter(c=>c))].sort();
  document.getElementById('filterCampaign').innerHTML = '<option value="">ทุก Campaign</option>'+campaigns.map(c=>`<option value="${c}">${c}</option>`).join('');
}

function applyFilters() {
  const campaign = document.getElementById('filterCampaign').value;
  filteredData = allData.filter(r=>{
    if (campaign && r.campaign !== campaign) return false;
    return true;
  });
  renderAll();
}

function clearFilters() {
  document.getElementById('filterCampaign').value = '';
  filteredData = [...allData];
  renderAll();
}

function handleFileUpload(e, platform) {
  const file = e.target.files[0]; 
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  
  uploadedFiles[platform] = file.name;
  if (platform === currentPlatform) {
    updateSourceLabel();
  }

  try {
    if (ext === 'xlsx' || ext === 'xls') {
      const reader = new FileReader();
      reader.onload = function(ev){ 
        try {
          const wb = XLSX.read(ev.target.result, {type:'array'}); 
          const ws = wb.Sheets[wb.SheetNames[0]]; 
          loadData(XLSX.utils.sheet_to_json(ws,{defval:0}), platform); 
        } catch(err) {
          showStatus('❌ ไม่สามารถอ่านไฟล์ Excel ได้ หรือไฟล์เสียหาย');
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      Papa.parse(file, {
        header:true, 
        skipEmptyLines:true, 
        complete: r => loadData(r.data, platform),
        error: err => showStatus('❌ เกิดข้อผิดพลาดในการอ่านไฟล์ CSV')
      });
    }
  } catch(err) {
    showStatus('❌ เกิดข้อผิดพลาดในระบบ: ' + err.message);
  }
}

document.getElementById('fbFileInput')?.addEventListener('change', (e) => handleFileUpload(e, 'facebook'));
document.getElementById('ggFileInput')?.addEventListener('change', (e) => handleFileUpload(e, 'google'));

function loadData(rows, platform, isInitial = false) {
  try {
    // Parser is generic enough to handle common FB/Google headers
    const parsedData = rows.map(r => ({
      campaign: r['Campaign name'] || r['Campaign'] || '',
      ad: (r['Ad name'] || r['Ad'] || r['Ad set name'] || r['Ad group'] || '').trim(),
      age: r['Age'] || 'Unknown',
      gender: r['Gender'] || 'Unknown',
      spend: parseFloat(r['Amount spent (THB)'] || r['Cost'] || 0) || 0,
      reach: parseInt(r['Reach'] || r['Unique users'] || 0) || 0,
      impressions: parseInt(r['Impressions'] || r['Impr.'] || 0) || 0,
      clicks: parseInt(r['Link clicks'] || r['Clicks'] || 0) || 0,
      lpViews: parseInt(r['Landing page views'] || 0) || 0,
      addCart: parseInt(r['Website adds to cart'] || 0) || 0,
      purchases: parseInt(r['Website purchases'] || r['Conversions'] || 0) || 0,
      convValue: parseFloat(r['Website purchases conversion value'] || r['Conv. value'] || 0) || 0
    })).filter(r => r.campaign || r.ad);
    
    platformData[platform] = parsedData;
    
    if (platform === currentPlatform) {
      allData = platformData[platform];
      filteredData = [...allData];
      populateFilters();
      renderAll();
    }
    
    if (!isInitial) {
      showStatus(`✅ โหลดข้อมูล ${platform.toUpperCase()} สำเร็จ: ${parsedData.length} แถว`);
      // Update UI timestamp if needed (optional)
    }
  } catch(e) { 
    if(!isInitial) showStatus('❌ เกิดข้อผิดพลาดในการอ่านไฟล์: '+e.message); 
  }
}

function showStatus(msg) {
  const bar = document.getElementById('statusBar');
  bar.textContent = msg; bar.style.display = 'block';
  setTimeout(()=>{ bar.style.display='none'; }, 4000);
}

// ─── UI & SETTINGS ─────────────────────────
const themeCheckbox = document.getElementById('themeToggleCheckbox');
function updateThemeIcon() {
  const isDark = document.body.classList.contains('dark-mode');
  const label = document.getElementById('themeSwitchText');
  if (label) {
    label.innerHTML = isDark
      ? '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg> โหมดมืด (Dark Mode)'
      : '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg> โหมดสว่าง (Light Mode)';
  }
  if (themeCheckbox) themeCheckbox.checked = isDark;
}
updateThemeIcon();

if (themeCheckbox) {
  themeCheckbox.addEventListener('change', (e) => {
    if (e.target.checked) document.body.classList.add('dark-mode');
    else document.body.classList.remove('dark-mode');
    updateThemeIcon();
    renderCharts();
  });
}

// Drawer
const settingsBtn = document.getElementById('settingsBtn');
const closeDrawerBtn = document.getElementById('closeDrawerBtn');
const settingsDrawer = document.getElementById('settingsDrawer');
const settingsOverlay = document.getElementById('settingsOverlay');
settingsBtn?.addEventListener('click', () => { settingsOverlay.classList.add('show'); settingsDrawer.classList.add('open'); });
const closeDrawer = () => { settingsOverlay.classList.remove('show'); settingsDrawer.classList.remove('open'); };
closeDrawerBtn?.addEventListener('click', closeDrawer);
settingsOverlay?.addEventListener('click', closeDrawer);
</script>
</body>
</html>
