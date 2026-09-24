import json
import base64
import os

with open("scratch/full_master_engine_data.json", "r", encoding="utf-8") as f:
    full_app_data = json.load(f)

app_data_json = json.dumps(full_app_data)

html_code = f'''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>AMFI Mutual Fund Rolling Return Analytics</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
  <script src="exceljs.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    body {{
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background-color: #f8fafc;
      color: #1e293b;
    }}
    .excel-header {{
      background-color: #1F4E79;
      color: #ffffff;
      font-weight: 700;
    }}
    .summary-row {{
      background-color: #EDF2F8;
      color: #1F4E79;
      font-weight: 600;
      border-bottom: 2px solid #1F4E79;
    }}
    .summary-row:hover {{
      background-color: #e2eaf4;
    }}
    .detail-row {{
      background-color: #ffffff;
    }}
    .detail-row-zebra {{
      background-color: #fafafa;
    }}
    .detail-row:hover {{
      background-color: #f1f5f9;
    }}
    .cagr-green {{
      color: #2E7D32;
      font-weight: 700;
    }}
    .recency-amber {{
      color: #B45309;
      font-weight: 700;
    }}
    .border-cell {{
      border: 1px solid #E2E8F0;
    }}
    .hidden {{
      display: none !important;
    }}
    
    .amfi-pill {{
      display: inline-flex;
      align-items: center;
      justify-content: space-between;
      min-width: 190px;
      padding: 0.65rem 1.4rem;
      background-color: #ffffff;
      border: 1.5px solid #d1d5db;
      border-radius: 9999px;
      font-size: 0.95rem;
      font-weight: 500;
      color: #334155;
      cursor: pointer;
      user-select: none;
      transition: all 0.15s ease;
      box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }}
    .amfi-pill:hover {{
      border-color: #94a3b8;
    }}
    .amfi-pill.active {{
      border-color: #00a896;
      box-shadow: 0 0 0 2px rgba(0, 168, 150, 0.15);
    }}
    
    .amfi-dropdown-menu {{
      position: absolute;
      top: calc(100% + 6px);
      left: 0;
      width: 100%;
      min-width: 250px;
      background-color: #ffffff;
      border: 1px solid #cbd5e1;
      border-radius: 14px;
      box-shadow: 0 12px 28px rgba(0,0,0,0.12);
      z-index: 50;
      padding: 8px;
    }}
    .amfi-dropdown-menu::-webkit-scrollbar {{
      width: 6px;
    }}
    .amfi-dropdown-menu::-webkit-scrollbar-track {{
      background: #f1f5f9;
      border-radius: 4px;
    }}
    .amfi-dropdown-menu::-webkit-scrollbar-thumb {{
      background: #cbd5e1;
      border-radius: 4px;
    }}
    .amfi-dropdown-menu::-webkit-scrollbar-thumb:hover {{
      background: #94a3b8;
    }}
    .amfi-dropdown-item {{
      padding: 9px 14px;
      font-size: 0.88rem;
      font-weight: 500;
      color: #334155;
      border-radius: 8px;
      cursor: pointer;
      transition: background-color 0.12s ease;
    }}
    .amfi-dropdown-item:hover {{
      background-color: #f1f5f9;
      color: #0f172a;
    }}
    .amfi-dropdown-item.selected {{
      background-color: #e6f7f5;
      color: #00897b;
      font-weight: 600;
    }}
    
    .toggle-btn {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 22px;
      height: 22px;
      border-radius: 4px;
      background-color: #1F4E79;
      color: #ffffff;
      font-weight: bold;
      font-size: 13px;
      cursor: pointer;
      user-select: none;
    }}
  </style>
</head>
<body class="p-4 sm:p-6 lg:p-8">
  <div class="max-w-[98%] mx-auto space-y-6">

    <!-- Top Navigation / Title Bar with Excel Download -->
    <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
      <div class="flex items-center gap-3">
        <div class="w-10 h-10 rounded-xl bg-[#1F4E79] flex items-center justify-center text-white font-black text-xl shadow-sm">
          AMFI
        </div>
        <div>
          <h1 class="text-xl sm:text-2xl font-bold text-slate-900 tracking-tight" id="mainHeaderTitle">Mutual Fund Rolling Return Analysis</h1>
          <p class="text-xs sm:text-sm text-slate-500 font-medium" id="mainHeaderSubtitle">10-Year Historical Performance Engine (Jan 2015 - Sep 2026) | Statutory AMFI Feed</p>
        </div>
      </div>

      <!-- Download Excel Button -->
      <div>
        <button onclick="downloadExcelFile()" class="inline-flex items-center gap-2.5 px-6 py-2.5 bg-emerald-700 hover:bg-emerald-800 text-white font-semibold text-sm rounded-full shadow-sm hover:shadow transition-all duration-200 cursor-pointer">
          <svg class="w-5 h-5" fill="currentColor" viewBox="0 0 24 24">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zM6 20V4h7v5h5v11H6zm4.5-3.5l1.8-3.1 1.7 3.1h1.9l-2.7-4.4 2.5-4.1h-1.9l-1.5 2.8-1.5-2.8H8.8l2.5 4.1-2.7 4.4h1.9z"/>
          </svg>
          <span id="excelBtnLabel">Download Excel (.xlsx)</span>
        </button>
      </div>
    </div>

    <!-- STATUTORY PILL FILTERS (Covering 100% of AMFI Master Data columns: Type, Group, Subcategory, Option) -->
    <div class="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm">
      <div class="flex flex-wrap items-center gap-4">
        
        <!-- Filter 1: Scheme Type -->
        <div class="relative" id="dropdownContainer1">
          <div class="amfi-pill" id="pillSchemeType" onclick="toggleDropdown('dropdown1', event)">
            <span id="labelSchemeType">Open Ended</span>
            <svg class="w-4 h-4 text-[#00a896] transition-transform duration-200 ml-4" id="chevron1" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" d="M19 9l-7 7-7-7"></path>
            </svg>
          </div>
          <div class="amfi-dropdown-menu hidden" id="dropdown1">
            <div class="amfi-dropdown-item selected" onclick="selectSchemeType('Open Ended', event)">Open Ended</div>
            <div class="amfi-dropdown-item" onclick="selectSchemeType('Close Ended', event)">Close Ended</div>
            <div class="amfi-dropdown-item" onclick="selectSchemeType('Interval Fund', event)">Interval Fund</div>
          </div>
        </div>

        <!-- Filter 2: Category Group -->
        <div class="relative" id="dropdownContainer2">
          <div class="amfi-pill" id="pillCategoryGroup" onclick="toggleDropdown('dropdown2', event)">
            <span id="labelCategoryGroup">Equity</span>
            <svg class="w-4 h-4 text-[#00a896] transition-transform duration-200 ml-4" id="chevron2" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" d="M19 9l-7 7-7-7"></path>
            </svg>
          </div>
          <div class="amfi-dropdown-menu hidden" id="dropdown2">
            <!-- Dynamically populated based on Scheme Type -->
          </div>
        </div>

        <!-- Filter 3: Sub Category -->
        <div class="relative" id="dropdownContainer3">
          <div class="amfi-pill" id="pillSubCategory" onclick="toggleDropdown('dropdown3', event)">
            <span id="labelSubCategory">Flexi Cap Fund</span>
            <svg class="w-4 h-4 text-[#00a896] transition-transform duration-200 ml-4" id="chevron3" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" d="M19 9l-7 7-7-7"></path>
            </svg>
          </div>
          <div class="amfi-dropdown-menu hidden max-h-72 overflow-y-auto" id="dropdown3">
            <!-- Dynamically populated via JS covering 100% of subcategories -->
          </div>
        </div>

        <!-- Filter 4: Scheme Option (Growth, IDCW, Bonus, Other, All Options) -->
        <div class="relative" id="dropdownContainer4">
          <div class="amfi-pill" id="pillSchemeOption" onclick="toggleDropdown('dropdown4', event)">
            <span id="labelSchemeOption">Growth</span>
            <svg class="w-4 h-4 text-[#00a896] transition-transform duration-200 ml-4" id="chevron4" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" d="M19 9l-7 7-7-7"></path>
            </svg>
          </div>
          <div class="amfi-dropdown-menu hidden" id="dropdown4">
            <div class="amfi-dropdown-item selected" onclick="selectSchemeOption('Growth', event)">Growth</div>
            <div class="amfi-dropdown-item" onclick="selectSchemeOption('IDCW', event)">IDCW (Dividend)</div>
            <div class="amfi-dropdown-item" onclick="selectSchemeOption('Bonus', event)">Bonus</div>
            <div class="amfi-dropdown-item" onclick="selectSchemeOption('Other', event)">Other</div>
            <div class="amfi-dropdown-item" onclick="selectSchemeOption('All Options', event)">All Options</div>
          </div>
        </div>

        <!-- Filter 5: Plan Type (Regular, Direct, All Plans) -->
        <div class="relative" id="dropdownContainer5">
          <div class="amfi-pill" id="pillSchemePlan" onclick="toggleDropdown('dropdown5', event)">
            <span id="labelSchemePlan">Regular</span>
            <svg class="w-4 h-4 text-[#00a896] transition-transform duration-200 ml-4" id="chevron5" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" d="M19 9l-7 7-7-7"></path>
            </svg>
          </div>
          <div class="amfi-dropdown-menu hidden" id="dropdown5">
            <div class="amfi-dropdown-item selected" onclick="selectSchemePlan('Regular', event)">Regular</div>
            <div class="amfi-dropdown-item" onclick="selectSchemePlan('Direct', event)">Direct</div>
            <div class="amfi-dropdown-item" onclick="selectSchemePlan('All Plans', event)">All Plans</div>
          </div>
        </div>
        <!-- Search Box & Expand All button (matching screenshot) -->
        <div class="ml-auto flex items-center gap-3">
          <div class="relative min-w-[240px]">
            <input type="text" id="fundSearchInput" oninput="filterFundNames()" placeholder="Search fund name..." class="w-full text-xs sm:text-sm pl-9 pr-4 py-2 border border-slate-300 rounded-full focus:outline-none focus:ring-2 focus:ring-[#00a896]/20 focus:border-[#00a896] transition">
            <svg class="w-4 h-4 text-slate-400 absolute left-3 top-2.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path>
            </svg>
          </div>
          <button onclick="toggleAllRows()" id="btnExpandAllPill" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold text-xs rounded-full border border-slate-300 transition cursor-pointer">
            Expand All [+]
          </button>
        </div>

      </div>
    </div>

    <!-- MAIN VIEW CONTAINER -->
    <div id="mainViewContainer">
      <!-- Open-Ended Rolling Return Tables -->
      <div class="space-y-8" id="openEndedTables">
        <div class="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
          <div class="p-6 border-b border-slate-200 bg-slate-50/50 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h2 id="table1Title" class="text-lg sm:text-xl font-bold text-[#1F4E79] tracking-tight">Flexi Cap Mutual Funds (Growth) — 3-Year Rolling Return Analysis (Regular Plans)</h2>
              <p id="table1Subtitle" class="text-xs sm:text-sm text-slate-500 font-medium italic mt-1">Half-Yearly Rolling Windows (Jan 2015 to Sep 2026) | Click [+] on any row to expand Top 25% funds</p>
            </div>
            <button onclick="toggleAllRows()" id="btnExpandAll" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold text-xs rounded-full border border-slate-300 transition cursor-pointer self-start sm:self-auto">
              Expand All [+]
            </button>
          </div>

          <div class="overflow-x-auto">
            <table class="w-full text-left text-xs sm:text-sm border-collapse" id="rollingTable">
              <thead>
                <tr class="excel-header text-center">
                  <th class="py-3 px-3 w-12 border-cell">Expand</th>
                  <th class="py-3 px-4 border-cell">Start Date</th>
                  <th class="py-3 px-4 border-cell">End Date</th>
                  <th class="py-3 px-3 border-cell">Quarter</th>
                  <th class="py-3 px-4 border-cell">Total Funds (n)</th>
                  <th class="py-3 px-4 border-cell">Top 25% Count</th>
                  <th class="py-3 px-6 text-left border-cell">Fund Name</th>
                  <th class="py-3 px-4 text-right border-cell">3Y CAGR (%)</th>
                </tr>
              </thead>
              <tbody id="rollingTableBody">
                <!-- Dynamically populated via JS -->
              </tbody>
            </table>
          </div>
        </div>

        <!-- TABLE 2: Consistency Summary with Recency & Composite Ranking (Filtered > 75% Windows) -->
        <div class="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
          <div class="p-6 border-b border-slate-200 bg-slate-50/50 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h2 id="table2Title" class="text-lg sm:text-xl font-bold text-[#1F4E79] tracking-tight">Flexi Cap Funds (Growth) — Top 25% Consistency Summary (Grouped by Name)</h2>
              <p id="table2Subtitle" class="text-xs sm:text-sm text-slate-500 font-medium italic mt-1">Ranked by Composite Score (65% Consistency + 35% Recency) | Top 25% Funds with ≥ 75% Windows | Click [+] to expand details</p>
            </div>
            <div class="flex items-center gap-3">
              <button onclick="toggleAllTable2Rows()" id="btnExpandAllT2" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold text-xs rounded-full border border-slate-300 transition cursor-pointer self-start sm:self-auto">
                Expand All [+]
              </button>
              <div class="text-xs font-semibold text-slate-500 bg-slate-100 px-3 py-1.5 rounded-full border border-slate-200">
                Click headers to sort
              </div>
            </div>
          </div>

          <div class="overflow-x-auto">
            <table class="w-full text-left text-xs sm:text-sm border-collapse" id="consistencyTable">
              <thead>
                <tr class="excel-header text-center cursor-pointer select-none">
                  <th class="py-3 px-3 w-12 border-cell">Expand</th>
                  <th onclick="sortTable2(0)" class="py-2.5 px-3 w-14 border-cell hover:bg-[#183e60] transition">Rank ⇅</th>
                  <th onclick="sortTable2(1)" class="py-2.5 px-4 border-cell hover:bg-[#183e60] transition min-w-[130px]" title="Composite Score = 65% Consistency Score + 35% Recency Score">Composite Score ⇅</th>
                  <th onclick="sortTable2(2)" class="py-2.5 px-5 border-cell hover:bg-[#183e60] transition min-w-[130px]">Consistency Score ⇅</th>
                  <th onclick="sortTable2(3)" class="py-2.5 px-4 text-right border-cell hover:bg-[#183e60] transition min-w-[95px]">Recency Score ⇅</th>
                  <th onclick="sortTable2(4)" class="py-2.5 px-6 text-left border-cell hover:bg-[#183e60] transition min-w-[280px]">Fund Name ⇅</th>
                  <th onclick="sortTable2(5)" class="py-2.5 px-4 text-right border-cell hover:bg-[#183e60] transition min-w-[105px]">Avg 3Y CAGR (%) ⇅</th>
                </tr>
              </thead>
              <tbody id="consistencyTableBody">
                <!-- Dynamically populated via JS -->
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- Close-Ended / Interval Table View -->
      <div id="closeEndedView" class="hidden bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden p-6 space-y-4">
        <div>
          <h2 id="closeEndedTitle" class="text-lg sm:text-xl font-bold text-[#1F4E79] tracking-tight">Close Ended Schemes</h2>
          <p class="text-xs sm:text-sm text-slate-500 font-medium italic mt-1">Statutory AMFI Disclosure: Schemes from master database matching your filters:</p>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-left text-xs sm:text-sm border-collapse" id="closeEndedTable">
            <thead>
              <tr class="excel-header text-center">
                <th class="py-2.5 px-4 border-cell w-24">Scheme Code</th>
                <th class="py-2.5 px-6 border-cell text-left">Scheme Name</th>
                <th class="py-2.5 px-6 border-cell text-left">AMC</th>
                <th class="py-2.5 px-4 border-cell text-center">Option</th>
                <th class="py-2.5 px-6 border-cell text-left">Category / Series</th>
                <th class="py-2.5 px-4 border-cell text-right">Latest NAV (₹)</th>
              </tr>
            </thead>
            <tbody id="closeEndedTableBody">
              <!-- Dynamically populated via JS -->
            </tbody>
          </table>
        </div>
      </div>
    </div>

  </div>

  <script>
    const appData = {app_data_json};

    let currentSchemeType = "Open Ended";
    let currentCategoryGroup = "Equity";
    let currentSubCategory = "Flexi Cap Fund";
    let currentSchemeOption = "Growth";
    let currentSchemePlan = "Regular";
    let isAllExpanded = false;
    let expandedMap = {{}};

    const categoryHierarchy = appData.hierarchy;
    const windowDates = appData.window_dates;
    const last6Labels = appData.last_6_labels;

    function initPage() {{
      updateCategoryGroupDropdown();
      updateSubCategoryDropdown();
      renderActiveView();
    }}

    function toggleDropdown(id, event) {{
      event.stopPropagation();
      const allMenus = document.querySelectorAll('.amfi-dropdown-menu');
      const allChevrons = document.querySelectorAll('.amfi-pill svg');
      const allPills = document.querySelectorAll('.amfi-pill');

      const target = document.getElementById(id);
      const isCurrentlyOpen = !target.classList.contains('hidden');

      allMenus.forEach(m => m.classList.add('hidden'));
      allChevrons.forEach(c => {{
        c.style.transform = "rotate(0deg)";
        c.classList.remove("text-slate-600");
        c.classList.add("text-[#00a896]");
      }});
      allPills.forEach(p => p.classList.remove('active'));

      if (!isCurrentlyOpen) {{
        target.classList.remove('hidden');
        const pill = target.previousElementSibling;
        const chevron = pill.querySelector('svg');
        pill.classList.add('active');
        chevron.style.transform = "rotate(180deg)";
        chevron.classList.remove("text-[#00a896]");
        chevron.classList.add("text-slate-600");
      }}
    }}

    document.addEventListener('click', () => {{
      document.querySelectorAll('.amfi-dropdown-menu').forEach(m => m.classList.add('hidden'));
      document.querySelectorAll('.amfi-pill svg').forEach(c => {{
        c.style.transform = "rotate(0deg)";
        c.classList.remove("text-slate-600");
        c.classList.add("text-[#00a896]");
      }});
      document.querySelectorAll('.amfi-pill').forEach(p => p.classList.remove('active'));
    }});

    function selectSchemeType(type, e) {{
      if (e) e.stopPropagation();
      currentSchemeType = type;
      document.getElementById('labelSchemeType').textContent = type;
      document.querySelectorAll('#dropdown1 .amfi-dropdown-item').forEach(el => {{
        el.classList.toggle('selected', el.textContent.trim() === type);
      }});

      document.getElementById('dropdown1').classList.add('hidden');
      document.getElementById('pillSchemeType').classList.remove('active');
      document.getElementById('chevron1').style.transform = "rotate(0deg)";

      updateCategoryGroupDropdown();
      updateSubCategoryDropdown();
      renderActiveView();
    }}

    function selectCategoryGroup(group, e) {{
      if (e) e.stopPropagation();
      currentCategoryGroup = group;
      document.getElementById('labelCategoryGroup').textContent = group;
      document.querySelectorAll('#dropdown2 .amfi-dropdown-item').forEach(el => {{
        el.classList.toggle('selected', el.textContent.trim() === group);
      }});

      document.getElementById('dropdown2').classList.add('hidden');
      document.getElementById('pillCategoryGroup').classList.remove('active');
      document.getElementById('chevron2').style.transform = "rotate(0deg)";

      updateSubCategoryDropdown();
      renderActiveView();
    }}

    function selectSubCategory(sub, e) {{
      if (e) e.stopPropagation();
      currentSubCategory = sub;
      document.getElementById('labelSubCategory').textContent = sub;
      document.querySelectorAll('#dropdown3 .amfi-dropdown-item').forEach(el => {{
        el.classList.toggle('selected', el.textContent.trim() === sub);
      }});

      document.getElementById('dropdown3').classList.add('hidden');
      document.getElementById('pillSubCategory').classList.remove('active');
      document.getElementById('chevron3').style.transform = "rotate(0deg)";

      renderActiveView();
    }}

    function selectSchemeOption(opt, e) {{
      if (e) e.stopPropagation();
      currentSchemeOption = opt;
      document.getElementById('labelSchemeOption').textContent = opt === 'IDCW' ? 'IDCW (Dividend)' : opt;
      document.querySelectorAll('#dropdown4 .amfi-dropdown-item').forEach(el => {{
        el.classList.toggle('selected', el.textContent.startsWith(opt));
      }});

      document.getElementById('dropdown4').classList.add('hidden');
      document.getElementById('pillSchemeOption').classList.remove('active');
      document.getElementById('chevron4').style.transform = "rotate(0deg)";

      renderActiveView();
    }}

    function selectSchemePlan(plan, e) {{
      if (e) e.stopPropagation();
      currentSchemePlan = plan;
      document.getElementById('labelSchemePlan').textContent = plan;
      document.querySelectorAll('#dropdown5 .amfi-dropdown-item').forEach(el => {{
        el.classList.toggle('selected', el.textContent.trim() === plan);
      }});

      document.getElementById('dropdown5').classList.add('hidden');
      document.getElementById('pillSchemePlan').classList.remove('active');
      document.getElementById('chevron5').style.transform = "rotate(0deg)";

      renderActiveView();
    }}

    function updateCategoryGroupDropdown() {{
      const menu = document.getElementById('dropdown2');
      menu.innerHTML = "";
      const grps = Object.keys(categoryHierarchy[currentSchemeType] || {{ "Equity": [] }});

      if (!grps.includes(currentCategoryGroup)) {{
        currentCategoryGroup = grps[0];
      }}
      document.getElementById('labelCategoryGroup').textContent = currentCategoryGroup;

      grps.forEach(g => {{
        const item = document.createElement('div');
        item.className = `amfi-dropdown-item ${{g === currentCategoryGroup ? 'selected' : ''}}`;
        item.textContent = g;
        item.onclick = (e) => selectCategoryGroup(g, e);
        menu.appendChild(item);
      }});
    }}

    function updateSubCategoryDropdown() {{
      const menu = document.getElementById('dropdown3');
      menu.innerHTML = "";
      const subs = (categoryHierarchy[currentSchemeType] && categoryHierarchy[currentSchemeType][currentCategoryGroup]) || ["Flexi Cap Fund"];

      if (!subs.includes(currentSubCategory)) {{
        currentSubCategory = subs[0];
      }}
      document.getElementById('labelSubCategory').textContent = currentSubCategory;

      subs.forEach(s => {{
        const item = document.createElement('div');
        item.className = `amfi-dropdown-item ${{s === currentSubCategory ? 'selected' : ''}}`;
        item.textContent = s;
        item.onclick = (e) => selectSubCategory(s, e);
        menu.appendChild(item);
      }});
    }}

    function getQuartileIcon(quartile) {{
      const topFill = quartile === 1 ? '<rect x="1.5" y="1.5" width="13" height="4" fill="#475569" stroke="none" />' : '';
      const q2Fill = quartile === 2 ? '<rect x="1.5" y="5.5" width="13" height="4.5" fill="#475569" stroke="none" />' : '';
      const q3Fill = quartile === 3 ? '<rect x="1.5" y="10" width="13" height="4.5" fill="#475569" stroke="none" />' : '';
      const botFill = quartile === 4 ? '<rect x="1.5" y="14.5" width="13" height="4" fill="#475569" stroke="none" />' : '';

      return `
        <span class="inline-flex items-center justify-center" title="Quartile ${{quartile}}">
          <svg class="w-3.5 h-4.5" viewBox="0 0 16 20" fill="none" stroke="#94a3b8" stroke-width="1.2">
            <rect x="1" y="1" width="14" height="18" rx="1.5" />
            <line x1="1" y1="5.5" x2="15" y2="5.5" />
            <line x1="1" y1="10" x2="15" y2="10" />
            <line x1="1" y1="14.5" x2="15" y2="14.5" />
            ${{topFill}}
            ${{q2Fill}}
            ${{q3Fill}}
            ${{botFill}}
          </svg>
        </span>
      `;
    }}

    function renderActiveView() {{
      const openView = document.getElementById("openEndedTables");
      const closeView = document.getElementById("closeEndedView");

      if (currentSchemeType === "Open Ended") {{
        openView.classList.remove("hidden");
        closeView.classList.add("hidden");
        renderOpenEndedData();
      }} else {{
        openView.classList.add("hidden");
        closeView.classList.remove("hidden");
        renderCloseEndedData();
      }}
    }}

    function renderCloseEndedData() {{
      const titleEl = document.getElementById("closeEndedTitle");
      const tbody = document.getElementById("closeEndedTableBody");
      tbody.innerHTML = "";

      const listMap = currentSchemeType === "Close Ended" ? appData.close_ended : appData.interval;
      const rawList = listMap[currentSubCategory] || Object.values(listMap)[0] || [];
      
      const list = rawList.filter(item => {{
        const matchOption = (currentSchemeOption === "All Options") || (item.option === currentSchemeOption);
        const matchPlan = (currentSchemePlan === "All Plans") || (item.plan === currentSchemePlan);
        return matchOption && matchPlan;
      }});

      const planLabel = currentSchemePlan === "All Plans" ? "Regular & Direct" : currentSchemePlan;
      titleEl.textContent = `${{currentSchemeType}} Schemes (${{currentSchemeOption}}, ${{planLabel}}) — ${{currentCategoryGroup}} / ${{currentSubCategory}}`;

      if (list.length === 0) {{
        tbody.innerHTML = `<tr><td colspan="6" class="py-8 text-center text-slate-400 italic">No active schemes found for plan '${{currentSchemePlan}}' and option '${{currentSchemeOption}}'. Try choosing 'All Plans' or 'All Options'.</td></tr>`;
        return;
      }}

      list.forEach((item, idx) => {{
        const tr = document.createElement("tr");
        tr.className = `detail-row ${{idx % 2 === 0 ? "detail-row-zebra" : ""}} hover:bg-blue-50/50`;
        tr.innerHTML = `
          <td class="py-2.5 px-4 font-mono text-center text-slate-600 border-cell">${{item.code}}</td>
          <td class="py-2.5 px-6 font-medium text-slate-900 border-cell text-left">${{item.name}}</td>
          <td class="py-2.5 px-6 text-slate-700 border-cell text-left">${{item.amc}}</td>
          <td class="py-2.5 px-4 text-center font-semibold text-xs border-cell text-slate-600">${{item.option}}</td>
          <td class="py-2.5 px-6 text-slate-600 border-cell text-left">${{item.category}}</td>
          <td class="py-2.5 px-4 text-right font-bold cagr-green border-cell">₹${{typeof item.nav === 'number' ? item.nav.toFixed(4) : item.nav}}</td>
        `;
        tbody.appendChild(tr);
      }});
    }}

    // In-memory active dataset for sorting & Excel download
    let activeComputedDataset = null;

    function renderOpenEndedData() {{
      const rawFunds = appData.open_ended[currentSubCategory] || 
                        appData.open_ended[currentSubCategory + " Fund"] || 
                        appData.open_ended[currentSubCategory.replace(" Fund", "")] || 
                        appData.open_ended["Flexi Cap Fund"] ||
                        appData.open_ended["Flexi Cap"] || [];

      // Filter by currentSchemeOption AND currentSchemePlan
      const funds = rawFunds.filter(f => {{
        const matchOption = (currentSchemeOption === "All Options") || (f.option === currentSchemeOption);
        const matchPlan = (currentSchemePlan === "All Plans") || (f.plan === currentSchemePlan);
        return matchOption && matchPlan;
      }});

      const optLabel = currentSchemeOption === "All Options" ? "All Options" : currentSchemeOption;
      const planLabel = currentSchemePlan === "All Plans" ? "Regular & Direct Plans" : `${{currentSchemePlan}} Plans`;
      document.getElementById("table1Title").textContent = `${{currentSubCategory}} Mutual Funds (${{optLabel}}) — 3-Year Rolling Return Analysis (${{planLabel}})`;
      document.getElementById("table1Subtitle").textContent = `Half-Yearly Rolling Windows (Jan 2015 to Sep 2026) | Click [+] on any row to expand Top 25% funds`;
      document.getElementById("table2Title").textContent = `${{currentSubCategory}} Funds (${{optLabel}}, ${{planLabel}}) — Top 25% Consistency Summary (Grouped by Name)`;
      document.getElementById("table2Subtitle").textContent = `Ranked by Composite Score (65% Consistency + 35% Recency) | Minimum 75% Eligible Windows (>13.5 of 18) Required`;

      const tbody1 = document.getElementById("rollingTableBody");
      const tbody2 = document.getElementById("consistencyTableBody");
      tbody1.innerHTML = "";
      tbody2.innerHTML = "";
      expandedMap = {{}};
      isAllExpanded = false;
      document.getElementById("btnExpandAll").textContent = "Expand All [+]";

      if (funds.length === 0) {{
        tbody1.innerHTML = `<tr><td colspan="8" class="py-10 text-center text-slate-400 italic">No funds found for plan '${{currentSchemePlan}}' and option '${{currentSchemeOption}}' in ${{currentSubCategory}}. Please select 'Regular', 'Direct', or 'All Plans'.</td></tr>`;
        tbody2.innerHTML = `<tr><td colspan="19" class="py-10 text-center text-slate-400 italic">No funds found for plan '${{currentSchemePlan}}' and option '${{currentSchemeOption}}' in ${{currentSubCategory}}. Please select 'Regular', 'Direct', or 'All Plans'.</td></tr>`;
        activeComputedDataset = null;
        return;
      }}

      // Compute rolling windows dynamically for filtered funds
      const computedWindows = [];
      windowDates.forEach((w, wIdx) => {{
        const winFunds = funds
          .filter(f => f.cagrs && f.cagrs[wIdx] !== null && f.cagrs[wIdx] !== undefined)
          .map(f => ({{
            schemeCode: f.code,
            schemeName: f.name,
            plan: f.plan,
            option: f.option,
            cagr: f.cagrs[wIdx]
          }}))
          .sort((a, b) => b.cagr - a.cagr);

        const n = winFunds.length;
        const top_25_pct = Math.max(1, Math.round(n * 0.25));
        const top_funds = winFunds.slice(0, top_25_pct);

        computedWindows.push({{
          start_date: w[0],
          end_date: w[1],
          quarter: w[2],
          n: n,
          top_25_pct: top_25_pct,
          win_funds: winFunds,
          top_funds: top_funds
        }});
      }});

      // Render Table 1
      computedWindows.forEach((win, winIdx) => {{
        const topFunds = win.top_funds || [];
        const bestFund = topFunds[0] || null;
        const bestName = bestFund ? `#1: ${{bestFund.schemeName}}` : "N/A";
        const bestCagr = (bestFund && bestFund.cagr !== null && bestFund.cagr !== undefined) ? `${{bestFund.cagr.toFixed(2)}}%` : "—";

        const summaryTr = document.createElement("tr");
        summaryTr.className = "summary-row cursor-pointer";
        summaryTr.onclick = () => toggleRow(winIdx);

        summaryTr.innerHTML = `
          <td class="py-2.5 px-3 text-center border-cell">
            <span class="toggle-btn" id="toggle-btn-${{winIdx}}">+</span>
          </td>
          <td class="py-2.5 px-4 text-center font-bold border-cell">${{win.start_date}}</td>
          <td class="py-2.5 px-4 text-center font-bold border-cell">${{win.end_date}}</td>
          <td class="py-2.5 px-3 text-center font-bold border-cell">${{win.quarter}}</td>
          <td class="py-2.5 px-4 text-center font-bold border-cell">${{win.n}}</td>
          <td class="py-2.5 px-4 text-center font-bold border-cell">${{win.top_25_pct}}</td>
          <td class="py-2.5 px-6 font-bold border-cell text-left">${{bestName}}</td>
          <td class="py-2.5 px-4 text-right font-bold border-cell">${{bestCagr}}</td>
        `;
        tbody1.appendChild(summaryTr);

        topFunds.slice(1).forEach((fund, rankIdx) => {{
          const rank = rankIdx + 2;
          const childTr = document.createElement("tr");
          childTr.className = `detail-row row-child-${{winIdx}} hidden ${{rank % 2 === 0 ? "detail-row-zebra" : ""}}`;
          const fundCagrStr = (fund && fund.cagr !== null && fund.cagr !== undefined) ? `${{fund.cagr.toFixed(2)}}%` : "—";

          childTr.innerHTML = `
            <td class="py-2 px-3 text-center text-slate-400 italic font-semibold border-cell">#${{rank}}</td>
            <td class="py-2 px-4 border-cell"></td>
            <td class="py-2 px-4 border-cell"></td>
            <td class="py-2 px-3 border-cell"></td>
            <td class="py-2 px-4 border-cell"></td>
            <td class="py-2 px-4 border-cell"></td>
            <td class="py-2 px-6 text-slate-700 font-medium border-cell text-left pl-8">${{fund.schemeName}}</td>
            <td class="py-2 px-4 text-right cagr-green border-cell">${{fundCagrStr}}</td>
          `;
          tbody1.appendChild(childTr);
        }});
      }});

      // Compute Table 2 Consistency Summary
      // Threshold: minimum eligible windows rounded to nearest integer (at least 75% of maximum windows)
      const maxEligible = windowDates.length; // 18
      const minRequired = Math.round(maxEligible * 0.75); // 14
      const top25DisplayLimit = Math.max(1, Math.round(funds.length * 0.25)); // Top 25% of funds in selected category
      const consistencyList = [];
      const last6Windows = computedWindows.slice(-6);

      funds.forEach(f => {{
        let eligible = 0;
        f.cagrs.forEach(c => {{
          if (c !== null && c !== undefined) eligible++;
        }});

        // Filter: only those with at least 75% of windows (rounded)
        if (eligible < minRequired) return;

        let topCount = 0;
        let sumCagr = 0;

        computedWindows.forEach(win => {{
          const c = f.cagrs[computedWindows.indexOf(win)];
          if (c !== null && c !== undefined) sumCagr += c;
          if (win.top_funds.some(x => x.schemeName === f.name)) {{
            topCount++;
          }}
        }});

        // Exclude funds with zero appearances in Top 25%
        if (topCount === 0) return;

        const consistencyRaw = (topCount / eligible) * 100.0;
        const consistencyScore = Math.round(consistencyRaw);
        const avgCagr = eligible > 0 ? (sumCagr / eligible) : 0;

        // Recency across last 6 windows
        let recTopCount = 0;
        const recDetails = [];

        last6Windows.forEach(lw => {{
          const sorted = lw.win_funds;
          const N = sorted.length;
          const q1 = Math.max(1, Math.round(N * 0.25));
          const q2 = Math.max(q1 + 1, Math.round(N * 0.50));
          const q3 = Math.max(q2 + 1, Math.round(N * 0.75));

          const matchIdx = sorted.findIndex(x => x.schemeName === f.name);
          if (matchIdx === -1) {{
            recDetails.push({{ cagr: null, quartile: 4 }});
          }} else {{
            const rankPos = matchIdx + 1;
            let qVal = 4;
            if (rankPos <= q1) {{
              qVal = 1;
              recTopCount++;
            }} else if (rankPos <= q2) {{
              qVal = 2;
            }} else if (rankPos <= q3) {{
              qVal = 3;
            }}
            recDetails.push({{ cagr: sorted[matchIdx].cagr, quartile: qVal }});
          }}
        }});

        const recencyRaw = (recTopCount / 6.0) * 100.0;
        const recencyScore = Math.round(recencyRaw);

        // Composite Score = 65% consistency score + 35% recency score
        const compositeScore = Math.round(0.65 * consistencyScore + 0.35 * recencyScore);

        // Exclude any fund with composite score of 0
        if (compositeScore === 0) return;

        consistencyList.push({{
          code: f.code,
          name: f.name,
          plan: f.plan,
          option: f.option,
          composite_score: compositeScore,
          consistency_score: consistencyScore,
          recency_score: recencyScore,
          eligible: eligible,
          avg_cagr: avgCagr,
          nav: f.nav,
          aum: f.aum,
          expense_ratio: f.expense_ratio,
          fund_manager: f.fund_manager,
          range: "Jan 2015 to Jul 2026",
          last_6_details: recDetails
        }});
      }});

      // Ranking based on Composite Score descending, then consistency, then recency, then avg cagr
      consistencyList.sort((a, b) => {{
        if (b.composite_score !== a.composite_score) return b.composite_score - a.composite_score;
        if (b.consistency_score !== a.consistency_score) return b.consistency_score - a.consistency_score;
        if (b.recency_score !== a.recency_score) return b.recency_score - a.recency_score;
        return b.avg_cagr - a.avg_cagr;
      }});

      // Display strictly the Top 25% of funds for this category
      const displayList = consistencyList.slice(0, top25DisplayLimit);

      document.getElementById("table2Subtitle").textContent = `Ranked by Composite Score (65% Consistency + 35% Recency) | Top 25% Funds (${{displayList.length}} of ${{funds.length}}) with ≥ 75% Windows (${{minRequired}} of ${{maxEligible}}) | Click [+] to expand details`;

      table2DisplayList = displayList;
      t2ExpandedMap = {{}};
      isAllT2Expanded = false;
      const btnExpT2 = document.getElementById("btnExpandAllT2");
      if (btnExpT2) btnExpT2.textContent = "Expand All [+]";

      activeComputedDataset = {{
        windows: computedWindows,
        consistency: displayList,
        subCategory: currentSubCategory,
        option: currentSchemeOption,
        plan: currentSchemePlan
      }};

      renderTable2Body();
    }}

    let table2DisplayList = [];
    let t2ExpandedMap = {{}};
    let isAllT2Expanded = false;
    let t2SortCol = 1;
    let t2SortAsc = false;

    function renderTable2Body() {{
      const tbody2 = document.getElementById("consistencyTableBody");
      tbody2.innerHTML = "";

      if (table2DisplayList.length === 0) {{
        tbody2.innerHTML = `<tr><td colspan="7" class="py-10 text-center text-slate-400 italic">No funds meet the minimum 75% window eligibility with positive scores in this category.</td></tr>`;
        return;
      }}

      const last6 = activeComputedDataset.windows.slice(-6);
      const endMonthYears = last6.map(w => w.end_date);

      table2DisplayList.forEach((item, idx) => {{
        const isExp = !!t2ExpandedMap[idx];
        const tr = document.createElement("tr");
        tr.className = `detail-row ${{idx % 2 === 0 ? "detail-row-zebra" : ""}} hover:bg-blue-50/50 cursor-pointer`;
        tr.onclick = () => toggleTable2Row(idx);

        const optionBadge = currentSchemeOption === "All Options" ? `<span class="ml-2 text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-600 font-mono">${{item.option}}</span>` : '';
        const planBadge = currentSchemePlan === "All Plans" ? `<span class="ml-1.5 text-[10px] px-1.5 py-0.5 rounded ${{item.plan === 'Direct' ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-100 text-slate-600'}} font-semibold">${{item.plan}}</span>` : '';

        tr.innerHTML = `
          <td class="py-2.5 px-3 text-center border-cell" onclick="event.stopPropagation(); toggleTable2Row(${{idx}});">
            <span class="toggle-btn" id="t2-toggle-btn-${{idx}}">${{isExp ? "−" : "+"}}</span>
          </td>
          <td class="py-2.5 px-3 text-center text-slate-600 font-semibold border-cell">#${{idx + 1}}</td>
          <td class="py-2.5 px-4 text-center border-cell font-black text-[#1F4E79] bg-slate-50/50">
            <span class="inline-flex items-center justify-center px-2.5 py-0.5 rounded-md bg-blue-100 text-[#1F4E79] font-bold text-xs">${{item.composite_score}}</span>
          </td>
          <td class="py-2.5 px-5 text-right border-cell">
            <div class="flex items-center justify-end gap-3">
              <div class="w-16 bg-slate-200 rounded-full h-1.5 overflow-hidden">
                <div class="bg-[#1F4E79] h-1.5 rounded-full" style="width: ${{Math.min(100, item.consistency_score)}}%"></div>
              </div>
              <span class="font-bold text-slate-900 w-7 text-right">${{item.consistency_score}}</span>
            </div>
          </td>
          <td class="py-2.5 px-4 text-right recency-amber border-cell font-bold text-sm">${{item.recency_score}}</td>
          <td class="py-2.5 px-6 font-medium text-slate-900 border-cell text-left">${{item.name}}${{planBadge}}${{optionBadge}}</td>
          <td class="py-2.5 px-4 text-right cagr-green border-cell font-bold">${{item.avg_cagr.toFixed(2)}}%</td>
        `;
        tbody2.appendChild(tr);

        // Nested Sub-Table Row
        const childTr = document.createElement("tr");
        childTr.id = `t2-child-row-${{idx}}`;
        childTr.className = `${{isExp ? "" : "hidden"}} bg-slate-50/80 border-b-2 border-slate-300`;

        const qDefs = [
          {{ id: 1, label: "Q1", subtitle: "Top 25%" }},
          {{ id: 2, label: "Q2", subtitle: "25% - 50%" }},
          {{ id: 3, label: "Q3", subtitle: "50% - 75%" }},
          {{ id: 4, label: "Q4", subtitle: "Bottom 25%" }}
        ];

        let quartileRowsHtml = "";
        qDefs.forEach(q => {{
          let cellsHtml = "";
          item.last_6_details.forEach(w => {{
            if (w.quartile === q.id && w.cagr !== null) {{
              cellsHtml += `
                <td class="py-2.5 px-3 border border-slate-200 text-center align-middle bg-slate-50/40">
                  <div class="bg-black text-white font-bold py-2.5 px-4 rounded-xl text-center shadow flex flex-col items-center justify-center mx-auto min-w-[110px] transition-transform hover:scale-105">
                    <span class="text-xs font-black tracking-wide">${{w.cagr.toFixed(2)}}%</span>
                    <span class="text-[9px] uppercase tracking-wider text-slate-300 font-semibold mt-0.5">${{q.label}}</span>
                  </div>
                </td>
              `;
            }} else {{
              cellsHtml += `
                <td class="py-2.5 px-3 border border-slate-200 text-center align-middle bg-white">
                  <div class="h-11 w-full min-w-[110px] rounded-xl bg-slate-50/50 border border-dashed border-slate-200 flex items-center justify-center text-slate-300 text-xs font-mono">
                    −
                  </div>
                </td>
              `;
            }}
          }});

          quartileRowsHtml += `
            <tr class="hover:bg-slate-50/60">
              <td class="py-2.5 px-4 border border-slate-200 text-left font-bold text-slate-800 bg-slate-50 text-xs w-28">
                ${{q.label}}
                <span class="text-[10px] text-slate-400 font-normal block">${{q.subtitle}}</span>
              </td>
              ${{cellsHtml}}
            </tr>
          `;
        }});

        const endHeadersHtml = endMonthYears.map(my => `
          <th class="py-2.5 px-3 font-bold text-[#1F4E79] bg-slate-100 border border-slate-200 text-center min-w-[125px] text-xs">
            ${{my}}
          </th>
        `).join("");

        const navFormatted = item.nav !== null && item.nav !== undefined ? `₹${{item.nav.toFixed(2)}}` : "—";
        const aumFormatted = item.aum !== null && item.aum !== undefined ? `₹${{Number(item.aum).toLocaleString('en-IN', {{maximumFractionDigits: 2}})}} Cr` : "—";
        const erFormatted = item.expense_ratio !== null && item.expense_ratio !== undefined ? `${{item.expense_ratio}}%` : "—";
        const fmFormatted = item.fund_manager && item.fund_manager !== "None" ? item.fund_manager : "—";

        childTr.innerHTML = `
          <td colspan="7" class="p-4 sm:p-6 bg-slate-50/80">
            <div class="max-w-[98%] mx-auto space-y-4">
              
              <!-- 1. Particular Fund Details Sub-Table -->
              <div class="bg-white border border-slate-200 rounded-xl p-4 shadow-sm">
                <div class="flex items-center justify-between border-b border-slate-100 pb-2 mb-3">
                  <div class="text-xs font-bold text-[#1F4E79] uppercase tracking-wider flex items-center gap-2">
                    <svg class="w-4 h-4 text-emerald-600" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
                    Fund Details & Operations Overview
                  </div>
                  <div class="text-[11px] text-slate-400 font-mono">Code: ${{item.code || 'AMFI'}}</div>
                </div>

                <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 text-center">
                  <div class="p-3 bg-slate-50 rounded-xl border border-slate-100">
                    <div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider mb-0.5">Latest NAV</div>
                    <div class="text-base font-bold text-slate-900">${{navFormatted}}</div>
                  </div>
                  <div class="p-3 bg-slate-50 rounded-xl border border-slate-100">
                    <div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider mb-0.5">Latest AUM</div>
                    <div class="text-base font-bold text-slate-900">${{aumFormatted}}</div>
                  </div>
                  <div class="p-3 bg-slate-50 rounded-xl border border-slate-100">
                    <div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider mb-0.5">Expense Ratio</div>
                    <div class="text-base font-bold text-slate-900">${{erFormatted}}</div>
                  </div>
                  <div class="p-3 bg-slate-50 rounded-xl border border-slate-100 text-center">
                    <div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider mb-0.5">Fund Manager</div>
                    <div class="text-xs font-semibold text-slate-800 line-clamp-2" title="${{fmFormatted}}">${{fmFormatted}}</div>
                  </div>
                </div>
              </div>

              <!-- 2. 2D Quartile Grid with Widened Black Boxes -->
              <div class="bg-white border border-slate-200 rounded-xl p-4 shadow-sm">
                <div class="border-b border-slate-100 pb-2 mb-3">
                  <h4 class="text-xs font-bold text-[#1F4E79] uppercase tracking-wider">
                    3-Year Rolling Return Quartile Tracking Matrix
                  </h4>
                  <p class="text-[11px] text-slate-500 mt-0.5">
                    Quartile standing across rolling 3-year windows (solid black box denotes the quartile tier achieved by the fund for each window)
                  </p>
                </div>

                <div class="overflow-x-auto">
                  <table class="w-full text-center border-collapse">
                    <thead>
                      <tr class="bg-slate-100 text-slate-700 text-xs font-bold border-b border-slate-200">
                        <th class="py-2.5 px-4 border border-slate-200 text-left w-28 font-bold">Quartile</th>
                        ${{endHeadersHtml}}
                      </tr>
                    </thead>
                    <tbody>
                      ${{quartileRowsHtml}}
                    </tbody>
                  </table>
                </div>
              </div>

            </div>
          </td>
        `;
        tbody2.appendChild(childTr);
      }});
    }}

    function toggleTable2Row(idx) {{
      const childTr = document.getElementById(`t2-child-row-${{idx}}`);
      const icon = document.getElementById(`t2-toggle-btn-${{idx}}`);
      const isOpen = t2ExpandedMap[idx];

      if (isOpen) {{
        if (childTr) childTr.classList.add("hidden");
        if (icon) icon.textContent = "+";
        t2ExpandedMap[idx] = false;
      }} else {{
        if (childTr) childTr.classList.remove("hidden");
        if (icon) icon.textContent = "−";
        t2ExpandedMap[idx] = true;
      }}
    }}

    function toggleAllTable2Rows() {{
      isAllT2Expanded = !isAllT2Expanded;
      const btn = document.getElementById("btnExpandAllT2");
      if (btn) btn.textContent = isAllT2Expanded ? "Collapse All [−]" : "Expand All [+]";

      table2DisplayList.forEach((_, idx) => {{
        const childTr = document.getElementById(`t2-child-row-${{idx}}`);
        const icon = document.getElementById(`t2-toggle-btn-${{idx}}`);
        if (isAllT2Expanded) {{
          if (childTr) childTr.classList.remove("hidden");
          if (icon) icon.textContent = "−";
          t2ExpandedMap[idx] = true;
        }} else {{
          if (childTr) childTr.classList.add("hidden");
          if (icon) icon.textContent = "+";
          t2ExpandedMap[idx] = false;
        }}
      }});
    }}

    function sortTable2(colIdx) {{
      if (t2SortCol === colIdx) {{
        t2SortAsc = !t2SortAsc;
      }} else {{
        t2SortCol = colIdx;
        t2SortAsc = (colIdx === 4); // Fund Name default asc, others desc
      }}

      table2DisplayList.sort((a, b) => {{
        let res = 0;
        if (colIdx === 0) {{
          res = 0;
        }} else if (colIdx === 1) {{
          res = a.composite_score - b.composite_score;
        }} else if (colIdx === 2) {{
          res = a.consistency_score - b.consistency_score;
        }} else if (colIdx === 3) {{
          res = a.recency_score - b.recency_score;
        }} else if (colIdx === 4) {{
          res = a.name.localeCompare(b.name);
        }} else if (colIdx === 5) {{
          res = a.avg_cagr - b.avg_cagr;
        }}
        return t2SortAsc ? res : -res;
      }});

      renderTable2Body();
    }}

    function toggleRow(winIdx) {{
      const children = document.querySelectorAll(`.row-child-${{winIdx}}`);
      const icon = document.getElementById(`toggle-btn-${{winIdx}}`);
      const isOpen = expandedMap[winIdx];

      if (isOpen) {{
        children.forEach(c => c.classList.add("hidden"));
        if (icon) icon.textContent = "+";
        expandedMap[winIdx] = false;
      }} else {{
        children.forEach(c => c.classList.remove("hidden"));
        if (icon) icon.textContent = "−";
        expandedMap[winIdx] = true;
      }}
    }}

    function toggleAllRows() {{
      isAllExpanded = !isAllExpanded;
      const btn = document.getElementById("btnExpandAll");
      if (btn) btn.textContent = isAllExpanded ? "Collapse All [−]" : "Expand All [+]";
      const btn2 = document.getElementById("btnExpandAllPill");
      if (btn2) btn2.textContent = isAllExpanded ? "Collapse All [−]" : "Expand All [+]";

      windowDates.forEach((_, winIdx) => {{
        const children = document.querySelectorAll(`.row-child-${{winIdx}}`);
        const icon = document.getElementById(`toggle-btn-${{winIdx}}`);
        if (isAllExpanded) {{
          children.forEach(c => c.classList.remove("hidden"));
          if (icon) icon.textContent = "−";
          expandedMap[winIdx] = true;
        }} else {{
          children.forEach(c => c.classList.add("hidden"));
          if (icon) icon.textContent = "+";
          expandedMap[winIdx] = false;
        }}
      }});
    }}

    function filterFundNames() {{
      const query = (document.getElementById('fundSearchInput')?.value || '').toLowerCase().trim();
      const t1Rows = document.querySelectorAll('#rollingTableBody tr');
      t1Rows.forEach(tr => {{
        if (!query) {{
          // Return to normal display
          const isChild = Array.from(tr.classList).some(c => c.startsWith('row-child-'));
          if (isChild) {{
            const match = Array.from(tr.classList).find(c => c.startsWith('row-child-'));
            const winIdx = match ? match.replace('row-child-', '') : null;
            tr.classList.toggle('hidden', !expandedMap[winIdx]);
          }} else {{
            tr.classList.remove('hidden');
          }}
          return;
        }}
        const text = tr.textContent.toLowerCase();
        tr.classList.toggle('hidden', !text.includes(query));
      }});

      const t2Rows = document.querySelectorAll('#table2Body tr');
      t2Rows.forEach(tr => {{
        if (!query) {{
          const isSub = tr.id && tr.id.startsWith('table2-sub-');
          if (isSub) {{
            const idx = tr.id.replace('table2-sub-', '');
            tr.classList.toggle('hidden', !t2ExpandedMap[idx]);
          }} else {{
            tr.classList.remove('hidden');
          }}
          return;
        }}
        const text = tr.textContent.toLowerCase();
        tr.classList.toggle('hidden', !text.includes(query));
      }});
    }}

    async function downloadExcelFile() {{
      if (!activeComputedDataset) {{
        alert("No active rolling data to export for this selection.");
        return;
      }}

      if (typeof ExcelJS === 'undefined') {{
        alert("Excel export library loading. Please try again in 2 seconds.");
        return;
      }}

      const wb = new ExcelJS.Workbook();
      wb.creator = 'AMFI Statutory Performance Engine';
      const ws = wb.addWorksheet('Rolling_Analysis', {{
        properties: {{
          outlineProperties: {{
            summaryBelow: false,
            summaryRight: false
          }},
          outlineLevelRow: 1
        }},
        views: [{{ showGridLines: true }}]
      }});

      // Typography
      const fontTitle = {{ name: 'Calibri', size: 14, bold: true, color: {{ argb: 'FF1F4E79' }} }};
      const fontSubtitle = {{ name: 'Calibri', size: 10, italic: true, color: {{ argb: 'FF595959' }} }};
      const fontHeader = {{ name: 'Calibri', size: 11, bold: true, color: {{ argb: 'FFFFFFFF' }} }};
      const fontSubHeader = {{ name: 'Calibri', size: 9, bold: true, color: {{ argb: 'FFFFFFFF' }} }};
      const fontWinSummary = {{ name: 'Calibri', size: 10, bold: true, color: {{ argb: 'FF1F4E79' }} }};
      const fontDetail = {{ name: 'Calibri', size: 10, color: {{ argb: 'FF333333' }} }};
      const fontRank = {{ name: 'Calibri', size: 9, italic: true, color: {{ argb: 'FF7F7F7F' }} }};
      const fontDetailCagr = {{ name: 'Calibri', size: 10, bold: true, color: {{ argb: 'FF2E7D32' }} }};
      const fontCardSection = {{ name: 'Calibri', size: 9.5, bold: true, color: {{ argb: 'FF1F4E79' }} }};
      const fontCardLabel = {{ name: 'Calibri', size: 8.5, bold: true, color: {{ argb: 'FF64748B' }} }};
      const fontCardVal = {{ name: 'Calibri', size: 11, bold: true, color: {{ argb: 'FF0F172A' }} }};
      const fontCardValFm = {{ name: 'Calibri', size: 9, bold: false, color: {{ argb: 'FF1E293B' }} }};
      const fontBlackBox = {{ name: 'Calibri', size: 10, bold: true, color: {{ argb: 'FFFFFFFF' }} }};
      const fontEmptyBox = {{ name: 'Calibri', size: 9, color: {{ argb: 'FF94A3B8' }} }};
      const fontQRowLabel = {{ name: 'Calibri', size: 9, bold: true, color: {{ argb: 'FF1E293B' }} }};
      const fontQStatus = {{ name: 'Calibri', size: 9, bold: true, color: {{ argb: 'FF1F4E79' }} }};

      // Fills
      const fillNavy = {{ type: 'pattern', pattern: 'solid', fgColor: {{ argb: 'FF1F4E79' }} }};
      const fillDarkSlate = {{ type: 'pattern', pattern: 'solid', fgColor: {{ argb: 'FF334155' }} }};
      const fillSectionDark = {{ type: 'pattern', pattern: 'solid', fgColor: {{ argb: 'FF1E293B' }} }};
      const fillIceBlue = {{ type: 'pattern', pattern: 'solid', fgColor: {{ argb: 'FFEDF2F8' }} }};
      const fillZebra = {{ type: 'pattern', pattern: 'solid', fgColor: {{ argb: 'FFFAFAFA' }} }};
      const fillCardHeader = {{ type: 'pattern', pattern: 'solid', fgColor: {{ argb: 'FFE8EEF5' }} }};
      const fillCardLabel = {{ type: 'pattern', pattern: 'solid', fgColor: {{ argb: 'FFF8FAFC' }} }};
      const fillBlack = {{ type: 'pattern', pattern: 'solid', fgColor: {{ argb: 'FF000000' }} }};
      const fillWhite = {{ type: 'pattern', pattern: 'solid', fgColor: {{ argb: 'FFFFFFFF' }} }};
      const fillQBadge = {{ type: 'pattern', pattern: 'solid', fgColor: {{ argb: 'FFE2E8F0' }} }};
      const fillQSummary = {{ type: 'pattern', pattern: 'solid', fgColor: {{ argb: 'FFF1F5F9' }} }};

      // Borders
      const borderThin = {{
        top: {{ style: 'thin', color: {{ argb: 'FFE0E0E0' }} }},
        bottom: {{ style: 'thin', color: {{ argb: 'FFE0E0E0' }} }},
        left: {{ style: 'thin', color: {{ argb: 'FFE0E0E0' }} }},
        right: {{ style: 'thin', color: {{ argb: 'FFE0E0E0' }} }}
      }};
      const borderDoubleBottom = {{
        top: {{ style: 'thin', color: {{ argb: 'FFE0E0E0' }} }},
        bottom: {{ style: 'double', color: {{ argb: 'FF1F4E79' }} }},
        left: {{ style: 'thin', color: {{ argb: 'FFE0E0E0' }} }},
        right: {{ style: 'thin', color: {{ argb: 'FFE0E0E0' }} }}
      }};
      const borderSummaryBottom = {{
        top: {{ style: 'thin', color: {{ argb: 'FFE0E0E0' }} }},
        bottom: {{ style: 'medium', color: {{ argb: 'FF1F4E79' }} }},
        left: {{ style: 'thin', color: {{ argb: 'FFE0E0E0' }} }},
        right: {{ style: 'thin', color: {{ argb: 'FFE0E0E0' }} }}
      }};

      // Alignments
      const alignCenter = {{ horizontal: 'center', vertical: 'middle' }};
      const alignLeft = {{ horizontal: 'left', vertical: 'middle' }};
      const alignRight = {{ horizontal: 'right', vertical: 'middle' }};
      const alignLeftWrap = {{ horizontal: 'left', vertical: 'middle', wrapText: true }};

      // 1. PAGE TITLE & FILTER METADATA
      const optLabel = currentSchemeOption === 'All Options' ? 'All Options' : currentSchemeOption;
      const planLabel = currentSchemePlan === 'All Plans' ? 'Regular & Direct Plans' : `${{currentSchemePlan}} Plans`;
      ws.getCell('A1').value = 'MUTUAL FUND ROLLING RETURN ANALYSIS — STATUTORY AMFI PERFORMANCE ENGINE';
      ws.getCell('A1').font = fontTitle;
      ws.getCell('A2').value = `Scheme Type: ${{currentSchemeType}} | Category Group: ${{currentCategoryGroup}} | Sub Category: ${{currentSubCategory}} | Option: ${{optLabel}} | Plan: ${{currentSchemePlan}}`;
      ws.getCell('A2').font = fontSubtitle;
      ws.getCell('A3').value = 'Analysis Period: 10 Years (Jan 2015 to Jul 2026) | Source: Statutory AMFI NAV Feed';
      ws.getCell('A3').font = fontSubtitle;

      // 2. TABLE 1: 3-YEAR ROLLING RETURN ANALYSIS (EXACT REPLICA OF TABLE 1)
      ws.getCell('A5').value = `${{currentSubCategory}} Mutual Funds (${{optLabel}}) — 3-Year Rolling Return Analysis (${{planLabel}})`;
      ws.getCell('A5').font = fontTitle;
      ws.getCell('A6').value = 'Half-Yearly Rolling Windows (Jan 2015 to Jul 2026) | Click [+] on left margin to expand Top 25% funds';
      ws.getCell('A6').font = fontSubtitle;

      const t1Defs = [
        {{ col: 1, label: 'Start Date', align: alignCenter }},
        {{ col: 2, label: 'End Date', align: alignCenter }},
        {{ col: 3, label: 'Quarter', align: alignCenter }},
        {{ col: 4, label: 'Total Funds (n)', align: alignCenter }},
        {{ col: 5, label: 'Top 25% Count', align: alignCenter }},
        {{ col: 6, label: 'Top Quartile Fund Name', align: alignLeft, mergeTo: 8 }},
        {{ col: 9, label: '3Y CAGR (%)', align: alignRight }}
      ];

      t1Defs.forEach(d => {{
        const c = ws.getCell(7, d.col);
        c.value = d.label;
        c.font = fontHeader;
        c.fill = fillNavy;
        c.alignment = d.align;
        c.border = borderDoubleBottom;
        if (d.mergeTo) {{
          for (let mc = d.col; mc <= d.mergeTo; mc++) {{
            const mcCell = ws.getCell(7, mc);
            mcCell.fill = fillNavy;
            mcCell.border = borderDoubleBottom;
          }}
          ws.mergeCells(7, d.col, 7, d.mergeTo);
        }}
      }});

      let currRow = 8;
      activeComputedDataset.windows.forEach((w, winIdx) => {{
        const topFunds = w.top_funds || [];
        const bestFund = topFunds[0] || null;
        const bestName = bestFund ? `#1: ${{bestFund.schemeName}}` : 'N/A';
        const bestCagr = bestFund ? (bestFund.cagr / 100.0) : 0;

        const sRow = currRow;
        ws.getCell(sRow, 1).value = w.start_date;
        ws.getCell(sRow, 2).value = w.end_date;
        ws.getCell(sRow, 3).value = w.quarter;
        ws.getCell(sRow, 4).value = w.n;
        ws.getCell(sRow, 5).value = w.top_25_pct;
        ws.getCell(sRow, 6).value = bestName;
        const cTop = ws.getCell(sRow, 9);
        cTop.value = bestCagr;
        cTop.numFmt = '0.00%';

        for (let c = 1; c <= 9; c++) {{
          const cell = ws.getCell(sRow, c);
          cell.font = fontWinSummary;
          cell.fill = fillIceBlue;
          cell.border = borderSummaryBottom;
          if (c <= 5) cell.alignment = alignCenter;
          else if (c === 6) cell.alignment = alignLeft;
          else if (c === 9) cell.alignment = alignRight;
        }}
        ws.mergeCells(sRow, 6, sRow, 8);

        // Parent row collapsed = true if it has children!
        if (topFunds.length > 1) {{
          const parentRowObj = ws.getRow(sRow);
          Object.defineProperty(parentRowObj, 'collapsed', {{ get: () => true, configurable: true }});
        }}

        currRow++;

        // Child detail rows with native Excel outline dropdown grouping!
        topFunds.slice(1).forEach((fund, rankIdx) => {{
          const rank = rankIdx + 2;
          const dRow = currRow;
          ws.getCell(dRow, 1).value = `#${{rank}}`;
          ws.getCell(dRow, 6).value = fund.schemeName;
          const cCagr = ws.getCell(dRow, 9);
          cCagr.value = fund.cagr / 100.0;
          cCagr.numFmt = '0.00%';

          for (let c = 1; c <= 9; c++) {{
            const cell = ws.getCell(dRow, c);
            cell.border = borderThin;
            if (c === 1) {{
              cell.font = fontRank;
              cell.alignment = alignCenter;
            }} else if (c === 6) {{
              cell.font = fontDetail;
              cell.alignment = alignLeft;
            }} else if (c === 9) {{
              cell.font = fontDetailCagr;
              cell.alignment = alignRight;
            }} else {{
              cell.alignment = alignCenter;
            }}
            if (rank % 2 === 0) cell.fill = fillZebra;
          }}
          ws.mergeCells(dRow, 6, dRow, 8);

          // Native Excel dropdown / outline grouping
          const rowObj = ws.getRow(dRow);
          rowObj.outlineLevel = 1;
          rowObj.hidden = true;
          Object.defineProperty(rowObj, 'collapsed', {{ get: () => false, configurable: true }});

          currRow++;
        }});
      }});

      // 3. TABLE 2: TOP 25% CONSISTENCY SUMMARY
      currRow += 2;
      const last6 = activeComputedDataset.windows.slice(-6);
      const endMonthYears = last6.map(w => w.end_date);
      const maxEligible = windowDates.length;
      const minRequired = Math.round(maxEligible * 0.75);

      ws.getCell(currRow, 1).value = `${{currentSubCategory}} Funds (${{optLabel}}, ${{planLabel}}) — Top 25% Consistency Summary (Grouped by Name)`;
      ws.getCell(currRow, 1).font = fontTitle;
      currRow++;
      ws.getCell(currRow, 1).value = `Ranked by Composite Score (65% Consistency + 35% Recency) | Top 25% Funds (${{activeComputedDataset.consistency.length}}) with ≥ 75% Windows (≥${{minRequired}} of ${{maxEligible}}) | Click [+] on left margin to view Fund Details & 2D Matrix`;
      ws.getCell(currRow, 1).font = fontSubtitle;
      currRow += 2;

      // Table 2 Headers (A to I)
      const t2Defs = [
        {{ col: 1, label: 'Rank', align: alignCenter }},
        {{ col: 2, label: 'Composite Score', align: alignCenter }},
        {{ col: 3, label: 'Consistency Score', align: alignCenter }},
        {{ col: 4, label: 'Recency Score', align: alignCenter }},
        {{ col: 5, label: 'Completed Windows', align: alignCenter }},
        {{ col: 6, label: 'Fund Name', align: alignLeft, mergeTo: 8 }},
        {{ col: 9, label: 'Avg 3Y CAGR (%)', align: alignRight }}
      ];

      t2Defs.forEach(d => {{
        const c = ws.getCell(currRow, d.col);
        c.value = d.label;
        c.font = fontHeader;
        c.fill = fillNavy;
        c.alignment = d.align;
        c.border = borderDoubleBottom;
        if (d.mergeTo) {{
          for (let mc = d.col; mc <= d.mergeTo; mc++) {{
            const mcCell = ws.getCell(currRow, mc);
            mcCell.fill = fillNavy;
            mcCell.border = borderDoubleBottom;
          }}
          ws.mergeCells(currRow, d.col, currRow, d.mergeTo);
        }}
      }});
      currRow++;

      // Table 2 Rows & Symmetrical Sub-Tables
      activeComputedDataset.consistency.forEach((item, idx) => {{
        const fundRow = currRow;
        ws.getCell(fundRow, 1).value = `#${{idx + 1}}`;
        ws.getCell(fundRow, 2).value = item.composite_score;
        ws.getCell(fundRow, 3).value = item.consistency_score;
        ws.getCell(fundRow, 4).value = item.recency_score;
        ws.getCell(fundRow, 5).value = `${{item.eligible || 18}} / 18`;
        ws.getCell(fundRow, 6).value = item.name;
        const cAvg = ws.getCell(fundRow, 9);
        cAvg.value = item.avg_cagr / 100.0;
        cAvg.numFmt = '0.00%';

        for (let c = 1; c <= 9; c++) {{
          const cell = ws.getCell(fundRow, c);
          cell.border = borderThin;
          cell.fill = fillIceBlue;
          if (c === 1) {{
            cell.font = fontWinSummary;
            cell.alignment = alignCenter;
          }} else if (c === 2) {{
            cell.font = {{ name: 'Calibri', size: 11, bold: true, color: {{ argb: 'FF0F172A' }} }};
            cell.alignment = alignCenter;
          }} else if (c === 3) {{
            cell.font = {{ name: 'Calibri', size: 11, bold: true, color: {{ argb: 'FF005F73' }} }};
            cell.alignment = alignCenter;
          }} else if (c === 4) {{
            cell.font = {{ name: 'Calibri', size: 11, bold: true, color: {{ argb: 'FFB45309' }} }};
            cell.alignment = alignCenter;
          }} else if (c === 5) {{
            cell.font = {{ name: 'Calibri', size: 10, bold: true, color: {{ argb: 'FF475569' }} }};
            cell.alignment = alignCenter;
          }} else if (c === 6) {{
            cell.font = {{ name: 'Calibri', size: 10, bold: true, color: {{ argb: 'FF1E293B' }} }};
            cell.alignment = alignLeft;
          }} else if (c === 9) {{
            cell.font = fontDetailCagr;
            cell.alignment = alignRight;
          }}
        }}
        ws.mergeCells(fundRow, 6, fundRow, 8);

        // Parent row collapsed = true for Fund row!
        const fundRowObj = ws.getRow(fundRow);
        Object.defineProperty(fundRowObj, 'collapsed', {{ get: () => true, configurable: true }});

        currRow++;

        function tagSubRow(r) {{
          r.outlineLevel = 1;
          r.hidden = true;
          Object.defineProperty(r, 'collapsed', {{ get: () => false, configurable: true }});
        }}

        // 1. CARD BANNER (Cols B to I)
        const cardBannerRow = currRow;
        ws.getRow(cardBannerRow).height = 20;
        for (let c = 1; c <= 9; c++) {{
          const cell = ws.getCell(cardBannerRow, c);
          cell.fill = fillCardHeader;
          cell.border = borderThin;
        }}
        const bCell = ws.getCell(cardBannerRow, 2);
        bCell.value = 'FUND OPERATIONS & PORTFOLIO METRICS';
        bCell.font = fontCardSection;
        bCell.alignment = alignLeft;
        ws.mergeCells(cardBannerRow, 2, cardBannerRow, 9);
        tagSubRow(ws.getRow(cardBannerRow));
        currRow++;

        // 2. CARD METRIC LABELS ROW (Cols B-C: NAV, D-E: AUM, F: ER, G-I: FM)
        const lblRow = currRow;
        ws.getRow(lblRow).height = 18;
        for (let c = 1; c <= 9; c++) {{
          const cell = ws.getCell(lblRow, c);
          cell.fill = fillCardLabel;
          cell.border = borderThin;
        }}
        ws.getCell(lblRow, 2).value = 'LATEST NAV';
        ws.getCell(lblRow, 2).font = fontCardLabel;
        ws.getCell(lblRow, 2).alignment = alignCenter;
        ws.mergeCells(lblRow, 2, lblRow, 3);

        ws.getCell(lblRow, 4).value = 'LATEST AUM (₹ CR)';
        ws.getCell(lblRow, 4).font = fontCardLabel;
        ws.getCell(lblRow, 4).alignment = alignCenter;
        ws.mergeCells(lblRow, 4, lblRow, 5);

        ws.getCell(lblRow, 6).value = 'EXPENSE RATIO';
        ws.getCell(lblRow, 6).font = fontCardLabel;
        ws.getCell(lblRow, 6).alignment = alignCenter;

        ws.getCell(lblRow, 7).value = 'FUND MANAGER(S)';
        ws.getCell(lblRow, 7).font = fontCardLabel;
        ws.getCell(lblRow, 7).alignment = alignLeft;
        ws.mergeCells(lblRow, 7, lblRow, 9);
        tagSubRow(ws.getRow(lblRow));
        currRow++;

        // 3. CARD METRIC VALUES ROW (Cols B-C: NAV val, D-E: AUM val, F: ER val, G-I: FM val)
        const valRow = currRow;
        ws.getRow(valRow).height = 26;
        for (let c = 1; c <= 9; c++) {{
          const cell = ws.getCell(valRow, c);
          cell.fill = fillWhite;
          cell.border = borderThin;
        }}

        const navVal = item.nav !== null && item.nav !== undefined ? `₹${{item.nav.toFixed(2)}}` : '—';
        const aumVal = item.aum !== null && item.aum !== undefined ? `₹${{Number(item.aum).toLocaleString('en-IN', {{maximumFractionDigits: 2}})}} Cr` : '—';
        const erVal = item.expense_ratio !== null && item.expense_ratio !== undefined ? `${{item.expense_ratio}}%` : '—';
        const fmVal = item.fund_manager && item.fund_manager !== 'None' ? item.fund_manager : '—';

        ws.getCell(valRow, 2).value = navVal;
        ws.getCell(valRow, 2).font = fontCardVal;
        ws.getCell(valRow, 2).alignment = alignCenter;
        ws.mergeCells(valRow, 2, valRow, 3);

        ws.getCell(valRow, 4).value = aumVal;
        ws.getCell(valRow, 4).font = fontCardVal;
        ws.getCell(valRow, 4).alignment = alignCenter;
        ws.mergeCells(valRow, 4, valRow, 5);

        ws.getCell(valRow, 6).value = erVal;
        ws.getCell(valRow, 6).font = fontCardVal;
        ws.getCell(valRow, 6).alignment = alignCenter;

        ws.getCell(valRow, 7).value = fmVal;
        ws.getCell(valRow, 7).font = fontCardValFm;
        ws.getCell(valRow, 7).alignment = alignLeftWrap;
        ws.mergeCells(valRow, 7, valRow, 9);
        tagSubRow(ws.getRow(valRow));
        currRow++;

        // 4. 2D QUARTILE MATRIX SECTION BANNER (Cols B to I)
        const matrixBannerRow = currRow;
        ws.getRow(matrixBannerRow).height = 22;
        for (let c = 1; c <= 9; c++) {{
          const cell = ws.getCell(matrixBannerRow, c);
          cell.fill = fillNavy;
          cell.border = borderThin;
        }}
        const mbCell = ws.getCell(matrixBannerRow, 2);
        mbCell.value = '3-YEAR ROLLING RETURN QUARTILE TRACKING MATRIX (LAST 6 WINDOWS)';
        mbCell.font = fontSubHeader;
        mbCell.alignment = alignCenter;
        ws.mergeCells(matrixBannerRow, 2, matrixBannerRow, 9);
        tagSubRow(ws.getRow(matrixBannerRow));
        currRow++;

        // 5. 2D MATRIX HEADER (Cols B to I)
        const gridHeadRow = currRow;
        ws.getRow(gridHeadRow).height = 20;
        for (let c = 1; c <= 9; c++) {{
          const cell = ws.getCell(gridHeadRow, c);
          cell.fill = fillDarkSlate;
          cell.border = borderThin;
        }}

        const qH = ws.getCell(gridHeadRow, 2);
        qH.value = 'Quartile Standing';
        qH.font = fontSubHeader;
        qH.alignment = alignCenter;

        endMonthYears.forEach((em, wIdx) => {{
          const col = 3 + wIdx; // Cols 3 to 8
          const c = ws.getCell(gridHeadRow, col);
          c.value = em;
          c.font = fontSubHeader;
          c.alignment = alignCenter;
        }});

        const cSumH = ws.getCell(gridHeadRow, 9);
        cSumH.value = 'Quartile Status';
        cSumH.font = fontSubHeader;
        cSumH.alignment = alignCenter;
        tagSubRow(ws.getRow(gridHeadRow));
        currRow++;

        // 6. 4 QUARTILE ROWS (Q1 to Q4)
        const qDefs = [
          {{ id: 1, label: 'Q1 (Top 25%)' }},
          {{ id: 2, label: 'Q2 (25% - 50%)' }},
          {{ id: 3, label: 'Q3 (50% - 75%)' }},
          {{ id: 4, label: 'Q4 (Bottom 25%)' }}
        ];

        qDefs.forEach(q => {{
          const qRow = currRow;
          ws.getRow(qRow).height = 22;

          ws.getCell(qRow, 1).border = borderThin;

          const lblCell = ws.getCell(qRow, 2);
          lblCell.value = q.label;
          lblCell.font = fontQRowLabel;
          lblCell.fill = fillQBadge;
          lblCell.alignment = alignCenter;
          lblCell.border = borderThin;

          let qCount = 0;
          item.last_6_details.forEach((w, wIdx) => {{
            const col = 3 + wIdx; // Cols C, D, E, F, G, H
            const cell = ws.getCell(qRow, col);
            cell.border = borderThin;

            if (w.quartile === q.id && w.cagr !== null) {{
              qCount++;
              // SOLID BLACK WIDENED BOX!
              cell.value = (w.cagr / 100.0);
              cell.numFmt = '0.00%';
              cell.fill = fillBlack;
              cell.font = fontBlackBox;
              cell.alignment = alignCenter;
            }} else {{
              cell.value = '-';
              cell.fill = fillZebra;
              cell.font = fontEmptyBox;
              cell.alignment = alignCenter;
            }}
          }});

          // Col 9 Status
          const statusCell = ws.getCell(qRow, 9);
          statusCell.border = borderThin;
          statusCell.alignment = alignCenter;
          if (qCount > 0) {{
            statusCell.value = `${{qCount}} of 6 in ${{q.label.slice(0, 2)}}`;
            statusCell.font = fontQStatus;
            statusCell.fill = fillQSummary;
          }} else {{
            statusCell.value = '-';
            statusCell.font = fontEmptyBox;
            statusCell.fill = fillWhite;
          }}

          tagSubRow(ws.getRow(qRow));
          currRow++;
        }});

        // 7. Spacer row
        const spacerRow = currRow;
        ws.getRow(spacerRow).height = 6;
        for (let c = 1; c <= 9; c++) {{
          ws.getCell(spacerRow, c).border = borderThin;
        }}
        tagSubRow(ws.getRow(spacerRow));
        currRow++;
      }});

      // UNIFORM, SYMMETRICAL COLUMN WIDTHS
      ws.columns = [
        {{ width: 12 }}, // Col A: Rank / Start Date
        {{ width: 16 }}, // Col B: Composite Score / Quartile Label / End Date
        {{ width: 14 }}, // Col C: Consistency Score / Jan 2024 / Quarter
        {{ width: 14 }}, // Col D: Recency Score / Jul 2024 / Total Funds
        {{ width: 14 }}, // Col E: Completed Windows / Jan 2025 / Top 25% Count
        {{ width: 14 }}, // Col F: Jul 2025 / Fund Name Part 1
        {{ width: 14 }}, // Col G: Jan 2026 / Fund Name Part 2
        {{ width: 14 }}, // Col H: Jul 2026 / Fund Name Part 3
        {{ width: 18 }}  // Col I: 3Y CAGR (%) / Fund Manager
      ];

      const safeSub = currentSubCategory.toLowerCase().replace(/[^a-z0-9]/g, "_");
      const safeOpt = currentSchemeOption.toLowerCase().replace(/[^a-z0-9]/g, "_");
      const safePlan = currentSchemePlan.toLowerCase().replace(/[^a-z0-9]/g, "_");
      const fileName = `${{safeSub}}_${{safePlan}}_${{safeOpt}}_3yr_rolling_analysis.xlsx`;

      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], {{ type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }});
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    }}

    if (document.readyState === 'loading') {{
      document.addEventListener('DOMContentLoaded', initPage);
    }} else {{
      initPage();
    }}
  </script>
</body>
</html>
'''

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_code)

brain_path = r"C:\Users\hiten\.gemini\antigravity\brain\94742509-c694-497e-9599-a7a4a826c346\index.html"
with open(brain_path, "w", encoding="utf-8") as f:
    f.write(html_code)

print("Updated index.html: Columns removed, >75% eligible windows threshold applied, ranking by Composite Score (65% consistency + 35% recency), all 3 scores rounded integers with no % sign!")
