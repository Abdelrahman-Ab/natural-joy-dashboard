(() => {
  'use strict';

  const COLORS = {
    forest: '#173c2d',
    leaf: '#3c7450',
    leaf2: '#5a9067',
    lightGreen: '#9acb87',
    gold: '#c89e4d',
    gold2: '#e4c982',
    sand: '#f7f2e8',
    muted: '#667368',
    red: '#b24438',
    amber: '#ce8e38',
    purple: '#77509b',
    grid: '#ece6d8'
  };
  const RISK_COLORS = {
    'تجاوز الميزانية': COLORS.red,
    'عالية المخاطر': '#d55b46',
    'تحتاج متابعة': COLORS.amber,
    'غير مخطط': COLORS.purple,
    'مكتملة': COLORS.leaf,
    'ضمن الخطة': '#87af7d'
  };

  let dashboard = null;
  let currentVersion = null;
  const state = { search: '', category: '', risk: '', status: '', selected: null };
  const el = id => document.getElementById(id);
  const money = value => new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 }).format(value || 0) + ' ج.م';
  const pct = value => new Intl.NumberFormat('en-US', { style: 'percent', maximumFractionDigits: 1 }).format(value || 0);
  const isPhone = () => window.innerWidth <= 650;
  const compactMoney = value => new Intl.NumberFormat('en-US', { notation: 'compact', maximumFractionDigits: 1 }).format(value || 0) + ' ج.م';

  const layoutBase = {
    font: { family: 'Segoe UI, Tahoma, Arial', color: COLORS.forest, size: 12 },
    paper_bgcolor: 'rgba(0,0,0,0)',
    plot_bgcolor: 'rgba(0,0,0,0)',
    margin: { t: 20, r: 20, b: 45, l: 60 },
    hoverlabel: { bgcolor: COLORS.forest, font: { color: '#fff' } },
    xaxis: { gridcolor: COLORS.grid, zerolinecolor: COLORS.forest, zerolinewidth: 1, fixedrange: true },
    yaxis: { gridcolor: COLORS.grid, zerolinecolor: COLORS.forest, zerolinewidth: 1, fixedrange: true },
    legend: { orientation: 'h', y: -0.18, x: 0 },
    dragmode: false,
    hovermode: 'closest'
  };
  const config = {
    responsive: true,
    displaylogo: false,
    displayModeBar: false,
    scrollZoom: false,
    doubleClick: false
  };

  function riskClass(risk) {
    return ({
      'تجاوز الميزانية': 'risk-over',
      'عالية المخاطر': 'risk-high',
      'تحتاج متابعة': 'risk-watch',
      'غير مخطط': 'risk-unplanned',
      'مكتملة': 'risk-done',
      'ضمن الخطة': 'risk-ok'
    })[risk] || 'risk-ok';
  }

  function taskStatus(task) {
    if (task.completed === 'نعم' || task.progress >= 0.999) return 'completed';
    if (task.progress > 0 || task.paid > 0) return 'inprogress';
    return 'notstarted';
  }

  function filteredTasks() {
    if (!dashboard) return [];
    const q = state.search.trim().toLowerCase();
    return dashboard.tasks.filter(task =>
      (!q || task.name.toLowerCase().includes(q) || task.category.toLowerCase().includes(q)) &&
      (!state.category || task.category === state.category) &&
      (!state.risk || task.risk === state.risk) &&
      (!state.status || taskStatus(task) === state.status)
    );
  }

  function aggregate(tasks) {
    const totalBudget = tasks.reduce((sum, task) => sum + task.budget, 0);
    const totalPaid = tasks.reduce((sum, task) => sum + task.paid, 0);
    const remaining = tasks.reduce((sum, task) => sum + task.remaining, 0);
    const savings = tasks.reduce((sum, task) => sum + task.savings, 0);
    const progressWeight = tasks.reduce((sum, task) => sum + (task.budget > 0 ? task.budget : 0), 0);
    const progressRatio = progressWeight
      ? tasks.reduce((sum, task) => sum + (Math.max(0, Math.min(task.progress, 1)) * Math.max(task.budget, 0)), 0) / progressWeight
      : (tasks.length ? tasks.reduce((sum, task) => sum + Math.max(0, Math.min(task.progress, 1)), 0) / tasks.length : 0);
    return {
      totalBudget,
      totalPaid,
      remaining,
      savings,
      spendRatio: totalBudget ? totalPaid / totalBudget : 0,
      progressRatio
    };
  }

  function showError(message) {
    const toast = el('errorToast');
    toast.textContent = message;
    toast.hidden = false;
    setTimeout(() => { toast.hidden = true; }, 7000);
    el('dataStatus').innerHTML = '<span class="dot" style="background:#e06756"></span> خطأ في البيانات';
  }

  async function loadData(force = false) {
    try {
      const response = await fetch('/api/data?ts=' + Date.now(), { cache: 'no-store' });
      const incoming = await response.json();
      if (!response.ok || incoming.error) throw new Error(incoming.error || 'Unable to read data.xlsx');
      if (!force && currentVersion === incoming.version) return;
      dashboard = incoming;
      currentVersion = incoming.version;
      el('lastUpdated').textContent = new Date(incoming.loaded_at).toLocaleString('en-GB');
      el('dataStatus').innerHTML = '<span class="dot"></span> متصل بـ OneDrive';
      buildFilters();
      renderAll();
    } catch (error) {
      showError('تعذر تحميل ملف Excel من OneDrive: ' + error.message);
    }
  }

  function buildFilters() {
    const category = el('categoryFilter');
    const risk = el('riskFilter');
    const categoryValue = state.category;
    const riskValue = state.risk;
    category.innerHTML = '<option value="">جميع التصنيفات</option>' +
      [...new Set(dashboard.tasks.map(task => task.category))].sort().map(value => `<option value="${value}">${value}</option>`).join('');
    risk.innerHTML = '<option value="">جميع مستويات المخاطر</option>' +
      [...new Set(dashboard.tasks.map(task => task.risk))].map(value => `<option value="${value}">${value}</option>`).join('');
    category.value = categoryValue;
    risk.value = riskValue;
  }

  function freshPlot(id, traces, layout) {
    const target = el(id);
    Plotly.purge(target);
    return Plotly.newPlot(target, traces, { ...layoutBase, ...layout }, config);
  }

  function niceStep(span) {
    if (!Number.isFinite(span) || span <= 0) return 100000;
    const rough = span / 5;
    const magnitude = Math.pow(10, Math.floor(Math.log10(rough)));
    const normalized = rough / magnitude;
    const factor = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
    return factor * magnitude;
  }

  function financialScale(values) {
    const valid = values.filter(value => Number.isFinite(value));
    const min = Math.min(0, ...(valid.length ? valid : [0]));
    const max = Math.max(0, ...(valid.length ? valid : [0]));
    const step = niceStep(Math.max(max - min, Math.abs(max), Math.abs(min)));
    return {
      range: [min < 0 ? -Math.ceil(Math.abs(min) / step) * step : 0, Math.max(step, Math.ceil(max / step) * step)],
      tick0: 0,
      dtick: step,
      tickformat: ',.0f',
      separatethousands: true,
      fixedrange: true,
      gridcolor: COLORS.grid,
      zeroline: true,
      zerolinecolor: COLORS.forest,
      zerolinewidth: 1
    };
  }

  function renderAll() {
    const tasks = filteredTasks();
    const summary = aggregate(tasks);
    const full = dashboard.overview;
    el('filterSummary').textContent = tasks.length === dashboard.tasks.length
      ? `عرض جميع بنود المشروع (${tasks.length} بندًا)`
      : `عرض ${tasks.length} من أصل ${dashboard.tasks.length} بندًا`;
    renderKpis(summary, full, tasks.length !== dashboard.tasks.length);
    renderGauge(summary);
    renderMonthly(tasks);
    renderCumulative(tasks);
    renderCash();
    renderPartners();
    renderPartnersGauge();
    renderUtilization(tasks);
    renderRisk(tasks);
    renderTreemap(tasks);
    renderMisc();
    renderHeatmap(tasks);
    renderTable(tasks);
    renderSelected();
  }

  function renderKpis(summary, full, isFiltered) {
    const cards = [
      ['إجمالي الميزانية', summary.totalBudget, 'ج.م', ''],
      ['إجمالي المصروفات', summary.totalPaid, 'ج.م', ''],
      ['إجمالي المتبقي', summary.remaining, 'ج.م', ''],
      ['نسبة الصرف', summary.spendRatio, '%', summary.spendRatio > 0.8 ? 'alert' : ''],
      ['رأس المال المسدد', full.capital_paid, 'ج.م', ''],
      ['النقد المتبقي', full.cash_remaining, 'ج.م', full.cash_remaining < 0 ? 'alert' : ''],
      ['إجمالي التوفير', summary.savings, 'ج.م', summary.savings < 0 ? 'alert' : ''],
      ['عدد البنود المعروضة', isFiltered ? filteredTasks().length : dashboard.tasks.length, 'بند', '']
    ];
    el('kpis').innerHTML = cards.map(([label, value, unit, cls]) =>
      `<article class="card kpi ${cls}"><span class="label">${label}</span><strong class="value">${unit === '%' ? pct(value) : (unit === 'بند' ? value : money(value))}</strong><span class="unit">${unit === '%' ? 'من الميزانية المعروضة' : unit === 'بند' ? 'بنود العمل' : 'جنيه مصري'}</span></article>`
    ).join('');
  }

  function renderGauge(summary) {
    const value = summary.progressRatio * 100;
    const barColor = value < 60 ? COLORS.red : value < 80 ? COLORS.amber : value < 90 ? COLORS.lightGreen : COLORS.leaf;
    freshPlot('gaugeChart', [{
      type: 'indicator',
      mode: 'gauge+number',
      value,
      number: { suffix: '%', font: { size: 38, color: COLORS.forest }, valueformat: '.1f' },
      gauge: {
        axis: { range: [0, 100], ticksuffix: '%', fixedrange: true },
        bar: { color: barColor },
        bgcolor: '#f1ede2',
        steps: [
          { range: [0, 60], color: '#f2d4d0' },
          { range: [60, 80], color: '#f4dfc0' },
          { range: [80, 90], color: '#dbead4' },
          { range: [90, 100], color: '#c8e0cc' }
        ]
      },
      hovertemplate: 'نسبة الإنجاز: %{value:.1f}%<extra></extra>'
    }], { margin: { t: 15, r: 30, b: 10, l: 30 } });
  }

  function monthSeries(tasks) {
    return dashboard.months.map(month => ({
      month,
      actual: tasks.reduce((sum, task) => sum + (task.monthly[month]?.actual || 0), 0),
      planned: tasks.reduce((sum, task) => sum + (task.monthly[month]?.planned || 0), 0)
    }));
  }

  function renderMonthly(tasks) {
    const series = monthSeries(tasks);
    const yAxis = financialScale(series.flatMap(month => [month.actual, month.planned]));
    freshPlot('monthlyChart', [
      {
        type: 'bar', name: 'الفعلي', x: series.map(month => month.month), y: series.map(month => month.actual),
        marker: { color: series.map(month => month.actual < 0 ? COLORS.red : COLORS.leaf) },
        hovertemplate: '%{x}<br>الفعلي: %{y:,.0f} ج.م<extra></extra>'
      },
      {
        type: 'bar', name: 'المخطط', x: series.map(month => month.month), y: series.map(month => month.planned),
        marker: { color: series.map(month => month.planned < 0 ? COLORS.red : COLORS.gold2) },
        hovertemplate: '%{x}<br>المخطط: %{y:,.0f} ج.م<extra></extra>'
      }
    ], {
      barmode: 'group',
      yaxis: { ...yAxis, title: 'جنيه مصري' },
      xaxis: { ...layoutBase.xaxis, tickangle: -25, fixedrange: true }
    });
  }

  function renderCumulative(tasks) {
    const series = monthSeries(tasks);
    let runningActual = 0;
    let runningPlanned = 0;
    const actual = series.map(month => (runningActual += month.actual));
    const planned = series.map(month => (runningPlanned += month.planned));
    const yAxis = financialScale([...actual, ...planned]);
    freshPlot('cumulativeChart', [
      {
        type: 'scatter', mode: 'lines+markers', name: 'التراكمي الفعلي', x: series.map(month => month.month), y: actual,
        line: { color: COLORS.leaf, width: 4, shape: 'spline' }, fill: 'tozeroy', fillcolor: 'rgba(60,116,80,.10)',
        hovertemplate: '%{x}<br>%{y:,.0f} ج.م<extra></extra>'
      },
      {
        type: 'scatter', mode: 'lines+markers', name: 'التراكمي المخطط', x: series.map(month => month.month), y: planned,
        line: { color: COLORS.gold, width: 3, dash: 'dash' },
        hovertemplate: '%{x}<br>%{y:,.0f} ج.م<extra></extra>'
      }
    ], {
      hovermode: 'x unified',
      yaxis: yAxis,
      xaxis: { ...layoutBase.xaxis, tickangle: -25, fixedrange: true }
    });
  }

  function renderCash() {
    const overview = dashboard.overview;
    const labels = ['رأس المال المسدد', 'المصروفات', 'النقد المتبقي'];
    const values = [overview.capital_paid, -overview.total_paid, overview.cash_remaining];
    const yAxis = financialScale(values);
    freshPlot('cashChart', [{
      type: 'bar',
      x: labels,
      y: values,
      marker: { color: [COLORS.leaf, COLORS.red, overview.cash_remaining < 0 ? COLORS.red : COLORS.gold] },
      hovertemplate: '%{x}<br>%{y:,.0f} ج.م<extra></extra>'
    }], {
      margin: { t: 20, r: 15, b: 75, l: 55 },
      showlegend: false,
      yaxis: yAxis,
      xaxis: { ...layoutBase.xaxis, tickangle: -18, fixedrange: true }
    });
  }

  function renderPartners() {
    const partners = [...(dashboard.partners || [])].sort((a, b) => b.share - a.share).reverse();
    const paidPositive = partners.map(partner => Math.max(partner.paid, 0));
    const remainingPositive = partners.map(partner => Math.max(partner.remaining, 0));
    const negative = partners.map(partner => Math.min(partner.paid, 0) + Math.min(partner.remaining, 0));
    const positiveTotal = partners.map((_, index) => paidPositive[index] + remainingPositive[index]);
    const xAxis = financialScale([...positiveTotal, ...negative]);
    freshPlot('partnersChart', [
      {
        type: 'bar', orientation: 'h', name: 'المدفوع', y: partners.map(partner => partner.name), x: paidPositive,
        marker: { color: COLORS.leaf },
        customdata: partners.map(partner => partner.share ? partner.paid / partner.share : 0),
        hovertemplate: '%{y}<br>المدفوع: %{x:,.0f} ج.م<br>نسبة السداد: %{customdata:.1%}<extra></extra>'
      },
      {
        type: 'bar', orientation: 'h', name: 'المتبقي', y: partners.map(partner => partner.name), x: remainingPositive,
        marker: { color: '#ded6c4' }, hovertemplate: '%{y}<br>المتبقي: %{x:,.0f} ج.م<extra></extra>'
      },
      {
        type: 'bar', orientation: 'h', name: 'قيمة سالبة', y: partners.map(partner => partner.name), x: negative,
        marker: { color: COLORS.red }, hovertemplate: '%{y}<br>قيمة سالبة: %{x:,.0f} ج.م<extra></extra>'
      }
    ], {
      barmode: 'relative',
      margin: { t: 12, r: 18, b: 48, l: isPhone() ? 104 : 175 },
      xaxis: { ...xAxis, title: 'جنيه مصري', tickfont: { size: isPhone() ? 10 : 12 } },
      yaxis: { automargin: true, fixedrange: true }
    });
  }

  function renderPartnersGauge() {
    const partners = dashboard.partners || [];
    const totalShare = partners.reduce((sum, partner) => sum + partner.share, 0);
    const totalPaid = partners.reduce((sum, partner) => sum + partner.paid, 0);
    const value = totalShare ? (totalPaid / totalShare) * 100 : 0;
    freshPlot('partnersGaugeChart', [{
      type: 'indicator',
      mode: 'gauge+number',
      value,
      number: { suffix: '%', font: { size: 38, color: COLORS.forest }, valueformat: '.1f' },
      gauge: {
        shape: 'angular',
        axis: { range: [0, 100], ticksuffix: '%', fixedrange: true },
        bar: { color: COLORS.leaf },
        bgcolor: '#ecebea',
        steps: [
          { range: [0, 100], color: '#e6e8e5' }
        ]
      },
      hovertemplate: `إجمالي المدفوع: ${money(totalPaid)}<br>إجمالي الحصص: ${money(totalShare)}<br>نسبة السداد: %{value:.1f}%<extra></extra>`
    }], { margin: { t: 15, r: 30, b: 10, l: 30 } });
  }

  function selectTask(name) {
    state.selected = dashboard.tasks.find(task => task.name === name) || null;
    renderSelected();
    if (state.selected) el('selectedCard').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  function renderUtilization(tasks) {
    const grouped = {};
    tasks.forEach(task => {
      if (!grouped[task.category]) grouped[task.category] = { category: task.category, paid: 0, remaining: 0, budget: 0 };
      grouped[task.category].paid += task.paid;
      grouped[task.category].remaining += task.remaining;
      grouped[task.category].budget += task.budget;
    });
    const data = Object.values(grouped).sort((a, b) => b.budget - a.budget).reverse();
    // When a category is overspent (negative remaining), show only its red
    // deficit bar. Do not stack the paid green segment before the red amount.
    const isOverspent = data.map(item => item.remaining < 0);
    const paidPositive = data.map((item, index) => isOverspent[index] ? 0 : Math.max(item.paid, 0));
    const remainingPositive = data.map((item, index) => isOverspent[index] ? 0 : Math.max(item.remaining, 0));
    const negativeShownRight = data.map((item, index) => isOverspent[index] ? Math.abs(item.remaining) : 0);
    const visibleTotal = data.map((_, index) => paidPositive[index] + remainingPositive[index] + negativeShownRight[index]);
    const xAxis = financialScale(visibleTotal);
    freshPlot('utilizationChart', [
      {
        type: 'bar', orientation: 'h', name: 'المدفوع', y: data.map(item => item.category), x: paidPositive,
        marker: { color: COLORS.leaf }, hovertemplate: '%{y}<br>المدفوع: %{x:,.0f} ج.م<extra></extra>'
      },
      {
        type: 'bar', orientation: 'h', name: 'المتبقي', y: data.map(item => item.category), x: remainingPositive,
        marker: { color: '#ded6c4' }, hovertemplate: '%{y}<br>المتبقي: %{x:,.0f} ج.م<extra></extra>'
      },
      {
        type: 'bar', orientation: 'h', name: 'متبقي سالب', y: data.map(item => item.category), x: negativeShownRight,
        marker: { color: COLORS.red }, hovertemplate: '%{y}<br>متبقي سالب: -%{x:,.0f} ج.م<extra></extra>'
      }
    ], {
      barmode: 'relative',
      margin: { t: 10, r: 12, b: 48, l: isPhone() ? 105 : 180 },
      xaxis: { ...xAxis, title: 'جنيه مصري', tickfont: { size: isPhone() ? 10 : 12 } },
      yaxis: { automargin: true, fixedrange: true, tickfont: { size: isPhone() ? 10 : 12 } }
    });
  }

  function renderRisk(tasks) {
    const grouped = {};
    tasks.forEach(task => { grouped[task.risk] = (grouped[task.risk] || 0) + task.paid; });
    const labels = Object.keys(grouped);
    const values = labels.map(label => grouped[label]);
    freshPlot('riskChart', [{
      type: 'pie', hole: 0.62, labels, values,
      marker: { colors: labels.map(label => RISK_COLORS[label]) },
      textinfo: 'percent',
      hovertemplate: '%{label}<br>%{value:,.0f} ج.م<br>%{percent}<extra></extra>'
    }], { margin: { t: 10, r: 20, b: 32, l: 20 }, legend: { orientation: 'h', y: -0.04 } });
  }

  function renderTreemap(tasks) {
    // Keep the detailed hierarchy: classification boxes contain their individual work items.
    // Labels are shortened only inside crowded boxes; hover always shows the full text and value.
    const rootId = '__budget_root__';
    const totalBudget = tasks.reduce((sum, task) => sum + Math.max(task.budget, 0), 0);
    const labels = ['إجمالي الميزانية'];
    const ids = [rootId];
    const parents = [''];
    const values = [totalBudget];
    const visibleText = [''];
    const hoverData = [['إجمالي الميزانية', tasks.length]];
    const nodeColors = ['#dfe9dc'];
    const categoryPalette = ['#9ebba2', '#adc6ae', '#c2d5c0', '#d6e1d2', '#e2ddca', '#d7c69f', '#c9b37c', '#b79c5f'];
    const groups = [...new Set(tasks.map(task => task.category))]
      .map(category => ({
        category,
        tasks: tasks.filter(task => task.category === category),
        value: tasks.filter(task => task.category === category).reduce((sum, task) => sum + Math.max(task.budget, 0), 0)
      }))
      .filter(group => group.value > 0)
      .sort((a, b) => b.value - a.value);

    const shortLabel = name => {
      const cleaned = String(name || '').replace(/\s+/g, ' ').trim();
      if (cleaned.length <= 28) return cleaned;
      return cleaned.slice(0, 27).trim() + '…';
    };

    groups.forEach((group, categoryIndex) => {
      const categoryId = `category-${categoryIndex}`;
      labels.push(group.category);
      ids.push(categoryId);
      parents.push(rootId);
      values.push(group.value);
      visibleText.push(`<b>${group.category}</b><br>${compactMoney(group.value)}`);
      hoverData.push([group.category, group.tasks.length]);
      nodeColors.push(categoryPalette[Math.min(categoryIndex, categoryPalette.length - 1)]);

      group.tasks.sort((a, b) => b.budget - a.budget).forEach((task, taskIndex) => {
        labels.push(task.name);
        ids.push(`${categoryId}-task-${taskIndex}`);
        parents.push(categoryId);
        values.push(Math.max(task.budget, 0));
        visibleText.push(shortLabel(task.name));
        hoverData.push([task.name, 1]);
        nodeColors.push('#d9e4d6');
      });
    });

    freshPlot('treemapChart', [{
      type: 'treemap',
      labels,
      ids,
      parents,
      values,
      text: visibleText,
      customdata: hoverData,
      branchvalues: 'total',
      texttemplate: '%{text}',
      textposition: 'middle center',
      textfont: { size: isPhone() ? 10 : 12, color: COLORS.forest },
      marker: { colors: nodeColors, line: { color: '#fffdf8', width: isPhone() ? 2 : 3 } },
      hovertemplate: '<b>%{customdata[0]}</b><br>الميزانية: %{value:,.0f} ج.م<extra></extra>',
      pathbar: { visible: false },
      tiling: { pad: isPhone() ? 2 : 3, squarifyratio: 1 },
      sort: false
    }], {
      margin: { t: 4, r: 4, b: 4, l: 4 },
      uniformtext: { minsize: isPhone() ? 9 : 10, mode: 'hide' }
    }).then(() => {
      const chart = el('treemapChart');
      chart.on('plotly_treemapclick', () => false);
    });
  }

  function renderMisc() {
    const misc = dashboard.misc;
    const yAxis = financialScale(misc.map(item => item.paid));
    freshPlot('miscChart', [{
      type: 'bar', x: misc.map(item => item.name), y: misc.map(item => item.paid),
      marker: { color: misc.map(item => item.paid < 0 ? COLORS.red : COLORS.purple) },
      hovertemplate: '%{x}<br>%{y:,.0f} ج.م<extra></extra>'
    }], {
      margin: { t: 15, r: 18, b: 80, l: 55 },
      xaxis: { ...layoutBase.xaxis, tickangle: -20, fixedrange: true },
      yaxis: yAxis,
      showlegend: false
    });
  }

  function renderHeatmap(tasks) {
    const top = [...tasks].sort((a, b) => b.paid - a.paid).slice(0, isPhone() ? 8 : 12);
    const z = top.map(task => dashboard.months.map(month => task.monthly[month]?.actual || 0));
    freshPlot('heatmapChart', [{
      type: 'heatmap', x: dashboard.months, y: top.map(task => task.name), z,
      colorscale: [[0, '#ffffff'], [0.10, '#f6f0e3'], [0.35, '#e7d3aa'], [0.62, '#b4c89d'], [0.82, COLORS.leaf2], [1, COLORS.forest]],
      hovertemplate: '%{y}<br>%{x}<br>%{z:,.0f} ج.م<extra></extra>'
    }], {
      paper_bgcolor: '#ffffff',
      plot_bgcolor: '#ffffff',
      margin: { t: 12, r: 10, b: 62, l: isPhone() ? 108 : 230 },
      xaxis: { ...layoutBase.xaxis, tickangle: isPhone() ? -45 : -25, fixedrange: true, tickfont: { size: isPhone() ? 9 : 12 } },
      yaxis: { automargin: true, fixedrange: true, tickfont: { size: isPhone() ? 9 : 12 } }
    });
  }

  function renderTable(tasks) {
    el('tasksBody').innerHTML = [...tasks].sort((a, b) => b.budget - a.budget).map(task => {
      const complete = task.progress >= 0.999;
      const rowClass = complete && task.remaining > 0 ? 'row-complete-positive' : complete && task.remaining < 0 ? 'row-complete-negative' : '';
      return `<tr class="${rowClass}" data-task="${encodeURIComponent(task.name)}"><td><strong>${task.name}</strong></td><td>${task.category}</td><td class="num">${money(task.budget)}</td><td class="num">${money(task.paid)}</td><td class="num">${money(task.remaining)}</td><td class="num">${pct(task.spend_ratio)}</td><td><span>${pct(task.progress)}</span><div class="progress"><i style="width:${Math.min(task.progress * 100, 100)}%"></i></div></td><td><span class="risk-pill ${riskClass(task.risk)}">${task.risk}</span></td></tr>`;
    }).join('');
    el('tasksBody').querySelectorAll('tr').forEach(row => row.addEventListener('click', () => selectTask(decodeURIComponent(row.dataset.task))));
  }

  function renderSelected() {
    const card = el('selectedCard');
    const task = state.selected;
    if (!task) { card.hidden = true; return; }
    card.hidden = false;
    el('selectedContent').innerHTML = `<div class="selected-grid"><div class="selected-name">${task.name}<br><span class="risk-pill ${riskClass(task.risk)}" style="margin-top:9px">${task.risk}</span></div><div class="selected-stat"><span>الميزانية</span><strong>${money(task.budget)}</strong></div><div class="selected-stat"><span>المدفوع</span><strong>${money(task.paid)}</strong></div><div class="selected-stat"><span>نسبة الصرف</span><strong>${pct(task.spend_ratio)}</strong></div><div class="selected-stat"><span>نسبة الإنجاز</span><strong>${pct(task.progress)}</strong></div></div>`;
  }

  function bindControls() {
    el('refreshBtn').addEventListener('click', () => loadData(true));
    el('searchFilter').addEventListener('input', event => { state.search = event.target.value; renderAll(); });
    el('categoryFilter').addEventListener('change', event => { state.category = event.target.value; renderAll(); });
    el('riskFilter').addEventListener('change', event => { state.risk = event.target.value; renderAll(); });
    el('statusFilter').addEventListener('change', event => { state.status = event.target.value; renderAll(); });
    el('clearFilters').addEventListener('click', () => {
      Object.assign(state, { search: '', category: '', risk: '', status: '', selected: null });
      ['searchFilter', 'categoryFilter', 'riskFilter', 'statusFilter'].forEach(id => { el(id).value = ''; });
      renderAll();
    });
    el('closeDetail').addEventListener('click', () => { state.selected = null; renderSelected(); });
  }

  let resizeTimer = null;
  let lastPhoneMode = isPhone();
  window.addEventListener('resize', () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
      const phoneMode = isPhone();
      if (dashboard && phoneMode !== lastPhoneMode) {
        lastPhoneMode = phoneMode;
        renderAll();
      } else {
        document.querySelectorAll('.js-plotly-plot').forEach(chart => Plotly.Plots.resize(chart));
      }
    }, 150);
  });

  window.addEventListener('DOMContentLoaded', () => {
    bindControls();
    loadData(true);
    setInterval(() => loadData(false), 60000);
  });
})();
