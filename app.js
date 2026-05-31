(() => {
  'use strict';
  const COLORS = { forest:'#173c2d', leaf:'#3c7450', leaf2:'#5a9067', gold:'#c89e4d', gold2:'#e4c982', sand:'#f7f2e8', muted:'#667368', red:'#b24438', amber:'#ce8e38', purple:'#77509b', grid:'#ece6d8' };
  const RISK_COLORS = { 'تجاوز الميزانية': COLORS.red, 'عالية المخاطر':'#d55b46', 'تحتاج متابعة':COLORS.amber, 'غير مخطط':COLORS.purple, 'مكتملة':COLORS.leaf, 'ضمن الخطة': '#87af7d' };
  let dashboard = null;
  let currentVersion = null;
  const state = { search:'', category:'', risk:'', status:'', selected:null };
  const el = id => document.getElementById(id);
  const money = value => new Intl.NumberFormat('en-US', {maximumFractionDigits:0}).format(value || 0) + ' ج.م';
  const pct = value => new Intl.NumberFormat('en-US', {style:'percent', maximumFractionDigits:1}).format(value || 0);
  const layoutBase = { font:{family:'Segoe UI, Tahoma, Arial', color:COLORS.forest, size:12}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', margin:{t:20, r:20, b:45, l:60}, hoverlabel:{bgcolor:COLORS.forest, font:{color:'#fff'}}, xaxis:{gridcolor:COLORS.grid, zerolinecolor:COLORS.grid}, yaxis:{gridcolor:COLORS.grid, zerolinecolor:COLORS.grid}, legend:{orientation:'h', y:-0.18, x:0} };
  const config = { responsive:true, displaylogo:false, modeBarButtonsToRemove:['lasso2d','select2d','autoScale2d'] };

  function riskClass(risk){ return ({'تجاوز الميزانية':'risk-over','عالية المخاطر':'risk-high','تحتاج متابعة':'risk-watch','غير مخطط':'risk-unplanned','مكتملة':'risk-done','ضمن الخطة':'risk-ok'})[risk] || 'risk-ok'; }
  function taskStatus(t){ if(t.completed === 'نعم' || t.progress >= .999) return 'completed'; if(t.progress > 0 || t.paid > 0) return 'inprogress'; return 'notstarted'; }
  function filteredTasks(){
    if(!dashboard) return [];
    const q = state.search.trim().toLowerCase();
    return dashboard.tasks.filter(t => (!q || t.name.toLowerCase().includes(q) || t.category.toLowerCase().includes(q)) && (!state.category || t.category===state.category) && (!state.risk || t.risk===state.risk) && (!state.status || taskStatus(t)===state.status));
  }
  function aggregate(tasks){
    const totalBudget = tasks.reduce((a,t)=>a+t.budget,0), totalPaid=tasks.reduce((a,t)=>a+t.paid,0), remaining=tasks.reduce((a,t)=>a+t.remaining,0), savings=tasks.reduce((a,t)=>a+t.savings,0);
    return { totalBudget, totalPaid, remaining, savings, spendRatio: totalBudget ? totalPaid/totalBudget : 0 };
  }
  function showError(message){ const toast=el('errorToast'); toast.textContent=message; toast.hidden=false; setTimeout(()=> toast.hidden=true, 7000); el('dataStatus').innerHTML='<span class="dot" style="background:#e06756"></span> خطأ في البيانات'; }
  async function loadData(force=false){
    try {
      const res = await fetch('/api/data?ts='+Date.now(), {cache:'no-store'});
      if(!res.ok) throw new Error('Unable to read data.xlsx');
      const incoming = await res.json();
      if(incoming.error) throw new Error(incoming.error);
      if(!force && currentVersion === incoming.version) return;
      dashboard = incoming; currentVersion = incoming.version;
      el('lastUpdated').textContent = new Date(incoming.loaded_at).toLocaleString('en-GB');
      el('dataStatus').innerHTML = '<span class="dot"></span> متصل بـ OneDrive';
      buildFilters(); renderAll();
    } catch(err){ showError('تعذر تحميل ملف Excel من OneDrive: ' + err.message); }
  }
  function buildFilters(){
    const category = el('categoryFilter'), risk = el('riskFilter');
    const catValue=category.value, riskValue=risk.value;
    category.innerHTML='<option value="">جميع التصنيفات</option>' + [...new Set(dashboard.tasks.map(t=>t.category))].sort().map(v=>`<option>${v}</option>`).join('');
    risk.innerHTML='<option value="">جميع مستويات المخاطر</option>' + [...new Set(dashboard.tasks.map(t=>t.risk))].map(v=>`<option>${v}</option>`).join('');
    category.value=catValue; risk.value=riskValue;
  }
  function renderAll(){
    const tasks = filteredTasks();
    const filtered = aggregate(tasks);
    const full = dashboard.overview;
    el('filterSummary').textContent = tasks.length === dashboard.tasks.length ? `عرض جميع بنود المشروع (${tasks.length} بندًا)` : `عرض ${tasks.length} من أصل ${dashboard.tasks.length} بندًا`;
    renderKpis(filtered, full, tasks.length !== dashboard.tasks.length);
    renderGauge(filtered);
    renderMonthly(tasks);
    renderCumulative(tasks);
    renderCash();
    renderUtilization(tasks);
    renderRisk(tasks);
    renderTreemap(tasks);
    renderScatter(tasks);
    renderMisc();
    renderHeatmap(tasks);
    renderTable(tasks);
    renderSelected();
  }
  function renderKpis(s, full, isFiltered){
    const cards = [
      ['إجمالي الميزانية', s.totalBudget, 'ج.م', ''], ['إجمالي المصروفات', s.totalPaid, 'ج.م', ''], ['إجمالي المتبقي', s.remaining, 'ج.م', ''], ['نسبة الصرف', s.spendRatio, '%', s.spendRatio>0.8?'alert':''],
      ['رأس المال المسدد', full.capital_paid, 'ج.م', ''], ['النقد المتبقي', full.cash_remaining, 'ج.م', full.cash_remaining<0?'alert':''], ['إجمالي التوفير', s.savings, 'ج.م', s.savings<0?'alert':''], ['عدد البنود المعروضة', isFiltered ? filteredTasks().length : dashboard.tasks.length, 'بند', '']
    ];
    el('kpis').innerHTML = cards.map(([label,val,unit,cls]) => `<article class="card kpi ${cls}"><span class="label">${label}</span><strong class="value">${unit==='%' ? pct(val) : (unit==='بند' ? val : money(val))}</strong><span class="unit">${unit==='%' ? 'من الميزانية المعروضة' : unit==='بند' ? 'بنود العمل' : 'جنيه مصري'}</span></article>`).join('');
  }
  function renderGauge(s){
    Plotly.react('gaugeChart', [{ type:'indicator', mode:'gauge+number', value:s.spendRatio*100, number:{suffix:'%', font:{size:38, color:COLORS.forest}}, gauge:{axis:{range:[0,110], ticksuffix:'%'}, bar:{color:s.spendRatio>0.8?COLORS.red:COLORS.leaf}, bgcolor:'#f1ede2', steps:[{range:[0,60],color:'#e6f0e6'},{range:[60,80],color:'#f8efd8'},{range:[80,100],color:'#fae5df'},{range:[100,110],color:'#f4cfca'}], threshold:{value:100, thickness:.7, line:{color:COLORS.red,width:3}}}}], {...layoutBase, margin:{t:15,r:30,b:10,l:30}}, config);
  }
  function monthSeries(tasks){ return dashboard.months.map(month => ({month, actual:tasks.reduce((a,t)=>a+(t.monthly[month]?.actual||0),0), planned:tasks.reduce((a,t)=>a+(t.monthly[month]?.planned||0),0)})); }
  function renderMonthly(tasks){
    const series=monthSeries(tasks); Plotly.react('monthlyChart', [
      {type:'bar', name:'الفعلي', x:series.map(m=>m.month), y:series.map(m=>m.actual), marker:{color:COLORS.leaf}, hovertemplate:'%{x}<br>الفعلي: %{y:,.0f} ج.م<extra></extra>'},
      {type:'bar', name:'المخطط', x:series.map(m=>m.month), y:series.map(m=>m.planned), marker:{color:COLORS.gold2}, hovertemplate:'%{x}<br>المخطط: %{y:,.0f} ج.م<extra></extra>'}
    ], {...layoutBase, barmode:'group', yaxis:{...layoutBase.yaxis, title:'جنيه مصري', tickformat:',.0f'}, xaxis:{...layoutBase.xaxis, tickangle:-25}}, config);
  }
  function renderCumulative(tasks){
    const series=monthSeries(tasks); let a=0,p=0; const actual=series.map(m=>a+=m.actual), planned=series.map(m=>p+=m.planned);
    Plotly.react('cumulativeChart', [
      {type:'scatter', mode:'lines+markers', name:'التراكمي الفعلي', x:series.map(m=>m.month), y:actual, line:{color:COLORS.leaf,width:4,shape:'spline'}, fill:'tozeroy', fillcolor:'rgba(60,116,80,.10)', hovertemplate:'%{x}<br>%{y:,.0f} ج.م<extra></extra>'},
      {type:'scatter', mode:'lines+markers', name:'التراكمي المخطط', x:series.map(m=>m.month), y:planned, line:{color:COLORS.gold,width:3,dash:'dash'}, hovertemplate:'%{x}<br>%{y:,.0f} ج.م<extra></extra>'}
    ], {...layoutBase, hovermode:'x unified', yaxis:{...layoutBase.yaxis,tickformat:',.0f'}, xaxis:{...layoutBase.xaxis,tickangle:-25}}, config);
  }
  function renderCash(){
    const o=dashboard.overview; Plotly.react('cashChart', [{type:'waterfall', orientation:'v', x:['رأس المال المسدد','المصروفات','النقد المتبقي'], y:[o.capital_paid,-o.total_paid,o.cash_remaining], measure:['absolute','relative','total'], connector:{line:{color:'#cfc6b5'}}, increasing:{marker:{color:COLORS.leaf}}, decreasing:{marker:{color:COLORS.red}}, totals:{marker:{color:COLORS.gold}}, hovertemplate:'%{x}<br>%{y:,.0f} ج.م<extra></extra>'}], {...layoutBase, margin:{t:20,r:15,b:55,l:55}, showlegend:false, yaxis:{...layoutBase.yaxis,tickformat:',.0f'}}, config);
  }
  function selectTask(name){ state.selected = dashboard.tasks.find(t=>t.name===name) || null; renderSelected(); if(state.selected) el('selectedCard').scrollIntoView({behavior:'smooth',block:'nearest'}); }
  function renderUtilization(tasks){
    const data=[...tasks].sort((a,b)=>b.budget-a.budget).slice(0,12).reverse();
    Plotly.react('utilizationChart', [
      {type:'bar', orientation:'h', name:'المدفوع', y:data.map(t=>t.name), x:data.map(t=>t.paid), marker:{color:COLORS.leaf}, customdata:data.map(t=>t.name), hovertemplate:'%{y}<br>المدفوع: %{x:,.0f} ج.م<extra></extra>'},
      {type:'bar', orientation:'h', name:'المتبقي', y:data.map(t=>t.name), x:data.map(t=>Math.max(t.remaining,0)), marker:{color:'#ded6c4'}, customdata:data.map(t=>t.name), hovertemplate:'%{y}<br>المتبقي: %{x:,.0f} ج.م<extra></extra>'}
    ], {...layoutBase, barmode:'stack', margin:{t:10,r:16,b:48,l:220}, xaxis:{...layoutBase.xaxis,tickformat:',.0f'}}, config).then(()=>{ const c=el('utilizationChart'); c.removeAllListeners?.('plotly_click'); c.on('plotly_click', ev => selectTask(ev.points[0].customdata)); });
  }
  function renderRisk(tasks){
    const grouped={}; tasks.forEach(t=> grouped[t.risk]=(grouped[t.risk]||0)+t.paid); const labels=Object.keys(grouped), vals=labels.map(x=>grouped[x]);
    Plotly.react('riskChart', [{type:'pie', hole:.62, labels, values:vals, marker:{colors:labels.map(l=>RISK_COLORS[l])}, textinfo:'percent', hovertemplate:'%{label}<br>%{value:,.0f} ج.م<br>%{percent}<extra></extra>'}], {...layoutBase, margin:{t:10,r:20,b:32,l:20}, legend:{orientation:'h', y:-.04}}, config).then(()=>{ const c=el('riskChart'); c.removeAllListeners?.('plotly_click'); c.on('plotly_click', ev=>{ state.risk=ev.points[0].label; el('riskFilter').value=state.risk; renderAll(); }); });
  }
  function renderTreemap(tasks){
    const labels=['الميزانية']; const parents=['']; const values=[tasks.reduce((a,t)=>a+t.budget,0)]; const custom=[''];
    [...new Set(tasks.map(t=>t.category))].forEach(cat=>{ const group=tasks.filter(t=>t.category===cat); labels.push(cat); parents.push('الميزانية'); values.push(group.reduce((a,t)=>a+t.budget,0)); custom.push('category:'+cat); group.forEach(t=>{ labels.push(t.name); parents.push(cat); values.push(t.budget); custom.push('task:'+t.name); }); });
    Plotly.react('treemapChart', [{type:'treemap', labels, parents, values, customdata:custom, branchvalues:'total', marker:{colorscale:[[0,'#dfe9dc'],[.6,COLORS.leaf2],[1,COLORS.forest]]}, hovertemplate:'%{label}<br>%{value:,.0f} ج.م<extra></extra>'}], {...layoutBase, margin:{t:8,r:8,b:8,l:8}}, config).then(()=>{ const c=el('treemapChart'); c.removeAllListeners?.('plotly_click'); c.on('plotly_click', ev=>{ const value=ev.points[0].customdata||''; if(value.startsWith('task:')) selectTask(value.slice(5)); else if(value.startsWith('category:')) {state.category=value.slice(9); el('categoryFilter').value=state.category; renderAll();} }); });
  }
  function renderScatter(tasks){
    const valid=tasks.filter(t=>t.budget>0); Plotly.react('scatterChart', [{type:'scatter', mode:'markers', x:valid.map(t=>t.progress*100), y:valid.map(t=>t.spend_ratio*100), text:valid.map(t=>t.name), customdata:valid.map(t=>t.name), marker:{size:valid.map(t=>Math.max(12,Math.sqrt(t.budget)/20)), sizemode:'diameter', color:valid.map(t=>RISK_COLORS[t.risk]), opacity:.82, line:{color:'#fff',width:1}}, hovertemplate:'%{text}<br>الإنجاز: %{x:.1f}%<br>الصرف: %{y:.1f}%<extra></extra>'}], {...layoutBase, margin:{t:12,r:20,b:55,l:58}, xaxis:{...layoutBase.xaxis,title:'نسبة الإنجاز (%)',range:[-4,105]}, yaxis:{...layoutBase.yaxis,title:'نسبة الصرف (%)'}, shapes:[{type:'line',x0:0,x1:100,y0:0,y1:100,line:{color:COLORS.gold,dash:'dot'}}]}, config).then(()=>{ const c=el('scatterChart'); c.removeAllListeners?.('plotly_click'); c.on('plotly_click', ev=>selectTask(ev.points[0].customdata)); });
  }
  function renderMisc(){
    const m=dashboard.misc; Plotly.react('miscChart', [{type:'bar', x:m.map(x=>x.name), y:m.map(x=>x.paid), marker:{color:COLORS.purple}, hovertemplate:'%{x}<br>%{y:,.0f} ج.م<extra></extra>'}], {...layoutBase,margin:{t:15,r:18,b:80,l:55},xaxis:{...layoutBase.xaxis,tickangle:-20},yaxis:{...layoutBase.yaxis,tickformat:',.0f'},showlegend:false}, config);
  }
  function renderHeatmap(tasks){
    const top=[...tasks].sort((a,b)=>b.paid-a.paid).slice(0,12); const z=top.map(t=>dashboard.months.map(m=>t.monthly[m]?.actual || 0));
    Plotly.react('heatmapChart', [{type:'heatmap', x:dashboard.months, y:top.map(t=>t.name), z, colorscale:[[0,'#f6f2e9'],[.25,'#d5c495'],[.55,'#7d9c70'],[1,COLORS.forest]], hovertemplate:'%{y}<br>%{x}<br>%{z:,.0f} ج.م<extra></extra>'}], {...layoutBase, margin:{t:12,r:15,b:56,l:230}, xaxis:{...layoutBase.xaxis,tickangle:-25}, yaxis:{automargin:true}}, config);
  }
  function renderTable(tasks){
    el('tasksBody').innerHTML=tasks.sort((a,b)=>b.budget-a.budget).map(t=>`<tr data-task="${encodeURIComponent(t.name)}"><td><strong>${t.name}</strong></td><td>${t.category}</td><td class="num">${money(t.budget)}</td><td class="num">${money(t.paid)}</td><td class="num">${money(t.remaining)}</td><td class="num">${pct(t.spend_ratio)}</td><td><span>${pct(t.progress)}</span><div class="progress"><i style="width:${Math.min(t.progress*100,100)}%"></i></div></td><td><span class="risk-pill ${riskClass(t.risk)}">${t.risk}</span></td></tr>`).join('');
    el('tasksBody').querySelectorAll('tr').forEach(row=>row.addEventListener('click',()=>selectTask(decodeURIComponent(row.dataset.task))));
  }
  function renderSelected(){
    const card=el('selectedCard'), t=state.selected; if(!t){card.hidden=true; return;} card.hidden=false;
    el('selectedContent').innerHTML=`<div class="selected-grid"><div class="selected-name">${t.name}<br><span class="risk-pill ${riskClass(t.risk)}" style="margin-top:9px">${t.risk}</span></div><div class="selected-stat"><span>الميزانية</span><strong>${money(t.budget)}</strong></div><div class="selected-stat"><span>المدفوع</span><strong>${money(t.paid)}</strong></div><div class="selected-stat"><span>نسبة الصرف</span><strong>${pct(t.spend_ratio)}</strong></div><div class="selected-stat"><span>نسبة الإنجاز</span><strong>${pct(t.progress)}</strong></div></div>`;
  }
  function bindControls(){
    el('refreshBtn').addEventListener('click',()=>loadData(true));
    el('searchFilter').addEventListener('input',e=>{state.search=e.target.value;renderAll();});
    el('categoryFilter').addEventListener('change',e=>{state.category=e.target.value;renderAll();});
    el('riskFilter').addEventListener('change',e=>{state.risk=e.target.value;renderAll();});
    el('statusFilter').addEventListener('change',e=>{state.status=e.target.value;renderAll();});
    el('clearFilters').addEventListener('click',()=>{Object.assign(state,{search:'',category:'',risk:'',status:'',selected:null}); ['searchFilter','categoryFilter','riskFilter','statusFilter'].forEach(id=>el(id).value=''); renderAll();});
    el('closeDetail').addEventListener('click',()=>{state.selected=null;renderSelected();});
  }
  window.addEventListener('DOMContentLoaded',()=>{ bindControls(); loadData(true); setInterval(()=>loadData(false),60000); });
})();
