/* ─── STATE ─────────────────────────────────────────── */
let currentMonth = new Date().toISOString().slice(0,7);
let currentType  = 'expense';
let categories   = [];
let charts       = {};

/* ─── INIT ──────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', async () => {
  updateMonthLabel();
  await loadCategories();
  switchSection('dashboard');
  document.querySelectorAll('.nav-btn').forEach(b => {
    b.addEventListener('click', () => {
      switchSection(b.dataset.section);
      closeSidebar();
    });
  });
  document.getElementById('openAddBtn').addEventListener('click', openAddModal);
  document.getElementById('exportExcel').addEventListener('click', exportExcel);
});

/* ─── NAVIGATION ────────────────────────────────────── */
function switchSection(name) {
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.toggle('active', b.dataset.section === name));
  document.getElementById(`sec-${name}`).classList.add('active');
  const titles = { dashboard:'Dashboard', transactions:'Movimientos', reports:'Reportes', categories:'Categorías' };
  document.getElementById('topbarTitle').textContent = titles[name];
  if (name === 'dashboard')    loadDashboard();
  if (name === 'transactions') loadTransactions();
  if (name === 'reports')      loadReports();
  if (name === 'categories')   renderCategories();
}

/* ─── MONTH ─────────────────────────────────────────── */
function updateMonthLabel() {
  const [y, m] = currentMonth.split('-');
  const names  = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  document.getElementById('monthLabel').textContent = `${names[+m-1]} ${y}`;
}
function changeMonth(delta) {
  const d = new Date(currentMonth + '-01');
  d.setMonth(d.getMonth() + delta);
  currentMonth = d.toISOString().slice(0,7);
  updateMonthLabel();
  const active = document.querySelector('.nav-btn.active')?.dataset.section;
  if (active === 'dashboard')    loadDashboard();
  if (active === 'transactions') loadTransactions();
  if (active === 'reports')      loadReports();
}

/* ─── SIDEBAR ───────────────────────────────────────── */
function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('overlay').classList.toggle('active');
}
function closeSidebar() {
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('overlay').classList.remove('active');
}

/* ─── CATEGORIES ────────────────────────────────────── */
async function loadCategories() {
  const res = await fetch('/api/categories');
  categories = await res.json();
  populateCategorySelects();
}
function populateCategorySelects() {
  const type = document.getElementById('editId').value
    ? (document.querySelector('.type-tab.active')?.dataset.type || currentType)
    : currentType;
  const fCat  = document.getElementById('fCategory');
  const fkCat = document.getElementById('filterCat');
  fCat.innerHTML  = '';
  fkCat.innerHTML = '<option value="">Todas las categorías</option>';
  categories
    .filter(c => c.type === type || c.type === 'both')
    .forEach(c => {
      const o = new Option(`${c.icon} ${c.name}`, c.name);
      fCat.appendChild(o.cloneNode(true));
    });
  categories.forEach(c => {
    const o = new Option(`${c.icon} ${c.name}`, c.name);
    fkCat.appendChild(o);
  });
}
function renderCategories() {
  const grid = document.getElementById('catGrid');
  grid.innerHTML = '';
  categories.forEach(c => {
    const div = document.createElement('div');
    div.className = 'cat-item';
    div.innerHTML = `
      <span class="cat-item-icon">${c.icon}</span>
      <div class="cat-item-info">
        <div class="cat-item-name">${c.name}</div>
        <div class="cat-item-type">${c.type==='income'?'Ingreso':c.type==='expense'?'Gasto':'Ambos'}</div>
      </div>
      <button class="cat-del" onclick="deleteCategory(${c.id})" title="Eliminar">✕</button>`;
    grid.appendChild(div);
  });
}
async function addCategory() {
  const name  = document.getElementById('catName').value.trim();
  const type  = document.getElementById('catType').value;
  const icon  = document.getElementById('catIcon').value.trim() || '📁';
  const color = document.getElementById('catColor').value;
  if (!name) { showToast('Ingresa un nombre de categoría','error'); return; }
  const res = await fetch('/api/categories', {
    method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({name, type, icon, color})
  });
  if (res.ok) {
    document.getElementById('catName').value = '';
    document.getElementById('catIcon').value = '';
    await loadCategories();
    renderCategories();
    showToast('Categoría agregada ✓');
  } else {
    const d = await res.json();
    showToast(d.error || 'Error al agregar','error');
  }
}
async function deleteCategory(id) {
  if (!confirm('¿Eliminar esta categoría?')) return;
  await fetch(`/api/categories/${id}`, {method:'DELETE'});
  await loadCategories();
  renderCategories();
  showToast('Categoría eliminada');
}

/* ─── DASHBOARD ─────────────────────────────────────── */
async function loadDashboard() {
  const [summary, trend, daily] = await Promise.all([
    fetch(`/api/stats/summary?month=${currentMonth}`).then(r=>r.json()),
    fetch('/api/stats/monthly_trend').then(r=>r.json()),
    fetch(`/api/stats/daily_trend?month=${currentMonth}`).then(r=>r.json()),
  ]);
  renderSummaryCards(summary);
  renderTrendChart(trend);
  renderDailyChart(daily);
  loadPieChart('expense', document.querySelector('.tog.active'));
}
function renderSummaryCards(s) {
  const fmt = v => 'S/ ' + Math.abs(v).toLocaleString('es-PE',{minimumFractionDigits:2,maximumFractionDigits:2});
  const [y, m] = currentMonth.split('-');
  const months = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  document.getElementById('summaryCards').innerHTML = `
    <div class="card income">
      <div class="card-label">Ingresos</div>
      <div class="card-amount">${fmt(s.income)}</div>
      <div class="card-sub">${months[+m-1]} ${y}</div>
      <div class="card-icon">💰</div>
    </div>
    <div class="card expense">
      <div class="card-label">Gastos</div>
      <div class="card-amount">${fmt(s.expense)}</div>
      <div class="card-sub">${months[+m-1]} ${y}</div>
      <div class="card-icon">💸</div>
    </div>
    <div class="card balance">
      <div class="card-label">Balance</div>
      <div class="card-amount" style="color:${s.balance>=0?'var(--income)':'var(--expense)'}">${s.balance>=0?'+':''}${fmt(s.balance)}</div>
      <div class="card-sub">Disponible</div>
      <div class="card-icon">⚖️</div>
    </div>`;
}
function renderTrendChart(data) {
  const ctx = document.getElementById('trendChart').getContext('2d');
  if (charts.trend) charts.trend.destroy();
  charts.trend = new Chart(ctx, {
    type: 'line',
    data: {
      labels: data.months.map(m => { const [y,mo]=m.split('-'); return ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'][+mo-1]+' '+y.slice(2); }),
      datasets: [
        { label:'Ingresos', data: data.income,  borderColor:'#10b981', backgroundColor:'rgba(16,185,129,.08)', tension:.35, fill:true, pointRadius:4, pointHoverRadius:6 },
        { label:'Gastos',   data: data.expense, borderColor:'#f43f5e', backgroundColor:'rgba(244,63,94,.08)',  tension:.35, fill:true, pointRadius:4, pointHoverRadius:6 },
      ]
    },
    options: chartOpts()
  });
}
async function loadPieChart(type, btn) {
  if (btn) {
    document.querySelectorAll('.tog').forEach(t=>t.classList.remove('active'));
    btn.classList.add('active');
  }
  const data = await fetch(`/api/stats/by_category?month=${currentMonth}&type=${type}`).then(r=>r.json());
  const ctx  = document.getElementById('pieChart').getContext('2d');
  if (charts.pie) charts.pie.destroy();
  const colors = ['#10b981','#3b82f6','#8b5cf6','#f59e0b','#ef4444','#ec4899','#14b8a6','#f97316','#6366f1','#a855f7','#22c55e','#0ea5e9'];
  charts.pie = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: data.map(d=>d.category),
      datasets: [{ data: data.map(d=>d.total), backgroundColor: colors, hoverOffset:6, borderWidth:2, borderColor:'#111827' }]
    },
    options: { ...pieOpts(), plugins: { legend:{ position:'right', labels:{ color:'#94a3b8', font:{size:11} } } } }
  });
}
function renderDailyChart(data) {
  const ctx = document.getElementById('dailyChart').getContext('2d');
  if (charts.daily) charts.daily.destroy();
  charts.daily = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: data.dates.map(d=>d.slice(5)),
      datasets: [
        { label:'Ingresos', data:data.income,  backgroundColor:'rgba(16,185,129,.7)',  borderRadius:4 },
        { label:'Gastos',   data:data.expense, backgroundColor:'rgba(244,63,94,.7)',   borderRadius:4 },
      ]
    },
    options: { ...chartOpts(), scales: { x:{...barXScale()}, y:{...barYScale()} } }
  });
}

/* ─── REPORTS ───────────────────────────────────────── */
async function loadReports() {
  const [catExp, catInc, trend] = await Promise.all([
    fetch(`/api/stats/by_category?month=${currentMonth}&type=expense`).then(r=>r.json()),
    fetch(`/api/stats/by_category?month=${currentMonth}&type=income`).then(r=>r.json()),
    fetch('/api/stats/monthly_trend').then(r=>r.json()),
  ]);
  renderBarChart('barCatChart', catExp, 'barCat', 'rgba(244,63,94,.7)');
  renderBarChart('barIncChart', catInc, 'barInc', 'rgba(16,185,129,.7)');
  renderAccumChart(trend);
}
function renderBarChart(id, data, key, color) {
  const ctx = document.getElementById(id).getContext('2d');
  if (charts[key]) charts[key].destroy();
  charts[key] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: data.map(d=>d.category),
      datasets: [{ label:'S/', data:data.map(d=>d.total), backgroundColor:color, borderRadius:6 }]
    },
    options: { ...chartOpts(), indexAxis:'y', scales:{ x:{...barYScale()}, y:{...barXScale()} } }
  });
}
function renderAccumChart(data) {
  const ctx = document.getElementById('accumChart').getContext('2d');
  if (charts.accum) charts.accum.destroy();
  let accum = 0;
  const accumData = data.months.map((_, i) => { accum += data.income[i] - data.expense[i]; return +accum.toFixed(2); });
  charts.accum = new Chart(ctx, {
    type: 'line',
    data: {
      labels: data.months.map(m => { const [y,mo]=m.split('-'); return ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'][+mo-1]+' '+y.slice(2); }),
      datasets: [{
        label:'Balance acumulado', data:accumData,
        borderColor:'#3b82f6', backgroundColor:'rgba(59,130,246,.08)',
        tension:.35, fill:true, pointRadius:4
      }]
    },
    options: chartOpts()
  });
}

/* ─── TRANSACTIONS ──────────────────────────────────── */
async function loadTransactions() {
  const params = new URLSearchParams();
  const search = document.getElementById('searchInput')?.value.trim();
  const type   = document.getElementById('filterType')?.value;
  const cat    = document.getElementById('filterCat')?.value;
  const from   = document.getElementById('filterFrom')?.value;
  const to     = document.getElementById('filterTo')?.value;
  if (search) params.set('search', search);
  if (type)   params.set('type',   type);
  if (cat)    params.set('category', cat);
  if (from)   params.set('date_from', from);
  if (to)     params.set('date_to', to);
  const txs = await fetch('/api/transactions?' + params).then(r=>r.json());
  const list = document.getElementById('txList');
  if (!txs.length) {
    list.innerHTML = '<div class="empty-state"><div class="big">📭</div><div>Sin movimientos encontrados</div></div>';
    return;
  }
  list.innerHTML = txs.map(tx => {
    const cat  = categories.find(c=>c.name===tx.category);
    const icon = cat?.icon || (tx.type==='income'?'💰':'💸');
    const fmt  = v => 'S/ ' + v.toLocaleString('es-PE',{minimumFractionDigits:2,maximumFractionDigits:2});
    return `
    <div class="tx-item">
      <div class="tx-icon ${tx.type}">${icon}</div>
      <div class="tx-info">
        <div class="tx-cat">${tx.category}</div>
        <div class="tx-desc">${tx.description || '—'}</div>
      </div>
      <div class="tx-date">${tx.date}</div>
      <div class="tx-amount ${tx.type}">${tx.type==='income'?'+':'-'}${fmt(tx.amount)}</div>
      <div class="tx-actions">
        <button class="tx-btn edit" onclick='openEditModal(${JSON.stringify(tx)})' title="Editar">✏</button>
        <button class="tx-btn del"  onclick="deleteTransaction(${tx.id})"         title="Eliminar">🗑</button>
      </div>
    </div>`;
  }).join('');
}
async function deleteTransaction(id) {
  if (!confirm('¿Eliminar este movimiento?')) return;
  await fetch(`/api/transactions/${id}`, {method:'DELETE'});
  showToast('Movimiento eliminado');
  loadTransactions();
  if (document.getElementById('sec-dashboard').classList.contains('active')) loadDashboard();
}

/* ─── MODAL ─────────────────────────────────────────── */
function openAddModal() {
  document.getElementById('editId').value = '';
  document.getElementById('fAmount').value      = '';
  document.getElementById('fDescription').value = '';
  document.getElementById('fDate').value = new Date().toISOString().slice(0,10);
  document.getElementById('modalTitle').textContent = 'Nueva transacción';
  document.getElementById('submitBtn').textContent  = 'Guardar';
  setType('expense', document.querySelector('[data-type="expense"]'));
  document.getElementById('modalBg').classList.add('active');
}
function openEditModal(tx) {
  document.getElementById('editId').value       = tx.id;
  document.getElementById('fAmount').value      = tx.amount;
  document.getElementById('fDescription').value = tx.description || '';
  document.getElementById('fDate').value        = tx.date;
  document.getElementById('modalTitle').textContent = 'Editar movimiento';
  document.getElementById('submitBtn').textContent  = 'Actualizar';
  setType(tx.type, document.querySelector(`[data-type="${tx.type}"]`));
  setTimeout(() => { document.getElementById('fCategory').value = tx.category; }, 50);
  document.getElementById('modalBg').classList.add('active');
}
function closeModal() { document.getElementById('modalBg').classList.remove('active'); }
document.getElementById('modalBg').addEventListener('click', e => { if (e.target === e.currentTarget) closeModal(); });

function setType(type, btn) {
  currentType = type;
  document.querySelectorAll('.type-tab').forEach(t => t.classList.remove('active'));
  if (btn) btn.classList.add('active');
  populateCategorySelects();
}
async function submitTransaction() {
  const id     = document.getElementById('editId').value;
  const amount = parseFloat(document.getElementById('fAmount').value);
  const cat    = document.getElementById('fCategory').value;
  const date   = document.getElementById('fDate').value;
  const desc   = document.getElementById('fDescription').value.trim();
  if (!amount || amount <= 0) { showToast('Ingresa un monto válido','error'); return; }
  if (!cat)  { showToast('Selecciona una categoría','error'); return; }
  if (!date) { showToast('Selecciona una fecha','error'); return; }
  const body = { type:currentType, amount, category:cat, date, description:desc };
  const url    = id ? `/api/transactions/${id}` : '/api/transactions';
  const method = id ? 'PUT' : 'POST';
  const res = await fetch(url, { method, headers:{'Content-Type':'application/json'}, body:JSON.stringify(body) });
  if (res.ok) {
    closeModal();
    showToast(id ? 'Movimiento actualizado ✓' : 'Movimiento registrado ✓');
    const active = document.querySelector('.nav-btn.active')?.dataset.section;
    if (active === 'transactions') loadTransactions();
    if (active === 'dashboard')    loadDashboard();
    if (active === 'reports')      loadReports();
  } else {
    showToast('Error al guardar','error');
  }
}

/* ─── EXPORT ────────────────────────────────────────── */
function exportExcel() {
  showToast('Generando Excel…');
  window.location.href = '/api/export/excel';
}

/* ─── TOAST ─────────────────────────────────────────── */
function showToast(msg, type='') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className   = 'toast show' + (type ? ' ' + type : '');
  setTimeout(() => t.classList.remove('show'), 2800);
}

/* ─── CHART HELPERS ─────────────────────────────────── */
const gridColor  = 'rgba(255,255,255,.05)';
const labelColor = '#64748b';
function chartOpts() {
  return {
    responsive: true,
    plugins: {
      legend:  { labels:{ color:labelColor, font:{size:11} } },
      tooltip: { backgroundColor:'#1e2740', titleColor:'#f1f5f9', bodyColor:'#94a3b8',
                 callbacks:{ label: ctx => ' S/ '+ctx.parsed.y?.toLocaleString('es-PE',{minimumFractionDigits:2})||'' } }
    },
    scales: {
      x: { grid:{color:gridColor}, ticks:{color:labelColor, font:{size:10}} },
      y: { grid:{color:gridColor}, ticks:{color:labelColor, font:{size:10},
             callback: v => 'S/'+v.toLocaleString('es-PE',{maximumFractionDigits:0}) } }
    }
  };
}
function pieOpts() {
  return {
    responsive:true,
    plugins:{
      legend:{ labels:{ color:labelColor, font:{size:11} } },
      tooltip:{ backgroundColor:'#1e2740', titleColor:'#f1f5f9', bodyColor:'#94a3b8',
                callbacks:{ label: ctx => ' S/ '+ctx.parsed.toLocaleString('es-PE',{minimumFractionDigits:2}) } }
    }
  };
}
function barXScale() { return { grid:{color:gridColor}, ticks:{color:labelColor, font:{size:10}} }; }
function barYScale() {
  return { grid:{color:gridColor}, ticks:{color:labelColor, font:{size:10},
           callback: v => 'S/'+Number(v).toLocaleString('es-PE',{maximumFractionDigits:0}) } };
}
