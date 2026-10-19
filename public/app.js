// ============================================================
// KONFIGURASI UPLOAD FOTO CLOUDINARY
// Isi CLOUDINARY_CLOUD_NAME dan CLOUDINARY_UPLOAD_PRESET dari
// Dashboard Cloudinary. Gunakan preset dengan mode Unsigned.
// Jangan pernah memasukkan API Secret ke file frontend ini.
// ============================================================
const CLOUDINARY_CLOUD_NAME = '';
const CLOUDINARY_UPLOAD_PRESET = '';

// SIPOND - client-side. Semua data diambil/disimpan lewat API server

if(localStorage.getItem('SIPOND_theme') === 'dark') {
  document.body.classList.add('dark-mode');
}

let state = {
  items: [],
  transactions: [],
  selectedItemId: null,
  view: 'idle',
  confirmingReturnCode: null,
  editingCode: null,
  selectedIds: new Set(),
  selectedHistoryCodes: new Set(),
  lastOperator: localStorage.getItem('SIPOND_lastOperator') || '',
  showDiscontinued: false
};

const API = {
  items: '/api/items',
  import: '/api/import',
  transactions: '/api/transactions',
  loans: '/api/loans',
  returns: (code) => '/api/returns/' + encodeURIComponent(code)
};

async function apiGet(url){
  const res = await fetch(url);
  if(res.status === 401){ redirectToLogin(); throw new Error('Sesi berakhir.'); }
  if(!res.ok) throw new Error('Gagal memuat data dari server.');
  return res.json();
}
async function apiSend(url, method, body){
  const res = await fetch(url, {
    method,
    headers: body ? { 'Content-Type': 'application/json' } : undefined,
    body: body ? JSON.stringify(body) : undefined
  });
  if(res.status === 401){ redirectToLogin(); throw new Error('Sesi berakhir, silakan login lagi.'); }
  const data = await res.json().catch(() => ({}));
  if(!res.ok) throw new Error(data.error || 'Terjadi kesalahan pada server.');
  return data;
}

function redirectToLogin(){
  window.location.href = '/login.html';
}

async function checkAuthAndInit(){
  try{
    const res = await fetch('/api/me');
    if(!res.ok){ redirectToLogin(); return; }
    const data = await res.json();
    renderSessionBar(data.username);
    loadAll();
  }catch(e){
    redirectToLogin();
  }
}

function toggleTheme(){
  const isDark = document.body.classList.toggle('dark-mode');
  localStorage.setItem('SIPOND_theme', isDark ? 'dark' : 'light');
  const btn = document.getElementById('themeToggleBtn');
  if(btn) btn.innerHTML = isDark ? '☀️ Siang' : '🌙 Malam';
}

function renderSessionBar(username){
  const el = document.getElementById('sessionBar');
  if(!el) return;
  const isDark = document.body.classList.contains('dark-mode');
  const themeIcon = isDark ? '☀️ Siang' : '🌙 Malam';
  el.innerHTML = '<div style="display:flex; align-items:center; gap:8px;">'
    +'<span style="margin-right:8px;">Masuk sebagai <b>'+username+'</b></span>'
    +'<button class="logout-link" onclick="toggleTheme()" id="themeToggleBtn">'+themeIcon+'</button>'
    +'<button class="logout-link" onclick="doLogout()">Keluar</button>'
    +'</div>';
}

async function doLogout(){
  try{ await fetch('/api/logout', { method:'POST' }); }catch(e){ }
  redirectToLogin();
}

async function toggleShowDiscontinued(checked){
  state.showDiscontinued = checked;
  await loadAll();
}

async function loadAll(){
  try{
    const itemsUrl = API.items + (state.showDiscontinued ? '?includeDiscontinued=true' : '');
    const [items, transactions] = await Promise.all([
      apiGet(itemsUrl),
      apiGet(API.transactions)
    ]);
    state.items = items;
    state.transactions = transactions;
  }catch(e){
    console.error(e);
    showGlobalError('Tidak bisa terhubung ke server SIPOND. Pastikan server sedang berjalan.');
  }
  render();
}

function showGlobalError(msg){
  const el = document.getElementById('importStatus');
  if(el){ el.className = 'import-status err'; el.textContent = msg; }
}

/* ---------- Helpers ---------- */
function fmtDate(d){
  if(!d) return '-';
  const [y,m,day] = d.split('-');
  return day+'/'+m+'/'+y;
}
function todayStr(){ return new Date().toISOString().slice(0,10); }
function monthsAgoStr(n){
  const d = new Date();
  d.setMonth(d.getMonth()-n);
  return d.toISOString().slice(0,10);
}
function itemDetailLine(i){
  return 'Tempat '+i.tempat+' · Rak '+i.rak+' · Mata '+i.mata+' · Papan '+i.papan;
}

/* ---------- Export data master pisau ke Excel ---------- */
async function exportKnifeInventory(){
  try{
    const res = await fetch('/api/export-items');
    if(res.status === 401){ redirectToLogin(); return; }
    if(!res.ok){
      const err = await res.json().catch(()=>({}));
      throw new Error(err.error || 'Gagal mengekspor data pisau.');
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'data-pisau-sipond-' + todayStr() + '.xlsx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }catch(err){
    alert(err.message || 'Gagal mengekspor data pisau.');
  }
}

/* ---------- Import Excel ---------- */
const importEl = document.getElementById('importFile');
if(importEl) {
  importEl.addEventListener('change', async function(e){
    const file = e.target.files[0];
    if(!file) return;
    const statusEl = document.getElementById('importStatus');
    statusEl.className = 'import-status';
    statusEl.textContent = 'Mengunggah & membaca file...';
    try{
      const fd = new FormData();
      fd.append('file', file);
      const res = await fetch(API.import, { method:'POST', body: fd });
      const data = await res.json();
      if(!res.ok) throw new Error(data.error || 'Gagal mengimpor file.');
      statusEl.className = 'import-status ok';
      statusEl.textContent = data.added+' pisau ditambahkan dari '+data.sheets+' sheet (data yang sudah ada, termasuk yang mirip, tidak ditimpa).';
      await loadAll();
    }catch(err){
      statusEl.className = 'import-status err';
      statusEl.textContent = err.message;
    }
    e.target.value = '';
  });
}

/* ---------- Rendering ---------- */
function getFilteredItems(){
  const searchInput = document.getElementById('searchInput');
  const locFilter = document.getElementById('locationFilter');
  const custFilter = document.getElementById('customerFilter');

  const q = (searchInput ? searchInput.value : '').toLowerCase();
  const loc = locFilter ? locFilter.value : '';
  const cust = custFilter ? custFilter.value : '';
  
  let filtered = state.items.filter(i=>{
    const matchQ = !q
      || i.name.toLowerCase().includes(q)
      || i.location.toLowerCase().includes(q)
      || (i.customer||'').toLowerCase().includes(q);
    const matchLoc = !loc || loc==='Semua lokasi' || i.location===loc;
    const matchCust = !cust || cust==='Semua customer' || i.customer===cust;
    return matchQ && matchLoc && matchCust;
  });

  if (state.showDiscontinued) {
    filtered.sort((a, b) => {
      if (a.status === 'discontinued' && b.status !== 'discontinued') return -1;
      if (a.status !== 'discontinued' && b.status === 'discontinued') return 1;
      return 0;
    });
  }
  return filtered;
}

function renderStats(){
  const filtered = getFilteredItems();
  const filteredIds = new Set(filtered.map(i=>i.id));
  const total = filtered.reduce((s,i)=>s+i.stock,0);
  const tersedia = filtered.reduce((s,i)=>s+i.available,0);
  const dipinjamHariIni = state.transactions.filter(t=>t.date===todayStr() && t.status==='dipinjam' && filteredIds.has(t.itemId)).length;
  
  const statsContainer = document.getElementById('statsRow');
  if(!statsContainer) return;
  let numbersBox = document.getElementById('statsNumbers');
  
  if(!numbersBox) {
    statsContainer.innerHTML = '<div id="statsNumbers" style="display:contents;"></div>'
      + '<div class="stat" style="flex:1; min-width:200px; display:flex; flex-direction:column; background:var(--accent-soft); border-color:var(--accent);">'
      + '<span style="font-size:11px;color:var(--accent-ink);margin-bottom:4px;font-weight:700;">📌 Pengingat / Catatan</span>'
      + '<textarea id="globalNoteInput" placeholder="Ketik pengingat di sini (otomatis tersimpan)..." style="flex:1; border:none; background:transparent; resize:none; padding:0; font-size:12.5px; color:var(--ink); font-family:inherit; outline:none;" oninput="localStorage.setItem(\'SIPOND_global_note\', this.value)"></textarea>'
      + '</div>';
    
    numbersBox = document.getElementById('statsNumbers');
    const noteInput = document.getElementById('globalNoteInput');
    if(noteInput) noteInput.value = localStorage.getItem('SIPOND_global_note') || '';
  }

  numbersBox.innerHTML = [
    ['Jenis pisau', filtered.length],
    ['Stok tersedia', tersedia+' / '+total],
    ['Dipinjam hari ini', dipinjamHariIni]
  ].map(([label,val])=>'<div class="stat"><b>'+val+'</b><span>'+label+'</span></div>').join('');
}

function renderLocationFilter(){
  const sel = document.getElementById('locationFilter');
  if(!sel) return;
  const locs = ['Semua lokasi', ...new Set(state.items.map(i=>i.location))].sort();
  const current = sel.value || 'Semua lokasi';
  sel.innerHTML = locs.map(r=>'<option value="'+r+'">'+r+'</option>').join('');
  sel.value = locs.includes(current) ? current : 'Semua lokasi';
}

function renderCustomerFilter(){
  const sel = document.getElementById('customerFilter');
  if(!sel) return;
  const custs = ['Semua customer', ...new Set(state.items.map(i=>i.customer).filter(c=>c))].sort();
  const current = sel.value || 'Semua customer';
  sel.innerHTML = custs.map(c=>'<option value="'+c+'">'+c+'</option>').join('');
  sel.value = custs.includes(current) ? current : 'Semua customer';
}

function renderCatalog(){
  const list = document.getElementById('catalogList');
  if(!list) return;
  const filtered = getFilteredItems();
  if(filtered.length===0){
    list.innerHTML = '<div class="empty-state">Pisau tidak ditemukan. Coba ubah kata kunci/filter, atau tambah pisau baru.</div>';
    updateBulkBar();
    return;
  }
  list.innerHTML = filtered.map(i=>{
    const isDiscontinued = i.status === 'discontinued';
    const out = i.available <= 0 && !isDiscontinued;
    const sel = i.id === state.selectedItemId ? ' selected' : '';
    const cls = out ? ' out' : sel;
    
    const reasonText = i.discontinueReason ? ' ('+i.discontinueReason+')' : '';
    const badge = isDiscontinued 
      ? '<span class="badge" style="background:var(--line);color:var(--ink-soft);">Discontinued'+reasonText+'</span>'
      : (out ? '<span class="badge out">Dipinjam</span>' : '<span class="badge avail">Tersedia</span>');
    
    let actionBtns = '';
    if (isDiscontinued) {
      actionBtns = '<button class="btn secondary small" style="color:var(--ok); border-color:var(--ok); margin-right:6px;" onclick="event.stopPropagation(); reactivateItem(\''+i.id+'\')">Aktifkan</button>'
                 + '<button class="btn secondary small" style="color:var(--warn); border-color:var(--warn);" onclick="event.stopPropagation(); deleteItemPermanently(\''+i.id+'\')">Hapus</button>';
    } else {
      actionBtns = '<button class="btn secondary small" style="margin-right:6px;" onclick="event.stopPropagation(); viewItemHistory(\''+i.id+'\')">Riwayat / Edit</button>'
                 + '<button class="btn secondary small" style="color:var(--warn); border-color:var(--warn);" onclick="event.stopPropagation(); promptManageItem(\''+i.id+'\')">Hapus / Discon</button>';
    }

    const clickAttr = ' onclick="previewItemForLoan(\''+i.id+'\')" style="cursor:pointer;"';
    const checked = state.selectedIds.has(i.id) ? ' checked' : '';
    
    const noteHtml = i.note ? '<div style="font-size:11px; color:var(--ink-soft); margin-top:5px; font-style:italic;">Catatan: '+i.note+'</div>' : '';
    const photoBadge = i.imageUrl
      ? '<span class="photo-status-badge has-photo" title="Foto pisau sudah tersedia">📷 Ada foto</span>'
      : '<span class="photo-status-badge no-photo" title="Foto pisau belum ditambahkan">📷 Belum ada foto</span>';
    return '<div class="item-tag'+cls+'"'+clickAttr+'>'
      +'<input type="checkbox" class="item-checkbox"'+checked+' onclick="event.stopPropagation()" onchange="toggleItemCheckbox(\''+i.id+'\', this.checked)">'
      +'<div class="hole"></div>'
      +'<div class="info"><div class="item-name-row"><b>'+i.name+'</b>'+photoBadge+'</div><span>'+itemDetailLine(i)+(i.customer?' · Customer '+i.customer:'')+'</span>'+noteHtml+'</div>'
      +'<div class="nourut-box"><label>No Urut</label><div class="nourut-value">'+(i.manualNo||'-')+'</div></div>'
      +'<div class="side" style="align-items:flex-end;">'+badge+'<div style="display:flex; gap:6px; margin-top:8px;">'+actionBtns+'</div></div>'
      +'</div>';
  }).join('');
  updateBulkBar();
}

/* ---------- Upload foto pisau ke Cloudinary ---------- */
function chooseKnifePhoto(prefix){
  const input = document.getElementById(prefix + 'PhotoFile');
  if(input) input.click();
}

function previewSelectedKnifePhoto(prefix){
  const input = document.getElementById(prefix + 'PhotoFile');
  const status = document.getElementById(prefix + 'PhotoStatus');
  const preview = document.getElementById(prefix + 'PhotoPreview');
  const file = input && input.files && input.files[0];
  if(!file) return;
  if(status) status.textContent = 'Foto dipilih: ' + file.name + '. Tekan Upload Foto untuk mengunggah.';
  if(preview && file.type && file.type.startsWith('image/')) {
    preview.src = URL.createObjectURL(file);
    preview.style.display = 'block';
  }
}

async function uploadKnifePhoto(prefix){
  const input = document.getElementById(prefix + 'PhotoFile');
  const urlInput = document.getElementById(prefix + 'ImageUrl');
  const status = document.getElementById(prefix + 'PhotoStatus');
  const preview = document.getElementById(prefix + 'PhotoPreview');
  const file = input && input.files && input.files[0];
  if(!file) {
    if(status) status.textContent = 'Pilih foto terlebih dahulu.';
    return;
  }
  if(!file.type || !file.type.startsWith('image/')) {
    if(status) status.textContent = 'File harus berupa gambar.';
    input.value = '';
    return;
  }
  if(file.size > 10 * 1024 * 1024) {
    if(status) status.textContent = 'Ukuran foto maksimal 10 MB.';
    input.value = '';
    return;
  }
  if(!CLOUDINARY_CLOUD_NAME || CLOUDINARY_CLOUD_NAME === 'ISI_CLOUD_NAME_ANDA' ||
     !CLOUDINARY_UPLOAD_PRESET || CLOUDINARY_UPLOAD_PRESET === 'ISI_UNSIGNED_UPLOAD_PRESET_ANDA') {
    if(status) status.textContent = 'Konfigurasi Cloudinary belum diisi di awal file public/app.js.';
    alert('Isi CLOUDINARY_CLOUD_NAME dan CLOUDINARY_UPLOAD_PRESET di awal file public/app.js terlebih dahulu.');
    return;
  }

  const button = document.getElementById(prefix + 'UploadPhotoBtn');
  if(button) { button.disabled = true; button.textContent = 'Mengunggah...'; }
  if(status) status.textContent = 'Sedang mengunggah foto ke Cloudinary...';
  try {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('upload_preset', CLOUDINARY_UPLOAD_PRESET);
    formData.append('folder', 'sipond/pisau');
    const endpoint = 'https://api.cloudinary.com/v1_1/' +
      encodeURIComponent(CLOUDINARY_CLOUD_NAME) + '/image/upload';
    const response = await fetch(endpoint, { method: 'POST', body: formData });
    const result = await response.json().catch(() => ({}));
    if(!response.ok || !result.secure_url) {
      throw new Error(result.error && result.error.message ? result.error.message : 'Cloudinary menolak upload. Periksa Cloud Name dan Unsigned Upload Preset.');
    }
    urlInput.value = result.secure_url;
    if(preview) {
      preview.src = result.secure_url;
      preview.style.display = 'block';
    }
    if(status) status.textContent = 'Foto berhasil diunggah. Tekan Simpan untuk mengaitkan foto dengan data pisau.';
  } catch(error) {
    if(status) status.textContent = 'Upload gagal: ' + error.message;
  } finally {
    if(button) { button.disabled = false; button.textContent = 'Upload Foto'; }
    if(input) input.value = '';
  }
}

function previewKnifePhoto(prefix){
  const urlInput = document.getElementById(prefix + 'ImageUrl');
  const preview = document.getElementById(prefix + 'PhotoPreview');
  if(!urlInput || !preview) return;
  const url = urlInput.value.trim();
  if(url) {
    preview.src = url;
    preview.style.display = 'block';
  } else {
    preview.removeAttribute('src');
    preview.style.display = 'none';
  }
}

function toggleItemCheckbox(id, checked){
  if(checked) state.selectedIds.add(id); else state.selectedIds.delete(id);
  updateBulkBar();
}

function toggleSelectAll(checked){
  const filtered = getFilteredItems();
  filtered.forEach(i => {
    if(checked) state.selectedIds.add(i.id); else state.selectedIds.delete(i.id);
  });
  renderCatalog();
}

function updateBulkBar(){
  const filtered = getFilteredItems();
  const countEl = document.getElementById('bulkCount');
  const btn = document.getElementById('bulkDeleteBtn');
  const allCb = document.getElementById('selectAllCheckbox');
  const n = state.selectedIds.size;
  if(countEl) countEl.textContent = n>0 ? n+' pisau dipilih' : '';
  if(btn) btn.style.display = n>0 ? 'inline-block' : 'none';
  if(allCb) allCb.checked = filtered.length>0 && filtered.every(i=>state.selectedIds.has(i.id));
}

async function bulkDeleteSelected(){
  const ids = Array.from(state.selectedIds);
  if(ids.length===0) return;
  if(!confirm('Yakin ingin menghapus '+ids.length+' pisau yang dipilih dari katalog? Riwayat transaksinya tetap tersimpan.')) return;
  try{
    await apiSend('/api/items/bulk-delete', 'POST', { ids });
    state.selectedIds.clear();
    await loadAll();
  }catch(e){
    alert(e.message);
  }
}

function previewItemForLoan(id){
  const item = state.items.find(i => i.id === id);
  if(!item) return;
  state.selectedItemId = id;
  state.view = 'photo-preview';
  render();
}

function continueToLoanForm(){
  const item = state.items.find(i => i.id === state.selectedItemId);
  if(!item || item.available <= 0 || item.status === 'discontinued') return;
  state.view = 'form';
  render();
}

function selectItem(id){
  previewItemForLoan(id);
}

function viewItemHistory(id){
  state.selectedItemId = id;
  state.view = 'history';
  render();
}

function cancelSelection(){
  state.selectedItemId = null;
  state.view = 'idle';
  render();
}

function cancelPhotoPreview(){
  state.selectedItemId = null;
  state.view = 'idle';
  render();
}

function openAddItemForm(){
  state.selectedItemId = null;
  state.view = 'add-item';
  render();
}

function promptManageItem(id){
  state.selectedItemId = id;
  state.view = 'manage-item';
  render();
}

function openDiscontinueReport(){
  state.selectedItemId = null;
  state.view = 'discontinue-report';
  render();
}

async function deleteItemPermanently(id){
  const item = state.items.find(i=>i.id===id);
  if(!item) return;
  const activeLoan = item.available < item.stock;
  const msg = activeLoan
    ? 'Pisau "'+item.name+'" sedang dipinjam! Yakin ingin menghapus permanen? (Riwayat tetap ada)'
    : 'Yakin ingin menghapus pisau "'+item.name+'" secara PERMANEN? (Riwayat tetap ada)';
  if(!confirm(msg)) return;
  try{
    await apiSend(API.items+'/'+encodeURIComponent(id), 'DELETE');
    if(state.selectedItemId===id){ state.selectedItemId=null; state.view='idle'; }
    state.selectedIds.delete(id);
    await loadAll();
  }catch(e){ alert(e.message); }
}

async function discontinueItem(id){
  const reasonEl = document.getElementById('mDiscontinueReason');
  const reason = reasonEl ? reasonEl.value.trim() : '';
  try{
    await apiSend(API.items+'/'+encodeURIComponent(id)+'/discontinue', 'POST', { reason });
    if(state.selectedItemId===id){ state.selectedItemId=null; state.view='idle'; }
    state.selectedIds.delete(id);
    await loadAll();
  }catch(e){ alert(e.message); }
}

async function reactivateItem(id){
  try{
    await apiSend(API.items+'/'+encodeURIComponent(id)+'/reactivate', 'POST');
    if(state.selectedItemId===id){ state.selectedItemId=null; state.view='idle'; }
    await loadAll();
  }catch(e){ alert(e.message); }
}

async function submitAddItem(){
  const name = document.getElementById('aName').value.trim();
  const tempat = document.getElementById('aTempat').value.trim();
  const rak = document.getElementById('aRak').value.trim();
  const mata = document.getElementById('aMata').value.trim();
  const papan = document.getElementById('aPapan').value.trim();
  const customer = document.getElementById('aCustomer').value.trim();
  const manualNo = document.getElementById('aManualNo').value.trim();
  const note = document.getElementById('aNote').value.trim();
  const imageUrl = document.getElementById('aImageUrl').value.trim();
  const errEl = document.getElementById('addItemErr');

  if(!name){ errEl.textContent='Nama pisau wajib diisi.'; errEl.style.display='block'; return; }
  if(!rak){ errEl.textContent='Rak wajib diisi.'; errEl.style.display='block'; return; }
  errEl.style.display='none';

  try{
    await apiSend(API.items, 'POST', { name, tempat, rak, mata, papan, customer, manualNo, note, imageUrl });
    state.view = 'idle';
    await loadAll();
  }catch(e){
    errEl.textContent = e.message;
    errEl.style.display='block';
  }
}

function openEditItemForm(id){
  state.selectedItemId = id;
  state.view = 'edit-item';
  render();
}

async function submitEditItem(){
  const item = state.items.find(i=>i.id===state.selectedItemId);
  if(!item) return;

  const name = document.getElementById('eName').value.trim();
  const tempat = document.getElementById('eTempat').value.trim();
  const rak = document.getElementById('eRak').value.trim();
  const mata = document.getElementById('eMata').value.trim();
  const papan = document.getElementById('ePapan').value.trim();
  const customer = document.getElementById('eCustomer').value.trim();
  const manualNo = document.getElementById('eManualNo').value.trim();
  const note = document.getElementById('eNote').value.trim();
  const imageUrl = document.getElementById('eImageUrl').value.trim();
  const errEl = document.getElementById('editItemErr');

  if(!name){ errEl.textContent='Nama pisau wajib diisi.'; errEl.style.display='block'; return; }
  if(!rak){ errEl.textContent='Rak wajib diisi.'; errEl.style.display='block'; return; }
  errEl.style.display='none';

  try{
    await apiSend(API.items+'/'+encodeURIComponent(item.id), 'PATCH', { name, tempat, rak, mata, papan, customer, manualNo, note, imageUrl });
    state.view = 'history';
    await loadAll();
  }catch(e){
    errEl.textContent = e.message;
    errEl.style.display='block';
  }
}

function renderOperatorPanel(){
  const body = document.getElementById('operatorBody');
  if(!body) return;

  if(state.view === 'photo-preview'){
    const item = state.items.find(i => i.id === state.selectedItemId);
    if(!item){ state.view = 'idle'; render(); return; }
    const canLoan = item.available > 0 && item.status !== 'discontinued';
    const photoContent = item.imageUrl
      ? '<img class="loan-photo-preview" src="'+String(item.imageUrl).replace(/\"/g, '&quot;')+'" alt="Foto pisau '+String(item.name).replace(/\"/g, '&quot;')+'" onerror="this.style.display=\'none\';document.getElementById(\'loanPhotoError\').style.display=\'block\';">'
        + '<div id="loanPhotoError" class="empty-state" style="display:none;">Foto gagal dimuat. Periksa URL foto Cloudinary.</div>'
      : '<div class="loan-photo-empty"><div style="font-size:42px;margin-bottom:10px;">📷</div><b>Foto belum tersedia</b><p>Tambahkan foto melalui menu Riwayat / Edit.</p></div>';
    body.innerHTML = '<div class="loan-photo-panel">'
      + '<h3 style="margin:0 0 6px;">Foto Pisau</h3>'
      + '<div class="loan-photo-name">'+item.name+'</div>'
      + '<div class="loan-photo-detail">'+itemDetailLine(item)+' · No Urut: '+(item.manualNo||'-')+'</div>'
      + '<div class="loan-photo-stage">'+photoContent+'</div>'
      + '<div class="loan-photo-actions">'
      + '<button class="btn accent" '+(canLoan?'':'disabled title="Pisau tidak tersedia untuk dipinjam"')+' onclick="continueToLoanForm()">Pinjam</button>'
      + '<button class="btn secondary" onclick="cancelPhotoPreview()">Batal</button>'
      + '</div></div>';
    return;
  }

  if(state.view === 'idle'){
    body.innerHTML = '<div class="empty-state">'
      +'<svg viewBox="0 0 24 24" fill="none" stroke-width="1.6"><rect x="3" y="7" width="18" height="13" rx="1"/><path d="M8 7V5a4 4 0 018 0v2"/></svg>'
      +'Pilih pisau dari katalog untuk mulai memproses peminjaman, atau klik "Riwayat" untuk melihat histori pinjam pisau tertentu.'
      +'<div style="margin-top:24px;"><button class="btn secondary small" onclick="openDiscontinueReport()">Lihat Laporan Discontinue</button></div>'
      +'</div>';
    return;
  }

  if(state.view === 'discontinue-report'){
    const reportItems = state.items.filter(i => i.discontinueHistory && i.discontinueHistory.length > 0);

    if (reportItems.length === 0) {
      body.innerHTML = '<div class="detail-panel"><h3>Laporan Discontinue</h3><p style="color:var(--ink-soft);font-size:13px;margin:12px 0;">Belum ada data riwayat discontinue pada pisau manapun.</p><div class="btn-row"><button class="btn secondary" onclick="cancelSelection()">Tutup</button></div></div>';
      return;
    }

    let tableRows = '';
    reportItems.forEach(i => {
      const hist = i.discontinueHistory;
      const totalCount = hist.length;
      hist.forEach((h, idx) => {
        const isFirst = idx === 0; 
        tableRows += '<tr>'
          + '<td style="vertical-align:top; padding:8px; border-bottom:1px solid var(--line);">'+(isFirst ? '<b>'+i.name+'</b><br><span style="font-size:11px;color:var(--ink-soft);">'+itemDetailLine(i)+'</span>' : '')+'</td>'
          + '<td class="mono" style="text-align:center; vertical-align:top; padding:8px; border-bottom:1px solid var(--line);">'+(isFirst ? totalCount+'x' : '')+'</td>'
          + '<td style="vertical-align:top; padding:8px; border-bottom:1px solid var(--line);">'+fmtDate(h.dateDiscontinued)+'</td>'
          + '<td style="vertical-align:top; padding:8px; border-bottom:1px solid var(--line);">'+(h.dateReactivated ? fmtDate(h.dateReactivated) : '<span style="color:var(--warn);font-size:11.5px;font-weight:500;">Masih nonaktif</span>')+'</td>'
          + '<td style="vertical-align:top; padding:8px; border-bottom:1px solid var(--line);">'+(h.reason || '-')+'</td>'
          + '</tr>';
      });
    });

    body.innerHTML = '<div class="detail-panel">'
      +'<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">'
      +'<h3 style="margin:0; font-size:16px;">Laporan Riwayat Discontinue</h3>'
      +'<button class="btn accent small" onclick="exportDiscontinueReport()">Export ke Excel</button>'
      +'</div>'
      +'<div style="max-height:380px; overflow-y:auto; border:1px solid var(--line); border-radius:2px;">'
      +'<table style="width:100%; border-collapse:collapse; font-size:12.5px; text-align:left;">'
      +'<thead style="background:var(--bg); position:sticky; top:0; z-index:1;"><tr>'
      +'<th style="padding:10px 8px; border-bottom:1px solid var(--line);">Nama & Deskripsi</th>'
      +'<th style="padding:10px 8px; border-bottom:1px solid var(--line); text-align:center;">Total</th>'
      +'<th style="padding:10px 8px; border-bottom:1px solid var(--line);">Tgl Discontinue</th>'
      +'<th style="padding:10px 8px; border-bottom:1px solid var(--line);">Tgl Aktif</th>'
      +'<th style="padding:10px 8px; border-bottom:1px solid var(--line);">Alasan</th>'
      +'</tr></thead>'
      +'<tbody>'
      + tableRows
      +'</tbody>'
      +'</table>'
      +'</div>'
      +'</div>'
      +'<div class="btn-row"><button class="btn secondary" onclick="cancelSelection()">Tutup</button></div>';
    return;
  }

  if(state.view === 'add-item'){
    body.innerHTML =
      '<h3 style="font-size:15px;margin-bottom:14px;">Tambah pisau baru</h3>'
      +'<div class="field"><label>Nama pisau</label><input type="text" id="aName" placeholder="Contoh: Box Polos"></div>'
      +'<div class="row2">'
      +'<div class="field"><label>Rak</label><input type="text" id="aRak" placeholder="Contoh: 1, atau 5A"></div>'
      +'<div class="field"><label>Tempat</label><input type="text" id="aTempat" placeholder="Contoh: 84"></div>'
      +'</div>'
      +'<div class="row2">'
      +'<div class="field"><label>Mata</label><input type="text" id="aMata" placeholder="Jumlah mata pisau"></div>'
      +'<div class="field"><label>Papan</label><input type="text" id="aPapan" placeholder="Jumlah papan penyangga"></div>'
      +'</div>'
      +'<div class="field"><label>No Urut (opsional)</label><input type="text" id="aManualNo" placeholder="Nomor urut, jika ada"></div>'
      +'<div class="field"><label>Customer (opsional)</label><input type="text" id="aCustomer" placeholder="Customer tetap pisau ini"></div>'
      +'<div class="field knife-photo-upload"><label>Foto pisau (opsional)</label>'
      +'<input type="file" id="aPhotoFile" accept="image/*" style="display:none" onchange="previewSelectedKnifePhoto(\'a\')">'
      +'<div class="photo-upload-row"><button type="button" class="btn secondary small" onclick="chooseKnifePhoto(\'a\')">Pilih Foto</button><button type="button" class="btn accent small" id="aUploadPhotoBtn" onclick="uploadKnifePhoto(\'a\')">Upload Foto</button></div>'
      +'<small class="field-help">Pilih gambar dari komputer (maksimal 10 MB), lalu upload ke Cloudinary.</small>'
      +'<div id="aPhotoStatus" class="photo-upload-status" aria-live="polite"></div>'
      +'<img id="aPhotoPreview" class="photo-upload-preview" alt="Pratinjau foto pisau" style="display:none">'
      +'<details class="photo-url-details"><summary>Atau tempel URL foto yang sudah ada</summary><input type="url" id="aImageUrl" oninput="previewKnifePhoto(\'a\')" placeholder="https://res.cloudinary.com/.../foto.jpg"></details>'
      +'</div>'
      +'<div class="field"><label>Catatan Singkat (opsional)</label><input type="text" id="aNote" placeholder="Misal: Sedikit tumpul, pisau sensitif, dll"></div>'
      +'<div id="addItemErr" class="err" style="display:none;"></div>'
      +'<div class="btn-row"><button class="btn accent" onclick="submitAddItem()">Simpan pisau</button>'
      +'<button class="btn secondary" onclick="cancelSelection()">Batal</button></div>';
    return;
  }

  if(state.view === 'edit-item'){
    const item = state.items.find(i=>i.id===state.selectedItemId);
    if(!item){ state.view='idle'; render(); return; }
    body.innerHTML =
      '<h3 style="font-size:15px;margin-bottom:14px;">Edit Data Pisau</h3>'
      +'<div class="field"><label>Nama pisau</label><input type="text" id="eName" value="'+item.name+'"></div>'
      +'<div class="row2">'
      +'<div class="field"><label>Rak</label><input type="text" id="eRak" value="'+item.rak+'"></div>'
      +'<div class="field"><label>Tempat</label><input type="text" id="eTempat" value="'+item.tempat+'"></div>'
      +'</div>'
      +'<div class="row2">'
      +'<div class="field"><label>Mata</label><input type="text" id="eMata" value="'+item.mata+'"></div>'
      +'<div class="field"><label>Papan</label><input type="text" id="ePapan" value="'+item.papan+'"></div>'
      +'</div>'
      +'<div class="field"><label>No Urut</label><input type="text" id="eManualNo" value="'+(item.manualNo||'')+'"></div>'
      +'<div class="field"><label>Customer (opsional)</label><input type="text" id="eCustomer" value="'+(item.customer||'')+'"></div>'
      +'<div class="field knife-photo-upload"><label>Foto pisau (opsional)</label>'
      +'<input type="file" id="ePhotoFile" accept="image/*" style="display:none" onchange="previewSelectedKnifePhoto(\'e\')">'
      +'<div class="photo-upload-row"><button type="button" class="btn secondary small" onclick="chooseKnifePhoto(\'e\')">Pilih Foto</button><button type="button" class="btn accent small" id="eUploadPhotoBtn" onclick="uploadKnifePhoto(\'e\')">Upload Foto</button></div>'
      +'<small class="field-help">Upload foto baru untuk mengganti foto lama.</small>'
      +'<div id="ePhotoStatus" class="photo-upload-status" aria-live="polite"></div>'
      +'<img id="ePhotoPreview" class="photo-upload-preview" alt="Pratinjau foto pisau" src="'+(item.imageUrl||'')+'" style="'+(item.imageUrl?'display:block':'display:none')+'">'
      +'<details class="photo-url-details"><summary>Atau tempel URL foto yang sudah ada</summary><input type="url" id="eImageUrl" value="'+(item.imageUrl||'')+'" oninput="previewKnifePhoto(\'e\')" placeholder="https://res.cloudinary.com/.../foto.jpg"></details>'
      +'</div>'
      +'<div class="field"><label>Catatan Singkat (opsional)</label><input type="text" id="eNote" value="'+(item.note||'')+'" placeholder="Misal: Sedikit tumpul, pisau sensitif, dll"></div>'
      +'<div id="editItemErr" class="err" style="display:none;"></div>'
      +'<div class="btn-row"><button class="btn accent" onclick="submitEditItem()">Simpan Perubahan</button>'
      +'<button class="btn secondary" onclick="viewItemHistory(\''+item.id+'\')">Batal</button></div>';
    return;
  }

  if(state.view === 'manage-item'){
    const item = state.items.find(i=>i.id===state.selectedItemId);
    if(!item){ state.view='idle'; render(); return; }
    body.innerHTML = '<div class="detail-panel">'
      +'<h3>Kelola Pisau: '+item.name+'</h3>'
      +'<p style="font-size:13.5px;color:var(--ink-soft);margin:10px 0 16px;line-height:1.5;">'
      +'Silakan pilih tindakan:<br>'
      +'<b>• Discontinue:</b> Menyembunyikan pisau dari katalog agar tidak bisa dipinjam, namun data dan riwayat tidak terhapus.<br>'
      +'<b>• Hapus Permanen:</b> Menghapus data pisau dari sistem selamanya.'
      +'</p>'
      +'<div class="field" style="margin-bottom:16px;"><label>Alasan Discontinue (jika pilih Discontinue)</label><input type="text" id="mDiscontinueReason" placeholder="Misal: Rusak, hilang, sedang diperbaiki..."></div>'
      +'</div>'
      +'<div class="btn-row">'
      +'<button class="btn secondary" onclick="discontinueItem(\''+item.id+'\')">Discontinue (Sembunyikan)</button>'
      +'<button class="btn accent" style="background:var(--warn);border-color:var(--warn);" onclick="deleteItemPermanently(\''+item.id+'\')">Hapus Permanen</button>'
      +'<button class="link-btn" onclick="cancelSelection()">Batal</button></div>';
    return;
  }

  if(state.view === 'history'){
    const item = state.items.find(i=>i.id===state.selectedItemId);
    if(!item){ state.view='idle'; render(); return; }
    const fullList = state.transactions.filter(t=>t.itemId===item.id).sort((a,b)=> (b.date+b.time).localeCompare(a.date+a.time));
    const list = fullList.slice(0, 3);
    const detailGrid = '<div class="detail-grid">'
      +'<div><span>Tempat</span><b>'+item.tempat+'</b></div>'
      +'<div><span>Rak</span><b>'+item.rak+'</b></div>'
      +'<div><span>Mata</span><b>'+item.mata+'</b></div>'
      +'<div><span>Papan</span><b>'+item.papan+'</b></div>'
      +'<div><span>No Urut</span><b>'+(item.manualNo||'-')+'</b></div>'
      +'</div>';
      
    const noteAlert = item.note ? '<div style="background:var(--warn-soft); color:var(--warn); padding:8px 12px; margin-bottom:12px; border-radius:2px; font-size:12.5px; border:1px solid var(--warn);"><b>Catatan Fisik Pisau:</b> '+item.note+'</div>' : '';

    const histHtml = list.length===0
      ? '<div class="empty-state" style="padding:18px 0;">Belum pernah dipinjam.</div>'
      : list.map(t=>{
          const statusTxt = t.status==='dipinjam'
            ? '<span class="status-pill dipinjam">Dipinjam</span>'
            : '<span class="status-pill dikembalikan">Kembali '+fmtDate(t.returnDate)+' '+(t.returnTime||'')+'</span>';
          return '<div class="hist-row"><span class="who">'+t.customer+' <span style="color:var(--ink-soft);font-weight:400;">(oleh '+t.operator+')</span></span><span class="when">'+fmtDate(t.date)+' '+(t.time||'')+'</span>'+statusTxt+'</div>';
        }).join('')
        + (fullList.length > 3 ? '<div class="hint" style="margin-top:8px;">Menampilkan 3 peminjaman terbaru dari total '+fullList.length+'.</div>' : '');
    
    body.innerHTML = '<div class="detail-panel">'
      +'<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">'
      +'<h3 style="margin:0; font-size:16px;">'+item.name+'</h3>'
      +'<button class="btn secondary small" onclick="openEditItemForm(\''+item.id+'\')">Edit Data</button>'
      +'</div>'
      +detailGrid
      +noteAlert
      +'<h3 style="font-size:13px;color:var(--ink-soft);margin-bottom:6px;">Riwayat peminjaman</h3>'
      +histHtml
      +'</div>'
      +'<div class="btn-row">'
      +(item.available>0 && item.status!=='discontinued' ? '<button class="btn accent" onclick="selectItem(\''+item.id+'\')">Pinjam pisau ini</button>' : '')
      +'<button class="btn secondary" onclick="cancelSelection()">Tutup</button></div>';
    return;
  }

  if(state.view === 'form'){
    const item = state.items.find(i=>i.id===state.selectedItemId);
    if(!item){ state.view='idle'; render(); return; }
    const defaultCount = defaultVoucherCount(item);
    
    const noteAlert = item.note ? '<div style="background:var(--warn-soft); color:var(--warn); padding:6px 10px; margin-top:8px; border-radius:2px; font-size:12px;"><b>Catatan Fisik Pisau:</b> '+item.note+'</div>' : '';

    body.innerHTML =
      '<div class="selected-strip"><div class="row1"><span>Meminjamkan: <b>'+item.name+'</b></span>'
      +'<button class="link-btn" onclick="cancelSelection()">Ganti</button></div>'
      +'<span class="meta">Detail Pisau: '+itemDetailLine(item)+'</span>'
      +'<span class="meta">No Urut: '+(item.manualNo||'-')+'</span>'
      +noteAlert
      +'</div>'
      +'<div class="field"><label>Operator</label><input type="text" id="fOperator" value="'+(state.lastOperator||'')+'" placeholder="Nama operator yang menginput"></div>'
      +'<div class="row2">'
      +'<div class="field"><label>Customer</label><input type="text" id="fCustomer" value="'+(item.customer||'')+'" placeholder="Nama customer"></div>'
      +'<div class="field"><label>No. SPK</label><input type="text" id="fSpk" placeholder="Contoh: SPK-0231"></div>'
      +'</div>'
      +'<div class="field"><label>Nama Produk</label><input type="text" id="fProduct" value="'+item.name+'" placeholder="Produk yang akan di-pond"></div>'
      +'<div class="field"><label>Tanggal pinjam</label><input type="date" id="fDate" value="'+todayStr()+'"></div>'
      +'<div class="field"><label>Catatan Peminjaman (opsional)</label><input type="text" id="fLoanNote" placeholder="Akan tercetak di voucher (Misal: Harap dikembalikan besok...)"></div>'
      +'<div id="formErr" class="err" style="display:none;"></div>'
      +'<div class="btn-row" style="align-items:center;">'
      +'<label class="hint" style="margin:0;display:flex;align-items:center;gap:6px;">Jumlah voucher (per papan)<input type="number" id="fVoucherCount" min="1" value="'+defaultCount+'" style="width:64px;"></label>'
      +'<button class="btn accent" onclick="submitLoan()">Simpan &amp; cetak voucher</button>'
      +'<button class="btn secondary" onclick="cancelSelection()">Batal</button></div>';
    return;
  }

  if(state.view === 'voucher'){
    const tx = state.transactions.find(t=>t.code===state.lastCode);
    if(!tx){ state.view='idle'; render(); return; }
    body.innerHTML = renderVoucherHtml(tx, false)
      +'<div class="btn-row"><button class="btn accent" onclick="printVoucher(\''+tx.code+'\')">Cetak voucher</button>'
      +'<button class="btn secondary" onclick="cancelSelection()">Transaksi baru</button></div>';
  }
}

function renderVoucherHtml(tx, forPrint, copyLabel){
  const brandClass = tx.status === 'dipinjam' ? 'SIPOND-brand-out' : 'SIPOND-brand-returned';
  const printStyle = forPrint ? 'padding:2.5mm 3.5mm; border:1.5px solid #000; width:100%; height:100%; box-sizing:border-box; display:flex; flex-direction:column; page-break-inside:avoid; color:#000; overflow:hidden;' : '';
  const rowStyle   = forPrint ? 'display:flex; justify-content:space-between; align-items:flex-start; border-bottom:1px dotted #000; padding-bottom:1.5px; margin-bottom:1.5px;' : '';
  const labelStyle = forPrint ? 'font-size:8px; font-weight:700; color:#000; display:inline-block; width:55px; flex-shrink:0; line-height:1.1;' : '';
  const valStyle   = forPrint ? 'font-size:8.5px; font-weight:800; color:#000; text-align:right; flex-grow:1; line-height:1.1; word-break:break-word;' : '';
  const titleStyle = forPrint ? 'font-size:12px; font-weight:900; color:#000;' : '';

  return '<div class="voucher" style="'+(forPrint? printStyle : '')+'">'
    +(copyLabel ? '<div class="voucher-copy-label" style="'+(forPrint?'font-size:7px; font-weight:900; color:#000; border-bottom:1.5px solid #000; padding-bottom:2px; margin-bottom:3px; text-align:center; text-transform:uppercase;':'')+'">'+copyLabel+'</div>' : '')
    +'<div class="vh" style="'+(forPrint?'margin-bottom:2.5px; display:flex; justify-content:space-between; align-items:center; border-bottom:1.5px solid #000; padding-bottom:2px; flex-shrink:0;':'')+'"><b class="SIPOND-brand '+brandClass+'" style="'+titleStyle+'">SIPOND</b><span style="'+(forPrint?'font-size:7px; font-weight:700; color:#000;':'')+'">Voucher Peminjaman</span></div>'
    +'<div style="flex-grow:1; display:flex; flex-direction:column; justify-content:center;">'
    +'<div class="vrow" style="'+rowStyle+'"><span style="'+labelStyle+'">Operator</span><span style="'+valStyle+'">'+tx.operator+'</span></div>'
    +'<div class="vrow" style="'+rowStyle+'"><span style="'+labelStyle+'">Customer</span><span style="'+valStyle+'">'+tx.customer+'</span></div>'
    +'<div class="vrow" style="'+rowStyle+'"><span style="'+labelStyle+'">No. SPK</span><span style="'+valStyle+'">'+tx.spk+'</span></div>'
    +'<div class="vrow" style="'+rowStyle+'"><span style="'+labelStyle+'">Nama Produk</span><span style="'+valStyle+'">'+tx.product+'</span></div>'
    +'<div class="vrow" style="'+rowStyle+'"><span style="'+labelStyle+'">Detail Pisau</span><span style="'+valStyle+'">'+tx.itemDetail+'</span></div>'
    +'<div class="vrow" style="'+rowStyle+'"><span style="'+labelStyle+'">No Urut</span><span style="'+valStyle+'">'+(tx.manualNo||'-')+'</span></div>'
    +'<div class="vrow" style="'+rowStyle+'"><span style="'+labelStyle+'">Catatan</span><span style="'+valStyle+'">'+(tx.note || '-')+'</span></div>'
    +'<div class="vrow" style="'+rowStyle+'"><span style="'+labelStyle+'">Tgl Pinjam</span><span style="'+valStyle+'">'+fmtDate(tx.date)+'</span></div>'
    +'</div>'
    +'<div class="code" style="'+(forPrint?'font-size:10px; margin-top:2px; color:#000; text-align:center; font-family:monospace; font-weight:900; flex-shrink:0; padding-top:2.5px; border-top:1.5px dashed #000;':'')+'">'+tx.code+'</div>'
    +'</div>';
}

function defaultVoucherCount(item){
  const n = parseInt(item && item.papan, 10);
  return (n && n>0) ? n : 1;
}

function getVoucherPages(tx, overrideCount){
  const n = overrideCount !== undefined ? overrideCount : (tx.voucherCount || 1);
  let pages = [];
  for(let i=1;i<=n;i++){
    pages.push(renderVoucherHtml(tx, true, 'Papan '+i+' dari '+n));
  }
  if(!tx.adminCopyPrinted){
    pages.push(renderVoucherHtml(tx, true, 'Lembar untuk Admin (Arsip)'));
  }
  return pages;
}

function chunkAndRenderPrintPages(pagesArray) {
  const chunkSize = 12; 
  let html = '';
  for(let i = 0; i < pagesArray.length; i += chunkSize) {
    const chunk = pagesArray.slice(i, i + chunkSize);
    html += '<div class="print-page">' + chunk.join('') + '</div>';
  }
  return html;
}

async function submitLoan(){
  const item = state.items.find(i=>i.id===state.selectedItemId);
  const operator = document.getElementById('fOperator').value.trim();
  const customer = document.getElementById('fCustomer').value.trim();
  const spk = document.getElementById('fSpk').value.trim();
  const product = document.getElementById('fProduct').value.trim();
  const date = document.getElementById('fDate').value;
  const loanNote = document.getElementById('fLoanNote').value.trim();
  const voucherCount = parseInt(document.getElementById('fVoucherCount').value, 10) || 1;
  const errEl = document.getElementById('formErr');

  if(!operator){ errEl.textContent='Nama operator wajib diisi.'; errEl.style.display='block'; return; }
  if(!customer){ errEl.textContent='Nama customer wajib diisi.'; errEl.style.display='block'; return; }
  if(!spk){ errEl.textContent='Nomor SPK wajib diisi.'; errEl.style.display='block'; return; }
  if(!product){ errEl.textContent='Nama produk wajib diisi.'; errEl.style.display='block'; return; }
  if(!date){ errEl.textContent='Tanggal pinjam wajib diisi.'; errEl.style.display='block'; return; }
  if(voucherCount<1){ errEl.textContent='Jumlah voucher minimal 1.'; errEl.style.display='block'; return; }
  if(item.available<=0){ errEl.textContent='Pisau ini sedang tidak tersedia.'; errEl.style.display='block'; return; }
  errEl.style.display='none';

  try{
    const tx = await apiSend(API.loans, 'POST', { operator, customer, spk, product, note: loanNote, itemId:item.id, date, voucherCount });
    state.lastOperator = operator;
    localStorage.setItem('SIPOND_lastOperator', operator);
    state.lastCode = tx.code;
    state.view = 'voucher';
    await loadAll();
    
    const refreshedTx = state.transactions.find(t => t.code === tx.code);
    if(refreshedTx && loanNote) {
      refreshedTx.note = loanNote;
    }
    setTimeout(()=>printVoucher(tx.code), 200);
  }catch(e){
    errEl.textContent = e.message;
    errEl.style.display='block';
  }
}

async function markAdminCopyPrinted(code){
  try{ await apiSend('/api/loans/'+encodeURIComponent(code)+'/mark-admin-printed', 'POST'); }catch(e){ }
}

async function printVoucher(code, overrideCount){
  const tx = state.transactions.find(t=>t.code===code);
  if(!tx) return;
  const includesAdminCopy = !tx.adminCopyPrinted;
  const pages = getVoucherPages(tx, overrideCount);
  const printArea = document.getElementById('printArea');
  if(printArea) printArea.innerHTML = chunkAndRenderPrintPages(pages);
  
  window.print();
  if(includesAdminCopy){
    tx.adminCopyPrinted = true;
    await markAdminCopyPrinted(code);
  }
}

function customReprint(code) {
  const tx = state.transactions.find(t=>t.code===code);
  if(!tx) return;
  const currentCount = tx.voucherCount || 1;
  const input = prompt('Berapa jumlah voucher (papan) yang ingin dicetak ulang untuk produk "' + tx.product + '"?', currentCount);
  if(input === null) return;
  const num = parseInt(input, 10);
  if(isNaN(num) || num < 1) {
    alert('Jumlah voucher tidak valid.');
    return;
  }
  if (!tx.note || tx.note.trim() === '') {
    const noteInput = prompt('Catatan tidak ditemukan (tidak tersimpan di server). Masukkan catatan untuk dicetak ulang (kosongkan jika tidak ada):', '');
    if (noteInput !== null) {
      tx.note = noteInput.trim();
    }
  }
  printVoucher(code, num);
}

async function bulkPrintHistory(){
  const codes = Array.from(state.selectedHistoryCodes);
  if(codes.length===0) return;
  const txs = codes.map(c => state.transactions.find(t=>t.code===c)).filter(Boolean);
  if(txs.length===0) return;
  const codesNeedingAdminCopy = txs.filter(t => !t.adminCopyPrinted).map(t=>t.code);
  let allPages = [];
  txs.forEach(t => { allPages = allPages.concat(getVoucherPages(t)); });
  const printArea = document.getElementById('printArea');
  if(printArea) printArea.innerHTML = chunkAndRenderPrintPages(allPages);
  window.print();
  if(codesNeedingAdminCopy.length>0){
    txs.forEach(t => { if(codesNeedingAdminCopy.includes(t.code)) t.adminCopyPrinted = true; });
    await Promise.all(codesNeedingAdminCopy.map(markAdminCopyPrinted));
  }
}

function startReturnConfirm(code){
  state.confirmingReturnCode = code;
  state.editingCode = null;
  render();
}
function cancelReturnConfirm(){
  state.confirmingReturnCode = null;
  render();
}
async function confirmReturn(code){
  const input = document.getElementById('returnDateInput_'+code);
  const returnDate = input ? input.value : todayStr();
  try{
    await apiSend(API.returns(code), 'POST', { returnDate: returnDate || todayStr() });
    state.confirmingReturnCode = null;
    await loadAll();
  }catch(e){
    alert(e.message);
  }
}

function startEditLoan(code){
  state.editingCode = code;
  state.confirmingReturnCode = null;
  renderHistory();
}
function cancelEditLoan(){
  state.editingCode = null;
  renderHistory();
}
async function confirmEditLoan(code){
  const opInput = document.getElementById('editOperator_'+code);
  const dateInput = document.getElementById('editDate_'+code);
  const operator = opInput ? opInput.value.trim() : '';
  const date = dateInput ? dateInput.value : '';
  if(!operator){ alert('Nama operator wajib diisi.'); return; }
  if(!date){ alert('Tanggal pinjam wajib diisi.'); return; }
  try{
    await apiSend('/api/loans/'+encodeURIComponent(code), 'PATCH', { operator, date });
    state.editingCode = null;
    state.lastOperator = operator;
    localStorage.setItem('SIPOND_lastOperator', operator);
    await loadAll();
  }catch(e){
    alert(e.message);
  }
}

async function deleteTransaction(code){
  const tx = state.transactions.find(t=>t.code===code);
  if(!tx) return;
  const msg = tx.status === 'dipinjam'
    ? 'Menghapus riwayat ini akan membuat pisau "'+tx.product+'" kembali berstatus tersedia. Yakin ingin menghapus riwayat peminjaman ini?'
    : 'Yakin ingin menghapus riwayat peminjaman ini? Tindakan ini tidak bisa dibatalkan.';
  if(!confirm(msg)) return;
  try{
    await apiSend('/api/loans/'+encodeURIComponent(code), 'DELETE');
    state.selectedHistoryCodes.delete(code);
    await loadAll();
  }catch(e){
    alert(e.message);
  }
}

/* ---------- Riwayat (global) ---------- */
function initHistoryDateDefaults(){
  const fromEl = document.getElementById('historyDateFrom');
  const toEl = document.getElementById('historyDateTo');
  if(fromEl && !fromEl.value) fromEl.value = monthsAgoStr(6);
  if(toEl && !toEl.value) toEl.value = todayStr();
}

function resetHistoryFilters(){
  const searchEl = document.getElementById('historySearch');
  const filterEl = document.getElementById('historyFilter');
  const custEl = document.getElementById('historyCustomerFilter');
  const locEl = document.getElementById('historyLocationFilter');
  const fromEl = document.getElementById('historyDateFrom');
  const toEl = document.getElementById('historyDateTo');

  if(searchEl) searchEl.value = '';
  if(filterEl) filterEl.value = 'all';
  if(custEl) custEl.value = 'Semua customer';
  if(locEl) locEl.value = 'Semua lokasi';
  if(fromEl) fromEl.value = monthsAgoStr(6);
  if(toEl) toEl.value = todayStr();
  renderHistory();
}

function renderHistoryCustomerFilter(){
  const sel = document.getElementById('historyCustomerFilter');
  if(!sel) return;
  const custs = ['Semua customer', ...new Set(state.transactions.map(t=>t.customer).filter(c=>c))].sort();
  const current = sel.value || 'Semua customer';
  sel.innerHTML = custs.map(c=>'<option value="'+c+'">'+c+'</option>').join('');
  sel.value = custs.includes(current) ? current : 'Semua customer';
}

function renderHistoryLocationFilter(){
  const sel = document.getElementById('historyLocationFilter');
  if(!sel) return;
  const locs = ['Semua lokasi', ...new Set(state.transactions.map(t=>t.location).filter(l=>l))].sort();
  const current = sel.value || 'Semua lokasi';
  sel.innerHTML = locs.map(l=>'<option value="'+l+'">'+l+'</option>').join('');
  sel.value = locs.includes(current) ? current : 'Semua lokasi';
}

function renderHistory(){
  initHistoryDateDefaults();
  const searchEl = document.getElementById('historySearch');
  const filterEl = document.getElementById('historyFilter');
  const custEl = document.getElementById('historyCustomerFilter');
  const locEl = document.getElementById('historyLocationFilter');
  const fromEl = document.getElementById('historyDateFrom');
  const toEl = document.getElementById('historyDateTo');

  const q = (searchEl ? searchEl.value : '').toLowerCase();
  const statusFilter = filterEl ? filterEl.value : 'all';
  const custFilter = custEl ? custEl.value : 'Semua customer';
  const locFilter = locEl ? locEl.value : 'Semua lokasi';
  const dateFrom = fromEl ? fromEl.value : '';
  const dateTo = toEl ? toEl.value : '';

  const rows = state.transactions.filter(t=>{
    const matchQ = !q
      || t.code.toLowerCase().includes(q)
      || (t.operator||'').toLowerCase().includes(q)
      || (t.customer||'').toLowerCase().includes(q)
      || (t.spk||'').toLowerCase().includes(q)
      || (t.product||'').toLowerCase().includes(q)
      || (t.note||'').toLowerCase().includes(q);
    const matchStatus = statusFilter==='all' || t.status===statusFilter;
    const matchCust = custFilter==='Semua customer' || t.customer===custFilter;
    const matchLoc = locFilter==='Semua lokasi' || t.location===locFilter;
    const matchFrom = !dateFrom || t.date >= dateFrom;
    const matchTo = !dateTo || t.date <= dateTo;
    return matchQ && matchStatus && matchCust && matchLoc && matchFrom && matchTo;
  });

  const body = document.getElementById('historyBody');
  if(!body) return;
  state.lastHistoryRows = rows;
  if(rows.length===0){
    body.innerHTML = '<tr><td colspan="13" style="color:var(--ink-soft);text-align:center;padding:22px;">Tidak ada transaksi yang cocok dengan filter.</td></tr>';
    updateHistoryBulkBar();
    return;
  }
  body.innerHTML = rows.map(t=>{
    const usageCount = state.transactions.filter(x => x.itemId === t.itemId).length;
    
    let actionCell;
    if(t.status==='dikembalikan'){
      actionCell = '<div class="action-cell">'
        +'<button class="action-btn act-reprint" onclick="customReprint(\''+t.code+'\')">Cetak ulang</button>'
        +'<button class="action-btn act-delete" onclick="deleteTransaction(\''+t.code+'\')">Hapus</button>'
        +'</div>';
    } else if(state.editingCode === t.code){
      actionCell = '<div class="return-confirm">'
        +'<input type="text" id="editOperator_'+t.code+'" value="'+t.operator+'" placeholder="Nama operator" style="width:110px;">'
        +'<input type="date" id="editDate_'+t.code+'" value="'+t.date+'">'
        +'<button class="btn small accent" onclick="confirmEditLoan(\''+t.code+'\')">Simpan</button>'
        +'<button class="link-btn" onclick="cancelEditLoan()">Batal</button>'
        +'</div>';
    } else if(state.confirmingReturnCode === t.code){
      actionCell = '<div class="return-confirm">'
        +'<input type="date" id="returnDateInput_'+t.code+'" value="'+todayStr()+'">'
        +'<button class="btn small accent" onclick="confirmReturn(\''+t.code+'\')">OK</button>'
        +'<button class="link-btn" onclick="cancelReturnConfirm()">Batal</button>'
        +'</div>';
    } else {
      actionCell = '<div class="action-cell">'
        +'<button class="action-btn act-return" onclick="startReturnConfirm(\''+t.code+'\')">Tandai kembali</button>'
        +'<button class="action-btn act-reprint" onclick="customReprint(\''+t.code+'\')">Cetak ulang</button>'
        +'<button class="action-btn act-edit" onclick="startEditLoan(\''+t.code+'\')">Edit</button>'
        +'<button class="action-btn act-delete" onclick="deleteTransaction(\''+t.code+'\')">Hapus</button>'
        +'</div>';
    }

    actionCell += '<div class="hint" style="margin-top:6px; font-weight:600; text-align:center;">Total Pakai: '+usageCount+'x</div>';

    return '<tr>'
      +'<td><input type="checkbox" class="history-checkbox"'+(state.selectedHistoryCodes.has(t.code)?' checked':'')+' onchange="toggleHistoryCheckbox(\''+t.code+'\', this.checked)"></td>'
      +'<td class="mono">'+t.code+'</td>'
      +'<td>'+t.operator+'</td>'
      +'<td>'+t.customer+'</td>'
      +'<td class="mono">'+t.spk+'</td>'
      +'<td>'+t.product+'</td>'
      +'<td>'+(t.itemDetail||'-')+'</td>'
      +'<td class="mono">'+(t.manualNo||'-')+'</td>'
      +'<td style="max-width:150px; white-space:normal; font-size:11.5px; line-height:1.4;">'+(t.note || '-')+'</td>'
      +'<td>'+fmtDate(t.date)+' '+(t.time||'')+'</td>'
      +'<td>'+(t.returnDate ? fmtDate(t.returnDate)+' '+(t.time||'') : '-')+'</td>'
      +'<td><span class="status-pill '+t.status+'">'+(t.status==='dipinjam'?'Dipinjam':'Dikembalikan')+'</span></td>'
      +'<td>'+actionCell+'</td>'
      +'</tr>';
  }).join('');
  updateHistoryBulkBar();
}

function toggleHistoryCheckbox(code, checked){
  if(checked) state.selectedHistoryCodes.add(code); else state.selectedHistoryCodes.delete(code);
  updateHistoryBulkBar();
}

function toggleSelectAllHistory(checked){
  const rows = state.lastHistoryRows || [];
  rows.forEach(t => {
    if(checked) state.selectedHistoryCodes.add(t.code); else state.selectedHistoryCodes.delete(t.code);
  });
  renderHistory();
}

function updateHistoryBulkBar(){
  const rows = state.lastHistoryRows || [];
  const countEl = document.getElementById('historyBulkCount');
  const btn = document.getElementById('historyBulkPrintBtn');
  const allCb = document.getElementById('historySelectAllCheckbox');
  const n = state.selectedHistoryCodes.size;
  if(countEl) countEl.textContent = n>0 ? n+' transaksi dipilih' : '';
  if(btn) btn.style.display = n>0 ? 'inline-block' : 'none';
  if(allCb) allCb.checked = rows.length>0 && rows.every(t=>state.selectedHistoryCodes.has(t.code));
}

async function exportHistory(){
  const rows = state.lastHistoryRows || [];
  if(rows.length === 0){
    alert('Tidak ada data untuk diekspor sesuai filter saat ini.');
    return;
  }
  try{
    const res = await fetch('/api/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ rows })
    });
    if(!res.ok){
      const err = await res.json().catch(()=>({}));
      throw new Error(err.error || 'Gagal mengekspor data.');
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'riwayat-peminjaman-'+todayStr()+'.xlsx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }catch(e){
    alert(e.message);
  }
}

async function exportDiscontinueReport(){
  const reportItems = state.items.filter(i => i.discontinueHistory && i.discontinueHistory.length > 0);
  if(reportItems.length === 0){
    alert('Tidak ada data laporan untuk diekspor.');
    return;
  }

  const rows = [];
  reportItems.forEach(i => {
    const hist = i.discontinueHistory;
    const totalCount = hist.length;
    hist.forEach((h, idx) => {
      const isFirst = idx === 0;
      rows.push({
        name: isFirst ? i.name : '',
        detail: isFirst ? itemDetailLine(i) : '',
        total: isFirst ? totalCount + 'x' : '',
        dateDiscontinued: fmtDate(h.dateDiscontinued),
        dateReactivated: h.dateReactivated ? fmtDate(h.dateReactivated) : '',
        reason: h.reason || '-'
      });
    });
  });

  try{
    const res = await fetch('/api/export-discontinue', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ rows })
    });
    if(!res.ok){
      const err = await res.json().catch(()=>({}));
      throw new Error(err.error || 'Gagal mengekspor data.');
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'laporan-discontinue-'+todayStr()+'.xlsx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }catch(e){
    alert(e.message);
  }
}

/* ---------- Root render + wiring ---------- */
function render(){
  renderLocationFilter();
  renderCustomerFilter();
  renderStats();
  renderCatalog();
  renderOperatorPanel();
  renderHistoryCustomerFilter();
  renderHistoryLocationFilter();
  renderHistory();
}

function updateCatalogAndStats(){ renderStats(); renderCatalog(); }

const searchInp = document.getElementById('searchInput');
if(searchInp) searchInp.addEventListener('input', updateCatalogAndStats);
const locFilt = document.getElementById('locationFilter');
if(locFilt) locFilt.addEventListener('change', updateCatalogAndStats);
const custFilt = document.getElementById('customerFilter');
if(custFilt) custFilt.addEventListener('change', updateCatalogAndStats);

const histFilt = document.getElementById('historyFilter');
if(histFilt) histFilt.addEventListener('change', renderHistory);
const histSearch = document.getElementById('historySearch');
if(histSearch) histSearch.addEventListener('input', renderHistory);
const histCustFilt = document.getElementById('historyCustomerFilter');
if(histCustFilt) histCustFilt.addEventListener('change', renderHistory);
const histLocFilt = document.getElementById('historyLocationFilter');
if(histLocFilt) histLocFilt.addEventListener('change', renderHistory);
const histDateFrom = document.getElementById('historyDateFrom');
if(histDateFrom) histDateFrom.addEventListener('change', renderHistory);
const histDateTo = document.getElementById('historyDateTo');
if(histDateTo) histDateTo.addEventListener('change', renderHistory);

checkAuthAndInit();

setInterval(() => {
  if(state.view === 'form' || state.view === 'add-item' || state.view === 'manage-item' || state.view === 'edit-item' || state.view === 'discontinue-report') return;
  if(state.editingCode || state.confirmingReturnCode) return;
  loadAll();
}, 8000);