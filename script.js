let products = JSON.parse(localStorage.getItem('valente_studio_products') || '[]');
let editingId = null;
let currentFilter = 'all';
let searchQuery = '';
let currentImageData = null;

function save() {
  try {
    localStorage.setItem('valente_studio_products', JSON.stringify(products));
    const lbl = document.getElementById('autosave-label');
    if (lbl) {
      const now = new Date();
      lbl.textContent = `Guardado a las ${now.getHours()}:${String(now.getMinutes()).padStart(2,'0')}:${String(now.getSeconds()).padStart(2,'0')}`;
      setTimeout(() => { lbl.textContent = 'Guardado automático activo'; }, 3000);
    }
    updateStorageBar();
  } catch (e) {
    alert('⚠️ Almacenamiento lleno. Elimina algunos productos o sus imágenes para continuar.');
  }
}

function updateStorageBar() {
  try {
    const data = localStorage.getItem('valente_studio_products') || '';
    const usedKB = Math.round((data.length * 2) / 1024);
    const limitKB = 5000;
    const pct = Math.min(100, Math.round((usedKB / limitKB) * 100));
    const bar = document.getElementById('storage-bar-fill');
    const label = document.getElementById('storage-label');
    if (!bar || !label) return;
    bar.style.width = pct + '%';
    bar.style.background = pct > 80 ? '#C0392B' : pct > 60 ? '#E67E22' : '#9CAF88';
    label.textContent = `Almacenamiento: ${usedKB} KB / ~5000 KB usados (${pct}%)`;
  } catch(e) {}
}

function exportToExcel() {
  if (products.length === 0) { alert('No hay productos para exportar.'); return; }

  const rows = products.map(p => ({
    'Nombre': p.name || '',
    'Categoría': p.category || '',
    'Código / SKU': p.code || '',
    'Enlace': p.link || '',
    'Precio Unitario (₡)': p.price || 0,
    'Cantidad': p.qty || 1,
    'Total (₡)': (p.price || 0) * (p.qty || 1),
    'Notas': p.notes || '',
    'Estado': p.purchased ? 'Comprado' : 'Por comprar',
    'Fecha agregado': p.createdAt ? new Date(p.createdAt).toLocaleDateString('es-CR') : ''
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [
    {wch:30},{wch:18},{wch:16},{wch:40},
    {wch:18},{wch:10},{wch:16},{wch:35},{wch:14},{wch:16}
  ];

  const headers = Object.keys(rows[0]);
  headers.forEach((h, i) => {
    const cell = XLSX.utils.encode_cell({r:0, c:i});
    if (!ws[cell]) return;
    ws[cell].s = { font: { bold: true }, fill: { fgColor: { rgb: "4A3728" } }, alignment: { horizontal: 'center' } };
  });

  const totalVal = products.reduce((a,p) => a + (p.price||0)*(p.qty||1), 0);
  const pendientes = products.filter(p => !p.purchased);
  const comprados = products.filter(p => p.purchased);
  const summaryData = [
    ['Valente Studio · Resumen de Inventario', ''],
    ['Fecha de exportación', new Date().toLocaleDateString('es-CR')],
    ['', ''],
    ['Total de productos', products.length],
    ['Por comprar', pendientes.length],
    ['Comprados', comprados.length],
    ['Monto total estimado (₡)', totalVal],
    ['Monto pendiente (₡)', pendientes.reduce((a,p) => a+(p.price||0)*(p.qty||1),0)],
    ['Monto comprado (₡)', comprados.reduce((a,p) => a+(p.price||0)*(p.qty||1),0)],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(summaryData);
  ws2['!cols'] = [{wch:32},{wch:24}];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws2, 'Resumen');
  XLSX.utils.book_append_sheet(wb, ws, 'Inventario');

  const fecha = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `Valente_Studio_Inventario_${fecha}.xlsx`);
}

function formatCurrency(val) {
  if (!val && val !== 0) return '—';
  return new Intl.NumberFormat('es-CR', { style: 'currency', currency: 'CRC', minimumFractionDigits: 0 }).format(val);
}

function updateStats() {
  const pending = products.filter(p => !p.purchased);
  const totalAmount = products.reduce((a, p) => a + (p.price * p.qty || 0), 0);
  document.getElementById('stat-total').textContent = products.length;
  document.getElementById('stat-amount').textContent = formatCurrency(totalAmount);
  document.getElementById('stat-pending').textContent = pending.length;
}

function filteredProducts() {
  let list = [...products];
  if (currentFilter === 'pending') list = list.filter(p => !p.purchased);
  if (currentFilter === 'purchased') list = list.filter(p => p.purchased);
  if (searchQuery) {
    const q = searchQuery.toLowerCase();
    list = list.filter(p =>
      (p.name||'').toLowerCase().includes(q) ||
      (p.code||'').toLowerCase().includes(q) ||
      (p.category||'').toLowerCase().includes(q)
    );
  }
  return list;
}

function render() {
  const grid = document.getElementById('grid');
  const list = filteredProducts();
  updateStats();

  if (list.length === 0) {
    grid.innerHTML = `
      <div class="empty-state">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.2">
          <path d="M6 2L3 6v14a2 2 0 002 2h14a2 2 0 002-2V6l-3-4z"/><line x1="3" y1="6" x2="21" y2="6"/>
          <path d="M16 10a4 4 0 01-8 0"/>
        </svg>
        <h3>Tu lista está vacía</h3>
        <p>Agrega productos que quieres comprar</p>
      </div>`;
    return;
  }

  grid.innerHTML = list.map(p => {
    const total = (p.price && p.qty) ? p.price * p.qty : null;
    const imgEl = p.image
      ? `<img class="card-img" src="${p.image}" alt="${p.name}">`
      : `<div class="card-img-placeholder"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.2"><path d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z"/></svg></div>`;
    return `
    <div class="card ${p.purchased ? 'purchased' : ''}" id="card-${p.id}">
      ${imgEl}
      <div class="card-body">
        ${p.purchased ? '<span class="purchased-badge">✓ Comprado</span>' : ''}
        ${p.category ? `<div class="card-category">${p.category}</div>` : ''}
        <div class="card-name">${p.name}</div>
        ${p.code ? `<div class="card-code">#${p.code}</div>` : ''}
        ${p.link ? `<a class="card-link" href="${p.link}" target="_blank" rel="noopener">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 13a5 5 0 007.54.54l3-3a5 5 0 00-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 00-7.54-.54l-3 3a5 5 0 007.07 7.07l1.71-1.71"/></svg>
          Ver producto
        </a>` : ''}
        ${p.notes ? `<div class="card-notes">${p.notes}</div>` : ''}
        <div class="card-meta">
          <div>
            <div class="card-price">${p.price ? formatCurrency(p.price) : '—'} <span>c/u</span></div>
            ${total ? `<div class="card-total">Total: ${formatCurrency(total)}</div>` : ''}
          </div>
          <div class="card-qty-badge">×${p.qty || 1}</div>
        </div>
        <div class="card-actions">
          ${!p.purchased ? `<button class="btn-icon btn-mark" onclick="markPurchased('${p.id}')">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" width="14" height="14"><polyline points="20 6 9 17 4 12"/></svg>
            Marcar comprado
          </button>` : `<button class="btn-icon btn-mark" onclick="markPurchased('${p.id}')" style="background:#C4B5A5;border-color:#C4B5A5;">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" width="14" height="14"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 102.13-9.36L1 10"/></svg>
            Desmarcar
          </button>`}
          <button class="btn-icon" onclick="editProduct('${p.id}')">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
          </button>
          <button class="btn-icon danger" onclick="deleteProduct('${p.id}')">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/></svg>
          </button>
        </div>
      </div>
    </div>`;
  }).join('');
}

function openModal(id = null) {
  editingId = id;
  currentImageData = null;
  const overlay = document.getElementById('modal-overlay');
  overlay.style.display = 'flex';

  if (id) {
    const p = products.find(x => x.id === id);
    document.getElementById('modal-title').textContent = 'Editar producto';
    document.getElementById('f-name').value = p.name || '';
    document.getElementById('f-category').value = p.category || '';
    document.getElementById('f-code').value = p.code || '';
    document.getElementById('f-link').value = p.link || '';
    document.getElementById('f-price').value = p.price || '';
    document.getElementById('f-qty').value = p.qty || 1;
    document.getElementById('f-notes').value = p.notes || '';
    if (p.image) {
      currentImageData = p.image;
      const prev = document.getElementById('img-preview');
      prev.src = p.image;
      prev.style.display = 'block';
    } else {
      document.getElementById('img-preview').style.display = 'none';
    }
  } else {
    document.getElementById('modal-title').textContent = 'Nuevo producto';
    ['f-name','f-category','f-code','f-link','f-price','f-notes'].forEach(id => {
      const el = document.getElementById(id);
      if (el.tagName === 'SELECT') el.value = '';
      else el.value = '';
    });
    document.getElementById('f-qty').value = 1;
    document.getElementById('img-preview').style.display = 'none';
    document.getElementById('img-input').value = '';
  }
}

function closeModal() {
  document.getElementById('modal-overlay').style.display = 'none';
  editingId = null;
  currentImageData = null;
}

function closeOnOverlay(e) {
  if (e.target === document.getElementById('modal-overlay')) closeModal();
}

function handleImage(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    const img = new Image();
    img.onload = () => {
      const MAX = 600;
      const scale = img.width > MAX ? MAX / img.width : 1;
      const canvas = document.createElement('canvas');
      canvas.width  = Math.round(img.width  * scale);
      canvas.height = Math.round(img.height * scale);
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      currentImageData = canvas.toDataURL('image/jpeg', 0.70);
      const prev = document.getElementById('img-preview');
      prev.src = currentImageData;
      prev.style.display = 'block';
      const kb = Math.round((currentImageData.length * 2) / 1024);
      prev.title = `Imagen comprimida: ~${kb} KB`;
    };
    img.src = ev.target.result;
  };
  reader.readAsDataURL(file);
}

function saveProduct() {
  const name = document.getElementById('f-name').value.trim();
  if (!name) { alert('El nombre del producto es obligatorio.'); return; }

  const product = {
    id: editingId || Date.now().toString(),
    name,
    category: document.getElementById('f-category').value,
    code: document.getElementById('f-code').value.trim(),
    link: document.getElementById('f-link').value.trim(),
    price: parseFloat(document.getElementById('f-price').value) || 0,
    qty: parseInt(document.getElementById('f-qty').value) || 1,
    notes: document.getElementById('f-notes').value.trim(),
    image: currentImageData || null,
    purchased: false,
    createdAt: new Date().toISOString()
  };

  if (editingId) {
    const existing = products.find(p => p.id === editingId);
    product.purchased = existing ? existing.purchased : false;
    product.createdAt = existing ? existing.createdAt : product.createdAt;
    products = products.map(p => p.id === editingId ? product : p);
  } else {
    products.unshift(product);
  }

  save();
  closeModal();
  render();
}

function editProduct(id) { openModal(id); }

function deleteProduct(id) {
  if (!confirm('¿Eliminar este producto de la lista?')) return;
  products = products.filter(p => p.id !== id);
  save();
  render();
}

function markPurchased(id) {
  products = products.map(p => p.id === id ? {...p, purchased: !p.purchased} : p);
  save();
  render();
}

function setFilter(f, btn) {
  currentFilter = f;
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  render();
}

function searchProducts(q) {
  searchQuery = q;
  render();
}

const uploadArea = document.getElementById('upload-area');
uploadArea.addEventListener('dragover', e => { e.preventDefault(); uploadArea.style.borderColor = 'var(--mocha)'; });
uploadArea.addEventListener('dragleave', () => { uploadArea.style.borderColor = 'var(--taupe)'; });
uploadArea.addEventListener('drop', e => {
  e.preventDefault();
  uploadArea.style.borderColor = 'var(--taupe)';
  const file = e.dataTransfer.files[0];
  if (file && file.type.startsWith('image/')) {
    const fakeEvent = { target: { files: [file] } };
    handleImage(fakeEvent);
  }
});

render();
updateStorageBar();
