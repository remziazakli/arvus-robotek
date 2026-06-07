// ─── AUTH GUARD ─────────────────────────────────────────
        // Auth guard disabled for Electron

        // ─── CATEGORIES ─────────────────────────────────────────
        const CATEGORIES = [
            'Çizgi İzleyen Robot', 'İnsansız Hava Araçları', 'Mini Sumo Robot',
            'Labirent Çözen', 'Savaşan Robot', 'Robo Futbol', 'Tasarla ve Geliştir', 'Serbest'
        ];
        const CAT_COLORS = {
            'Çizgi İzleyen Robot': '#4f8ef7', 'İnsansız Hava Araçları': '#7c6dfa',
            'Mini Sumo Robot': '#3fb950', 'Labirent Çözen': '#f0883e',
            'Savaşan Robot': '#f85149', 'Robo Futbol': '#4dd0e1',
            'Tasarla ve Geliştir': '#b95ff7', 'Serbest': '#e1914d'
        };

        // İndirilecek kategoriler (Konfigüre edilebilir)
        const DOWNLOAD_CATEGORIES = ['İnsansız Hava Araçları', 'Tasarla ve Geliştir', 'Savaşan Robot', 'Serbest'];

        // ─── DATA STORE ──────────────────────────────────────────
        function loadApps() { return JSON.parse(localStorage.getItem('bs_apps') || '[]'); }
        function saveApps(arr) { localStorage.setItem('bs_apps', JSON.stringify(arr)); }

        let apps = loadApps();
        let editingId = null;
        let confirmCb = null;
        let pendingImportRows = [];

        // ─── VIEW SWITCH ─────────────────────────────────────────
        const VIEW_TITLES = { dashboard: 'Dashboard', applications: 'Başvurular', import: 'Excel Yükle', settings: 'Ayarlar', sync: 'Senkronizasyon' };
        let activeCat = ''; // boş = Tümü
        function switchView(name) {
            document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            document.getElementById('view-' + name).classList.add('active');
            document.querySelector(`[data-view="${name}"]`).classList.add('active');
            document.getElementById('pageTitle').textContent = VIEW_TITLES[name];
            if (name === 'dashboard') { renderDashboard(); renderReportStatusGrid(); }
            if (name === 'applications') { renderCatTabs(); renderAppTable(); }
        }
        function toggleSidebar() {
            const s = document.getElementById('sidebar'), m = document.getElementById('mainContent');
            s.classList.toggle('collapsed'); m.classList.toggle('expanded');
        }
        // Kategori sekmeleri
        function renderCatTabs() {
            const tabs = document.getElementById('catTabs');
            const allCount = apps.length;
            const tabDefs = [{ label: '📋 Tümü', cat: '' }, ...CATEGORIES.map(c => ({ label: c, cat: c }))];
            tabs.innerHTML = tabDefs.map(({ label, cat }) => {
                const count = cat ? apps.filter(a => a.category === cat).length : allCount;
                const color = cat ? CAT_COLORS[cat] : '#8b949e';
                const active = activeCat === cat ? 'active' : '';
                return `<button class="cat-tab ${active}" style="--tc:${color}" onclick="setCatTab('${cat}')">${label}<span class="tab-count">${count}</span></button>`;
            }).join('');
        }
        function setCatTab(cat) {
            activeCat = cat;
            renderCatTabs();
            renderAppTable();
            const btn = document.getElementById('btnSyncCat');
            if (btn) {
                if (cat && localStorage.getItem('bs_script_url')) {
                    btn.style.display = 'inline-flex';
                } else {
                    btn.style.display = 'none';
                }
            }
        }
        // Sadece aktif kategori için Excel aktar
        function exportCatExcel() {
            const list = activeCat ? apps.filter(a => a.category === activeCat) : apps;
            if (!list.length) { showToast('Bu kategoride kayıt yok!', 'error'); return; }
            const wb = XLSX.utils.book_new();
            const cols = ['#', 'Ad Soyad', 'E-posta', 'Kategori', 'Okul', 'Doğum Tarihi', 'Veli İletişim', 'Takım Adı', 'Kayıt Tarihi', 'Durum'];
            const rows = [cols, ...list.map((a, i) => [i + 1, a.name, a.email, a.category, a.school, a.birth, a.parent, a.team, fmtDate(a.createdAt), a.status || 'Bekliyor'])];
            const ws = XLSX.utils.aoa_to_sheet(rows);
            ws['!cols'] = [{ wch: 5 }, { wch: 25 }, { wch: 30 }, { wch: 25 }, { wch: 35 }, { wch: 15 }, { wch: 18 }, { wch: 20 }, { wch: 18 }, { wch: 12 }];
            const sheetName = (activeCat || 'Tüm Kayıtlar').substring(0, 31);
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
            const date = new Date().toLocaleDateString('tr-TR').replace(/\./g, '-');
            XLSX.writeFile(wb, `basvurular_${sheetName}_${date}.xlsx`);
            showToast('Excel indirildi ✅', 'success');
        }

        // ─── DASHBOARD ───────────────────────────────────────────
        function renderDashboard() {
            const total = apps.length;
            const pending = apps.filter(a => a.status === 'Bekliyor').length;
            const today = apps.filter(a => a.createdAt?.startsWith(todayStr())).length;
            const GOAL = 1000, DAILY = 25;

            document.getElementById('statTotal').textContent = total;
            document.getElementById('statPending').textContent = pending;
            document.getElementById('statToday').textContent = today;

            setProg('progTotal', 'progTotalPct', total, GOAL);
            setProg('progPending', 'progPendingPct', pending, total || 1);
            setProg('progToday', 'progTodayPct', today, DAILY);

            document.getElementById('statTotalTarget').textContent = `Hedef: ${GOAL}`;
            document.getElementById('statPendingTarget').textContent = `Tüm Kayıtlar: ${total}`;

            // dash table — son 20
            renderDashTable();

            // Charts
            const catData = {};
            CATEGORIES.forEach(c => catData[c] = 0);
            apps.forEach(a => { if (catData[a.category] !== undefined) catData[a.category]++; else catData[a.category] = (catData[a.category] || 0) + 1; });
            renderBarChart('catChart', catData, CAT_COLORS);

            const stData = { 'Bekliyor': 0, 'Onaylandı': 0, 'Reddedildi': 0 };
            apps.forEach(a => { stData[a.status || 'Bekliyor'] = (stData[a.status || 'Bekliyor'] || 0) + 1; });
            renderBarChart('statusChart', stData, { 'Bekliyor': '#f0883e', 'Onaylandı': '#3fb950', 'Reddedildi': '#f85149' });
        }
        function setProg(fillId, pctId, val, max) {
            const pct = max > 0 ? Math.min(100, Math.round((val / max) * 100)) : 0;
            document.getElementById(fillId).style.width = pct + '%';
            document.getElementById(pctId).textContent = pct + '%';
        }
        function renderBarChart(id, data, colors) {
            const el = document.getElementById(id);
            const entries = Object.entries(data).filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]);
            if (!entries.length) { el.innerHTML = '<div style="color:var(--muted);font-size:13px;padding:8px">Veri yok</div>'; return; }
            const max = Math.max(...entries.map(e => e[1]));
            el.innerHTML = entries.map(([l, c]) => `
    <div class="bar-row">
      <div class="bar-label" title="${l}">${l}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${(c / max) * 100}%;background:${colors[l] || '#4f8ef7'}"></div></div>
      <div class="bar-count">${c}</div>
    </div>`).join('');
        }
        function renderDashTable() {
            const q = (document.getElementById('dashSearch')?.value || '').toLowerCase();
            const list = apps.filter(a => !q || JSON.stringify(a).toLowerCase().includes(q)).slice(0, 50);
            const tbody = document.getElementById('dashTbody');
            if (!list.length) { tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:20px;color:var(--muted)">Kayıt bulunamadı</td></tr>'; return; }
            tbody.innerHTML = list.map(a => `
    <tr onclick="openEditModal('${a.id}')" style="cursor:pointer">
      <td><b>${a.name || '—'}</b></td>
      <td><span class="cat-badge" style="background:${CAT_COLORS[a.category] || '#4f8ef7'}22;color:${CAT_COLORS[a.category] || '#4f8ef7'}">${a.category || '—'}</span></td>
      <td style="color:var(--muted);font-size:13px">${fmtDate(a.createdAt)}</td>
      <td><span class="status-badge badge-${stCls(a.status)}">${a.status || 'Bekliyor'}</span></td>
      <td onclick="event.stopPropagation()" style="display:flex;gap:6px">
        <button class="btn-icon2" onclick="openEditModal('${a.id}')">👁️</button>
        <button class="btn-icon2 del" onclick="confirmDelete('${a.id}','${escQ(a.name)}')">🗑️</button>
      </td>
    </tr>`).join('');
        }

        // ─── APPLICATIONS TABLE ──────────────────────────────────
        function renderAppTable() {
            const q = (document.getElementById('appSearch')?.value || '').toLowerCase();
            const st = document.getElementById('filterStatus')?.value || '';
            const rpt = document.getElementById('filterReport')?.value || '';
            const list = apps.filter(a => {
                const m = !q || (a.name || '').toLowerCase().includes(q) || (a.email || '').toLowerCase().includes(q) || (a.school || '').toLowerCase().includes(q) || (a.team || '').toLowerCase().includes(q);
                const mc = !activeCat || a.category === activeCat;
                const ms = !st || a.status === st;
                const mr = !rpt || (rpt === 'yes' && a.hasFiles) || (rpt === 'no' && !a.hasFiles);
                return m && mc && ms && mr;
            });
            const tbody = document.getElementById('appTbody');
            if (!list.length) { tbody.innerHTML = '<tr><td colspan="12" style="text-align:center;padding:24px;color:var(--muted)">Bu kategoride kayıt bulunamadı</td></tr>'; return; }
            tbody.innerHTML = list.map((a, i) => `
    <tr>
      <td style="color:var(--muted)">${i + 1}</td>
      <td><b>${a.name || '—'}</b></td>
      <td style="font-size:12px;color:var(--muted)">${a.email || '—'}</td>
      <td><span class="cat-badge" style="background:${CAT_COLORS[a.category] || '#4f8ef7'}22;color:${CAT_COLORS[a.category] || '#4f8ef7'}">${a.category || '—'}</span></td>
      <td style="font-size:13px">${a.school || '—'}</td>
      <td style="font-size:13px">${a.birth || '—'}</td>
      <td style="font-size:13px">${a.parent || '—'}</td>
      <td style="font-size:13px">${a.team || '—'}</td>
      <td>${a.hasFiles ? '<span class="report-badge rpt-yes">📎 ' + (a.fileCount || '?') + ' dosya</span>' : (a.siteId ? '<span class="report-badge rpt-no">❌ Yok</span>' : '<span class="report-badge rpt-unknown">—</span>')}</td>
      <td style="color:var(--muted);font-size:12px">${fmtDate(a.createdAt)}</td>
      <td>
        <select class="status-select badge-${stCls(a.status)}" onchange="changeStatus('${a.id}',this.value)">
          <option ${a.status === 'Bekliyor' ? 'selected' : ''}>Bekliyor</option>
          <option ${a.status === 'Onaylandı' ? 'selected' : ''}>Onaylandı</option>
          <option ${a.status === 'Reddedildi' ? 'selected' : ''}>Reddedildi</option>
        </select>
      </td>
      <td style="display:flex;gap:6px">
        <button class="btn-icon2" onclick="openEditModal('${a.id}')">✏️</button>
        <button class="btn-icon2" onclick="downloadRegistrationFiles(${a.siteId || 0},'${escQ(a.name)}')" title="Dosyaları İndir" ${!a.siteId ? 'disabled style="opacity:0.3"' : ''}>📎</button>
        <button class="btn-icon2 del" onclick="confirmDelete('${a.id}','${escQ(a.name)}')">🗑️</button>
      </td>
    </tr>`).join('');
            const catLabel = activeCat || 'Tüm Kategoriler';
            document.getElementById('appFooter').textContent = `${list.length} kayıt — ${catLabel} (toplam: ${apps.length})`;
        }
        function clearAppFilters() {
            document.getElementById('appSearch').value = '';
            document.getElementById('filterStatus').value = '';
            document.getElementById('filterReport').value = '';
            activeCat = '';
            renderCatTabs();
            renderAppTable();
        }
        function changeStatus(id, status) {
            apps = apps.map(a => a.id === id ? { ...a, status } : a);
            saveApps(apps);
            showToast(`Durum güncellendi: ${status}`, 'success');
            renderDashboard();
            renderCatTabs();
        }

        // ─── ADD/EDIT MODAL ──────────────────────────────────────
        function openAddModal() {
            editingId = null;
            document.getElementById('appModalTitle').textContent = 'Yeni Başvuru';
            document.getElementById('appForm').reset();
            document.getElementById('fStatus').value = 'Bekliyor';
            document.getElementById('appModalOverlay').classList.add('open');
        }
        function openEditModal(id) {
            const a = apps.find(x => x.id === id); if (!a) return;
            editingId = id;
            document.getElementById('appModalTitle').textContent = 'Başvuru Düzenle';
            document.getElementById('fName').value = a.name || '';
            document.getElementById('fEmail').value = a.email || '';
            document.getElementById('fCat').value = a.category || '';
            document.getElementById('fSchool').value = a.school || '';
            document.getElementById('fBirth').value = a.birth || '';
            document.getElementById('fParent').value = a.parent || '';
            document.getElementById('fTeam').value = a.team || '';
            document.getElementById('fStatus').value = a.status || 'Bekliyor';
            document.getElementById('fNotes').value = a.notes || '';
            document.getElementById('appModalOverlay').classList.add('open');
        }
        function closeAppModal() { document.getElementById('appModalOverlay').classList.remove('open'); editingId = null; }
        function saveApp(e) {
            e.preventDefault();
            const data = {
                name: document.getElementById('fName').value.trim(),
                email: document.getElementById('fEmail').value.trim(),
                category: document.getElementById('fCat').value,
                school: document.getElementById('fSchool').value.trim(),
                birth: document.getElementById('fBirth').value.trim(),
                parent: document.getElementById('fParent').value.trim(),
                team: document.getElementById('fTeam').value.trim(),
                status: document.getElementById('fStatus').value,
                notes: document.getElementById('fNotes').value.trim(),
            };
            if (editingId) {
                apps = apps.map(a => a.id === editingId ? { ...a, ...data } : a);
                showToast('Başvuru güncellendi ✅', 'success');
            } else {
                apps.unshift({ ...data, id: uid(), createdAt: isoNow() });
                showToast('Başvuru eklendi 🎉', 'success');
            }
            saveApps(apps);
            closeAppModal();
            renderDashboard();
            renderAppTable();
        }

        // ─── DELETE ──────────────────────────────────────────────
        function confirmDelete(id, name) {
            document.getElementById('confirmMsg').innerHTML = `<b>${name}</b> silinecek. Bu işlem geri alınamaz!`;
            confirmCb = () => { apps = apps.filter(a => a.id !== id); saveApps(apps); showToast('Silindi', 'info'); renderDashboard(); renderAppTable(); };
            document.getElementById('confirmOverlay').classList.add('open');
        }
        function closeConfirm() { document.getElementById('confirmOverlay').classList.remove('open'); confirmCb = null; }
        function doConfirm() { if (confirmCb) confirmCb(); closeConfirm(); }

        // ─── EXCEL EXPORT ──────────────────────────────────────────
        const COL_WIDTHS = [{ wch: 5 }, { wch: 25 }, { wch: 32 }, { wch: 26 }, { wch: 36 }, { wch: 15 }, { wch: 18 }, { wch: 22 }, { wch: 18 }, { wch: 12 }];
        const HEADERS = ['#', 'Ad Soyad', 'E-posta', 'Kategori', 'Okul / Üniversite', 'Doğum Tarihi', 'Veli İletişim', 'Takım Adı', 'Kayıt Tarihi', 'Durum'];

        function appToRow(a, i) {
            return [
                i + 1,
                a.name || '',
                a.email || '',
                a.category || '',
                a.school || '',
                a.birth || '',
                a.parent || '',
                a.team || '',
                fmtDate(a.createdAt),
                a.status || 'Bekliyor'
            ];
        }
        function makeSheet(list) {
            const rows = [HEADERS, ...list.map((a, i) => appToRow(a, i))];
            const ws = XLSX.utils.aoa_to_sheet(rows);
            ws['!cols'] = COL_WIDTHS;
            return ws;
        }
        // Sheet adı için yasak karakterleri temizle ve 31 kar sınırına uyu
        function safeSheetName(name) {
            return name.replace(/[\\/\*\[\]\?:]/g, '').substring(0, 31);
        }

        // Güvenilir indirme — file:// protokolünde de çalışır
        function downloadXLSX(wb, filename) {
            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(url), 1500);
        }

        function exportExcel() {
            if (!apps.length) { showToast('Dışa aktarılacak kayıt yok!', 'error'); return; }
            const wb = XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(wb, makeSheet(apps), 'Tum Kayitlar');

            const realCats = [...new Set(apps.map(a => a.category).filter(Boolean))].sort();
            realCats.forEach(cat => {
                const catApps = apps.filter(a => a.category === cat);
                XLSX.utils.book_append_sheet(wb, makeSheet(catApps), safeSheetName(cat));
            });

            const date = new Date().toLocaleDateString('tr-TR').replace(/\./g, '-');
            downloadXLSX(wb, `basvurular_${date}.xlsx`);
            showToast(`Excel indirildi ✅ (${XLSX.utils.book_get_sheet_names(wb).length} sheet)`, 'success');
        }

        function exportCatExcel() {
            const list = activeCat ? apps.filter(a => a.category === activeCat) : apps;
            if (!list.length) { showToast('Bu kategoride kayıt yok!', 'error'); return; }
            const wb = XLSX.utils.book_new();
            const sheetName = safeSheetName(activeCat || 'Tum Kayitlar');
            XLSX.utils.book_append_sheet(wb, makeSheet(list), sheetName);
            const date = new Date().toLocaleDateString('tr-TR').replace(/\./g, '-');
            downloadXLSX(wb, `basvurular_${sheetName}_${date}.xlsx`);
            showToast('Excel indirildi ✅', 'success');
        }


        // ─── EXCEL IMPORT ───────────────────────────────────────
        // Sütun başlık eşleştirme (büyük/küçük harf duyarsız)
        const COL_MAP = {
            // Kategori
            'kategori': 'category', 'category': 'category', 'cat': 'category',
            // Ad Soyad
            'ad soyad': 'name', 'name': 'name', 'adsoyad': 'name', 'ad': 'name', 'full name': 'name', 'fullname': 'name',
            // E-posta
            'e-posta': 'email', 'eposta': 'email', 'email': 'email', 'mail': 'email', 'e posta': 'email',
            // Okul — TÜM varyantlar
            'okul': 'school', 'üniversite': 'school', 'okul/üniversite': 'school', 'üniversite/okul': 'school',
            'üniversite / okul': 'school', 'okul / üniversite': 'school',
            'university': 'school', 'school': 'school', 'university/school': 'school',
            'okul adı': 'school', 'üniversite adı': 'school', 'kurum': 'school',
            // Doğum tarihi
            'doğum tarihi': 'birth', 'dogum tarihi': 'birth', 'birth': 'birth', 'doğum': 'birth',
            'doğum tarihi ': 'birth', 'birth date': 'birth', 'birthdate': 'birth',
            // Veli
            'veli iletişim': 'parent', 'veli': 'parent', 'parent': 'parent', 'veli iletisim': 'parent',
            'veli bilgi': 'parent', 'guardian': 'parent', 'veli tel': 'parent', 'veli telefon': 'parent',
            // Takım
            'takım adı': 'team', 'takim adi': 'team', 'team': 'team', 'takım': 'team',
            'team name': 'team', 'teamname': 'team',
        };
        function normalizeHeader(h) { return (h || '').toString().toLowerCase().trim().replace(/\s+/g, ' '); }

        function handleFile(file) {
            if (!file) return;
            const reader = new FileReader();
            reader.onload = e => {
                try {
                    const wb = XLSX.read(e.target.result, { type: 'binary' });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                    if (raw.length < 2) { showToast('Dosya boş veya hatalı!', 'error'); return; }

                    // Header mapping
                    const headers = raw[0].map(normalizeHeader);
                    const fieldIdx = {};
                    headers.forEach((h, i) => { const f = COL_MAP[h]; if (f) fieldIdx[f] = i; });

                    // Hangi alanların bulunduğunu logla (debug)
                    console.log('Excel başlıkları:', raw[0]);
                    console.log('Alan eşleştirmesi:', fieldIdx);

                    // Satırları dönüştür
                    const rows = raw.slice(1).filter(r => r.some(c => c !== '')).map(r => ({
                        category: (r[fieldIdx.category] || '').toString().trim() || 'Serbest',
                        name: (r[fieldIdx.name] || '').toString().trim(),
                        email: (r[fieldIdx.email] || '').toString().trim().toLowerCase(),
                        school: fieldIdx.school !== undefined ? (r[fieldIdx.school] || '').toString().trim() : '',
                        birth: (r[fieldIdx.birth] || '').toString().trim(),
                        parent: (r[fieldIdx.parent] || '').toString().trim(),
                        team: (r[fieldIdx.team] || '').toString().trim(),
                    })).filter(r => r.name || r.email);

                    if (!rows.length) { showToast('İçe aktarılacak geçerli satır bulunamadı!', 'error'); return; }

                    // Kategori filtresi uygulanmış mı?
                    const importCatFilter = document.getElementById('importCatFilter').value;
                    const filteredRows = importCatFilter
                        ? rows.filter(r => r.category === importCatFilter)
                        : rows;

                    pendingImportRows = filteredRows;

                    // Önizleme hesapla
                    const existingKeys = new Set(apps.map(a => `${a.email}__${a.category}`.toLowerCase()));
                    const newRows = filteredRows.filter(r => !existingKeys.has(`${r.email}__${r.category}`.toLowerCase()));
                    const dupRows = filteredRows.filter(r => existingKeys.has(`${r.email}__${r.category}`.toLowerCase()));
                    // Mevcut kayıtlar içinde okul boş olanlar — güncellenecek
                    const toUpdateCount = dupRows.filter(r => {
                        const existing = apps.find(a => `${a.email}__${a.category}`.toLowerCase() === `${r.email}__${r.category}`.toLowerCase());
                        return existing && (!existing.school && r.school);
                    }).length;

                    const res = document.getElementById('importResult');
                    res.style.display = 'block';
                    res.className = 'import-result ' + (newRows.length > 0 || toUpdateCount > 0 ? 'res-ok' : 'res-warn');
                    const filterInfo = importCatFilter ? ` (» ${importCatFilter})` : ' (Tüm kategoriler)';
                    res.innerHTML = `📊 <b>${rows.length}</b> toplam satır &nbsp;|&nbsp; 🔍 <b>${filteredRows.length}</b> süzüldükte${filterInfo} &nbsp;|&nbsp; 🆕 <b>${newRows.length}</b> yeni &nbsp;|&nbsp; 🔄 <b>${toUpdateCount}</b> eksik okul/bilgi güncellenecek`;

                    // Okul alanı gelmediyse uyar
                    if (fieldIdx.school === undefined) {
                        res.innerHTML += `<br>⚠️ <b>Dikkat:</b> Excel'de okul sütunu eşleştirilemedi. Başlık: "${raw[0].join(' | ')}"`;
                    }

                    // Önizleme (ilk 5 — yeni + güncellenecekler karışık)
                    const prev = document.getElementById('importPreview');
                    const pt = document.getElementById('previewTable');
                    const previewRows = [...newRows, ...dupRows.filter(r => {
                        const ex = apps.find(a => `${a.email}__${a.category}`.toLowerCase() === `${r.email}__${r.category}`.toLowerCase());
                        return ex && (!ex.school && r.school);
                    })].slice(0, 5);
                    pt.innerHTML = `<thead><tr><th>Ad Soyad</th><th>E-posta</th><th>Kategori</th><th>Okul</th><th>Takım</th></tr></thead>
        <tbody>${previewRows.map(r => `<tr><td>${r.name}</td><td>${r.email}</td><td>${r.category}</td><td>${r.school || '—'}</td><td>${r.team || '—'}</td></tr>`).join('')}</tbody>`;
                    prev.style.display = (newRows.length > 0 || toUpdateCount > 0) ? 'block' : 'none';
                    document.getElementById('importBtn').dataset.count = newRows.length;
                } catch (err) { showToast('Dosya okunamadı: ' + err.message, 'error'); }
            };
            reader.readAsBinaryString(file);
        }

        function confirmImport() {
            // Mevcut kayıtların key map'i (email__category → id)
            const keyToId = {};
            apps.forEach(a => { keyToId[`${a.email}__${a.category}`.toLowerCase()] = a.id; });

            const newRows = [];
            let updatedCount = 0;

            pendingImportRows.forEach(r => {
                const key = `${r.email}__${r.category}`.toLowerCase();
                if (keyToId[key]) {
                    // Mevcut kayıt var — boş alanları güncelle
                    apps = apps.map(a => {
                        if (a.id !== keyToId[key]) return a;
                        let changed = false;
                        const update = { ...a };
                        if (!a.school && r.school) { update.school = r.school; changed = true; }
                        if (!a.parent && r.parent) { update.parent = r.parent; changed = true; }
                        if (!a.team && r.team) { update.team = r.team; changed = true; }
                        if (!a.birth && r.birth) { update.birth = r.birth; changed = true; }
                        if (changed) updatedCount++;
                        return update;
                    });
                } else {
                    newRows.push(r);
                }
            });

            const newApps = newRows.map(r => ({ ...r, id: uid(), status: 'Bekliyor', createdAt: isoNow() }));
            apps = [...newApps, ...apps];
            saveApps(apps);

            let msg = '';
            if (newApps.length) msg += `✅ ${newApps.length} yeni kayıt eklendi`;
            if (updatedCount) msg += (msg ? ', ' : '') + `🔄 ${updatedCount} kayıttaki eksik okul/veli/takım bilgisi güncellendi`;
            if (!newApps.length && !updatedCount) msg = 'ℹ️ Eklenecek veya güncellenecek kayıt bulunamadı';
            showToast(msg, newApps.length || updatedCount ? 'success' : 'info');

            // Reset
            document.getElementById('importResult').style.display = 'none';
            document.getElementById('importPreview').style.display = 'none';
            document.getElementById('fileInput').value = '';
            pendingImportRows = [];
            renderDashboard();
            renderCatTabs();
        }

        // Drag & drop
        const dz = document.getElementById('dropZone');
        dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
        dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
        dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag-over'); handleFile(e.dataTransfer.files[0]); });

        // ─── SETTINGS ────────────────────────────────────────────
        function saveSettings() {
            const u = document.getElementById('newUser').value.trim();
            const p = document.getElementById('newPass').value;
            if (!u || p.length < 6) { showToast('En az 6 karakterli şifre girin!', 'error'); return; }
            localStorage.setItem('bs_admin_user', u);
            localStorage.setItem('bs_admin_pass', p);
            showToast('Ayarlar kaydedildi ✅', 'success');
        }
        function clearAllData() {
            if (!confirm('TÜM başvurular silinecek! Emin misiniz?')) return;
            apps = [];
            saveApps(apps);
            showToast('Tüm veriler silindi', 'info');
            renderDashboard();
        }

        // ─── MÜKERRERLERİ TARA / TEMİZLE ────────────────────────
        function scanDuplicates() {
            // E-posta bazında gruplama (büyük/küçük harf duyarsız)
            const groups = {};
            apps.forEach(a => {
                const key = (a.email || '').toLowerCase().trim();
                if (!key) return;
                if (!groups[key]) groups[key] = [];
                groups[key].push(a);
            });
            const dupGroups = Object.values(groups).filter(g => g.length > 1);
            const dupTotal = dupGroups.reduce((s, g) => s + g.length - 1, 0);
            const info = document.getElementById('dupInfo');
            const btn = document.getElementById('btnRemoveDups');
            if (dupTotal === 0) {
                info.innerHTML = '✅ Mükerrer kayıt yok, her e-posta benzersiz.';
                info.style.color = '#3fb950';
                btn.style.display = 'none';
            } else {
                info.innerHTML = `⚠️ <b>${dupGroups.length}</b> e-postada <b>${dupTotal}</b> fazladan kayıt bulundu.`;
                info.style.color = '#f0883e';
                btn.style.display = 'inline-flex';
            }
        }
        function removeDuplicates() {
            if (!confirm('Mükerrer kayıtlar silinecek. En eski tarihli kayıt korunur. Devam edilsin mi?')) return;
            const groups = {};
            apps.forEach(a => {
                const key = (a.email || '').toLowerCase().trim();
                if (!key) { groups['__no_email_' + a.id] = [a]; return; }
                if (!groups[key]) groups[key] = [];
                groups[key].push(a);
            });
            const cleaned = [];
            let removed = 0;
            Object.values(groups).forEach(g => {
                if (g.length === 1) { cleaned.push(g[0]); return; }
                // En eski tarihi tut (createdAt küçük = daha eski)
                g.sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
                cleaned.push(g[0]);
                removed += g.length - 1;
            });
            apps = cleaned;
            saveApps(apps);
            showToast(`🧹 ${removed} mükerrer kayıt silindi!`, 'success');
            document.getElementById('dupInfo').innerHTML = `✅ ${removed} kayıt temizlendi. Kalan: ${apps.length} benzersiz kayıt.`;
            document.getElementById('dupInfo').style.color = '#3fb950';
            document.getElementById('btnRemoveDups').style.display = 'none';
            renderDashboard();
            renderCatTabs();
        }

        // ─── TAM SİSTEM SIFIRLAMA ────────────────────────────────
        function fullReset() {
            if (!confirm('⚠️ TÜM veriler, admin şifresi ve Drive ayarları silinecek!\nSistem fabrika ayarlarına dönecek.\n\nDevam edilsin mi?')) return;
            if (!confirm('Son kez onaylıyın: Her şey kalıcı olarak silinecek!')) return;
            localStorage.clear();
            showToast('Sistem sıfırlandı, giriş sayfasına yönlendiriliyorsunuz...', 'info');
            setTimeout(() => window.location.href = 'index.html', 1500);
        }


        // ─── GOOGLE SHEETS ENTEGRASYONU ─────────────────────────────
        const SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';
        let sheetsToken = null;

        function loadDriveSettings() {
            const url = localStorage.getItem('bs_script_url') || '';
            const el = document.getElementById('scriptUrl');
            if (el) el.value = url;
            document.getElementById('btnDrive').style.display = url ? 'inline-flex' : 'none';
        }
        function saveDriveSettings() {
            const url = document.getElementById('scriptUrl').value.trim();
            if (!url.startsWith('https://')) { showToast('Geçerli bir URL girin!', 'error'); return; }
            localStorage.setItem('bs_script_url', url);
            loadDriveSettings();
            showToast('Kaydedildi ✅ Sheets\'e Aktar butonu aktif!', 'success');
        }

        async function syncToSheets() {
            if (!apps.length) { showToast('Aktarılacak kayıt yok!', 'error'); return; }
            const scriptUrl = localStorage.getItem('bs_script_url');
            if (!scriptUrl) { showToast("Önce Ayarlar'dan Script URL'sini girin!", 'error'); return; }

            const btn = document.getElementById('btnDrive');
            btn.textContent = '⏳ Aktarılıyor...';
            btn.disabled = true;
            try {
                const rows = [HEADERS, ...apps.map((a, i) => appToRow(a, i))];
                // Google Apps Script no-cors için JSON gövdesinin gitmesi adına text/plain kullanıyoruz
                await fetch(scriptUrl, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify({ sheetName: 'Tüm Kayıtlar', rows: rows })
                });
                const now = new Date().toLocaleString('tr-TR');
                const link = 'https://docs.google.com/spreadsheets/d/1Sjd2CA2vD6GV7zKXOatiipniFxLD2m0L16aQnccau9o/edit';
                showToast(`📊 Tüm Kayıtlar Sheets'e aktarıldı!`, 'success');
                const st = document.getElementById('driveStatus');
                if (st) { st.style.display = 'block'; st.innerHTML = `✅ Son aktarım: <b>${now}</b> — <a href="${link}" target="_blank" style="color:#4f8ef7">Tabloyu aç →</a>`; }
            } catch (err) {
                showToast('Hata: ' + err.message, 'error');
            } finally {
                btn.textContent = "📊 Sheets'e Aktar";
                btn.disabled = false;
            }
        }

        async function syncCatToSheets() {
            const cat = activeCat;
            if (!cat) { showToast('Lütfen önce bir kategori seçin!', 'error'); return; }
            const catApps = apps.filter(a => a.category === cat);
            if (!catApps.length) { showToast('Bu kategoride aktarılacak kayıt yok!', 'error'); return; }

            const scriptUrl = localStorage.getItem('bs_script_url');
            if (!scriptUrl) { showToast("Önce Ayarlar'dan Script URL'sini girin!", 'error'); return; }

            const btn = document.getElementById('btnSyncCat');
            const origText = btn.textContent;
            btn.textContent = '⏳ Aktarılıyor...';
            btn.disabled = true;
            try {
                const rows = [HEADERS, ...catApps.map((a, i) => appToRow(a, i))];
                await fetch(scriptUrl, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify({ sheetName: cat.substring(0, 90), rows: rows })
                });
                showToast(`📊 "${cat}" kategorisi Sheets'e aktarıldı!`, 'success');
            } catch (err) {
                showToast('Hata: ' + err.message, 'error');
            } finally {
                btn.textContent = origText;
                btn.disabled = false;
            }
        }


        // ─── LOGOUT ──────────────────────────────────────────────
        function doLogout() { sessionStorage.removeItem('bs_auth'); window.location.href = 'index.html'; }

        // ─── HELPERS ─────────────────────────────────────────────
        function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
        function isoNow() { return new Date().toISOString(); }
        function todayStr() { return new Date().toISOString().split('T')[0]; }
        function parseTurkishDate(dStr) {
            if(!dStr) return isoNow();
            try {
                const parts = dStr.split(' ');
                const dp = parts[0].split(/[./-]/);
                if(dp.length===3) {
                    const y=dp[2].length===2?'20'+dp[2]:dp[2], m=dp[1], d=dp[0], t=parts[1]||'00:00';
                    const dt = new Date(`${y}-${m}-${d}T${t}:00`);
                    if(!isNaN(dt.getTime())) return dt.toISOString();
                }
                const fallback = new Date(dStr);
                if(!isNaN(fallback.getTime())) return fallback.toISOString();
            } catch(e) {}
            return isoNow();
        }
        function fmtDate(iso) { if (!iso) return '—'; try { const d = new Date(iso); return d.toLocaleDateString('tr-TR') + ' ' + d.toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' }); } catch { return iso; } }
        function stCls(s) { if (!s || s === 'Bekliyor') return 'pending'; if (s === 'Onaylandı') return 'approved'; return 'rejected'; }
        function escQ(s) { return (s || '').replace(/'/g, "\\'"); }
        function showToast(msg, type = 'info') { const t = document.getElementById('toast'); t.textContent = msg; t.className = `toast show ${type}`; clearTimeout(window._tt); window._tt = setTimeout(() => t.className = 'toast', 3000); }

        // ─── SENKRONIZASYON ──────────────────────────────────────
        let syncConnected = false;

        async function connectToRobotek() {
          const email = document.getElementById('syncEmail').value.trim();
          const pass = document.getElementById('syncPassword').value.trim();
          if (!email || !pass) { showToast('E-posta ve şifre girin!', 'error'); return; }
          
          const btn = document.getElementById('btnConnect');
          btn.textContent = '⏳ Bağlanıyor...';
          btn.disabled = true;
          
          try {
            if (window.robotekAPI) {
              const result = await window.robotekAPI.login(email, pass);
              if (result.success) {
                syncConnected = true;
                updateSyncUI(true);
                showToast('ROBOTEK sitesine bağlandı! ✅', 'success');
                addSyncLog('✅ Siteye başarıyla bağlandı');
                // Save email (not password) for convenience
                localStorage.setItem('bs_sync_email', email);
              } else {
                showToast('Bağlantı hatası: ' + result.message, 'error');
                addSyncLog('❌ Bağlantı hatası: ' + result.message);
              }
            } else {
              showToast('Bu özellik sadece Electron uygulamasında çalışır!', 'error');
            }
          } catch (err) {
            showToast('Hata: ' + err.message, 'error');
          } finally {
            btn.textContent = '🔗 Bağlan';
            btn.disabled = false;
          }
        }

        function disconnectRobotek() {
          if (window.robotekAPI) window.robotekAPI.logout();
          syncConnected = false;
          updateSyncUI(false);
          showToast('Bağlantı kesildi', 'info');
          addSyncLog('🔌 Bağlantı kesildi');
        }

        function updateSyncUI(connected) {
          const dot = document.querySelector('.sync-dot');
          const statusText = document.querySelector('#syncStatusIndicator span');
          const loginForm = document.getElementById('syncLoginForm');
          const actionsCard = document.getElementById('syncActionsCard');
          const btnConnect = document.getElementById('btnConnect');
          const btnDisconnect = document.getElementById('btnDisconnect');
          
          if (connected) {
            dot.className = 'sync-dot online';
            statusText.textContent = 'Bağlı';
            actionsCard.style.opacity = '1';
            actionsCard.style.pointerEvents = 'auto';
            btnConnect.style.display = 'none';
            btnDisconnect.style.display = 'inline-flex';
          } else {
            dot.className = 'sync-dot offline';
            statusText.textContent = 'Bağlı Değil';
            actionsCard.style.opacity = '0.5';
            actionsCard.style.pointerEvents = 'none';
            btnConnect.style.display = 'inline-flex';
            btnDisconnect.style.display = 'none';
          }
        }

        async function syncAllRegistrations() {
          if (!window.robotekAPI) { showToast('Electron gerekli!', 'error'); return; }
          
          showSyncProgress(true);
          updateSyncProgress(0, 'Başvurular çekiliyor...');
          addSyncLog('🔄 Senkronizasyon başlatıldı...');
          
          try {
            const result = await window.robotekAPI.fetchAllRegistrations();
            
            if (result.success && result.registrations) {
              const siteRegs = result.registrations;
              addSyncLog(`📊 ${siteRegs.length} başvuru siteden çekildi`);
              
              let merged = 0, updated = 0;
              
              const existingNames = new Map();
              apps.forEach(a => {
                  const n = (a.name || '').toLowerCase().replace(/\s+/g, '');
                  if (!existingNames.has(n)) existingNames.set(n, []);
                  existingNames.get(n).push(a);
              });
              
              siteRegs.forEach(reg => {
                  const n = (reg.name || '').toLowerCase().replace(/\s+/g, '');
                  if (existingNames.has(n)) {
                      existingNames.get(n).forEach(a => {
                          let isChanged = false;
                          if (a.siteId !== reg.siteId) {
                              a.siteId = reg.siteId;
                              isChanged = true;
                          }
                          if (a.hasFiles !== (reg.hasFiles || false)) {
                              a.hasFiles = reg.hasFiles || false;
                              isChanged = true;
                          }
                          if (reg.category) {
                              const mappedCat = mapCategory(reg.category);
                              if (a.category !== mappedCat) {
                                  a.category = mappedCat;
                                  isChanged = true;
                              }
                          }
                          if (isChanged) {
                              updated++;
                          }
                      });
                      merged++;
                  }
              });

              saveApps(apps);
              
              const msg = `✅ Senkronizasyon tamamlandı: ${updated} güncellendi, ${merged} eşleşti`;
              addSyncLog(msg);
              showToast(msg, 'success');
              
              document.getElementById('syncLastTime').textContent = 'Son senkronizasyon: ' + new Date().toLocaleString('tr-TR');
              localStorage.setItem('bs_last_sync', new Date().toISOString());
              
              renderDashboard();
              renderCatTabs();
              renderAppTable();
            } else {
              showToast('Senkronizasyon hatası: ' + (result.error || 'Bilinmeyen hata'), 'error');
              addSyncLog('❌ Hata: ' + (result.error || 'Bilinmeyen hata'));
            }
          } catch (err) {
            showToast('Hata: ' + err.message, 'error');
            addSyncLog('❌ Hata: ' + err.message);
          } finally {
            showSyncProgress(false);
          }
        }

        async function syncAndDownloadFiles() {
          addSyncLog('🔄 Aşama 1/2: Siteden kayıtlar çekiliyor...');
          await syncAllRegistrations();
          
          addSyncLog('📋 Kategoriler: ' + DOWNLOAD_CATEGORIES.join(', '));
          
          const withFiles = apps.filter(a => {
            if (!a.siteId) return false;
            const cat = (a.category || '').toLowerCase().trim();
            return DOWNLOAD_CATEGORIES.some(c => {
              const cLower = c.toLowerCase();
              if (cat.includes(cLower)) return true;
              if (c === 'İnsansız Hava Araçları' && (cat.includes('iha') || cat.includes('insansiz') || cat.includes('insansız') || cat.includes('hava'))) return true;
              if (c === 'Tasarla ve Geliştir' && (cat.includes('tasarla') || cat.includes('geliştir') || cat.includes('gelistir'))) return true;
              if (c === 'Savaşan Robot' && (cat.includes('savasan') || cat.includes('savaşan'))) return true;
              if (c === 'Serbest' && cat.includes('serbest')) return true;
              return false;
            });
          });
          
          if (!withFiles.length) {
            showToast('Bu 4 kategoride siteId eşleşen kayıt yok!', 'info');
            addSyncLog('⚠️ İndirilecek kayıt yok.');
            return;
          }
          
          const DELAY_SEC = 8;
          addSyncLog(`📎 Aşama 2/2: ${withFiles.length} kayıt için ZIP indiriliyor...`);
          addSyncLog(`⏱️ Tahmini süre: ~${Math.ceil(withFiles.length * DELAY_SEC / 60)} dk`);
          
          showSyncProgress(true);
          let downloaded = 0, failed = 0;
          const startTime = Date.now();
          
          for (let i = 0; i < withFiles.length; i++) {
            const entry = withFiles[i];
            const etaMin = Math.ceil((withFiles.length - i) * DELAY_SEC / 60);
            
            updateSyncProgress(
              ((i + 1) / withFiles.length) * 100,
              `[${i + 1}/${withFiles.length}] ${entry.name || '?'} — ~${etaMin} dk kaldı`
            );
            
            if (i > 0) {
              await new Promise(r => setTimeout(r, DELAY_SEC * 1000));
            }
            
            try {
              addSyncLog(`📥 [${i + 1}/${withFiles.length}] ${entry.name} (${entry.category})`);
              const result = await window.robotekAPI.downloadAllFiles(entry.siteId, entry.category, entry.name);
              
              if (result.success) {
                downloaded++;
                addSyncLog(`✅ ${entry.name} — indirildi`);
              } else {
                failed++;
                addSyncLog(`❌ ${entry.name} — ${result.error}`);
              }
            } catch (err) {
              failed++;
              addSyncLog(`❌ ${entry.name} — ${err.message}`);
            }
          }
          
          const totalSec = Math.round((Date.now() - startTime) / 1000);
          showSyncProgress(false);
          
          const summary = `🏁 TAMAMLANDI: ${downloaded}/${withFiles.length} başarılı — ${Math.floor(totalSec/60)}dk ${totalSec%60}sn`;
          addSyncLog(summary);
          showToast(summary, downloaded > 0 ? 'success' : 'error');
        }

        // Category mapping from site names to our categories
        function mapCategory(siteCat) {
          if (!siteCat) return 'Serbest';
          const clean = siteCat.toLowerCase().replace(/_/g, ' ').trim();
          
          const map = {
            'cizgi izleyen': 'Çizgi İzleyen Robot',
            'çizgi izleyen': 'Çizgi İzleyen Robot',
            'cizgi_izleyen': 'Çizgi İzleyen Robot',
            'insansiz hava': 'İnsansız Hava Araçları',
            'insansız hava': 'İnsansız Hava Araçları',
            'savasan iha': 'İnsansız Hava Araçları',
            'savaşan iha': 'İnsansız Hava Araçları',
            'iha': 'İnsansız Hava Araçları',
            'mini sumo': 'Mini Sumo Robot',
            'mini_sumo': 'Mini Sumo Robot',
            'labirent': 'Labirent Çözen',
            'labirent çözen': 'Labirent Çözen',
            'savasan robot': 'Savaşan Robot',
            'savaşan robot': 'Savaşan Robot',
            'robo futbol': 'Robo Futbol',
            'robo_futbol': 'Robo Futbol',
            'robo futbol lise': 'Robo Futbol',
            'robo_football_lise_universite': 'Robo Futbol',
            'tasarla': 'Tasarla ve Geliştir',
            'tasarla ve geliştir': 'Tasarla ve Geliştir',
            'tasarla gelistir': 'Tasarla ve Geliştir',
            'serbest': 'Serbest'
          };
          
          for (const [key, val] of Object.entries(map)) {
            if (clean.includes(key)) return val;
          }
          
          // Try exact match with existing categories
          const found = CATEGORIES.find(c => c.toLowerCase() === clean);
          if (found) return found;
          
          return siteCat; // Return as-is if no match
        }

        function showSyncProgress(show) {
          document.getElementById('syncProgress').style.display = show ? 'block' : 'none';
        }

        function updateSyncProgress(pct, text) {
          document.getElementById('syncProgressFill').style.width = Math.min(100, pct) + '%';
          document.getElementById('syncProgressText').textContent = text;
        }

        function addSyncLog(msg) {
          const log = document.getElementById('syncLog');
          const entries = document.getElementById('syncLogEntries');
          log.style.display = 'block';
          const time = new Date().toLocaleTimeString('tr-TR');
          entries.innerHTML = `<div class="sync-log-entry"><span class="sync-log-time">${time}</span> ${msg}</div>` + entries.innerHTML;
          // Keep only last 30
          while (entries.children.length > 30) entries.removeChild(entries.lastChild);
        }

        async function openDownloadFolder() {
          if (window.robotekAPI) {
            await window.robotekAPI.openDownloadFolder();
          } else {
            showToast('Bu özellik sadece Electron\'da çalışır', 'error');
          }
        }

        // ─── DOSYA İNDİRME (Tek kayıt) ──────────────────────────
        async function downloadRegistrationFiles(siteId, name) {
          if (!window.robotekAPI) { showToast('Electron gerekli!', 'error'); return; }
          if (!siteId) { showToast('Bu kaydın site ID\'si yok. Önce senkronize edin.', 'error'); return; }
          
          showToast(`${name} dosyaları indiriliyor...`, 'info');
          try {
            const result = await window.robotekAPI.downloadAllFiles(siteId);
            if (result.success) {
              showToast(`${result.downloadedCount} dosya indirildi! 📂`, 'success');
            } else {
              showToast('Hata: ' + result.error, 'error');
            }
          } catch (err) {
            showToast('Hata: ' + err.message, 'error');
          }
        }

        // Setup sync event listeners
        if (window.robotekAPI) {
          window.robotekAPI.onSyncProgress((data) => {
            if (data.phase === 'details') {
              updateSyncProgress((data.current / data.total) * 100, `Detay çekiliyor: ${data.current}/${data.total} — ${data.name || ''}`);
            } else {
              updateSyncProgress((data.fetched / (data.total || 1)) * 100, `Sayfa ${data.page}/${data.totalPages || '?'} — ${data.fetched} kayıt çekildi`);
            }
          });

          // Backend loglarını da senkronizasyon loguna aktar
          window.robotekAPI.onLog((msg) => {
            console.log('[Scraper]', msg);
            // Sadece önemli mesajları ekrana yansıt
            if (msg.includes('Login') || msg.includes('ZIP') || msg.includes('hatası') || msg.includes('başarılı') || msg.includes('Oturum')) {
              addSyncLog('🔧 ' + msg);
            }
          });
        }

        // ─── INIT ────────────────────────────────────────────────
        renderDashboard();
        renderReportStatusGrid();
        loadDriveSettings();

        // Örnek veri yükle (ilk açılışta)
        if (!apps.length) {
            const sampleData = [
                { id: uid(), name: 'Azra Selin Türk', email: 'ozelemkarasc7@gmail.com', category: 'Robo Futbol', school: 'Mustafa Kutlu İlkokulu', birth: '03.01.2019', parent: '05392778121', team: 'HAROT', status: 'Bekliyor', createdAt: new Date('2026-04-16T13:39:00').toISOString() },
                { id: uid(), name: 'Halil', email: 'halilkus9999@gmail.com', category: 'İnsansız Hava Araçları', school: 'Mustafa barut anadolu lisesi', birth: '19.02.2007', parent: 'Feconi', team: 'den', status: 'Bekliyor', createdAt: new Date('2026-04-16T12:58:00').toISOString() },
                { id: uid(), name: 'Rahman', email: 'rahmanşahin040@hotmail.com', category: 'Çizgi İzleyen Robot', school: 'Bolu Sosyal Bilimler Lisesi', birth: '05.04.2011', parent: '05315341419', team: 'rtons', status: 'Bekliyor', createdAt: new Date('2026-04-15T19:58:00').toISOString() },
                { id: uid(), name: 'Şükrü Acıyan', email: 'subruaciyan@gmail.com', category: 'Mini Sumo Robot', school: '—', birth: '15.04.2024', parent: 'Rehman', team: '—', status: 'Bekliyor', createdAt: new Date('2026-04-15T22:16:00').toISOString() },
                { id: uid(), name: 'Melik Halis Ateş', email: 'melikhalisates5@gmail.com', category: 'Robo Futbol', school: 'Borsa İstanbul Mehmet Akif Ersoy Mesleki Teknik Anadolu Lisesi', birth: '14.08.2010', parent: '03869044890', team: 'RoboForce FC', status: 'Bekliyor', createdAt: new Date('2026-04-15T16:50:00').toISOString() },
                { id: uid(), name: 'Kerem Sanca', email: 'sarcabaris811@gmail.com', category: 'Serbest', school: 'Özel Adem Ceylan Final Mesleki ve Teknik Anadolu Lisesi', birth: '04.12.2009', parent: '05421831976', team: 'Morduk-41', status: 'Onaylandı', createdAt: new Date('2026-04-14T10:00:00').toISOString() },
            ];
            apps = sampleData;
            saveApps(apps);
            renderDashboard();
        }

        // Load sync email
        const savedSyncEmail = localStorage.getItem('bs_sync_email');
        if (savedSyncEmail) {
          const el = document.getElementById('syncEmail');
          if (el) el.value = savedSyncEmail;
        }
        const lastSync = localStorage.getItem('bs_last_sync');
        if (lastSync) {
          const el = document.getElementById('syncLastTime');
          if (el) el.textContent = 'Son senkronizasyon: ' + new Date(lastSync).toLocaleString('tr-TR');
        }
        if (window.robotekAPI && window.robotekAPI.getSyncSettings) {
          window.robotekAPI.getSyncSettings().then(function(settings) {
            var pathEl = document.getElementById('syncDownloadPath');
            if (pathEl && settings && settings.downloadFolder) {
              pathEl.value = settings.downloadFolder;
            }
          });
        }
    
        // ─── RAPOR DURUMU FONKSİYONLARI ──────────────────────────────
        function loadCatReportStatus() {
            try { return JSON.parse(localStorage.getItem('bs_cat_report') || '{}'); } catch(e) { return {}; }
        }
        function saveCatReportStatus(obj) {
            localStorage.setItem('bs_cat_report', JSON.stringify(obj));
        }
        function toggleCatReport(cat) {
            var s = loadCatReportStatus();
            var key = 'CAT__' + cat;
            if (s[key]) delete s[key]; else s[key] = true;
            saveCatReportStatus(s);
            renderReportStatusGrid();
        }
        function renderReportStatusGrid() {
            var grid = document.getElementById('reportStatusGrid');
            if (!grid) return;
            var s = loadCatReportStatus();
            var html2 = '';
            for (var i = 0; i < CATEGORIES.length; i++) {
                var cat = CATEGORIES[i];
                var key = 'CAT__' + cat;
                var uploaded = !!s[key];
                var cnt = apps.filter(function(a) { return a.category === cat; }).length;
                html2 += '<div style="display:flex;align-items:center;gap:12px;padding:11px 16px;border-bottom:1px solid var(--border)">';
                html2 += '<div style="flex:1">';
                html2 += '<div style="font-size:13px;font-weight:600">' + cat + '</div>';
                html2 += '<div style="font-size:11px;color:var(--muted);margin-top:2px">' + cnt + ' başvuru</div>';
                html2 += '</div>';
                html2 += '<button ';
                html2 += 'onclick="toggleCatReport(' + JSON.stringify(cat) + ')" ';
                if (uploaded) {
                    html2 += 'style="padding:6px 14px;border:none;border-radius:8px;cursor:pointer;font-size:12px;font-weight:600;background:rgba(63,185,80,.15);color:#3fb950">';
                    html2 += '✅ Yüklendi';
                } else {
                    html2 += 'style="padding:6px 14px;border:none;border-radius:8px;cursor:pointer;font-size:12px;font-weight:600;background:rgba(240,136,62,.15);color:#f0883e">';
                    html2 += '⏳ Yüklenmedi';
                }
                html2 += '</button>';
                html2 += '</div>';
            }
            grid.innerHTML = html2;
        }
        function buildReportText() {
            var s = loadCatReportStatus();
            var now = new Date().toLocaleString('tr-TR');
            var lines2 = [];
            lines2.push('═══════════════════════════════════════════════════');
            lines2.push('     BAŞVURU SİSTEMİ — DETAYLI TAKIM RAPOR RAPORU');
            lines2.push('  Tarih: ' + now);
            lines2.push('═══════════════════════════════════════════════════');
            lines2.push('');

            for (var i = 0; i < CATEGORIES.length; i++) {
                var cat = CATEGORIES[i];
                var key = 'CAT__' + cat;
                var catUploaded = !!s[key];
                var catApps = apps.filter(function(a) { return a.category === cat; });

                lines2.push('📂 ' + cat.toUpperCase());
                lines2.push('   Kategori Durumu: ' + (catUploaded ? '✅ Raporlar Yüklendi' : '⏳ Raporlar Tamamlanmadı'));
                lines2.push('   Toplam Takım: ' + catApps.length);
                lines2.push('  ─────────────────────────────────────────────────');

                if (catApps.length === 0) {
                    lines2.push('     (Bu kategoride henüz başvuru yok)');
                } else {
                    for (var j = 0; j < catApps.length; j++) {
                        var a = catApps[j];
                        var teamName = a.team ? a.team + ' (' + a.name + ')' : a.name;
                        var reportStatus = a.hasFiles ? 'RAPOR YÜKLENDİ' : 'RAPOR YÜKLENMEDİ';
                        var reportIcon = a.hasFiles ? '✅' : '❌';
                        lines2.push('     ' + (j + 1) + '. ' + teamName + ' — ' + reportIcon + ' ' + reportStatus);
                    }
                }
                lines2.push('');
            }

            lines2.push('═══════════════════════════════════════════════════');
            lines2.push('ÖZET:');
            var totalUploaded = CATEGORIES.filter(function(c) { return s['CAT__' + c]; }).length;
            lines2.push('  Kategori Rapor Tamamlanma: ' + totalUploaded + '/' + CATEGORIES.length);
            var totalWithReport = apps.filter(function(a) { return a.hasFiles; }).length;
            lines2.push('  Raporu Yüklenen Toplam Takım: ' + totalWithReport + '/' + apps.length);
            lines2.push('═══════════════════════════════════════════════════');

            return lines2.join('\n');
        }
        function openReportModal() {
            var ov = document.getElementById('reportModalOverlay');
            if (!ov) return;
            ov.style.display = 'flex';
            document.getElementById('reportTextContent').value = buildReportText();
            document.getElementById('reportCopyStatus').textContent = '';
        }
        function closeReportModal() {
            var ov = document.getElementById('reportModalOverlay');
            if (ov) ov.style.display = 'none';
        }
        function copyReportText() {
            var text = buildReportText();
            navigator.clipboard.writeText(text).then(function() {
                document.getElementById('reportCopyStatus').textContent = '✅ Kopyalandı!';
                showToast('Rapor panoya kopyalandı!', 'success');
            }).catch(function() {
                var ta = document.createElement('textarea');
                ta.value = text;
                document.body.appendChild(ta);
                ta.select();
                document.execCommand('copy');
                document.body.removeChild(ta);
                document.getElementById('reportCopyStatus').textContent = '✅ Kopyalandı!';
            });
        }
        // ─── İLK YÜKLEME ─────────────────────────────────────────────