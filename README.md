<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Nazhwa Soraya - Sistem Absensi Siswa</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {
            --primary: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            --success: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
            --danger: linear-gradient(135deg, #ff6b6b 0%, #ee5a52 100%);
            --warning: linear-gradient(135deg, #feca57 0%, #ff9ff3 100%);
            --info: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            --purple: linear-gradient(135deg, #a855f7 0%, #ec4899 100%);
            --shadow: 0 20px 40px rgba(0,0,0,0.1);
            --border-radius: 20px;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: var(--border-radius);
            box-shadow: var(--shadow);
            overflow: hidden;
        }

        .header {
            background: var(--primary);
            color: white;
            padding: 40px;
            text-align: center;
        }

        .header h1 {
            font-size: clamp(2.2rem, 5vw, 3.8rem);
            font-weight: 800;
            margin-bottom: 10px;
        }

        .header p {
            font-size: 1.3rem;
            opacity: 0.95;
        }

        .content {
            padding: 50px;
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 25px;
            margin-bottom: 40px;
        }

        .stat-card {
            background: white;
            padding: 30px;
            border-radius: 20px;
            text-align: center;
            box-shadow: 0 10px 30px rgba(0,0,0,0.08);
            transition: all 0.3s ease;
            border-top: 4px solid #667eea;
        }

        .stat-card:hover {
            transform: translateY(-5px);
        }

        .stat-icon {
            font-size: 3rem;
            margin-bottom: 15px;
        }

        .stat-number {
            font-size: 3rem;
            font-weight: 800;
            margin-bottom: 5px;
            background: var(--info);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }

        .input-section {
            background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
            padding: 40px;
            border-radius: 20px;
            margin-bottom: 40px;
            border: 1px solid rgba(102, 126, 234, 0.1);
        }

        .section-title {
            font-size: 1.8rem;
            color: #1e293b;
            margin-bottom: 30px;
            display: flex;
            align-items: center;
            gap: 15px;
        }

        .input-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 25px;
            margin-bottom: 30px;
        }

        .form-group label {
            font-weight: 600;
            color: #374151;
            margin-bottom: 10px;
            font-size: 0.95rem;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .form-control {
            padding: 16px 20px;
            border: 2px solid #e2e8f0;
            border-radius: 15px;
            font-size: 1rem;
            transition: all 0.3s ease;
            background: white;
        }

        .form-control:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 4px rgba(102, 126, 234, 0.1);
        }

        .keterangan-group {
            grid-column: 1 / -1;
        }

        .keterangan-group textarea {
            min-height: 100px;
            resize: vertical;
        }

        .btn {
            padding: 18px 30px;
            border: none;
            border-radius: 15px;
            font-size: 1.1rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            display: inline-flex;
            align-items: center;
            gap: 12px;
            text-decoration: none;
        }

        .btn-primary {
            background: var(--primary);
            color: white;
        }

        .btn-success {
            background: var(--success);
            color: white;
        }

        .btn-purple {
            background: var(--purple);
            color: white;
        }

        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 25px rgba(0,0,0,0.15);
        }

        .export-buttons {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 40px;
        }

        @media (max-width: 768px) {
            .export-buttons {
                grid-template-columns: 1fr;
            }
        }

        .table-container {
            background: white;
            border-radius: 20px;
            box-shadow: 0 15px 35px rgba(0,0,0,0.08);
            overflow: hidden;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }

        th {
            background: var(--primary);
            color: white;
            padding: 20px;
            font-weight: 600;
            text-align: left;
        }

        td {
            padding: 20px;
            border-bottom: 1px solid #f1f5f9;
            vertical-align: middle;
        }

        tr:hover {
            background: #f8fafc;
        }

        .status-badge {
            padding: 8px 16px;
            border-radius: 25px;
            font-weight: 600;
            font-size: 0.85rem;
        }

        .status-hadir { background: #dcfce7; color: #166534; }
        .status-alpha { background: #fee2e2; color: #dc2626; }
        .status-izin { background: #fef3c7; color: #d97706; }
        .status-sakit { background: #dbeafe; color: #1e40af; }

        .btn-sm {
            padding: 10px 16px;
            font-size: 0.9rem;
            margin-right: 10px;
        }

        .btn-delete {
            background: var(--danger);
            color: white;
            padding: 10px 16px;
            border-radius: 10px;
            font-size: 0.9rem;
        }

        .siswa-info {
            background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
            padding: 20px;
            border-radius: 15px;
            margin-bottom: 20px;
            border-left: 5px solid #0ea5e9;
        }

        .siswa-name {
            font-size: 1.3rem;
            font-weight: 700;
            color: #0369a1;
            margin-bottom: 5px;
        }

        .siswa-stats {
            display: flex;
            gap: 30px;
            font-size: 0.95rem;
            color: #64748b;
        }

        .empty-state {
            text-align: center;
            padding: 80px 40px;
            color: #64748b;
        }

        .notification {
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 16px 24px;
            border-radius: 12px;
            color: white;
            font-weight: 500;
            z-index: 10000;
            transform: translateX(400px);
            transition: all 0.3s ease;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }

        @media (max-width: 768px) {
            .content { padding: 30px 20px; }
            .input-grid { grid-template-columns: 1fr; }
            .stats-grid { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i class="fas fa-graduation-cap"></i> Nazhwa Soraya</h1>
            <p><strong>Sistem Absensi Siswa</strong> • Real-time • Download Per Siswa & Total</p>
        </div>

        <div class="content">
            <!-- Statistik Real-time -->
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-icon" style="color: #10b981;"><i class="fas fa-check-circle"></i></div>
                    <div class="stat-number" id="totalHadir">0</div>
                    <div>Total Hadir</div>
                </div>
                <div class="stat-card">
                    <div class="stat-icon"><i class="fas fa-users"></i></div>
                    <div class="stat-number" id="totalSiswa">0</div>
                    <div>Total Siswa</div>
                </div>
                <div class="stat-card">
                    <div class="stat-icon" style="color: #f59e0b;"><i class="fas fa-chart-line"></i></div>
                    <div class="stat-number" id="persentaseHadir">0%</div>
                    <div>Persentase</div>
                </div>
                <div class="stat-card">
                    <div class="stat-icon" style="color: #8b5cf6;"><i class="fas fa-file-excel"></i></div>
                    <div class="stat-number" id="totalSiswaUnik">0</div>
                    <div>Siswa Unik</div>
                </div>
            </div>

            <!-- Form Input -->
            <div class="input-section">
                <h2 class="section-title">
                    <i class="fas fa-plus-circle"></i> Input Absensi Baru
                </h2>
                <div class="input-grid">
                    <div class="form-group">
                        <label><i class="fas fa-user"></i> Nama Siswa *</label>
                        <input type="text" id="namaSiswa" class="form-control" placeholder="Nama lengkap siswa" required>
                    </div>
                    <div class="form-group">
                        <label><i class="fas fa-clipboard-check"></i> Status *</label>
                        <select id="status" class="form-control" required>
                            <option value="">Pilih Status</option>
                            <option value="Hadir">✅ Hadir</option>
                            <option value="Alpha">❌ Alpha</option>
                            <option value="Izin">📄 Izin</option>
                            <option value="Sakit">🤒 Sakit</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label><i class="fas fa-calendar"></i> Tanggal *</label>
                        <input type="date" id="tanggal" class="form-control" required>
                    </div>
                    <div class="form-group keterangan-group">
                        <label><i class="fas fa-sticky-note"></i> Keterangan</label>
                        <textarea id="keterangan" class="form-control" placeholder="Keterangan tambahan (opsional)"></textarea>
                    </div>
                </div>
                <button class="btn btn-primary" onclick="tambahAbsensi()">
                    <i class="fas fa-save"></i> Simpan Absensi
                </button>
            </div>

            <!-- Tombol Export -->
            <div class="export-buttons">
                <button class="btn btn-success" onclick="exportAllToExcel()">
                    <i class="fas fa-file-excel"></i> Download Semua Data
                </button>
                <button class="btn btn-purple" onclick="exportSelectedSiswa()">
                    <i class="fas fa-user-times"></i> Download Per Siswa
                </button>
            </div>

            <!-- Filter Siswa -->
            <div style="margin-bottom: 20px;">
                <select id="filterSiswa" class="form-control" style="max-width: 300px; display: inline-block;">
                    <option value="">📋 Semua Siswa</option>
                </select>
                <button class="btn btn-sm btn-primary" onclick="clearFilter()" style="margin-left: 10px;">
                    <i class="fas fa-redo"></i> Reset
                </button>
            </div>

            <!-- Tabel -->
            <div class="table-container">
                <table id="tabelAbsensi">
                    <thead>
                        <tr>
                            <th>No</th>
                            <th>Nama Siswa</th>
                            <th>Status</th>
                            <th>Tanggal</th>
                            <th>Keterangan</th>
                            <th>Aksi</th>
                        </tr>
                    </thead>
                    <tbody id="tbodyAbsensi"></tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
        let dataAbsensi = JSON.parse(localStorage.getItem('absensi_nazhwa_soraya')) || [];
        let nextId = dataAbsensi.length ? Math.max(...dataAbsensi.map(a => a.id)) + 1 : 1;
        let filteredData = [];

        document.addEventListener('DOMContentLoaded', function() {
            document.getElementById('tanggal').valueAsDate = new Date();
            initApp();
        });

        function initApp() {
            updateSiswaFilter();
            renderTabel();
            updateStats();
        }

        function tambahAbsensi() {
            const nama = document.getElementById('namaSiswa').value.trim();
            const status = document.getElementById('status').value;
            const tanggal = document.getElementById('tanggal').value;
            const keterangan = document.getElementById('keterangan').value.trim();

            if (!nama || !status || !tanggal) {
                showNotification('Lengkapi field wajib (*)!', 'error');
                return;
            }

            const absensi = {
                id: nextId++,
                nama,
                status,
                tanggal,
                keterangan: keterangan || '-',
                timestamp: new Date().toISOString()
            };

            dataAbsensi.unshift(absensi);
            saveData();
            updateSiswaFilter();
            renderTabel();
            updateStats();
            resetForm();
            showNotification(`Absensi ${nama} berhasil disimpan!`, 'success');
        }

        function hapusAbsensi(id) {
            if (confirm('Hapus data absensi ini?')) {
                dataAbsensi = dataAbsensi.filter(item => item.id !== id);
                saveData();
                updateSiswaFilter();
                renderTabel();
                updateStats();
                showNotification('Data dihapus!', 'success');
            }
        }

        function updateSiswaFilter() {
            const siswaUnik = [...new Set(dataAbsensi.map(item => item.nama))];
            const select = document.getElementById('filterSiswa');
            select.innerHTML = '<option value="">📋 Semua Siswa</option>';
            
            siswaUnik.forEach(nama => {
                const option = document.createElement('option');
                option.value = nama;
                option.textContent = nama;
                select.appendChild(option);
            });

            select.onchange = function() {
                filterData(this.value);
            };
        }

        function filterData(namaSiswa) {
            if (!namaSiswa) {
                filteredData = dataAbsensi;
                renderSiswaInfo(null);
            } else {
                filteredData = dataAbsensi.filter(item => item.nama === namaSiswa);
                renderSiswaInfo(namaSiswa);
            }
            renderTabel();
        }

        function renderSiswaInfo(namaSiswa) {
            const container = document.querySelector('.export-buttons');
            if (namaSiswa && filteredData.length > 0) {
                const stats = getSiswaStats(namaSiswa);
                container.insertAdjacentHTML('beforebegin', `
                    <div class="siswa-info">
                        <div class="siswa-name">${namaSiswa}</div>
                        <div class="siswa-stats">
                            <span><i class="fas fa-check-circle"></i> Hadir: ${stats.hadir}</span>
                            <span><i class="fas fa-times-circle"></i> Alpha: ${stats.alpha}</span>
                            <span><i class="fas fa-file-alt"></i> Total: ${stats.total}</span>
                        </div>
                    </div>
                `);
            } else {
                document.querySelector('.siswa-info')?.remove();
            }
        }

        function getSiswaStats(nama) {
            const siswaData = dataAbsensi.filter(item => item.nama === nama);
            return {
                total: siswaData.length,
                hadir: siswaData.filter(item => item.status === 'Hadir').length,
                alpha: siswaData.filter(item => item.status === 'Alpha').length
            };
        }

        function renderTabel() {
            const tbody = document.getElementById('tbodyAbsensi');
            const dataToShow = filteredData.length > 0 ? filteredData : dataAbsensi;
            
            if (dataToShow.length === 0) {
                tbody.innerHTML = `
                    <tr>
                        <td colspan="6" class="empty-state">
                            <i class="fas fa-clipboard-list" style="font-size: 4rem; margin-bottom: 20px;"></i>
                            <h3>Belum ada data absensi</h3>
                            <p>Masukkan data absensi pertama!</p>
                        </td>
                    </tr>
                `;
                return;
            }

            tbody.innerHTML = dataToShow.map((item, index) => `
                <tr>
                    <td><strong>${index + 1}</strong></td>
                    <td style="font-weight: 600;">${item.nama}</td>
                    <td><span class="status-badge status-${item.status.toLowerCase()}">${item.status}</span></td>
                    <td>${formatTanggal(item.tanggal)}</td>
                    <td>${item.keterangan}</td>
                    <td>
                        <button class="btn-sm btn-delete" onclick="hapusAbsensi(${item.id})">
                            <i class="fas fa-trash"></i>
                        </button>
                    </td>
                </tr>
            `).join('');
        }

        function updateStats() {
            const totalSiswa = dataAbsensi.length;
            const totalHadir = dataAbsensi.filter(item => item.status === 'Hadir').length;
            const persentase = totalSiswa > 0 ? Math.round((totalHadir / totalSiswa) * 100) : 0;
            const siswaUnik = [...new Set(dataAbsensi.map(item => item.nama))].length;

            document.getElementById('totalHadir').textContent = totalHadir;
            document.getElementById('totalSiswa').textContent = totalSiswa;
            document.getElementById('persentaseHadir').textContent = persentase + '%';
            document.getElementById('totalSiswaUnik').textContent = siswaUnik;
        }

        function resetForm() {
            document.getElementById('namaSiswa').value = '';
            document.getElementById('status').value = '';
            document.getElementById('keterangan').value = '';
            document.getElementById('tanggal').valueAsDate = new Date();
        }

        function clearFilter() {
            document.getElementById('filterSiswa').value = '';
            filterData('');
        }

        function formatTanggal(tanggal) {
            return new Date(tanggal).toLocaleDateString('id-ID', {
                weekday: 'short',
                year: 'numeric',
                month: 'short',
                day: 'numeric'
            });
        }

        function exportAllToExcel() {
            if (dataAbsensi.length === 0) {
                showNotification('Tidak ada data untuk diexport!', 'error');
                return;
            }
            exportData(dataAbsensi, 'SEMUA_DATA');
        }

        function exportSelectedSiswa() {
            const selectedSiswa = document.getElementById('filterSiswa').value;
            if (!selectedSiswa || filteredData.length === 0) {
                showNotification('Pilih siswa terlebih dahulu!', 'error');
                return;
            }
            exportData(filteredData, selectedSiswa);
        }

        function exportData(dataToExport, namaFile) {
            const wsData = [
                [`ABSENSI NAZHWA SORAYA - ${namaFile}`],
                [],
                ['No', 'Nama Siswa', 'Status', 'Tanggal', 'Keterangan']
            ];
            
            dataToExport.forEach((item, index) => {
                wsData.push([index + 1, item.nama, item.status, formatTanggal(item.tanggal), item.keterangan]);
            });

            const ws = XLSX.utils.aoa_to_sheet(wsData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Absensi");
            
            const filename = `Absensi_Nazhwa_Soraya_${namaFile}_${new Date().toISOString().split('T')[0]}.xlsx`;
            XLSX.writeFile(wb, filename);
            
            showNotification(`✅ Excel "${filename}" berhasil diunduh!`, 'success');
        }

        function saveData() {
            localStorage.setItem('absensi_nazhwa_soraya', JSON.stringify(dataAbsensi));
        }

        function showNotification(message, type = 'info') {
            const notification = document.createElement('div');
            notification.className = 'notification';
            notification.style.background = type === 'success' ? '#10b981' : '#ef4444';
            notification.textContent = message;
            document.body.appendChild(notification);
            
            setTimeout(() => notification.style.transform = 'translateX(0)', 100);
            setTimeout(() => {
                notification.style.transform = 'translateX(400px)';
                setTimeout(() => document.body.removeChild(notification), 300);
            }, 3000);
        }
    </script>
</body>
</html>
