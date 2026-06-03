
      // Load embedded data
      function loadDataFromFile() {
        if (window.EMBEDDED_DATA && window.EMBEDDED_DATA.records) {
          initializeTable(window.EMBEDDED_DATA.records);
          alert('Dữ liệu đã được tải lại từ hệ thống');
        } else {
          alert('Không tìm thấy dữ liệu');
        }
      }

      // Initialize table
      function initializeTable(records) {
        const tableBody = document.getElementById('tableBody');
        if (!tableBody) return;

        tableBody.innerHTML = '';

        if (!records || records.length === 0) {
          tableBody.innerHTML = '<tr><td colspan="12" style="text-align:center;padding:20px;">Không có dữ liệu</td></tr>';
          return;
        }

        records.forEach((record, index) => {
          const row = document.createElement('tr');
          row.style.cursor = 'pointer';
          row.onclick = () => showModal(record._record_id);
          row.innerHTML = `
            <td>${record.MaChungKhoan || record.code || 'N/A'}</td>
            <td>${record.LoaiQuyen || 'N/A'}</td>
            <td>${record.ThiTruong || 'N/A'}</td>
            <td>${record.SanGD || 'N/A'}</td>
            <td>${record.NoiQuanLy || 'N/A'}</td>
            <td>${record.NgayChot || 'N/A'}</td>
            <td>${record.NgayThucHien || 'N/A'}</td>
            <td>${record.NgayThanhToan || 'N/A'}</td>
            <td>${record.NgayHetHanCNQuyenMua || 'N/A'}</td>
            <td>${record.NgayHetHanDKQuyenMua || 'N/A'}</td>
            <td><div style="max-height: 100px; overflow-y: auto; font-size: 13px;">${record.NoiDung || 'N/A'}</div></td>
            <td><span class="status-badge status-${record.status || 'pending'}">${record.status || 'pending'}</span></td>
          `;
          tableBody.appendChild(row);
        });

        updateStats(records);
      }

      // Update statistics
      function updateStats(records) {
        const total = records.length;
        const pending = records.filter(r => r.status === 'pending').length;
        const confirmed = records.filter(r => r.status === 'confirmed').length;
        const rejected = records.filter(r => r.status === 'rejected').length;

        document.getElementById('totalRecords').textContent = total;
        document.getElementById('pendingCount').textContent = pending;
        document.getElementById('confirmedCount').textContent = confirmed;
        document.getElementById('rejectedCount').textContent = rejected;
      }

      // Show record modal
      function showModal(recordId) {
        const records = window.EMBEDDED_DATA?.records || [];
        const record = records.find(r => r._record_id === recordId);

        if (!record) {
          alert('Không tìm thấy record: ' + recordId);
          return;
        }

        // Update modal with record data
        document.getElementById('detailCode').textContent = record.code || 'N/A';
        document.getElementById('detailName').textContent = record.tên_chứng_khoán || 'N/A';
        document.getElementById('detailOrg').textContent = record.tên_tổ_chức_đăng_ký || 'N/A';
        document.getElementById('detailISIN').textContent = record.mã_isin || 'N/A';
        document.getElementById('detailExchange').textContent = record.nơi_giao_dịch || 'N/A';
        document.getElementById('detailType').textContent = record.loại_chứng_khoán || 'N/A';
        document.getElementById('detailRegDate').textContent = record.ngày_đăng_ký_cuối || 'N/A';
        document.getElementById('detailTime').textContent = record.thời_gian_thực_hiện || 'N/A';
        document.getElementById('detailLocation').textContent = record.địa_điểm_thực_hiện || 'N/A';
        document.getElementById('detailPurpose').textContent = record.lý_do_mục_đích || 'N/A';
        document.getElementById('detailRatio').textContent = record.tỷ_lệ_thực_hiện || 'N/A';

        // Mapped Derived Fields
        document.getElementById('detailLoaiQuyen').textContent = record.LoaiQuyen || 'N/A';
        document.getElementById('detailThiTruong').textContent = record.ThiTruong || 'N/A';
        document.getElementById('detailSanGD').textContent = record.SanGD || 'N/A';
        document.getElementById('detailNoiQuanLy').textContent = record.NoiQuanLy || 'N/A';
        document.getElementById('detailNgayChot').textContent = record.NgayChot || 'N/A';
        document.getElementById('detailNgayThucHien').textContent = record.NgayThucHien || 'N/A';
        document.getElementById('detailNgayThanhToan').textContent = record.NgayThanhToan || 'N/A';
        document.getElementById('detailNgayHetHanCN').textContent = record.NgayHetHanCNQuyenMua || 'N/A';
        document.getElementById('detailNgayHetHanDK').textContent = record.NgayHetHanDKQuyenMua || 'N/A';
        document.getElementById('detailNoiDung').textContent = record.NoiDung || 'N/A';

        // Rights and benefits section
        document.getElementById('detailRightShMeeting').textContent = record.quyền_họp_đại_hội_cổ_đông || 'N/A';
        document.getElementById('detailRightDividendCash').textContent = record.quyền_cổ_tức_tiền || 'N/A';
        document.getElementById('detailRightDividendShare').textContent = record.quyền_cổ_tức_cổ_phiếu || 'N/A';
        document.getElementById('detailRightPurchase').textContent = record.quyền_mua || 'N/A';
        document.getElementById('detailRightSwapConversion').textContent = record.quyền_hoán_đổi_chuyển_đổi || 'N/A';
        document.getElementById('detailWarrant').textContent = record.chứng_quyền || 'N/A';
        document.getElementById('detailRegistrationApproval').textContent = record.chấp_thuận_đăng_ký || 'N/A';
        document.getElementById('detailCancellation').textContent = record.tin_húy || 'N/A';
        document.getElementById('detailChange').textContent = record.thay_đổi || 'N/A';

        // Full text content
        const fullTextSection = document.getElementById('fullTextSection');
        if (record.text_content && fullTextSection) {
          document.getElementById('detailFullText').textContent = record.text_content;
          fullTextSection.style.display = 'block';
        } else if (fullTextSection) {
          fullTextSection.style.display = 'none';
        }

        document.getElementById('detailUrl').href = record.url || '#';
        document.getElementById('detailUrl').textContent = 'Xem trên nguồn gốc';

        // Show modal
        const modal = document.getElementById('detailModal');
        if (modal) {
          modal.classList.add('show');
        }
      }

      // Close modal
      function closeModal() {
        const modal = document.getElementById('detailModal');
        if (modal) {
          modal.classList.remove('show');
        }
      }

      // Confirm/Reject buttons
      function confirmRecord() {
        alert('Record xác nhận (tính năng sắp tới)');
        closeModal();
      }

      function rejectRecord() {
        alert('Record bị từ chối (tính năng sắp tới)');
        closeModal();
      }

      // Initialize table on load
      if (window.EMBEDDED_DATA && window.EMBEDDED_DATA.records) {
        initializeTable(window.EMBEDDED_DATA.records);
      }

      // Filter table by search text and date range
      function filterTable() {
        const searchInput = document.getElementById('searchInput')?.value.toLowerCase() || '';
        const startDate = document.getElementById('startDate')?.value;
        const endDate = document.getElementById('endDate')?.value;
        const loaiQuyenFilter = document.getElementById('loaiQuyenFilter')?.value;
        const sanGDFilter = document.getElementById('sanGDFilter')?.value;
        const tableBody = document.getElementById('tableBody');
        
        if (!tableBody) return;

        const records = window.EMBEDDED_DATA?.records || [];
        let filtered = records;

        // Filter by search text
        if (searchInput) {
          filtered = filtered.filter(r => 
            (r.MaChungKhoan && r.MaChungKhoan.toLowerCase().includes(searchInput)) ||
            (r.code && r.code.toLowerCase().includes(searchInput)) ||
            (r.NoiDung && r.NoiDung.toLowerCase().includes(searchInput)) ||
            (r.tên_chứng_khoán && r.tên_chứng_khoán.toLowerCase().includes(searchInput))
          );
        }

        // Filter by Loại Quyền
        if (loaiQuyenFilter) {
          filtered = filtered.filter(r => r.LoaiQuyen === loaiQuyenFilter);
        }

        // Filter by Sàn GD
        if (sanGDFilter) {
          filtered = filtered.filter(r => r.SanGD === sanGDFilter);
        }

        // Filter by date range (using NgayChot primarily)
        if (startDate || endDate) {
          filtered = filtered.filter(r => {
            const recordDate = r.NgayChot || r.ngày_đăng_ký_cuối || r.date; // format: DD/MM/YYYY
            if (!recordDate) return false;
            
            const [day, month, year] = recordDate.split('/');
            const dateObj = new Date(year, month - 1, day);
            
            const start = startDate ? new Date(startDate) : new Date('1900-01-01');
            const end = endDate ? new Date(endDate) : new Date('2099-12-31');
            
            return dateObj >= start && dateObj <= end;
          });
        }

        // Render filtered results
        tableBody.innerHTML = '';
        if (filtered.length === 0) {
          tableBody.innerHTML = '<tr><td colspan="12" style="text-align:center;padding:20px;">Không tìm thấy kết quả</td></tr>';
          return;
        }

        filtered.forEach(record => {
          const row = document.createElement('tr');
          row.style.cursor = 'pointer';
          row.onclick = () => showModal(record._record_id);
          row.innerHTML = `
            <td>${record.MaChungKhoan || record.code || 'N/A'}</td>
            <td>${record.LoaiQuyen || 'N/A'}</td>
            <td>${record.ThiTruong || 'N/A'}</td>
            <td>${record.SanGD || 'N/A'}</td>
            <td>${record.NoiQuanLy || 'N/A'}</td>
            <td>${record.NgayChot || 'N/A'}</td>
            <td>${record.NgayThucHien || 'N/A'}</td>
            <td>${record.NgayThanhToan || 'N/A'}</td>
            <td>${record.NgayHetHanCNQuyenMua || 'N/A'}</td>
            <td>${record.NgayHetHanDKQuyenMua || 'N/A'}</td>
            <td><div style="max-height: 100px; overflow-y: auto; font-size: 13px;">${record.NoiDung || 'N/A'}</div></td>
            <td><span class="status-badge status-${record.status || 'pending'}">${record.status || 'pending'}</span></td>
          `;
          tableBody.appendChild(row);
        });
      }

      // Filter by status
      let currentStatus = 'all';
      function filterByStatus(status) {
        currentStatus = status;
        
        // Update button styles
        document.querySelectorAll('.filter-buttons .btn').forEach(btn => {
          btn.classList.remove('active');
        });
        event.target.classList.add('active');

        // Filter table
        const tableBody = document.getElementById('tableBody');
        if (!tableBody) return;

        const records = window.EMBEDDED_DATA?.records || [];
        let filtered = records;

        if (status !== 'all') {
          filtered = filtered.filter(r => r.status === status);
        }

        tableBody.innerHTML = '';
        if (filtered.length === 0) {
          tableBody.innerHTML = '<tr><td colspan="12" style="text-align:center;padding:20px;">Không có records</td></tr>';
          return;
        }

        filtered.forEach(record => {
          const row = document.createElement('tr');
          row.style.cursor = 'pointer';
          row.onclick = () => showModal(record._record_id);
          row.innerHTML = `
            <td>${record.MaChungKhoan || record.code || 'N/A'}</td>
            <td>${record.LoaiQuyen || 'N/A'}</td>
            <td>${record.ThiTruong || 'N/A'}</td>
            <td>${record.SanGD || 'N/A'}</td>
            <td>${record.NoiQuanLy || 'N/A'}</td>
            <td>${record.NgayChot || 'N/A'}</td>
            <td>${record.NgayThucHien || 'N/A'}</td>
            <td>${record.NgayThanhToan || 'N/A'}</td>
            <td>${record.NgayHetHanCNQuyenMua || 'N/A'}</td>
            <td>${record.NgayHetHanDKQuyenMua || 'N/A'}</td>
            <td><div style="max-height: 100px; overflow-y: auto; font-size: 13px;">${record.NoiDung || 'N/A'}</div></td>
            <td><span class="status-badge status-${record.status || 'pending'}">${record.status || 'pending'}</span></td>
          `;
          tableBody.appendChild(row);
        });
      }

      // Clear date filter
      function clearDateFilter() {
        document.getElementById('startDate').value = '';
        document.getElementById('endDate').value = '';
        filterTable();
      }
    