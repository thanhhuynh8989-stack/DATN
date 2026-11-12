// js/main.js
import { loadTemplatesSilently } from './templates.js';
import { initGlobalSortables, createHoiDongBox, renderHoidongFromExcel } from './hoidong-manager.js';
import { exportWordZIP } from './docx-exporter.js';
import { exportExcelReport, exportDSHoiDongExcel, exportPhanBienExcel } from './excel-exporter.js';
import { 
    readExcelFile, parseSinhVienRows, parseGiangVienRows, parseHoiDongRows,
    loadDanhSachNganhSilently, danhMucNganh, findNganhInfoByLop 
} from './excel-reader.js';

document.addEventListener("DOMContentLoaded", async () => {
    try {
        await Promise.all([
            loadTemplatesSilently(),
            loadDanhSachNganhSilently('./sample_files/DANH_SACH_NGANH.xlsx')
        ]);
    } catch (e) {
        console.error("Lỗi tải dữ liệu khởi tạo:", e);
    }

    initKhoaNganhDropdowns();
    initGlobalSortables();
    bindEvents();
});

function bindEvents() {
    // 0. XỬ LÝ CHUYỂN TAB TRÊN ĐIỆN THOẠI ĐẶT DỌC (MOBILE VIEW)
    const tabSV = document.getElementById('tab-btn-sv');
    const tabHD = document.getElementById('tab-btn-hd');
    const tabGV = document.getElementById('tab-btn-gv');

    const colSV = document.getElementById('col-sv');
    const colHD = document.getElementById('col-hd');
    const colGV = document.getElementById('col-gv');

    const switchTab = (activeBtn, activeCol) => {
        [tabSV, tabHD, tabGV].forEach(btn => btn?.classList.remove('active'));
        [colSV, colHD, colGV].forEach(col => col?.classList.remove('active'));

        activeBtn?.classList.add('active');
        activeCol?.classList.add('active');
    };

    tabSV?.addEventListener('click', () => switchTab(tabSV, colSV));
    tabHD?.addEventListener('click', () => switchTab(tabHD, colHD));
    tabGV?.addEventListener('click', () => switchTab(tabGV, colGV));

    // 1. SỰ KIỆN NẠP DỮ LIỆU EXCEL
    const btnReadFiles = document.getElementById('btn-read-files');
    btnReadFiles?.addEventListener('click', async () => {
        const fileSV = document.getElementById('upload-sv')?.files[0];
        const fileGV = document.getElementById('upload-gv')?.files[0];
        const fileHD = document.getElementById('upload-hd')?.files[0];

        if (!fileSV && !fileGV && !fileHD) {
            return alert("Vui lòng chọn ít nhất 1 file Excel để nạp dữ liệu!");
        }

        try {
            if (fileSV) {
                const rows = await readExcelFile(fileSV);
                const svList = parseSinhVienRows(rows);
                renderSinhVienPool(svList);

                // Tự động chọn đúng Khoa & Ngành trên thanh cài đặt dựa vào mã lớp SV
                const firstSVWithLop = svList.find(s => s.lop);
                if (firstSVWithLop && typeof window.syncKhoaNganhFromLop === 'function') {
                    window.syncKhoaNganhFromLop(firstSVWithLop.lop);
                }
            }
            if (fileGV) {
                const rows = await readExcelFile(fileGV);
                const gvList = parseGiangVienRows(rows);
                renderGiangVienPool(gvList);
            }
            if (fileHD) {
                const rows = await readExcelFile(fileHD);
                const parsedHoidongs = parseHoiDongRows(rows);
                renderHoidongFromExcel(parsedHoidongs);
            }
            alert("✅ Nạp dữ liệu Excel thành công!");
        } catch (err) {
            console.error("Lỗi đọc Excel:", err);
            alert("Lỗi khi đọc file Excel: " + err.message);
        }
    });

    // 2. SỰ KIỆN TẠO HỘI ĐỒNG MỚI
    const btnAddHD = document.getElementById('btn-add-hd');
    btnAddHD?.addEventListener('click', () => {
        createHoiDongBox();
    });

    // 3. SỰ KIỆN MỞ MODAL NHẬP ĐIỂM TRỰC TIẾP
    const btnNhapDiem = document.getElementById('btn-nhap-diem');
    btnNhapDiem?.addEventListener('click', () => {
        const meta = getSetupMetadata();
        const exportData = collectExportData(meta);

        if (exportData.length === 0) {
            return alert("Vui lòng phân công sinh viên vào Hội đồng trước khi nhập điểm!");
        }

        const tbody = document.getElementById('tbody-nhap-diem');
        if (!tbody) return;
        tbody.innerHTML = '';

        exportData.forEach((sv, idx) => {
            const tr = `
                <tr>
                    <td>${idx + 1}</td>
                    <td class="fw-bold">${sv.ma_sv}</td>
                    <td class="text-start">${sv.ho_ten}</td>
                    <td><span class="badge bg-info text-dark">${sv.so_hd}</span></td>
                    <td><input type="number" step="0.1" min="0" max="10" class="form-control form-control-sm text-center input-diem-gvhd" data-masv="${sv.ma_sv}" value="${sv.diem_gvhd || ''}" placeholder="0.0"></td>
                    <td><input type="number" step="0.1" min="0" max="10" class="form-control form-control-sm text-center input-diem-ct" data-masv="${sv.ma_sv}" value="${sv.diem_ct || ''}" placeholder="0.0"></td>
                    <td><input type="number" step="0.1" min="0" max="10" class="form-control form-control-sm text-center input-diem-pb" data-masv="${sv.ma_sv}" value="${sv.diem_pb || ''}" placeholder="0.0"></td>
                    <td><input type="number" step="0.1" min="0" max="10" class="form-control form-control-sm text-center input-diem-tk" data-masv="${sv.ma_sv}" value="${sv.diem_tk || ''}" placeholder="0.0"></td>
                </tr>
            `;
            tbody.insertAdjacentHTML('beforeend', tr);
        });

        const modalDiem = new bootstrap.Modal(document.getElementById('modalNhapDiem'));
        modalDiem.show();
    });

    // 4. SỰ KIỆN NÚT "LƯU ĐIỂM" TRONG MODAL
    const btnSaveDiem = document.getElementById('btn-save-diem');
    btnSaveDiem?.addEventListener('click', () => {
        document.querySelectorAll('#tbody-nhap-diem tr').forEach(tr => {
            const maSV = tr.querySelector('.input-diem-gvhd')?.getAttribute('data-masv');
            if (!maSV) return;

            const diemGVHD = tr.querySelector('.input-diem-gvhd')?.value || '';
            const diemCT = tr.querySelector('.input-diem-ct')?.value || '';
            const diemPB = tr.querySelector('.input-diem-pb')?.value || '';
            const diemTK = tr.querySelector('.input-diem-tk')?.value || '';

            // Lưu điểm vào thuộc tính HTML của thẻ Sinh viên trong DOM
            const svCard = document.querySelector(`.dropzone-sv .card-item[data-id="${maSV}"]`);
            if (svCard) {
                svCard.setAttribute('data-diem-gvhd', diemGVHD);
                svCard.setAttribute('data-diem-ct', diemCT);
                svCard.setAttribute('data-diem-pb', diemPB);
                svCard.setAttribute('data-diem-tk', diemTK);
            }
        });

        // Ẩn Modal sau khi lưu
        const modalEl = document.getElementById('modalNhapDiem');
        const modalInstance = bootstrap.Modal.getInstance(modalEl);
        modalInstance?.hide();

        alert("✅ Đã lưu điểm trực tiếp thành công!");
    });

    // 5. SỰ KIỆN XUẤT FILE WORD & 3 FILE EXCEL (.ZIP)
    // 5A. Xuất Gói Hồ sơ Word (.ZIP)
    document.getElementById('btn-export-word')?.addEventListener('click', async () => {
        const meta = getSetupMetadata();
        const exportData = collectExportData(meta);
        if (exportData.length === 0) return alert("Chưa có sinh viên nào được phân công vào Hội đồng!");

        try {
            await exportWordZIP(exportData, meta);
            alert("🎉 Đã xuất thành công gói Hồ sơ Word (.ZIP)!");
        } catch (err) {
            console.error("Lỗi xuất Word:", err);
            alert("Có lỗi xuất Hồ sơ Word: " + err.message);
        }
    });

    // 5B. Xuất Bảng điểm Kết quả chấm (.XLSX)
    document.getElementById('btn-export-diem')?.addEventListener('click', async () => {
        const meta = getSetupMetadata();
        const exportData = collectExportData(meta);
        if (exportData.length === 0) return alert("Chưa có sinh viên nào được phân công vào Hội đồng!");

        try {
            await exportExcelReport(exportData, meta);
        } catch (err) {
            console.error("Lỗi xuất Bảng điểm:", err);
            alert("Có lỗi xuất Bảng điểm: " + err.message);
        }
    });

    // 5C. Xuất Danh sách Hội đồng đã lập (.XLSX)
    document.getElementById('btn-export-hd')?.addEventListener('click', async () => {
        const meta = getSetupMetadata();
        const exportData = collectExportData(meta);
        if (exportData.length === 0) return alert("Chưa có sinh viên nào được phân công vào Hội đồng!");

        try {
            await exportDSHoiDongExcel(exportData, meta);
        } catch (err) {
            console.error("Lỗi xuất DS Hội đồng:", err);
            alert("Có lỗi xuất Danh sách Hội đồng: " + err.message);
        }
    });

    // 5D. Xuất Danh sách Phân công Phản biện (.XLSX)
    document.getElementById('btn-export-pb')?.addEventListener('click', async () => {
        const meta = getSetupMetadata();
        const exportData = collectExportData(meta);
        if (exportData.length === 0) return alert("Chưa có sinh viên nào được phân công vào Hội đồng!");

        try {
            await exportPhanBienExcel(exportData, meta);
        } catch (err) {
            console.error("Lỗi xuất DS Phản biện:", err);
            alert("Có lỗi xuất Danh sách Phản biện: " + err.message);
        }
    });
}

function initKhoaNganhDropdowns() {
    const khoaSelect = document.getElementById('khoa-select');
    const nganhSelect = document.getElementById('nganh-select');
    if (!khoaSelect || !nganhSelect) return;

    if (danhMucNganh.length === 0) {
        khoaSelect.innerHTML = `<option value="Khoa Công nghệ Thực phẩm">Khoa Công nghệ Thực phẩm</option>`;
        nganhSelect.innerHTML = `<option value="Công nghệ Thực phẩm">Công nghệ Thực phẩm</option>`;
        return;
    }

    // Lấy danh sách Khoa duy nhất
    const uniqueKhoas = [...new Set(danhMucNganh.map(item => item.khoa))];
    khoaSelect.innerHTML = uniqueKhoas
        .map(k => `<option value="${k}">${k}</option>`)
        .join('');

    // Hàm cập nhật danh sách Ngành theo Khoa đang chọn
    const updateNganhOptions = (selectedKhoa, preselectNganh = '') => {
        const filteredNganhs = danhMucNganh.filter(item => item.khoa === selectedKhoa);
        nganhSelect.innerHTML = filteredNganhs
            .map(n => `<option value="${n.chuyenNganh}">${n.chuyenNganh} (${n.maChuyenNganh})</option>`)
            .join('');

        if (preselectNganh) {
            const matchOpt = Array.from(nganhSelect.options).find(
                opt => opt.value.toLowerCase() === preselectNganh.toLowerCase()
            );
            if (matchOpt) nganhSelect.value = matchOpt.value;
        }
    };

    // Khởi tạo danh sách Ngành cho Khoa đầu tiên
    updateNganhOptions(khoaSelect.value);

    // Lắng nghe sự kiện đổi Khoa -> tự động lọc lại danh sách Ngành
    khoaSelect.addEventListener('change', () => {
        updateNganhOptions(khoaSelect.value);
    });

    // Lưu helper lên window để dùng khi nạp file Excel SV tự động nhảy đúng Khoa/Ngành
    window.syncKhoaNganhFromLop = (lopStr) => {
        const info = findNganhInfoByLop(lopStr);
        if (info) {
            khoaSelect.value = info.khoa;
            updateNganhOptions(info.khoa, info.chuyenNganh);
        }
    };
}

function getSetupMetadata() {
    return {
        khoa: (document.getElementById('khoa-select')?.value || '').trim(),
        nganh: (document.getElementById('nganh-select')?.value || '').trim(),
        dotThang: (document.getElementById('dot-thang')?.value || '').trim(),
        namCham: (document.getElementById('nam-cham')?.value || '').trim(),
        soQD: (document.getElementById('so-qd')?.value || '').trim(),
        ngayKyQD: (document.getElementById('ngay-ky-qd')?.value || '').trim()
    };
}

function renderSinhVienPool(svList) {
    const pool = document.getElementById('pool-sv');
    if (!pool) return;
    pool.innerHTML = '';
    svList.forEach(sv => {
        const html = `
            <div class="card-item mb-2 p-2 border rounded bg-white shadow-sm" 
                 data-id="${sv.maSV}" data-name="${sv.hoTen}" data-lop="${sv.lop}" data-detai="${sv.deTai}" data-gvhd="${sv.gvhd}"> data-nganh="${sv.nganh || ''}"
                <div class="fw-bold text-primary">${sv.hoTen} (${sv.maSV})</div>
                <div class="text-muted small">Lớp: ${sv.lop} | GVHD: <span class="fw-bold text-dark">${sv.gvhd}</span></div>
                <div class="text-dark small text-truncate" title="${sv.deTai}">Đề tài: ${sv.deTai}</div>
            </div>`;
        pool.insertAdjacentHTML('beforeend', html);
    });
}

function renderGiangVienPool(gvList) {
    const pool = document.getElementById('pool-gv');
    if (!pool) return;
    pool.innerHTML = '';
    gvList.forEach(gv => {
        const html = `
            <div class="card-item bg-light border-info mb-2 p-2 rounded shadow-sm" data-id="${gv.maGV}" data-name="${gv.hoTenGV}">
                <div class="fw-bold text-dark">${gv.hoTenGV}</div>
                <div class="text-muted small">${gv.donVi} - Mã: ${gv.maGV}</div>
            </div>`;
        pool.insertAdjacentHTML('beforeend', html);
    });
}

function collectExportData(meta) {
    const exportData = [];
    document.querySelectorAll('.box-hoidong').forEach(box => {
        const hdName = box.querySelector('.hd-title')?.value || 'Hội Đồng';
        const dateStr = box.querySelector('.hd-date')?.value || '';
        const time = box.querySelector('.hd-time')?.value || '';
        const room = box.querySelector('.hd-room')?.value || '';
        const location = box.querySelector('.hd-location')?.value || '';

        const getGVInfo = (role) => {
            const el = box.querySelector(`.dropzone-gv[data-role="${role}"] .card-item`);
            return {
                name: el ? el.getAttribute('data-name') : '',
                id: el ? el.getAttribute('data-id') : ''
            };
        };

        const ct = getGVInfo('Chủ tịch');
        const pb = getGVInfo('Phản biện');
        const tk = getGVInfo('Thư ký');

        let formattedDate = dateStr;
        if (dateStr.includes('-')) {
            const parts = dateStr.split('-');
            formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
        }

        box.querySelectorAll('.dropzone-sv .card-item').forEach(svEl => {
            const svNganh = (svEl.getAttribute('data-nganh') || '').trim();
            exportData.push({
                khoa: meta.khoa,
                nganh: meta.nganh || svNganh || '',
                so_hd: hdName,
                ngay: formattedDate,
                gio: time,
                phong: room,
                dia_diem: location,
                chu_tich: ct.name, chu_tich_id: ct.id,
                phan_bien: pb.name, phan_bien_id: pb.id,
                thu_ky: tk.name, thu_ky_id: tk.id,
                ma_sv: svEl.getAttribute('data-id'),
                ho_ten: svEl.getAttribute('data-name'),
                de_tai: svEl.getAttribute('data-detai'),
                lop: svEl.getAttribute('data-lop'),
                gvhd: svEl.getAttribute('data-gvhd'),
                diem_gvhd: svEl.getAttribute('data-diem-gvhd') || '',
                diem_ct: svEl.getAttribute('data-diem-ct') || '',
                diem_pb: svEl.getAttribute('data-diem-pb') || '',
                diem_tk: svEl.getAttribute('data-diem-tk') || ''
            });
        });
    });
    return exportData;
}
