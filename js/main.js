// js/main.js
import { loadTemplatesSilently } from './templates.js';
import { readExcelFile, parseSinhVienRows, parseGiangVienRows } from './excel-reader.js';
import { initGlobalSortables, createHoiDongBox, renderHoidongFromExcel } from './hoidong-manager.js';
import { exportWordZIP } from './docx-exporter.js';
import { exportExcelReport } from './excel-exporter.js';

document.addEventListener("DOMContentLoaded", async () => {
    try {
        await loadTemplatesSilently();
    } catch (e) {
        console.error("Lỗi tải template Word:", e);
    }

    initGlobalSortables();
    bindEvents();
});

// Bổ sung vào hàm bindEvents() trong file js/main.js

// 4. Sự kiện Nhập Điểm Trực Tiếp
const btnNhapDiem = document.getElementById('btn-nhap-diem');
btnNhapDiem?.addEventListener('click', () => {
    const khoaSelect = document.getElementById('khoa-select');
    const khoa = khoaSelect ? khoaSelect.value : 'Khoa';
    const exportData = collectExportData(khoa);

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
                <td><input type="number" step="0.1" min="0" max="10" class="form-control form-control-sm text-center input-diem-gvhd" data-masv="${sv.ma_sv}" placeholder="0.0"></td>
                <td><input type="number" step="0.1" min="0" max="10" class="form-control form-control-sm text-center input-diem-ct" data-masv="${sv.ma_sv}" placeholder="0.0"></td>
                <td><input type="number" step="0.1" min="0" max="10" class="form-control form-control-sm text-center input-diem-pb" data-masv="${sv.ma_sv}" placeholder="0.0"></td>
                <td><input type="number" step="0.1" min="0" max="10" class="form-control form-control-sm text-center input-diem-tk" data-masv="${sv.ma_sv}" placeholder="0.0"></td>
            </tr>
        `;
        tbody.insertAdjacentHTML('beforeend', tr);
    });

    // Mở Bootstrap Modal
    const modalDiem = new bootstrap.Modal(document.getElementById('modalNhapDiem'));
    modalDiem.show();
});

function bindEvents() {
    // 1. Sự kiện Nạp dữ liệu Excel
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
            }
            if (fileGV) {
                const rows = await readExcelFile(fileGV);
                const gvList = parseGiangVienRows(rows);
                renderGiangVienPool(gvList);
            }
            if (fileHD) {
                const rows = await readExcelFile(fileHD);
                renderHoidongFromExcel(rows);
            }
            alert("✅ Nạp dữ liệu Excel thành công!");
        } catch (err) {
            console.error("Lỗi đọc Excel:", err);
            alert("Lỗi khi đọc file Excel: " + err.message);
        }
    });

    // 2. Sự kiện Tạo Hội Đồng
    const btnAddHD = document.getElementById('btn-add-hd');
    btnAddHD?.addEventListener('click', () => {
        createHoiDongBox();
    });

    // 3. Sự kiện Xuất File Word & Excel
    const btnExport = document.getElementById('btn-export');
    btnExport?.addEventListener('click', async () => {
        const meta = getSetupMetadata();
        const exportData = collectExportData(meta);

        if (exportData.length === 0) {
            return alert("Chưa có sinh viên nào được phân công vào Hội đồng!");
        }

        try {
            await exportWordZIP(exportData);
            await exportExcelReport(exportData, meta);
            alert("🎉 Xuất Hồ sơ Word (.ZIP) và Bảng điểm Excel thành công!");
        } catch (err) {
            console.error("Lỗi xuất file:", err);
            alert("Có lỗi xuất file: " + err.message);
        }
    });
}

function getSetupMetadata() {
    return {
        khoa: document.getElementById('khoa-select')?.value || 'Khoa Công nghệ Thực phẩm',
        dotThang: document.getElementById('dot-thang')?.value || '6',
        namCham: document.getElementById('nam-cham')?.value || '2026',
        soQD: document.getElementById('so-qd')?.value || '2071/QĐ-ĐHĐT',
        ngayKyQD: document.getElementById('ngay-ky-qd')?.value || '5/8/2026'
    };
}

function renderSinhVienPool(svList) {
    const pool = document.getElementById('pool-sv');
    if (!pool) return;
    pool.innerHTML = '';
    svList.forEach(sv => {
        const html = `
            <div class="card-item mb-2 p-2 border rounded bg-white shadow-sm" 
                 data-id="${sv.maSV}" data-name="${sv.hoTen}" data-lop="${sv.lop}" data-detai="${sv.deTai}" data-gvhd="${sv.gvhd}">
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

        const getGV = (role) => {
            const el = box.querySelector(`.dropzone-gv[data-role="${role}"] .card-item`);
            return el ? el.getAttribute('data-name') : '';
        };

        let formattedDate = dateStr;
        if (dateStr.includes('-')) {
            const parts = dateStr.split('-');
            formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
        }

        box.querySelectorAll('.dropzone-sv .card-item').forEach(svEl => {
            exportData.push({
                khoa: meta.khoa, 
                so_hd: hdName, 
                ngay: formattedDate, 
                gio: time, 
                phong: room, 
                dia_diem: location,
                chu_tich: getGV('Chủ tịch'), 
                phan_bien: getGV('Phản biện'), 
                thu_ky: getGV('Thư ký'),
                ma_sv: svEl.getAttribute('data-id'), 
                ho_ten: svEl.getAttribute('data-name'),
                de_tai: svEl.getAttribute('data-detai'), 
                lop: svEl.getAttribute('data-lop'),
                gvhd: svEl.getAttribute('data-gvhd')
            });
        });
    });
    return exportData;
}
