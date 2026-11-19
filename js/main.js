// js/main.js
import { initGlobalSortables, createHoiDongBox, renderHoidongFromExcel } from './hoidong-manager.js';
import { exportExcelReport, exportDSHoiDongExcel, exportPhanBienExcel } from './excel-exporter.js';
import { 
    readExcelFile, parseSinhVienRows, parseGiangVienRows, parseHoiDongRows,
    loadDanhSachNganhSilently, danhMucNganh, findNganhInfoByLop 
} from './excel-reader.js';
import { 
    loadTemplatesSilently, templates,
    updateScoringTemplate, resetAllScoringTemplates 
} from './templates.js';
import { 
    getAdminConfig, saveAdminConfig, resetAdminConfigText, 
    saveAdminBinaryFile, deleteAdminBinaryFile 
} from './admin-config.js';
import { exportWordZIP } from './docx-exporter.js';
import { detectDuplicateTeacherNameDifferentId } from './validators.js';
import { 
    getGasWebAppUrl, setGasWebAppUrl, getSavedSessionMeta, saveSessionMeta,
    syncSessionToCloud, loadSessionFromCloud, restoreSessionFromPayload 
} from './session-sync.js';

window.studentScores = window.studentScores || {};

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

// Hàm áp dụng cấu hình Admin lên giao diện chính ngay lập tức (đuôi QĐ + dropdown Cơ sở)
function applyAdminConfigToUI() {
    const cfg = getAdminConfig();
    const qdSuffixEl = document.getElementById('so-qd')?.nextElementSibling;
    if (qdSuffixEl) qdSuffixEl.textContent = cfg.qdSuffix;

    document.querySelectorAll('.hd-location').forEach(sel => {
        const currVal = sel.value;
        sel.innerHTML = [
            `<option value="">-- Chọn cơ sở --</option>`,
            ...cfg.campuses.map(c => `<option value="${c}" ${c === currVal ? 'selected' : ''}>${c}</option>`)
        ].join('');
    });
}

// Hàm hiển thị Modal báo cáo rà soát rộng rãi, có thanh cuộn (trả về Promise<boolean>: true = Tiếp tục, false = Hủy)
function showExcelReviewModal({ title, requireConfirm = false, htmlSections = [] }) {
    return new Promise((resolve) => {
        const modalEl = document.getElementById('modalExcelReviewReport');
        if (!modalEl || !window.bootstrap) {
            alert(htmlSections.map(s => s.replace(/<[^>]+>/g, ' ')).join('\n\n'));
            return resolve(true);
        }

        const titleEl = document.getElementById('review-modal-title');
        const bodyEl = document.getElementById('review-modal-body');
        const btnCancel = document.getElementById('btn-review-cancel');
        const btnConfirm = document.getElementById('btn-review-confirm');
        const btnCloseX = document.getElementById('btn-review-close-x');

        if (titleEl) titleEl.textContent = title || '📋 KẾT QUẢ RÀ SOÁT & NẠP DỮ LIỆU EXCEL';
        if (bodyEl) {
            bodyEl.innerHTML = htmlSections.join('');
            bodyEl.scrollTop = 0;
        }

        if (requireConfirm) {
            btnCancel?.classList.remove('d-none');
            btnCloseX?.classList.add('d-none');
            if (btnConfirm) {
                btnConfirm.textContent = '✅ Xác nhận file đã đúng, tiếp tục nạp';
                btnConfirm.className = 'btn btn-success btn-sm fw-bold px-3';
            }
        } else {
            btnCancel?.classList.add('d-none');
            btnCloseX?.classList.remove('d-none');
            if (btnConfirm) {
                btnConfirm.textContent = '✅ Đã hiểu';
                btnConfirm.className = 'btn btn-primary btn-sm fw-bold px-3';
            }
        }

        const modalInstance = bootstrap.Modal.getOrCreateInstance(modalEl);
        let settled = false;

        const finish = (result) => {
            if (settled) return;
            settled = true;
            btnConfirm?.removeEventListener('click', onConfirm);
            btnCancel?.removeEventListener('click', onCancel);
            modalEl.removeEventListener('hidden.bs.modal', onHidden);
            modalInstance.hide();
            resolve(result);
        };

        const onConfirm = () => finish(true);
        const onCancel = () => finish(false);
        const onHidden = () => finish(!requireConfirm);

        btnConfirm?.addEventListener('click', onConfirm);
        btnCancel?.addEventListener('click', onCancel);
        modalEl.addEventListener('hidden.bs.modal', onHidden);

        modalInstance.show();
    });
}

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

    // 1. SỰ KIỆN NẠP DỮ LIỆU EXCEL (RÀ SOÁT CHÉO TOÀN BỘ FILE & HIỂN THỊ MODAL LỚN CÓ THANH CUỘN)
    const btnReadFiles = document.getElementById('btn-read-files');
    btnReadFiles?.addEventListener('click', async () => {
        const fileSV = document.getElementById('upload-sv')?.files[0];
        const fileGV = document.getElementById('upload-gv')?.files[0];
        const fileHD = document.getElementById('upload-hd')?.files[0];

        if (!fileSV && !fileGV && !fileHD) {
            return alert("Vui lòng chọn ít nhất 1 file Excel để nạp dữ liệu!");
        }

        // BẮT BUỘC: Chặn nếu chưa chọn file Danh sách SV & Đề tài VÀ trên giao diện cũng chưa từng nạp SV nào
        const existingSVCount = document.querySelectorAll('#pool-sv .card-item, .dropzone-sv .card-item').length;
        if (!fileSV && existingSVCount === 0) {
            return alert("⚠️ Vui lòng chọn file (1) Nhập Danh sách SV & Đồ án trước khi bấm Nạp dữ liệu Excel!");
        }

        try {
            let gvMapById = {};
            let gvList = [];
            let svList = [];
            let parsedHoidongs = {};

            if (fileGV) {
                const rows = await readExcelFile(fileGV);
                gvList = parseGiangVienRows(rows);
                gvList.forEach(gv => {
                    if (gv.maGV) gvMapById[gv.maGV.toUpperCase()] = gv.hoTenGV;
                });
                renderGiangVienPool(gvList);
            }

            if (fileSV) {
                const rows = await readExcelFile(fileSV);
                svList = parseSinhVienRows(rows);

                svList.forEach(sv => {
                    if (!sv.gvhd && sv.maGVHD && gvMapById[sv.maGVHD.toUpperCase()]) {
                        sv.gvhd = gvMapById[sv.maGVHD.toUpperCase()];
                    }
                });

                renderSinhVienPool(svList);

                const firstSVWithLop = svList.find(s => s.lop);
                if (firstSVWithLop && typeof window.syncKhoaNganhFromLop === 'function') {
                    window.syncKhoaNganhFromLop(firstSVWithLop.lop);
                }
            }

            if (fileHD) {
                const rows = await readExcelFile(fileHD);
                parsedHoidongs = parseHoiDongRows(rows);
            }

            // BƯỚC RÀ SOÁT CHÉO 1: GIẢNG VIÊN TRÙNG TÊN CHUẨN HÓA (BỎ HỌC VỊ, BỎ DẤU, BỎ KHOẢNG TRẮNG) NHƯNG KHÁC MÃ GV
            const dupNameDiffIdList = detectDuplicateTeacherNameDifferentId({ svList, gvList, parsedHoidongs });

            if (dupNameDiffIdList.length > 0 && fileHD) {
                const dupItemsHTML = dupNameDiffIdList.map((item, idx) => {
                    const variantsHTML = item.variants.map(v => `
                        <li class="mb-1">
                            <span class="badge bg-dark font-monospace">Mã GV: ${v.maGV}</span> 
                            <strong>${v.names}</strong>
                            <div class="text-muted small ps-2">↳ Xuất hiện tại: ${v.locations.join(' | ')}</div>
                        </li>
                    `).join('');

                    return `
                        <div class="bg-white border border-warning rounded p-2 mb-2 shadow-sm">
                            <div class="fw-bold text-danger small mb-1">#${idx + 1}. Trùng họ tên giảng viên nhưng khác Mã GV:</div>
                            <ul class="mb-0 ps-3 small">${variantsHTML}</ul>
                        </div>
                    `;
                }).join('');

                const confirmSection = `
                    <div class="alert alert-warning py-2 small mb-2">
                        <strong>⚠️ PHÁT HIỆN GIẢNG VIÊN TRÙNG TÊN NHƯNG KHÁC MÃ GV:</strong><br>
                        Vui lòng kiểm tra các vị trí bên dưới xem đây là <strong>2 giảng viên khác nhau trùng họ tên</strong> hay <strong>nhập sai Mã GV</strong>. Nếu file đã chính xác, bấm <em>"Xác nhận file đã đúng, tiếp tục nạp"</em>; nếu cần sửa lại file Excel, bấm <em>"Hủy"</em>.
                    </div>
                    ${dupItemsHTML}
                `;

                const userConfirmed = await showExcelReviewModal({
                    title: '⚠️ XÁC NHẬN RÀ SOÁT MÃ GIẢNG VIÊN TRÙNG TÊN',
                    requireConfirm: true,
                    htmlSections: [confirmSection]
                });

                if (!userConfirmed) {
                    // Người dùng bấm Hủy -> Dừng không render Hội đồng, vẫn giữ nguyên SV & GV đã nạp ở bước 1 & 2
                    return;
                }
            }

            // BƯỚC 2: RENDER HỘI ĐỒNG & TỔNG HỢP BÁO CÁO (GỘP HĐ, TRÙNG MÃ SV, VI PHẠM GVHD)
            const reportSections = [
                `<div class="alert alert-success py-2 small mb-2 fw-bold">✅ Đã nạp dữ liệu Excel thành công!</div>`
            ];

            if (dupNameDiffIdList.length > 0 && !fileHD) {
                const dupHTML = dupNameDiffIdList.map(item =>
                    `<li>${item.variants.map(v => `<strong>${v.names}</strong> (Mã: <code>${v.maGV}</code> — ${v.locations.join(', ')})`).join(' <strong>vs</strong> ')}</li>`
                ).join('');
                reportSections.push(`
                    <div class="card border-warning mb-2">
                        <div class="card-header bg-warning text-dark fw-bold small py-1">⚠️ Cảnh báo Giảng viên trùng tên nhưng khác Mã GV</div>
                        <div class="card-body p-2 small"><ul class="mb-0 ps-3">${dupHTML}</ul></div>
                    </div>
                `);
            }

            if (fileHD) {
                const { mergeLogs, conflictLogs, duplicateSVLogs } = renderHoidongFromExcel(parsedHoidongs);

                if (conflictLogs.length > 0) {
                    reportSections.push(`
                        <div class="card border-danger mb-2">
                            <div class="card-header bg-danger text-white fw-bold small py-1">⛔ Vi phạm GVHD trùng Thành viên Hội đồng (Đã đẩy về Kho chờ)</div>
                            <div class="card-body p-2 small"><ul class="mb-0 ps-3">${conflictLogs.map(l => `<li class="mb-1">${l.replace(/^•\s*/, '')}</li>`).join('')}</ul></div>
                        </div>
                    `);
                }

                if (duplicateSVLogs.length > 0) {
                    reportSections.push(`
                        <div class="card border-warning mb-2">
                            <div class="card-header bg-warning text-dark fw-bold small py-1">⚠️ Trùng lặp Sinh viên trong file Hội đồng</div>
                            <div class="card-body p-2 small"><ul class="mb-0 ps-3">${duplicateSVLogs.map(l => `<li class="mb-1">${l.replace(/^•\s*/, '')}</li>`).join('')}</ul></div>
                        </div>
                    `);
                }

                if (mergeLogs.length > 0) {
                    reportSections.push(`
                        <div class="card border-info mb-2">
                            <div class="card-header bg-info text-dark fw-bold small py-1">🔄 Tự động gộp & Đánh lại số Hội đồng</div>
                            <div class="card-body p-2 small"><ul class="mb-0 ps-3">${mergeLogs.map(l => `<li class="mb-1">${l.replace(/^•\s*/, '')}</li>`).join('')}</ul></div>
                        </div>
                    `);
                }
            }

            await showExcelReviewModal({
                title: '📋 KẾT QUẢ RÀ SOÁT & NẠP DỮ LIỆU EXCEL',
                requireConfirm: false,
                htmlSections: reportSections
            });

        } catch (err) {
            console.error("Lỗi đọc Excel:", err);
            alert("Lỗi khi đọc file Excel: " + err.message);
        }
    });

    // 1B. SỰ KIỆN TẢI 3 FILE MẪU EXCEL (.XLSX)
    const downloadExcelArray = (aoaData, colWidths, sheetName, fileName) => {
        if (typeof XLSX === 'undefined') {
            return alert("Thư viện XLSX chưa được nạp!");
        }
        const ws = XLSX.utils.aoa_to_sheet(aoaData);
        ws['!cols'] = colWidths.map(w => ({ wch: w }));
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        XLSX.writeFile(wb, fileName);
    };

    document.getElementById('btn-sample-sv')?.addEventListener('click', () => {
        const dataSV = [
            ["STT", "MSSV", "Họ và tên", "Lớp", "Tên đề tài", "GVHD", "Mã GVHD"],
            [1, "27206635615", "Võ Thị Thùy Diễm", "K27 TNM", "Nghiên cứu quy trình sản xuất trà Cascara từ vỏ cà phê", "ThS. Phạm Thị Nga", "230114084"],
            [2, "27216601361", "Trần Nhật Hoàng", "K27 TNM", "Nghiên cứu quy trình sản xuất sữa đậu lăng kết hợp hạt sen", "ThS. Nguyễn Thị Phúc Lộc", "240114106"]
        ];
        downloadExcelArray(dataSV, [6, 15, 24, 12, 48, 24, 15], "DS_Sinh_Vien", "Mau_Danh_Sach_SV_DoAn.xlsx");
    });

    document.getElementById('btn-sample-gv')?.addEventListener('click', () => {
        const dataGV = [
            ["Mã GV", "Họ và tên Giảng viên", "Đơn vị công tác"],
            ["230114084", "ThS. Phạm Thị Nga", "Khoa Công nghệ Thực phẩm - ĐH Duy Tân"],
            ["240114106", "ThS. Nguyễn Thị Phúc Lộc", "Khoa Công nghệ Thực phẩm - ĐH Duy Tân"],
            ["110114025", "TS. Phan Thị Việt Hà", "Khoa Công nghệ Thực phẩm - ĐH Duy Tân"]
        ];
        downloadExcelArray(dataGV, [15, 28, 38], "DS_Giang_Vien", "Mau_Danh_Sach_GV.xlsx");
    });

    document.getElementById('btn-sample-hd')?.addEventListener('click', () => {
        const dataHD = [
            ["Hội đồng", "Mã GV", "Họ và tên GV", "Đơn vị", "Ngày bảo vệ", "Chức danh", "Tên đề tài", "MSSV", "Họ và tên SV", "Lớp"],
            ["1", "110114025", "TS. Phan Thị Việt Hà", "Đại học Duy Tân", "20/06/2026 - 17:00 - Phòng 111 - 120 Hoàng Minh Thảo", "Chủ tịch", "Nghiên cứu quy trình sản xuất trà Cascara từ vỏ cà phê", "27206635615", "Võ Thị Thùy Diễm", "K27 TNM"],
            ["1", "240114106", "ThS. Nguyễn Thị Phúc Lộc", "Đại học Duy Tân", "", "Phản biện", "", "", "", ""],
            ["1", "230114084", "ThS. Phạm Thị Nga", "Đại học Duy Tân", "", "Thư ký", "", "", "", ""]
        ];
        downloadExcelArray(dataHD, [10, 14, 24, 20, 42, 14, 42, 15, 22, 12], "DS_Hoi_Dong", "Mau_Danh_Sach_Hoi_Dong.xlsx");
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

            window.studentScores[maSV] = {
                ...(window.studentScores[maSV] || {}),
                gvhd: diemGVHD,
                ct: diemCT,
                pb: diemPB,
                tk: diemTK
            };

            const svCard = document.querySelector(`.dropzone-sv .card-item[data-id="${maSV}"]`);
            if (svCard) {
                svCard.setAttribute('data-diem-gvhd', diemGVHD);
                svCard.setAttribute('data-diem-ct', diemCT);
                svCard.setAttribute('data-diem-pb', diemPB);
                svCard.setAttribute('data-diem-tk', diemTK);
            }
        });

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
        const exportData = collectExportData(meta, true);
        if (exportData.length === 0) return alert("Chưa có Hội đồng nào được tạo!");

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

    // 6. XÓA TỪNG THẺ SV HOẶC GV Ở 2 KHUNG KHO CHỜ
    document.addEventListener('click', (e) => {
        const btnRemove = e.target.closest('.btn-remove-card');
        if (btnRemove) {
            e.stopPropagation();
            const card = btnRemove.closest('.card-item');
            if (card) card.remove();
        }
    });

    // 7. ĐIỀU KHIỂN MODAL TIPS HƯỚNG DẪN THEO KHU VỰC & PHÍM TẮT F1
    let lastActiveTipTab = 'tab-setup';
    document.addEventListener('click', (e) => {
        if (e.target.closest('#pool-sv') || e.target.closest('#pool-gv')) {
            lastActiveTipTab = 'tab-pool';
        } else if (e.target.closest('#workspace-hd')) {
            lastActiveTipTab = 'tab-hd';
        } else if (e.target.closest('.footer-section')) {
            lastActiveTipTab = 'tab-export';
        } else if (e.target.closest('.setup-fields-row') || e.target.closest('#btn-read-files')) {
            lastActiveTipTab = 'tab-setup';
        }
    });

    const tipsModalEl = document.getElementById('tipsModal');
    if (tipsModalEl) {
        tipsModalEl.addEventListener('show.bs.modal', () => {
            const targetTabBtn = document.getElementById(lastActiveTipTab);
            if (targetTabBtn && window.bootstrap) {
                bootstrap.Tab.getOrCreateInstance(targetTabBtn).show();
            }
        });

        document.addEventListener('keydown', (e) => {
            if (e.key === 'F1') {
                e.preventDefault();
                if (window.bootstrap) {
                    bootstrap.Modal.getOrCreateInstance(tipsModalEl).toggle();
                }
            }
        });
    }

    // 8. QUẢN LÝ TẢI VỀ & UPLOAD MẪU PHIẾU CHẤM (GVHD, GVPB, HỘI ĐỒNG)
    const scoringTplFileNames = {
        phieuGVHD: '2_Phieu_GVHD.docx',
        phieuGVPB: '3_Phieu_GVPB.docx',
        phieuHD: '4_Phieu_Hoi_Dong.docx'
    };

    const syncScoringTemplateBadges = () => {
        let anyModified = false;
        Object.keys(scoringTplFileNames).forEach(key => {
            const isMod = Boolean(templates.modifiedFlags?.[key]);
            if (isMod) anyModified = true;
            const statusEl = document.getElementById(`status-tpl-${key}`);
            if (statusEl) {
                statusEl.textContent = isMod ? 'Đã tải mẫu mới' : 'Mặc định';
                statusEl.className = isMod ? 'badge bg-warning text-dark ms-1' : 'badge bg-secondary ms-1';
            }
        });

        const mainBadge = document.getElementById('badge-template-modified');
        if (mainBadge) {
            mainBadge.classList.toggle('d-none', !anyModified);
        }
    };

    document.getElementById('btn-open-upload-templates')?.addEventListener('click', () => {
        syncScoringTemplateBadges();
        const modalEl = document.getElementById('modalUploadScoringTemplates');
        if (modalEl && window.bootstrap) {
            bootstrap.Modal.getOrCreateInstance(modalEl).show();
        }
    });

    document.querySelectorAll('.btn-dl-scoring-tpl').forEach(btn => {
        btn.addEventListener('click', () => {
            const key = btn.getAttribute('data-tpl');
            const bin = templates[key];
            if (!bin) return alert("Chưa nạp được dữ liệu file mẫu này!");
            const blob = new Blob([bin], {
                type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            });
            saveAs(blob, scoringTplFileNames[key]);
        });
    });

    document.querySelectorAll('.btn-apply-scoring-tpl').forEach(btn => {
        btn.addEventListener('click', async () => {
            const key = btn.getAttribute('data-tpl');
            const inputEl = document.querySelector(`.input-up-scoring-tpl[data-tpl="${key}"]`);
            const file = inputEl?.files[0];

            if (!file) {
                return alert("Vui lòng chọn file Word (.docx) trước khi bấm Tải lên & Áp dụng!");
            }
            if (!file.name.toLowerCase().endsWith('.docx')) {
                return alert("⚠️ Chỉ hỗ trợ định dạng file Word (.docx)!");
            }

            try {
                const arrayBuffer = await file.arrayBuffer();
                const ok = updateScoringTemplate(key, arrayBuffer);
                if (!ok) {
                    return alert("❌ File .docx không hợp lệ hoặc bị hỏng. Vui lòng kiểm tra lại!");
                }
                syncScoringTemplateBadges();
                if (inputEl) inputEl.value = '';
                alert(`✅ Đã áp dụng mẫu "${file.name}" cho ${scoringTplFileNames[key]} trong phiên làm việc này!`);
            } catch (err) {
                alert("Lỗi khi đọc file mẫu: " + err.message);
            }
        });
    });

    document.getElementById('btn-reset-all-scoring-tpl')?.addEventListener('click', () => {
        if (confirm("Khôi phục toàn bộ 3 mẫu Phiếu chấm (GVHD, GVPB, Hội đồng) về đúng nguyên bản gốc?")) {
            resetAllScoringTemplates();
            document.querySelectorAll('.input-up-scoring-tpl').forEach(inp => inp.value = '');
            syncScoringTemplateBadges();
            alert("🔄 Đã khôi phục toàn bộ mẫu phiếu chấm về nguyên bản gốc!");
        }
    });

    // 9. XỬ LÝ GIAO DIỆN QUẢN TRỊ (ADMIN) CÓ MẬT KHẨU
    applyAdminConfigToUI();

    document.getElementById('btn-open-admin')?.addEventListener('click', () => {
        document.getElementById('admin-pass-input').value = '';
        document.getElementById('admin-login-error')?.classList.add('d-none');
        bootstrap.Modal.getOrCreateInstance(document.getElementById('adminLoginModal')).show();
    });

    document.getElementById('form-admin-login')?.addEventListener('submit', (e) => {
        e.preventDefault();
        const inputPass = document.getElementById('admin-pass-input')?.value || '';
        const cfg = getAdminConfig();

        if (inputPass === cfg.password) {
            bootstrap.Modal.getInstance(document.getElementById('adminLoginModal'))?.hide();

            document.getElementById('adm-uni-name').value = cfg.universityName;
            document.getElementById('adm-school-name').value = cfg.schoolName;
            document.getElementById('adm-qd-suffix').value = cfg.qdSuffix;
            document.getElementById('adm-campuses').value = cfg.campuses.join('\n');

            bootstrap.Modal.getOrCreateInstance(document.getElementById('adminPanelModal')).show();
        } else {
            document.getElementById('admin-login-error')?.classList.remove('d-none');
        }
    });

    // Lưu Tab 1: Header, Đuôi QĐ & Cơ sở
    document.getElementById('btn-adm-save-general')?.addEventListener('click', () => {
        const campuses = document.getElementById('adm-campuses').value
            .split('\n').map(s => s.trim()).filter(Boolean);

        saveAdminConfig({
            universityName: document.getElementById('adm-uni-name').value.trim() || 'ĐẠI HỌC DUY TÂN',
            schoolName: document.getElementById('adm-school-name').value.trim() || 'TRƯỜNG CÔNG NGHỆ & KỸ THUẬT',
            qdSuffix: document.getElementById('adm-qd-suffix').value.trim() || '/QĐ-ĐHDT-HĐTN',
            campuses: campuses.length > 0 ? campuses : ['120 Hoàng Minh Thảo']
        });
        applyAdminConfigToUI();
        alert("✅ Đã lưu cấu hình Header, đuôi Quyết định và Danh sách Cơ sở!");
    });

    document.getElementById('btn-adm-reset-general')?.addEventListener('click', () => {
        if (confirm("Khôi phục Header, đuôi QĐ và Cơ sở về mặc định?")) {
            const cfg = resetAdminConfigText();
            document.getElementById('adm-uni-name').value = cfg.universityName;
            document.getElementById('adm-school-name').value = cfg.schoolName;
            document.getElementById('adm-qd-suffix').value = cfg.qdSuffix;
            document.getElementById('adm-campuses').value = cfg.campuses.join('\n');
            applyAdminConfigToUI();
        }
    });

    // Tab 2: Tải về & Cập nhật Danh sách Ngành
    document.getElementById('btn-adm-download-nganh')?.addEventListener('click', () => {
        const aoa = [
            ['Ma_Chuyen_nganh', 'Chuyen_nganh', 'Khoa'],
            ...danhMucNganh.map(item => [item.maChuyenNganh, item.chuyenNganh, item.khoa])
        ];
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'DANH_SACH_NGANH');
        XLSX.writeFile(wb, 'DANH_SACH_NGANH.xlsx');
    });

    document.getElementById('btn-adm-save-nganh')?.addEventListener('click', async () => {
        const file = document.getElementById('adm-upload-nganh')?.files[0];
        if (!file) return alert("Vui lòng chọn file Excel Danh sách ngành!");
        const buf = await file.arrayBuffer();
        await saveAdminBinaryFile('danh_sach_nganh', buf);
        await loadDanhSachNganhSilently();
        initKhoaNganhDropdowns();
        alert("✅ Đã cập nhật Danh sách Ngành & Khoa thành công!");
    });

    document.getElementById('btn-adm-reset-nganh')?.addEventListener('click', async () => {
        if (confirm("Xóa file Danh sách ngành tùy chỉnh và dùng lại file gốc?")) {
            await deleteAdminBinaryFile('danh_sach_nganh');
            await loadDanhSachNganhSilently();
            initKhoaNganhDropdowns();
            alert("🔄 Đã khôi phục Danh sách ngành gốc!");
        }
    });

    // Tab 3: Tải về & Cập nhật 4 file Template Word
    const tplFileNames = {
        bienBan: '1_BIEN_BAN_BAO_VE_DATN.docx',
        phieuGVHD: '2_Phieu_GVHD.docx',
        phieuGVPB: '3_Phieu_GVPB.docx',
        phieuHD: '4_Phieu_Hoi_Dong.docx'
    };

    document.querySelectorAll('.btn-adm-dl-tpl').forEach(btn => {
        btn.addEventListener('click', () => {
            const key = btn.getAttribute('data-tpl');
            const bin = templates[key];
            if (!bin) return alert("Chưa có dữ liệu file mẫu này!");
            const blob = new Blob([bin], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
            saveAs(blob, tplFileNames[key]);
        });
    });

    document.querySelectorAll('.btn-adm-save-tpl').forEach(btn => {
        btn.addEventListener('click', async () => {
            const key = btn.getAttribute('data-tpl');
            const input = document.querySelector(`.adm-up-tpl[data-tpl="${key}"]`);
            const file = input?.files[0];
            if (!file) return alert("Vui lòng chọn file .docx trước khi bấm Lưu!");
            const buf = await file.arrayBuffer();
            await saveAdminBinaryFile(`tpl_${key}`, buf);
            await loadTemplatesSilently();
            syncScoringTemplateBadges();
            alert(`✅ Đã cập nhật mẫu ${tplFileNames[key]} thành công!`);
        });
    });

    document.getElementById('btn-adm-reset-all-tpl')?.addEventListener('click', async () => {
        if (confirm("Khôi phục toàn bộ 4 file Word Template về bản gốc trên hệ thống?")) {
            for (const key of Object.keys(tplFileNames)) {
                await deleteAdminBinaryFile(`tpl_${key}`);
            }
            await loadTemplatesSilently();
            syncScoringTemplateBadges();
            alert("🔄 Đã khôi phục toàn bộ 4 file Template gốc!");
        }
    });

    // Tab 4: Đổi mật khẩu Admin
    document.getElementById('btn-adm-change-pass')?.addEventListener('click', () => {
        const p1 = document.getElementById('adm-new-pass').value;
        const p2 = document.getElementById('adm-confirm-pass').value;
        if (!p1 || p1.length < 4) return alert("Mật khẩu mới phải có ít nhất 4 ký tự!");
        if (p1 !== p2) return alert("Mật khẩu xác nhận không khớp!");
        saveAdminConfig({ password: p1 });
        document.getElementById('adm-new-pass').value = '';
        document.getElementById('adm-confirm-pass').value = '';
        alert("🔑 Đã đổi mật khẩu quản trị thành công!");
    });
}

// 10. ĐỒNG BỘ PHIÊN LÀM VIỆC TRỰC TUYẾN QUA GOOGLE APPS SCRIPT
    const showCloudStatus = (msg, type = 'info') => {
        const el = document.getElementById('cloud-sync-status');
        if (!el) return;
        el.className = `alert alert-${type} py-2 px-2 small mb-0`;
        el.textContent = msg;
        el.classList.remove('d-none');
    };

    document.getElementById('btn-open-cloud-sync')?.addEventListener('click', () => {
        const savedMeta = getSavedSessionMeta();
        document.getElementById('cloud-gas-url').value = getGasWebAppUrl();
        document.getElementById('cloud-session-id').value = savedMeta.sessionId || '';
        document.getElementById('cloud-session-pin').value = savedMeta.pin || '';
        document.getElementById('cloud-sync-status')?.classList.add('d-none');
        bootstrap.Modal.getOrCreateInstance(document.getElementById('modalCloudSession')).show();
    });

    document.getElementById('btn-cloud-save')?.addEventListener('click', async () => {
        const gasUrl = document.getElementById('cloud-gas-url').value.trim();
        const sessionId = document.getElementById('cloud-session-id').value.trim().toUpperCase();
        const pin = document.getElementById('cloud-session-pin').value.trim();

        if (!gasUrl) return showCloudStatus("⚠️ Vui lòng nhập URL Web App Google Apps Script!", "warning");
        if (!sessionId) return showCloudStatus("⚠️ Vui lòng nhập Mã phiên làm việc!", "warning");

        setGasWebAppUrl(gasUrl);
        saveSessionMeta({ sessionId, pin });
        showCloudStatus("⏳ Đang đồng bộ dữ liệu lên Google Drive...", "info");

        try {
            const res = await syncSessionToCloud(gasUrl, sessionId, pin);
            if (res.ok) {
                showCloudStatus(`✅ ${res.message}`, "success");
            } else {
                showCloudStatus(`❌ ${res.message}`, "danger");
            }
        } catch (err) {
            showCloudStatus(`❌ Lỗi kết nối GAS: ${err.message}`, "danger");
        }
    });

    document.getElementById('btn-cloud-load')?.addEventListener('click', async () => {
        const gasUrl = document.getElementById('cloud-gas-url').value.trim();
        const sessionId = document.getElementById('cloud-session-id').value.trim().toUpperCase();
        const pin = document.getElementById('cloud-session-pin').value.trim();

        if (!gasUrl) return showCloudStatus("⚠️ Vui lòng nhập URL Web App Google Apps Script!", "warning");
        if (!sessionId) return showCloudStatus("⚠️ Vui lòng nhập Mã phiên làm việc!", "warning");

        setGasWebAppUrl(gasUrl);
        saveSessionMeta({ sessionId, pin });
        showCloudStatus("⏳ Đang tải phiên làm việc từ Google Drive...", "info");

        try {
            const res = await loadSessionFromCloud(gasUrl, sessionId, pin);
            if (res.ok && res.payload) {
                restoreSessionFromPayload(res.payload);
                syncScoringTemplateBadges();
                showCloudStatus(`✅ Đã khôi phục phiên [${res.sessionId}] (${res.summary})!`, "success");
            } else {
                showCloudStatus(`❌ ${res.message}`, "danger");
            }
        } catch (err) {
            showCloudStatus(`❌ Lỗi kết nối GAS: ${err.message}`, "danger");
        }
    });

function initKhoaNganhDropdowns() {
    const khoaSelect = document.getElementById('khoa-select');
    const nganhSelect = document.getElementById('nganh-select');
    if (!khoaSelect || !nganhSelect) return;

    if (danhMucNganh.length === 0) {
        khoaSelect.innerHTML = `<option value="Khoa Công nghệ Thực phẩm">Khoa Công nghệ Thực phẩm</option>`;
        nganhSelect.innerHTML = `<option value="Công nghệ Thực phẩm">Công nghệ Thực phẩm</option>`;
        return;
    }

    const uniqueKhoas = [...new Set(danhMucNganh.map(item => item.khoa))];
    khoaSelect.innerHTML = uniqueKhoas
        .map(k => `<option value="${k}">${k}</option>`)
        .join('');

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

    updateNganhOptions(khoaSelect.value);

    khoaSelect.addEventListener('change', () => {
        updateNganhOptions(khoaSelect.value);
    });

    window.syncKhoaNganhFromLop = (lopStr) => {
        const info = findNganhInfoByLop(lopStr);
        if (info) {
            khoaSelect.value = info.khoa;
            updateNganhOptions(info.khoa, info.chuyenNganh);
        }
    };
}

function getSetupMetadata() {
    const cfg = getAdminConfig();
    const rawNgayKy = (document.getElementById('ngay-ky-qd')?.value || '').trim();
    let formattedNgayKy = rawNgayKy;
    if (rawNgayKy.includes('-')) {
        const parts = rawNgayKy.split('-');
        if (parts.length === 3) {
            formattedNgayKy = `${parts[2]}/${parts[1]}/${parts[0]}`;
        }
    }

    const rawSoQD = (document.getElementById('so-qd')?.value || '').trim();
    const soPart = rawSoQD || '................';
    const fullSoQD = soPart.includes('/') ? soPart : `${soPart}${cfg.qdSuffix || '/QĐ-ĐHDT-HĐTN'}`;

    return {
        khoa: (document.getElementById('khoa-select')?.value || '').trim(),
        nganh: (document.getElementById('nganh-select')?.value || '').trim(),
        dotThang: (document.getElementById('dot-thang')?.value || '').trim(),
        namCham: (document.getElementById('nam-cham')?.value || '').trim(),
        soQD: fullSoQD,
        ngayKyQD: formattedNgayKy
    };
}

function renderSinhVienPool(svList) {
    const pool = document.getElementById('pool-sv');
    if (!pool) return;
    pool.innerHTML = '';

    svList.forEach(sv => {
        const maGVDisplay = sv.maGVHD ? ` [${sv.maGVHD}]` : '';
        const cardHTML = `
            <div class="card-item position-relative mb-2 p-2 pe-4 border rounded bg-white shadow-sm" 
                 data-id="${sv.maSV}" 
                 data-name="${sv.hoTen}" 
                 data-lop="${sv.lop}" 
                 data-detai="${sv.deTai}" 
                 data-magvhd="${sv.maGVHD || ''}" 
                 data-gvhd="${sv.gvhd || ''}" 
                 data-nganh="${sv.nganh || ''}">
                <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card" title="Xóa sinh viên này">✕</button>
                <div class="fw-bold text-primary">${sv.hoTen} (${sv.maSV})</div>
                <div class="text-muted small">Lớp: ${sv.lop} | GVHD: <span class="fw-bold text-dark">${sv.gvhd || 'Chưa rõ'}${maGVDisplay}</span></div>
                <div class="text-dark small text-truncate" title="${sv.deTai}">Đề tài: ${sv.deTai}</div>
            </div>`;
        pool.insertAdjacentHTML('beforeend', cardHTML);
    });
}

function renderGiangVienPool(gvList) {
    const pool = document.getElementById('pool-gv');
    if (!pool) return;
    pool.innerHTML = '';

    gvList.forEach(gv => {
        const cardHTML = `
            <div class="card-item position-relative mb-2 p-2 pe-4 border rounded bg-white shadow-sm" 
                 data-id="${gv.maGV}" data-name="${gv.hoTenGV}" data-donvi="${gv.donVi || ''}">
                <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card" title="Xóa giảng viên này">✕</button>
                <div class="fw-bold text-dark small">${gv.hoTenGV}</div>
                <div class="text-muted text-truncate" style="font-size: 0.75rem;">${gv.donVi || 'Đại học Duy Tân'} - Mã: ${gv.maGV}</div>
            </div>`;
        pool.insertAdjacentHTML('beforeend', cardHTML);
    });
}

function collectExportData(meta, includeEmptyHD = false) {
    const exportData = [];

    document.querySelectorAll('.box-hoidong').forEach(box => {
        const hdName = box.querySelector('.hd-title')?.value.trim() || 'Hội Đồng';
        
        const gio = box.querySelector('.hd-time')?.value || '';
        const phong = box.querySelector('.hd-room')?.value.trim() || '';
        const diaDiem = box.querySelector('.hd-location')?.value || '';

        const rawDate = (box.querySelector('.hd-date')?.value || '').trim();
        let ngay = rawDate;
        let dPart = '', mPart = '', yPart = '';

        if (rawDate.includes('-')) {
            const parts = rawDate.split('-'); // [yyyy, mm, dd]
            if (parts.length === 3) {
                [yPart, mPart, dPart] = parts;
                ngay = `${dPart}/${mPart}/${yPart}`;
            }
        } else if (rawDate.includes('/')) {
            const parts = rawDate.split('/'); // [dd, mm, yyyy]
            if (parts.length === 3) {
                [dPart, mPart, yPart] = parts;
                ngay = `${dPart.padStart(2, '0')}/${mPart.padStart(2, '0')}/${yPart}`;
            }
        }

        const getGVInfo = (role) => {
            const el = box.querySelector(`.dropzone-gv[data-role="${role}"] .card-item`);
            return {
                name: el ? (el.getAttribute('data-name') || '') : '',
                id: el ? (el.getAttribute('data-id') || '') : '',
                donVi: el ? (el.getAttribute('data-donvi') || 'Đại học Duy Tân') : 'Đại học Duy Tân'
            };
        };

        const ctInfo = getGVInfo('Chủ tịch');
        const pbInfo = getGVInfo('Phản biện');
        const tkInfo = getGVInfo('Thư ký');

        const baseHD = {
            khoa: meta.khoa,
            nganh: meta.nganh || '',
            so_hd: hdName,
            ngay: ngay,
            ngay_so: dPart,
            thang_so: mPart,
            nam_so: yPart,
            gio: gio,
            phong: phong,
            dia_diem: diaDiem,
            chu_tich: ctInfo.name,
            chu_tich_id: ctInfo.id,
            chu_tich_don_vi: ctInfo.donVi,
            phan_bien: pbInfo.name,
            phan_bien_id: pbInfo.id,
            phan_bien_don_vi: pbInfo.donVi,
            thu_ky: tkInfo.name,
            thu_ky_id: tkInfo.id,
            thu_ky_don_vi: tkInfo.donVi
        };

        const svCards = box.querySelectorAll('.dropzone-sv .card-item');

        if (svCards.length > 0) {
            svCards.forEach(svEl => {
                const svId = svEl.getAttribute('data-id') || '';
                const svNganh = (svEl.getAttribute('data-nganh') || '').trim();
                const scoreObj = window.studentScores ? (window.studentScores[svId] || {}) : {};

                exportData.push({
                    ...baseHD,
                    nganh: meta.nganh || svNganh || '',
                    ho_ten: svEl.getAttribute('data-name') || '',
                    ma_sv: svId,
                    lop: svEl.getAttribute('data-lop') || '',
                    de_tai: svEl.getAttribute('data-detai') || '',
                    gvhd: svEl.getAttribute('data-gvhd') || '',
                    diem_gvhd: scoreObj.gvhd ?? svEl.getAttribute('data-diem-gvhd') ?? '',
                    diem_ct: scoreObj.ct ?? svEl.getAttribute('data-diem-ct') ?? '',
                    diem_pb: scoreObj.pb ?? svEl.getAttribute('data-diem-pb') ?? '',
                    diem_tk: scoreObj.tk ?? svEl.getAttribute('data-diem-tk') ?? '',
                    diem_chu: scoreObj.chu ?? ''
                });
            });
        } else if (includeEmptyHD) {
            exportData.push({
                ...baseHD,
                ho_ten: '',
                ma_sv: '',
                lop: '',
                de_tai: '',
                gvhd: '',
                isEmptyHD: true
            });
        }
    });

    return exportData;
}
