// js/main.js
import { initGlobalSortables, createHoiDongBox, renderHoidongFromExcel } from './hoidong-manager.js';
import { exportExcelReport, exportDSHoiDongExcel, exportPhanBienExcel } from './excel-exporter.js';
import { 
    readExcelFile, parseSinhVienRows, parseGiangVienRows, parseHoiDongRows,
    loadDanhSachNganhSilently, danhMucNganh, findNganhInfoByLop 
} from './excel-reader.js';
import { 
    loadTemplatesSilently, templates,
    extractPhieuHDCriteriaTable, applyPhieuHDCriteriaTable, resetPhieuHDToOriginal 
} from './templates.js';
import { 
    getAdminConfig, saveAdminConfig, resetAdminConfigText, 
    saveAdminBinaryFile, deleteAdminBinaryFile 
} from './admin-config.js';
import { exportWordZIP, previewPhieuHDTemplate } from './docx-exporter.js';

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

    // Cập nhật lại dropdown Cơ sở trên tất cả các Box Hội đồng đang mở
    document.querySelectorAll('.hd-location').forEach(sel => {
        const currVal = sel.value;
        sel.innerHTML = [
            `<option value="">-- Chọn cơ sở --</option>`,
            ...cfg.campuses.map(c => `<option value="${c}" ${c === currVal ? 'selected' : ''}>${c}</option>`)
        ].join('');
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

    // 1. SỰ KIỆN NẠP DỮ LIỆU EXCEL
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
            if (fileGV) {
                const rows = await readExcelFile(fileGV);
                const gvList = parseGiangVienRows(rows);
                gvList.forEach(gv => {
                    if (gv.maGV) gvMapById[gv.maGV.toUpperCase()] = gv.hoTenGV;
                });
                renderGiangVienPool(gvList);
            }

            if (fileSV) {
                const rows = await readExcelFile(fileSV);
                const svList = parseSinhVienRows(rows);

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

            let hdSummaryMsg = '';
            if (fileHD) {
                const rows = await readExcelFile(fileHD);
                const parsedHoidongs = parseHoiDongRows(rows);
                const { mergeLogs, conflictLogs, duplicateSVLogs } = renderHoidongFromExcel(parsedHoidongs);

                if (mergeLogs.length > 0) {
                    hdSummaryMsg += `\n\n🔄 TỰ ĐỘNG GỘP & ĐÁNH LẠI SỐ HỘI ĐỒNG:\n${mergeLogs.join('\n')}`;
                }
                if (duplicateSVLogs.length > 0) {
                    hdSummaryMsg += `\n\n⚠️ CẢNH BÁO TRÙNG LẶP SINH VIÊN TRONG FILE HỘI ĐỒNG:\n${duplicateSVLogs.join('\n')}`;
                }
                if (conflictLogs.length > 0) {
                    hdSummaryMsg += `\n\n⛔ CẢNH BÁO VI PHẠM GVHD TRONG FILE HỘI ĐỒNG:\n${conflictLogs.join('\n')}`;
                }
            }

            alert(`✅ Nạp dữ liệu Excel thành công!${hdSummaryMsg}`);

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

            // Đồng bộ lưu điểm vào cả window.studentScores lẫn thuộc tính DOM của thẻ SV
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

        // 8. CHỈNH SỬA TẠM THỜI BẢNG MỤC II - MẪU 4. PHIẾU HỘI ĐỒNG (2 CỘT, THÊM/XÓA HÀNG, TỰ TÍNH TỔNG)
    const btnEditTplHD = document.getElementById('btn-edit-template-hd');
    const modalEditTplEl = document.getElementById('modalEditTemplateHD');
    const badgeModified = document.getElementById('badge-template-hd-modified');

    // Hàm tự động cộng tổng cột Thang điểm (hỗ trợ cả dấu phẩy "4,0" lẫn dấu chấm "4.0")
    const recalcTemplateHDTotal = () => {
        let sum = 0;
        let hasComma = false;
        document.querySelectorAll('.tpl-hd-score-input').forEach(inp => {
            const raw = inp.value.trim();
            if (raw.includes(',')) hasComma = true;
            const num = parseFloat(raw.replace(',', '.'));
            if (!isNaN(num)) sum += num;
        });

        const totalInput = document.getElementById('tpl-hd-total-score');
        if (totalInput) {
            const formatted = Number.isInteger(sum) ? String(sum) : sum.toFixed(1);
            totalInput.value = hasComma ? formatted.replace('.', ',') : formatted;
        }
    };

    // Hàm tạo 1 hàng HTML (2 cột: Nội dung | Thang điểm + Nút xóa)
    const createCriteriaRowHTML = (content = '', score = '') => `
        <tr class="tpl-hd-row">
            <td class="align-middle p-1">
                <textarea class="form-control form-control-sm tpl-hd-content-input" 
                          rows="${content.length > 55 ? 2 : 1}" 
                          placeholder="Nhập nội dung đánh giá...">${content}</textarea>
            </td>
            <td class="align-middle p-1" style="width: 105px;">
                <input type="text" class="form-control form-control-sm text-center fw-bold tpl-hd-score-input" 
                       value="${score}" placeholder="0,0">
            </td>
            <td class="align-middle text-center p-1" style="width: 42px;">
                <button type="button" class="btn btn-sm btn-outline-danger px-2 py-0 fw-bold btn-remove-tpl-row" title="Xóa hàng này">✕</button>
            </td>
        </tr>
    `;

    const renderTemplateHDEditor = () => {
        const listContainer = document.getElementById('template-hd-editor-list');
        if (!listContainer) return;

        const { rows, totalScore } = extractPhieuHDCriteriaTable();
        if (rows.length === 0) {
            listContainer.innerHTML = '<div class="text-danger text-center py-3">Không đọc được bảng Mục II trong file 4_Phieu_Hoi_Dong.docx!</div>';
            return;
        }

        const rowsHTML = rows.map(r => createCriteriaRowHTML(r.content, r.score)).join('');

        listContainer.innerHTML = `
            <div class="table-responsive bg-white rounded border shadow-sm">
                <table class="table table-bordered table-sm mb-0 align-middle">
                    <thead class="table-primary text-center small">
                        <tr>
                            <th>Nội dung đánh giá</th>
                            <th style="width: 105px;">Thang điểm</th>
                            <th style="width: 42px;">Xóa</th>
                        </tr>
                    </thead>
                    <tbody id="tpl-hd-tbody">
                        ${rowsHTML}
                    </tbody>
                    <tfoot class="table-light">
                        <tr>
                            <td class="text-end fw-bold small pe-2">TỔNG ĐIỂM:</td>
                            <td class="p-1">
                                <input type="text" id="tpl-hd-total-score" class="form-control form-control-sm text-center fw-bold text-danger bg-white" value="${totalScore}">
                            </td>
                            <td></td>
                        </tr>
                    </tfoot>
                </table>
            </div>
            <div class="mt-2">
                <button type="button" id="btn-add-tpl-row" class="btn btn-outline-primary btn-sm fw-bold">
                    ➕ Chèn thêm hàng tiêu chí
                </button>
            </div>
        `;
    };

    // Lắng nghe thêm hàng, xóa hàng và tự động cộng tổng điểm ngay trong Modal
    document.addEventListener('click', (e) => {
        if (e.target.closest('#btn-add-tpl-row')) {
            const tbody = document.getElementById('tpl-hd-tbody');
            tbody?.insertAdjacentHTML('beforeend', createCriteriaRowHTML('', '0,0'));
            recalcTemplateHDTotal();
        }
        const btnDelRow = e.target.closest('.btn-remove-tpl-row');
        if (btnDelRow) {
            const tbody = document.getElementById('tpl-hd-tbody');
            if (tbody && tbody.querySelectorAll('.tpl-hd-row').length <= 1) {
                return alert("Bảng đánh giá phải có ít nhất 1 hàng nội dung!");
            }
            btnDelRow.closest('.tpl-hd-row')?.remove();
            recalcTemplateHDTotal();
        }
    });

    document.addEventListener('input', (e) => {
        if (e.target.classList.contains('tpl-hd-score-input')) {
            recalcTemplateHDTotal();
        }
    });

    // Thu thập dữ liệu bảng từ Modal
    const collectEditedTableData = () => {
        const editedRows = [];
        document.querySelectorAll('#tpl-hd-tbody .tpl-hd-row').forEach(tr => {
            const content = tr.querySelector('.tpl-hd-content-input')?.value.trim() || '';
            const score = tr.querySelector('.tpl-hd-score-input')?.value.trim() || '';
            if (content || score) {
                editedRows.push({ content, score });
            }
        });
        const totalScore = document.getElementById('tpl-hd-total-score')?.value.trim() || '10';
        return { editedRows, totalScore };
    };

    btnEditTplHD?.addEventListener('click', () => {
        renderTemplateHDEditor();
        if (modalEditTplEl && window.bootstrap) {
            bootstrap.Modal.getOrCreateInstance(modalEditTplEl).show();
        }
    });

    // Áp dụng các thay đổi vào bộ nhớ phiên làm việc hiện tại
    document.getElementById('btn-apply-template-hd')?.addEventListener('click', () => {
        const { editedRows, totalScore } = collectEditedTableData();
        if (editedRows.length === 0) return alert("Vui lòng nhập ít nhất 1 hàng tiêu chí đánh giá!");

        const ok = applyPhieuHDCriteriaTable(editedRows, totalScore);
        if (ok) {
            badgeModified?.classList.remove('d-none');
            if (modalEditTplEl && window.bootstrap) {
                bootstrap.Modal.getInstance(modalEditTplEl)?.hide();
            }
            alert("✅ Đã áp dụng chỉnh sửa tạm thời cho Mẫu 4. Phiếu Hội đồng trong phiên làm việc này!");
        } else {
            alert("❌ Có lỗi khi cập nhật nội dung file mẫu!");
        }
    });

    // Tải thử 1 bản xem trước Mẫu 4 đã chỉnh sửa
    document.getElementById('btn-preview-template-hd')?.addEventListener('click', async () => {
        const { editedRows, totalScore } = collectEditedTableData();
        if (editedRows.length === 0) return alert("Vui lòng nhập ít nhất 1 hàng tiêu chí đánh giá!");

        const ok = applyPhieuHDCriteriaTable(editedRows, totalScore);
        if (ok) {
            badgeModified?.classList.remove('d-none');
            if (typeof previewPhieuHDTemplate === 'function') {
                await previewPhieuHDTemplate();
            }
        }
    });

    // Khôi phục Mẫu 4 về nguyên bản gốc
    document.getElementById('btn-reset-template-hd')?.addEventListener('click', () => {
        if (confirm("Khôi phục Mẫu 4. Phiếu Hội đồng về đúng nguyên bản ban đầu?")) {
            const ok = resetPhieuHDToOriginal();
            if (ok) {
                badgeModified?.classList.add('d-none');
                renderTemplateHDEditor();
                alert("🔄 Đã khôi phục Mẫu 4 về nguyên bản gốc!");
            }
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

            // Đổ dữ liệu cấu hình hiện tại vào form Admin
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
            alert(`✅ Đã cập nhật mẫu ${tplFileNames[key]} thành công!`);
        });
    });

    document.getElementById('btn-adm-reset-all-tpl')?.addEventListener('click', async () => {
        if (confirm("Khôi phục toàn bộ 4 file Word Template về bản gốc trên hệ thống?")) {
            for (const key of Object.keys(tplFileNames)) {
                await deleteAdminBinaryFile(`tpl_${key}`);
            }
            await loadTemplatesSilently();
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
    const fullSoQD = soPart.includes('/') ? soPart : `${soPart}/QĐ-ĐHDT-HĐTN`;

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
        const rawDate = box.querySelector('.hd-date')?.value || '';
        const gio = box.querySelector('.hd-time')?.value || '';
        const phong = box.querySelector('.hd-room')?.value.trim() || '';
        const diaDiem = box.querySelector('.hd-location')?.value || '';

        let ngay = rawDate;
        if (rawDate.includes('-')) {
            const parts = rawDate.split('-');
            if (parts.length === 3) ngay = `${parts[2]}/${parts[1]}/${parts[0]}`;
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
            // Hội đồng chưa có sinh viên nhưng vẫn xuất khung Hội đồng & Giảng viên
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
