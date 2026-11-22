// js/main.js (Phần 2/3)
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
import { detectDuplicateTeacherNameDifferentId, detectDuplicateStudentNameDifferentId, detectDuplicateStudentId, collectMissingCodeWarnings } from './validators.js';
import {
    sessionState, getGasWebAppUrl, setGasWebAppUrl, getSavedSessionMeta, saveSessionMeta,
    createCloudSession, loadCloudSession, saveCloudHoiDong, toggleCloudHDLock,
    pullLatestScoresFromCloud, pushDeltaScoresToCloud, restoreSessionFromPayload,
    applySessionPermissionsToUI, setupAutoSaveAndLockOnClose, disconnectCloudSession
} from './session-sync.js';

window.studentScores = window.studentScores || {};

const SCORING_TEMPLATE_FILENAMES = {
    phieuGVHD: '2_Phieu_GVHD.docx',
    phieuGVPB: '3_Phieu_GVPB.docx',
    phieuHD: '4_Phieu_Hoi_Dong.docx'
};

const ADMIN_TEMPLATE_FILENAMES = {
    bienBan: '1_BIEN_BAN_BAO_VE_DATN.docx',
    phieuGVHD: '2_Phieu_GVHD.docx',
    phieuGVPB: '3_Phieu_GVPB.docx',
    phieuHD: '4_Phieu_Hoi_Dong.docx'
};

// Đồng bộ trạng thái badge "Đã tải mẫu mới / Mặc định" cho cả modal phiếu chấm và nút ngoài
export function syncScoringTemplateBadges() {
    let anyModified = false;
    Object.keys(SCORING_TEMPLATE_FILENAMES).forEach(key => {
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
}

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

function showExcelReviewModal({ title, requireConfirm = false, htmlSections = [] }) {
    return new Promise((resolve) => {
        const modalEl = document.getElementById('modalExcelReviewReport');
        if (!modalEl || !window.bootstrap) {
            if (requireConfirm) {
                return resolve(confirm(htmlSections.map(s => s.replace(/<[^>]+>/g, ' ')).join('\n\n')));
            }
            alert(htmlSections.map(s => s.replace(/<[^>]+>/g, ' ')).join('\n\n'));
            return resolve(true);
        }

        const openModalNow = () => {
            const titleEl = document.getElementById('review-modal-title');
            const bodyEl = document.getElementById('review-modal-body');
            const rawBtnCancel = document.getElementById('btn-review-cancel');
            const rawBtnConfirm = document.getElementById('btn-review-confirm');
            const btnCloseX = document.getElementById('btn-review-close-x');

            const btnConfirm = rawBtnConfirm ? rawBtnConfirm.cloneNode(true) : null;
            const btnCancel = rawBtnCancel ? rawBtnCancel.cloneNode(true) : null;
            if (rawBtnConfirm && btnConfirm) rawBtnConfirm.replaceWith(btnConfirm);
            if (rawBtnCancel && btnCancel) rawBtnCancel.replaceWith(btnCancel);

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

            const oldInst = bootstrap.Modal.getInstance(modalEl);
            if (oldInst) oldInst.dispose();

            const modalInstance = new bootstrap.Modal(modalEl, {
                backdrop: requireConfirm ? 'static' : true,
                keyboard: !requireConfirm
            });

            let settled = false;
            let chosenValue = !requireConfirm;

            const finish = (result) => {
                if (settled) return;
                settled = true;
                btnConfirm?.removeEventListener('click', onConfirm);
                btnCancel?.removeEventListener('click', onCancel);
                modalEl.removeEventListener('hidden.bs.modal', onHidden);
                resolve(result);
            };

            const onConfirm = () => {
                chosenValue = true;
                modalInstance.hide();
            };

            const onCancel = () => {
                chosenValue = false;
                modalInstance.hide();
            };

            const onHidden = () => {
                finish(chosenValue);
            };

            btnConfirm?.addEventListener('click', onConfirm);
            btnCancel?.addEventListener('click', onCancel);
            modalEl.addEventListener('hidden.bs.modal', onHidden);

            modalInstance.show();
        };

        // Nếu modal đang mở từ bước trước, đợi đóng xong hẳn rồi mới mở modal kế tiếp
        if (modalEl.classList.contains('show')) {
            modalEl.addEventListener('hidden.bs.modal', openModalNow, { once: true });
            bootstrap.Modal.getInstance(modalEl)?.hide();
        } else {
            openModalNow();
        }
    });
}

async function loadHtmlPartials() {
    const container = document.getElementById('partials-container');
    if (!container) return;

    const partialFiles = [
        './partials/modal-nhap-diem-gan-sv.html',
        './partials/modal-tips-templates.html',
        './partials/modal-admin.html',
        './partials/modal-excel-cloud.html'
    ];

    const htmlChunks = await Promise.all(
        partialFiles.map(async (url) => {
            const resp = await fetch(url);
            if (!resp.ok) throw new Error(`Không tải được partial: ${url} (${resp.status})`);
            return await resp.text();
        })
    );

    container.innerHTML = htmlChunks.join('\n');
}

function bindMobileTabEvents() {
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
}

function bindExcelAndHoiDongEvents() {
    const btnReadFiles = document.getElementById('btn-read-files');
    btnReadFiles?.addEventListener('click', async () => {
        const fileSV = document.getElementById('upload-sv')?.files[0];
        const fileGV = document.getElementById('upload-gv')?.files[0];
        const fileHD = document.getElementById('upload-hd')?.files[0];

        if (!fileSV && !fileGV && !fileHD) {
            return alert("Vui lòng chọn ít nhất 1 file Excel để nạp dữ liệu!");
        }

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
            }

            if (fileSV) {
                const rows = await readExcelFile(fileSV);
                svList = parseSinhVienRows(rows);

                svList.forEach(sv => {
                    if (!sv.gvhd && sv.maGVHD && gvMapById[sv.maGVHD.toUpperCase()]) {
                        sv.gvhd = gvMapById[sv.maGVHD.toUpperCase()];
                    }
                });
            }

            if (fileHD) {
                const rows = await readExcelFile(fileHD);
                parsedHoidongs = parseHoiDongRows(rows);
                console.table(
                    Object.keys(parsedHoidongs)
                        .filter(k => !k.startsWith('__'))
                        .map(k => ({
                            key: k,
                            soHD: parsedHoidongs[k].soHD,
                            rawIndex: parsedHoidongs[k].rawIndex,
                            ngayBV: parsedHoidongs[k].ngayBV,
                            gioBV: parsedHoidongs[k].gioBV,
                            thanhVien: parsedHoidongs[k].thanhVien.map(g => `${g.chucDanh}: ${g.hoTenGV} [${g.maGV}]`).join(' | '),
                            soSV: parsedHoidongs[k].sinhViens.length
                        }))
                );
            }

            // 1. Chặn nếu có tên nhưng thiếu Mã SV / Mã GV (không render)
            const missingCodeWarnings = collectMissingCodeWarnings({ svList, gvList, parsedHoidongs });
            if (missingCodeWarnings.length > 0) {
                const missingHtml = `
                    <div class="alert alert-danger py-2 px-3 mb-2 small">
                        <div class="fw-bold mb-1">❌ Phát hiện dòng có tên nhưng thiếu Mã SV / Mã GV:</div>
                        <ul class="mb-0 ps-3">
                            ${missingCodeWarnings.map(w => `<li class="mb-1">${w}</li>`).join('')}
                        </ul>
                        <div class="mt-2 fw-semibold">Vui lòng bổ sung đầy đủ Mã SV / Mã GV vào file Excel và tải lại.</div>
                    </div>
                `;

                await showExcelReviewModal({
                    title: '⚠️ CẢNH BÁO THIẾU MÃ ĐỊNH DANH TRONG FILE EXCEL',
                    requireConfirm: false,
                    htmlSections: [missingHtml]
                });
                return;
            }

            // 2. Chặn nếu trùng Mã SV trong file SV hoặc file HĐ (không render)
            const dupStudentIdList = detectDuplicateStudentId({ svList, parsedHoidongs });
            if (dupStudentIdList.length > 0) {
                const dupIdHtml = dupStudentIdList.map((c, idx) => {
                    const detailLines = c.items.map(it =>
                        `<li><strong>${it.hoTen}</strong>${it.lop ? ` — Lớp: <code>${it.lop}</code>` : ''} <span class="text-muted">(${it.location})</span></li>`
                    ).join('');

                    return `
                        <div class="bg-white border border-danger rounded p-2 mb-2 shadow-sm">
                            <div class="fw-bold text-danger small mb-1">
                                #${idx + 1}. Trùng Mã SV <span class="badge bg-danger font-monospace">${c.maSV}</span> trong <u>${c.source}</u> (${c.items.length} lần):
                            </div>
                            <ul class="mb-0 ps-3 small">${detailLines}</ul>
                        </div>
                    `;
                }).join('');

                await showExcelReviewModal({
                    title: '❌ LỖI TRÙNG MÃ SINH VIÊN — CẦN SỬA VÀ TẢI LẠI FILE',
                    requireConfirm: false,
                    htmlSections: [
                        `
                        <div class="alert alert-danger py-2 px-3 mb-2 small">
                            <div class="fw-bold mb-1">⛔ Phát hiện Mã SV bị trùng lặp trong file Excel:</div>
                            <div>Vui lòng kiểm tra và sửa lại Mã SV bị trùng trong file Excel, sau đó tải lại file trước khi tiếp tục nạp.</div>
                        </div>
                        ${dupIdHtml}
                        `
                    ]
                });
                return;
            }

            // 3. Xác nhận nếu GV trùng tên nhưng khác Mã GV
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

                if (!userConfirmed) return;
            }

            // 4. Xác nhận nếu SV trùng tên nhưng khác Mã SV
            const dupSVNameDiffIdList = detectDuplicateStudentNameDifferentId({ svList, parsedHoidongs });
            if (dupSVNameDiffIdList.length > 0 && (fileSV || fileHD)) {
                const dupSVItemsHTML = dupSVNameDiffIdList.map((item, idx) => {
                    const variantsHTML = item.variants.map(v => `
                        <li class="mb-1">
                            <span class="badge bg-dark font-monospace">Mã SV: ${v.maSV}</span>
                            <strong>${v.names}</strong>
                            <div class="text-muted small ps-2">↳ Xuất hiện tại: ${v.locations.join(' | ')}</div>
                        </li>
                    `).join('');

                    return `
                        <div class="bg-white border border-warning rounded p-2 mb-2 shadow-sm">
                            <div class="fw-bold text-danger small mb-1">#${idx + 1}. Trùng họ tên sinh viên nhưng khác Mã SV:</div>
                            <ul class="mb-0 ps-3 small">${variantsHTML}</ul>
                        </div>
                    `;
                }).join('');

                const confirmSVSection = `
                    <div class="alert alert-warning py-2 small mb-2">
                        <strong>⚠️ PHÁT HIỆN SINH VIÊN TRÙNG TÊN NHƯNG KHÁC MÃ SV:</strong><br>
                        Vui lòng đối chiếu giữa file SV và file Hội đồng xem đây là <strong>2 sinh viên khác nhau trùng tên</strong> hay <strong>nhập sai Mã SV</strong>. Bấm <em>"Xác nhận file đã đúng, tiếp tục nạp"</em> để tiếp tục, hoặc bấm <em>"Hủy"</em> để sửa lại file Excel.
                    </div>
                    ${dupSVItemsHTML}
                `;

                const userConfirmedSV = await showExcelReviewModal({
                    title: '⚠️ XÁC NHẬN RÀ SOÁT MÃ SINH VIÊN TRÙNG TÊN',
                    requireConfirm: true,
                    htmlSections: [confirmSVSection]
                });

                if (!userConfirmedSV) return;
            }

            // 5. Đã qua toàn bộ rà soát & xác nhận -> Render Kho GV, Kho SV, và Hội đồng
            if (fileGV) {
                renderGiangVienPool(gvList);
            }

            if (fileSV) {
                renderSinhVienPool(svList);

                const firstSVWithLop = svList.find(s => s.lop);
                if (firstSVWithLop && typeof window.syncKhoaNganhFromLop === 'function') {
                    window.syncKhoaNganhFromLop(firstSVWithLop.lop);
                }
            }

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
                const hdKeys = Object.keys(parsedHoidongs || {}).filter(k => !k.startsWith('__'));
                console.log('[RENDER HĐ] Số hội đồng bóc tách được:', hdKeys.length, parsedHoidongs);

                if (hdKeys.length === 0) {
                    reportSections.push(`
                        <div class="alert alert-warning py-2 small mb-2">
                            ⚠️ File Hội đồng đã được đọc nhưng <strong>không bóc tách được Hội đồng nào</strong>. Vui lòng kiểm tra lại cấu trúc cột file Excel Hội đồng.
                        </div>
                    `);
                }

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

    document.getElementById('btn-add-hd')?.addEventListener('click', () => {
        createHoiDongBox();
    });

    document.addEventListener('click', (e) => {
        const btnRemove = e.target.closest('.btn-remove-card');
        if (btnRemove) {
            e.stopPropagation();
            const card = btnRemove.closest('.card-item');
            if (card) card.remove();
        }
    });
}
// js/main.js (Phần 3/3)
const normalizeSingleScore = (rawVal) => {
    const str = String(rawVal ?? '').trim().replace(',', '.');
    if (str === '') return '';
    const num = parseFloat(str);
    if (isNaN(num)) return '';
    const clamped = Math.min(10, Math.max(0, num));
    return clamped.toFixed(1);
};

function bindScoreAndExportEvents() {
    const btnNhapDiem = document.getElementById('btn-nhap-diem');
    btnNhapDiem?.addEventListener('click', async () => {
        if (sessionState.isCloudConnected) {
            if (!sessionState.isHDLocked) {
                return alert("⛔ Vùng Hội đồng chưa được Chủ phiên Chốt & Khóa! Chưa thể nhập điểm.");
            }
            btnNhapDiem.disabled = true;
            const oldText = btnNhapDiem.textContent;
            btnNhapDiem.textContent = "⏳ Đang tải điểm mới nhất...";
            try {
                const pullRes = await pullLatestScoresFromCloud();
                if (!pullRes.ok) {
                    return alert("❌ Không thể đồng bộ điểm từ Cloud: " + pullRes.message);
                }
            } catch (err) {
                return alert("❌ Lỗi kết nối khi kéo điểm mới nhất: " + err.message);
            } finally {
                btnNhapDiem.disabled = false;
                btnNhapDiem.textContent = oldText;
            }
        }

        const meta = getSetupMetadata();
        const exportData = collectExportData(meta);

        if (exportData.length === 0) {
            return alert("Vui lòng phân công sinh viên vào Hội đồng trước khi nhập điểm!");
        }

        const tbody = document.getElementById('tbody-nhap-diem');
        if (!tbody) return;
        tbody.innerHTML = '';

        exportData.forEach((sv, idx) => {
            const valGVHD = normalizeSingleScore(sv.diem_gvhd);
            const valCT = normalizeSingleScore(sv.diem_ct);
            const valPB = normalizeSingleScore(sv.diem_pb);
            const valTK = normalizeSingleScore(sv.diem_tk);

            const tr = `
                <tr>
                    <td>${idx + 1}</td>
                    <td class="fw-bold">${sv.ma_sv}</td>
                    <td class="text-start">${sv.ho_ten}</td>
                    <td><span class="badge bg-info text-dark">${sv.so_hd}</span></td>
                    <td><input type="text" inputmode="decimal" class="form-control form-control-sm text-center score-decimal-input input-diem-gvhd" data-masv="${sv.ma_sv}" value="${valGVHD}" placeholder=""></td>
                    <td><input type="text" inputmode="decimal" class="form-control form-control-sm text-center score-decimal-input input-diem-ct" data-masv="${sv.ma_sv}" value="${valCT}" placeholder=""></td>
                    <td><input type="text" inputmode="decimal" class="form-control form-control-sm text-center score-decimal-input input-diem-pb" data-masv="${sv.ma_sv}" value="${valPB}" placeholder=""></td>
                    <td><input type="text" inputmode="decimal" class="form-control form-control-sm text-center score-decimal-input input-diem-tk" data-masv="${sv.ma_sv}" value="${valTK}" placeholder=""></td>
                </tr>
            `;
            tbody.insertAdjacentHTML('beforeend', tr);
        });

        tbody.querySelectorAll('.score-decimal-input').forEach(inp => {
            inp.addEventListener('blur', () => {
                inp.value = normalizeSingleScore(inp.value);
            });
        });

        const modalDiem = new bootstrap.Modal(document.getElementById('modalNhapDiem'));
        modalDiem.show();
    });

    const btnSaveDiem = document.getElementById('btn-save-diem');
    btnSaveDiem?.addEventListener('click', async () => {
        document.querySelectorAll('#tbody-nhap-diem tr').forEach(tr => {
            const maSV = tr.querySelector('.input-diem-gvhd')?.getAttribute('data-masv');
            if (!maSV) return;

            const diemGVHD = normalizeSingleScore(tr.querySelector('.input-diem-gvhd')?.value);
            const diemCT = normalizeSingleScore(tr.querySelector('.input-diem-ct')?.value);
            const diemPB = normalizeSingleScore(tr.querySelector('.input-diem-pb')?.value);
            const diemTK = normalizeSingleScore(tr.querySelector('.input-diem-tk')?.value);

            window.studentScores[maSV] = {
                ...(window.studentScores[maSV] || {}),
                gvhd: diemGVHD,
                ct: diemCT,
                pb: diemPB,
                tk: diemTK
            };

            document.querySelectorAll(`.dropzone-sv .card-item[data-id="${maSV}"]`).forEach(svCard => {
                svCard.setAttribute('data-diem-gvhd', diemGVHD);
                svCard.setAttribute('data-diem-ct', diemCT);
                svCard.setAttribute('data-diem-pb', diemPB);
                svCard.setAttribute('data-diem-tk', diemTK);
            });
        });

        if (sessionState.isCloudConnected) {
            btnSaveDiem.disabled = true;
            const oldText = btnSaveDiem.textContent;
            btnSaveDiem.textContent = "⏳ Đang đồng bộ điểm...";
            try {
                const pushRes = await pushDeltaScoresToCloud();
                if (!pushRes.ok) {
                    return alert("❌ Lỗi đồng bộ điểm lên Cloud: " + pushRes.message);
                }
                bootstrap.Modal.getInstance(document.getElementById('modalNhapDiem'))?.hide();
                alert(pushRes.message || "✅ Đã lưu & đồng bộ điểm mới nhất lên Cloud!");
            } catch (err) {
                alert("❌ Lỗi kết nối Cloud khi lưu điểm: " + err.message);
            } finally {
                btnSaveDiem.disabled = false;
                btnSaveDiem.textContent = oldText;
            }
            return;
        }

        bootstrap.Modal.getInstance(document.getElementById('modalNhapDiem'))?.hide();
        alert("✅ Đã lưu điểm trực tiếp thành công!");
    });

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

    // Mở Modal thống kê Vai trò Giảng viên
    let cachedLecturerStats = [];

    const renderLecturerStatsTable = (keyword = '') => {
        const tbody = document.getElementById('tbody-thong-ke-gv');
        const tfoot = document.getElementById('tfoot-thong-ke-gv');
        const summaryEl = document.getElementById('summary-thong-ke-gv');
        if (!tbody || !tfoot) return;

        const q = keyword.trim().toLowerCase();
        const filtered = cachedLecturerStats.filter(item =>
            !q ||
            item.maGV.toLowerCase().includes(q) ||
            item.hoTen.toLowerCase().includes(q) ||
            item.donVi.toLowerCase().includes(q)
        );

        let sumHD = 0, sumCT = 0, sumPB = 0, sumTK = 0, sumAll = 0;

        if (filtered.length === 0) {
            tbody.innerHTML = `<tr><td colspan="9" class="text-muted py-3">Không có dữ liệu giảng viên phù hợp.</td></tr>`;
            tfoot.innerHTML = '';
        } else {
            tbody.innerHTML = filtered.map((item, idx) => {
                sumHD += item.huongDan;
                sumCT += item.chuTich;
                sumPB += item.phanBien;
                sumTK += item.thuKy;
                sumAll += item.tongLuot;

                return `
                    <tr>
                        <td>${idx + 1}</td>
                        <td class="font-monospace">${item.maGV || '<span class="text-muted">—</span>'}</td>
                        <td class="text-start fw-bold">${item.hoTen}</td>
                        <td class="text-start small text-muted">${item.donVi || 'Đại học Duy Tân'}</td>
                        <td>${item.huongDan || 0}</td>
                        <td>${item.chuTich || 0}</td>
                        <td>${item.phanBien || 0}</td>
                        <td>${item.thuKy || 0}</td>
                        <td class="fw-bold text-primary">${item.tongLuot}</td>
                    </tr>
                `;
            }).join('');

            tfoot.innerHTML = `
                <tr>
                    <td colspan="4" class="text-end">TỔNG CỘNG (${filtered.length} giảng viên):</td>
                    <td>${sumHD}</td>
                    <td>${sumCT}</td>
                    <td>${sumPB}</td>
                    <td>${sumTK}</td>
                    <td class="text-primary">${sumAll}</td>
                </tr>
            `;
        }

        if (summaryEl) {
            summaryEl.textContent = `Tổng số: ${cachedLecturerStats.length} giảng viên`;
        }
    };

    document.getElementById('btn-thong-ke-gv')?.addEventListener('click', () => {
        cachedLecturerStats = buildLecturerRoleStats();
        if (cachedLecturerStats.length === 0) {
            return alert("Chưa có dữ liệu Giảng viên hoặc Sinh viên để thống kê!");
        }
        const searchInput = document.getElementById('search-thong-ke-gv');
        if (searchInput) searchInput.value = '';
        renderLecturerStatsTable('');

        const modalEl = document.getElementById('modalThongKeGV');
        if (modalEl && window.bootstrap) {
            bootstrap.Modal.getOrCreateInstance(modalEl).show();
        }
    });

    document.getElementById('search-thong-ke-gv')?.addEventListener('input', (e) => {
        renderLecturerStatsTable(e.target.value);
    });

    document.getElementById('btn-export-thong-ke-gv')?.addEventListener('click', () => {
        if (cachedLecturerStats.length === 0) {
            cachedLecturerStats = buildLecturerRoleStats();
        }
        if (cachedLecturerStats.length === 0) {
            return alert("Không có dữ liệu để xuất Excel!");
        }

        const meta = getSetupMetadata();
        const headerRows = [
            ['BẢNG THỐNG KÊ DANH SÁCH GIẢNG VIÊN THEO VAI TRÒ HỘI ĐỒNG ĐATN'],
            [`Khoa: ${meta.khoa || '...'} | Ngành: ${meta.nganh || '...'} | Đợt: ${meta.dotThang || '...'} - Năm: ${meta.namCham || '...'}`],
            [],
            ['STT', 'Mã GV', 'Họ và tên Giảng viên', 'Đơn vị công tác', 'Hướng dẫn (SV)', 'Chủ tịch (HĐ)', 'Phản biện (HĐ)', 'Thư ký (HĐ)', 'Tổng lượt']
        ];

        let tHD = 0, tCT = 0, tPB = 0, tTK = 0, tAll = 0;
        const bodyRows = cachedLecturerStats.map((item, idx) => {
            tHD += item.huongDan;
            tCT += item.chuTich;
            tPB += item.phanBien;
            tTK += item.thuKy;
            tAll += item.tongLuot;
            return [
                idx + 1,
                item.maGV,
                item.hoTen,
                item.donVi || 'Đại học Duy Tân',
                item.huongDan,
                item.chuTich,
                item.phanBien,
                item.thuKy,
                item.tongLuot
            ];
        });

        bodyRows.push(['', '', 'TỔNG CỘNG', '', tHD, tCT, tPB, tTK, tAll]);

        const ws = XLSX.utils.aoa_to_sheet([...headerRows, ...bodyRows]);
        ws['!cols'] = [
            { wch: 6 }, { wch: 15 }, { wch: 28 }, { wch: 25 },
            { wch: 16 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 13 }
        ];

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Thong_Ke_GV');
        XLSX.writeFile(wb, `Thong_Ke_Vai_Tro_Giang_Vien_Dot_${meta.dotThang || 'DATN'}.xlsx`);
    });
}

function bindTipsAndScoringTemplatesEvents() {
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
            saveAs(blob, SCORING_TEMPLATE_FILENAMES[key]);
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
                alert(`✅ Đã áp dụng mẫu "${file.name}" cho ${SCORING_TEMPLATE_FILENAMES[key]} trong phiên làm việc này!`);
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
}

function bindAdminEvents() {
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

    document.querySelectorAll('.btn-adm-dl-tpl').forEach(btn => {
        btn.addEventListener('click', () => {
            const key = btn.getAttribute('data-tpl');
            const bin = templates[key];
            if (!bin) return alert("Chưa có dữ liệu file mẫu này!");
            const blob = new Blob([bin], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
            saveAs(blob, ADMIN_TEMPLATE_FILENAMES[key]);
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
            alert(`✅ Đã cập nhật mẫu ${ADMIN_TEMPLATE_FILENAMES[key]} thành công!`);
        });
    });

    document.getElementById('btn-adm-reset-all-tpl')?.addEventListener('click', async () => {
        if (confirm("Khôi phục toàn bộ 4 file Word Template về bản gốc trên hệ thống?")) {
            for (const key of Object.keys(ADMIN_TEMPLATE_FILENAMES)) {
                await deleteAdminBinaryFile(`tpl_${key}`);
            }
            await loadTemplatesSilently();
            syncScoringTemplateBadges();
            alert("🔄 Đã khôi phục toàn bộ 4 file Template gốc!");
        }
    });

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

function bindCloudSessionEvents() {
    const showCloudStatus = (msg, type = 'info') => {
        const el = document.getElementById('cloud-sync-status');
        if (!el) return;
        el.className = `alert alert-${type} py-2 px-2 small mt-2 mb-0`;
        el.textContent = msg;
        el.classList.remove('d-none');
    };

    const syncCloudModalButtons = () => {
        const btnSaveHD = document.getElementById('btn-cloud-save-hd');
        const btnToggleLock = document.getElementById('btn-cloud-toggle-lock');
        const btnDisconnect = document.getElementById('btn-cloud-disconnect');

        if (btnDisconnect) {
            btnDisconnect.classList.toggle('d-none', !sessionState.isCloudConnected);
        }

        if (!btnSaveHD || !btnToggleLock) return;

        if (sessionState.isCloudConnected && sessionState.canEditHD) {
            btnToggleLock.classList.remove('d-none');
            btnToggleLock.textContent = sessionState.isHDLocked
                ? '🔓 Mở khóa sửa Hội đồng'
                : '🔒 Chốt & Khóa Hội đồng';
            btnToggleLock.className = sessionState.isHDLocked
                ? 'btn btn-outline-warning btn-sm fw-bold flex-fill'
                : 'btn btn-warning btn-sm fw-bold flex-fill';

            btnSaveHD.classList.toggle('d-none', sessionState.isHDLocked);
        } else {
            btnSaveHD.classList.add('d-none');
            btnToggleLock.classList.add('d-none');
        }
    };

    document.getElementById('btn-open-cloud-sync')?.addEventListener('click', () => {
        const savedMeta = getSavedSessionMeta();
        document.getElementById('cloud-gas-url').value = getGasWebAppUrl();
        document.getElementById('cloud-session-id').value = savedMeta.sessionId || '';
        document.getElementById('cloud-session-pin').value = savedMeta.pin || '';
        document.getElementById('cloud-session-otp').value = savedMeta.otp || '';
        document.getElementById('cloud-sync-status')?.classList.add('d-none');
        syncCloudModalButtons();
        bootstrap.Modal.getOrCreateInstance(document.getElementById('modalCloudSession')).show();
    });

    document.getElementById('btn-cloud-create')?.addEventListener('click', async () => {
        const gasUrl = document.getElementById('cloud-gas-url').value.trim();
        const sessionId = document.getElementById('create-session-id').value.trim().toUpperCase();
        const pin = document.getElementById('create-session-pin').value.trim();
        const email = document.getElementById('create-session-email').value.trim();

        if (!gasUrl) return showCloudStatus("⚠️ Vui lòng nhập URL Web App Google Apps Script!", "warning");
        if (!sessionId) return showCloudStatus("⚠️ Vui lòng đặt Mã phiên làm việc!", "warning");
        if (!pin) return showCloudStatus("⚠️ Vui lòng đặt Mã PIN cho các thư ký nhập điểm!", "warning");
        if (!email || !email.includes('@')) return showCloudStatus("⚠️ Vui lòng nhập Email hợp lệ để nhận mã OTP!", "warning");

        setGasWebAppUrl(gasUrl);
        showCloudStatus("⏳ Đang khởi tạo phiên & gửi mã OTP qua Email...", "info");

        try {
            const res = await createCloudSession(gasUrl, sessionId, pin, email);
            if (res.ok) {
                sessionState.isCloudConnected = true;
                sessionState.sessionId = sessionId;
                sessionState.pin = pin;
                sessionState.otp = res.otp || '';
                sessionState.canEditHD = true;
                sessionState.isHDLocked = false;

                saveSessionMeta({ sessionId, pin, otp: sessionState.otp });
                document.getElementById('cloud-session-id').value = sessionId;
                document.getElementById('cloud-session-pin').value = pin;
                document.getElementById('cloud-session-otp').value = sessionState.otp;

                applySessionPermissionsToUI();
                syncCloudModalButtons();
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
        const otp = document.getElementById('cloud-session-otp').value.trim();

        if (!gasUrl) return showCloudStatus("⚠️ Vui lòng nhập URL Web App Google Apps Script!", "warning");
        if (!sessionId) return showCloudStatus("⚠️ Vui lòng nhập Mã phiên làm việc!", "warning");
        if (!pin) return showCloudStatus("⚠️ Vui lòng nhập Mã PIN!", "warning");

        setGasWebAppUrl(gasUrl);
        showCloudStatus("⏳ Đang tải phiên làm việc từ Cloud...", "info");

        try {
            const res = await loadCloudSession(gasUrl, sessionId, pin, otp);
            if (res.ok && res.payload) {
                sessionState.isCloudConnected = true;
                sessionState.sessionId = res.sessionId;
                sessionState.pin = pin;
                sessionState.otp = res.canEditHD ? otp : '';
                sessionState.canEditHD = Boolean(res.canEditHD);
                sessionState.isHDLocked = Boolean(res.isHDLocked);

                saveSessionMeta({ sessionId: res.sessionId, pin, otp: sessionState.otp });
                restoreSessionFromPayload(res.payload);
                applySessionPermissionsToUI();
                syncScoringTemplateBadges();
                syncCloudModalButtons();

                bootstrap.Modal.getInstance(document.getElementById('modalCloudSession'))?.hide();
            } else {
                showCloudStatus(`❌ ${res.message}`, "danger");
            }
        } catch (err) {
            showCloudStatus(`❌ Lỗi kết nối GAS: ${err.message}`, "danger");
        }
    });

    document.getElementById('btn-cloud-save-hd')?.addEventListener('click', async () => {
        const gasUrl = getGasWebAppUrl();
        showCloudStatus("⏳ Đang lưu thay đổi Hội đồng lên Cloud...", "info");
        try {
            const res = await saveCloudHoiDong(gasUrl, sessionState.sessionId, sessionState.pin, sessionState.otp);
            if (res.ok) {
                showCloudStatus(`✅ ${res.message}`, "success");
            } else {
                showCloudStatus(`❌ ${res.message}`, "danger");
            }
        } catch (err) {
            showCloudStatus(`❌ Lỗi kết nối GAS: ${err.message}`, "danger");
        }
    });

    document.getElementById('btn-cloud-toggle-lock')?.addEventListener('click', async () => {
        const targetLockState = !sessionState.isHDLocked;
        const confirmMsg = targetLockState
            ? "Xác nhận CHỐT & KHÓA vùng Hội đồng? Sau khi khóa, không ai sửa được lịch/thành viên HĐ và các Thư ký sẽ bắt đầu Nhập điểm được."
            : "Mở khóa vùng Hội đồng để chỉnh sửa? Lưu ý: Trong lúc mở khóa, các Thư ký sẽ tạm thời không nhập điểm hay xuất file được.";
        if (!confirm(confirmMsg)) return;

        const gasUrl = getGasWebAppUrl();
        const bannerLockBtn = document.getElementById('btn-banner-toggle-lock');
        if (bannerLockBtn) bannerLockBtn.disabled = true;
        showCloudStatus("⏳ Đang cập nhật trạng thái khóa Hội đồng...", "info");

        try {
            if (targetLockState) {
                const saveRes = await saveCloudHoiDong(gasUrl, sessionState.sessionId, sessionState.pin, sessionState.otp);
                if (!saveRes.ok) {
                    showCloudStatus(`❌ Không thể lưu dữ liệu trước khi khóa: ${saveRes.message}`, "danger");
                    return;
                }
            }

            const res = await toggleCloudHDLock(gasUrl, sessionState.sessionId, sessionState.pin, sessionState.otp, targetLockState);
            if (res.ok) {
                sessionState.isHDLocked = Boolean(res.isHDLocked);
                applySessionPermissionsToUI();
                syncCloudModalButtons();
                showCloudStatus(res.message, "success");
            } else {
                showCloudStatus(`❌ ${res.message}`, "danger");
            }
        } catch (err) {
            showCloudStatus(`❌ Lỗi kết nối GAS: ${err.message}`, "danger");
        } finally {
            if (bannerLockBtn) bannerLockBtn.disabled = false;
        }
    });

    document.getElementById('btn-cloud-disconnect')?.addEventListener('click', () => {
        disconnectCloudSession();
        syncCloudModalButtons();
        showCloudStatus("🔌 Đã ngắt kết nối phiên Cloud. Bạn đang làm việc độc lập cục bộ.", "secondary");
    });

    document.addEventListener('click', (e) => {
        if (e.target.closest('#btn-banner-toggle-lock')) {
            document.getElementById('btn-cloud-toggle-lock')?.click();
        }
    });

    setupAutoSaveAndLockOnClose((autoSaveText) => {
        const banner = document.getElementById('session-role-banner');
        if (!banner || !sessionState.isCloudConnected) return;
        let statusTag = banner.querySelector('.autosave-indicator');
        if (!statusTag) {
            statusTag = document.createElement('span');
            statusTag.className = 'autosave-indicator fw-bold text-primary';
            banner.appendChild(statusTag);
        }
        statusTag.textContent = autoSaveText;
    });
}

function bindEvents() {
    bindMobileTabEvents();
    bindExcelAndHoiDongEvents();
    bindScoreAndExportEvents();
    bindTipsAndScoringTemplatesEvents();
    bindAdminEvents();
    bindCloudSessionEvents();
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
            const parts = rawDate.split('-');
            if (parts.length === 3) {
                [yPart, mPart, dPart] = parts;
                ngay = `${dPart}/${mPart}/${yPart}`;
            }
        } else if (rawDate.includes('/')) {
            const parts = rawDate.split('/');
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

function normalizeTeacherMatchKey(maGV, hoTen) {
    const cleanId = String(maGV || '').trim().toUpperCase();
    if (cleanId && !cleanId.startsWith('GV_')) {
        return `ID:${cleanId}`;
    }
    const cleanName = String(hoTen || '').toLowerCase()
        .replace(/gs\.?|pgs\.?|ts\.?|ths\.?|gvc\.?|gv\.?|ks\.?|cn\.?|bs\.?|prof\.?|dr\.?/g, ' ')
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd')
        .replace(/[^a-z0-9]/g, '');
    return cleanName ? `NAME:${cleanName}` : '';
}

function buildLecturerRoleStats() {
    const statsMap = new Map();

    const ensureTeacher = ({ maGV = '', hoTen = '', donVi = '' }) => {
        const key = normalizeTeacherMatchKey(maGV, hoTen);
        if (!key) return null;

        if (!statsMap.has(key)) {
            const validId = String(maGV || '').trim();
            statsMap.set(key, {
                maGV: (validId && !validId.toUpperCase().startsWith('GV_')) ? validId : '',
                hoTen: String(hoTen || '').trim(),
                donVi: String(donVi || '').trim(),
                huongDan: 0,
                chuTich: 0,
                phanBien: 0,
                thuKy: 0
            });
        } else {
            const item = statsMap.get(key);
            const validId = String(maGV || '').trim();
            if (!item.maGV && validId && !validId.toUpperCase().startsWith('GV_')) item.maGV = validId;
            if ((!item.hoTen || String(hoTen || '').trim().length > item.hoTen.length) && hoTen) {
                item.hoTen = String(hoTen).trim();
            }
            if (!item.donVi && donVi) item.donVi = String(donVi).trim();
        }
        return statsMap.get(key);
    };

    // 1. Nạp danh sách từ Kho Giảng viên (để không sót GV có 0 lượt)
    document.querySelectorAll('#pool-gv .card-item').forEach(el => {
        ensureTeacher({
            maGV: el.getAttribute('data-id') || '',
            hoTen: el.getAttribute('data-name') || '',
            donVi: el.getAttribute('data-donvi') || ''
        });
    });

    // 2. Đếm số lượt Hướng dẫn từ tất cả SV (cả trong Kho chờ lẫn đã xếp vào Hội đồng)
    document.querySelectorAll('#pool-sv .card-item, .dropzone-sv .card-item').forEach(svEl => {
        const maGVHD = svEl.getAttribute('data-magvhd') || '';
        const gvhd = svEl.getAttribute('data-gvhd') || '';
        const rec = ensureTeacher({ maGV: maGVHD, hoTen: gvhd });
        if (rec) rec.huongDan += 1;
    });

    // 3. Đếm số lượt Chủ tịch, Phản biện, Thư ký theo từng Hội đồng đã lập
    document.querySelectorAll('.box-hoidong').forEach(box => {
        const hasStudentsInHD = box.querySelectorAll('.dropzone-sv .card-item').length > 0;
        if (!hasStudentsInHD) return;

        const roleMappings = [
            { role: 'Chủ tịch', field: 'chuTich' },
            { role: 'Phản biện', field: 'phanBien' },
            { role: 'Thư ký', field: 'thuKy' }
        ];

        roleMappings.forEach(({ role, field }) => {
            const gvEl = box.querySelector(`.dropzone-gv[data-role="${role}"] .card-item`);
            if (!gvEl) return;
            const rec = ensureTeacher({
                maGV: gvEl.getAttribute('data-id') || '',
                hoTen: gvEl.getAttribute('data-name') || '',
                donVi: gvEl.getAttribute('data-donvi') || ''
            });
            if (rec) rec[field] += 1;
        });
    });

    return Array.from(statsMap.values())
        .map(item => ({
            ...item,
            tongLuot: item.huongDan + item.chuTich + item.phanBien + item.thuKy
        }))
        .sort((a, b) => b.tongLuot - a.tongLuot || a.hoTen.localeCompare(b.hoTen, 'vi'));
}

document.addEventListener("DOMContentLoaded", async () => {
    try {
        await Promise.all([
            loadHtmlPartials(),
            loadTemplatesSilently(),
            loadDanhSachNganhSilently()
        ]);
    } catch (e) {
        console.error("Lỗi tải dữ liệu khởi tạo:", e);
    }

    initKhoaNganhDropdowns();
    initGlobalSortables();
    bindEvents();
});
