// js/hoidong-manager.js
import { validateGVHDConflictWithHD, validateTeacherIsGVHDOfAnyStudent, checkScheduleConflict, isSameTeacher } from './validators.js';
import { getAdminConfig } from './admin-config.js';
import { buildMembersSignature, getSessionKey } from './excel-reader.js';

let hoidongCount = 0;

function findPoolSVItemById(maSV) {
    const target = String(maSV || '').trim().toUpperCase();
    if (!target) return null;
    const items = document.querySelectorAll('#pool-sv .card-item');
    return Array.from(items).find(el => String(el.getAttribute('data-id') || '').trim().toUpperCase() === target) || null;
}

function escapeHtmlAttr(val) {
    return String(val ?? '')
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

/**
 * Quét toàn bộ các box Hội đồng hiện có trên màn hình để lấy số HĐ kế tiếp (max + 1),
 * đảm bảo tạo mới không bao giờ bị trùng số hoặc nhảy số vô lý sau khi xóa/gộp.
 */
export function getNextAvailableHoiDongNumber() {
    let maxNum = 0;
    document.querySelectorAll('#workspace-hd .hd-title').forEach(inp => {
        const m = String(inp.value || '').match(/\d+/);
        if (m) {
            const n = parseInt(m[0], 10);
            if (!isNaN(n) && n > maxNum) maxNum = n;
        }
    });
    return maxNum + 1;
}

export function createGVCardInRoleZone({ maGV = '', hoTenGV = '', donVi = 'Đại học Duy Tân' }) {
    const card = document.createElement('div');
    card.className = 'card-item d-flex align-items-start justify-content-between gap-1 bg-white border border-info p-1 rounded shadow-sm';
    card.style.userSelect = 'none';
    card.style.webkitUserSelect = 'none';
    card.setAttribute('draggable', 'false');
    card.setAttribute('data-id', maGV);
    card.setAttribute('data-name', hoTenGV);
    card.setAttribute('data-donvi', donVi || 'Đại học Duy Tân');

    card.innerHTML = `
        <div class="overflow-hidden" style="min-width: 0; flex: 1;">
            <div class="fw-bold text-dark small text-truncate" title="${escapeHtmlAttr(hoTenGV)}">${escapeHtmlAttr(hoTenGV)}</div>
            <div class="text-muted text-truncate" style="font-size: 0.72rem;" title="${escapeHtmlAttr(donVi || 'Đại học Duy Tân')} - Mã: ${escapeHtmlAttr(maGV)}">
                Mã: ${escapeHtmlAttr(maGV)}
            </div>
        </div>
        <button type="button"
                class="btn btn-sm btn-outline-danger py-0 px-1 fw-bold flex-shrink-0 btn-remove-gv-role"
                style="line-height: 1.2; font-size: 0.75rem;"
                title="Xóa giảng viên khỏi chức danh này">✕</button>
    `;
    return card;
}

// Hàm đồng bộ nút ✕ trên thẻ SV theo vị trí (trong HĐ -> trả về Kho chờ; ngoài Kho chờ -> xóa hẳn)
export function syncSVCardActionButton(svCard, isInHoiDong) {
    if (!svCard) return;
    let btn = svCard.querySelector('.btn-remove-card, .btn-remove-sv-from-hd');
    if (!btn) {
        svCard.classList.add('position-relative', 'pe-4');
        btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1';
        btn.textContent = '✕';
        svCard.prepend(btn);
    }

    if (isInHoiDong) {
        btn.classList.remove('btn-remove-card');
        btn.classList.add('btn-remove-sv-from-hd');
        btn.setAttribute('title', 'Đưa sinh viên này khỏi Hội đồng về lại Kho chờ');
    } else {
        btn.classList.remove('btn-remove-sv-from-hd');
        btn.classList.add('btn-remove-card');
        btn.setAttribute('title', 'Xóa sinh viên này');
    }
}

// Export hàm 1: Khởi tạo Sortable toàn cục & Lắng nghe sự kiện click nút "+ Thêm SV", "+ Gán"
export function initGlobalSortables() {
    const poolSV = document.getElementById('pool-sv');
    if (poolSV) {
        Sortable.create(poolSV, {
            group: 'sinhvien',
            animation: 150,
            onAdd: function (evt) {
                syncSVCardActionButton(evt.item, false);
            }
        });
    }

    initAddSVModalEvents();
    initAssignGVRoleEvents();
}

// Export hàm 2: Tạo Box Hội Đồng mới
export function createHoiDongBox(titleText = null, initDate = '', initTime = '', initRoom = '', initLocation = '') {
    const nextSuggestedNum = getNextAvailableHoiDongNumber();
    hoidongCount = nextSuggestedNum;

    const cleanTitle = (titleText && String(titleText).trim())
        ? String(titleText).trim()
        : `Hội Đồng ${nextSuggestedNum}`;
    const hdTitle = escapeHtmlAttr(cleanTitle);
    const boxId = `hd-box-${Date.now()}-${Math.floor(Math.random() * 1000)}`;

    let formattedTime = initTime || '';
    if (formattedTime.includes('h')) {
        const parts = formattedTime.split('h');
        formattedTime = `${parts[0].padStart(2, '0')}:${(parts[1] || '00').padStart(2, '0')}`;
    }

    const locations = getAdminConfig().campuses || [];
    const locationOptions = [
        `<option value="" ${!initLocation ? 'selected' : ''}>-- Chọn cơ sở --</option>`,
        ...locations.map(loc =>
            `<option value="${escapeHtmlAttr(loc)}" ${loc === initLocation ? 'selected' : ''}>${escapeHtmlAttr(loc)}</option>`
        )
    ].join('');

    const renderRoleSlot = (roleLabel, textColorClass) => `
        <div class="mb-2">
            <div class="d-flex justify-content-between align-items-center mb-1 gap-1">
                <small class="${textColorClass} fw-bold">${roleLabel}:</small>
                <button type="button"
                        class="btn btn-outline-primary btn-sm py-0 px-1 btn-open-assign-gv"
                        style="font-size: 0.72rem; line-height: 1.3;"
                        data-hd-id="${boxId}"
                        data-role="${roleLabel}"
                        title="Chọn/Thay thế Giảng viên vào vị trí ${roleLabel}">+ Gán</button>
            </div>
            <div class="dropzone-gv border rounded p-1 bg-light" data-role="${roleLabel}" style="min-height: 40px;"></div>
        </div>
    `;

    const boxHTML = `
        <div class="box-hoidong p-2 mb-3 border rounded bg-white shadow-sm" id="${boxId}">
            <div class="d-flex gap-2 mb-2 bg-light p-2 rounded align-items-center flex-wrap">
                <input type="text" class="form-control form-control-sm fw-bold hd-title" value="${hdTitle}" placeholder="Số/Tên HĐ" style="width: 120px;">
                <input type="date" class="form-control form-control-sm hd-date" value="${escapeHtmlAttr(initDate)}" style="width: 125px;" title="Ngày bảo vệ">
                <input type="time" class="form-control form-control-sm hd-time" value="${escapeHtmlAttr(formattedTime)}" style="width: 110px;" title="Giờ:Phút bảo vệ">
                <input type="text" class="form-control form-control-sm hd-room" value="${escapeHtmlAttr(initRoom)}" placeholder="Phòng (111)" style="width: 105px;">
                <select class="form-select form-select-sm hd-location" style="width: 155px;">
                    ${locationOptions}
                </select>
                <button type="button" class="btn btn-sm btn-outline-danger fw-bold ms-auto btn-delete-hd" title="Xóa Hội đồng này">✕ Xóa</button>
            </div>
            <div class="row">
                <div class="col-5 col-md-4 border-end">
                    ${renderRoleSlot('Chủ tịch', 'text-danger')}
                    ${renderRoleSlot('Phản biện', 'text-success')}
                    ${renderRoleSlot('Thư ký', 'text-primary')}
                </div>
                <div class="col-7 col-md-8">
                    <div class="d-flex justify-content-between align-items-center mb-1">
                        <small class="fw-bold">Sinh viên bảo vệ:</small>
                        <button type="button" class="btn btn-outline-success btn-sm py-0 px-2 btn-open-add-sv" data-hd-id="${boxId}">
                            + Thêm SV
                        </button>
                    </div>
                    <div class="dropzone-sv p-2 bg-light border rounded" style="min-height: 140px;"></div>
                </div>
            </div>
        </div>
    `;

    const workspace = document.getElementById('workspace-hd');
    if (!workspace) return null;
    workspace.insertAdjacentHTML('beforeend', boxHTML);
    const boxEl = document.getElementById(boxId);

    boxEl?.querySelector('.btn-delete-hd')?.addEventListener('click', () => {
        boxEl.remove();
        hoidongCount = getNextAvailableHoiDongNumber() - 1;
    });

    attachSortableToHoiDong(boxId);
    return boxId;
}

// Hàm hậu xử lý:
// 1. Gộp các ca bảo vệ trùng (Thành viên + Chức danh, Ngày, Buổi, Phòng)
// 2. Đánh số Hội đồng nhất quán:
//    - Nếu preserveExplicitNumbers = true: giữ nguyên số HĐ gốc từ file Excel (cột A), đồng bộ tên "Hội Đồng <Số>"
//    - Nếu preserveExplicitNumbers = false: đánh số tuần tự 1 -> N (cùng bộ Thành viên + Chức danh thì dùng chung 1 số HĐ)
export function finalizeAndRenumberHoiDongs(preserveExplicitNumbers = false) {
    const workspace = document.getElementById('workspace-hd');
    if (!workspace) return [];

    const mergeLogs = [];
    const boxes = Array.from(workspace.querySelectorAll('.box-hoidong'));
    const keptBoxes = [];

    const readBoxMembers = (box) => {
        const members = [];
        ['Chủ tịch', 'Phản biện', 'Thư ký'].forEach(role => {
            const card = box.querySelector(`.dropzone-gv[data-role="${role}"] .card-item`);
            if (card) {
                members.push({
                    chucDanh: role,
                    maGV: card.getAttribute('data-id') || '',
                    hoTenGV: card.getAttribute('data-name') || ''
                });
            }
        });
        return members;
    };

    const extractExplicitNum = (titleStr) => {
        const cleaned = String(titleStr || '').trim();
        const m = cleaned.match(/(?:hội\s*đồng|hoi\s*dong|hđ|hd)?\s*(?:số|so|#|:|-|\.)?\s*(\d+[a-zA-Z0-9\-]*)/i);
        return m ? m[1].trim() : '';
    };

    // Bước A: Gộp ca bảo vệ trùng (Thành viên + Chức danh, Ngày, Buổi, Phòng)
    boxes.forEach((box, idx) => {
        const oldTitle = box.querySelector('.hd-title')?.value.trim() || `Hội Đồng ${idx + 1}`;
        const explicitNum = extractExplicitNum(oldTitle);
        const memSig = buildMembersSignature(readBoxMembers(box));
        const dateKey = (box.querySelector('.hd-date')?.value || '').trim();
        const timeVal = (box.querySelector('.hd-time')?.value || '').trim();
        const sessionKey = getSessionKey(timeVal);
        const roomKey = (box.querySelector('.hd-room')?.value || '').trim().toUpperCase();

        const sameSlotTarget = keptBoxes.find(target =>
            memSig &&
            target.memSig === memSig &&
            target.dateKey === dateKey &&
            target.sessionKey === sessionKey &&
            target.roomKey === roomKey
        );

        if (sameSlotTarget) {
            const targetSVZone = sameSlotTarget.box.querySelector('.dropzone-sv');
            const currSVCards = Array.from(box.querySelectorAll('.dropzone-sv .card-item'));
            currSVCards.forEach(svCard => {
                syncSVCardActionButton(svCard, true);
                targetSVZone?.appendChild(svCard);
            });

            if (!sameSlotTarget.box.querySelector('.hd-time')?.value && timeVal) {
                sameSlotTarget.box.querySelector('.hd-time').value = timeVal;
            }
            const currLoc = box.querySelector('.hd-location')?.value || '';
            const targetLocEl = sameSlotTarget.box.querySelector('.hd-location');
            if (targetLocEl && !targetLocEl.value && currLoc) {
                targetLocEl.value = currLoc;
            }

            box.remove();
            mergeLogs.push(`• Gộp ca bảo vệ của khối #${idx + 1} vào ${sameSlotTarget.oldTitle} (trùng Thành viên, Ngày, Buổi & Phòng).`);
            return;
        }

        keptBoxes.push({
            box,
            oldTitle,
            explicitNum,
            memSig,
            dateKey,
            sessionKey,
            roomKey
        });
    });

    // Bước B: Đánh số Hội đồng chuẩn xác
    const sigToNumber = new Map();
    let nextNum = 1;

    keptBoxes.forEach((entry, idx) => {
        let assignedNum;

        if (preserveExplicitNumbers && entry.explicitNum) {
            assignedNum = entry.explicitNum;
            if (entry.memSig && !sigToNumber.has(entry.memSig)) {
                sigToNumber.set(entry.memSig, assignedNum);
            }
        } else if (entry.memSig && sigToNumber.has(entry.memSig)) {
            assignedNum = sigToNumber.get(entry.memSig);
            mergeLogs.push(`• Khối #${idx + 1} dùng chung thành viên với Hội Đồng ${assignedNum} (khác buổi/phòng) → Đánh số Hội Đồng ${assignedNum}.`);
        } else {
            assignedNum = String(nextNum++);
            if (entry.memSig) sigToNumber.set(entry.memSig, assignedNum);
        }

        const titleInput = entry.box.querySelector('.hd-title');
        if (titleInput) {
            titleInput.value = `Hội Đồng ${assignedNum}`;
        }
    });

    hoidongCount = getNextAvailableHoiDongNumber() - 1;
    return mergeLogs;
}

// Export hàm 3: Render Hội đồng từ Excel, rút SV từ pool, rồi mới đánh lại số HĐ chính xác
export function renderHoidongFromExcel(parsedHoidongs) {
    const workspace = document.getElementById('workspace-hd');
    if (!workspace) return { mergeLogs: [], conflictLogs: [], duplicateSVLogs: [] };
    workspace.innerHTML = '';
    hoidongCount = 0;

    const safeParsed = parsedHoidongs && typeof parsedHoidongs === 'object' ? parsedHoidongs : {};
    const preserveExplicitNumbers = Boolean(safeParsed.__preserveExplicitNumbers);
    const conflictLogs = [];
    const duplicateSVLogs = [];

    // Bước tiền rà soát: Thống kê mỗi Mã SV xuất hiện ở những khối Hội đồng nào
    const svAppearances = new Map();
    Object.keys(safeParsed).forEach(hdKey => {
        if (hdKey.startsWith('__')) return;
        const hd = safeParsed[hdKey];
        if (!hd || typeof hd !== 'object') return;
        const hdLabel = `Khối #${hd.soHD || hdKey}${hd.ngayBV ? ` (${hd.gioBV ? hd.gioBV + ' ' : ''}${hd.ngayBV})` : ''}`;

        (Array.isArray(hd.sinhViens) ? hd.sinhViens : []).forEach(sv => {
            const ma = String(sv.maSV || '').trim().toUpperCase();
            if (!ma) return;
            if (!svAppearances.has(ma)) {
                svAppearances.set(ma, { hoTen: sv.hoTenSV || sv.hoTen || ma, hoidongs: [] });
            }
            svAppearances.get(ma).hoidongs.push(hdLabel);
        });
    });

    svAppearances.forEach((info, ma) => {
        if (info.hoidongs.length > 1) {
            duplicateSVLogs.push(
                `• SV ${info.hoTen} (${ma}) bị lặp lại ${info.hoidongs.length} lần tại: ${info.hoidongs.join(' ; ')} → Đã giữ ở khối đầu tiên và loại bỏ bản lặp phía sau!`
            );
        }
    });

    const renderedSVSet = new Set();

    Object.keys(safeParsed).forEach(hdKey => {
        if (hdKey.startsWith('__')) return;
        const hd = safeParsed[hdKey];
        if (!hd || typeof hd !== 'object') return;

        const thanhVien = Array.isArray(hd.thanhVien) ? hd.thanhVien : [];
        const sinhViens = Array.isArray(hd.sinhViens) ? hd.sinhViens : [];

        const formattedDate = formatDateToInput(hd.ngayBV);
        const boxId = createHoiDongBox(
            `Hội Đồng ${hd.soHD}`,
            formattedDate,
            hd.gioBV || '',
            hd.phongBV || '',
            hd.diaDiemBV || ''
        );

        const boxEl = boxId ? document.getElementById(boxId) : null;
        if (!boxEl) return;

        // 1. Render Giảng viên vào đúng chức danh (mỗi chức danh chỉ có 1 GV + icon ✕)
        thanhVien.forEach(gv => {
            let roleName = "";
            const chucDanhLower = String(gv.chucDanh || '').toLowerCase();
            if (chucDanhLower.includes("chủ tịch")) roleName = "Chủ tịch";
            else if (chucDanhLower.includes("phản biện")) roleName = "Phản biện";
            else if (chucDanhLower.includes("thư ký")) roleName = "Thư ký";

            if (roleName) {
                const zone = boxEl.querySelector(`.dropzone-gv[data-role="${roleName}"]`);
                if (zone) {
                    zone.innerHTML = '';
                    zone.appendChild(createGVCardInRoleZone({
                        maGV: gv.maGV,
                        hoTenGV: gv.hoTenGV,
                        donVi: gv.donVi || 'Đại học Duy Tân'
                    }));
                }
            }
        });

        // 2. Render Sinh viên, rút từ #pool-sv & RÀ SOÁT TRÙNG GVHD
        const svDropzone = boxEl.querySelector('.dropzone-sv');
        const poolSV = document.getElementById('pool-sv');

        sinhViens.forEach(sv => {
            const maSVNorm = String(sv.maSV || '').trim().toUpperCase();
            if (!maSVNorm) return;

            if (renderedSVSet.has(maSVNorm)) return;
            renderedSVSet.add(maSVNorm);

            const poolItem = findPoolSVItemById(sv.maSV);

            const hoTen = poolItem ? poolItem.getAttribute('data-name') : (sv.hoTenSV || sv.hoTen || '');
            const lop = poolItem ? poolItem.getAttribute('data-lop') : (sv.lop || '');
            const deTai = poolItem ? poolItem.getAttribute('data-detai') : (sv.deTai || '');
            const maGVHD = poolItem ? (poolItem.getAttribute('data-magvhd') || '') : (sv.maGVHD || '');
            const gvhd = poolItem ? (poolItem.getAttribute('data-gvhd') || '') : (sv.gvhd || '');
            const nganh = poolItem ? (poolItem.getAttribute('data-nganh') || '') : (sv.nganh || '');
            const maGVDisplay = (maGVHD && !/^AUTO_GV_/i.test(maGVHD)) ? ` [${maGVHD}]` : '';

            let violatedMember = null;
            for (const member of thanhVien) {
                if (isSameTeacher({ id1: maGVHD, name1: gvhd }, { id2: member.maGV, name2: member.hoTenGV })) {
                    violatedMember = member;
                    break;
                }
            }

            if (violatedMember) {
                if (!poolItem && poolSV) {
                    const returnHTML = `
                        <div class="card-item position-relative mb-2 p-2 pe-4 border border-danger rounded bg-white shadow-sm" 
                             data-id="${escapeHtmlAttr(sv.maSV)}" data-name="${escapeHtmlAttr(hoTen)}" data-lop="${escapeHtmlAttr(lop)}" data-detai="${escapeHtmlAttr(deTai)}" data-magvhd="${escapeHtmlAttr(maGVHD)}" data-gvhd="${escapeHtmlAttr(gvhd)}" data-nganh="${escapeHtmlAttr(nganh)}">
                            <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card" title="Xóa sinh viên này">✕</button>
                            <div class="fw-bold text-primary">${escapeHtmlAttr(hoTen)} (${escapeHtmlAttr(sv.maSV)})</div>
                            <div class="text-muted small">Lớp: ${escapeHtmlAttr(lop)} | GVHD: <span class="fw-bold text-danger">${escapeHtmlAttr(gvhd || 'Chưa rõ')}${escapeHtmlAttr(maGVDisplay)}</span></div>
                            <div class="text-dark small text-truncate" title="${escapeHtmlAttr(deTai)}">Đề tài: ${escapeHtmlAttr(deTai)}</div>
                        </div>`;
                    poolSV.insertAdjacentHTML('afterbegin', returnHTML);
                } else if (poolItem) {
                    poolItem.classList.add('border-danger');
                    syncSVCardActionButton(poolItem, false);
                    poolSV?.prepend(poolItem);
                }

                conflictLogs.push(
                    `• SV ${hoTen} (${sv.maSV}) - GVHD: ${gvhd || maGVHD} trùng với ${violatedMember.chucDanh} (${violatedMember.hoTenGV}) → Đã đẩy về Kho chờ, đề nghị xếp sang Hội đồng khác!`
                );
                return;
            }

            // Hợp lệ -> Đưa vào Hội đồng (có nút ✕ trả về Kho chờ) và rút khỏi Kho chờ (#pool-sv)
            const svHTML = `
                <div class="card-item position-relative mb-2 p-2 pe-4 border rounded bg-white shadow-sm" 
                     data-id="${escapeHtmlAttr(sv.maSV)}" data-name="${escapeHtmlAttr(hoTen)}" data-lop="${escapeHtmlAttr(lop)}" data-detai="${escapeHtmlAttr(deTai)}" data-magvhd="${escapeHtmlAttr(maGVHD)}" data-gvhd="${escapeHtmlAttr(gvhd)}" data-nganh="${escapeHtmlAttr(nganh)}">
                    <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-sv-from-hd" title="Đưa sinh viên này khỏi Hội đồng về lại Kho chờ">✕</button>
                    <div class="fw-bold text-primary">${escapeHtmlAttr(hoTen)} (${escapeHtmlAttr(sv.maSV)})</div>
                    <div class="text-muted small">Lớp: ${escapeHtmlAttr(lop)} | GVHD: <span class="fw-bold text-dark">${escapeHtmlAttr(gvhd || 'Chưa rõ')}${escapeHtmlAttr(maGVDisplay)}</span></div>
                    <div class="text-dark small text-truncate" title="${escapeHtmlAttr(deTai)}">Đề tài: ${escapeHtmlAttr(deTai)}</div>
                </div>`;

            svDropzone?.insertAdjacentHTML('beforeend', svHTML);

            if (poolItem) {
                poolItem.remove();
            }
        });
    });

    // SAU KHI RENDER XONG: Gộp ca trùng và đánh số HĐ nhất quán
    // SAU KHI RENDER XONG: Chỉ chạy gộp/đánh số tự động nếu chưa được chuẩn hóa trước đó
    const mergeLogs = safeParsed.__skipAutoMergeInRenderer
        ? []
        : finalizeAndRenumberHoiDongs(preserveExplicitNumbers);

    return { mergeLogs, conflictLogs, duplicateSVLogs };
}

// ----------------------------------------------------
// MODAL GÁN / THAY THẾ GIẢNG VIÊN VÀO CHỨC VỤ HỘI ĐỒNG (+ GÁN)
// ----------------------------------------------------
let targetBoxIdForAssignGV = null;
let targetRoleForAssignGV = null;

function ensureAssignGVModalExists() {
    let modalEl = document.getElementById('modalGanGiangVien');
    if (modalEl) return modalEl;

    const modalHTML = `
        <div class="modal fade" id="modalGanGiangVien" tabindex="-1" aria-hidden="true">
            <div class="modal-dialog modal-dialog-scrollable">
                <div class="modal-content">
                    <div class="modal-header bg-primary text-white py-2">
                        <h6 class="modal-title fw-bold mb-0" id="modalGanGVTitle">👨‍🏫 CHỌN GIẢNG VIÊN</h6>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Đóng"></button>
                    </div>
                    <div class="modal-body p-2">
                        <input type="text" id="search-gv-assign-modal" class="form-control form-control-sm mb-2" placeholder="🔍 Tìm nhanh theo Mã GV, Họ tên hoặc Đơn vị...">
                        <div id="list-gv-phu-hop" class="list-group"></div>
                    </div>
                    <div class="modal-footer py-2">
                        <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Đóng</button>
                    </div>
                </div>
            </div>
        </div>
    `;
    document.body.insertAdjacentHTML('beforeend', modalHTML);
    return document.getElementById('modalGanGiangVien');
}

function initAssignGVRoleEvents() {
    document.addEventListener('click', (e) => {
        // 1. Nút ✕ xóa GV khỏi chức danh trong HĐ
        const btnRemoveGVRole = e.target.closest('.btn-remove-gv-role');
        if (btnRemoveGVRole) {
            e.stopPropagation();
            if (window.cloudSession?.active && !window.cloudSession?.canEditStructure) {
                alert('⚠️ Phiên Cloud đang khóa cấu trúc Hội đồng hoặc bạn đang đăng nhập quyền Thư ký nên không thể sửa thành viên HĐ!');
                return;
            }
            const gvCard = btnRemoveGVRole.closest('.card-item');
            if (gvCard) gvCard.remove();
            return;
        }

        // 2. Nút "+ Gán" ở ô chức vụ GV trong HĐ
        const btnAssignGV = e.target.closest('.btn-open-assign-gv');
        if (btnAssignGV) {
            e.stopPropagation();
            if (window.cloudSession?.active && !window.cloudSession?.canEditStructure) {
                alert('⚠️ Phiên Cloud đang khóa cấu trúc Hội đồng hoặc bạn đang đăng nhập quyền Thư ký nên không thể sửa thành viên HĐ!');
                return;
            }
            targetBoxIdForAssignGV = btnAssignGV.getAttribute('data-hd-id');
            targetRoleForAssignGV = btnAssignGV.getAttribute('data-role');
            openAssignGVModal(targetBoxIdForAssignGV, targetRoleForAssignGV);
        }
    });

    // 3. Lọc nhanh danh sách GV khi gõ tìm kiếm trong Modal
    document.addEventListener('input', (e) => {
        if (e.target && e.target.id === 'search-gv-assign-modal') {
            renderAssignGVList(e.target.value);
        }
    });
}

function openAssignGVModal(boxId, roleName) {
    const box = document.getElementById(boxId);
    if (!box || !roleName) return;

    const modalEl = ensureAssignGVModalExists();
    const hdTitle = box.querySelector('.hd-title')?.value || 'Hội Đồng';
    const titleEl = document.getElementById('modalGanGVTitle');
    if (titleEl) {
        titleEl.innerText = `👨‍🏫 GÁN ${roleName.toUpperCase()} — ${hdTitle.toUpperCase()}`;
    }

    const searchInput = document.getElementById('search-gv-assign-modal');
    if (searchInput) searchInput.value = '';

    renderAssignGVList('');

    if (modalEl && window.bootstrap) {
        bootstrap.Modal.getOrCreateInstance(modalEl).show();
    }
}

function renderAssignGVList(keyword = '') {
    const container = document.getElementById('list-gv-phu-hop');
    const box = targetBoxIdForAssignGV ? document.getElementById(targetBoxIdForAssignGV) : null;
    if (!container || !box || !targetRoleForAssignGV) return;

    const poolGVItems = Array.from(document.querySelectorAll('#pool-gv .card-item'));
    if (poolGVItems.length === 0) {
        container.innerHTML = '<div class="text-center text-muted py-3">Kho Giảng viên đang trống!</div>';
        return;
    }

    const q = keyword.trim().toLowerCase();
    const targetZone = box.querySelector(`.dropzone-gv[data-role="${targetRoleForAssignGV}"]`);
    const currentCardInTarget = targetZone?.querySelector('.card-item');
    const currentIdInTarget = currentCardInTarget?.getAttribute('data-id') || '';
    const currentNameInTarget = currentCardInTarget?.getAttribute('data-name') || '';

    const otherRoleCards = Array.from(box.querySelectorAll('.dropzone-gv .card-item'))
        .filter(node => node.parentElement !== targetZone);

    container.innerHTML = '';
    let matchedCount = 0;

    poolGVItems.forEach(gvItem => {
        const maGV = (gvItem.getAttribute('data-id') || '').trim();
        const hoTen = (gvItem.getAttribute('data-name') || '').trim();
        const donVi = (gvItem.getAttribute('data-donvi') || 'Đại học Duy Tân').trim();

        if (q && !maGV.toLowerCase().includes(q) && !hoTen.toLowerCase().includes(q) && !donVi.toLowerCase().includes(q)) {
            return;
        }
        matchedCount++;

        const isCurrentHolder = currentCardInTarget && isSameTeacher(
            { id1: maGV, name1: hoTen },
            { id2: currentIdInTarget, name2: currentNameInTarget }
        );

        const isAlreadyOtherRole = otherRoleCards.some(node => isSameTeacher(
            { id1: maGV, name1: hoTen },
            { id2: node.getAttribute('data-id') || '', name2: node.getAttribute('data-name') || '' }
        ));
        const violatedSVName = validateTeacherIsGVHDOfAnyStudent(hoTen, box, maGV);
        const conflictHDTitle = checkScheduleConflict(maGV, hoTen, box);

        let blockReason = '';
        if (isAlreadyOtherRole) blockReason = 'Đã giữ vai trò khác trong HĐ này';
        else if (violatedSVName) blockReason = `Là GVHD của SV ${violatedSVName}`;
        else if (conflictHDTitle) blockReason = `Trùng lịch tại ${conflictHDTitle}`;

        const btnItem = document.createElement('button');
        btnItem.type = 'button';
        btnItem.className = `list-group-item list-group-item-action d-flex justify-content-between align-items-center gap-2 ${blockReason ? 'list-group-item-light text-muted' : ''}`;
        if (blockReason) btnItem.disabled = true;

        btnItem.innerHTML = `
            <div class="text-start">
                <div class="fw-bold ${blockReason ? 'text-muted' : 'text-dark'}">
                    ${escapeHtmlAttr(hoTen)} <span class="badge bg-secondary font-monospace">${escapeHtmlAttr(maGV)}</span>
                </div>
                <div class="small text-muted">${escapeHtmlAttr(donVi)}</div>
                ${blockReason ? `<div class="small text-danger fw-semibold">⛔ ${escapeHtmlAttr(blockReason)}</div>` : ''}
            </div>
            <span class="badge ${isCurrentHolder ? 'bg-success' : (blockReason ? 'bg-secondary' : 'bg-primary')}">
                ${isCurrentHolder ? 'Đang chọn' : (blockReason ? 'Không khả dụng' : 'Chọn')}
            </span>
        `;

        if (!blockReason) {
            btnItem.addEventListener('click', () => {
                if (targetZone) {
                    targetZone.innerHTML = '';
                    targetZone.appendChild(createGVCardInRoleZone({
                        maGV,
                        hoTenGV: hoTen,
                        donVi
                    }));
                }
                const modalEl = document.getElementById('modalGanGiangVien');
                const modalInstance = modalEl && window.bootstrap ? bootstrap.Modal.getInstance(modalEl) : null;
                modalInstance?.hide();
            });
        }

        container.appendChild(btnItem);
    });

    if (matchedCount === 0) {
        container.innerHTML = '<div class="text-center text-muted py-3">Không tìm thấy Giảng viên phù hợp!</div>';
    }
}

// ----------------------------------------------------
// THUẬT TOÁN MỞ MODAL & LỌC SINH VIÊN PHÙ HỢP (CLICK-TO-ASSIGN)
// ----------------------------------------------------
let targetBoxIdForAddSV = null;

function initAddSVModalEvents() {
    document.addEventListener('click', (e) => {
        const btnAdd = e.target.closest('.btn-open-add-sv');
        if (btnAdd) {
            targetBoxIdForAddSV = btnAdd.getAttribute('data-hd-id');
            openSelectSVModal(targetBoxIdForAddSV);
        }
    });

    document.getElementById('btn-confirm-gan-sv')?.addEventListener('click', () => {
        if (!targetBoxIdForAddSV) return;
        const box = document.getElementById(targetBoxIdForAddSV);
        const svDropzone = box?.querySelector('.dropzone-sv');
        if (!svDropzone) return;

        const checkedInputs = document.querySelectorAll('#list-sv-phu-hop .chk-select-sv:checked');
        checkedInputs.forEach(chk => {
            const poolItem = findPoolSVItemById(chk.value);
            if (poolItem) {
                syncSVCardActionButton(poolItem, true);
                svDropzone.appendChild(poolItem);
            }
        });

        const modalEl = document.getElementById('modalGanSinhVien');
        const modalInstance = modalEl && window.bootstrap ? bootstrap.Modal.getInstance(modalEl) : null;
        modalInstance?.hide();
    });
}

function openSelectSVModal(boxId) {
    const box = document.getElementById(boxId);
    if (!box) return;

    const hdTitle = box.querySelector('.hd-title')?.value || 'Hội Đồng';
    const titleEl = document.getElementById('modalGanSVTitle');
    if (titleEl) titleEl.innerText = `📋 CHỌN SINH VIÊN VÀO ${hdTitle.toUpperCase()}`;

    const poolSVItems = Array.from(document.querySelectorAll('#pool-sv .card-item'));
    const container = document.getElementById('list-sv-phu-hop');
    if (!container) return;
    container.innerHTML = '';

    if (poolSVItems.length === 0) {
        container.innerHTML = '<div class="text-center text-muted py-3">Không có sinh viên nào đang chờ phân công!</div>';
    } else {
        let countAvailable = 0;
        poolSVItems.forEach(svItem => {
            const maSV = svItem.getAttribute('data-id');
            const hoTen = svItem.getAttribute('data-name');
            const lop = svItem.getAttribute('data-lop');
            const deTai = svItem.getAttribute('data-detai');
            const maGVHD = (svItem.getAttribute('data-magvhd') || '').trim();
            const gvhd = (svItem.getAttribute('data-gvhd') || '').trim();

            const isConflicted = Boolean(validateGVHDConflictWithHD(gvhd, box, maGVHD));

            if (!isConflicted) {
                countAvailable++;
                const maGVDisplay = maGVHD ? ` [${maGVHD}]` : '';
                const itemHTML = `
                    <label class="list-group-item d-flex align-items-center gap-2 cursor-pointer">
                        <input class="form-check-input flex-shrink-0 chk-select-sv" type="checkbox" value="${escapeHtmlAttr(maSV)}">
                        <div class="w-100">
                            <div class="fw-bold text-primary mb-0">${escapeHtmlAttr(hoTen)} (${escapeHtmlAttr(maSV)})</div>
                            <div class="small text-muted">Lớp: ${escapeHtmlAttr(lop)} | GVHD: <span class="fw-bold text-dark">${escapeHtmlAttr(gvhd || 'Chưa rõ')}${escapeHtmlAttr(maGVDisplay)}</span></div>
                            <div class="small text-dark text-truncate">Đề tài: ${escapeHtmlAttr(deTai)}</div>
                        </div>
                    </label>`;
                container.insertAdjacentHTML('beforeend', itemHTML);
            }
        });

        if (countAvailable === 0) {
            container.innerHTML = `
                <div class="text-center text-danger py-3 fw-bold">
                    ⛔ Tất cả SV trong kho chờ đều bị trùng GVHD với Thành viên của ${escapeHtmlAttr(hdTitle)}!
                </div>`;
        }
    }

    const modalEl = document.getElementById('modalGanSinhVien');
    if (modalEl && window.bootstrap) {
        const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
        modal.show();
    }
}

function formatDateToInput(dateStr) {
    if (!dateStr) return '';
    if (dateStr.includes('/')) {
        const p = dateStr.split('/');
        if (p.length === 3) return `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
    }
    return dateStr;
}

function attachSortableToHoiDong(boxId) {
    const box = document.getElementById(boxId);
    if (!box) return;

    // CHỈ gắn Sortable cho khu vực Sinh viên (.dropzone-sv).
    // Tuyệt đối KHÔNG gắn Sortable cho .dropzone-gv để:
    // 1. Thẻ SV không bao giờ lọt được vào ô Chức vụ GV.
    // 2. Thẻ GV trong HĐ không kích hoạt hành vi nhấn giữ / kéo thả trên điện thoại.
    const svZoneEl = box.querySelector('.dropzone-sv');
    if (svZoneEl) {
        Sortable.create(svZoneEl, {
            group: 'sinhvien',
            animation: 150,
            onAdd: function (evt) {
                const svCard = evt.item;
                const svName = svCard.getAttribute('data-name');
                const gvhdName = svCard.getAttribute('data-gvhd') || '';
                const maGVHD = svCard.getAttribute('data-magvhd') || '';

                const conflictedMember = validateGVHDConflictWithHD(gvhdName, box, maGVHD);
                if (conflictedMember) {
                    alert(`⛔ VI PHẠM NGUYÊN TẮC: GVHD (${gvhdName || maGVHD}) đang làm thành viên HĐ (${conflictedMember}). Không thể xếp sinh viên ${svName} vào HĐ này!`);
                    evt.from?.appendChild(svCard);
                } else {
                    syncSVCardActionButton(svCard, true);
                }
            }
        });
    }
}
