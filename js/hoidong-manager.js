// js/hoidong-manager.js
import { validateGVHDConflictWithHD, validateTeacherIsGVHDOfAnyStudent, checkScheduleConflict, isSameTeacher } from './validators.js';
import { getAdminConfig } from './admin-config.js';

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

// Export hàm 1: Khởi tạo Sortable toàn cục & Lắng nghe sự kiện click nút "+ Thêm SV"
export function initGlobalSortables() {
    const poolSV = document.getElementById('pool-sv');
    if (poolSV) Sortable.create(poolSV, { group: 'sinhvien', animation: 150 });

    const poolGV = document.getElementById('pool-gv');
    if (poolGV) {
        Sortable.create(poolGV, {
            group: { name: 'giangvien', pull: 'clone', put: false },
            animation: 150, sort: false
        });
    }

    initAddSVModalEvents();
}

// Export hàm 2: Tạo Box Hội Đồng mới (Lấy danh sách Cơ sở động từ cấu hình Admin)
export function createHoiDongBox(titleText = null, initDate = '', initTime = '', initRoom = '', initLocation = '') {
    hoidongCount++;
    const hdTitle = escapeHtmlAttr(titleText || `HĐ ${hoidongCount}`);
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

    const boxHTML = `
        <div class="box-hoidong p-2 mb-3 border rounded bg-white shadow-sm" id="${boxId}">
            <div class="d-flex gap-2 mb-2 bg-light p-2 rounded align-items-center flex-wrap">
                <input type="text" class="form-control form-control-sm fw-bold hd-title" value="${hdTitle}" placeholder="Số/Tên HĐ" style="width: 110px;">
                <input type="date" class="form-control form-control-sm hd-date" value="${escapeHtmlAttr(initDate)}" style="width: 125px;" title="Ngày bảo vệ">
                <input type="time" class="form-control form-control-sm hd-time" value="${escapeHtmlAttr(formattedTime)}" style="width: 110px;" title="Giờ:Phút bảo vệ">
                <input type="text" class="form-control form-control-sm hd-room" value="${escapeHtmlAttr(initRoom)}" placeholder="Phòng (111)" style="width: 105px;">
                <select class="form-select form-select-sm hd-location" style="width: 155px;">
                    ${locationOptions}
                </select>
                <button type="button" class="btn btn-sm btn-outline-danger fw-bold ms-auto btn-delete-hd" title="Xóa Hội đồng này">✕ Xóa</button>
            </div>
            <div class="row">
                <div class="col-4 border-end">
                    <div class="mb-1"><small class="text-danger fw-bold">Chủ tịch:</small><div class="dropzone-gv border rounded p-1 bg-light" data-role="Chủ tịch" style="min-height: 38px;"></div></div>
                    <div class="mb-1"><small class="text-success fw-bold">Phản biện:</small><div class="dropzone-gv border rounded p-1 bg-light" data-role="Phản biện" style="min-height: 38px;"></div></div>
                    <div><small class="text-primary fw-bold">Thư ký:</small><div class="dropzone-gv border rounded p-1 bg-light" data-role="Thư ký" style="min-height: 38px;"></div></div>
                </div>
                <div class="col-8">
                    <div class="d-flex justify-content-between align-items-center mb-1">
                        <small class="fw-bold">Sinh viên bảo vệ:</small>
                        <button type="button" class="btn btn-outline-success btn-sm py-0 px-2 btn-open-add-sv" data-hd-id="${boxId}">
                            + Thêm SV
                        </button>
                    </div>
                    <div class="dropzone-sv p-2 bg-light border rounded" style="min-height: 120px;"></div>
                </div>
            </div>
        </div>
    `;

    const workspace = document.getElementById('workspace-hd');
    if (!workspace) return null;
    workspace.insertAdjacentHTML('beforeend', boxHTML);
    const boxEl = document.getElementById(boxId);

    boxEl?.querySelector('.btn-delete-hd')?.addEventListener('click', () => boxEl.remove());

    attachSortableToHoiDong(boxId);
    return boxId;
}

// Export hàm 3: Bóc tách file Excel Hội đồng, rà soát trùng Mã SV + trùng GVHD & đẩy SV vi phạm GVHD về Kho chờ
export function renderHoidongFromExcel(parsedHoidongs) {
    const workspace = document.getElementById('workspace-hd');
    if (!workspace) return { mergeLogs: [], conflictLogs: [], duplicateSVLogs: [] };
    workspace.innerHTML = '';
    hoidongCount = 0;

    const safeParsed = parsedHoidongs && typeof parsedHoidongs === 'object' ? parsedHoidongs : {};
    const mergeLogs = Array.isArray(safeParsed.__mergeLogs) ? safeParsed.__mergeLogs : [];
    const conflictLogs = [];
    const duplicateSVLogs = [];

    // Bước tiền rà soát: Thống kê mỗi Mã SV xuất hiện ở những Hội đồng nào
    const svAppearances = new Map(); // maSV -> { hoTen, hoidongs: [] }
    Object.keys(safeParsed).forEach(hdKey => {
        if (hdKey.startsWith('__')) return;
        const hd = safeParsed[hdKey];
        if (!hd || typeof hd !== 'object') return;
        const hdLabel = `${hd.soHD || hdKey}${hd.ngayBV ? ` (${hd.gioBV ? hd.gioBV + ' ' : ''}${hd.ngayBV})` : ''}`;

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
                `• SV ${info.hoTen} (${ma}) bị lặp lại ${info.hoidongs.length} lần tại: ${info.hoidongs.join(' ; ')} → Đã giữ ở HĐ đầu tiên và loại bỏ bản lặp phía sau, đề nghị kiểm tra sửa lại!`
            );
        }
    });

    const renderedSVSet = new Set(); // Đảm bảo mỗi Mã SV chỉ render 1 lần duy nhất trên giao diện

    Object.keys(safeParsed).forEach(hdKey => {
        if (hdKey.startsWith('__')) return;
        const hd = safeParsed[hdKey];
        if (!hd || typeof hd !== 'object') return;

        const thanhVien = Array.isArray(hd.thanhVien) ? hd.thanhVien : [];
        const sinhViens = Array.isArray(hd.sinhViens) ? hd.sinhViens : [];

        const formattedDate = formatDateToInput(hd.ngayBV);
        const boxId = createHoiDongBox(
            hd.soHD,
            formattedDate,
            hd.gioBV || '',
            hd.phongBV || '',
            hd.diaDiemBV || ''
        );

        const boxEl = boxId ? document.getElementById(boxId) : null;
        if (!boxEl) return;

        // 1. Render Giảng viên vào đúng chức danh
        thanhVien.forEach(gv => {
            let roleName = "";
            const chucDanhStr = String(gv.chucDanh || '');
            if (chucDanhStr.includes("Chủ tịch")) roleName = "Chủ tịch";
            else if (chucDanhStr.includes("Phản biện")) roleName = "Phản biện";
            else if (chucDanhStr.includes("Thư ký")) roleName = "Thư ký";

            if (roleName) {
                const zone = boxEl.querySelector(`.dropzone-gv[data-role="${roleName}"]`);
                if (zone) {
                    zone.innerHTML = `
                        <div class="card-item position-relative bg-light border-info p-1 pe-4 rounded shadow-sm" 
                             data-id="${escapeHtmlAttr(gv.maGV)}" data-name="${escapeHtmlAttr(gv.hoTenGV)}" data-donvi="${escapeHtmlAttr(gv.donVi || 'Đại học Duy Tân')}">
                            <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card" title="Xóa giảng viên này">✕</button>
                            <div class="fw-bold text-dark small">${escapeHtmlAttr(gv.hoTenGV)}</div>
                            <div class="text-muted text-truncate" style="font-size: 0.75rem;">${escapeHtmlAttr(gv.donVi || 'Đại học Duy Tân')} - Mã: ${escapeHtmlAttr(gv.maGV)}</div>
                        </div>
                    `;
                }
            }
        });

        // 2. Render Sinh viên & RÀ SOÁT TRÙNG MÃ SV + TRÙNG GVHD
        const svDropzone = boxEl.querySelector('.dropzone-sv');
        const poolSV = document.getElementById('pool-sv');

        sinhViens.forEach(sv => {
            const maSVNorm = String(sv.maSV || '').trim().toUpperCase();
            if (!maSVNorm) return;

            // Bỏ qua nếu Mã SV này đã được xử lý ở Hội đồng trước đó (tránh lặp trùng)
            if (renderedSVSet.has(maSVNorm)) return;
            renderedSVSet.add(maSVNorm);

            const poolItem = findPoolSVItemById(sv.maSV);

            const hoTen = poolItem ? poolItem.getAttribute('data-name') : (sv.hoTenSV || sv.hoTen || '');
            const lop = poolItem ? poolItem.getAttribute('data-lop') : (sv.lop || '');
            const deTai = poolItem ? poolItem.getAttribute('data-detai') : (sv.deTai || '');
            const maGVHD = poolItem ? (poolItem.getAttribute('data-magvhd') || '') : (sv.maGVHD || '');
            const gvhd = poolItem ? (poolItem.getAttribute('data-gvhd') || '') : (sv.gvhd || '');
            const nganh = poolItem ? (poolItem.getAttribute('data-nganh') || '') : (sv.nganh || '');
            const maGVDisplay = maGVHD ? ` [${maGVHD}]` : '';

            // Kiểm tra GVHD của SV có trùng với thành viên nào trong HĐ này không
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
                    poolSV?.prepend(poolItem);
                }

                conflictLogs.push(
                    `• SV ${hoTen} (${sv.maSV}) - GVHD: ${gvhd || maGVHD} trùng với ${violatedMember.chucDanh} (${violatedMember.hoTenGV}) tại ${hd.soHD} → Đã đẩy về Kho chờ, đề nghị xếp sang Hội đồng khác!`
                );
                return;
            }

            // Hợp lệ -> Đưa vào Hội đồng và xóa khỏi Kho chờ
            const svHTML = `
                <div class="card-item position-relative mb-2 p-2 pe-4 border rounded bg-white shadow-sm" 
                     data-id="${escapeHtmlAttr(sv.maSV)}" data-name="${escapeHtmlAttr(hoTen)}" data-lop="${escapeHtmlAttr(lop)}" data-detai="${escapeHtmlAttr(deTai)}" data-magvhd="${escapeHtmlAttr(maGVHD)}" data-gvhd="${escapeHtmlAttr(gvhd)}" data-nganh="${escapeHtmlAttr(nganh)}">
                    <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card" title="Xóa sinh viên này">✕</button>
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

    return { mergeLogs, conflictLogs, duplicateSVLogs };
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

            // LỌC CHỐNG VI PHẠM 2 LỚP: Kiểm tra bằng Mã GVHD trước -> dự phòng Tên chuẩn hóa
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
                }
            }
        });
    }

    box.querySelectorAll('.dropzone-gv').forEach(zone => {
        Sortable.create(zone, {
            group: 'giangvien',
            animation: 150,
            onAdd: function (evt) {
                const item = evt.item;
                const gvId = item.getAttribute('data-id') || '';
                const gvName = item.getAttribute('data-name') || '';
                const currBox = evt.to.closest('.box-hoidong');
                if (!currBox) return;

                if (evt.to.children.length > 1) {
                    const firstChild = evt.to.children[0];
                    if (firstChild && firstChild !== item) evt.to.removeChild(firstChild);
                }

                const otherGVInThisHD = Array.from(currBox.querySelectorAll('.dropzone-gv .card-item')).filter(node => node !== item);
                if (otherGVInThisHD.some(node => isSameTeacher(
                    { id1: gvId, name1: gvName },
                    { id2: node.getAttribute('data-id') || '', name2: node.getAttribute('data-name') || '' }
                ))) {
                    alert(`LỖI: Giảng viên ${gvName} đã có vị trí khác trong Hội đồng này!`);
                    item.remove();
                    return;
                }

                const violatedSVName = validateTeacherIsGVHDOfAnyStudent(gvName, currBox, gvId);
                if (violatedSVName) {
                    alert(`⛔ VI PHẠM NGUYÊN TẮC: ${gvName} là GVHD của sinh viên ${violatedSVName} đang thuộc Hội đồng này!`);
                    item.remove();
                    return;
                }

                const conflictHDTitle = checkScheduleConflict(gvId, gvName, currBox);
                if (conflictHDTitle) {
                    alert(`CẢNH BÁO TRÙNG LỊCH: GV ${gvName} đang có lịch tại ${conflictHDTitle} trong cùng khung giờ!`);
                    item.remove();
                    return;
                }
            }
        });
    });
}
