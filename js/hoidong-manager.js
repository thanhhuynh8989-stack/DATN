// js/hoidong-manager.js
import { validateGVHDConflictWithHD, validateTeacherIsGVHDOfAnyStudent, checkScheduleConflict } from './validators.js';

let hoidongCount = 0;

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

    // Lắng nghe sự kiện click nút "+ Thêm SV" và "Xác nhận Gán" (Chỉ đăng ký 1 lần)
    initAddSVModalEvents();
}

// Export hàm 2: Tạo Box Hội Đồng mới
export function createHoiDongBox(titleText = null, initDate = '', initTime = '17:00', initRoom = '', initLocation = '120 Hoàng Minh Thảo') {
    hoidongCount++;
    const hdTitle = titleText || `Hội Đồng ${hoidongCount}`;
    const boxId = `hd-box-${Date.now()}-${Math.floor(Math.random() * 1000)}`;

    let formattedTime = initTime || '17:00';
    if (formattedTime.includes('h')) {
        const parts = formattedTime.split('h');
        formattedTime = `${parts[0].padStart(2, '0')}:${(parts[1] || '00').padStart(2, '0')}`;
    }

    const locations = [
        "120 Hoàng Minh Thảo",
        "254 Nguyễn Văn Linh",
        "137 Nguyễn Văn Linh",
        "03 Quang Trung",
        "K7/25 Quang Trung",
        "Phan Văn Trị"
    ];
    const locationOptions = locations.map(loc => 
        `<option value="${loc}" ${loc === initLocation ? 'selected' : ''}>${loc}</option>`
    ).join('');

    const boxHTML = `
        <div class="box-hoidong p-2 mb-3 border rounded bg-white shadow-sm" id="${boxId}">
            <div class="d-flex gap-2 mb-2 bg-light p-2 rounded align-items-center flex-wrap">
                <input type="text" class="form-control form-control-sm fw-bold hd-title" value="${hdTitle}" placeholder="Số/Tên HĐ" style="width: 110px;">
                <input type="date" class="form-control form-control-sm hd-date" value="${initDate}" style="width: 125px;" title="Ngày bảo vệ">
                <input type="time" class="form-control form-control-sm hd-time" value="${formattedTime}" style="width: 110px;" title="Giờ:Phút bảo vệ">
                <input type="text" class="form-control form-control-sm hd-room" value="${initRoom}" placeholder="Phòng (111)" style="width: 105px;">
                <select class="form-select form-select-sm hd-location" style="width: 150px;">
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
                        <!-- NÚT GÁN SV NHANH CHỐNG VI PHẠM -->
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
    workspace.insertAdjacentHTML('beforeend', boxHTML);
    const boxEl = document.getElementById(boxId);

    boxEl.querySelector('.btn-delete-hd')?.addEventListener('click', () => boxEl.remove());

    attachSortableToHoiDong(boxId);
    return boxId;
}

// Export hàm 3: Bóc tách file Excel Hội đồng, đổ thông tin đầy đủ & XÓA SV ĐÃ PHÂN CÔNG KHỎI KHO CHỜ
export function renderHoidongFromExcel(parsedHoidongs) {
    const workspace = document.getElementById('workspace-hd');
    if (!workspace) return;
    workspace.innerHTML = '';
    hoidongCount = 0;

    Object.keys(parsedHoidongs).forEach(hdKey => {
        const hd = parsedHoidongs[hdKey];
        const formattedDate = formatDateToInput(hd.ngayBV);
        const boxId = createHoiDongBox(
            hd.soHD, 
            formattedDate, 
            hd.gioBV || '17:00', 
            hd.phongBV || '', 
            hd.diaDiemBV || '120 Hoàng Minh Thảo'
        );

        const boxEl = document.getElementById(boxId);
        if (!boxEl) return;

        // 1. Render Giảng viên vào đúng chức danh theo Mã GV
        hd.thanhVien.forEach(gv => {
            let roleName = "";
            if (gv.chucDanh.includes("Chủ tịch")) roleName = "Chủ tịch";
            else if (gv.chucDanh.includes("Phản biện")) roleName = "Phản biện";
            else if (gv.chucDanh.includes("Thư ký")) roleName = "Thư ký";

            if (roleName) {
                const zone = boxEl.querySelector(`.dropzone-gv[data-role="${roleName}"]`);
                if (zone) {
                    zone.innerHTML = `
                        <div class="card-item bg-light border-info p-1 rounded shadow-sm" data-id="${gv.maGV}" data-name="${gv.hoTenGV}">
                            <div class="fw-bold text-dark small">${gv.hoTenGV}</div>
                            <div class="text-muted text-truncate" style="font-size: 0.75rem;">${gv.donVi || ''} - Mã: ${gv.maGV}</div>
                        </div>
                    `;
                }
            }
        });

        // 2. Render Sinh viên & XÓA KHỎI KHO CHỜ THEO MÃ SV (maSV)
        const svDropzone = boxEl.querySelector('.dropzone-sv');
        hd.sinhViens.forEach(sv => {
            const poolItem = document.querySelector(`#pool-sv .card-item[data-id="${sv.maSV}"]`);
            
            const hoTen = poolItem ? poolItem.getAttribute('data-name') : (sv.hoTenSV || sv.hoTen);
            const lop = poolItem ? poolItem.getAttribute('data-lop') : (sv.lop || '');
            const deTai = poolItem ? poolItem.getAttribute('data-detai') : (sv.deTai || '');
            const gvhd = poolItem ? poolItem.getAttribute('data-gvhd') : '';

            const svHTML = `
                <div class="card-item mb-2 p-2 border rounded bg-white shadow-sm" 
                     data-id="${sv.maSV}" data-name="${hoTen}" data-lop="${lop}" data-detai="${deTai}" data-gvhd="${gvhd}">
                    <div class="fw-bold text-primary">${hoTen} (${sv.maSV})</div>
                    <div class="text-muted small">Lớp: ${lop} | GVHD: <span class="fw-bold text-dark">${gvhd || 'Chưa rõ'}</span></div>
                    <div class="text-dark small text-truncate" title="${deTai}">Đề tài: ${deTai}</div>
                </div>`;
            
            svDropzone?.insertAdjacentHTML('beforeend', svHTML);

            if (poolItem) {
                poolItem.remove();
            }
        });
    });
}

// ----------------------------------------------------
// THUẬT TOÁN MỞ MODAL & LỌC SINH VIÊN PHÙ HỢP (CLICK-TO-ASSIGN)
// ----------------------------------------------------
let targetBoxIdForAddSV = null;

function initAddSVModalEvents() {
    // Lắng nghe click mở Modal từ các box HĐ
    document.addEventListener('click', (e) => {
        const btnAdd = e.target.closest('.btn-open-add-sv');
        if (btnAdd) {
            targetBoxIdForAddSV = btnAdd.getAttribute('data-hd-id');
            openSelectSVModal(targetBoxIdForAddSV);
        }
    });

    // Lắng nghe click nút "➕ Gán vào Hội đồng" trong Modal
    document.getElementById('btn-confirm-gan-sv')?.addEventListener('click', () => {
        if (!targetBoxIdForAddSV) return;
        const box = document.getElementById(targetBoxIdForAddSV);
        const svDropzone = box?.querySelector('.dropzone-sv');
        if (!svDropzone) return;

        const checkedInputs = document.querySelectorAll('#list-sv-phu-hop .chk-select-sv:checked');
        checkedInputs.forEach(chk => {
            const maSV = chk.value;
            const poolItem = document.querySelector(`#pool-sv .card-item[data-id="${maSV}"]`);
            if (poolItem) {
                svDropzone.appendChild(poolItem);
            }
        });

        const modalEl = document.getElementById('modalGanSinhVien');
        const modalInstance = bootstrap.Modal.getInstance(modalEl);
        modalInstance?.hide();
    });
}

function openSelectSVModal(boxId) {
    const box = document.getElementById(boxId);
    if (!box) return;

    const hdTitle = box.querySelector('.hd-title')?.value || 'Hội Đồng';
    const titleEl = document.getElementById('modalGanSVTitle');
    if (titleEl) titleEl.innerText = `📋 CHỌN SINH VIÊN VÀO ${hdTitle.toUpperCase()}`;

    // 1. Trích xuất TÊN các Giảng viên đang có trong Hội đồng này (Chủ tịch, Phản biện, Thư ký)
    const activeGVNames = Array.from(box.querySelectorAll('.dropzone-gv .card-item'))
                               .map(item => (item.getAttribute('data-name') || '').trim())
                               .filter(name => name.length > 0);

    // 2. Lấy danh sách SV còn trong Kho chờ (pool-sv)
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
            const gvhd = (svItem.getAttribute('data-gvhd') || '').trim();

            // 3. LỌC CHỐNG VI PHẠM: GVHD của SV không được nằm trong danh sách GV đang thuộc HĐ này
            const isConflicted = activeGVNames.some(gvName => gvName.toLowerCase() === gvhd.toLowerCase());

            if (!isConflicted) {
                countAvailable++;
                const itemHTML = `
                    <label class="list-group-item d-flex align-items-center gap-2 cursor-pointer">
                        <input class="form-check-input flex-shrink-0 chk-select-sv" type="checkbox" value="${maSV}">
                        <div class="w-100">
                            <div class="fw-bold text-primary mb-0">${hoTen} (${maSV})</div>
                            <div class="small text-muted">Lớp: ${lop} | GVHD: <span class="fw-bold text-dark">${gvhd || 'Chưa rõ'}</span></div>
                            <div class="small text-dark text-truncate">Đề tài: ${deTai}</div>
                        </div>
                    </label>`;
                container.insertAdjacentHTML('beforeend', itemHTML);
            }
        });

        if (countAvailable === 0) {
            container.innerHTML = `
                <div class="text-center text-danger py-3 fw-bold">
                    ⛔ Tất cả SV trong kho chờ đều bị trùng GVHD với Thành viên của ${hdTitle}!
                </div>`;
        }
    }

    const modal = new bootstrap.Modal(document.getElementById('modalGanSinhVien'));
    modal.show();
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

    Sortable.create(box.querySelector('.dropzone-sv'), {
        group: 'sinhvien',
        animation: 150,
        onAdd: function (evt) {
            const svCard = evt.item;
            const svName = svCard.getAttribute('data-name');
            const gvhdName = svCard.getAttribute('data-gvhd');

            const conflictedMember = validateGVHDConflictWithHD(gvhdName, box);
            if (conflictedMember) {
                alert(`⛔ VI PHẠM NGUYÊN TẮC: GVHD (${gvhdName}) đang làm thành viên HĐ (${conflictedMember}). Không thể xếp sinh viên ${svName} vào HĐ này!`);
                evt.from.appendChild(svCard);
            }
        }
    });

    box.querySelectorAll('.dropzone-gv').forEach(zone => {
        Sortable.create(zone, {
            group: 'giangvien',
            animation: 150,
            onAdd: function (evt) {
                const item = evt.item;
                const gvId = item.getAttribute('data-id');
                const gvName = item.getAttribute('data-name');
                const currBox = evt.to.closest('.box-hoidong');

                if (evt.to.children.length > 1) evt.to.removeChild(evt.to.children[0]);

                const allGVInThisHD = Array.from(currBox.querySelectorAll('.dropzone-gv .card-item'));
                if (allGVInThisHD.filter(node => node.getAttribute('data-id') === gvId).length > 1) {
                    alert(`LỖI: Giảng viên ${gvName} đã có vị trí trong Hội đồng này!`);
                    item.remove(); return;
                }

                const violatedSVName = validateTeacherIsGVHDOfAnyStudent(gvName, currBox);
                if (violatedSVName) {
                    alert(`⛔ VI PHẠM NGUYÊN TẮC: ${gvName} là GVHD của sinh viên ${violatedSVName} đang thuộc Hội đồng này!`);
                    item.remove(); return;
                }

                const conflictHDTitle = checkScheduleConflict(gvId, gvName, currBox);
                if (conflictHDTitle) {
                    alert(`CẢNH BÁO TRÙNG LỊCH: GV ${gvName} đang có lịch tại ${conflictHDTitle} trong cùng khung giờ!`);
                    item.remove(); return;
                }
            }
        });
    });
}
