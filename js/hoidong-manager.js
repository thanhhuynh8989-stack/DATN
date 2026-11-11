// js/hoidong-manager.js
import { validateGVHDConflictWithHD, validateTeacherIsGVHDOfAnyStudent, checkScheduleConflict } from './validators.js';

let hoidongCount = 0;

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
}

export function createHoiDongBox(titleText = null, initDate = '', initTime = '', initRoom = '') {
    hoidongCount++;
    const hdTitle = titleText || `Hội Đồng ${hoidongCount}`;
    const boxId = `hd-box-${Date.now()}-${Math.floor(Math.random() * 1000)}`;

    const boxHTML = `
        <div class="box-hoidong p-2 mb-3 border rounded bg-white shadow-sm" id="${boxId}">
            <div class="d-flex gap-2 mb-2 bg-light p-2 rounded align-items-center">
                <input type="text" class="form-control form-control-sm fw-bold hd-title" value="${hdTitle}" style="width: 120px;">
                <input type="date" class="form-control form-control-sm hd-date" value="${initDate}" style="width: 130px;">
                <select class="form-select form-select-sm hd-time" style="width: 95px;">
                    <option value="">--Giờ--</option>
                    <option value="07h00" ${initTime==='07h00'?'selected':''}>07h00</option>
                    <option value="08h00" ${initTime==='08h00'?'selected':''}>08h00</option>
                    <option value="13h00" ${initTime==='13h00'?'selected':''}>13h00</option>
                    <option value="14h00" ${initTime==='14h00'?'selected':''}>14h00</option>
                    <option value="17h00" ${initTime==='17h00'?'selected':''}>17h00</option>
                </select>
                <input type="text" class="form-control form-control-sm hd-room" value="${initRoom}" placeholder="Phòng" style="width: 110px;">
                <select class="form-select form-select-sm hd-location">
                    <option value="120 Hoàng Minh Thảo">120 Hoàng Minh Thảo</option>
                    <option value="254 Nguyễn Văn Linh">254 Nguyễn Văn Linh</option>
                    <option value="137 Nguyễn Văn Linh">137 Nguyễn Văn Linh</option>
                    <option value="03 Quang Trung">03 Quang Trung</option>
                    <option value="K7/25 Quang Trung">K7/25 Quang Trung</option>
                    <option value="Phan Văn Trị">Phan Văn Trị</option>
                </select>
                <button type="button" class="btn btn-sm btn-outline-danger fw-bold ms-auto btn-delete-hd">✕ Xóa</button>
            </div>
            <div class="row">
                <div class="col-4 border-end">
                    <div class="mb-1"><small class="text-danger fw-bold">Chủ tịch:</small><div class="dropzone-gv" data-role="Chủ tịch"></div></div>
                    <div class="mb-1"><small class="text-success fw-bold">Phản biện:</small><div class="dropzone-gv" data-role="Phản biện"></div></div>
                    <div><small class="text-primary fw-bold">Thư ký:</small><div class="dropzone-gv" data-role="Thư ký"></div></div>
                </div>
                <div class="col-8">
                    <small class="fw-bold">Sinh viên bảo vệ:</small>
                    <div class="dropzone-sv p-2 bg-light border rounded" style="min-height: 120px;"></div>
                </div>
            </div>
        </div>
    `;

    const workspace = document.getElementById('workspace-hd');
    workspace.insertAdjacentHTML('beforeend', boxHTML);
    const boxEl = document.getElementById(boxId);

    // Gán sự kiện Xóa
    boxEl.querySelector('.btn-delete-hd').addEventListener('click', () => boxEl.remove());

    attachSortableToHoiDong(boxId);
    return boxId;
}

// Export hàm 3: Bóc tách file Excel Hội đồng và tạo danh sách Hội đồng
export function renderHoidongFromExcel(rows) {
    const workspace = document.getElementById('workspace-hd');
    if (!workspace) return;
    workspace.innerHTML = '';
    hoidongCount = 0;

    let defaultTime = "", defaultDate = "", defaultRoom = "";

    // Bóc tách Thời gian bảo vệ từ các dòng tiêu đề file Excel
    for (let i = 0; i < Math.min(rows.length, 7); i++) {
        const rowStr = (rows[i] || []).join(' ');
        if (rowStr.includes('Thời gian bảo vệ')) {
            const timeMatch = rowStr.match(/(\d{1,2}h\d{2})/);
            if (timeMatch) defaultTime = timeMatch[1];
            
            const dayMatch = rowStr.match(/ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})/);
            if (dayMatch) {
                const d = dayMatch[1].padStart(2, '0');
                const m = dayMatch[2].padStart(2, '0');
                const y = dayMatch[3];
                defaultDate = `${y}-${m}-${d}`;
            }
            break;
        }
    }

    let currentBoxId = null;
    let currentHDNum = null;

    rows.forEach((row, idx) => {
        if (idx < 7 || !row) return;

        const hdVal = String(row[0] || '').trim();
        const maGV = String(row[1] || '').trim();
        const hoTenGV = String(row[2] || '').trim();
        const chucDanh = String(row[5] || '').trim();

        if (hdVal && hdVal !== currentHDNum && !isNaN(hdVal)) {
            currentHDNum = hdVal;
            currentBoxId = createHoiDongBox(`Hội Đồng ${hdVal}`, defaultDate, defaultTime, defaultRoom);
        }

        if (currentBoxId && hoTenGV && chucDanh) {
            const box = document.getElementById(currentBoxId);
            let roleName = "";
            if (chucDanh.includes("Chủ tịch")) roleName = "Chủ tịch";
            else if (chucDanh.includes("Phản biện")) roleName = "Phản biện";
            else if (chucDanh.includes("Thư ký")) roleName = "Thư ký";

            if (roleName) {
                const zone = box.querySelector(`.dropzone-gv[data-role="${roleName}"]`);
                if (zone) {
                    zone.innerHTML = `
                        <div class="card-item bg-light border-info p-1 rounded" data-id="${maGV}" data-name="${hoTenGV}">
                            <div class="fw-bold text-dark small">${hoTenGV}</div>
                        </div>
                    `;
                }
            }
        }
    });
}

function attachSortableToHoiDong(boxId) {
    const box = document.getElementById(boxId);
    if (!box) return;

    // Kéo thả Sinh Viên & Kiểm tra Ràng buộc GVHD
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

    // Kéo thả Giảng Viên & Kiểm tra các Ràng buộc trùng lặp
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
