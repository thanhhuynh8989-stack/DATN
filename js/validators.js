// js/validators.js

// Hàm làm sạch và so sánh tên Giảng viên (tự bỏ qua ThS., TS., GS., PGS., KS., CN...)
export function isSameTeacherName(name1, name2) {
    if (!name1 || !name2) return false;
    const clean = (str) => String(str).toLowerCase()
        .replace(/gs\.?|pgs\.?|ts\.?|ths\.?|gsv\.?|ks\.?|cn\.?|gv\.?|prof\.?|dr\.?/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    return clean(name1) === clean(name2);
}

// Hàm cốt lõi: So khớp 2 Giảng viên (Ưu tiên khớp Mã GV trước -> Dự phòng khớp Tên chuẩn hóa)
export function isSameTeacher({ id1 = '', name1 = '' }, { id2 = '', name2 = '' }) {
    const cleanId1 = String(id1 || '').trim().toUpperCase();
    const cleanId2 = String(id2 || '').trim().toUpperCase();

    // 1. Ưu tiên tuyệt đối: So sánh Mã GV nếu cả hai đều có Mã GV hợp lệ (loại trừ mã tự sinh GV_...)
    if (cleanId1 && cleanId2 && !cleanId1.startsWith('GV_') && !cleanId2.startsWith('GV_')) {
        return cleanId1 === cleanId2;
    }

    // 2. Dự phòng: So sánh tên sau khi loại bỏ học hàm/học vị
    return isSameTeacherName(name1, name2);
}

// Kiểm tra SV có trùng GVHD với Thành viên HĐ không (hỗ trợ cả Tên lẫn Mã GVHD)
export function validateGVHDConflictWithHD(gvhdName, hdBox, maGVHD = '') {
    if (!gvhdName && !maGVHD) return false;
    const hdGVNodes = hdBox.querySelectorAll('.dropzone-gv .card-item');

    for (let gvNode of hdGVNodes) {
        const memberId = (gvNode.getAttribute('data-id') || '').trim();
        const memberName = (gvNode.getAttribute('data-name') || '').trim();

        if (isSameTeacher({ id1: maGVHD, name1: gvhdName }, { id2: memberId, name2: memberName })) {
            return memberName || memberId; // Trả về tên GV bị trùng để hiển thị thông báo
        }
    }
    return false;
}

// Kiểm tra GV được kéo thả có trùng GVHD của bất kỳ SV nào trong HĐ không (hỗ trợ cả Tên lẫn Mã GV)
export function validateTeacherIsGVHDOfAnyStudent(gvName, hdBox, gvId = '') {
    const svCards = hdBox.querySelectorAll('.dropzone-sv .card-item');
    for (let svNode of svCards) {
        const maGVHDOfSV = (svNode.getAttribute('data-magvhd') || '').trim();
        const gvhdOfSV = (svNode.getAttribute('data-gvhd') || '').trim();

        if (isSameTeacher({ id1: maGVHDOfSV, name1: gvhdOfSV }, { id2: gvId, name2: gvName })) {
            return svNode.getAttribute('data-name'); // Trả về tên SV bị vi phạm
        }
    }
    return false;
}

// Kiểm tra cảnh báo trùng lịch GV ở HĐ khác (Cùng Ngày + Giờ)
export function checkScheduleConflict(gvId, gvName, currBox) {
    const thisDate = currBox.querySelector('.hd-date')?.value;
    const thisTime = currBox.querySelector('.hd-time')?.value;

    if (!thisDate || !thisTime) return false;

    const allHDs = document.querySelectorAll('.box-hoidong');
    for (let hd of allHDs) {
        if (hd.id === currBox.id) continue;
        const thatDate = hd.querySelector('.hd-date')?.value;
        const thatTime = hd.querySelector('.hd-time')?.value;

        if (thatDate === thisDate && thatTime === thisTime) {
            const otherGVNodes = hd.querySelectorAll('.dropzone-gv .card-item');
            for (let otherNode of otherGVNodes) {
                const otherId = otherNode.getAttribute('data-id') || '';
                const otherName = otherNode.getAttribute('data-name') || '';
                if (isSameTeacher({ id1: gvId, name1: gvName }, { id2: otherId, name2: otherName })) {
                    return hd.querySelector('.hd-title')?.value || 'Hội đồng khác';
                }
            }
        }
    }
    return false;
}