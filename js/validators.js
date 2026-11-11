// js/validators.js

// Hàm làm sạch và so sánh tên Giảng viên (tự bỏ qua ThS., TS., GS., PGS...)
export function isSameTeacherName(name1, name2) {
    if (!name1 || !name2) return false;
    const clean = (str) => str.toLowerCase()
        .replace(/ths\.|ts\.|gsv\.|pgs\.|prof\.|dr\./g, '')
        .replace(/\s+/g, ' ')
        .trim();
    return clean(name1) === clean(name2);
}

// Kiểm tra SV có trùng GVHD với Thành viên HĐ không
export function validateGVHDConflictWithHD(gvhdName, hdBox) {
    if (!gvhdName) return false;
    const cleanGVHD = gvhdName.toLowerCase().trim();
    const hdGVNodes = hdBox.querySelectorAll('.dropzone-gv .card-item');

    for (let gvNode of hdGVNodes) {
        const memberName = (gvNode.getAttribute('data-name') || '').toLowerCase().trim();
        if (isSameTeacherName(cleanGVHD, memberName)) {
            return memberName; // Trả về tên GV bị trùng để hiển thị thông báo
        }
    }
    return false;
}

// Kiểm tra GV được kéo thả có trùng GVHD của bất kỳ SV nào trong HĐ không
export function validateTeacherIsGVHDOfAnyStudent(gvName, hdBox) {
    const svCards = hdBox.querySelectorAll('.dropzone-sv .card-item');
    for (let svNode of svCards) {
        const gvhdOfSV = (svNode.getAttribute('data-gvhd') || '').toLowerCase().trim();
        if (isSameTeacherName(gvhdOfSV, gvName)) {
            return svNode.getAttribute('data-name'); // Trả về tên SV bị vi phạm
        }
    }
    return false;
}

// Kiểm tra cảnh báo trùng lịch GV ở HĐ khác (Cùng Ngày + Giờ)
export function checkScheduleConflict(gvId, gvName, currBox) {
    const thisDate = currBox.querySelector('.hd-date').value;
    const thisTime = currBox.querySelector('.hd-time').value;

    if (!thisDate || !thisTime) return false;

    const allHDs = document.querySelectorAll('.box-hoidong');
    for (let hd of allHDs) {
        if (hd.id === currBox.id) continue;
        const thatDate = hd.querySelector('.hd-date').value;
        const thatTime = hd.querySelector('.hd-time').value;

        if (thatDate === thisDate && thatTime === thisTime) {
            const existInOther = hd.querySelector(`.dropzone-gv .card-item[data-id="${gvId}"]`);
            if (existInOther) {
                return hd.querySelector('.hd-title').value; // Trả về tên HĐ bị trùng
            }
        }
    }
    return false;
}
