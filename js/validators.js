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

// Chuẩn hóa tên GV: Bỏ học hàm/học vị -> Bỏ dấu tiếng Việt -> Xóa toàn bộ khoảng trắng -> Viết thường
export function normalizeTeacherKeyForDupCheck(rawName) {
    if (!rawName) return '';
    let s = String(rawName).toLowerCase().trim();
    // Bỏ học hàm, học vị phổ biến
    s = s.replace(/\b(gs|pgs|ts|ths|gvc|gv|ks|cn|bs|dsc|prof|dr|msc|mr|ms|mrs)\.?\b/gi, ' ');
    // Bỏ dấu tiếng Việt
    s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd');
    // Xóa toàn bộ khoảng trắng và ký tự đặc biệt
    s = s.replace(/[^a-z0-9]/g, '');
    return s;
}

// Gom danh sách cảnh báo thiếu Mã GV hoặc thiếu Mã SV (bỏ qua nếu HĐ chưa có SV)
export function collectMissingCodeWarnings({ svList = [], gvList = [], parsedHoidongs = {} }) {
    const warnings = [];

    if (Array.isArray(svList.__missingCodeIssues)) {
        warnings.push(...svList.__missingCodeIssues);
    } else {
        svList.forEach((sv, idx) => {
            if (sv.hoTen && !sv.maSV) {
                warnings.push(`SV #${idx + 1}: "${sv.hoTen}" bị thiếu Mã SV.`);
            }
            if (sv.gvhd && !sv.maGVHD) {
                warnings.push(`SV #${idx + 1} ("${sv.hoTen || sv.maSV || ''}"): GVHD "${sv.gvhd}" bị thiếu Mã GV.`);
            }
        });
    }

    if (Array.isArray(gvList.__missingCodeIssues)) {
        warnings.push(...gvList.__missingCodeIssues);
    } else {
        gvList.forEach((gv, idx) => {
            const id = String(gv.maGV || '').trim();
            if (gv.hoTenGV && (!id || id.toUpperCase().startsWith('GV_'))) {
                warnings.push(`GV #${idx + 1}: "${gv.hoTenGV}" bị thiếu Mã GV.`);
            }
        });
    }

    if (Array.isArray(parsedHoidongs.__missingCodeIssues)) {
        warnings.push(...parsedHoidongs.__missingCodeIssues);
    } else {
        Object.keys(parsedHoidongs).forEach(key => {
            if (key === '__mergeLogs' || key === '__missingCodeIssues') return;
            const hd = parsedHoidongs[key];
            const hdLabel = `HĐ ${hd.soHD || key}`;
            (hd.thanhVien || []).forEach(tv => {
                const id = String(tv.maGV || '').trim();
                if (tv.hoTenGV && (!id || id.toUpperCase().startsWith('GV_'))) {
                    warnings.push(`${hdLabel}: Giảng viên "${tv.hoTenGV}" (${tv.chucDanh || ''}) bị thiếu Mã GV.`);
                }
            });
            (hd.sinhViens || []).forEach(sv => {
                if (sv.hoTenSV && !sv.maSV) {
                    warnings.push(`${hdLabel}: Sinh viên "${sv.hoTenSV}" bị thiếu Mã SV.`);
                }
            });
        });
    }

    return [...new Set(warnings)];
}

// Đối chiếu chéo tất cả các nguồn (DS SV, DS GV, DS HĐ) để tìm GV trùng tên chuẩn hóa nhưng khác Mã GV
export function detectDuplicateTeacherNameDifferentId({ svList = [], gvList = [], parsedHoidongs = {} }) {
    // Map: normalizedName -> Map(maGV -> Set(locations))
    const nameMap = new Map();

    const addRecord = (rawName, rawId, locationDesc) => {
        const normName = normalizeTeacherKeyForDupCheck(rawName);
        const id = String(rawId || '').trim().toUpperCase();
        // Bỏ qua nếu thiếu tên hoặc thiếu mã thật (mã tự sinh GV_... không tính là mã xung đột)
        if (!normName || !id || id.startsWith('GV_')) return;

        if (!nameMap.has(normName)) {
            nameMap.set(normName, new Map());
        }
        const idMap = nameMap.get(normName);
        if (!idMap.has(id)) {
            idMap.set(id, { displayNames: new Set(), locations: [] });
        }
        const entry = idMap.get(id);
        entry.displayNames.add(String(rawName).trim());
        if (!entry.locations.includes(locationDesc)) {
            entry.locations.push(locationDesc);
        }
    };

    // 1. Nguồn từ Danh sách Sinh viên & Đồ án (cột GVHD + Mã GVHD)
    svList.forEach(sv => {
        if (sv.gvhd && sv.maGVHD) {
            addRecord(sv.gvhd, sv.maGVHD, `File SV: SV ${sv.hoTen} (${sv.maSV})`);
        }
    });
    // Cả những SV đã có sẵn trên giao diện (nếu nạp bổ sung file HĐ sau)
    document.querySelectorAll('#pool-sv .card-item, .dropzone-sv .card-item').forEach(el => {
        const gvhd = el.getAttribute('data-gvhd') || '';
        const ma = el.getAttribute('data-magvhd') || '';
        const svName = el.getAttribute('data-name') || '';
        const svId = el.getAttribute('data-id') || '';
        if (gvhd && ma) addRecord(gvhd, ma, `Kho/HĐ hiện tại: SV ${svName} (${svId})`);
    });

    // 2. Nguồn từ Danh sách Giảng viên
    gvList.forEach(gv => {
        addRecord(gv.hoTenGV, gv.maGV, `File GV: Danh mục Giảng viên`);
    });
    document.querySelectorAll('#pool-gv .card-item').forEach(el => {
        const name = el.getAttribute('data-name') || '';
        const id = el.getAttribute('data-id') || '';
        if (name && id) addRecord(name, id, `Kho GV hiện tại`);
    });

    // 3. Nguồn từ Danh sách Hội đồng upload
    Object.keys(parsedHoidongs).forEach(key => {
        if (key === '__mergeLogs') return;
        const hd = parsedHoidongs[key];
        const hdLabel = `${hd.soHD}${hd.ngayBV ? ` (${hd.gioBV ? hd.gioBV + ' ' : ''}${hd.ngayBV})` : ''}`;
        (hd.thanhVien || []).forEach(tv => {
            addRecord(tv.hoTenGV, tv.maGV, `File HĐ: ${hdLabel} - ${tv.chucDanh}`);
        });
    });

    // Lọc ra các giảng viên có >= 2 Mã GV khác nhau cho cùng 1 tên chuẩn hóa
    const conflicts = [];
    nameMap.forEach((idMap, normName) => {
        if (idMap.size > 1) {
            const variants = [];
            idMap.forEach((info, id) => {
                variants.push({
                    maGV: id,
                    names: Array.from(info.displayNames).join(' / '),
                    locations: info.locations
                });
            });
            conflicts.push({ normName, variants });
        }
    });

    return conflicts;
}

// Chuẩn hóa tên SV: Bỏ dấu tiếng Việt -> Xóa toàn bộ khoảng trắng & ký tự đặc biệt -> Viết thường
export function normalizeStudentKeyForDupCheck(rawName) {
    if (!rawName) return '';
    return String(rawName)
        .toLowerCase()
        .trim()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd')
        .replace(/[^a-z0-9]/g, '');
}

// Đối chiếu chéo giữa DS SV, DS HĐ và SV hiện có để tìm SV trùng tên chuẩn hóa nhưng khác Mã SV
export function detectDuplicateStudentNameDifferentId({ svList = [], parsedHoidongs = {} }) {
    // Map: normalizedName -> Map(maSV -> { displayNames: Set, locations: Array })
    const nameMap = new Map();

    const addRecord = (rawName, rawId, locationDesc) => {
        const normName = normalizeStudentKeyForDupCheck(rawName);
        const id = String(rawId || '').trim().toUpperCase();
        if (!normName || !id) return;

        if (!nameMap.has(normName)) {
            nameMap.set(normName, new Map());
        }
        const idMap = nameMap.get(normName);
        if (!idMap.has(id)) {
            idMap.set(id, { displayNames: new Set(), locations: [] });
        }
        const entry = idMap.get(id);
        entry.displayNames.add(String(rawName).trim());
        if (!entry.locations.includes(locationDesc)) {
            entry.locations.push(locationDesc);
        }
    };

    // 1. Nguồn từ file Danh sách SV & Đồ án
    svList.forEach(sv => {
        if (sv.hoTen && sv.maSV) {
            addRecord(sv.hoTen, sv.maSV, `File SV${sv.lop ? ` (Lớp ${sv.lop})` : ''}`);
        }
    });

    // 2. Nguồn từ SV đã có sẵn trên giao diện (kho chờ hoặc hội đồng)
    document.querySelectorAll('#pool-sv .card-item, .dropzone-sv .card-item').forEach(el => {
        const svName = el.getAttribute('data-name') || '';
        const svId = el.getAttribute('data-id') || '';
        const svLop = el.getAttribute('data-lop') || '';
        if (svName && svId) {
            addRecord(svName, svId, `Kho/HĐ hiện tại${svLop ? ` (Lớp ${svLop})` : ''}`);
        }
    });

    // 3. Nguồn từ file Danh sách Hội đồng upload
    Object.keys(parsedHoidongs || {}).forEach(key => {
        if (key.startsWith('__')) return;
        const hd = parsedHoidongs[key];
        if (!hd || typeof hd !== 'object') return;
        const hdLabel = `${hd.soHD}${hd.ngayBV ? ` (${hd.gioBV ? hd.gioBV + ' ' : ''}${hd.ngayBV})` : ''}`;
        (Array.isArray(hd.thanhVien) ? hd.thanhVien : []).forEach(tv => {
            addRecord(tv.hoTenGV, tv.maGV, `File HĐ: ${hdLabel} - ${tv.chucDanh}`);
        });
    });

    // Lọc các sinh viên có >= 2 Mã SV khác nhau cho cùng 1 tên chuẩn hóa
    const conflicts = [];
    nameMap.forEach((idMap, normName) => {
        if (idMap.size > 1) {
            const variants = [];
            idMap.forEach((info, id) => {
                variants.push({
                    maSV: id,
                    names: Array.from(info.displayNames).join(' / '),
                    locations: info.locations
                });
            });
            conflicts.push({ normName, variants });
        }
    });

    return conflicts;
}

// Rà soát trùng Mã SV trong file Sinh viên, file Hội đồng (và đối chiếu chéo nếu cần)
export function detectDuplicateStudentId({ svList = [], parsedHoidongs = {} }) {
    const conflicts = [];

    // 1. Kiểm tra trùng Mã SV trong nội bộ file Danh sách SV
    const svMap = new Map();
    (svList || []).forEach((sv, idx) => {
        const id = String(sv.maSV || '').trim().toUpperCase();
        if (!id) return;
        const entry = {
            rowIndex: idx + 1,
            hoTen: String(sv.hoTen || '').trim() || '(Chưa rõ tên)',
            lop: String(sv.lop || '').trim(),
            location: `File SV (dòng dữ liệu #${idx + 1})`
        };
        if (!svMap.has(id)) {
            svMap.set(id, [entry]);
        } else {
            svMap.get(id).push(entry);
        }
    });

    svMap.forEach((items, maSV) => {
        if (items.length > 1) {
            conflicts.push({
                source: 'File Danh sách SV',
                maSV,
                items
            });
        }
    });

    // 2. Kiểm tra trùng Mã SV trong file Danh sách Hội đồng
    const hdSvMap = new Map();
    Object.keys(parsedHoidongs || {}).forEach((key) => {
        if (key.startsWith('__')) return;
        const hd = parsedHoidongs[key];
        if (!hd || typeof hd !== 'object') return;
        const hdLabel = `HĐ ${hd.soHD || key}${hd.ngayBV ? ` (${hd.gioBV ? hd.gioBV + ' ' : ''}${hd.ngayBV})` : ''}`;

        (hd.sinhViens || []).forEach((sv, sIdx) => {
            const id = String(sv.maSV || '').trim().toUpperCase();
            if (!id) return;
            const entry = {
                hoTen: String(sv.hoTenSV || '').trim() || '(Chưa rõ tên)',
                lop: String(sv.lop || '').trim(),
                location: `${hdLabel} (SV #${sIdx + 1})`
            };
            if (!hdSvMap.has(id)) {
                hdSvMap.set(id, [entry]);
            } else {
                hdSvMap.get(id).push(entry);
            }
        });
    });

    hdSvMap.forEach((items, maSV) => {
        if (items.length > 1) {
            conflicts.push({
                source: 'File Danh sách Hội đồng',
                maSV,
                items
            });
        }
    });

    return conflicts;
}
