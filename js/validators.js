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

    if (cleanId1 && cleanId2 && !cleanId1.startsWith('GV_') && !cleanId2.startsWith('GV_')) {
        return cleanId1 === cleanId2;
    }

    return isSameTeacherName(name1, name2);
}

// Kiểm tra SV có trùng GVHD với Thành viên HĐ không
export function validateGVHDConflictWithHD(gvhdName, hdBox, maGVHD = '') {
    if (!gvhdName && !maGVHD) return false;
    const hdGVNodes = hdBox.querySelectorAll('.dropzone-gv .card-item');

    for (let gvNode of hdGVNodes) {
        const memberId = (gvNode.getAttribute('data-id') || '').trim();
        const memberName = (gvNode.getAttribute('data-name') || '').trim();

        if (isSameTeacher({ id1: maGVHD, name1: gvhdName }, { id2: memberId, name2: memberName })) {
            return memberName || memberId;
        }
    }
    return false;
}

// TẦNG 2: Kiểm tra 1 Mã GV xuất hiện >= 2 lần (giữ nhiều vai trò) trong CÙNG 1 Hội đồng (Chặn & yêu cầu sửa file)
export function detectDuplicateTeacherIdInSameHD(parsedHoidongs = {}) {
    const conflicts = [];

    Object.keys(parsedHoidongs || {}).forEach(key => {
        if (key.startsWith('__')) return;
        const hd = parsedHoidongs[key];
        if (!hd || typeof hd !== 'object') return;

        const hdLabel = `Khối HĐ #${hd.soHD || key}${hd.ngayBV ? ` (${hd.gioBV ? hd.gioBV + ' ' : ''}${hd.ngayBV})` : ''}`;
        const idMap = new Map(); // maGV -> [{ hoTenGV, chucDanh }]

        (Array.isArray(hd.thanhVien) ? hd.thanhVien : []).forEach(tv => {
            const id = String(tv.maGV || '').trim().toUpperCase();
            if (!id || id.startsWith('GV_')) return;

            if (!idMap.has(id)) {
                idMap.set(id, []);
            }
            idMap.get(id).push({
                hoTenGV: String(tv.hoTenGV || '').trim(),
                chucDanh: String(tv.chucDanh || '').trim()
            });
        });

        idMap.forEach((roles, maGV) => {
            if (roles.length > 1) {
                conflicts.push({
                    hdLabel,
                    maGV,
                    roles
                });
            }
        });
    });

    return conflicts;
}

// Kiểm tra GV được kéo thả có trùng GVHD của bất kỳ SV nào trong HĐ không
export function validateTeacherIsGVHDOfAnyStudent(gvName, hdBox, gvId = '') {
    const svCards = hdBox.querySelectorAll('.dropzone-sv .card-item');
    for (let svNode of svCards) {
        const maGVHDOfSV = (svNode.getAttribute('data-magvhd') || '').trim();
        const gvhdOfSV = (svNode.getAttribute('data-gvhd') || '').trim();

        if (isSameTeacher({ id1: maGVHDOfSV, name1: gvhdOfSV }, { id2: gvId, name2: gvName })) {
            return svNode.getAttribute('data-name');
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

// Chuẩn hóa tên GV: Bỏ học hàm/học vị -> Bỏ dấu tiếng Việt -> Xóa ký tự đặc biệt -> Viết thường
export function normalizeTeacherKeyForDupCheck(rawName) {
    if (!rawName) return '';
    let s = String(rawName).toLowerCase().trim();
    s = s.replace(/\b(gs|pgs|ts|ths|gvc|gv|ks|cn|bs|dsc|prof|dr|msc|mr|ms|mrs)\.?\b/gi, ' ');
    s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd');
    s = s.replace(/[^a-z0-9]/g, '');
    return s;
}

// Chuẩn hóa tên SV: Bỏ dấu tiếng Việt -> Xóa ký tự đặc biệt -> Viết thường
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

// Gom danh sách cảnh báo thiếu Mã GV hoặc thiếu Mã SV (bỏ qua nếu HĐ chưa có SV)
export function collectMissingCodeWarnings({ svList = [], gvList = [], parsedHoidongs = {} }) {
    const warnings = [];

    if (Array.isArray(svList.__missingCodeIssues)) {
        warnings.push(...svList.__missingCodeIssues);
    } else {
        (svList || []).forEach((sv, idx) => {
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
        (gvList || []).forEach((gv, idx) => {
            const id = String(gv.maGV || '').trim();
            if (gv.hoTenGV && (!id || id.toUpperCase().startsWith('GV_'))) {
                warnings.push(`GV #${idx + 1}: "${gv.hoTenGV}" bị thiếu Mã GV.`);
            }
        });
    }

    if (Array.isArray(parsedHoidongs.__missingCodeIssues)) {
        warnings.push(...parsedHoidongs.__missingCodeIssues);
    } else {
        Object.keys(parsedHoidongs || {}).forEach(key => {
            if (key.startsWith('__')) return;
            const hd = parsedHoidongs[key];
            if (!hd || typeof hd !== 'object') return;
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

// Rà soát trùng Mã SV trong file Sinh viên và file Hội đồng (chặn & yêu cầu sửa file)
export function detectDuplicateStudentId({ svList = [], parsedHoidongs = {} }) {
    const conflicts = [];

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

// Đối chiếu chéo tìm Mã GV bị dùng cho >= 2 Giảng viên khác tên chuẩn hóa (Chặn & yêu cầu sửa file)
export function detectDuplicateTeacherIdDifferentName({ svList = [], gvList = [], parsedHoidongs = {} }) {
    // Map: maGV -> Map(normName -> { displayNames: Set, locations: [] })
    const idMap = new Map();

    const addRecord = (rawId, rawName, locationDesc) => {
        const id = String(rawId || '').trim().toUpperCase();
        const normName = normalizeTeacherKeyForDupCheck(rawName);
        if (!id || id.startsWith('GV_') || !normName) return;

        if (!idMap.has(id)) {
            idMap.set(id, new Map());
        }
        const nameGroup = idMap.get(id);
        if (!nameGroup.has(normName)) {
            nameGroup.set(normName, { displayNames: new Set(), locations: [] });
        }
        const entry = nameGroup.get(normName);
        entry.displayNames.add(String(rawName).trim());
        if (!entry.locations.includes(locationDesc)) {
            entry.locations.push(locationDesc);
        }
    };

    // 1. Nguồn từ Danh sách Giảng viên
    (gvList || []).forEach((gv, idx) => {
        if (gv.maGV && gv.hoTenGV) {
            addRecord(gv.maGV, gv.hoTenGV, `File GV (dòng dữ liệu #${idx + 1})`);
        }
    });
    document.querySelectorAll('#pool-gv .card-item').forEach(el => {
        const id = el.getAttribute('data-id') || '';
        const name = el.getAttribute('data-name') || '';
        if (id && name) addRecord(id, name, `Kho GV hiện tại`);
    });

    // 2. Nguồn từ Danh sách Sinh viên & Đồ án (cột Mã GVHD + GVHD)
    (svList || []).forEach(sv => {
        if (sv.maGVHD && sv.gvhd) {
            addRecord(sv.maGVHD, sv.gvhd, `File SV: SV ${sv.hoTen} (${sv.maSV})`);
        }
    });
    document.querySelectorAll('#pool-sv .card-item, .dropzone-sv .card-item').forEach(el => {
        const ma = el.getAttribute('data-magvhd') || '';
        const gvhd = el.getAttribute('data-gvhd') || '';
        const svName = el.getAttribute('data-name') || '';
        const svId = el.getAttribute('data-id') || '';
        if (ma && gvhd) addRecord(ma, gvhd, `Kho/HĐ hiện tại: SV ${svName} (${svId})`);
    });

    // 3. Nguồn từ Danh sách Hội đồng upload
    Object.keys(parsedHoidongs || {}).forEach(key => {
        if (key.startsWith('__')) return;
        const hd = parsedHoidongs[key];
        if (!hd || typeof hd !== 'object') return;
        const hdLabel = `Khối HĐ #${hd.soHD || key}${hd.ngayBV ? ` (${hd.gioBV ? hd.gioBV + ' ' : ''}${hd.ngayBV})` : ''}`;
        (hd.thanhVien || []).forEach(tv => {
            if (tv.maGV && tv.hoTenGV) {
                addRecord(tv.maGV, tv.hoTenGV, `File HĐ: ${hdLabel} - ${tv.chucDanh}`);
            }
        });
    });

    const conflicts = [];
    idMap.forEach((nameGroup, maGV) => {
        if (nameGroup.size > 1) {
            const variants = [];
            nameGroup.forEach((info, normName) => {
                variants.push({
                    normName,
                    names: Array.from(info.displayNames).join(' / '),
                    locations: info.locations
                });
            });
            conflicts.push({ maGV, variants });
        }
    });

    return conflicts;
}

// Đối chiếu chéo tìm GV trùng tên chuẩn hóa nhưng khác Mã GV
export function detectDuplicateTeacherNameDifferentId({ svList = [], gvList = [], parsedHoidongs = {} }) {
    const nameMap = new Map();
    const sameHDMap = new Map(); // normName -> Set(hdLabel)

    const addRecord = (rawName, rawId, locationDesc) => {
        const normName = normalizeTeacherKeyForDupCheck(rawName);
        const id = String(rawId || '').trim().toUpperCase();
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

    (svList || []).forEach(sv => {
        if (sv.gvhd && sv.maGVHD) {
            addRecord(sv.gvhd, sv.maGVHD, `File SV: SV ${sv.hoTen} (${sv.maSV})`);
        }
    });

    document.querySelectorAll('#pool-sv .card-item, .dropzone-sv .card-item').forEach(el => {
        const gvhd = el.getAttribute('data-gvhd') || '';
        const ma = el.getAttribute('data-magvhd') || '';
        const svName = el.getAttribute('data-name') || '';
        const svId = el.getAttribute('data-id') || '';
        if (gvhd && ma) addRecord(gvhd, ma, `Kho/HĐ hiện tại: SV ${svName} (${svId})`);
    });

    (gvList || []).forEach(gv => {
        addRecord(gv.hoTenGV, gv.maGV, `File GV: Danh mục Giảng viên`);
    });

    document.querySelectorAll('#pool-gv .card-item').forEach(el => {
        const name = el.getAttribute('data-name') || '';
        const id = el.getAttribute('data-id') || '';
        if (name && id) addRecord(name, id, `Kho GV hiện tại`);
    });

    Object.keys(parsedHoidongs || {}).forEach(key => {
        if (key.startsWith('__')) return;
        const hd = parsedHoidongs[key];
        if (!hd || typeof hd !== 'object') return;
        const hdLabel = `Khối HĐ #${hd.soHD || key}${hd.ngayBV ? ` (${hd.gioBV ? hd.gioBV + ' ' : ''}${hd.ngayBV})` : ''}`;

        const seenNamesInThisHD = new Set();
        (hd.thanhVien || []).forEach(tv => {
            addRecord(tv.hoTenGV, tv.maGV, `File HĐ: ${hdLabel} - ${tv.chucDanh}`);
            const norm = normalizeTeacherKeyForDupCheck(tv.hoTenGV);
            if (norm) {
                if (seenNamesInThisHD.has(norm)) {
                    if (!sameHDMap.has(norm)) sameHDMap.set(norm, new Set());
                    sameHDMap.get(norm).add(hdLabel);
                }
                seenNamesInThisHD.add(norm);
            }
        });
    });

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
            conflicts.push({
                normName,
                variants,
                sameHDLabels: sameHDMap.has(normName) ? Array.from(sameHDMap.get(normName)) : []
            });
        }
    });

    return conflicts;
}

// Đối chiếu chéo tìm SV trùng tên chuẩn hóa nhưng khác Mã SV
export function detectDuplicateStudentNameDifferentId({ svList = [], parsedHoidongs = {} }) {
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

    (svList || []).forEach(sv => {
        if (sv.hoTen && sv.maSV) {
            addRecord(sv.hoTen, sv.maSV, `File SV${sv.lop ? ` (Lớp ${sv.lop})` : ''}`);
        }
    });

    document.querySelectorAll('#pool-sv .card-item, .dropzone-sv .card-item').forEach(el => {
        const svName = el.getAttribute('data-name') || '';
        const svId = el.getAttribute('data-id') || '';
        const svLop = el.getAttribute('data-lop') || '';
        if (svName && svId) {
            addRecord(svName, svId, `Kho/HĐ hiện tại${svLop ? ` (Lớp ${svLop})` : ''}`);
        }
    });

    Object.keys(parsedHoidongs || {}).forEach(key => {
        if (key.startsWith('__')) return;
        const hd = parsedHoidongs[key];
        if (!hd || typeof hd !== 'object') return;
        const hdLabel = `HĐ ${hd.soHD}${hd.ngayBV ? ` (${hd.gioBV ? hd.gioBV + ' ' : ''}${hd.ngayBV})` : ''}`;
        (hd.sinhViens || []).forEach(sv => {
            if (sv.hoTenSV && sv.maSV) {
                addRecord(sv.hoTenSV, sv.maSV, `File HĐ: ${hdLabel}${sv.lop ? ` - Lớp ${sv.lop}` : ''}`);
            }
        });
    });

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
