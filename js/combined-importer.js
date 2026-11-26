// js/combined-importer.js
// Module độc lập xử lý luồng nhập 2 file: Danh sách GV (tùy chọn) + Danh sách Phân công tổng hợp (SV, GVHD, HĐ)

import { renderHoidongFromExcel } from './hoidong-manager.js';
import { findNganhInfoByLop } from './excel-reader.js';
import { validateStudentsAgainstSchoolMaster } from './student-master-validator.js';

/**
 * Chuẩn hóa chuỗi khóa tra cứu Giảng viên:
 * - Đưa về Unicode NFC, chữ thường, rút gọn khoảng trắng
 * - Loại bỏ học hàm / học vị ở đầu (PGS, GS, TSKH, TS, ThS, Th.S, GVC, GV, KS, CN...)
 * - Bỏ dấu ngoặc đơn (nếu có) nhưng GIỮ NGUYÊN ký hiệu phân biệt như A, B
 *   VD: "ThS. Nguyễn Đức Hoàng (A)" hoặc "Nguyễn Đức Hoàng A" -> "nguyễn đức hoàng a"
 */
export function buildStrictTeacherLookupKey(rawName) {
    let s = String(rawName || '').normalize('NFC').trim();
    if (!s) return '';

    s = s.replace(
        /^(?:PGS\.?\s*TSKH\.?|PGS\.?\s*TS\.?|GS\.?\s*TSKH\.?|GS\.?\s*TS\.?|TSKH\.?|TS\.?|Th\.?\s*S\.?|GVC\.?|GV\.?|KS\.?|CN\.?|BS\.?|Dr\.?|Prof\.?)+\s*/i,
        ''
    );

    s = s
        .replace(/[()]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();

    return s;
}

/**
 * Tách tên gốc (không kèm ký hiệu phân biệt A/B/(A)/(B)/(1)/(2) ở cuối) và phần ký hiệu phân biệt
 * Dùng để phát hiện trường hợp cùng 1 họ tên nhưng lúc có ký hiệu lúc không có ký hiệu
 */
function splitTeacherBaseAndTag(rawName) {
    const key = buildStrictTeacherLookupKey(rawName);
    if (!key) return { baseKey: '', tag: '' };
    const m = key.match(/^(.*?)\s+([a-z0-9]{1,3})$/i);
    if (m && m[1].trim().includes(' ')) {
        return { baseKey: m[1].trim(), tag: m[2].toLowerCase() };
    }
    return { baseKey: key, tag: '' };
}

/**
 * Làm sạch chuỗi hiển thị tên GV (giữ nguyên Học hàm/Học vị và ký hiệu (A)/(B), chỉ bỏ khoảng trắng thừa)
 */
function cleanDisplayTeacherName(rawName) {
    return String(rawName || '')
        .normalize('NFC')
        .replace(/\s*\(\s*/g, ' (')
        .replace(/\s*\)\s*/g, ')')
        .replace(/\s+/g, ' ')
        .trim();
}

function escapeHtml(str) {
    return String(str ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function normalizeHeaderCell(val) {
    return String(val || '')
        .normalize('NFC')
        .replace(/\s+/g, ' ')
        .trim()
        .toUpperCase();
}

function detectNganhFromLop(lop) {
    const u = String(lop || '').toUpperCase().replace(/\s+/g, '');
    if (u.includes('CKO') || u.includes('CK0')) return 'Công nghệ Kỹ thuật Ô tô';
    if (u.includes('CDO') || u.includes('CD0')) return 'Điện cơ Ô tô';
    if (u.includes('CCM') || u.includes('CKM')) return 'Công nghệ Chế tạo máy';
    if (u.includes('COD')) return 'Công nghệ ô tô điện';
    if (u.includes('CTP')) return 'Công nghệ Thực phẩm';
    if (u.includes('SHU')) return 'Công nghệ Sinh học';
    if (u.includes('KMT')) return 'Quản lý Tài nguyên & Môi trường';
    return '';
}

function normalizeExcelDateValue(raw) {
    if (raw === null || raw === undefined || raw === '') return '';
    if (raw instanceof Date && !isNaN(raw.getTime())) {
        const d = String(raw.getDate()).padStart(2, '0');
        const m = String(raw.getMonth() + 1).padStart(2, '0');
        const y = raw.getFullYear();
        return `${d}/${m}/${y}`;
    }
    if (typeof raw === 'number' && raw > 20000) {
        const utcDays = Math.floor(raw - 25569);
        const dt = new Date(utcDays * 86400 * 1000);
        const d = String(dt.getUTCDate()).padStart(2, '0');
        const m = String(dt.getUTCMonth() + 1).padStart(2, '0');
        const y = dt.getUTCFullYear();
        return `${d}/${m}/${y}`;
    }
    const s = String(raw).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
        const [y, m, d] = s.split('-');
        return `${d}/${m}/${y}`;
    }
    return s;
}

function normalizeExcelTimeValue(raw) {
    if (raw === null || raw === undefined || raw === '') return '';
    if (raw instanceof Date && !isNaN(raw.getTime())) {
        const hh = String(raw.getHours()).padStart(2, '0');
        const mm = String(raw.getMinutes()).padStart(2, '0');
        return `${hh}:${mm}`;
    }
    if (typeof raw === 'number' && raw >= 0 && raw < 1) {
        const totalMins = Math.round(raw * 24 * 60);
        const hh = String(Math.floor(totalMins / 60)).padStart(2, '0');
        const mm = String(totalMins % 60).padStart(2, '0');
        return `${hh}:${mm}`;
    }
    return String(raw).trim();
}

function detectSessionPeriod(gioStr) {
    const clean = String(gioStr || '').toLowerCase().replace('h', ':').trim();
    const hour = parseInt(clean.split(':')[0], 10);
    if (isNaN(hour)) return 'SANG';
    if (hour < 12) return 'SANG';
    if (hour < 17) return 'CHIEU';
    return 'TOI';
}

function formatSessionLabel(sessionCode) {
    if (sessionCode === 'SANG') return 'Buổi Sáng';
    if (sessionCode === 'CHIEU') return 'Buổi Chiều';
    return 'Buổi Tối';
}

function parseMinutesForSort(gioStr) {
    const clean = String(gioStr || '').toLowerCase().replace('h', ':').trim();
    const m = clean.match(/(\d{1,2})\s*:\s*(\d{1,2})?/);
    if (!m) return 9999;
    return parseInt(m[1], 10) * 60 + parseInt(m[2] || '0', 10);
}

function parseVNDateTimestamp(dateStr) {
    const p = String(dateStr || '').split('/');
    if (p.length !== 3) return Number.MAX_SAFE_INTEGER;
    return new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0])).getTime();
}

async function readFirstSheetRowsFromFile(file, inputElementToReset = null) {
    if (!file) {
        throw new Error('Không tìm thấy dữ liệu tệp tin đầu vào.');
    }
    let buffer;
    try {
        buffer = await file.arrayBuffer();
    } catch (err) {
        if (inputElementToReset) {
            try { inputElementToReset.value = ''; } catch { /* ignore */ }
        }
        const isLockOrModified =
            err?.name === 'NotReadableError' ||
            String(err?.message || '').includes('NotReadableError') ||
            String(err?.message || '').includes('could not be read');

        if (isLockOrModified) {
            throw new Error(
                `Không thể đọc file "${file.name}" do file vừa bị chỉnh sửa/lưu đè hoặc đang bị khóa trên máy tính. Vui lòng chọn lại file "${file.name}" và thử lại!`
            );
        }
        throw new Error(`Lỗi đọc file "${file.name}": ${err?.message || err}`);
    }

    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
    const firstSheetName = wb.SheetNames[0];
    const ws = wb.Sheets[firstSheetName];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
}

/**
 * 1. Đọc & kiểm tra nội bộ File Danh sách Giảng viên (Mã GV | Họ và Tên | Đơn vị Công tác)
 */
export async function parseStrictTeacherMasterFile(file, inputElementToReset = null) {
    const rows = await readFirstSheetRowsFromFile(file, inputElementToReset);
    const lookupMap = new Map();
    const internalErrors = [];
    const seenIds = new Map();
    const rawList = [];

    let headerRowIdx = -1;
    let colMa = 0, colTen = 1, colDonVi = 2;

    for (let i = 0; i < Math.min(rows.length, 10); i++) {
        const rUpper = (rows[i] || []).map(normalizeHeaderCell);
        const idxMa = rUpper.findIndex(c => c.includes('MÃ GV') || c.includes('MÃ SỐ GV'));
        const idxTen = rUpper.findIndex(c => c.includes('HỌ VÀ TÊN') || c.includes('HỌ TÊN'));
        if (idxMa !== -1 && idxTen !== -1) {
            headerRowIdx = i;
            colMa = idxMa;
            colTen = idxTen;
            const idxDV = rUpper.findIndex(c => c.includes('ĐƠN VỊ') || c.includes('KHOA') || c.includes('BỘ MÔN'));
            colDonVi = idxDV !== -1 ? idxDV : idxTen + 1;
            break;
        }
    }

    const startRow = headerRowIdx >= 0 ? headerRowIdx + 1 : 0;

    for (let r = startRow; r < rows.length; r++) {
        const row = rows[r] || [];
        const maGV = String(row[colMa] || '').trim();
        const hoTenRaw = cleanDisplayTeacherName(row[colTen]);
        const donVi = String(row[colDonVi] || '').trim() || 'Đại học Duy Tân';

        if (!maGV && !hoTenRaw) continue;
        if (normalizeHeaderCell(maGV).includes('MÃ GV')) continue;

        if (!maGV || !hoTenRaw) {
            internalErrors.push(`• Dòng ${r + 1} (File Danh sách GV): Thiếu ${!maGV ? 'Mã GV' : 'Họ và tên'} ("${maGV || hoTenRaw}").`);
            continue;
        }

        const maUpper = maGV.toUpperCase();
        if (seenIds.has(maUpper)) {
            internalErrors.push(`• Dòng ${r + 1} (File Danh sách GV): Trùng Mã GV "${maGV}" với dòng ${seenIds.get(maUpper)}.`);
        } else {
            seenIds.set(maUpper, r + 1);
        }

        const lookupKey = buildStrictTeacherLookupKey(hoTenRaw);
        if (lookupMap.has(lookupKey)) {
            const prev = lookupMap.get(lookupKey);
            internalErrors.push(
                `• Dòng ${r + 1} (File Danh sách GV): Trùng tên giảng viên "${hoTenRaw}" với dòng ${prev.rowNum} ("${prev.hoTenGoc}"). Vui lòng thêm ký hiệu phân biệt (A), (B) vào file Danh sách GV!`
            );
        } else {
            const record = { maGV, hoTenGoc: hoTenRaw, donVi, rowNum: r + 1 };
            lookupMap.set(lookupKey, record);
            rawList.push(record);
        }
    }

    if (rawList.length === 0 && internalErrors.length === 0) {
        internalErrors.push('• File Danh sách GV không có dữ liệu giảng viên hợp lệ!');
    }

    return { lookupMap, rawList, internalErrors };
}

/**
 * 2. Đọc & đối soát File Phân công tổng hợp:
 * - Hỗ trợ 2 chế độ:
 *   (A) Có teacherLookupMap (đã tải File Danh sách GV): đối soát chặt chẽ Mã GV & Đơn vị.
 *   (B) Không có teacherLookupMap (bỏ qua File Danh sách GV): tự trích xuất GV từ tên + ký hiệu phân biệt,
 *       kiểm tra mâu thuẫn ký hiệu (có (A)/(B) nhưng nơi khác lại thiếu ký hiệu) và phát hiện nghi vấn trùng tên chấm cùng ca.
 */
export async function parseAndValidateCombinedAssignment(fileOrRows, teacherLookupMap = null, inputElementToReset = null) {
    const rows = Array.isArray(fileOrRows)
        ? fileOrRows
        : await readFirstSheetRowsFromFile(fileOrRows, inputElementToReset);

    const hasMasterList = teacherLookupMap instanceof Map && teacherLookupMap.size > 0;
    const blockingErrors = [];
    const consistencyWarnings = [];

    // Dò dòng tiêu đề
    let headerIdx = -1;
    let colMap = null;

    for (let i = 0; i < Math.min(rows.length, 15); i++) {
        const rUpper = (rows[i] || []).map(normalizeHeaderCell);
        const idxMSSV = rUpper.findIndex(c => c === 'MSSV' || c === 'MÃ SV' || c === 'MÃ SỐ SV' || c === 'MÃ SINH VIÊN');
        const idxGVHD = rUpper.findIndex(c => c.includes('GV HƯỚNG DẪN') || c === 'GVHD');

        if (idxMSSV !== -1 && idxGVHD !== -1) {
            headerIdx = i;
            const idxHoTen = rUpper.findIndex(c => c.includes('HỌ VÀ TÊN') || c.includes('HỌ TÊN'));
            const idxLop = rUpper.findIndex(c => c === 'LỚP' || c.includes('LỚP SH'));
            const idxDeTai = rUpper.findIndex(c => c.includes('ĐỀ TÀI'));

            const isSplitNameCol = idxHoTen !== -1 && idxLop !== -1 && (idxLop - idxHoTen === 2) && !rUpper[idxHoTen + 1];

            const findColAfterGVHD = (keywords, fallbackOffset) => {
                for (let c = idxGVHD + 1; c < rUpper.length; c++) {
                    if (keywords.some(kw => rUpper[c].includes(kw))) return c;
                }
                return idxGVHD + fallbackOffset;
            };

            const idxPB = findColAfterGVHD(['PHẢN BIỆN', 'GVPB'], 1);
            const idxCT = findColAfterGVHD(['CHỦ TỊCH'], 2);
            const idxTK = findColAfterGVHD(['THƯ KÝ'], 3);
            const idxHD = findColAfterGVHD(['HỘI ĐỒNG'], 4);
            const idxNgay = findColAfterGVHD(['NGÀY BẢO VỆ', 'NGÀY BV'], 5);
            const idxGio = findColAfterGVHD(['GIỜ BẢO VỆ', 'GIỜ BV', 'GIỜ'], 6);
            const idxPhong = findColAfterGVHD(['PHÒNG'], 7);
            const idxCoSo = findColAfterGVHD(['CƠ SỞ', 'ĐỊA ĐIỂM'], 8);

            colMap = {
                idxMSSV,
                idxHoTen,
                isSplitNameCol,
                idxLop,
                idxDeTai,
                idxGVHD,
                idxPB,
                idxCT,
                idxTK,
                idxHD,
                idxNgay,
                idxGio,
                idxPhong,
                idxCoSo
            };
            break;
        }
    }

    if (!colMap) {
        return {
            ok: false,
            hasMasterList,
            blockingErrors: ['• Không tìm thấy dòng tiêu đề hợp lệ (cần có cột MSSV / Mã SV và cột GV HƯỚNG DẪN) trong file Phân công tổng hợp!'],
            consistencyWarnings: [],
            parsedHoidongs: {},
            studentList: [],
            teacherPoolList: [],
            detectedKhoas: [],
            detectedNganhs: []
        };
    }

    // Nếu chạy chế độ không có File Danh sách GV: tiền quét toàn bộ tên GV để phát hiện mâu thuẫn ký hiệu
    // VD: dòng 3 ghi "Nguyễn Đức Hoàng (A)" nhưng dòng 15 ghi "Nguyễn Đức Hoàng" (thiếu ký hiệu)
    const autoTeacherMap = new Map(); // lookupKey -> { maGV, hoTenGV, donVi }
    let autoTeacherSeq = 1;

    if (!hasMasterList) {
        const baseNameStats = new Map(); // baseKey -> { taggedNames: Set, untaggedRows: Array }
        for (let r = headerIdx + 1; r < rows.length; r++) {
            const row = rows[r] || [];
            const mssv = String(row[colMap.idxMSSV] || '').trim();
            const upperLine = row.map(normalizeHeaderCell).join(' ');
            if (
                upperLine.includes('HIỆU TRƯỞNG') ||
                upperLine.includes('TRƯỞNG KHOA') ||
                upperLine.includes('BAN ĐÀO TẠO') ||
                upperLine.includes('GIÁM ĐỐC ĐẠI HỌC')
            ) {
                break;
            }
            if (!mssv && !String(row[colMap.idxHoTen] || '').trim()) continue;

            const roleCols = [
                { col: colMap.idxGVHD, label: 'GV Hướng dẫn' },
                { col: colMap.idxPB, label: 'Phản biện' },
                { col: colMap.idxCT, label: 'Chủ tịch' },
                { col: colMap.idxTK, label: 'Thư ký' }
            ];

            roleCols.forEach(({ col, label }) => {
                const disp = cleanDisplayTeacherName(row[col]);
                if (!disp) return;
                const { baseKey, tag } = splitTeacherBaseAndTag(disp);
                if (!baseKey) return;
                if (!baseNameStats.has(baseKey)) {
                    baseNameStats.set(baseKey, { taggedExamples: new Set(), untaggedLocations: [] });
                }
                const st = baseNameStats.get(baseKey);
                if (tag) {
                    st.taggedExamples.add(disp);
                } else {
                    st.untaggedLocations.push(`Dòng ${r + 1} (MSSV ${mssv || 'N/A'} — ${label}): "${disp}"`);
                }
            });
        }

        baseNameStats.forEach((st) => {
            if (st.taggedExamples.size > 0 && st.untaggedLocations.length > 0) {
                const examples = Array.from(st.taggedExamples).join(', ');
                st.untaggedLocations.forEach(loc => {
                    blockingErrors.push(
                        `• ${loc} → Tên giảng viên này có phân biệt ký hiệu ở dòng khác (${examples}) nhưng tại đây lại thiếu ký hiệu phân biệt!`
                    );
                });
            }
        });
    }

    const resolveTeacher = (rawCellVal, rowNumber, mssv, roleLabel) => {
        const displayName = cleanDisplayTeacherName(rawCellVal);
        if (!displayName) return null;

        const key = buildStrictTeacherLookupKey(displayName);

        if (hasMasterList) {
            const found = teacherLookupMap.get(key);
            if (!found) {
                const { tag } = splitTeacherBaseAndTag(displayName);
                const hint = tag
                    ? `Không tìm thấy giảng viên mang ký hiệu phân biệt này trong Danh sách GV (cần khớp cả ký hiệu (A)/(B)).`
                    : `Tên giảng viên chưa có trong file Danh sách GV hoặc thiếu ký hiệu phân biệt (A)/(B).`;
                blockingErrors.push(
                    `• Dòng ${rowNumber} (MSSV: ${mssv || 'N/A'} — ${roleLabel}): "${displayName}" → ${hint}`
                );
                return null;
            }
            return {
                maGV: found.maGV,
                hoTenGV: displayName,
                donVi: found.donVi
            };
        }

        // Chế độ không có File Danh sách GV: cấp mã tạm nội bộ AUTO_GV_xx theo khóa chuẩn hóa tên + ký hiệu
        if (!autoTeacherMap.has(key)) {
            const tempId = `AUTO_GV_${String(autoTeacherSeq++).padStart(3, '0')}`;
            autoTeacherMap.set(key, {
                maGV: tempId,
                hoTenGV: displayName,
                donVi: 'Đại học Duy Tân'
            });
        } else {
            const existing = autoTeacherMap.get(key);
            if (displayName.length > existing.hoTenGV.length) {
                existing.hoTenGV = displayName;
            }
        }
        return { ...autoTeacherMap.get(key) };
    };

    const studentList = [];
    const memberGroupMap = new Map();
    const usedTeachersMap = new Map();
    const detectedKhoaList = [];
    const detectedNganhList = [];

    const registerTeacherToPool = (tObj) => {
        if (!tObj || !tObj.maGV) return;
        const idKey = tObj.maGV.toUpperCase();
        const existing = usedTeachersMap.get(idKey);
        if (!existing || tObj.hoTenGV.length > existing.hoTenGV.length) {
            usedTeachersMap.set(idKey, { ...tObj });
        }
    };

    if (hasMasterList) {
        teacherLookupMap.forEach(rec => {
            usedTeachersMap.set(rec.maGV.toUpperCase(), {
                maGV: rec.maGV,
                hoTenGV: rec.hoTenGoc,
                donVi: rec.donVi
            });
        });
    }

    for (let r = headerIdx + 1; r < rows.length; r++) {
        const row = rows[r] || [];
        const mssv = String(row[colMap.idxMSSV] || '').trim();
        const rawHoTen1 = String(row[colMap.idxHoTen] || '').trim();
        const rawHoTen2 = colMap.isSplitNameCol ? String(row[colMap.idxHoTen + 1] || '').trim() : '';
        const hoTenSV = colMap.isSplitNameCol
            ? `${rawHoTen1} ${rawHoTen2}`.replace(/\s+/g, ' ').trim()
            : rawHoTen1.replace(/\s+/g, ' ').trim();

        const upperLine = row.map(normalizeHeaderCell).join(' ');
        if (
            upperLine.includes('HIỆU TRƯỞNG') ||
            upperLine.includes('TRƯỞNG KHOA') ||
            upperLine.includes('BAN ĐÀO TẠO') ||
            upperLine.includes('GIÁM ĐỐC ĐẠI HỌC')
        ) {
            break;
        }

        if (!mssv && !hoTenSV) continue;

        const lop = String(row[colMap.idxLop] || '').replace(/\s+/g, '').trim();
        const deTai = String(row[colMap.idxDeTai] || '').replace(/\s+/g, ' ').trim();
        const rawGVHD = row[colMap.idxGVHD];
        const rawPB = row[colMap.idxPB];
        const rawCT = row[colMap.idxCT];
        const rawTK = row[colMap.idxTK];
        const rawHD = String(row[colMap.idxHD] || '').trim();
        const ngayBV = normalizeExcelDateValue(row[colMap.idxNgay]);
        const gioBV = normalizeExcelTimeValue(row[colMap.idxGio]);
        const phongBV = String(row[colMap.idxPhong] || '').trim();
        const diaDiemBV = String(row[colMap.idxCoSo] || '').trim();

        const rowNum = r + 1;

        if (!mssv || !hoTenSV) {
            blockingErrors.push(`• Dòng ${rowNum}: Thiếu ${!mssv ? 'MSSV' : 'Họ và tên sinh viên'}.`);
            continue;
        }

        const gvhdObj = resolveTeacher(rawGVHD, rowNum, mssv, 'GV Hướng dẫn');
        if (!cleanDisplayTeacherName(rawGVHD)) {
            blockingErrors.push(`• Dòng ${rowNum} (MSSV: ${mssv}): Bị trống tên GV Hướng dẫn.`);
        }

        const pbObj = resolveTeacher(rawPB, rowNum, mssv, 'Phản biện');
        const ctObj = resolveTeacher(rawCT, rowNum, mssv, 'Chủ tịch');
        const tkObj = resolveTeacher(rawTK, rowNum, mssv, 'Thư ký');

        if (gvhdObj) registerTeacherToPool(gvhdObj);
        if (pbObj) registerTeacherToPool(pbObj);
        if (ctObj) registerTeacherToPool(ctObj);
        if (tkObj) registerTeacherToPool(tkObj);

        const nganhInfo = typeof findNganhInfoByLop === 'function' ? findNganhInfoByLop(lop) : null;
        const nganh = (nganhInfo?.chuyenNganh || detectNganhFromLop(lop) || '').trim();
        const khoa = (nganhInfo?.khoa || '').trim();

        if (khoa && !detectedKhoaList.includes(khoa)) {
            detectedKhoaList.push(khoa);
        }
        if (nganh && !detectedNganhList.includes(nganh)) {
            detectedNganhList.push(nganh);
        }

        const svRecord = {
            maSV: mssv,
            hoTenSV,
            lop,
            deTai,
            maGVHD: gvhdObj?.maGV || '',
            gvhd: gvhdObj?.hoTenGV || cleanDisplayTeacherName(rawGVHD),
            nganh,
            ngayBV,
            gioBV,
            phongBV,
            diaDiemBV
        };
        studentList.push(svRecord);

        // Gom nhóm thành 1 Hội đồng theo BỘ 3 THÀNH VIÊN + CHỨC DANH (CT + PB + TK)
        if (rawHD || ctObj || pbObj || tkObj) {
            const hdNumClean = rawHD
                .replace(/^(?:hội\s*đồng|hoi\s*dong|hđ|hd)\s*(?:số|so|#|:|-|\.)?\s*/i, '')
                .trim();

            const ctKey = ctObj?.maGV?.toUpperCase() || '';
            const pbKey = pbObj?.maGV?.toUpperCase() || '';
            const tkKey = tkObj?.maGV?.toUpperCase() || '';
            const hasAnyMember = Boolean(ctKey || pbKey || tkKey);

            const memberGroupKey = hasAnyMember
                ? `MEM__${ctKey}|${pbKey}|${tkKey}`
                : `NO_MEM__${hdNumClean || `ROW_${rowNum}`}`;

            if (!memberGroupMap.has(memberGroupKey)) {
                memberGroupMap.set(memberGroupKey, {
                    memberGroupKey,
                    ct: ctObj,
                    pb: pbObj,
                    tk: tkObj,
                    rawSoHDs: [],
                    firstRowNum: rowNum,
                    rows: []
                });
            }

            const mGroup = memberGroupMap.get(memberGroupKey);
            if (!mGroup.ct && ctObj) mGroup.ct = ctObj;
            if (!mGroup.pb && pbObj) mGroup.pb = pbObj;
            if (!mGroup.tk && tkObj) mGroup.tk = tkObj;

            if (hdNumClean && !mGroup.rawSoHDs.some(x => x.toUpperCase() === hdNumClean.toUpperCase())) {
                mGroup.rawSoHDs.push(hdNumClean);
            }

            mGroup.rows.push({
                rowNum,
                hdNumRaw: hdNumClean,
                sv: svRecord
            });
        }
    }

    // Đối soát toàn bộ Mã SV & Họ tên SV với DANH_SACH_SV_TOAN_TRUONG.xlsx
    const schoolCheck = await validateStudentsAgainstSchoolMaster(
        studentList.map((sv, idx) => ({
            maSV: sv.maSV,
            hoTen: sv.hoTenSV,
            location: `MSSV ${sv.maSV}`
        }))
    );

    if (!schoolCheck.ok) {
        blockingErrors.push(...schoolCheck.errors);
    } else if (schoolCheck.skipped && schoolCheck.warning) {
        consistencyWarnings.unshift(`• ${schoolCheck.warning}`);
    }

    if (blockingErrors.length > 0) {
        return {
            ok: false,
            hasMasterList,
            blockingErrors,
            consistencyWarnings,
            parsedHoidongs: {},
            studentList: [],
            teacherPoolList: [],
            detectedKhoas: [],
            detectedNganhs: []
        };
    }

    const memberGroups = Array.from(memberGroupMap.values());

    let maxNumericSoHD = 0;
    memberGroups.forEach(g => {
        g.rawSoHDs.forEach(s => {
            const n = parseInt(s, 10);
            if (!isNaN(n) && n > maxNumericSoHD) maxNumericSoHD = n;
        });
    });

    const allocatedSoHDs = new Set();

    memberGroups.forEach(g => {
        const memberDesc = [
            g.ct ? `CT: ${g.ct.hoTenGV}` : '',
            g.pb ? `PB: ${g.pb.hoTenGV}` : '',
            g.tk ? `TK: ${g.tk.hoTenGV}` : ''
        ].filter(Boolean).join(', ');

        let preferredSoHD = '';
        if (g.rawSoHDs.length === 1) {
            preferredSoHD = g.rawSoHDs[0];
        } else if (g.rawSoHDs.length > 1) {
            const sortedCandidates = [...g.rawSoHDs].sort((a, b) => {
                const na = parseInt(a, 10);
                const nb = parseInt(b, 10);
                if (!isNaN(na) && !isNaN(nb)) return na - nb;
                return 0;
            });
            preferredSoHD = sortedCandidates[0];
            consistencyWarnings.push(
                `• Gộp Hội đồng (${memberDesc}): Trong Excel các dòng của bộ thành viên này đang ghi nhiều số HĐ khác nhau (${g.rawSoHDs.map(x => `HĐ ${x}`).join(', ')}) → Hệ thống đã gộp chung thành 1 Hội đồng và giữ số HĐ ${preferredSoHD}.`
            );
        }

        const normKey = preferredSoHD.toUpperCase();
        if (preferredSoHD && !allocatedSoHDs.has(normKey)) {
            g.finalSoHD = preferredSoHD;
            allocatedSoHDs.add(normKey);
        } else {
            let nextNum = Math.max(maxNumericSoHD + 1, allocatedSoHDs.size + 1);
            while (allocatedSoHDs.has(String(nextNum))) {
                nextNum++;
            }
            maxNumericSoHD = Math.max(maxNumericSoHD, nextNum);
            g.finalSoHD = String(nextNum);
            allocatedSoHDs.add(g.finalSoHD);

            if (preferredSoHD) {
                consistencyWarnings.push(
                    `• Đổi số Hội đồng trùng (dòng ${g.firstRowNum}): Bộ thành viên (${memberDesc}) bị ghi trùng số "HĐ ${preferredSoHD}" với một Hội đồng khác thành viên đã có trước đó → Hệ thống tự động đổi nhóm này thành HĐ ${g.finalSoHD}.`
                );
            } else {
                consistencyWarnings.push(
                    `• Cấp số Hội đồng (dòng ${g.firstRowNum}): Bộ thành viên (${memberDesc}) chưa có số HĐ trong Excel → Hệ thống tự động gắn số HĐ ${g.finalSoHD}.`
                );
            }
        }
    });

    memberGroups.sort((a, b) => {
        const na = parseInt(a.finalSoHD, 10);
        const nb = parseInt(b.finalSoHD, 10);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return String(a.finalSoHD).localeCompare(String(b.finalSoHD), 'vi');
    });

    // Tách box hiển thị theo suất chấm (Buổi + Ngày + Phòng + Cơ sở) và giữ nguyên cùng finalSoHD
    const renderedSlots = [];

    memberGroups.forEach(g => {
        const slotMap = new Map();

        g.rows.forEach(({ sv }) => {
            const sessionCode = detectSessionPeriod(sv.gioBV);
            const dateNorm = (sv.ngayBV || '').trim();
            const roomNorm = (sv.phongBV || '').trim().toUpperCase();
            const campusNorm = (sv.diaDiemBV || '').trim().toUpperCase();

            const slotKey = `${dateNorm}__${sessionCode}__${roomNorm}__${campusNorm}`;

            if (!slotMap.has(slotKey)) {
                slotMap.set(slotKey, {
                    soHD: g.finalSoHD,
                    ngayBV: sv.ngayBV,
                    gioBV: sv.gioBV,
                    sessionCode,
                    phongBV: sv.phongBV,
                    diaDiemBV: sv.diaDiemBV,
                    ct: g.ct,
                    pb: g.pb,
                    tk: g.tk,
                    sinhViens: []
                });
            }

            const slotRec = slotMap.get(slotKey);
            if (sv.gioBV) {
                if (!slotRec.gioBV || parseMinutesForSort(sv.gioBV) < parseMinutesForSort(slotRec.gioBV)) {
                    slotRec.gioBV = sv.gioBV;
                }
            }
            slotRec.sinhViens.push(sv);
        });

        const groupSlots = Array.from(slotMap.values()).sort((a, b) => {
            const da = parseVNDateTimestamp(a.ngayBV);
            const db = parseVNDateTimestamp(b.ngayBV);
            if (da !== db) return da - db;
            const ta = parseMinutesForSort(a.gioBV);
            const tb = parseMinutesForSort(b.gioBV);
            if (ta !== tb) return ta - tb;
            return String(a.phongBV).localeCompare(String(b.phongBV), 'vi');
        });

        if (groupSlots.length > 1) {
            const slotSummary = groupSlots
                .map(s => `${formatSessionLabel(s.sessionCode)}${s.ngayBV ? ` ${s.ngayBV}` : ''}${s.phongBV ? ` (${s.phongBV})` : ''}: ${s.sinhViens.length} SV`)
                .join('; ');
            consistencyWarnings.push(
                `• Tách suất chấm HĐ ${g.finalSoHD}: Hội đồng này chấm ở ${groupSlots.length} buổi/ngày/phòng khác nhau → Đã tách thành ${groupSlots.length} khung chấm cùng mang số HĐ ${g.finalSoHD} (${slotSummary}).`
            );
        }

        renderedSlots.push(...groupSlots);
    });

    // Kiểm tra cảnh báo trùng lịch chấm cùng ca (Ngày + Buổi) giữa các Hội đồng khác nhau
    // Đặc biệt hữu ích khi không có Danh sách GV để phát hiện 2 người trùng tên nhưng quên gắn (A)/(B)
    const scheduleClashMap = new Map(); // "ngay|session|maGV" -> Array<{ soHD, phongBV, hoTenGV, role }>
    renderedSlots.forEach(slot => {
        if (!slot.ngayBV) return;
        const members = [
            slot.ct ? { ...slot.ct, role: 'Chủ tịch' } : null,
            slot.pb ? { ...slot.pb, role: 'Phản biện' } : null,
            slot.tk ? { ...slot.tk, role: 'Thư ký' } : null
        ].filter(Boolean);

        members.forEach(m => {
            const clashKey = `${slot.ngayBV}__${slot.sessionCode}__${m.maGV.toUpperCase()}`;
            if (!scheduleClashMap.has(clashKey)) {
                scheduleClashMap.set(clashKey, []);
            }
            scheduleClashMap.get(clashKey).push({
                soHD: slot.soHD,
                phongBV: slot.phongBV,
                hoTenGV: m.hoTenGV,
                role: m.role,
                ngayBV: slot.ngayBV,
                sessionLabel: formatSessionLabel(slot.sessionCode)
            });
        });
    });

    scheduleClashMap.forEach((entries) => {
        const distinctHDs = new Set(entries.map(e => `${e.soHD}_${e.phongBV}`));
        if (distinctHDs.size > 1) {
            const first = entries[0];
            const places = entries.map(e => `HĐ ${e.soHD} (${e.role}${e.phongBV ? `, Phòng ${e.phongBV}` : ''})`).join(' và ');
            const extraHint = !hasMasterList
                ? ' (Nếu đây là 2 giảng viên khác nhau trùng họ tên, hãy thêm ký hiệu (A)/(B) hoặc tải kèm File Danh sách GV!)'
                : '';
            consistencyWarnings.unshift(
                `• ⚠️ Trùng lịch chấm cùng ca: GV "${first.hoTenGV}" xuất hiện tại ${places} vào ${first.sessionLabel} ngày ${first.ngayBV}${extraHint}`
            );
        }
    });

    const parsedHoidongs = {
        __preserveExplicitNumbers: true,
        __skipAutoMergeInRenderer: true,
        __mergeLogs: [],
        __missingCodeIssues: []
    };

    renderedSlots.forEach((slot, idx) => {
        const thanhVien = [];
        if (slot.ct) thanhVien.push({ ...slot.ct, chucDanh: 'Chủ tịch' });
        if (slot.pb) thanhVien.push({ ...slot.pb, chucDanh: 'Phản biện' });
        if (slot.tk) thanhVien.push({ ...slot.tk, chucDanh: 'Thư ký' });

        parsedHoidongs[`hd_${idx}`] = {
            rawIndex: String(idx + 1),
            soHD: slot.soHD,
            hasExplicitSoHD: true,
            ngayBV: slot.ngayBV,
            gioBV: slot.gioBV,
            phongBV: slot.phongBV,
            diaDiemBV: slot.diaDiemBV,
            thanhVien,
            sinhViens: slot.sinhViens
        };
    });

    return {
        ok: true,
        hasMasterList,
        blockingErrors: [],
        consistencyWarnings,
        parsedHoidongs,
        studentList,
        teacherPoolList: Array.from(usedTeachersMap.values()),
        detectedKhoas: detectedKhoaList,
        detectedNganhs: detectedNganhList
    };
}

/**
 * 3. Render Kho GV & Kho SV lên giao diện (tương thích 100% DOM hiện tại)
 */
function renderPoolsForCombinedMode(teacherPoolList, studentList) {
    const poolGV = document.getElementById('pool-gv');
    const poolSV = document.getElementById('pool-sv');

    if (poolGV) {
        poolGV.innerHTML = teacherPoolList.map(gv => {
            const isTempId = /^AUTO_GV_/i.test(String(gv.maGV || ''));
            const subLine = isTempId
                ? `${escapeHtml(gv.donVi)} <span class="text-warning-emphasis">(Chưa có Mã GV)</span>`
                : `${escapeHtml(gv.donVi)} - Mã: ${escapeHtml(gv.maGV)}`;

            return `
                <div class="card-item position-relative mb-2 p-2 pe-4 border rounded bg-white shadow-sm"
                     data-id="${escapeHtml(gv.maGV)}"
                     data-name="${escapeHtml(gv.hoTenGV)}"
                     data-donvi="${escapeHtml(gv.donVi)}">
                    <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card" title="Xóa giảng viên này">✕</button>
                    <div class="fw-bold text-dark">${escapeHtml(gv.hoTenGV)}</div>
                    <div class="text-muted small">${subLine}</div>
                </div>
            `;
        }).join('');
    }

    if (poolSV) {
        poolSV.innerHTML = studentList.map(sv => {
            const isTempId = /^AUTO_GV_/i.test(String(sv.maGVHD || ''));
            const maGVDisplay = (sv.maGVHD && !isTempId) ? ` [${sv.maGVHD}]` : '';
            return `
                <div class="card-item position-relative mb-2 p-2 pe-4 border rounded bg-white shadow-sm"
                     data-id="${escapeHtml(sv.maSV)}"
                     data-name="${escapeHtml(sv.hoTenSV)}"
                     data-lop="${escapeHtml(sv.lop)}"
                     data-detai="${escapeHtml(sv.deTai)}"
                     data-magvhd="${escapeHtml(sv.maGVHD)}"
                     data-gvhd="${escapeHtml(sv.gvhd)}"
                     data-nganh="${escapeHtml(sv.nganh)}">
                    <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card" title="Xóa sinh viên này">✕</button>
                    <div class="fw-bold text-primary">${escapeHtml(sv.hoTenSV)} (${escapeHtml(sv.maSV)})</div>
                    <div class="text-muted small">Lớp: ${escapeHtml(sv.lop)} | GVHD: <span class="fw-bold text-dark">${escapeHtml(sv.gvhd)}${escapeHtml(maGVDisplay)}</span></div>
                    <div class="text-dark small text-truncate" title="${escapeHtml(sv.deTai)}">Đề tài: ${escapeHtml(sv.deTai)}</div>
                </div>
            `;
        }).join('');
    }
}

/**
 * 4. Khởi tạo bộ lắng nghe sự kiện cho Bảng tùy chọn & 2 ô upload của luồng Tổng hợp
 */
export function initCombinedImportWorkflow() {
    const radioNo = document.getElementById('opt-has-combined-no');
    const radioYes = document.getElementById('opt-has-combined-yes');
    const legacySection = document.getElementById('legacy-upload-section');
    const combinedSection = document.getElementById('combined-upload-section');

    const inputGVFile = document.getElementById('file-combined-gv');
    const inputAllFile = document.getElementById('file-combined-all');
    const btnReadCombined = document.getElementById('btn-read-combined-files');
    const statusBox = document.getElementById('combined-import-status');

    if (!radioNo || !radioYes || !legacySection || !combinedSection) return;

    let cachedTeacherMaster = null;
    let cachedCombinedRows = null;

    const switchMode = (hasCombined) => {
        if (hasCombined === true) {
            legacySection.classList.add('d-none');
            combinedSection.classList.remove('d-none');
        } else if (hasCombined === false) {
            combinedSection.classList.add('d-none');
            legacySection.classList.remove('d-none');
        } else {
            legacySection.classList.add('d-none');
            combinedSection.classList.add('d-none');
        }
    };

    if (radioYes.checked) {
        switchMode(true);
    } else if (radioNo.checked) {
        switchMode(false);
    } else {
        switchMode(null);
    }

    radioNo.addEventListener('change', () => {
        if (radioNo.checked) switchMode(false);
    });
    radioYes.addEventListener('change', () => {
        if (radioYes.checked) switchMode(true);
    });

    const showStatusAlert = (type, title, lines = []) => {
        if (!statusBox) return;
        const hasLines = Array.isArray(lines) && lines.length > 0;
        const toggleBtnHtml = hasLines
            ? `<button type="button" id="btn-toggle-combined-status"
                       class="btn btn-sm btn-outline-secondary bg-white py-0 px-2 fw-semibold"
                       style="font-size: 12px; white-space: nowrap;">
                   Thu gọn ▲
               </button>`
            : '';

        statusBox.className = `alert alert-${type} py-2 px-3 small mt-2 mb-0`;
        statusBox.innerHTML = `
            <div class="d-flex justify-content-between align-items-center gap-2">
                <div class="fw-bold">${escapeHtml(title)}</div>
                ${toggleBtnHtml}
            </div>
            ${hasLines ? `<div id="combined-status-details" class="mt-2 pt-2 border-top" style="max-height: 180px; overflow-y: auto; white-space: pre-line;">${escapeHtml(lines.join('\n'))}</div>` : ''}
        `;
        statusBox.classList.remove('d-none');

        const toggleBtn = document.getElementById('btn-toggle-combined-status');
        const detailsEl = document.getElementById('combined-status-details');
        if (toggleBtn && detailsEl) {
            toggleBtn.addEventListener('click', () => {
                const isHidden = detailsEl.classList.contains('d-none');
                if (isHidden) {
                    detailsEl.classList.remove('d-none');
                    toggleBtn.textContent = 'Thu gọn ▲';
                } else {
                    detailsEl.classList.add('d-none');
                    toggleBtn.textContent = `Mở rộng (${lines.length}) ▼`;
                }
            });
        }
    };

    const clearRenderedWorkspace = () => {
        const ws = document.getElementById('workspace-hd');
        if (ws) ws.innerHTML = '';
        const poolSV = document.getElementById('pool-sv');
        if (poolSV) poolSV.innerHTML = '';
    };

    const runCombinedPipeline = async () => {
        if (!cachedTeacherMaster && !cachedCombinedRows) return;

        // Nếu người dùng CÓ chọn File Danh sách GV và file đó có lỗi nội bộ -> chặn trước
        if (cachedTeacherMaster && cachedTeacherMaster.internalErrors.length > 0) {
            clearRenderedWorkspace();
            showStatusAlert(
                'danger',
                `⛔ TẠM CHẶN DỮ LIỆU: File Danh sách GV có ${cachedTeacherMaster.internalErrors.length} lỗi cần chỉnh sửa trước khi tiếp tục:`,
                cachedTeacherMaster.internalErrors
            );
            return;
        }

        // Nếu mới chỉ chọn File Danh sách GV mà chưa chọn File Phân công tổng hợp
        if (cachedTeacherMaster && !cachedCombinedRows) {
            showStatusAlert(
                'info',
                `✅ Đã nạp hợp lệ ${cachedTeacherMaster.rawList.length} Giảng viên. Vui lòng chọn tiếp File Phân công tổng hợp!`
            );
            return;
        }

        // Chạy đối soát File Phân công tổng hợp (có hoặc không có teacherLookupMap)
        const lookupMap = cachedTeacherMaster ? cachedTeacherMaster.lookupMap : null;
        const result = await parseAndValidateCombinedAssignment(cachedCombinedRows, lookupMap);

        if (!result.ok) {
            clearRenderedWorkspace();
            const failTitle = result.hasMasterList
                ? `⛔ TẠM CHẶN RENDER HỘI ĐỒNG: Phát hiện ${result.blockingErrors.length} giảng viên chưa khớp với File Danh sách GV (thiếu ký hiệu (A)/(B) hoặc sai tên):`
                : `⛔ TẠM CHẶN RENDER HỘI ĐỒNG: Phát hiện ${result.blockingErrors.length} lỗi dữ liệu / thiếu ký hiệu phân biệt tên Giảng viên trong File Phân công tổng hợp:`;

            showStatusAlert('danger', failTitle, result.blockingErrors);
            return;
        }

        renderPoolsForCombinedMode(result.teacherPoolList, result.studentList);

        if (typeof window.syncKhoaNganhFromLop === 'function') {
            if (result.detectedNganhs && result.detectedNganhs.length > 0) {
                window.syncKhoaNganhFromLop({
                    khoa: result.detectedKhoas?.[0] || '',
                    nganhs: result.detectedNganhs
                });
            } else if (result.studentList[0]?.lop) {
                window.syncKhoaNganhFromLop(result.studentList[0].lop);
            }
        }

        const checkLogs = renderHoidongFromExcel(result.parsedHoidongs);

        // Ẩn mã tạm AUTO_GV_xxx trong thẻ GV vừa vẽ trong các khung Hội đồng
        document.querySelectorAll('.box-hoidong .dropzone-gv .card-item').forEach(cardEl => {
            const id = cardEl.getAttribute('data-id') || '';
            if (/^AUTO_GV_/i.test(id)) {
                cardEl.querySelectorAll('.text-muted, small, div').forEach(sub => {
                    if (sub.textContent.includes(id)) {
                        sub.textContent = sub.textContent.replace(new RegExp(`\\s*-?\\s*Mã:\\s*${id}`, 'gi'), '').trim();
                    }
                });
            }
        });

        const modeNotice = !result.hasMasterList
            ? ['• ℹ️ Đang chạy ở chế độ KHÔNG có File Danh sách GV: Hệ thống tự nhận diện GV theo tên & ký hiệu phân biệt (chưa có Mã GV chuẩn và Đơn vị công tác khi xuất Bảng thanh toán).']
            : [];

        const allWarnings = [
            ...modeNotice,
            ...result.consistencyWarnings,
            ...(checkLogs.duplicateSVLogs || []),
            ...(checkLogs.conflictLogs || []),
            ...(checkLogs.mergeLogs || [])
        ];

        if (allWarnings.length > 0) {
            showStatusAlert(
                'warning',
                `✅ Đã nạp thành công ${result.studentList.length} SV (${result.detectedNganhs.length || 1} ngành) và dựng Hội đồng. Lưu ý rà soát ${allWarnings.length} ghi chú/cảnh báo bên dưới:`,
                allWarnings
            );
        } else {
            showStatusAlert(
                'success',
                `✅ Đối soát hoàn tất: 100% Giảng viên khớp Mã GV (${result.studentList.length} Sinh viên — ${result.detectedNganhs.length || 1} ngành đã được xếp vào các Hội đồng hợp lệ)!`
            );
        }
    };

    inputGVFile?.addEventListener('change', async (e) => {
        const file = e.target.files?.[0];
        if (!file) {
            cachedTeacherMaster = null;
            return;
        }
        try {
            cachedTeacherMaster = await parseStrictTeacherMasterFile(file, inputGVFile);
            await runCombinedPipeline();
        } catch (err) {
            cachedTeacherMaster = null;
            showStatusAlert('danger', `❌ ${err.message || err}`);
        }
    });

    inputAllFile?.addEventListener('change', async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        try {
            cachedCombinedRows = await readFirstSheetRowsFromFile(file, inputAllFile);
            await runCombinedPipeline();
        } catch (err) {
            cachedCombinedRows = null;
            showStatusAlert('danger', `❌ ${err.message || err}`);
        }
    });

    btnReadCombined?.addEventListener('click', async () => {
        const gvFile = inputGVFile?.files?.[0];
        const allFile = inputAllFile?.files?.[0];

        if (!allFile && !cachedCombinedRows) {
            showStatusAlert('warning', '⚠️ Vui lòng chọn File (2) Phân công tổng hợp trước khi bấm Nạp dữ liệu!');
            return;
        }

        try {
            if (gvFile) {
                cachedTeacherMaster = await parseStrictTeacherMasterFile(gvFile, inputGVFile);
            } else if (!inputGVFile?.value) {
                cachedTeacherMaster = null;
            }
            if (allFile) {
                cachedCombinedRows = await readFirstSheetRowsFromFile(allFile, inputAllFile);
            }
            await runCombinedPipeline();
        } catch (err) {
            showStatusAlert('danger', `❌ ${err.message || err}`);
        }
    });
}
