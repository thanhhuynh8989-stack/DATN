// js/combined-importer.js
// Module độc lập xử lý luồng nhập 2 file: Danh sách GV + Danh sách Phân công tổng hợp (SV, GVHD, HĐ)

import { renderHoidongFromExcel } from './hoidong-manager.js';
import { findNganhInfoByLop } from './excel-reader.js';

/**
 * Chuẩn hóa chuỗi khóa tra cứu Giảng viên:
 * - Đưa về Unicode NFC, chữ thường, rút gọn khoảng trắng
 * - Loại bỏ học hàm / học vị ở đầu (PGS, GS, TSKH, TS, ThS, Th.S, GVC, GV, KS, CN...)
 * - GIỮ NGUYÊN ký hiệu phân biệt trong ngoặc như (a), (b)
 */
export function buildStrictTeacherLookupKey(rawName) {
    let s = String(rawName || '').normalize('NFC').trim();
    if (!s) return '';

    // Loại bỏ học hàm / học vị ở đầu chuỗi
    s = s.replace(
        /^(?:PGS\.?\s*TSKH\.?|PGS\.?\s*TS\.?|GS\.?\s*TSKH\.?|GS\.?\s*TS\.?|TSKH\.?|TS\.?|Th\.?\s*S\.?|GVC\.?|GV\.?|KS\.?|CN\.?|BS\.?|Dr\.?|Prof\.?)+\s*/i,
        ''
    );

    // Loại bỏ dấu ngoặc đơn (nếu có) và chuẩn hóa khoảng trắng:
    // "Nguyễn Đức Hoàng (A)" hoặc "Nguyễn Đức Hoàng A" -> "nguyễn đức hoàng a"
    s = s
        .replace(/[()]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();

    return s;
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
    if (u.includes('CKM')) return 'Công nghệ Chế tạo máy';
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
 * 2. Đọc & đối soát chặt chẽ File Phân công tổng hợp:
 * - Gộp thành 1 Hội đồng khi trùng đúng bộ 3 thành viên theo chức danh (CT + PB + TK), bất kể cột HỘI ĐỒNG ghi số gì.
 * - Khi gộp nhiều số HĐ khác nhau về cùng 1 HĐ, ưu tiên giữ số HĐ nhỏ nhất / xuất hiện đầu tiên.
 * - Khi 2 nhóm KHÁC thành viên bị ghi trùng số HĐ trong Excel: nhóm đầu giữ nguyên, nhóm sau tự động cấp số kế tiếp chưa dùng & ghi cảnh báo.
 * - Tách thành từng khung chấm trên giao diện nếu 1 HĐ có nhiều suất (Buổi + Ngày + Phòng + Cơ sở) khác nhau, nhưng giữ nguyên cùng Số HĐ.
 */
export async function parseAndValidateCombinedAssignment(fileOrRows, teacherLookupMap, inputElementToReset = null) {
    const rows = Array.isArray(fileOrRows)
        ? fileOrRows
        : await readFirstSheetRowsFromFile(fileOrRows, inputElementToReset);

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
            blockingErrors: ['• Không tìm thấy dòng tiêu đề hợp lệ (cần có cột MSSV / Mã SV và cột GV HƯỚNG DẪN) trong file Phân công tổng hợp!'],
            consistencyWarnings: [],
            parsedHoidongs: {},
            studentList: [],
            teacherPoolList: [],
            detectedKhoas: [],
            detectedNganhs: []
        };
    }

    const resolveTeacherStrict = (rawCellVal, rowNumber, mssv, roleLabel) => {
        const displayName = cleanDisplayTeacherName(rawCellVal);
        if (!displayName) return null;

        const key = buildStrictTeacherLookupKey(displayName);
        const found = teacherLookupMap.get(key);

        if (!found) {
            const hasSuffix = /\([a-z0-9]+\)/i.test(displayName);
            const hint = hasSuffix
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
    };

    const studentList = [];
    const memberGroupMap = new Map(); // Key = "CT_ID|PB_ID|TK_ID"
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

    teacherLookupMap.forEach(rec => {
        usedTeachersMap.set(rec.maGV.toUpperCase(), {
            maGV: rec.maGV,
            hoTenGV: rec.hoTenGoc,
            donVi: rec.donVi
        });
    });

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

        const gvhdObj = resolveTeacherStrict(rawGVHD, rowNum, mssv, 'GV Hướng dẫn');
        if (!cleanDisplayTeacherName(rawGVHD)) {
            blockingErrors.push(`• Dòng ${rowNum} (MSSV: ${mssv}): Bị trống tên GV Hướng dẫn.`);
        }

        const pbObj = resolveTeacherStrict(rawPB, rowNum, mssv, 'Phản biện');
        const ctObj = resolveTeacherStrict(rawCT, rowNum, mssv, 'Chủ tịch');
        const tkObj = resolveTeacherStrict(rawTK, rowNum, mssv, 'Thư ký');

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

            // Nếu hàng có đủ/có ít nhất 1 thành viên HĐ -> gom theo đúng tổ hợp (CT | PB | TK)
            // Nếu hàng hoàn toàn trống cả 3 cột CT, PB, TK -> gom tạm theo số HĐ
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

    if (blockingErrors.length > 0) {
        return {
            ok: false,
            blockingErrors,
            consistencyWarnings,
            parsedHoidongs: {},
            studentList: [],
            teacherPoolList: [],
            detectedKhoas: [],
            detectedNganhs: []
        };
    }

    // =========================================================================
    // BƯỚC XỬ LÝ SỐ HỘI ĐỒNG CHO TỪNG NHÓM THÀNH VIÊN:
    // 1. Trong cùng 1 nhóm thành viên: nếu có nhiều số HĐ khác nhau -> giữ số nhỏ nhất / xuất hiện đầu tiên.
    // 2. Giữa các nhóm KHÁC thành viên: nếu trùng số HĐ đã chiếm trước đó -> giữ nguyên nhóm đầu,
    //    cấp số HĐ kế tiếp chưa dùng cho nhóm sau & ghi thông báo rõ ràng.
    // =========================================================================
    const memberGroups = Array.from(memberGroupMap.values());

    // Tìm số HĐ lớn nhất hiện có trong file để cấp số kế tiếp khi bị trùng (vd: max=20 -> cấp 21)
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

        // Chọn số HĐ đại diện cho nhóm này: ưu tiên số nhỏ nhất (nếu là số) hoặc số xuất hiện đầu tiên
        let preferredSoHD = '';
        if (g.rawSoHDs.length === 1) {
            preferredSoHD = g.rawSoHDs[0];
        } else if (g.rawSoHDs.length > 1) {
            const sortedCandidates = [...g.rawSoHDs].sort((a, b) => {
                const na = parseInt(a, 10);
                const nb = parseInt(b, 10);
                if (!isNaN(na) && !isNaN(nb)) return na - nb;
                return 0; // Giữ thứ tự xuất hiện đầu tiên nếu không phải số
            });
            preferredSoHD = sortedCandidates[0];
            consistencyWarnings.push(
                `• Gộp Hội đồng (${memberDesc}): Trong Excel các dòng của bộ thành viên này đang ghi nhiều số HĐ khác nhau (${g.rawSoHDs.map(x => `HĐ ${x}`).join(', ')}) → Hệ thống đã gộp chung thành 1 Hội đồng và giữ số HĐ ${preferredSoHD}.`
            );
        }

        // Kiểm tra xem preferredSoHD đã bị nhóm thành viên khác chiếm chưa
        const normKey = preferredSoHD.toUpperCase();
        if (preferredSoHD && !allocatedSoHDs.has(normKey)) {
            g.finalSoHD = preferredSoHD;
            allocatedSoHDs.add(normKey);
        } else {
            // Cấp số HĐ kế tiếp chưa dùng
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

    // Sắp xếp các Hội đồng theo finalSoHD tăng dần
    memberGroups.sort((a, b) => {
        const na = parseInt(a.finalSoHD, 10);
        const nb = parseInt(b.finalSoHD, 10);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return String(a.finalSoHD).localeCompare(String(b.finalSoHD), 'vi');
    });

    // =========================================================================
    // TÁCH BOX HIỂN THỊ THEO SUẤT CHẤM (BUỔI + NGÀY + PHÒNG + CƠ SỞ)
    // Giữ nguyên cùng finalSoHD cho các suất chấm thuộc cùng 1 Hội đồng
    // =========================================================================
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
            // Lấy giờ bắt đầu sớm nhất trong cùng buổi nếu có nhiều mốc giờ
            if (sv.gioBV) {
                if (!slotRec.gioBV || parseMinutesForSort(sv.gioBV) < parseMinutesForSort(slotRec.gioBV)) {
                    slotRec.gioBV = sv.gioBV;
                }
            }
            slotRec.sinhViens.push(sv);
        });

        // Sắp xếp các suất chấm của cùng 1 HĐ theo Ngày -> Buổi -> Phòng
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

    const parsedHoidongs = {
        __preserveExplicitNumbers: true,
        __skipAutoMergeInRenderer: true, // Đã chuẩn hóa nhóm theo thành viên & tách theo suất chấm ở trên
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
        poolGV.innerHTML = teacherPoolList.map(gv => `
            <div class="card-item position-relative mb-2 p-2 pe-4 border rounded bg-white shadow-sm"
                 data-id="${escapeHtml(gv.maGV)}"
                 data-name="${escapeHtml(gv.hoTenGV)}"
                 data-donvi="${escapeHtml(gv.donVi)}">
                <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card" title="Xóa giảng viên này">✕</button>
                <div class="fw-bold text-dark">${escapeHtml(gv.hoTenGV)}</div>
                <div class="text-muted small">${escapeHtml(gv.donVi)} - Mã: ${escapeHtml(gv.maGV)}</div>
            </div>
        `).join('');
    }

    if (poolSV) {
        poolSV.innerHTML = studentList.map(sv => {
            const maGVDisplay = sv.maGVHD ? ` [${sv.maGVHD}]` : '';
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
            // Mặc định khi chưa chọn radio nào: ẩn cả 2 khu vực upload
            legacySection.classList.add('d-none');
            combinedSection.classList.add('d-none');
        }
    };

    // Khởi tạo trạng thái hiển thị ban đầu theo radio đang được chọn (nếu có)
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

        if (!cachedTeacherMaster) {
            showStatusAlert('warning', '⏳ Đã nhận File Phân công tổng hợp. Vui lòng tải lên File Danh sách GV để đối soát Mã GV!');
            return;
        }

        if (cachedTeacherMaster.internalErrors.length > 0) {
            clearRenderedWorkspace();
            showStatusAlert(
                'danger',
                `⛔ TẠM CHẶN DỮ LIỆU: File Danh sách GV có ${cachedTeacherMaster.internalErrors.length} lỗi cần chỉnh sửa trước khi tiếp tục:`,
                cachedTeacherMaster.internalErrors
            );
            return;
        }

        if (!cachedCombinedRows) {
            showStatusAlert(
                'info',
                `✅ Đã nạp hợp lệ ${cachedTeacherMaster.rawList.length} Giảng viên. Vui lòng chọn tiếp File Phân công tổng hợp!`
            );
            return;
        }

        const result = await parseAndValidateCombinedAssignment(cachedCombinedRows, cachedTeacherMaster.lookupMap);

        if (!result.ok) {
            clearRenderedWorkspace();
            showStatusAlert(
                'danger',
                `⛔ TẠM CHẶN RENDER HỘI ĐỒNG: Phát hiện ${result.blockingErrors.length} giảng viên chưa khớp với File Danh sách GV (thiếu ký hiệu (A)/(B) hoặc sai tên). Vui lòng cập nhật lại file Excel và tải lại:`,
                result.blockingErrors
            );
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

        const allWarnings = [
            ...result.consistencyWarnings,
            ...(checkLogs.duplicateSVLogs || []),
            ...(checkLogs.conflictLogs || []),
            ...(checkLogs.mergeLogs || [])
        ];

        if (allWarnings.length > 0) {
            showStatusAlert(
                'warning',
                `✅ Đã nạp thành công ${result.studentList.length} SV (${result.detectedNganhs.length || 1} ngành) và dựng Hội đồng. Lưu ý rà soát ${allWarnings.length} cảnh báo nghiệp vụ bên dưới:`,
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
        if (!file) return;
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

        if (!gvFile && !cachedTeacherMaster) {
            showStatusAlert('warning', '⚠️ Vui lòng chọn File (1) Danh sách GV!');
            return;
        }
        if (!allFile && !cachedCombinedRows) {
            showStatusAlert('warning', '⚠️ Vui lòng chọn File (2) Phân công tổng hợp!');
            return;
        }

        try {
            if (gvFile) {
                cachedTeacherMaster = await parseStrictTeacherMasterFile(gvFile, inputGVFile);
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
