// js/combined-importer.js
// Module độc lập xử lý luồng nhập 2 file: Danh sách GV + Danh sách Phân công tổng hợp (SV, GVHD, HĐ)

import { renderHoidongFromExcel } from './hoidong-manager.js';

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

    // Chuẩn hóa khoảng trắng quanh dấu ngoặc đơn: "Nguyễn Đức Hoàng  ( B )" -> "nguyễn đức hoàng (b)"
    s = s
        .replace(/\s*\(\s*/g, ' (')
        .replace(/\s*\)\s*/g, ')')
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
    if (u.includes('CTP')) return 'Công nghệ Thực phẩm';
    if (u.includes('SHU')) return 'Công nghệ Sinh học';
    if (u.includes('KMT')) return 'Quản lý Tài nguyên & Môi trường';
    return '';
}

function normalizeExcelDateValue(raw) {
    if (raw === null || raw === undefined || raw === '') return '';
    if (typeof raw === 'number' && raw > 20000) {
        // Excel serial date -> dd/mm/yyyy
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
    if (typeof raw === 'number' && raw >= 0 && raw < 1) {
        const totalMins = Math.round(raw * 24 * 60);
        const hh = String(Math.floor(totalMins / 60)).padStart(2, '0');
        const mm = String(totalMins % 60).padStart(2, '0');
        return `${hh}:${mm}`;
    }
    return String(raw).trim();
}

async function readFirstSheetRowsFromFile(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const firstSheetName = wb.SheetNames[0];
    const ws = wb.Sheets[firstSheetName];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
}

/**
 * 1. Đọc & kiểm tra nội bộ File Danh sách Giảng viên (Mã GV | Họ và Tên | Đơn vị Công tác)
 */
export async function parseStrictTeacherMasterFile(file) {
    const rows = await readFirstSheetRowsFromFile(file);
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
 * 2. Đọc & đối soát chặt chẽ File Phân công tổng hợp với lookupMap từ Danh sách GV
 */
export async function parseAndValidateCombinedAssignment(file, teacherLookupMap) {
    const rows = await readFirstSheetRowsFromFile(file);
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

            // Kiểm tra xem cột Họ và Tên có bị tách làm 2 cột (Họ đệm + Tên) giống file Khoa Cơ khí không
            const isSplitNameCol = idxHoTen !== -1 && idxLop !== -1 && (idxLop - idxHoTen === 2) && !rUpper[idxHoTen + 1];

            // Các cột ngay sau cột GVHD: Phản biện, Chủ tịch, Thư ký, Hội đồng, Ngày bảo vệ, Giờ bảo vệ, Phòng, Cơ sở
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
            teacherPoolList: []
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

        // Giữ nguyên Học hàm/Học vị + Tên + (A)/(B) như file tổng hợp upload
        return {
            maGV: found.maGV,
            hoTenGV: displayName,
            donVi: found.donVi
        };
    };

    const studentList = [];
    const hdGroupMap = new Map();
    const usedTeachersMap = new Map(); // Lưu các GV tham gia kèm Học hàm/Học vị chuẩn để nạp Kho GV

    const registerTeacherToPool = (tObj) => {
        if (!tObj || !tObj.maGV) return;
        const idKey = tObj.maGV.toUpperCase();
        const existing = usedTeachersMap.get(idKey);
        // Ưu tiên bản ghi có học hàm/học vị đầy đủ hơn (chuỗi dài hơn)
        if (!existing || tObj.hoTenGV.length > existing.hoTenGV.length) {
            usedTeachersMap.set(idKey, { ...tObj });
        }
    };

    // Khởi tạo trước Kho GV từ toàn bộ file Danh sách GV (để người dùng có thể gán thêm GV dự phòng nếu cần)
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

        // Đối soát bắt buộc 4 cột GV với Danh sách GV
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

        const nganh = detectNganhFromLop(lop);

        studentList.push({
            maSV: mssv,
            hoTenSV,
            lop,
            deTai,
            maGVHD: gvhdObj?.maGV || '',
            gvhd: gvhdObj?.hoTenGV || cleanDisplayTeacherName(rawGVHD),
            nganh
        });

        // Gom nhóm vào Hội đồng nếu cột HỘI ĐỒNG hoặc thành viên HĐ có dữ liệu
        if (rawHD || ctObj || pbObj || tkObj) {
            const hdNumClean = rawHD.replace(/^(?:hội\s*đồng|hoi\s*dong|hđ|hd)\s*(?:số|so|#|:|-|\.)?\s*/i, '').trim() || 'Chưa_Đánh_Số';
            // Tách theo (Số HĐ + Ngày bảo vệ) nếu cùng 1 HĐ bảo vệ nhiều ngày khác nhau
            const groupKey = `${hdNumClean}__${ngayBV || 'NO_DATE'}`;

            if (!hdGroupMap.has(groupKey)) {
                hdGroupMap.set(groupKey, {
                    soHD: hdNumClean,
                    ngayBV,
                    gioBV,
                    phongBV,
                    diaDiemBV,
                    ct: ctObj,
                    pb: pbObj,
                    tk: tkObj,
                    sinhViens: []
                });
            }

            const hdRec = hdGroupMap.get(groupKey);
            if (!hdRec.gioBV && gioBV) hdRec.gioBV = gioBV;
            if (!hdRec.phongBV && phongBV) hdRec.phongBV = phongBV;
            if (!hdRec.diaDiemBV && diaDiemBV) hdRec.diaDiemBV = diaDiemBV;

            // Kiểm tra tính nhất quán thành viên giữa các dòng cùng 1 Hội đồng
            const checkRoleConsistency = (roleName, existingMember, newMember) => {
                if (!existingMember && newMember) return newMember;
                if (existingMember && newMember && existingMember.maGV.toUpperCase() !== newMember.maGV.toUpperCase()) {
                    consistencyWarnings.push(
                        `• Dòng ${rowNum} (MSSV ${mssv}, HĐ ${hdNumClean}): Cột ${roleName} ghi "${newMember.hoTenGV}" khác với dòng đầu của HĐ (${existingMember.hoTenGV}). Hệ thống giữ theo dòng đầu tiên.`
                    );
                }
                return existingMember;
            };

            hdRec.ct = checkRoleConsistency('Chủ tịch', hdRec.ct, ctObj);
            hdRec.pb = checkRoleConsistency('Phản biện', hdRec.pb, pbObj);
            hdRec.tk = checkRoleConsistency('Thư ký', hdRec.tk, tkObj);

            hdRec.sinhViens.push({
                maSV: mssv,
                hoTenSV,
                lop,
                deTai,
                maGVHD: gvhdObj?.maGV || '',
                gvhd: gvhdObj?.hoTenGV || cleanDisplayTeacherName(rawGVHD),
                nganh
            });
        }
    }

    // Nếu có bất kỳ lỗi không khớp Mã GV / thiếu (A)/(B) nào -> CHẶN CỨNG, không dựng Hội đồng
    if (blockingErrors.length > 0) {
        return {
            ok: false,
            blockingErrors,
            consistencyWarnings,
            parsedHoidongs: {},
            studentList: [],
            teacherPoolList: []
        };
    }

    // Sắp xếp danh sách Hội đồng theo số tăng dần và chuyển sang định dạng chuẩn của renderHoidongFromExcel
    const sortedGroups = Array.from(hdGroupMap.values()).sort((a, b) => {
        const na = parseInt(a.soHD, 10);
        const nb = parseInt(b.soHD, 10);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return String(a.soHD).localeCompare(String(b.soHD), 'vi');
    });

    const parsedHoidongs = {
        __preserveExplicitNumbers: true,
        __mergeLogs: [],
        __missingCodeIssues: []
    };

    sortedGroups.forEach((g, idx) => {
        const thanhVien = [];
        if (g.ct) thanhVien.push({ ...g.ct, chucDanh: 'Chủ tịch' });
        if (g.pb) thanhVien.push({ ...g.pb, chucDanh: 'Phản biện' });
        if (g.tk) thanhVien.push({ ...g.tk, chucDanh: 'Thư ký' });

        parsedHoidongs[`hd_${idx}`] = {
            rawIndex: String(idx + 1),
            soHD: g.soHD,
            hasExplicitSoHD: true,
            ngayBV: g.ngayBV,
            gioBV: g.gioBV,
            phongBV: g.phongBV,
            diaDiemBV: g.diaDiemBV,
            thanhVien,
            sinhViens: g.sinhViens
        };
    });

    return {
        ok: true,
        blockingErrors: [],
        consistencyWarnings,
        parsedHoidongs,
        studentList,
        teacherPoolList: Array.from(usedTeachersMap.values())
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
    let cachedCombinedFile = null;

    const switchMode = (hasCombined) => {
        if (hasCombined) {
            legacySection.classList.add('d-none');
            combinedSection.classList.remove('d-none');
        } else {
            combinedSection.classList.add('d-none');
            legacySection.classList.remove('d-none');
        }
    };

    radioNo.addEventListener('change', () => {
        if (radioNo.checked) switchMode(false);
    });
    radioYes.addEventListener('change', () => {
        if (radioYes.checked) switchMode(true);
    });

    const showStatusAlert = (type, title, lines = []) => {
        if (!statusBox) return;
        statusBox.className = `alert alert-${type} py-2 px-3 small mt-2 mb-0`;
        statusBox.innerHTML = `
            <div class="fw-bold mb-1">${escapeHtml(title)}</div>
            ${lines.length > 0 ? `<div style="max-height: 180px; overflow-y: auto; white-space: pre-line;">${escapeHtml(lines.join('\n'))}</div>` : ''}
        `;
        statusBox.classList.remove('d-none');
    };

    const clearRenderedWorkspace = () => {
        const ws = document.getElementById('workspace-hd');
        if (ws) ws.innerHTML = '';
        const poolSV = document.getElementById('pool-sv');
        if (poolSV) poolSV.innerHTML = '';
    };

    const runCombinedPipeline = async () => {
        if (!cachedTeacherMaster && !cachedCombinedFile) return;

        if (!cachedTeacherMaster) {
            showStatusAlert('warning', '⏳ Đã nhận File Phân công tổng hợp. Vui lòng tải lên File Danh sách GV để đối soát Mã GV!');
            return;
        }

        if (cachedTeacherMaster.internalErrors.length > 0) {
            // Chặn render Hội đồng khi File Danh sách GV bị trùng tên chưa gắn (A)/(B) hoặc trùng Mã GV
            clearRenderedWorkspace();
            showStatusAlert(
                'danger',
                `⛔ TẠM CHẶN DỮ LIỆU: File Danh sách GV có ${cachedTeacherMaster.internalErrors.length} lỗi cần chỉnh sửa trước khi tiếp tục:`,
                cachedTeacherMaster.internalErrors
            );
            return;
        }

        if (!cachedCombinedFile) {
            showStatusAlert(
                'info',
                `✅ Đã nạp hợp lệ ${cachedTeacherMaster.rawList.length} Giảng viên. Vui lòng chọn tiếp File Phân công tổng hợp!`
            );
            return;
        }

        const result = await parseAndValidateCombinedAssignment(cachedCombinedFile, cachedTeacherMaster.lookupMap);

        if (!result.ok) {
            // Chặn cứng: xóa sạch khu vực Hội đồng để buộc người dùng cập nhật khớp dữ liệu
            clearRenderedWorkspace();
            showStatusAlert(
                'danger',
                `⛔ TẠM CHẶN RENDER HỘI ĐỒNG: Phát hiện ${result.blockingErrors.length} giảng viên chưa khớp với File Danh sách GV (thiếu ký hiệu (A)/(B) hoặc sai tên). Vui lòng cập nhật lại file Excel và tải lại:`,
                result.blockingErrors
            );
            return;
        }

        // 100% khớp Mã GV -> Nạp Kho GV/SV và chạy toàn bộ bộ kiểm tra trùng lặp của renderHoidongFromExcel
        renderPoolsForCombinedMode(result.teacherPoolList, result.studentList);
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
                `✅ Đã nạp thành công ${result.studentList.length} SV và dựng Hội đồng. Lưu ý rà soát ${allWarnings.length} cảnh báo nghiệp vụ bên dưới:`,
                allWarnings
            );
        } else {
            showStatusAlert(
                'success',
                `✅ Đối soát hoàn tất: 100% Giảng viên khớp Mã GV (${result.studentList.length} Sinh viên đã được xếp vào các Hội đồng hợp lệ)!`
            );
        }
    };

    inputGVFile?.addEventListener('change', async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        cachedTeacherMaster = await parseStrictTeacherMasterFile(file);
        await runCombinedPipeline();
    });

    inputAllFile?.addEventListener('change', async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        cachedCombinedFile = file;
        await runCombinedPipeline();
    });

    btnReadCombined?.addEventListener('click', async () => {
        const gvFile = inputGVFile?.files?.[0];
        const allFile = inputAllFile?.files?.[0];

        if (!gvFile || !allFile) {
            showStatusAlert('warning', '⚠️ Vui lòng chọn đầy đủ cả 2 file: (1) Danh sách GV và (2) File Phân công tổng hợp!');
            return;
        }

        // Đọc lại trực tiếp từ disk để lấy phiên bản file mới nhất nếu người dùng vừa sửa & lưu trên Excel
        cachedTeacherMaster = await parseStrictTeacherMasterFile(gvFile);
        cachedCombinedFile = allFile;
        await runCombinedPipeline();
    });
}