// js/excel-reader.js
import { getAdminBinaryFile } from './admin-config.js';
export let danhMucNganh = []; // Mảng lưu { maChuyenNganh, chuyenNganh, khoa }

// 1. Tự động nạp file DANH_SACH_NGANH.xlsx khi khởi động ứng dụng
export async function loadDanhSachNganhSilently(url = './sample_files/DANH_SACH_NGANH.xlsx') {
    try {
        let arrayBuffer = await getAdminBinaryFile('danh_sach_nganh');
        if (!arrayBuffer) {
            const response = await fetch(url);
            if (!response.ok) throw new Error(`Không tìm thấy file ${url}`);
            arrayBuffer = await response.arrayBuffer();
        }
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1 });

        let headerRowIdx = -1;
        let colMa = 0, colNganh = 1, colKhoa = 2;

        for (let i = 0; i < Math.min(rows.length, 10); i++) {
            const r = rows[i] || [];
            const upperCells = r.map(c => String(c || '').toUpperCase().trim());
            const idxNganh = upperCells.findIndex(c => c.includes('NGÀNH'));
            const idxKhoa = upperCells.findIndex(c => c.includes('KHOA'));

            if (idxNganh >= 0 && idxKhoa >= 0) {
                headerRowIdx = i;
                colNganh = idxNganh;
                colKhoa = idxKhoa;
                const idxMa = upperCells.findIndex(c => c.includes('MÃ') || c.includes('MA_'));
                if (idxMa >= 0) colMa = idxMa;
                break;
            }
        }

        const startRow = headerRowIdx >= 0 ? headerRowIdx + 1 : 1;
        const parsed = [];

        rows.forEach((row, idx) => {
            if (idx < startRow || !row || row.length < 2) return;
            const ma = String(row[colMa] || '').trim().toUpperCase();
            const tenNganh = String(row[colNganh] || '').trim();
            const tenKhoa = String(row[colKhoa] || '').trim();

            if (!ma || ma.includes('MA_CHUYEN_NGANH') || ma.includes('MÃ') || ma === 'STT') return;
            if (tenNganh && tenKhoa) {
                parsed.push({
                    maChuyenNganh: ma,
                    chuyenNganh: tenNganh,
                    khoa: tenKhoa
                });
            }
        });

        danhMucNganh = parsed;
        console.log(`✅ Đã nạp ${danhMucNganh.length} ngành/chuyên ngành từ danh mục.`);
        return danhMucNganh;
    } catch (err) {
        console.warn("Cảnh báo tải DANH_SACH_NGANH.xlsx:", err.message);
        return [];
    }
}

// 2. Tra cứu thông tin { chuyenNganh, khoa } từ chuỗi Lớp
export function findNganhInfoByLop(lopStr) {
    const cleanLop = String(lopStr || '').toUpperCase().trim();
    if (!cleanLop || danhMucNganh.length === 0) return null;

    const tokens = cleanLop.split(/[^A-Z0-9]+/).filter(Boolean);

    for (const item of danhMucNganh) {
        if (tokens.includes(item.maChuyenNganh)) {
            return item;
        }
    }

    const sorted = [...danhMucNganh].sort((a, b) => b.maChuyenNganh.length - a.maChuyenNganh.length);
    for (const item of sorted) {
        if (cleanLop.includes(item.maChuyenNganh)) {
            return item;
        }
    }

    return null;
}

// 3. Dịch mã Lớp sang tên Ngành chuẩn từ danhMucNganh
export function detectNganhFromLop(lopStr) {
    const found = findNganhInfoByLop(lopStr);
    return found ? found.chuyenNganh.toUpperCase() : '';
}

// 4. Đọc file Excel chung
export function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: false });
                const firstSheet = workbook.SheetNames[0];
                const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1, raw: false });
                resolve(jsonData);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

// 5. Bóc tách Danh sách Sinh viên & Đề tài (Hỗ trợ cả cột Họ & Tên tách rời hoặc Merge ngang)
export function parseSinhVienRows(rows) {
    let headerIdx = -1;
    let detectedNganhFromHeader = '';
    let colMap = {
        maSV: 1,
        hoDem: -1,
        ten: -1,
        hoTen: 2,
        lop: 3,
        deTai: 4,
        gvhd: 5,
        maGVHD: -1,
        nganh: -1
    };
    let hasCustomColMap = false;

    for (let i = 0; i < Math.min(rows.length, 15); i++) {
        const row = rows[i] || [];
        const rowStr = row.join(' ').trim();
        const upperStr = rowStr.toUpperCase();

        const nganhMatch = rowStr.match(/NGÀNH\s*[:\-]\s*([^\n\r]+)/i);
        if (nganhMatch && nganhMatch[1]) {
            detectedNganhFromHeader = nganhMatch[1].trim().toUpperCase();
        }

        if (upperStr.includes('MSSV') || upperStr.includes('MÃ SỐ SV') || upperStr.includes('MÃ SV')) {
            headerIdx = i;
            hasCustomColMap = true;

            row.forEach((cellVal, cIdx) => {
                const h = String(cellVal || '').toUpperCase().replace(/\s+/g, ' ').trim();
                if (!h) return;

                if (h.includes('MSSV') || h.includes('MÃ SỐ SV') || h.includes('MÃ SV')) {
                    colMap.maSV = cIdx;
                } else if (h === 'HỌ' || h.includes('HỌ ĐỆM') || h.includes('HỌ VÀ TÊN ĐỆM')) {
                    colMap.hoDem = cIdx;
                } else if (h === 'TÊN') {
                    colMap.ten = cIdx;
                } else if (h.includes('HỌ VÀ TÊN') || h.includes('HỌ TÊN')) {
                    colMap.hoTen = cIdx;
                } else if (h.includes('LỚP') || h.includes('KHÓA')) {
                    colMap.lop = cIdx;
                } else if (h.includes('ĐỀ TÀI')) {
                    colMap.deTai = cIdx;
                } else if (
                    h.includes('MÃ GV') ||
                    h.includes('MÃ SỐ GV') ||
                    h.includes('MÃ GIẢNG VIÊN') ||
                    h.includes('MÃ SỐ GIẢNG VIÊN')
                ) {
                    colMap.maGVHD = cIdx;
                } else if (
                    h.includes('GVHD') ||
                    h.includes('GIẢNG VIÊN HƯỚNG DẪN') ||
                    h.includes('GV HƯỚNG DẪN') ||
                    h.includes('HƯỚNG DẪN')
                ) {
                    colMap.gvhd = cIdx;
                } else if (h.includes('NGÀNH')) {
                    colMap.nganh = cIdx;
                }
            });

            // Xử lý ô tiêu đề "HỌ VÀ TÊN" bị merge ngang 2 cột (VD: merge C & D -> C là Họ đệm, D là Tên)
            if (colMap.hoTen >= 0 && colMap.hoDem < 0 && colMap.ten < 0) {
                const nextCol = colMap.hoTen + 1;
                const nextHeaderEmpty = !String(row[nextCol] || '').trim();
                const isNextColUnassigned = ![
                    colMap.maSV,
                    colMap.lop,
                    colMap.deTai,
                    colMap.gvhd,
                    colMap.maGVHD,
                    colMap.nganh
                ].includes(nextCol);

                if (nextHeaderEmpty && isNextColUnassigned) {
                    colMap.hoDem = colMap.hoTen;
                    colMap.ten = nextCol;
                }
            }
            break;
        }
    }

    const startRow = headerIdx >= 0 ? headerIdx + 1 : 1;
    const result = [];
    const missingCodeIssues = [];

    rows.forEach((row, idx) => {
        if (idx < startRow || !row || row.length === 0) return;

        const upperRow = row.map(c => String(c || '').trim()).join(' ').toUpperCase();
        if (
            upperRow.includes('HỘI ĐỒNG THI') ||
            upperRow.includes('XÉT CNTN') ||
            upperRow.includes('TRƯỞNG KHOA') ||
            upperRow.includes('BAN THƯ KÝ') ||
            upperRow.includes('ĐÀ NẴNG, NGÀY') ||
            upperRow.includes('CT. HỘI ĐỒNG')
        ) {
            return;
        }

        let maSV = '', hoTen = '', lop = '', deTai = '', gvhd = '', maGVHD = '', nganhCol = '';

        if (hasCustomColMap) {
            maSV = String(row[colMap.maSV] || '').trim();
            if (colMap.hoDem >= 0 && colMap.ten >= 0) {
                hoTen = `${String(row[colMap.hoDem] || '').trim()} ${String(row[colMap.ten] || '').trim()}`
                    .replace(/\s+/g, ' ')
                    .trim();
            } else {
                hoTen = String(row[colMap.hoTen] || '').trim();
            }
            lop = String(row[colMap.lop] || '').trim();
            deTai = String(row[colMap.deTai] || '').trim();
            gvhd = String(row[colMap.gvhd] || '').trim();
            maGVHD = colMap.maGVHD >= 0 ? String(row[colMap.maGVHD] || '').trim() : '';
            nganhCol = colMap.nganh >= 0 ? String(row[colMap.nganh] || '').trim().toUpperCase() : '';
        } else {
            maSV = String(row[1] || '').trim();
            const col2 = String(row[2] || '').trim();
            const col3 = String(row[3] || '').trim();
            if (col3 && !col3.startsWith('K2') && !col3.startsWith('K1') && !col3.startsWith('K3')) {
                hoTen = `${col2} ${col3}`.replace(/\s+/g, ' ').trim();
                lop = String(row[4] || '').trim();
                deTai = String(row[5] || '').trim();
                gvhd = String(row[6] || '').trim();
                maGVHD = String(row[7] || '').trim();
            } else {
                hoTen = col2;
                lop = col3;
                deTai = String(row[4] || '').trim();
                gvhd = String(row[5] || '').trim();
                maGVHD = String(row[6] || '').trim();
            }
        }

        const isHeaderLike =
            maSV.toUpperCase().includes('MSSV') ||
            maSV.toUpperCase().includes('MÃ') ||
            hoTen.toUpperCase().includes('HỌ VÀ TÊN') ||
            hoTen.toUpperCase().includes('HỌ TÊN');

        if (isHeaderLike || (!maSV && !hoTen && !gvhd && !maGVHD)) return;

        if (hoTen && !maSV) {
            missingCodeIssues.push(`Dòng ${idx + 1} (File SV): Sinh viên "${hoTen}" bị thiếu Mã SV.`);
        }

        if (gvhd && !maGVHD) {
            missingCodeIssues.push(`Dòng ${idx + 1} (File SV): GVHD "${gvhd}"${hoTen ? ` (SV: ${hoTen})` : ''} bị thiếu Mã GV.`);
        }

        if (!maSV && !hoTen) return;

        const nganh = nganhCol || detectedNganhFromHeader || detectNganhFromLop(lop);
        result.push({ maSV, hoTen, lop, deTai, maGVHD, gvhd, nganh });
    });

    result.__missingCodeIssues = missingCodeIssues;
    return result;
}

// 6. Bóc tách Danh sách Giảng viên
export function parseGiangVienRows(rows) {
    const result = [];
    const missingCodeIssues = [];

    rows.forEach((row, idx) => {
        if (idx < 1 || !row || row.length === 0) return;
        const col0 = String(row[0] || '').trim();
        const hoTenGV = String(row[1] || '').trim();
        const donVi = String(row[2] || 'Đại học Duy Tân').trim();

        if (
            col0.toUpperCase().includes('STT') ||
            col0.toUpperCase().includes('MÃ') ||
            hoTenGV.toUpperCase().includes('HỌ VÀ TÊN')
        ) {
            return;
        }

        if (!col0 && !hoTenGV) return;

        const isAutoId = !col0 || col0.toUpperCase().startsWith('GV_');
        if (hoTenGV && isAutoId) {
            missingCodeIssues.push(`Dòng ${idx + 1} (File GV): Giảng viên "${hoTenGV}" bị thiếu Mã GV.`);
        }

        const maGV = col0 || `GV_${idx}`;
        if (hoTenGV) result.push({ maGV, hoTenGV, donVi });
    });

    result.__missingCodeIssues = missingCodeIssues;
    return result;
}

// 7. Bóc tách Giờ, Ngày, Phòng, Cơ sở
function extractScheduleAndPlace(rawStr) {
    const text = String(rawStr || '').trim();
    let gio = '';
    let ngay = '';
    let phong = '';
    let diaDiem = '';

    const timeMatch = text.match(/(\d{1,2})\s*[:hH]\s*(\d{2})/);
    if (timeMatch) {
        gio = `${timeMatch[1].padStart(2, '0')}:${timeMatch[2]}`;
    }

    const dateDMY = text.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    const dateYMD = text.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (dateDMY) {
        ngay = `${dateDMY[3]}-${dateDMY[2].padStart(2, '0')}-${dateDMY[1].padStart(2, '0')}`;
    } else if (dateYMD) {
        ngay = `${dateYMD[1]}-${dateYMD[2].padStart(2, '0')}-${dateYMD[3].padStart(2, '0')}`;
    }

    const roomMatch = text.match(/(?:Phòng|P\.)\s*([A-Za-z0-9\-\.]+)/i);
    if (roomMatch && !roomMatch[1].includes('...')) {
        phong = roomMatch[1].trim();
    }

    const campuses = [
        '120 Hoàng Minh Thảo',
        '254 Nguyễn Văn Linh',
        '137 Nguyễn Văn Linh',
        '03 Quang Trung',
        'K7/25 Quang Trung',
        'Phan Văn Trị'
    ];
    for (const c of campuses) {
        if (text.toLowerCase().includes(c.toLowerCase())) {
            diaDiem = c;
            break;
        }
    }

    return { gio, ngay, phong, diaDiem };
}

export function getSessionKey(gioStr) {
    if (!gioStr || !gioStr.includes(':')) return 'MAC_DINH';
    const hour = parseInt(gioStr.split(':')[0], 10);
    if (isNaN(hour)) return 'MAC_DINH';
    if (hour < 12) return 'SANG';
    if (hour < 17) return 'CHIEU';
    return 'TOI';
}

export function buildMembersSignature(thanhVienList) {
    const cleanName = (str) => String(str || '').toLowerCase()
        .replace(/gs\.?|pgs\.?|ts\.?|ths\.?|gsv\.?|ks\.?|cn\.?|gv\.?|prof\.?|dr\.?/g, '')
        .replace(/\s+/g, ' ')
        .trim();

    const getRoleKey = (cd) => {
        const l = String(cd || '').toLowerCase();
        if (l.includes('chủ tịch')) return 'CT';
        if (l.includes('phản biện')) return 'PB';
        if (l.includes('thư ký')) return 'TK';
        return l;
    };

    const parts = (thanhVienList || []).map(gv => {
        const role = getRoleKey(gv.chucDanh);
        const id = String(gv.maGV || '').trim().toUpperCase();
        const validId = (id && !id.startsWith('GV_')) ? id : cleanName(gv.hoTenGV);
        return `${role}:${validId}`;
    }).sort();

    return parts.join('|');
}

// 8. Bóc tách file Excel Danh sách Hội đồng
// LƯU Ý: Không đọc/không phụ thuộc cột A; tách mỗi HĐ theo chức danh "Chủ tịch", sau khi render xong mới đánh lại số HĐ.
export function parseHoiDongRows(rows) {
    const rawList = [];
    const missingCodeIssues = [];
    let currentObj = null;
    let defaultHeaderPlace = { gio: '', ngay: '', phong: '', diaDiem: '' };

    for (let i = 0; i < Math.min(rows.length, 8); i++) {
        const rowStr = (rows[i] || []).join(' ');
        if (rowStr.toLowerCase().includes('thời gian bảo vệ')) {
            defaultHeaderPlace = extractScheduleAndPlace(rowStr);
            break;
        }
    }

    const normalizeRole = (cd) => {
        const l = String(cd || '').toLowerCase().trim();
        if (l.includes('chủ tịch')) return 'Chủ tịch';
        if (l.includes('phản biện')) return 'Phản biện';
        if (l.includes('thư ký')) return 'Thư ký';
        return '';
    };

    rows.forEach((row, idx) => {
        if (!row || row.length === 0) return;

        // Bỏ qua cột A (row[0]) khi xác định HĐ
        const maGV = String(row[1] || '').trim();
        const hoTenGV = String(row[2] || '').trim();
        const donVi = String(row[3] || '').trim();
        const ngayBV = String(row[4] || '').trim();
        const chucDanh = String(row[5] || '').trim();
        const deTai = String(row[6] || '').trim();
        const maSV = String(row[7] || '').trim();
        const hoTenSV = String(row[8] || '').trim();
        const lop = String(row[9] || '').trim();

        const upperRow = row.map(c => String(c || '').trim()).join(' ').toUpperCase();

        // Chặn dòng tiêu đề và phần chữ ký cuối trang
        if (
            upperRow.includes('MÃ SỐ GV') ||
            upperRow.includes('MÃ GV') ||
            upperRow.includes('CHỨC DANH') ||
            upperRow.includes('TÊN ĐỀ TÀI') ||
            upperRow.includes('TRƯỞNG KHOA') ||
            upperRow.includes('BAN THƯ KÝ') ||
            upperRow.includes('HỘI ĐỒNG THI') ||
            upperRow.includes('XÉT CNTN') ||
            upperRow.includes('ĐÀ NẴNG, NGÀY') ||
            upperRow.includes('CT. HỘI ĐỒNG')
        ) {
            return;
        }

        if (!maGV && !hoTenGV && !chucDanh && !maSV && !hoTenSV && !deTai) return;

        const roleKey = normalizeRole(chucDanh);

        // QUY TẮC TÁCH HĐ: Cứ gặp chức danh "Chủ tịch" (hoặc chưa có block nào) là bắt đầu 1 khối HĐ mới
        if (roleKey === 'Chủ tịch' || (!currentObj && (maGV || hoTenGV || maSV || hoTenSV))) {
            const parsedInfo = extractScheduleAndPlace(ngayBV);
            const seqIndex = String(rawList.length + 1);
            currentObj = {
                rawIndex: seqIndex,
                soHD: seqIndex,
                ngayBV: parsedInfo.ngay || defaultHeaderPlace.ngay || '',
                gioBV: parsedInfo.gio || defaultHeaderPlace.gio || '',
                phongBV: parsedInfo.phong || defaultHeaderPlace.phong || '',
                diaDiemBV: parsedInfo.diaDiem || defaultHeaderPlace.diaDiem || '',
                thanhVien: [],
                sinhViens: []
            };
            rawList.push(currentObj);
        }

        if (currentObj) {
            const hdLabel = `Khối HĐ #${currentObj.rawIndex} (dòng ${idx + 1})`;

            if (ngayBV) {
                const parsedInfo = extractScheduleAndPlace(ngayBV);
                if (!currentObj.ngayBV && parsedInfo.ngay) currentObj.ngayBV = parsedInfo.ngay;
                if (!currentObj.gioBV && parsedInfo.gio) currentObj.gioBV = parsedInfo.gio;
                if (!currentObj.phongBV && parsedInfo.phong) currentObj.phongBV = parsedInfo.phong;
                if (!currentObj.diaDiemBV && parsedInfo.diaDiem) currentObj.diaDiemBV = parsedInfo.diaDiem;
            }

            if (hoTenGV && !maGV) {
                missingCodeIssues.push(`Dòng ${idx + 1} (${hdLabel}): Giảng viên "${hoTenGV}"${roleKey ? ` (${roleKey})` : ''} bị thiếu Mã GV.`);
            }

            if (maGV && hoTenGV && roleKey) {
                const existRole = currentObj.thanhVien.find(g => normalizeRole(g.chucDanh) === roleKey);
                if (!existRole) {
                    currentObj.thanhVien.push({ maGV, hoTenGV, donVi, chucDanh: roleKey });
                }
            }

            // Chỉ cảnh báo thiếu Mã SV nếu dòng có tên SV (hoTenSV); nếu HĐ chưa có SV thì bỏ qua
            if (hoTenSV && !maSV) {
                missingCodeIssues.push(`Dòng ${idx + 1} (${hdLabel}): Sinh viên "${hoTenSV}" bị thiếu Mã SV.`);
            }

            if (maSV && hoTenSV) {
                const nganh = detectNganhFromLop(lop);
                currentObj.sinhViens.push({ maSV, hoTenSV, lop, deTai, nganh });
            }
        }
    });

    const validBlocks = rawList.filter(item => item.thanhVien.length > 0 || item.sinhViens.length > 0);

    const resultMap = {};
    validBlocks.forEach((hd, idx) => {
        const num = String(idx + 1);
        resultMap[`hd_${idx}`] = {
            ...hd,
            rawIndex: num,
            soHD: num
        };
    });
    resultMap.__mergeLogs = [];
    resultMap.__missingCodeIssues = missingCodeIssues;

    return resultMap;
}
