// js/excel-reader.js
import { getAdminBinaryFile, getAdminConfig } from './admin-config.js';
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

        // Tự động dò dòng tiêu đề và vị trí cột (Mã | Chuyên ngành | Khoa)
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

// 2. Tra cứu thông tin { chuyenNganh, khoa } từ chuỗi Lớp (VD: "K27 CTP" -> khớp Ma_Chuyen_nganh "CTP")
export function findNganhInfoByLop(lopStr) {
    const cleanLop = String(lopStr || '').toUpperCase().trim();
    if (!cleanLop || danhMucNganh.length === 0) return null;

    const tokens = cleanLop.split(/[^A-Z0-9]+/).filter(Boolean);

    // Ưu tiên khớp chính xác token (VD: "CTP" === "CTP")
    for (const item of danhMucNganh) {
        if (tokens.includes(item.maChuyenNganh)) {
            return item;
        }
    }

    // Dự phòng: khớp chuỗi con (sắp xếp mã dài trước để tránh "MT" đè "QLMT")
    const sorted = [...danhMucNganh].sort((a, b) => b.maChuyenNganh.length - a.maChuyenNganh.length);
    for (const item of sorted) {
        if (cleanLop.includes(item.maChuyenNganh)) {
            return item;
        }
    }

    return null;
}

// 3. Hàm duy nhất dịch mã Lớp sang tên Ngành chuẩn từ danhMucNganh
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

// 5. Bóc tách Danh sách Sinh viên & Đề tài (Tự động nhận diện cột Mã GVHD / Mã GV / Mã Giảng viên)
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
            break;
        }
    }

    let startRow = headerIdx >= 0 ? headerIdx + 1 : 1;
    const result = [];

    rows.forEach((row, idx) => {
        if (idx < startRow || !row || row.length === 0) return;
        const maSV = String(row[colMap.maSV] || '').trim();
        if (!maSV || maSV.toUpperCase().includes('MSSV') || maSV.toUpperCase().includes('MÃ')) return;

        let hoTen = '', lop = '', deTai = '', gvhd = '', maGVHD = '', nganhCol = '';

        if (hasCustomColMap) {
            if (colMap.hoDem >= 0 && colMap.ten >= 0) {
                hoTen = `${String(row[colMap.hoDem] || '').trim()} ${String(row[colMap.ten] || '').trim()}`.trim();
            } else {
                hoTen = String(row[colMap.hoTen] || '').trim();
            }
            lop = String(row[colMap.lop] || '').trim();
            deTai = String(row[colMap.deTai] || '').trim();
            gvhd = String(row[colMap.gvhd] || '').trim();
            maGVHD = colMap.maGVHD >= 0 ? String(row[colMap.maGVHD] || '').trim() : '';
            nganhCol = colMap.nganh >= 0 ? String(row[colMap.nganh] || '').trim().toUpperCase() : '';
        } else {
            // Dự phòng nếu file không có dòng tiêu đề chuẩn
            const col2 = String(row[2] || '').trim();
            const col3 = String(row[3] || '').trim();
            if (col3 && !col3.startsWith('K2') && !col3.startsWith('K1') && !col3.startsWith('K3')) {
                hoTen = `${col2} ${col3}`.trim();
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

        const nganh = nganhCol || detectedNganhFromHeader || detectNganhFromLop(lop);
        result.push({ maSV, hoTen, lop, deTai, maGVHD, gvhd, nganh });
    });

    return result;
}

// 6. Bóc tách Danh sách Giảng viên
export function parseGiangVienRows(rows) {
    const result = [];
    rows.forEach((row, idx) => {
        if (idx < 1 || !row || row.length === 0) return;
        const col0 = String(row[0] || '').trim();
        if (col0.toUpperCase().includes('STT') || col0.toUpperCase().includes('MÃ')) return;

        const maGV = col0 || `GV_${idx}`;
        const hoTenGV = String(row[1] || '').trim();
        const donVi = String(row[2] || 'Đại học Duy Tân').trim();

        if (hoTenGV) result.push({ maGV, hoTenGV, donVi });
    });
    return result;
}

// 7. Hàm phụ trợ bóc tách Giờ, Ngày, Phòng, Cơ sở từ chuỗi thời gian & địa điểm (Không tự gán mặc định)
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

// 8. Bóc tách file Excel Danh sách Hội đồng
export function parseHoiDongRows(rows) {
    const hoidongs = {};
    let currentHD = null;
    let defaultHeaderPlace = { gio: '', ngay: '', phong: '', diaDiem: '' };

    for (let i = 0; i < Math.min(rows.length, 8); i++) {
        const rowStr = (rows[i] || []).join(' ');
        if (rowStr.toLowerCase().includes('thời gian bảo vệ')) {
            defaultHeaderPlace = extractScheduleAndPlace(rowStr);
            break;
        }
    }

    rows.forEach((row, idx) => {
        if (idx < 1 || !row || row.length === 0) return;

        const colA = String(row[0] || '').trim();
        const maGV = String(row[1] || '').trim();
        const hoTenGV = String(row[2] || '').trim();
        const donVi = String(row[3] || '').trim();
        const ngayBV = String(row[4] || '').trim();
        const chucDanh = String(row[5] || '').trim();
        const deTai = String(row[6] || '').trim();
        const maSV = String(row[7] || '').trim();
        const hoTenSV = String(row[8] || '').trim();
        const lop = String(row[9] || '').trim();

        const upperRow = row.join(' ').toUpperCase();

        // Chặn dòng tiêu đề bảng và toàn bộ các dòng chữ ký Footer ở cuối file
        if (
            colA.toUpperCase().includes('HỘI ĐỒNG') ||
            maGV.toUpperCase().includes('MÃ') ||
            upperRow.includes('TRƯỞNG KHOA') ||
            upperRow.includes('BAN THƯ KÝ') ||
            upperRow.includes('HỘI ĐỒNG THI') ||
            upperRow.includes('XÉT CNTN') ||
            upperRow.includes('ĐÀ NẴNG, NGÀY')
        ) {
            return;
        }

        if (colA && !isNaN(colA)) {
            currentHD = colA;
        }

        if (currentHD) {
            if (!hoidongs[currentHD]) {
                const parsedInfo = extractScheduleAndPlace(ngayBV);
                hoidongs[currentHD] = {
                    soHD: `Hội Đồng ${currentHD}`,
                    ngayBV: parsedInfo.ngay || defaultHeaderPlace.ngay || '',
                    gioBV: parsedInfo.gio || defaultHeaderPlace.gio || '',
                    phongBV: parsedInfo.phong || defaultHeaderPlace.phong || '',
                    diaDiemBV: parsedInfo.diaDiem || defaultHeaderPlace.diaDiem || '',
                    thanhVien: [],
                    sinhViens: []
                };
            }

            if (maGV && hoTenGV) {
                const existGV = hoidongs[currentHD].thanhVien.find(g => g.maGV === maGV);
                if (!existGV) {
                    hoidongs[currentHD].thanhVien.push({ maGV, hoTenGV, donVi, chucDanh });
                }
            }

            // Chỉ nhận Sinh viên khi có cả Mã SV và Họ tên SV hợp lệ
            if (maSV && hoTenSV) {
                const existSV = hoidongs[currentHD].sinhViens.find(s => s.maSV === maSV);
                if (!existSV) {
                    const nganh = detectNganhFromLop(lop);
                    hoidongs[currentHD].sinhViens.push({ maSV, hoTenSV, lop, deTai, nganh });
                }
            }
        }
    });

    return hoidongs;
}
