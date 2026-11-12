// js/excel-reader.js

export let danhMucNganh = []; // Mảng lưu { maChuyenNganh, chuyenNganh, khoa }

// Tự động nạp file DANH_SACH_NGANH.xlsx khi khởi động ứng dụng
export async function loadDanhSachNganhSilently(url = './sample_files/DANH_SACH_NGANH.xlsx') {
    try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`Không tìm thấy file ${url}`);
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1 });

        const parsed = [];
        rows.forEach((row, idx) => {
            if (idx < 1 || !row || row.length < 3) return;
            const ma = String(row[0] || '').trim().toUpperCase();
            const tenNganh = String(row[1] || '').trim();
            const tenKhoa = String(row[2] || '').trim();

            if (!ma || ma.includes('MA_CHUYEN_NGANH') || ma.includes('MÃ')) return;
            if (tenNganh && tenKhoa) {
                parsed.push({
                    maChuyenNganh: ma,
                    chuyenNganh: tenNganh,
                    khoa: tenKhoa
                });
            }
        });

        danhMucNganh = parsed;
        return danhMucNganh;
    } catch (err) {
        console.warn("Cảnh báo tải DANH_SACH_NGANH.xlsx:", err.message);
        return [];
    }
}

// Tra cứu thông tin { chuyenNganh, khoa } từ chuỗi Lớp (VD: "K27 CTP" -> khớp Ma_Chuyen_nganh "CTP")
export function findNganhInfoByLop(lopStr) {
    const cleanLop = String(lopStr || '').toUpperCase().trim();
    if (!cleanLop || danhMucNganh.length === 0) return null;

    // Tách các token chữ cái trong mã lớp (VD: "K27 CTP" -> ["K", "CTP"])
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

export function detectNganhFromLop(lopStr) {
    const found = findNganhInfoByLop(lopStr);
    return found ? found.chuyenNganh.toUpperCase() : '';
}

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

// Hàm tự động dịch mã Lớp sang tên Ngành chuẩn
export function detectNganhFromLop(lopStr) {
    const u = String(lopStr || '').toUpperCase();
    if (u.includes('CTP') || u.includes('CNTP')) return 'CÔNG NGHỆ THỰC PHẨM';
    if (u.includes('KTD')) return 'KỸ THUẬT ĐIỆN';
    if (u.includes('TDH')) return 'KỸ THUẬT ĐIỀU KHIỂN & TỰ ĐỘNG HÓA';
    if (u.includes('DTVT') || u.includes('KTVT')) return 'KỸ THUẬT ĐIỆN TỬ - VIỄN THÔNG';
    if (u.includes('MT') || u.includes('QLMT')) return 'QUẢN LÝ TÀI NGUYÊN & MÔI TRƯỜNG';
    if (u.includes('KTSH') || u.includes('CNSH')) return 'CÔNG NGHỆ SINH HỌC';
    return '';
}

export function parseSinhVienRows(rows) {
    let headerIdx = -1;
    let detectedNganhFromHeader = '';

    for (let i = 0; i < Math.min(rows.length, 15); i++) {
        const rowStr = (rows[i] || []).join(' ').trim();
        const upperStr = rowStr.toUpperCase();

        // Tìm tên Ngành nếu file Excel SV có dòng "NGÀNH: ..." ở phần tiêu đề
        const nganhMatch = rowStr.match(/NGÀNH\s*[:\-]\s*([^\n\r]+)/i);
        if (nganhMatch && nganhMatch[1]) {
            detectedNganhFromHeader = nganhMatch[1].trim().toUpperCase();
        }

        if (upperStr.includes('MSSV') || upperStr.includes('MÃ SỐ SV')) {
            headerIdx = i;
            break;
        }
    }

    let startRow = headerIdx >= 0 ? headerIdx + 1 : 1;
    const result = [];

    rows.forEach((row, idx) => {
        if (idx < startRow || !row || row.length === 0) return;
        const maSV = String(row[1] || '').trim();
        if (!maSV || maSV.toUpperCase().includes('MSSV')) return;

        const col2 = String(row[2] || '').trim();
        const col3 = String(row[3] || '').trim();
        let hoTen = '', lop = '', deTai = '', gvhd = '', nganhCol = '';

        if (col3 && !col3.startsWith('K2') && !col3.startsWith('K1') && !col3.startsWith('K3')) {
            hoTen = `${col2} ${col3}`.trim();
            lop = String(row[4] || '').trim();
            deTai = String(row[5] || '').trim();
            gvhd = String(row[6] || '').trim();
            nganhCol = String(row[7] || '').trim().toUpperCase();
        } else {
            hoTen = col2;
            lop = col3;
            deTai = String(row[4] || '').trim();
            gvhd = String(row[5] || '').trim();
            nganhCol = String(row[6] || '').trim().toUpperCase();
        }

        // Ưu tiên: Cột Ngành trong file -> Dòng header NGÀNH -> Dịch từ mã Lớp
        const nganh = nganhCol || detectedNganhFromHeader || detectNganhFromLop(lop);

        result.push({ maSV, hoTen, lop, deTai, gvhd, nganh });
    });
    return result;
}

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

// Hàm phụ trợ bóc tách Giờ, Ngày, Phòng, Cơ sở từ chuỗi thời gian & địa điểm
function extractScheduleAndPlace(rawStr) {
    const text = String(rawStr || '').trim();
    let gio = '17:00';
    let ngay = '';
    let phong = '';
    let diaDiem = '';

    // 1. Bắt Giờ (ví dụ: 17:00 hoặc 17h00 hoặc 7h30)
    const timeMatch = text.match(/(\d{1,2})\s*[:hH]\s*(\d{2})/);
    if (timeMatch) {
        gio = `${timeMatch[1].padStart(2, '0')}:${timeMatch[2]}`;
    }

    // 2. Bắt Ngày (dd/mm/yyyy hoặc dd-mm-yyyy hoặc yyyy-mm-dd)
    const dateDMY = text.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    const dateYMD = text.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (dateDMY) {
        ngay = `${dateDMY[3]}-${dateDMY[2].padStart(2, '0')}-${dateDMY[1].padStart(2, '0')}`;
    } else if (dateYMD) {
        ngay = `${dateYMD[1]}-${dateYMD[2].padStart(2, '0')}-${dateYMD[3].padStart(2, '0')}`;
    }

    // 3. Bắt Phòng (sau chữ "Phòng" hoặc "P.")
    const roomMatch = text.match(/(?:Phòng|P\.)\s*([A-Za-z0-9\-\.]+)/i);
    if (roomMatch && !roomMatch[1].includes('...')) {
        phong = roomMatch[1].trim();
    }

    // 4. Bắt Cơ sở (khớp với danh sách cơ sở ĐH Duy Tân)
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

// Bóc tách file Excel Danh sách Hội đồng (10 Cột + Đọc cả dòng Thời gian/Địa điểm ở header)
export function parseHoiDongRows(rows) {
    const hoidongs = {};
    let currentHD = null;
    let defaultHeaderPlace = { gio: '17:00', ngay: '', phong: '', diaDiem: '' };

    // Quét 8 dòng đầu xem có dòng "Thời gian bảo vệ: ... tại Phòng ... cơ sở ..." không
    for (let i = 0; i < Math.min(rows.length, 8); i++) {
        const rowStr = (rows[i] || []).join(' ');
        if (rowStr.toLowerCase().includes('thời gian bảo vệ')) {
            defaultHeaderPlace = extractScheduleAndPlace(rowStr);
            break;
        }
    }

    rows.forEach((row, idx) => {
        if (idx < 1 || !row || row.length === 0) return;

        const colA = String(row[0] || '').trim();       // HỘI ĐỒNG (1, 2, 3...)
        const maGV = String(row[1] || '').trim();       // MÃ SỐ GV
        const hoTenGV = String(row[2] || '').trim();    // HỌ VÀ TÊN GV
        const donVi = String(row[3] || '').trim();      // ĐƠN VỊ CÔNG TÁC
        const ngayBV = String(row[4] || '').trim();     // NGÀY BẢO VỆ (có thể chứa cả giờ, phòng, cơ sở)
        const chucDanh = String(row[5] || '').trim();   // CHỨC DANH
        const deTai = String(row[6] || '').trim();      // TÊN ĐỀ TÀI ĐỒ ÁN
        const maSV = String(row[7] || '').trim();       // MÃ SỐ SV
        const hoTenSV = String(row[8] || '').trim();    // HỌ VÀ TÊN SV
        const lop = String(row[9] || '').trim();        // LỚP

        if (colA.toUpperCase().includes('HỘI ĐỒNG') || maGV.toUpperCase().includes('MÃ')) return;

        if (colA && !isNaN(colA)) {
            currentHD = colA;
        }

        if (currentHD) {
            if (!hoidongs[currentHD]) {
                const parsedInfo = extractScheduleAndPlace(ngayBV);
                hoidongs[currentHD] = {
                    soHD: `Hội Đồng ${currentHD}`,
                    ngayBV: parsedInfo.ngay || defaultHeaderPlace.ngay,
                    gioBV: parsedInfo.gio || defaultHeaderPlace.gio || '17:00',
                    phongBV: parsedInfo.phong || defaultHeaderPlace.phong,
                    diaDiemBV: parsedInfo.diaDiem || defaultHeaderPlace.diaDiem || '120 Hoàng Minh Thảo',
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

            if (maSV) {
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
