// js/excel-reader.js

export function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.SheetNames[0];
                const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1 });
                resolve(jsonData);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

export function parseSinhVienRows(rows) {
    let headerIdx = -1;
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
        const rowStr = (rows[i] || []).join(' ').toUpperCase();
        if (rowStr.includes('MSSV') || rowStr.includes('MÃ SỐ SV')) {
            headerIdx = i; break;
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
        let hoTen = '', lop = '', deTai = '', gvhd = '';

        if (col3 && !col3.startsWith('K2') && !col3.startsWith('K1') && !col3.startsWith('K3')) {
            hoTen = `${col2} ${col3}`.trim();
            lop = String(row[4] || '').trim();
            deTai = String(row[5] || '').trim();
            gvhd = String(row[6] || '').trim();
        } else {
            hoTen = col2;
            lop = col3;
            deTai = String(row[4] || '').trim();
            gvhd = String(row[5] || '').trim();
        }

        result.push({ maSV, hoTen, lop, deTai, gvhd });
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

// BỔ SUNG: Bóc tách file Excel Danh sách Hội đồng (10 Cột đầy đủ thông tin)
export function parseHoiDongRows(rows) {
    const hoidongs = {};
    let currentHD = null;

    rows.forEach((row, idx) => {
        if (idx < 1 || !row || row.length === 0) return;

        const colA = String(row[0] || '').trim();      // HỘI ĐỒNG (1, 2, 3...)
        const maGV = String(row[1] || '').trim();      // MÃ SỐ GV
        const hoTenGV = String(row[2] || '').trim();   // HỌ VÀ TÊN GV
        const donVi = String(row[3] || '').trim();     // ĐƠN VỊ CÔNG TÁC
        const ngayBV = String(row[4] || '').trim();    // NGÀY BẢO VỆ
        const chucDanh = String(row[5] || '').trim();  // CHỨC DANH
        const deTai = String(row[6] || '').trim();     // TÊN ĐỀ TÀI ĐỒ ÁN
        const maSV = String(row[7] || '').trim();      // MÃ SỐ SV
        const hoTenSV = String(row[8] || '').trim();   // HỌ VÀ TÊN SV
        const lop = String(row[9] || '').trim();       // LỚP

        // Bỏ qua dòng tiêu đề
        if (colA.toUpperCase().includes('HỘI ĐỒNG') || maGV.toUpperCase().includes('MÃ')) return;

        if (colA && !isNaN(colA)) {
            currentHD = colA;
        }

        if (currentHD) {
            if (!hoidongs[currentHD]) {
                hoidongs[currentHD] = {
                    soHD: `Hội Đồng ${currentHD}`,
                    ngayBV: ngayBV,
                    thanhVien: [],
                    sinhViens: []
                };
            }

            // Lưu thông tin Giảng viên theo Mã GV
            if (maGV && hoTenGV) {
                const existGV = hoidongs[currentHD].thanhVien.find(g => g.maGV === maGV);
                if (!existGV) {
                    hoidongs[currentHD].thanhVien.push({ maGV, hoTenGV, donVi, chucDanh });
                }
            }

            // Lưu thông tin Sinh viên theo Mã SV
            if (maSV) {
                const existSV = hoidongs[currentHD].sinhViens.find(s => s.maSV === maSV);
                if (!existSV) {
                    hoidongs[currentHD].sinhViens.push({ maSV, hoTenSV, lop, deTai });
                }
            }
        }
    });

    return hoidongs;
}
