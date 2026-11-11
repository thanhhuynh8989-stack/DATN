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
