// js/payment-exporter.js
import { normalizeTeacherKeyForDupCheck } from './validators.js';

const STORAGE_KEY_PAYMENT_CFG = 'datn_payment_export_cfg_v1';

const DEFAULT_PAYMENT_CONFIG = {
    truongThanhVien: 'TRƯỜNG CN&KT',
    donGia: 185000,
    heSoHD: 18,
    heSoPB: 1.6,
    heSoCT: 0.84,
    heSoHDPB: 0.7,
    heSoHDTK: 0.77
};

function loadSavedPaymentConfig() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY_PAYMENT_CFG);
        if (!raw) return { ...DEFAULT_PAYMENT_CONFIG };
        return { ...DEFAULT_PAYMENT_CONFIG, ...JSON.parse(raw) };
    } catch {
        return { ...DEFAULT_PAYMENT_CONFIG };
    }
}

function savePaymentConfig(cfg) {
    try {
        localStorage.setItem(STORAGE_KEY_PAYMENT_CFG, JSON.stringify({
            truongThanhVien: cfg.truongThanhVien,
            donGia: cfg.donGia,
            heSoHD: cfg.heSoHD,
            heSoPB: cfg.heSoPB,
            heSoCT: cfg.heSoCT,
            heSoHDPB: cfg.heSoHDPB,
            heSoHDTK: cfg.heSoHDTK
        }));
    } catch {
        // Bỏ qua nếu trình duyệt chặn storage
    }
}

function escapeHtml(str) {
    return String(str ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function formatVND(num) {
    return Math.round(Number(num) || 0).toLocaleString('vi-VN');
}

function formatDecimalVN(num) {
    const val = Math.round((Number(num) || 0) * 100) / 100;
    return val.toLocaleString('vi-VN', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

// Tách Họ đệm (kèm học hàm/học vị) và Tên để đưa vào 2 cột theo mẫu Duy Tân
function splitTeacherNameAndTitle(rawFullName) {
    const clean = String(rawFullName || '').trim().replace(/\s+/g, ' ');
    if (!clean) return { hoDem: '', ten: '', hocVi: '' };

    const titleMatch = clean.match(/^((?:PGS\.?\s*TS\.?|GS\.?\s*TSKH\.?|GS\.?\s*TS\.?|TSKH\.?|TS\.?|Th\.?\s*S\.?|GVC\.?|GV\.?|KS\.?|CN\.?|BS\.?|Dr\.?|Prof\.?)+)\s*(.*)$/i);
    let hocVi = '';
    if (titleMatch && titleMatch[1]) {
        hocVi = titleMatch[1].replace(/\.+$/, '').replace(/\s+/g, ' ').trim();
    }

    const parts = clean.split(' ');
    if (parts.length === 1) {
        return { hoDem: '', ten: parts[0], hocVi };
    }
    const ten = parts[parts.length - 1];
    const hoDem = parts.slice(0, -1).join(' ');
    return { hoDem, ten, hocVi };
}

function extractHoiDongNumber(rawTitle, fallbackIdx) {
    const cleaned = String(rawTitle || '').trim();
    if (!cleaned) return String(fallbackIdx);
    const stripped = cleaned
        .replace(/^(?:hội\s*đồng|hoi\s*dong|hđ|hd)\s*(?:số|so|#|:|-|\.)?\s*/i, '')
        .trim();
    if (stripped) return stripped;
    const m = cleaned.match(/\d+/);
    return m ? m[0] : String(fallbackIdx);
}

function formatDateVN(rawDate) {
    const s = String(rawDate || '').trim();
    if (!s) return '';
    if (s.includes('-')) {
        const p = s.split('-');
        if (p.length === 3) return `${p[2].padStart(2, '0')}/${p[1].padStart(2, '0')}/${p[0]}`;
    }
    return s;
}

function sanitizeFileSlug(str) {
    return String(str || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd')
        .replace(/Đ/g, 'D')
        .replace(/[^a-zA-Z0-9]+/g, '_')
        .replace(/^_+|_+$/g, '');
}

/**
 * Thu thập dữ liệu nền từ DOM Hội đồng:
 * - Chỉ tính SV đã xếp vào Hội đồng
 * - Tự động lấy Ngành & tập hợp Lớp của tất cả SV làm Khóa
 * - Gom HĐ theo Số Hội đồng (không phân biệt ngày, địa điểm)
 */
export function collectPaymentSourceFromDOM(meta = {}) {
    const khoGVMap = new Map();
    document.querySelectorAll('#pool-gv .card-item').forEach(el => {
        const maGV = (el.getAttribute('data-id') || '').trim();
        const hoTen = (el.getAttribute('data-name') || '').trim();
        const donVi = (el.getAttribute('data-donvi') || '').trim();
        const key = normalizeTeacherKeyForDupCheck(hoTen) || maGV.toUpperCase();
        if (key) {
            khoGVMap.set(key, { maGV, hoTen, donVi });
        }
    });

    const resolveTeacherMeta = (maGV, hoTen, fallbackDonVi = '') => {
        const key = normalizeTeacherKeyForDupCheck(hoTen) || String(maGV || '').trim().toUpperCase();
        const fromPool = key ? khoGVMap.get(key) : null;
        const finalId = (maGV && !String(maGV).toUpperCase().startsWith('GV_'))
            ? String(maGV).trim()
            : (fromPool?.maGV || '');
        const finalName = (fromPool?.hoTen && fromPool.hoTen.length > String(hoTen || '').length)
            ? fromPool.hoTen
            : (String(hoTen || '').trim() || fromPool?.hoTen || '');
        const finalDonVi = fallbackDonVi || fromPool?.donVi || (meta.khoa ? `Khoa ${meta.khoa}` : 'Đại học Duy Tân');
        return { maGV: finalId, hoTen: finalName, donVi: finalDonVi };
    };

    const classSet = new Set();
    const nganhSet = new Set();
    const hdGroupedMap = new Map();
    const assignedStudents = [];

    document.querySelectorAll('.box-hoidong').forEach((box, idx) => {
        const svCards = Array.from(box.querySelectorAll('.dropzone-sv .card-item'));
        if (svCards.length === 0) return;

        const rawTitle = box.querySelector('.hd-title')?.value.trim() || '';
        const soHD = extractHoiDongNumber(rawTitle, idx + 1);
        const ngayVN = formatDateVN(box.querySelector('.hd-date')?.value || '');

        const getMember = (role) => {
            const card = box.querySelector(`.dropzone-gv[data-role="${role}"] .card-item`);
            if (!card) return { maGV: '', hoTen: '', donVi: '' };
            return resolveTeacherMeta(
                card.getAttribute('data-id') || '',
                card.getAttribute('data-name') || '',
                card.getAttribute('data-donvi') || ''
            );
        };

        const ct = getMember('Chủ tịch');
        const pb = getMember('Phản biện');
        const tk = getMember('Thư ký');

        if (!hdGroupedMap.has(soHD)) {
            hdGroupedMap.set(soHD, {
                soHD,
                ct,
                pb,
                tk,
                dates: new Set(),
                totalSV: 0
            });
        }
        const hdRec = hdGroupedMap.get(soHD);
        if (!hdRec.ct.hoTen && ct.hoTen) hdRec.ct = ct;
        if (!hdRec.pb.hoTen && pb.hoTen) hdRec.pb = pb;
        if (!hdRec.tk.hoTen && tk.hoTen) hdRec.tk = tk;
        if (ngayVN) hdRec.dates.add(ngayVN);
        hdRec.totalSV += svCards.length;

        svCards.forEach(svEl => {
            const lop = (svEl.getAttribute('data-lop') || '').trim();
            const nganhSV = (svEl.getAttribute('data-nganh') || '').trim();
            if (lop) classSet.add(lop);
            if (nganhSV) nganhSet.add(nganhSV);

            assignedStudents.push({
                maSV: (svEl.getAttribute('data-id') || '').trim(),
                hoTenSV: (svEl.getAttribute('data-name') || '').trim(),
                lop,
                maGVHD: (svEl.getAttribute('data-magvhd') || '').trim(),
                gvhd: (svEl.getAttribute('data-gvhd') || '').trim(),
                soHD,
                pb
            });
        });
    });

    const detectedNganh = (meta.nganh || '').trim() || Array.from(nganhSet).join(', ');
    const detectedKhoaHoc = Array.from(classSet).sort((a, b) => a.localeCompare(b, 'vi')).join(', ');

    return {
        khoa: (meta.khoa || '').trim(),
        nganh: detectedNganh,
        khoaHoc: detectedKhoaHoc,
        assignedStudents,
        hdGroupedList: Array.from(hdGroupedMap.values()).sort((a, b) => {
            const na = parseInt(a.soHD, 10);
            const nb = parseInt(b.soHD, 10);
            if (!isNaN(na) && !isNaN(nb)) return na - nb;
            return String(a.soHD).localeCompare(String(b.soHD), 'vi');
        }),
        resolveTeacherMeta
    };
}

/**
 * Tính toán chi tiết số liệu cho cả 3 bảng thanh toán từ nguồn dữ liệu & cấu hình đơn giá/hệ số
 */
export function computePaymentTables(sourceData, config) {
    const donGia = Number(config.donGia) || 0;
    const heSoHD = Number(config.heSoHD) || 0;
    const heSoPB = Number(config.heSoPB) || 0;
    const heSoCT = Number(config.heSoCT) || 0;
    const heSoHDPB = Number(config.heSoHDPB) || 0;
    const heSoHDTK = Number(config.heSoHDTK) || 0;

    // 1. Bảng Hướng dẫn (chỉ tính SV đã xếp vào HĐ)
    const hdTeacherMap = new Map();
    sourceData.assignedStudents.forEach(sv => {
        const tMeta = sourceData.resolveTeacherMeta(sv.maGVHD, sv.gvhd);
        const key = normalizeTeacherKeyForDupCheck(tMeta.hoTen) || tMeta.maGV.toUpperCase();
        if (!key) return;
        if (!hdTeacherMap.has(key)) {
            const nameParts = splitTeacherNameAndTitle(tMeta.hoTen);
            hdTeacherMap.set(key, {
                maGV: tMeta.maGV,
                hoTenFull: tMeta.hoTen,
                hoDem: nameParts.hoDem,
                ten: nameParts.ten,
                hocVi: nameParts.hocVi,
                donVi: tMeta.donVi,
                slDeTai: 0
            });
        }
        hdTeacherMap.get(key).slDeTai += 1;
    });

    const rowsHuongDan = Array.from(hdTeacherMap.values())
        .sort((a, b) => a.ten.localeCompare(b.ten, 'vi') || a.hoDem.localeCompare(b.hoDem, 'vi'))
        .map((item, idx) => {
            const quyChuan = Math.round(item.slDeTai * heSoHD * 100) / 100;
            const soTien = Math.round(quyChuan * donGia);
            return {
                tt: idx + 1,
                ...item,
                heSo: heSoHD,
                quyChuan,
                donGia,
                soTien
            };
        });

    // 2. Bảng Phản biện (cộng tổng SV theo GV Phản biện ở các HĐ)
    const pbTeacherMap = new Map();
    sourceData.hdGroupedList.forEach(hd => {
        if (!hd.pb || (!hd.pb.hoTen && !hd.pb.maGV)) return;
        const key = normalizeTeacherKeyForDupCheck(hd.pb.hoTen) || hd.pb.maGV.toUpperCase();
        if (!key) return;
        if (!pbTeacherMap.has(key)) {
            const nameParts = splitTeacherNameAndTitle(hd.pb.hoTen);
            pbTeacherMap.set(key, {
                maGV: hd.pb.maGV,
                hoTenFull: hd.pb.hoTen,
                hoDem: nameParts.hoDem,
                ten: nameParts.ten,
                hocVi: nameParts.hocVi,
                donVi: hd.pb.donVi,
                slDeTai: 0
            });
        }
        pbTeacherMap.get(key).slDeTai += hd.totalSV;
    });

    const rowsPhanBien = Array.from(pbTeacherMap.values())
        .sort((a, b) => a.ten.localeCompare(b.ten, 'vi') || a.hoDem.localeCompare(b.hoDem, 'vi'))
        .map((item, idx) => {
            const quyChuan = Math.round(item.slDeTai * heSoPB * 100) / 100;
            const soTien = Math.round(quyChuan * donGia);
            return {
                tt: idx + 1,
                ...item,
                heSo: heSoPB,
                quyChuan,
                donGia,
                soTien
            };
        });

    // 3. Bảng Chi tiết Hội đồng chấm Tốt nghiệp (mỗi HĐ 1 block 3 dòng, SLSV = tổng SV của HĐ)
    const rowsHoiDong = [];
    let sttCounter = 1;
    sourceData.hdGroupedList.forEach(hd => {
        const thoiGianStr = Array.from(hd.dates).join(', ');
        const roles = [
            { chucDanh: 'Chủ tịch', gv: hd.ct, heSo: heSoCT },
            { chucDanh: 'UV-Phản biện', gv: hd.pb, heSo: heSoHDPB },
            { chucDanh: 'UV-Thư ký', gv: hd.tk, heSo: heSoHDTK }
        ];

        roles.forEach((r, roleIdx) => {
            const nameParts = splitTeacherNameAndTitle(r.gv?.hoTen || '');
            const gioQD = Math.round(hd.totalSV * r.heSo * 100) / 100;
            const thanhTien = Math.round(gioQD * donGia);
            rowsHoiDong.push({
                stt: sttCounter++,
                soHD: hd.soHD,
                isFirstInBlock: roleIdx === 0,
                blockSize: 3,
                maGV: r.gv?.maGV || '',
                hoTenFull: r.gv?.hoTen || '',
                hoDem: nameParts.hoDem,
                ten: nameParts.ten,
                donVi: r.gv?.donVi || '',
                chucDanh: r.chucDanh,
                thoiGian: thoiGianStr,
                slsv: hd.totalSV,
                heSo: r.heSo,
                gioQD,
                donGia,
                thanhTien
            });
        });
    });

    const summary = {
        tongSLHuongDan: rowsHuongDan.reduce((s, r) => s + r.slDeTai, 0),
        tongTienHuongDan: rowsHuongDan.reduce((s, r) => s + r.soTien, 0),
        tongSLPhanBien: rowsPhanBien.reduce((s, r) => s + r.slDeTai, 0),
        tongTienPhanBien: rowsPhanBien.reduce((s, r) => s + r.soTien, 0),
        tongSLSVHoiDong: sourceData.hdGroupedList.reduce((s, h) => s + h.totalSV, 0),
        tongGioQDHoiDong: Math.round(rowsHoiDong.reduce((s, r) => s + r.gioQD, 0) * 100) / 100,
        tongTienHoiDong: rowsHoiDong.reduce((s, r) => s + r.thanhTien, 0)
    };

    return { rowsHuongDan, rowsPhanBien, rowsHoiDong, summary };
}

/**
 * Tạo Workbook XLSX (SheetJS) gồm 3 sheet chính + 1 sheet ẩn _META_DATA phục vụ ghép Tổng hợp liên ngành
 */
function buildPaymentWorkbook(sourceData, config, computed) {
    const wb = XLSX.utils.book_new();
    const khoaUpper = (config.khoa || sourceData.khoa || '').toUpperCase();
    const nganhUpper = (config.nganh || sourceData.nganh || '').toUpperCase();
    const khoaHocUpper = (config.khoaHoc || sourceData.khoaHoc || '').toUpperCase();
    const truongUpper = (config.truongThanhVien || 'TRƯỜNG CN&KT').toUpperCase();

    const makeFooterRows = (totalColsCount) => {
        const rowDate = new Array(totalColsCount).fill('');
        rowDate[Math.max(0, totalColsCount - 4)] = 'Đà Nẵng, ngày .... tháng .... năm ......';

        const rowSign1 = new Array(totalColsCount).fill('');
        const rowSign2 = new Array(totalColsCount).fill('');

        if (totalColsCount >= 12) {
            rowSign1[1] = 'LẬP BẢNG';
            rowSign1[4] = 'KT. TRƯỞNG KHOA';
            rowSign2[4] = 'PHÓ TRƯỞNG KHOA';
            rowSign1[6] = `BGH ${truongUpper}`;
            rowSign2[6] = 'PHÓ HIỆU TRƯỞNG';
            rowSign1[8] = 'BAN KH - TC';
            rowSign1[10] = 'BAN GIÁM ĐỐC';
        } else {
            rowSign1[1] = 'LẬP BẢNG';
            rowSign1[3] = 'KT. TRƯỞNG KHOA';
            rowSign2[3] = 'PHÓ TRƯỞNG KHOA';
            rowSign1[5] = `BGH ${truongUpper}`;
            rowSign2[5] = 'PHÓ HIỆU TRƯỞNG';
            rowSign1[7] = 'BAN KH - TC';
            rowSign1[9] = 'BAN GIÁM ĐỐC';
        }

        return [[], rowDate, rowSign1, rowSign2];
    };

    // SHEET 1: Thanh toán HD
    const ws1Data = [
        [{ t: 's', f: '"SET@DTU "&DATEVALUE(NOW())' }],
        ['ĐẠI HỌC DUY TÂN', '', '', '', 'BẢNG THANH TOÁN TIỀN HƯỚNG DẪN ĐATN'],
        [khoaUpper ? `KHOA ${khoaUpper.replace(/^KHOA\s+/i, '')}` : '', '', '', '', nganhUpper ? `NGÀNH ${nganhUpper.replace(/^NGÀNH\s+/i, '')}` : ''],
        ['', '', '', '', khoaHocUpper ? `KHÓA ${khoaHocUpper.replace(/^KHÓA\s+/i, '')}` : ''],
        [],
        ['TT', 'HỌ VÀ TÊN', '', 'Học vị', 'ĐƠN VỊ', 'SL ĐỀ TÀI', 'HỆ SỐ', 'QUY CHUẨN', 'ĐƠN GIÁ', 'SỐ TIỀN', 'KÝ TÊN', 'GHI CHÚ']
    ];

    const startRow1 = 7;
    computed.rowsHuongDan.forEach((r, idx) => {
        const excelRow = startRow1 + idx;
        ws1Data.push([
            r.tt,
            r.hoDem,
            r.ten,
            r.hocVi,
            r.donVi,
            r.slDeTai,
            r.heSo,
            { t: 'n', f: `F${excelRow}*G${excelRow}` },
            r.donGia,
            { t: 'n', f: `H${excelRow}*I${excelRow}` },
            '',
            ''
        ]);
    });

    const endRow1 = startRow1 + computed.rowsHuongDan.length - 1;
    ws1Data.push([
        '',
        'Tổng',
        '',
        '',
        '',
        computed.rowsHuongDan.length > 0 ? { t: 'n', f: `SUM(F${startRow1}:F${endRow1})` } : 0,
        '',
        '',
        '',
        computed.rowsHuongDan.length > 0 ? { t: 'n', f: `SUM(J${startRow1}:J${endRow1})` } : 0,
        '',
        ''
    ]);
    ws1Data.push(...makeFooterRows(12));

    const ws1 = XLSX.utils.aoa_to_sheet(ws1Data);
    ws1['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 11 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } },
        { s: { r: 1, c: 4 }, e: { r: 1, c: 11 } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: 3 } },
        { s: { r: 2, c: 4 }, e: { r: 2, c: 11 } },
        { s: { r: 3, c: 4 }, e: { r: 3, c: 11 } },
        { s: { r: 5, c: 1 }, e: { r: 5, c: 2 } }
    ];
    ws1['!cols'] = [
        { wch: 5 }, { wch: 22 }, { wch: 11 }, { wch: 9 }, { wch: 18 },
        { wch: 11 }, { wch: 9 }, { wch: 12 }, { wch: 13 }, { wch: 15 }, { wch: 12 }, { wch: 14 }
    ];
    XLSX.utils.book_append_sheet(wb, ws1, 'Thanh toán HD');

    // SHEET 2: Chấm PB
    const ws2Data = [
        [{ t: 's', f: '"SET@DTU "&DATEVALUE(NOW())' }],
        ['ĐẠI HỌC DUY TÂN', '', '', '', 'BẢNG THANH TOÁN TIỀN PHẢN BIỆN ĐATN'],
        [khoaUpper ? `KHOA ${khoaUpper.replace(/^KHOA\s+/i, '')}` : '', '', '', '', nganhUpper ? `NGÀNH ${nganhUpper.replace(/^NGÀNH\s+/i, '')}` : ''],
        ['', '', '', '', khoaHocUpper ? `KHÓA ${khoaHocUpper.replace(/^KHÓA\s+/i, '')}` : ''],
        [],
        ['TT', 'HỌ VÀ TÊN', '', 'Học vị', 'SL ĐỀ TÀI', 'HỆ SỐ', 'QUY CHUẨN', 'ĐƠN GIÁ', 'SỐ TIỀN', 'KÝ TÊN', 'GHI CHÚ']
    ];

    const startRow2 = 7;
    computed.rowsPhanBien.forEach((r, idx) => {
        const excelRow = startRow2 + idx;
        ws2Data.push([
            r.tt,
            r.hoDem,
            r.ten,
            r.hocVi,
            r.slDeTai,
            r.heSo,
            { t: 'n', f: `E${excelRow}*F${excelRow}` },
            r.donGia,
            { t: 'n', f: `G${excelRow}*H${excelRow}` },
            '',
            ''
        ]);
    });

    const endRow2 = startRow2 + computed.rowsPhanBien.length - 1;
    ws2Data.push([
        '',
        'TỔNG',
        '',
        '',
        computed.rowsPhanBien.length > 0 ? { t: 'n', f: `SUM(E${startRow2}:E${endRow2})` } : 0,
        '',
        '',
        '',
        computed.rowsPhanBien.length > 0 ? { t: 'n', f: `SUM(I${startRow2}:I${endRow2})` } : 0,
        '',
        ''
    ]);
    ws2Data.push(...makeFooterRows(11));

    const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
    ws2['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 10 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } },
        { s: { r: 1, c: 4 }, e: { r: 1, c: 10 } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: 3 } },
        { s: { r: 2, c: 4 }, e: { r: 2, c: 10 } },
        { s: { r: 3, c: 4 }, e: { r: 3, c: 10 } },
        { s: { r: 5, c: 1 }, e: { r: 5, c: 2 } }
    ];
    ws2['!cols'] = [
        { wch: 5 }, { wch: 22 }, { wch: 11 }, { wch: 9 },
        { wch: 11 }, { wch: 9 }, { wch: 12 }, { wch: 13 }, { wch: 15 }, { wch: 12 }, { wch: 14 }
    ];
    XLSX.utils.book_append_sheet(wb, ws2, 'Chấm PB');

    // SHEET 3: Chấm BV
    const ws3Data = [
        [{ t: 's', f: '"SET@DTU "&DATEVALUE(NOW())' }],
        ['ĐẠI HỌC DUY TÂN', '', '', '', 'BẢNG CHI TIẾT HỘI ĐỒNG CHẤM TỐT NGHIỆP'],
        [khoaUpper ? `KHOA ${khoaUpper.replace(/^KHOA\s+/i, '')}` : '', '', '', '', nganhUpper ? `NGÀNH ${nganhUpper.replace(/^NGÀNH\s+/i, '')}` : ''],
        ['', '', '', '', khoaHocUpper ? `KHÓA ${khoaHocUpper.replace(/^KHÓA\s+/i, '')}` : ''],
        [],
        ['STT', 'HỘI ĐỒNG SỐ', 'HỌ VÀ TÊN', '', 'CHỨC DANH', 'T.GIAN THAM GIA HĐ TỐT NGHIỆP', 'SLSV', 'HỆ SỐ', 'Giờ QĐ', 'ĐƠN GIÁ', 'THÀNH TIỀN', 'GHI CHÚ']
    ];

    const startRow3 = 7;
    const ws3Merges = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 11 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } },
        { s: { r: 1, c: 4 }, e: { r: 1, c: 11 } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: 3 } },
        { s: { r: 2, c: 4 }, e: { r: 2, c: 11 } },
        { s: { r: 3, c: 4 }, e: { r: 3, c: 11 } },
        { s: { r: 5, c: 2 }, e: { r: 5, c: 3 } }
    ];

    computed.rowsHoiDong.forEach((r, idx) => {
        const excelRow = startRow3 + idx;
        const zeroBasedRow = startRow3 - 1 + idx;
        ws3Data.push([
            r.stt,
            r.isFirstInBlock ? r.soHD : '',
            r.hoDem,
            r.ten,
            r.chucDanh,
            r.isFirstInBlock ? r.thoiGian : '',
            r.slsv,
            r.heSo,
            { t: 'n', f: `G${excelRow}*H${excelRow}` },
            r.donGia,
            { t: 'n', f: `I${excelRow}*J${excelRow}` },
            ''
        ]);

        if (r.isFirstInBlock) {
            ws3Merges.push({ s: { r: zeroBasedRow, c: 1 }, e: { r: zeroBasedRow + 2, c: 1 } });
            ws3Merges.push({ s: { r: zeroBasedRow, c: 5 }, e: { r: zeroBasedRow + 2, c: 5 } });
        }
    });

    const endRow3 = startRow3 + computed.rowsHoiDong.length - 1;
    ws3Data.push([
        '',
        'TỔNG',
        '',
        '',
        '',
        '',
        computed.summary.tongSLSVHoiDong,
        '',
        computed.rowsHoiDong.length > 0 ? { t: 'n', f: `SUM(I${startRow3}:I${endRow3})` } : 0,
        '',
        computed.rowsHoiDong.length > 0 ? { t: 'n', f: `SUM(K${startRow3}:K${endRow3})` } : 0,
        ''
    ]);
    ws3Data.push(...makeFooterRows(12));

    const ws3 = XLSX.utils.aoa_to_sheet(ws3Data);
    ws3['!merges'] = ws3Merges;
    ws3['!cols'] = [
        { wch: 5 }, { wch: 13 }, { wch: 22 }, { wch: 11 }, { wch: 14 },
        { wch: 18 }, { wch: 8 }, { wch: 9 }, { wch: 10 }, { wch: 13 }, { wch: 15 }, { wch: 14 }
    ];
    XLSX.utils.book_append_sheet(wb, ws3, 'Chấm BV');

    // SHEET 4 (ẨN): _META_DATA phục vụ ghép Tổng hợp liên ngành sau này
    const metaPayload = {
        version: 1,
        exportedAt: new Date().toISOString(),
        khoa: config.khoa || sourceData.khoa || '',
        nganh: config.nganh || sourceData.nganh || '',
        khoaHoc: config.khoaHoc || sourceData.khoaHoc || '',
        truongThanhVien: config.truongThanhVien || '',
        donGia: config.donGia,
        rowsHuongDan: computed.rowsHuongDan,
        rowsPhanBien: computed.rowsPhanBien,
        rowsHoiDong: computed.rowsHoiDong,
        summary: computed.summary
    };
    const wsMeta = XLSX.utils.aoa_to_sheet([
        ['DATN_PAYMENT_META_V1'],
        [JSON.stringify(metaPayload)]
    ]);
    XLSX.utils.book_append_sheet(wb, wsMeta, '_META_DATA');
    wb.Workbook = wb.Workbook || {};
    wb.Workbook.Sheets = [
        { name: 'Thanh toán HD', Hidden: 0 },
        { name: 'Chấm PB', Hidden: 0 },
        { name: 'Chấm BV', Hidden: 0 },
        { name: '_META_DATA', Hidden: 1 }
    ];

    return wb;
}

/**
 * Lưu bản sao file Excel (.xlsx) lên đúng thư mục chứa file JSON Phiên làm việc trên Drive (lưu trữ 60 ngày)
 */
async function uploadPaymentExcelToSessionFolder(wb, fileName, metadata) {
    const gasUrl =
        document.getElementById('cloud-gas-url')?.value.trim() ||
        window.GAS_API_URL ||
        window.CLOUD_API_URL ||
        localStorage.getItem('datn_cloud_gas_url') ||
        localStorage.getItem('datn_gas_api_url') ||
        '';

    if (!gasUrl) {
        return { saved: false, reason: 'Chưa cấu hình URL Google Apps Script' };
    }

    const sessionId =
        window.cloudSession?.sessionId ||
        document.getElementById('cloud-session-id')?.value.trim() ||
        '';
    const pin =
        window.cloudSession?.pin ||
        document.getElementById('cloud-session-pin')?.value.trim() ||
        '';

    const base64Data = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });

    const payload = {
        action: 'savePaymentExcelToSessionFolder',
        sessionId,
        pin,
        fileName,
        base64Data,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        retentionDays: 60,
        meta: metadata
    };

    const res = await fetch(gasUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload)
    });
    const json = await res.json();
    return {
        saved: Boolean(json && json.ok),
        fileUrl: json?.fileUrl || '',
        message: json?.message || ''
    };
}

function ensurePaymentPreviewModal() {
    let modalEl = document.getElementById('modalPreviewThanhToan');
    if (modalEl) return modalEl;

    const html = `
        <div class="modal fade" id="modalPreviewThanhToan" tabindex="-1" aria-hidden="true">
            <div class="modal-dialog modal-xl modal-dialog-scrollable">
                <div class="modal-content">
                    <div class="modal-header bg-success text-white py-2">
                        <h6 class="modal-title fw-bold mb-0">💰 XEM TRƯỚC BẢNG THANH TOÁN ĐỒ ÁN TỐT NGHIỆP (3 SHEET)</h6>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Đóng"></button>
                    </div>
                    <div class="modal-body p-3">
                        <div class="card border-success mb-3">
                            <div class="card-header bg-light py-1 px-3 fw-bold small text-success">
                                ⚙️ Thông tin Header & Định mức thanh toán (Tự động cập nhật bảng xem trước bên dưới)
                            </div>
                            <div class="card-body p-2">
                                <div class="row g-2 small">
                                    <div class="col-12 col-md-4">
                                        <label class="form-label mb-0 fw-semibold">Ngành đào tạo:</label>
                                        <input type="text" id="pay-cfg-nganh" class="form-control form-control-sm pay-cfg-input">
                                    </div>
                                    <div class="col-6 col-md-4">
                                        <label class="form-label mb-0 fw-semibold">Khóa (từ Lớp của SV):</label>
                                        <input type="text" id="pay-cfg-khoahoc" class="form-control form-control-sm pay-cfg-input">
                                    </div>
                                    <div class="col-6 col-md-4">
                                        <label class="form-label mb-0 fw-semibold">Trường thành viên (BGH):</label>
                                        <input type="text" id="pay-cfg-truong" class="form-control form-control-sm pay-cfg-input">
                                    </div>
                                    <div class="col-6 col-md-2">
                                        <label class="form-label mb-0 fw-semibold">Đơn giá (VNĐ):</label>
                                        <input type="number" id="pay-cfg-dongia" class="form-control form-control-sm pay-cfg-input" step="1000">
                                    </div>
                                    <div class="col-6 col-md-2">
                                        <label class="form-label mb-0 fw-semibold">Hệ số HD:</label>
                                        <input type="number" id="pay-cfg-hs-hd" class="form-control form-control-sm pay-cfg-input" step="0.1">
                                    </div>
                                    <div class="col-4 col-md-2">
                                        <label class="form-label mb-0 fw-semibold">Hệ số PB:</label>
                                        <input type="number" id="pay-cfg-hs-pb" class="form-control form-control-sm pay-cfg-input" step="0.1">
                                    </div>
                                    <div class="col-4 col-md-2">
                                        <label class="form-label mb-0 fw-semibold">HĐ - Chủ tịch:</label>
                                        <input type="number" id="pay-cfg-hs-ct" class="form-control form-control-sm pay-cfg-input" step="0.01">
                                    </div>
                                    <div class="col-4 col-md-2">
                                        <label class="form-label mb-0 fw-semibold">HĐ - Phản biện:</label>
                                        <input type="number" id="pay-cfg-hs-hdpb" class="form-control form-control-sm pay-cfg-input" step="0.01">
                                    </div>
                                    <div class="col-4 col-md-2">
                                        <label class="form-label mb-0 fw-semibold">HĐ - Thư ký:</label>
                                        <input type="number" id="pay-cfg-hs-hdtk" class="form-control form-control-sm pay-cfg-input" step="0.01">
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div id="pay-preview-summary-bar" class="alert alert-info py-2 px-3 small mb-2 d-flex flex-wrap justify-content-between gap-2"></div>

                        <ul class="nav nav-tabs small fw-bold" id="payPreviewTabs" role="tablist">
                            <li class="nav-item" role="presentation">
                                <button class="nav-link active" data-bs-toggle="tab" data-bs-target="#tab-pay-hd" type="button">1. Thanh toán HD</button>
                            </li>
                            <li class="nav-item" role="presentation">
                                <button class="nav-link" data-bs-toggle="tab" data-bs-target="#tab-pay-pb" type="button">2. Chấm PB</button>
                            </li>
                            <li class="nav-item" role="presentation">
                                <button class="nav-link" data-bs-toggle="tab" data-bs-target="#tab-pay-bv" type="button">3. Chấm BV (Hội đồng)</button>
                            </li>
                        </ul>

                        <div class="tab-content border border-top-0 p-2 bg-white" style="max-height: 46vh; overflow: auto;">
                            <div class="tab-pane fade show active" id="tab-pay-hd"></div>
                            <div class="tab-pane fade" id="tab-pay-pb"></div>
                            <div class="tab-pane fade" id="tab-pay-bv"></div>
                        </div>
                    </div>
                    <div class="modal-footer py-2 d-flex justify-content-between flex-wrap gap-2">
                        <div class="small text-muted">
                            📌 Mỗi sheet tự động gắn mã <code>SET@DTU</code> ở dòng 1 &amp; lưu bản sao vào thư mục Phiên (60 ngày).
                        </div>
                        <div class="d-flex gap-2">
                            <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Đóng</button>
                            <button type="button" class="btn btn-success btn-sm fw-bold" id="btn-confirm-export-payment">
                                📥 Xuất Excel &amp; Lưu Phiên
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    document.body.insertAdjacentHTML('beforeend', html);
    return document.getElementById('modalPreviewThanhToan');
}

/**
 * Mở Modal xem trước Bảng thanh toán 3 sheet trước khi xuất & lưu vào thư mục Phiên (60 ngày)
 */
export function openPaymentPreviewModal(getMetaFn) {
    const meta = typeof getMetaFn === 'function' ? getMetaFn() : {};
    const sourceData = collectPaymentSourceFromDOM(meta);

    if (sourceData.assignedStudents.length === 0) {
        alert('⚠️ Chưa có sinh viên nào được xếp vào Hội đồng để lập bảng thanh toán!');
        return;
    }

    const modalEl = ensurePaymentPreviewModal();
    const savedCfg = loadSavedPaymentConfig();

    document.getElementById('pay-cfg-nganh').value = sourceData.nganh || '';
    document.getElementById('pay-cfg-khoahoc').value = sourceData.khoaHoc || '';
    document.getElementById('pay-cfg-truong').value = savedCfg.truongThanhVien || 'TRƯỜNG CN&KT';
    document.getElementById('pay-cfg-dongia').value = savedCfg.donGia;
    document.getElementById('pay-cfg-hs-hd').value = savedCfg.heSoHD;
    document.getElementById('pay-cfg-hs-pb').value = savedCfg.heSoPB;
    document.getElementById('pay-cfg-hs-ct').value = savedCfg.heSoCT;
    document.getElementById('pay-cfg-hs-hdpb').value = savedCfg.heSoHDPB;
    document.getElementById('pay-cfg-hs-hdtk').value = savedCfg.heSoHDTK;

    const readCurrentConfig = () => ({
        khoa: sourceData.khoa,
        nganh: document.getElementById('pay-cfg-nganh').value.trim(),
        khoaHoc: document.getElementById('pay-cfg-khoahoc').value.trim(),
        truongThanhVien: document.getElementById('pay-cfg-truong').value.trim() || 'TRƯỜNG CN&KT',
        donGia: Number(document.getElementById('pay-cfg-dongia').value) || 0,
        heSoHD: Number(document.getElementById('pay-cfg-hs-hd').value) || 0,
        heSoPB: Number(document.getElementById('pay-cfg-hs-pb').value) || 0,
        heSoCT: Number(document.getElementById('pay-cfg-hs-ct').value) || 0,
        heSoHDPB: Number(document.getElementById('pay-cfg-hs-hdpb').value) || 0,
        heSoHDTK: Number(document.getElementById('pay-cfg-hs-hdtk').value) || 0
    });

    const renderPreviewContent = () => {
        const cfg = readCurrentConfig();
        const computed = computePaymentTables(sourceData, cfg);
        const s = computed.summary;
        const tongTatCa = s.tongTienHuongDan + s.tongTienPhanBien + s.tongTienHoiDong;

        const headerBanner = `
            <div class="border rounded p-2 mb-2 bg-light small">
                <div class="text-muted font-monospace" style="font-size: 0.72rem;">Dòng 1 Excel: <strong>="SET@DTU "&amp;DATEVALUE(NOW())</strong></div>
                <div class="d-flex justify-content-between flex-wrap mt-1">
                    <div><strong>ĐẠI HỌC DUY TÂN</strong> — KHOA ${escapeHtml((cfg.khoa || '...').toUpperCase())}</div>
                    <div class="text-end">
                        <div><strong>NGÀNH ${escapeHtml((cfg.nganh || '...').toUpperCase())}</strong></div>
                        <div><strong>KHÓA ${escapeHtml((cfg.khoaHoc || '...').toUpperCase())}</strong></div>
                    </div>
                </div>
            </div>
        `;

        document.getElementById('pay-preview-summary-bar').innerHTML = `
            <span>👨‍🏫 Hướng dẫn (${s.tongSLHuongDan} SV): <strong>${formatVND(s.tongTienHuongDan)} đ</strong></span>
            <span>📝 Phản biện (${s.tongSLPhanBien} SV): <strong>${formatVND(s.tongTienPhanBien)} đ</strong></span>
            <span>🏛️ Hội đồng (${s.tongSLSVHoiDong} SV / ${formatDecimalVN(s.tongGioQDHoiDong)} giờ): <strong>${formatVND(s.tongTienHoiDong)} đ</strong></span>
            <span class="text-primary">💰 TỔNG CỘNG: <strong>${formatVND(tongTatCa)} đ</strong></span>
        `;

        document.getElementById('tab-pay-hd').innerHTML = `
            ${headerBanner}
            <table class="table table-bordered table-sm text-center align-middle small mb-0">
                <thead class="table-secondary">
                    <tr>
                        <th>TT</th><th colspan="2">HỌ VÀ TÊN</th><th>Học vị</th><th>ĐƠN VỊ</th>
                        <th>SL ĐỀ TÀI</th><th>HỆ SỐ</th><th>QUY CHUẨN</th><th>ĐƠN GIÁ</th><th>SỐ TIỀN</th>
                    </tr>
                </thead>
                <tbody>
                    ${computed.rowsHuongDan.map(r => `
                        <tr>
                            <td>${r.tt}</td>
                            <td class="text-start">${escapeHtml(r.hoDem)}</td>
                            <td class="text-start fw-bold">${escapeHtml(r.ten)}</td>
                            <td>${escapeHtml(r.hocVi)}</td>
                            <td class="text-start">${escapeHtml(r.donVi)}</td>
                            <td>${r.slDeTai}</td>
                            <td>${formatDecimalVN(r.heSo)}</td>
                            <td>${formatDecimalVN(r.quyChuan)}</td>
                            <td class="text-end">${formatVND(r.donGia)}</td>
                            <td class="text-end fw-bold">${formatVND(r.soTien)}</td>
                        </tr>
                    `).join('')}
                </tbody>
                <tfoot class="table-light fw-bold">
                    <tr>
                        <td colspan="5">Tổng</td>
                        <td>${s.tongSLHuongDan}</td>
                        <td colspan="3"></td>
                        <td class="text-end text-primary">${formatVND(s.tongTienHuongDan)}</td>
                    </tr>
                </tfoot>
            </table>
        `;

        document.getElementById('tab-pay-pb').innerHTML = `
            ${headerBanner}
            <table class="table table-bordered table-sm text-center align-middle small mb-0">
                <thead class="table-secondary">
                    <tr>
                        <th>TT</th><th colspan="2">HỌ VÀ TÊN</th><th>Học vị</th>
                        <th>SL ĐỀ TÀI</th><th>HỆ SỐ</th><th>QUY CHUẨN</th><th>ĐƠN GIÁ</th><th>SỐ TIỀN</th>
                    </tr>
                </thead>
                <tbody>
                    ${computed.rowsPhanBien.map(r => `
                        <tr>
                            <td>${r.tt}</td>
                            <td class="text-start">${escapeHtml(r.hoDem)}</td>
                            <td class="text-start fw-bold">${escapeHtml(r.ten)}</td>
                            <td>${escapeHtml(r.hocVi)}</td>
                            <td>${r.slDeTai}</td>
                            <td>${formatDecimalVN(r.heSo)}</td>
                            <td>${formatDecimalVN(r.quyChuan)}</td>
                            <td class="text-end">${formatVND(r.donGia)}</td>
                            <td class="text-end fw-bold">${formatVND(r.soTien)}</td>
                        </tr>
                    `).join('')}
                </tbody>
                <tfoot class="table-light fw-bold">
                    <tr>
                        <td colspan="4">TỔNG</td>
                        <td>${s.tongSLPhanBien}</td>
                        <td colspan="3"></td>
                        <td class="text-end text-primary">${formatVND(s.tongTienPhanBien)}</td>
                    </tr>
                </tfoot>
            </table>
        `;

        document.getElementById('tab-pay-bv').innerHTML = `
            ${headerBanner}
            <table class="table table-bordered table-sm text-center align-middle small mb-0">
                <thead class="table-secondary">
                    <tr>
                        <th>STT</th><th>HỘI ĐỒNG SỐ</th><th colspan="2">HỌ VÀ TÊN</th><th>CHỨC DANH</th>
                        <th>T.GIAN THAM GIA HĐ</th><th>SLSV</th><th>HỆ SỐ</th><th>Giờ QĐ</th><th>ĐƠN GIÁ</th><th>THÀNH TIỀN</th>
                    </tr>
                </thead>
                <tbody>
                    ${computed.rowsHoiDong.map(r => `
                        <tr>
                            <td>${r.stt}</td>
                            ${r.isFirstInBlock ? `<td rowspan="3" class="fw-bold bg-light">${escapeHtml(r.soHD)}</td>` : ''}
                            <td class="text-start">${escapeHtml(r.hoDem)}</td>
                            <td class="text-start fw-bold">${escapeHtml(r.ten)}</td>
                            <td>${escapeHtml(r.chucDanh)}</td>
                            ${r.isFirstInBlock ? `<td rowspan="3" class="bg-light">${escapeHtml(r.thoiGian)}</td>` : ''}
                            <td>${r.slsv}</td>
                            <td>${formatDecimalVN(r.heSo)}</td>
                            <td>${formatDecimalVN(r.gioQD)}</td>
                            <td class="text-end">${formatVND(r.donGia)}</td>
                            <td class="text-end fw-bold">${formatVND(r.thanhTien)}</td>
                        </tr>
                    `).join('')}
                </tbody>
                <tfoot class="table-light fw-bold">
                    <tr>
                        <td colspan="6">TỔNG</td>
                        <td>${s.tongSLSVHoiDong}</td>
                        <td></td>
                        <td>${formatDecimalVN(s.tongGioQDHoiDong)}</td>
                        <td></td>
                        <td class="text-end text-primary">${formatVND(s.tongTienHoiDong)}</td>
                    </tr>
                </tfoot>
            </table>
        `;
    };

    modalEl.querySelectorAll('.pay-cfg-input').forEach(inp => {
        inp.oninput = renderPreviewContent;
    });

    const btnConfirm = document.getElementById('btn-confirm-export-payment');
    btnConfirm.onclick = async () => {
        const cfg = readCurrentConfig();
        savePaymentConfig(cfg);
        const computed = computePaymentTables(sourceData, cfg);
        const wb = buildPaymentWorkbook(sourceData, cfg, computed);

        const nganhSlug = sanitizeFileSlug(cfg.nganh || 'Nganh');
        const khoaSlug = sanitizeFileSlug(cfg.khoaHoc || 'DATN');
        const fileName = `Thanh_Toan_DATN_${nganhSlug}_${khoaSlug}.xlsx`;

        // 1. Tải file Excel về máy người dùng
        XLSX.writeFile(wb, fileName);

        // 2. Tự động lưu bản sao vào thư mục Phiên làm việc trên Drive (lưu trữ 60 ngày)
        const origText = btnConfirm.innerHTML;
        btnConfirm.disabled = true;
        btnConfirm.innerHTML = '⏳ Đang lưu bản sao lên Phiên Cloud...';
        try {
            const cloudRes = await uploadPaymentExcelToSessionFolder(wb, fileName, {
                khoa: cfg.khoa,
                nganh: cfg.nganh,
                khoaHoc: cfg.khoaHoc,
                tongTien: computed.summary.tongTienHuongDan + computed.summary.tongTienPhanBien + computed.summary.tongTienHoiDong
            });
            if (cloudRes.saved) {
                alert(`✅ Đã xuất file "${fileName}" và lưu tự động vào thư mục Phiên làm việc (thời hạn lưu trữ 60 ngày)!`);
            } else {
                alert(`✅ Đã tải file "${fileName}" về máy.\nℹ️ Lưu ý Cloud: ${cloudRes.reason || cloudRes.message || 'Chưa kết nối phiên Cloud nên chỉ lưu bản trên máy.'}`);
            }
        } catch (err) {
            console.warn('Lỗi lưu bản sao Excel lên thư mục Phiên:', err);
            alert(`✅ Đã tải file "${fileName}" về máy.\n⚠️ Không thể đồng bộ bản sao lên Drive: ${err.message}`);
        } finally {
            btnConfirm.disabled = false;
            btnConfirm.innerHTML = origText;
            bootstrap.Modal.getInstance(modalEl)?.hide();
        }
    };

    renderPreviewContent();
    bootstrap.Modal.getOrCreateInstance(modalEl).show();
}
