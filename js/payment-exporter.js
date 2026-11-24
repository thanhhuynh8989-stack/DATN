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
    return val.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Chuẩn hóa danh sách nhiều ngành: tách theo dòng, loại trùng lặp, đảm bảo mỗi ngành 1 dòng
function normalizeMultiNganhLines(rawNganh) {
    const rawStr = String(rawNganh ?? '').trim();
    if (!rawStr) return [];
    const seen = new Set();
    const lines = [];
    rawStr
        .split(/\r?\n/)
        .map(s => s.trim())
        .filter(Boolean)
        .forEach(item => {
            const key = item.toLowerCase();
            if (!seen.has(key)) {
                seen.add(key);
                lines.push(item);
            }
        });
    return lines;
}

/**
- Tính giá trị số serial ngày theo chuẩn Excel (tương đương giá trị của DATEVALUE(TODAY()))
- Trả về chuỗi giá trị tĩnh: "SET@DTU <serial>"
 */
function getSetDtuSerialStamp() {
    const now = new Date();
    const localMidnightUtc = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
    const excelEpochUtc = Date.UTC(1899, 11, 30);
    const excelSerial = Math.floor((localMidnightUtc - excelEpochUtc) / 86400000);
    return `SET@DTU ${excelSerial}`;
}

async function ensureExcelJSLoaded() {
    if (window.ExcelJS) return window.ExcelJS;
    await new Promise((resolve, reject) => {
        const existing = document.querySelector('script[data-exceljs-cdn="1"]');
        if (existing) {
            existing.addEventListener('load', resolve);
            existing.addEventListener('error', reject);
            return;
        }
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
        script.setAttribute('data-exceljs-cdn', '1');
        script.onload = resolve;
        script.onerror = () => reject(new Error('Không thể tải thư viện định dạng ExcelJS từ CDN.'));
        document.head.appendChild(script);
    });
    return window.ExcelJS;
}

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

    const hdNumMatch = cleaned.match(/(?:hội\s*đồng|hoi\s*dong|hđ|hd)\s*(?:số|so|#|:|-|\.)?\s*([0-9]+[a-zA-Z0-9\-]*)/i);
    if (hdNumMatch && hdNumMatch[1]) {
        return hdNumMatch[1].trim();
    }

    const stripped = cleaned
        .replace(/^(?:hội\s*đồng|hoi\s*dong|hđ|hd)\s*(?:số|so|#|:|-|\.)?\s*/i, '')
        .trim();
    if (stripped) return stripped;

    const anyNum = cleaned.match(/\d+/);
    return anyNum ? anyNum[0] : String(fallbackIdx);
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
        const finalDonVi = fallbackDonVi || fromPool?.donVi || (meta.khoa ? `Khoa ${meta.khoa.replace(/^Khoa\s+/i, '')}` : 'Đại học Duy Tân');
        return { maGV: finalId, hoTen: finalName, donVi: finalDonVi };
    };

    const classSet = new Set();
    const nganhList = [];
    const nganhLowerSet = new Set();
    const hdGroupedMap = new Map();
    const assignedStudents = [];

    document.querySelectorAll('.box-hoidong').forEach((box, idx) => {
        const svCards = Array.from(box.querySelectorAll('.dropzone-sv .card-item'));
        if (svCards.length === 0) return;

        const titleEl = box.querySelector('.hd-title, .hd-name, .hoidong-title, input[type="text"]');
        const rawTitle = (
            titleEl?.value ??
            titleEl?.textContent ??
            box.getAttribute('data-hd-so') ??
            box.getAttribute('data-hd-name') ??
            box.querySelector('.card-header, .box-header, h5, h6')?.textContent ??
            ''
        ).trim();

        const soHD = extractHoiDongNumber(rawTitle, idx + 1) || String(idx + 1);
        const ngayVN = formatDateVN(box.querySelector('.hd-date')?.value || box.getAttribute('data-date') || '');

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

        // Tạo khóa định danh duy nhất theo bộ 3 thành viên + chức danh
        const ctKey = normalizeTeacherKeyForDupCheck(ct.hoTen) || ct.maGV.toUpperCase();
        const pbKey = normalizeTeacherKeyForDupCheck(pb.hoTen) || pb.maGV.toUpperCase();
        const tkKey = normalizeTeacherKeyForDupCheck(tk.hoTen) || tk.maGV.toUpperCase();
        const hasAnyMember = Boolean(ctKey || pbKey || tkKey);
        const groupKey = hasAnyMember ? `MEM_${ctKey}|${pbKey}|${tkKey}` : `HD_${soHD}_${idx}`;

        if (!hdGroupedMap.has(groupKey)) {
            hdGroupedMap.set(groupKey, {
                soHD,
                ct,
                pb,
                tk,
                dates: new Set(),
                totalSV: 0
            });
        }
        const hdRec = hdGroupedMap.get(groupKey);

        // Nếu gộp nhiều box có số HĐ khác nhau (vd: 12 và 13), ưu tiên giữ số HĐ nhỏ hơn (12)
        const currNum = parseInt(soHD, 10);
        const existingNum = parseInt(hdRec.soHD, 10);
        if (!isNaN(currNum) && !isNaN(existingNum) && currNum < existingNum) {
            hdRec.soHD = soHD;
        }

        if (!hdRec.ct.hoTen && ct.hoTen) hdRec.ct = ct;
        if (!hdRec.pb.hoTen && pb.hoTen) hdRec.pb = pb;
        if (!hdRec.tk.hoTen && tk.hoTen) hdRec.tk = tk;
        if (ngayVN) hdRec.dates.add(ngayVN);
        hdRec.totalSV += svCards.length;

        svCards.forEach(svEl => {
            const lop = (svEl.getAttribute('data-lop') || '').trim();
            const nganhSV = (svEl.getAttribute('data-nganh') || '').trim();
            if (lop) classSet.add(lop);
            if (nganhSV && !nganhLowerSet.has(nganhSV.toLowerCase())) {
                nganhLowerSet.add(nganhSV.toLowerCase());
                nganhList.push(nganhSV);
            }

            assignedStudents.push({
                maSV: (svEl.getAttribute('data-id') || '').trim(),
                hoTenSV: (svEl.getAttribute('data-name') || '').trim(),
                lop,
                maGVHD: (svEl.getAttribute('data-magvhd') || '').trim(),
                gvhd: (svEl.getAttribute('data-gvhd') || '').trim(),
                soHD: hdRec.soHD,
                pb
            });
        });
    });

    // Sắp xếp lại thứ tự ngày tăng dần bên trong từng Hội đồng đã gộp (vd: "23/12/2025, 24/12/2025")
    const parseVNDate = (dStr) => {
        const p = String(dStr || '').split('/');
        if (p.length !== 3) return 0;
        return new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0])).getTime();
    };

    const hdGroupedList = Array.from(hdGroupedMap.values())
        .map(hd => {
            const sortedDates = Array.from(hd.dates).sort((a, b) => parseVNDate(a) - parseVNDate(b));
            return {
                ...hd,
                dates: new Set(sortedDates)
            };
        })
        .sort((a, b) => {
            const na = parseInt(a.soHD, 10);
            const nb = parseInt(b.soHD, 10);
            if (!isNaN(na) && !isNaN(nb)) return na - nb;
            return String(a.soHD).localeCompare(String(b.soHD), 'vi');
        });

    const metaNganhLines = normalizeMultiNganhLines(meta.nganh);
    const detectedNganh = metaNganhLines.length > 0 ? metaNganhLines.join('\n') : nganhList.join('\n');
    const detectedKhoaHoc = Array.from(classSet).sort((a, b) => a.localeCompare(b, 'vi')).join(', ');

    return {
        khoa: (meta.khoa || '').trim(),
        nganh: detectedNganh,
        khoaHoc: detectedKhoaHoc,
        assignedStudents,
        hdGroupedList,
        resolveTeacherMeta
    };
}

export function computePaymentTables(sourceData, config) {
    const donGia = Number(config.donGia) || 0;
    const heSoHD = Number(config.heSoHD) || 0;
    const heSoPB = Number(config.heSoPB) || 0;
    const heSoCT = Number(config.heSoCT) || 0;
    const heSoHDPB = Number(config.heSoHDPB) || 0;
    const heSoHDTK = Number(config.heSoHDTK) || 0;

    // 1. Bảng Hướng dẫn
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

    // 2. Bảng Phản biện
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

    // 3. Bảng Chi tiết Hội đồng chấm Tốt nghiệp
    const rowsHoiDong = [];
    let sttCounter = 1;
    sourceData.hdGroupedList.forEach((hd, hdIdx) => {
        const thoiGianStr = Array.from(hd.dates).join(', ');
        const soHDStr = String(hd.soHD || (hdIdx + 1));
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
                soHD: soHDStr,
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

async function buildStyledPaymentWorkbook(sourceData, config, computed) {
    const ExcelJS = await ensureExcelJSLoaded();
    const wb = new ExcelJS.Workbook();
    wb.creator = 'SET@DTU - Hệ thống Quản lý ĐATN';
    wb.created = new Date();

    const stampText = getSetDtuSerialStamp();
    const khoaUpper = (config.khoa || sourceData.khoa || '').toUpperCase().replace(/^KHOA\s+/i, '');
    const nganhLinesUpper = normalizeMultiNganhLines(config.nganh || sourceData.nganh || '')
        .map(s => s.toUpperCase().replace(/^NGÀNH\s+/i, ''));
    const khoaHocUpper = (config.khoaHoc || sourceData.khoaHoc || '').toUpperCase().replace(/^KHÓA\s+/i, '');
    const truongUpper = (config.truongThanhVien || 'TRƯỜNG CN&KT').toUpperCase();

    const FONT_BASE = { name: 'Times New Roman', size: 12 };
    const FONT_BOLD = { name: 'Times New Roman', size: 12, bold: true };
    const FONT_ITALIC = { name: 'Times New Roman', size: 12, italic: true };
    const FONT_STAMP = { name: 'Times New Roman', size: 11, italic: true, color: { argb: 'FF444444' } };

    const THIN_BORDER = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
    };

    const NAME_LEFT_BORDER = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } }
    };
    const NAME_RIGHT_BORDER = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } }
    };

    const HEADER_FILL = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9D9D9' }
    };

    const setupHeaderAndFooter = (ws, totalCols, titleText) => {
        ws.pageSetup = {
            paperSize: 9,
            orientation: 'landscape',
            fitToPage: true,
            fitToWidth: 1,
            fitToHeight: 0,
            margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 }
        };

        ws.mergeCells(1, 1, 1, totalCols);
        const r1 = ws.getRow(1);
        r1.height = 19;
        const cStamp = ws.getCell(1, 1);
        cStamp.value = stampText;
        cStamp.font = FONT_STAMP;
        cStamp.alignment = { horizontal: 'right', vertical: 'middle' };

        ws.mergeCells(2, 1, 2, 4);
        ws.mergeCells(2, 5, 2, totalCols);
        const r2 = ws.getRow(2);
        r2.height = 22;
        ws.getCell(2, 1).value = 'ĐẠI HỌC DUY TÂN';
        ws.getCell(2, 1).font = FONT_BOLD;
        ws.getCell(2, 1).alignment = { horizontal: 'center', vertical: 'middle' };

        ws.getCell(2, 5).value = titleText;
        ws.getCell(2, 5).font = FONT_BOLD;
        ws.getCell(2, 5).alignment = { horizontal: 'center', vertical: 'middle' };

        ws.mergeCells(3, 1, 3, 4);
        ws.mergeCells(3, 5, 3, totalCols);
        const r3 = ws.getRow(3);

        // Tự động liệt kê từng ngành xuống dòng và căn chỉnh chiều cao dòng 3 (dòng NGÀNH)
        let nganhCellText = '';
        let visualLines = 1;
        if (nganhLinesUpper.length === 1) {
            nganhCellText = `NGÀNH ${nganhLinesUpper[0]}`;
            visualLines = Math.max(1, Math.ceil(nganhCellText.length / 65));
        } else if (nganhLinesUpper.length > 1) {
            nganhCellText = `NGÀNH:\n${nganhLinesUpper.join('\n')}`;
            visualLines = 1 + nganhLinesUpper.reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / 65)), 0);
        }
        r3.height = Math.max(21, visualLines * 18 + 4);

        const cKhoa = ws.getCell(3, 1);
        cKhoa.value = khoaUpper ? `KHOA ${khoaUpper}` : '';
        cKhoa.font = { ...FONT_BOLD, underline: true };
        cKhoa.alignment = { horizontal: 'center', vertical: 'middle' };

        const cNganh = ws.getCell(3, 5);
        cNganh.value = nganhCellText;
        cNganh.font = FONT_BOLD;
        cNganh.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

        ws.mergeCells(4, 5, 4, totalCols);
        const r4 = ws.getRow(4);
        r4.height = 21;
        const cKhoaHoc = ws.getCell(4, 5);
        cKhoaHoc.value = khoaHocUpper ? `KHÓA ${khoaHocUpper}` : '';
        cKhoaHoc.font = FONT_BOLD;
        cKhoaHoc.alignment = { horizontal: 'center', vertical: 'middle' };

        ws.getRow(5).height = 10;
    };

    const appendSignatureFooter = (ws, afterRowIdx, totalCols) => {
        ws.getRow(afterRowIdx + 1).height = 12;

        const dateRowIdx = afterRowIdx + 2;
        const sign1RowIdx = afterRowIdx + 3;
        const sign2RowIdx = afterRowIdx + 4;

        ws.getRow(dateRowIdx).height = 20;
        ws.getRow(sign1RowIdx).height = 20;
        ws.getRow(sign2RowIdx).height = 20;

        ws.mergeCells(dateRowIdx, totalCols - 3, dateRowIdx, totalCols);
        const cDate = ws.getCell(dateRowIdx, totalCols - 3);
        cDate.value = 'Đà Nẵng, ngày .... tháng .... năm ......';
        cDate.font = FONT_ITALIC;
        cDate.alignment = { horizontal: 'right', vertical: 'middle' };

        const signBlocks = [
            { c1: 1, c2: 3, line1: 'LẬP BẢNG', line2: '' },
            { c1: 4, c2: 5, line1: 'TRƯỞNG KHOA', line2: '' },
            { c1: 6, c2: 8, line1: `BGH ${truongUpper}`, line2: '' },
            { c1: 9, c2: 10, line1: 'BAN KH - TC', line2: '' },
            { c1: 11, c2: 12, line1: 'BAN GIÁM ĐỐC', line2: '' }
        ];

        signBlocks.forEach(b => {
            ws.mergeCells(sign1RowIdx, b.c1, sign1RowIdx, b.c2);
            const cell1 = ws.getCell(sign1RowIdx, b.c1);
            cell1.value = b.line1;
            cell1.font = FONT_BOLD;
            cell1.alignment = { horizontal: 'center', vertical: 'middle' };

            if (b.line2) {
                ws.mergeCells(sign2RowIdx, b.c1, sign2RowIdx, b.c2);
                const cell2 = ws.getCell(sign2RowIdx, b.c1);
                cell2.value = b.line2;
                cell2.font = FONT_BOLD;
                cell2.alignment = { horizontal: 'center', vertical: 'middle' };
            }
        });
    };

    // =========================================================================
    // SHEET 1: Thanh toán HD
    // =========================================================================
    const ws1 = wb.addWorksheet('Thanh toán HD');
    setupHeaderAndFooter(ws1, 12, 'BẢNG THANH TOÁN TIỀN HƯỚNG DẪN ĐATN');

    ws1.columns = [
        { width: 5 }, { width: 22 }, { width: 11 }, { width: 9 }, { width: 19 },
        { width: 11 }, { width: 9 }, { width: 12 }, { width: 13 }, { width: 16 }, { width: 13 }, { width: 14 }
    ];

    const hdrRow1 = ws1.getRow(6);
    hdrRow1.height = 30;
    const headers1 = ['TT', 'HỌ VÀ TÊN', '', 'Học vị', 'ĐƠN VỊ', 'SL ĐỀ TÀI', 'HỆ SỐ', 'QUY CHUẨN', 'ĐƠN GIÁ', 'SỐ TIỀN', 'KÝ TÊN', 'GHI CHÚ'];
    headers1.forEach((h, i) => {
        const cell = hdrRow1.getCell(i + 1);
        cell.value = h;
        cell.font = FONT_BOLD;
        cell.fill = HEADER_FILL;
        cell.border = THIN_BORDER;
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    ws1.mergeCells(6, 2, 6, 3);

    const startRow1 = 7;
    computed.rowsHuongDan.forEach((r, idx) => {
        const rowNum = startRow1 + idx;
        const row = ws1.getRow(rowNum);
        row.height = 22;

        const vals = [
            r.tt,
            r.hoDem,
            r.ten,
            r.hocVi,
            r.donVi,
            r.slDeTai,
            r.heSo,
            { formula: `F${rowNum}*G${rowNum}`, result: r.quyChuan },
            r.donGia,
            { formula: `H${rowNum}*I${rowNum}`, result: r.soTien },
            '',
            ''
        ];

        vals.forEach((v, cIdx) => {
            const colNum = cIdx + 1;
            const cell = row.getCell(colNum);
            cell.value = v;
            cell.font = FONT_BASE;

            if (colNum === 2) {
                cell.border = NAME_LEFT_BORDER;
                cell.alignment = { horizontal: 'left', vertical: 'middle' };
            } else if (colNum === 3) {
                cell.border = NAME_RIGHT_BORDER;
                cell.alignment = { horizontal: 'left', vertical: 'middle' };
            } else if (colNum === 5) {
                cell.border = THIN_BORDER;
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            } else if (colNum === 9 || colNum === 10) {
                cell.border = THIN_BORDER;
                cell.numFmt = '#,##0';
                cell.alignment = { horizontal: 'right', vertical: 'middle' };
            } else if (colNum === 7 || colNum === 8) {
                cell.border = THIN_BORDER;
                cell.numFmt = '0.00'; // Luôn giữ đủ 2 chữ số thập phân (tránh lỗi 8, hoặc 2,1)
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            } else {
                cell.border = THIN_BORDER;
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            }
        });
    });

    const totalRowIdx1 = startRow1 + computed.rowsHuongDan.length;
    const totalRow1 = ws1.getRow(totalRowIdx1);
    totalRow1.height = 24;
    for (let c = 1; c <= 12; c++) {
        const cell = totalRow1.getCell(c);
        cell.font = FONT_BOLD;
        cell.border = THIN_BORDER;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    }
    ws1.mergeCells(totalRowIdx1, 2, totalRowIdx1, 3);
    totalRow1.getCell(2).value = 'Tổng';
    totalRow1.getCell(6).value = computed.rowsHuongDan.length > 0
        ? { formula: `SUM(F${startRow1}:F${totalRowIdx1 - 1})`, result: computed.summary.tongSLHuongDan }
        : 0;
    totalRow1.getCell(6).numFmt = '#,##0';

    totalRow1.getCell(10).value = computed.rowsHuongDan.length > 0
        ? { formula: `SUM(J${startRow1}:J${totalRowIdx1 - 1})`, result: computed.summary.tongTienHuongDan }
        : 0;
    totalRow1.getCell(10).numFmt = '#,##0';
    totalRow1.getCell(10).alignment = { horizontal: 'right', vertical: 'middle' };

    appendSignatureFooter(ws1, totalRowIdx1, 12);

    // =========================================================================
    // SHEET 2: Chấm PB
    // =========================================================================
    const ws2 = wb.addWorksheet('Chấm PB');
    setupHeaderAndFooter(ws2, 12, 'BẢNG THANH TOÁN TIỀN PHẢN BIỆN ĐATN');

    ws2.columns = [
        { width: 5 }, { width: 22 }, { width: 11 }, { width: 9 }, { width: 19 },
        { width: 11 }, { width: 9 }, { width: 12 }, { width: 13 }, { width: 16 }, { width: 13 }, { width: 14 }
    ];

    const hdrRow2 = ws2.getRow(6);
    hdrRow2.height = 30;
    const headers2 = ['TT', 'HỌ VÀ TÊN', '', 'Học vị', 'ĐƠN VỊ', 'SL ĐỀ TÀI', 'HỆ SỐ', 'QUY CHUẨN', 'ĐƠN GIÁ', 'SỐ TIỀN', 'KÝ TÊN', 'GHI CHÚ'];
    headers2.forEach((h, i) => {
        const cell = hdrRow2.getCell(i + 1);
        cell.value = h;
        cell.font = FONT_BOLD;
        cell.fill = HEADER_FILL;
        cell.border = THIN_BORDER;
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    ws2.mergeCells(6, 2, 6, 3);

    const startRow2 = 7;
    computed.rowsPhanBien.forEach((r, idx) => {
        const rowNum = startRow2 + idx;
        const row = ws2.getRow(rowNum);
        row.height = 22;

        const vals = [
            r.tt,
            r.hoDem,
            r.ten,
            r.hocVi,
            r.donVi,
            r.slDeTai,
            r.heSo,
            { formula: `F${rowNum}*G${rowNum}`, result: r.quyChuan },
            r.donGia,
            { formula: `H${rowNum}*I${rowNum}`, result: r.soTien },
            '',
            ''
        ];

        vals.forEach((v, cIdx) => {
            const colNum = cIdx + 1;
            const cell = row.getCell(colNum);
            cell.value = v;
            cell.font = FONT_BASE;

            if (colNum === 2) {
                cell.border = NAME_LEFT_BORDER;
                cell.alignment = { horizontal: 'left', vertical: 'middle' };
            } else if (colNum === 3) {
                cell.border = NAME_RIGHT_BORDER;
                cell.alignment = { horizontal: 'left', vertical: 'middle' };
            } else if (colNum === 5) {
                cell.border = THIN_BORDER;
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            } else if (colNum === 9 || colNum === 10) {
                cell.border = THIN_BORDER;
                cell.numFmt = '#,##0';
                cell.alignment = { horizontal: 'right', vertical: 'middle' };
            } else if (colNum === 7 || colNum === 8) {
                cell.border = THIN_BORDER;
                cell.numFmt = '0.00'; // Hiển thị 1,60 và 8,00 thay vì 8,
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            } else {
                cell.border = THIN_BORDER;
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            }
        });
    });

    const totalRowIdx2 = startRow2 + computed.rowsPhanBien.length;
    const totalRow2 = ws2.getRow(totalRowIdx2);
    totalRow2.height = 24;
    for (let c = 1; c <= 12; c++) {
        const cell = totalRow2.getCell(c);
        cell.font = FONT_BOLD;
        cell.border = THIN_BORDER;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    }
    ws2.mergeCells(totalRowIdx2, 2, totalRowIdx2, 3);
    totalRow2.getCell(2).value = 'TỔNG';
    totalRow2.getCell(6).value = computed.rowsPhanBien.length > 0
        ? { formula: `SUM(F${startRow2}:F${totalRowIdx2 - 1})`, result: computed.summary.tongSLPhanBien }
        : 0;
    totalRow2.getCell(6).numFmt = '#,##0';

    totalRow2.getCell(10).value = computed.rowsPhanBien.length > 0
        ? { formula: `SUM(J${startRow2}:J${totalRowIdx2 - 1})`, result: computed.summary.tongTienPhanBien }
        : 0;
    totalRow2.getCell(10).numFmt = '#,##0';
    totalRow2.getCell(10).alignment = { horizontal: 'right', vertical: 'middle' };

    appendSignatureFooter(ws2, totalRowIdx2, 12);

    // =========================================================================
    // SHEET 3: Chấm BV
    // =========================================================================
    const ws3 = wb.addWorksheet('Chấm BV');
    setupHeaderAndFooter(ws3, 12, 'BẢNG CHI TIẾT HỘI ĐỒNG CHẤM TỐT NGHIỆP');

    ws3.columns = [
        { width: 5 }, { width: 12 }, { width: 22 }, { width: 11 }, { width: 14 },
        { width: 20 }, { width: 8 }, { width: 9 }, { width: 10 }, { width: 13 }, { width: 16 }, { width: 14 }
    ];

    const hdrRow3 = ws3.getRow(6);
    hdrRow3.height = 42;
    const headers3 = ['STT', 'HỘI ĐỒNG SỐ', 'HỌ VÀ TÊN', '', 'CHỨC DANH', 'T.GIAN THAM GIA HĐ TỐT NGHIỆP', 'SLSV', 'HỆ SỐ', 'Giờ QĐ', 'ĐƠN GIÁ', 'THÀNH TIỀN', 'GHI CHÚ'];
    headers3.forEach((h, i) => {
        const cell = hdrRow3.getCell(i + 1);
        cell.value = h;
        cell.font = FONT_BOLD;
        cell.fill = HEADER_FILL;
        cell.border = THIN_BORDER;
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    ws3.mergeCells(6, 3, 6, 4);

    const startRow3 = 7;
    const blocksToMerge = [];

    computed.rowsHoiDong.forEach((r, idx) => {
        const rowNum = startRow3 + idx;
        const row = ws3.getRow(rowNum);
        row.height = 22;

        // Điền trước các cột đơn lẻ (chừa cột 2 & 6 để merge riêng sau vòng lặp)
        const cellStt = row.getCell(1);
        cellStt.value = r.stt;
        cellStt.font = FONT_BASE;
        cellStt.border = THIN_BORDER;
        cellStt.alignment = { horizontal: 'center', vertical: 'middle' };

        const cellHoDem = row.getCell(3);
        cellHoDem.value = r.hoDem;
        cellHoDem.font = FONT_BASE;
        cellHoDem.border = NAME_LEFT_BORDER;
        cellHoDem.alignment = { horizontal: 'left', vertical: 'middle' };

        const cellTen = row.getCell(4);
        cellTen.value = r.ten;
        cellTen.font = FONT_BASE;
        cellTen.border = NAME_RIGHT_BORDER;
        cellTen.alignment = { horizontal: 'left', vertical: 'middle' };

        const cellChucDanh = row.getCell(5);
        cellChucDanh.value = r.chucDanh;
        cellChucDanh.font = FONT_BASE;
        cellChucDanh.border = THIN_BORDER;
        cellChucDanh.alignment = { horizontal: 'center', vertical: 'middle' };

        const cellSlsv = row.getCell(7);
        cellSlsv.value = r.slsv;
        cellSlsv.font = FONT_BASE;
        cellSlsv.border = THIN_BORDER;
        cellSlsv.alignment = { horizontal: 'center', vertical: 'middle' };

        const cellHeSo = row.getCell(8);
        cellHeSo.value = r.heSo;
        cellHeSo.font = FONT_BASE;
        cellHeSo.border = THIN_BORDER;
        cellHeSo.numFmt = '0.00';
        cellHeSo.alignment = { horizontal: 'center', vertical: 'middle' };

        const cellGioQD = row.getCell(9);
        cellGioQD.value = { formula: `G${rowNum}*H${rowNum}`, result: r.gioQD };
        cellGioQD.font = FONT_BASE;
        cellGioQD.border = THIN_BORDER;
        cellGioQD.numFmt = '0.00';
        cellGioQD.alignment = { horizontal: 'center', vertical: 'middle' };

        const cellDonGia = row.getCell(10);
        cellDonGia.value = r.donGia;
        cellDonGia.font = FONT_BASE;
        cellDonGia.border = THIN_BORDER;
        cellDonGia.numFmt = '#,##0';
        cellDonGia.alignment = { horizontal: 'right', vertical: 'middle' };

        const cellThanhTien = row.getCell(11);
        cellThanhTien.value = { formula: `I${rowNum}*J${rowNum}`, result: r.thanhTien };
        cellThanhTien.font = FONT_BASE;
        cellThanhTien.border = THIN_BORDER;
        cellThanhTien.numFmt = '#,##0';
        cellThanhTien.alignment = { horizontal: 'right', vertical: 'middle' };

        const cellGhiChu = row.getCell(12);
        cellGhiChu.value = '';
        cellGhiChu.font = FONT_BASE;
        cellGhiChu.border = THIN_BORDER;
        cellGhiChu.alignment = { horizontal: 'center', vertical: 'middle' };

        if (r.isFirstInBlock) {
            blocksToMerge.push({
                rowStart: rowNum,
                rowEnd: rowNum + 2,
                soHD: r.soHD,
                thoiGian: r.thoiGian
            });
        }
    });

    // Thực hiện gộp ô dọc 3 dòng và gán giá trị bằng địa chỉ B{row} / F{row} sau khi tất cả dòng đã dựng xong
    blocksToMerge.forEach(b => {
        ws3.mergeCells(`B${b.rowStart}:B${b.rowEnd}`);
        const hdCell = ws3.getCell(`B${b.rowStart}`);
        hdCell.value = b.soHD;
        hdCell.font = FONT_BASE;
        hdCell.border = THIN_BORDER;
        hdCell.alignment = { horizontal: 'center', vertical: 'middle' };

        ws3.mergeCells(`F${b.rowStart}:F${b.rowEnd}`);
        const timeCell = ws3.getCell(`F${b.rowStart}`);
        timeCell.value = b.thoiGian;
        timeCell.font = FONT_BASE;
        timeCell.border = THIN_BORDER;
        timeCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });

    const totalRowIdx3 = startRow3 + computed.rowsHoiDong.length;
    const totalRow3 = ws3.getRow(totalRowIdx3);
    totalRow3.height = 24;
    for (let c = 1; c <= 12; c++) {
        const cell = totalRow3.getCell(c);
        cell.font = FONT_BOLD;
        cell.border = THIN_BORDER;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    }
    ws3.mergeCells(totalRowIdx3, 2, totalRowIdx3, 6);
    totalRow3.getCell(2).value = 'TỔNG';
    totalRow3.getCell(7).value = computed.summary.tongSLSVHoiDong;
    totalRow3.getCell(7).numFmt = '#,##0';

    totalRow3.getCell(9).value = computed.rowsHoiDong.length > 0
        ? { formula: `SUM(I${startRow3}:I${totalRowIdx3 - 1})`, result: computed.summary.tongGioQDHoiDong }
        : 0;
    totalRow3.getCell(9).numFmt = '0.00';

    totalRow3.getCell(11).value = computed.rowsHoiDong.length > 0
        ? { formula: `SUM(K${startRow3}:K${totalRowIdx3 - 1})`, result: computed.summary.tongTienHoiDong }
        : 0;
    totalRow3.getCell(11).numFmt = '#,##0';
    totalRow3.getCell(11).alignment = { horizontal: 'right', vertical: 'middle' };

    appendSignatureFooter(ws3, totalRowIdx3, 12);

    // =========================================================================
    // SHEET 4 (ẨN): _META_DATA
    // =========================================================================
    const wsMeta = wb.addWorksheet('_META_DATA', { state: 'hidden' });
    wsMeta.getCell('A1').value = 'DATN_PAYMENT_META_V1';
    wsMeta.getCell('A2').value = JSON.stringify({
        version: 1,
        exportedAt: new Date().toISOString(),
        stamp: stampText,
        khoa: config.khoa || sourceData.khoa || '',
        nganh: config.nganh || sourceData.nganh || '',
        khoaHoc: config.khoaHoc || sourceData.khoaHoc || '',
        truongThanhVien: config.truongThanhVien || '',
        donGia: config.donGia,
        rowsHuongDan: computed.rowsHuongDan,
        rowsPhanBien: computed.rowsPhanBien,
        rowsHoiDong: computed.rowsHoiDong,
        summary: computed.summary
    });

    return wb;
}

function arrayBufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
        binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
    }
    return btoa(binary);
}

async function uploadPaymentExcelToSessionFolder(xlsxBuffer, fileName, metadata) {
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

    const base64Data = arrayBufferToBase64(xlsxBuffer);

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
                                        <label class="form-label mb-0 fw-semibold">Ngành đào tạo (mỗi ngành 1 dòng):</label>
                                        <textarea id="pay-cfg-nganh" rows="2" class="form-control form-control-sm pay-cfg-input"></textarea>
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
                            📌 Font <strong>Times New Roman 12</strong>, tự động gắn giá trị <code>${escapeHtml(getSetDtuSerialStamp())}</code> căn phải ở dòng 1 &amp; lưu Phiên (60 ngày).
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

export function openPaymentPreviewModal(getMetaFn) {
    const meta = typeof getMetaFn === 'function' ? getMetaFn() : {};
    const sourceData = collectPaymentSourceFromDOM(meta);

    if (sourceData.assignedStudents.length === 0) {
        alert('⚠️ Chưa có sinh viên nào được xếp vào Hội đồng để lập bảng thanh toán!');
        return;
    }

    const modalEl = ensurePaymentPreviewModal();
    const savedCfg = loadSavedPaymentConfig();

    const nganhLinesInit = normalizeMultiNganhLines(sourceData.nganh);
    const nganhTextarea = document.getElementById('pay-cfg-nganh');
    nganhTextarea.value = nganhLinesInit.join('\n');
    nganhTextarea.rows = Math.min(4, Math.max(1, nganhLinesInit.length));

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
        nganh: normalizeMultiNganhLines(document.getElementById('pay-cfg-nganh').value).join('\n'),
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
        const stampText = getSetDtuSerialStamp();
        const khoaCleanUpper = (cfg.khoa || '...').toUpperCase().replace(/^KHOA\s+/i, '');

        const previewNganhLines = normalizeMultiNganhLines(cfg.nganh)
            .map(line => line.toUpperCase().replace(/^NGÀNH\s+/i, ''));
        const previewNganhHtml = previewNganhLines.length === 0
            ? 'NGÀNH ...'
            : previewNganhLines.length === 1
                ? `NGÀNH ${escapeHtml(previewNganhLines[0])}`
                : `NGÀNH:<br>${previewNganhLines.map(escapeHtml).join('<br>')}`;

        const headerBanner = `
            <div class="border rounded p-2 mb-2 bg-light" style="font-family: 'Times New Roman', serif; font-size: 13px;">
                <div class="text-end text-muted fst-italic" style="font-size: 11px;">${escapeHtml(stampText)}</div>
                <div class="row text-center fw-bold mt-1 align-items-center">
                    <div class="col-5">
                        <div>ĐẠI HỌC DUY TÂN</div>
                        <div class="text-decoration-underline">KHOA ${escapeHtml(khoaCleanUpper)}</div>
                    </div>
                    <div class="col-7">
                        <div>${previewNganhHtml}</div>
                        <div>KHÓA ${escapeHtml((cfg.khoaHoc || '...').toUpperCase().replace(/^KHÓA\s+/i, ''))}</div>
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
            <table class="table table-bordered table-sm text-center align-middle mb-0" style="font-family: 'Times New Roman', serif; font-size: 13px;">
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
                            <td class="text-start border-end-0">${escapeHtml(r.hoDem)}</td>
                            <td class="text-start border-start-0">${escapeHtml(r.ten)}</td>
                            <td>${escapeHtml(r.hocVi)}</td>
                            <td>${escapeHtml(r.donVi)}</td>
                            <td>${r.slDeTai}</td>
                            <td>${formatDecimalVN(r.heSo)}</td>
                            <td>${formatDecimalVN(r.quyChuan)}</td>
                            <td class="text-end">${formatVND(r.donGia)}</td>
                            <td class="text-end">${formatVND(r.soTien)}</td>
                        </tr>
                    `).join('')}
                </tbody>
                <tfoot class="fw-bold">
                    <tr>
                        <td colspan="5">Tổng</td>
                        <td>${s.tongSLHuongDan}</td>
                        <td colspan="3"></td>
                        <td class="text-end">${formatVND(s.tongTienHuongDan)}</td>
                    </tr>
                </tfoot>
            </table>
        `;

        document.getElementById('tab-pay-pb').innerHTML = `
            ${headerBanner}
            <table class="table table-bordered table-sm text-center align-middle mb-0" style="font-family: 'Times New Roman', serif; font-size: 13px;">
                <thead class="table-secondary">
                    <tr>
                        <th>TT</th><th colspan="2">HỌ VÀ TÊN</th><th>Học vị</th><th>ĐƠN VỊ</th>
                        <th>SL ĐỀ TÀI</th><th>HỆ SỐ</th><th>QUY CHUẨN</th><th>ĐƠN GIÁ</th><th>SỐ TIỀN</th>
                    </tr>
                </thead>
                <tbody>
                    ${computed.rowsPhanBien.map(r => `
                        <tr>
                            <td>${r.tt}</td>
                            <td class="text-start border-end-0">${escapeHtml(r.hoDem)}</td>
                            <td class="text-start border-start-0">${escapeHtml(r.ten)}</td>
                            <td>${escapeHtml(r.hocVi)}</td>
                            <td>${escapeHtml(r.donVi)}</td>
                            <td>${r.slDeTai}</td>
                            <td>${formatDecimalVN(r.heSo)}</td>
                            <td>${formatDecimalVN(r.quyChuan)}</td>
                            <td class="text-end">${formatVND(r.donGia)}</td>
                            <td class="text-end">${formatVND(r.soTien)}</td>
                        </tr>
                    `).join('')}
                </tbody>
                <tfoot class="fw-bold">
                    <tr>
                        <td colspan="5">TỔNG</td>
                        <td>${s.tongSLPhanBien}</td>
                        <td colspan="3"></td>
                        <td class="text-end">${formatVND(s.tongTienPhanBien)}</td>
                    </tr>
                </tfoot>
            </table>
        `;

        document.getElementById('tab-pay-bv').innerHTML = `
            ${headerBanner}
            <table class="table table-bordered table-sm text-center align-middle mb-0" style="font-family: 'Times New Roman', serif; font-size: 13px;">
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
                            ${r.isFirstInBlock ? `<td rowspan="3">${escapeHtml(r.soHD)}</td>` : ''}
                            <td class="text-start border-end-0">${escapeHtml(r.hoDem)}</td>
                            <td class="text-start border-start-0">${escapeHtml(r.ten)}</td>
                            <td>${escapeHtml(r.chucDanh)}</td>
                            ${r.isFirstInBlock ? `<td rowspan="3">${escapeHtml(r.thoiGian)}</td>` : ''}
                            <td>${r.slsv}</td>
                            <td>${formatDecimalVN(r.heSo)}</td>
                            <td>${formatDecimalVN(r.gioQD)}</td>
                            <td class="text-end">${formatVND(r.donGia)}</td>
                            <td class="text-end">${formatVND(r.thanhTien)}</td>
                        </tr>
                    `).join('')}
                </tbody>
                <tfoot class="fw-bold">
                    <tr>
                        <td colspan="6">TỔNG</td>
                        <td>${s.tongSLSVHoiDong}</td>
                        <td></td>
                        <td>${formatDecimalVN(s.tongGioQDHoiDong)}</td>
                        <td></td>
                        <td class="text-end">${formatVND(s.tongTienHoiDong)}</td>
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

        const origText = btnConfirm.innerHTML;
        btnConfirm.disabled = true;
        btnConfirm.innerHTML = '⏳ Đang tạo file Excel chuẩn định dạng...';

        try {
            const wb = await buildStyledPaymentWorkbook(sourceData, cfg, computed);
            const xlsxBuffer = await wb.xlsx.writeBuffer();

            const nganhSlug = sanitizeFileSlug(cfg.nganh.replace(/\r?\n/g, '_') || 'Nganh').slice(0, 60);
            const khoaSlug = sanitizeFileSlug(cfg.khoaHoc || 'DATN').slice(0, 40);
            const fileName = `Thanh_Toan_DATN_${nganhSlug}_${khoaSlug}.xlsx`;

            const blob = new Blob([xlsxBuffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 5000);

            btnConfirm.innerHTML = '⏳ Đang lưu bản sao lên Phiên Cloud...';
            const cloudRes = await uploadPaymentExcelToSessionFolder(xlsxBuffer, fileName, {
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
            console.error('Lỗi xuất bảng thanh toán:', err);
            alert(`❌ Có lỗi khi xuất file Excel thanh toán: ${err.message}`);
        } finally {
            btnConfirm.disabled = false;
            btnConfirm.innerHTML = origText;
            bootstrap.Modal.getInstance(modalEl)?.hide();
        }
    };

    renderPreviewContent();
    bootstrap.Modal.getOrCreateInstance(modalEl).show();
}
