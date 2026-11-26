// js/student-master-validator.js
// Tự động nạp ./sample_files/DANH_SACH_SV_TOAN_TRUONG.xlsx và đối soát Mã SV + Họ và tên SV

let cachedMasterPromise = null;

export function normalizeStudentNameKey(rawName) {
    return String(rawName || '')
        .normalize('NFC')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
}

function normalizeHeader(val) {
    return String(val || '')
        .normalize('NFC')
        .replace(/\s+/g, ' ')
        .trim()
        .toUpperCase();
}

/**
 * Tải và lập chỉ mục file ./sample_files/DANH_SACH_SV_TOAN_TRUONG.xlsx (chỉ tải 1 lần, cache trong RAM)
 */
export async function loadSchoolStudentMaster(url = './sample_files/DANH_SACH_SV_TOAN_TRUONG.xlsx') {
    if (cachedMasterPromise) return cachedMasterPromise;

    cachedMasterPromise = (async () => {
        try {
            if (typeof XLSX === 'undefined') {
                console.warn('[StudentMaster] Chưa nạp thư viện XLSX.');
                return { loaded: false, byId: new Map(), error: 'Thiếu thư viện XLSX.' };
            }

            const resp = await fetch(url);
            if (!resp.ok) {
                throw new Error(`HTTP ${resp.status} khi tải ${url}`);
            }

            const buffer = await resp.arrayBuffer();
            const wb = XLSX.read(buffer, { type: 'array' });
            const byId = new Map(); // maSV_upper -> { maSV, hoTen, nameKey, lop, nganh }

            wb.SheetNames.forEach(sheetName => {
                const ws = wb.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                if (!rows || rows.length === 0) return;

                let headerIdx = -1;
                let colMSSV = -1;
                let colHoTen = -1;
                let colHoDem = -1;
                let colTen = -1;
                let colLop = -1;
                let colNganh = -1;

                for (let i = 0; i < Math.min(rows.length, 15); i++) {
                    const rUp = (rows[i] || []).map(normalizeHeader);
                    const idxId = rUp.findIndex(c =>
                        c === 'MSSV' || c === 'MÃ SV' || c === 'MÃ SỐ SV' || c === 'MÃ SINH VIÊN' || c === 'MASV'
                    );
                    if (idxId === -1) continue;

                    const idxFull = rUp.findIndex(c => c.includes('HỌ VÀ TÊN') || c.includes('HỌ TÊN'));
                    const idxLast = rUp.findIndex(c => c === 'HỌ ĐỆM' || c === 'HỌ LÓT' || c === 'HỌ');
                    const idxFirst = rUp.findIndex(c => c === 'TÊN');

                    if (idxFull !== -1 || (idxLast !== -1 && idxFirst !== -1)) {
                        headerIdx = i;
                        colMSSV = idxId;
                        colHoTen = idxFull;
                        colHoDem = idxLast;
                        colTen = idxFirst;
                        colLop = rUp.findIndex(c => c === 'LỚP' || c.includes('LỚP SH') || c === 'KHÓA');
                        colNganh = rUp.findIndex(c => c.includes('NGÀNH') || c.includes('CHUYÊN NGÀNH'));
                        break;
                    }
                }

                if (headerIdx === -1) return;

                // Kiểm tra trường hợp cột Họ và tên gộp 2 ô liền nhau (ô sau trống tiêu đề, ô tiếp theo là Lớp)
                const isSplitFullCol =
                    colHoTen !== -1 &&
                    colLop !== -1 &&
                    colLop - colHoTen === 2 &&
                    !normalizeHeader(rows[headerIdx][colHoTen + 1]);

                for (let r = headerIdx + 1; r < rows.length; r++) {
                    const row = rows[r] || [];
                    const maSV = String(row[colMSSV] || '').trim();
                    if (!maSV || normalizeHeader(maSV).includes('MÃ SV') || normalizeHeader(maSV) === 'MSSV') continue;

                    let hoTen = '';
                    if (colHoTen !== -1) {
                        const p1 = String(row[colHoTen] || '').trim();
                        const p2 = isSplitFullCol ? String(row[colHoTen + 1] || '').trim() : '';
                        hoTen = `${p1} ${p2}`.replace(/\s+/g, ' ').trim();
                    } else if (colHoDem !== -1 && colTen !== -1) {
                        hoTen = `${String(row[colHoDem] || '').trim()} ${String(row[colTen] || '').trim()}`
                            .replace(/\s+/g, ' ')
                            .trim();
                    }

                    if (!hoTen) continue;

                    const lop = colLop !== -1 ? String(row[colLop] || '').replace(/\s+/g, '').trim() : '';
                    const nganh = colNganh !== -1 ? String(row[colNganh] || '').replace(/\s+/g, ' ').trim() : '';

                    byId.set(maSV.toUpperCase(), {
                        maSV,
                        hoTen,
                        nameKey: normalizeStudentNameKey(hoTen),
                        lop,
                        nganh
                    });
                }
            });

            console.log(`✅ [StudentMaster] Đã nạp ${byId.size} sinh viên từ ${url}`);
            return { loaded: byId.size > 0, byId, error: byId.size === 0 ? 'File danh sách SV toàn trường trống.' : '' };
        } catch (err) {
            console.error('[StudentMaster] Lỗi nạp danh sách SV toàn trường:', err);
            cachedMasterPromise = null; // Cho phép thử tải lại ở lần gọi sau
            return { loaded: false, byId: new Map(), error: err.message || String(err) };
        }
    })();

    return cachedMasterPromise;
}

/**
 * Đối soát danh sách bản ghi sinh viên với DANH_SACH_SV_TOAN_TRUONG.xlsx
 * @param {Array<{ maSV: string, hoTen: string, location?: string }>} entries
 * @returns {Promise<{ ok: boolean, skipped: boolean, errors: string[] }>}
 */
export async function validateStudentsAgainstSchoolMaster(entries) {
    const master = await loadSchoolStudentMaster();
    if (!master.loaded || master.byId.size === 0) {
        return {
            ok: true,
            skipped: true,
            warning: `⚠️ Không thể tải file đối soát DANH_SACH_SV_TOAN_TRUONG.xlsx (${master.error || 'không có dữ liệu'}). Bỏ qua bước đối soát toàn trường.`,
            errors: []
        };
    }

    const errors = [];

    (entries || []).forEach(item => {
        const rawId = String(item.maSV || '').trim();
        const rawName = String(item.hoTen || '').replace(/\s+/g, ' ').trim();
        const locPrefix = item.location ? `${item.location}: ` : '';

        if (!rawId) return;

        const found = master.byId.get(rawId.toUpperCase());
        if (!found) {
            errors.push(
                `• ${locPrefix}Mã SV "${rawId}" (${rawName || 'Chưa rõ tên'}) KHÔNG TỒN TẠI trong DANH_SACH_SV_TOAN_TRUONG.xlsx.`
            );
            return;
        }

        const inputNameKey = normalizeStudentNameKey(rawName);
        if (inputNameKey && found.nameKey && inputNameKey !== found.nameKey) {
            errors.push(
                `• ${locPrefix}Sai Họ tên của Mã SV "${rawId}": File nạp vào ghi "${rawName}" ≠ Danh sách toàn trường ghi "${found.hoTen}".`
            );
        }
    });

    return {
        ok: errors.length === 0,
        skipped: false,
        errors
    };
}
