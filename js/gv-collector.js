// js/gv-collector.js
import { normalizeTeacherKeyForDupCheck } from './validators.js';

function escapeHtmlAttr(val) {
    return String(val ?? '')
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function isValidTeacherCode(maGV) {
    const id = String(maGV || '').trim().toUpperCase();
    return Boolean(id && !id.startsWith('GV_'));
}

function getMergeKey(maGV, hoTenGV) {
    const id = String(maGV || '').trim().toUpperCase();
    if (isValidTeacherCode(id)) {
        return `ID:${id}`;
    }
    const normName = normalizeTeacherKeyForDupCheck(hoTenGV);
    return normName ? `NAME:${normName}` : '';
}

/**
 * Thu thập & hợp nhất danh sách Giảng viên từ:
 * 1. Kho GV đang hiển thị trên DOM (#pool-gv)
 * 2. File Danh sách Giảng viên (gvList)
 * 3. File Danh sách Sinh viên & Đề tài (svList - cột Mã GVHD & GVHD)
 * 4. File Danh sách Hội đồng (parsedHoidongs - thành viên các HĐ)
 * Tự động tránh trùng lặp theo Mã GV và làm giàu thông tin (Đơn vị, Họ tên đầy đủ học vị).
 */
export function collectUniqueGiangVien({
    gvList = [],
    svList = [],
    parsedHoidongs = {},
    includeExistingPool = true
} = {}) {
    const gvMap = new Map(); // key -> { maGV, hoTenGV, donVi }
    let autoSeq = 1;

    const upsertTeacher = ({ maGV = '', hoTenGV = '', donVi = '' }) => {
        const cleanId = String(maGV || '').trim();
        const cleanName = String(hoTenGV || '').trim();
        const cleanDonVi = String(donVi || '').trim();

        if (!cleanId && !cleanName) return;

        const key = getMergeKey(cleanId, cleanName);
        if (!key) return;

        if (!gvMap.has(key)) {
            gvMap.set(key, {
                maGV: cleanId || `GV_AUTO_${autoSeq++}`,
                hoTenGV: cleanName || cleanId,
                donVi: cleanDonVi || 'Đại học Duy Tân'
            });
            return;
        }

        const existing = gvMap.get(key);

        // Ưu tiên cập nhật Mã GV thật nếu trước đó chỉ có mã tạm
        if (!isValidTeacherCode(existing.maGV) && isValidTeacherCode(cleanId)) {
            existing.maGV = cleanId;
        }

        // Ưu tiên giữ tên đầy đủ hơn (ví dụ có học hàm/học vị "TS. Nguyễn Văn A" thay vì "Nguyễn Văn A")
        if (cleanName && (!existing.hoTenGV || cleanName.length > existing.hoTenGV.length)) {
            existing.hoTenGV = cleanName;
        }

        // Ưu tiên Đơn vị cụ thể thay vì giá trị mặc định
        if (
            cleanDonVi &&
            cleanDonVi !== 'Đại học Duy Tân' &&
            (!existing.donVi || existing.donVi === 'Đại học Duy Tân')
        ) {
            existing.donVi = cleanDonVi;
        }
    };

    // Nguồn 0: Dữ liệu GV đang có sẵn trong Kho (#pool-gv) để không làm mất khi nạp bổ sung file
    if (includeExistingPool) {
        document.querySelectorAll('#pool-gv .card-item').forEach(el => {
            upsertTeacher({
                maGV: el.getAttribute('data-id') || '',
                hoTenGV: el.getAttribute('data-name') || '',
                donVi: el.getAttribute('data-donvi') || ''
            });
        });
    }

    // Nguồn 1: File Danh sách Giảng viên
    (Array.isArray(gvList) ? gvList : []).forEach(gv => {
        upsertTeacher({
            maGV: gv.maGV,
            hoTenGV: gv.hoTenGV,
            donVi: gv.donVi
        });
    });

    // Nguồn 2: File Danh sách Sinh viên & Đề tài (GVHD)
    (Array.isArray(svList) ? svList : []).forEach(sv => {
        upsertTeacher({
            maGV: sv.maGVHD,
            hoTenGV: sv.gvhd,
            donVi: sv.donViGVHD || 'Đại học Duy Tân'
        });
    });

    // Nguồn 3: File Danh sách Hội đồng (Chủ tịch, Phản biện, Thư ký)
    Object.keys(parsedHoidongs || {}).forEach(key => {
        if (key.startsWith('__')) return;
        const hd = parsedHoidongs[key];
        if (!hd || typeof hd !== 'object') return;

        (Array.isArray(hd.thanhVien) ? hd.thanhVien : []).forEach(tv => {
            upsertTeacher({
                maGV: tv.maGV,
                hoTenGV: tv.hoTenGV,
                donVi: tv.donVi
            });
        });
    });

    return Array.from(gvMap.values()).sort((a, b) =>
        a.hoTenGV.localeCompare(b.hoTenGV, 'vi')
    );
}

/**
 * Thu thập toàn bộ GV từ các file đã đọc, khử trùng theo Mã GV và render thẳng vào #pool-gv
 * @returns {{ totalCount: number, addedCount: number, mergedList: Array }}
 */
export function syncCollectedGiangVienToPool({
    gvList = [],
    svList = [],
    parsedHoidongs = {},
    includeExistingPool = true
} = {}) {
    const pool = document.getElementById('pool-gv');
    const beforeCount = pool ? pool.querySelectorAll('.card-item').length : 0;

    const mergedList = collectUniqueGiangVien({
        gvList,
        svList,
        parsedHoidongs,
        includeExistingPool
    });

    if (pool) {
        pool.innerHTML = mergedList.map(gv => `
            <div class="card-item position-relative mb-2 p-2 pe-4 border border-info rounded bg-white shadow-sm" 
                 data-id="${escapeHtmlAttr(gv.maGV)}" data-name="${escapeHtmlAttr(gv.hoTenGV)}" data-donvi="${escapeHtmlAttr(gv.donVi || 'Đại học Duy Tân')}">
                <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card" title="Xóa giảng viên này">✕</button>
                <div class="fw-bold text-dark">${escapeHtmlAttr(gv.hoTenGV)}</div>
                <div class="text-muted small">${escapeHtmlAttr(gv.donVi || 'Đại học Duy Tân')} (Mã: ${escapeHtmlAttr(gv.maGV)})</div>
            </div>
        `).join('');
    }

    return {
        totalCount: mergedList.length,
        addedCount: Math.max(0, mergedList.length - beforeCount),
        mergedList
    };
}