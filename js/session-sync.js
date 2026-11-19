// js/session-sync.js
import { createHoiDongBox } from './hoidong-manager.js';
import { templates } from './templates.js';

const GAS_URL_STORAGE_KEY = 'datn_gas_webapp_url';
const LAST_SESSION_META_KEY = 'datn_last_cloud_session';

// Lưu URL Web App GAS vào cấu hình trình duyệt (hoặc gán URL mặc định của bạn tại đây)
const DEFAULT_GAS_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbxWmG5jDte-EkGQdvchrD1wQjw_8jRdC3XoOT2ly9K_30op7El4GAsZ87GWDr7CUiLS/exec'; // Dán URL https://script.google.com/macros/s/.../exec của bạn vào đây nếu muốn dùng chung cho mọi máy

export function getGasWebAppUrl() {
    return localStorage.getItem(GAS_URL_STORAGE_KEY) || DEFAULT_GAS_WEBAPP_URL;
}

export function setGasWebAppUrl(url) {
    localStorage.setItem(GAS_URL_STORAGE_KEY, String(url || '').trim());
}

export function getSavedSessionMeta() {
    try {
        return JSON.parse(localStorage.getItem(LAST_SESSION_META_KEY) || '{}');
    } catch {
        return {};
    }
}

export function saveSessionMeta(meta) {
    localStorage.setItem(LAST_SESSION_META_KEY, JSON.stringify(meta));
}

// Helper chuyển ArrayBuffer <-> Base64 để đồng bộ cả file Word Template tùy chỉnh nếu có
function arrayBufferToBase64(buffer) {
    if (!buffer) return null;
    const bytes = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
    return btoa(binary);
}

function base64ToArrayBuffer(base64) {
    if (!base64) return null;
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes.buffer;
}

// 1. ĐÓNG GÓI TOÀN BỘ PHIÊN LÀM VIỆC HIỆN TẠI THÀNH JSON
export function serializeCurrentSession() {
    const setup = {
        khoa: document.getElementById('khoa-select')?.value || '',
        nganh: document.getElementById('nganh-select')?.value || '',
        dotThang: document.getElementById('dot-thang')?.value || '',
        namCham: document.getElementById('nam-cham')?.value || '',
        soQD: document.getElementById('so-qd')?.value || '',
        ngayKyQD: document.getElementById('ngay-ky-qd')?.value || ''
    };

    const readSVCard = (el) => ({
        maSV: el.getAttribute('data-id') || '',
        hoTen: el.getAttribute('data-name') || '',
        lop: el.getAttribute('data-lop') || '',
        deTai: el.getAttribute('data-detai') || '',
        maGVHD: el.getAttribute('data-magvhd') || '',
        gvhd: el.getAttribute('data-gvhd') || '',
        nganh: el.getAttribute('data-nganh') || '',
        diem_gvhd: el.getAttribute('data-diem-gvhd') || '',
        diem_ct: el.getAttribute('data-diem-ct') || '',
        diem_pb: el.getAttribute('data-diem-pb') || '',
        diem_tk: el.getAttribute('data-diem-tk') || ''
    });

    const readGVCard = (el) => ({
        maGV: el.getAttribute('data-id') || '',
        hoTenGV: el.getAttribute('data-name') || '',
        donVi: el.getAttribute('data-donvi') || 'Đại học Duy Tân'
    });

    const poolSV = Array.from(document.querySelectorAll('#pool-sv .card-item')).map(readSVCard);
    const poolGV = Array.from(document.querySelectorAll('#pool-gv .card-item')).map(readGVCard);

    const hoidongs = Array.from(document.querySelectorAll('.box-hoidong')).map(box => {
        const getMember = (role) => {
            const card = box.querySelector(`.dropzone-gv[data-role="${role}"] .card-item`);
            return card ? readGVCard(card) : null;
        };
        return {
            title: box.querySelector('.hd-title')?.value || '',
            date: box.querySelector('.hd-date')?.value || '',
            time: box.querySelector('.hd-time')?.value || '',
            room: box.querySelector('.hd-room')?.value || '',
            location: box.querySelector('.hd-location')?.value || '',
            chuTich: getMember('Chủ tịch'),
            phanBien: getMember('Phản biện'),
            thuKy: getMember('Thư ký'),
            sinhViens: Array.from(box.querySelectorAll('.dropzone-sv .card-item')).map(readSVCard)
        };
    });

    // Lưu kèm các file mẫu phiếu chấm nếu người dùng có tải mẫu riêng trong phiên này
    const customTemplates = {};
    ['phieuGVHD', 'phieuGVPB', 'phieuHD'].forEach(k => {
        if (templates.modifiedFlags?.[k] && templates[k]) {
            customTemplates[k] = arrayBufferToBase64(templates[k]);
        }
    });

    const summary = `${setup.nganh || setup.khoa || 'Chưa chọn ngành'} | ${hoidongs.length} HĐ | ${poolSV.length} SV chờ`;

    return {
        summary,
        payload: {
            version: 1,
            setup,
            poolSV,
            poolGV,
            hoidongs,
            studentScores: window.studentScores || {},
            customTemplates
        }
    };
}

// 2. KHÔI PHỤC TOÀN BỘ PHIÊN LÀM VIỆC TỪ DỮ LIỆU JSON CLOUD
export function restoreSessionFromPayload(payload) {
    if (!payload) return;

    // A. Khôi phục Setup Metadata
    const s = payload.setup || {};
    const khoaSel = document.getElementById('khoa-select');
    if (khoaSel && s.khoa) {
        khoaSel.value = s.khoa;
        khoaSel.dispatchEvent(new Event('change'));
    }
    const nganhSel = document.getElementById('nganh-select');
    if (nganhSel && s.nganh) nganhSel.value = s.nganh;
    if (document.getElementById('dot-thang')) document.getElementById('dot-thang').value = s.dotThang || '';
    if (document.getElementById('nam-cham')) document.getElementById('nam-cham').value = s.namCham || '';
    if (document.getElementById('so-qd')) document.getElementById('so-qd').value = s.soQD || '';
    if (document.getElementById('ngay-ky-qd')) document.getElementById('ngay-ky-qd').value = s.ngayKyQD || '';

    // B. Khôi phục Bảng điểm
    window.studentScores = payload.studentScores || {};

    // C. Khôi phục Kho SV
    const poolSVEl = document.getElementById('pool-sv');
    if (poolSVEl) {
        poolSVEl.innerHTML = '';
        (payload.poolSV || []).forEach(sv => {
            const maGVDisplay = sv.maGVHD ? ` [${sv.maGVHD}]` : '';
            poolSVEl.insertAdjacentHTML('beforeend', `
                <div class="card-item position-relative mb-2 p-2 pe-4 border rounded bg-white shadow-sm" 
                     data-id="${sv.maSV}" data-name="${sv.hoTen}" data-lop="${sv.lop}" data-detai="${sv.deTai}" 
                     data-magvhd="${sv.maGVHD || ''}" data-gvhd="${sv.gvhd || ''}" data-nganh="${sv.nganh || ''}">
                    <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card">✕</button>
                    <div class="fw-bold text-primary">${sv.hoTen} (${sv.maSV})</div>
                    <div class="text-muted small">Lớp: ${sv.lop} | GVHD: <span class="fw-bold text-dark">${sv.gvhd || 'Chưa rõ'}${maGVDisplay}</span></div>
                    <div class="text-dark small text-truncate" title="${sv.deTai}">Đề tài: ${sv.deTai}</div>
                </div>`);
        });
    }

    // D. Khôi phục Kho GV
    const poolGVEl = document.getElementById('pool-gv');
    if (poolGVEl) {
        poolGVEl.innerHTML = '';
        (payload.poolGV || []).forEach(gv => {
            poolGVEl.insertAdjacentHTML('beforeend', `
                <div class="card-item position-relative mb-2 p-2 pe-4 border rounded bg-white shadow-sm" 
                     data-id="${gv.maGV}" data-name="${gv.hoTenGV}" data-donvi="${gv.donVi || ''}">
                    <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card">✕</button>
                    <div class="fw-bold text-dark small">${gv.hoTenGV}</div>
                    <div class="text-muted text-truncate" style="font-size: 0.75rem;">${gv.donVi || 'Đại học Duy Tân'} - Mã: ${gv.maGV}</div>
                </div>`);
        });
    }

    // E. Khôi phục Khu vực Hội đồng
    const workspace = document.getElementById('workspace-hd');
    if (workspace) {
        workspace.innerHTML = '';
        (payload.hoidongs || []).forEach(hd => {
            const boxId = createHoiDongBox(hd.title, hd.date, hd.time, hd.room, hd.location);
            const boxEl = document.getElementById(boxId);
            if (!boxEl) return;

            const renderMember = (role, gv) => {
                if (!gv) return;
                const zone = boxEl.querySelector(`.dropzone-gv[data-role="${role}"]`);
                if (zone) {
                    zone.innerHTML = `
                        <div class="card-item position-relative bg-light border-info p-1 pe-4 rounded shadow-sm" 
                             data-id="${gv.maGV}" data-name="${gv.hoTenGV}" data-donvi="${gv.donVi || 'Đại học Duy Tân'}">
                            <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card">✕</button>
                            <div class="fw-bold text-dark small">${gv.hoTenGV}</div>
                            <div class="text-muted text-truncate" style="font-size: 0.75rem;">${gv.donVi || 'Đại học Duy Tân'} - Mã: ${gv.maGV}</div>
                        </div>`;
                }
            };
            renderMember('Chủ tịch', hd.chuTich);
            renderMember('Phản biện', hd.phanBien);
            renderMember('Thư ký', hd.thuKy);

            const svZone = boxEl.querySelector('.dropzone-sv');
            (hd.sinhViens || []).forEach(sv => {
                const maGVDisplay = sv.maGVHD ? ` [${sv.maGVHD}]` : '';
                svZone?.insertAdjacentHTML('beforeend', `
                    <div class="card-item position-relative mb-2 p-2 pe-4 border rounded bg-white shadow-sm" 
                         data-id="${sv.maSV}" data-name="${sv.hoTen}" data-lop="${sv.lop}" data-detai="${sv.deTai}" 
                         data-magvhd="${sv.maGVHD || ''}" data-gvhd="${sv.gvhd || ''}" data-nganh="${sv.nganh || ''}"
                         data-diem-gvhd="${sv.diem_gvhd || ''}" data-diem-ct="${sv.diem_ct || ''}" 
                         data-diem-pb="${sv.diem_pb || ''}" data-diem-tk="${sv.diem_tk || ''}">
                        <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card">✕</button>
                        <div class="fw-bold text-primary">${sv.hoTen} (${sv.maSV})</div>
                        <div class="text-muted small">Lớp: ${sv.lop} | GVHD: <span class="fw-bold text-dark">${sv.gvhd || 'Chưa rõ'}${maGVDisplay}</span></div>
                        <div class="text-dark small text-truncate" title="${sv.deTai}">Đề tài: ${sv.deTai}</div>
                    </div>`);
            });
        });
    }

    // F. Khôi phục các file Word Template tùy chỉnh nếu có
    if (payload.customTemplates) {
        ['phieuGVHD', 'phieuGVPB', 'phieuHD'].forEach(k => {
            if (payload.customTemplates[k]) {
                templates[k] = base64ToArrayBuffer(payload.customTemplates[k]);
                if (templates.modifiedFlags) templates.modifiedFlags[k] = true;
            }
        });
    }
}

// 3. GỌI API GOOGLE APPS SCRIPT (LƯU & TẢI PHIÊN)
export async function syncSessionToCloud(gasUrl, sessionId, pin) {
    const { summary, payload } = serializeCurrentSession();
    const res = await fetch(gasUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'save', sessionId, pin, summary, payload })
    });
    return await res.json();
}

export async function loadSessionFromCloud(gasUrl, sessionId, pin) {
    const res = await fetch(gasUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'load', sessionId, pin })
    });
    return await res.json();
}
