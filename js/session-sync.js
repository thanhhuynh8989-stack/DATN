// js/session-sync.js
import { createHoiDongBox, initGlobalSortables } from './hoidong-manager.js';
import { templates } from './templates.js';

const GAS_URL_STORAGE_KEY = 'datn_gas_webapp_url';
const LAST_SESSION_META_KEY = 'datn_last_cloud_session';
const DEFAULT_GAS_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbxWmG5jDte-EkGQdvchrD1wQjw_8jRdC3XoOT2ly9K_30op7El4GAsZ87GWDr7CUiLS/exec';

// Cờ nội bộ để tạm tắt AutoSave khi đang khôi phục DOM từ payload Cloud
let isRestoringSessionDOM = false;
let cancelPendingAutoSaveFn = null;

// Trạng thái phiên làm việc hiện tại trên Client
export const sessionState = {
    isCloudConnected: false,
    sessionId: '',
    pin: '',
    otp: '',
    canEditHD: true,     // Mặc định true khi làm việc offline; khi lên Cloud chỉ true nếu có OTP hợp lệ
    isHDLocked: false,   // Khi true: khóa vùng Hội đồng, mở khóa Nhập điểm & Xuất file
    lastSyncedScores: {} // Bản chụp điểm lần gần nhất đồng bộ với Cloud để tính Delta Diff
};

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

export function clearCloudAutoSaveTimer() {
    if (typeof cancelPendingAutoSaveFn === 'function') {
        cancelPendingAutoSaveFn();
    }
}

export function disconnectCloudSession() {
    clearCloudAutoSaveTimer();
    sessionState.isCloudConnected = false;
    sessionState.sessionId = '';
    sessionState.pin = '';
    sessionState.otp = '';
    sessionState.canEditHD = true;
    sessionState.isHDLocked = false;
    sessionState.lastSyncedScores = {};
    applySessionPermissionsToUI();
}

function clonePlainScores(scoresObj) {
    const clean = {};
    Object.keys(scoresObj || {}).forEach(maSV => {
        const s = scoresObj[maSV] || {};
        clean[maSV] = {
            gvhd: String(s.gvhd ?? ''),
            ct: String(s.ct ?? ''),
            pb: String(s.pb ?? ''),
            tk: String(s.tk ?? ''),
            chu: String(s.chu ?? '')
        };
    });
    return clean;
}

// Áp dụng điểm từ Cloud lên cả window.studentScores lẫn toàn bộ thẻ SV khớp mã trên DOM
export function applyScoresToUI(scoresObj) {
    window.studentScores = scoresObj || {};
    sessionState.lastSyncedScores = clonePlainScores(window.studentScores);

    Object.keys(window.studentScores).forEach(maSV => {
        const sc = window.studentScores[maSV] || {};
        document.querySelectorAll(`.dropzone-sv .card-item[data-id="${maSV}"]`).forEach(svCard => {
            svCard.setAttribute('data-diem-gvhd', sc.gvhd ?? '');
            svCard.setAttribute('data-diem-ct', sc.ct ?? '');
            svCard.setAttribute('data-diem-pb', sc.pb ?? '');
            svCard.setAttribute('data-diem-tk', sc.tk ?? '');
        });
    });
}

// Tính toán phần điểm thực sự thay đổi so với lần đồng bộ gần nhất (Delta Sync)
export function computeScoreDelta(currentScores) {
    const delta = {};
    let changedCount = 0;

    Object.keys(currentScores || {}).forEach(maSV => {
        const curr = currentScores[maSV] || {};
        const prev = sessionState.lastSyncedScores[maSV] || {};
        const diffFields = {};

        ['gvhd', 'ct', 'pb', 'tk', 'chu'].forEach(field => {
            const currVal = String(curr[field] ?? '').trim();
            const prevVal = String(prev[field] ?? '').trim();
            if (currVal !== prevVal) {
                diffFields[field] = currVal;
                changedCount++;
            }
        });

        if (Object.keys(diffFields).length > 0) {
            delta[maSV] = diffFields;
        }
    });

    return { delta, changedCount };
}

// Áp dụng trạng thái Khóa/Mở khóa giao diện theo đúng phân quyền & trạng thái phiên
export function applySessionPermissionsToUI() {
    const hdEditable = !sessionState.isCloudConnected
        ? true
        : (sessionState.canEditHD && !sessionState.isHDLocked);

    // 1. Vùng Thiết lập, Nạp Excel, Kho SV, Kho GV & Khu vực Hội đồng
    const setupInputs = document.querySelectorAll(
        '#khoa-select, #nganh-select, #dot-thang, #nam-cham, #so-qd, #ngay-ky-qd, #upload-sv, #upload-gv, #upload-hd, #btn-read-files, #btn-add-hd'
    );
    setupInputs.forEach(el => { el.disabled = !hdEditable; });

    document.querySelectorAll('.box-hoidong input, .box-hoidong select, .box-hoidong button').forEach(el => {
        el.disabled = !hdEditable;
    });

    document.querySelectorAll('.btn-remove-card').forEach(btn => {
        btn.classList.toggle('d-none', !hdEditable);
    });

    document.querySelectorAll('#pool-sv, #pool-gv, .dropzone-sv, .dropzone-gv').forEach(zone => {
        zone.style.pointerEvents = hdEditable ? 'auto' : 'none';
        zone.style.opacity = hdEditable ? '1' : '0.88';
    });
    if (hdEditable && typeof initGlobalSortables === 'function') {
        initGlobalSortables();
    }

    // 2. Nhóm tác vụ Nhập điểm & Xuất file
    const allowScoringAndExport = !sessionState.isCloudConnected
        ? true
        : (sessionState.canEditHD || sessionState.isHDLocked);

    const actionBtnIds = [
        'btn-nhap-diem',
        'btn-export-pb',
        'btn-open-upload-templates',
        'btn-export-hd',
        'btn-export-word',
        'btn-export-diem'
    ];
    actionBtnIds.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
            btn.disabled = !allowScoringAndExport;
            btn.title = allowScoringAndExport
                ? ''
                : 'Chỉ mở khi Chủ phiên đã Chốt & Khóa vùng Hội đồng!';
        }
    });

    // 3. Thanh trạng thái phiên trên Header (kèm nút Chốt & Khóa nhanh cho Chủ phiên)
    const banner = document.getElementById('session-role-banner');
    if (banner) {
        if (!sessionState.isCloudConnected) {
            banner.classList.add('d-none');
            banner.innerHTML = '';
        } else {
            banner.classList.remove('d-none');
            const roleText = sessionState.canEditHD
                ? '👑 Chủ phiên (Có OTP)'
                : '✍️ Thư ký nhập điểm (Chỉ có PIN)';
            const lockText = sessionState.isHDLocked
                ? '🔒 Hội đồng ĐÃ KHÓA (Thư ký đang được Nhập điểm & Xuất file)'
                : '🔓 Hội đồng ĐANG MỞ SỬA (Thư ký chưa thể nhập điểm)';

            const quickLockBtnHTML = sessionState.canEditHD
                ? `<button type="button" id="btn-banner-toggle-lock" class="btn btn-xs ${sessionState.isHDLocked ? 'btn-outline-dark' : 'btn-danger'} py-0 px-2 fw-bold ms-2" style="font-size: 0.75rem;">
                       ${sessionState.isHDLocked ? '🔓 Mở khóa sửa HĐ' : '🔒 Chốt & Khóa HĐ'}
                   </button>`
                : '';

            banner.innerHTML = `
                <div class="d-flex align-items-center flex-wrap gap-1">
                    <span>Phiên: <strong>${sessionState.sessionId}</strong> | ${roleText} | <strong>${lockText}</strong></span>
                    ${quickLockBtnHTML}
                </div>
            `;
            banner.className = `alert ${sessionState.isHDLocked ? 'alert-success' : 'alert-warning'} py-1 px-2 mb-1 small d-flex justify-content-between align-items-center`;
        }
    }
}

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
        nganh: el.getAttribute('data-nganh') || ''
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
            version: 2,
            setup,
            poolSV,
            poolGV,
            hoidongs,
            studentScores: window.studentScores || {},
            customTemplates
        }
    };
}

export function restoreSessionFromPayload(payload) {
    if (!payload) return;

    clearCloudAutoSaveTimer();
    isRestoringSessionDOM = true;

    try {
        const s = payload.setup || {};
        const khoaSel = document.getElementById('khoa-select');
        const nganhSel = document.getElementById('nganh-select');

        if (khoaSel && s.khoa) {
            if (!Array.from(khoaSel.options).some(opt => opt.value === s.khoa)) {
                khoaSel.insertAdjacentHTML('beforeend', `<option value="${s.khoa}">${s.khoa}</option>`);
            }
            khoaSel.value = s.khoa;
            khoaSel.dispatchEvent(new Event('change'));
        }

        if (nganhSel && s.nganh) {
            const matchedOpt = Array.from(nganhSel.options).find(
                opt => opt.value === s.nganh || opt.textContent.toLowerCase().includes(s.nganh.toLowerCase())
            );
            if (matchedOpt) {
                nganhSel.value = matchedOpt.value;
            } else {
                nganhSel.insertAdjacentHTML('beforeend', `<option value="${s.nganh}">${s.nganh}</option>`);
                nganhSel.value = s.nganh;
            }
        }

        const dotEl = document.getElementById('dot-thang');
        if (dotEl) dotEl.value = s.dotThang ?? '';

        const namEl = document.getElementById('nam-cham');
        if (namEl) namEl.value = s.namCham ?? '';

        const soQdEl = document.getElementById('so-qd');
        if (soQdEl) soQdEl.value = s.soQD ?? '';

        const ngayKyEl = document.getElementById('ngay-ky-qd');
        if (ngayKyEl) {
            let dateVal = String(s.ngayKyQD ?? '').trim();
            if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateVal)) {
                const [d, m, y] = dateVal.split('/');
                dateVal = `${y}-${m}-${d}`;
            }
            ngayKyEl.value = dateVal;
        }

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
                             data-magvhd="${sv.maGVHD || ''}" data-gvhd="${sv.gvhd || ''}" data-nganh="${sv.nganh || ''}">
                            <button type="button" class="btn btn-sm btn-link text-danger text-decoration-none fw-bold position-absolute top-0 end-0 px-2 py-1 btn-remove-card">✕</button>
                            <div class="fw-bold text-primary">${sv.hoTen} (${sv.maSV})</div>
                            <div class="text-muted small">Lớp: ${sv.lop} | GVHD: <span class="fw-bold text-dark">${sv.gvhd || 'Chưa rõ'}${maGVDisplay}</span></div>
                            <div class="text-dark small text-truncate" title="${sv.deTai}">Đề tài: ${sv.deTai}</div>
                        </div>`);
                });
            });
        }

        if (payload.customTemplates) {
            ['phieuGVHD', 'phieuGVPB', 'phieuHD'].forEach(k => {
                if (payload.customTemplates[k]) {
                    templates[k] = base64ToArrayBuffer(payload.customTemplates[k]);
                    if (templates.modifiedFlags) templates.modifiedFlags[k] = true;
                }
            });
        }

        applyScoresToUI(payload.studentScores || {});
        applySessionPermissionsToUI();
    } finally {
        // Nhả cờ sau khi MutationObserver đã qua lượt microtask hiện tại
        setTimeout(() => {
            isRestoringSessionDOM = false;
        }, 0);
    }
}

// --- CÁC HÀM GỌI API GOOGLE APPS SCRIPT (KIỂM TRA HTTP & JSON AN TOÀN) ---
async function callGasApi(gasUrl, bodyObj) {
    const res = await fetch(gasUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(bodyObj)
    });

    if (!res.ok) {
        throw new Error(`Máy chủ GAS phản hồi mã lỗi HTTP ${res.status}`);
    }

    const rawText = await res.text();
    try {
        return JSON.parse(rawText);
    } catch {
        throw new Error('Phản hồi từ Google Apps Script không đúng định dạng JSON (kiểm tra lại URL Web App hoặc quyền truy cập Everyone).');
    }
}

export async function createCloudSession(gasUrl, sessionId, pin, email) {
    clearCloudAutoSaveTimer();
    const { summary, payload } = serializeCurrentSession();
    return await callGasApi(gasUrl, {
        action: 'create_session',
        sessionId,
        pin,
        email,
        summary,
        payload
    });
}

export async function loadCloudSession(gasUrl, sessionId, pin, otp = '') {
    clearCloudAutoSaveTimer();
    return await callGasApi(gasUrl, {
        action: 'load',
        sessionId,
        pin,
        otp
    });
}

export async function saveCloudHoiDong(gasUrl, sessionId, pin, otp) {
    const { summary, payload } = serializeCurrentSession();
    return await callGasApi(gasUrl, {
        action: 'save_hd',
        sessionId,
        pin,
        otp,
        summary,
        payload
    });
}

export async function toggleCloudHDLock(gasUrl, sessionId, pin, otp, isHDLocked) {
    return await callGasApi(gasUrl, {
        action: 'toggle_hd_lock',
        sessionId,
        pin,
        otp,
        isHDLocked
    });
}

export async function pullLatestScoresFromCloud() {
    if (!sessionState.isCloudConnected) return { ok: true, offline: true };
    const gasUrl = getGasWebAppUrl();
    const res = await callGasApi(gasUrl, {
        action: 'pull_scores',
        sessionId: sessionState.sessionId,
        pin: sessionState.pin
    });
    if (res.ok && res.studentScores && typeof res.studentScores === 'object') {
        applyScoresToUI(res.studentScores);
    }
    return res;
}

export async function pushDeltaScoresToCloud() {
    if (!sessionState.isCloudConnected) return { ok: true, offline: true };
    const { delta, changedCount } = computeScoreDelta(window.studentScores);
    if (changedCount === 0) {
        return { ok: true, noChange: true, message: 'Không có ô điểm nào thay đổi so với Cloud.' };
    }

    const gasUrl = getGasWebAppUrl();
    const res = await callGasApi(gasUrl, {
        action: 'merge_scores',
        sessionId: sessionState.sessionId,
        pin: sessionState.pin,
        deltaScores: delta
    });
    if (res.ok && res.studentScores && typeof res.studentScores === 'object') {
        applyScoresToUI(res.studentScores);
    }
    return res;
}

// Tự động lưu vùng Hội đồng sau mỗi thao tác (debounce 2.5s) & Tự động lưu + Khóa HĐ khi đóng trang thực sự
export function setupAutoSaveAndLockOnClose(onStatusChange) {
    let autoSaveTimer = null;
    let isSavingInFlight = false;
    let hasPendingAfterSave = false;
    let isClosingSession = false;
    let lockedDueToPageHide = false;

    cancelPendingAutoSaveFn = () => {
        clearTimeout(autoSaveTimer);
        autoSaveTimer = null;
        hasPendingAfterSave = false;
    };

    const setupFieldIds = ['khoa-select', 'nganh-select', 'dot-thang', 'nam-cham', 'so-qd', 'ngay-ky-qd'];

    const isTrackedTarget = (target) => {
        if (!target || typeof target.closest !== 'function') return false;
        return (
            setupFieldIds.includes(target.id) ||
            Boolean(target.closest('.box-hoidong')) ||
            Boolean(target.closest('#pool-sv')) ||
            Boolean(target.closest('#pool-gv')) ||
            Boolean(target.closest('.setup-section'))
        );
    };

    const runSaveCycle = async () => {
        if (isClosingSession || isRestoringSessionDOM) return;
        if (!sessionState.isCloudConnected || !sessionState.canEditHD || sessionState.isHDLocked) return;

        if (isSavingInFlight) {
            hasPendingAfterSave = true;
            return;
        }

        isSavingInFlight = true;
        hasPendingAfterSave = false;

        try {
            const res = await saveCloudHoiDong(
                getGasWebAppUrl(),
                sessionState.sessionId,
                sessionState.pin,
                sessionState.otp
            );
            if (!isClosingSession && typeof onStatusChange === 'function') {
                onStatusChange(
                    res.ok
                        ? `☁️ Đã tự động lưu lúc ${new Date().toLocaleTimeString('vi-VN')}`
                        : `⚠️ ${res.message}`
                );
            }
        } catch (err) {
            if (!isClosingSession && typeof onStatusChange === 'function') {
                onStatusChange(`⚠️ Lỗi tự động lưu: ${err.message}`);
            }
        } finally {
            isSavingInFlight = false;
            if (!isClosingSession && hasPendingAfterSave) {
                hasPendingAfterSave = false;
                triggerAutoSaveHD();
            }
        }
    };

    const triggerAutoSaveHD = () => {
        if (isClosingSession || isRestoringSessionDOM) return;
        if (!sessionState.isCloudConnected || !sessionState.canEditHD || sessionState.isHDLocked) return;

        clearTimeout(autoSaveTimer);
        if (typeof onStatusChange === 'function') {
            onStatusChange('⏳ Đang chờ lưu thay đổi...');
        }

        autoSaveTimer = setTimeout(() => {
            runSaveCycle();
        }, 2500);
    };

    window.triggerCloudSessionAutoSave = triggerAutoSaveHD;

    // 1. Theo dõi thêm/bớt/kéo-thả thẻ trong Khu vực Hội đồng, Kho SV và Kho GV
    const observer = new MutationObserver((mutations) => {
        if (isClosingSession || isRestoringSessionDOM) return;
        const hasMeaningfulChange = mutations.some(m => m.addedNodes.length > 0 || m.removedNodes.length > 0);
        if (hasMeaningfulChange) {
            triggerAutoSaveHD();
        }
    });

    ['workspace-hd', 'pool-sv', 'pool-gv'].forEach(id => {
        const el = document.getElementById(id);
        if (el) observer.observe(el, { childList: true, subtree: true });
    });

    // 2. Theo dõi mọi thay đổi giá trị ở Thông tin chung & các ô trong Hội đồng
    document.addEventListener('change', (e) => {
        if (isTrackedTarget(e.target)) triggerAutoSaveHD();
    });

    document.addEventListener('input', (e) => {
        if (isTrackedTarget(e.target)) triggerAutoSaveHD();
    });

    // 3. Khi tắt tab / đóng trình duyệt: Lưu bản mới nhất & khóa chỉnh sửa HĐ
    const lockSessionOnClose = () => {
        if (isClosingSession) return;
        if (!sessionState.isCloudConnected || !sessionState.canEditHD || sessionState.isHDLocked) return;

        const gasUrl = getGasWebAppUrl();
        if (!gasUrl) return;

        isClosingSession = true;
        lockedDueToPageHide = true;
        clearTimeout(autoSaveTimer);
        sessionState.isHDLocked = true;

        const { summary, payload } = serializeCurrentSession();
        const bodyStr = JSON.stringify({
            action: 'save_and_lock_hd',
            sessionId: sessionState.sessionId,
            pin: sessionState.pin,
            otp: sessionState.otp,
            summary,
            payload
        });

        try {
            fetch(gasUrl, {
                method: 'POST',
                keepalive: true,
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: bodyStr
            }).catch(() => {});
        } catch {
            if (navigator.sendBeacon) {
                const blob = new Blob([bodyStr], { type: 'text/plain;charset=utf-8' });
                navigator.sendBeacon(gasUrl, blob);
            }
        }
    };

    // Nếu trình duyệt chỉ tạm ẩn trang (bfcache / mobile chuyển app) rồi quay lại thì cập nhật lại UI
    const handlePageShow = () => {
        if (lockedDueToPageHide) {
            isClosingSession = false;
            lockedDueToPageHide = false;
            applySessionPermissionsToUI();
        }
    };

    window.addEventListener('pagehide', lockSessionOnClose);
    window.addEventListener('pageshow', handlePageShow);
}
