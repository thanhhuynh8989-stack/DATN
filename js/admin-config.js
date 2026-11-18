// js/admin-config.js

const CONFIG_STORAGE_KEY = 'hd_datn_admin_config_v1';
const DB_NAME = 'HD_DATN_Admin_Files_DB';
const DB_STORE = 'files';

export const DEFAULT_ADMIN_CONFIG = {
    password: 'admin123',
    universityName: 'ĐẠI HỌC DUY TÂN',
    schoolName: 'TRƯỜNG CÔNG NGHỆ & KỸ THUẬT',
    qdSuffix: '/QĐ-ĐHDT-HĐTN',
    campuses: [
        '120 Hoàng Minh Thảo',
        '254 Nguyễn Văn Linh',
        '137 Nguyễn Văn Linh',
        '03 Quang Trung',
        'K7/25 Quang Trung',
        'Phan Văn Trị'
    ]
};

// 1. Đọc & Lưu cấu hình Text (Header, Đuôi QĐ, Cơ sở, Mật khẩu) vào localStorage
export function getAdminConfig() {
    try {
        const raw = localStorage.getItem(CONFIG_STORAGE_KEY);
        if (!raw) return { ...DEFAULT_ADMIN_CONFIG, campuses: [...DEFAULT_ADMIN_CONFIG.campuses] };
        const parsed = JSON.parse(raw);
        return {
            ...DEFAULT_ADMIN_CONFIG,
            ...parsed,
            campuses: Array.isArray(parsed.campuses) && parsed.campuses.length > 0
                ? parsed.campuses
                : [...DEFAULT_ADMIN_CONFIG.campuses]
        };
    } catch {
        return { ...DEFAULT_ADMIN_CONFIG, campuses: [...DEFAULT_ADMIN_CONFIG.campuses] };
    }
}

export function saveAdminConfig(newConfig) {
    const merged = { ...getAdminConfig(), ...newConfig };
    localStorage.setItem(CONFIG_STORAGE_KEY, JSON.stringify(merged));
    window.appAdminConfig = merged;
    return merged;
}

export function resetAdminConfigText() {
    localStorage.removeItem(CONFIG_STORAGE_KEY);
    window.appAdminConfig = { ...DEFAULT_ADMIN_CONFIG, campuses: [...DEFAULT_ADMIN_CONFIG.campuses] };
    return window.appAdminConfig;
}

// 2. Lưu & Đọc File nhị phân (.xlsx, .docx) vào IndexedDB của trình duyệt
function openFilesDB() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(DB_NAME, 1);
        req.onupgradeneeded = (e) => {
            const db = e.target.result;
            if (!db.objectStoreNames.contains(DB_STORE)) {
                db.createObjectStore(DB_STORE);
            }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}

export async function saveAdminBinaryFile(key, arrayBuffer) {
    const db = await openFilesDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(DB_STORE, 'readwrite');
        tx.objectStore(DB_STORE).put(arrayBuffer, key);
        tx.oncomplete = () => resolve(true);
        tx.onerror = () => reject(tx.error);
    });
}

export async function getAdminBinaryFile(key) {
    try {
        const db = await openFilesDB();
        return new Promise((resolve) => {
            const tx = db.transaction(DB_STORE, 'readonly');
            const req = tx.objectStore(DB_STORE).get(key);
            req.onsuccess = () => resolve(req.result || null);
            req.onerror = () => resolve(null);
        });
    } catch {
        return null;
    }
}

export async function deleteAdminBinaryFile(key) {
    const db = await openFilesDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(DB_STORE, 'readwrite');
        tx.objectStore(DB_STORE).delete(key);
        tx.oncomplete = () => resolve(true);
        tx.onerror = () => reject(tx.error);
    });
}

// Khởi tạo biến toàn cục ngay khi import
window.appAdminConfig = getAdminConfig();
