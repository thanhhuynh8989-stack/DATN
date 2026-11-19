// js/templates.js
import { getAdminBinaryFile } from './admin-config.js';

export const templates = {
    bienBan: null,
    phieuGVHD: null,
    phieuGVPB: null,
    phieuHD: null,
    originalPhieuGVHD: null,
    originalPhieuGVPB: null,
    originalPhieuHD: null,
    modifiedFlags: {
        phieuGVHD: false,
        phieuGVPB: false,
        phieuHD: false
    }
};

export async function loadTemplatesSilently() {
    const loadFileFromServer = (url) => new Promise((resolve, reject) => {
        if (typeof PizZipUtils === 'undefined') {
            return reject(new Error("PizZipUtils chưa được nạp từ lib/pizzip-utils.min.js"));
        }
        PizZipUtils.getBinaryContent(url, (err, data) => {
            if (err) reject(err); else resolve(data);
        });
    });

    const loadTemplateWithAdminOverride = async (adminKey, fallbackUrl) => {
        const customBinary = await getAdminBinaryFile(`tpl_${adminKey}`);
        if (customBinary) {
            return customBinary;
        }
        return await loadFileFromServer(fallbackUrl);
    };

    try {
        templates.bienBan = await loadTemplateWithAdminOverride('bienBan', './templates/1_BIEN_BAN_BAO_VE_DATN.docx');
        const gvhdBinary = await loadTemplateWithAdminOverride('phieuGVHD', './templates/2_Phieu_GVHD.docx');
        const gvpbBinary = await loadTemplateWithAdminOverride('phieuGVPB', './templates/3_Phieu_GVPB.docx');
        const hdBinary = await loadTemplateWithAdminOverride('phieuHD', './templates/4_Phieu_Hoi_Dong.docx');

        templates.originalPhieuGVHD = gvhdBinary;
        templates.phieuGVHD = gvhdBinary.slice(0);

        templates.originalPhieuGVPB = gvpbBinary;
        templates.phieuGVPB = gvpbBinary.slice(0);

        templates.originalPhieuHD = hdBinary;
        templates.phieuHD = hdBinary.slice(0);

        templates.modifiedFlags = {
            phieuGVHD: false,
            phieuGVPB: false,
            phieuHD: false
        };

        console.log("✅ [Module Templates] Đã tải xong 4 file Word template (đã đồng bộ cấu hình Admin)!");
        return true;
    } catch (e) {
        console.error("❌ [Module Templates] Lỗi tải template:", e);
        throw e;
    }
}

// Cập nhật file mẫu .docx mới cho 1 trong 3 loại phiếu chấm trong phiên làm việc hiện tại
export function updateScoringTemplate(key, arrayBuffer) {
    if (!['phieuGVHD', 'phieuGVPB', 'phieuHD'].includes(key) || !arrayBuffer) {
        return false;
    }
    try {
        // Kiểm tra cấu trúc ZIP/DOCX hợp lệ trước khi áp dụng
        new PizZip(arrayBuffer);
    } catch (e) {
        console.error("File .docx không hợp lệ:", e);
        return false;
    }

    templates[key] = arrayBuffer.slice(0);
    templates.modifiedFlags[key] = true;
    return true;
}

// Khôi phục toàn bộ 3 mẫu phiếu chấm về đúng nguyên bản ban đầu của phiên
export function resetAllScoringTemplates() {
    if (templates.originalPhieuGVHD) {
        templates.phieuGVHD = templates.originalPhieuGVHD.slice(0);
    }
    if (templates.originalPhieuGVPB) {
        templates.phieuGVPB = templates.originalPhieuGVPB.slice(0);
    }
    if (templates.originalPhieuHD) {
        templates.phieuHD = templates.originalPhieuHD.slice(0);
    }
    templates.modifiedFlags = {
        phieuGVHD: false,
        phieuGVPB: false,
        phieuHD: false
    };
    return true;
}
