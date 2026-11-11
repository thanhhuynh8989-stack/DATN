// js/templates.js

export const templates = {
    bienBan: null,
    phieuGVHD: null,
    phieuGVPB: null,
    phieuHD: null
};

export async function loadTemplatesSilently() {
    const loadFile = (url) => new Promise((resolve, reject) => {
        if (typeof PizZipUtils === 'undefined') {
            return reject(new Error("PizZipUtils chưa được nạp từ lib/pizzip-utils.min.js"));
        }
        PizZipUtils.getBinaryContent(url, (err, data) => {
            if (err) reject(err); else resolve(data);
        });
    });

    try {
        templates.bienBan = await loadFile('./templates/1_BIEN_BAN_BAO_VE_DATN.docx');
        templates.phieuGVHD = await loadFile('./templates/2_Phieu_GVHD.docx');
        templates.phieuGVPB = await loadFile('./templates/3_Phieu_GVPB.docx');
        templates.phieuHD = await loadFile('./templates/4_Phieu_Hoi_Dong.docx');
        console.log("✅ [Module Templates] Đã tải xong 4 file Word template!");
        return true;
    } catch (e) {
        console.error("❌ [Module Templates] Lỗi tải template:", e);
        throw e;
    }
}
