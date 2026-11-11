// Biến toàn cục lưu trữ nội dung file Word (dạng nhị phân arraybuffer)
const templates = {
    bienBan: null,
    phieuGVHD: null,
    phieuGVPB: null,
    phieuHoiDong: null
};

async function loadTemplatesSilently() {
    try {
        const fetchFile = async (url) => {
            const res = await fetch(url);
            return await res.arrayBuffer();
        };
        
        templates.bienBan = await fetchFile('templates/1_BIEN_BAN_BAO_VE_DATN.docx');
        templates.phieuGVHD = await fetchFile('templates/2_Phieu_GVHD.docx');
        templates.phieuGVPB = await fetchFile('templates/3_Phieu_GVPB.docx');
        templates.phieuHoiDong = await fetchFile('templates/4_Phieu_Hoi_Dong.docx');
        
        console.log("Đã tải xong 4 template Word ngầm.");
    } catch (error) {
        console.error("Lỗi khi tải template ngầm:", error);
    }
}

// Chạy ngay khi load trang
document.addEventListener("DOMContentLoaded", () => {
    loadTemplatesSilently();
});
