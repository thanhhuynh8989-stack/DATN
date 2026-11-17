// js/templates.js

export const templates = {
    bienBan: null,
    phieuGVHD: null,
    phieuGVPB: null,
    phieuHD: null,
    originalPhieuHD: null, // Lưu bản gốc để khôi phục bất kỳ lúc nào trong phiên
    isPhieuHDModified: false
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
        const hdBinary = await loadFile('./templates/4_Phieu_Hoi_Dong.docx');

        templates.originalPhieuHD = hdBinary;
        templates.phieuHD = hdBinary.slice(0);
        templates.isPhieuHDModified = false;

        console.log("✅ [Module Templates] Đã tải xong 4 file Word template!");
        return true;
    } catch (e) {
        console.error("❌ [Module Templates] Lỗi tải template:", e);
        throw e;
    }
}

// Trích xuất danh sách các đoạn văn bản (<w:p>) từ word/document.xml của Mẫu 4 để chỉnh sửa trên Web
export function extractPhieuHDParagraphs() {
    if (!templates.phieuHD) return [];
    const zip = new PizZip(templates.phieuHD);
    const xmlStr = zip.file('word/document.xml').asText();
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');

    const wNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS(wNS, 'p'));

    const items = [];
    paragraphs.forEach((pNode, index) => {
        const tNodes = Array.from(pNode.getElementsByTagNameNS(wNS, 't'));
        const fullText = tNodes.map(t => t.textContent || '').join('');
        // Chỉ lấy những đoạn có nội dung chữ để giao diện gọn gàng
        if (fullText.trim().length > 0) {
            const isInTable = Boolean(pNode.closest('tc') || pNode.parentNode?.localName === 'tc');
            items.push({
                pIndex: index,
                text: fullText,
                isInTable
            });
        }
    });

    return items;
}

// Áp dụng các đoạn văn bản đã chỉnh sửa vào bản sao tạm thời templates.phieuHD (giữ nguyên font/style)
export function applyPhieuHDParagraphEdits(editedMap) {
    if (!templates.phieuHD) return false;
    const zip = new PizZip(templates.phieuHD);
    const xmlStr = zip.file('word/document.xml').asText();
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');

    const wNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS(wNS, 'p'));

    Object.keys(editedMap).forEach(key => {
        const pIdx = Number(key);
        const newText = editedMap[key];
        const pNode = paragraphs[pIdx];
        if (!pNode) return;

        const tNodes = Array.from(pNode.getElementsByTagNameNS(wNS, 't'));
        if (tNodes.length > 0) {
            // Giữ nguyên format của run đầu tiên, gán toàn bộ chuỗi đã sửa và giữ khoảng trắng
            tNodes[0].setAttribute('xml:space', 'preserve');
            tNodes[0].textContent = newText;
            // Xóa nội dung các run phụ phía sau trong cùng đoạn để tránh lặp chữ
            for (let i = 1; i < tNodes.length; i++) {
                tNodes[i].textContent = '';
            }
        }
    });

    const serializer = new XMLSerializer();
    const updatedXml = serializer.serializeToString(xmlDoc);
    zip.file('word/document.xml', updatedXml);

    templates.phieuHD = zip.generate({ type: 'arraybuffer' });
    templates.isPhieuHDModified = true;
    return true;
}

// Khôi phục Mẫu 4 về nguyên bản ban đầu
export function resetPhieuHDToOriginal() {
    if (!templates.originalPhieuHD) return false;
    templates.phieuHD = templates.originalPhieuHD.slice(0);
    templates.isPhieuHDModified = false;
    return true;
}