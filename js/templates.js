// js/templates.js
import { getAdminBinaryFile } from './admin-config.js';

export const templates = {
    bienBan: null,
    phieuGVHD: null,
    phieuGVPB: null,
    phieuHD: null,
    originalPhieuHD: null, // Lưu bản gốc để khôi phục bất kỳ lúc nào trong phiên
    isPhieuHDModified: false
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
        templates.phieuGVHD = await loadTemplateWithAdminOverride('phieuGVHD', './templates/2_Phieu_GVHD.docx');
        templates.phieuGVPB = await loadTemplateWithAdminOverride('phieuGVPB', './templates/3_Phieu_GVPB.docx');
        const hdBinary = await loadTemplateWithAdminOverride('phieuHD', './templates/4_Phieu_Hoi_Dong.docx');

        templates.originalPhieuHD = hdBinary;
        templates.phieuHD = hdBinary.slice(0);
        templates.isPhieuHDModified = false;

        console.log("✅ [Module Templates] Đã tải xong 4 file Word template (đã đồng bộ cấu hình Admin)!");
        return true;
    } catch (e) {
        console.error("❌ [Module Templates] Lỗi tải template:", e);
        throw e;
    }
}

// Tìm bảng Mục II. Kết quả đánh giá trong word/document.xml
function findSection2Table(xmlDoc, wNS) {
    const tables = Array.from(xmlDoc.getElementsByTagNameNS(wNS, 'tbl'));
    for (const tbl of tables) {
        const text = tbl.textContent.toLowerCase();
        if (text.includes('nội dung') && text.includes('thang điểm')) {
            return tbl;
        }
    }
    return null;
}

// Helper: Ghi đè nội dung text sạch vào 1 ô <w:tc> (giữ nguyên định dạng font/căn lề của ô)
function setCellTextPreserveStyle(tcNode, newText, wNS, xmlDoc) {
    const pNodes = Array.from(tcNode.getElementsByTagNameNS(wNS, 'p'));
    if (pNodes.length === 0) return;

    const lines = String(newText ?? '').split('\n');

    // Đảm bảo đủ số thẻ <w:p> tương ứng số dòng xuống hàng
    while (pNodes.length < lines.length) {
        const clonedP = pNodes[0].cloneNode(true);
        tcNode.appendChild(clonedP);
        pNodes.push(clonedP);
    }
    while (pNodes.length > lines.length && pNodes.length > 1) {
        const lastP = pNodes.pop();
        tcNode.removeChild(lastP);
    }

    lines.forEach((lineText, idx) => {
        const p = pNodes[idx];
        let tNodes = Array.from(p.getElementsByTagNameNS(wNS, 't'));
        if (tNodes.length === 0) {
            const r = xmlDoc.createElementNS(wNS, 'w:r');
            const t = xmlDoc.createElementNS(wNS, 'w:t');
            r.appendChild(t);
            p.appendChild(r);
            tNodes = [t];
        }
        tNodes[0].setAttribute('xml:space', 'preserve');
        tNodes[0].textContent = lineText;
        for (let i = 1; i < tNodes.length; i++) {
            tNodes[i].textContent = '';
        }
    });
}

// 1. Trích xuất danh sách các hàng tiêu chí và Tổng điểm của bảng Mục II
export function extractPhieuHDCriteriaTable() {
    if (!templates.phieuHD) return { rows: [], totalScore: '10' };
    const zip = new PizZip(templates.phieuHD);
    const xmlStr = zip.file('word/document.xml').asText();
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');

    const wNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const targetTable = findSection2Table(xmlDoc, wNS);
    if (!targetTable) return { rows: [], totalScore: '10' };

    const trList = Array.from(targetTable.getElementsByTagNameNS(wNS, 'tr'));
    const rows = [];
    let totalScore = '10';

    trList.forEach((tr) => {
        const tcList = Array.from(tr.getElementsByTagNameNS(wNS, 'tc'));
        if (tcList.length === 0) return;

        const rowTextLower = tr.textContent.replace(/\s+/g, ' ').trim().toLowerCase();
        // Bỏ qua hàng tiêu đề cột (TT | Nội dung | Thang điểm | Điểm chấm)
        if (rowTextLower.includes('nội dung') && rowTextLower.includes('thang điểm')) return;

        // Hàng TỔNG ĐIỂM
        if (rowTextLower.includes('tổng điểm')) {
            if (tcList.length >= 2) {
                const val = tcList[1].textContent.trim();
                if (val) totalScore = val;
            }
            return;
        }

        // Hàng tiêu chí thông thường: [0: TT] | [1: Nội dung] | [2: Thang điểm] | [3: Điểm chấm]
        if (tcList.length >= 3) {
            const pNodes = Array.from(tcList[1].getElementsByTagNameNS(wNS, 'p'));
            const content = pNodes
                .map(p => Array.from(p.getElementsByTagNameNS(wNS, 't')).map(t => t.textContent || '').join(''))
                .join('\n')
                .trim();

            const score = tcList[2].textContent.trim();
            rows.push({ content, score });
        }
    });

    return { rows, totalScore };
}

// 2. Áp dụng danh sách hàng (thêm/xóa/sửa) và Tổng điểm mới vào file Word Mẫu 4
export function applyPhieuHDCriteriaTable(editedRows, totalScore) {
    if (!templates.phieuHD) return false;
    const zip = new PizZip(templates.phieuHD);
    const xmlStr = zip.file('word/document.xml').asText();
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');

    const wNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const targetTable = findSection2Table(xmlDoc, wNS);
    if (!targetTable) return false;

    const trList = Array.from(targetTable.getElementsByTagNameNS(wNS, 'tr'));
    const criteriaTrs = [];
    let totalTr = null;

    trList.forEach((tr) => {
        const tcList = Array.from(tr.getElementsByTagNameNS(wNS, 'tc'));
        if (tcList.length === 0) return;

        const rowTextLower = tr.textContent.replace(/\s+/g, ' ').trim().toLowerCase();
        if (rowTextLower.includes('nội dung') && rowTextLower.includes('thang điểm')) return;

        if (rowTextLower.includes('tổng điểm')) {
            totalTr = tr;
        } else if (tcList.length >= 3) {
            criteriaTrs.push(tr);
        }
    });

    if (criteriaTrs.length === 0) return false;

    // Dùng hàng tiêu chí đầu tiên làm khuôn mẫu (clone) khi chèn thêm hàng mới
    const templateTr = criteriaTrs[0].cloneNode(true);

    // Xóa toàn bộ các hàng tiêu chí cũ trong bảng
    criteriaTrs.forEach(tr => targetTable.removeChild(tr));

    // Chèn lại đúng các hàng người dùng đã chỉnh sửa/thêm mới (tự đánh số TT 1, 2, 3...)
    editedRows.forEach((rowObj, idx) => {
        const newTr = templateTr.cloneNode(true);
        const tcList = Array.from(newTr.getElementsByTagNameNS(wNS, 'tc'));

        setCellTextPreserveStyle(tcList[0], String(idx + 1), wNS, xmlDoc); // Cột TT
        setCellTextPreserveStyle(tcList[1], rowObj.content, wNS, xmlDoc);  // Cột Nội dung
        setCellTextPreserveStyle(tcList[2], rowObj.score, wNS, xmlDoc);    // Cột Thang điểm
        if (tcList[3]) setCellTextPreserveStyle(tcList[3], '', wNS, xmlDoc); // Cột Điểm chấm để trống

        if (totalTr) {
            targetTable.insertBefore(newTr, totalTr);
        } else {
            targetTable.appendChild(newTr);
        }
    });

    // Cập nhật ô TỔNG ĐIỂM
    if (totalTr) {
        const totalTcs = Array.from(totalTr.getElementsByTagNameNS(wNS, 'tc'));
        if (totalTcs.length >= 2) {
            setCellTextPreserveStyle(totalTcs[1], String(totalScore || '10'), wNS, xmlDoc);
        }
    }

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
