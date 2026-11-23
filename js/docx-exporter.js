// js/docx-exporter.js
import { templates } from './templates.js';

// Helper: Làm sạch các ký tự không hợp lệ trong tên thư mục (\ / : * ? " < > |)
function sanitizeFolderName(name, fallback = 'Thu_Muc') {
    const cleaned = String(name || '')
        .replace(/[\\/:*?"<>|]/g, '-')
        .replace(/\s+/g, ' ')
        .trim();
    return cleaned || fallback;
}

function escapeXml(str) {
    return String(str ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

function extractOnlyHoiDongNumber(rawSoHD) {
    const s = String(rawSoHD || '').trim();
    if (!s) return '';
    const m = s.match(/(?:hội\s*đồng|hoi\s*dong|hđ|hd)?\s*(?:số|so|#|:|-|\.)?\s*(\d+[a-zA-Z0-9\-]*)/i);
    return m && m[1] ? m[1].trim() : s;
}

/**
 * Chèn dòng đánh dấu lên đầu phần thân văn bản (word/document.xml) của file .docx
 */
function prependMarkerLineToDocxZip(zipInstance, markerText) {
    const docFile = zipInstance.file('word/document.xml');
    if (!docFile) return;

    const xmlContent = docFile.asText();
    const markerParagraphXml =
        `<w:p>` +
            `<w:pPr>` +
                `<w:jc w:val="right"/>` +
                `<w:spacing w:before="0" w:after="60" w:line="240" w:lineRule="auto"/>` +
            `</w:pPr>` +
            `<w:r>` +
                `<w:rPr>` +
                    `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>` +
                    `<w:i/>` +
                    `<w:color w:val="595959"/>` +
                    `<w:sz w:val="20"/>` +
                    `<w:szCs w:val="20"/>` +
                `</w:rPr>` +
                `<w:t xml:space="preserve">${escapeXml(markerText)}</w:t>` +
            `</w:r>` +
        `</w:p>`;

    const updatedXml = xmlContent.replace(/(<w:body[^>]*>)/i, `$1${markerParagraphXml}`);
    zipInstance.file('word/document.xml', updatedXml);
}

export async function exportWordZIP(exportData, meta = {}) {
    const zip = new JSZip();

    // Custom parser giúp nhận diện chính xác các placeholder có dấu cách hoặc bị xuống dòng trong Word
    const customParser = (tag) => {
        const cleanTag = tag.replace(/\s+/g, ' ').trim();
        return {
            get: (scope) => {
                if (cleanTag === '.') return scope;
                return scope[cleanTag] !== undefined ? scope[cleanTag] : '';
            }
        };
    };

    // Bộ đếm STT sinh viên theo từng Hội đồng (nếu dữ liệu đầu vào chưa có sẵn trường stt)
    const sttCounterByHD = new Map();

    // Tạo cây thư mục: [Tên Hội Đồng] / [MSSV - Họ Tên SV] / [6 file Word]
    exportData.forEach(data => {
        const hdFolderName = sanitizeFolderName(data.so_hd, 'Hoi_Dong');
        const hdFolder = zip.folder(hdFolderName);

        const svFolderName = sanitizeFolderName(`${data.ma_sv} - ${data.ho_ten}`, 'Sinh_Vien');
        const svFolder = hdFolder.folder(svFolderName);

        const hdNumOnly = extractOnlyHoiDongNumber(data.so_hd);
        const hdGroupKey = hdNumOnly || hdFolderName;
        const nextStt = (sttCounterByHD.get(hdGroupKey) || 0) + 1;
        sttCounterByHD.set(hdGroupKey, nextStt);

        const sttTrongHD = data.stt_hd ?? data.stt ?? nextStt;
        
        const topMarkerLine = `HĐ số ${hdNumOnly || data.so_hd || ''} - STT ${sttTrongHD}`;

        const baseContext = {
            KHOA: data.khoa || '',
            SO_HOI_DONG: data.so_hd || '',
            STT_HOI_DONG: String(sttTrongHD),
            NGÀY__BẢO_VỆ: data.ngay || '',
            NGÀY_BẢO_VỆ: data.ngay || '',
            GIỜ: data.gio || '',
            ĐỊA_ĐIỂM: `Phòng ${data.phong || ''}, cơ sở ${data.dia_diem || ''}`,
            Chủ_tịch: data.chu_tich || '',
            Phản_biện: data.phan_bien || '',
            Thư_ký: data.thu_ky || '',
            GVHD: data.gvhd || '',
            HỌ_VÀ_TÊN: data.ho_ten || '',
            MÃ_SỐ_SV: data.ma_sv || '',
            LỚP: data.lop || '',
            TÊN_ĐỀ_TÀI_ĐỒ_ÁN_TỐT_NGHIỆP: data.de_tai || '',
            'Ghi tên thành viên HĐ tại đây': ''
        };

        const renderWord = (templateBinary, fileName, extraContext = {}) => {
            if (!templateBinary) return;
            try {
                const doc = new window.docxtemplater(new PizZip(templateBinary), {
                    paragraphLoop: true,
                    linebreaks: true,
                    delimiters: { start: '{{', end: '}}' },
                    parser: customParser
                });

                doc.render({ ...baseContext, ...extraContext });

                // Chèn dòng đánh dấu lên đầu trang của file Word xuất ra
                const renderedZip = doc.getZip();
                prependMarkerLineToDocxZip(renderedZip, topMarkerLine);

                const outBlob = renderedZip.generate({ type: "blob" });
                svFolder.file(fileName, outBlob);
            } catch (error) {
                console.error(`Lỗi render file ${fileName}:`, error);
            }
        };

        // Mẫu 1, 2, 3
        renderWord(templates.bienBan, '1_BIEN_BAN_BAO_VE.docx');
        renderWord(templates.phieuGVHD, '2_Phieu_GVHD.docx');
        renderWord(templates.phieuGVPB, '3_Phieu_GVPB.docx');

        // Mẫu 4: Thay lần lượt {{Ghi tên thành viên HĐ tại đây}} bằng tên từng thành viên
        renderWord(templates.phieuHD, '4a_Phieu_HD_ChuTich.docx', {
            'Ghi tên thành viên HĐ tại đây': data.chu_tich || ''
        });
        renderWord(templates.phieuHD, '4b_Phieu_HD_PhanBien.docx', {
            'Ghi tên thành viên HĐ tại đây': data.phan_bien || ''
        });
        renderWord(templates.phieuHD, '4c_Phieu_HD_ThuKy.docx', {
            'Ghi tên thành viên HĐ tại đây': data.thu_ky || ''
        });
    });

    // Tải về file ZIP Hồ sơ Word
    const zipContent = await zip.generateAsync({ type: "blob" });
    saveAs(zipContent, "Ho_So_Hoi_Dong_DATN.zip");
}

// Xem trước Mẫu 4. Phiếu Hội đồng của phiên hiện tại
export function previewPhieuHDTemplate() {
    if (!templates.phieuHD) return alert("Chưa nạp được mẫu 4. Phiếu Hội đồng!");
    const blob = new Blob([templates.phieuHD], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });
    saveAs(blob, "4_Phieu_Hoi_Dong_Phien_Hien_Tai.docx");
}
