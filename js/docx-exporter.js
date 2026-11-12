// js/docx-exporter.js
import { templates } from './templates.js';

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

    // Tạo các thư mục Hồ sơ Word cho từng Sinh viên
    exportData.forEach(data => {
        const svFolder = zip.folder(`${data.ma_sv} - ${data.ho_ten}`);
        
        const baseContext = {
            KHOA: data.khoa || '',
            SO_HOI_DONG: data.so_hd || '',
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
                const outBlob = doc.getZip().generate({ type: "blob" });
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
