// js/docx-exporter.js
import { templates } from './templates.js';

export async function exportWordZIP(exportData) {
    const zip = new JSZip();

    exportData.forEach(data => {
        const svFolder = zip.folder(`${data.ma_sv} - ${data.ho_ten}`);
        
        const context = {
            KHOA: data.khoa || '',
            SO_HOI_DONG: data.so_hd || '',
            NGÀY__BẢO_VỆ: data.ngay || '',
            GIỜ: data.gio || '',
            ĐỊA_ĐIỂM: `Phòng ${data.phong || ''}, cơ sở ${data.dia_diem || ''}`,
            Chủ_tịch: data.chu_tich || '',
            Phản_biện: data.phan_bien || '',
            Thư_ký: data.thu_ky || '',
            GVHD: data.gvhd || '',
            HỌ_VÀ_TÊN: data.ho_ten || '',
            MÃ_SỐ_SV: data.ma_sv || '',
            LỚP: data.lop || '',
            TÊN_ĐỀ_TÀI_ĐỒ_ÁN_TỐT_NGHIỆP: data.de_tai || ''
        };

        const renderWord = (templateBinary, fileName) => {
            if (!templateBinary) return;
            try {
                const doc = new window.docxtemplater(new PizZip(templateBinary), { 
                    paragraphLoop: true, 
                    linebreaks: true,
                    delimiters: { start: '{{', end: '}}' } // Đảm bảo nhận diện chuẩn {{ tag }}
                });
                
                doc.render(context);
                const outBlob = doc.getZip().generate({ type: "blob" });
                svFolder.file(fileName, outBlob);
            } catch (error) {
                console.error(`Lỗi render file ${fileName}:`, error);
                throw new Error(`File template ${fileName} bị lỗi cú pháp tag {{ }}. Vui lòng kiểm tra lại file Word!`);
            }
        };

        renderWord(templates.bienBan, '1_BIEN_BAN_BAO_VE.docx');
        renderWord(templates.phieuGVHD, '2_Phieu_GVHD.docx');
        renderWord(templates.phieuGVPB, '3_Phieu_GVPB.docx');
        renderWord(templates.phieuHD, '4a_Phieu_HD_ChuTich.docx');
        renderWord(templates.phieuHD, '4b_Phieu_HD_PhanBien.docx');
        renderWord(templates.phieuHD, '4c_Phieu_HD_ThuKy.docx');
    });

    const zipContent = await zip.generateAsync({ type: "blob" });
    saveAs(zipContent, "Ho_So_Tot_Nghiep_DATN.zip");
}
