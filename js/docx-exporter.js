// js/docx-exporter.js
import { templates } from './templates.js';
import { buildExcelReportWorkbook, buildDSDBWorkbook } from './excel-exporter.js';

export async function exportWordZIP(exportData, meta = {}) {
    const zip = new JSZip();

    // 1. Tạo các thư mục Hồ sơ Word cho từng Sinh viên
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
                    delimiters: { start: '{{', end: '}}' }
                });
                
                doc.render(context);
                const outBlob = doc.getZip().generate({ type: "blob" });
                svFolder.file(fileName, outBlob);
            } catch (error) {
                console.error(`Lỗi render file ${fileName}:`, error);
            }
        };

        renderWord(templates.bienBan, '1_BIEN_BAN_BAO_VE.docx');
        renderWord(templates.phieuGVHD, '2_Phieu_GVHD.docx');
        renderWord(templates.phieuGVPB, '3_Phieu_GVPB.docx');
        renderWord(templates.phieuHD, '4a_Phieu_HD_ChuTich.docx');
        renderWord(templates.phieuHD, '4b_Phieu_HD_PhanBien.docx');
        renderWord(templates.phieuHD, '4c_Phieu_HD_ThuKy.docx');
    });

    // 2. TỰ ĐỘNG ĐÓNG GÓI 2 FILE EXCEL VÀO TRỰC TIẾP THƯ MỤC GỐC CỦA FILE ZIP
    try {
        // File Excel 1: Bảng điểm tổng hợp HĐ
        const excelReportWB = await buildExcelReportWorkbook(exportData, meta);
        const bufferReport = await excelReportWB.xlsx.writeBuffer();
        zip.file("Bang_Diem_Tong_Hop_DATN.xlsx", bufferReport);

        // File Excel 2: Danh sách Hội đồng đã lập
        const excelDSDBWB = await buildDSDBWorkbook(exportData, meta);
        const bufferDSDB = await excelDSDBWB.xlsx.writeBuffer();
        zip.file("Danh_Sach_Hoi_Dong_Da_Lap.xlsx", bufferDSDB);
    } catch (e) {
        console.error("Lỗi đóng gói file Excel vào ZIP:", e);
    }

    // 3. Tải về file ZIP tổng hợp
    const zipContent = await zip.generateAsync({ type: "blob" });
    saveAs(zipContent, "Ho_So_Va_Bang_Diem_Hoi_Dong_DATN.zip");
}
