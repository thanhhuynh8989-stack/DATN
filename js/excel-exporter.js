// js/excel-exporter.js

export async function exportExcelReport(exportData) {
    if (typeof ExcelJS === 'undefined') return;

    const workbook = new ExcelJS.Workbook();
    const groups = {};

    exportData.forEach(item => {
        let nganh = "ĐỒ ÁN TỐT NGHIỆP";
        if (item.lop.includes("CTP")) nganh = "CÔNG NGHỆ THỰC PHẨM";
        else if (item.lop.includes("EKD") || item.lop.includes("KTD")) nganh = "KỸ THUẬT ĐIỆN";
        else if (item.lop.includes("EDK")) nganh = "KỸ THUẬT ĐIỀU KHIỂN & TỰ ĐỘNG HÓA";
        else if (item.lop.includes("EDT")) nganh = "ĐIỆN TỰ ĐỘNG";

        const groupKey = `${nganh}_${item.phong}_${item.ngay}`;
        if (!groups[groupKey]) {
            groups[groupKey] = { nganh, phong: item.phong, ngay: item.ngay, gio: item.gio, diaDiem: item.dia_diem, students: [] };
        }
        groups[groupKey].students.push(item);
    });

    Object.keys(groups).forEach((key, sheetIdx) => {
        const grp = groups[key];
        const ws = workbook.addWorksheet(`${grp.phong || 'Phong'}_${sheetIdx + 1}`.substring(0, 30));

        ws.mergeCells('A1:C1'); ws.getCell('A1').value = 'BỘ GIÁO DỤC VÀ ĐÀO TẠO';
        ws.mergeCells('A2:C2'); ws.getCell('A2').value = 'ĐẠI HỌC DUY TÂN';
        ws.mergeCells('A3:C3'); ws.getCell('A3').value = 'TRƯỜNG CÔNG NGHỆ VÀ KỸ THUẬT';

        ws.mergeCells('D1:N1'); ws.getCell('D1').value = 'BẢNG ĐIỂM GHI KẾT QUẢ BẢO VỆ ĐỒ ÁN TỐT NGHIỆP';
        ws.getCell('D1').font = { bold: true, size: 14 };

        ws.getCell('D2').value = `NGÀNH: ${grp.nganh}`;
        ws.getCell('D3').value = `Thời gian: ${grp.gio} ngày ${grp.ngay}`;
        ws.getCell('H3').value = `Địa điểm: Phòng ${grp.phong}, cơ sở ${grp.diaDiem}`;

        ws.addRow(["STT", "MSSV", "HỌ VÀ TÊN SINH VIÊN", "KHÓA", "TÊN GVHD", "ĐIỂM GVHD (30%)", "ĐIỂM CỦA THÀNH VIÊN HĐ", "", "", "", "ĐHĐ (70%)", "ĐIỂM TỔNG KẾT 100%", "", "TÊN THÀNH VIÊN HĐ THEO CHỨC DANH"]);
        ws.addRow(["", "", "", "", "", "", "CT", "PB", "UV", "TK", "", "SỐ", "CHỮ", ""]);

        ws.mergeCells('A7:A8'); ws.mergeCells('B7:B8'); ws.mergeCells('C7:C8');
        ws.mergeCells('D7:D8'); ws.mergeCells('E7:E8'); ws.mergeCells('F7:F8');
        ws.mergeCells('G7:J7'); ws.mergeCells('K7:K8'); ws.mergeCells('L7:M7'); ws.mergeCells('N7:N8');

        let currentRow = 9;
        grp.students.forEach((sv, svIdx) => {
            const startR = currentRow;

            ws.getCell(`A${startR}`).value = svIdx + 1;
            ws.getCell(`B${startR}`).value = sv.ma_sv;
            ws.getCell(`C${startR}`).value = sv.ho_ten;
            ws.getCell(`D${startR}`).value = sv.lop;
            ws.getCell(`E${startR}`).value = sv.gvhd;

            ws.getCell(`N${startR}`).value = `CT: ${sv.chu_tich}`;
            ws.getCell(`N${startR + 1}`).value = `UVPB: ${sv.phan_bien}`;
            ws.getCell(`N${startR + 2}`).value = `UVTK: ${sv.thu_ky}`;

            ws.mergeCells(`A${startR}:A${startR + 2}`);
            ws.mergeCells(`B${startR}:B${startR + 2}`);
            ws.mergeCells(`C${startR}:C${startR + 2}`);
            ws.mergeCells(`D${startR}:D${startR + 2}`);
            ws.mergeCells(`E${startR}:E${startR + 2}`);
            ws.mergeCells(`F${startR}:F${startR + 2}`);
            ws.mergeCells(`K${startR}:K${startR + 2}`);
            ws.mergeCells(`L${startR}:L${startR + 2}`);
            ws.mergeCells(`M${startR}:M${startR + 2}`);

            currentRow += 3;
        });

        ws.columns.forEach(col => col.width = 15);
        ws.getColumn(3).width = 25;
        ws.getColumn(5).width = 25;
        ws.getColumn(14).width = 32;
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, "Bang_Diem_Tong_Hop_DATN.xlsx");
}
