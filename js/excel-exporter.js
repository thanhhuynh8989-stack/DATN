// js/excel-exporter.js

export async function exportExcelReport(exportData) {
    if (typeof ExcelJS === 'undefined') {
        alert("Thiếu thư viện ExcelJS từ lib/exceljs.min.js");
        return;
    }

    const workbook = new ExcelJS.Workbook();
    const groups = {};

    // 1. Phân loại Sinh viên theo từng Hội đồng
    exportData.forEach(item => {
        let nganh = "CÔNG NGHỆ THỰC PHẨM";
        if (item.lop.includes("CTP")) nganh = "CÔNG NGHỆ THỰC PHẨM";
        else if (item.lop.includes("EKD") || item.lop.includes("KTD")) nganh = "KỸ THUẬT ĐIỆN";
        else if (item.lop.includes("EDK")) nganh = "KỸ THUẬT ĐIỀU KHIỂN & TỰ ĐỘNG HÓA";
        else if (item.lop.includes("EDT")) nganh = "ĐIỆN TỰ ĐỘNG";

        const groupKey = `${item.so_hd}_${grpKeyClean(item.phong)}_${grpKeyClean(item.ngay)}`;
        if (!groups[groupKey]) {
            groups[groupKey] = {
                nganh,
                soHD: item.so_hd,
                phong: item.phong,
                ngay: item.ngay,
                gio: item.gio,
                diaDiem: item.dia_diem,
                chuTich: item.chu_tich,
                phanBien: item.phan_bien,
                thuKy: item.thu_ky,
                students: []
            };
        }
        groups[groupKey].students.push(item);
    });

    const ws = workbook.addWorksheet("Bảng điểm HĐ IN", { views: [{ showGridLines: true }] });

    const thinBorder = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    let currentRow = 1;

    Object.keys(groups).forEach((key) => {
        const grp = groups[key];

        // --- PHẦN HEADER (Dòng 1 đến 5) ---
        ws.mergeCells(`A${currentRow}:C${currentRow}`);
        ws.getCell(`A${currentRow}`).value = 'BỘ GIÁO DỤC VÀ ĐÀO TẠO';
        ws.getCell(`A${currentRow}`).font = { name: 'Times New Roman', size: 10, bold: true };
        ws.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };

        ws.mergeCells(`A${currentRow+1}:C${currentRow+1}`);
        ws.getCell(`A${currentRow+1}`).value = 'ĐẠI HỌC DUY TÂN';
        ws.getCell(`A${currentRow+1}`).font = { name: 'Times New Roman', size: 10, bold: true };
        ws.getCell(`A${currentRow+1}`).alignment = { horizontal: 'center' };

        ws.mergeCells(`A${currentRow+2}:C${currentRow+2}`);
        ws.getCell(`A${currentRow+2}`).value = 'TRƯỜNG CÔNG NGHỆ VÀ KỸ THUẬT';
        ws.getCell(`A${currentRow+2}`).font = { name: 'Times New Roman', size: 10, bold: true };
        ws.getCell(`A${currentRow+2}`).alignment = { horizontal: 'center' };

        ws.mergeCells(`D${currentRow}:O${currentRow}`);
        ws.getCell(`D${currentRow}`).value = 'BẢNG ĐIỂM GHI KẾT QUẢ BẢO VỆ ĐỒ ÁN TỐT NGHIỆP';
        ws.getCell(`D${currentRow}`).font = { name: 'Times New Roman', size: 13, bold: true };
        ws.getCell(`D${currentRow}`).alignment = { horizontal: 'left', vertical: 'middle' };

        ws.getCell(`D${currentRow+1}`).value = `NGÀNH :   ${grp.nganh.toUpperCase()}`;
        ws.getCell(`D${currentRow+1}`).font = { name: 'Times New Roman', size: 11, bold: true };

        const dateParts = (grp.ngay || '').split('/');
        const monthStr = dateParts[1] || '6';
        const yearStr = dateParts[2] || '2026';

        ws.getCell(`D${currentRow+2}`).value = `Đợt:            Tháng ${monthStr}`;
        ws.getCell(`D${currentRow+2}`).font = { name: 'Times New Roman', size: 11 };
        ws.getCell(`F${currentRow+2}`).value = `Năm : ${yearStr}`;
        ws.getCell(`F${currentRow+2}`).font = { name: 'Times New Roman', size: 11, bold: true };

        // Dòng 4: Hội đồng bảo vệ kèm theo Quyết định số...
        ws.getCell(`D${currentRow+3}`).value = `Hội đồng bảo vệ kèm theo Quyết định số  2071/QĐ-ĐHĐT`;
        ws.getCell(`D${currentRow+3}`).font = { name: 'Times New Roman', size: 11, bold: true };
        ws.getCell(`F${currentRow+3}`).value = `ký ngày    5/8/2026`;
        ws.getCell(`F${currentRow+3}`).font = { name: 'Times New Roman', size: 11 };

        // Dòng 5: Thời gian & Địa điểm
        ws.getCell(`D${currentRow+4}`).value = `Thời gian:  ${grp.gio || '15h30'} ngày ${grp.ngay || '17/05/2026'}`;
        ws.getCell(`D${currentRow+4}`).font = { name: 'Times New Roman', size: 11, bold: true };

        ws.getCell(`F${currentRow+4}`).value = `Địa điểm: Phòng ${grp.phong || '403'}, cơ sở ${grp.diaDiem || 'K7/25 Quang Trung'}`;
        ws.getCell(`F${currentRow+4}`).font = { name: 'Times New Roman', size: 11, bold: true };

        currentRow += 6; // Đưa con trỏ tới tiêu đề bảng (Dòng 7)

        // --- PHẦN TIÊU ĐỀ BẢNG (Dòng 7 & 8) ---
        const hRow1 = currentRow;
        const hRow2 = currentRow + 1;

        ws.getCell(`A${hRow1}`).value = "STT";
        ws.getCell(`B${hRow1}`).value = "MSSV";
        ws.getCell(`C${hRow1}`).value = "HỌ VÀ TÊN SINH VIÊN";
        ws.getCell(`D${hRow1}`).value = "KHÓA";
        ws.getCell(`E${hRow1}`).value = "TÊN GVHD";
        ws.getCell(`F${hRow1}`).value = "ĐIỂM GVHD (30%)";
        ws.getCell(`G${hRow1}`).value = "ĐIỂM CỦA THÀNH VIÊN HĐ";
        ws.getCell(`L${hRow1}`).value = "ĐIỂM TỔNG KẾT 100%";
        ws.getCell(`N${hRow1}`).value = "TÊN THÀNH VIÊN HĐ THEO CHỨC DANH";
        ws.getCell(`O${hRow1}`).value = "THÀNH VIÊN HĐ KÝ TÊN";

        ws.getCell(`G${hRow2}`).value = "CT";
        ws.getCell(`H${hRow2}`).value = "PB";
        ws.getCell(`I${hRow2}`).value = "UV";
        ws.getCell(`J${hRow2}`).value = "TK";
        ws.getCell(`K${hRow2}`).value = "ĐHĐ 70%";
        ws.getCell(`L${hRow2}`).value = "SỐ";
        ws.getCell(`M${hRow2}`).value = "CHỮ";

        ws.mergeCells(`A${hRow1}:A${hRow2}`);
        ws.mergeCells(`B${hRow1}:B${hRow2}`);
        ws.mergeCells(`C${hRow1}:C${hRow2}`);
        ws.mergeCells(`D${hRow1}:D${hRow2}`);
        ws.mergeCells(`E${hRow1}:E${hRow2}`);
        ws.mergeCells(`F${hRow1}:F${hRow2}`);
        ws.mergeCells(`G${hRow1}:K${hRow1}`);
        ws.mergeCells(`L${hRow1}:M${hRow1}`);
        ws.mergeCells(`N${hRow1}:N${hRow2}`);
        ws.mergeCells(`O${hRow1}:O${hRow2}`);

        for (let r = hRow1; r <= hRow2; r++) {
            for (let col = 1; col <= 15; col++) {
                const cell = ws.getCell(r, col);
                cell.font = { name: 'Times New Roman', size: 9, bold: true };
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                cell.border = thinBorder;
            }
        }

        currentRow += 2; // Dòng sinh viên đầu tiên

        // --- PHẦN ĐỔ DỮ LIỆU SINH VIÊN (MỖI SV = 3 DÒNG) ---
        grp.students.forEach((sv, svIdx) => {
            const startR = currentRow;
            const endR = startR + 2;

            // Gán dữ liệu vào dòng đầu tiên của Sinh viên
            ws.getCell(`A${startR}`).value = svIdx + 1;
            ws.getCell(`B${startR}`).value = sv.ma_sv;
            ws.getCell(`C${startR}`).value = sv.ho_ten;
            ws.getCell(`D${startR}`).value = sv.lop;
            ws.getCell(`E${startR}`).value = sv.gvhd;

            // ĐIỂM
            const diemGVHD = sv.diem_gvhd !== undefined ? Number(sv.diem_gvhd) : null;
            ws.getCell(`F${startR}`).value = diemGVHD !== null ? diemGVHD : (sv.diem_gvhd || '');

            ws.getCell(`G${startR}`).value = sv.diem_ct !== undefined ? sv.diem_ct : '';
            ws.getCell(`H${startR}`).value = sv.diem_pb !== undefined ? sv.diem_pb : '';
            ws.getCell(`I${startR}`).value = '-';
            ws.getCell(`J${startR}`).value = sv.diem_tk !== undefined ? sv.diem_tk : '';

            // Công thức tính ĐHĐ (70%) và Điểm tổng kết
            ws.getCell(`K${startR}`).value = { formula: `AVERAGE(G${startR}:J${startR})` };
            ws.getCell(`L${startR}`).value = { formula: `F${startR}*0.3 + K${startR}*0.7` };
            ws.getCell(`M${startR}`).value = sv.diem_chu || '';

            // 3 dòng chi tiết Cột N (Chức danh HĐ)
            ws.getCell(`N${startR}`).value = `CT: ${grp.chuTich || ''}`;
            ws.getCell(`N${startR+1}`).value = `UVPB: ${grp.phanBien || ''}`;
            ws.getCell(`N${startR+2}`).value = `UVTK: ${grp.thuKy || ''}`;

            // Merge dọc 3 dòng cho các cột từ A -> M và cột O
            ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'O'].forEach(col => {
                ws.mergeCells(`${col}${startR}:${col}${endR}`);
            });

            // Format Border và Căn chỉnh
            for (let r = startR; r <= endR; r++) {
                for (let col = 1; col <= 15; col++) {
                    const cell = ws.getCell(r, col);
                    cell.border = thinBorder;
                    cell.font = { name: 'Times New Roman', size: 10 };
                    
                    if (col === 3 || col === 5 || col === 14) {
                        cell.alignment = { horizontal: 'left', vertical: 'middle' };
                        if (col === 3 || col === 14) cell.font = { name: 'Times New Roman', size: 10, bold: true };
                    } else {
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    }
                }
            }

            currentRow += 3; // Chuyển sang sinh viên tiếp theo
        });

        // --- PHẦN FOOTER CHỮ KÝ ---
        currentRow += 1;
        ws.mergeCells(`L${currentRow}:O${currentRow}`);
        ws.getCell(`L${currentRow}`).value = `Đà Nẵng, ngày ..... tháng .... năm .......`;
        ws.getCell(`L${currentRow}`).font = { name: 'Times New Roman', size: 10, italic: true };
        ws.getCell(`L${currentRow}`).alignment = { horizontal: 'center' };

        currentRow += 1;
        ws.mergeCells(`A${currentRow}:C${currentRow}`);
        ws.getCell(`A${currentRow}`).value = 'TRƯỞNG KHOA';
        ws.getCell(`A${currentRow}`).font = { name: 'Times New Roman', size: 10, bold: true };
        ws.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };

        ws.mergeCells(`E${currentRow}:H${currentRow}`);
        ws.getCell(`E${currentRow}`).value = 'TRƯỞNG PHÒNG TỔNG HỢP';
        ws.getCell(`E${currentRow}`).font = { name: 'Times New Roman', size: 10, bold: true };
        ws.getCell(`E${currentRow}`).alignment = { horizontal: 'center' };

        ws.mergeCells(`L${currentRow}:O${currentRow}`);
        ws.getCell(`L${currentRow}`).value = 'HIỆU TRƯỜNG TRƯỜNG CN&KT';
        ws.getCell(`L${currentRow}`).font = { name: 'Times New Roman', size: 10, bold: true };
        ws.getCell(`L${currentRow}`).alignment = { horizontal: 'center' };

        currentRow += 6;
    });

    // Cài đặt độ rộng cột chuẩn
    ws.getColumn(1).width = 5;   // STT
    ws.getColumn(2).width = 14;  // MSSV
    ws.getColumn(3).width = 24;  // HỌ VÀ TÊN SINH VIÊN
    ws.getColumn(4).width = 10;  // KHÓA
    ws.getColumn(5).width = 24;  // TÊN GVHD
    ws.getColumn(6).width = 11;  // ĐIỂM GVHD (30%)
    ws.getColumn(7).width = 5;   // CT
    ws.getColumn(8).width = 5;   // PB
    ws.getColumn(9).width = 5;   // UV
    ws.getColumn(10).width = 5;  // TK
    ws.getColumn(11).width = 8;  // ĐHĐ 70%
    ws.getColumn(12).width = 7;  // SỐ
    ws.getColumn(13).width = 12; // CHỮ
    ws.getColumn(14).width = 28; // TÊN THÀNH VIÊN HĐ THEO CHỨC DANH
    ws.getColumn(15).width = 15; // THÀNH VIÊN HĐ KÝ TÊN

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, "Bang_Diem_Tong_Hop_DATN.xlsx");
}

function grpKeyClean(str) {
    return String(str || '').replace(/[^a-zA-Z0-9]/g, '');
}
