// js/excel-exporter.js

export async function exportExcelReport(exportData, meta = {}) {
    if (typeof ExcelJS === 'undefined') {
        alert("Thiếu thư viện ExcelJS từ lib/exceljs.min.js");
        return;
    }
    const workbook = await buildExcelReportWorkbook(exportData, meta);
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, "Bang_Diem_Tong_Hop_DATN.xlsx");
}

// 1. WORKBOOK BẢNG ĐIỂM TỔNG HỢP (TÁCH SHEET THEO BUỔI + PHÒNG + CƠ SỞ)
export async function buildExcelReportWorkbook(exportData, meta = {}) {
    const workbook = new ExcelJS.Workbook();
    const sheetsMap = {};

    const getBuoi = (timeStr) => {
        if (!timeStr) return "Sáng";
        const cleanTime = String(timeStr).toLowerCase().replace('h', ':').trim();
        const hour = parseInt(cleanTime.split(':')[0], 10);
        if (isNaN(hour)) return "Sáng";
        if (hour < 12) return "Sáng";
        if (hour < 17) return "Chiều";
        return "Tối";
    };

    exportData.forEach(item => {
        let nganh = meta.khoa || "CÔNG NGHỆ THỰC PHẨM";
        if (item.lop && item.lop.includes("CTP")) nganh = "CÔNG NGHỆ THỰC PHẨM";

        const buoi = getBuoi(item.gio);
        const phong = item.phong || 'Phong';
        const coSo = item.dia_diem || 'CoSo';
        
        const sheetNameClean = `${buoi}_P${grpKeyClean(phong)}_${grpKeyClean(coSo)}`.substring(0, 30);

        if (!sheetsMap[sheetNameClean]) {
            sheetsMap[sheetNameClean] = { sheetName: sheetNameClean, hoidongs: {} };
        }

        const hdKey = item.so_hd || 'HĐ';
        if (!sheetsMap[sheetNameClean].hoidongs[hdKey]) {
            sheetsMap[sheetNameClean].hoidongs[hdKey] = {
                nganh, soHD: item.so_hd, phong: item.phong, ngay: item.ngay, gio: item.gio,
                diaDiem: item.dia_diem, chuTich: item.chu_tich, phanBien: item.phan_bien, thuKy: item.thu_ky,
                students: []
            };
        }
        sheetsMap[sheetNameClean].hoidongs[hdKey].students.push(item);
    });

    const thinBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

    Object.keys(sheetsMap).forEach(sheetKey => {
        const sheetData = sheetsMap[sheetKey];
        const ws = workbook.addWorksheet(sheetData.sheetName, { views: [{ showGridLines: true }] });
        let currentRow = 1;

        Object.keys(sheetData.hoidongs).forEach((hdKey) => {
            const grp = sheetData.hoidongs[hdKey];

            // HEADER DÒNG 1 -> 5
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

            const dotThangStr = meta.dotThang || '6';
            const namChamStr = meta.namCham || '2026';
            const soQDStr = meta.soQD || '2071/QĐ-ĐHĐT';
            const ngayKyQDStr = meta.ngayKyQD || '5/8/2026';

            ws.getCell(`D${currentRow+2}`).value = `Đợt:            Tháng ${dotThangStr}`;
            ws.getCell(`D${currentRow+2}`).font = { name: 'Times New Roman', size: 11 };
            ws.getCell(`I${currentRow+2}`).value = `Năm : ${namChamStr}`;
            ws.getCell(`I${currentRow+2}`).font = { name: 'Times New Roman', size: 11, bold: true };

            ws.getCell(`D${currentRow+3}`).value = `Hội đồng bảo vệ kèm theo Quyết định số  ${soQDStr}`;
            ws.getCell(`D${currentRow+3}`).font = { name: 'Times New Roman', size: 11, bold: true };
            ws.getCell(`H${currentRow+3}`).value = `ký ngày    ${ngayKyQDStr}`;
            ws.getCell(`H${currentRow+3}`).font = { name: 'Times New Roman', size: 11 };

            ws.getCell(`D${currentRow+4}`).value = `Thời gian:  ${grp.gio || '17:00'} ngày ${grp.ngay || '11/11/2026'}`;
            ws.getCell(`D${currentRow+4}`).font = { name: 'Times New Roman', size: 11, bold: true };

            ws.getCell(`I${currentRow+4}`).value = `Địa điểm: Phòng ${grp.phong || '111'}, cơ sở ${grp.diaDiem || '120 Hoàng Minh Thảo'}`;
            ws.getCell(`I${currentRow+4}`).font = { name: 'Times New Roman', size: 11, bold: true };

            currentRow += 6;

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

            currentRow += 2;
            const startStudentRow = currentRow;

            grp.students.forEach((sv, svIdx) => {
                const startR = currentRow;
                const endR = startR + 2;

                ws.getCell(`A${startR}`).value = svIdx + 1;
                ws.getCell(`B${startR}`).value = sv.ma_sv;
                ws.getCell(`C${startR}`).value = sv.ho_ten;
                ws.getCell(`D${startR}`).value = sv.lop;
                ws.getCell(`E${startR}`).value = sv.gvhd;

                const diemGVHD = (sv.diem_gvhd !== undefined && sv.diem_gvhd !== '') ? Number(sv.diem_gvhd) : '';
                ws.getCell(`F${startR}`).value = diemGVHD;
                ws.getCell(`G${startR}`).value = (sv.diem_ct !== undefined && sv.diem_ct !== '') ? Number(sv.diem_ct) : '';
                ws.getCell(`H${startR}`).value = (sv.diem_pb !== undefined && sv.diem_pb !== '') ? Number(sv.diem_pb) : '';
                ws.getCell(`I${startR}`).value = '-';
                ws.getCell(`J${startR}`).value = (sv.diem_tk !== undefined && sv.diem_tk !== '') ? Number(sv.diem_tk) : '';

                ws.getCell(`K${startR}`).value = { formula: `IF(COUNT(G${startR}:J${startR})>0, AVERAGE(G${startR}:J${startR}), "")` };
                ws.getCell(`L${startR}`).value = { formula: `IF(AND(ISNUMBER(F${startR}), ISNUMBER(K${startR})), F${startR}*0.3 + K${startR}*0.7, "")` };
                ws.getCell(`M${startR}`).value = sv.diem_chu || '';

                ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'].forEach(col => {
                    ws.mergeCells(`${col}${startR}:${col}${endR}`);
                });

                for (let r = startR; r <= endR; r++) {
                    for (let col = 1; col <= 13; col++) {
                        const cell = ws.getCell(r, col);
                        cell.border = thinBorder;
                        cell.font = { name: 'Times New Roman', size: 10 };
                        if (col === 3 || col === 5) {
                            cell.alignment = { horizontal: 'left', vertical: 'middle' };
                            if (col === 3) cell.font = { name: 'Times New Roman', size: 10, bold: true };
                        } else {
                            cell.alignment = { horizontal: 'center', vertical: 'middle' };
                        }
                    }
                }
                currentRow += 3;
            });

            const endStudentRow = currentRow - 1;

            if (startStudentRow <= endStudentRow) {
                const totalRows = (endStudentRow - startStudentRow + 1);
                const step = Math.floor(totalRows / 3);

                const r1_start = startStudentRow;
                const r1_end = r1_start + step - 1;
                const r2_start = r1_end + 1;
                const r2_end = r2_start + step - 1;
                const r3_start = r2_end + 1;
                const r3_end = endStudentRow;

                ws.mergeCells(`N${r1_start}:N${r1_end}`);
                ws.getCell(`N${r1_start}`).value = `CT: ${grp.chuTich || ''}`;

                ws.mergeCells(`N${r2_start}:N${r2_end}`);
                ws.getCell(`N${r2_start}`).value = `UVPB: ${grp.phanBien || ''}`;

                ws.mergeCells(`N${r3_start}:N${r3_end}`);
                ws.getCell(`N${r3_start}`).value = `UVTK: ${grp.thuKy || ''}`;

                ws.mergeCells(`O${startStudentRow}:O${endStudentRow}`);
                ws.getCell(`O${startStudentRow}`).value = "";

                for (let r = startStudentRow; r <= endStudentRow; r++) {
                    const cellN = ws.getCell(`N${r}`);
                    cellN.border = thinBorder;
                    cellN.font = { name: 'Times New Roman', size: 10, bold: true };
                    cellN.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

                    const cellO = ws.getCell(`O${r}`);
                    cellO.border = thinBorder;
                }
            }

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
            ws.getCell(`E${currentRow}`).value = 'TRƯỜNG PHÒNG TỔNG HỢP';
            ws.getCell(`E${currentRow}`).font = { name: 'Times New Roman', size: 10, bold: true };
            ws.getCell(`E${currentRow}`).alignment = { horizontal: 'center' };

            ws.mergeCells(`L${currentRow}:O${currentRow}`);
            ws.getCell(`L${currentRow}`).value = 'HIỆU TRƯỜNG TRƯỜNG CN&KT';
            ws.getCell(`L${currentRow}`).font = { name: 'Times New Roman', size: 10, bold: true };
            ws.getCell(`L${currentRow}`).alignment = { horizontal: 'center' };

            currentRow += 6;
        });

        ws.getColumn(1).width = 5;
        ws.getColumn(2).width = 14;
        ws.getColumn(3).width = 24;
        ws.getColumn(4).width = 10;
        ws.getColumn(5).width = 24;
        ws.getColumn(6).width = 11;
        ws.getColumn(7).width = 5;
        ws.getColumn(8).width = 5;
        ws.getColumn(9).width = 5;
        ws.getColumn(10).width = 5;
        ws.getColumn(11).width = 8;
        ws.getColumn(12).width = 7;
        ws.getColumn(13).width = 12;
        ws.getColumn(14).width = 28;
        ws.getColumn(15).width = 15;
    });

    return workbook;
}

// 2. WORKBOOK DANH SÁCH HỘI ĐỒNG ĐÃ LẬP (KHỚP 100% MẪU GỒM CẢ GIẢNG VIÊN VÀ SINH VIÊN)
export async function buildDSDBWorkbook(exportData, meta = {}) {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("DS_Hoi_Dong_Da_Lap", { views: [{ showGridLines: true }] });

    const thinBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

    // 1. Trích xuất danh sách các Lớp duy nhất từ toàn bộ Sinh viên để điền vào mục KHÓA
    const uniqueClasses = [...new Set(
        exportData
            .map(item => (item.lop || '').trim())
            .filter(lop => lop.length > 0)
    )];
    const khoaListStr = uniqueClasses.length > 0 ? uniqueClasses.join(' + ') : 'K26 + K27';

    // 2. Gom nhóm dữ liệu theo từng Hội đồng
    const hdGroups = {};
    exportData.forEach(item => {
        const key = item.so_hd || 'Hội Đồng';
        if (!hdGroups[key]) {
            hdGroups[key] = {
                soHD: item.so_hd,
                phong: item.phong,
                ngay: item.ngay,
                gio: item.gio,
                diaDiem: item.dia_diem,
                chuTich: item.chu_tich,
                phanBien: item.phan_bien,
                thuKy: item.thu_ky,
                chuTichID: item.chu_tich_id || '',
                phanBienID: item.phan_bien_id || '',
                thuKyID: item.thu_ky_id || '',
                students: []
            };
        }
        hdGroups[key].students.push(item);
    });

    // 3. HEADER QUỐC HIỆU & TIÊU ĐỀ (Dòng 1 -> 5)
    ws.mergeCells('A1:C1');
    ws.getCell('A1').value = 'BỘ GIÁO DỤC & ĐÀO TẠO';
    ws.getCell('A1').font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells('E1:I1');
    ws.getCell('E1').value = 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM';
    ws.getCell('E1').font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell('E1').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells('A2:C2');
    ws.getCell('A2').value = 'ĐẠI HỌC DUY TÂN';
    ws.getCell('A2').font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell('A2').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells('E2:I2');
    ws.getCell('E2').value = 'Độc lập - Tự do - Hạnh phúc';
    ws.getCell('E2').font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell('E2').alignment = { horizontal: 'center', vertical: 'middle' };

    // Dòng 4: Tiêu đề chính với chữ KHÓA mặc định xuống dòng
    ws.mergeCells('A4:I4');
    const titleText = `DANH SÁCH HỘI ĐỒNG BẢO VỆ ĐỒ ÁN TỐT NGHIỆP\nKHÓA ${khoaListStr}`;
    const cellTitle = ws.getCell('A4');
    cellTitle.value = titleText;
    cellTitle.font = { name: 'Times New Roman', size: 12, bold: true };
    cellTitle.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    // Mặc định 2 dòng (height = 34), tự động tăng lên 48 nếu danh sách lớp ở dòng 2 quá dài
    ws.getRow(4).height = khoaListStr.length > 65 ? 48 : 34;

    // Dòng 5: Thời gian bảo vệ
    ws.mergeCells('A5:I5');
    const dotThangStr = meta.dotThang || '6';
    const namChamStr = meta.namCham || '2026';
    const cellSubTitle = ws.getCell('A5');
    cellSubTitle.value = `Thời gian bảo vệ: ngày .... tháng ${dotThangStr} năm ${namChamStr} tại Phòng ............. cơ sở ...................`;
    cellSubTitle.font = { name: 'Times New Roman', size: 11, bold: true, italic: true };
    cellSubTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(5).height = 20;

    // 4. HEADER BẢNG 9 CỘT CHUẨN (Dòng 8)
    const headers = [
        "HỘI\nĐỒNG", "MÃ SỐ\nGV", "HỌ VÀ TÊN", "NGÀY\nBẢO VỆ", "CHỨC DANH",
        "TÊN ĐỀ TÀI ĐỒ ÁN TỐT\nNGHIỆP", "MÃ SỐ SV", "HỌ VÀ TÊN", "LỚP"
    ];

    ws.getRow(8).height = 34;
    headers.forEach((h, idx) => {
        const cell = ws.getCell(8, idx + 1);
        cell.value = h;
        cell.font = { name: 'Times New Roman', size: 10, bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFFFF' } }; // Xanh ngọc nhạt như ảnh
        cell.border = thinBorder;
    });

    // 5. ĐỔ DỮ LIỆU TỪNG HỘI ĐỒNG (Bắt đầu từ Dòng 9)
    let rowIdx = 9;
    let hdIndex = 1;

    Object.keys(hdGroups).forEach(hdKey => {
        const hd = hdGroups[hdKey];
        const gvMembers = [
            { role: 'Chủ tịch', name: hd.chuTich, id: hd.chuTichID },
            { role: 'Phản biện', name: hd.phanBien, id: hd.phanBienID },
            { role: 'Thư ký', name: hd.thuKy, id: hd.thuKyID }
        ];

        const totalRows = Math.max(3, hd.students.length);
        const startR = rowIdx;
        const endR = startR + totalRows - 1;

        for (let offset = 0; offset < totalRows; offset++) {
            const r = startR + offset;

            // Cột B, C, E: Thông tin Giảng viên
            if (offset < 3) {
                const gv = gvMembers[offset];
                ws.getCell(`B${r}`).value = gv.id || '';
                ws.getCell(`C${r}`).value = gv.name || '';
                ws.getCell(`E${r}`).value = gv.role;
            }

            // Cột F, G, H, I: Thông tin Sinh viên
            if (offset < hd.students.length) {
                const sv = hd.students[offset];
                ws.getCell(`F${r}`).value = sv.de_tai || '';
                ws.getCell(`G${r}`).value = sv.ma_sv || '';
                ws.getCell(`H${r}`).value = sv.ho_ten || '';
                ws.getCell(`I${r}`).value = sv.lop || '';
            }

            // Định dạng viền và căn lề từng ô
            for (let c = 1; c <= 9; c++) {
                const cell = ws.getCell(r, c);
                cell.border = thinBorder;
                cell.font = { name: 'Times New Roman', size: 10 };
                if (c === 3 || c === 6 || c === 8) {
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                } else {
                    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                }
            }
            ws.getRow(r).height = 24;
        }

        // Merge Cột A (HỘI ĐỒNG)
        ws.mergeCells(`A${startR}:A${endR}`);
        const cellA = ws.getCell(`A${startR}`);
        cellA.value = hdIndex;
        cellA.font = { name: 'Times New Roman', size: 11, bold: true };
        cellA.alignment = { horizontal: 'center', vertical: 'middle' };

        // Merge Cột D (NGÀY BẢO VỆ)
        ws.mergeCells(`D${startR}:D${endR}`);
        const cellD = ws.getCell(`D${startR}`);
        const timeInfo = [hd.gio, hd.ngay].filter(Boolean).join(' ngày ');
        const placeInfo = [hd.phong ? `Phòng ${hd.phong}` : '', hd.diaDiem].filter(Boolean).join(', ');
        cellD.value = [timeInfo, placeInfo].filter(Boolean).join('\n');
        cellD.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

        rowIdx += totalRows;
        hdIndex++;
    });

    // 6. FOOTER CHỮ KÝ 3 CHỨC DANH
    rowIdx += 1;
    ws.mergeCells(`G${rowIdx}:I${rowIdx}`);
    const cellDate = ws.getCell(`G${rowIdx}`);
    cellDate.font = { name: 'Times New Roman', size: 10, italic: true };
    cellDate.alignment = { horizontal: 'center', vertical: 'middle' };

    rowIdx += 1;
    ws.mergeCells(`A${rowIdx}:C${rowIdx}`);
    const cellSignLeft = ws.getCell(`A${rowIdx}`);
    cellSignLeft.value = 'TRƯỞNG KHOA';
    cellSignLeft.font = { name: 'Times New Roman', size: 11, bold: true };
    cellSignLeft.alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells(`D${rowIdx}:F${rowIdx}`);
    const cellSignMid = ws.getCell(`D${rowIdx}`);
    cellSignMid.value = 'BAN THƯ KÝ';
    cellSignMid.font = { name: 'Times New Roman', size: 11, bold: true };
    cellSignMid.alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells(`G${rowIdx}:I${rowIdx}`);
    const cellSignRight = ws.getCell(`G${rowIdx}`);
    cellSignRight.value = 'CT. HỘI ĐỒNG THI & XÉT CNTN';
    cellSignRight.font = { name: 'Times New Roman', size: 11, bold: true };
    cellSignRight.alignment = { horizontal: 'center', vertical: 'middle' };

    // 7. ĐỘ RỘNG 9 CỘT
    ws.getColumn(1).width = 9;   // A: HỘI ĐỒNG
    ws.getColumn(2).width = 13;  // B: MÃ SỐ GV
    ws.getColumn(3).width = 23;  // C: HỌ VÀ TÊN GV
    ws.getColumn(4).width = 20;  // D: NGÀY BẢO VỆ
    ws.getColumn(5).width = 14;  // E: CHỨC DANH
    ws.getColumn(6).width = 36;  // F: TÊN ĐỀ TÀI ĐỒ ÁN TỐT NGHIỆP
    ws.getColumn(7).width = 14;  // G: MÃ SỐ SV
    ws.getColumn(8).width = 23;  // H: HỌ VÀ TÊN SV
    ws.getColumn(9).width = 13;  // I: LỚP

    return workbook;
}

function grpKeyClean(str) {
    return String(str || '').replace(/[^a-zA-Z0-9]/g, '');
}
