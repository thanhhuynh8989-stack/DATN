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

// 1. TẠO WORKBOOK BẢNG ĐIỂM TỔNG HỢP (TÁCH SHEET THEO BUỔI + PHÒNG + CƠ SỞ)
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

    // Phân loại nhóm và tách Sheet
    exportData.forEach(item => {
        let nganh = meta.khoa || "CÔNG NGHỆ THỰC PHẨM";
        if (item.lop && item.lop.includes("CTP")) nganh = "CÔNG NGHỆ THỰC PHẨM";

        const buoi = getBuoi(item.gio);
        const phong = item.phong || 'Phong';
        const coSo = item.dia_diem || 'CoSo';
        
        // Tên Sheet chuẩn tối đa 31 ký tự: "Sáng_P111_120HMT"
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

    // Duyệt vẽ từng Sheet
    Object.keys(sheetsMap).forEach(sheetKey => {
        const sheetData = sheetsMap[sheetKey];
        const ws = workbook.addWorksheet(sheetData.sheetName, { views: [{ showGridLines: true }] });
        let currentRow = 1;

        Object.keys(sheetData.hoidongs).forEach((hdKey) => {
            const grp = sheetData.hoidongs[hdKey];

            // --- HEADER DÒNG 1 -> 5 CHUẨN XÁC VỊ TRÍ TỪNG CỘT ---
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

            // Dòng 3: Đợt đặt ô D3, Năm đặt đúng ô I3
            ws.getCell(`D${currentRow+2}`).value = `Đợt:            Tháng ${dotThangStr}`;
            ws.getCell(`D${currentRow+2}`).font = { name: 'Times New Roman', size: 11 };
            ws.getCell(`I${currentRow+2}`).value = `Năm : ${namChamStr}`;
            ws.getCell(`I${currentRow+2}`).font = { name: 'Times New Roman', size: 11, bold: true };

            // Dòng 4: QĐ đặt ô D4, ký ngày đặt đúng ô H4
            ws.getCell(`D${currentRow+3}`).value = `Hội đồng bảo vệ kèm theo Quyết định số  ${soQDStr}`;
            ws.getCell(`D${currentRow+3}`).font = { name: 'Times New Roman', size: 11, bold: true };
            ws.getCell(`H${currentRow+3}`).value = `ký ngày    ${ngayKyQDStr}`;
            ws.getCell(`H${currentRow+3}`).font = { name: 'Times New Roman', size: 11 };

            // Dòng 5: Thời gian đặt ô D5, Địa điểm đặt đúng ô I5
            ws.getCell(`D${currentRow+4}`).value = `Thời gian:  ${grp.gio || '17:00'} ngày ${grp.ngay || '11/11/2026'}`;
            ws.getCell(`D${currentRow+4}`).font = { name: 'Times New Roman', size: 11, bold: true };

            ws.getCell(`I${currentRow+4}`).value = `Địa điểm: Phòng ${grp.phong || '111'}, cơ sở ${grp.diaDiem || '120 Hoàng Minh Thảo'}`;
            ws.getCell(`I${currentRow+4}`).font = { name: 'Times New Roman', size: 11, bold: true };

            currentRow += 6;

            // --- TIÊU ĐỀ BẢNG (Dòng 7 & 8) ---
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

            // --- ĐỔ DỮ LIỆU SINH VIÊN (3 DÒNG / SV) ---
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

// 2. TẠO WORKBOOK DANH SÁCH HỘI ĐỒNG ĐÃ LẬP (XUẤT CHUẨN MÃ GV VÀ ĐƠN VỊ CÔNG TÁC)
export async function buildDSDBWorkbook(exportData, meta = {}) {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("DS_Hoi_Dong_Da_Lap", { views: [{ showGridLines: true }] });

    const thinBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

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
                thuKyID: item.thu_ky_id || ''
            };
        }
    });

    ws.mergeCells('A1:C1'); ws.getCell('A1').value = 'BỘ GIÁO DỤC VÀ ĐÀO TẠO';
    ws.mergeCells('A2:C2'); ws.getCell('A2').value = 'ĐẠI HỌC DUY TÂN';
    ws.mergeCells('A3:C3'); ws.getCell('A3').value = (meta.khoa || 'KHOA MÔI TRƯỜNG & KHTN').toUpperCase();
    ['A1', 'A2', 'A3'].forEach(c => {
        ws.getCell(c).font = { name: 'Times New Roman', size: 10, bold: true };
        ws.getCell(c).alignment = { horizontal: 'center' };
    });

    ws.mergeCells('D1:G1');
    ws.getCell('D1').value = 'DANH SÁCH HỘI ĐỒNG BẢO VỆ ĐỒ ÁN TỐT NGHIỆP ĐÃ LẬP';
    ws.getCell('D1').font = { name: 'Times New Roman', size: 13, bold: true };
    ws.getCell('D1').alignment = { horizontal: 'left', vertical: 'middle' };

    ws.getCell('D2').value = `Đợt: Tháng ${meta.dotThang || '12'} / Năm ${meta.namCham || '2026'}`;
    ws.getCell('D2').font = { name: 'Times New Roman', size: 11, italic: true };

    ws.getCell('D3').value = `Quyết định số: ${meta.soQD || '2071/QĐ-ĐHĐT'} ký ngày ${meta.ngayKyQD || '5/12/2026'}`;
    ws.getCell('D3').font = { name: 'Times New Roman', size: 11, bold: true };

    const headers = ["STT", "HỘI ĐỒNG", "MÃ GV", "HỌ VÀ TÊN GIẢNG VIÊN", "ĐƠN VỊ CÔNG TÁC", "CHỨC DANH HỘI ĐỒNG", "THỜI GIAN & ĐỊA ĐIỂM BẢO VỆ"];
    headers.forEach((h, idx) => {
        const cell = ws.getCell(5, idx + 1);
        cell.value = h;
        cell.font = { name: 'Times New Roman', size: 10, bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6EE9C' } };
        cell.border = thinBorder;
    });

    let rowIdx = 6;
    let stt = 1;

    Object.keys(hdGroups).forEach(hdKey => {
        const hd = hdGroups[hdKey];
        const members = [
            { role: 'Chủ tịch HĐ', name: hd.chuTich, id: hd.chuTichID },
            { role: 'Ủy viên Phản biện', name: hd.phanBien, id: hd.phanBienID },
            { role: 'Ủy viên Thư ký', name: hd.thuKy, id: hd.thuKyID }
        ];

        const startR = rowIdx;
        members.forEach((m, mIdx) => {
            const r = rowIdx + mIdx;
            ws.getCell(`C${r}`).value = m.id || `GV_${stt}_${mIdx+1}`;
            ws.getCell(`D${r}`).value = m.name || '';
            ws.getCell(`E${r}`).value = 'Đại học Duy Tân';
            ws.getCell(`F${r}`).value = m.role;

            for (let c = 1; c <= 7; c++) {
                const cell = ws.getCell(r, c);
                cell.border = thinBorder;
                cell.font = { name: 'Times New Roman', size: 10 };
                if (c === 4 || c === 5) cell.alignment = { horizontal: 'left', vertical: 'middle' };
                else cell.alignment = { horizontal: 'center', vertical: 'middle' };
            }
        });

        const endR = startR + 2;
        ws.mergeCells(`A${startR}:A${endR}`); ws.getCell(`A${startR}`).value = stt;
        ws.mergeCells(`B${startR}:B${endR}`); ws.getCell(`B${startR}`).value = hd.soHD;
        ws.mergeCells(`G${startR}:G${endR}`); 
        ws.getCell(`G${startR}`).value = `${hd.gio || '17:00'} ngày ${hd.ngay || ''}\nPhòng ${hd.phong || ''}, ${hd.diaDiem || ''}`;
        ws.getCell(`G${startR}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

        rowIdx += 3;
        stt++;
    });

    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 15;
    ws.getColumn(3).width = 14;
    ws.getColumn(4).width = 25;
    ws.getColumn(5).width = 22;
    ws.getColumn(6).width = 20;
    ws.getColumn(7).width = 32;

    return workbook;
}

function grpKeyClean(str) {
    return String(str || '').replace(/[^a-zA-Z0-9]/g, '');
}