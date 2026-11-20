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

// Chuẩn hóa điểm (chấp nhận dấu . hoặc , | trả về null nếu trống)
function parseScoreNum(rawVal) {
    const str = String(rawVal ?? '').trim().replace(',', '.');
    if (str === '') return null;
    const num = parseFloat(str);
    if (isNaN(num)) return null;
    const clamped = Math.min(10, Math.max(0, num));
    return Number(clamped.toFixed(1));
}

// Đọc điểm thập phân 1 chữ số thang 10 thành chữ tiếng Việt (VD: 8.6 -> "tám phẩy sáu", 7.0 -> "bảy phẩy không")
function scoreToVietnameseWords(scoreNum) {
    if (scoreNum === null || scoreNum === undefined || isNaN(Number(scoreNum))) return '';
    const digitWords = ['không', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín', 'mười'];
    const fixedStr = Number(scoreNum).toFixed(1);
    const [intPartStr, decPartStr] = fixedStr.split('.');
    const intPart = parseInt(intPartStr, 10);
    const decPart = parseInt(decPartStr, 10);

    const intWord = digitWords[intPart] ?? String(intPart);
    const decWord = digitWords[decPart] ?? String(decPart);
    return `${intWord} phẩy ${decWord}`;
}

// 1. WORKBOOK BẢNG ĐIỂM TỔNG HỢP (TÁCH SHEET THEO BUỔI + PHÒNG + CƠ SỞ)
export async function buildExcelReportWorkbook(exportData, meta = {}) {
    const workbook = new ExcelJS.Workbook();
    const sheetsMap = {};
    const DOTS = '..................................';

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
        const nganh = (item.nganh || meta.nganh || DOTS).toUpperCase();
        const buoi = getBuoi(item.gio);
        const phong = item.phong || 'Phong';
        const coSo = item.dia_diem || 'CoSo';
        
        const sheetNameClean = `${buoi}_P${grpKeyClean(phong)}_${grpKeyClean(coSo)}`.substring(0, 30);

        if (!sheetsMap[sheetNameClean]) {
            sheetsMap[sheetNameClean] = {
                sheetName: sheetNameClean,
                nganh,
                phong: item.phong,
                ngay: item.ngay,
                gio: item.gio,
                diaDiem: item.dia_diem,
                hoidongs: {}
            };
        }

        const hdKey = item.so_hd || 'HĐ';
        if (!sheetsMap[sheetNameClean].hoidongs[hdKey]) {
            sheetsMap[sheetNameClean].hoidongs[hdKey] = {
                soHD: item.so_hd,
                chuTich: item.chu_tich,
                phanBien: item.phan_bien,
                thuKy: item.thu_ky,
                students: []
            };
        }
        sheetsMap[sheetNameClean].hoidongs[hdKey].students.push(item);
    });

    const thinBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    const admCfg = window.appAdminConfig || {};
    const uniName = (admCfg.universityName || 'ĐẠI HỌC DUY TÂN').toUpperCase();
    const schoolName = (admCfg.schoolName || 'TRƯỜNG CÔNG NGHỆ VÀ KỸ THUẬT').toUpperCase();

    Object.keys(sheetsMap).forEach(sheetKey => {
        const sheetData = sheetsMap[sheetKey];
        const ws = workbook.addWorksheet(sheetData.sheetName, { views: [{ showGridLines: true }] });
        let currentRow = 1;

        // =========================================================================
        // 1. VẼ HEADER CHUNG DUY NHẤT 1 LẦN Ở ĐẦU SHEET (DÒNG 1 -> 5)
        // =========================================================================
        ws.mergeCells(`A${currentRow}:C${currentRow}`);
        ws.getCell(`A${currentRow}`).value = 'BỘ GIÁO DỤC VÀ ĐÀO TẠO';
        ws.getCell(`A${currentRow}`).font = { name: 'Times New Roman', size: 10, bold: true };
        ws.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };

        ws.mergeCells(`A${currentRow+1}:C${currentRow+1}`);
        ws.getCell(`A${currentRow+1}`).value = uniName;
        ws.getCell(`A${currentRow+1}`).font = { name: 'Times New Roman', size: 10, bold: true };
        ws.getCell(`A${currentRow+1}`).alignment = { horizontal: 'center' };

        ws.mergeCells(`A${currentRow+2}:C${currentRow+2}`);
        ws.getCell(`A${currentRow+2}`).value = schoolName;
        ws.getCell(`A${currentRow+2}`).font = { name: 'Times New Roman', size: 10, bold: true };
        ws.getCell(`A${currentRow+2}`).alignment = { horizontal: 'center' };

        ws.mergeCells(`D${currentRow}:O${currentRow}`);
        ws.getCell(`D${currentRow}`).value = 'BẢNG ĐIỂM GHI KẾT QUẢ BẢO VỆ ĐỒ ÁN TỐT NGHIỆP';
        ws.getCell(`D${currentRow}`).font = { name: 'Times New Roman', size: 13, bold: true };
        ws.getCell(`D${currentRow}`).alignment = { horizontal: 'left', vertical: 'middle' };

        ws.getCell(`D${currentRow+1}`).value = `NGÀNH :   ${sheetData.nganh}`;
        ws.getCell(`D${currentRow+1}`).font = { name: 'Times New Roman', size: 11, bold: true };

        const dotThangStr = (meta.dotThang || '').trim() || '......';
        const namChamStr = (meta.namCham || '').trim() || '........';
        const soQDStr = (meta.soQD || '').trim() || DOTS;
        const ngayKyQDStr = (meta.ngayKyQD || '').trim() || '................';

        ws.getCell(`D${currentRow+2}`).value = `Đợt:            Tháng ${dotThangStr}`;
        ws.getCell(`D${currentRow+2}`).font = { name: 'Times New Roman', size: 11 };
        ws.getCell(`F${currentRow+2}`).value = `Năm : ${namChamStr}`;
        ws.getCell(`F${currentRow+2}`).font = { name: 'Times New Roman', size: 11, bold: true };

        ws.getCell(`D${currentRow+3}`).value = `Hội đồng bảo vệ kèm theo Quyết định số  ${soQDStr}`;
        ws.getCell(`D${currentRow+3}`).font = { name: 'Times New Roman', size: 11, bold: true };
        ws.getCell(`L${currentRow+3}`).value = `ký ngày    ${ngayKyQDStr}`;
        ws.getCell(`L${currentRow+3}`).font = { name: 'Times New Roman', size: 11 };

       const parseMinutes = (rawTime) => {
            const str = String(rawTime || '').trim();
            if (!str) return null;
            const startPart = str.split(/[-–]/)[0].trim();
            const match = startPart.match(/(\d{1,2})\s*[:hH]\s*(\d{1,2})?/);
            if (!match) return null;
            const hh = parseInt(match[1], 10);
            const mm = parseInt(match[2] || '0', 10);
            if (isNaN(hh) || isNaN(mm) || hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
            return {
                totalMins: hh * 60 + mm,
                formatted: `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`
            };
        };

        const rawTimes = [
            ...Object.values(sheetData.hoidongs).flatMap(h => h.students.map(s => s.gio)),
            sheetData.gio
        ].filter(Boolean);

        const parsedTimes = rawTimes
            .map(parseMinutes)
            .filter(Boolean)
            .sort((a, b) => a.totalMins - b.totalMins);

        const gioStr = parsedTimes.length > 0
            ? parsedTimes[0].formatted
            : ((rawTimes[0] || '').trim() || '......');
        const ngayStr = (sheetData.ngay || '').trim() || '................';
        const phongStr = (sheetData.phong || '').trim() || '......';
        const diaDiemStr = (sheetData.diaDiem || '').trim() || DOTS;

        ws.getCell(`D${currentRow+4}`).value = `Thời gian:  ${gioStr} ngày ${ngayStr}`;
        ws.getCell(`D${currentRow+4}`).font = { name: 'Times New Roman', size: 11, bold: true };

        ws.getCell(`G${currentRow+4}`).value = `Địa điểm: Phòng ${phongStr}, cơ sở ${diaDiemStr}`;
        ws.getCell(`G${currentRow+4}`).font = { name: 'Times New Roman', size: 11, bold: true };

        currentRow += 6;

        // =========================================================================
        // 2. VẼ TIÊU ĐỀ CỘT BẢNG DUY NHẤT 1 LẦN (DÒNG 7 & 8)
        // =========================================================================
        const hRow1 = currentRow;
        const hRow2 = currentRow + 1;

        ws.getRow(hRow1).height = 25;
        ws.getRow(hRow2).height = 20;

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
        let globalSvIdx = 1;

        // =========================================================================
        // 3. GẮN LIÊN TIẾP DỮ LIỆU CÁC HỘI ĐỒNG TRONG CÙNG 1 SHEET
        // =========================================================================
        Object.keys(sheetData.hoidongs).forEach((hdKey) => {
            const grp = sheetData.hoidongs[hdKey];
            const startStudentRow = currentRow;

            grp.students.forEach((sv) => {
                const startR = currentRow;
                const endR = startR + 2;

                ws.getCell(`A${startR}`).value = globalSvIdx++;
                ws.getCell(`B${startR}`).value = sv.ma_sv;
                ws.getCell(`C${startR}`).value = sv.ho_ten;
                ws.getCell(`D${startR}`).value = sv.lop;
                ws.getCell(`E${startR}`).value = sv.gvhd;

                const diemGVHD = parseScoreNum(sv.diem_gvhd);
                const diemCT = parseScoreNum(sv.diem_ct);
                const diemPB = parseScoreNum(sv.diem_pb);
                const diemTK = parseScoreNum(sv.diem_tk);

                ws.getCell(`F${startR}`).value = diemGVHD !== null ? diemGVHD : '';
                ws.getCell(`G${startR}`).value = diemCT !== null ? diemCT : '';
                ws.getCell(`H${startR}`).value = diemPB !== null ? diemPB : '';
                ws.getCell(`I${startR}`).value = '-';
                ws.getCell(`J${startR}`).value = diemTK !== null ? diemTK : '';

                // Tính giá trị thực tế ĐHĐ 70% và Điểm Tổng Kết 100% để sinh chữ chính xác cho cột CHỮ (cột M)
                const hdScores = [diemCT, diemPB, diemTK].filter(v => v !== null);
                const avgHD = hdScores.length > 0
                    ? Number((hdScores.reduce((a, b) => a + b, 0) / hdScores.length).toFixed(1))
                    : null;
                const finalTongKet = (diemGVHD !== null && avgHD !== null)
                    ? Number((diemGVHD * 0.3 + avgHD * 0.7).toFixed(1))
                    : null;

                ws.getCell(`K${startR}`).value = {
                    formula: `IF(COUNT(G${startR}:J${startR})>0, ROUND(AVERAGE(G${startR}:J${startR}), 1), "")`,
                    result: avgHD !== null ? avgHD : undefined
                };
                ws.getCell(`L${startR}`).value = {
                    formula: `IF(AND(ISNUMBER(F${startR}), ISNUMBER(K${startR})), ROUND(F${startR}*0.3 + K${startR}*0.7, 1), "")`,
                    result: finalTongKet !== null ? finalTongKet : undefined
                };

                // Cột CHỮ (cột M): ghi ra chữ giá trị ở cột SỐ nếu cột SỐ có điểm, ngược lại để trống
                ws.getCell(`M${startR}`).value = finalTongKet !== null
                    ? scoreToVietnameseWords(finalTongKet)
                    : (sv.diem_chu || '');

                // Định dạng hiển thị 1 chữ số thập phân (0.0) cho các ô điểm số (F, G, H, J, K, L)
                ['F', 'G', 'H', 'J', 'K', 'L'].forEach(colLetter => {
                    ws.getCell(`${colLetter}${startR}`).numFmt = '0.0';
                });

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
                            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: col === 13 };
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
        });

        // =========================================================================
        // 4. VẼ FOOTER CHỮ KÝ DUY NHẤT 1 LẦN Ở CUỐI SHEET
        // =========================================================================
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
        ws.getCell(`L${currentRow}`).value = 'HIỆU TRƯỞNG TRƯỜNG CN&KT';
        ws.getCell(`L${currentRow}`).font = { name: 'Times New Roman', size: 10, bold: true };
        ws.getCell(`L${currentRow}`).alignment = { horizontal: 'center' };

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
        ws.getColumn(12).width = 8;
        ws.getColumn(13).width = 16;
        ws.getColumn(14).width = 28;
        ws.getColumn(15).width = 15;
    });

    return workbook;
}

// Helper: Gom nhóm các HĐ trùng nhau (cùng tên/số HĐ hoặc cùng bộ Thành viên & Chức danh), gộp thời gian và đánh lại số 1, 2, 3...
function consolidateAndRenumberHoidongs(exportData) {
    const groups = [];
    const keyToGroup = new Map();
    const renumberLogs = [];

    exportData.forEach(item => {
        const hdTitleKey = String(item.so_hd || '').trim().toUpperCase();
        const memKey = `${item.chu_tich_id || item.chu_tich}|${item.phan_bien_id || item.phan_bien}|${item.thu_ky_id || item.thu_ky}`;
        const groupKey = `${hdTitleKey}__${memKey}`;

        const timeInfo = [item.gio ? item.gio.replace(':', 'h') : '', item.ngay].filter(Boolean).join(' ngày ');
        const placeInfo = [item.phong ? `Phòng ${item.phong}` : '', item.dia_diem].filter(Boolean).join(', ');
        const fullTimeLabel = [timeInfo, placeInfo].filter(Boolean).join(' - ');

        if (!keyToGroup.has(groupKey)) {
            const newGroup = {
                originalTitle: item.so_hd,
                chuTich: item.chu_tich,
                phanBien: item.phan_bien,
                thuKy: item.thu_ky,
                chuTichID: item.chu_tich_id || '',
                phanBienID: item.phan_bien_id || '',
                thuKyID: item.thu_ky_id || '',
                chuTichDonVi: item.chu_tich_don_vi || 'Đại học Duy Tân',
                phanBienDonVi: item.phan_bien_don_vi || 'Đại học Duy Tân',
                thuKyDonVi: item.thu_ky_don_vi || 'Đại học Duy Tân',
                timeLabels: fullTimeLabel ? [fullTimeLabel] : [],
                students: []
            };
            keyToGroup.set(groupKey, newGroup);
            groups.push(newGroup);
        } else {
            const existing = keyToGroup.get(groupKey);
            if (fullTimeLabel && !existing.timeLabels.includes(fullTimeLabel)) {
                existing.timeLabels.push(fullTimeLabel);
            }
        }

        const grp = keyToGroup.get(groupKey);
        if (!item.isEmptyHD && (item.ma_sv || item.ho_ten)) {
            if (!grp.students.some(s => s.ma_sv === item.ma_sv)) {
                grp.students.push(item);
            }
        }
    });

    groups.forEach((grp, idx) => {
        const newNum = idx + 1;
        grp.finalNumber = newNum;
        grp.mergedTimeText = grp.timeLabels.join('\n');

        const oldNumMatch = String(grp.originalTitle || '').match(/\d+/);
        const oldNum = oldNumMatch ? oldNumMatch[0] : grp.originalTitle;
        if (String(oldNum) !== String(newNum) || grp.timeLabels.length > 1) {
            renumberLogs.push(
                `• "${grp.originalTitle}" → Đánh số lại thành Hội đồng ${newNum}` +
                (grp.timeLabels.length > 1 ? ` (Đã gộp chung ${grp.timeLabels.length} ca bảo vệ)` : '')
            );
        }
    });

    return { groups, renumberLogs };
}

// 2. WORKBOOK DANH SÁCH HỘI ĐỒNG ĐÃ LẬP (GỘP HĐ TRÙNG & ĐÁNH LẠI SỐ 1, 2, 3...)
export async function buildDSDBWorkbook(exportData, meta = {}) {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("DS_Hoi_Dong_Da_Lap", { views: [{ showGridLines: true }] });

    const thinBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    const admCfg = window.appAdminConfig || {};
    const uniName = (admCfg.universityName || 'ĐẠI HỌC DUY TÂN').toUpperCase();

    // 1. Trích xuất danh sách các Lớp duy nhất từ toàn bộ Sinh viên để điền vào mục KHÓA
    const uniqueClasses = [...new Set(
        exportData
            .map(item => (item.lop || '').trim())
            .filter(lop => lop.length > 0)
    )];
    const khoaListStr = uniqueClasses.length > 0 ? uniqueClasses.join(' + ') : '................';

    // 2. Gom nhóm & đánh lại số Hội đồng 1, 2, 3...
    const { groups, renumberLogs } = consolidateAndRenumberHoidongs(exportData);

    // 3. HEADER QUỐC HIỆU & TIÊU ĐỀ (Dòng 1 -> 5, giãn đều từ A đến J)
    ws.mergeCells('A1:C1');
    ws.getCell('A1').value = 'BỘ GIÁO DỤC & ĐÀO TẠO';
    ws.getCell('A1').font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells('F1:J1');
    ws.getCell('F1').value = 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM';
    ws.getCell('F1').font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell('F1').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells('A2:C2');
    ws.getCell('A2').value = uniName;
    ws.getCell('A2').font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell('A2').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells('F2:J2');
    ws.getCell('F2').value = 'Độc lập - Tự do - Hạnh phúc';
    ws.getCell('F2').font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell('F2').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells('A4:J4');
    const titleText = `DANH SÁCH HỘI ĐỒNG BẢO VỆ ĐỒ ÁN TỐT NGHIỆP\nKHÓA ${khoaListStr}`;
    const cellTitle = ws.getCell('A4');
    cellTitle.value = titleText;
    cellTitle.font = { name: 'Times New Roman', size: 12, bold: true };
    cellTitle.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    ws.getRow(4).height = khoaListStr.length > 75 ? 48 : 34;

    ws.mergeCells('A5:J5');
    const dotThangStr = meta.dotThang || '6';
    const namChamStr = meta.namCham || '2026';
    const cellSubTitle = ws.getCell('A5');
    cellSubTitle.value = `Thời gian bảo vệ: ngày .... tháng ${dotThangStr} năm ${namChamStr} tại Phòng ............. cơ sở ...................`;
    cellSubTitle.font = { name: 'Times New Roman', size: 11, bold: true, italic: true };
    cellSubTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(5).height = 20;

    // 4. HEADER BẢNG 10 CỘT CHUẨN (Dòng 8)
    const headers = [
        "HỘI\nĐỒNG", "MÃ SỐ\nGV", "HỌ VÀ TÊN", "ĐƠN VỊ\nCÔNG TÁC", "NGÀY\nBẢO VỆ", 
        "CHỨC DANH", "TÊN ĐỀ TÀI ĐỒ ÁN TỐT\nNGHIỆP", "MÃ SỐ SV", "HỌ VÀ TÊN", "LỚP"
    ];

    ws.getRow(8).height = 34;
    headers.forEach((h, idx) => {
        const cell = ws.getCell(8, idx + 1);
        cell.value = h;
        cell.font = { name: 'Times New Roman', size: 10, bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFFFF' } };
        cell.border = thinBorder;
    });

    // 5. ĐỔ DỮ LIỆU TỪNG HỘI ĐỒNG ĐÃ GỘP & ĐÁNH SỐ LẠI (Bắt đầu từ Dòng 9)
    let rowIdx = 9;

    groups.forEach(hd => {
        const gvMembers = [
            { role: 'Chủ tịch', name: hd.chuTich, id: hd.chuTichID, donVi: hd.chuTichDonVi },
            { role: 'Phản biện', name: hd.phanBien, id: hd.phanBienID, donVi: hd.phanBienDonVi },
            { role: 'Thư ký', name: hd.thuKy, id: hd.thuKyID, donVi: hd.thuKyDonVi }
        ];

        const totalRows = Math.max(3, hd.students.length);
        const startR = rowIdx;
        const endR = startR + totalRows - 1;

        for (let offset = 0; offset < totalRows; offset++) {
            const r = startR + offset;

            if (offset < 3) {
                const gv = gvMembers[offset];
                ws.getCell(`B${r}`).value = gv.id || '';
                ws.getCell(`C${r}`).value = gv.name || '';
                ws.getCell(`D${r}`).value = gv.donVi || 'Đại học Duy Tân';
                ws.getCell(`F${r}`).value = gv.role;
            }

            if (offset < hd.students.length) {
                const sv = hd.students[offset];
                ws.getCell(`G${r}`).value = sv.de_tai || '';
                ws.getCell(`H${r}`).value = sv.ma_sv || '';
                ws.getCell(`I${r}`).value = sv.ho_ten || '';
                ws.getCell(`J${r}`).value = sv.lop || '';
            }

            for (let c = 1; c <= 10; c++) {
                const cell = ws.getCell(r, c);
                cell.border = thinBorder;
                cell.font = { name: 'Times New Roman', size: 10 };
                if (c === 3 || c === 4 || c === 7 || c === 9) {
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                } else {
                    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                }
            }
            ws.getRow(r).height = 26;
        }

        // Merge Cột A (HỘI ĐỒNG - đánh số lại 1, 2, 3...)
        ws.mergeCells(`A${startR}:A${endR}`);
        const cellA = ws.getCell(`A${startR}`);
        cellA.value = hd.finalNumber;
        cellA.font = { name: 'Times New Roman', size: 11, bold: true };
        cellA.alignment = { horizontal: 'center', vertical: 'middle' };

        // Merge Cột E (NGÀY BẢO VỆ - đã gộp các ca bảo vệ nếu có)
        ws.mergeCells(`E${startR}:E${endR}`);
        const cellE = ws.getCell(`E${startR}`);
        cellE.value = hd.mergedTimeText;
        cellE.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

        rowIdx += totalRows;
    });

    // 6. FOOTER CHỮ KÝ 3 CHỨC DANH
    rowIdx += 1;
    ws.mergeCells(`H${rowIdx}:J${rowIdx}`);
    const cellDate = ws.getCell(`H${rowIdx}`);
    cellDate.font = { name: 'Times New Roman', size: 10, italic: true };
    cellDate.alignment = { horizontal: 'center', vertical: 'middle' };

    rowIdx += 1;
    ws.mergeCells(`A${rowIdx}:C${rowIdx}`);
    const cellSignLeft = ws.getCell(`A${rowIdx}`);
    cellSignLeft.value = 'TRƯỞNG KHOA';
    cellSignLeft.font = { name: 'Times New Roman', size: 11, bold: true };
    cellSignLeft.alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells(`D${rowIdx}:G${rowIdx}`);
    const cellSignMid = ws.getCell(`D${rowIdx}`);
    cellSignMid.value = 'BAN THƯ KÝ';
    cellSignMid.font = { name: 'Times New Roman', size: 11, bold: true };
    cellSignMid.alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells(`H${rowIdx}:J${rowIdx}`);
    const cellSignRight = ws.getCell(`H${rowIdx}`);
    cellSignRight.value = 'CT. HỘI ĐỒNG THI & XÉT CNTN';
    cellSignRight.font = { name: 'Times New Roman', size: 11, bold: true };
    cellSignRight.alignment = { horizontal: 'center', vertical: 'middle' };

    // 7. ĐỘ RỘNG 10 CỘT
    ws.getColumn(1).width = 9;
    ws.getColumn(2).width = 13;
    ws.getColumn(3).width = 23;
    ws.getColumn(4).width = 22;
    ws.getColumn(5).width = 24;
    ws.getColumn(6).width = 14;
    ws.getColumn(7).width = 36;
    ws.getColumn(8).width = 14;
    ws.getColumn(9).width = 23;
    ws.getColumn(10).width = 13;

    return { workbook, renumberLogs };
}

function grpKeyClean(str) {
    return String(str || '').replace(/[^a-zA-Z0-9]/g, '');
}

// Xuất độc lập File 2: Danh sách Hội đồng đã lập (kèm thông báo gộp & đánh lại số)
export async function exportDSHoiDongExcel(exportData, meta = {}) {
    if (typeof ExcelJS === 'undefined') return alert("Thiếu thư viện ExcelJS!");
    const { workbook, renumberLogs } = await buildDSDBWorkbook(exportData, meta);
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), "Danh_Sach_Hoi_Dong_Da_Lap.xlsx");

    if (renumberLogs.length > 0) {
        alert(`✅ Đã xuất Danh sách Hội đồng!\n\n📋 THÔNG BÁO ĐÁNH LẠI SỐ & GỘP HỘI ĐỒNG:\n${renumberLogs.join('\n')}`);
    }
}

// Xuất độc lập File 3: Danh sách Phân công Phản biện Tốt nghiệp
export async function exportPhanBienExcel(exportData, meta = {}) {
    if (typeof ExcelJS === 'undefined') return alert("Thiếu thư viện ExcelJS!");
    const workbook = await buildPhanBienWorkbook(exportData, meta);
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), "Danh_Sach_Phan_Cong_Phan_Bien.xlsx");
}

// 3. WORKBOOK DANH SÁCH PHÂN CÔNG PHẢN BIỆN TỐT NGHIỆP
export async function buildPhanBienWorkbook(exportData, meta = {}) {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("DS_Phan_Bien", { views: [{ showGridLines: true }] });

    const thinBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    const DOTS = '..................................';
    const admCfg = window.appAdminConfig || {};
    const uniName = (admCfg.universityName || 'ĐẠI HỌC DUY TÂN').toUpperCase();

    // --- HEADER DÒNG 1 -> 5 ---
    ws.mergeCells('A1:C1');
    ws.getCell('A1').value = 'BỘ GIÁO DỤC VÀ ĐÀO TẠO';
    ws.getCell('A1').font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells('E1:H1');
    ws.getCell('E1').value = 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM';
    ws.getCell('E1').font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell('E1').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells('A2:C2');
    ws.getCell('A2').value = uniName;
    ws.getCell('A2').font = { name: 'Times New Roman', size: 11, bold: true, underline: true };
    ws.getCell('A2').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells('E2:H2');
    ws.getCell('E2').value = 'Độc lập - Tự do - Hạnh phúc';
    ws.getCell('E2').font = { name: 'Times New Roman', size: 11, bold: true, underline: true };
    ws.getCell('E2').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells('C3:H3');
    ws.getCell('C3').value = 'DANH SÁCH PHÂN CÔNG PHẢN BIỆN TỐT NGHIỆP';
    ws.getCell('C3').font = { name: 'Times New Roman', size: 12, bold: true };
    ws.getCell('C3').alignment = { horizontal: 'center', vertical: 'middle' };

    const nganhRaw = ((meta.nganh || '').trim() || exportData.find(s => s.nganh)?.nganh || DOTS).toUpperCase();
    ws.mergeCells('C4:H4');
    ws.getCell('C4').value = `NGÀNH: ${nganhRaw}`;
    ws.getCell('C4').font = { name: 'Times New Roman', size: 12, bold: true };
    ws.getCell('C4').alignment = { horizontal: 'center', vertical: 'middle' };

    const soQD = (meta.soQD || '').trim() || DOTS;
    const ngayKy = (meta.ngayKyQD || '').trim() || DOTS;
    ws.mergeCells('B5:H5');
    ws.getCell('B5').value = `(Ban hành kèm theo Quyết định số : ${soQD} ngày ${ngayKy} của Giám đốc ${uniName})`;
    ws.getCell('B5').font = { name: 'Times New Roman', size: 11, italic: true };
    ws.getCell('B5').alignment = { horizontal: 'center', vertical: 'middle' };

    // --- HEADER BẢNG (DÒNG 7 - MERGE C7:D7 CHO HỌ VÀ TÊN) ---
    ws.getRow(7).height = 28;
    ws.getCell('A7').value = 'STT';
    ws.getCell('B7').value = 'MSSV';
    ws.mergeCells('C7:D7');
    ws.getCell('C7').value = 'HỌ VÀ TÊN';
    ws.getCell('E7').value = 'LỚP';
    ws.getCell('F7').value = 'TÊN ĐỀ TÀI';
    ws.getCell('G7').value = 'GV PHẢN BIỆN';
    ws.getCell('H7').value = 'GHI CHÚ';

    for (let c = 1; c <= 8; c++) {
        const cell = ws.getCell(7, c);
        cell.font = { name: 'Times New Roman', size: 11, bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = thinBorder;
    }

    // --- ĐỔ DỮ LIỆU SINH VIÊN & GV PHẢN BIỆN (BẮT ĐẦU TỪ DÒNG 8) ---
    let rowIdx = 8;
    exportData.forEach((sv, idx) => {
        const r = rowIdx;
        ws.getRow(r).height = 34;

        const fullName = (sv.ho_ten || '').trim();
        const nameParts = fullName.split(/\s+/);
        const firstName = nameParts.length > 1 ? nameParts.pop() : fullName;
        const lastName = nameParts.length > 0 ? nameParts.join(' ') : '';

        ws.getCell(`A${r}`).value = idx + 1;
        ws.getCell(`B${r}`).value = sv.ma_sv || '';
        ws.getCell(`C${r}`).value = lastName;
        ws.getCell(`D${r}`).value = firstName;
        ws.getCell(`E${r}`).value = sv.lop || '';
        ws.getCell(`F${r}`).value = sv.de_tai || '';
        ws.getCell(`G${r}`).value = sv.phan_bien || '';
        ws.getCell(`H${r}`).value = '';

        for (let c = 1; c <= 8; c++) {
            const cell = ws.getCell(r, c);
            cell.font = { name: 'Times New Roman', size: 11 };

            cell.border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                left: c === 4 ? undefined : { style: 'thin' },
                right: c === 3 ? undefined : { style: 'thin' }
            };

            if (c === 3 || c === 4 || c === 6 || c === 7) {
                cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
            } else {
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            }
        }
        rowIdx++;
    });

    // --- FOOTER CHỮ KÝ 4 CHỨC DANH ---
    const footerRow = rowIdx;
    ws.getRow(footerRow).height = 28;

    ws.mergeCells(`A${footerRow}:C${footerRow}`);
    ws.getCell(`A${footerRow}`).value = 'TRƯỞNG KHOA';
    ws.getCell(`A${footerRow}`).font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell(`A${footerRow}`).alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells(`D${footerRow}:F${footerRow}`);
    ws.getCell(`D${footerRow}`).value = 'HIỆU TRƯỞNG TRƯỜNG ĐÀO TẠO';
    ws.getCell(`D${footerRow}`).font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell(`D${footerRow}`).alignment = { horizontal: 'center', vertical: 'middle' };

    ws.getCell(`G${footerRow}`).value = 'BAN ĐÀO TẠO';
    ws.getCell(`G${footerRow}`).font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell(`G${footerRow}`).alignment = { horizontal: 'center', vertical: 'middle' };

    ws.getCell(`H${footerRow}`).value = 'GIÁM ĐỐC';
    ws.getCell(`H${footerRow}`).font = { name: 'Times New Roman', size: 11, bold: true };
    ws.getCell(`H${footerRow}`).alignment = { horizontal: 'center', vertical: 'middle' };

    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 16;
    ws.getColumn(3).width = 18;
    ws.getColumn(4).width = 10;
    ws.getColumn(5).width = 11;
    ws.getColumn(6).width = 45;
    ws.getColumn(7).width = 24;
    ws.getColumn(8).width = 14;

    return workbook;
}
