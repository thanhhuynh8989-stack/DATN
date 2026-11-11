// ==========================================
// 1. BIẾN TOÀN CỤC & TẢI TEMPLATE MẶC ĐỊNH
// ==========================================
let templates = {};
let hoidongCount = 0;

document.addEventListener("DOMContentLoaded", async () => {
    // Tải ngầm 4 file Word templates từ thư mục relative
    const loadFile = (url) => new Promise((resolve, reject) => {
        if (typeof PizZipUtils === 'undefined') {
            return reject(new Error("PizZipUtils chưa được nạp. Kiểm tra file lib/pizzip-utils.min.js"));
        }
        PizZipUtils.getBinaryContent(url, (err, data) => {
            if (err) reject(err); else resolve(data);
        });
    });

    try {
        templates.bienBan = await loadFile('./templates/1_BIEN_BAN_BAO_VE_DATN.docx');
        templates.phieuGVHD = await loadFile('./templates/2_Phieu_GVHD.docx');
        templates.phieuGVPB = await loadFile('./templates/3_Phieu_GVPB.docx');
        templates.phieuHD = await loadFile('./templates/4_Phieu_Hoi_Dong.docx');
        console.log("✅ Load templates thành công!");
    } catch (e) {
        console.error("Lỗi tải template:", e);
        alert("Không tải được file template Word trong thư mục ./templates/. Vui lòng kiểm tra lại file trên GitHub!");
    }

    initGlobalSortables();
    bindEvents();
});

function bindEvents() {
    // Sự kiện khi bấm nút "⚡ NẠP DỮ LIỆU EXCEL"
    const btnRead = document.getElementById('btn-read-files');
    if (btnRead) {
        btnRead.addEventListener('click', async () => {
            const fileSV = document.getElementById('upload-sv').files[0];
            const fileGV = document.getElementById('upload-gv').files[0];
            const fileHD = document.getElementById('upload-hd').files[0];

            if (!fileSV && !fileGV && !fileHD) {
                alert("Vui lòng chọn ít nhất 1 file Excel để nạp dữ liệu!");
                return;
            }

            try {
                if (fileSV) {
                    const dataSV = await readExcelFile(fileSV);
                    renderSinhVienPool(dataSV);
                }

                if (fileGV) {
                    const dataGV = await readExcelFile(fileGV);
                    renderGiangVienPool(dataGV);
                }

                if (fileHD) {
                    const dataHD = await readExcelFile(fileHD);
                    renderHoidongFromExcel(dataHD);
                }

                alert("✅ Nạp dữ liệu Excel thành công!");
            } catch (err) {
                console.error("Lỗi khi đọc file Excel:", err);
                alert("Đã xảy ra lỗi khi đọc file Excel. Vui lòng kiểm tra lại định dạng file!");
            }
        });
    }

    // Nút Tạo Hội Đồng mới thủ công
    const btnAddHD = document.getElementById('btn-add-hd');
    if (btnAddHD) {
        btnAddHD.addEventListener('click', () => {
            hoidongCount++;
            createNewHoiDongBox(`Hội Đồng ${hoidongCount}`);
        });
    }

    // Nút Xuất Hồ Sơ & Bảng Điểm
    const btnExport = document.getElementById('btn-export');
    if (btnExport) {
        btnExport.addEventListener('click', handleExport);
    }
}

// ==========================================
// 2. ĐỌC FILE EXCEL & RENDER DỮ LIỆU
// ==========================================
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.SheetNames[0];
                const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1 });
                resolve(jsonData);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

// Render Thẻ Sinh Viên vào cột trái (tự động ghép Họ lót + Tên)
function renderSinhVienPool(rows) {
    const poolSV = document.getElementById('pool-sv');
    poolSV.innerHTML = ''; 

    // Tìm dòng chứa header để dò vị trí cột
    let headerIdx = -1;
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
        const rowStr = (rows[i] || []).join(' ').toUpperCase();
        if (rowStr.includes('MSSV') || rowStr.includes('MÃ SỐ SV')) {
            headerIdx = i;
            break;
        }
    }

    let startRow = headerIdx >= 0 ? headerIdx + 1 : 1;

    rows.forEach((row, idx) => {
        if (idx < startRow || !row || row.length === 0) return;

        // Bóc tách cột theo cấu trúc file Danh Sách SV
        const maSV = String(row[1] || '').trim();
        if (!maSV || maSV.toUpperCase().includes('MSSV')) return;

        const col2 = String(row[2] || '').trim();
        const col3 = String(row[3] || '').trim();
        
        let hoTen = '', lop = '', deTai = '', gvhd = '';

        // Trường hợp Họ và Tên bị tách làm 2 cột
        if (col3 && !col3.startsWith('K2') && !col3.startsWith('K1') && !col3.startsWith('K3')) {
            hoTen = `${col2} ${col3}`.trim();
            lop = String(row[4] || '').trim();
            deTai = String(row[5] || '').trim();
            gvhd = String(row[6] || '').trim();
        } else {
            hoTen = col2;
            lop = col3;
            deTai = String(row[4] || '').trim();
            gvhd = String(row[5] || '').trim();
        }

        const svCard = `
            <div class="card-item mb-2 p-2 border rounded bg-white" 
                 data-id="${maSV}" 
                 data-name="${hoTen}" 
                 data-lop="${lop}" 
                 data-detai="${deTai}" 
                 data-gvhd="${gvhd}">
                <div class="fw-bold text-primary">${hoTen} (${maSV})</div>
                <div class="text-muted small">Lớp: ${lop} | GVHD: ${gvhd}</div>
                <div class="text-dark small text-truncate" title="${deTai}">Đề tài: ${deTai}</div>
            </div>
        `;
        poolSV.insertAdjacentHTML('beforeend', svCard);
    });
}

// Render Thẻ Giảng Viên vào cột phải
function renderGiangVienPool(rows) {
    const poolGV = document.getElementById('pool-gv');
    poolGV.innerHTML = '';

    rows.forEach((row, idx) => {
        if (idx < 1 || !row || row.length === 0) return;
        const col0 = String(row[0] || '').trim();
        if (col0.toUpperCase().includes('STT') || col0.toUpperCase().includes('MÃ')) return;

        const maGV = col0 || `GV_${idx}`;
        const hoTenGV = String(row[1] || '').trim();
        const donVi = String(row[2] || 'Đại học Duy Tân').trim();

        if (!hoTenGV) return;

        const gvCard = `
            <div class="card-item bg-light border-info mb-2 p-2 rounded" data-id="${maGV}" data-name="${hoTenGV}">
                <div class="fw-bold text-dark">${hoTenGV}</div>
                <div class="text-muted small">${donVi} - Mã: ${maGV}</div>
            </div>
        `;
        poolGV.insertAdjacentHTML('beforeend', gvCard);
    });
}

// Render các Box Hội Đồng nếu tải file Hội Đồng sẵn
function renderHoidongFromExcel(rows) {
    let currentHDTitle = "";
    let currentBoxId = null;

    rows.forEach((row, idx) => {
        if (idx < 5 || !row) return; // Bỏ qua tiêu đề
        const hdVal = String(row[0] || '').trim();
        if (hdVal && hdVal !== currentHDTitle && !isNaN(hdVal)) {
            currentHDTitle = `Hội Đồng ${hdVal}`;
            hoidongCount++;
            currentBoxId = createNewHoiDongBox(currentHDTitle);
        }
    });
}

// ==========================================
// 3. KHỞI TẠO KÉO THẢ & RÀNG BUỘC LOGIC (DRAG & DROP)
// ==========================================
function initGlobalSortables() {
    const poolSV = document.getElementById('pool-sv');
    if (poolSV) {
        Sortable.create(poolSV, { group: 'sinhvien', animation: 150 });
    }

    const poolGV = document.getElementById('pool-gv');
    if (poolGV) {
        Sortable.create(poolGV, {
            group: { name: 'giangvien', pull: 'clone', put: false },
            animation: 150,
            sort: false
        });
    }
}

function createNewHoiDongBox(titleText) {
    hoidongCount++;
    const boxId = `hd-box-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
    const boxHTML = `
        <div class="box-hoidong p-2 mb-3 border rounded bg-white" id="${boxId}">
            <div class="d-flex gap-2 mb-2 bg-light p-2 rounded align-items-center">
                <input type="text" class="form-control form-control-sm fw-bold hd-title" value="${titleText}" placeholder="Số/Tên HĐ" style="width: 130px;">
                <input type="date" class="form-control form-control-sm hd-date" style="width: 130px;">
                <select class="form-select form-select-sm hd-time" style="width: 100px;">
                    <option value="">--Giờ--</option>
                    <option value="07h00">07h00</option>
                    <option value="08h00">08h00</option>
                    <option value="13h00">13h00</option>
                    <option value="14h00">14h00</option>
                    <option value="17h00">17h00</option>
                </select>
                <input type="text" class="form-control form-control-sm hd-room" placeholder="Phòng (VD: C120)" style="width: 120px;">
                <select class="form-select form-select-sm hd-location">
                    <option value="120 Hoàng Minh Thảo">120 Hoàng Minh Thảo</option>
                    <option value="254 Nguyễn Văn Linh">254 Nguyễn Văn Linh</option>
                    <option value="137 Nguyễn Văn Linh">137 Nguyễn Văn Linh</option>
                    <option value="03 Quang Trung">03 Quang Trung</option>
                    <option value="K7/25 Quang Trung">K7/25 Quang Trung</option>
                    <option value="Phan Văn Trị">Phan Văn Trị</option>
                </select>
                <button type="button" class="btn btn-outline-danger btn-sm ms-auto" onclick="document.getElementById('${boxId}').remove()" title="Xóa Hội đồng">✕</button>
            </div>
            <div class="row">
                <div class="col-4 border-end">
                    <div class="mb-1"><small class="text-danger fw-bold">Chủ tịch:</small><div class="dropzone-gv" data-role="Chủ tịch"></div></div>
                    <div class="mb-1"><small class="text-success fw-bold">Phản biện:</small><div class="dropzone-gv" data-role="Phản biện"></div></div>
                    <div><small class="text-primary fw-bold">Thư ký:</small><div class="dropzone-gv" data-role="Thư ký"></div></div>
                </div>
                <div class="col-8">
                    <small class="fw-bold">Sinh viên bảo vệ:</small>
                    <div class="dropzone-sv p-2 bg-light border rounded" style="min-height: 120px;"></div>
                </div>
            </div>
        </div>
    `;
    
    const workspace = document.getElementById('workspace-hd');
    workspace.insertAdjacentHTML('beforeend', boxHTML);
    attachSortableToHoiDong(boxId);
    return boxId;
}

function attachSortableToHoiDong(boxId) {
    const box = document.getElementById(boxId);
    if (!box) return;
    
    // Vùng thả Sinh viên
    Sortable.create(box.querySelector('.dropzone-sv'), {
        group: 'sinhvien',
        animation: 150
    });

    // Vùng thả Giảng viên (3 vị trí)
    box.querySelectorAll('.dropzone-gv').forEach(zone => {
        Sortable.create(zone, {
            group: 'giangvien',
            animation: 150,
            onAdd: function (evt) {
                const item = evt.item;
                const gvId = item.getAttribute('data-id');
                const gvName = item.getAttribute('data-name');
                const currBox = evt.to.closest('.box-hoidong');
                
                // RÀNG BUỘC 1: Mỗi vị trí chỉ nhận 1 Giảng viên
                if (evt.to.children.length > 1) {
                    evt.to.removeChild(evt.to.children[0]);
                }

                // RÀNG BUỘC 2: GV không được kiêm nhiệm 2 vị trí trong 1 HĐ
                const allGVInThisHD = Array.from(currBox.querySelectorAll('.dropzone-gv .card-item'));
                const duplicates = allGVInThisHD.filter(node => node.getAttribute('data-id') === gvId);
                if (duplicates.length > 1) {
                    alert(`LỖI: Giảng viên ${gvName} đã có chức danh trong Hội đồng này rồi!`);
                    item.remove(); 
                    return;
                }

                // RÀNG BUỘC 3: Cảnh báo trùng lịch (Cùng Ngày + Cùng Giờ) ở HĐ khác
                const thisDate = currBox.querySelector('.hd-date').value;
                const thisTime = currBox.querySelector('.hd-time').value;
                if (thisDate && thisTime) {
                    const allHDs = document.querySelectorAll('.box-hoidong');
                    for (let hd of allHDs) {
                        if (hd.id === currBox.id) continue;
                        const thatDate = hd.querySelector('.hd-date').value;
                        const thatTime = hd.querySelector('.hd-time').value;

                        if (thatDate === thisDate && thatTime === thisTime) {
                            const existInOther = hd.querySelector(`.dropzone-gv .card-item[data-id="${gvId}"]`);
                            if (existInOther) {
                                const otherTitle = hd.querySelector('.hd-title').value;
                                alert(`CẢNH BÁO TRÙNG LỊCH: GV ${gvName} đang bận tại ${otherTitle} vào khung giờ ${thisTime} ngày ${thisDate}!`);
                                item.remove(); 
                                return;
                            }
                        }
                    }
                }
            }
        });
    });
}

// ==========================================
// 4. XỬ LÝ XUẤT HỒ SƠ WORD (ZIP) & EXCEL
// ==========================================
async function handleExport() {
    const khoa = document.getElementById('khoa-select').value;
    let exportData = []; 

    // Quét tất cả các Box Hội Đồng
    document.querySelectorAll('.box-hoidong').forEach(box => {
        const hdName = box.querySelector('.hd-title').value || 'Hội Đồng';
        const dateStr = box.querySelector('.hd-date').value || '';
        const time = box.querySelector('.hd-time').value || '';
        const room = box.querySelector('.hd-room').value || '';
        const location = box.querySelector('.hd-location').value || '';

        const getGV = (role) => {
            const el = box.querySelector(`.dropzone-gv[data-role="${role}"] .card-item`);
            return el ? el.getAttribute('data-name') : '';
        };

        const chuTich = getGV('Chủ tịch');
        const phanBien = getGV('Phản biện');
        const thuKy = getGV('Thư ký');

        // Định dạng lại Ngày bảo vệ DD/MM/YYYY
        let formattedDate = dateStr;
        if (dateStr.includes('-')) {
            const parts = dateStr.split('-');
            formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
        }

        box.querySelectorAll('.dropzone-sv .card-item').forEach(svEl => {
            exportData.push({
                khoa: khoa,
                so_hd: hdName,
                ngay: formattedDate,
                gio: time,
                phong: room,
                dia_diem: location,
                chu_tich: chuTich,
                phan_bien: phanBien,
                thu_ky: thuKy,
                ma_sv: svEl.getAttribute('data-id'),
                ho_ten: svEl.getAttribute('data-name'),
                de_tai: svEl.getAttribute('data-detai'),
                lop: svEl.getAttribute('data-lop'),
                gvhd: svEl.getAttribute('data-gvhd')
            });
        });
    });

    if (exportData.length === 0) {
        alert("Chưa có sinh viên nào được phân công vào Hội đồng!");
        return;
    }

    try {
        // A. TẠO FILE ZIP CHỨA TOÀN BỘ FILE WORD
        const zip = new JSZip();

        exportData.forEach(data => {
            const svFolder = zip.folder(`${data.ma_sv} - ${data.ho_ten}`);

            const context = {
                KHOA: data.khoa,
                SO_HOI_DONG: data.so_hd,
                NGÀY__BẢO_VỆ: data.ngay,
                GIỜ: data.gio,
                ĐỊA_ĐIỂM: `Phòng ${data.phong}, cơ sở ${data.dia_diem}`,
                Chủ_tịch: data.chu_tich,
                Phản_biện: data.phan_bien,
                Thư_ký: data.thu_ky,
                GVHD: data.gvhd,
                HỌ_VÀ_TÊN: data.ho_ten,
                MÃ_SỐ_SV: data.ma_sv,
                LỚP: data.lop,
                TÊN_ĐỀ_TÀI_ĐỒ_ÁN_TỐT_NGHIỆP: data.de_tai
            };

            const renderWord = (templateBinary, fileName) => {
                if (!templateBinary) return;
                const doc = new window.docxtemplater(new PizZip(templateBinary), { 
                    paragraphLoop: true, 
                    linebreaks: true 
                });
                doc.render(context);
                const outBlob = doc.getZip().generate({ type: "blob" });
                svFolder.file(fileName, outBlob);
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

        // B. TẠO FILE EXCEL BẢNG ĐIỂM BẰNG EXCELJS
        await generateExcelReport(exportData);

        alert("🎉 Xuất Hồ sơ Word (.ZIP) và Bảng điểm Excel (.XLSX) thành công!");
    } catch (err) {
        console.error("Lỗi khi xuất file:", err);
        alert("Có lỗi xảy ra trong quá trình tạo file: " + err.message);
    }
}

// ==========================================
// 5. THUẬT TOÁN GOM NHÓM & TẠO FILE EXCEL
// ==========================================
async function generateExcelReport(exportData) {
    if (typeof ExcelJS === 'undefined') {
        console.warn("ExcelJS chưa được nạp, bỏ qua bước tạo Excel.");
        return;
    }

    const workbook = new ExcelJS.Workbook();

    // Gom nhóm Sinh viên theo Key: Nganh_Phong_Ngay_Gio
    const groups = {};
    exportData.forEach(item => {
        // Tự động phân tách ngành từ lớp (VD: K27 CTP -> CNTP)
        let nganh = "ĐỒ ÁN TỐT NGHIỆP";
        if (item.lop.includes("CTP")) nganh = "CÔNG NGHỆ THỰC PHẨM";
        else if (item.lop.includes("EKD") || item.lop.includes("KTD")) nganh = "KỸ THUẬT ĐIỆN";
        else if (item.lop.includes("EDK")) nganh = "KỸ THUẬT ĐIỀU KHIỂN & TỰ ĐỘNG HÓA";
        else if (item.lop.includes("EDT")) nganh = "ĐIỆN TỰ ĐỘNG";

        const groupKey = `${nganh}_${item.phong}_${item.ngay}`;
        if (!groups[groupKey]) {
            groups[groupKey] = {
                nganh: nganh,
                phong: item.phong,
                ngay: item.ngay,
                gio: item.gio,
                diaDiem: item.dia_diem,
                students: []
            };
        }
        groups[groupKey].students.push(item);
    });

    // Tạo từng Sheet tương ứng mỗi Nhóm
    Object.keys(groups).forEach((key, sheetIdx) => {
        const grp = groups[key];
        const sheetName = `${grp.phong || 'Phong'}_${sheetIdx + 1}`.substring(0, 30);
        const ws = workbook.addWorksheet(sheetName);

        // Header Tiêu đề trường
        ws.mergeCells('A1:C1'); ws.getCell('A1').value = 'BỘ GIÁO DỤC VÀ ĐÀO TẠO';
        ws.mergeCells('A2:C2'); ws.getCell('A2').value = 'ĐẠI HỌC DUY TÂN';
        ws.mergeCells('A3:C3'); ws.getCell('A3').value = 'TRƯỜNG CÔNG NGHỆ VÀ KỸ THUẬT';

        ws.mergeCells('D1:N1'); ws.getCell('D1').value = 'BẢNG ĐIỂM GHI KẾT QUẢ BẢO VỆ ĐỒ ÁN TỐT NGHIỆP';
        ws.getCell('D1').font = { bold: true, size: 14 };

        ws.getCell('D2').value = `NGÀNH: ${grp.nganh}`;
        ws.getCell('D3').value = `Thời gian: ${grp.gio} ngày ${grp.ngay}`;
        ws.getCell('H3').value = `Địa điểm: Phòng ${grp.phong}, cơ sở ${grp.diaDiem}`;

        // Header Bảng Dữ Liệu (Dòng 7 & 8)
        const headers1 = ["STT", "MSSV", "HỌ VÀ TÊN SINH VIÊN", "KHÓA", "TÊN GVHD", "ĐIỂM GVHD (30%)", "ĐIỂM CỦA THÀNH VIÊN HĐ", "", "", "", "ĐHĐ (70%)", "ĐIỂM TỔNG KẾT 100%", "", "TÊN THÀNH VIÊN HĐ THEO CHỨC DANH"];
        const headers2 = ["", "", "", "", "", "", "CT", "PB", "UV", "TK", "", "SỐ", "CHỮ", ""];
        
        ws.addRow(headers1);
        ws.addRow(headers2);

        // Merge các ô Header
        ws.mergeCells('A7:A8'); ws.mergeCells('B7:B8'); ws.mergeCells('C7:C8');
        ws.mergeCells('D7:D8'); ws.mergeCells('E7:E8'); ws.mergeCells('F7:F8');
        ws.mergeCells('G7:J7'); ws.mergeCells('K7:K8'); ws.mergeCells('L7:M7');
        ws.mergeCells('N7:N8');

        let currentRow = 9;
        grp.students.forEach((sv, svIdx) => {
            const startR = currentRow;

            // Mỗi sinh viên chiếm 3 dòng để vừa khớp 3 chức danh Hội đồng (CT, PB, TK)
            ws.getCell(`A${startR}`).value = svIdx + 1;
            ws.getCell(`B${startR}`).value = sv.ma_sv;
            ws.getCell(`C${startR}`).value = sv.ho_ten;
            ws.getCell(`D${startR}`).value = sv.lop;
            ws.getCell(`E${startR}`).value = sv.gvhd;

            // Gán tên 3 thành viên Hội đồng vào cột N (Dòng 1: CT, Dòng 2: PB, Dòng 3: TK)
            ws.getCell(`N${startR}`).value = `CT: ${sv.chu_tich}`;
            ws.getCell(`N${startR + 1}`).value = `UVPB: ${sv.phan_bien}`;
            ws.getCell(`N${startR + 2}`).value = `UVTK: ${sv.thu_ky}`;

            // Merge các ô thông tin SV qua 3 dòng
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

        // Căn chỉnh độ rộng cột
        ws.columns.forEach(column => {
            column.width = 15;
        });
        ws.getColumn(3).width = 25; // Họ tên SV
        ws.getColumn(5).width = 25; // Tên GVHD
        ws.getColumn(14).width = 32; // Thành viên HĐ
    });

    // Xuất file Excel
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, "Bang_Diem_Tong_Hop_DATN.xlsx");
}
