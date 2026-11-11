// ==========================================
// 1. BIẾN TOÀN CỤC & LOAD TEMPLATE MẶC ĐỊNH
// ==========================================
let templates = {};
let hoidongCount = 0;

document.addEventListener("DOMContentLoaded", async () => {
    // Tải ngầm 4 file Word templates từ thư mục
    const loadFile = (url) => new Promise((resolve, reject) => {
        PizZipUtils.getBinaryContent(url, (err, data) => {
            if (err) reject(err); else resolve(data);
        });
    });

    try {
        templates.bienBan = await loadFile('templates/1_BIEN_BAN_BAO_VE_DATN.docx');
        templates.phieuGVHD = await loadFile('templates/2_Phieu_GVHD.docx');
        templates.phieuGVPB = await loadFile('templates/3_Phieu_GVPB.docx');
        templates.phieuHD = await loadFile('templates/4_Phieu_Hoi_Dong.docx');
        console.log("✅ Load templates nội bộ thành công!");
    } catch (e) {
        alert("Lỗi tải template: Đảm bảo chạy qua Local Server (VSCode Live Server). Không mở trực tiếp file html (file://)");
    }

    initGlobalSortables();
});

// ==========================================
// 2. KHỞI TẠO KÉO THẢ & RÀNG BUỘC LOGIC (DRAG & DROP)
// ==========================================
function initGlobalSortables() {
    // Pool Sinh Viên (Trái)
    Sortable.create(document.getElementById('pool-sv'), {
        group: 'sinhvien', animation: 150
    });

    // Pool Giảng Viên (Phải) - Sử dụng 'clone' để 1 GV kéo được nhiều HĐ
    Sortable.create(document.getElementById('pool-gv'), {
        group: { name: 'giangvien', pull: 'clone', put: false },
        animation: 150, sort: false
    });
}

function attachSortableToHoiDong(boxId) {
    const box = document.getElementById(boxId);
    
    // Vùng thả SV
    Sortable.create(box.querySelector('.dropzone-sv'), {
        group: 'sinhvien', animation: 150
    });

    // Vùng thả GV (3 vị trí)
    box.querySelectorAll('.dropzone-gv').forEach(zone => {
        Sortable.create(zone, {
            group: 'giangvien', animation: 150,
            onAdd: function (evt) {
                const item = evt.item;
                const gvId = item.getAttribute('data-id');
                const gvName = item.getAttribute('data-name');
                const currBox = evt.to.closest('.box-hoidong');
                
                // RÀNG BUỘC 1: Chỉ cho phép 1 GV / 1 Dropzone (xóa người cũ nếu có)
                if (evt.to.children.length > 1) {
                    evt.to.removeChild(evt.to.children[0]);
                }

                // RÀNG BUỘC 2: GV không được kiêm nhiệm 2 vị trí trong 1 HĐ
                const allGVInThisHD = Array.from(currBox.querySelectorAll('.dropzone-gv .card-item'));
                const duplicates = allGVInThisHD.filter(node => node.getAttribute('data-id') === gvId);
                if (duplicates.length > 1) {
                    alert(`LỖI: ${gvName} đã có chức danh trong Hội đồng này!`);
                    item.remove(); return;
                }

                // RÀNG BUỘC 3: Check trùng lịch (Cùng ngày + Cùng giờ)
                const thisDate = currBox.querySelector('.hd-date').value;
                const thisTime = currBox.querySelector('.hd-time').value;
                if (thisDate && thisTime) {
                    const allHDs = document.querySelectorAll('.box-hoidong');
                    for (let hd of allHDs) {
                        if (hd.id === currBox.id) continue;
                        if (hd.querySelector('.hd-date').value === thisDate && 
                            hd.querySelector('.hd-time').value === thisTime) {
                            
                            const existInOther = hd.querySelector(`.dropzone-gv .card-item[data-id="${gvId}"]`);
                            if (existInOther) {
                                alert(`CẢNH BÁO: ${gvName} đang bận tại ${hd.querySelector('.hd-title').value} vào khung giờ này!`);
                                item.remove(); return;
                            }
                        }
                    }
                }
            }
        });
    });
}

// ==========================================
// 3. TẠO GIAO DIỆN HỘI ĐỒNG ĐỘNG
// ==========================================
document.getElementById('btn-add-hd').addEventListener('click', () => {
    hoidongCount++;
    const boxId = `hd-${hoidongCount}`;
    const boxHTML = `
        <div class="box-hoidong p-2" id="${boxId}">
            <div class="d-flex gap-2 mb-2 bg-light p-2 rounded">
                <input type="text" class="form-control form-control-sm fw-bold hd-title" value="Hội Đồng ${hoidongCount}" placeholder="Số/Tên HĐ">
                <input type="date" class="form-control form-control-sm hd-date">
                <select class="form-select form-select-sm hd-time">
                    <option value="">--Giờ--</option>
                    <option value="07h00">07h00</option>
                    <option value="08h00">08h00</option>
                    <option value="13h00">13h00</option>
                    <option value="14h00">14h00</option>
                    <option value="17h00">17h00</option>
                </select>
                <input type="text" class="form-control form-control-sm hd-room" placeholder="Phòng (VD: C120)">
                <select class="form-select form-select-sm hd-location">
                    <option>120 Hoàng Minh Thảo</option>
                    <option>254 Nguyễn Văn Linh</option>
                    <option>03 Quang Trung</option>
                    <option>Phan Văn Trị</option>
                </select>
            </div>
            <div class="row">
                <div class="col-4 border-end">
                    <div class="mb-1"><small class="text-danger fw-bold">Chủ tịch:</small><div class="dropzone-gv" data-role="Chủ tịch"></div></div>
                    <div class="mb-1"><small class="text-success fw-bold">Phản biện:</small><div class="dropzone-gv" data-role="Phản biện"></div></div>
                    <div><small class="text-primary fw-bold">Thư ký:</small><div class="dropzone-gv" data-role="Thư ký"></div></div>
                </div>
                <div class="col-8">
                    <small class="fw-bold">Sinh viên bảo vệ:</small>
                    <div class="dropzone-sv"></div>
                </div>
            </div>
        </div>
    `;
    document.getElementById('workspace-hd').insertAdjacentHTML('beforeend', boxHTML);
    attachSortableToHoiDong(boxId);
});

// ==========================================
// 4. XỬ LÝ DỮ LIỆU ĐẦU RA KHI BẤM NÚT XUẤT
// ==========================================
document.getElementById('btn-export').addEventListener('click', async () => {
    // 1. Trích xuất toàn bộ dữ liệu từ Giao diện DOM hiện tại
    const khoa = document.getElementById('khoa-select').value;
    let exportData = []; // Mảng chứa dữ liệu chuẩn bị xuất

    document.querySelectorAll('.box-hoidong').forEach(box => {
        const hdName = box.querySelector('.hd-title').value;
        const dateStr = box.querySelector('.hd-date').value;
        const time = box.querySelector('.hd-time').value;
        const room = box.querySelector('.hd-room').value;
        const location = box.querySelector('.hd-location').value;

        // Lấy GV (Xử lý an toàn nếu chưa kéo thả)
        const getGV = (role) => {
            const el = box.querySelector(`.dropzone-gv[data-role="${role}"] .card-item`);
            return el ? el.getAttribute('data-name') : `[Thiếu ${role}]`;
        };
        const chuTich = getGV('Chủ tịch');
        const phanBien = getGV('Phản biện');
        const thuKy = getGV('Thư ký');

        // Lấy DS Sinh viên trong HĐ này
        box.querySelectorAll('.dropzone-sv .card-item').forEach(svEl => {
            exportData.push({
                khoa: khoa,
                so_hd: hdName,
                ngay: dateStr, // Có thể format lại DD/MM/YYYY
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

    if(exportData.length === 0) return alert("Chưa có sinh viên nào được phân công!");

    // 2. TẠO FILE ZIP CHỨA CÁC FILE WORD
    const zip = new JSZip();
    exportData.forEach(data => {
        const svFolder = zip.folder(`${data.ma_sv} - ${data.ho_ten}`);
        
        // Dictionary Map Variable -> Data
        const context = {
            KHOA: data.khoa, SO_HOI_DONG: data.so_hd,
            NGÀY__BẢO_VỆ: data.ngay, GIỜ: data.gio, ĐỊA_ĐIỂM: `${data.phong} - ${data.dia_diem}`,
            Chủ_tịch: data.chu_tich, Phản_biện: data.phan_bien, Thư_ký: data.thu_ky,
            GVHD: data.gvhd, HỌ_VÀ_TÊN: data.ho_ten, MÃ_SỐ_SV: data.ma_sv, LỚP: data.lop,
            TÊN_ĐỀ_TÀI_ĐỒ_ÁN_TỐT_NGHIỆP: data.de_tai
        };

        // Hàm render Word
        const renderWord = (templateBinary, fileName) => {
            const doc = new window.docxtemplater(new PizZip(templateBinary), { paragraphLoop: true, linebreaks: true });
            doc.render(context);
            svFolder.file(fileName, doc.getZip().generate({ type: "blob" }));
        };

        renderWord(templates.bienBan, '1_BIEN_BAN_BAO_VE.docx');
        renderWord(templates.phieuGVHD, '2_Phieu_GVHD.docx');
        renderWord(templates.phieuGVPB, '3_Phieu_GVPB.docx');
        // Tách 3 biểu mẫu hội đồng
        renderWord(templates.phieuHD, '4a_Phieu_HD_ChuTich.docx'); // (Sẽ cần template riêng nếu bạn muốn ẩn text ko lq)
        renderWord(templates.phieuHD, '4b_Phieu_HD_PhanBien.docx');
        renderWord(templates.phieuHD, '4c_Phieu_HD_ThuKy.docx');
    });

    // Tải ZIP về
    zip.generateAsync({ type: "blob" }).then(content => {
        saveAs(content, "Ho_So_Tot_Nghiep.zip");
    });

    // 3. TẠO BÁO CÁO EXCEL BẰNG EXCELJS (Gom nhóm Sheet & Merge)
    // - Dùng Thuật toán gom nhóm Ngành + Ngày + Phòng
    // - Code gen ExcelJS khá dài, sẽ xử lý Workbook.addWorksheet, mergeCells('M8:M10') cho CT, PB, TK.
    alert("Đã tạo xong File ZIP Hồ sơ Word! Báo cáo Excel sẽ được triển khai theo cấu trúc gom nhóm.");
});
