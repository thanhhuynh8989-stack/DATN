// Hàm lõi: Loại bỏ triệt để mọi biến thể học hàm, học vị, danh hiệu (GS, PGS, TS, ThS, Th.S, PGS.TS, GVC, KS, CN, KTS, BS, DS, NGƯT...)
export function stripAcademicTitles(rawName) {
    if (!rawName) return '';
    let s = String(rawName).trim();

    // 1. Xóa các cụm viết đầy đủ bằng tiếng Việt (nếu có)
    s = s.replace(/\b(phó\s+giáo\s+sư|giáo\s+sư|tiến\s+sĩ|tiến\s+sỹ|thạc\s+sĩ|thạc\s+sỹ|giảng\s+viên\s+chính|giảng\s+viên|kỹ\s+sư|kiến\s+trúc\s+sư|cử\s+nhân|bác\s+sĩ|bác\s+sỹ|dược\s+sĩ|dược\s+sỹ|nhà\s+giáo\s+nhân\s+dân|nhà\s+giáo\s+ưu\s+tú)\b/gi, ' ');

    // 2. Xóa các biến thể viết tắt có dấu chấm/gạch ngang xen giữa (VD: Th.S, Th.S., PGS.TS, GS-TS, PGS.TS.BS, TS.KTS...)
    // Lặp vài vòng để bóc sạch chuỗi học hàm/học vị ghép ở đầu tên (VD: "PGS. TS. Th.S. Nguyễn Văn A")
    const titleTokenRegex = /(?:^|[\s.,;\-\/()])(?:p\.?\s*g\.?\s*s|g\.?\s*s|t\.?\s*s|th\.?\s*s|g\.?\s*v\.?\s*c|g\.?\s*v|k\.?\s*t\.?\s*s|k\.?\s*s|c\.?\s*n|b\.?\s*s|d\.?\s*s\.?\s*c|d\.?\s*s|n\.?\s*g\.?\s*n\.?\s*d|n\.?\s*g\.?\s*ư\.?\s*t|n\.?\s*g\.?\s*u\.?\s*t|prof|assoc\.?\s*prof|dr|ph\.?\s*d|m\.?\s*sc|msc|m\.?\s*a|b\.?\s*sc|mr|ms|mrs)(?=[\s.,;\-\/()]|$)/gi;

    let prev;
    do {
        prev = s;
        s = s.replace(titleTokenRegex, ' ');
    } while (s !== prev);

    // 3. Làm sạch dấu chấm/gạch ngang/khoảng trắng dư thừa còn sót ở đầu hoặc giữa tên
    s = s.replace(/^[\s.,;\-\/()]+|[\s.,;\-\/()]+$/g, '')
         .replace(/\s+/g, ' ')
         .trim();

    return s;
}

// Chuẩn hóa tên GV: Bỏ học hàm/học vị -> Bỏ dấu tiếng Việt -> Xóa ký tự đặc biệt -> Viết thường
export function normalizeTeacherKeyForDupCheck(rawName) {
    const stripped = stripAcademicTitles(rawName);
    if (!stripped) return '';
    return stripped
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd')
        .replace(/[^a-z0-9]/g, '');
}

// Hàm làm sạch và so sánh tên Giảng viên (dùng chung chuẩn hóa với normalizeTeacherKeyForDupCheck)
export function isSameTeacherName(name1, name2) {
    const k1 = normalizeTeacherKeyForDupCheck(name1);
    const k2 = normalizeTeacherKeyForDupCheck(name2);
    return Boolean(k1 && k2 && k1 === k2);
}

// Chuẩn hóa tên SV: Bỏ dấu tiếng Việt -> Xóa ký tự đặc biệt -> Viết thường
export function normalizeStudentKeyForDupCheck(rawName) {
    if (!rawName) return '';
    return String(rawName)
        .toLowerCase()
        .trim()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd')
        .replace(/[^a-z0-9]/g, '');
}