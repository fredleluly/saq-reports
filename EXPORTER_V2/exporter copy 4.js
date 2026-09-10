'use strict';

/**
* main.js — Konversi 1:1 dari main.py ke Node.js
*
* Dependencies (install via npm):
*   npm install pizzip docxtemplater exceljs
*   npm install image-size   <-- opsional, untuk dimensi gambar otomatis
*
* Versi:
*   - pizzip          ^3.x
*   - docxtemplater   ^3.x
*   - exceljs         ^4.x
*   - image-size      ^7.x  (opsional)
*/

const fs   = require('fs');
const path = require('path');
const PizZip        = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ExcelJS       = require('exceljs');

// Load env (Railway usually injects it; local still relies on .env)
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

// Use the same repository layer as the web app (now MongoDB-backed)
const { connectDB } = require('../config/database');
const {
    initRepositories,
    assistantRepo,
    lecturerRepo,
    courseRepo,
    scheduleRepo,
    reportRepo,
    globalConfigRepo,
} = require('../repositories');

// Opsional: image-size untuk mendapatkan dimensi gambar secara otomatis
let sizeOf = null;
try { sizeOf = require('image-size'); } catch (_) {}

// =============================================================================
// == KONFIGURASI PATH ==
// =============================================================================

const BASE_DIR            = __dirname;

// Input
const REPORTS_PATH        = path.join(BASE_DIR, '../reports.json');
const SCHEDULES_PATH      = path.join(BASE_DIR, '../schedules.json');
const COURSES_PATH        = path.join(BASE_DIR, '../courses.json');
const LECTURERS_PATH      = path.join(BASE_DIR, '../lecturers.json');
const ASSISTANTS_PATH     = path.join(BASE_DIR, '../config_asisten.json');
const GLOBAL_CONFIG_PATH  = path.join(BASE_DIR, '../config_global.json');

// Output
const TEMPLATE_PATH_DOCX      = path.join(BASE_DIR, 'template_saq.docx');
const OUTPUT_BASE_DIR_DOCX    = path.join(BASE_DIR, 'output_export');
const TEMPLATE_PATH_XLSX      = path.join(BASE_DIR, 'template_berita_acara.xlsx');
const OUTPUT_DIR_XLSX         = path.join(OUTPUT_BASE_DIR_DOCX, 'berita_acara');

const TABLE_ANCHOR_PLACEHOLDER = '###TABLE_ANCHOR###';

const ASSETS_DIR        = path.join(BASE_DIR, 'assets');
const LOGO_ITPLN_PATH   = path.join(ASSETS_DIR, 'ITPLN.webp');
const LOGO_SAQ_PATH     = path.join(ASSETS_DIR, 'SAQ.webp');
const TEMPLATE_REKAP_PATH = path.join(BASE_DIR, 'Rekap_Ngawas_Lab.xlsx');
const OUTPUT_DIR_REKAP  = path.join(OUTPUT_BASE_DIR_DOCX, 'rekap_kehadiran');

// =============================================================================
// == UTILITAS TANGGAL BAHASA INDONESIA ==
// Setara: locale.setlocale(locale.LC_TIME, 'id_ID.UTF-8')
// =============================================================================

// Setara Python: day_map_id = ["Senin","Selasa","Rabu","Kamis","Jumat","Sabtu","Minggu"]
// Python weekday(): 0=Senin … 6=Minggu
const DAY_MAP_ID   = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];
const MONTH_MAP_ID = [
    'Januari', 'Februari', 'Maret',    'April',   'Mei',      'Juni',
    'Juli',    'Agustus',  'September','Oktober',  'November', 'Desember',
];

/**
* Mengambil nama hari dalam Bahasa Indonesia dari objek Date.
* Setara: day_map_id[current_date.weekday()]
*/
function getDayNameID(date) {
    // JS getDay(): 0=Minggu, 1=Senin … 6=Sabtu
    // Python weekday(): 0=Senin … 6=Minggu
    const jsDay      = date.getDay();
    const pyWeekday  = jsDay === 0 ? 6 : jsDay - 1;
    return DAY_MAP_ID[pyWeekday];
}

/**
* Setara Python: current_date.strftime('%A, %d %B %Y')  (locale id_ID)
* Contoh output: "Senin, 03 Maret 2025"
*/
function formatDateIndonesian(date) {
    const day   = getDayNameID(date);
    const tgl   = String(date.getDate()).padStart(2, '0');
    const bulan = MONTH_MAP_ID[date.getMonth()];
    const tahun = date.getFullYear();
    return `${day}, ${tgl} ${bulan} ${tahun}`;
}

/**
* Setara Python: datetime.strptime(s, "%A, %d %B %Y")  (locale id_ID)
* Contoh input: "Senin, 03 Maret 2025"
*/
function parseDateIndonesian(dateStr) {
    const cleaned = dateStr.replace(',', '').trim();
    const parts   = cleaned.split(/\s+/);
    // parts: [hari, dd, bulan, yyyy]
    const tgl      = parseInt(parts[1], 10);
    const bulanIdx = MONTH_MAP_ID.indexOf(parts[2]);
    const tahun    = parseInt(parts[3], 10);
    if (bulanIdx === -1) throw new Error(`Nama bulan tidak dikenal: "${parts[2]}"`);
    return new Date(tahun, bulanIdx, tgl);
}

/**
* Setara Python: dt_object.strftime("%B").lower()  (locale id_ID)
* Contoh output: "maret"
*/
function getMonthNameIDLower(date) {
    return MONTH_MAP_ID[date.getMonth()].toLowerCase();
}

/**
* Setara Python: dt_object.strftime('%Y_%m_%d_%A')  (locale id_ID)
* Contoh output: "2025_03_03_Senin"
*/
function formatFilenameDate(date) {
    const yyyy    = date.getFullYear();
    const mm      = String(date.getMonth() + 1).padStart(2, '0');
    const dd      = String(date.getDate()).padStart(2, '0');
    const dayName = getDayNameID(date);
    return `${yyyy}_${mm}_${dd}_${dayName}`;
}

/**
* Helper untuk format YYYY-MM-DD
*/
function isoDateString(date) {
    const yyyy = date.getFullYear();
    const mm   = String(date.getMonth() + 1).padStart(2, '0');
    const dd   = String(date.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

// =============================================================================
// == FUNGSI INTI: MEMBANGUN DAFTAR LAPORAN KOMPREHENSIF ==
// =============================================================================

/**
* Setara Python: load_json_data(path, key_name=None)
*/
function loadJsonData(filePath, keyName = null) {
    try {
        const content = fs.readFileSync(filePath, 'utf-8');
        const data    = JSON.parse(content);
        console.log(`Berhasil memuat data dari ${filePath}`);
        return keyName ? data[keyName] : data;
    } catch (err) {
        if (err.code === 'ENOENT') {
            console.log(`Error: File tidak ditemukan di ${filePath}`);
        } else {
            console.log(`Error: Gagal memuat JSON dari ${filePath}. Pastikan formatnya benar.`);
        }
        return null;
    }
}

/**
* Setara Python: clean_cell_value(value)
*/
function cleanCellValue(value) {
    if (value === null || value === undefined) return '';
    const cleaned = String(value).trim().toLowerCase();
    if (cleaned === 'none' || cleaned === 'n/a') return '';
    return String(value).trim();
}

// =============================================================================
// == FUNGSI INTI: MEMBANGUN DAFTAR LAPORAN KOMPREHENSIF ==
// =============================================================================

/**
* Setara Python: build_comprehensive_list_with_placeholders()
*/
function buildComprehensiveListWithPlaceholders() {
    console.log('\n>>> Memulai proses pembangunan daftar laporan komprehensif...');
    
    // 1. Muat semua data master
    const assistantsArr = assistantRepo.read();
    const lecturersArr  = lecturerRepo.read();
    const coursesArr    = courseRepo.read();
    const schedules     = scheduleRepo.read();
    const globalConfig  = globalConfigRepo.read();
    const actualReports = reportRepo.read();
    
    // Setara: if not all([assistants, lecturers, courses, schedules, global_config, actual_reports is not None])
    if (!assistantsArr || !lecturersArr || !coursesArr || !schedules || !globalConfig) {
        console.log('!!! Gagal memuat satu atau lebih file data master penting (asisten/dosen/matkul/jadwal/config). Proses dibatalkan.');
        return null;
    }
    
    // Setara: {a['id']: a for a in ...}
    const assistants = Object.fromEntries(assistantsArr.map(a => [a.id, a]));
    const lecturers  = Object.fromEntries(lecturersArr.map(l  => [l.id, l]));
    const courses    = Object.fromEntries(coursesArr.map(c   => [c.id, c]));
    
    const semesterInfo   = globalConfig.semester || {};
    const semesterYear   = semesterInfo.year;
    const semesterMonths = semesterInfo.months || [];
    
    if (!semesterYear || !semesterMonths.length) {
        console.log('!!! Info semester (tahun/bulan) tidak ditemukan di config_global.json.');
        return null;
    }
    
    // 2. Buat lookup map untuk laporan yang sudah ada
    // Setara: key = (report.get('schedule_id'), report.get('date'))
    const reportsMap = {};
    for (const report of actualReports) {
        const key = `${report.schedule_id}__${report.date}`;
        reportsMap[key] = report;
    }
    console.log(`Ditemukan ${Object.keys(reportsMap).length} laporan aktual.`);
    
    // 3. Generate semua kemungkinan jadwal (placeholder)
    const allPossibleReports = [];
    
    // Setara: start_date = datetime(semester_year, semester_months[0], 1)
    const startDate = new Date(semesterYear, semesterMonths[0] - 1, 1);
    
    // Setara:
    //   last_month = semester_months[-1]
    //   last_day   = (datetime(year, last_month+1, 1) - timedelta(days=1)).day
    //   end_date   = datetime(year, last_month, last_day)
    const lastMonth = semesterMonths[semesterMonths.length - 1];
    // new Date(year, lastMonth, 0) = hari terakhir bulan lastMonth (1-based)
    const endDate = new Date(semesterYear, lastMonth, 0);
    
    const currentDate = new Date(startDate);
    while (currentDate <= endDate) {
        const currentMonth1Based = currentDate.getMonth() + 1;
        
        if (semesterMonths.includes(currentMonth1Based)) {
            const dayName = getDayNameID(currentDate);
            const isoDate = isoDateString(currentDate);
            
            // Ambil semua laporan aktual untuk tanggal ini
            const reportsForThisDate = actualReports.filter(r => r.date === isoDate);
            
            // Ambil jadwal rutin untuk hari ini
            const schedulesForDay = schedules.filter(s => s.day === dayName);
            
            // JIKA tidak ada jadwal rutin DAN tidak ada laporan aktual, JANGAN buat DOCX (skip Sabtu/Minggu)
            if (schedulesForDay.length === 0 && reportsForThisDate.length === 0) {
                currentDate.setDate(currentDate.getDate() + 1);
                continue;
            }
            
            for (const scheduleInfo of schedulesForDay) {
                const scheduleId   = scheduleInfo.id;
                const mapKey       = `${scheduleId}__${isoDate}`;
                const actualReport = reportsMap[mapKey] || null;
                
                const courseInfo   = courses[scheduleInfo.course_id]     || {};
                const lecturerInfo = lecturers[scheduleInfo.lecturer_id] || {};
                
                let combinedRecord;
                if (actualReport) {
                    combinedRecord = {
                        hariTanggal:        actualReport.formattedDate,
                        materi:             actualReport.materi,
                        deskripsiKegiatan:  actualReport.deskripsiKegiatan,
                        keterangan:         actualReport.keterangan,
                        status:             actualReport.status || 'hadir',
                        imagePath:          actualReport.imagePath,
                        asisten: (actualReport.attending_assistant_ids || [])
                        .map(id => (assistants[id] || {}).panggilan || 'N/A'),
                        mataKuliah: courseInfo.name,
                        kelas:      scheduleInfo.class,
                        jam:        scheduleInfo.time,
                        sks:        courseInfo.sks,
                        dosen:      lecturerInfo.name,
                    };
                } else {
                    combinedRecord = {
                        hariTanggal:        formatDateIndonesian(currentDate),
                        materi:             null,
                        deskripsiKegiatan:  null,
                        keterangan:         null,
                        status:             'belum ada laporan',
                        imagePath:          null,
                        asisten: (scheduleInfo.assistant_ids || [])
                        .map(id => (assistants[id] || {}).panggilan || 'N/A'),
                        mataKuliah: courseInfo.name,
                        kelas:      scheduleInfo.class,
                        jam:        scheduleInfo.time,
                        sks:        courseInfo.sks,
                        dosen:      lecturerInfo.name,
                    };
                }
                allPossibleReports.push(combinedRecord);
            }
        }
        currentDate.setDate(currentDate.getDate() + 1);
    }
    
    console.log(`>>> Pembangunan daftar komprehensif selesai. Total ${allPossibleReports.length} entri (termasuk placeholder).`);
    return allPossibleReports;
}

// =============================================================================
// == BAGIAN PEMBUATAN LAPORAN DOCX ==
// =============================================================================

/**
* Escape karakter khusus XML.
*/
function escapeXml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
* Membangun XML tabel OOXML.
* Setara Python: blok pembuatan table_element dengan OxmlElement.
*
* column_widths_pct = [5, 15, 7, 13, 10, 22, 14, 14]
* headers = ["No","Mata Kuliah","Kelas","Dosen","Materi","Deskripsi Kegiatan","Asisten","Keterangan"]
*/
function buildTableXml(dataRecords) {
    const columnWidthsPct = [5, 15, 7, 13, 10, 22, 14, 14];
    const headers         = ['No', 'Mata Kuliah', 'Kelas', 'Dosen', 'Materi', 'Deskripsi Kegiatan', 'Asisten', 'Keterangan'];
    
    // tblGrid — setara: gridCol.set(qn('w:w'), str(int(pct_width * 50)))
    const gridColsXml = columnWidthsPct
    .map(pct => `<w:gridCol w:w="${Math.round(pct * 50)}"/>`)
    .join('');
    
    // tblBorders — setara: for border_name in ['top','left','bottom','right','insideH','insideV']
    const borderNames  = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
    const bordersXml   = borderNames
    .map(name => `<w:${name} w:val="single" w:sz="4" w:color="auto"/>`)
    .join('');
    
    // Header row — setara: tc_hdr dengan bold + center
    const headerCellsXml = headers
    .map(h =>
        `<w:tc>` +
        `<w:p>` +
        `<w:pPr><w:jc w:val="center"/></w:pPr>` +
        `<w:r><w:rPr><w:b/></w:rPr><w:t>${escapeXml(h)}</w:t></w:r>` +
        `</w:p>` +
        `</w:tc>`
    )
    .join('');
    const headerRowXml = `<w:tr>${headerCellsXml}</w:tr>`;
    
    // Data rows — setara: for i, record in enumerate(data_records_for_day)
    const dataRowsXml = dataRecords
    .map((record, i) => {
        const asistanStr = (record.asisten || []).join(', ');
        const values = [
            String(i + 1),
            cleanCellValue(record.mataKuliah),
            cleanCellValue(record.kelas),
            cleanCellValue(record.dosen),
            cleanCellValue(record.materi),
            cleanCellValue(record.deskripsiKegiatan),
            cleanCellValue(asistanStr),
            cleanCellValue(record.keterangan),
        ];
        
        const cellsXml = values
        .map((val, idx) => {
            // Setara: if val_idx == 0: center
            const centerXml = idx === 0
            ? `<w:pPr><w:jc w:val="center"/></w:pPr>`
            : '';
            return (
                `<w:tc>` +
                `<w:tcPr><w:vAlign w:val="center"/></w:tcPr>` +
                `<w:p>${centerXml}` +
                `<w:r><w:t xml:space="preserve">${escapeXml(val)}</w:t></w:r>` +
                `</w:p>` +
                `</w:tc>`
            );
        })
        .join('');
        
        return `<w:tr>${cellsXml}</w:tr>`;
    })
    .join('');
    
    return (
        `<w:tbl>` +
        `<w:tblPr>` +
        `<w:tblW w:w="5000" w:type="pct"/>` +
        `<w:tblLayout w:type="fixed"/>` +
        `<w:tblGrid>${gridColsXml}</w:tblGrid>` +
        `<w:tblBorders>${bordersXml}</w:tblBorders>` +
        `</w:tblPr>` +
        headerRowXml +
        dataRowsXml +
        `</w:tbl>`
    );
}

/**
* Mengganti paragraf yang mengandung anchorText dengan tableXml.
* Setara Python:
*   p_element.addprevious(table_element)
*   p_element.getparent().remove(p_element)
*/
function replaceAnchorParagraphWithTable(docXml, anchorText, tableXml) {
    // Cari paragraf yang teksnya mengandung anchorText
    // (anchor mungkin terbagi antar <w:r> run, jadi strip tag dulu)
    const pRegex = /<w:p[ >][\s\S]*?<\/w:p>/g;
    let match;
    
    while ((match = pRegex.exec(docXml)) !== null) {
        const pContent  = match[0];
        const plainText = pContent.replace(/<[^>]+>/g, ''); // strip XML tags
        if (plainText.includes(anchorText)) {
            // Ganti paragraf anchor dengan XML tabel
            return (
                docXml.slice(0, match.index) +
                tableXml +
                docXml.slice(match.index + pContent.length)
            );
        }
    }
    
    console.log(`  [!] Peringatan: Placeholder '${anchorText}' tidak ditemukan di template.`);
    return docXml;
}

/**
* Menambahkan paragraf teks di akhir body.
* Setara Python: doc.add_paragraph(caption_text)
*/
function appendParagraphXml(docXml, text) {
    const paraXml = `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
    return docXml.replace(/<\/w:body>/, paraXml + '</w:body>');
}

/**
* Menambahkan gambar inline di akhir body.
* Setara Python: doc.add_picture(full_image_path, width=Cm(8))
*   - width=Cm(8) → 2 880 000 EMU  (1 cm = 360 000 EMU)
*   - tinggi dihitung proporsional dari dimensi asli gambar
*/
function appendImageToDocXml(zip, docXml, fullImagePath, imgCounter) {
    try {
        const imageBuffer = fs.readFileSync(fullImagePath);
        const ext         = path.extname(fullImagePath).toLowerCase();
        
        const extMap = { '.jpg': 'jpeg', '.jpeg': 'jpeg', '.png': 'png', '.gif': 'gif', '.bmp': 'bmp' };
        const extNorm = extMap[ext] || 'jpeg';
        const ctMap   = { jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', bmp: 'image/bmp' };
        const contentType = ctMap[extNorm] || 'image/jpeg';
        
        // Lebar target: 8 cm = 2 880 000 EMU
        // const targetWidthEmu  = 2880000;
        const targetWidthEmu  = 1500000;
        const targetHeightEmu = 900000; // untuk landscape
        // 
        // default fallback: 6 cm
        
        
        if (sizeOf) {
            try {
                const dims = sizeOf(imageBuffer);
                if (dims.width && dims.height) {
                    targetHeightEmu = Math.round(targetWidthEmu * (dims.height / dims.width));
                }
            } catch (_) {}
        }
        
        // Tambahkan berkas gambar ke dalam zip (word/media/)
        const mediaName = `image_${imgCounter}${ext}`;
        const mediaPath = `word/media/${mediaName}`;
        zip.file(mediaPath, imageBuffer);
        
        // Tambahkan relationship ke word/_rels/document.xml.rels
        const relsPath = 'word/_rels/document.xml.rels';
        let relsXml    = zip.files[relsPath].asText();
        const rId      = `rIdImg${imgCounter}`;
        const newRel   =
        `<Relationship Id="${rId}" ` +
        `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ` +
        `Target="media/${mediaName}"/>`;
        relsXml = relsXml.replace('</Relationships>', newRel + '</Relationships>');
        zip.file(relsPath, relsXml);
        
        // Tambahkan Content-Type jika belum ada
        const ctPath = '[Content_Types].xml';
        let   ctXml  = zip.files[ctPath].asText();
        if (!ctXml.includes(`Extension="${extNorm}"`)) {
            const newDefault = `<Default Extension="${extNorm}" ContentType="${contentType}"/>`;
            ctXml = ctXml.replace('</Types>', newDefault + '</Types>');
            zip.file(ctPath, ctXml);
        }
        
        // XML drawing (inline image) — setara python-docx add_picture internal XML
        const drawingXml =
        `<w:p><w:r><w:drawing>` +
        `<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
        `distT="0" distB="0" distL="0" distR="0">` +
        `<wp:extent cx="${targetWidthEmu}" cy="${targetHeightEmu}"/>` +
        `<wp:docPr id="${imgCounter}" name="Picture ${imgCounter}"/>` +
        `<wp:cNvGraphicFramePr>` +
        `<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>` +
        `</wp:cNvGraphicFramePr>` +
        `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
        `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
        `<pic:nvPicPr>` +
        `<pic:cNvPr id="${imgCounter}" name="Picture ${imgCounter}"/>` +
        `<pic:cNvPicPr/>` +
        `</pic:nvPicPr>` +
        `<pic:blipFill>` +
        `<a:blip r:embed="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>` +
        `<a:stretch><a:fillRect/></a:stretch>` +
        `</pic:blipFill>` +
        `<pic:spPr>` +
        `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${targetWidthEmu}" cy="${targetHeightEmu}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `</pic:spPr>` +
        `</pic:pic>` +
        `</a:graphicData>` +
        `</a:graphic>` +
        `</wp:inline>` +
        `</w:drawing></w:r></w:p>`;
        
        return docXml.replace(/<\/w:body>/, drawingXml + '</w:body>');
        
    } catch (err) {
        console.log(`  (-) Gagal menambahkan gambar ${path.basename(fullImagePath)}: ${err.message}`);
        return docXml;
    }
}

/**
* Setara Python: create_single_report_for_day(hari_tanggal_laporan, data_records_for_day, global_config)
*/
function createSingleReportForDay(hariTanggalLaporan, dataRecordsForDay, globalConfig) {
    console.log(`\n${'='.repeat(20)} MEMBUAT LAPORAN DOCX UNTUK: ${hariTanggalLaporan} ${'='.repeat(20)}`);
    
    try {
        const dtObject = parseDateIndonesian(hariTanggalLaporan);
        
        // Setara: month_name = dt_object.strftime("%B").lower()
        const monthName = getMonthNameIDLower(dtObject);
        
        // Setara: output_dir_for_month = os.path.join(OUTPUT_BASE_DIR_DOCX, month_name)
        //         os.makedirs(output_dir_for_month, exist_ok=True)
        const outputDirForMonth = path.join(OUTPUT_BASE_DIR_DOCX, monthName);
        fs.mkdirSync(outputDirForMonth, { recursive: true });
        
        const semesterName = (globalConfig.semester || {}).name || 'N/A';
        
        // Setara: output_filename = f"{dt_object.strftime('%Y_%m_%d_%A')}.docx"
        const outputFilename = `${formatFilenameDate(dtObject)}.docx`;
        const outputPath     = path.join(outputDirForMonth, outputFilename);
        
        console.log(`Akan menyimpan laporan ke: ${outputPath}`);
        
        if (!fs.existsSync(TEMPLATE_PATH_DOCX)) {
            console.log(`!!! ERROR: Template DOCX '${TEMPLATE_PATH_DOCX}' tidak ditemukan.`);
            return;
        }
        
        // Setara: tpl = DocxTemplate(TEMPLATE_PATH_DOCX)
        const templateBuf = fs.readFileSync(TEMPLATE_PATH_DOCX);
        const zip         = new PizZip(templateBuf);
        let renderedZip   = zip;
        let docXml        = zip.files['word/document.xml'].asText();
        
        try {
            const tpl = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
            tpl.render({
                hari_tanggal:           hariTanggalLaporan,
                tahun_periode_semester: semesterName,
            });
            renderedZip = tpl.getZip();
            docXml      = renderedZip.files['word/document.xml'].asText();
        } catch (templateErr) {
            console.log(`  [!] Peringatan: Template tersendat (Word metadata issue), menggunakan sistem perbaikan otomatis.`);
            // Fallback: Ganti placeholder secara manual di XML jika Docxtemplater gagal
            // Gunakan regex non-greedy yang mencakup kemungkinan tag XML di tengah variabel
            docXml = docXml
            .replace(/\{\{.*?hari_tanggal.*?\}\}/g, escapeXml(hariTanggalLaporan))
            .replace(/\{\{.*?tahun_periode_semester.*?\}\}/g, escapeXml(semesterName));
        }
        
        // Setara: table_anchor_paragraph = next((p for p in doc.paragraphs if TABLE_ANCHOR_PLACEHOLDER in p.text), None)
        const tableXml = buildTableXml(dataRecordsForDay);
        docXml = replaceAnchorParagraphWithTable(docXml, TABLE_ANCHOR_PLACEHOLDER, tableXml);
        
        // Setara: Penambahan Lampiran (gambar)
        let imageCounter = 0;
        for (const record of dataRecordsForDay) {
            const imagePathStr = record.imagePath;
            if (imagePathStr) {
                const fullImagePath = path.join(BASE_DIR, '..', imagePathStr);
                if (fs.existsSync(fullImagePath)) {
                    imageCounter++;
                    const captionText =
                    `${imageCounter}. ${record.mataKuliah || 'N/A'} ` +
                    `${record.kelas || ''} (${record.jam || ''})`;
                    
                    docXml = appendParagraphXml(docXml, captionText);
                    docXml = appendImageToDocXml(renderedZip, docXml, fullImagePath, imageCounter);
                    console.log(`  (+) Menambahkan gambar: ${path.basename(fullImagePath)}`);
                }
            }
        }
        
        // Simpan XML yang sudah dimodifikasi kembali ke zip
        renderedZip.file('word/document.xml', docXml);
        
        // Simpan file
        const outputBuffer = renderedZip.generate({ type: 'nodebuffer' });
        fs.writeFileSync(outputPath, outputBuffer);
        console.log(`--- Berhasil membuat laporan DOCX: ${outputPath} ---`);
        
    } catch (err) {
        console.log(`!!! KESALAHAN FATAL saat membuat laporan DOCX untuk ${hariTanggalLaporan}: ${err.message}`);
        // Setara: traceback.print_exc()
        console.error(err.stack);
    }
}

/**
* Normalize date format to ensure consistent grouping.
* Convert "Senin, 4 Mei 2026" to "Senin, 04 Mei 2026"
*/
function normalizeDateFormat(dateStr) {
    if (!dateStr) return dateStr;
    // Replace ", X " with ", 0X " (e.g., ", 4 " => ", 04 ")
    return dateStr.replace(/,\s(\d)\s/, ', 0$1 ');
}

/**
* Setara Python: generate_all_docx_reports(all_comprehensive_records, global_config)
*/
function generateAllDocxReports(allComprehensiveRecords, globalConfig) {
    if (!allComprehensiveRecords || allComprehensiveRecords.length === 0) {
        console.log('Tidak ada data komprehensif untuk diproses menjadi DOCX.');
        return;
    }
    
    // Setara: grouped_by_date = defaultdict(list)
    const groupedByDate = {};
    for (const record of allComprehensiveRecords) {
        let dateKey = record.hariTanggal;
        if (dateKey) {
            // Normalize date format to handle both "Senin, 4 Mei" and "Senin, 04 Mei"
            dateKey = normalizeDateFormat(dateKey);
            if (!groupedByDate[dateKey]) groupedByDate[dateKey] = [];
            groupedByDate[dateKey].push(record);
        }
    }
    
    // Setara: for hari_tanggal, records_for_day in grouped_by_date.items():
    for (const [hariTanggal, recordsForDay] of Object.entries(groupedByDate)) {
        createSingleReportForDay(hariTanggal, recordsForDay, globalConfig);
    }
}

// =============================================================================
// == BAGIAN PEMBUATAN BERITA ACARA (XLSX) ==
// =============================================================================

/**
* Setara Python: create_single_berita_acara(class_key, records_for_class, global_config)
*/
async function createSingleBeritaAcara(classKey, recordsForClass, globalConfig) {
    const [mataKuliah, kelas] = classKey;
    console.log(
        `\n${'='.repeat(20)} MEMBUAT BERITA ACARA UNTUK: ${mataKuliah} - Kelas ${kelas} ${'='.repeat(20)}`
    );
    
    // Setara: workbook = openpyxl.load_workbook(TEMPLATE_PATH_XLSX)
    if (!fs.existsSync(TEMPLATE_PATH_XLSX)) {
        console.log(`!!! ERROR: Template Berita Acara '${TEMPLATE_PATH_XLSX}' tidak ditemukan.`);
        return;
    }
    
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(TEMPLATE_PATH_XLSX);
    } catch (err) {
        console.log(`!!! ERROR: Template Berita Acara '${TEMPLATE_PATH_XLSX}' tidak ditemukan.`);
        return;
    }
    
    // Setara: sheet = workbook.active
    const sheet = workbook.worksheets[0];
    
    const firstRecord  = recordsForClass[0];
    const semesterName = (globalConfig.semester || {}).name || 'N/A';
    
    // Setara: sheet['E5'] = first_record.get('mataKuliah', 'N/A')
    sheet.getCell('E5').value = firstRecord.mataKuliah || 'N/A';
    sheet.getCell('E6').value = firstRecord.sks        || 'N/A';
    sheet.getCell('E8').value = firstRecord.dosen       || 'N/A';
    
    // Setara: hari_jadwal = first_record.get('day', '')
    //         sheet['E9'] = f"{hari_jadwal}, {first_record.get('jam', '')}".strip(', ')
    const hariJadwal = firstRecord.day || '';
    const jamJadwal  = firstRecord.jam  || '';
    // Python .strip(', ') removes leading/trailing commas and spaces
    const e9Value = `${hariJadwal}, ${jamJadwal}`.replace(/^[, ]+|[, ]+$/g, '');
    sheet.getCell('E9').value = e9Value;
    
    // Setara: sheet['E10'] = ", ".join(first_record.get('asisten', []))
    sheet.getCell('E10').value = (firstRecord.asisten || []).join(', ');
    sheet.getCell('E11').value = semesterName;
    
    const startRow = 17;
    
    // Setara: sorted_records = sorted(records_for_class, key=lambda r: datetime.strptime(r['hariTanggal'], "%A, %d %B %Y"))
    const sortedRecords = [...recordsForClass].sort((a, b) => {
        const da = parseDateIndonesian(a.hariTanggal);
        const db = parseDateIndonesian(b.hariTanggal);
        return da - db;
    });
    
    // Setara: records_with_data = [r for r in sorted_records if r.get('status') != 'belum ada laporan']
    const recordsWithData = sortedRecords.filter(r => r.status !== 'belum ada laporan' && r.status !== 'tidak ada jadwal');
    
    for (let i = 0; i < recordsWithData.length; i++) {
        const record     = recordsWithData[i];
        const currentRow = startRow + i;
        
        // Setara: if current_row > sheet.max_row: break
        if (currentRow > sheet.rowCount) break;
        
        sheet.getCell(`C${currentRow}`).value = i + 1;
        
        // Setara: ruang = "LAB" if record.get('deskripsiKegiatan','').lower() != 'online' else 'TEAMS'
        const deskripsi = (record.deskripsiKegiatan || '').toLowerCase();
        const ruang     = deskripsi !== 'online' ? 'LAB' : 'TEAMS';
        sheet.getCell(`D${currentRow}`).value = `${record.hariTanggal || ''}/${ruang}`;
        
        // Setara: jam_full = record.get('jam',' - '); jam_split = jam_full.split(' - ')
        const jamFull  = record.jam || ' - ';
        const jamSplit = jamFull.split(' - ');
        sheet.getCell(`E${currentRow}`).value = jamSplit.length > 0 ? jamSplit[0].trim() : '';
        sheet.getCell(`F${currentRow}`).value = jamSplit.length > 1 ? jamSplit[1].trim() : '';
        
        sheet.getCell(`G${currentRow}`).value = record.materi || '';
        sheet.getCell(`O${currentRow}`).value = '*';
        sheet.getCell(`K${currentRow}`).value = '*';
    }
    
    // Setara: safe_matkul_name = mata_kuliah.replace('.','').replace('/','_')
    const safeMatkul = mataKuliah.replace(/\./g, '').replace(/\//g, '_');
    const outputFilename = `${safeMatkul}_${kelas}.xlsx`;
    
    // Setara: os.makedirs(OUTPUT_DIR_XLSX, exist_ok=True)
    fs.mkdirSync(OUTPUT_DIR_XLSX, { recursive: true });
    
    const outputPath = path.join(OUTPUT_DIR_XLSX, outputFilename);
    await workbook.xlsx.writeFile(outputPath);
    console.log(`--- Berhasil membuat Berita Acara: ${outputPath} ---`);
}

/**
* Setara Python: generate_all_berita_acara(all_comprehensive_records, global_config)
*/
async function generateAllBeritaAcara(allComprehensiveRecords, globalConfig) {
    if (!allComprehensiveRecords || allComprehensiveRecords.length === 0) {
        console.log('Tidak ada data untuk diproses menjadi Berita Acara.');
        return;
    }
    
    console.log('\n\n>>> MEMULAI PROSES PEMBUATAN SEMUA BERITA ACARA (XLSX) <<<');
    
    // Setara: grouped_by_class = defaultdict(list)
    //         key = (record.get("mataKuliah"), record.get("kelas"))
    //         if all(key): grouped_by_class[key].append(record)
    const groupedByClass = {};
    for (const record of allComprehensiveRecords) {
        const mk  = record.mataKuliah;
        const kls = record.kelas;
        if (mk && kls) { // setara: if all(key)
            const key = `${mk}|||${kls}`;
            if (!groupedByClass[key]) {
                groupedByClass[key] = { classKey: [mk, kls], records: [] };
            }
            groupedByClass[key].records.push(record);
        }
    }
    
    // Setara: for class_key, records in grouped_by_class.items():
    for (const { classKey, records } of Object.values(groupedByClass)) {
        await createSingleBeritaAcara(classKey, records, globalConfig);
    }
}

// =============================================================================
// == BAGIAN PEMBUATAN REKAP KEHADIRAN BULANAN (XLSX) ==
// =============================================================================

const { buildAttendanceGrid, formatJam, CELL_STATUS } = require('../services/attendanceService');

// Status-cell fill colors (data-driven, tidak mengubah warna template)
const STATUS_FILLS = {
    sunday:    { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B82F6' } },
    tambahan:  { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBBF24' } },
    diganti:   { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92400E' } },
    pengganti: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5E7EB' } },
    libur:     { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEF4444' } },
};

const DATA_START = 19;   // Row 19 = first data row (TOTAL row gets pushed down)
const COL_E      = 5;    // Column E = Day 1  (template has E, F as placeholders)

/**
* Generate Rekap Kehadiran XLSX for one month using template.
*/
async function generateRekapKehadiranForMonth(month, year, globalConfig) {
    const gridData = buildAttendanceGrid(month, year);
    if (!gridData || !gridData.assistants.length) {
        console.log(`  [!] Tidak ada data asisten untuk bulan ${month}/${year}. Skip.`);
        return;
    }
    
    const { monthInfo, assistants, grandTotalHari, grandTotalJam } = gridData;
    const daysInMonth = monthInfo.daysInMonth;
    const semesterName = (globalConfig.semester || {}).name || 'N/A';
    
    // Dynamic column positions (after inserting extra day columns)
    // Template has cols: A(1)=NO, B(2)=NAMA, C(3)=NIM, D(4)=MK, E(5)=day1, F(6)=day2, G(7)=TOTAL_HARI, H(8)=TOTAL_JAM
    // We need to INSERT (daysInMonth - 2) extra columns after E to make E..E+daysInMonth-1 = day columns
    const EXTRA_COLS     = daysInMonth - 2;  // template already has 2 day cols (E, F)
    const COL_TOTAL_HARI = COL_E + daysInMonth;     // after last day col
    const COL_TOTAL_JAM  = COL_TOTAL_HARI + 1;
    
    // 1. Baca template
    if (!fs.existsSync(TEMPLATE_REKAP_PATH)) {
        console.log(`!!! ERROR: Template '${TEMPLATE_REKAP_PATH}' tidak ditemukan.`);
        return;
    }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(TEMPLATE_REKAP_PATH);
    const ws = wb.worksheets[0];
    
    console.log(`\n--- Membuat Rekap Kehadiran: ${monthInfo.monthName} ${year} ---`);
    
    // 2. Capture dynamic values from template before clearing merges
    const skValue = ws.getCell(13, 1).value || '';
    
    // 2b. Clear ALL existing merges
    const allMergeKeys = Object.keys(ws._merges || {});
    for (const key of allMergeKeys) {
        try { ws.unMergeCells(key); } catch (e) { /* ignore */ }
    }
    if (ws._merges) {
        for (const key of Object.keys(ws._merges)) {
            delete ws._merges[key];
        }
    }
    
    // 2b. Remove existing images to prevent overlap
    // Note: ws.getImages() returns current images, but ExcelJS doesn't have a direct "remove"
    // However, we can re-initialize the header area.
    
    // 3. Update column widths
    ws.getColumn(1).width = 5;    // NO
    ws.getColumn(2).width = 28;   // NAMA
    ws.getColumn(3).width = 14;   // NIM
    ws.getColumn(4).width = 42;   // MATA KULIAH
    for (let d = 0; d < daysInMonth; d++) {
        ws.getColumn(COL_E + d).width = 5.5;
    }
    ws.getColumn(COL_TOTAL_HARI).width = 12;
    ws.getColumn(COL_TOTAL_JAM).width = 10;
    
    const lastCol = COL_TOTAL_JAM;
    
    // 4. Build Logo Areas (Merge 1,1 to 9,2 for ITPLN and 1,lastCol-1 to 9,lastCol for SAQ)
    ws.mergeCells(3, 1, 5, 2);
    ws.mergeCells(3, lastCol - 1, 5, lastCol);
    
    // Add ITPLN Logo
    if (fs.existsSync(LOGO_ITPLN_PATH)) {
        const logoId = wb.addImage({
            filename: LOGO_ITPLN_PATH,
            extension: 'png', // ExcelJS might treat webp as png/jpeg buffer internally if extension is set carefully
        });
        ws.addImage(logoId, {
            tl: { col: 0, row: 0 },
            br: { col: 2, row: 9 },
            editAs: 'oneCell'
        });
    }
    
    // Add SAQ Logo
    if (fs.existsSync(LOGO_SAQ_PATH)) {
        const logoId = wb.addImage({
            filename: LOGO_SAQ_PATH,
            extension: 'png',
        });
        ws.addImage(logoId, {
            tl: { col: lastCol - 2, row: 0 },
            br: { col: lastCol, row: 9 },
            editAs: 'oneCell'
        });
    }
    
    // 5. Re-merge Center Institution Header Rows
    
    for (let r = 3; r <= 5; r++) {
        let rowVal = '';
        for (let c = 1; c <= 8; c++) {
            const v = ws.getCell(r, c).value;
            if (v !== null && v !== undefined && v !== '') {
                const s = String(typeof v === 'object' ? JSON.stringify(v) : v);
                if (s.length > String(rowVal).length) {
                    rowVal = v;
                }
            }
        }
        
        // Clear side columns text if any
        ws.getCell(r, 1).value = '';
        ws.getCell(r, 2).value = '';
        ws.getCell(r, lastCol - 1).value = '';
        ws.getCell(r, lastCol).value = '';
        
        const cellCenter = ws.getCell(r, 3);
        cellCenter.value = rowVal;
        ws.mergeCells(r, 3, r, lastCol - 2);
        cellCenter.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cellCenter.font = { name: 'Times New Roman', size: 24, bold: true };
    }
    
    
    for (let r = 6; r <= 9; r++) {
        let rowVal = '';
        for (let c = 1; c <= 8; c++) {
            const v = ws.getCell(r, c).value;
            if (v !== null && v !== undefined && v !== '') {
                const s = String(typeof v === 'object' ? JSON.stringify(v) : v);
                if (s.length > String(rowVal).length) {
                    rowVal = v;
                }
            }
        }
        
        // Clear side columns text if any
        ws.getCell(r, 1).value = '';
        ws.getCell(r, 2).value = '';
        ws.getCell(r, lastCol - 1).value = '';
        ws.getCell(r, lastCol).value = '';
        
        const cellCenter = ws.getCell(r, 3);
        cellCenter.value = rowVal;
        ws.mergeCells(r, 3, r, lastCol - 2);
        cellCenter.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cellCenter.font = { name: 'Calibri', size: 10, bold: true };
    }
    
    // Row 10: Full width Blue Bar
    ws.mergeCells(10, 1, 10, lastCol);
    const blueRowStyle = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4F81BD' }
    };
    for(let c = 1; c <= lastCol; c++) {
        ws.getCell(10, c).fill = blueRowStyle;
    }
    
    // Rows 11-14: Titles (Full width)
    for (let r = 11; r <= 14; r++) {
        ws.mergeCells(r, 1, r, lastCol);
        ws.getCell(r, 1).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(r, 1).font = { name: 'Calibri', size: 11, bold: true };
    }
    ws.getCell(11, 1).value = 'REKAPITULASI KEHADIRAN ASISTEN LABORATORIUM SOFTWARE ARCHITECTURE & QUALITY';
    ws.getCell(12, 1).value = `SEMESTER ${semesterName.toUpperCase()}`;
    ws.getCell(13, 1).value = skValue || 'SESUAI SK NOMOR : ...........................................'; 
    ws.getCell(14, 1).value = `PERIODE BULAN : ${monthInfo.monthName.toUpperCase()} ${year}`;
    
    // Row 16-18: Column headers
    ws.mergeCells(16, 1, 18, 1);  // NO
    ws.getCell(16, 1).value = 'NO';
    ws.mergeCells(16, 2, 18, 2);  // NAMA
    ws.getCell(16, 2).value = 'NAMA MAHASISWA';
    ws.mergeCells(16, 3, 18, 3);  // NIM
    ws.getCell(16, 3).value = 'NIM';
    ws.mergeCells(16, 4, 18, 4);  // MK
    ws.getCell(16, 4).value = 'MATA KULIAH';
    
    // Day group header
    ws.mergeCells(16, COL_E, 16, COL_E + daysInMonth - 1);
    ws.getCell(16, COL_E).value = 'BULAN, TANGGAL DAN JAM HADIR PELAKSANAAN (DIISI JUMLAH JAM PER HARI)';
    
    // Month name sub-header
    ws.mergeCells(17, COL_E, 17, COL_E + daysInMonth - 1);
    ws.getCell(17, COL_E).value = monthInfo.monthName.toUpperCase();
    
    // Style Constants
    const centerAlign = { horizontal: 'center', vertical: 'middle', wrapText: true };
    const dataFont    = { name: 'Calibri', size: 9 };
    const boldFont    = { name: 'Calibri', size: 9, bold: true };
    const thinBorder  = {
        top:    { style: 'thin', color: { argb: 'FFB0B0B0' } },
        bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
        left:   { style: 'thin', color: { argb: 'FFB0B0B0' } },
        right:  { style: 'thin', color: { argb: 'FFB0B0B0' } },
    };
    
    // TOTAL HARI header
    const totalHeaderStyle = {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } },
        alignment: { ...centerAlign, wrapText: true },
        font: boldFont,
        border: thinBorder
    };
    
    ws.mergeCells(16, COL_TOTAL_HARI, 18, COL_TOTAL_HARI);
    const thTitleCell = ws.getCell(16, COL_TOTAL_HARI);
    thTitleCell.value = 'TOTAL JUMLAH HARI KEHADIRAN';
    thTitleCell.style = totalHeaderStyle;
    
    // TOTAL JAM header
    ws.mergeCells(16, COL_TOTAL_JAM, 18, COL_TOTAL_JAM);
    const tjTitleCell = ws.getCell(16, COL_TOTAL_JAM);
    tjTitleCell.value = 'TOTAL JAM HADIR';
    tjTitleCell.style = totalHeaderStyle;
    
    // Row 18: Day numbers
    for (let d = 0; d < daysInMonth; d++) {
        ws.getCell(18, COL_E + d).value = d + 1;
    }
    
    // 6. Calculate total data rows and insert them
    let totalDataRows = 0;
    assistants.forEach(asst => {
        totalDataRows += Math.max(asst.mataKuliahRows.length, 1);
    });
    const emptyRows = Array(totalDataRows).fill([]);
    ws.spliceRows(DATA_START, 0, ...emptyRows);
    
    // 7. Fill data rows
    let currentRow = DATA_START;
    
    for (let aIdx = 0; aIdx < assistants.length; aIdx++) {
        const asst = assistants[aIdx];
        const rowCount = Math.max(asst.mataKuliahRows.length, 1);
        const startRow = currentRow;
        const endRow   = currentRow + rowCount - 1;
        
        // Merge NO, NAMA, NIM, TOTAL JAM for multi-course assistants
        if (rowCount > 1) {
            ws.mergeCells(startRow, 1, endRow, 1);
            ws.mergeCells(startRow, 2, endRow, 2);
            ws.mergeCells(startRow, 3, endRow, 3);
            ws.mergeCells(startRow, COL_TOTAL_JAM, endRow, COL_TOTAL_JAM);
        }
        
        // NO
        const noCell = ws.getCell(startRow, 1);
        noCell.value = aIdx + 1;  noCell.font = boldFont;
        noCell.alignment = centerAlign;  noCell.border = thinBorder;
        
        // NAMA
        const namaCell = ws.getCell(startRow, 2);
        namaCell.value = asst.nama;  namaCell.font = { ...dataFont, bold: true };
        namaCell.alignment = { vertical: 'middle', wrapText: true };
        namaCell.border = thinBorder;
        
        // NIM
        const nimCell = ws.getCell(startRow, 3);
        nimCell.value = asst.nim;  nimCell.font = dataFont;
        nimCell.alignment = centerAlign;  nimCell.border = thinBorder;
        
        // TOTAL JAM HADIR (merged)
        const tjCell = ws.getCell(startRow, COL_TOTAL_JAM);
        tjCell.value = asst.totalJam > 0 ? Math.round(asst.totalJam) : 0;
        tjCell.font = boldFont;  tjCell.alignment = centerAlign;
        tjCell.border = thinBorder;
        
        // Per mata kuliah rows
        for (let mkIdx = 0; mkIdx < rowCount; mkIdx++) {
            const r = startRow + mkIdx;
            const mkRow = asst.mataKuliahRows[mkIdx];
            
            // MATA KULIAH
            const mkCell = ws.getCell(r, 4);
            mkCell.value = mkRow ? `${mkRow.name} (${mkRow.kelas})` : '';
            mkCell.font  = (mkRow && mkRow.isTambahan)
            ? { ...dataFont, color: { argb: 'FF92400E' }, bold: true }
            : dataFont;
            mkCell.alignment = { vertical: 'middle', wrapText: true };
            mkCell.border = thinBorder;
            
            // Day columns
            for (let d = 0; d < daysInMonth; d++) {
                const dayInfo  = monthInfo.days[d];
                const col      = COL_E + d;
                const cell     = ws.getCell(r, col);
                const cellData = mkRow
                ? (mkRow.cells[dayInfo.date] || { jam: 0, status: 'empty' })
                : { jam: 0, status: dayInfo.isSunday ? 'sunday' : 'empty' };
                
                if (cellData.jam > 0) {
                    cell.value  = parseFloat(formatJam(cellData.jam));
                    cell.numFmt = '0.00';
                }
                
                const fill = STATUS_FILLS[cellData.status];
                if (fill) cell.fill = fill;
                
                if (cellData.status === 'diganti')
                    cell.font = { ...dataFont, color: { argb: 'FFFEF3C7' }, bold: true };
                else if (cellData.status === 'tambahan')
                    cell.font = { ...dataFont, color: { argb: 'FF92400E' }, bold: true };
                else if (cellData.status === 'pengganti')
                    cell.font = { ...dataFont, bold: true };
                else if (cellData.jam > 0)
                    cell.font = { ...dataFont, color: { argb: 'FF1E3A5F' }, bold: true };
                else
                    cell.font = dataFont;
                
                cell.alignment = centerAlign;
                cell.border    = thinBorder;
            }
            
            // TOTAL HARI per course
            const thCell = ws.getCell(r, COL_TOTAL_HARI);
            thCell.value = (mkRow && mkRow.totalHari > 0) ? mkRow.totalHari : '';
            thCell.font  = boldFont;  thCell.alignment = centerAlign;
            thCell.border = thinBorder;
            
            // Sub-row borders for merged cells
            if (mkIdx > 0) {
                ws.getCell(r, 1).border = thinBorder;
                ws.getCell(r, 2).border = thinBorder;
                ws.getCell(r, 3).border = thinBorder;
                ws.getCell(r, COL_TOTAL_JAM).border = thinBorder;
            }
        }
        
        currentRow = endRow + 1;
    }
    
    // 8. Update TOTAL row (pushed down by spliceRows)
    const totalRowIdx = DATA_START + totalDataRows;
    ws.getCell(totalRowIdx, COL_TOTAL_HARI).value = grandTotalHari;
    ws.getCell(totalRowIdx, COL_TOTAL_JAM).value = grandTotalJam > 0 ? Math.round(grandTotalJam) : 0;
    
    // 9. Save
    fs.mkdirSync(OUTPUT_DIR_REKAP, { recursive: true });
    const outputFilename = `Rekap_Kehadiran_${monthInfo.monthName}_${year}.xlsx`;
    const outputPath = path.join(OUTPUT_DIR_REKAP, outputFilename);
    await wb.xlsx.writeFile(outputPath);
    console.log(`--- Berhasil membuat Rekap Kehadiran: ${outputPath} ---`);
}

/**
* Generate all Rekap Kehadiran XLSX files for all semester months.
*/
async function generateAllRekapKehadiran(globalConfig) {
    const semesterInfo = globalConfig.semester || {};
    const months = semesterInfo.months || [];
    const year   = semesterInfo.year;
    
    if (!year || !months.length) {
        console.log('[REKAP KEHADIRAN] Info semester tidak lengkap. Skip.');
        return;
    }
    
    console.log(`\n\n>>> MEMULAI PROSES PEMBUATAN REKAP KEHADIRAN (XLSX) <<<`);
    console.log(`    Semester: ${semesterInfo.name || 'N/A'} | Months: [${months.join(', ')}] | Year: ${year}`);
    
    for (const month of months) {
        await generateRekapKehadiranForMonth(month, year, globalConfig);
    }
    
    console.log('>>> REKAP KEHADIRAN SELESAI.\n');
}


// =============================================================================
// == FUNGSI UTAMA ==
// =============================================================================

/**
* Setara Python: main()
*/
async function main() {
    // Setara: set_locale()  → tidak diperlukan karena kita handle manual
    console.log("Locale diatur secara manual ke Bahasa Indonesia.");
    
    // Ensure MongoDB connection + repository cache are ready
    await connectDB();
    await initRepositories();
    
    // 1. Bangun struktur data komprehensif dari file-file JSON
    const comprehensiveData = buildComprehensiveListWithPlaceholders();
    const globalConfig      = globalConfigRepo.read();
    
    if (!comprehensiveData || !globalConfig) {
        console.log('Proses dihentikan karena data inti gagal dibangun/dimuat.');
        return;
    }
    
    // 2. Hasilkan Laporan Harian (DOCX)
    generateAllDocxReports(comprehensiveData, globalConfig);
    
    // 3. Hasilkan Berita Acara (XLSX)
    await generateAllBeritaAcara(comprehensiveData, globalConfig);
    
    // 4. Hasilkan Rekap Kehadiran Bulanan (XLSX)
    await generateAllRekapKehadiran(globalConfig);
    
    console.log('\n\n>>> SEMUA PROSES SELESAI. <<<');
}

// Setara: if __name__ == "__main__": main()
if (require.main === module) {
    main().catch(err => {
        console.error('Unhandled error:', err);
        process.exit(1);
    });
}

module.exports = {
    generateAllRekapKehadiran,
    loadJsonData,
    GLOBAL_CONFIG_PATH
};