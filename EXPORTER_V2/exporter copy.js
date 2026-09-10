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
    const assistantsArr = loadJsonData(ASSISTANTS_PATH, 'assistants');
    const lecturersArr  = loadJsonData(LECTURERS_PATH,  'lecturers');
    const coursesArr    = loadJsonData(COURSES_PATH,    'courses');
    const schedules     = loadJsonData(SCHEDULES_PATH,  'schedules');
    const globalConfig  = loadJsonData(GLOBAL_CONFIG_PATH);
    let actualReports = loadJsonData(REPORTS_PATH);
    if (actualReports === null) {
        console.log('--- Warning: reports.json tidak ditemukan atau kosong. Menggunakan daftar kosong. ---');
        actualReports = [];
    }

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
        const targetWidthEmu  = 2880000;
        let   targetHeightEmu = 2160000; // default fallback: 6 cm

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
        const dateKey = record.hariTanggal;
        if (dateKey) {
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

const TEMPLATE_REKAP_PATH = path.join(BASE_DIR, 'Rekap_Ngawas_Lab.xlsx');
const OUTPUT_DIR_REKAP    = path.join(OUTPUT_BASE_DIR_DOCX, 'rekap_kehadiran');

// ── Warna fill per status cell ──────────────────────────────────────────────────
const FILL_COLORS = {
    sunday:    { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B82F6' } },
    normal:    null, // no fill, just value
    tambahan:  { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBBF24' } },
    diganti:   { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92400E' } },
    pengganti: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9CA3AF' } },
    libur:     { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEF4444' } },
};

const HEADER_FILL     = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
const HEADER_FILL_ALT = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F2942' } };
const TOTAL_ROW_FILL  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
const SUNDAY_HEADER   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B82F6' } };

const HEADER_FONT = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
const DATA_FONT   = { name: 'Calibri', size: 9 };
const BOLD_FONT   = { name: 'Calibri', size: 9, bold: true };
const TITLE_FONT  = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF1E3A5F' } };

const THIN_BORDER = {
    top:    { style: 'thin', color: { argb: 'FF94A3B8' } },
    bottom: { style: 'thin', color: { argb: 'FF94A3B8' } },
    left:   { style: 'thin', color: { argb: 'FF94A3B8' } },
    right:  { style: 'thin', color: { argb: 'FF94A3B8' } },
};

/**
 * Apply common cell styling
 */
function styleCell(cell, { font, fill, alignment, border, numFmt } = {}) {
    if (font) cell.font = font;
    if (fill) cell.fill = fill;
    if (alignment) cell.alignment = alignment;
    if (border) cell.border = border;
    if (numFmt) cell.numFmt = numFmt;
}

/**
 * Generate Rekap Kehadiran XLSX for one month.
 */
async function generateRekapKehadiranForMonth(month, year, globalConfig) {
    const gridData = buildAttendanceGrid(month, year);
    if (!gridData || !gridData.assistants.length) {
        console.log(`  [!] Tidak ada data asisten untuk ${gridData?.monthInfo?.monthName || month}/${year}. Skip.`);
        return;
    }

    const { monthInfo, assistants, grandTotalHari, grandTotalJam } = gridData;
    const daysInMonth = monthInfo.daysInMonth;
    const semesterName = (globalConfig.semester || {}).name || 'N/A';

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('REKAPITULASI', {
        views: [{ state: 'frozen', xSplit: 4, ySplit: 18 }],
    });

    // ── Column widths ──────────────────────────────────────────────────────────
    // A=NO(4), B=NAMA(28), C=NIM(14), D=MATA KULIAH(38), E..=Days(5.5 each), AJ=TOTAL HARI(11), AK=TOTAL JAM(9)
    const COL_DAY_START = 5; // Column E
    const COL_TOTAL_HARI = COL_DAY_START + daysInMonth;     // after last day
    const COL_TOTAL_JAM  = COL_TOTAL_HARI + 1;

    ws.getColumn(1).width = 5;   // NO
    ws.getColumn(2).width = 28;  // NAMA
    ws.getColumn(3).width = 14;  // NIM
    ws.getColumn(4).width = 42;  // MATA KULIAH
    for (let d = 0; d < daysInMonth; d++) {
        ws.getColumn(COL_DAY_START + d).width = 5.5;
    }
    ws.getColumn(COL_TOTAL_HARI).width = 12;
    ws.getColumn(COL_TOTAL_JAM).width = 10;

    // ── Header rows 1–10: Logo area (basic text) ──────────────────────────────
    // Row 2-7: Institution text (simplified — no logo image)
    ws.mergeCells(2, 2, 7, COL_TOTAL_JAM - 2);
    const instCell = ws.getCell(2, 2);
    instCell.value = 'SOFTWARE ARCHITECTURE AND\nQUALITY LABORATORY\nINSTITUT TEKNOLOGI PLN JAKARTA\nJalan Lingkar Luar Barat, Duri Kosambi, Cengkareng, Jakarta';
    styleCell(instCell, {
        font: { name: 'Calibri', size: 14, bold: true, color: { argb: 'FF1E3A5F' } },
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
    });

    // Row 9: Blue line separator
    for (let c = 1; c <= COL_TOTAL_JAM; c++) {
        const cell = ws.getCell(9, c);
        cell.border = { bottom: { style: 'medium', color: { argb: 'FF1E3A5F' } } };
    }

    // Row 11-14: Title area
    ws.mergeCells(11, 1, 11, COL_TOTAL_JAM);
    const titleCell = ws.getCell(11, 1);
    titleCell.value = 'REKAPITULASI KEHADIRAN ASISTEN LABORATORIUM SOFTWARE ARCHITECTURE & QUALITY';
    styleCell(titleCell, { font: TITLE_FONT, alignment: { horizontal: 'center' } });

    ws.mergeCells(12, 1, 12, COL_TOTAL_JAM);
    const semCell = ws.getCell(12, 1);
    semCell.value = `SEMESTER ${semesterName.toUpperCase()}`;
    styleCell(semCell, { font: { ...TITLE_FONT, size: 11 }, alignment: { horizontal: 'center' } });

    ws.mergeCells(13, 1, 13, COL_TOTAL_JAM);
    const skCell = ws.getCell(13, 1);
    skCell.value = 'SESUAI SK NOMOR : NO.01/SK/GANJIL/SAQLab/2025';
    styleCell(skCell, { font: { ...TITLE_FONT, size: 10 }, alignment: { horizontal: 'center' } });

    ws.mergeCells(14, 1, 14, COL_TOTAL_JAM);
    const periodeCell = ws.getCell(14, 1);
    periodeCell.value = `PERIODE BULAN : ${monthInfo.monthName.toUpperCase()} ${year}`;
    styleCell(periodeCell, { font: { ...TITLE_FONT, size: 10 }, alignment: { horizontal: 'center' } });

    // ── Row 15: Blank separator line ──────────────────────────────────────────
    ws.mergeCells(15, 1, 15, COL_TOTAL_JAM);
    for (let c = 1; c <= COL_TOTAL_JAM; c++) {
        ws.getCell(15, c).fill = HEADER_FILL;
    }

    // ── Row 16: Group headers ─────────────────────────────────────────────────
    const centerAlign = { horizontal: 'center', vertical: 'middle', wrapText: true };

    // NO (merge 16-18)  
    ws.mergeCells(16, 1, 18, 1);
    const noHeader = ws.getCell(16, 1);
    noHeader.value = 'NO';
    styleCell(noHeader, { font: HEADER_FONT, fill: HEADER_FILL, alignment: centerAlign, border: THIN_BORDER });

    // NAMA MAHASISWA (merge 16-18)
    ws.mergeCells(16, 2, 18, 2);
    const namaHeader = ws.getCell(16, 2);
    namaHeader.value = 'NAMA MAHASISWA';
    styleCell(namaHeader, { font: HEADER_FONT, fill: HEADER_FILL, alignment: centerAlign, border: THIN_BORDER });

    // NIM (merge 16-18)
    ws.mergeCells(16, 3, 18, 3);
    const nimHeader = ws.getCell(16, 3);
    nimHeader.value = 'NIM';
    styleCell(nimHeader, { font: HEADER_FONT, fill: HEADER_FILL, alignment: centerAlign, border: THIN_BORDER });

    // MATA KULIAH (merge 16-18)
    ws.mergeCells(16, 4, 18, 4);
    const mkHeader = ws.getCell(16, 4);
    mkHeader.value = 'MATA KULIAH';
    styleCell(mkHeader, { font: HEADER_FONT, fill: HEADER_FILL, alignment: centerAlign, border: THIN_BORDER });

    // Day group header: "BULAN, TANGGAL DAN JAM HADIR PELAKSANAAN (DIISI JUMLAH JAM PER HARI)" (row 16, merge day cols)
    ws.mergeCells(16, COL_DAY_START, 16, COL_DAY_START + daysInMonth - 1);
    const dayGroupHeader = ws.getCell(16, COL_DAY_START);
    dayGroupHeader.value = 'BULAN, TANGGAL DAN JAM HADIR PELAKSANAAN (DIISI JUMLAH JAM PER HARI)';
    styleCell(dayGroupHeader, { font: HEADER_FONT, fill: HEADER_FILL, alignment: centerAlign, border: THIN_BORDER });

    // TOTAL JUMLAH HARI KEHADIRAN (merge 16-18)
    ws.mergeCells(16, COL_TOTAL_HARI, 18, COL_TOTAL_HARI);
    const totHariHead = ws.getCell(16, COL_TOTAL_HARI);
    totHariHead.value = 'TOTAL JUMLAH\nHARI KEHADIRAN';
    styleCell(totHariHead, { font: HEADER_FONT, fill: HEADER_FILL_ALT, alignment: centerAlign, border: THIN_BORDER });

    // TOTAL JAM HADIR (merge 16-18)
    ws.mergeCells(16, COL_TOTAL_JAM, 18, COL_TOTAL_JAM);
    const totJamHead = ws.getCell(16, COL_TOTAL_JAM);
    totJamHead.value = 'TOTAL JAM\nHADIR';
    styleCell(totJamHead, { font: HEADER_FONT, fill: HEADER_FILL_ALT, alignment: centerAlign, border: THIN_BORDER });

    // ── Row 17: Month name (merge day cols) ───────────────────────────────────
    ws.mergeCells(17, COL_DAY_START, 17, COL_DAY_START + daysInMonth - 1);
    const monthNameHeader = ws.getCell(17, COL_DAY_START);
    monthNameHeader.value = monthInfo.monthName.toUpperCase();
    styleCell(monthNameHeader, { font: HEADER_FONT, fill: HEADER_FILL, alignment: centerAlign, border: THIN_BORDER });

    // ── Row 18: Day numbers (1, 2, 3, ...) ────────────────────────────────────
    for (let d = 0; d < daysInMonth; d++) {
        const dayInfo = monthInfo.days[d];
        const col = COL_DAY_START + d;
        const cell = ws.getCell(18, col);
        cell.value = dayInfo.day;
        const isSun = dayInfo.isSunday;
        styleCell(cell, {
            font: { ...HEADER_FONT, size: 8 },
            fill: isSun ? SUNDAY_HEADER : HEADER_FILL,
            alignment: centerAlign,
            border: THIN_BORDER,
        });
    }

    // ── Data rows (row 19 onward) ─────────────────────────────────────────────
    let currentRow = 19;

    for (let aIdx = 0; aIdx < assistants.length; aIdx++) {
        const asst = assistants[aIdx];
        const rowCount = Math.max(asst.mataKuliahRows.length, 1);
        const startRow = currentRow;
        const endRow   = currentRow + rowCount - 1;

        // Merge NO, NAMA, NIM, TOTAL JAM across all mataKuliah rows
        if (rowCount > 1) {
            ws.mergeCells(startRow, 1, endRow, 1); // NO
            ws.mergeCells(startRow, 2, endRow, 2); // NAMA
            ws.mergeCells(startRow, 3, endRow, 3); // NIM
            ws.mergeCells(startRow, COL_TOTAL_JAM, endRow, COL_TOTAL_JAM); // TOTAL JAM
        }

        // NO
        const noCell = ws.getCell(startRow, 1);
        noCell.value = aIdx + 1;
        styleCell(noCell, { font: BOLD_FONT, alignment: centerAlign, border: THIN_BORDER });

        // NAMA
        const namaCell = ws.getCell(startRow, 2);
        namaCell.value = asst.nama;
        styleCell(namaCell, {
            font: { ...DATA_FONT, bold: true },
            alignment: { vertical: 'middle', wrapText: true },
            border: THIN_BORDER,
        });

        // NIM
        const nimCell = ws.getCell(startRow, 3);
        nimCell.value = asst.nim;
        styleCell(nimCell, { font: DATA_FONT, alignment: centerAlign, border: THIN_BORDER });

        // TOTAL JAM (merged)
        const totalJamCell = ws.getCell(startRow, COL_TOTAL_JAM);
        totalJamCell.value = asst.totalJam > 0 ? Math.round(asst.totalJam) : 0;
        styleCell(totalJamCell, { font: BOLD_FONT, alignment: centerAlign, border: THIN_BORDER });

        // Per mata kuliah row
        for (let mkIdx = 0; mkIdx < rowCount; mkIdx++) {
            const r = startRow + mkIdx;
            const mkRow = asst.mataKuliahRows[mkIdx];

            if (!mkRow) {
                // No courses — empty row
                const mkCell = ws.getCell(r, 4);
                mkCell.value = '';
                styleCell(mkCell, { font: DATA_FONT, border: THIN_BORDER });
                for (let d = 0; d < daysInMonth; d++) {
                    const dayCol = COL_DAY_START + d;
                    const dayInfo = monthInfo.days[d];
                    const cell = ws.getCell(r, dayCol);
                    if (dayInfo.isSunday) cell.fill = FILL_COLORS.sunday;
                    cell.border = THIN_BORDER;
                }
                // TOTAL HARI
                const thCell = ws.getCell(r, COL_TOTAL_HARI);
                thCell.value = 0;
                styleCell(thCell, { font: BOLD_FONT, alignment: centerAlign, border: THIN_BORDER });
                continue;
            }

            // MATA KULIAH
            const mkCell = ws.getCell(r, 4);
            mkCell.value = `${mkRow.name} (${mkRow.kelas})`;
            styleCell(mkCell, {
                font: mkRow.isTambahan
                    ? { ...DATA_FONT, color: { argb: 'FF92400E' }, bold: true }
                    : DATA_FONT,
                alignment: { vertical: 'middle', wrapText: true },
                border: THIN_BORDER,
            });

            // Day columns
            for (let d = 0; d < daysInMonth; d++) {
                const dayInfo = monthInfo.days[d];
                const dayCol  = COL_DAY_START + d;
                const cell    = ws.getCell(r, dayCol);
                const cellData = mkRow.cells[dayInfo.date] || { jam: 0, status: 'empty' };

                // Value
                if (cellData.jam > 0) {
                    cell.value = parseFloat(formatJam(cellData.jam));
                    cell.numFmt = '0.00';
                }

                // Fill color
                const fillDef = FILL_COLORS[cellData.status];
                if (fillDef) {
                    cell.fill = fillDef;
                }

                // Font for diganti (light text on dark bg)
                if (cellData.status === 'diganti') {
                    cell.font = { ...DATA_FONT, color: { argb: 'FFFEF3C7' }, bold: true };
                } else if (cellData.status === 'tambahan') {
                    cell.font = { ...DATA_FONT, color: { argb: 'FF92400E' }, bold: true };
                } else if (cellData.jam > 0) {
                    cell.font = { ...DATA_FONT, bold: true };
                } else {
                    cell.font = DATA_FONT;
                }

                cell.alignment = centerAlign;
                cell.border = THIN_BORDER;
            }

            // TOTAL HARI (per course row)
            const thCell = ws.getCell(r, COL_TOTAL_HARI);
            thCell.value = mkRow.totalHari > 0 ? mkRow.totalHari : '';
            styleCell(thCell, { font: BOLD_FONT, alignment: centerAlign, border: THIN_BORDER });

            // Make borders on sub-rows for merged cells
            if (mkIdx > 0) {
                // Borders for merged cells (they still need borders on sub-rows)
                ws.getCell(r, 1).border = THIN_BORDER;
                ws.getCell(r, 2).border = THIN_BORDER;
                ws.getCell(r, 3).border = THIN_BORDER;
                ws.getCell(r, COL_TOTAL_JAM).border = THIN_BORDER;
            }
        }

        currentRow = endRow + 1;
    }

    // ── TOTAL row ─────────────────────────────────────────────────────────────
    const totalRow = currentRow;
    ws.getCell(totalRow, 1).border = THIN_BORDER;

    ws.mergeCells(totalRow, 2, totalRow, 4);
    const totalLabel = ws.getCell(totalRow, 2);
    totalLabel.value = 'TOTAL';
    styleCell(totalLabel, {
        font: { ...BOLD_FONT, size: 10 },
        fill: TOTAL_ROW_FILL,
        alignment: { horizontal: 'center', vertical: 'middle' },
        border: THIN_BORDER,
    });
    ws.getCell(totalRow, 1).fill = TOTAL_ROW_FILL;

    // Day totals
    for (let d = 0; d < daysInMonth; d++) {
        const dayInfo = monthInfo.days[d];
        const col = COL_DAY_START + d;
        const cell = ws.getCell(totalRow, col);

        let dayTotal = 0;
        assistants.forEach(asst => {
            asst.mataKuliahRows.forEach(mkRow => {
                const c = mkRow.cells[dayInfo.date];
                if (c && c.jam > 0) dayTotal += c.jam;
            });
        });

        if (dayTotal > 0) {
            cell.value = parseFloat(dayTotal.toFixed(2));
            cell.numFmt = '0.00';
        }
        if (dayInfo.isSunday) cell.fill = FILL_COLORS.sunday;
        else cell.fill = TOTAL_ROW_FILL;
        styleCell(cell, { font: BOLD_FONT, alignment: centerAlign, border: THIN_BORDER });
    }

    // Grand total hari
    const gtHariCell = ws.getCell(totalRow, COL_TOTAL_HARI);
    gtHariCell.value = grandTotalHari;
    styleCell(gtHariCell, { font: { ...BOLD_FONT, size: 10 }, fill: TOTAL_ROW_FILL, alignment: centerAlign, border: THIN_BORDER });

    // Grand total jam
    const gtJamCell = ws.getCell(totalRow, COL_TOTAL_JAM);
    gtJamCell.value = grandTotalJam > 0 ? Math.round(grandTotalJam) : 0;
    styleCell(gtJamCell, { font: { ...BOLD_FONT, size: 10 }, fill: TOTAL_ROW_FILL, alignment: centerAlign, border: THIN_BORDER });

    // ── Save ──────────────────────────────────────────────────────────────────
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
        console.log('[REKAP KEHADIRAN] Info semester tidak lengkap di config_global.json. Skip.');
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

    // 1. Bangun struktur data komprehensif dari file-file JSON
    const comprehensiveData = buildComprehensiveListWithPlaceholders();
    const globalConfig      = loadJsonData(GLOBAL_CONFIG_PATH);

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
main().catch(err => {
    console.error('Unhandled error:', err);
    process.exit(1);
});