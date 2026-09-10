'use strict';

const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

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

// Copy fungsi dari exporter.js
const DAY_MAP_ID = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];
const MONTH_MAP_ID = [
    'Januari', 'Februari', 'Maret',    'April',   'Mei',      'Juni',
    'Juli',    'Agustus',  'September','Oktober',  'November', 'Desember',
];

function getDayNameID(date) {
    const jsDay = date.getDay();
    const pyWeekday = jsDay === 0 ? 6 : jsDay - 1;
    return DAY_MAP_ID[pyWeekday];
}

function formatDateIndonesian(date) {
    const day = getDayNameID(date);
    const tgl = String(date.getDate()).padStart(2, '0');
    const bulan = MONTH_MAP_ID[date.getMonth()];
    const tahun = date.getFullYear();
    return `${day}, ${tgl} ${bulan} ${tahun}`;
}

function isoDateString(date) {
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
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

function buildComprehensiveListWithPlaceholders() {
    console.log('\n>>> Memulai proses pembangunan daftar laporan komprehensif...');

    const assistantsArr = assistantRepo.read();
    const lecturersArr = lecturerRepo.read();
    const coursesArr = courseRepo.read();
    const schedules = scheduleRepo.read();
    const globalConfig = globalConfigRepo.read();
    const actualReports = reportRepo.read();

    if (!assistantsArr || !lecturersArr || !coursesArr || !schedules || !globalConfig) {
        console.log('!!! Gagal memuat data master.');
        return null;
    }

    const assistants = Object.fromEntries(assistantsArr.map(a => [a.id, a]));
    const lecturers = Object.fromEntries(lecturersArr.map(l => [l.id, l]));
    const courses = Object.fromEntries(coursesArr.map(c => [c.id, c]));

    const semesterInfo = globalConfig.semester || {};
    const semesterYear = semesterInfo.year;
    const semesterMonths = semesterInfo.months || [];

    const reportsMap = {};
    for (const report of actualReports) {
        const key = `${report.schedule_id}__${report.date}`;
        reportsMap[key] = report;
    }

    const allPossibleReports = [];
    const startDate = new Date(semesterYear, semesterMonths[0] - 1, 1);
    const lastMonth = semesterMonths[semesterMonths.length - 1];
    const endDate = new Date(semesterYear, lastMonth, 0);

    const currentDate = new Date(startDate);
    while (currentDate <= endDate) {
        const currentMonth1Based = currentDate.getMonth() + 1;

        if (semesterMonths.includes(currentMonth1Based)) {
            const dayName = getDayNameID(currentDate);
            const isoDate = isoDateString(currentDate);

            const reportsForThisDate = actualReports.filter(r => r.date === isoDate);
            const schedulesForDay = schedules.filter(s => s.day === dayName);

            if (schedulesForDay.length === 0 && reportsForThisDate.length === 0) {
                currentDate.setDate(currentDate.getDate() + 1);
                continue;
            }

            for (const scheduleInfo of schedulesForDay) {
                const scheduleId = scheduleInfo.id;
                const mapKey = `${scheduleId}__${isoDate}`;
                const actualReport = reportsMap[mapKey] || null;

                const courseInfo = courses[scheduleInfo.course_id] || {};
                const lecturerInfo = lecturers[scheduleInfo.lecturer_id] || {};

                let combinedRecord;
                if (actualReport) {
                    combinedRecord = {
                        hariTanggal: actualReport.formattedDate,
                        materi: actualReport.materi,
                        deskripsiKegiatan: actualReport.deskripsiKegiatan,
                        keterangan: actualReport.keterangan,
                        status: actualReport.status || 'hadir',
                        imagePath: actualReport.imagePath,
                        asisten: (actualReport.attending_assistant_ids || [])
                            .map(id => (assistants[id] || {}).panggilan || 'N/A'),
                        mataKuliah: courseInfo.name,
                        kelas: scheduleInfo.class,
                        jam: scheduleInfo.time,
                        sks: courseInfo.sks,
                        dosen: lecturerInfo.name,
                    };
                } else {
                    combinedRecord = {
                        hariTanggal: formatDateIndonesian(currentDate),
                        materi: null,
                        deskripsiKegiatan: null,
                        keterangan: null,
                        status: 'belum ada laporan',
                        imagePath: null,
                        asisten: (scheduleInfo.assistant_ids || [])
                            .map(id => (assistants[id] || {}).panggilan || 'N/A'),
                        mataKuliah: courseInfo.name,
                        kelas: scheduleInfo.class,
                        jam: scheduleInfo.time,
                        sks: courseInfo.sks,
                        dosen: lecturerInfo.name,
                    };
                }
                allPossibleReports.push(combinedRecord);
            }
        }
        currentDate.setDate(currentDate.getDate() + 1);
    }

    return allPossibleReports;
}

async function checkMaySeninOutput() {
    console.log('\n=== TRACING ACTUAL EXPORTER LOGIC ===\n');

    await connectDB();
    await initRepositories();

    const allRecords = buildComprehensiveListWithPlaceholders();
    const globalConfig = globalConfigRepo.read();

    // Group by date (same as generateAllDocxReports) - WITH NORMALIZATION
    const groupedByDate = {};
    for (const record of allRecords) {
        let dateKey = record.hariTanggal;
        if (dateKey) {
            // Normalize date format to handle both "Senin, 4 Mei" and "Senin, 04 Mei"
            dateKey = normalizeDateFormat(dateKey);
            if (!groupedByDate[dateKey]) groupedByDate[dateKey] = [];
            groupedByDate[dateKey].push(record);
        }
    }

    console.log(`Total unique dates: ${Object.keys(groupedByDate).length}`);

    // Find May 4 Senin
    const maySeninKey = 'Senin, 04 Mei 2026';
    if (groupedByDate[maySeninKey]) {
        console.log(`\n[${maySeninKey}]`);
        console.log(`Records count: ${groupedByDate[maySeninKey].length}`);
        groupedByDate[maySeninKey].forEach((record, idx) => {
            console.log(`\n  Record ${idx + 1}:`);
            console.log(`    Mata Kuliah: ${record.mataKuliah}`);
            console.log(`    Kelas: ${record.kelas}`);
            console.log(`    Jam: ${record.jam}`);
            console.log(`    Dosen: ${record.dosen}`);
            console.log(`    Materi: ${record.materi || '(tidak ada)'}`);
            console.log(`    Deskripsi: ${record.deskripsiKegiatan || '(tidak ada)'}`);
            console.log(`    Status: ${record.status}`);
            console.log(`    Asisten: ${JSON.stringify(record.asisten)}`);
            console.log(`    Keterangan: ${record.keterangan || '(tidak ada)'}`);
        });
    } else {
        console.log(`\n!!! Key '${maySeninKey}' NOT FOUND in groupedByDate`);
        console.log('\nAvailable keys containing "Mei":');
        Object.keys(groupedByDate)
            .filter(k => k.includes('Mei'))
            .forEach(k => {
                console.log(`  - ${k} (${groupedByDate[k].length} records)`);
            });
    }

    process.exit(0);
}

checkMaySeninOutput().catch(err => {
    console.error('ERROR:', err);
    process.exit(1);
});
