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
    const day   = getDayNameID(date);
    const tgl   = String(date.getDate()).padStart(2, '0');
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

async function debugSchedulesAndBuilding() {
    console.log('\n=== STEP 1: CHECK SCHEDULES FOR SENIN ===\n');

    await connectDB();
    await initRepositories();

    const schedules = scheduleRepo.read();
    const courses = courseRepo.read();
    const lecturers = lecturerRepo.read();
    const assistants = assistantRepo.read();
    const reports = reportRepo.read();

    const coursesMap = Object.fromEntries(courses.map(c => [c.id, c]));
    const lecturersMap = Object.fromEntries(lecturers.map(l => [l.id, l]));
    const assistantsMap = Object.fromEntries(assistants.map(a => [a.id, a]));

    // Filter schedules untuk Senin
    const senin = schedules.filter(s => s.day === 'Senin');
    console.log(`Total schedules untuk SENIN: ${senin.length}\n`);

    senin.forEach((sched, idx) => {
        console.log(`[Schedule ${idx + 1}]`);
        console.log(`  ID: ${sched.id}`);
        console.log(`  Course: ${coursesMap[sched.course_id]?.name || 'N/A'}`);
        console.log(`  Class: ${sched.class}`);
        console.log(`  Time: ${sched.time}`);
        console.log(`  Lecturer: ${lecturersMap[sched.lecturer_id]?.name || 'N/A'}`);
        console.log(`  assistant_ids (from schedule): ${JSON.stringify(sched.assistant_ids || [])}`);
        const asistanNames = (sched.assistant_ids || [])
            .map(id => assistantsMap[id]?.panggilan || 'N/A');
        console.log(`  Resolved asistan names: ${JSON.stringify(asistanNames)}`);
        console.log();
    });

    console.log('\n=== STEP 2: CHECK COMPREHENSIVE LIST BUILDING ===\n');

    const globalConfig = globalConfigRepo.read();
    const semesterInfo = globalConfig.semester || {};
    const semesterYear = semesterInfo.year;
    const semesterMonths = semesterInfo.months || [];

    // Build reportsMap
    const reportsMap = {};
    for (const report of reports) {
        const key = `${report.schedule_id}__${report.date}`;
        reportsMap[key] = report;
    }

    // Check APRIL Senin (2026-04-06)
    console.log(`>>> APRIL SENIN (2026-04-06) <<<`);
    const aprilSeninDate = new Date(2026, 3, 6);
    const aprilSeninIso = isoDateString(aprilSeninDate);
    console.log(`ISO date: ${aprilSeninIso}`);

    senin.forEach((sched, idx) => {
        const mapKey = `${sched.id}__${aprilSeninIso}`;
        const actualReport = reportsMap[mapKey];
        console.log(`\nSchedule ${idx + 1} (${sched.id}):`);
        console.log(`  Report found: ${actualReport ? 'YES' : 'NO'}`);
        if (actualReport) {
            console.log(`    - materi: ${actualReport.materi}`);
            console.log(`    - imagePath: ${actualReport.imagePath}`);
            console.log(`    - attending_assistant_ids: ${JSON.stringify(actualReport.attending_assistant_ids)}`);
        } else {
            console.log(`  PLACEHOLDER would be created with:`);
            const asistanNames = (sched.assistant_ids || [])
                .map(id => assistantsMap[id]?.panggilan || 'N/A');
            console.log(`    - asisten: ${JSON.stringify(asistanNames)}`);
        }
    });

    // Check MAY Senin (2026-05-04)
    console.log(`\n\n>>> MAY SENIN (2026-05-04) <<<`);
    const maySeninDate = new Date(2026, 4, 4);
    const maySeninIso = isoDateString(maySeninDate);
    console.log(`ISO date: ${maySeninIso}`);

    senin.forEach((sched, idx) => {
        const mapKey = `${sched.id}__${maySeninIso}`;
        const actualReport = reportsMap[mapKey];
        console.log(`\nSchedule ${idx + 1} (${sched.id}):`);
        console.log(`  Report found: ${actualReport ? 'YES' : 'NO'}`);
        if (actualReport) {
            console.log(`    - materi: ${actualReport.materi}`);
            console.log(`    - imagePath: ${actualReport.imagePath}`);
            console.log(`    - attending_assistant_ids: ${JSON.stringify(actualReport.attending_assistant_ids)}`);
        } else {
            console.log(`  PLACEHOLDER would be created with:`);
            const asistanNames = (sched.assistant_ids || [])
                .map(id => assistantsMap[id]?.panggilan || 'N/A');
            console.log(`    - asisten: ${JSON.stringify(asistanNames)}`);
        }
    });

    console.log('\n\n=== STEP 3: SIMULATE buildComprehensiveListWithPlaceholders ===\n');

    // Simulate the ACTUAL building logic
    const allPossibleReports = [];
    const startDate = new Date(semesterYear, semesterMonths[0] - 1, 1);
    const lastMonth = semesterMonths[semesterMonths.length - 1];
    const endDate = new Date(semesterYear, lastMonth, 0);

    const currentDate = new Date(startDate);
    let countMaySeninRecords = 0;

    while (currentDate <= endDate) {
        const currentMonth1Based = currentDate.getMonth() + 1;

        if (semesterMonths.includes(currentMonth1Based)) {
            const dayName = getDayNameID(currentDate);
            const isoDate = isoDateString(currentDate);

            // Only check May Senin dates
            if (isoDate.includes('2026-05') && dayName === 'Senin') {
                console.log(`\n[${isoDate} - ${dayName}]`);
                
                const reportsForThisDate = reports.filter(r => r.date === isoDate);
                const schedulesForDay = schedules.filter(s => s.day === dayName);

                console.log(`  - schedulesForDay.length: ${schedulesForDay.length}`);
                console.log(`  - reportsForThisDate.length: ${reportsForThisDate.length}`);

                if (schedulesForDay.length === 0 && reportsForThisDate.length === 0) {
                    console.log(`  >>> SKIPPED (no schedule AND no report)`);
                    currentDate.setDate(currentDate.getDate() + 1);
                    continue;
                }

                for (const scheduleInfo of schedulesForDay) {
                    const scheduleId = scheduleInfo.id;
                    const mapKey = `${scheduleId}__${isoDate}`;
                    const actualReport = reportsMap[mapKey] || null;

                    const courseInfo = coursesMap[scheduleInfo.course_id] || {};
                    const lecturerInfo = lecturersMap[scheduleInfo.lecturer_id] || {};

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
                                .map(id => (assistantsMap[id] || {}).panggilan || 'N/A'),
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
                                .map(id => (assistantsMap[id] || {}).panggilan || 'N/A'),
                            mataKuliah: courseInfo.name,
                            kelas: scheduleInfo.class,
                            jam: scheduleInfo.time,
                            sks: courseInfo.sks,
                            dosen: lecturerInfo.name,
                        };
                    }
                    
                    allPossibleReports.push(combinedRecord);
                    countMaySeninRecords++;
                    
                    console.log(`  [Record ${countMaySeninRecords}] ${combinedRecord.mataKuliah} ${combinedRecord.kelas}`);
                    console.log(`    - status: ${combinedRecord.status}`);
                    console.log(`    - asisten: ${JSON.stringify(combinedRecord.asisten)}`);
                }
            }
        }

        currentDate.setDate(currentDate.getDate() + 1);
    }

    console.log(`\n\nTotal May Senin records built: ${countMaySeninRecords}`);
    console.log('Expected: 3 (like April Senin)');

    process.exit(0);
}

debugSchedulesAndBuilding().catch(err => {
    console.error('ERROR:', err);
    process.exit(1);
});
