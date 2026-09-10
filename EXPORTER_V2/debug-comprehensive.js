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

function isoDateString(date) {
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

async function debugComprehensiveList() {
    console.log('\n=== DEBUG: COMPREHENSIVE LIST FOR APRIL vs MAY ===\n');

    await connectDB();
    await initRepositories();

    const globalConfig = globalConfigRepo.read();
    const schedules = scheduleRepo.read();
    const reports = reportRepo.read();
    const assistants = assistantRepo.read();
    const courses = courseRepo.read();
    const lecturers = lecturerRepo.read();

    // Build maps
    const assistantsMap = Object.fromEntries(assistants.map(a => [a.id, a]));
    const coursesMap = Object.fromEntries(courses.map(c => [c.id, c]));
    const lecturersMap = Object.fromEntries(lecturers.map(l => [l.id, l]));

    const semesterInfo = globalConfig.semester || {};
    const semesterYear = semesterInfo.year;
    const semesterMonths = semesterInfo.months || [];

    console.log(`Semester: ${semesterInfo.name}`);
    console.log(`Year: ${semesterYear}, Months: [${semesterMonths.join(', ')}]\n`);

    // Build reportsMap
    const reportsMap = {};
    for (const report of reports) {
        const key = `${report.schedule_id}__${report.date}`;
        reportsMap[key] = report;
    }

    // Check April and May specifically
    [4, 5].forEach(month => {
        const monthName = month === 4 ? 'APRIL' : 'MAY';
        console.log(`\n${'='.repeat(60)}`);
        console.log(`CHECKING ${monthName} (month ${month})`);
        console.log(`${'='.repeat(60)}\n`);

        const startDate = new Date(semesterYear, month - 1, 1);
        const endDate = new Date(semesterYear, month, 0);

        let dateCount = 0;
        let scheduleCount = 0;
        let reportCount = 0;

        const currentDate = new Date(startDate);
        while (currentDate <= endDate) {
            const currentMonth1Based = currentDate.getMonth() + 1;
            
            if (semesterMonths.includes(currentMonth1Based)) {
                const dayName = getDayNameID(currentDate);
                const isoDate = isoDateString(currentDate);
                
                const reportsForThisDate = reports.filter(r => r.date === isoDate);
                const schedulesForDay = schedules.filter(s => s.day === dayName);

                if (schedulesForDay.length > 0 || reportsForThisDate.length > 0) {
                    dateCount++;
                    scheduleCount += schedulesForDay.length;
                    reportCount += reportsForThisDate.length;

                    // Show detail for first few dates
                    if (dateCount <= 3 || reportsForThisDate.length > 0) {
                        console.log(`  ${isoDate} (${dayName})`);
                        console.log(`    - Schedules for this day: ${schedulesForDay.length}`);
                        console.log(`    - Reports for this date: ${reportsForThisDate.length}`);
                        
                        if (reportsForThisDate.length > 0) {
                            reportsForThisDate.forEach((rpt, idx) => {
                                const mkName = schedulesForDay.find(s => s.id === rpt.schedule_id)
                                    ? coursesMap[schedules.find(s => s.id === rpt.schedule_id)?.course_id]?.name 
                                    : 'UNKNOWN';
                                console.log(`      Report ${idx+1}: imagePath=${rpt.imagePath ? '✓' : '✗'}, asistans=${rpt.attending_assistant_ids?.length || 0}`);
                            });
                        }
                    }
                }
            }
            
            currentDate.setDate(currentDate.getDate() + 1);
        }

        console.log(`\n  [SUMMARY]`);
        console.log(`    - Dates with data: ${dateCount}`);
        console.log(`    - Total schedule slots: ${scheduleCount}`);
        console.log(`    - Total actual reports: ${reportCount}`);
    });

    console.log('\n\n=== CHECKING ASISTEN DATA IN REPORTS ===\n');
    
    [4, 5].forEach(month => {
        const monthName = month === 4 ? 'APRIL' : 'MAY';
        const monthStr = month === 4 ? '2026-04' : '2026-05';
        const reportsForMonth = reports.filter(r => r.date && r.date.includes(monthStr));

        console.log(`\n[${monthName}] Sample report asisten mapping:`);
        if (reportsForMonth.length > 0) {
            const sample = reportsForMonth[0];
            console.log(`  Report date: ${sample.date}`);
            console.log(`  attending_assistant_ids: ${JSON.stringify(sample.attending_assistant_ids)}`);
            const asistanNames = (sample.attending_assistant_ids || [])
                .map(id => (assistantsMap[id] || {}).panggilan || 'N/A');
            console.log(`  Resolved names: ${JSON.stringify(asistanNames)}`);
        }
    });

    process.exit(0);
}

debugComprehensiveList().catch(err => {
    console.error('ERROR:', err);
    process.exit(1);
});
