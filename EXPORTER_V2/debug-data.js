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

async function debugData() {
    console.log('\n=== DEBUG: CHECKING DATA FOR APRIL vs MAY ===\n');

    await connectDB();
    await initRepositories();

    const globalConfig = globalConfigRepo.read();
    const schedules = scheduleRepo.read();
    const reports = reportRepo.read();
    const assistants = assistantRepo.read();

    console.log('\n--- GLOBAL CONFIG ---');
    console.log('Semester:', globalConfig.semester);

    console.log('\n--- SCHEDULES COUNT ---');
    console.log(`Total schedules: ${schedules.length}`);
    
    // Group by day
    const schedByDay = {};
    schedules.forEach(s => {
        if (!schedByDay[s.day]) schedByDay[s.day] = [];
        schedByDay[s.day].push(s);
    });
    console.log('Schedules by day:', Object.keys(schedByDay));

    console.log('\n--- REPORTS: APRIL vs MAY ---');
    const aprilReports = reports.filter(r => r.date && r.date.includes('2026-04'));
    const mayReports = reports.filter(r => r.date && r.date.includes('2026-05'));
    
    console.log(`April reports: ${aprilReports.length}`);
    if (aprilReports.length > 0) {
        console.log(`  Sample:`, aprilReports[0]);
        console.log(`  imagePath exists?`, aprilReports[0].imagePath ? 'YES' : 'NO');
        console.log(`  attending_assistant_ids:`, aprilReports[0].attending_assistant_ids);
    }

    console.log(`\nMay reports: ${mayReports.length}`);
    if (mayReports.length > 0) {
        console.log(`  Sample:`, mayReports[0]);
        console.log(`  imagePath exists?`, mayReports[0].imagePath ? 'YES' : 'NO');
        console.log(`  attending_assistant_ids:`, mayReports[0].attending_assistant_ids);
    } else {
        console.log('  >>> EMPTY! No May reports found.');
    }

    console.log('\n--- ASSISTANTS ---');
    console.log(`Total assistants: ${assistants.length}`);
    if (assistants.length > 0) {
        console.log(`  Sample:`, assistants[0]);
    }

    console.log('\n--- BUILDING COMPREHENSIVE LIST FOR COMPARISON ---');
    
    // Simulate: buildComprehensiveListWithPlaceholders()
    const assistantsMap = Object.fromEntries(assistants.map(a => [a.id, a]));
    const coursesMap = Object.fromEntries(courseRepo.read().map(c => [c.id, c]));
    const lecturersMap = Object.fromEntries(lecturerRepo.read().map(l => [l.id, l]));

    const semesterInfo = globalConfig.semester || {};
    const semesterMonths = semesterInfo.months || [];
    const semesterYear = semesterInfo.year;

    console.log(`\nSemester months: [${semesterMonths.join(', ')}]`);
    console.log(`Semester year: ${semesterYear}`);

    // Check April vs May specifically
    [4, 5].forEach(month => {
        const monthName = month === 4 ? 'APRIL' : 'MAY';
        const monthStr = month === 4 ? '2026-04' : '2026-05';
        const reportsForMonth = reports.filter(r => r.date && r.date.includes(monthStr));
        const isInSemester = semesterMonths.includes(month);

        console.log(`\n[${monthName}] Is in semester? ${isInSemester}`);
        console.log(`[${monthName}] Report count: ${reportsForMonth.length}`);
        
        if (reportsForMonth.length > 0) {
            const withImages = reportsForMonth.filter(r => r.imagePath).length;
            const withAsistants = reportsForMonth.filter(r => r.attending_assistant_ids && r.attending_assistant_ids.length > 0).length;
            console.log(`[${monthName}] Reports with imagePath: ${withImages}`);
            console.log(`[${monthName}] Reports with asistants: ${withAsistants}`);
        }
    });

    console.log('\n=== END DEBUG ===\n');
    process.exit(0);
}

debugData().catch(err => {
    console.error('ERROR:', err);
    process.exit(1);
});
