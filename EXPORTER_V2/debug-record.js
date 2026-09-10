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

async function debugMayRecord() {
    console.log('\n=== DEBUG: TRACKING MAY 4 RECORD ===\n');

    await connectDB();
    await initRepositories();

    const assistantsArr = assistantRepo.read();
    const lecturersArr = lecturerRepo.read();
    const coursesArr = courseRepo.read();
    const schedules = scheduleRepo.read();
    const reports = reportRepo.read();

    const coursesMap = Object.fromEntries(coursesArr.map(c => [c.id, c]));
    const lecturersMap = Object.fromEntries(lecturersArr.map(l => [l.id, l]));
    const assistantsMap = Object.fromEntries(assistantsArr.map(a => [a.id, a]));

    // Check May 4 SCHED_001 report
    const mayReport = reports.find(r => r.date === '2026-05-04' && r.schedule_id === 'SCHED_001');
    console.log('[May 4 SCHED_001 Report from DB]');
    console.log('  date:', mayReport?.date);
    console.log('  formattedDate:', mayReport?.formattedDate);
    console.log('  schedule_id:', mayReport?.schedule_id);
    console.log('  materi:', mayReport?.materi);

    // Simulate building record from this report
    const sched = schedules.find(s => s.id === 'SCHED_001');
    const courseInfo = coursesMap[sched.course_id];
    const lecturerInfo = lecturersMap[sched.lecturer_id];

    console.log('\n[Building Record for May 4 SCHED_001]');
    const hariTanggalFromReport = mayReport?.formattedDate;
    console.log(`  Using formattedDate from report: "${hariTanggalFromReport}"`);

    const record = {
        hariTanggal: hariTanggalFromReport,
        materi: mayReport?.materi,
        mataKuliah: courseInfo.name,
        kelas: sched.class,
        asisten: (mayReport?.attending_assistant_ids || [])
            .map(id => (assistantsMap[id] || {}).panggilan || 'N/A'),
    };

    console.log('\n[Constructed Record]');
    console.log('  hariTanggal:', record.hariTanggal);
    console.log('  mataKuliah:', record.mataKuliah);
    console.log('  asisten:', JSON.stringify(record.asisten));

    // Check grouping key
    const dateKey = record.hariTanggal;
    console.log('\n[Will be grouped under dateKey]');
    console.log(`  Key: "${dateKey}"`);
    console.log(`  Key === "Senin, 04 Mei 2026"? ${dateKey === 'Senin, 04 Mei 2026'}`);

    // Check if key exists in groupedByDate
    console.log('\n[Checking formattedDate formatting]');
    const currentDate = new Date(2026, 4, 4);
    const formattedManual = formatDateIndonesian(currentDate);
    console.log(`  Manual format result: "${formattedManual}"`);
    console.log(`  DB formattedDate:     "${mayReport?.formattedDate}"`);
    console.log(`  Match? ${formattedManual === mayReport?.formattedDate}`);

    // Check differences
    if (formattedManual !== mayReport?.formattedDate) {
        console.log('\n  !!! MISMATCH FOUND !!!');
        console.log(`  Manual chars: [${[...formattedManual].map(c => `${c}(${c.charCodeAt(0)})`).join(', ')}]`);
        console.log(`  DB chars:     [${[...(mayReport?.formattedDate || '')].map(c => `${c}(${c.charCodeAt(0)})`).join(', ')}]`);
    }

    process.exit(0);
}

debugMayRecord().catch(err => {
    console.error('ERROR:', err);
    process.exit(1);
});
