const { Router } = require('express');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');
const { courseRepo, assistantRepo, scheduleRepo, reportRepo } = require('../repositories');
const { getEnrichedSchedules, getEnrichedReports } = require('../services/scheduleService');
const { buildAttendanceGrid, formatJam } = require('../services/attendanceService');
const { asyncHandler } = require('./_helpers');

const router = Router();

// Halaman pilihan rekap
router.get('/', asyncHandler(async (req, res) => {
    const stats = {
        totalCourses:    courseRepo.read().length,
        totalAssistants: assistantRepo.read().length,
        totalSchedules:  scheduleRepo.read().length,
        totalReports:    reportRepo.read().length,
    };
    res.render('rekap-pilihan', {
        title:      'Rekap Bendahara - Pilih Jenis Rekap',
        stats,
        activePage: 'rekap',
    });
}));

// Rekap SKS & Kehadiran Asisten
router.get('/asisten', asyncHandler(async (req, res) => {
    const assistants = assistantRepo.read();
    const schedules  = getEnrichedSchedules();
    const reports    = getEnrichedReports();

    const rekapAsisten = assistants.map(assistant => {
        const assignedSchedules = schedules.filter(s =>
            s.assistant_ids.includes(assistant.id) || s.pj_assistant_id === assistant.id
        );
        const attendedReports = reports.filter(r =>
            r.attending_assistant_ids?.includes(assistant.id)
        );

        let totalSKS       = 0;
        const detailKehadiran = [];

        attendedReports.forEach(report => {
            if (report.course?.sks) {
                totalSKS += report.course.sks;
                detailKehadiran.push({
                    tanggal:         report.date,
                    tanggalFormatted: report.formattedDate ||
                        new Date(report.date).toLocaleDateString('id-ID'),
                    courseCode:      report.course.code,
                    courseName:      report.course.name,
                    sks:             report.course.sks,
                    kelas:           report.schedule?.class ?? 'N/A',
                    status:          report.status,
                });
            }
        });

        const isPjSchedules     = schedules.filter(s => s.pj_assistant_id === assistant.id);
        const regularSchedules  = schedules.filter(s =>
            s.assistant_ids.includes(assistant.id) && s.pj_assistant_id !== assistant.id
        );

        return {
            ...assistant,
            totalSKS,
            totalKehadiran:      attendedReports.length,
            assignedSchedules:   assignedSchedules.length,
            isPjCount:           isPjSchedules.length,
            regularAssignmentCount: regularSchedules.length,
            detailKehadiran:     detailKehadiran.sort((a, b) => new Date(a.tanggal) - new Date(b.tanggal)),
            averageSKSPerReport: attendedReports.length > 0
                ? (totalSKS / attendedReports.length).toFixed(2) : 0,
        };
    }).sort((a, b) => b.totalSKS - a.totalSKS);

    res.render('rekap-asisten', {
        title:                 'Rekap SKS Asisten',
        rekapAsisten,
        totalAsisten:          assistants.length,
        totalSKSKeseluruhan:   rekapAsisten.reduce((sum, a) => sum + a.totalSKS, 0),
        activePage:            'rekap',
    });
}));

// Rekap Laporan Per Kelas
router.get('/kelas', asyncHandler(async (req, res) => {
    const schedules = getEnrichedSchedules();
    const reports   = getEnrichedReports();

    const rekapKelas = schedules.map(schedule => {
        const classReports = reports.filter(r => r.schedule_id === schedule.id);
        const sorted       = [...classReports].sort((a, b) => new Date(b.date) - new Date(a.date));

        const statusBreakdown = {
            hadir:           classReports.filter(r => r.status === 'hadir').length,
            libur:           classReports.filter(r => r.status === 'libur').length,
            pengganti:       classReports.filter(r => r.status === 'pengganti').length,
            belumAdaLaporan: classReports.filter(r => r.status === 'belum ada laporan').length,
        };

        const pertemuanEfektif  = statusBreakdown.hadir + statusBreakdown.pengganti;
        const totalPertemuan    = classReports.length;
        const latestReport      = sorted[0] ?? null;

        const detailLaporan = sorted.map(r => ({
            tanggal:          r.date,
            tanggalFormatted: r.formattedDate || new Date(r.date).toLocaleDateString('id-ID'),
            status:           r.status,
            materi:           r.materi        || '-',
            jumlahAsisten:    r.attending_assistant_ids?.length ?? 0,
            keterangan:       r.keterangan    || '-',
        }));

        return {
            ...schedule,
            totalLaporan:         totalPertemuan,
            pertemuanEfektif,
            tingkatKehadiran:     totalPertemuan > 0
                ? ((pertemuanEfektif / totalPertemuan) * 100).toFixed(1) : 0,
            statusBreakdown,
            latestReportDate:     latestReport?.date ?? null,
            latestReportFormatted: latestReport
                ? (latestReport.formattedDate || new Date(latestReport.date).toLocaleDateString('id-ID'))
                : 'Belum ada laporan',
            detailLaporan,
            totalSKSGenerated:    totalPertemuan * (schedule.course?.sks ?? 0),
        };
    }).sort((a, b) => b.totalLaporan - a.totalLaporan);

    res.render('rekap-kelas', {
        title:                   'Rekap Laporan Kelas',
        rekapKelas,
        totalKelas:              schedules.length,
        totalLaporanKeseluruhan: rekapKelas.reduce((sum, k) => sum + k.totalLaporan, 0),
        activePage:              'rekap',
    });
}));

// Rekap Kehadiran Bulanan (spreadsheet-style grid)
router.get('/kehadiran', asyncHandler(async (req, res) => {
    const now   = new Date();
    const month = parseInt(req.query.month) || (now.getMonth() + 1);
    const year  = parseInt(req.query.year)  || now.getFullYear();

    const gridData = buildAttendanceGrid(month, year);

    // Navigasi bulan sebelumnya / berikutnya
    const prevMonth = month > 1  ? month - 1 : 12;
    const prevYear  = month > 1  ? year      : year - 1;
    const nextMonth = month < 12 ? month + 1 : 1;
    const nextYear  = month < 12 ? year      : year + 1;

    res.render('rekap-kehadiran', {
        title:       `Rekap Kehadiran — ${gridData.monthInfo.monthName} ${year}`,
        activePage:  'rekap',
        gridData,
        formatJam,
        currentMonth: month,
        currentYear:  year,
        prevMonth, prevYear,
        nextMonth, nextYear,
    });
}));

// Trigger Export XLSX (Semua bulan di semester)
router.get('/export-xlsx', asyncHandler(async (req, res) => {
    const { generateAllRekapKehadiran, loadJsonData, GLOBAL_CONFIG_PATH } = require('../EXPORTER_V2/exporter');
    const globalConfig = loadJsonData(GLOBAL_CONFIG_PATH);
    
    // Jalankan export
    await generateAllRekapKehadiran(globalConfig);

    // Redirect kembali ke halaman asal (atau halaman rekap) dengan notifikasi sukses
    res.redirect('/rekap/kehadiran?exportSuccess=true');
}));

// Download ZIP berisi semua folder di EXPORTER_V2/output_export
router.get('/download-exports-zip', asyncHandler(async (req, res) => {
    const exportsRoot = path.join(__dirname, '..', 'EXPORTER_V2', 'output_export');

    if (!fs.existsSync(exportsRoot)) {
        return res.status(404).send('Folder export belum tersedia.');
    }

    // Jika folder kosong, berikan respon yang jelas.
    const hasAnyEntry = fs.readdirSync(exportsRoot).some((name) => !name.startsWith('.'));
    if (!hasAnyEntry) {
        return res.status(404).send('Tidak ada file export untuk di-zip.');
    }

    const zipFileName = `saq_export_${new Date().toISOString().replace(/[:.]/g, '-')}.zip`;

    res.setHeader('Content-Type', 'application/zip');
    res.attachment(zipFileName);

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.on('error', (err) => {
        // Jika error terjadi sebelum response selesai, handler express akan menangkap reject.
        console.error('[ZIP] Error:', err);
    });

    // Stream hasil zip langsung ke browser.
    archive.pipe(res);

    const addDir = (dirPath) => {
        const entries = fs.readdirSync(dirPath, { withFileTypes: true });
        for (const entry of entries) {
            if (entry.name.startsWith('.')) continue;

            const fullPath = path.join(dirPath, entry.name);
            const relPath  = path.relative(exportsRoot, fullPath).replace(/\\/g, '/');

            if (entry.isDirectory()) {
                addDir(fullPath);
            } else if (entry.isFile()) {
                archive.file(fullPath, { name: relPath });
            }
        }
    };

    addDir(exportsRoot);

    // Finalize setelah semua file ditambahkan.
    await new Promise((resolve, reject) => {
        archive.on('error', reject);
        res.on('close', resolve);
        res.on('finish', resolve);
        archive.finalize().catch(reject);
    });
}));

module.exports = router;
