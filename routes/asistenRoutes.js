const { Router } = require('express');
const { assistantRepo }   = require('../repositories');
const AppError            = require('../config/AppError');
const { asyncHandler, generateId } = require('./_helpers');

const router = Router();

// READ
router.get('/', asyncHandler(async (req, res) => {
    const assistants = assistantRepo.read();
    res.render('asisten', { assistants, title: 'Manajemen Data Asisten', activePage: 'asisten' });
}));

// CREATE - Form
router.get('/add', (req, res) => {
    res.render('add-asisten', { title: 'Tambah Asisten Baru', activePage: 'asisten' });
});

// CREATE - Proses
router.post('/add', asyncHandler(async (req, res) => {
    const assistants   = assistantRepo.read();
    const newAssistant = {
        id:           generateId('ASST'),
        nama_asisten: req.body.nama_asisten.toUpperCase(),
        nama_panjang: req.body.nama_panjang,
        panggilan:    req.body.panggilan,
        no_telp:      req.body.no_telp,
    };
    assistants.push(newAssistant);
    assistantRepo.write(assistants);
    res.redirect('/asisten');
}));

// UPDATE - Form
router.get('/edit/:id', asyncHandler(async (req, res) => {
    const assistants = assistantRepo.read();
    const assistant  = assistants.find(a => a.id === req.params.id);
    if (!assistant) throw new AppError('Asisten tidak ditemukan', 404);
    res.render('edit-asisten', { title: 'Edit Data Asisten', assistant, activePage: 'asisten' });
}));

// UPDATE - Proses
router.post('/update/:id', asyncHandler(async (req, res) => {
    const assistants = assistantRepo.read();
    const idx        = assistants.findIndex(a => a.id === req.params.id);
    if (idx === -1) throw new AppError('Asisten tidak ditemukan', 404);

    assistants[idx] = {
        ...assistants[idx],
        nama_asisten: req.body.nama_asisten.toUpperCase(),
        nama_panjang: req.body.nama_panjang,
        panggilan:    req.body.panggilan,
        no_telp:      req.body.no_telp,
    };
    assistantRepo.write(assistants);
    res.redirect('/asisten');
}));

// DELETE
router.post('/delete/:id', asyncHandler(async (req, res) => {
    const assistants   = assistantRepo.read();
    const filtered     = assistants.filter(a => a.id !== req.params.id);
    if (filtered.length === assistants.length) throw new AppError('Asisten tidak ditemukan', 404);
    assistantRepo.write(filtered);
    res.redirect('/asisten');
}));

module.exports = router;
