'use strict';
const path = require('path');
const { connectDB } = require('./config/database');
const { initRepositories, reportRepo } = require('./repositories');

async function main() {
    require('dotenv').config();

    await connectDB();
    await initRepositories();

    const reports = reportRepo.read();
    console.log(`Total reports: ${reports.length}`);

    const withImage = reports.filter(r => r.imagePath || r.imageFilename);
    console.log(`Reports with imagePath or imageFilename: ${withImage.length}`);
    for (const r of withImage) {
        console.log(`ID: ${r.id}, Date: ${r.date}, ImagePath: ${r.imagePath}, ImageFilename: ${r.imageFilename}`);
    }
    process.exit(0);
}

main().catch(err => {
    console.error(err);
    process.exit(1);
});
