const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const toIco = require('to-ico');

const icnsPath = path.join(__dirname, '..', 'icon 2.icns');
const tempDir = path.join(__dirname, 'temp_icon');
const icoOutput = path.join(__dirname, '..', 'icon.ico');

if (fs.existsSync(tempDir)) {
    fs.rmSync(tempDir, { recursive: true, force: true });
}
fs.mkdirSync(tempDir, { recursive: true });

console.log('Extracting PNG from ICNS...');

try {
    execSync(`iconutil -p -x "${icnsPath}"`, { cwd: tempDir, encoding: null });

    const files = fs.readdirSync(tempDir);
    console.log('Extracted files:', files);

    const pngFiles = files.filter(f => f.endsWith('.png'));

    if (pngFiles.length === 0) {
        console.error('No PNG files found in ICNS');
        process.exit(1);
    }

    const pngBuffers = pngFiles.map(f => fs.readFileSync(path.join(tempDir, f)));
    console.log(`Found ${pngBuffers.length} PNG sizes`);

    toIco(pngBuffers).then(icoBuffer => {
        fs.writeFileSync(icoOutput, icoBuffer);
        console.log(`Created ${icoOutput}`);
        fs.rmSync(tempDir, { recursive: true, force: true });
        console.log('Done!');
    });
} catch (e) {
    console.error('Error:', e.message);
    process.exit(1);
}
