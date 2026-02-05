import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function analyze(filePath) {
    if (!fs.existsSync(filePath)) {
        console.error(`File not found: ${filePath}`);
        return;
    }
    console.log(`\n--- Analyzing: ${path.basename(filePath)} ---`);
    try {
        const data = fs.readFileSync(filePath);
        const zip = await JSZip.loadAsync(data);
        const slideFiles = Object.keys(zip.files).filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml')).sort();

        for (const slidePath of slideFiles) {
            const content = await zip.file(slidePath).async('string');
            const tNodes = content.match(/<a:t>([^<]*)<\/a:t>/g) || [];
            const nodes = tNodes.map(n => n.replace(/<\/?a:t>/g, ''));
            nodes.forEach((txt, i) => {
                if (txt.includes('010') || txt.includes('@') || txt.includes('연락')) {
                    console.log(`[SLIDE ${slidePath}] Index ${i}: "${txt}"`);
                    console.log(`   Next nodes: [${i + 1}] "${nodes[i + 1]}", [${i + 2}] "${nodes[i + 2]}"`);
                }
            });
        }
    } catch (e) { console.error(e); }
}

const samples = [
    path.join(__dirname, 'samples', 'B유형', '김민옥 프로필 .pptx'),
    path.join(__dirname, 'samples', 'E유형', '김다정_프로필.pptx')
];

for (const s of samples) {
    await analyze(s);
}
