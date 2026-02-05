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
    console.log(`Analyzing: ${filePath}`);
    try {
        const data = fs.readFileSync(filePath);
        const zip = await JSZip.loadAsync(data);

        const slideFiles = Object.keys(zip.files).filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml')).sort();

        for (const slidePath of slideFiles) {
            const content = await zip.file(slidePath).async('string');

            // 이름이나 생년월일 관련 키워드 검색 (조각조각 나있을 수 있으므로 유의)
            if (content.includes('남') || content.includes('미') || content.includes('예') || content.includes('성') || content.includes('명') || content.includes('생')) {
                console.log(`--- Slide: ${slidePath} ---`);
                const tNodes = content.match(/<a:t>([^<]*)<\/a:t>/g) || [];
                console.log('Text fragments found:');
                tNodes.forEach((n, i) => {
                    const txt = n.replace(/<\/?a:t>/g, '');
                    if (txt.trim()) console.log(`[${i}] "${txt.trim()}"`);
                });
            }
        }
    } catch (e) {
        console.error(`Error: ${e.message}`);
    }
}

const target = path.join(__dirname, 'samples', 'E유형', '김다정_프로필.pptx');
analyze(target);
