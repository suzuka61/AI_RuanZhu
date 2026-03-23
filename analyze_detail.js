const JSZip = require('jszip');
const fs = require('fs');

const docPath = '/Users/songtao/Documents/设计说明书.docx';

async function analyzeDetail() {
  try {
    const data = fs.readFileSync(docPath);
    const zip = await JSZip.loadAsync(data);
    
    console.log('=== 详细分析参考文档 ===\n');
    
    // 分析页眉1
    console.log('1. 页眉1完整XML:');
    const header1 = await zip.file('word/header1.xml').async('string');
    console.log(header1);
    
    console.log('\n\n2. 页眉2完整XML:');
    const header2 = await zip.file('word/header2.xml').async('string');
    console.log(header2);
    
    console.log('\n\n3. 页脚完整XML:');
    const footer1 = await zip.file('word/footer1.xml').async('string');
    console.log(footer1);
    
    console.log('\n\n4. 文档关系文件:');
    const rels = await zip.file('word/_rels/document.xml.rels').async('string');
    console.log(rels);
    
  } catch (e) {
    console.error('分析失败:', e);
  }
}

analyzeDetail();
