const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

const docPath = '/Users/songtao/Documents/设计说明书.docx';

console.log('=== 解析参考文档样式 ===\n');

async function parseDocx() {
  try {
    const data = fs.readFileSync(docPath);
    const zip = await JSZip.loadAsync(data);
    
    console.log('1. 文档结构:');
    Object.keys(zip.files).filter(f => !f.endsWith('/')).forEach(f => {
      console.log(`   - ${f}`);
    });
    
    console.log('\n2. 读取 document.xml:');
    const docXml = await zip.file('word/document.xml').async('string');
    
    console.log('   文档XML长度:', docXml.length);
    
    console.log('\n3. 查找页眉/页脚:');
    const headerFiles = Object.keys(zip.files).filter(f => f.startsWith('word/header'));
    const footerFiles = Object.keys(zip.files).filter(f => f.startsWith('word/footer'));
    
    console.log(`   页眉文件: ${headerFiles.length > 0 ? headerFiles.join(', ') : '无'}`);
    console.log(`   页脚文件: ${footerFiles.length > 0 ? footerFiles.join(', ') : '无'}`);
    
    if (headerFiles.length > 0) {
      console.log('\n4. 页眉内容预览:');
      for (const hf of headerFiles) {
        const headerXml = await zip.file(hf).async('string');
        console.log(`   ${hf}:`);
        const textContent = headerXml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
        console.log(`   ${textContent.slice(0, 200)}...`);
      }
    }
    
    console.log('\n5. 读取样式:');
    if (zip.file('word/styles.xml')) {
      const stylesXml = await zip.file('word/styles.xml').async('string');
      console.log('   styles.xml 存在');
      const styleMatches = [...stylesXml.matchAll(/w:styleId="([^"]+)"/g)];
      console.log(`   找到 ${styleMatches.length} 个样式`);
    }
    
    console.log('\n6. 文档内容预览 (文本):');
    const textOnly = docXml.replace(/<[^>]+>/g, '\n').replace(/\n\s*\n/g, '\n').trim();
    console.log(textOnly.slice(0, 1000));
    
  } catch (e) {
    console.error('解析失败:', e.message);
  }
}

parseDocx();
