// 秒著 - 一体化服务器 v3
// node proxy.js
const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const PORT = 3000;
const HTML_FILE = path.join(__dirname, '秒著.html');
const TARGET_HOST = 'ark.cn-beijing.volces.com';
const TARGET_PATH = '/api/coding/v3/chat/completions';

function ensureDocx() {
  if (!fs.existsSync(path.join(__dirname, 'node_modules', 'docx'))) {
    console.log('  正在安装 docx 依赖包...');
    try { execSync('npm install docx', { cwd: __dirname, stdio: 'inherit' }); console.log('  ✓ docx 安装完成'); }
    catch(e) { console.error('  ✗ docx 安装失败，请手动运行: npm install docx'); }
  }
}
ensureDocx();

// ── 工具函数 ──────────────────────────────────────────────────────────────
function makeBorder() {
  const b = { style: 'single', size: 1, color: '888888' };
  return { top: b, bottom: b, left: b, right: b };
}

function makeHeaderCell(text, widthDxa) {
  const { TableCell, Paragraph, TextRun, ShadingType, WidthType } = require('docx');
  return new TableCell({
    borders: makeBorder(),
    width: { size: widthDxa, type: WidthType.DXA },
    shading: { fill: 'DDEEFF', type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [new TextRun({ text: String(text||''), bold: true, size: 22, font: '宋体' })] })]
  });
}

function makeDataCell(text, widthDxa) {
  const { TableCell, Paragraph, TextRun, WidthType } = require('docx');
  return new TableCell({
    borders: makeBorder(),
    width: { size: widthDxa, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [new TextRun({ text: String(text||''), size: 22, font: '宋体' })] })]
  });
}

function makePara(text, opts = {}) {
  const { Paragraph, TextRun, AlignmentType } = require('docx');
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: opts.before || 0, after: opts.after !== undefined ? opts.after : 100 },
    children: [new TextRun({ text: String(text||''), bold: opts.bold, size: opts.size || 22, font: '宋体', color: opts.color })]
  });
}

function makeHeading(text, level) {
  const { Paragraph, TextRun, HeadingLevel } = require('docx');
  const lvlMap = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
  return new Paragraph({
    heading: lvlMap[level] || HeadingLevel.HEADING_2,
    spacing: { before: level === 1 ? 360 : 240, after: 120 },
    children: [new TextRun({ text: String(text||''), font: '宋体', bold: true, size: level === 1 ? 28 : level === 2 ? 24 : 22 })]
  });
}

// 将 base64 图片转为 ImageRun（带尺寸限制）
// 从PNG/JPEG二进制读取原始宽高
function getImageSize(buf, mime) {
  try {
    if (mime && (mime.includes('png'))) {
      // PNG: width at offset 16, height at offset 20 (big-endian)
      if (buf[0]===0x89 && buf[1]===0x50 && buf[2]===0x4E && buf[3]===0x47) {
        const w = buf.readUInt32BE(16);
        const h = buf.readUInt32BE(20);
        if (w > 0 && h > 0) return { w, h };
      }
    }
    if (mime && (mime.includes('jpeg') || mime.includes('jpg'))) {
      // JPEG: scan for SOF marker
      let i = 2;
      while (i < buf.length - 8) {
        if (buf[i] === 0xFF) {
          const marker = buf[i+1];
          if (marker >= 0xC0 && marker <= 0xC3) {
            const h = buf.readUInt16BE(i+5);
            const w = buf.readUInt16BE(i+7);
            if (w > 0 && h > 0) return { w, h };
          }
          const segLen = buf.readUInt16BE(i+2);
          i += 2 + segLen;
        } else { i++; }
      }
    }
  } catch(e) {}
  return null;
}

function makeImageRun(imgObj) {
  const { ImageRun } = require('docx');
  const buf = Buffer.from(imgObj.b64, 'base64');
  const mime = imgObj.mime || 'image/png';
  const type = mime.includes('jpeg') || mime.includes('jpg') ? 'jpg' : 'png';

  // 最大可用宽度（A4 减去页边距）约 500px
  const MAX_W = 500;
  const MAX_H = 600;

  const size = getImageSize(buf, mime);
  let width, height;
  if (size && size.w > 0 && size.h > 0) {
    const ratio = size.w / size.h;
    if (size.w <= MAX_W && size.h <= MAX_H) {
      width = size.w;
      height = size.h;
    } else if (size.w / MAX_W > size.h / MAX_H) {
      width = MAX_W;
      height = Math.round(MAX_W / ratio);
    } else {
      height = MAX_H;
      width = Math.round(MAX_H * ratio);
    }
  } else {
    width = MAX_W;
    height = Math.round(MAX_W * 0.6);
  }

  return new ImageRun({ data: buf, transformation: { width, height }, type });
}

// 图片段落
function makeImagePara(imgObj, caption) {
  const { Paragraph, TextRun, AlignmentType } = require('docx');
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: 60 },
      children: [makeImageRun(imgObj)]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: caption || '', size: 18, font: '宋体', color: '666666', italics: true })]
    })
  ];
}

// ── 页眉页脚工具 ──────────────────────────────────────────────────────────
// 页眉：左侧软件名，右侧"第x页 共x页"
function makeDocHeader(titleText) {
  const { Header, Paragraph, TextRun, TabStopPosition, TabStopType, PageNumber } = require('docx');

  return new Header({
    children: [
      new Paragraph({
        tabStops: [
          { type: TabStopType.RIGHT, position: TabStopPosition.MAX }
        ],
        children: [
          new TextRun({ 
            text: titleText, 
            size: 18, 
            font: '宋体'
          }),
          new TextRun({ text: '\t', size: 18, font: '宋体' }),
          new TextRun({ text: '第', size: 18, font: '宋体' }),
          new TextRun({ children: [PageNumber.CURRENT], size: 18, font: '宋体' }),
          new TextRun({ text: '页 共', size: 18, font: '宋体' }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: '宋体' }),
          new TextRun({ text: '页', size: 18, font: '宋体' }),
        ]
      })
    ]
  });
}

// 页脚为空（不使用页脚）
function makeEmptyFooter() {
  const { Footer, Paragraph } = require('docx');
  return new Footer({
    children: [new Paragraph({ children: [] })]
  });
}

function makeEmptyHeaderFooter() {
  const { Header, Footer, Paragraph } = require('docx');
  return {
    header: new Header({ 
      children: [new Paragraph({ children: [] })] 
    }),
    footer: new Footer({ 
      children: [new Paragraph({ children: [] })] 
    }),
  };
}

// ── 生成软件概述 ──────────────────────────────────────────────────────────
async function buildOverviewDocx(data) {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          AlignmentType, WidthType, ShadingType, PageBreak, HeaderFooterType } = require('docx');

  const info = data.info || {};
  const images = data.images || [];
  const softwareName = info.software_name || '软件概述';

  function fieldRow(label, value) {
    return new TableRow({
      children: [
        new TableCell({
          borders: makeBorder(), width: { size: 3200, type: WidthType.DXA },
          shading: { fill: 'F5F5F5', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 22, font: '宋体' })] })]
        }),
        new TableCell({
          borders: makeBorder(), width: { size: 6800, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: String(value||'待填写'), size: 22, font: '宋体' })] })]
        }),
      ]
    });
  }

  // 功能模块列表
  const modules = (info.main_features || '').split('\n').filter(l => l.trim()).slice(0, 10);
  const funcRows = [
    new TableRow({ children: ['序号','功能名称','功能描述'].map((h,i) => makeHeaderCell(h,[1000,2800,6200][i])) }),
    ...modules.map((mod, i) => {
      const clean = mod.replace(/^[\d\.\-\*①②③④⑤⑥⑦⑧⑨⑩、]+/, '').trim();
      const parts = clean.split(/[：:]/);
      const name = (parts[0]||'').trim() || ('功能'+(i+1));
      const desc = (parts[1]||clean).trim();
      return new TableRow({ children: [makeDataCell(i+1,1000), makeDataCell(name,2800), makeDataCell(desc,6200)] });
    })
  ];

  // 查找架构图
  const archImg = images.find(img => img.section && (img.section.includes('架构') || img.section.includes('功能'))) || images[0];

  const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 } } };
  const emptyHF = makeEmptyHeaderFooter();

  // 封面 section
  const coverChildren = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 3000, after: 600 }, children: [new TextRun({ text: softwareName, bold: true, size: 48, font: '宋体' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 800 }, children: [new TextRun({ text: '软件概述', bold: true, size: 40, font: '宋体' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: info.developer || '', size: 26, font: '宋体' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: info.completion_date || '', size: 24, font: '宋体' })] }),
  ];

  // 正文内容
  const bodyChildren = [
    new Table({
      width: { size: 10000, type: WidthType.DXA },
      rows: [
        fieldRow('软件全称', info.software_name),
        fieldRow('软件简称', info.short_name),
        fieldRow('版本号', info.version),
        fieldRow('软件分类', '应用软件'),
        fieldRow('软件说明', '①原创'),
        fieldRow('开发方式', '独立开发'),
        fieldRow('开发完成日期', info.completion_date),
        fieldRow('首次发表日期', info.publish_date),
        fieldRow('首次发表地点', info.publish_location),
        fieldRow('开发的硬件环境', info.dev_hardware),
        fieldRow('运行的硬件环境', info.dev_hardware),
        fieldRow('开发该软件的操作系统', info.dev_os),
        fieldRow('软件开发环境/开发工具', info.dev_tools),
        fieldRow('该软件的运行平台/操作系统', info.run_platform),
        fieldRow('软件运行支撑环境/支持软件', info.runtime_env),
        fieldRow('编程语言', info.programming_langs),
        fieldRow('源程序量', (info.source_lines||'') + ' 行'),
      ]
    }),
    makePara(''),
    makeHeading('开发目的', 2),
    ...(info.purpose||'').split('\n').filter(l=>l.trim()).map(l => makePara(l)),
    makePara(''),
    makeHeading('面向领域/行业', 2),
    makePara(info.industry || '通用'),
    makePara(''),
    makeHeading('软件的主要功能', 2),
    new Table({ width: { size: 10000, type: WidthType.DXA }, rows: funcRows }),
    makePara(''),
    makeHeading('软件的技术特点', 2),
    ...(info.architecture||'').split('\n').filter(l=>l.trim()).map(l => makePara(l)),
  ];

  if (archImg) {
    bodyChildren.push(makePara(''));
    bodyChildren.push(makeHeading('产品架构图', 2));
    bodyChildren.push(...makeImagePara(archImg, '图1 产品功能架构图'));
  }

  const doc = new Document({
    sections: [{
      properties: { ...pageProps },
      headers: { default: makeDocHeader(softwareName) },
      children: [
        ...coverChildren,
        new Paragraph({ children: [new PageBreak()] }),
        ...bodyChildren,
      ],
    }]
  });
  return Packer.toBuffer(doc);
}

// ── 生成设计说明书 ─────────────────────────────────────────────────────────
async function buildDesignDocx(data) {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, AlignmentType,
          HeadingLevel, WidthType, PageBreak } = require('docx');

  const info = data.info || {};
  const images = data.images || [];
  const aiContent = data.content || '';

  // 调试日志：确认收到内容
  console.log('[调试] aiContent 长度:', aiContent.length);
  console.log('[调试] aiContent 前300字:', aiContent.slice(0, 300));

  // 人员分工表
  const members = info.team_members || [];
  const validMembers = members.filter(m => m.name && m.name.trim() && m.name !== '待填写');
  const memberWidths = [1500, 2500, 2000, 2000, 2000];
  const memberRows = [
    new TableRow({ children: ['姓名','项目组职务','参与时间','研究阶段','开发阶段'].map((h,i) => makeHeaderCell(h, memberWidths[i])) }),
    ...(validMembers.length > 0 
      ? validMembers.map(m => new TableRow({ children: [m.name||'', m.role||'', m.period||'', m.research||m.period||'', m.dev||m.period||''].map((v,i) => makeDataCell(v, memberWidths[i])) }))
      : [new TableRow({ children: ['','','','',''].map((v,i) => makeDataCell('', memberWidths[i])) })]
    )
  ];

  // 功能目标表
  const modules = (info.main_features||'').split('\n').filter(l=>l.trim()).slice(0,10);
  const goalRows = [
    new TableRow({ children: ['序号','名称','功能'].map((h,i) => makeHeaderCell(h,[1000,2800,6200][i])) }),
    ...modules.map((mod,i) => {
      const clean = mod.replace(/^[\d\.\-\*①-⑩、]+/,'').trim();
      const parts = clean.split(/[：:]/);
      const name = (parts[0]||'').trim() || ('功能'+(i+1));
      const desc = (parts[1]||clean).trim();
      return new TableRow({ children: [makeDataCell(i+1,1000), makeDataCell(name,2800), makeDataCell(desc,6200)] });
    })
  ];

  // ── 图片智能匹配 ──
  function textSimilarity(a, b) {
    if (!a || !b) return 0;
    a = a.replace(/[\s\d\.\、\#\*]+/g, '').toLowerCase();
    b = b.replace(/[\s\d\.\、\#\*]+/g, '').toLowerCase();
    if (!a || !b) return 0;
    if (a === b) return 1;
    if (a.includes(b) || b.includes(a)) return 0.8;
    let common = 0;
    const shorter = a.length < b.length ? a : b;
    const longer = a.length < b.length ? b : a;
    for (const ch of shorter) {
      if (longer.includes(ch)) common++;
    }
    return common / Math.max(a.length, b.length);
  }

  function findNearestHeading(lines, lineIndex) {
    for (let i = lineIndex; i >= 0; i--) {
      const t = lines[i].trim();
      if (/^#{1,3}\s/.test(t)) return t.replace(/^#+\s*/, '');
      if (/^[\d]+[\.\、]/.test(t) && t.length < 60) return t.replace(/^[\d]+[\.\、]\s*/, '');
    }
    return '';
  }

  function matchImagesToPlaceholders(lines, images, archImg) {
    const placeholders = [];
    for (let i = 0; i < lines.length; i++) {
      const trimmed = lines[i].trim();
      const wireframeMatch = trimmed.match(/^\[WIREFRAME(?::(.+))?\]$/);
      if (wireframeMatch || trimmed === '[ARCHITECTURE_IMAGE]') {
        const heading = findNearestHeading(lines, i);
        const explicitName = wireframeMatch ? (wireframeMatch[1] || '').trim() : '';
        placeholders.push({ 
          lineIndex: i, 
          type: wireframeMatch ? '[WIREFRAME]' : trimmed, 
          heading: explicitName || heading
        });
      }
    }

    // 使用图片索引来追踪已使用的图片（避免对象引用问题）
    const usedImageIndices = new Set();
    const mapping = {};

    // 找到架构图在 images 数组中的索引
    let archImgIndex = -1;
    if (archImg) {
      archImgIndex = images.findIndex(img => img.name === archImg.name && img.from === archImg.from);
    }

    for (const ph of placeholders) {
      if (ph.type === '[ARCHITECTURE_IMAGE]' && archImgIndex >= 0) {
        mapping[ph.lineIndex] = { img: images[archImgIndex], caption: '产品功能架构图' };
        usedImageIndices.add(archImgIndex);
      }
    }

    const wireframePhs = placeholders.filter(ph => ph.type === '[WIREFRAME]');
    let wireframeIndex = 0;

    // 策略 1：首先尝试通过标题相似度匹配图片和占位符
    for (const ph of wireframePhs) {
      if (!ph.heading) continue;
      
      // 找到最匹配的图片
      let bestMatch = { index: -1, similarity: 0 };
      for (let i = 0; i < images.length; i++) {
        if (usedImageIndices.has(i)) continue;
        const img = images[i];
        if (!img.section) continue;
        
        const sim = textSimilarity(ph.heading, img.section);
        if (sim > bestMatch.similarity) {
          bestMatch = { index: i, similarity: sim };
        }
      }
      
      // 如果相似度足够高（>0.5），使用匹配的图片
      if (bestMatch.index >= 0 && bestMatch.similarity > 0.5) {
        const img = images[bestMatch.index];
        const caption = img.section ? img.section.replace(/^[\d\.\、\s]+/, '').slice(0, 40) : (ph.heading || '界面示意图');
        mapping[ph.lineIndex] = { img, caption };
        usedImageIndices.add(bestMatch.index);
      }
    }

    // 策略 2：对于未匹配的占位符，按顺序匹配剩余图片
    for (const ph of wireframePhs) {
      if (mapping[ph.lineIndex]) continue; // 已经匹配的跳过
      
      // 找到下一个未使用的图片
      while (wireframeIndex < images.length && usedImageIndices.has(wireframeIndex)) {
        wireframeIndex++;
      }
      if (wireframeIndex < images.length) {
        const img = images[wireframeIndex];
        const caption = img.section ? img.section.replace(/^[\d\.\、\s]+/, '').slice(0, 40) : (ph.heading || '界面示意图');
        mapping[ph.lineIndex] = { img, caption };
        usedImageIndices.add(wireframeIndex);
        wireframeIndex++;
      }
    }

    const unmatchedImgs = images.filter((img, idx) => !usedImageIndices.has(idx));
    return { mapping, unmatchedImgs };
  }

  const archImg = images.find(img => img.section && (img.section.includes('架构') || img.section.includes('功能模块') || img.section.includes('产品功能')));
  const wireframeImgs = images.filter(img => img !== archImg);

  // ── 清理 markdown 格式符号 ──
  function cleanMarkdown(txt) {
    return (txt || '').replace(/\*\*/g, '').replace(/\*/g, '').replace(/`/g, '').trim();
  }

  // ── AI内容转docx段落 ──
  function contentToParasWithImages(text, images) {
    const lines = (text||'').split('\n');
    const { mapping, unmatchedImgs } = matchImagesToPlaceholders(lines, images, archImg);
    const result = [];
    let skipTocSection = false;
    
    for (let i = 0; i < lines.length; i++) {
      const trimmed = lines[i].trim();
      
      if (/^\[WIREFRAME(?::.+)?\]$/.test(trimmed) || trimmed === '[ARCHITECTURE_IMAGE]') {
        if (mapping[i]) {
          result.push(...makeImagePara(mapping[i].img, mapping[i].caption));
        }
        continue;
      }
      
      if (!trimmed) {
        if (!skipTocSection) {
          result.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
        }
        continue;
      }
      
      if (/^#{1,3}\s/.test(trimmed)) {
        const title = cleanMarkdown(trimmed.replace(/^#+\s/, ''));
        const lvl = trimmed.startsWith('###') ? 3 : trimmed.startsWith('##') ? 2 : 1;
        
        if (title.includes('目录')) {
          skipTocSection = true;
          continue;
        }
        if (skipTocSection && lvl === 2) {
          skipTocSection = false;
        }
        if (!skipTocSection) {
          result.push(makeHeading(title, lvl));
        }
      } else if (!skipTocSection) {
        if (/^[-\*•]\s/.test(trimmed)) {
          result.push(new Paragraph({ bullet: { level: 0 }, spacing: { after: 60 }, children: [new TextRun({ text: cleanMarkdown(trimmed.replace(/^[-\*•]\s/, '')), size: 22, font: '宋体' })] }));
        } else if (/^\d+\.\s/.test(trimmed)) {
          result.push(new Paragraph({ numbering: { reference: 'default', level: 0 }, spacing: { after: 60 }, children: [new TextRun({ text: cleanMarkdown(trimmed.replace(/^\d+\.\s/, '')), size: 22, font: '宋体' })] }));
        } else {
          result.push(makePara(cleanMarkdown(trimmed)));
        }
      }
    }
    
    return { paragraphs: result, unmatchedImgs };
  }

  const { paragraphs: requirementChildren, unmatchedImgs } = contentToParasWithImages(aiContent, wireframeImgs);

  if (unmatchedImgs.length > 0) {
    requirementChildren.push(makePara(''));
    requirementChildren.push(makeHeading('附录：其他界面截图', 2));
    for (const img of unmatchedImgs) {
      const caption = img.section ? img.section.replace(/^[\d\.\、\s]+/, '').slice(0, 40) : '界面示意图';
      requirementChildren.push(...makeImagePara(img, caption));
    }
  }

  const { HeaderFooterType, PageNumberFormat } = require('docx');
  const softwareName = info.software_name || '设计说明书';
  const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 } } };
  const emptyHF = makeEmptyHeaderFooter();
  
  // 调试：确认页眉创建
  const docHeader = makeDocHeader(softwareName);
  console.log('[调试] 设计说明书页眉创建成功:', softwareName);

  // 封面内容
  const coverChildren = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 3000, after: 600 }, children: [new TextRun({ text: softwareName, bold: true, size: 48, font: '宋体' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 800 }, children: [new TextRun({ text: '设计说明书', bold: true, size: 40, font: '宋体' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: info.developer||'', size: 26, font: '宋体' })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: info.completion_date||'', size: 24, font: '宋体' })] }),
  ];

  const { TableOfContents } = require('docx');
  
  // 从AI内容中提取标题，生成静态目录
  function extractTocFromContent(text) {
    const lines = (text || '').split('\n');
    const tocItems = [];
    for (const line of lines) {
      const trimmed = line.trim();
      if (/^#{1,3}\s/.test(trimmed)) {
        const title = trimmed.replace(/^#+\s*/, '');
        const level = trimmed.startsWith('###') ? 3 : trimmed.startsWith('##') ? 2 : 1;
        if (!title.includes('目录')) {
          tocItems.push({ title, level });
        }
      } else if (/^\d+[\.\、]/.test(trimmed) && trimmed.length < 60) {
        const title = trimmed.replace(/^\d+[\.\、]\s*/, '');
        if (!title.includes('目录')) {
          tocItems.push({ title, level: 2 });
        }
      }
    }
    return tocItems.slice(0, 20);
  }

  const tocItems = extractTocFromContent(aiContent);
  const indentMap = { 1: 0, 2: 400, 3: 800 };
  const tocParagraphs = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 400 }, children: [new TextRun({ text: '目录', bold: true, size: 32, font: '宋体' })] }),
    ...tocItems.map(item => new Paragraph({
      spacing: { before: 80, after: 80 },
      indent: { left: indentMap[item.level] || 0 },
      children: [new TextRun({ text: item.title, size: 24, font: '宋体' })]
    })),
  ];

  const doc = new Document({
    styles: {
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 32, bold: true, font: '宋体' }, paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 26, bold: true, font: '宋体' }, paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } },
        { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 24, bold: true, font: '宋体' }, paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
      ]
    },
    sections: [
      // 所有内容统一在一个section中，使用分页符分隔
      {
        properties: { ...pageProps },
        headers: { default: docHeader },
        children: [
          ...coverChildren,
          new Paragraph({ children: [new PageBreak()] }),
          ...tocParagraphs,
          new Paragraph({ children: [new PageBreak()] }),
          ...requirementChildren,
        ],
      },
    ]
  });
  return Packer.toBuffer(doc);
}

// ── HTTP 服务器 ──────────────────────────────────────────────────────────
const server = http.createServer((req, res) => {

  if (req.method === 'OPTIONS') {
    res.writeHead(204, { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'POST, GET, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type, Authorization' });
    res.end(); return;
  }

  // /proxy → 火山引擎
  if (req.url === '/proxy' && req.method === 'POST') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => {
      let parsedBody;
      try {
        parsedBody = JSON.parse(body);
      } catch(e) {
        res.writeHead(400, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
        res.end(JSON.stringify({ error: '无效的JSON格式' }));
        return;
      }
      
      const apiKey = parsedBody.apiKey || '';
      if (!apiKey) {
        res.writeHead(401, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
        res.end(JSON.stringify({ error: '缺少API密钥' }));
        return;
      }
      
      const forwardBody = JSON.stringify({
        model: parsedBody.model,
        messages: parsedBody.messages
      });
      
      const opts = { 
        hostname: TARGET_HOST, 
        path: TARGET_PATH, 
        method: 'POST', 
        headers: { 
          'Content-Type': 'application/json', 
          'Authorization': 'Bearer ' + apiKey, 
          'Content-Length': Buffer.byteLength(forwardBody) 
        } 
      };
      const pr = https.request(opts, pRes => {
        res.writeHead(pRes.statusCode, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
        pRes.pipe(res);
        pRes.on('end', () => console.log('[' + new Date().toLocaleTimeString() + '] API ' + pRes.statusCode));
      });
      pr.on('error', e => { res.writeHead(502); res.end(JSON.stringify({ error: { message: e.message } })); });
      console.log('[' + new Date().toLocaleTimeString() + '] → 火山引擎');
      pr.write(forwardBody); pr.end();
    });
    return;
  }

  // /generate-docx → 生成Word
  if (req.url === '/generate-docx' && req.method === 'POST') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', async () => {
      try {
        const data = JSON.parse(body);
        console.log('[' + new Date().toLocaleTimeString() + '] 生成docx: ' + data.type + ' | 图片数: ' + (data.images||[]).length + ' | 人员数: ' + (data.info?.team_members||[]).length);
        const buffer = data.type === 'overview' ? await buildOverviewDocx(data) : await buildDesignDocx(data);
        const fname = data.type === 'overview' ? '%E8%BD%AF%E4%BB%B6%E6%A6%82%E8%BF%B0.docx' : '%E8%AE%BE%E8%AE%A1%E8%AF%B4%E6%98%8E%E4%B9%A6.docx';
        res.writeHead(200, {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          'Content-Disposition': 'attachment; filename*=UTF-8\'\'' + fname,
          'Access-Control-Allow-Origin': '*',
          'Content-Length': buffer.length,
        });
        res.end(buffer);
        console.log('[' + new Date().toLocaleTimeString() + '] ✓ docx生成 ' + buffer.length + ' bytes');
      } catch(e) {
        console.error('生成docx失败:', e);
        res.writeHead(500, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
        res.end(JSON.stringify({ error: e.message }));
      }
    });
    return;
  }

  // GET / → HTML
  if (req.method === 'GET' && (req.url === '/' || req.url === '/index.html')) {
    if (!fs.existsSync(HTML_FILE)) { res.writeHead(404); res.end('找不到秒著.html'); return; }
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(fs.readFileSync(HTML_FILE, 'utf-8')); return;
  }

  res.writeHead(404); res.end('Not found');
});

server.listen(PORT, () => {
  console.log('\n  🌋 秒著 v3 已启动');
  console.log('  ──────────────────────────────────');
  console.log('  👉 http://localhost:' + PORT);
  console.log('  ──────────────────────────────────');
  console.log('  · 自动提取上传文档中的图片并嵌入docx');
  console.log('  · 人员信息从文档原文提取，不编造');
  console.log('  · 按 Ctrl+C 停止\n');
  require('child_process').exec('open http://localhost:' + PORT);
});