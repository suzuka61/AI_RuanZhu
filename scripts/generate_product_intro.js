const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  Table, TableRow, TableCell, WidthType, BorderStyle,
  AlignmentType, PageBreak, TabStopType, TabStopPosition,
  ImageRun, Header, Footer, PageNumber, NumberFormat,
  LevelFormat, convertInchesToTwip, ShadingType,
  TableOfContents, StyleLevel, UnderlineType,
  VerticalAlign, PageOrientation
} = require('docx');
const fs = require('fs');
const path = require('path');

const outputPath = '/Users/songtao/Documents/软著生成器/秒著产品介绍.docx';
const imageBasePath = '/Users/songtao/Documents/软著生成器/软件界面';

function createFeatureTable(headers, data) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 8, color: 'BFBFBF' },
      bottom: { style: BorderStyle.SINGLE, size: 8, color: 'BFBFBF' },
      left: { style: BorderStyle.NIL },
      right: { style: BorderStyle.NIL },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: 'D9D9D9' },
      insideVertical: { style: BorderStyle.NIL },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map(h => new TableCell({
          children: [new Paragraph({
            children: [new TextRun({ text: h, bold: true })],
          })],
          borders: {
            bottom: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
          },
        })),
      }),
      ...data.map((row, i) => new TableRow({
        children: row.map(cell => new TableCell({
          children: [new Paragraph({
            children: [new TextRun({ text: cell })],
          })],
          shading: i % 2 === 1 ? { fill: 'F2F2F2', type: ShadingType.CLEAR } : undefined,
        })),
      })),
    ],
  });
}

function createSectionHeading(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 120 },
    children: [new TextRun({ text, font: 'Aptos Display' })],
  });
}

function createHeading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text })],
  });
}

function createBodyParagraph(text) {
  return new Paragraph({
    spacing: { after: 160 },
    children: [new TextRun({ text })],
  });
}

function createBulletItem(text) {
  return new Paragraph({
    bullet: { level: 0 },
    children: [new TextRun({ text })],
  });
}

function createWarningParagraph(text) {
  return new Paragraph({
    spacing: { after: 160 },
    children: [new TextRun({ text, italics: true, color: '78786C' })],
  });
}

function createImageParagraph(imagePath, caption) {
  const fullPath = path.join(imageBasePath, imagePath);
  if (!fs.existsSync(fullPath)) {
    return new Paragraph({
      children: [new TextRun({ text: `[图片未找到: ${imagePath}]`, color: 'FF0000' })],
    });
  }

  const imageBuffer = fs.readFileSync(fullPath);
  const ext = path.extname(imagePath).toLowerCase();

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 120 },
    children: [
      new ImageRun({
        data: imageBuffer,
        transformation: { width: 550, height: 350 },
        type: ext === '.png' ? 'png' : 'image',
      }),
    ],
  });
}

const doc = new Document({
  styles: {
    paragraphStyles: [
      {
        id: 'Normal',
        name: 'Normal',
        run: {
          font: 'Aptos',
          size: 22,
        },
        paragraph: {
          spacing: { line: 276, after: 160 },
        },
      },
      {
        id: 'Heading1',
        name: 'Heading 1',
        basedOn: 'Normal',
        next: 'Normal',
        run: {
          font: 'Aptos Display',
          size: 40,
          color: '1F3864',
          bold: false,
        },
        paragraph: {
          spacing: { before: 480, after: 120 },
          outlineLevel: 0,
        },
      },
      {
        id: 'Heading2',
        name: 'Heading 2',
        basedOn: 'Normal',
        next: 'Normal',
        run: {
          font: 'Aptos Display',
          size: 32,
          color: '1F3864',
          bold: false,
        },
        paragraph: {
          spacing: { before: 360, after: 80 },
          outlineLevel: 1,
        },
      },
      {
        id: 'Heading3',
        name: 'Heading 3',
        basedOn: 'Normal',
        next: 'Normal',
        run: {
          font: 'Aptos',
          size: 26,
          color: '1F3864',
          bold: true,
        },
        paragraph: {
          spacing: { before: 240, after: 80 },
          outlineLevel: 2,
        },
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 },
      },
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [
              new TextRun({ text: '第 ', size: 18, color: '808080' }),
              new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '808080' }),
              new TextRun({ text: ' 页', size: 18, color: '808080' }),
            ],
          }),
        ],
      }),
    },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 2400, after: 400 },
        children: [
          new TextRun({
            text: '秒著',
            font: 'Aptos Display',
            size: 72,
            color: '1F3864',
            bold: true,
          }),
        ],
      }),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [
          new TextRun({
            text: 'AI软著材料生成工具',
            size: 32,
            color: '5D7052',
          }),
        ],
      }),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 800 },
        children: [
          new TextRun({
            text: '产品介绍 | 版本 v1.8 | 2026年3月',
            size: 24,
            color: '78786C',
          }),
        ],
      }),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 1200 },
        children: [
          new TextRun({
            text: '让软著申请更简单',
            size: 22,
            color: '78786C',
          }),
        ],
      }),

      new Paragraph({ children: [] }),

      createSectionHeading('产品概述'),
      createBodyParagraph('秒著是一款基于AI的软件著作权申请材料自动生成工具。用户上传项目文档（.doc/.docx/.pdf），系统通过AI提取关键信息，自动生成符合国家版权局要求的「软件概述」和「设计说明书」两份核心申请材料，并导出为标准格式的Word文档。'),

      createSectionHeading('核心价值'),
      createFeatureTable(
        ['核心价值', '说明'],
        [
          ['高效便捷', '将传统需要1-3天的软著材料撰写工作缩短至几分钟'],
          ['降低门槛', '降低软著申请的专业门槛，无需专业写作背景'],
          ['规范合规', '生成符合官方规范的标准化文档格式'],
          ['多模型支持', '支持多种AI模型，包括免费模型'],
        ]
      ),

      createSectionHeading('主要功能'),

      createHeading3('1. 免费模型支持'),
      createBodyParagraph('秒著内置 Step 3.5 Flash 免费模型，用户无需配置API Key即可使用。同时支持火山引擎、DeepSeek、智谱GLM、Kimi、MiniMax等多种AI模型，满足不同用户需求。'),

      createHeading3('2. 智能文档解析'),
      createBodyParagraph('支持上传 .doc、.docx、.pdf 格式的项目文档，系统自动解析文档内容并提取关键信息，包括软件名称、版本号、开发者信息、主要功能描述等。'),

      createHeading3('3. 自动生成材料'),
      createBodyParagraph('AI自动生成符合国家版权局要求的软著申请材料：'),
      createBulletItem('软件概述：包含软件基本信息表，开发目的、主要功能、技术特点'),
      createBulletItem('设计说明书：包含项目背景、目标、范围、时间表、需求描述及界面展示'),

      createHeading3('4. 图片智能匹配'),
      createBodyParagraph('系统自动从上传文档中提取图片，并根据章节标题进行智能匹配，将图片嵌入到生成的文档对应位置。未匹配的图片会自动归入附录。'),

      createHeading3('5. 软件类型模板系统'),
      createBodyParagraph('内置10种软件类型模板，自动匹配并填充缺失的技术信息：'),
      createFeatureTable(
        ['模板类型', '匹配关键词'],
        [
          ['管理系统', '管理、ERP、CRM、OA、HRM、进销存、仓储、财务'],
          ['微信小程序', '小程序、微信、扫码、点餐、预约、商城'],
          ['移动APP', 'APP、Android、iOS、手机、移动端'],
          ['网站平台', '网站、门户、官网、商城、电商'],
          ['教育系统', '教育、学习、培训、课程、考试、在线教育'],
          ['医疗健康', '医疗、医院、健康、诊疗、挂号、病历'],
          ['金融系统', '金融、银行、支付、理财、贷款、保险'],
          ['物流系统', '物流、快递、配送、仓储、运输、供应链'],
          ['物联网系统', '物联网、IoT、智能、设备、传感器'],
          ['通用软件', '默认模板，适用于其他类型软件'],
        ]
      ),

      createSectionHeading('产品界面'),
      createBodyParagraph('秒著采用简洁直观的四步操作流程：'),

      createImageParagraph('1.首页.jpg', '秒著首页 - 配置API'),
      createHeading3('Step 1：配置API'),
      createBodyParagraph('选择AI提供商并输入API Key。推荐使用免费的 Step 3.5 Flash 模型，无需配置即可使用。'),

      createImageParagraph('2.第二页上传文件.jpg', '上传文件界面'),
      createHeading3('Step 2：上传文件'),
      createBodyParagraph('支持拖拽或点击上传项目文档（.doc/.docx/.pdf），也可以直接粘贴文档文本内容。'),

      createImageParagraph('3.第三页确认信息页.jpg', '确认信息界面'),
      createHeading3('Step 3：确认信息'),
      createBodyParagraph('AI自动提取文档信息，用户可以核对并修改各项内容，确保信息准确无误。'),

      createImageParagraph('4.第四页生成中.jpg', '文档生成界面'),
      createHeading3('Step 4：生成文档'),
      createBodyParagraph('实时预览生成进度，支持流式输出，字数统计和预估剩余时间显示。生成完成后可预览、复制或下载文档。'),

      createImageParagraph('5.4.第五页生成结果.jpg', '生成结果界面'),

      createSectionHeading('技术架构'),
      createFeatureTable(
        ['组件', '技术选型', '说明'],
        [
          ['前端', '单HTML文件（原生JS）', '零依赖，浏览器直接运行'],
          ['后端代理', 'Node.js (proxy.js)', '解决跨域 + 生成docx'],
          ['AI服务', '火山引擎API', '支持多模型切换'],
          ['文档解析', 'mammoth.js / pdf.js', '前端解析docx和pdf'],
          ['文档生成', 'docx (npm)', '服务端生成Word文档'],
          ['图片处理', 'JSZip', '从docx中提取嵌入图片'],
        ]
      ),

      createSectionHeading('适用场景'),
      createBodyParagraph('秒著适用于以下用户群体：'),
      createBulletItem('需要申请软著的个人开发者'),
      createBulletItem('中小企业技术负责人'),
      createBulletItem('知识产权代理机构'),

      createSectionHeading('使用流程'),
      createFeatureTable(
        ['步骤', '操作', '说明'],
        [
          ['1', '启动应用', '运行 node proxy.js，访问 http://localhost:3000'],
          ['2', '配置模型', '选择API提供商，输入API Key（可选免费模型）'],
          ['3', '上传文档', '上传项目文档或粘贴文本内容'],
          ['4', '提取信息', 'AI自动解析并提取关键信息'],
          ['5', '确认修改', '核对并修改提取的信息'],
          ['6', '生成材料', '一键生成软著申请材料'],
          ['7', '下载使用', '下载生成的Word文档进行后续申请'],
        ]
      ),

      createSectionHeading('支持的文件格式'),
      createBodyParagraph('输入文件：'),
      createBulletItem('.doc - Microsoft Word 97-2003 文档'),
      createBulletItem('.docx - Microsoft Word 2007 及以上文档'),
      createBulletItem('.pdf - 便携式文档格式'),
      createBulletItem('单文件大小限制：10MB'),
      new Paragraph({ spacing: { after: 80 }, children: [] }),
      createBodyParagraph('输出文档：'),
      createBulletItem('.docx - 符合国家版权局要求的软著申请材料'),

      createSectionHeading('注意事项'),
      createWarningParagraph('免责声明：生成的文档内容仅供参考和学习使用，用户需自行核对内容准确性，并确保符合国家版权局的相关规定。'),
    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log(`文档已生成: ${outputPath}`);
}).catch(err => {
  console.error('生成失败:', err);
  process.exit(1);
});