const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, AlignmentType, PageBreak } = require('docx');
const fs = require('fs');
const path = require('path');

const imageFiles = [
  '软件界面/1.首页.jpg',
  '软件界面/2.第二页上传文件.jpg',
  '软件界面/3.第三页确认信息页.jpg',
  '软件界面/4.第四页生成中.jpg',
  '软件界面/5.4.第五页生成结果.jpg'
];

const imageNames = [
  '图1：首页 - 配置AI模型',
  '图2：上传文件页 - 上传项目文档',
  '图3：确认信息页 - 核对提取的信息',
  '图4：生成中页 - 实时预览生成进度',
  '图5：生成结果页 - 下载软著材料'
];

async function createImageParagraph(imagePath, caption, width, height) {
  const data = fs.readFileSync(imagePath);
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [
      new ImageRun({
        type: 'jpg',
        data: data,
        transformation: { width: width, height: height },
        altText: { title: caption, description: caption, name: caption }
      })
    ]
  });
}

async function generateDoc() {
  const children = [];

  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 200 },
    children: [
      new TextRun({
        text: '秒著',
        bold: true,
        size: 72,
        font: 'Arial'
      })
    ]
  }));

  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [
      new TextRun({
        text: 'AI软著材料生成工具',
        size: 36,
        font: 'Arial',
        color: '666666'
      })
    ]
  }));

  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 800, after: 100 },
    children: [
      new TextRun({
        text: '产品介绍文档',
        size: 28,
        font: 'Arial'
      })
    ]
  }));

  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 100 },
    children: [
      new TextRun({
        text: '内部使用',
        size: 24,
        font: 'Arial',
        color: '999999'
      })
    ]
  }));

  children.push(new Paragraph({
    children: [new PageBreak()]
  }));

  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun('目录')]
  }));

  const tocItems = [
    '1. 产品简介',
    '2. 主要功能',
    '3. 软件界面',
    '4. 使用指南'
  ];

  tocItems.forEach(item => {
    children.push(new Paragraph({
      spacing: { before: 120, after: 60 },
      children: [new TextRun({ text: item, size: 24, font: 'Arial' })]
    }));
  });

  children.push(new Paragraph({
    children: [new PageBreak()]
  }));

  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun('1. 产品简介')]
  }));

  children.push(new Paragraph({
    spacing: { after: 200 },
    children: [
      new TextRun({
        text: '秒著是一款基于AI的软件著作权申请材料自动生成工具，支持一键生成软件概述和设计说明书。',
        size: 24,
        font: 'Arial'
      })
    ]
  }));

  children.push(new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [
      new TextRun({
        text: '核心优势：',
        bold: true,
        size: 24,
        font: 'Arial'
      })
    ]
  }));

  const advantages = [
    '免费模型：Step 3.5 Flash 完全免费，无需配置API Key',
    '多AI模型支持：火山引擎、DeepSeek、智谱GLM、Kimi、MiniMax',
    '流式输出：实时预览生成内容，降低等待焦虑',
    '安全可靠：API Key仅在当前会话使用，不保存到本地'
  ];

  advantages.forEach(item => {
    children.push(new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: '• ' + item, size: 24, font: 'Arial' })]
    }));
  });

  children.push(new Paragraph({
    children: [new PageBreak()]
  }));

  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun('2. 主要功能')]
  }));

  const features = [
    { title: '自动信息提取', desc: '从项目文档中自动提取软件名称、版本、开发环境等信息' },
    { title: '软著材料生成', desc: '自动生成符合规范的软件概述和设计说明书' },
    { title: '图片嵌入', desc: '自动提取并嵌入文档中的图片到生成材料中' },
    { title: '智能目录生成', desc: '自动生成规范的项目目录和页码' },
    { title: '多格式支持', desc: '支持.doc、.docx、.pdf格式文档上传' },
    { title: '实时预览', desc: '流式输出，实时显示生成进度和预估剩余时间' }
  ];

  features.forEach((feature, index) => {
    children.push(new Paragraph({
      spacing: { before: 200, after: 60 },
      children: [
        new TextRun({ text: (index + 1) + '. ' + feature.title, bold: true, size: 24, font: 'Arial' })
      ]
    }));
    children.push(new Paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun({ text: feature.desc, size: 24, font: 'Arial' })
      ]
    }));
  });

  children.push(new Paragraph({
    children: [new PageBreak()]
  }));

  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun('3. 软件界面')]
  }));

  for (let i = 0; i < imageFiles.length; i++) {
    const imagePath = imageFiles[i];
    if (fs.existsSync(imagePath)) {
      children.push(await createImageParagraph(imagePath, imageNames[i], 450, 280));
      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [
          new TextRun({
            text: imageNames[i],
            size: 20,
            font: 'Arial',
            color: '666666',
            italics: true
          })
        ]
      }));
    }
  }

  children.push(new Paragraph({
    children: [new PageBreak()]
  }));

  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun('4. 使用指南')]
  }));

  const steps = [
    { step: '第一步：配置模型', content: '选择AI提供商（推荐使用免费的Step 3.5 Flash），输入API Key，点击"测试连接"验证后保存。' },
    { step: '第二步：上传文件', content: '上传项目文档（支持.doc、.docx、.pdf格式），或直接粘贴文档内容，点击"AI提取信息"。' },
    { step: '第三步：确认信息', content: '核对AI提取的软件信息，如有错误可手动修改，确认无误后点击"生成软著材料"。' },
    { step: '第四步：下载文档', content: '生成完成后，可预览生成的软件概述和设计说明书，点击下载即可获取完整材料。' }
  ];

  steps.forEach((item, index) => {
    children.push(new Paragraph({
      spacing: { before: 240, after: 80 },
      children: [
        new TextRun({ text: item.step, bold: true, size: 24, font: 'Arial' })
      ]
    }));
    children.push(new Paragraph({
      spacing: { after: 160 },
      children: [
        new TextRun({ text: item.content, size: 24, font: 'Arial' })
      ]
    }));
  });

  children.push(new Paragraph({
    spacing: { before: 300 },
    children: [
      new TextRun({
        text: '如有问题，请联系开发者。',
        size: 24,
        font: 'Arial',
        color: '666666'
      })
    ]
  }));

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: 'Arial', size: 24 }
        }
      },
      paragraphStyles: [
        {
          id: 'Heading1',
          name: 'Heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: { size: 36, bold: true, font: 'Arial' },
          paragraph: { spacing: { before: 300, after: 200 }, outlineLevel: 0 }
        }
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: children
    }]
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync('秒著产品介绍.docx', buffer);
  console.log('文档已生成：秒著产品介绍.docx');
}

generateDoc().catch(console.error);