const XLSX = require('xlsx');
const PptxGenJS = require('pptxgenjs');
const path = require('path');

// ============ READ & PARSE DATA ============
const filePath = path.join(process.env.USERPROFILE, 'Desktop', 'create_recrod_brazil6031659.xlsx');
const wb = XLSX.readFile(filePath);
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws);
const record = data.filter(r => r.countryid === 'japan06032028' && String(r.number) === '1')[0];
const content = record.information;

const lines = content.split('\n');
const subsections = [];
let currentSub = null;

lines.forEach(line => {
  line = line.trim();
  if (!line || line.startsWith('### ')) return;
  if (line.startsWith('#### ')) {
    if (currentSub) subsections.push(currentSub);
    currentSub = { title: line.replace('#### ', ''), points: [] };
  } else if (currentSub) {
    let cleaned = line.replace(/【来源：.+?】/g, '').replace(/【发布时间：.+?】/g, '');
    cleaned = cleaned.trim();
    if (cleaned && cleaned.length > 10) {
      const sentences = cleaned.split('。').filter(s => s.trim());
      const take = Math.min(sentences.length, 3);
      for (let j = 0; j < take; j++) {
        const s2 = sentences[j].trim();
        if (s2 && s2.length > 6) currentSub.points.push(s2 + '。');
      }
    }
  }
});
if (currentSub) subsections.push(currentSub);

const part1 = subsections.slice(0, 4);   // macro environment
const part2 = subsections.slice(4, 7);   // social culture
const part3 = subsections.slice(7);      // China-Japan relations

console.log('Part 1 (macro):', part1.length, 'subsections,', part1.reduce((a,b)=>a+b.points.length,0), 'points');
console.log('Part 2 (social):', part2.length, 'subsections,', part2.reduce((a,b)=>a+b.points.length,0), 'points');
console.log('Part 3 (relations):', part3.length, 'subsections,', part3.reduce((a,b)=>a+b.points.length,0), 'points');

// ============ PPT CONFIG ============
const pptx = new PptxGenJS();
const SW = 13.33;
const SH = 7.5;

const C = {
  navy:      '1B2A4A',
  blue:      '3A7BD5',
  lightBlue: 'E8F1FC',
  red:       'D94040',
  gold:      'C49B4C',
  white:     'FFFFFF',
  offWhite:  'F7F8FA',
  dark:      '1A1A2E',
  gray:      '6B7280',
  lightGray: 'D1D5DB',
  green:     '059669',
  teal:      '0D9488',
};
const F_T = 'Microsoft YaHei';
const F_B = 'Microsoft YaHei';

// ===== HELPERS =====
function addCard(slide, x, y, w, h, opts = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: opts.fill || C.offWhite },
    shadow: { type: 'outer', blur: 3, offset: 1, color: '000000', opacity: 0.06 },
    rectRadius: 0.06
  });
  if (opts.accentColor) {
    slide.addShape(pptx.ShapeType.rect, { x, y, w, h: 0.04, fill: { color: opts.accentColor } });
  }
}

function addKeyInsight(slide, x, y, w, text, color) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h: 0.65,
    fill: { color: color, transparency: 92 },
    line: { color: color, width: 1, transparency: 50 },
    rectRadius: 0.04
  });
  slide.addText('💡 ' + text, {
    x: x + 0.15, y: y + 0.02, w: w - 0.3, h: 0.6,
    fontSize: 11, color: C.dark, fontFace: F_B, valign: 'middle'
  });
}

// ===== SECTION DIVIDER =====
function addSectionDivider(num, title, subtitle, color, icon) {
  const s = pptx.addSlide();
  s.background = { fill: color };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.06, fill: { color: 'FFFFFF', transparency: 60 } });
  s.addText(icon, { x: 1.0, y: 1.8, w: 2, h: 2, fontSize: 48 });
  s.addText('PART ' + num, {
    x: 1.0, y: 3.0, w: 5, h: 0.5,
    fontSize: 13, color: 'FFFFFF', fontFace: 'Arial', charSpacing: 5, transparency: 40
  });
  s.addText(title, {
    x: 1.0, y: 3.4, w: 11, h: 1.2,
    fontSize: 36, bold: true, color: C.white, fontFace: F_T
  });
  if (subtitle) {
    s.addText(subtitle, {
      x: 1.0, y: 4.5, w: 11, h: 0.5,
      fontSize: 13, color: 'FFFFFF', fontFace: F_B, transparency: 30
    });
  }
  s.addShape(pptx.ShapeType.rect, { x: 0, y: SH - 0.06, w: SW, h: 0.06, fill: { color: 'FFFFFF', transparency: 60 } });
}

// ===== CONTENT SLIDE (1-2 column layout) =====
function addContentSlide(subsection, index, accentColor) {
  const s = pptx.addSlide();
  s.background = { fill: C.white };

  // Top colored header
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SW, h: 1.1, fill: { color: accentColor } });

  // Number badge
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.5, y: 0.18, w: 0.68, h: 0.68,
    fill: { color: 'FFFFFF', transparency: 80 },
    rectRadius: 0.34
  });
  s.addText(String(index + 1), {
    x: 0.5, y: 0.18, w: 0.68, h: 0.68,
    fontSize: 18, bold: true, color: C.white, align: 'center', valign: 'middle', fontFace: 'Arial'
  });

  // Title in header
  s.addText(subsection.title, {
    x: 1.4, y: 0.10, w: 11, h: 0.85,
    fontSize: 20, bold: true, color: C.white, fontFace: F_T, valign: 'middle'
  });

  // Content area
  const pts = subsection.points;
  const mid = Math.ceil(pts.length / 2);
  const leftPts = pts.slice(0, mid);
  const rightPts = pts.slice(mid);
  const useTwoCol = rightPts.length > 0;
  const colW = useTwoCol ? 5.7 : 11.8;

  function drawPoints(slide, points, x, y) {
    points.forEach((pt, i) => {
      // dot
      slide.addShape(pptx.ShapeType.ellipse, {
        x: x, y: y + i * 1.15 + 0.12, w: 0.16, h: 0.16,
        fill: { color: accentColor }
      });
      // text
      slide.addText(pt, {
        x: x + 0.32, y: y + i * 1.15, w: colW - 0.38, h: 0.7,
        fontSize: 12.5, color: C.dark, fontFace: F_B, valign: 'middle', lineSpacing: 20
      });
    });
  }

  const startY = 1.5;
  drawPoints(s, leftPts, 0.7, startY);

  if (useTwoCol) {
    s.addShape(pptx.ShapeType.rect, {
      x: 6.55, y: startY, w: 0.015, h: Math.max(leftPts.length, rightPts.length) * 1.15,
      fill: { color: C.lightGray }
    });
    drawPoints(s, rightPts, 6.8, startY);
  }

  // Key insight at bottom
  const maxRows = Math.max(leftPts.length, rightPts.length);
  const bottomY = startY + maxRows * 1.15 + 0.3;
  const firstPt = pts[0] || '';
  const short = firstPt.length > 65 ? firstPt.substring(0, 65) + '...' : firstPt;
  addKeyInsight(s, 0.7, bottomY, 11.8, short, accentColor);

  // Page number
  s.addText(String(index + 3), {
    x: 12.2, y: SH - 0.45, w: 0.8, h: 0.35,
    fontSize: 9, color: C.lightGray, align: 'right', fontFace: 'Arial'
  });
}

// ===== KEY TAKEAWAYS SLIDE =====
function addTakeawaysSlide(partTitle, takeaways, accentColor) {
  const s = pptx.addSlide();
  s.background = { fill: C.white };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.06, fill: { color: accentColor } });
  s.addText('核心洞察', {
    x: 0.6, y: 0.25, w: 5, h: 0.75, fontSize: 26, bold: true, color: C.navy, fontFace: F_T
  });
  s.addText(partTitle + '  /  KEY INSIGHTS', {
    x: 0.6, y: 0.85, w: 8, h: 0.35, fontSize: 11, color: C.gray, fontFace: 'Arial'
  });
  s.addShape(pptx.ShapeType.rect, { x: 0.6, y: 1.3, w: 2.5, h: 0.03, fill: { color: accentColor } });

  takeaways.forEach((t, i) => {
    const y = 1.7 + i * 1.7;
    addCard(s, 0.7, y, 11.8, 1.5, { accentColor: accentColor, fill: C.offWhite });
    s.addText('0' + (i + 1), {
      x: 1.0, y: y + 0.15, w: 1.0, h: 1.1,
      fontSize: 34, bold: true, color: accentColor, fontFace: 'Arial', valign: 'middle'
    });
    // Vertical accent line
    s.addShape(pptx.ShapeType.rect, {
      x: 2.1, y: y + 0.3, w: 0.04, h: 0.9,
      fill: { color: accentColor, transparency: 50 }
    });
    s.addText(t, {
      x: 2.4, y: y + 0.15, w: 9.5, h: 1.1,
      fontSize: 13.5, color: C.dark, fontFace: F_B, valign: 'middle', lineSpacing: 22
    });
  });
}

// ===== SUMMARY SLIDE =====
function addSummarySlide(allTakeaways) {
  const s = pptx.addSlide();
  s.background = { fill: C.white };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.06, fill: { color: C.navy } });

  s.addText('总  结', {
    x: 0.6, y: 0.25, w: 5, h: 0.8, fontSize: 30, bold: true, color: C.navy, fontFace: F_T
  });
  s.addText('SUMMARY', {
    x: 0.6, y: 0.85, w: 5, h: 0.35, fontSize: 11, color: C.gray, fontFace: 'Arial', charSpacing: 4
  });
  s.addShape(pptx.ShapeType.rect, { x: 0.6, y: 1.3, w: 2.5, h: 0.03, fill: { color: C.red } });

  const colors = [C.navy, C.blue, C.teal];
  const icons = ['01', '02', '03'];
  const labels = ['宏观环境', '社会文化', '中日政商'];

  allTakeaways.forEach((t, i) => {
    const y = 1.7 + i * 1.7;
    addCard(s, 0.7, y, 11.8, 1.5, { accentColor: colors[i], fill: C.offWhite });

    // Icon circle
    s.addShape(pptx.ShapeType.ellipse, {
      x: 1.1, y: y + 0.35, w: 0.8, h: 0.8,
      fill: { color: colors[i] }
    });
    s.addText(icons[i], {
      x: 1.1, y: y + 0.35, w: 0.8, h: 0.8,
      fontSize: 18, bold: true, color: C.white, align: 'center', valign: 'middle', fontFace: 'Arial'
    });

    // Label
    s.addText(labels[i], {
      x: 2.2, y: y + 0.15, w: 2.5, h: 0.45,
      fontSize: 14, bold: true, color: colors[i], fontFace: F_T
    });
    // Content
    s.addText(t, {
      x: 2.2, y: y + 0.6, w: 9.5, h: 0.75,
      fontSize: 12.5, color: C.dark, fontFace: F_B, valign: 'top', lineSpacing: 20
    });
  });
}

// ===== THANK YOU SLIDE =====
function addThankYouSlide() {
  const s = pptx.addSlide();
  s.background = { fill: C.navy };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.06, fill: { color: C.gold } });
  s.addShape(pptx.ShapeType.rect, { x: 5.5, y: 2.3, w: 2.3, h: 0.03, fill: { color: C.red } });
  s.addText('感谢聆听', {
    x: 0, y: 2.6, w: SW, h: 1.5,
    fontSize: 42, bold: true, color: C.white, align: 'center', fontFace: F_T
  });
  s.addText('THANK YOU', {
    x: 0, y: 3.8, w: SW, h: 0.6,
    fontSize: 15, color: C.lightGray, align: 'center', fontFace: 'Arial', charSpacing: 7
  });
  s.addText('日本国情概况  |  第一章 完', {
    x: 0, y: 5.5, w: SW, h: 0.5,
    fontSize: 13, color: C.gray, align: 'center', fontFace: F_T
  });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: SH - 0.06, w: SW, h: 0.06, fill: { color: C.gold } });
}

// ==================== BUILD ALL SLIDES ====================

console.log('Generating slides...');

// SLIDE 1: COVER
{
  const s = pptx.addSlide();
  s.background = { fill: C.navy };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.06, fill: { color: C.gold } });
  s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.8, w: 0.06, h: 2.6, fill: { color: C.red } });
  s.addText('日本国情概况', {
    x: 1.3, y: 1.5, w: 10, h: 1.4,
    fontSize: 44, bold: true, color: C.white, fontFace: F_T
  });
  s.addText('JAPAN  /  COUNTRY PROFILE', {
    x: 1.3, y: 2.7, w: 10, h: 0.55,
    fontSize: 14, color: C.lightGray, fontFace: 'Arial', charSpacing: 5
  });
  s.addShape(pptx.ShapeType.roundRect, {
    x: 1.3, y: 4.0, w: 6.8, h: 0.8,
    fill: { color: 'FFFFFF', transparency: 92 },
    line: { color: C.gold, width: 1 },
    rectRadius: 0.04
  });
  s.addText('日本市场进入指南  |  第一章  |  建议课时：30分钟', {
    x: 1.5, y: 4.1, w: 6.3, h: 0.6,
    fontSize: 13, color: C.lightGray, fontFace: F_T
  });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: SH - 0.06, w: SW, h: 0.06, fill: { color: C.gold } });
  s.addShape(pptx.ShapeType.ellipse, {
    x: 10.5, y: 1.5, w: 2.2, h: 2.2,
    fill: { color: C.red }
  });
}

// SLIDE 2: AGENDA
{
  const s = pptx.addSlide();
  s.background = { fill: C.white };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.06, fill: { color: C.navy } });
  s.addText('目  录', { x: 0.6, y: 0.3, w: 5, h: 0.8, fontSize: 30, bold: true, color: C.navy, fontFace: F_T });
  s.addText('AGENDA', { x: 0.6, y: 0.85, w: 5, h: 0.35, fontSize: 11, color: C.gray, fontFace: 'Arial', charSpacing: 4 });
  s.addShape(pptx.ShapeType.rect, { x: 0.6, y: 1.3, w: 2.5, h: 0.03, fill: { color: C.red } });

  const items = [
    { num: '01', title: '宏观环境', sub: 'PEST Analysis', tags: '历史快览  |  政经格局  |  经济脉搏  |  宗教底色', color: C.navy, icon: '🌏' },
    { num: '02', title: '社会文化', sub: 'Social & Culture', tags: '价值观核心  |  行为规范  |  节庆风俗', color: C.blue, icon: '🏯' },
    { num: '03', title: '中日政商关系', sub: 'Diplomacy & Business', tags: '外交脉络  |  经贸版图  |  政策环境  |  风险提示', color: C.teal, icon: '🤝' },
  ];

  items.forEach((item, i) => {
    const x = 0.6 + i * 4.1;
    const y = 2.1;
    addCard(s, x, y, 3.8, 4.2, { accentColor: item.color, fill: C.offWhite });
    s.addText(item.num, { x: x + 0.3, y: y + 0.25, w: 1.5, h: 1.0, fontSize: 42, bold: true, color: item.color, fontFace: 'Arial' });
    s.addText(item.icon, { x: x + 2.0, y: y + 0.25, w: 1.5, h: 1.0, fontSize: 36 });
    s.addText(item.title, { x: x + 0.3, y: y + 1.5, w: 3.2, h: 0.7, fontSize: 22, bold: true, color: C.dark, fontFace: F_T });
    s.addText(item.sub, { x: x + 0.3, y: y + 2.1, w: 3.2, h: 0.35, fontSize: 11, color: C.gray, fontFace: 'Arial' });
    s.addShape(pptx.ShapeType.rect, { x: x + 0.3, y: y + 2.6, w: 1.5, h: 0.02, fill: { color: item.color } });
    s.addText(item.tags, { x: x + 0.3, y: y + 2.9, w: 3.2, h: 1.0, fontSize: 10, color: C.gray, fontFace: F_B, lineSpacing: 20 });
  });
}

// PART 1: Macro Environment
addSectionDivider('01', '宏观环境', 'PEST Analysis — 历史 · 政经 · 经济 · 宗教', C.navy, '🌏');
part1.forEach((sub, i) => addContentSlide(sub, i, C.navy));
addTakeawaysSlide('宏观环境', [
  '日本政局稳定、法制健全，为中资企业提供了低风险的投资环境',
  '制造业与高科技是核心产业，GDP增长稳定，营商环境全球排名领先',
  '神道教与佛教影响温和，商业节奏不受宗教干扰，但需尊重当地习俗',
  '中日历史渊源深厚，理解和尊重当地文化是业务成功的关键第一步',
], C.navy);

// PART 2: Social Culture
addSectionDivider('02', '社会文化', 'Social & Culture — 价值观 · 行为规范 · 节庆风俗', C.blue, '🏯');
part2.forEach((sub, i) => addContentSlide(sub, i, C.blue));
addTakeawaysSlide('社会文化', [
  '集体主义文化明显，团队决策和上级指示在职场中占主导地位',
  '时间观念极强、迟到容忍度低；性别角色趋于平等但仍需注意传统规范',
  '法定节假日和宗教节日影响工作节奏，避免在此期间安排重要商务活动',
], C.blue);

// PART 3: China-Japan Relations
addSectionDivider('03', '中日政商关系', 'Diplomacy & Business — 外交 · 经贸 · 政策 · 风险', C.teal, '🤝');
part3.forEach((sub, i) => addContentSlide(sub, i, C.teal));
addTakeawaysSlide('中日政商关系', [
  '中日双边贸易持续增长，中国是日本主要贸易伙伴，投资涵盖制造与服务领域',
  '日本对外资准入政策较为开放，无专门负面清单，近期政策环境相对稳定',
  '需关注合规风险与地缘政治变化，善用驻日使馆经商处和JETRO等本地资源',
], C.teal);

// SUMMARY
addSummarySlide([
  '宏观环境：政局稳定、经济发达、宗教影响温和，为中资企业提供了良好的投资基础',
  '社会文化：集体主义导向、时间严谨、尊重节庆，外派员工需充分适应本地行为规范',
  '中日关系：贸易紧密、政策开放、本地资源丰富，充分利用使馆经商处与JETRO渠道',
]);

// THANK YOU
addThankYouSlide();

// ===== SAVE =====
const outputPath = path.join(process.env.USERPROFILE, 'Desktop', '日本国情概况.pptx');
pptx.writeFile({ fileName: outputPath }).then(() => {
  console.log('DONE: PPT saved to', outputPath);
}).catch(err => {
  console.error('Error:', err);
});
