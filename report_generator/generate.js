// Self-contained industrial training report generator.

const path = require('path');

const fs = require('fs');
const D = require('docx');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, ImageRun, Footer, PageNumber, TableOfContents, HeadingLevel, LevelFormat,
  NumberFormat, VerticalAlign, PageBreak, SectionType
} = D;

const FONT = 'STIX Two Text';
const TW = 8666; // A4 text width with 1.25" left and 1" right margins

// Replace the highlighted values below before generating a completed report.
const INFO = {
  university: 'Islamic University of Technology',
  department: 'Department of Computer Science and Engineering',
  degree: 'BSc in Computer Science and Engineering',
  studentName: '[FILL: Student name]',
  studentId: '[FILL: Student ID]',
  organization: '[FILL: Organization name]',
  program: '[FILL: Training program]',
  group: '[FILL: Group or cohort]',
  startDate: '[FILL: Training start date]',
  endDate: '[FILL: Training end date]',
  mode: '[FILL: Training mode]',
  venue: '[FILL: Training venue]',
  supervisor: '[FILL: Industry supervisor]',
  submissionDate: '[FILL: Date of submission]',
  projectName: '[FILL: Main project name]',
};
const programLabel = `${INFO.program} (${INFO.group})`;

// ---------- inline parser: **bold**, *italic*, §code§, [FILL: ...] ----------
function runs(text, base = {}) {
  const parts = text.split(/(\[FILL:[^\]]*\]|\*\*[^*]+\*\*|\*[^*]+\*|§[^§]+§)/g).filter(s => s !== '');
  return parts.map(p => {
    if (p.startsWith('[FILL:')) return new TextRun({ ...base, text: p, highlight: 'yellow' });
    if (p.startsWith('**')) return new TextRun({ ...base, text: p.slice(2, -2), bold: true });
    if (p.startsWith('§')) return new TextRun({ ...base, text: p.slice(1, -1), font: 'Consolas', size: 20 });
    if (p.startsWith('*')) return new TextRun({ ...base, text: p.slice(1, -1), italics: true });
    return new TextRun({ ...base, text: p });
  });
}

const P = (text, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { before: o.before ?? 0, after: o.after ?? 120, line: 253 },
  keepNext: o.keepNext, indent: o.indent,
  children: runs(text, o.run || {}),
});
const H1 = (text, brk = true, chapter = false) => {
  const match = text.match(/^(\d+)\.\s*(.+)$/);
  const isChapter = match && chapter;
  const title = match ? (isChapter ? match[2].toUpperCase() : text.toUpperCase()) : text.toUpperCase();
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, pageBreakBefore: brk, alignment: AlignmentType.LEFT,
    spacing: { before: 480, after: 280, line: 253 },
    children: isChapter
      ? [new TextRun({ text: `CHAPTER ${match[1]}`, bold: true }), new TextRun({ text: title, bold: true, break: 1 })]
      : [new TextRun({ text: title, bold: true })],
  });
};
const H2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
const H3 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
const B = (text, lvl = 0) => new Paragraph({
  numbering: { reference: 'bul', level: lvl },
  alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, line: 253 }, children: runs(text),
});
const TCAP = (text) => new Paragraph({
  keepNext: true, alignment: AlignmentType.LEFT, spacing: { before: 200, after: 120 },
  children: [new TextRun({ text, bold: true, size: 20 })],
});
const FCAP = (text) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 60, after: 200 }, children: [new TextRun({ text, italics: true, size: 20 })],
});
const SP = (after = 120) => new Paragraph({ spacing: { after }, children: [] });

const bd = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const borders = { top: bd, left: bd, bottom: bd, right: bd };
function cell(text, w, o = {}) {
  const paras = (Array.isArray(text) ? text : [text]).map(t => new Paragraph({
    alignment: o.align || AlignmentType.LEFT, spacing: { after: 0 },
    children: runs(t, { size: 20, bold: o.bold }),
  }));
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: paras,
    columnSpan: o.span,
  });
}
function tbl(widths, header, rows, o = {}) {
  const sum = widths.reduce((a, b) => a + b, 0);
  const trs = [];
  if (header) trs.push(new TableRow({ tableHeader: true, cantSplit: true,
    children: header.map((h, i) => cell(h, widths[i], { bold: true, align: AlignmentType.CENTER })) }));
  rows.forEach(r => trs.push(new TableRow({ cantSplit: true,
    children: r.map((c, i) => cell(c, widths[i], { align: (o.center || []).includes(i) ? AlignmentType.CENTER : AlignmentType.LEFT, bold: o.boldFirst && i === 0 })) })));
  return new Table({ width: { size: sum, type: WidthType.DXA }, columnWidths: widths, rows: trs });
}
// single-cell placeholder frame (for images the student will insert)
function frame(text, h = 3000) {
  return new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: [TW], rows: [new TableRow({
    height: { value: h, rule: 'atLeast' }, cantSplit: true,
    children: [new TableCell({ width: { size: TW, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
      margins: { top: 100, bottom: 100, left: 200, right: 200 },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: runs(text, { size: 20 }) })] })],
  })] });
}
const rule = (after = 300, before = 100) => new Paragraph({
  spacing: { before, after }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: '000000', space: 4 } },
  indent: { left: 2300, right: 2300 }, children: [],
});
const CP = (text, o = {}) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: o.before ?? 0, after: o.after ?? 120 },
  keepNext: o.keepNext, children: Array.isArray(text) ? text : [new TextRun({ text, bold: o.bold, italics: o.italics, size: o.size })] });

function footer() {
  return new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
    new TextRun({ children: [PageNumber.CURRENT], size: 18 }),
  ] })] });
}

const styles = {
  default: { document: { run: { font: FONT, size: 22, color: '000000' } } },
  paragraphStyles: [
    { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: FONT, bold: true, size: 32 },
      paragraph: { alignment: AlignmentType.LEFT, spacing: { before: 480, after: 280, line: 253 }, outlineLevel: 0 } },
    { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: FONT, bold: true, size: 26 }, paragraph: { spacing: { before: 280, after: 120, line: 253 }, outlineLevel: 1, keepNext: true } },
    { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: FONT, bold: true, size: 24 }, paragraph: { spacing: { before: 200, after: 100, line: 253 }, outlineLevel: 2, keepNext: true } },
  ],
};
const numbering = { config: [{ reference: 'bul', levels: [
  { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } },
  { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 270 } } } },
] }] };

const lab = (label, value) => CP([
  new TextRun({ text: `${label}: `, bold: true }),
  new TextRun({ text: value }),
], { after: 150 });
const logoPath = __dirname + '/iut_logo.png';

const cover = [
  CP(INFO.university.toUpperCase(), { bold: true, size: 32, before: 220, after: 180 }),
  ...(fs.existsSync(logoPath) ? [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [
    new ImageRun({ type: 'png', data: fs.readFileSync(logoPath),
      transformation: { width: 88, height: 144 },
      altText: { title: 'IUT logo', description: 'Islamic University of Technology logo', name: 'IUT logo' } })] })] : []),
  CP(INFO.department.toUpperCase(), { bold: true, size: 28, after: 70 }),
  CP(INFO.degree, { bold: true, size: 26, after: 140 }),
  rule(220, 80),
  CP('INDUSTRIAL TRAINING REPORT', { bold: true, size: 30, before: 80, after: 140 }),
  CP('[FILL: Report title]', { bold: true, size: 42, after: 90 }),
  CP('[FILL: Report subtitle or internship area]', { italics: true, size: 26, after: 120 }),
  rule(140, 100),
  CP('SUBMITTED BY', { bold: true, before: 80, after: 70 }),
  CP(INFO.studentName.toUpperCase(), { bold: true, size: 24, after: 40 }),
  CP(`Student ID: ${INFO.studentId}`, { after: 220 }),
  lab('Organization', INFO.organization),
  lab('Training Program', programLabel),
  lab('Training Period', `${INFO.startDate} to ${INFO.endDate}`),
  lab('Training Mode', INFO.mode),
  lab('Venue', INFO.venue),
  lab('Industry Supervisor', INFO.supervisor),
  lab('Date of Submission', INFO.submissionDate),
];

const offer = [
  H1('2. Offer Letter', false),
  P(`The industrial training placement was arranged through the university, and no separate offer letter was issued by ${INFO.organization}. The particulars of the placement are summarized in Table 2.1.`),
  TCAP('Table 2.1: Placement particulars'),
  tbl([2800, 6560], ['Item', 'Detail'], [
    ['Student name', INFO.studentName],
    ['Student ID', INFO.studentId],
    ['University and program', `${INFO.university}; ${INFO.degree}`],
    ['Organization', INFO.organization],
    ['Training program', programLabel],
    ['Training period', `${INFO.startDate} to ${INFO.endDate}`],
    ['Session schedule', '[FILL: Working days and hours]'],
    ['Training mode', INFO.mode],
    ['Venue', INFO.venue],
  ], { boldFirst: true }),
  SP(160),
  P('[FILL: If the university issued a placement or allocation letter, insert a scan of it on this page.]'),
];

const cert = [
  H1('3. Certificate / Confirmation'),
  P(`The internship was completed under the **${INFO.program}** at ${INFO.organization} from ${INFO.startDate} to ${INFO.endDate}. Replace this paragraph with the applicable confirmation details for the internship.`),
  SP(120),
  P('[FILL: State whether a completion certificate or confirmation letter is available. Attach it here only if permission has been granted.]'),
];

const ack = [
  H1('4. Acknowledgement'),
  P('I would like to express my sincere gratitude to the Almighty for giving me the opportunity, health and patience to complete this industrial training successfully.'),
  P(`I am deeply thankful to **${INFO.organization}** for offering me a place in the *${INFO.program}*. I am especially grateful to my industry supervisor and the program facilitators, whose live demonstrations, honest case studies and direct feedback shaped the way I now think about building software products with artificial intelligence.`),
  P(`I also thank the faculty members of ${INFO.department} at ${INFO.university}, whose teaching gave me the foundation to benefit fully from this training. I am grateful to the Industrial Training course coordinator and faculty members who supervised the arrangement and documentation of the attachment.`),
  P('My thanks go to my teammate on the project, and to my university friends and fellow trainees who completed my validation survey, tested my minimum viable product and shared candid feedback in group discussions. Finally, I am grateful to my family for their constant support and encouragement throughout the training period.'),
  SP(240),
  P(INFO.studentName, { align: AlignmentType.LEFT, after: 0, run: { bold: true } }),
  P(INFO.submissionDate, { align: AlignmentType.LEFT }),
];

const toc = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 320, after: 200 },
    border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: '000000', space: 2 } },
    children: [new TextRun({ text: 'Table of Contents', bold: true, size: 26 })] }),
  new TableOfContents('Table of Contents', {
    hyperlink: true,
    headingStyleRange: '1-3',
    beginDirty: true,
  }),
];

const templateChapters = [
  H1('1. Organization Profile', true, true),
  H2('1.1 Organization Overview'),
  P('[FILL: Describe the organization, location, nature of business, products or services, and the department in which the internship was completed.]'),
  H2('1.2 Internship Environment'),
  P('[FILL: Describe the workplace, supervision arrangement, working hours, team structure and resources available during the internship.]'),
  H2('1.3 Relevant Activities'),
  P('[FILL: Explain how the organization and department were relevant to your academic program and internship objectives.]'),

  H1('2. Internship Activities and Responsibilities', true, true),
  H2('2.1 Overview of Activities'),
  P('[FILL: Summarize the main activities performed during the internship. Distinguish clearly between work personally performed, work completed with a team and activities only observed.]'),
  H2('2.2 Technical Activities'),
  P('[FILL: Describe the technical tasks completed, such as development, testing, analysis, design, documentation, networking, data handling, support or research.]'),
  H2('2.3 Professional and Operational Activities'),
  P('[FILL: Describe meetings, reporting, communication, teamwork, time management, organizational procedures and other workplace responsibilities.]'),
  H2('2.4 Tools and Methods'),
  P('[FILL: List the tools, technologies, methods, standards or procedures actually used and explain how they supported the work.]'),
  H2('2.5 Responsibilities Summary'),
  tbl([1800, 3400, 3460], ['Area', 'Work performed', 'Evidence or outcome'], [
    ['[FILL: Area]', '[FILL: Your responsibility]', '[FILL: Deliverable, record or result]'],
    ['[FILL: Area]', '[FILL: Your responsibility]', '[FILL: Deliverable, record or result]'],
    ['[FILL: Area]', '[FILL: Your responsibility]', '[FILL: Deliverable, record or result]'],
  ], { boldFirst: true }),

  H1(`3. Technical Work / Project: ${INFO.projectName}`, true, true),
  H2('3.1 Problem or Objective'),
  P('[FILL: State the problem, business need or technical objective addressed by the work.]'),
  H2('3.2 Requirements and Scope'),
  P('[FILL: Describe functional requirements, constraints, assumptions, exclusions and expected outcomes.]'),
  H2('3.3 Method and Implementation'),
  P('[FILL: Explain the approach, workflow, architecture, tools, technologies and your specific contribution.]'),
  H2('3.4 Results and Evaluation'),
  P('[FILL: Describe the result, tests performed, feedback received, measurable outcomes and remaining limitations.]'),
  H2('3.5 Problems and Solutions'),
  P('[FILL: Explain important problems encountered using the structure: challenge, cause, response and lesson.]'),
  H2('3.6 Security, Safety, Privacy and Quality'),
  P('[FILL: Discuss the relevant security, privacy, safety, ethical, quality or confidentiality considerations for this work.]'),

  H1('4. Knowledge and Skills Acquired', true, true),
  H2('4.1 Technical Knowledge'),
  P('[FILL: Explain the technical knowledge gained and connect it to relevant academic courses.]'),
  H2('4.2 Tools and Development Practices'),
  P('[FILL: Describe new tools, software-engineering practices, testing methods, documentation practices or standards learned.]'),
  H2('4.3 Communication and Teamwork'),
  P('[FILL: Explain how the internship improved communication, collaboration, presentation, feedback and reporting skills.]'),
  H2('4.4 Professional Development'),
  P('[FILL: Explain how the internship changed your understanding of workplace responsibilities, quality, deadlines and professional ethics.]'),

  H1('5. Challenges and Problem Solving', true, true),
  H2('5.1 Challenge One'),
  P('[FILL: Describe the challenge, its cause, the response taken and the lesson learned.]'),
  H2('5.2 Challenge Two'),
  P('[FILL: Describe another meaningful challenge and how it was addressed.]'),
  H2('5.3 Scope, Time or Resource Constraints'),
  P('[FILL: Explain how limited time, resources, information or changing requirements were managed.]'),
  H2('5.4 Lessons from Problem Solving'),
  P('[FILL: Summarize the problem-solving methods that you will carry into future professional work.]'),

  H1('6. Professional Experience and Reflection', true, true),
  H2('6.1 Changes in Professional Practice'),
  P('[FILL: Reflect on how the internship changed your approach to planning, implementation, verification, documentation or decision-making.]'),
  H2('6.2 Team Collaboration'),
  P('[FILL: Reflect on supervision, teamwork, communication, feedback and responsibility for deliverables.]'),
  H2('6.3 Academic and Industrial Learning'),
  P('[FILL: Connect academic knowledge with workplace practice and identify areas for further learning.]'),
  H2('6.4 Work Habits Developed'),
  P('[FILL: List practical habits developed during the internship, such as planning, version control, testing, record keeping or meeting deadlines.]'),

  H1('7. Conclusion', true, true),
  P('[FILL: State what you did during the internship, what you learned, how your professional practice changed and the lasting value of the experience. Do not introduce new technical information here.]'),

  H1('8. References', true, true),
  P('[FILL: List only the books, websites, documentation, papers, software documentation and organizational materials actually consulted. Use one consistent citation style and include URLs and access dates where required.]'),

  H1('9. Appendices', true, true),
  H2('Appendix A: Internship Information'),
  P(`**Organization:** ${INFO.organization} | **Program:** ${programLabel} | **Period:** ${INFO.startDate} – ${INFO.endDate} | **Mode:** ${INFO.mode} | **Venue:** ${INFO.venue}.`),
  H2('Appendix B: Activity Log'),
  P('[FILL: Add a dated activity log, schedule or weekly summary if required by your department.]'),
  H2('Appendix C: Supporting Evidence'),
  P('[FILL: Add approved screenshots, diagrams, tables, forms, test records or other supporting material. Do not include confidential information without permission.]'),
  H2('Appendix D: Final Checklist'),
  P('[FILL: Confirm that personal contribution, technical work, learning outcomes, challenges, references, figures, tables and confidentiality requirements have all been reviewed.]'),
];
const templateAcknowledgement = [
  H1('4. Acknowledgement'),
  P(`I would like to express my sincere gratitude to the Almighty for giving me the opportunity, health and patience to complete this internship successfully.`),
  P(`I am deeply thankful to **${INFO.organization}** for providing the internship opportunity and to my industry supervisor, colleagues and mentors for their guidance, feedback and support throughout the placement.`),
  P(`I also thank the faculty members of ${INFO.department} at ${INFO.university}, whose teaching provided the foundation to benefit from this practical experience. I am grateful to my family, teammates and everyone who supported the completion of the internship.`),
  SP(240),
  P(INFO.studentName, { align: AlignmentType.LEFT, after: 0, run: { bold: true } }),
  P(INFO.submissionDate, { align: AlignmentType.LEFT }),
];
const templateExecutiveSummary = [
  H1('5. Executive Summary'),
  P(`This report documents the internship completed by **${INFO.studentName}** at **${INFO.organization}** under the *${INFO.program}*. The placement took place from ${INFO.startDate} to ${INFO.endDate} in ${INFO.mode.toLowerCase()} mode at ${INFO.venue}.`),
  P('[FILL: Summarize the purpose of the internship, the department or team involved, the main responsibilities performed and the principal work completed.]'),
  P('[FILL: Summarize the main project or technical work, including its objective, approach, tools and outcome. Keep claims limited to work actually completed.]'),
  P('[FILL: Summarize the most important technical, professional and personal learning outcomes, together with any significant limitations or recommendations.]'),
];

const outputPath = path.join(__dirname, 'industrial-training-report.docx');
const coverPageBreak = new D.Paragraph({ children: [new D.PageBreak()] });
const page = {
  size: { width: 11906, height: 16838 },
  margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 },
};
const document = new D.Document({
  styles,
  numbering,
  features: { updateFields: true },
  sections: [{
    properties: {
      type: D.SectionType.NEXT_PAGE,
      page: { ...page, pageNumbers: { start: 1, formatType: D.NumberFormat.LOWER_ROMAN } },
    },
    footers: { default: footer() },
    children: [...cover, coverPageBreak, ...offer, ...cert, ...templateAcknowledgement, ...templateExecutiveSummary, ...toc],
  }, {
    properties: {
      page: { ...page, pageNumbers: { start: 1, formatType: D.NumberFormat.DECIMAL } },
    },
    footers: { default: footer() },
    children: templateChapters,
  }],
});

D.Packer.toBuffer(document).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log(`Generated ${outputPath} (${buffer.length} bytes)`);
}).catch(error => {
  console.error('Failed to generate the DOCX file:', error);
  process.exitCode = 1;
});
