const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, LevelFormat,
} = require('docx');
const fs = require('fs');
const { convert } = require('docx-to-pdf');

// Load resume data from JSON
const resumeData = JSON.parse(fs.readFileSync('./resume-data.json', 'utf8'));

const NAME_COLOR   = "1F4464";  // #23496D name - dark navy  
const BLUE         = "3D71A3";  // section headers, blue rule line, skills left col text
const SUBTITLE_BLU = "4A6E8E";  // subtitle pipe-separated line (slightly muted blue)
const CONTACT_GRAY = "555555";  // contact line gray
const BODY_TEXT    = "333333";  // body paragraphs
const COMPANY_BLK  = "000000";  // company name bold black
const ROLE_BLU     = "3A5163";  // role italic line (darker blue-gray)
const EDU_BLU      = "3D71A3";  // education degree bold blue (same as section headers)
const TBL_BORDER   = "C8C8C8";  // skills table border light gray
const PAGE_W  = 12240;
const PAGE_H  = 15840;
const MARGIN  = 1080;           // 0.75 inch
const CW      = PAGE_W - MARGIN * 2; // 10080

// ── HELPERS ──────────────────────────────────────────────────────────────────

// Section header: bold blue ALL-CAPS text + blue border-bottom rule
function sectionHeader(text) {
  return new Paragraph({
    spacing: { before: 300, after: 100 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 }
    },
    children: [new TextRun({
      text: text.toUpperCase(),
      bold: true, size: 22, font: "Calibri", color: BLUE,
    })]
  });
}

// Thin BLACK rule line between job entries
function thinRule() {
  return new Paragraph({
    spacing: { before: 0, after: 0 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000", space: 1 }
    },
    children: []
  });
}

function bullet(parts) {
  const runs = parts.map(p => new TextRun({
    text: p.text, bold: p.bold || false,
    size: 20, font: "Calibri", color: BODY_TEXT,
  }));
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: runs
  });
}

const b = text => ({ text, bold: true });
const t = text => ({ text });

// Company line + date (2-col table)
function companyLine(company, location, dates) {
  const LW = Math.round(CW * 0.72), RW = CW - LW;
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [LW, RW],
    borders: { top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},insideH:{style:BorderStyle.NONE},insideV:{style:BorderStyle.NONE} },
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: LW, type: WidthType.DXA },
        borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},
        children: [new Paragraph({
          spacing: { before: 140, after: 20 },
          children: [
            new TextRun({ text: company, bold: true, size: 21, font: "Calibri", color: COMPANY_BLK }),
            new TextRun({ text: "  |  " + location, size: 20, font: "Calibri", color: CONTACT_GRAY }),
          ]
        })]
      }),
      new TableCell({
        width: { size: RW, type: WidthType.DXA },
        borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 140, after: 20 },
          children: [new TextRun({ text: dates, italics: true, size: 20, font: "Calibri", color: CONTACT_GRAY })]
        })]
      }),
    ]})]
  });
}

function roleLine(text) {
  return new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text, italics: true, bold: true, size: 20, font: "Calibri", color: ROLE_BLU })]
  });
}

// Skills table — NO cell shading, light gray border, blue bold left col text (NO background tint)
function skillsTable(rows) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: TBL_BORDER };
  const borders = { top: border, bottom: border, left: border, right: border };
  const LW = Math.round(CW * 0.20), RW = CW - LW;
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [LW, RW],
    borders: { top: border, bottom: border, left: border, right: border, insideH: border, insideV: border },
    rows: rows.map(([label, value]) => new TableRow({ children: [
      new TableCell({
        width: { size: LW, type: WidthType.DXA }, borders,
        margins: { top: 80, bottom: 80, left: 140, right: 80 },
        children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 20, font: "Calibri", color: BLUE })] })]
      }),
      new TableCell({
        width: { size: RW, type: WidthType.DXA }, borders,
        margins: { top: 80, bottom: 80, left: 140, right: 80 },
        children: [new Paragraph({ children: [new TextRun({ text: value, size: 20, font: "Calibri", color: BODY_TEXT })] })]
      }),
    ]}))
  });
}

function eduEntry(degree, institution, dates) {
  const LW = Math.round(CW * 0.72), RW = CW - LW;
  return [
    new Table({
      width: { size: CW, type: WidthType.DXA }, columnWidths: [LW, RW],
      borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},insideH:{style:BorderStyle.NONE},insideV:{style:BorderStyle.NONE}},
      rows: [new TableRow({ children: [
        new TableCell({
          width:{size:LW,type:WidthType.DXA},
          borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},
          children: [new Paragraph({ spacing:{before:80,after:10}, children:[new TextRun({text:degree,bold:true,size:21,font:"Calibri",color:EDU_BLU})] })]
        }),
        new TableCell({
          width:{size:RW,type:WidthType.DXA},
          borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},
          children: [new Paragraph({ alignment:AlignmentType.RIGHT, spacing:{before:80,after:10}, children:[new TextRun({text:dates,italics:true,size:20,font:"Calibri",color:CONTACT_GRAY})] })]
        }),
      ]})]
    }),
    new Paragraph({ spacing:{before:0,after:80}, children:[new TextRun({text:institution,size:20,font:"Calibri",color:CONTACT_GRAY})] }),
  ];
}

// ── DOCUMENT ─────────────────────────────────────────────────────────────────

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "\u2022",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 480, hanging: 240 } } }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Calibri", size: 20, color: BODY_TEXT } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
      }
    },
    children: [

      // ── NAME ──────────────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 30 },
        children: [new TextRun({
          text: resumeData.personal.name,
          bold: true, size: 52, font: "Calibri", color: NAME_COLOR,
        })]
      }),

      // ── SUBTITLE (blue, centered, pipe-separated) ──────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 20 },
        children: [new TextRun({
          text: resumeData.personal.subtitle,
          size: 21, font: "Calibri", color: SUBTITLE_BLU,
        })]
      }),

      // ── CONTACT (gray, centered) ───────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [
          new TextRun({ text: `${resumeData.personal.location}  |  ${resumeData.personal.phone}  |  ${resumeData.personal.email}  |  `, size: 20, font: "Calibri", color: CONTACT_GRAY }),
          new TextRun({ text: resumeData.personal.linkedin, size: 20, font: "Calibri", color: BLUE }),
        ]
      }),

      // ── PROFESSIONAL SUMMARY ──────────────────────────────────────────────
      sectionHeader("Professional Summary"),
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [new TextRun({ text: resumeData.summary.content, size: 20, font: "Calibri", color: BODY_TEXT })]
      }),

      // ── TECHNICAL SKILLS ──────────────────────────────────────────────────
      sectionHeader("Technical Skills"),
      new Paragraph({ spacing: { before: 80, after: 40 }, children: [] }),
      skillsTable(resumeData.skills.map(s => [s.label, s.value])),
      new Paragraph({ spacing: { before: 60, after: 0 }, children: [] }),

      // ── PROFESSIONAL EXPERIENCE ───────────────────────────────────────────
      sectionHeader("Professional Experience"),

      ...resumeData.experience.flatMap((job, idx) => [
        ...(idx > 0 ? [thinRule()] : []),
        companyLine(job.company, job.location, job.dates),
        roleLine(job.role),
        ...job.bullets.map(bulletText => bullet([t(bulletText)])),
      ]),

      // ── KEY PROJECTS ──────────────────────────────────────────────────────
      sectionHeader("Key Projects"),
      new Paragraph({ spacing: { before: 80, after: 0 }, children: [] }),
      ...resumeData.projects.map(project => 
        bullet([b(project.name), t(" — " + project.description)])
      ),

      // ── AWARDS & RECOGNITION ──────────────────────────────────────────────
      sectionHeader("Awards & Recognition"),
      new Paragraph({ spacing: { before: 80, after: 0 }, children: [] }),
      ...resumeData.awards.map(award => 
        bullet([b(award.title), t(award.description ? " — " + award.description : "")])
      ),

      // ── CERTIFICATIONS & LEARNING ─────────────────────────────────────────
      sectionHeader("Certifications & Learning"),
      new Paragraph({ spacing: { before: 80, after: 0 }, children: [] }),
      ...resumeData.certifications.map(cert => 
        bullet([b(cert.title), t(cert.description ? " — " + cert.description : "")])
      ),

      // ── EDUCATION ─────────────────────────────────────────────────────────
      sectionHeader("Education"),
      new Paragraph({ spacing: { before: 80, after: 0 }, children: [] }),
      ...resumeData.education.flatMap(edu => eduEntry(edu.degree, edu.institution, edu.dates)),

    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  const docxPath = "/mnt/user-data/outputs/Krushna_Thube_Resume.docx";
  const pdfPath = "/mnt/user-data/outputs/Krushna_Thube_Resume.pdf";
  
  fs.writeFileSync(docxPath, buf);
  console.log("Generated DOCX: Krushna_Thube_Resume.docx");
  
  // Convert DOCX to PDF
  convert({
    input: docxPath,
    output: pdfPath
  }).then(() => {
    console.log("Generated PDF: Krushna_Thube_Resume.pdf");
    console.log("Done - Both DOCX and PDF generated successfully");
  }).catch(err => {
    console.error("PDF conversion failed:", err);
  });
});