const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, LevelFormat,
} = require('docx');
const fs = require('fs');

// ── EXACT COLORS from pixel analysis of Virendra's PDF ──────────────────────
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

// Section header: bold blue ALL-CAPS text + blue border-bottom rule (EXACTLY like Virendra)
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

// Thin BLACK rule line between job entries (like Virendra's inter-job separators)
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
          text: "KRUSHNA THUBE",
          bold: true, size: 52, font: "Calibri", color: NAME_COLOR,
        })]
      }),

      // ── SUBTITLE (blue, centered, pipe-separated) ──────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 20 },
        children: [new TextRun({
          text: "Senior Full Stack Developer  |  Node.js  |  Angular 16+  |  React  |  AWS Serverless  |  Microservices  |  GenAI",
          size: 21, font: "Calibri", color: SUBTITLE_BLU,
        })]
      }),

      // ── CONTACT (gray, centered) ───────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [
          new TextRun({ text: "Pune, Maharashtra  |  +91 7387476746  |  krushnathube19@gmail.com  |  ", size: 20, font: "Calibri", color: CONTACT_GRAY }),
          new TextRun({ text: "linkedin.com/in/krushna-thube/", size: 20, font: "Calibri", color: BLUE }),
        ]
      }),

      // ── PROFESSIONAL SUMMARY ──────────────────────────────────────────────
      sectionHeader("Professional Summary"),
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [
          new TextRun({ text: "Senior Full Stack Developer with ", size: 20, font: "Calibri", color: BODY_TEXT }),
          new TextRun({ text: "11+ years", bold: true, size: 20, font: "Calibri", color: BODY_TEXT }),
          new TextRun({ text: " of experience building scalable SaaS platforms, cloud-native microservices, and high-performance enterprise applications across healthcare, fintech, IoT, ERP, and enterprise collaboration domains. Strong expertise in ", size: 20, font: "Calibri", color: BODY_TEXT }),
          new TextRun({ text: "Node.js, TypeScript, Angular 16+, React, AWS Serverless (Lambda, DynamoDB, API Gateway, S3, CloudWatch, KMS)", bold: true, size: 20, font: "Calibri", color: BODY_TEXT }),
          new TextRun({ text: ", Kafka-based event-driven systems, Docker, and Kubernetes. Proven ability to deliver production-grade REST APIs, real-time systems, and containerized full-stack applications. Currently advancing expertise in ", size: 20, font: "Calibri", color: BODY_TEXT }),
          new TextRun({ text: "Generative AI, LLMs, RAG pipelines, and AI-assisted development", bold: true, size: 20, font: "Calibri", color: BODY_TEXT }),
          new TextRun({ text: " workflows. AWS Certified Solutions Architect — Associate. Consistently recognized for performance optimization, team leadership, and delivering enterprise-grade solutions in Agile environments.", size: 20, font: "Calibri", color: BODY_TEXT }),
        ]
      }),

      // ── TECHNICAL SKILLS ──────────────────────────────────────────────────
      sectionHeader("Technical Skills"),
      new Paragraph({ spacing: { before: 80, after: 40 }, children: [] }),
      skillsTable([
        ["Languages",        "JavaScript (ES6+), TypeScript, SQL, HTML5, CSS3"],
        ["Frontend",         "Angular 16+, TypeScript, React.js, AG Grid, AG Chart, jQuery, Responsive UI"],
        ["Backend",          "Node.js, Express.js, NestJS, REST APIs, GraphQL, WebSockets, Middleware Development, Backend Architecture"],
        ["Cloud / DevOps",   "AWS Lambda, API Gateway, DynamoDB, S3, CloudWatch, KMS, IAM, SQS, SNS, EventBridge, Step Functions, ECS, EC2, Cognito, Docker, Kubernetes, Terraform, Jenkins, GitHub Actions, CI/CD"],
        ["Databases",        "MongoDB, PostgreSQL, MySQL, DynamoDB, Redis, Mongoose, TypeORM"],
        ["Architecture",     "Microservices, Event-Driven Architecture, Kafka, Distributed Systems, Serverless, Monolith-to-Microservices, SaaS Platform Design"],
        ["Testing",          "Jest, Mocha, Chai, Sinon, Supertest, Cypress, Unit Testing, Integration Testing, TDD, ESLint, Prettier"],
        ["Security & Auth",  "JWT, OAuth 2.0, API Security, Rate Limiting, AWS KMS, AWS Cognito"],
        ["AI / GenAI",       "LLMs, RAG Pipelines, Prompt Engineering, Claude API, Amazon Q, GitHub Copilot, AI-Assisted Development"],
        ["Practices",        "System Design, SOLID Principles, OOP, Performance Optimization, Agile / Scrum, Code Reviews, Mentoring"],
      ]),
      new Paragraph({ spacing: { before: 60, after: 0 }, children: [] }),

      // ── PROFESSIONAL EXPERIENCE ───────────────────────────────────────────
      sectionHeader("Professional Experience"),

      // --- Guidesly ---
      companyLine("Guidesly India Pvt. Ltd.", "Bangalore, India", "Apr 2025 – Present"),
      roleLine("Sr. Backend Developer  |  SaaS Platform Engineering"),
      bullet([t("Designed and developed scalable backend services and RESTful APIs using Node.js, Express.js, and PostgreSQL for a SaaS-based fishing guide platform serving web and mobile clients.")]),
      bullet([t("Built cloud-native serverless solutions using "), b("AWS Lambda, API Gateway, DynamoDB, S3, and CloudWatch"), t(" for scalable event-driven backend workflows.")]),
      bullet([t("Implemented third-party insurance API integrations and developed Classes & Certifications modules for the Guidesly SaaS ecosystem.")]),
      bullet([t("Developed microservice-oriented backend components with focus on scalability, maintainability, and production reliability.")]),
      bullet([t("Containerized backend services using Docker and supported CI/CD deployment automation workflows.")]),
      bullet([t("Improved application quality through automated unit and integration testing using Jest, Mocha, and Chai.")]),
      bullet([t("Collaborated in Agile delivery teams using JIRA and Confluence for sprint planning and release execution.")]),

      // --- Hexaware ---
      thinRule(),
      companyLine("Hexaware Technologies", "Pune, India", "Aug 2021 – Mar 2025"),
      roleLine("Sr. Software Engineer — MEAN Stack Developer  |  Healthcare Enterprise Platform"),
      bullet([t("Architected and delivered scalable "), b("Node.js backend services with TypeScript"), t(" and enterprise REST APIs supporting high-availability healthcare web and mobile platforms.")]),
      bullet([t("Designed and implemented "), b("AWS Serverless solutions (Lambda, S3, CloudWatch, KMS)"), t(" for secure healthcare data processing and event-driven workflows.")]),
      bullet([t("Developed "), b("Angular 16+ frontend modules"), t(" with data-rich dashboards using "), b("AG Grid and AG Chart"), t(" for real-time clinical and operational data visualization.")]),
      bullet([t("Built high-throughput Kafka-based event-driven pipelines for distributed healthcare data processing across web and mobile platforms.")]),
      bullet([t("Authored comprehensive unit tests using "), b("Jest"), t(" and end-to-end test suites using "), b("Cypress"), t(", maintaining high coverage across critical modules.")]),
      bullet([t("Integrated AI-assisted development tools ("), b("Amazon Q, GitHub Copilot"), t(") to accelerate feature delivery, code review, and debugging efficiency in Agile CI/CD environments.")]),
      bullet([t("Contributed to containerized application deployment using Docker and participated in system design, API optimization, and backend architecture planning.")]),

      // --- Bridgetek ---
      thinRule(),
      companyLine("Bridgetek Pte. Ltd.", "Singapore", "Jan 2020 – Aug 2021"),
      roleLine("Sr. Software Engineer  |  Enterprise Room Management Platform"),
      bullet([t("Designed and delivered scalable microservices-based full-stack applications using Node.js, Angular, and PostgreSQL.")]),
      bullet([t("Worked on "), b("Kubernetes"), t("-based container orchestration and deployment pipelines; leveraged "), b("Terraform"), t(" for infrastructure-as-code automation.")]),
      bullet([t("Improved backend modularity, deployment scalability, and team velocity through microservices adoption and Docker containerization.")]),

      // --- Mobiquity ---
      thinRule(),
      companyLine("Mobiquity India", "India", "Mar 2018 – Dec 2019"),
      roleLine("Sr. Software Developer  |  AWS Serverless Platform"),
      bullet([t("Developed serverless REST APIs for web and mobile applications using "), b("Node.js and TypeScript"), t(" on AWS.")]),
      bullet([b("Built AWS cloud-native solutions using API Gateway, Lambda, DynamoDB, and S3"), t(" — key foundation of extensive AWS Serverless experience.")]),
      bullet([t("Implemented scalable backend services optimized for low operational overhead and high availability.")]),
      bullet([t("Contributed to backend architecture and API lifecycle management for enterprise applications at scale.")]),

      // --- Nihilent ---
      thinRule(),
      companyLine("Nihilent Ltd.", "India", "Feb 2017 – Mar 2018"),
      roleLine("System Analyst  |  Banking Enterprise Applications"),
      bullet([t("Analyzed business requirements and delivered enterprise frontend features using Angular and JavaScript for Nedbank, a major financial services client.")]),
      bullet([t("Worked on frontend development with Angular alongside .NET backend integration for enterprise banking workflows.")]),

      // --- TAS India ---
      thinRule(),
      companyLine("TAS India Pvt. Ltd.", "India", "Mar 2015 – Feb 2017"),
      roleLine("Software Developer  |  IIoT & Real-Time Systems"),
      bullet([t("Developed full-stack IIoT and telemetry solutions for solar energy monitoring using Node.js, Angular, and MongoDB.")]),
      bullet([t("Implemented real-time device communication using "), b("MQTT"), t(" and "), b("Modbus"), t(" protocols for industrial IoT deployments.")]),
      bullet([t("Built cloud-connected IoT systems with real-time data ingestion, telemetry processing, and analytics dashboards.")]),

      // --- HashTrix ---
      thinRule(),
      companyLine("HashTrix Technologies Pvt. Ltd.", "Pune, India", "Jun 2014 – Apr 2015"),
      roleLine("Software Developer  |  ERP Development"),
      bullet([t("Developed ERP application modules using JavaScript and SQL; contributed to backend feature implementation, business logic, and database operations.")]),

      // ── KEY PROJECTS ──────────────────────────────────────────────────────
      sectionHeader("Key Projects"),
      new Paragraph({ spacing: { before: 80, after: 0 }, children: [] }),
      bullet([b("Healthcare Enterprise Platform (Hexaware / Otsuka)"), t(" — Full-stack platform with Node.js + TypeScript backend, Angular 16+ with AG Grid/AG Chart dashboards, AWS Lambda/S3/CloudWatch/KMS, Kafka event-driven pipelines, Jest + Cypress testing. Served web and mobile platforms at enterprise scale.")]),
      bullet([b("Guidesly SaaS Backend (Guidesly)"), t(" — Scalable REST APIs using Node.js, Express.js, PostgreSQL, AWS Lambda, API Gateway, DynamoDB. Third-party insurance API integrations, microservice components, Docker-based CI/CD deployment.")]),
      bullet([b("AWS Serverless Platform (Mobiquity / AWS re:Invent)"), t(" — Serverless REST APIs using Node.js + TypeScript, API Gateway, Lambda, DynamoDB, S3. Optimized for high availability and low operational overhead.")]),
      bullet([b("PanL Room Manager (Bridgetek)"), t(" — Microservices-based enterprise full-stack platform with Node.js, Docker, Kubernetes, Terraform IaC, Angular, and PostgreSQL.")]),
      bullet([b("TASM2M Solar IoT (TAS India)"), t(" — IIoT telemetry platform using MQTT, Modbus, Node.js, Angular, and MongoDB with real-time monitoring dashboards and cloud-connected device management.")]),

      // ── AWARDS & RECOGNITION ──────────────────────────────────────────────
      sectionHeader("Awards & Recognition"),
      new Paragraph({ spacing: { before: 80, after: 0 }, children: [] }),
      bullet([b("AWS Certified Solutions Architect – Associate"), t(" — Certified hands-on expertise across Lambda, DynamoDB, API Gateway, S3, CloudWatch, KMS, IAM, and core AWS infrastructure.")]),
      bullet([b("Delivery Excellence — Hexaware Technologies (2023)"), t(" — Recognized for leading healthcare platform backend modernization with Angular 16+ frontend and full AWS Serverless integration.")]),

      // ── CERTIFICATIONS & LEARNING ─────────────────────────────────────────
      sectionHeader("Certifications & Learning"),
      new Paragraph({ spacing: { before: 80, after: 0 }, children: [] }),
      bullet([b("AWS Certified Solutions Architect – Associate")]),
      bullet([b("Generative AI & LLMs (In Progress, 2026)"), t(" — Focus: RAG Pipelines, Prompt Engineering, Multi-Agent Systems, Claude API, AI Workflow Automation.")]),
      bullet([t("AI-Assisted Development — Active use of "), b("Amazon Q, GitHub Copilot, and Claude API"), t(" for productivity acceleration, automated testing, and debugging.")]),

      // ── EDUCATION ─────────────────────────────────────────────────────────
      sectionHeader("Education"),
      new Paragraph({ spacing: { before: 80, after: 0 }, children: [] }),
      ...eduEntry(
        "Master of Computer Science (MCS)",
        "Savitribai Phule Pune University, Pune",
        "2012 – 2014"
      ),
      ...eduEntry(
        "Bachelor of Computer Science (BCS)",
        "Pune University, Pune",
        "2009 – 2012"
      ),

    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/mnt/user-data/outputs/KrushnaThube_Updated.docx", buf);
  console.log("Done");
});