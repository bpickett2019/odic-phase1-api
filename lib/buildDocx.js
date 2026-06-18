/**
 * buildDocx.js
 * Assembles the ODIC Phase I ESA Word document from form data + Claude-generated sections.
 * Matches the structure and formatting of the ODIC template (ASTM E1527-21).
 */

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  HeadingLevel,
  AlignmentType,
  WidthType,
  BorderStyle,
  Header,
  Footer,
  PageNumber,
  NumberFormat,
  ShadingType,
  UnderlineType,
  PageBreak,
  Tab,
  TabStopPosition,
  TabStopType,
  convertInchesToTwip,
  LevelFormat,
} from 'docx';

// ─── Color / Style Constants ────────────────────────────────────────────────
const ODIC_RED = 'C00000';
const BLACK = '000000';
const WHITE = 'FFFFFF';
const LIGHT_GRAY = 'D9D9D9';
const DARK_GRAY = '404040';

const BODY_FONT = 'Times New Roman';
const HEADING_FONT = 'Times New Roman';
const BODY_SIZE = 20; // half-points (= 10pt)
const NORMAL_SIZE = 22; // 11pt
const HEADING_SIZE = 24; // 12pt
const SMALL_SIZE = 18; // 9pt

const MARGIN = convertInchesToTwip(1);
const HALF_INCH = convertInchesToTwip(0.5);

// ─── Helpers ────────────────────────────────────────────────────────────────

/** Plain body paragraph */
function bodyPara(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, font: BODY_FONT, size: NORMAL_SIZE, color: BLACK, ...opts })],
    spacing: { after: 120, line: 276 },
    alignment: AlignmentType.JUSTIFIED,
  });
}

/** Italic body paragraph (for quotes / block text) */
function italicPara(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: BODY_FONT, size: NORMAL_SIZE, italics: true, color: BLACK })],
    spacing: { after: 120, line: 276 },
    indent: { left: convertInchesToTwip(0.5), right: convertInchesToTwip(0.5) },
    alignment: AlignmentType.JUSTIFIED,
  });
}

/** Bold label + normal text on same line */
function labelPara(label, value) {
  return new Paragraph({
    children: [
      new TextRun({ text: `${label}: `, font: BODY_FONT, size: NORMAL_SIZE, bold: true }),
      new TextRun({ text: value || 'N/A', font: BODY_FONT, size: NORMAL_SIZE }),
    ],
    spacing: { after: 80 },
  });
}

/** Section heading (e.g., "2.0 PROPERTY DESCRIPTION") */
function sectionHeading(text) {
  return new Paragraph({
    children: [new TextRun({ text: text.toUpperCase(), font: HEADING_FONT, size: HEADING_SIZE, bold: true, color: BLACK })],
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120 },
    alignment: AlignmentType.CENTER,
  });
}

/** Subsection heading (e.g., "2.1 Project Information") */
function subHeading(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: HEADING_FONT, size: NORMAL_SIZE, bold: true, italics: true, color: BLACK })],
    spacing: { before: 200, after: 80 },
  });
}

/** Bullet paragraph */
function bullet(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: BODY_FONT, size: NORMAL_SIZE })],
    bullet: { level: 0 },
    spacing: { after: 80 },
  });
}

/** Bold underlined subsection label */
function boldUnderline(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: BODY_FONT, size: NORMAL_SIZE, bold: true, underline: { type: UnderlineType.SINGLE } })],
    spacing: { before: 160, after: 80 },
  });
}

/** Empty spacer paragraph */
function spacer(pts = 120) {
  return new Paragraph({ children: [new TextRun('')], spacing: { after: pts } });
}

/** Page break paragraph */
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ─── Table Builders ─────────────────────────────────────────────────────────

/** Standard two-column key-value table (ODIC style) */
function kvTable(title, rows) {
  const headerRow = new TableRow({
    children: [
      new TableCell({
        children: [new Paragraph({
          children: [new TextRun({ text: 'ITEM', font: BODY_FONT, size: NORMAL_SIZE, bold: true, color: WHITE })],
        })],
        columnSpan: 2,
        shading: { fill: DARK_GRAY, type: ShadingType.SOLID, color: DARK_GRAY },
        width: { size: 100, type: WidthType.PERCENTAGE },
      }),
    ],
  });

  const titleRow = new TableRow({
    children: [
      new TableCell({
        children: [new Paragraph({
          children: [new TextRun({ text: title, font: BODY_FONT, size: NORMAL_SIZE, bold: true })],
          alignment: AlignmentType.CENTER,
        })],
        columnSpan: 2,
        shading: { fill: LIGHT_GRAY, type: ShadingType.SOLID },
      }),
    ],
  });

  const dataRows = rows.map(([key, value]) =>
    new TableRow({
      children: [
        new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: key, font: BODY_FONT, size: NORMAL_SIZE })] })],
          width: { size: 30, type: WidthType.PERCENTAGE },
          shading: { fill: 'F2F2F2', type: ShadingType.SOLID },
        }),
        new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: value || '', font: BODY_FONT, size: NORMAL_SIZE })] })],
          width: { size: 70, type: WidthType.PERCENTAGE },
        }),
      ],
    })
  );

  return new Table({
    rows: [titleRow, headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
  });
}

/** Full-width shaded heading row for reconnaissance table */
function reconTable(rows) {
  const tableRows = rows.map(([header, value]) => {
    if (!value) {
      // This is a section header row
      return new TableRow({
        children: [
          new TableCell({
            children: [new Paragraph({
              children: [new TextRun({ text: header, font: BODY_FONT, size: NORMAL_SIZE, bold: true })],
            })],
            shading: { fill: LIGHT_GRAY, type: ShadingType.SOLID },
            columnSpan: 1,
          }),
        ],
      });
    }
    return new TableRow({
      children: [
        new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: value, font: BODY_FONT, size: NORMAL_SIZE })] })],
        }),
      ],
    });
  });

  return new Table({
    rows: tableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
  });
}

/** Summary of Findings table (Executive Summary style) */
function summaryTable(rows) {
  const tableRows = rows.map(([header, content, isHeader]) => {
    if (isHeader) {
      return new TableRow({
        children: [
          new TableCell({
            children: [new Paragraph({
              children: [new TextRun({ text: header, font: BODY_FONT, size: NORMAL_SIZE, bold: true, color: WHITE })],
              alignment: AlignmentType.CENTER,
            })],
            shading: { fill: DARK_GRAY, type: ShadingType.SOLID, color: DARK_GRAY },
          }),
        ],
      });
    }
    return new TableRow({
      children: [
        new TableCell({
          children: [
            new Paragraph({
              children: [new TextRun({ text: header, font: BODY_FONT, size: NORMAL_SIZE, bold: true })],
              shading: { fill: LIGHT_GRAY, type: ShadingType.SOLID },
            }),
            ...splitParagraphs(content).map(t => bodyPara(t)),
          ],
        }),
      ],
    });
  });

  return new Table({
    rows: tableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
  });
}

/** Split a multi-paragraph string into array of strings */
function splitParagraphs(text) {
  if (!text) return [''];
  return text.split(/\n\n+/).map(t => t.trim()).filter(Boolean);
}

/** Convert multi-paragraph text to array of bodyPara() elements */
function textToParas(text) {
  return splitParagraphs(text || '').map(t => bodyPara(t));
}

// ─── ODIC Standard Header ───────────────────────────────────────────────────

function makeHeader() {
  return new Header({
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: 'Odic ', font: 'Arial', size: 28, bold: true, color: ODIC_RED }),
          new TextRun({ text: 'Environmental', font: 'Arial', size: 28, bold: true, color: BLACK }),
        ],
        alignment: AlignmentType.RIGHT,
      }),
      new Paragraph({
        children: [new TextRun({ text: 'Environmental Consulting & Real Estate Due Diligence', font: 'Arial', size: 16, color: BLACK })],
        alignment: AlignmentType.RIGHT,
      }),
      new Paragraph({
        children: [new TextRun({ text: '407 West Imperial Suite H #303', font: 'Arial', size: 16, color: BLACK })],
        alignment: AlignmentType.RIGHT,
      }),
      new Paragraph({
        children: [new TextRun({ text: 'Brea, CA 92821', font: 'Arial', size: 16, color: BLACK })],
        alignment: AlignmentType.RIGHT,
      }),
      new Paragraph({
        children: [new TextRun({ text: 'Tel 213-380-0090', font: 'Arial', size: 16, color: BLACK })],
        alignment: AlignmentType.RIGHT,
      }),
      new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLACK } },
        children: [],
      }),
    ],
  });
}

/** Running header for report body pages */
function makeBodyHeader(projectNumber, reportTitle) {
  return new Header({
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: reportTitle, font: BODY_FONT, size: SMALL_SIZE, italics: true }),
          new TextRun({ text: '\t', font: BODY_FONT, size: SMALL_SIZE }),
          new TextRun({ text: `Project No. ${projectNumber}`, font: BODY_FONT, size: SMALL_SIZE, italics: true }),
        ],
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLACK } },
      }),
    ],
  });
}

/** Footer with page number and ODIC branding */
function makeFooter() {
  return new Footer({
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: 'O D I C   E n v i r o n m e n t a l', font: BODY_FONT, size: SMALL_SIZE, color: BLACK }),
        ],
        alignment: AlignmentType.RIGHT,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: BLACK } },
      }),
    ],
  });
}

// ─── Section Builders ────────────────────────────────────────────────────────

function buildTitlePage(d) {
  return [
    spacer(1440), // ~1 inch
    new Paragraph({
      children: [new TextRun({ text: 'PHASE I ENVIRONMENTAL SITE ASSESSMENT', font: BODY_FONT, size: 28, bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 480 },
    }),
    new Paragraph({
      children: [new TextRun({ text: 'Subject Property Address', font: BODY_FONT, size: NORMAL_SIZE, italics: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: d.propertyAddress || '', font: BODY_FONT, size: HEADING_SIZE, bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }),
    new Paragraph({
      children: [new TextRun({ text: 'Odic Project Number', font: BODY_FONT, size: NORMAL_SIZE, italics: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: d.projectNumber || '', font: BODY_FONT, size: HEADING_SIZE, bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }),
    new Paragraph({
      children: [new TextRun({ text: 'Report Date', font: BODY_FONT, size: NORMAL_SIZE, italics: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: d.reportDate || '', font: BODY_FONT, size: HEADING_SIZE, bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }),
    new Paragraph({
      children: [new TextRun({ text: 'Prepared for', font: BODY_FONT, size: NORMAL_SIZE, italics: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: d.clientName || '', font: BODY_FONT, size: HEADING_SIZE, bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: d.clientCompany || '', font: BODY_FONT, size: HEADING_SIZE, bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: d.clientAddress || '', font: BODY_FONT, size: HEADING_SIZE, bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 480 },
    }),
    new Paragraph({
      children: [new TextRun({ text: 'Odic Environmental', font: BODY_FONT, size: HEADING_SIZE, bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
    }),
    new Paragraph({
      border: { top: { style: BorderStyle.DOUBLE, size: 6, color: BLACK } },
      children: [],
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: 'Environmental Consulting and Real Estate Due Diligence', font: BODY_FONT, size: NORMAL_SIZE, italics: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
    }),
    new Paragraph({
      children: [new TextRun({ text: '407 W. Imperial Hwy Suite H #303, Brea, CA92821', font: BODY_FONT, size: NORMAL_SIZE, italics: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
    }),
    new Paragraph({
      children: [new TextRun({ text: '213.380.0090', font: BODY_FONT, size: NORMAL_SIZE, italics: true })],
      alignment: AlignmentType.CENTER,
    }),
    pageBreak(),
  ];
}

function buildTransmittalLetter(d) {
  const lenderLines = (d.lenders || []).flatMap(l => [
    bodyPara(`${l.name}`),
    bodyPara(`${l.company}`),
    bodyPara(`${l.address}`),
    spacer(80),
  ]);

  return [
    spacer(240),
    bodyPara(d.reportDate || ''),
    spacer(80),
    bodyPara(d.clientName || ''),
    bodyPara(d.clientCompany || ''),
    bodyPara(d.clientAddress || ''),
    spacer(240),
    bodyPara('Attached please find our PHASE I ENVIRONMENTAL SITE ASSESSMENT, ("the Report") for the above-mentioned Subject Property. This report has been prepared by Odic for the Client under the professional supervision of the principal and/or senior staff . Neither Odic, nor any staff member assigned to this investigation has any interest or contemplated interest, financial or otherwise, in the subject or surrounding properties, or in any entity which owns, leases, or occupies the subject or surrounding properties, and has no personal bias with respect to the parties involved.'),
    spacer(200),
    bodyPara('The assessment was conducted in a manner consistent with the level of care and skill ordinarily exercised by members of the profession, and in accordance with generally accepted practices of other consultants currently practicing in the same locality under similar conditions. No other representation, expressed or implied, and no warranty or guarantee is included or intended. The Report speaks only as of its date, in the absence of a specific written update of the Report, signed and delivered by Odic.'),
    spacer(200),
    bodyPara('There are no intended or unintended third party beneficiaries to this Report, unless specifically named. Odic is an independent contractor, not an employee of either the issuer or the borrower, and its compensation was not based on the findings or recommendations made in the Report or on the closing of any business transaction. Thank you for the opportunity to prepare this Report, and assist you with this project. Please call us if you have any questions or if we may be of further assistance.'),
    spacer(200),
    bodyPara('By signing below, Odic declares that, to the best of our professional knowledge and belief, the undersigned meet the definition of an Environmental Professional as defined in §312.10 of 40 CFR 312 and have the specific qualifications based on education, training, and experience to assess a property of the nature, history, and setting of the Subject Property. Odic has developed and performed the all appropriate inquiries in conformance with the standards and practices set forth in 40 CFR Part 312.'),
    spacer(200),
    bodyPara('Respectfully Submitted,'),
    spacer(600),
    bodyPara(d.epName || 'Michael Miller'),
    bodyPara(d.epTitle || 'Senior Consultant per §312.10 of 40 CFR 312'),
    pageBreak(),
  ];
}

function buildTableOfContents(d) {
  const tocEntries = [
    ['EXECUTIVE SUMMARY', '2'],
    ['FINDINGS AND RECOMMENDATIONS', '5'],
    ['1.0 INTRODUCTION', '10'],
    ['2.0 PROPERTY DESCRIPTION', '16'],
    ['  2.1 Project Information', '16'],
    ['  2.2 Property Improvements', '16'],
    ['  2.3 Property Occupants and Use', '17'],
    ['  2.4 Municipal Services and Utilities', '17'],
    ['  2.5 Physical Setting', '17'],
    ['3.0 PROPERTY RECONNAISSANCE', '19'],
    ['  3.1 Limiting Conditions', '19'],
    ['  3.2 Property Reconnaissance', '19'],
    ['  3.3 Detailed Description of Site Reconnaissance and Environmental Conditions', '20'],
    ['  3.4 Current Uses of Adjoining Properties', '21'],
    ['  3.5 Non-Scope (Non-ASTM) Considerations', '21'],
    ['4.0 PROPERTY AND VICINITY HISTORY', '26'],
    ['  4.1 Previous Environmental Reports', '26'],
    ['  4.2 Sanborn Map Company Fire Insurance Maps', '26'],
    ['  4.3 Historical Aerial Photographs', '26'],
    ['  4.4 Local Street Directories / Historical City Directories', '26'],
    ['  4.5 City/County Building Department, Zoning/Land Use, Property Tax Records, Profiles', '27'],
    ['  4.6 Historical Topographic Maps', '27'],
    ['  4.7 Oil & Gas Maps', '28'],
    ['  4.8 Other Historical Records', '28'],
    ['5.0 STANDARD ENVIRONMENTAL RECORDS RESEARCH', '29'],
    ['  5.1 Procedure', '29'],
    ['  5.2 Property Listing(s)', '29'],
    ['  5.3 Surrounding Sites: Federal Agency Listings', '38'],
    ['  5.4 Surrounding Sites: State Agency Listings', '42'],
    ['  5.5 Vapor Encroachment Condition', '47'],
    ['6.0 USER PROVIDED INFORMATION', '50'],
    ['  6.1 User Provided Information', '50'],
    ['  6.2 Preliminary Title Report or Land Title Records', '51'],
    ['  6.3 Interviews', '52'],
    ['7.0 REFERENCES', '53'],
  ];

  return [
    new Paragraph({
      children: [new TextRun({ text: 'TABLE OF CONTENTS', font: BODY_FONT, size: HEADING_SIZE, bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
    }),
    ...tocEntries.map(([label, pg]) =>
      new Paragraph({
        children: [
          new TextRun({ text: label, font: BODY_FONT, size: NORMAL_SIZE }),
          new TextRun({ text: `\t${pg}`, font: BODY_FONT, size: NORMAL_SIZE }),
        ],
        spacing: { after: 60 },
        tabStops: [{ type: TabStopType.RIGHT, position: convertInchesToTwip(6.5) }],
      })
    ),
    spacer(200),
    bodyPara('APPENDIX A – PROPERTY LOCATION MAP & PLOT PLAN'),
    bodyPara('APPENDIX B – PROPERTY & VICINITY PHOTOGRAPHS'),
    bodyPara('APPENDIX C – DATABASE REPORT'),
    bodyPara('APPENDIX D – HISTORICAL RECORDS RESEARCH'),
    bodyPara('APPENDIX E – PUBLIC AGENCY RECORDS / OTHER RELEVANT DOCUMENTS'),
    bodyPara('APPENDIX F – QUALIFICATIONS OF ENVIRONMENTAL PROFESSIONAL'),
    pageBreak(),
  ];
}

function buildExecutiveSummary(d, s) {
  const es = s.execSummary || {};
  return [
    sectionHeading('EXECUTIVE SUMMARY'),
    bodyPara(`ODIC Environmental (hereinafter referred to as ODIC) performed a Phase I Environmental Site Assessment (hereinafter Phase I ESA or Report) of the subject property (hereinafter referred to as the Property) in conformance with the scope and limitations of ASTM Standard Practice E1527-21. Any exceptions to or deletions from this practice are described in the individual sections of this Report. A summary of findings is provided below:`),
    spacer(160),
    summaryTable([
      ['SUMMARY OF FINDINGS', '', true],
      ['CURRENT PROPERTY ADDRESS', `The Property address is ${d.propertyAddress || ''}.`],
      ['PROPERTY DESCRIPTION', es.propertyDescription || ''],
      ['SUMMARY OF PROPERTY RECONNAISSANCE', es.propertyReconnaissance || ''],
      ['HISTORICAL USE OF THE PROPERTY', es.historicalUse || ''],
      ['HISTORICAL USE OF ADJOINING PROPERTIES / VICINITY', es.historicalAdjoining || ''],
      ['PFAS', es.pfas || 'There was no evidence from any researched source that per- and polyfluoroalkyl substances (PFAS), perfluorooctanoic acid, or perfluorooctanesulfonic acid were historically used at the Property.'],
      ['FEDERAL, STATE, AND LOCAL AGENCY RECORDS RESEARCH', es.federalStateLocalRecords || ''],
      ['POTENTIAL OFF-SITE CONCERNS', es.potentialOffsiteConcerns || ''],
      ['NON-SCOPE ITEMS', es.nonScopeItems || 'Unless the Client contracted ODIC to investigate specific Non-Scope or Non-CERCLA items, evaluation of Non-Scope or Non-CERCLA items, including those addressed in Section 3.5 of this Report, is not required nor relevant for compliance with the AAI Rule or ASTM Standard Practice E1527-21.'],
      ['INACCESSIBLE OR UN-SURVEYED PORTIONS OF THE PROPERTY', es.inaccessiblePortions || 'Full access to the entire Property was provided to ODIC, and there were no notable portions of the Property excluded from the survey and field inspection.'],
      ['SIGNIFICANT DATA GAP', es.significantDataGap || 'No significant data gaps were identified during the course of this assessment.'],
    ]),
    pageBreak(),
  ];
}

function buildFindingsAndRecommendations(d, s) {
  const fr = s.findingsAndRecommendations || {};
  return [
    sectionHeading('FINDINGS AND RECOMMENDATIONS'),
    spacer(120),
    summaryTable([
      ['SUMMARY OF FINDINGS', '', true],
      ['CONCLUSIONS AND FINDINGS', buildFindingsText(d, fr)],
      ['OTHER ENVIRONMENTAL ISSUES', fr.otherEnvironmentalIssues || 'None were identified.'],
      ['RECOMMENDATIONS AND OPINIONS', buildRecommendationsText(d, fr)],
    ]),
    spacer(200),
    boldUnderline('SBA Mitigating Factors:'),
    spacer(80),
    bodyPara('Per SBA SOP 50 10 8 Chapter 5, §E Environmental Policies and Procedures (refer to of SBA SOP 50 10 8 for full text):'),
    spacer(120),
    bodyPara('Loans may not be approved or disbursed if there is known Contamination or on-going Remediation at the Property unless the risks have been minimized to the satisfaction of SBA. SBA Lenders seeking loan approval or disbursement authority despite Contamination or on-going Remediation at the Property must submit a recommendation to SBA that includes, at a minimum, a discussion of the following:'),
    spacer(120),
    ...textToParas(fr.sbaMitigatingFactorsDiscussion || ''),
    pageBreak(),
  ];
}

function buildFindingsText(d, fr) {
  return `ODIC Environmental (ODIC) performed a Phase I Environmental Site Assessment of the Property in conformance with the scope and limitations of ASTM Standard Practice E1527-21.\n\n` +
    `ASTM Standard Practice E1527-21 defines a Recognized Environmental Condition (REC) as: (1) the presence of hazardous substances or petroleum products in, on, or at the subject property due to a release to the environment; (2) the likely presence of any hazardous substances or petroleum products in, on, or at the subject property (1) due to release to the environment; (a release) under conditions indicative of a or likely release to the environment; or (3) the presence of hazardous substances or petroleum products in, on, or at the subject property under conditions that pose a material threat of a future release to the environment.\n\n` +
    (fr.recNarrative || '') + '\n\n' +
    `ASTM Standard Practice E1527-21 defines a Historical Recognized Environmental Condition (HREC) as: a previous release of hazardous substances or petroleum products affecting the subject property that has been addressed to the satisfaction of the applicable regulatory authority or authorities and meeting unrestricted use criteria established by the applicable regulatory authority or authorities without subjecting the subject property to any controls.\n\n` +
    (fr.hrecNarrative || 'This environmental assessment has revealed no HRECs in connection with the Property.') + '\n\n' +
    `ASTM Standard Practice E1527-21 defines a Controlled Recognized Environmental Condition (CREC) as: recognized environmental condition affecting the subject property that has been addressed to the satisfaction of the applicable regulatory authority or authorities with hazardous substances or petroleum products allowed to remain in place subject to implementation of required controls.\n\n` +
    (fr.crecNarrative || 'This environmental assessment has revealed no CRECs in connection with the Property.');
}

function buildRecommendationsText(d, fr) {
  return `ODIC performed a Phase I Environmental Site Assessment of the Property in conformance with the scope and limitations of ASTM Standard Practice E1527-21.\n\n` +
    (fr.recommendation || 'ODIC recommends No Further Investigation at this time.');
}

function buildIntroduction(d) {
  return [
    sectionHeading('1.0 INTRODUCTION'),
    spacer(120),
    bodyPara(`ODIC Environmental (ODIC) performed a Phase I Environmental Site Assessment Report (hereinafter referred as "Phase I ESA" or "Report") of the Property in conformance with the scope and limitations of the ASTM International, formerly known as the American Society for Testing and Materials (ASTM), Standard Practice for Environmental Site Assessments: Phase I Environmental Site Assessment Process, ASTM Designation E1527-21.`),
    bodyPara(`This Report documents the methods and findings of the Phase I ESA performed in general conformance with the scope and limitations of ASTM Standard Practice E1527-21 and the Environmental Protection Agency Standards and Practices for All Appropriate Inquiries (40 CFR Part 312) for the Property.`),
    bodyPara(`This Report has been prepared by ODIC for the Client under the professional supervision of the principal and/or senior staff whose seal(s) and signature(s) appear hereon. Neither ODIC, nor any staff member assigned to this investigation has any interest or contemplated interest, financial or otherwise, in the subject or surrounding properties, or in any entity which owns, leases, or occupies the subject or surrounding properties or which may be responsible for environmental issues identified during the course of this investigation, and has no personal bias with respect to the parties involved.`),
    spacer(160),
    subHeading('PURPOSE AND OBJECTIVE'),
    bodyPara(`The purpose of this practice is to define good commercial and customary practice for conducting an environmental site assessment of a parcel(s) of commercial real estate with respect to the range of contaminants within the scope of the Comprehensive Environmental Response, Compensation and Liability Act (CERCLA) (42 U.S.C. §9601) and petroleum products. As such, this practice is intended to permit a User (Client, Purchaser, Lender, Owner) to satisfy one of the requirements to qualify for the innocent landowner, contiguous property owner, or bona fide prospective purchaser limitations on CERCLA liability (hereinafter, the "landowner liability protections," or "LLPs"): that is, the practice that constitutes "all appropriate inquiry" into the previous ownership and uses of the Property consistent with good commercial or customary standards and practices" as defined at 42 U.S.C. §9601(35)(B).`),
    bodyPara(`Another purpose of this Phase I ESA may be to assist the Client in its underwriting of a proposed mortgage loan on the Property, if this Report is prepared as a part of a pre-financing environmental due diligence, and to identify Recognized Environmental Conditions (RECs) in connection with the Property described in this Report.`),
    spacer(160),
    subHeading('SCOPE OF WORK'),
    bodyPara(`This Report was prepared for the exclusive use of the Client or User of this Report. The information reported was obtained through sources deemed reasonably ascertainable, as defined in ASTM Standard Practice E1527-21; a visual survey of areas readily observable, easily accessible or made accessible by the Property contact, and interviews with owners, agents, occupants, or other appropriate persons involved with the Property. Municipal information was obtained through file reviews of reasonably ascertainable standard government record sources, and interviews with the authorities having jurisdiction over the Property. Findings, conclusions, and recommendations included in the Report are based on our visual observations in the field, the municipal information reasonably obtained, information provided by the Client (or User), and/or a review of readily available and supplied documents.`),
    bodyPara(`The scope of work for this Phase I ESA is in general accordance with the requirements of ASTM Standard Practice E 1527-21. This assessment included: 1) Property and adjoining site reconnaissance; 2) interviews with key personnel; 3) a review of standard historical sources; 4) a review of standard regulatory agency records; and 5) a review of a regulatory database report provided by a third-party company such as Environmental Data Resources (EDR).`),
    spacer(160),
    subHeading('LIMITATIONS AND EXCEPTIONS'),
    bodyPara(`ODIC renders no opinion as to the Property condition at un-surveyed and/or inaccessible portions of the Property, which are described below. ODIC relies completely on the information, whether written, graphic or verbal, provided by the Property contact or as shown on any documents reviewed or received from the Property contact, owner or agent, or municipal source, and assumes that information to be true and correct. The observations in this Report are valid on the date of the Property reconnaissance. Note: Typically lenders have environmental policies where due diligence reports are valid for one year from the report date. However, such policies and standards can vary from each lender or User. For CERCLA landowner liability protection, Phase I ESA reports are valid for 180 days, per ASTM Standard Practice E1527-21.`),
    pageBreak(),
  ];
}

function buildPropertyDescription(d, s) {
  const ps = s.physicalSetting || {};
  return [
    sectionHeading('2.0 PROPERTY DESCRIPTION'),
    spacer(120),
    subHeading('2.1   PROJECT INFORMATION'),
    spacer(80),
    kvTable('Project Information', [
      ['Project Number', d.projectNumber],
      ['Property Address(es)', d.propertyAddress],
      ['Historical/Alternate Property Addresses', d.alternateAddresses || 'None were identified.'],
      ['Tax Assessor\'s Parcel Number', d.apn],
      ['Property Inspection Date', d.inspectionDate],
      ['Weather Condition', d.weatherCondition],
      ['Site Visit Conducted by', d.siteVisitBy],
      ['Report Author', d.reportAuthor],
      ['QA/QC Environmental Professional', d.qaqcEP],
      ['Property Location', d.propertyLocation],
      ['General Setting', d.generalSetting],
    ]),
    spacer(200),
    subHeading('2.2   PROPERTY IMPROVEMENTS'),
    spacer(80),
    kvTable('Property Improvements and Building / Land Description', [
      ['Property Description', d.propertyDescription],
      ['Estimated Year of Construction', d.constructionYear],
      ['Improvement Description', d.improvementDescription],
      ['Other Improvements & Features', d.otherImprovements || 'None were observed.'],
    ]),
    spacer(200),
    subHeading('2.3   PROPERTY OCCUPANTS AND USE'),
    spacer(80),
    kvTable('Property Occupants and Use', [
      ['Present Occupant(s) and Detailed Description of Business Operation(s)', d.presentOccupants],
    ]),
    spacer(200),
    subHeading('2.4   MUNICIPAL SERVICES AND UTILITIES'),
    spacer(80),
    kvTable('Municipal Services and Utilities', [
      ['Potable Water Supply', d.waterSupply],
      ['Natural Gas Utility Provider', d.gasUtility],
      ['Electrical Utility Provider', d.electricUtility],
      ['Sewage Disposal System', d.sewageSystem],
      ['Solid Waste Disposal', d.solidWaste],
      ['Any Septic System, Cesspool, Seepage Pits', d.septicSystem || 'None were identified.'],
      ['Private Water Well', d.privateWell || 'None were identified.'],
      ['Heating/Cooling System', d.hvacSystem],
    ]),
    spacer(200),
    subHeading('2.5   PHYSICAL SETTING'),
    spacer(80),
    boldUnderline('TOPOGRAPHY'),
    ...textToParas(ps.topography || ''),
    spacer(120),
    boldUnderline('GEOLOGY AND HYDROGEOLOGY'),
    ...textToParas(ps.geologyHydrogeology || ''),
    pageBreak(),
  ];
}

function buildPropertyReconnaissance(d, s) {
  const recon = d.reconItems || {};
  return [
    sectionHeading('3.0 PROPERTY RECONNAISSANCE'),
    spacer(120),
    subHeading('3.1   LIMITING CONDITIONS'),
    bodyPara('The information reported herein was obtained through sources deemed reliable, a visual site survey of areas readily observable, easily accessible or made accessible by the Property contact, and interviews with owners, agents, occupants, or other appropriate persons involved with the Property.'),
    bodyPara('No limitations imposed by physical obstructions such as adjacent buildings, bodies of water, asphalt, or other paved areas, and other physical constraints were identified unless indicated below.'),
    bodyPara('No disassembly of systems or building components or physical or invasive testing was performed. ODIC renders no opinion as to the Property condition at un-surveyed and/or inaccessible portions of the Property. ODIC relies completely on the information, whether written, graphic, or verbal, provided by the Property contact or as shown on any documents reviewed or received from the Property contact, owner or agent, or municipal source, and assumes that information to be true and correct.'),
    spacer(160),
    subHeading('3.2   PROPERTY RECONNAISSANCE'),
    bodyPara('ODIC conducted interior and exterior observations of the Property with the intent to identify releases or material threat of future releases of hazardous substances or petroleum products to the environment. The table below lists items visually and/or physically observed.'),
    spacer(80),
    reconTable([
      ['PROPERTY RECONNAISSANCE', null],
      ['Hazardous Substances and Petroleum Products', null],
      [null, recon.hazardousSubstances || 'None were observed.'],
      ['Underground Storage Tanks (USTs), vent pipes, fill pipes, or access ways indicating USTs', null],
      [null, recon.usts || 'None were observed.'],
      ['Aboveground Storage Tanks (ASTs)', null],
      [null, recon.asts || 'None were observed.'],
      ['Drums, Totes, and Intermediate Bulk Containers', null],
      [null, recon.drums || 'None were observed.'],
      ['Standing Surface Water and Pools or Sumps Containing Liquids Likely to be Hazardous Substances or Petroleum Products', null],
      [null, recon.standingWater || 'None were observed.'],
      ['Hazardous Substance and Petroleum Product Containers Not in Connection With Identified Uses', null],
      [null, recon.unidentifiedContainers || 'None were observed.'],
      ['Unidentified Substance Containers', null],
      [null, recon.unknownContainers || 'None were observed.'],
      ['Stains or Corrosion on Floors, Walls, or Ceilings (except for staining from water)', null],
      [null, recon.stains || 'None were observed.'],
      ['Stained Soil or Pavement', null],
      [null, recon.stainedSoil || 'None were observed.'],
      ['Drains, Sumps, Wastewater Treatment Units, Clarifiers', null],
      [null, recon.drains || 'None were observed.'],
      ['Pits, Ponds, or Lagoons', null],
      [null, recon.pits || 'None were observed.'],
      ['Stressed Vegetation (other than from insufficient water)', null],
      [null, recon.stressedVegetation || 'None were observed.'],
      ['Areas that are apparently graded by non-natural causes suggesting solid waste disposal', null],
      [null, recon.gradedAreas || 'None were observed.'],
      ['Water/Wastewater — other liquid discharged from or to the Property', null],
      [null, recon.wastewater || 'None were observed, other than typical stormwater runoff.'],
      ['Wells (including dry wells, irrigation wells, injection wells, monitoring wells, abandoned wells)', null],
      [null, recon.wells || 'None were observed.'],
      ['Septic Systems or Cesspools', null],
      [null, recon.septic || 'None were observed.'],
      ['Unusual Areas of Pavement Patching (including possible boring locations)', null],
      [null, recon.pavementPatching || 'None were observed.'],
    ]),
    spacer(200),
    subHeading('3.3   DETAILED DESCRIPTION OF SITE RECONNAISSANCE AND ENVIRONMENTAL CONDITIONS'),
    ...textToParas(s.siteReconnaissanceDetailed || ''),
    spacer(200),
    subHeading('3.4   CURRENT USES OF ADJOINING PROPERTIES'),
    italicPara('For the scope of this assessment, properties are defined and categorized based upon their physical proximity to the Property. An adjoining property is defined as any real property or properties the border of which is contiguous or partially contiguous with that of the Property, or that would be contiguous or partially contiguous with that of the Property but for a street, road, or other public thoroughfare separating them.'),
    spacer(80),
    kvTable('Adjoining Properties', [
      ['North', d.adjNorth],
      ['Northeast', d.adjNE || `Refer to the north-adjacent property.`],
      ['East', d.adjEast],
      ['Southeast', d.adjSE],
      ['South', d.adjSouth],
      ['Southwest', d.adjSW || `Refer to the south-adjacent property.`],
      ['West', d.adjWest],
      ['Northwest', d.adjNW || `Refer to the north-adjacent property.`],
    ]),
    spacer(200),
    subHeading('3.5   NON-SCOPE (NON-ASTM) CONSIDERATIONS'),
    bodyPara('Evaluation of Non-Scope or Non-CERCLA items, including those addressed in Section 3.4 of this Report, is not required nor relevant for compliance with the AAI Rule or ASTM Standard Practice E1527-21. Inclusion of any non-scope item in a Phase I Environmental Site Assessment report is entirely within the discretion of the User based on its own risk tolerance.'),
    spacer(80),
    ...buildNonScopeTable(d),
    pageBreak(),
  ];
}

function buildNonScopeTable(d) {
  const buildYear = parseInt(d.constructionYear) || 1978;
  const preACBM = buildYear < 1978;
  const items = [
    ['Suspect asbestos-containing building materials (ACBM) unrelated to releases into the environment',
      preACBM
        ? `Since an asbestos survey is not included in the scope of services for this Phase I ESA, ODIC did not test suspect asbestos-containing building materials (ACBM) at the Property. However, because improvements at the Property were constructed prior to 1978, the presence of ACBM is possible. It is important to note that State and Federal Laws impose special requirements for handling these materials.`
        : `Since an asbestos survey is not included in the scope of services for this Phase I ESA, ODIC did not test for ACBM at the Property.`],
    ['Lead-based paint (LBP) unrelated to releases into the environment',
      preACBM
        ? `Since a lead-based paint survey is not included in the scope of services for this Phase I ESA, ODIC did not test suspect lead-based paint (LBP) at the Property. However, because improvements at the Property were constructed prior to 1978, the presence of LBP is possible.`
        : `A lead-based paint survey is not included in the scope of services for this Phase I ESA.`],
    ['Naturally-occurring radon', `Since a radon survey is not included in the scope of services for this Phase I ESA, ODIC did not test for radon at the Property.`],
    ['PCB-containing building materials', `Since a PCB survey is not included in the scope of services for this Phase I ESA, ODIC did not test for PCBs at the Property.`],
    ['Flood Zone', `The Federal Emergency Management Agency Flood Insurance Rate Map is typically used to determine if the Property is located within a flood zone. A flood zone evaluation is not included in the scope of work for this Phase I ESA.`],
    ['Methane Gas', `Since a methane gas survey is not included in the scope of services for this Phase I ESA, ODIC did not test for methane gas at the Property.`],
    ['Mold or microbial growth conditions', `Since a mold/microbial matter survey is not included in the scope of services for this Phase I ESA, ODIC did not test for mold/microbial matter at the Property.`],
  ];

  const rows = items.map(([label, text]) =>
    new TableRow({
      children: [
        new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: label, font: BODY_FONT, size: NORMAL_SIZE })] })],
          width: { size: 30, type: WidthType.PERCENTAGE },
          shading: { fill: 'F2F2F2', type: ShadingType.SOLID },
        }),
        new TableCell({
          children: textToParas(text),
          width: { size: 70, type: WidthType.PERCENTAGE },
        }),
      ],
    })
  );

  const headerRow = new TableRow({
    children: [
      new TableCell({
        children: [new Paragraph({
          children: [new TextRun({ text: 'NON-SCOPE, NON-CERCLA ITEMS', font: BODY_FONT, size: NORMAL_SIZE, bold: true, color: WHITE })],
          alignment: AlignmentType.CENTER,
        })],
        columnSpan: 2,
        shading: { fill: DARK_GRAY, type: ShadingType.SOLID, color: DARK_GRAY },
      }),
    ],
  });

  const subHeaderRow = new TableRow({
    children: [
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: 'ITEM', font: BODY_FONT, size: NORMAL_SIZE, bold: true })] })],
        shading: { fill: LIGHT_GRAY, type: ShadingType.SOLID },
      }),
      new TableCell({
        children: [new Paragraph({ children: [] })],
        shading: { fill: LIGHT_GRAY, type: ShadingType.SOLID },
      }),
    ],
  });

  return [
    new Table({
      rows: [headerRow, subHeaderRow, ...rows],
      width: { size: 100, type: WidthType.PERCENTAGE },
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    }),
  ];
}

function buildPropertyHistory(d, s) {
  const ph = s.propertyHistory || {};
  const dirs = d.cityDirectories || [];

  return [
    sectionHeading('4.0 PROPERTY AND VICINITY HISTORY'),
    bodyPara('The objective of consulting historical sources is to develop a history of the previous uses of the Property and surrounding area, in order to help identify the likelihood of past uses having led to Recognized Environmental Conditions (RECs), Controlled Recognized Environmental Conditions (CRECs), or Historical Recognized Environmental Conditions (HRECs) in connection with the Property.'),
    spacer(160),
    subHeading('4.1   PREVIOUS ENVIRONMENTAL REPORTS'),
    ...textToParas(ph.previousReports || 'See Section 5.2.'),
    spacer(160),
    subHeading('4.2   SANBORN MAP COMPANY FIRE INSURANCE MAPS'),
    ...textToParas(ph.sanbornMaps || 'Sanborn Map Company fire insurance maps were reviewed for the Property and vicinity. See findings below.'),
    spacer(160),
    subHeading('4.3   HISTORICAL AERIAL PHOTOGRAPHS'),
    ...textToParas(ph.aerialPhotographs || ''),
    spacer(160),
    subHeading('4.4   LOCAL STREET DIRECTORIES / HISTORICAL CITY DIRECTORIES'),
    ...textToParas(ph.cityDirectories || ''),
    spacer(80),
    ...(dirs.length > 0 ? [
      kvTable(`Historical City Directories\n\nProperty Address: ${d.propertyAddress || ''}`, [
        ['YEAR', 'LISTINGS'],
        ...dirs.map(item => [item.year || item.yearRange, item.listings]),
      ]),
    ] : []),
    spacer(160),
    subHeading('4.5   CITY/COUNTY BUILDING DEPARTMENT, ZONING/LAND USE, PROPERTY TAX RECORDS, PROFILES'),
    boldUnderline('Building Department Records:'),
    ...textToParas(ph.buildingDepartment || 'Building department records were obtained.'),
    spacer(80),
    ...(d.buildingPermits && d.buildingPermits.length > 0 ? [
      new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: 'Building Department Records', font: BODY_FONT, size: NORMAL_SIZE, bold: true })], alignment: AlignmentType.CENTER })],
                columnSpan: 3,
                shading: { fill: LIGHT_GRAY, type: ShadingType.SOLID },
              }),
            ],
          }),
          new TableRow({
            children: ['YEAR', 'DESCRIPTION', 'OWNER'].map(h =>
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: h, font: BODY_FONT, size: NORMAL_SIZE, bold: true })] })], shading: { fill: 'F2F2F2', type: ShadingType.SOLID } })
            ),
          }),
          ...d.buildingPermits.map(p =>
            new TableRow({
              children: [p.year, p.description, p.owner].map(v =>
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: v || '', font: BODY_FONT, size: NORMAL_SIZE })] })] })
              ),
            })
          ),
        ],
        width: { size: 100, type: WidthType.PERCENTAGE },
        margins: { top: 40, bottom: 40, left: 80, right: 80 },
      }),
    ] : []),
    spacer(120),
    boldUnderline('Property Profile:'),
    bodyPara('ODIC obtained the following Property information from a title company property profile:'),
    ...(d.propertyProfile ? Object.entries(d.propertyProfile).map(([k, v]) =>
      new Paragraph({
        children: [
          new TextRun({ text: `${k}: `, font: BODY_FONT, size: NORMAL_SIZE, bold: true }),
          new TextRun({ text: v || '', font: BODY_FONT, size: NORMAL_SIZE }),
        ],
        bullet: { level: 0 },
        spacing: { after: 60 },
      })
    ) : []),
    spacer(160),
    subHeading('4.6   HISTORICAL TOPOGRAPHIC MAPS'),
    ...textToParas(ph.topoMaps || 'Development history of the Property and surrounding area was researched using historical 7.5 minute USGS topographic maps. No features of environmental concern were identified.'),
    spacer(160),
    subHeading('4.7   OIL & GAS MAPS'),
    ...textToParas(ph.oilGasMaps || 'ODIC reviewed the Geologic Energy Management Division\'s (CalGEM) online mapping application Well Finder maps for the Property and immediate vicinity via the CalGEM Online Mapping System. No active or abandoned oil and/or gas wells were identified on the Property or in the immediate vicinity.'),
    spacer(160),
    subHeading('4.8   OTHER HISTORICAL RECORDS'),
    ...textToParas(ph.otherHistoricalRecords || 'An internet search of the Property address(es) was conducted to identify additional former Property occupant information that may not have been identified via other sources. None were identified.'),
    pageBreak(),
  ];
}

function buildDatabaseResearch(d, s) {
  const db = s.databaseFindings || {};
  const propListings = d.propertyDatabaseListings || {};

  return [
    sectionHeading('5.0 STANDARD ENVIRONMENTAL RECORDS RESEARCH'),
    spacer(120),
    subHeading('5.1   PROCEDURE'),
    bodyPara('The most current databases sources maintained by state and federal offices were provided by governmental record search database suppliers, such as Environmental Data Resources (EDR). For definitions of database acronyms, review the database report in Appendix C and/or refer to ASTM Standard Practice E1527-21, Sections 3.3 and 8.2.'),
    bodyPara('Database sources maintained by local offices were obtained via records requests. Databases were searched for properties with reported environmental listings within distances specified by ASTM Standard Practice E1527-21.'),
    spacer(160),
    subHeading('5.2   PROPERTY LISTING(S)'),
    spacer(80),
    ...(buildDatabaseListingsTable(propListings)),
    spacer(160),
    ...textToParas(db.propertyListingsNarrative || ''),
    spacer(200),
    subHeading('5.3   SURROUNDING SITES: FEDERAL AGENCY LISTINGS'),
    ...textToParas(db.surroundingFederalSites || ''),
    spacer(200),
    subHeading('5.4   SURROUNDING SITES: STATE AGENCY LISTINGS'),
    ...textToParas(db.surroundingStateSites || ''),
    spacer(200),
    subHeading('5.5   VAPOR ENCROACHMENT CONDITION'),
    ...textToParas(db.vaporEncroachment || ''),
    pageBreak(),
  ];
}

function buildDatabaseListingsTable(listings) {
  const federalDbs = [
    ['NPL', listings.npl],
    ['De-listed NPL', listings.delistedNpl],
    ['CERCLIS/SEMS', listings.cerclis],
    ['CERCLIS-NFRAP/SEMS-ARCHIVE', listings.cerclisNfrap],
    ['RCRA-CORRACTS', listings.rcraCorracts],
    ['RCRA-TSDF', listings.rcraTsdf],
    ['RCRA-Generator', listings.rcraGenerator],
    ['ERNS', listings.erns],
    ['Federal IC/EC Registries', listings.federalIcEc],
    ['Other Federal Listings', listings.otherFederal],
  ];

  const stateDbs = [
    ['State/Tribal Equivalent NPL', listings.stateNpl],
    ['State/Tribal Equivalent CERCLIS', listings.stateCerclis],
    ['State/Tribal SWLF', listings.stateSwlf],
    ['State/Tribal Voluntary Cleanup Sites', listings.stateVoluntary],
    ['State/Tribal Brownfield Sites', listings.stateBrownfield],
    ['State/Tribal Leaking Storage Tank', listings.stateLust],
    ['State/Tribal CPS / SLIC', listings.stateSlic],
    ['State/Tribal Registered Storage Tank', listings.stateRst],
    ['State/Tribal IC/EC Registries', listings.stateIcEc],
    ['Other State Listings', listings.otherState],
  ];

  const makeRow = ([label, val]) => new TableRow({
    children: [
      new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: label, font: BODY_FONT, size: SMALL_SIZE })] })], width: { size: 60, type: WidthType.PERCENTAGE } }),
      new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: val || 'No', font: BODY_FONT, size: SMALL_SIZE })] })], width: { size: 40, type: WidthType.PERCENTAGE } }),
    ],
  });

  const makeHeader = (text) => new TableRow({
    children: [
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text, font: BODY_FONT, size: SMALL_SIZE, bold: true, color: WHITE })], alignment: AlignmentType.CENTER })],
        columnSpan: 2,
        shading: { fill: DARK_GRAY, type: ShadingType.SOLID, color: DARK_GRAY },
      }),
    ],
  });

  const makeSubHeader = (cols) => new TableRow({
    children: cols.map(t => new TableCell({
      children: [new Paragraph({ children: [new TextRun({ text: t, font: BODY_FONT, size: SMALL_SIZE, bold: true })] })],
      shading: { fill: LIGHT_GRAY, type: ShadingType.SOLID },
    })),
  });

  return [
    new Table({
      rows: [
        makeHeader('FEDERAL AGENCY LISTINGS'),
        makeSubHeader(['DATABASE', 'PROPERTY LISTED']),
        ...federalDbs.map(makeRow),
      ],
      width: { size: 48, type: WidthType.PERCENTAGE },
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    }),
    spacer(120),
    new Table({
      rows: [
        makeHeader('STATE AGENCY LISTINGS'),
        makeSubHeader(['DATABASE', 'PROPERTY LISTED']),
        ...stateDbs.map(makeRow),
      ],
      width: { size: 48, type: WidthType.PERCENTAGE },
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
    }),
  ];
}

function buildUserProvidedInfo(d, s) {
  const ui = s.userProvidedInfo || {};
  return [
    sectionHeading('6.0 USER PROVIDED INFORMATION'),
    spacer(120),
    subHeading('6.1   USER PROVIDED INFORMATION'),
    ...textToParas(ui.userInfo || ''),
    spacer(160),
    subHeading('6.2   PRELIMINARY TITLE REPORT OR LAND TITLE RECORDS'),
    ...textToParas(ui.titleReport || ''),
    spacer(160),
    subHeading('6.3   INTERVIEWS'),
    ...textToParas(ui.interviews || ''),
    pageBreak(),
  ];
}

function buildReferences(d, s) {
  return [
    sectionHeading('7.0 REFERENCES'),
    ...textToParas(s.references || 'Environmental Data Resources (EDR) database report; USGS topographic maps; California State Water Resources Control Board GeoTracker; California DTSC ENVIROSTOR; CalGEM Well Finder.'),
    pageBreak(),
  ];
}

// ─── Main Builder ────────────────────────────────────────────────────────────

/**
 * Assembles the complete ODIC Phase I ESA DOCX document.
 * @param {object} formData - All form inputs
 * @param {object} sections - Claude-generated section text
 * @returns {Promise<Buffer>} - DOCX file as a Buffer
 */
export async function buildDocx(formData, sections) {
  const d = formData;
  const s = sections;
  const projectNumber = d.projectNumber || 'ESAI';
  const reportTitle = 'Phase I Environmental Site Assessment Report';

  const doc = new Document({
    creator: 'ODIC Environmental',
    title: `Phase I ESA – ${d.propertyAddress}`,
    description: `Phase I Environmental Site Assessment Report, Project No. ${projectNumber}`,
    styles: {
      default: {
        document: {
          run: { font: BODY_FONT, size: NORMAL_SIZE, color: BLACK },
          paragraph: { spacing: { line: 276 } },
        },
      },
    },
    sections: [
      // ── Section 1: Title Page (no header/footer) ──
      {
        properties: {
          page: {
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
          },
        },
        headers: { default: makeHeader() },
        footers: { default: makeFooter() },
        children: [
          ...buildTitlePage(d),
          ...buildTransmittalLetter(d),
          ...buildTableOfContents(d),
        ],
      },
      // ── Section 2: Report Body ──
      {
        properties: {
          page: {
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
          },
        },
        headers: {
          default: makeBodyHeader(projectNumber, reportTitle),
        },
        footers: { default: makeFooter() },
        children: [
          ...buildExecutiveSummary(d, s),
          ...buildFindingsAndRecommendations(d, s),
          ...buildIntroduction(d),
          ...buildPropertyDescription(d, s),
          ...buildPropertyReconnaissance(d, s),
          ...buildPropertyHistory(d, s),
          ...buildDatabaseResearch(d, s),
          ...buildUserProvidedInfo(d, s),
          ...buildReferences(d, s),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
}
