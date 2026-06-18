/**
 * buildBodyPdf.js
 * Generates a clean narrative PDF of the Phase I ESA report body using pdf-lib.
 * Pure JS — no Puppeteer, no binary deps, works in Vercel serverless.
 *
 * Returns a Buffer containing the PDF bytes.
 */

import { PDFDocument, rgb, StandardFonts, PageSizes } from 'pdf-lib';

// ── Brand Colors ──────────────────────────────────────────────────────────────
const RED   = rgb(0.75, 0, 0);       // #C00000
const BLACK = rgb(0, 0, 0);
const GRAY  = rgb(0.25, 0.25, 0.25);
const LGRAY = rgb(0.85, 0.85, 0.85);

// ── Layout ────────────────────────────────────────────────────────────────────
const MARGIN     = 72;       // 1 inch
const LINE_H     = 14;       // normal line height
const BODY_SIZE  = 11;
const H1_SIZE    = 14;
const H2_SIZE    = 12;
const SMALL_SIZE = 9;

// ── Text Utilities ────────────────────────────────────────────────────────────

/**
 * Wrap a string to fit within maxWidth pixels given a font + size.
 * Returns array of line strings.
 */
async function wrapText(text, font, size, maxWidth) {
  if (!text) return [''];
  const words = text.split(' ');
  const lines = [];
  let current = '';

  for (const word of words) {
    const candidate = current ? `${current} ${word}` : word;
    const w = font.widthOfTextAtSize(candidate, size);
    if (w > maxWidth && current) {
      lines.push(current);
      current = word;
    } else {
      current = candidate;
    }
  }
  if (current) lines.push(current);
  return lines.length ? lines : [''];
}

// ── Drawing Context ───────────────────────────────────────────────────────────

class PageWriter {
  constructor(doc, fonts, header) {
    this.doc     = doc;
    this.fonts   = fonts;
    this.header  = header; // { projectNumber, propertyAddress }
    this.page    = null;
    this.y       = 0;
    this.pageNum = 0;
    this._newPage();
  }

  _newPage() {
    this.page    = this.doc.addPage(PageSizes.Letter);
    this.pageNum += 1;
    const { width, height } = this.page.getSize();
    this.width  = width;
    this.height = height;
    this.y      = height - MARGIN;

    // Running header (skip title page = page 1)
    if (this.pageNum > 1) {
      this._drawRunningHeader();
    }
    // Footer
    this._drawFooter();
  }

  _drawRunningHeader() {
    const { bold, regular } = this.fonts;
    const headerY = this.height - 40;

    // Red top bar
    this.page.drawRectangle({
      x: MARGIN,
      y: headerY + 10,
      width: this.width - 2 * MARGIN,
      height: 2,
      color: RED,
    });

    // Project number left, address right
    this.page.drawText(`Project No. ${this.header.projectNumber}`, {
      x: MARGIN,
      y: headerY - 4,
      font: bold,
      size: SMALL_SIZE,
      color: GRAY,
    });

    const addrText = this.header.propertyAddress || '';
    const addrW = bold.widthOfTextAtSize(addrText, SMALL_SIZE);
    this.page.drawText(addrText, {
      x: this.width - MARGIN - addrW,
      y: headerY - 4,
      font: bold,
      size: SMALL_SIZE,
      color: GRAY,
    });

    this.y = headerY - 20;
  }

  _drawFooter() {
    const { regular } = this.fonts;
    const footerY = MARGIN - 20;

    this.page.drawRectangle({
      x: MARGIN,
      y: footerY + 12,
      width: this.width - 2 * MARGIN,
      height: 1,
      color: LGRAY,
    });

    this.page.drawText(`Page ${this.pageNum}`, {
      x: this.width / 2 - 20,
      y: footerY - 2,
      font: regular,
      size: SMALL_SIZE,
      color: GRAY,
    });

    this.page.drawText('ODIC Environmental', {
      x: MARGIN,
      y: footerY - 2,
      font: regular,
      size: SMALL_SIZE,
      color: GRAY,
    });

    this.page.drawText('DRAFT — NOT FOR DISTRIBUTION', {
      x: this.width - MARGIN - 160,
      y: footerY - 2,
      font: regular,
      size: SMALL_SIZE,
      color: GRAY,
    });
  }

  ensureSpace(needed) {
    if (this.y - needed < MARGIN + 30) {
      this._newPage();
    }
  }

  drawLine(text, font, size, color = BLACK, indent = 0) {
    if (this.y - LINE_H < MARGIN + 30) this._newPage();
    this.page.drawText(text || '', {
      x: MARGIN + indent,
      y: this.y,
      font,
      size,
      color,
    });
    this.y -= LINE_H;
  }

  space(amount = LINE_H * 0.5) {
    this.y -= amount;
    if (this.y < MARGIN + 30) this._newPage();
  }

  async drawParagraph(text, font, size, color = BLACK, indent = 0) {
    if (!text) return;
    const maxW = this.width - 2 * MARGIN - indent;
    const lines = await wrapText(text, font, size, maxW);
    for (const line of lines) {
      this.drawLine(line, font, size, color, indent);
    }
  }

  drawH1(text) {
    this.space(LINE_H);
    this.ensureSpace(H1_SIZE + LINE_H * 2);

    // Red underline bar
    const textW = this.fonts.bold.widthOfTextAtSize(text, H1_SIZE);
    this.page.drawRectangle({
      x: MARGIN,
      y: this.y - 3,
      width: this.width - 2 * MARGIN,
      height: 1.5,
      color: RED,
    });
    this.page.drawText(text, {
      x: MARGIN,
      y: this.y,
      font: this.fonts.bold,
      size: H1_SIZE,
      color: RED,
    });
    this.y -= H1_SIZE + 6;
    this.space(4);
  }

  drawH2(text) {
    this.space(LINE_H * 0.5);
    this.ensureSpace(H2_SIZE + LINE_H);
    this.page.drawText(text, {
      x: MARGIN,
      y: this.y,
      font: this.fonts.bold,
      size: H2_SIZE,
      color: BLACK,
    });
    this.y -= H2_SIZE + 4;
  }

  drawBullet(text) {
    return this.drawParagraph(`•  ${text}`, this.fonts.regular, BODY_SIZE, BLACK, 12);
  }

  drawKV(label, value) {
    return this.drawParagraph(`${label}: ${value || 'N/A'}`, this.fonts.regular, BODY_SIZE);
  }
}

// ── Title Page ────────────────────────────────────────────────────────────────

async function drawTitlePage(doc, fonts, form) {
  const page = doc.addPage(PageSizes.Letter);
  const { width, height } = page.getSize();
  const { bold, regular } = fonts;

  // Red header bar
  page.drawRectangle({ x: 0, y: height - 90, width, height: 90, color: RED });

  page.drawText('ODIC ENVIRONMENTAL', {
    x: MARGIN,
    y: height - 45,
    font: bold,
    size: 22,
    color: rgb(1, 1, 1),
  });
  page.drawText('Phase I Environmental Site Assessment', {
    x: MARGIN,
    y: height - 68,
    font: regular,
    size: 12,
    color: rgb(1, 1, 1),
  });

  // Main content block
  let y = height - 160;

  const field = (label, val) => {
    page.drawText(label, { x: MARGIN, y, font: bold, size: 11, color: GRAY });
    page.drawText(val || 'N/A', { x: MARGIN + 130, y, font: regular, size: 11, color: BLACK });
    y -= 20;
  };

  page.drawText(form.propertyAddress || '', {
    x: MARGIN,
    y,
    font: bold,
    size: 16,
    color: BLACK,
  });
  y -= 28;

  page.drawRectangle({ x: MARGIN, y, width: width - 2 * MARGIN, height: 1.5, color: RED });
  y -= 20;

  field('Project Number:', form.projectNumber);
  field('Report Date:', form.reportDate);
  field('Prepared For:', form.clientName);
  field('Property Type:', form.propertyType);
  field('REC Determination:', form.recDetermination);

  // Footer
  page.drawRectangle({ x: MARGIN, y: MARGIN + 40, width: width - 2 * MARGIN, height: 1, color: LGRAY });
  page.drawText('ASTM E1527-21 | SBA SOP 50 10 8', {
    x: MARGIN,
    y: MARGIN + 24,
    font: regular,
    size: SMALL_SIZE,
    color: GRAY,
  });
  page.drawText('CONFIDENTIAL', {
    x: width - MARGIN - 80,
    y: MARGIN + 24,
    font: bold,
    size: SMALL_SIZE,
    color: RED,
  });
}

// ── Main Export ───────────────────────────────────────────────────────────────

export async function buildBodyPdf(formData, sections) {
  const doc = await PDFDocument.create();

  const [regularFont, boldFont, italicFont] = await Promise.all([
    doc.embedFont(StandardFonts.TimesRoman),
    doc.embedFont(StandardFonts.TimesRomanBold),
    doc.embedFont(StandardFonts.TimesRomanItalic),
  ]);

  const fonts = { regular: regularFont, bold: boldFont, italic: italicFont };

  // ── Title Page ────────────────────────────────────────────────────────────
  await drawTitlePage(doc, fonts, formData);

  const header = {
    projectNumber: formData.projectNumber,
    propertyAddress: formData.propertyAddress,
  };
  const w = new PageWriter(doc, fonts, header);

  // ── Executive Summary ─────────────────────────────────────────────────────
  const es = sections.execSummary || {};
  w.drawH1('EXECUTIVE SUMMARY');

  const rows = [
    ['Property Address', formData.propertyAddress],
    ['Project Number', formData.projectNumber],
    ['Report Date', formData.reportDate],
    ['Client', formData.clientName],
    ['Property Type', formData.propertyType],
    ['Site Area (approx.)', formData.siteArea],
    ['Current Use', formData.currentUse],
    ['REC Determination', formData.recDetermination],
  ];
  for (const [label, val] of rows) {
    await w.drawKV(label, val);
  }
  w.space();

  if (es.summaryNarrative) {
    await w.drawParagraph(es.summaryNarrative, fonts.regular, BODY_SIZE);
    w.space();
  }

  if (es.recDetails) {
    w.drawH2('REC Details');
    await w.drawParagraph(es.recDetails, fonts.regular, BODY_SIZE);
    w.space();
  }

  if (es.hrecDetails) {
    w.drawH2('HREC Details');
    await w.drawParagraph(es.hrecDetails, fonts.regular, BODY_SIZE);
    w.space();
  }

  if (es.crecDetails) {
    w.drawH2('CREC Details');
    await w.drawParagraph(es.crecDetails, fonts.regular, BODY_SIZE);
    w.space();
  }

  // ── Findings & Recommendations ────────────────────────────────────────────
  const fr = sections.findingsAndRecommendations || {};
  w.drawH1('FINDINGS AND RECOMMENDATIONS');

  if (fr.findings) {
    w.drawH2('Findings');
    await w.drawParagraph(fr.findings, fonts.regular, BODY_SIZE);
    w.space();
  }
  if (fr.recommendations) {
    w.drawH2('Recommendations');
    await w.drawParagraph(fr.recommendations, fonts.regular, BODY_SIZE);
    w.space();
  }
  if (fr.sbaCompliance) {
    w.drawH2('SBA Compliance');
    await w.drawParagraph(fr.sbaCompliance, fonts.regular, BODY_SIZE);
    w.space();
  }
  if (fr.sbaMitigatingFactorsDiscussion) {
    w.drawH2('Mitigating Factors');
    await w.drawParagraph(fr.sbaMitigatingFactorsDiscussion, fonts.regular, BODY_SIZE);
    w.space();
  }

  // ── Section 1: Introduction ───────────────────────────────────────────────
  w.drawH1('SECTION 1 – INTRODUCTION');
  const intro = [
    ['Purpose', formData.purposeScope],
    ['Limiting Conditions', formData.limitingConditions],
    ['Site Visit Date', formData.siteVisitDate],
    ['Environmental Professional', formData.epName],
  ];
  for (const [label, val] of intro) {
    if (val) {
      w.drawH2(label);
      await w.drawParagraph(val, fonts.regular, BODY_SIZE);
      w.space();
    }
  }

  // ── Section 2: Property Description ──────────────────────────────────────
  w.drawH1('SECTION 2 – PROPERTY DESCRIPTION');
  const ps = sections.physicalSetting || '';
  await w.drawParagraph(ps, fonts.regular, BODY_SIZE);
  w.space();

  // ── Section 3: Site Reconnaissance ───────────────────────────────────────
  w.drawH1('SECTION 3 – SITE RECONNAISSANCE');
  const sr = sections.siteReconnaissanceDetailed || '';
  await w.drawParagraph(sr, fonts.regular, BODY_SIZE);
  w.space();

  // Non-Scope items
  if (formData.nonScopeItems && formData.nonScopeItems.length > 0) {
    w.drawH2('Non-Scope Considerations');
    for (const item of formData.nonScopeItems) {
      if (item.item) await w.drawBullet(`${item.item}: ${item.observed || 'N/A'} — ${item.remarks || ''}`);
    }
    w.space();
  }

  // ── Section 4: Property History ───────────────────────────────────────────
  w.drawH1('SECTION 4 – PROPERTY HISTORY');
  const ph = sections.propertyHistory || {};
  const histItems = [
    ['Historical Overview', ph.historicalOverview],
    ['Sanborn Maps', ph.sanbornMaps],
    ['Aerial Photographs', ph.aerialPhotographs],
    ['City Directories', ph.cityDirectories],
    ['Topographic Maps', ph.topoMaps],
    ['Previous Reports', ph.previousReports],
    ['Chain of Title', ph.chainOfTitle],
    ['Summary', ph.historySummary],
  ];
  for (const [label, val] of histItems) {
    if (val) {
      w.drawH2(label);
      await w.drawParagraph(val, fonts.regular, BODY_SIZE);
      w.space();
    }
  }

  // ── Section 5: Database Research ──────────────────────────────────────────
  w.drawH1('SECTION 5 – DATABASE RESEARCH');
  const db = sections.databaseFindings || {};
  const dbItems = [
    ['Database Overview', db.databaseOverview],
    ['Federal Listings', db.federalListings],
    ['State Listings', db.stateListings],
    ['Findings Summary', db.findingsSummary],
  ];
  for (const [label, val] of dbItems) {
    if (val) {
      w.drawH2(label);
      await w.drawParagraph(val, fonts.regular, BODY_SIZE);
      w.space();
    }
  }

  // Database sites table (if any)
  if (formData.databaseSites && formData.databaseSites.length > 0) {
    w.drawH2('Identified Database Sites');
    for (const site of formData.databaseSites) {
      await w.drawBullet(`${site.name || ''} (${site.distance || ''}) — ${site.database || ''}: ${site.status || ''}`);
    }
    w.space();
  }

  // ── Section 6: Additional Information ────────────────────────────────────
  const ui = sections.userProvidedInfo || {};
  if (ui.ownerInterviewSummary || ui.governmentRecordsSummary || ui.additionalInfo) {
    w.drawH1('SECTION 6 – ADDITIONAL INFORMATION');
    if (ui.ownerInterviewSummary) {
      w.drawH2('Owner/Occupant Interview');
      await w.drawParagraph(ui.ownerInterviewSummary, fonts.regular, BODY_SIZE);
      w.space();
    }
    if (ui.governmentRecordsSummary) {
      w.drawH2('Government Records');
      await w.drawParagraph(ui.governmentRecordsSummary, fonts.regular, BODY_SIZE);
      w.space();
    }
    if (ui.additionalInfo) {
      w.drawH2('Additional Information');
      await w.drawParagraph(ui.additionalInfo, fonts.regular, BODY_SIZE);
      w.space();
    }
  }

  // ── Section 7: References ─────────────────────────────────────────────────
  const refs = sections.references || '';
  if (refs) {
    w.drawH1('SECTION 7 – REFERENCES');
    await w.drawParagraph(refs, fonts.regular, BODY_SIZE);
  }

  // ── Serialize ─────────────────────────────────────────────────────────────
  const pdfBytes = await doc.save();
  return Buffer.from(pdfBytes);
}
