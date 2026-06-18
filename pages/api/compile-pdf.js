/**
 * POST /api/compile-pdf
 *
 * Merges multiple PDF files (base64) into one final report PDF,
 * inserting labeled section divider pages between each section.
 *
 * Request body:
 * {
 *   projectInfo: {
 *     projectNumber:   string
 *     propertyAddress: string
 *     clientName:      string
 *     reportDate:      string
 *   },
 *   files: [
 *     {
 *       section:    string   — section key e.g. "REPORT_BODY"
 *       label:      string   — human label e.g. "Write Up / Main Report Body"
 *       order:      number   — sort order
 *       pdfBase64:  string   — base64-encoded PDF bytes
 *       filename:   string   — original filename (for logging)
 *     },
 *     ...
 *   ],
 *   mode: "draft" | "final"   — draft adds watermark (default: "final")
 * }
 *
 * Response:
 * {
 *   success: true,
 *   compiledPdfBase64: string,
 *   filename: string,
 *   pageCount: number,
 *   sectionsCompiled: string[]
 * }
 */

import { PDFDocument, rgb, StandardFonts, PageSizes } from 'pdf-lib';

// ── Brand ─────────────────────────────────────────────────────────────────────
const RED   = rgb(0.75, 0, 0);
const BLACK = rgb(0, 0, 0);
const WHITE = rgb(1, 1, 1);
const GRAY  = rgb(0.3, 0.3, 0.3);
const LGRAY = rgb(0.85, 0.85, 0.85);

// ── CORS ──────────────────────────────────────────────────────────────────────

function setCors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
}

export const config = {
  api: {
    bodyParser: { sizeLimit: '50mb' },
    responseLimit: '60mb',
  },
};

// ── Handler ───────────────────────────────────────────────────────────────────

export default async function handler(req, res) {
  setCors(res);
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { projectInfo = {}, files = [], mode = 'final' } = req.body || {};

  if (!Array.isArray(files) || files.length === 0) {
    return res.status(400).json({ error: 'files array is required and must not be empty' });
  }

  try {
    const { pdfBytes, pageCount, sectionsCompiled } = await compilePdf({ projectInfo, files, mode });
    const base64 = Buffer.from(pdfBytes).toString('base64');
    const slug = slugify(projectInfo.propertyAddress || projectInfo.projectNumber || 'report');
    const filename = `${slug}_Phase1ESA_${mode}_${dateStamp()}.pdf`;

    return res.status(200).json({
      success: true,
      compiledPdfBase64: base64,
      filename,
      pageCount,
      sectionsCompiled,
    });
  } catch (err) {
    console.error('[compile-pdf] error:', err);
    return res.status(500).json({ error: 'PDF compilation failed', details: err.message });
  }
}

// ── Core Compilation ──────────────────────────────────────────────────────────

async function compilePdf({ projectInfo, files, mode }) {
  const merged = await PDFDocument.create();

  // Embed fonts once — reused across divider pages
  const [boldFont, regularFont] = await Promise.all([
    merged.embedFont(StandardFonts.HelveticaBold),
    merged.embedFont(StandardFonts.Helvetica),
  ]);

  const isDraft = mode === 'draft';

  // Sort by order
  const sorted = [...files].sort((a, b) => (a.order ?? 999) - (b.order ?? 999));

  // Build cover page first
  await drawCoverPage(merged, boldFont, regularFont, projectInfo, mode);

  const sectionsCompiled = [];

  for (const file of sorted) {
    if (!file.pdfBase64) {
      console.warn(`[compile-pdf] Skipping ${file.filename} — no pdfBase64`);
      continue;
    }

    // Section divider page
    await drawDividerPage(merged, boldFont, regularFont, file.label || file.section, projectInfo);

    // Merge the PDF
    let srcBytes;
    try {
      srcBytes = Buffer.from(file.pdfBase64, 'base64');
    } catch {
      console.warn(`[compile-pdf] Could not decode base64 for ${file.filename}`);
      continue;
    }

    let srcDoc;
    try {
      srcDoc = await PDFDocument.load(srcBytes, { ignoreEncryption: true });
    } catch (e) {
      console.warn(`[compile-pdf] Could not parse PDF ${file.filename}: ${e.message}`);
      // Insert error page instead of crashing
      await drawErrorPage(merged, boldFont, regularFont, file.filename, e.message);
      continue;
    }

    const pageIndices = srcDoc.getPageIndices();
    const copiedPages = await merged.copyPages(srcDoc, pageIndices);

    for (const page of copiedPages) {
      if (isDraft) {
        // Overlay watermark text on each page
        addWatermark(page, boldFont);
      }
      merged.addPage(page);
    }

    sectionsCompiled.push(file.section || file.label || file.filename);
  }

  const pdfBytes = await merged.save();
  return { pdfBytes, pageCount: merged.getPageCount(), sectionsCompiled };
}

// ── Cover Page ────────────────────────────────────────────────────────────────

async function drawCoverPage(doc, boldFont, regularFont, projectInfo, mode) {
  const page = doc.addPage(PageSizes.Letter);
  const { width, height } = page.getSize();

  // Red header bar
  page.drawRectangle({ x: 0, y: height - 100, width, height: 100, color: RED });
  page.drawText('ODIC ENVIRONMENTAL', {
    x: 54,
    y: height - 52,
    font: boldFont,
    size: 24,
    color: WHITE,
  });
  page.drawText('Phase I Environmental Site Assessment — Compiled Report', {
    x: 54,
    y: height - 78,
    font: regularFont,
    size: 11,
    color: WHITE,
  });

  let y = height - 160;
  const field = (label, value) => {
    page.drawText(label, { x: 54, y, font: boldFont, size: 11, color: GRAY });
    page.drawText(value || 'N/A', { x: 200, y, font: regularFont, size: 11, color: BLACK });
    y -= 22;
  };

  page.drawText(projectInfo.propertyAddress || '', {
    x: 54, y, font: boldFont, size: 16, color: BLACK,
  });
  y -= 30;

  page.drawRectangle({ x: 54, y: y + 4, width: width - 108, height: 1.5, color: RED });
  y -= 18;

  field('Project Number:', projectInfo.projectNumber);
  field('Report Date:', projectInfo.reportDate);
  field('Client:', projectInfo.clientName);
  field('Mode:', mode.toUpperCase());

  // Footer
  page.drawRectangle({ x: 54, y: 60, width: width - 108, height: 1, color: LGRAY });
  page.drawText('ASTM E1527-21 | SBA SOP 50 10 8', {
    x: 54, y: 42, font: regularFont, size: 9, color: GRAY,
  });
  if (mode === 'draft') {
    page.drawText('DRAFT — NOT FOR DELIVERY', {
      x: width - 54 - 160, y: 42, font: boldFont, size: 9, color: RED,
    });
  }
}

// ── Section Divider ───────────────────────────────────────────────────────────

async function drawDividerPage(doc, boldFont, regularFont, label, projectInfo) {
  const page = doc.addPage(PageSizes.Letter);
  const { width, height } = page.getSize();

  // Full red background
  page.drawRectangle({ x: 0, y: 0, width, height, color: RED });

  // White centered label
  const fontSize = 22;
  const textWidth = boldFont.widthOfTextAtSize(label, fontSize);
  const x = Math.max(54, (width - textWidth) / 2);

  page.drawText(label, {
    x,
    y: height / 2,
    font: boldFont,
    size: fontSize,
    color: WHITE,
  });

  if (projectInfo.projectNumber) {
    const sub = `Project No. ${projectInfo.projectNumber}`;
    const subWidth = regularFont.widthOfTextAtSize(sub, 11);
    page.drawText(sub, {
      x: (width - subWidth) / 2,
      y: height / 2 - 32,
      font: regularFont,
      size: 11,
      color: rgb(1, 0.8, 0.8),
    });
  }
}

// ── Error Page ────────────────────────────────────────────────────────────────

async function drawErrorPage(doc, boldFont, regularFont, filename, errorMsg) {
  const page = doc.addPage(PageSizes.Letter);
  const { width, height } = page.getSize();

  page.drawText('⚠ File Could Not Be Included', {
    x: 54,
    y: height - 100,
    font: boldFont,
    size: 16,
    color: RED,
  });
  page.drawText(`Filename: ${filename}`, {
    x: 54,
    y: height - 140,
    font: regularFont,
    size: 11,
    color: BLACK,
  });
  page.drawText(`Error: ${(errorMsg || '').slice(0, 200)}`, {
    x: 54,
    y: height - 170,
    font: regularFont,
    size: 10,
    color: GRAY,
  });
}

// ── Watermark ─────────────────────────────────────────────────────────────────

function addWatermark(page, boldFont) {
  const { width, height } = page.getSize();
  try {
    page.drawText('DRAFT — NOT FOR DELIVERY', {
      x: width * 0.1,
      y: height * 0.45,
      font: boldFont,
      size: 40,
      color: rgb(0.8, 0, 0),
      opacity: 0.22,
      rotate: { type: 'degrees', angle: 35 },
    });
  } catch {
    // Some pages may not support opacity — skip gracefully
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function slugify(str) {
  return (str || 'report')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 60);
}

function dateStamp() {
  return new Date().toISOString().slice(0, 10).replace(/-/g, '');
}
