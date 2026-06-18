/**
 * POST /api/classify-file
 *
 * Accepts a filename + extracted text content and uses Claude to classify
 * which section of the ODIC Phase I ESA report package the file belongs in.
 *
 * Request body:
 * {
 *   filename:       string   — original filename (e.g. "EDR_Report.pdf")
 *   textContent:    string   — first ~10,000 chars of extracted text (optional)
 *   contentType:    string   — MIME type (optional)
 *   projectContext: {        — optional project info for better classification
 *     projectNumber:    string
 *     propertyAddress:  string
 *     clientName:       string
 *   }
 * }
 *
 * Response:
 * {
 *   success: true,
 *   section: "APPENDIX_C",       // one of ALLOWED_SECTIONS
 *   label:   "APPENDIX C – DATABASE REPORT",
 *   confidence: 0.95,
 *   reason:  "...",
 *   excludeByDefault: false,
 *   risks: []
 * }
 */

import Anthropic from '@anthropic-ai/sdk';

// Mirrors ChatGPT's section definitions — same keys so frontends can merge
const SECTIONS = [
  { key: 'E_AND_O_INSURANCE',       label: 'E&O Insurance',                                     order: 10  },
  { key: 'RELIANCE_LETTER',         label: 'Reliance Letter',                                    order: 20  },
  { key: 'COVER',                   label: 'Cover Page / Transmittal',                           order: 30  },
  { key: 'REPORT_BODY',             label: 'Write Up / Main Report Body',                        order: 40  },
  { key: 'APPENDIX_A',             label: 'Appendix A – Property Location Map & Plot Plan',     order: 100 },
  { key: 'APPENDIX_B',             label: 'Appendix B – Property & Vicinity Photographs',       order: 200 },
  { key: 'APPENDIX_C',             label: 'Appendix C – Database Report (EDR/Radius)',          order: 300 },
  { key: 'APPENDIX_D_SANBORN',     label: 'Appendix D – Historical Records: Sanborn Maps',     order: 410 },
  { key: 'APPENDIX_D_AERIALS',     label: 'Appendix D – Historical Records: Aerials',          order: 420 },
  { key: 'APPENDIX_D_TOPOS',       label: 'Appendix D – Historical Records: Topos',            order: 430 },
  { key: 'APPENDIX_D_CITY_DIRECTORIES', label: 'Appendix D – Historical Records: City Directories', order: 440 },
  { key: 'APPENDIX_D_OTHER',       label: 'Appendix D – Historical Records: Other',            order: 490 },
  { key: 'APPENDIX_E_AGENCY',      label: 'Appendix E – Public Agency Records',                order: 510 },
  { key: 'APPENDIX_E_PRIOR_REPORTS', label: 'Appendix E – Prior Reports',                      order: 520 },
  { key: 'APPENDIX_E_OTHER',       label: 'Appendix E – Other Relevant Documents',             order: 590 },
  { key: 'APPENDIX_F',             label: 'Appendix F – Qualifications of Environmental Professional', order: 600 },
  { key: 'NEEDS_REVIEW',           label: 'Needs Review',                                      order: 900 },
  { key: 'EXCLUDED',               label: 'Excluded',                                          order: 990 },
];

const SECTION_MAP = Object.fromEntries(SECTIONS.map((s) => [s.key, s]));
const ALLOWED_KEYS = SECTIONS.map((s) => s.key);

// ── CORS ──────────────────────────────────────────────────────────────────────

function setCors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
}

export const config = {
  api: { bodyParser: { sizeLimit: '2mb' } },
};

// ── Handler ───────────────────────────────────────────────────────────────────

export default async function handler(req, res) {
  setCors(res);
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(500).json({ error: 'ANTHROPIC_API_KEY not configured' });

  const { filename, textContent = '', contentType = '', projectContext = {} } = req.body || {};

  if (!filename) return res.status(400).json({ error: 'filename is required' });

  try {
    const result = await classifyFile({ filename, textContent, contentType, projectContext }, apiKey);
    return res.status(200).json({ success: true, ...result });
  } catch (err) {
    console.error('[classify-file] error:', err);
    return res.status(500).json({ error: 'Classification failed', details: err.message });
  }
}

// ── Core Classification ───────────────────────────────────────────────────────

async function classifyFile({ filename, textContent, contentType, projectContext }, apiKey) {
  const client = new Anthropic({ apiKey });

  const sectionList = SECTIONS.map((s) => `  "${s.key}": "${s.label}"`).join('\n');

  const systemPrompt = `You are a Phase I Environmental Site Assessment (ESA) document classifier for ODIC Environmental.
Your job is to classify uploaded project files into the correct section of the ODIC report package.

ALLOWED SECTIONS (use exact keys):
${sectionList}

Rules:
- EDR, Radius Map, FirstSearch = APPENDIX_C
- Sanborn fire insurance maps = APPENDIX_D_SANBORN
- Aerial photographs = APPENDIX_D_AERIALS
- Topographic maps = APPENDIX_D_TOPOS
- City directories = APPENDIX_D_CITY_DIRECTORIES
- Site photos, property photographs = APPENDIX_B
- Property location maps, plot plans, site maps = APPENDIX_A
- EP qualifications, resumes, certifications = APPENDIX_F
- E&O insurance certificates = E_AND_O_INSURANCE
- Reliance letters, third-party reliance = RELIANCE_LETTER
- Prior Phase I reports = APPENDIX_E_PRIOR_REPORTS
- Fire department records, agency letters = APPENDIX_E_AGENCY
- The main written report narrative = REPORT_BODY
- Cover page, transmittal letter = COVER
- If you cannot determine the type, use NEEDS_REVIEW

Return ONLY a JSON object with no markdown:
{
  "section": "<SECTION_KEY>",
  "confidence": <0.0–1.0>,
  "reason": "<brief operator-facing reason>",
  "excludeByDefault": <true|false>,
  "risks": ["<risk tag>", ...]
}`;

  const userPrompt = JSON.stringify({
    filename,
    contentType: contentType || 'unknown',
    project: projectContext,
    extractedTextSample: (textContent || '').slice(0, 10000),
  });

  const message = await client.messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 512,
    temperature: 0,
    system: systemPrompt,
    messages: [{ role: 'user', content: `Classify this file. Return JSON only.\n\n${userPrompt}` }],
  });

  const rawText = message.content.find((b) => b.type === 'text')?.text || '';
  const parsed = extractJson(rawText);

  // Validate section key
  let section = parsed.section || 'NEEDS_REVIEW';
  if (!ALLOWED_KEYS.includes(section)) section = 'NEEDS_REVIEW';

  const sectionDef = SECTION_MAP[section];

  return {
    section,
    label: sectionDef?.label || section,
    order: sectionDef?.order || 900,
    confidence: Number(parsed.confidence ?? 0),
    reason: String(parsed.reason || 'No reason provided'),
    excludeByDefault: Boolean(parsed.excludeByDefault),
    risks: Array.isArray(parsed.risks) ? parsed.risks : [],
  };
}

// ── JSON Extraction ───────────────────────────────────────────────────────────

function extractJson(text) {
  let cleaned = text.trim();
  // Strip markdown fences
  cleaned = cleaned.replace(/^```(?:json)?/m, '').replace(/```$/m, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch {
    const match = cleaned.match(/\{[\s\S]*\}/);
    if (match) {
      try { return JSON.parse(match[0]); } catch { /* fall through */ }
    }
    return {};
  }
}
