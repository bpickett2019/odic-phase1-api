/**
 * POST /api/smart-generate
 *
 * Minimal-input Phase I ESA generation.
 * Staff provides: project number, address, date, client name, and raw field notes.
 * Claude extracts all structured data AND writes all narrative sections.
 *
 * Request body:
 * {
 *   projectNumber:   string   — e.g. "6384578ESAI"
 *   propertyAddress: string   — e.g. "1212 E Ash Ave., Fullerton, CA 92831"
 *   reportDate:      string   — e.g. "January 6, 2026"
 *   clientName:      string?  — client name (optional)
 *   rawProjectData:  string?  — paste all field notes, database findings, history, etc.
 * }
 *
 * Response: same shape as /api/generate-report
 * {
 *   success: true,
 *   docxBase64, bodyPdfBase64, sections, formData, filename, pdfFilename
 * }
 */

import Anthropic from '@anthropic-ai/sdk';
import { buildDocx } from '../../lib/buildDocx';
import { buildBodyPdf } from '../../lib/buildBodyPdf';

export const config = {
  api: {
    bodyParser: { sizeLimit: '5mb' },
    responseLimit: '30mb',
  },
};

// Allow up to 5 minutes for Claude to generate the full report
export const maxDuration = 300;

// ── Handler ───────────────────────────────────────────────────────────────────

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', process.env.ALLOWED_ORIGIN || '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { projectNumber, propertyAddress, reportDate, clientName = '', rawProjectData = '' } = req.body || {};

  if (!projectNumber || !propertyAddress || !reportDate) {
    return res.status(400).json({
      error: 'projectNumber, propertyAddress, and reportDate are required',
    });
  }

  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY not configured' });
  }

  try {
    console.log(`[smart-generate] Starting for project ${projectNumber}`);

    const { formData, sections } = await extractAndGenerate({
      projectNumber, propertyAddress, reportDate, clientName, rawProjectData,
    });

    console.log('[smart-generate] Claude call complete — building documents...');
    const [docxBuffer, pdfBuffer] = await Promise.all([
      buildDocx(formData, sections),
      buildBodyPdf(formData, sections),
    ]);

    const slug = (projectNumber || 'report').replace(/[^a-zA-Z0-9]/g, '-');
    return res.status(200).json({
      success: true,
      docxBase64: docxBuffer.toString('base64'),
      bodyPdfBase64: pdfBuffer.toString('base64'),
      sections,
      formData,
      filename: `${slug}-Phase1-ESA.docx`,
      pdfFilename: `${slug}-Phase1-ESA-body.pdf`,
    });

  } catch (err) {
    console.error('[smart-generate] error:', err);
    return res.status(500).json({ error: 'Report generation failed', details: err.message });
  }
}

// ── Core AI call ──────────────────────────────────────────────────────────────

async function extractAndGenerate({ projectNumber, propertyAddress, reportDate, clientName, rawProjectData }) {
  const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

  const today = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });

  const systemPrompt = `You are a senior environmental professional at ODIC Environmental, specializing in Phase I Environmental Site Assessments (ESAs) under ASTM E1527-21 and SBA SOP 50 10 8.

Your task: given raw project data and minimal structured inputs, produce a complete, professional Phase I ESA report by extracting all relevant information and writing thorough, ASTM-compliant narratives.

Rules:
- Extract every piece of information you can find in the raw data
- Write professional, formal narrative prose for all section fields
- Use standard ASTM E1527-21 language and terminology
- Where data is not provided, use appropriate professional default language (e.g. "not identified during this assessment", "not observed during the site reconnaissance")
- RECs, HRECs, and CRECs should only be listed if explicitly mentioned in the raw data
- If no RECs are identified, recommendation should be "No Further Investigation is recommended"
- Return ONLY a valid JSON object — no markdown, no commentary, no preamble`;

  const userPrompt = `Generate a complete Phase I ESA report JSON from this project data.

REQUIRED FIELDS:
- Project Number: ${projectNumber}
- Property Address: ${propertyAddress}
- Report Date: ${reportDate}
- Client Name: ${clientName || 'Not provided'}
- Today's Date: ${today}

RAW PROJECT DATA / FIELD NOTES:
${rawProjectData || '(No additional data provided — use professional defaults throughout)'}

Return a JSON object with EXACTLY this structure. All string fields must be populated with professional content:

{
  "formData": {
    "projectNumber": "${projectNumber}",
    "reportDate": "${reportDate}",
    "propertyAddress": "${propertyAddress}",
    "apn": "",
    "alternateAddresses": null,
    "inspectionDate": "",
    "weatherCondition": "",
    "siteVisitBy": "Neil Kuemerle, Senior Environmental Consultant",
    "reportAuthor": "ODIC Environmental",
    "qaqcEP": "ODIC Environmental QA/QC Review",
    "propertyLocation": "",
    "generalSetting": "",
    "clientName": "${clientName || ''}",
    "clientCompany": "",
    "clientAddress": "",
    "borrowerName": null,
    "borrowerCompany": null,
    "lenders": [],
    "epName": "Neil Kuemerle",
    "epTitle": "Senior Consultant per §312.10 of 40 CFR 312",
    "firmRepName": "ODIC Environmental",
    "firmRepTitle": "Authorized Representative",
    "propertyDescription": "",
    "lotSizeAcres": "",
    "constructionYear": "",
    "improvementDescription": "",
    "otherImprovements": null,
    "presentOccupants": "",
    "waterSupply": "",
    "gasUtility": "",
    "electricUtility": "",
    "sewageSystem": "",
    "solidWaste": "",
    "septicSystem": null,
    "privateWell": null,
    "hvacSystem": "",
    "elevationFt": "",
    "groundwaterDepthFt": "",
    "groundwaterFlowDirection": "",
    "geologyNotes": null,
    "reconItems": {
      "hazardousSubstances": "None observed",
      "usts": "None observed",
      "asts": "None observed",
      "drums": "None observed",
      "standingWater": "None observed",
      "unidentifiedContainers": "None observed",
      "unknownContainers": "None observed",
      "stains": "None observed",
      "stainedSoil": "None observed",
      "drains": "None observed",
      "pits": "None observed",
      "stressedVegetation": "None observed",
      "gradedAreas": "None observed",
      "wastewater": "None observed",
      "wells": "None observed",
      "septic": "None observed",
      "pavementPatching": "None observed"
    },
    "siteReconNarrative": "",
    "adjNorth": "",
    "adjNE": null,
    "adjEast": "",
    "adjSE": null,
    "adjSouth": "",
    "adjSW": null,
    "adjWest": "",
    "adjNW": null,
    "previousEnvReports": null,
    "sanbornMapNote": "",
    "aerialPhotosYears": "",
    "aerialPhotoFindings": "",
    "cityDirectories": [],
    "buildingPermits": [],
    "propertyProfile": {
      "Current Property Owner": "",
      "Assessor Parcel Number": "",
      "Lot Size": "",
      "Building Size": "",
      "Construction Date": "",
      "Site Use / Use Code": ""
    },
    "topoMapFindings": null,
    "oilGasMapFindings": null,
    "otherHistoricalRecords": null,
    "propertyDatabaseListings": {
      "npl": "No",
      "delistedNpl": "No",
      "cerclis": "No",
      "cerclisNfrap": "No",
      "rcraCorracts": "No",
      "rcraTsdf": "No",
      "rcraGenerator": "No",
      "erns": "No",
      "federalIcEc": "No",
      "otherFederal": "No",
      "stateNpl": "No",
      "stateCerclis": "No",
      "stateSwlf": "No",
      "stateVoluntary": "No",
      "stateBrownfield": "No",
      "stateLust": "No",
      "stateSlic": "No",
      "stateRst": "No",
      "stateIcEc": "No",
      "otherState": "No"
    },
    "databaseFindingsNarrative": "",
    "surroundingFederalSites": "",
    "surroundingStateSites": "",
    "vaporEncroachment": "",
    "userProvidedInfo": "",
    "titleReportInfo": "",
    "interviewInfo": "",
    "recs": [],
    "hrecs": [],
    "crecs": [],
    "otherEnvironmentalIssues": null,
    "recommendation": "No Further Investigation is recommended at this time.",
    "sbaMitigatingFactors": ""
  },
  "sections": {
    "execSummary": "",
    "findingsAndRecommendations": "",
    "physicalSetting": "",
    "siteReconnaissanceDetailed": "",
    "propertyHistory": "",
    "databaseFindings": "",
    "userProvidedInfo": "",
    "references": "This Phase I Environmental Site Assessment was conducted in accordance with the scope and limitations of ASTM Practice E1527-21 Standard Practice for Environmental Site Assessments: Phase I Environmental Site Assessment Process. Additional references include the Small Business Administration Standard Operating Procedure (SBA SOP 50 10 8) and 40 CFR Part 312."
  }
}`;

  const message = await client.messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 8192,
    system: systemPrompt,
    messages: [{ role: 'user', content: userPrompt }],
  });

  const rawText = message.content.find((b) => b.type === 'text')?.text || '';
  const parsed = extractJson(rawText);

  if (!parsed.formData || !parsed.sections) {
    throw new Error('Claude response missing formData or sections');
  }

  // Ensure required fields are present
  parsed.formData.projectNumber = projectNumber;
  parsed.formData.propertyAddress = propertyAddress;
  parsed.formData.reportDate = reportDate;

  return parsed;
}

// ── JSON extraction ───────────────────────────────────────────────────────────

function extractJson(text) {
  let cleaned = text.trim().replace(/^```(?:json)?\s*/m, '').replace(/\s*```$/m, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch {
    const match = cleaned.match(/\{[\s\S]*\}/);
    if (match) {
      try { return JSON.parse(match[0]); } catch { /* fall through */ }
    }
    throw new Error(`Could not parse JSON from Claude response. Preview: ${cleaned.slice(0, 200)}`);
  }
}
