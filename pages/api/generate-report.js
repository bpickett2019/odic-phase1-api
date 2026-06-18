/**
 * POST /api/generate-report
 *
 * Accepts Phase I ESA form data, calls Claude to generate narrative sections,
 * assembles a Word document (.docx), and returns:
 *   - docxBase64: base64-encoded DOCX for frontend download
 *   - sections:   structured JSON of all generated text (use for PDF preview)
 *   - filename:   suggested download filename
 *
 * ──────────────────────────────────────────────────────────────────────────────
 * REQUEST BODY SCHEMA (application/json)
 * ──────────────────────────────────────────────────────────────────────────────
 * {
 *   // ── Project / Client ──────────────────────────────────────────────────
 *   projectNumber:       string   // e.g. "6384578ESAI"
 *   reportDate:          string   // e.g. "January 6, 2026"
 *   propertyAddress:     string   // e.g. "1212 E Ash Ave., Fullerton, CA 92831"
 *   apn:                 string   // Tax Assessor's Parcel Number
 *   alternateAddresses:  string?  // Historical/alternate addresses (or null)
 *   inspectionDate:      string   // e.g. "December 22, 2025"
 *   weatherCondition:    string   // e.g. "Partly cloudy, approximately 60°F"
 *   siteVisitBy:         string   // "Neil Kuemerle, Senior Environmental Consultant"
 *   reportAuthor:        string
 *   qaqcEP:              string   // QA/QC Environmental Professional
 *   propertyLocation:    string   // Descriptive location statement
 *   generalSetting:      string   // "Industrial" | "Commercial" | "Residential" etc.
 *
 *   // ── Client / Lender ───────────────────────────────────────────────────
 *   clientName:          string
 *   clientCompany:       string
 *   clientAddress:       string
 *   borrowerName:        string?
 *   borrowerCompany:     string?
 *   lenders:             Array<{ name: string, company: string, address: string, role?: string }>
 *
 *   // ── EP / Firm ─────────────────────────────────────────────────────────
 *   epName:              string   // Environmental Professional name
 *   epTitle:             string   // e.g. "Senior Consultant per §312.10 of 40 CFR 312"
 *   firmRepName:         string   // Authorized firm representative
 *   firmRepTitle:        string
 *
 *   // ── Property Description ──────────────────────────────────────────────
 *   propertyDescription: string   // Full narrative property description
 *   lotSizeAcres:        string   // e.g. "0.36"
 *   constructionYear:    string   // e.g. "1954/1955"
 *   improvementDescription: string
 *   otherImprovements:   string?
 *   presentOccupants:    string
 *   waterSupply:         string
 *   gasUtility:          string
 *   electricUtility:     string
 *   sewageSystem:        string
 *   solidWaste:          string
 *   septicSystem:        string?
 *   privateWell:         string?
 *   hvacSystem:          string
 *
 *   // ── Physical Setting ──────────────────────────────────────────────────
 *   elevationFt:         string   // e.g. "170"
 *   groundwaterDepthFt:  string   // e.g. "108"
 *   groundwaterFlowDirection: string  // e.g. "southwest"
 *   geologyNotes:        string?  // Any known geology / prior report data
 *
 *   // ── Property Reconnaissance ───────────────────────────────────────────
 *   reconItems: {
 *     hazardousSubstances: string
 *     usts:               string
 *     asts:               string
 *     drums:              string
 *     standingWater:      string
 *     unidentifiedContainers: string
 *     unknownContainers:  string
 *     stains:             string
 *     stainedSoil:        string
 *     drains:             string
 *     pits:               string
 *     stressedVegetation: string
 *     gradedAreas:        string
 *     wastewater:         string
 *     wells:              string
 *     septic:             string
 *     pavementPatching:   string
 *   }
 *   siteReconNarrative:  string   // Your field notes / descriptive observations
 *
 *   // ── Adjoining Properties ──────────────────────────────────────────────
 *   adjNorth:   string
 *   adjNE:      string?
 *   adjEast:    string
 *   adjSE:      string
 *   adjSouth:   string
 *   adjSW:      string?
 *   adjWest:    string
 *   adjNW:      string?
 *
 *   // ── Property History ─────────────────────────────────────────────────
 *   previousEnvReports:   string?  // Summary of any prior Phase I/II reports
 *   sanbornMapNote:       string?  // What Sanborn maps showed (or none available)
 *   aerialPhotosYears:    string?  // e.g. "1938, 1947, 1953, 1963, 1972, 1977..."
 *   aerialPhotoFindings:  string   // What aerials showed by decade
 *   cityDirectories: Array<{ yearRange: string, listings: string }>
 *   buildingPermits: Array<{ year: string, description: string, owner: string }>
 *   propertyProfile: {             // From title company
 *     "Current Property Owner": string
 *     "Assessor Parcel Number":  string
 *     "Lot Size":                string
 *     "Building Size":           string
 *     "Construction Date":       string
 *     "Site Use / Use Code":     string
 *   }
 *   topoMapFindings:      string?
 *   oilGasMapFindings:    string?
 *   otherHistoricalRecords: string?
 *
 *   // ── Database Research ─────────────────────────────────────────────────
 *   propertyDatabaseListings: {
 *     // Federal (Yes | No | "Yes - [detail]")
 *     npl:           string
 *     delistedNpl:   string
 *     cerclis:       string
 *     cerclisNfrap:  string
 *     rcraCorracts:  string
 *     rcraTsdf:      string
 *     rcraGenerator: string
 *     erns:          string
 *     federalIcEc:   string
 *     otherFederal:  string
 *     // State
 *     stateNpl:      string
 *     stateCerclis:  string
 *     stateSwlf:     string
 *     stateVoluntary: string
 *     stateBrownfield: string
 *     stateLust:     string
 *     stateSlic:     string
 *     stateRst:      string
 *     stateIcEc:     string
 *     otherState:    string
 *   }
 *   databaseFindingsNarrative: string  // Detailed narrative of each DB finding
 *   surroundingFederalSites:   string  // Summary of surrounding federal listings
 *   surroundingStateSites:     string  // Summary of surrounding state listings
 *   vaporEncroachment:         string  // VEC analysis narrative
 *
 *   // ── User Provided Info ────────────────────────────────────────────────
 *   userProvidedInfo:    string
 *   titleReportInfo:     string
 *   interviewInfo:       string
 *
 *   // ── RECs / Findings ───────────────────────────────────────────────────
 *   recs:  Array<string>   // Each REC description
 *   hrecs: Array<string>   // Each HREC description (often empty)
 *   crecs: Array<string>   // Each CREC description
 *   otherEnvironmentalIssues: string?
 *   recommendation: string  // e.g. "No Further Investigation"
 *   sbaMitigatingFactors: string  // Discussion of SBA mitigating factors
 * }
 * ──────────────────────────────────────────────────────────────────────────────
 * RESPONSE BODY
 * ──────────────────────────────────────────────────────────────────────────────
 * {
 *   success: true,
 *   docxBase64: string,   // base64 DOCX — decode and offer as download
 *   sections: object,     // all generated text by section key
 *   filename: string      // e.g. "6384578ESAI-Phase1-Report.docx"
 * }
 * ──────────────────────────────────────────────────────────────────────────────
 */

import { generateAllSections } from '../../lib/generateSections';
import { buildDocx } from '../../lib/buildDocx';
import { buildBodyPdf } from '../../lib/buildBodyPdf';

export const config = {
  api: {
    bodyParser: {
      sizeLimit: '10mb',
    },
    responseLimit: '30mb',
  },
};

export default async function handler(req, res) {
  // ── CORS headers (adjust origin in production) ────────────────────────────
  res.setHeader('Access-Control-Allow-Origin', process.env.ALLOWED_ORIGIN || '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed. Use POST.' });
  }

  // ── Basic validation ──────────────────────────────────────────────────────
  const formData = req.body;
  if (!formData || typeof formData !== 'object') {
    return res.status(400).json({ error: 'Request body must be a JSON object.' });
  }

  const required = ['projectNumber', 'propertyAddress', 'reportDate'];
  const missing = required.filter(k => !formData[k]);
  if (missing.length > 0) {
    return res.status(400).json({ error: `Missing required fields: ${missing.join(', ')}` });
  }

  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY environment variable is not set.' });
  }

  try {
    console.log(`[generate-report] Starting report for project ${formData.projectNumber}`);

    // ── Step 1: Generate all narrative sections via Claude ────────────────
    console.log('[generate-report] Calling Claude for section generation...');
    const sections = await generateAllSections(formData);
    console.log('[generate-report] Section generation complete.');

    // ── Step 2: Build DOCX + PDF in parallel ─────────────────────────────
    console.log('[generate-report] Building DOCX and body PDF...');
    const [docxBuffer, pdfBuffer] = await Promise.all([
      buildDocx(formData, sections),
      buildBodyPdf(formData, sections),
    ]);
    const docxBase64 = docxBuffer.toString('base64');
    const bodyPdfBase64 = pdfBuffer.toString('base64');
    console.log('[generate-report] DOCX + PDF build complete.');

    const slug = (formData.projectNumber || 'report').replace(/[^a-zA-Z0-9]/g, '-');
    const docxFilename = `${slug}-Phase1-ESA.docx`;
    const pdfFilename  = `${slug}-Phase1-ESA-body.pdf`;

    return res.status(200).json({
      success: true,
      docxBase64,
      bodyPdfBase64,
      sections,
      filename: docxFilename,
      pdfFilename,
    });
  } catch (err) {
    console.error('[generate-report] Error:', err);
    return res.status(500).json({
      error: 'Report generation failed.',
      details: err.message,
    });
  }
}
