/**
 * POST /api/build-only
 *
 * Takes already-generated formData + sections JSON and builds DOCX + PDF.
 * No AI call — just document assembly. Completes in ~2-3 seconds.
 *
 * Used by compile_esai.py which calls Anthropic directly (no timeout),
 * then calls this endpoint to get the final documents.
 *
 * Request body: { formData, sections }
 * Response:     { success, docxBase64, bodyPdfBase64, filename, pdfFilename }
 */

import { buildDocx } from '../../lib/buildDocx';
import { buildBodyPdf } from '../../lib/buildBodyPdf';

export const config = {
  api: {
    bodyParser: { sizeLimit: '10mb' },
    responseLimit: '30mb',
  },
};

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', process.env.ALLOWED_ORIGIN || '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { formData, sections } = req.body || {};

  if (!formData || !sections) {
    return res.status(400).json({ error: 'formData and sections are required' });
  }

  try {
    const [docxBuffer, pdfBuffer] = await Promise.all([
      buildDocx(formData, sections),
      buildBodyPdf(formData, sections),
    ]);

    const slug = (formData.projectNumber || 'report').replace(/[^a-zA-Z0-9]/g, '-');

    return res.status(200).json({
      success: true,
      docxBase64: docxBuffer.toString('base64'),
      bodyPdfBase64: pdfBuffer.toString('base64'),
      filename: `${slug}-Phase1-ESA.docx`,
      pdfFilename: `${slug}-Phase1-ESA-body.pdf`,
    });
  } catch (err) {
    console.error('[build-only] error:', err);
    return res.status(500).json({ error: 'Document build failed', details: err.message });
  }
}
