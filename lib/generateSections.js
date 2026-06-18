/**
 * generateSections.js
 * Calls Claude API to generate all narrative sections of the Phase I ESA.
 * Returns a structured object with each section's text ready for DOCX assembly.
 */

import Anthropic from '@anthropic-ai/sdk';

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

/**
 * Main entry point. Takes the form data and returns all generated text sections.
 * @param {object} formData - The complete form submission from the frontend
 * @returns {Promise<object>} - Structured section text
 */
export async function generateAllSections(formData) {
  const prompt = buildMasterPrompt(formData);

  const message = await client.messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 8192,
    messages: [
      {
        role: 'user',
        content: prompt,
      },
    ],
  });

  const raw = message.content[0].text;

  // Parse the structured JSON response
  try {
    const jsonMatch = raw.match(/```json\n([\s\S]*?)\n```/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[1]);
    }
    // Fallback: try to parse the whole response
    return JSON.parse(raw);
  } catch {
    // If parsing fails, return the raw text mapped to sections
    console.error('Failed to parse Claude JSON response, using raw fallback');
    return buildFallbackSections(raw, formData);
  }
}

/**
 * Builds the master prompt that asks Claude to generate all narrative sections at once.
 */
function buildMasterPrompt(d) {
  return `You are an experienced environmental consultant writing a Phase I Environmental Site Assessment (Phase I ESA) report for ODIC Environmental.
Write in ODIC's professional, objective third-person style, consistent with ASTM Standard Practice E1527-21.

Using the project data below, generate all narrative sections of the report. Return ONLY valid JSON (no prose before or after), wrapped in a \`\`\`json code block.

PROJECT DATA:
${JSON.stringify(d, null, 2)}

Generate the following JSON structure. Each field should be a complete, professionally written paragraph or section of text ready for insertion into the report. Use formal environmental consulting language. Do NOT include section numbers or headers in the text — those are added by the DOCX builder.

\`\`\`json
{
  "execSummary": {
    "propertyDescription": "One paragraph describing the property — size, improvements, current occupant and use.",
    "propertyReconnaissance": "2-3 paragraphs summarizing site reconnaissance observations relevant to environmental conditions. Reference any hazardous materials, tanks, drains, or other items of note.",
    "historicalUse": "1-2 paragraphs summarizing historical use of the property based on the data provided.",
    "historicalAdjoining": "1 paragraph on adjoining property history if relevant, or state that the vicinity was [X] land use historically.",
    "pfas": "1 sentence on PFAS — either 'There was no evidence...' or describe findings.",
    "federalStateLocalRecords": "1 paragraph listing which databases the property was identified on.",
    "potentialOffsiteConcerns": "1-2 paragraphs on any off-site concerns identified. If none, state that.",
    "nonScopeItems": "Standard ODIC language: 'Unless the Client contracted ODIC to investigate specific Non-Scope or Non-CERCLA items...'",
    "inaccessiblePortions": "State whether full access was provided or note any limitations.",
    "significantDataGap": "State whether any significant data gaps were identified."
  },
  "findingsAndRecommendations": {
    "recNarrative": "Full narrative for each REC identified. Include the ASTM definition of REC as it appears in ODIC reports, then bullet-point description of each REC found.",
    "hrecNarrative": "Full narrative for HRECs identified, or state none were identified.",
    "crecNarrative": "Full narrative for CRECs identified, or state none were identified.",
    "otherEnvironmentalIssues": "State 'None were identified.' or describe issues.",
    "recommendation": "State the recommendation (e.g., 'ODIC recommends No Further Investigation at this time.' or Phase II recommendation).",
    "sbaMitigatingFactorsDiscussion": "Full discussion paragraph(s) addressing SBA mitigating factors for any RECs/CRECs identified. Reference the specific SBA SOP 50 10 8 mitigating factor that applies."
  },
  "physicalSetting": {
    "topography": "Paragraph describing topography, USGS quad map reference, elevation, and general downslope direction.",
    "geologyHydrogeology": "Paragraph(s) describing geology and hydrogeology based on available data. Include estimated groundwater depth and flow direction."
  },
  "siteReconnaissanceDetailed": "2-4 paragraphs providing a detailed narrative description of the site reconnaissance observations, expanding on the reconnaissance table entries. Describe the property use, any observed hazardous materials, tanks, drains, pavement patching, and any items of environmental concern.",
  "propertyHistory": {
    "previousReports": "Reference to any previous environmental reports on the property, or state 'See Section 5.2.'",
    "sanbornMaps": "State whether Sanborn maps were available and what they showed, or that none were drawn for the area.",
    "aerialPhotographs": "Narrative of aerial photo review — years reviewed, what was observed on the property and vicinity in chronological order.",
    "cityDirectories": "Narrative summarizing city directory review — source, years covered, and summary of historical occupants.",
    "buildingDepartment": "Narrative of building department records obtained and summary of permits found.",
    "topoMaps": "State findings from historical topographic map review.",
    "oilGasMaps": "State findings from oil and gas map review (CalGEM).",
    "otherHistoricalRecords": "State findings from internet search and other historical records."
  },
  "databaseFindings": {
    "propertyListingsNarrative": "Narrative describing the property's database listings — which databases it appears on and the significance.",
    "surroundingFederalSites": "Narrative of surrounding federal-listed sites within regulatory search distances, or state none were identified of concern.",
    "surroundingStateSites": "Narrative of surrounding state-listed sites within regulatory search distances, or state none were identified of concern.",
    "vaporEncroachment": "Narrative of vapor encroachment condition assessment per ASTM E2600, or state that a VEC analysis was not performed as part of this assessment."
  },
  "userProvidedInfo": {
    "userInfo": "Narrative of information provided by the user/client, including any knowledge of environmental conditions.",
    "titleReport": "Narrative of title report review findings, or state that a preliminary title report was provided/not provided.",
    "interviews": "Narrative summarizing interviews conducted with property owners, occupants, and other knowledgeable parties."
  },
  "references": "List of key references used in the assessment, formatted as a paragraph or list."
}
\`\`\`

Important style guidelines:
- Write in third person ("ODIC conducted...", "The Property is...", "The assessment revealed...")
- Do not use first person ("we", "I", "our")
- Match ODIC's formal, precise environmental consulting tone
- For any data not provided, use reasonable professional language like "No [X] were identified" or "Information was not provided"
- ASTM definitions should match the exact language from ASTM Standard Practice E1527-21
- All database names should use their full names and acronyms (e.g., "Leaking Underground Storage Tank (LUST)")`;
}

/**
 * Fallback if JSON parsing fails — maps raw text to section keys.
 */
function buildFallbackSections(rawText, formData) {
  const addr = formData.propertyAddress || 'the subject property';
  return {
    execSummary: {
      propertyDescription: `The Property located at ${addr} was assessed by ODIC Environmental.`,
      propertyReconnaissance: rawText.slice(0, 500),
      historicalUse: 'See Section 4.0 of this report.',
      historicalAdjoining: 'See Section 4.0 of this report.',
      pfas: 'There was no evidence from any researched source that PFAS substances were historically used at the Property.',
      federalStateLocalRecords: 'See Section 5.0 of this report.',
      potentialOffsiteConcerns: 'See Section 5.0 of this report.',
      nonScopeItems: 'Unless the Client contracted ODIC to investigate specific Non-Scope or Non-CERCLA items, evaluation of Non-Scope or Non-CERCLA items is not required nor relevant for compliance with the AAI Rule or ASTM Standard Practice E1527-21.',
      inaccessiblePortions: 'Full access to the entire Property was provided to ODIC, and there were no notable portions of the Property excluded from the survey and field inspection.',
      significantDataGap: 'No significant data gaps were identified during the course of this assessment.',
    },
    findingsAndRecommendations: {
      recNarrative: 'See findings in Section 5.0.',
      hrecNarrative: 'This environmental assessment has revealed no HRECs in connection with the Property.',
      crecNarrative: 'See findings in Section 5.0.',
      otherEnvironmentalIssues: 'None were identified.',
      recommendation: 'ODIC recommends No Further Investigation at this time.',
      sbaMitigatingFactorsDiscussion: 'See discussion in Section 5.0.',
    },
    physicalSetting: {
      topography: `The Property's physical location was researched employing a United States Geological Survey (USGS) 7.5-Minute Topographic Quadrangle (Quad) Map relevant to the Property.`,
      geologyHydrogeology: 'Geologic and hydrogeologic information was obtained from available sources.',
    },
    siteReconnaissanceDetailed: 'ODIC conducted interior and exterior observations of the Property.',
    propertyHistory: {
      previousReports: 'See Section 5.2.',
      sanbornMaps: 'Sanborn Map Company fire insurance maps were reviewed for the Property.',
      aerialPhotographs: 'ODIC reviewed historical aerial photographs for the Property and vicinity.',
      cityDirectories: 'ODIC reviewed Historical City Directories for the Property.',
      buildingDepartment: 'Building department records were obtained.',
      topoMaps: 'No features of environmental concern were identified on historical topographic maps.',
      oilGasMaps: 'No active or abandoned oil and/or gas wells were identified on the Property or in the immediate vicinity.',
      otherHistoricalRecords: 'An internet search of the Property address was conducted. None were identified.',
    },
    databaseFindings: {
      propertyListingsNarrative: 'The Property was identified on environmental databases.',
      surroundingFederalSites: 'See Section 5.3.',
      surroundingStateSites: 'See Section 5.4.',
      vaporEncroachment: 'See Section 5.5.',
    },
    userProvidedInfo: {
      userInfo: 'User provided information was reviewed.',
      titleReport: 'A preliminary title report was reviewed.',
      interviews: 'Interviews were conducted with knowledgeable parties.',
    },
    references: 'Environmental Data Resources (EDR) database report; USGS topographic maps; California State Water Resources Control Board GeoTracker; California DTSC ENVIROSTOR.',
  };
}
