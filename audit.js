require('dotenv').config();
const OpenAI = require('openai');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const fieldDefs = require('./server/data/field-definitions.json');
const conditionLookup = require('./server/data/condition-lookup.json');
const formConfig = require('./server/data/form-config.json');

// Client/company name is configurable so this code is reusable across engagements.
const COMPANY_NAME = process.env.COMPANY_NAME || 'CLIENT';

function loadTaxonomy() {
  const filePath = path.join(__dirname, 'data/taxonomy_data.xlsx');
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets['Field_options'];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  const tags = rows
    .map(row => row['Tag'])
    .filter(tag => tag && tag.toString().trim() !== '');

  return 'SUPPORT TICKET TAXONOMY TAGS:\n' + tags.join('\n');
}

const TAXONOMY = loadTaxonomy();

const CHECKBOX_MAP = {
  ti_techtype_hardware: { fieldId: 45511186410267 },
  ti_techtype_software: { fieldId: 45511186410267 },
  fm01_app:             { fieldId: 45511202974107 },
  account:              { fieldId: 45511202974107 },
  retention_ticket:     { fieldId: 45511187352219 },
};

const MAX_PASSES = 7;
const SKIP_IDS = new Set(formConfig.skip_field_ids.map(String));

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

function stripPII(text) {
  return text
    .replace(/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g, '[EMAIL]')
    .replace(/\b(\+1[-.\s]?)?(\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})\b/g, '[PHONE]')
    .replace(/\b\d+\s+\w+(\s+\w+){0,3}\s+(St|Ave|Rd|Blvd|Dr|Ln|Way|Ct|Pl|Street|Avenue|Road|Boulevard|Drive|Lane|Court|Place)\.?\b/gi, '[ADDRESS]')
    .replace(/\b(?!(Agent|Customer|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b)[A-Z][a-z]+ [A-Z][a-z]+\b/g, '[NAME]');
}

function buildFieldPromptBlock(fieldIds) {
  const lines = [];
  for (const fid of fieldIds) {
    const def = fieldDefs[String(fid)];
    if (!def) continue;
    if (SKIP_IDS.has(String(fid))) continue;
    let desc = `- "${fid}" (${def.title})`;
    if (def.type === 'checkbox') {
      desc += ' — value must be true or false';
    } else if ((def.type === 'tagger' || def.type === 'multiselect') && def.options) {
      const optList = def.options.map(o => `"${o.value}"`).join(', ');
      desc += ` — pick ONE of: [${optList}]`;
    } else if (def.type === 'integer' || def.type === 'decimal') {
      desc += ' — numeric value';
    } else if (def.type === 'date') {
      desc += ' — ISO date string (YYYY-MM-DD)';
    } else {
      desc += ' — free text string';
    }
    lines.push(desc);
  }
  return lines.join('\n');
}

function buildDynamicSystemPrompt(fieldIds, previouslyFilled, passNumber) {
  const fieldBlock = buildFieldPromptBlock(fieldIds);
  let contextBlock = '';
  if (Object.keys(previouslyFilled).length > 0) {
    const lines = Object.entries(previouslyFilled).map(([fid, val]) => {
      const def = fieldDefs[String(fid)];
      const title = def ? def.title : fid;
      return `  "${title}": ${JSON.stringify(val)}`;
    });
    contextBlock = `\nThe following fields were already determined in previous passes:\n${lines.join('\n')}\n\nUse this context to inform your answers for the fields below.\n`;
  }
  return `You are a Zendesk ticket assistant for ${COMPANY_NAME}, a customer support operation.

Analyze the support call transcript and extract ticket field values.
Return ONLY a valid JSON object — no markdown, no explanation.
${contextBlock}
${passNumber === 0 ? `This is the initial pass. Focus on top-level category and checkbox fields.

IMPORTANT for 01 - Product List: Read the transcript carefully to identify the EXACT product model mentioned. The product names, variants, and their option values are configurable taxonomy values loaded from the taxonomy file. Pay close attention to distinguishing similar models, for example:
- "Product A v3" = product_a_v3 (NOT product_a_v4 — these are different products)
- "Product A v4" = product_a_v4
- "Product B variant" = product_b_variant (NOT product_b — these are different products)
- "Product B" alone = product_b
- If the customer names a product line without a version number and you cannot determine the exact variant, return null — do not guess.
- Always pick the most specific product tag that matches what was said. If unsure between two similar products, return null.


In addition to the fields listed below, also extract these fixed metadata fields and include them in your response:
- first_name     : string — customer's first name if they introduce themselves (e.g. "My name is Jane Smith"); null if absent
- last_name      : string — customer's last name from the same intro phrase; null if absent
- channel        : one of "Phone", "Email", "Chat", "Web Form" — infer from transcript context; null if unclear
- wrong_number   : boolean — true ONLY if the customer clearly reached the wrong department or company
- retention      : boolean — true if the customer expresses intent to cancel, return, or shows churn risk
- tech_issue     : boolean — true if the issue involves a technical problem with a device or software
- services_issue : boolean — true if the issue involves a subscription plan or account service
- operations_issue: boolean — true if the issue involves shipping, orders, billing, or inventory

Return these metadata fields using the same shape as all other fields:
{
  "first_name":       { "value": "Jane",   "confidence": "high", "reason": null },
  "last_name":        { "value": "Smith",  "confidence": "high", "reason": null },
  "channel":          { "value": "Phone",  "confidence": "high", "reason": null },
  "wrong_number":     { "value": false,    "confidence": "high", "reason": null },
  "retention":        { "value": false,    "confidence": "high", "reason": null },
  "tech_issue":       { "value": true,     "confidence": "high", "reason": null },
  "services_issue":   { "value": false,    "confidence": "high", "reason": null },
  "operations_issue": { "value": false,    "confidence": "high", "reason": null }
}\n` : `This is pass ${passNumber + 1}. These are sub-category fields revealed by previous selections.\n`}
Fill these fields:
${fieldBlock}

For EVERY field listed above, return this shape using the field ID as the key:
{
  "<field_id>": {
    "value": <extracted value or null>,
    "confidence": "high" | "medium" | "low",
    "reason": <string if confidence is low or value is null, otherwise null>
  }
}

Rules:
- For tagger/dropdown fields return exactly one of the listed option values, or null.
- For checkbox fields return true or false (boolean).
- For multiselect fields return an array of one or more option values, or null.
- For the "TI - Software List" field, ONLY return a value if TI - Technical Issue Type is "ti_techtype_software". In all other cases return null.
- If you cannot determine a value return null with a reason.
- For field 45511093273883 (Issue Sub-type): Only fill this field if you can identify a specific phrase in the transcript that maps to one of the listed values. If the issue is clearly a camera/doorbell problem but doesn't match any specific sub-type, "camera_doorbell_other" is acceptable. Do NOT use connectivity or feature sub-types unless the transcript explicitly mentions connection drops or specific feature problems by name.
- Return ONLY the JSON object.
- If two option values look similar, always pick the one whose wording most closely matches exact phrases used in the transcript. Do not infer or generalize.
- For product fields, match the exact model name stated by the customer. Similar-sounding variants (e.g. "Product A v3" and "Product A v4") are different products — do not confuse them.
- If you are not at least 80% confident in a value, return null with a reason rather than guessing. A null caught by the agent is better than a wrong value written silently to the ticket.
- Never return a value you cannot directly support with a specific phrase or statement from the transcript.
- For boolean/checkbox fields, default to false — only return true if there is clear explicit evidence in the transcript.
- For field 45511187352219 (no_pm_response): only set to true if the customer EXPLICITLY states they contacted ${COMPANY_NAME} before AND received no reply. Both conditions must be present. Phrases like "I've been waiting", "I sent an email", "I called before", or "I've had this problem for a while" do NOT qualify on their own. The customer must clearly say something like "I reached out and never heard back" or "I emailed and got no response". If only one condition is present, return false.`.trim();
}

function resolveChildFields(filledFields) {
  const triggered = new Map();
  for (const [fid, value] of Object.entries(filledFields)) {
    if (value === null || value === undefined) continue;
    const def = fieldDefs[String(fid)];
    if (!def) continue;
    let lookupValue;
    if (def.type === 'checkbox') {
      lookupValue = value === true || value === 'true' ? 'true' : 'false';
    } else if (Array.isArray(value)) {
      for (const v of value) {
        const key = `${fid}:${v}`;
        const cond = conditionLookup[key];
        if (cond) {
          for (const ch of cond.child_fields) {
            triggered.set(String(ch.id), { is_required: ch.is_required });
          }
        }
      }
      continue;
    } else {
      lookupValue = String(value);
    }
    const key = `${fid}:${lookupValue}`;
    const cond = conditionLookup[key];
    if (cond) {
      for (const ch of cond.child_fields) {
        triggered.set(String(ch.id), { is_required: ch.is_required });
      }
    }
  }
  return triggered;
}

function zdAuth() {
  const { ZENDESK_SUBDOMAIN, ZENDESK_EMAIL, ZENDESK_API_TOKEN } = process.env;
  const credentials = Buffer.from(`${ZENDESK_EMAIL}/token:${ZENDESK_API_TOKEN}`).toString('base64');
  return { subdomain: ZENDESK_SUBDOMAIN, headers: { 'Authorization': `Basic ${credentials}`, 'Content-Type': 'application/json' } };
}

async function runAIAnalysis(transcript) {
  const stripped = stripPII(transcript);
  const allResults = {};
  const filledValues = {};
  const seenFieldIds = new Set();

  let currentFieldIds = formConfig.always_visible_field_ids
    .map(String)
    .filter(id => !SKIP_IDS.has(id) && fieldDefs[id]);

  for (let pass = 0; pass < MAX_PASSES; pass++) {
    const newFieldIds = currentFieldIds.filter(id => !seenFieldIds.has(id));
    if (newFieldIds.length === 0) break;
    for (const id of newFieldIds) seenFieldIds.add(id);

    const systemPrompt = buildDynamicSystemPrompt(newFieldIds, filledValues, pass);
    console.log(`Pass ${pass + 1}: processing ${newFieldIds.length} fields`);

    const response = await openai.chat.completions.create({
      model: 'gpt-4o',
      temperature: 0,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: `Here is the tag taxonomy:\n\n${TAXONOMY}\n\nTranscript:\n${stripped}` },
      ],
    });

    const passFields = JSON.parse(response.choices[0].message.content);

    for (const [fid, info] of Object.entries(passFields)) {
      allResults[fid] = {
        ...info,
        field_title: fieldDefs[fid]?.title || fid,
        field_type: fieldDefs[fid]?.type || 'unknown',
        pass,
      };
      if (info.value !== null && info.value !== undefined) {
        filledValues[fid] = info.value;
      }
    }

    const triggeredChildren = resolveChildFields(filledValues);
    currentFieldIds = [...triggeredChildren.keys()].filter(
      id => !seenFieldIds.has(id) && !SKIP_IDS.has(id) && fieldDefs[id]
    );
    console.log(`Pass ${pass + 1} complete. Triggered ${currentFieldIds.length} new fields.`);
    if (currentFieldIds.length === 0) break;
  }

  const taxonomyTags = new Set(TAXONOMY.split('\n').slice(1));
  for (const [fid, info] of Object.entries(allResults)) {
    if (info.value === null || info.value === undefined) continue;
    const def = fieldDefs[String(fid)];
    if (Array.isArray(info.value)) {
      info.value = info.value.filter(tag => taxonomyTags.has(tag));
      if (info.value.length === 0) info.value = null;
    } else if (def?.options && def.options.length > 0) {
      const validValues = new Set(def.options.map(o => o.value));
      if (!validValues.has(info.value)) {
        info.value = null;
        info.confidence = 'low';
        info.reason = 'GPT returned a value not in the allowed option list';
      }
    }
  }

  return allResults;
}

async function fetchTicketData(ticketId) {
  const { subdomain, headers } = zdAuth();
  const res = await fetch(`https://${subdomain}/api/v2/tickets/${ticketId}.json`, { headers });
  if (!res.ok) throw new Error(`Zendesk ticket fetch failed: ${res.status}`);
  const data = await res.json();
  const flatTags = data.ticket.tags || [];
  const customFields = (data.ticket.custom_fields || [])
    .filter(f => f.value !== null && f.value !== false && f.value !== '')
    .map(f => ({ id: String(f.id), value: f.value }));
  return { flatTags, customFields };
}

async function fetchTranscript(ticketId) {
  const { subdomain, headers } = zdAuth();
  const res = await fetch(`https://${subdomain}/api/v2/tickets/${ticketId}/comments.json`, { headers });
  if (!res.ok) throw new Error(`Zendesk comments fetch failed: ${res.status}`);
  const data = await res.json();
  const comments = data.comments || [];
  if (comments.length === 0) throw new Error('No transcript found');
  const longest = comments.reduce((best, c) => {
    const body = c.plain_body || c.body || '';
    return body.length > (best.plain_body || best.body || '').length ? c : best;
  });
  const body = longest.plain_body || longest.body || '';
  if (!body) throw new Error('No transcript found');
  return body;
}

function compareTags(csrTags, aiTags) {
  const csrSet = new Set(csrTags);
  const aiSet  = new Set(aiTags);
  const matched = csrTags.filter(t => aiSet.has(t));
  const missing = csrTags.filter(t => !aiSet.has(t));
  const extra   = aiTags.filter(t => !csrSet.has(t));
  return {
    matched, missing, extra,
    matchCount:   matched.length,
    missingCount: missing.length,
    extraCount:   extra.length,
    totalCSR:     csrTags.length,
    totalAI:      aiTags.length,
  };
}

function compareFields(csrCustomFields, aiCustomFields) {
  const csrMap = new Map(csrCustomFields.map(f => [f.id, String(f.value)]));
  const aiMap  = new Map(aiCustomFields.map(f => [f.id, String(f.value)]));

  const matchedFields = [];
  const missingFields = [];
  const extraFields   = [];

  for (const [id, val] of csrMap) {
    if (aiMap.get(id) === val) {
      matchedFields.push(`${id}:${val}`);
    } else {
      missingFields.push(`${id}:${val}`);
    }
  }
  for (const [id, val] of aiMap) {
    if (csrMap.get(id) !== val) {
      extraFields.push(`${id}:${val}`);
    }
  }

  return {
    matchedFields, missingFields, extraFields,
    fieldMatchCount:   matchedFields.length,
    fieldMissingCount: missingFields.length,
    fieldExtraCount:   extraFields.length,
    totalCSRFields:    csrCustomFields.length,
    totalAIFields:     aiCustomFields.length,
  };
}

function buildExcel(rows) {
  const wb = XLSX.utils.book_new();

  const HEADERS = [
    'Ticket ID',        // A  0
    'CSR Tags',         // B  1
    'AI Tags',          // C  2
    'Matched Tags',     // D  3
    'Missing Tags',     // E  4
    'Extra Tags',       // F  5
    'Match Count',      // G  6
    'Missing Count',    // H  7
    'Extra Count',      // I  8
    'Total CSR Tags',   // J  9
    'Total AI Tags',    // K  10
    'Match Rate %',     // L  11
    'Notes',            // M  12
    'CSR Custom Fields',     // N  13
    'AI Custom Fields',      // O  14
    'Matched Fields',        // P  15
    'Missing Fields',        // Q  16
    'Extra Fields',          // R  17
    'Field Match Count',     // S  18
    'Field Missing Count',   // T  19
    'Field Extra Count',     // U  20
    'Total CSR Fields',      // V  21
    'Total AI Fields',       // W  22
    'Field Match Rate %',    // X  23
  ];

  const col = i => String.fromCharCode(65 + i); // A=0 … X=23

  const headerStyle = {
    font:  { name: 'Arial', bold: true },
    fill:  { fgColor: { rgb: 'BDD7EE' }, patternType: 'solid' },
    alignment: { horizontal: 'center' },
  };
  const baseFont = { name: 'Arial' };
  const yellowFill = { fgColor: { rgb: 'FFFF00' }, patternType: 'solid' };
  const orangeFill = { fgColor: { rgb: 'FFE0B2' }, patternType: 'solid' };
  const totalStyle = {
    font: { name: 'Arial', bold: true },
    fill: { fgColor: { rgb: 'D9D9D9' }, patternType: 'solid' },
  };

  const ws = {};

  // Header row (row 1)
  HEADERS.forEach((h, ci) => {
    ws[`${col(ci)}1`] = { v: h, t: 's', s: headerStyle };
  });

  // Data rows (starting row 2)
  rows.forEach((r, ri) => {
    const rowNum = ri + 2;
    const cells = [
      { v: r.ticketId,                          t: 's', s: { font: baseFont } },
      { v: (r.csrTags        || []).join(', '), t: 's', s: { font: baseFont } },
      { v: (r.aiTaxonomyTags || []).join(', '), t: 's', s: { font: baseFont } },
      { v: (r.matched        || []).join(', '), t: 's', s: { font: baseFont } },
      { v: (r.missing        || []).join(', '), t: 's', s: r.missingCount > 0 ? { font: baseFont, fill: yellowFill } : { font: baseFont } },
      { v: (r.extra          || []).join(', '), t: 's', s: r.extraCount   > 0 ? { font: baseFont, fill: orangeFill } : { font: baseFont } },
      { v: r.matchCount,        t: 'n', s: { font: baseFont } },
      { v: r.missingCount,      t: 'n', s: { font: baseFont } },
      { v: r.extraCount,        t: 'n', s: { font: baseFont } },
      { v: r.totalCSR,          t: 'n', s: { font: baseFont } },
      { v: r.totalAI,           t: 'n', s: { font: baseFont } },
      { f: `IFERROR(G${rowNum}/J${rowNum},"N/A")`, t: 'n', z: '0.0%', s: { font: baseFont } },
      { v: r.notes || '',                       t: 's', s: { font: baseFont } },
      { v: (r.csrCustomFields  || []).map(f => `${f.id}:${f.value}`).join(', '), t: 's', s: { font: baseFont } },
      { v: (r.aiCustomFields   || []).map(f => `${f.id}:${f.value}`).join(', '), t: 's', s: { font: baseFont } },
      { v: (r.matchedFields    || []).join(', '), t: 's', s: { font: baseFont } },
      { v: (r.missingFields    || []).join(', '), t: 's', s: r.fieldMissingCount > 0 ? { font: baseFont, fill: yellowFill } : { font: baseFont } },
      { v: (r.extraFields      || []).join(', '), t: 's', s: r.fieldExtraCount   > 0 ? { font: baseFont, fill: orangeFill } : { font: baseFont } },
      { v: r.fieldMatchCount,   t: 'n', s: { font: baseFont } },
      { v: r.fieldMissingCount, t: 'n', s: { font: baseFont } },
      { v: r.fieldExtraCount,   t: 'n', s: { font: baseFont } },
      { v: r.totalCSRFields,    t: 'n', s: { font: baseFont } },
      { v: r.totalAIFields,     t: 'n', s: { font: baseFont } },
      { f: `IFERROR(S${rowNum}/V${rowNum},"N/A")`, t: 'n', z: '0.0%', s: { font: baseFont } },
    ];
    cells.forEach((cell, ci) => {
      ws[`${col(ci)}${rowNum}`] = cell;
    });
  });

  // Totals row
  const totalRow = rows.length + 2;
  const dataStart = 2;
  const dataEnd = rows.length + 1;
  // indices of numeric columns that get SUM: G(6),H(7),I(8),J(9),K(10), S(18),T(19),U(20),V(21),W(22)
  const sumCols = new Set([6, 7, 8, 9, 10, 18, 19, 20, 21, 22]);
  // indices of % columns that get AVERAGE: L(11), X(23)
  const avgCols = new Set([11, 23]);

  HEADERS.forEach((_, ci) => {
    const addr = `${col(ci)}${totalRow}`;
    if (ci === 0) {
      ws[addr] = { v: 'TOTALS', t: 's', s: totalStyle };
    } else if (sumCols.has(ci)) {
      ws[addr] = { f: `SUM(${col(ci)}${dataStart}:${col(ci)}${dataEnd})`, t: 'n', s: totalStyle };
    } else if (avgCols.has(ci)) {
      ws[addr] = { f: `AVERAGE(${col(ci)}${dataStart}:${col(ci)}${dataEnd})`, t: 'n', z: '0.0%', s: totalStyle };
    } else {
      ws[addr] = { v: '', t: 's', s: totalStyle };
    }
  });

  ws['!ref'] = `A1:${col(HEADERS.length - 1)}${totalRow}`;
  ws['!cols'] = HEADERS.map(() => ({ wch: 20 }));

  XLSX.utils.book_append_sheet(wb, ws, 'Audit Results');
  XLSX.writeFile(wb, path.join(__dirname, 'accuracy_report.xlsx'));
  console.log(`Saved accuracy_report.xlsx (${rows.length} ticket${rows.length !== 1 ? 's' : ''})`);
}

const APP_MANAGED_FIELD_IDS = new Set([
  ...formConfig.always_visible_field_ids.map(String),
  ...Object.keys(conditionLookup).map(entry => entry.split(':')[0]),
]);

async function main() {
  const ids = fs.readFileSync(path.join(__dirname, 'ticket_ids.txt'), 'utf8')
    .split('\n')
    .map(l => l.trim())
    .filter(l => l.length > 0 && !l.startsWith('#'));

  console.log(`Starting audit for ${ids.length} tickets...`);

  const rows = [];

  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    console.log(`Processing ticket ${id} (${i + 1}/${ids.length})...`);

    try {
      const { flatTags, customFields: csrCustomFields } = await fetchTicketData(id);
      const transcript = await fetchTranscript(id);
      const allResults = await runAIAnalysis(transcript);

      const taxonomyTagSet = new Set(TAXONOMY.split('\n').slice(1).map(t => t.trim()).filter(Boolean));

      const aiTaxonomyTags = [];
      for (const [, info] of Object.entries(allResults)) {
        if (info.value === null || info.value === undefined) continue;
        if (typeof info.value === 'string' && taxonomyTagSet.has(info.value)) {
          aiTaxonomyTags.push(info.value);
        } else if (Array.isArray(info.value)) {
          aiTaxonomyTags.push(...info.value.filter(t => taxonomyTagSet.has(t)));
        }
      }

      const aiCustomFields = [];
      for (const [fid, info] of Object.entries(allResults)) {
        if (!fieldDefs[fid]) continue;
        if (info.value === null || info.value === undefined || info.value === false) continue;
        aiCustomFields.push({ id: fid, value: String(info.value) });
      }

      const aiFieldIds = new Set(aiCustomFields.map(f => f.id));
      const comparableCSRFields = csrCustomFields.filter(f => APP_MANAGED_FIELD_IDS.has(f.id));

      const tagComparison   = compareTags(flatTags, aiTaxonomyTags);
      const fieldComparison = compareFields(comparableCSRFields, aiCustomFields);

      rows.push({
        ticketId:          id,
        csrTags:           flatTags,
        aiTaxonomyTags,
        matched:           tagComparison.matched,
        missing:           tagComparison.missing,
        extra:             tagComparison.extra,
        matchCount:        tagComparison.matchCount,
        missingCount:      tagComparison.missingCount,
        extraCount:        tagComparison.extraCount,
        totalCSR:          tagComparison.totalCSR,
        totalAI:           tagComparison.totalAI,
        notes:             '',
        csrCustomFields,
        aiCustomFields,
        matchedFields:     fieldComparison.matchedFields,
        missingFields:     fieldComparison.missingFields,
        extraFields:       fieldComparison.extraFields,
        fieldMatchCount:   fieldComparison.fieldMatchCount,
        fieldMissingCount: fieldComparison.fieldMissingCount,
        fieldExtraCount:   fieldComparison.fieldExtraCount,
        totalCSRFields:    fieldComparison.totalCSRFields,
        totalAIFields:     fieldComparison.totalAIFields,
      });
    } catch (err) {
      console.log(`Error on ticket ${id}: ${err.message}`);
      rows.push({
        ticketId:          id,
        csrTags:           [],
        aiTaxonomyTags:    [],
        matched:           [],
        missing:           [],
        extra:             [],
        matchCount:        0,
        missingCount:      0,
        extraCount:        0,
        totalCSR:          0,
        totalAI:           0,
        notes:             `ERROR: ${err.message}`,
        csrCustomFields:   [],
        aiCustomFields:    [],
        matchedFields:     [],
        missingFields:     [],
        extraFields:       [],
        fieldMatchCount:   0,
        fieldMissingCount: 0,
        fieldExtraCount:   0,
        totalCSRFields:    0,
        totalAIFields:     0,
      });
    }

    if (i < ids.length - 1) await new Promise(r => setTimeout(r, 500));
  }

  buildExcel(rows);
  console.log('Done! Report saved to accuracy_report.xlsx');
}

main();
