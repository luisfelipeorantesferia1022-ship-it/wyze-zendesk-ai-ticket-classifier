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
    .replace(/\b\d+\s+\w+(\s+\w+){0,3}\s+(St|Ave|Rd|Blvd|Dr|Ln|Way|Ct|Pl|Street|Avenue|Road|Boulevard|Drive|Lane|Court|Place)\.?\b/gi, '[ADDRESS]');
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


CRITICAL CATEGORY ROUTING:
- Account ACCESS problems are TECH issues, not services. "Can't log in", "password reset", "account hacked", "locked out", "2FA not working" → tech_issue = true (software/account). Do NOT set services_issue for these.
- Subscription problems are SERVICES issues. "My plan isn't showing", "I paid for the premium plan but it shows the free tier", "cancel my subscription", "wrong plan on my account" → services_issue = true. The word "account" in a subscription context does NOT make it a tech issue.
- Account verification is a normal tech-support step. Do NOT switch a tech call to services just because the agent looked up the customer's email, order, or account details during troubleshooting.
- Warranty, return, or replacement workflows add operations ON TOP of the root cause. If the agent starts a warranty replacement, return, or order process for a defective device, set BOTH tech_issue = true (the root problem) AND operations_issue = true (the resolution workflow). This is the one expected case where two categories are true at once.
- If the customer mentions a subscription plan, billing, charges, or any subscription issue, you MUST set services_issue to true. If you see the fields for '01 - Service Plan Name' (45511155209371) or '02 - Service Issue' (45511050711835) in any pass, you MUST populate them (e.g., selecting the matching plan or issue value) if the transcript supports it.
- REFUND vs CANCELLATION: A customer asking for money back or disputing a charge is a CHARGE issue, NOT a cancellation. "I want a refund", "what is this charge", "I was charged after I canceled" → treat as a billing/charge issue. Only treat it as a cancellation if the customer clearly wants to stop the subscription going forward.
- CHANNEL DETECTION: Infer channel from how the conversation runs. A spoken call transcript ("Thank you for calling") = "Phone". Back-and-forth typed messages with an agent = "Chat". A single written message = "Email". Automated routing text at the very start does not decide the channel.
- CATEGORY DISCIPLINE: Do not tick more than one category checkbox unless the transcript clearly spans more than one. The only routine exception is a defective-device call that triggers a warranty/return, which is tech_issue AND operations_issue (per the routing guidance above). A return being mentioned does not by itself make a tech call an operations call — an actual order/return/warranty action must be taken.

In addition to the fields listed below, also extract these fixed metadata fields and include them in your response:
- full_name      : string — the customer's full name exactly as they state it (e.g. "Barbara Jenkins"); null if they never give a name
- channel        : one of "Phone", "Email", "Chat", "Web Form" — infer from transcript context; null if unclear
- wrong_number   : boolean — true ONLY if the customer clearly reached the wrong department or company
- retention      : boolean — true ONLY if the customer is visibly angry/frustrated AND explicitly threatens to cancel their subscription, demand a refund, or leave ${COMPANY_NAME} for a competitor. Do NOT set to true for standard, polite return requests or basic subscription cancellations.
- tech_issue     : boolean — true if the issue involves a technical problem with a device or software, including account access problems (login, password reset, 2FA)
- services_issue : boolean — true if the issue involves a subscription or billing (subscription plan, charges, cancellation). Do NOT set true for account access/login problems — those are tech issues.
- operations_issue: boolean — true if the call involves shipping, physical hardware orders, or a warranty/return/replacement workflow. Can be true alongside tech_issue when the agent initiates a return or replacement for a defective device. Do NOT set true for subscription billing or digital service charges.

Return these metadata fields using the same shape as all other fields:
{
  "full_name":        { "value": "Barbara Jenkins", "confidence": "high", "reason": null },
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
- For field 45511093273883 (Issue Sub-type): Only fill this field if you can identify a specific phrase in the transcript that maps to one of the listed values. Do NOT use connectivity or feature sub-types unless the transcript explicitly mentions connection drops or specific feature problems by name. You must NEVER use 'camera_doorbell_other' if the primary issue is related to an Account, the mobile App, or a Subscription. Only use it if the issue is strictly a physical camera hardware problem that does not fit any other hardware category.
- Do NOT set the field for 'purchase inquiry' (or any variant of it) unless the customer is explicitly asking a pre-sales question about buying a new product. Do not trigger it if they are just mentioning that they previously bought something or are asking for a warranty replacement.
- Return ONLY the JSON object.
- If two option values look similar, always pick the one whose wording most closely matches exact phrases used in the transcript. Do not infer or generalize.
- For product fields, match the exact model name stated by the customer. Similar-sounding variants (e.g. "Product A v3" and "Product A v4") are different products — do not confuse them.
- If you are not at least 80% confident in a value, return null with a reason rather than guessing. A null caught by the agent is better than a wrong value written silently to the ticket.
- Never return a value you cannot directly support with a specific phrase or statement from the transcript.
- For boolean/checkbox fields, default to false — only return true if there is clear explicit evidence in the transcript.
- For field 45511186554395 (no_pm_response): only set to true if the customer EXPLICITLY states they contacted ${COMPANY_NAME} before AND received no reply. Both conditions must be present. Phrases like "I've been waiting", "I sent an email", "I called before", or "I've had this problem for a while" do NOT qualify on their own. The customer must clearly say something like "I reached out and never heard back" or "I emailed and got no response". If only one condition is present, return false. If you flag this as true, you MUST be able to quote the exact phrase where the customer says they were previously ignored.`.trim();
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

    // Split full_name into first/last so downstream + frontend still work
    if (passFields.full_name && passFields.full_name.value) {
      const parts = String(passFields.full_name.value).trim().split(/\s+/);
      const conf = passFields.full_name.confidence || 'high';
      passFields.first_name = { value: parts[0] || null, confidence: conf, reason: null };
      passFields.last_name  = { value: parts.length > 1 ? parts.slice(1).join(' ') : null, confidence: conf, reason: null };
    }

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
    'Ticket ID',           // A   0
    'CSR Tags',            // B   1
    'AI Tags',             // C   2
    'Matched Tags',        // D   3
    'Missing Tags',        // E   4
    'Extra Tags',          // F   5
    'Match Count',         // G   6
    'Missing Count',       // H   7
    'Extra Count',         // I   8
    'Total CSR Tags',      // J   9
    'Total AI Tags',       // K  10
    'Match Rate %',        // L  11
    'High Conf AI Tags',   // M  12
    'High Matched',        // N  13
    'High Missing',        // O  14
    'High Extra',          // P  15
    'High Match Rate %',   // Q  16
    'Med Conf AI Tags',    // R  17
    'Med Matched',         // S  18
    'Med Missing',         // T  19
    'Med Extra',           // U  20
    'Med Match Rate %',    // V  21
    'Low Conf AI Tags',    // W  22
    'Low Matched',         // X  23
    'Low Missing',         // Y  24
    'Low Extra',           // Z  25
    'Low Match Rate %',    // AA 26
    'Notes',               // AB 27
    'CSR Custom Fields',   // AC 28
    'AI Custom Fields',    // AD 29
    'Matched Fields',      // AE 30
    'Missing Fields',      // AF 31
    'Extra Fields',        // AG 32
    'Field Match Count',   // AH 33
    'Field Missing Count', // AI 34
    'Field Extra Count',   // AJ 35
    'Total CSR Fields',        // AK 36
    'Total AI Fields',         // AL 37
    'Field Match Rate %',      // AM 38
    'High Conf AI Fields',     // AN 39
    'High Matched Fields',     // AO 40
    'High Missing Fields',     // AP 41
    'High Extra Fields',       // AQ 42
    'High Field Match Rate %', // AR 43
    'High Field Recall %',     // AS 44
    'Med Conf AI Fields',      // AT 45
    'Med Matched Fields',      // AU 46
    'Med Missing Fields',      // AV 47
    'Med Extra Fields',        // AW 48
    'Med Field Match Rate %',  // AX 49
    'Low Conf AI Fields',      // AY 50
    'Low Matched Fields',      // AZ 51
    'Low Missing Fields',      // BA 52
    'Low Extra Fields',        // BB 53
    'Low Field Match Rate %',  // BC 54
  ];

  const col = i => i < 26
    ? String.fromCharCode(65 + i)
    : String.fromCharCode(64 + Math.floor(i / 26)) + String.fromCharCode(65 + (i % 26));

  const ws = {};

  // Header row (row 1)
  HEADERS.forEach((h, ci) => {
    ws[`${col(ci)}1`] = { v: h, t: 's' };
  });

  // Data rows (starting row 2)
  rows.forEach((r, ri) => {
    const rowNum = ri + 2;
    const cells = [
      { v: r.ticketId,                           t: 's' },
      { v: (r.csrTags        || []).join(', '),  t: 's' },
      { v: (r.aiTaxonomyTags || []).join(', '),  t: 's' },
      { v: (r.matched        || []).join(', '),  t: 's' },
      { v: (r.missing        || []).join(', '),  t: 's' },
      { v: (r.extra          || []).join(', '),  t: 's' },
      { v: r.matchCount,         t: 'n' },
      { v: r.missingCount,       t: 'n' },
      { v: r.extraCount,         t: 'n' },
      { v: r.totalCSR,           t: 'n' },
      { v: r.totalAI,            t: 'n' },
    ];
    const rateAll         = r.totalCSR > 0 ? r.matchCount / r.totalCSR : null;
    const rateHigh        = (r.aiTagsHigh      || []).length > 0 ? r.matchCountHigh        / r.aiTagsHigh.length        : null;
    const rateMedium      = (r.aiTagsMedium    || []).length > 0 ? r.matchCountMedium      / r.aiTagsMedium.length      : null;
    const rateLow         = (r.aiTagsLow       || []).length > 0 ? r.matchCountLow         / r.aiTagsLow.length         : null;
    const rateField       = r.totalCSRFields > 0 ? r.fieldMatchCount / r.totalCSRFields : null;
    const rateFieldHigh   = (r.aiFieldsHigh   || []).length > 0 ? r.fieldMatchCountHigh   / r.aiFieldsHigh.length   : null;
    const recallFieldHigh = r.totalCSRFields  > 0               ? r.fieldMatchCountHigh   / r.totalCSRFields         : null;
    const rateFieldMedium = (r.aiFieldsMedium || []).length > 0 ? r.fieldMatchCountMedium / r.aiFieldsMedium.length : null;
    const rateFieldLow    = (r.aiFieldsLow    || []).length > 0 ? r.fieldMatchCountLow    / r.aiFieldsLow.length    : null;
    cells.push(
      { v: rateAll    ?? '', t: rateAll    !== null ? 'n' : 's', z: '0.0%' },
      { v: (r.aiTagsHigh    || []).join(', '),   t: 's' },
      { v: (r.matchedHigh   || []).join(', '),   t: 's' },
      { v: (r.missingHigh   || []).join(', '),   t: 's' },
      { v: (r.extraHigh     || []).join(', '),   t: 's' },
      { v: rateHigh   ?? '', t: rateHigh   !== null ? 'n' : 's', z: '0.0%' },
      { v: (r.aiTagsMedium  || []).join(', '),   t: 's' },
      { v: (r.matchedMedium || []).join(', '),   t: 's' },
      { v: (r.missingMedium || []).join(', '),   t: 's' },
      { v: (r.extraMedium   || []).join(', '),   t: 's' },
      { v: rateMedium ?? '', t: rateMedium !== null ? 'n' : 's', z: '0.0%' },
      { v: (r.aiTagsLow     || []).join(', '),   t: 's' },
      { v: (r.matchedLow    || []).join(', '),   t: 's' },
      { v: (r.missingLow    || []).join(', '),   t: 's' },
      { v: (r.extraLow      || []).join(', '),   t: 's' },
      { v: rateLow    ?? '', t: rateLow    !== null ? 'n' : 's', z: '0.0%' },
      { v: r.notes || '',                        t: 's' },
      { v: (r.csrCustomFields || []).map(f => `${f.id}:${f.value}`).join(', '), t: 's' },
      { v: (r.aiCustomFields  || []).map(f => `${f.id}:${f.value}`).join(', '), t: 's' },
      { v: (r.matchedFields   || []).join(', '), t: 's' },
      { v: (r.missingFields   || []).join(', '), t: 's' },
      { v: (r.extraFields     || []).join(', '), t: 's' },
      { v: r.fieldMatchCount,    t: 'n' },
      { v: r.fieldMissingCount,  t: 'n' },
      { v: r.fieldExtraCount,    t: 'n' },
      { v: r.totalCSRFields,     t: 'n' },
      { v: r.totalAIFields,      t: 'n' },
      { v: rateField       ?? '', t: rateField       !== null ? 'n' : 's', z: '0.0%' },
      { v: (r.aiFieldsHigh      || []).map(f => `${f.id}:${f.value}`).join(', '), t: 's' },
      { v: (r.matchedFieldsHigh || []).join(', '), t: 's' },
      { v: (r.missingFieldsHigh || []).join(', '), t: 's' },
      { v: (r.extraFieldsHigh   || []).join(', '), t: 's' },
      { v: rateFieldHigh   ?? '', t: rateFieldHigh   !== null ? 'n' : 's', z: '0.0%' },
      { v: recallFieldHigh ?? '', t: recallFieldHigh !== null ? 'n' : 's', z: '0.0%' },
      { v: (r.aiFieldsMedium      || []).map(f => `${f.id}:${f.value}`).join(', '), t: 's' },
      { v: (r.matchedFieldsMedium || []).join(', '), t: 's' },
      { v: (r.missingFieldsMedium || []).join(', '), t: 's' },
      { v: (r.extraFieldsMedium   || []).join(', '), t: 's' },
      { v: rateFieldMedium ?? '', t: rateFieldMedium !== null ? 'n' : 's', z: '0.0%' },
      { v: (r.aiFieldsLow      || []).map(f => `${f.id}:${f.value}`).join(', '), t: 's' },
      { v: (r.matchedFieldsLow || []).join(', '), t: 's' },
      { v: (r.missingFieldsLow || []).join(', '), t: 's' },
      { v: (r.extraFieldsLow   || []).join(', '), t: 's' },
      { v: rateFieldLow    ?? '', t: rateFieldLow    !== null ? 'n' : 's', z: '0.0%' },
    );
    cells.forEach((cell, ci) => {
      ws[`${col(ci)}${rowNum}`] = cell;
    });
  });

  // Totals row
  const totalRow = rows.length + 2;
  const dataStart = 2;
  const dataEnd = rows.length + 1;
  // indices of numeric columns that get SUM: G(6)-K(10) tag counts, AH(33)-AL(37) field counts
  const sumCols = new Set([6, 7, 8, 9, 10, 33, 34, 35, 36, 37]);

  const avgNonNull = vals => {
    const valid = vals.filter(v => v !== null);
    return valid.length > 0 ? valid.reduce((a, b) => a + b, 0) / valid.length : null;
  };
  const totalsAvg = {
    11: avgNonNull(rows.map(r => r.totalCSR > 0 ? r.matchCount / r.totalCSR : null)),
    16: avgNonNull(rows.map(r => (r.aiTagsHigh   || []).length > 0 ? r.matchCountHigh   / r.aiTagsHigh.length   : null)),
    21: avgNonNull(rows.map(r => (r.aiTagsMedium || []).length > 0 ? r.matchCountMedium / r.aiTagsMedium.length : null)),
    26: avgNonNull(rows.map(r => (r.aiTagsLow    || []).length > 0 ? r.matchCountLow    / r.aiTagsLow.length    : null)),
    38: avgNonNull(rows.map(r => r.totalCSRFields > 0 ? r.fieldMatchCount / r.totalCSRFields : null)),
    43: avgNonNull(rows.map(r => (r.aiFieldsHigh   || []).length > 0 ? r.fieldMatchCountHigh   / r.aiFieldsHigh.length   : null)),
    44: avgNonNull(rows.map(r => r.totalCSRFields  > 0               ? r.fieldMatchCountHigh   / r.totalCSRFields         : null)),
    49: avgNonNull(rows.map(r => (r.aiFieldsMedium || []).length > 0 ? r.fieldMatchCountMedium / r.aiFieldsMedium.length : null)),
    54: avgNonNull(rows.map(r => (r.aiFieldsLow    || []).length > 0 ? r.fieldMatchCountLow    / r.aiFieldsLow.length    : null)),
  };

  HEADERS.forEach((_, ci) => {
    const addr = `${col(ci)}${totalRow}`;
    if (ci === 0) {
      ws[addr] = { v: 'TOTALS', t: 's' };
    } else if (sumCols.has(ci)) {
      ws[addr] = { f: `SUM(${col(ci)}${dataStart}:${col(ci)}${dataEnd})`, t: 'n' };
    } else if (totalsAvg[ci] !== undefined) {
      const tv = totalsAvg[ci];
      ws[addr] = { v: tv ?? '', t: tv !== null ? 'n' : 's', z: '0.0%' };
    } else {
      ws[addr] = { v: '', t: 's' };
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
      const filteredCSRTags = flatTags.filter(t => taxonomyTagSet.has(t));

      const aiTaxonomyTags = [];
      const aiTagsByConf = { high: [], medium: [], low: [] };
      for (const [, info] of Object.entries(allResults)) {
        if (info.value === null || info.value === undefined) continue;
        const bucket = aiTagsByConf[info.confidence] || aiTagsByConf.low;
        if (typeof info.value === 'string' && taxonomyTagSet.has(info.value)) {
          aiTaxonomyTags.push(info.value);
          bucket.push(info.value);
        } else if (Array.isArray(info.value)) {
          const valid = info.value.filter(t => taxonomyTagSet.has(t));
          aiTaxonomyTags.push(...valid);
          bucket.push(...valid);
        }
      }

      // Deterministic tag mapping from metadata and custom field values
      const channelTagMap = { 'Phone': 'phone', 'Email': 'email', 'Chat': 'chat', 'Web Form': 'web_form' };
      const channelTag = channelTagMap[allResults.channel?.value];
      if (channelTag) { aiTaxonomyTags.push(channelTag); aiTagsByConf.high.push(channelTag); }
      if (allResults.tech_issue?.value === true)       { aiTaxonomyTags.push('00_tech_issues');      aiTagsByConf.high.push('00_tech_issues'); }
      if (allResults.services_issue?.value === true)   { aiTaxonomyTags.push('00_services_issue');   aiTagsByConf.high.push('00_services_issue'); }
      if (allResults.operations_issue?.value === true) { aiTaxonomyTags.push('00_operations_issue'); aiTagsByConf.high.push('00_operations_issue'); }
      for (const [fid, info] of Object.entries(allResults)) {
        if (!fieldDefs[fid]) continue;
        if (info.value === null || info.value === undefined || info.value === false) continue;
        const strVal = String(info.value);
        if (taxonomyTagSet.has(strVal) && !aiTaxonomyTags.includes(strVal)) {
          aiTaxonomyTags.push(strVal);
          aiTagsByConf.high.push(strVal);
        }
      }

      const aiCustomFields = [];
      const aiFieldsByConf = { high: [], medium: [], low: [] };
      for (const [fid, info] of Object.entries(allResults)) {
        if (!fieldDefs[fid]) continue;
        if (info.value === null || info.value === undefined || info.value === false) continue;
        const entry = { id: fid, value: String(info.value) };
        aiCustomFields.push(entry);
        (aiFieldsByConf[info.confidence] || aiFieldsByConf.low).push(entry);
      }

      const comparableCSRFields = csrCustomFields.filter(f => APP_MANAGED_FIELD_IDS.has(f.id));

      const tagComparison   = compareTags(filteredCSRTags, aiTaxonomyTags);
      const tagCompHigh     = compareTags(filteredCSRTags, aiTagsByConf.high);
      const tagCompMedium   = compareTags(filteredCSRTags, aiTagsByConf.medium);
      const tagCompLow      = compareTags(filteredCSRTags, aiTagsByConf.low);
      const fieldComparison = compareFields(comparableCSRFields, aiCustomFields);
      const fieldCompHigh   = compareFields(comparableCSRFields, aiFieldsByConf.high);
      const fieldCompMedium = compareFields(comparableCSRFields, aiFieldsByConf.medium);
      const fieldCompLow    = compareFields(comparableCSRFields, aiFieldsByConf.low);

      rows.push({
        ticketId:           id,
        csrTags:            flatTags,
        aiTaxonomyTags,
        matched:            tagComparison.matched,
        missing:            tagComparison.missing,
        extra:              tagComparison.extra,
        matchCount:         tagComparison.matchCount,
        missingCount:       tagComparison.missingCount,
        extraCount:         tagComparison.extraCount,
        totalCSR:           tagComparison.totalCSR,
        totalAI:            tagComparison.totalAI,
        aiTagsHigh:         aiTagsByConf.high,
        matchedHigh:        tagCompHigh.matched,
        missingHigh:        tagCompHigh.missing,
        extraHigh:          tagCompHigh.extra,
        matchCountHigh:     tagCompHigh.matchCount,
        missingCountHigh:   tagCompHigh.missingCount,
        extraCountHigh:     tagCompHigh.extraCount,
        aiTagsMedium:       aiTagsByConf.medium,
        matchedMedium:      tagCompMedium.matched,
        missingMedium:      tagCompMedium.missing,
        extraMedium:        tagCompMedium.extra,
        matchCountMedium:   tagCompMedium.matchCount,
        missingCountMedium: tagCompMedium.missingCount,
        extraCountMedium:   tagCompMedium.extraCount,
        aiTagsLow:          aiTagsByConf.low,
        matchedLow:         tagCompLow.matched,
        missingLow:         tagCompLow.missing,
        extraLow:           tagCompLow.extra,
        matchCountLow:      tagCompLow.matchCount,
        missingCountLow:    tagCompLow.missingCount,
        extraCountLow:      tagCompLow.extraCount,
        notes:              '',
        csrCustomFields,
        aiCustomFields,
        matchedFields:      fieldComparison.matchedFields,
        missingFields:      fieldComparison.missingFields,
        extraFields:        fieldComparison.extraFields,
        fieldMatchCount:    fieldComparison.fieldMatchCount,
        fieldMissingCount:  fieldComparison.fieldMissingCount,
        fieldExtraCount:    fieldComparison.fieldExtraCount,
        totalCSRFields:     fieldComparison.totalCSRFields,
        totalAIFields:      fieldComparison.totalAIFields,
        aiFieldsHigh:           aiFieldsByConf.high,
        matchedFieldsHigh:      fieldCompHigh.matchedFields,
        missingFieldsHigh:      fieldCompHigh.missingFields,
        extraFieldsHigh:        fieldCompHigh.extraFields,
        fieldMatchCountHigh:    fieldCompHigh.fieldMatchCount,
        aiFieldsMedium:         aiFieldsByConf.medium,
        matchedFieldsMedium:    fieldCompMedium.matchedFields,
        missingFieldsMedium:    fieldCompMedium.missingFields,
        extraFieldsMedium:      fieldCompMedium.extraFields,
        fieldMatchCountMedium:  fieldCompMedium.fieldMatchCount,
        aiFieldsLow:            aiFieldsByConf.low,
        matchedFieldsLow:       fieldCompLow.matchedFields,
        missingFieldsLow:       fieldCompLow.missingFields,
        extraFieldsLow:         fieldCompLow.extraFields,
        fieldMatchCountLow:     fieldCompLow.fieldMatchCount,
      });
    } catch (err) {
      console.log(`Error on ticket ${id}: ${err.message}`);
      rows.push({
        ticketId:           id,
        csrTags:            [],
        aiTaxonomyTags:     [],
        matched:            [],
        missing:            [],
        extra:              [],
        matchCount:         0,
        missingCount:       0,
        extraCount:         0,
        totalCSR:           0,
        totalAI:            0,
        aiTagsHigh:         [],
        matchedHigh:        [],
        missingHigh:        [],
        extraHigh:          [],
        matchCountHigh:     0,
        missingCountHigh:   0,
        extraCountHigh:     0,
        aiTagsMedium:       [],
        matchedMedium:      [],
        missingMedium:      [],
        extraMedium:        [],
        matchCountMedium:   0,
        missingCountMedium: 0,
        extraCountMedium:   0,
        aiTagsLow:          [],
        matchedLow:         [],
        missingLow:         [],
        extraLow:           [],
        matchCountLow:      0,
        missingCountLow:    0,
        extraCountLow:      0,
        notes:              `ERROR: ${err.message}`,
        csrCustomFields:    [],
        aiCustomFields:     [],
        matchedFields:      [],
        missingFields:      [],
        extraFields:        [],
        fieldMatchCount:    0,
        fieldMissingCount:  0,
        fieldExtraCount:    0,
        totalCSRFields:     0,
        totalAIFields:      0,
      });
    }

    if (i < ids.length - 1) await new Promise(r => setTimeout(r, 500));
  }

  buildExcel(rows);
  console.log('Done! Report saved to accuracy_report.xlsx');
}

main();
