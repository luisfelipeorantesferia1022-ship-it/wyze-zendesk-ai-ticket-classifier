require('dotenv').config();
const express = require('express');
const cors = require('cors');
const OpenAI = require('openai');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const fieldDefs = require('./data/field-definitions.json');
const conditionLookup = require('./data/condition-lookup.json');
const formConfig = require('./data/form-config.json');

// Client/company name is configurable so this code is reusable across engagements.
const COMPANY_NAME = process.env.COMPANY_NAME || 'CLIENT';

const TIMING_LOG_PATH = path.join(__dirname, 'data', 'timing-log.jsonl');

function loadTaxonomy() {
  const filePath = path.join(__dirname, '../data/taxonomy_data.xlsx');
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets['Field_options'];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  const tags = rows
    .map(row => row['Tag'])
    .filter(tag => tag && tag.toString().trim() !== '');

  return 'SUPPORT TICKET TAXONOMY TAGS:\n' + tags.join('\n');
}

const TAXONOMY = loadTaxonomy();


const MAX_PASSES = 7;
const SKIP_IDS = new Set(formConfig.skip_field_ids.map(String));

const app = express();
app.use(express.json());

// CORS 
// Allows the Zendesk sidebar (any origin) to reach this server.
app.use(cors());

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
- For field 45511186554395 (no_pm_response): only set to true if the customer EXPLICITLY states they contacted ${COMPANY_NAME} before AND received no reply. Both conditions must be present. Phrases like "I've been waiting", "I sent an email", "I called before", or "I've had this problem for a while" do NOT qualify on their own. The customer must clearly say something like "I reached out and never heard back" or "I emailed and got no response". If only one condition is present, return false. If you flag this as true, you MUST be able to quote the exact phrase where the customer says they were previously ignored.
- DEFECTIVE ITEM: Set field 45511052447003 ("OR - Defective Item Check") to true ONLY if the agent confirms the product is genuinely defective or broken and starts a warranty/replacement. If the transcript indicates the customer damaged the item themselves (dropped it, water damage, misuse), do NOT set this — that is a different field. Default false.
- HAPPINESS ADJUSTMENT: Set field 45511118507931 ("OR - Happiness adjustment") to true ONLY if the agent offers a concrete goodwill gesture — a refund, discount, free product, or replacement specifically to make a frustrated customer whole. Do not set it for routine warranty replacement of a defective unit (that is just standard service, not goodwill). Do not attempt to fill the action or reason sub-fields; those appear in a later pass once this box is checked. Default false.`.trim();
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

async function writeToZendesk(ticketId, fieldUpdates) {
  const { ZENDESK_SUBDOMAIN, ZENDESK_EMAIL, ZENDESK_API_TOKEN } = process.env;
  const url = `https://${ZENDESK_SUBDOMAIN}/api/v2/tickets/${ticketId}.json`;
  const credentials = Buffer.from(`${ZENDESK_EMAIL}/token:${ZENDESK_API_TOKEN}`).toString('base64');

  const ticketBody = {};

  if (fieldUpdates.standard) {
    const s = fieldUpdates.standard;
    if (s.subject)     ticketBody.subject  = s.subject;
    if (s.priority)    ticketBody.priority = s.priority;
    if (s.type)        ticketBody.type     = s.type;
    if (s.status)      ticketBody.status   = s.status;
    if (s.tags)        ticketBody.tags     = s.tags;
    if (s.description) ticketBody.comment  = { body: s.description, public: false };
  }

  if (fieldUpdates.custom && fieldUpdates.custom.length > 0) {
    ticketBody.custom_fields = fieldUpdates.custom.map(f => ({
      id: Number(f.fieldId),
      value: f.value,
    }));
  }

  console.log('Writing to Zendesk - ticketId:', ticketId, 'body:', JSON.stringify(ticketBody));

  const response = await fetch(url, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Basic ${credentials}`,
    },
    body: JSON.stringify({ ticket: ticketBody }),
  });

  const responseBody = await response.json().catch(() => ({}));
  console.log('Zendesk response status:', response.status);

  if (!response.ok) {
    throw new Error(`Zendesk API error ${response.status}: ${JSON.stringify(responseBody)}`);
  }

  return responseBody;
}

// Health check
app.get('/', (req, res) => {
  res.send(`${COMPANY_NAME} Ticket Assistant server is running.`);
});

// Write a single custom field to a Zendesk ticket
app.post('/write-field', async (req, res) => {
  const { ticketId, fieldId, value } = req.body;
  try {
    await writeToZendesk(ticketId, {
      custom: [{ fieldId, value }]
    });
    res.json({ success: true });
  } catch (error) {
    console.error('write-field error:', error.message);
    res.status(500).json({ error: error.message });
  }
});

app.post('/analyze-transcript-v2', async (req, res) => {
  const { transcript: rawTranscript, ticketId } = req.body;
  if (!rawTranscript) {
    return res.status(400).json({ error: 'No transcript provided.' });
  }
  const transcript = stripPII(rawTranscript);
  console.log('--- v2 endpoint called ---');

  const serverStartTime = Date.now();
  const perPassMs = [];

  try {
    const allResults = {};
    const filledValues = {};
    const seenFieldIds = new Set();

    let currentFieldIds = formConfig.always_visible_field_ids
      .map(String)
      .filter(id => !SKIP_IDS.has(id) && fieldDefs[id]);

    for (let pass = 0; pass < MAX_PASSES; pass++) {
      const passStart = Date.now();
      const newFieldIds = currentFieldIds.filter(id => !seenFieldIds.has(id));
      if (newFieldIds.length === 0) break;
      for (const id of newFieldIds) seenFieldIds.add(id);

      const systemPrompt = buildDynamicSystemPrompt(newFieldIds, filledValues, pass);
      console.log(`v2 Pass ${pass + 1}: processing ${newFieldIds.length} fields`);

      const response = await openai.chat.completions.create({
        model: 'gpt-4o',
        temperature: 0,
        response_format: { type: 'json_object' },
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: `Here is the tag taxonomy:\n\n${TAXONOMY}\n\nTranscript:\n${transcript}` },
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

      const META_FIELD_MAP = {
        first_name: '45511081093275',
        last_name:  '45511066969243',
      };
      for (const [metaKey, realId] of Object.entries(META_FIELD_MAP)) {
        if (passFields[metaKey] && passFields[metaKey].value != null && !passFields[realId]) {
          passFields[realId] = passFields[metaKey];
        }
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

      const standard = {};
      const custom = [];
      const STANDARD_KEYS = ['subject', 'priority', 'type', 'status', 'description', 'tags'];
      for (const [fid, info] of Object.entries(passFields)) {
        if (info.value === null || info.value === undefined) continue;
        if (STANDARD_KEYS.includes(fid)) {
          standard[fid] = info.value;
        } else if (fieldDefs[fid]) {
          custom.push({ fieldId: fid, value: info.value });
        }
        // keys with no fieldDefs entry (Pass 0 metadata) are intentionally skipped —
        // they are routing signals, not Zendesk fields.
      }
      if (ticketId) {
        if (Object.keys(standard).length > 0) await writeToZendesk(ticketId, { standard });
        if (custom.length > 0) await writeToZendesk(ticketId, { custom });
      }

      const triggeredChildren = resolveChildFields(filledValues);
      currentFieldIds = [...triggeredChildren.keys()].filter(
        id => !seenFieldIds.has(id) && !SKIP_IDS.has(id) && fieldDefs[id]
      );
      console.log(`v2 Pass ${pass + 1} complete. Triggered ${currentFieldIds.length} new fields.`);
      perPassMs.push(Date.now() - passStart);
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

    console.log(`v2 complete. Total fields: ${Object.keys(allResults).length}`);

    const serverDurationMs = Date.now() - serverStartTime;

    res.json({
      success: true,
      fields: allResults,
      metadata: {
        total_fields: Object.keys(allResults).length,
        passes: [...new Set(Object.values(allResults).map(f => f.pass))].length,
        duration_ms: serverDurationMs,
        per_pass_ms: perPassMs,
        transcript_length: transcript.length,
      },
    });

  } catch (error) {
    console.error('v2 error:', error.message);
    res.status(500).json({ error: 'Failed to analyze transcript.' });
  }
});

// Record a timing entry
app.post('/timing-data', (req, res) => {
  const record = {
    timestamp: new Date().toISOString(),
    ticket_id: req.body.ticket_id || null,
    total_duration_ms: req.body.total_duration_ms || 0,
    server_duration_ms: req.body.server_duration_ms || 0,
    per_pass_ms: req.body.per_pass_ms || [],
    num_passes: req.body.num_passes || 0,
    num_fields_filled: req.body.num_fields_filled || 0,
    transcript_length: req.body.transcript_length || 0,
  };
  try {
    fs.appendFileSync(TIMING_LOG_PATH, JSON.stringify(record) + '\n');
    res.json({ success: true });
  } catch (err) {
    console.error('timing-data write error:', err.message);
    res.status(500).json({ error: 'Failed to write timing data.' });
  }
});

// Export timing data as CSV
app.get('/timing-data', (req, res) => {
  const headers = ['timestamp', 'ticket_id', 'total_duration_ms', 'server_duration_ms', 'per_pass_ms', 'num_passes', 'num_fields_filled', 'transcript_length'];
  let records = [];
  try {
    if (fs.existsSync(TIMING_LOG_PATH)) {
      const lines = fs.readFileSync(TIMING_LOG_PATH, 'utf8').trim().split('\n').filter(Boolean);
      records = lines.map(line => JSON.parse(line));
    }
  } catch (err) {
    console.error('timing-data read error:', err.message);
  }
  const csvRows = [headers.join(',')];
  for (const r of records) {
    csvRows.push(headers.map(h => {
      const val = r[h];
      if (Array.isArray(val)) return `"${val.join(';')}"`;
      if (val === null || val === undefined) return '';
      return String(val).includes(',') ? `"${val}"` : String(val);
    }).join(','));
  }
  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', 'attachment; filename="timing-data.csv"');
  res.send(csvRows.join('\n'));
});

// Start
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});