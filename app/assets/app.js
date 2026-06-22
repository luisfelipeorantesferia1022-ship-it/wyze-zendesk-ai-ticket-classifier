/* Ticket Assistant — Zendesk Sidebar App
 *
 * State machine:  idle → loading → review → cleanup
 */

const SERVER_URL = 'http://localhost:3000';

const META_KEYS = new Set([
  'full_name', 'first_name', 'last_name', 'channel', 'wrong_number',
  'retention', 'tech_issue', 'services_issue', 'operations_issue'
]);

// Maps AI-returned field names → ZAF SDK ticket field paths.
// Extend this for custom Zendesk fields.
const ZAF_FIELD_MAP = {
  subject:     'ticket.subject',
  priority:    'ticket.priority',
  type:        'ticket.type',
  status:      'ticket.status',
  description: 'ticket.description',
  tags:        'ticket.tags',
};

// ── State ─────────────────────────────────────────────────────────────────────

let client = null;

// allFields: { [fieldName]: { value, confidence, reason } }
// value === null means AI could not fill it.
let allFields = {};

// stores [{ label, fieldId, value }, ...] for pass2/3/4 results
let backendFields = [];
let lastMetadata = null;

// ── Bootstrap ─────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  client = window.ZAFClient ? ZAFClient.init() : null;

  if (client) {
    client.invoke('resize', { width: '100%', height: '600px' });
  }

  document.getElementById('btn-fill').addEventListener('click', handleFill);
  document.getElementById('btn-confirm').addEventListener('click', handleConfirm);
  document.getElementById('btn-finish').addEventListener('click', handleFinish);
});

// ── View helpers ──────────────────────────────────────────────────────────────

function showView(name) {
  ['idle', 'loading', 'review', 'cleanup'].forEach(v => {
    document.getElementById(`view-${v}`).classList.toggle('active', v === name);
  });
  document.getElementById('footer-review').style.display  = name === 'review'  ? 'block' : 'none';
  document.getElementById('footer-cleanup').style.display = name === 'cleanup' ? 'block' : 'none';
  clearError();
}

function setLoadingLabel(text) {
  document.getElementById('loading-label').textContent = text;
}

function showError(msg) {
  document.getElementById('error-text').textContent = msg;
  document.getElementById('error-bar').classList.add('visible');
}

function clearError() {
  document.getElementById('error-bar').classList.remove('visible');
}

// ── Step 1: Fill out ticket (idle → loading → review) ────────────────────────

async function handleFill() {
  showView('loading');
  clearError();

  try {
    const fillStart = Date.now();
    // 1. Fetch transcript from Zendesk ticket conversation
    setLoadingLabel('Fetching conversation…');
    const transcript = await fetchTranscript();

    if (!transcript) {
      showView('idle');
      showError('No conversation found on this ticket.');
      return;
    }

    // 2. Send to backend for AI analysis
    setLoadingLabel('Analyzing with AI…');
    const ticketData = await client?.get('ticket.id');
    const ticketId = ticketData?.['ticket.id'] || null;
    const fields = await analyzeTranscript(transcript, ticketId);

    // 3. Store and render review panel
    allFields = fields;
    renderReview(fields);
    await applyBackendFieldsToForm(fields);   // populate live form in-place
    showView('review');
    logTiming(ticketId, Date.now() - fillStart);

  } catch (err) {
    showView('idle');
    showError(err.message);
  }
}

// ── Fetch ticket conversation via ZAF ─────────────────────────────────────────

async function fetchTranscript() {
  if (!client) {
    return '[DEV MODE] Agent: Hello, how can I help?\nCustomer: My Product A camera has been offline since the firmware update yesterday. I tried rebooting it but it still won\'t connect.';
  }

  const ticketData = await client.get('ticket.id');
  const ticketId = ticketData['ticket.id'];

  const commentsRes = await client.request({
    url: `/api/v2/tickets/${ticketId}/comments.json`,
    type: 'GET',
  });

  const comments = commentsRes.comments;
  if (!comments || comments.length === 0) return null;

  return comments
    .map(c => {
      const role = c.via?.channel === 'api' || c.public === false ? 'Agent' : 'Customer';
      let body = c.plain_body || c.body || '';
      body = body.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/\s+/g, ' ').trim();
      return `${role}: ${body}`;
    })
    .join('\n\n');
}

// ── POST to backend ────────────────────────────────────────────────────────────

async function analyzeTranscript(transcript, ticketId) {
  const res = await fetch(`${SERVER_URL}/analyze-transcript-v2`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ transcript: stripPII(transcript), ticketId }),
  });

  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body.error || `Server error ${res.status}`);
  }

  const data = await res.json();
  if (!data.success || !data.fields) throw new Error('Unexpected response from server.');

  lastMetadata = data.metadata || null;

  backendFields = Object.entries(data.fields)
    .filter(([, info]) => info.value !== null && info.value !== undefined)
    .map(([fieldId, info]) => ({
      label: info.field_title || fieldId,
      fieldId,
      value: info.value,
    }));

  return data.fields;
}

// ── Review filter state ───────────────────────────────────────────────────────

let currentFilter = 'all';

function setFilter(f) {
  currentFilter = f;
  ['all', 'high', 'medium', 'empty'].forEach(p => {
    document.getElementById('fpill-' + p).classList.toggle('active', p === f);
  });
  applyReviewFilter();
}

function applyReviewFilter() {
  let anyVisible = false;
  document.querySelectorAll('.rv-section').forEach(sec => {
    const items = sec.querySelectorAll('.field-item');
    let visible = 0;
    items.forEach(item => {
      const match = currentFilter === 'all' || item.dataset.conf === currentFilter;
      item.classList.toggle('fitem-hidden', !match);
      if (match) visible++;
    });
    sec.classList.toggle('sec-hidden', visible === 0);
    if (visible > 0) anyVisible = true;
  });
  document.getElementById('review-no-results').style.display = anyVisible ? 'none' : 'block';
}

// ── Render review panel ───────────────────────────────────────────────────────

// Maps field names to their display section key
const SECTION_ORDER = [
  { key: 'standard',  label: 'Standard fields',        match: name => ['subject','priority','type','status','description','tags'].includes(name) },
  { key: 'chk00',     label: '00 — Checkboxes',         match: (name, info) => typeof (info.value) === 'boolean' || String(info.field_title || name).startsWith('00') },
  { key: 'prod01',    label: '01 — Product list',       match: (name, info) => String(info.field_title || name).startsWith('01') },
  { key: 'cat02',     label: '02 — Category',           match: (name, info) => String(info.field_title || name).startsWith('02') },
  { key: 'subcat03',  label: '03 — Sub-category',       match: (name, info) => String(info.field_title || name).startsWith('03') },
  { key: 'ti',        label: 'TI — Technical issue',    match: (name, info) => String(info.field_title || name).toUpperCase().startsWith('TI') },
  { key: 'other',     label: 'Other',                   match: () => true },
];

function renderReview(fields) {
  currentFilter = 'all';
  ['all','high','medium','empty'].forEach(p => {
    const pill = document.getElementById('fpill-' + p);
    if (pill) pill.classList.toggle('active', p === 'all');
  });

  const container = document.getElementById('review-sections');
  // Remove old sections but keep the no-results div
  container.querySelectorAll('.rv-section').forEach(s => s.remove());

  // Group fields into buckets
  const buckets = {};
  SECTION_ORDER.forEach(s => { buckets[s.key] = []; });

  Object.entries(fields).forEach(([name, info]) => {
    if (META_KEYS.has(name)) return;   // internal signals — not shown in panel
    // Hide false booleans — don't show unchecked checkboxes
    if (info.value === false) return;

    for (const sec of SECTION_ORDER) {
      if (sec.match(name, info)) {
        buckets[sec.key].push([name, info]);
        break;
      }
    }
  });

  // Update filter pill counts
  let counts = { high: 0, medium: 0, empty: 0 };
  Object.values(fields).forEach(info => {
    if (info.value === false) return;
    const conf = (info.confidence || '').toLowerCase();
    if (info.value === null || info.value === undefined) counts.empty++;
    else if (conf === 'high') counts.high++;
    else if (conf === 'medium') counts.medium++;
  });
  document.getElementById('fpill-high').textContent   = `${counts.high} High`;
  document.getElementById('fpill-medium').textContent = `${counts.medium} Medium`;
  document.getElementById('fpill-empty').textContent  = `${counts.empty} Empty`;

  const noResults = document.getElementById('review-no-results');

  SECTION_ORDER.forEach(sec => {
    const entries = buckets[sec.key];
    if (entries.length === 0) return;

    const section = el('div', 'rv-section');
    section.id = 'rv-sec-' + sec.key;

    // Header
    const hdr = el('div', 'rv-section-header');
    hdr.onclick = () => toggleRvSection(sec.key);
    const left = el('div', 'rv-section-header-left');
    left.appendChild(el('span', 'rv-section-name', sec.label));
    const filled = entries.filter(([, i]) => i.value !== null && i.value !== undefined).length;
    const empty  = entries.length - filled;
    const countText = empty > 0 ? `${filled} filled · ${empty} empty` : `${filled} field${filled !== 1 ? 's' : ''}`;
    left.appendChild(el('span', 'rv-section-badge', countText));
    hdr.appendChild(left);
    const chev = el('span', 'rv-chevron open', '▾');
    chev.id = 'rv-chev-' + sec.key;
    hdr.appendChild(chev);
    section.appendChild(hdr);

    // Body
    const body = el('div', 'rv-section-body');
    body.id = 'rv-body-' + sec.key;

    entries.forEach(([name, info]) => {
      const isEmpty = info.value === null || info.value === undefined;
      const conf = isEmpty ? 'empty' : (info.confidence || 'low').toLowerCase();

      const item = el('div', 'field-item');
      item.dataset.conf = conf;

      if (isEmpty) {
        item.appendChild(makeZdUnfilledRow(name, info));
      } else if (typeof info.value === 'boolean') {
        item.appendChild(makeZdCheckboxRow(name, info));
      } else if (name === 'tags' || (Array.isArray(info.value))) {
        item.appendChild(makeZdTagsRow(name, info));
      } else {
        item.appendChild(makeZdInputRow(name, info));
      }

      body.appendChild(item);
    });

    section.appendChild(body);
    container.insertBefore(section, noResults);
  });

  updateFooterNote();
}

// ── Zendesk-style field row builders ─────────────────────────────────────────

function makeZdInputRow(name, info) {
  const conf = (info.confidence || 'low').toLowerCase();
  const row = el('div', 'field-row');

  const cb = el('input');
  cb.type = 'checkbox'; cb.className = 'field-checkbox'; cb.checked = true;
  cb.dataset.field = name;
  if (info._isBackend) { cb.dataset.fieldtype = 'backend'; cb.dataset.fieldid = info._fieldId; cb.dataset.fieldvalue = info.value; }
  cb.addEventListener('change', function() { toggleFieldRow(row, this); });

  const main = el('div', 'field-main');
  main.appendChild(el('span', 'zd-label', info.field_title || label(name)));
  const input = el('div', 'zd-input ai-filled zd-dropdown', String(info.value));
  main.appendChild(input);
  main.appendChild(makeConfRow(conf));

  row.appendChild(cb);
  row.appendChild(main);
  return row;
}

function makeZdTagsRow(name, info) {
  const conf = (info.confidence || 'low').toLowerCase();
  const row = el('div', 'field-row');

  const cb = el('input');
  cb.type = 'checkbox'; cb.className = 'field-checkbox'; cb.checked = true;
  cb.dataset.field = name;
  cb.addEventListener('change', function() { toggleFieldRow(row, this); });

  const main = el('div', 'field-main');
  main.appendChild(el('span', 'zd-label', info.field_title || label(name)));

  const box = el('div', 'zd-tags-box');
  const tags = Array.isArray(info.value) ? info.value : [String(info.value)];
  tags.forEach(t => box.appendChild(el('span', 'zd-tag', t)));
  main.appendChild(box);
  main.appendChild(makeConfRow(conf));

  row.appendChild(cb);
  row.appendChild(main);
  return row;
}

function makeZdCheckboxRow(name, info) {
  const conf = (info.confidence || 'low').toLowerCase();
  const row = el('div', 'field-row');

  const cb = el('input');
  cb.type = 'checkbox'; cb.className = 'field-checkbox'; cb.checked = true;
  cb.dataset.field = name;
  cb.addEventListener('change', function() { toggleFieldRow(row, this); });

  const main = el('div', 'field-main');
  const cbRow = el('div', 'zd-cb-row');
  const nativeCb = el('input');
  nativeCb.type = 'checkbox'; nativeCb.className = 'zd-cb-native';
  nativeCb.checked = !!info.value;
  cbRow.appendChild(nativeCb);
  cbRow.appendChild(el('span', 'zd-cb-label', info.field_title || label(name)));
  main.appendChild(cbRow);
  main.appendChild(makeConfRow(conf));

  row.appendChild(cb);
  row.appendChild(main);
  return row;
}

function makeZdUnfilledRow(name, info) {
  const row = el('div', 'field-row');
  const spacer = el('div', 'field-spacer');
  const main = el('div', 'field-main');
  main.appendChild(el('span', 'zd-label', info.field_title || label(name)));
  const input = el('div', 'zd-input zd-unfilled zd-dropdown', 'Empty');
  main.appendChild(input);
  main.appendChild(el('span', 'unfilled-reason', info.reason || 'Not found in transcript'));
  row.appendChild(spacer);
  row.appendChild(main);
  return row;
}

function makeConfRow(conf) {
  const row = el('div', 'conf-row');
  const dot = el('span', `conf-dot ${conf}`);
  const lbl = el('span', 'conf-lbl', 'Confidence:');
  const val = el('span', `conf-${conf}`, capitalize(conf));
  row.appendChild(dot); row.appendChild(lbl); row.appendChild(val);
  return row;
}

function toggleFieldRow(row, cb) {
  row.classList.toggle('unchecked', !cb.checked);
  updateFooterNote();
}

function updateFooterNote() {
  const all     = document.querySelectorAll('#review-sections .field-checkbox');
  const checked = [...all].filter(c => c.checked).length;
  const note    = document.getElementById('footer-field-note');
  if (note) note.textContent = `${checked} of ${all.length} fields will be written to ticket`;
}

// ── Section collapse toggle ───────────────────────────────────────────────────

function toggleRvSection(key) {
  const body = document.getElementById('rv-body-' + key);
  const chev = document.getElementById('rv-chev-' + key);
  const collapsed = body.classList.toggle('collapsed');
  chev.classList.toggle('open', !collapsed);
}

// ── Step 2: Confirm (review → cleanup) ───────────────────────────────────────

async function handleConfirm() {
  const btn = document.getElementById('btn-confirm');
  btn.disabled = true;

  clearError();

  // Collect checkbox states from new section structure
  const checkboxes = document.querySelectorAll('#review-sections .field-checkbox');
  const checked   = [];
  const unchecked = [];

  checkboxes.forEach(cb => {
    const name = cb.dataset.field;
    (cb.checked ? checked : unchecked).push(name);
  });

  // Apply checked fields to ticket
  if (checked.length > 0) {
    try {
      await applyFields(checked);
    } catch (err) {
      showError('Could not apply some fields: ' + err.message);
      btn.disabled = false;
      return;
    }
  }

  // Write checked backend fields to Zendesk via /write-field
  let ticketId = null;
  if (client) {
    const tidData = await client.get('ticket.id');
    ticketId = tidData['ticket.id'];
  }

  const backendCheckboxes = document.querySelectorAll('#review-sections .field-checkbox[data-fieldtype="backend"]');
  for (const cb of backendCheckboxes) {
    if (cb.checked) {
      await writeBackendField(ticketId, cb.dataset.fieldid, cb.dataset.fieldvalue);
    }
  }

  // Build cleanup panel and show it
  renderCleanup(unchecked);
  showView('cleanup');
  btn.disabled = false;
}

async function applyFields(fieldNames) {
  if (!client) return; // dev mode: skip

  const updates = {};
  fieldNames.forEach(name => {
    const path = ZAF_FIELD_MAP[name.toLowerCase()];
    const val  = allFields[name]?.value;
    if (path && val !== null && val !== undefined) {
      updates[path] = val;
    }
  });

  if (Object.keys(updates).length > 0) {
    await client.set(updates);
  }
}

async function applyBackendFieldsToForm(fields) {
  if (!client) return; // dev mode
  for (const [key, info] of Object.entries(fields)) {
    if (META_KEYS.has(key)) continue;        // skip routing signals (first_name, etc.)
    if (!/^\d+$/.test(key)) continue;        // only real numeric Zendesk field IDs
    const val = info.value;
    if (val === null || val === undefined || val === false) continue;
    try {
      await client.set('ticket.customField:custom_field_' + key, val);
    } catch (e) {
      console.warn('Could not set field ' + key + ':', e.message);
    }
  }
}

// ── Render cleanup panel ──────────────────────────────────────────────────────

function renderCleanup(uncheckedNames) {
  const unfilledNames = Object.entries(allFields)
    .filter(([, f]) => f.value === null || f.value === undefined)
    .map(([name]) => name);

  // ── Unchecked section ──
  const uncheckedSec  = document.getElementById('section-unchecked');
  const uncheckedBody = document.getElementById('body-unchecked');
  const uncheckedCnt  = document.getElementById('count-unchecked');

  uncheckedBody.innerHTML = '';

  if (uncheckedNames.length > 0) {
    uncheckedSec.style.display = 'block';
    uncheckedCnt.textContent = uncheckedNames.length;

    uncheckedNames.forEach(name => {
      const info = allFields[name];
      uncheckedBody.appendChild(makeCleanupUnchecked(name, info));
    });
  } else {
    uncheckedSec.style.display = 'none';
  }

  // ── Unfilled section ──
  const unfilledSec  = document.getElementById('section-unfilled');
  const unfilledBody = document.getElementById('body-unfilled');
  const unfilledCnt  = document.getElementById('count-unfilled');

  unfilledBody.innerHTML = '';

  if (unfilledNames.length > 0) {
    unfilledSec.style.display = 'block';
    unfilledCnt.textContent = unfilledNames.length;

    unfilledNames.forEach(name => {
      const info = allFields[name];
      unfilledBody.appendChild(makeCleanupUnfilled(name, info));
    });
  } else {
    unfilledSec.style.display = 'none';
  }

  // If nothing to show
  const isEmpty = uncheckedNames.length === 0 && unfilledNames.length === 0;
  document.getElementById('cleanup-empty').style.display = isEmpty ? 'block' : 'none';
}

function makeCleanupUnchecked(name, info) {
  const row = el('div', 'cleanup-row');

  const infoDiv = el('div', 'cleanup-info');
  infoDiv.appendChild(el('span', 'cleanup-name', label(name)));
  infoDiv.appendChild(el('span', 'cleanup-suggestion', `"${info.value}"`));

  const actions = el('div', 'cleanup-actions');
  const acceptBtn = el('button', 'btn-sm btn-accept', 'Accept');

  acceptBtn.addEventListener('click', async () => {
    try {
      await applyFields([name]);
      acceptBtn.textContent = 'Applied';
      acceptBtn.classList.add('done');
      acceptBtn.disabled = true;
    } catch (err) {
      showError('Could not apply: ' + err.message);
    }
  });

  actions.appendChild(acceptBtn);
  row.appendChild(infoDiv);
  row.appendChild(actions);
  return row;
}

function makeCleanupUnfilled(name, info) {
  const row = el('div', 'cleanup-row');

  const infoDiv = el('div', 'cleanup-info');
  infoDiv.appendChild(el('span', 'cleanup-name', label(name)));
  infoDiv.appendChild(el('span', 'cleanup-reason', info.reason || 'Not found in transcript'));

  const actions = el('div', 'cleanup-actions');
  const jumpBtn = el('button', 'btn-sm btn-jump', 'Jump to field');

  jumpBtn.addEventListener('click', () => jumpToField(name));

  actions.appendChild(jumpBtn);
  row.appendChild(infoDiv);
  row.appendChild(actions);
  return row;
}

// ── Jump to field ─────────────────────────────────────────────────────────────

function jumpToField(fieldName) {
  if (!client) {
    showError(`Fill in "${label(fieldName)}" manually.`);
    return;
  }
  // ZAF does not expose a universal field-focus API; we highlight it via a
  // notification so the agent knows which field to fill.
  client.invoke('notify', `Please fill in: ${label(fieldName)}`, 'notice', 3000).catch(() => {
    showError(`Please fill in "${label(fieldName)}" manually.`);
  });
}

// ── Step 3: Finish ────────────────────────────────────────────────────────────

function handleFinish() {
  // Reset to idle so the panel can be reused
  allFields = {};
  showView('idle');
}




async function logTiming(ticketId, totalDurationMs) {
  if (!lastMetadata) return;
  const numFilled = Object.values(allFields).filter(
    f => f.value !== null && f.value !== undefined && f.value !== false
  ).length;
  try {
    await fetch(`${SERVER_URL}/timing-data`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ticket_id: ticketId,
        total_duration_ms: totalDurationMs,
        server_duration_ms: lastMetadata.duration_ms || 0,
        per_pass_ms: lastMetadata.per_pass_ms || [],
        num_passes: lastMetadata.passes || 0,
        num_fields_filled: numFilled,
        transcript_length: lastMetadata.transcript_length || 0,
      }),
    });
  } catch (err) {
    console.warn('Timing log failed (non-blocking):', err.message);
  }
}

async function writeBackendField(ticketId, fieldId, value) {
  if (!ticketId) return;
  await fetch(`${SERVER_URL}/write-field`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ticketId, fieldId: Number(fieldId), value }),
  });
}

// ── Small DOM utilities ───────────────────────────────────────────────────────

function el(tag, className = '', text = '') {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text)      node.textContent = text;
  return node;
}

function label(name) {
  return name.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
}

function capitalize(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function stripPII(text) {
  return text
    .replace(/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g, '[EMAIL]')
    .replace(/\b(\+1[-.\s]?)?(\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})\b/g, '[PHONE]')
    .replace(/\b\d+\s+\w+(\s+\w+){0,3}\s+(St|Ave|Rd|Blvd|Dr|Ln|Way|Ct|Pl|Street|Avenue|Road|Boulevard|Drive|Lane|Court|Place)\.?\b/gi, '[ADDRESS]');
}
