/**
 * NinjaClaw settings UI controller.
 *
 * - Loads the schema + current values from /api/settings.
 * - Renders one section per category, with optional sub-groups.
 * - Tracks dirty state so we only POST changed fields.
 * - Sends `<clear>` for secrets the user explicitly cleared, and skips
 *   secrets the user didn't touch (so we never overwrite a real secret
 *   with the masked placeholder).
 */

const TOKEN_KEY = 'ninjaclaw_web_token';

const CATEGORY_META = {
  general:   { title: 'General',                          order: 1 },
  github:    { title: 'GitHub Models',                    order: 2 },
  web:       { title: 'Web UI',                           order: 3 },
  telegram:  { title: 'Telegram',                         order: 4 },
  teams:     { title: 'Microsoft Teams (legacy)',         order: 5 },
  agent365:  { title: 'Microsoft Agent 365 (preview)',    order: 6 },
  onecli:    { title: 'OneCLI Credential Gateway',        order: 7 },
};

const state = {
  token: localStorage.getItem(TOKEN_KEY) || '',
  schema: [],
  values: new Map(), // key -> { value, masked, isSet }
  dirty: new Map(),  // key -> new value
  envPath: '',
};

const els = {
  status:    document.getElementById('status'),
  authPrompt: document.getElementById('auth-prompt'),
  authInput:  document.getElementById('auth-token'),
  authBtn:    document.getElementById('btn-auth'),
  content:    document.getElementById('settings-content'),
  envPath:    document.getElementById('env-path'),
  categories: document.getElementById('categories'),
  saveBtn:    document.getElementById('btn-save'),
  reloadBtn:  document.getElementById('btn-reload'),
  result:     document.getElementById('save-result'),
};

async function fetchSettings() {
  const res = await fetch('/api/settings', {
    headers: state.token ? { Authorization: `Bearer ${state.token}` } : {},
  });
  if (res.status === 401) {
    showAuthPrompt();
    return;
  }
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data = await res.json();
  state.schema = data.schema || [];
  state.envPath = data.envPath || '';
  state.values = new Map((data.values || []).map((v) => [v.key, v]));
  state.dirty.clear();
  hideAuthPrompt();
  render();
  setStatus('Loaded', 'online');
}

function showAuthPrompt() {
  els.authPrompt.classList.remove('hidden');
  els.content.classList.add('hidden');
  setStatus('Authentication required', 'error');
}
function hideAuthPrompt() {
  els.authPrompt.classList.add('hidden');
  els.content.classList.remove('hidden');
}

function setStatus(text, cls = '') {
  els.status.textContent = text;
  els.status.className = `status ${cls}`;
}

function render() {
  els.envPath.textContent = state.envPath || '.env';
  els.categories.innerHTML = '';

  // Group by category, then by optional `group` field.
  const byCategory = new Map();
  for (const field of state.schema) {
    if (!byCategory.has(field.category)) byCategory.set(field.category, []);
    byCategory.get(field.category).push(field);
  }

  const ordered = Array.from(byCategory.keys()).sort((a, b) => {
    const oa = CATEGORY_META[a]?.order ?? 99;
    const ob = CATEGORY_META[b]?.order ?? 99;
    return oa - ob;
  });

  for (const category of ordered) {
    const fields = byCategory.get(category);
    els.categories.appendChild(renderCategory(category, fields));
  }
}

function renderCategory(category, fields) {
  const setCount = fields.filter((f) => state.values.get(f.key)?.isSet).length;
  const section = document.createElement('section');
  section.className = 'category';
  section.innerHTML = `
    <h2>
      <span>${escapeHtml(CATEGORY_META[category]?.title ?? category)}</span>
      <span class="badge">${setCount}/${fields.length} configured</span>
    </h2>
  `;

  // Group fields by .group
  const byGroup = new Map();
  for (const f of fields) {
    const g = f.group || '';
    if (!byGroup.has(g)) byGroup.set(g, []);
    byGroup.get(g).push(f);
  }

  for (const [groupName, groupFields] of byGroup.entries()) {
    const groupEl = document.createElement('div');
    groupEl.className = 'field-group';
    if (groupName) {
      const h3 = document.createElement('h3');
      h3.textContent = groupName;
      groupEl.appendChild(h3);
    }
    for (const field of groupFields) {
      groupEl.appendChild(renderField(field));
    }
    section.appendChild(groupEl);
  }

  return section;
}

function renderField(field) {
  const current = state.values.get(field.key) || { isSet: false };
  const wrapper = document.createElement('div');
  wrapper.className = 'field';

  const inputType = field.type === 'number' ? 'number'
    : field.type === 'url' ? 'url'
    : field.secret ? 'password' : 'text';

  const placeholderAttr = field.placeholder
    ? ` placeholder="${escapeHtml(field.placeholder)}"` : '';

  if (field.secret) {
    wrapper.innerHTML = `
      <label>
        <span>${escapeHtml(field.label)}</span>
        <span class="key">${escapeHtml(field.key)}</span>
      </label>
      <p class="description">${escapeHtml(field.description)}</p>
      <div class="secret-row">
        <input type="${inputType}" data-key="${escapeHtml(field.key)}" data-secret="1"${placeholderAttr}>
        <button type="button" class="clear-btn" data-clear="${escapeHtml(field.key)}">Clear</button>
      </div>
      <span class="current-state ${current.isSet ? 'set' : 'unset'}">
        ${current.isSet ? `Currently set (${escapeHtml(current.masked || '••••')}). Leave blank to keep.`
                        : 'Not set.'}
      </span>
    `;
  } else {
    wrapper.innerHTML = `
      <label>
        <span>${escapeHtml(field.label)}</span>
        <span class="key">${escapeHtml(field.key)}</span>
      </label>
      <p class="description">${escapeHtml(field.description)}</p>
      <input type="${inputType}" data-key="${escapeHtml(field.key)}"
             value="${escapeHtml(current.value ?? '')}"${placeholderAttr}>
    `;
  }

  const input = wrapper.querySelector('input[data-key]');
  input.addEventListener('input', () => {
    state.dirty.set(field.key, input.value);
    els.saveBtn.disabled = state.dirty.size === 0;
  });

  const clearBtn = wrapper.querySelector('button[data-clear]');
  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      state.dirty.set(field.key, '<clear>');
      input.value = '';
      input.placeholder = '(will be cleared on save)';
      els.saveBtn.disabled = false;
    });
  }

  return wrapper;
}

function escapeHtml(str) {
  return String(str ?? '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  })[c]);
}

async function saveSettings() {
  if (state.dirty.size === 0) return;
  els.saveBtn.disabled = true;
  els.result.classList.add('hidden');

  const updates = {};
  for (const [key, value] of state.dirty.entries()) {
    updates[key] = value;
  }

  try {
    const res = await fetch('/api/settings', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        ...(state.token ? { Authorization: `Bearer ${state.token}` } : {}),
      },
      body: JSON.stringify({ updates }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);
    showResult(
      `Saved ${data.written?.length ?? 0} setting(s). ${data.notice || ''}`,
      'success',
    );
    state.dirty.clear();
    await fetchSettings();
  } catch (err) {
    showResult(`Save failed: ${err.message}`, 'error');
    els.saveBtn.disabled = false;
  }
}

function showResult(text, cls) {
  els.result.textContent = text;
  els.result.className = `save-result ${cls}`;
  els.result.classList.remove('hidden');
}

els.authBtn.addEventListener('click', () => {
  state.token = els.authInput.value.trim();
  if (state.token) localStorage.setItem(TOKEN_KEY, state.token);
  fetchSettings().catch((err) => setStatus(err.message, 'error'));
});

els.saveBtn.addEventListener('click', saveSettings);
els.reloadBtn.addEventListener('click', () => fetchSettings().catch((err) => setStatus(err.message, 'error')));
els.saveBtn.disabled = true;

fetchSettings().catch((err) => setStatus(err.message, 'error'));
