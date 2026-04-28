/**
 * NinjaClaw Agent 365 admin console controller.
 *
 * Three flows:
 *   1. Sign-in: starts an MSAL device-code flow on the server, surfaces
 *      the user code + verification URL, and polls until success.
 *   2. Status: shows whether an admin account is cached.
 *   3. Publish: dry-run or live POST of agent365/blueprint.json to the
 *      admin plane using the cached admin token.
 */

const TOKEN_KEY = 'ninjaclaw_web_token';

const state = {
  token: localStorage.getItem(TOKEN_KEY) || '',
  account: null,
  flow: null,
  pollTimer: null,
  settings: null,
};

const els = {
  status:        document.getElementById('status'),
  authPrompt:    document.getElementById('auth-prompt'),
  authInput:     document.getElementById('auth-token'),
  authBtn:       document.getElementById('btn-auth'),
  content:       document.getElementById('content'),
  signedOut:     document.getElementById('signed-out'),
  signedIn:      document.getElementById('signed-in'),
  signInBtn:     document.getElementById('btn-signin'),
  signOutBtn:    document.getElementById('btn-signout'),
  username:      document.getElementById('account-username'),
  tenantSpan:    document.getElementById('account-tenant'),
  deviceFlow:    document.getElementById('device-flow'),
  deviceUri:     document.getElementById('device-uri'),
  deviceCode:    document.getElementById('device-code'),
  deviceStatus:  document.getElementById('device-status'),
  copyBtn:       document.getElementById('btn-copy-code'),
  deployList:    document.getElementById('deployment-context'),
  dryRunBtn:     document.getElementById('btn-dry-run'),
  publishBtn:    document.getElementById('btn-publish'),
  publishResult: document.getElementById('publish-result'),
  // Instances
  instancesList: document.getElementById('instances-list'),
  ciDisplay:     document.getElementById('ci-display'),
  ciHosting:     document.getElementById('ci-hosting'),
  ciEndpoint:    document.getElementById('ci-endpoint'),
  ciTenant:      document.getElementById('ci-tenant'),
  ciSubscription:document.getElementById('ci-subscription'),
  ciRg:          document.getElementById('ci-rg'),
  ciLocation:    document.getElementById('ci-location'),
  ciBotname:     document.getElementById('ci-botname'),
  ciActive:      document.getElementById('ci-active'),
  bfFields:      document.getElementById('bf-fields'),
  createBtn:     document.getElementById('btn-create-instance'),
  createResult:  document.getElementById('create-result'),
};

function authHeaders() {
  return state.token ? { Authorization: `Bearer ${state.token}` } : {};
}

async function api(method, url, body) {
  const init = { method, headers: { ...authHeaders() } };
  if (body !== undefined) {
    init.headers['Content-Type'] = 'application/json';
    init.body = JSON.stringify(body);
  }
  const res = await fetch(url, init);
  if (res.status === 401) {
    showAuthPrompt();
    throw new Error('unauthorized');
  }
  const text = await res.text();
  let data;
  try { data = text ? JSON.parse(text) : {}; } catch { data = { raw: text }; }
  if (!res.ok) {
    const err = new Error(data.error || `HTTP ${res.status}`);
    err.payload = data;
    throw err;
  }
  return data;
}

function setStatus(text, cls = '') {
  els.status.textContent = text;
  els.status.className = `status ${cls}`;
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

function prefillInstanceForm() {
  // Auto-fill messaging endpoint from the page's own origin when served
  // behind a public tunnel (e.g. Cloudflare). Only prefill if empty so we
  // don't clobber a manual override.
  try {
    if (els.ciEndpoint && !els.ciEndpoint.value) {
      const origin = window.location.origin;
      if (/^https:\/\//.test(origin)) {
        els.ciEndpoint.value = `${origin}/api/messages`;
      }
    }
  } catch {}
  // Default display name + location hints.
  if (els.ciDisplay && !els.ciDisplay.value) {
    els.ciDisplay.value = 'NinjaClaw-Nano';
  }
  if (els.ciLocation && !els.ciLocation.value) {
    els.ciLocation.value = 'global';
  }
}

async function bootstrap() {
  try {
    const [account, settings] = await Promise.all([
      api('GET', '/api/agent365/account'),
      api('GET', '/api/settings'),
    ]);
    state.account = account;
    state.settings = settings;
    prefillInstanceForm();
    hideAuthPrompt();
    renderAccount();
    renderDeploymentContext();
    await loadInstances();
    setStatus('Ready', 'online');
  } catch (err) {
    if (err.message !== 'unauthorized') {
      setStatus(err.message, 'error');
    }
  }
}

async function loadInstances() {
  try {
    const data = await api('GET', '/api/agent365/instances');
    renderInstances(data);
  } catch (err) {
    els.instancesList.innerHTML =
      `<p class="muted">Failed to load: ${escapeHtml(err.message)}</p>`;
  }
}

function renderInstances(data) {
  const { instances = [], activeInstanceId } = data;
  if (!instances.length) {
    els.instancesList.innerHTML =
      '<p class="muted">No instances yet. Create one below.</p>';
    return;
  }
  els.instancesList.innerHTML = '';
  for (const inst of instances) {
    const row = document.createElement('div');
    row.className = `instance-row${inst.isActive ? ' active' : ''}`;
    row.innerHTML = `
      <div>
        <div class="name">${escapeHtml(inst.displayName)}
          ${inst.isActive ? '<span class="badge">active</span>' : ''}
        </div>
        <div class="meta">
          <span>id: ${escapeHtml(inst.id)}</span>
          <span>app: ${escapeHtml(inst.clientId)}</span>
          <span>mode: ${escapeHtml(inst.hostingMode)}</span>
          ${inst.botName ? `<span>bot: ${escapeHtml(inst.botName)}</span>` : ''}
          ${inst.messagingEndpoint ? `<span>endpoint: ${escapeHtml(inst.messagingEndpoint)}</span>` : ''}
        </div>
      </div>
      <div class="actions">
        ${inst.isActive ? '' : `<button data-action="activate" data-id="${inst.id}">Activate</button>`}
        <button data-action="delete" data-id="${inst.id}" class="danger">Delete</button>
        <button data-action="delete-cascade" data-id="${inst.id}" class="danger" title="Also delete Entra app + Bot resource">Delete (cascade)</button>
      </div>
    `;
    els.instancesList.appendChild(row);
  }
}

els.instancesList.addEventListener('click', async (e) => {
  const btn = e.target.closest('button[data-action]');
  if (!btn) return;
  const { action, id } = btn.dataset;
  btn.disabled = true;
  try {
    if (action === 'activate') {
      await api('POST', `/api/agent365/instances/${id}/activate`);
    } else if (action === 'delete') {
      if (!confirm('Remove this instance from the local store? Entra app + Bot resource will remain.')) return;
      await api('DELETE', `/api/agent365/instances/${id}`);
    } else if (action === 'delete-cascade') {
      if (!confirm('Delete this instance AND its Entra app + Bot Framework resource? This cannot be undone.')) return;
      await api('DELETE', `/api/agent365/instances/${id}?cascade=true`);
    }
    await loadInstances();
    state.settings = await api('GET', '/api/settings');
    renderDeploymentContext();
  } catch (err) {
    alert(`Action failed: ${err.message}`);
  } finally {
    btn.disabled = false;
  }
});

els.ciHosting.addEventListener('change', () => {
  els.bfFields.style.display =
    els.ciHosting.value === 'bot-framework' ? '' : 'none';
});

els.createBtn.addEventListener('click', async () => {
  const body = {
    displayName: els.ciDisplay.value.trim(),
    hostingMode: els.ciHosting.value,
    messagingEndpoint: els.ciEndpoint.value.trim() || undefined,
    tenantId: els.ciTenant.value.trim() || undefined,
    setActive: els.ciActive.checked,
  };
  if (!body.displayName) { alert('Display name is required'); return; }
  if (body.hostingMode === 'bot-framework') {
    const sub = els.ciSubscription.value.trim();
    const rg = els.ciRg.value.trim();
    if (sub && rg) {
      body.botFramework = {
        subscriptionId: sub,
        resourceGroup: rg,
        location: els.ciLocation.value.trim() || 'global',
        botName: els.ciBotname.value.trim() || undefined,
      };
    }
  }

  els.createResult.classList.remove('hidden', 'error', 'success');
  els.createResult.textContent = 'Creating instance…';
  els.createBtn.disabled = true;
  try {
    const result = await api('POST', '/api/agent365/instances', body);
    els.createResult.classList.add('success');
    els.createResult.textContent = JSON.stringify(result, null, 2);
    await loadInstances();
    state.settings = await api('GET', '/api/settings');
    renderDeploymentContext();
  } catch (err) {
    els.createResult.classList.add('error');
    els.createResult.textContent =
      `${err.message}\n\n${JSON.stringify(err.payload || {}, null, 2)}`;
  } finally {
    els.createBtn.disabled = false;
  }
});

function renderAccount() {
  if (state.account?.signedIn) {
    els.signedIn.classList.remove('hidden');
    els.signedOut.classList.add('hidden');
    els.username.textContent = state.account.username || '(unknown)';
    els.tenantSpan.textContent = state.account.tenantId
      ? `(tenant ${state.account.tenantId.slice(0, 8)}…)` : '';
    els.publishBtn.disabled = false;
  } else {
    els.signedIn.classList.add('hidden');
    els.signedOut.classList.remove('hidden');
    els.publishBtn.disabled = true;
  }
}

function renderDeploymentContext() {
  const wanted = [
    'AGENT365_CLIENT_ID',
    'AGENT365_TENANT_ID',
    'AGENT365_OBJECT_ID',
    'AGENT365_HOSTING_MODE',
    'AGENT365_MESSAGING_ENDPOINT',
    'AGENT365_BOT_NAME',
    'AGENT365_PUBLISH_API',
    'AGENT365_BLUEPRINT_ID',
  ];
  const byKey = new Map((state.settings.values || []).map((v) => [v.key, v]));
  els.deployList.innerHTML = '';
  for (const key of wanted) {
    const v = byKey.get(key) || { isSet: false };
    const li = document.createElement('li');
    const valueDisplay = v.isSet
      ? (v.value ?? v.masked ?? '(set)')
      : '(not set)';
    li.innerHTML = `
      <span class="key">${key}</span>
      <span class="val ${v.isSet ? '' : 'unset'}">${escapeHtml(valueDisplay)}</span>
    `;
    els.deployList.appendChild(li);
  }
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  })[c]);
}

async function startSignIn() {
  els.signInBtn.disabled = true;
  els.deviceFlow.classList.remove('hidden');
  els.deviceStatus.textContent = 'Requesting code…';
  try {
    const flow = await api('POST', '/api/agent365/signin');
    state.flow = flow;
    els.deviceUri.textContent = flow.verificationUri;
    els.deviceUri.href = flow.verificationUri;
    els.deviceCode.textContent = flow.userCode;
    els.deviceStatus.textContent = 'Waiting for sign-in (poll every 3s)…';
    pollSignIn();
  } catch (err) {
    els.deviceStatus.textContent = `Error: ${err.message}`;
    els.signInBtn.disabled = false;
  }
}

function pollSignIn() {
  if (state.pollTimer) clearTimeout(state.pollTimer);
  state.pollTimer = setTimeout(async () => {
    try {
      const status = await api('GET', `/api/agent365/signin/${state.flow.flowId}`);
      if (status.status === 'success') {
        els.deviceStatus.textContent = `Signed in as ${status.username}`;
        await bootstrap();
        els.deviceFlow.classList.add('hidden');
        els.signInBtn.disabled = false;
      } else if (status.status === 'error') {
        els.deviceStatus.textContent = `Sign-in failed: ${status.error || 'unknown'}`;
        els.signInBtn.disabled = false;
      } else {
        pollSignIn();
      }
    } catch (err) {
      els.deviceStatus.textContent = `Poll error: ${err.message}`;
      els.signInBtn.disabled = false;
    }
  }, 3000);
}

async function doSignOut() {
  await api('POST', '/api/agent365/signout');
  await bootstrap();
}

async function publish(dryRun) {
  els.publishResult.classList.remove('hidden', 'error', 'success');
  els.publishResult.textContent = dryRun ? 'Building dry-run payload…' : 'Publishing…';
  els.publishBtn.disabled = true;
  els.dryRunBtn.disabled = true;
  try {
    const result = await api('POST', '/api/agent365/blueprint/publish', { dryRun });
    const cls = result.status === 'error' ? 'error' : 'success';
    els.publishResult.classList.add(cls);
    els.publishResult.textContent = JSON.stringify(result, null, 2);
    if (result.status === 'success') {
      // Refresh deployment context so the new blueprint id shows up.
      state.settings = await api('GET', '/api/settings');
      renderDeploymentContext();
    }
  } catch (err) {
    els.publishResult.classList.add('error');
    els.publishResult.textContent = `${err.message}\n\n${JSON.stringify(err.payload || {}, null, 2)}`;
  } finally {
    els.publishBtn.disabled = !state.account?.signedIn;
    els.dryRunBtn.disabled = false;
  }
}

els.authBtn.addEventListener('click', () => {
  state.token = els.authInput.value.trim();
  if (state.token) localStorage.setItem(TOKEN_KEY, state.token);
  bootstrap();
});

els.signInBtn.addEventListener('click', startSignIn);
els.signOutBtn.addEventListener('click', doSignOut);
els.dryRunBtn.addEventListener('click', () => publish(true));
els.publishBtn.addEventListener('click', () => publish(false));

els.copyBtn.addEventListener('click', async () => {
  if (!state.flow?.userCode) return;
  try {
    await navigator.clipboard.writeText(state.flow.userCode);
    els.copyBtn.textContent = '✓';
    setTimeout(() => { els.copyBtn.textContent = '⧉'; }, 1500);
  } catch { /* ignore */ }
});

// ---------------------------------------------------------------------
// User-delegated Microsoft 365 sign-in (per active instance).
// ---------------------------------------------------------------------

const userEls = {
  noInstance:    document.getElementById('user-no-instance'),
  signedOut:     document.getElementById('user-signed-out'),
  signedIn:      document.getElementById('user-signed-in'),
  signInBtn:     document.getElementById('btn-user-signin'),
  signOutBtn:    document.getElementById('btn-user-signout'),
  enableBtn:     document.getElementById('btn-enable-delegated'),
  username:      document.getElementById('user-username'),
  displayName:   document.getElementById('user-displayname'),
  mailBtn:       document.getElementById('btn-user-mail'),
  eventsBtn:     document.getElementById('btn-user-events'),
  result:        document.getElementById('user-result'),
  deviceFlow:    document.getElementById('user-device-flow'),
  deviceUri:     document.getElementById('user-device-uri'),
  deviceCode:    document.getElementById('user-device-code'),
  deviceStatus:  document.getElementById('user-device-status'),
  copyBtn:       document.getElementById('btn-user-copy-code'),
};

const userState = { flow: null, pollTimer: null };

async function refreshUserCard() {
  // Hide everything by default; the response decides what shows.
  userEls.noInstance.classList.add('hidden');
  userEls.signedOut.classList.add('hidden');
  userEls.signedIn.classList.add('hidden');
  try {
    const data = await api('GET', '/api/agent365/user/account');
    if (data.error === 'no active instance' || !data.instanceId) {
      userEls.noInstance.classList.remove('hidden');
      return;
    }
    if (data.signedIn) {
      userEls.signedIn.classList.remove('hidden');
      userEls.username.textContent = data.account?.username || '(unknown)';
      userEls.displayName.textContent = data.me?.displayName ? ` — ${data.me.displayName}` : '';
    } else {
      userEls.signedOut.classList.remove('hidden');
    }
  } catch (err) {
    userEls.signedOut.classList.remove('hidden');
    userEls.result.classList.remove('hidden');
    userEls.result.textContent = `account check failed: ${err.message}`;
  }
}

async function startUserSignIn() {
  userEls.signInBtn.disabled = true;
  userEls.deviceFlow.classList.add('hidden');
  userEls.result.classList.add('hidden');
  try {
    const flow = await api('POST', '/api/agent365/user/signin', {});
    userState.flow = flow;
    userEls.deviceFlow.classList.remove('hidden');
    userEls.deviceUri.href = flow.verificationUri;
    userEls.deviceUri.textContent = flow.verificationUri;
    userEls.deviceCode.textContent = flow.userCode;
    userEls.deviceStatus.textContent = 'Waiting for sign-in…';
    pollUserSignIn();
  } catch (err) {
    userEls.result.classList.remove('hidden');
    userEls.result.textContent = `sign-in failed: ${err.message}`;
  } finally {
    userEls.signInBtn.disabled = false;
  }
}

function pollUserSignIn() {
  if (userState.pollTimer) clearInterval(userState.pollTimer);
  userState.pollTimer = setInterval(async () => {
    if (!userState.flow?.flowId) return;
    try {
      const status = await api('GET', `/api/agent365/user/signin/${userState.flow.flowId}`);
      if (status.status === 'success') {
        clearInterval(userState.pollTimer);
        userState.pollTimer = null;
        userEls.deviceFlow.classList.add('hidden');
        await refreshUserCard();
      } else if (status.status === 'error') {
        clearInterval(userState.pollTimer);
        userState.pollTimer = null;
        userEls.deviceStatus.textContent = `sign-in failed: ${status.error || 'unknown'}`;
      }
    } catch (err) {
      userEls.deviceStatus.textContent = `poll failed: ${err.message}`;
    }
  }, 3000);
}

async function userSignOut() {
  try {
    await api('POST', '/api/agent365/user/signout', {});
  } finally {
    await refreshUserCard();
  }
}

async function showRecentMail() {
  userEls.result.classList.remove('hidden');
  userEls.result.textContent = 'Loading…';
  try {
    const data = await api('GET', '/api/agent365/user/mail?top=10');
    const lines = (data.items || []).map(m =>
      `• ${m.receivedDateTime ? m.receivedDateTime.slice(0,16).replace('T',' ') : '?'}` +
      `  ${m.from || '?'}\n  ${m.subject || '(no subject)'}` +
      (m.bodyPreview ? `\n  ${m.bodyPreview.slice(0, 120)}` : '')
    );
    userEls.result.textContent = lines.length ? lines.join('\n\n') : '(no messages)';
  } catch (err) {
    userEls.result.textContent = `failed: ${err.message}`;
  }
}

async function showUpcomingEvents() {
  userEls.result.classList.remove('hidden');
  userEls.result.textContent = 'Loading…';
  try {
    const data = await api('GET', '/api/agent365/user/events?top=10&daysAhead=7');
    const lines = (data.items || []).map(e =>
      `• ${e.start ? e.start.slice(0,16).replace('T',' ') : '?'}  ${e.subject || '(no title)'}` +
      (e.location ? `  @ ${e.location}` : '')
    );
    userEls.result.textContent = lines.length ? lines.join('\n') : '(no upcoming events)';
  } catch (err) {
    userEls.result.textContent = `failed: ${err.message}`;
  }
}

async function enableDelegatedOnActive() {
  userEls.result.classList.remove('hidden');
  userEls.result.textContent = 'Patching active instance…';
  try {
    // Look up the active instance id from the instances list.
    const data = await api('GET', '/api/agent365/instances');
    const id = data.activeInstanceId;
    if (!id) { userEls.result.textContent = 'no active instance'; return; }
    const result = await api('POST', `/api/agent365/instances/${id}/enable-user-delegated`, {});
    userEls.result.textContent =
      `Patched app ${result.clientId}. The next sign-in will prompt for the new delegated permissions.`;
  } catch (err) {
    userEls.result.textContent = `failed: ${err.message}`;
  }
}

userEls.signInBtn?.addEventListener('click', startUserSignIn);
userEls.signOutBtn?.addEventListener('click', userSignOut);
userEls.enableBtn?.addEventListener('click', enableDelegatedOnActive);
userEls.mailBtn?.addEventListener('click', showRecentMail);
userEls.eventsBtn?.addEventListener('click', showUpcomingEvents);
userEls.copyBtn?.addEventListener('click', async () => {
  if (!userState.flow?.userCode) return;
  try {
    await navigator.clipboard.writeText(userState.flow.userCode);
    userEls.copyBtn.textContent = '✓';
    setTimeout(() => { userEls.copyBtn.textContent = '⧉'; }, 1500);
  } catch { /* ignore */ }
});

// After the initial bootstrap() call below finishes, also load the user card.
// The bootstrap() function is a declaration so we wrap it via an extra call.
async function bootstrapAll() {
  await refreshUserCard();
}

bootstrap();
bootstrapAll();
