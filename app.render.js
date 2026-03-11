// Simple local API calls to /api/roster endpoints

const AUTO_SYNC_SESSION_KEY = 'fwpd_auto_sync_done';
const LOCAL_SYNC_TABS_KEY = 'fwpd_sync_tabs_v1';
const AUTH_TOKEN_KEY = 'fwpd_auth_token';
const COMMAND_USERS_SOURCE_URL_KEY = 'fwpd_command_users_source_url';
const DISCIPLINE_SOURCE_URL_KEY = 'fwpd_discipline_source_url';
const EVALUATION_SOURCE_URL_KEY = 'fwpd_evaluation_source_url';
const CALENDAR_VIEW_TIMEZONE_KEY = 'fwpd_calendar_view_timezone';
const AUTO_COMMAND_USERS_LINK_KEY = 'fwpd_command_users_auto_linked';
const DEFAULT_COMMAND_USERS_SOURCE_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vR6_40O35zd-9GMo_nTg5KS76Svzt1P8ZKrfBQwPAtLloGFtpE1r4JBP3t-F-meLlDKCpvWzZkhMlOb/pubhtml?gid=1476592599&single=true';
const DEFAULT_ROSTER_SOURCE_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vR6_40O35zd-9GMo_nTg5KS76Svzt1P8ZKrfBQwPAtLloGFtpE1r4JBP3t-F-meLlDKCpvWzZkhMlOb/pubhtml?gid=757275616&single=true';
const DEFAULT_DISCIPLINE_SOURCE_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vS4seHseWTG0lk0IBKCetQqz2elv2_QRVtFRaCbJIMbONhvsixRjc7VrERdyaW2tqUv6ZUfIA-4EztK/pubhtml?gid=10995956&single=true';
const DEFAULT_EVALUATION_SOURCE_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vR6_40O35zd-9GMo_nTg5KS76Svzt1P8ZKrfBQwPAtLloGFtpE1r4JBP3t-F-meLlDKCpvWzZkhMlOb/pub?output=csv&gid=1513386776';
const PORTAL_OWNER_EMAILS = ['mattprz89@gmail.com'];
const APP_BUILD = '20260309z32';
const MESSAGE_POLL_MS = 45000;
const DISCUSSION_POLL_MS = 4000;
const DEFAULT_CALENDAR_TIMEZONE = 'America/New_York';

const CALENDAR_TIMEZONE_OPTIONS = [
  { value: 'America/New_York', label: 'EST / EDT (US Eastern)' },
  { value: 'America/Chicago', label: 'CST / CDT (US Central)' },
  { value: 'America/Denver', label: 'MST / MDT (US Mountain)' },
  { value: 'America/Los_Angeles', label: 'PST / PDT (US Pacific)' },
  { value: 'UTC', label: 'UTC' }
];

let currentUser = null;
let unreadMessageCount = 0;
let messagePollTimer = null;
let lastLoadedReportItems = [];
let dataAutoSyncPromise = null;
let ftoListLoadToken = 0;
let headerAlertLines = [];
let discussionPollTimer = null;
let headerClockTimer = null;

function formatUserDisplayName(user) {
  const rank = String((user && user.rank) || '').trim();
  const name = String((user && user.characterName) || '').trim();
  if (rank && name) return rank + ' ' + name;
  return name || rank || 'Officer';
}

function isPrivilegedRoleClient(roleText) {
  const role = String(roleText || '').trim().toLowerCase();
  if (!role) return false;
  if (role.includes('admin')) return true;
  if (role.includes('chief')) return true;
  if (role.includes('commander')) return true;
  if (role.includes('supervisor')) return true;
  return false;
}

function canPromoteOfficersClient(roleText) {
  const role = String(roleText || '').trim().toLowerCase();
  if (!role) return false;
  if (role.includes('admin')) return true;
  if (role.includes('chief')) return true;
  if (role.includes('commander')) return true;
  return false;
}

function normalizeEmailClient(email) {
  return String(email || '').trim().toLowerCase();
}

function isPortalOwnerClient(user) {
  const email = normalizeEmailClient(user && user.email);
  if (!email) return false;
  return PORTAL_OWNER_EMAILS.includes(email);
}

function hasLeadershipAccessClient(user) {
  const role = String((user && user.role) || '');
  return canPromoteOfficersClient(role) || isPortalOwnerClient(user);
}

function hasAdminAccessClient(user) {
  const role = String((user && user.role) || '');
  return normalizeAccessRoleClient(role) === 'admin' || isPortalOwnerClient(user);
}

function normalizeAccessRoleClient(roleText) {
  const role = String(roleText || '').trim().toLowerCase();
  if (role === 'admin') return 'admin';
  if (role === 'chief') return 'chief';
  if (role === 'commander') return 'commander';
  return 'command';
}

function getAuthToken() {
  return localStorage.getItem(AUTH_TOKEN_KEY) || '';
}

function setAuthToken(token) {
  if (token) localStorage.setItem(AUTH_TOKEN_KEY, token);
  else localStorage.removeItem(AUTH_TOKEN_KEY);
}

function authHeaders(extra = {}) {
  const token = getAuthToken();
  const headers = Object.assign({}, extra);
  if (token) headers.Authorization = 'Bearer ' + token;
  return headers;
}

async function authFetch(url, options = {}) {
  const headers = authHeaders(options.headers || {});
  const response = await fetch(url, Object.assign({}, options, { headers }));
  if (response.status === 401) {
    setAuthToken('');
    stopMessagePolling();
    unreadMessageCount = 0;
    currentUser = null;
    showAuthBanner();
    applyRuntimeLayoutFixes();
    renderLoginScreen('Session expired or unauthorized. Please log in again.');
  }
  return response;
}

function isLoggedIn() {
  return !!currentUser;
}

function showAuthBanner() {
  const existing = document.getElementById('authBanner');
  if (existing) existing.remove();

  const tools = ensureHeaderTools();
  if (!tools) return;

  const accountBtn = document.getElementById('headerAccountBtn');
  const logoutBtn = document.getElementById('headerLogoutBtn');
  const notifyBtn = document.getElementById('headerNotifyBtn');
  const notifyPanel = document.getElementById('headerNotifyPanel');

  if (currentUser) {
    tools.style.display = 'flex';
    if (accountBtn) {
      accountBtn.textContent = 'Account';
      accountBtn.disabled = false;
      accountBtn.title = formatUserDisplayName(currentUser);
    }
    if (logoutBtn) {
      logoutBtn.disabled = false;
      logoutBtn.style.display = 'inline-block';
    }
  } else {
    tools.style.display = 'none';
    if (accountBtn) {
      accountBtn.textContent = 'Account';
      accountBtn.disabled = true;
      accountBtn.title = 'Login required';
    }
    if (logoutBtn) {
      logoutBtn.disabled = true;
      logoutBtn.style.display = 'inline-block';
    }
    if (notifyBtn) notifyBtn.disabled = true;
    if (notifyPanel) notifyPanel.hidden = true;
  }

  setHeaderAlerts(headerAlertLines);
}

function ensureHeaderTools() {
  const header = document.querySelector('.header');
  if (!header) return null;

  let tools = document.getElementById('headerTools');
  if (!tools) {
    tools = document.createElement('div');
    tools.id = 'headerTools';
    tools.className = 'header-tools';
    tools.innerHTML = '' +
      '<button id="headerNotifyBtn" class="header-tool-btn header-bell-btn" type="button" aria-label="Notifications"><span class="tool-icon" aria-hidden="true">&#128276;</span><span class="tool-count" aria-hidden="true"></span></button>' +
      '<div id="headerNotifyPanel" class="header-notify-panel" hidden>No alerts.</div>' +
      '<button id="headerAccountBtn" class="header-tool-btn" type="button">Account</button>' +
      '<button id="headerLogoutBtn" class="header-tool-btn" type="button">Logout</button>';
    header.appendChild(tools);
  }

  const notifyBtn = tools.querySelector('#headerNotifyBtn');
  const notifyPanel = tools.querySelector('#headerNotifyPanel');
  const accountBtn = tools.querySelector('#headerAccountBtn');
  const logoutBtn = tools.querySelector('#headerLogoutBtn');
  const initialized = tools.getAttribute('data-initialized') === '1';

  if (notifyBtn) notifyBtn.classList.add('header-bell-btn');

  if (initialized) return tools;

  if (notifyBtn && notifyPanel) {
    notifyBtn.addEventListener('click', (event) => {
      event.stopPropagation();
      notifyPanel.hidden = !notifyPanel.hidden;
      if (!notifyPanel.hidden) loadDashboardAlerts();
    });
  }

  if (accountBtn) {
    accountBtn.addEventListener('click', () => {
      if (!isLoggedIn()) return;
      loadPage('account');
    });
  }

  if (logoutBtn) {
    logoutBtn.addEventListener('click', () => {
      if (!isLoggedIn()) return;
      logoutAccount();
    });
  }

  document.addEventListener('click', (event) => {
    if (!notifyPanel || notifyPanel.hidden) return;
    if (!tools.contains(event.target)) notifyPanel.hidden = true;
  });

  tools.setAttribute('data-initialized', '1');

  return tools;
}

function setHeaderAlerts(lines) {
  headerAlertLines = Array.isArray(lines) ? lines.filter(Boolean) : [];

  const notifyBtn = document.getElementById('headerNotifyBtn');
  const notifyPanel = document.getElementById('headerNotifyPanel');
  if (!notifyBtn || !notifyPanel) return;

  const alertCount = headerAlertLines.length;
  notifyBtn.disabled = !isLoggedIn();
  const countBadge = alertCount > 0 ? ('<span class="tool-count" aria-hidden="true">' + alertCount + '</span>') : '<span class="tool-count" aria-hidden="true"></span>';
  notifyBtn.innerHTML = '<span class="tool-icon" aria-hidden="true">&#128276;</span>' + countBadge;
  notifyBtn.setAttribute('aria-label', alertCount > 0 ? ('Notifications (' + alertCount + ')') : 'Notifications');
  notifyBtn.title = 'Unread messages: ' + String(unreadMessageCount || 0);

  if (!headerAlertLines.length) {
    notifyPanel.innerHTML = '<div class="notify-empty">No current alerts.</div>';
    return;
  }

  notifyPanel.innerHTML = headerAlertLines
    .slice(0, 12)
    .map((line) => '<div class="notify-item">' + escapeHtml(String(line || '')) + '</div>')
    .join('');
}

function applyRuntimeLayoutFixes() {
  const sidebar = document.querySelector('.sidebar');
  if (sidebar) {
    let links = Array.from(sidebar.querySelectorAll('a'));
    let messagesLink = null;
    let adminLink = null;
    let adminChatLink = null;
    let discussionsLink = null;
    let calendarLink = null;
    let ftoLink = null;
    let promotionRecLink = null;
    let highCommandApprovalLink = null;
    links.forEach((link) => {
      const text = String(link.textContent || '').trim().toLowerCase();
      if (text === 'sheet tabs') {
        link.remove();
      }
      if (text.startsWith('messages')) {
        messagesLink = link;
      }
      if (text === 'admin') {
        adminLink = link;
      }
      if (text === 'admin chat') {
        adminChatLink = link;
      }
      if (text === 'discussions' || text === 'message board' || text === 'chat room' || text === 'chat') {
        discussionsLink = link;
      }
      if (text === 'calendar' || text === 'event calendar') {
        calendarLink = link;
      }
      if (text === 'fto') {
        ftoLink = link;
      }
      if (text === 'promotion recommendations') {
        promotionRecLink = link;
      }
      if (text === 'high command approval') {
        highCommandApprovalLink = link;
      }
    });

    links = Array.from(sidebar.querySelectorAll('a'));
    const accountLink = links.find((link) => String(link.textContent || '').trim().toLowerCase() === 'account');
    if (accountLink) accountLink.remove();

    if (!messagesLink) {
      messagesLink = document.createElement('a');
      messagesLink.setAttribute('href', "javascript:loadPage('messages')");
      messagesLink.textContent = 'Messages';
      sidebar.appendChild(messagesLink);
    }

    if (!discussionsLink) {
      discussionsLink = document.createElement('a');
      discussionsLink.setAttribute('href', "javascript:loadPage('discussions')");
      discussionsLink.textContent = 'Chat Room';
      sidebar.appendChild(discussionsLink);
    } else {
      discussionsLink.setAttribute('href', "javascript:loadPage('discussions')");
      discussionsLink.textContent = 'Chat Room';
    }

    if (!calendarLink) {
      calendarLink = document.createElement('a');
      calendarLink.setAttribute('href', "javascript:loadPage('calendar')");
      calendarLink.textContent = 'Calendar';
      sidebar.appendChild(calendarLink);
    }

    const showAdmin = !!currentUser && hasAdminAccessClient(currentUser);
    const showFto = !!currentUser && hasLeadershipAccessClient(currentUser);
    const showHighCommand = !!currentUser && hasLeadershipAccessClient(currentUser);
    if (showAdmin && !adminLink) {
      adminLink = document.createElement('a');
      adminLink.setAttribute('href', "javascript:loadPage('admin')");
      adminLink.textContent = 'Admin';
      sidebar.appendChild(adminLink);
    }
    if (!showAdmin && adminLink) {
      adminLink.remove();
    }

    if (showAdmin && !adminChatLink) {
      adminChatLink = document.createElement('a');
      adminChatLink.setAttribute('href', "javascript:loadPage('adminchat')");
      adminChatLink.textContent = 'Admin Chat';
      sidebar.appendChild(adminChatLink);
    }
    if (!showAdmin && adminChatLink) {
      adminChatLink.remove();
    }

    if (showFto && !ftoLink) {
      ftoLink = document.createElement('a');
      ftoLink.setAttribute('href', "javascript:loadPage('fto')");
      ftoLink.textContent = 'FTO';
      sidebar.appendChild(ftoLink);
    }
    if (!showFto && ftoLink) {
      ftoLink.remove();
    }

    if (showHighCommand && !promotionRecLink) {
      promotionRecLink = document.createElement('a');
      promotionRecLink.setAttribute('href', "javascript:loadPage('promotion-recommendations')");
      promotionRecLink.textContent = 'Promotion Recommendations';
      sidebar.appendChild(promotionRecLink);
    }
    if (showHighCommand && promotionRecLink) {
      promotionRecLink.setAttribute('href', "javascript:loadPage('promotion-recommendations')");
      promotionRecLink.textContent = 'Promotion Recommendations';
    }
    if (!showHighCommand && promotionRecLink) {
      promotionRecLink.remove();
    }

    if (showHighCommand && !highCommandApprovalLink) {
      highCommandApprovalLink = document.createElement('a');
      highCommandApprovalLink.setAttribute('href', "javascript:loadPage('high-command-approval')");
      highCommandApprovalLink.textContent = 'High Command Approval';
      sidebar.appendChild(highCommandApprovalLink);
    }
    if (showHighCommand && highCommandApprovalLink) {
      highCommandApprovalLink.setAttribute('href', "javascript:loadPage('high-command-approval')");
      highCommandApprovalLink.textContent = 'High Command Approval';
    }
    if (!showHighCommand && highCommandApprovalLink) {
      highCommandApprovalLink.remove();
    }

    const countText = unreadMessageCount > 0 ? ('Messages (' + unreadMessageCount + ')') : 'Messages';
    messagesLink.textContent = countText;

    let footer = document.getElementById('sidebarBuildTag');
    if (!footer) {
      footer = document.createElement('div');
      footer.id = 'sidebarBuildTag';
      footer.className = 'sidebar-build-tag';
      sidebar.appendChild(footer);
    } else {
      // Move to end if not already last
      if (sidebar.lastElementChild !== footer) {
        sidebar.appendChild(footer);
      }
    }
    footer.textContent = 'Build ' + APP_BUILD;
  }

  const title = document.querySelector('.title');
  if (title) {
    let dateTimeTag = document.getElementById('headerDateTimeTag');
    if (!dateTimeTag) {
      dateTimeTag = document.createElement('div');
      dateTimeTag.id = 'headerDateTimeTag';
      dateTimeTag.className = 'header-datetime-tag';
      title.appendChild(dateTimeTag);
    }

    const updateDateTime = () => {
      const now = new Date();
      dateTimeTag.textContent = now.toLocaleString('en-US', {
        month: 'long',
        day: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
    };
    updateDateTime();
    if (!headerClockTimer) {
      headerClockTimer = setInterval(updateDateTime, 30000);
    }
  }
}

function stopMessagePolling() {
  if (messagePollTimer) {
    clearInterval(messagePollTimer);
    messagePollTimer = null;
  }
}

function startDiscussionLiveRefresh() {
  stopDiscussionLiveRefresh();
  discussionPollTimer = setInterval(() => {
    if (!isLoggedIn()) return;
    const thread = document.getElementById('discussionMessages');
    if (!thread) return;
    loadDiscussionMessages({ preserveScroll: true, silent: true });
  }, DISCUSSION_POLL_MS);
}

function stopDiscussionLiveRefresh() {
  if (discussionPollTimer) {
    clearInterval(discussionPollTimer);
    discussionPollTimer = null;
  }
}

function startMessagePolling() {
  stopMessagePolling();
  messagePollTimer = setInterval(() => {
    if (!isLoggedIn()) return;
    loadUnreadMessageCount();
    checkCalendarReminders();
  }, MESSAGE_POLL_MS);
}

async function loadUnreadMessageCount() {
  if (!isLoggedIn()) {
    unreadMessageCount = 0;
    applyRuntimeLayoutFixes();
    showAuthBanner();
    return;
  }
  try {
    const response = await authFetch('/api/messages/unread-count');
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed unread count');
    unreadMessageCount = Number(data.unread || 0);
  } catch (e) {
    unreadMessageCount = 0;
  }
  applyRuntimeLayoutFixes();
  showAuthBanner();
}

function setAuthLockedLayout(locked) {
  const sidebar = document.querySelector('.sidebar');
  if (sidebar) sidebar.style.display = locked ? 'none' : 'block';
}

function renderLoginScreen(statusText = '') {
  setAuthLockedLayout(true);
  document.getElementById('content').innerHTML = `
    <div class="login-shell" style="max-width:1060px;margin:20px auto;">
      <div class="login-cli-header">MDT ACCESS TERMINAL :: AUTHENTICATION</div>
      <div class="login-cli-subheader">SECURE SESSION REQUIRED</div>
      <h2>Command Login</h2>
      <div class="login-legal-disclaimer">
        <div class="login-legal-title">LEGAL NOTICE</div>
        <div class="login-legal-summary">ROLEPLAY USE ONLY. THIS PORTAL IS NOT AFFILIATED WITH THE FORT WORTH POLICE DEPARTMENT OR ANY GOVERNMENT AGENCY.</div>
        <details class="login-legal-details">
          <summary>VIEW FULL LEGAL DISCLAIMER</summary>
          <div class="login-legal-fulltext">FOR THE AVOIDANCE OF DOUBT, THIS PORTAL IS A PRIVATE DIGITAL RESOURCE CREATED SOLELY FOR A FIVEM ROLEPLAY COMMUNITY AND IS INTENDED EXCLUSIVELY FOR FICTIONAL, ENTERTAINMENT, AND TRAINING-STYLE ROLEPLAY PURPOSES. THIS PROJECT IS NOT AFFILIATED WITH, ENDORSED BY, SPONSORED BY, OR OTHERWISE ASSOCIATED WITH THE FORT WORTH POLICE DEPARTMENT, NOR WITH ANY MUNICIPAL, COUNTY, STATE, FEDERAL, OR OTHER GOVERNMENTAL AGENCY OR ENTITY. ANY USE OF NAMES, TERMINOLOGY, TITLES, MARKINGS, OR INSIGNIA-LIKE REFERENCES IS STRICTLY FOR NON-OFFICIAL ROLEPLAY CONTEXT AND SHALL NOT BE CONSTRUED AS REPRESENTING REAL-WORLD AUTHORITY, OFFICIAL POLICY, OR GOVERNMENT ACTION. NORTH TEXAS ROLEPLAY (NTXRP), INCLUDING ITS OWNERS, STAFF, AND AFFILIATES, DISCLAIMS RESPONSIBILITY FOR ANY MISINTERPRETATION OF CONTENT OR PRESENTATION WITHIN THIS PORTAL AND SHALL COMPLY WITH ANY VALID LEGAL NOTICE, INCLUDING CEASE-AND-DESIST DEMANDS, AS REQUIRED BY APPLICABLE LAW.</div>
        </details>
      </div>
      <p>Only users listed in <b>Command_Users</b> can create accounts and log in.</p>

      <div style="display:flex;gap:10px;flex-wrap:wrap;margin:10px 0 16px 0;">
        <button id="showLoginPane">Login</button>
        <button id="showCreatePane">Create Account</button>
      </div>

      <div style="border:1px solid rgba(255,255,255,.2);padding:10px;margin-bottom:12px;">
        <b>Command_Users Access</b><br>
        <span style="font-size:13px;opacity:.9">If your email is missing, click Sync Command_Users or use Link Command_Users below.</span>
        <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
          <button id="syncCommandUsersBtn">Sync Command_Users</button>
        </div>
        <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
          <input id="commandUsersUrl" type="text" style="min-width:260px;flex:1" placeholder="Paste Command_Users tab CSV/Google link">
          <button id="linkCommandUsersBtn">Link Command_Users</button>
        </div>
      </div>

      <div id="loginPane" style="display:block">
        <h3>Login</h3>
        <label>Email</label><br>
        <input id="loginEmail" type="email" style="width:100%;max-width:360px"><br>
        <label>Password</label><br>
        <input id="loginPassword" type="password" style="width:100%;max-width:360px"><br><br>
        <button id="loginBtn">Login</button>
      </div>

      <div id="createPane" style="display:none">
        <h3>Create Account</h3>
        <label>Email</label><br>
        <input id="createEmail" type="email" style="width:100%;max-width:360px"><br>
        <label>Password</label><br>
        <input id="createPassword" type="password" style="width:100%;max-width:360px"><br><br>
        <label>Verify Password</label><br>
        <input id="createPasswordVerify" type="password" style="width:100%;max-width:360px"><br><br>
        <button id="createAccountBtn">Create Account</button>
      </div>

      <pre id="accountStatus" style="margin-top:14px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">${statusText || 'Please login to continue.'}</pre>
    </div>
  `;

  const loginPane = document.getElementById('loginPane');
  const createPane = document.getElementById('createPane');
  document.getElementById('showLoginPane').addEventListener('click', () => {
    loginPane.style.display = 'block';
    createPane.style.display = 'none';
  });
  document.getElementById('showCreatePane').addEventListener('click', () => {
    loginPane.style.display = 'none';
    createPane.style.display = 'block';
  });
  document.getElementById('createAccountBtn').addEventListener('click', createAccount);
  document.getElementById('loginBtn').addEventListener('click', loginAccount);
  const syncBtn = document.getElementById('syncCommandUsersBtn');
  if (syncBtn) syncBtn.addEventListener('click', syncCommandUsersHardset);
  const linkBtn = document.getElementById('linkCommandUsersBtn');
  if (linkBtn) linkBtn.addEventListener('click', linkCommandUsersTab);
  const commandUsersUrlEl = document.getElementById('commandUsersUrl');
  if (commandUsersUrlEl) {
    const savedCommandUsersUrl = localStorage.getItem(COMMAND_USERS_SOURCE_URL_KEY) || DEFAULT_COMMAND_USERS_SOURCE_URL;
    if (savedCommandUsersUrl) commandUsersUrlEl.value = savedCommandUsersUrl;
  }
  autoLinkCommandUsersOnLogin();
}

async function syncCommandUsersHardset() {
  const status = document.getElementById('accountStatus');
  const saved = String(localStorage.getItem(COMMAND_USERS_SOURCE_URL_KEY) || '').trim();
  const defaultUrl = String(DEFAULT_COMMAND_USERS_SOURCE_URL || '').trim();
  const sourceUrl = saved || defaultUrl;

  try {
    if (sourceUrl) {
      const data = await linkCommandUsersTabByUrl(sourceUrl);
      const rows = (((data || {}).import || {}).result || []).find(x => x.name === 'command_users');
      const rowCount = rows && typeof rows.rows === 'number' ? rows.rows : 0;
      if (status) status.textContent = 'Command_Users synced (' + rowCount + ' rows). You can create account now.';
      return;
    }

    const response = await fetch('/api/auth/ensure-command-users', { method: 'POST' });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || 'Sync failed');
    const rows = Number(data.afterCount || 0);
    if (status) status.textContent = 'Command_Users synced (' + rows + ' rows). You can create account now.';
  } catch (err) {
    if (status) status.textContent = 'Command_Users sync failed: ' + err.message;
  }
}

async function linkCommandUsersTabByUrl(url) {
  const sourceUrl = String(url || '').trim();
  if (!sourceUrl) return null;
  const response = await fetch('/api/auth/link-command-users', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ url: sourceUrl })
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || 'Failed to link Command_Users tab');
  return data;
}

async function autoLinkCommandUsersOnLogin() {
  if (sessionStorage.getItem(AUTO_COMMAND_USERS_LINK_KEY) === '1') return;

  try {
    const response = await fetch('/api/auth/ensure-command-users', { method: 'POST' });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to sync Command_Users');
    sessionStorage.setItem(AUTO_COMMAND_USERS_LINK_KEY, '1');
  } catch (err) {
    // Silent by design so login UI stays clean for non-admin users.
  }
}

async function linkCommandUsersTab() {
  const url = String((document.getElementById('commandUsersUrl') || {}).value || '').trim();
  const status = document.getElementById('accountStatus');
  if (!url) {
    if (status) status.textContent = 'Please paste a Command_Users tab link first.';
    return;
  }

  try {
    const data = await linkCommandUsersTabByUrl(url);
    try { localStorage.setItem(COMMAND_USERS_SOURCE_URL_KEY, url); } catch (e) {}

    const rows = (((data || {}).import || {}).result || []).find(x => x.name === 'command_users');
    const rowCount = rows && typeof rows.rows === 'number' ? rows.rows : 0;
    if (status) status.textContent = 'Command_Users linked successfully (' + rowCount + ' rows). You can now create your account.';
  } catch (err) {
    if (status) status.textContent = 'Link failed: ' + err.message;
  }
}

async function refreshAuthSession() {
  applyRuntimeLayoutFixes();
  const token = getAuthToken();
  if (!token) {
    currentUser = null;
    showAuthBanner();
    renderLoginScreen();
    return;
  }

  try {
    const response = await authFetch('/api/auth/me');
    const data = await response.json();
    if (!response.ok) {
      setAuthToken('');
      currentUser = null;
    } else {
      currentUser = data.user || null;
    }
  } catch (e) {
    currentUser = null;
  }
  showAuthBanner();
  if (!currentUser) {
    stopMessagePolling();
    unreadMessageCount = 0;
    renderLoginScreen();
    return;
  }

  await loadUnreadMessageCount();
  await checkCalendarReminders();
  startMessagePolling();

  // On initial authenticated load, land on dashboard.
  const content = document.getElementById('content');
  if (content && /Command Login/i.test(String(content.textContent || ''))) {
    loadPage('dashboard');
  }
}

function getLocalSyncTabs() {
  try {
    const raw = localStorage.getItem(LOCAL_SYNC_TABS_KEY);
    const parsed = JSON.parse(raw || '[]');
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(t => t && t.name && t.url);
  } catch (e) {
    return [];
  }
}

function setLocalSyncTabs(tabs) {
  try {
    const safe = Array.isArray(tabs) ? tabs.filter(t => t && t.name && t.url) : [];
    localStorage.setItem(LOCAL_SYNC_TABS_KEY, JSON.stringify(safe));
  } catch (e) {
    // Ignore localStorage errors (private mode/quota).
  }
}

async function ensureDataTabsSynced() {
  if (!isLoggedIn()) return;
  if (dataAutoSyncPromise) {
    await dataAutoSyncPromise;
    return;
  }

  // Prevent duplicate sync requests when multiple pages initialize together.
  dataAutoSyncPromise = (async () => {
    await autoSyncOnLoad();
  })();

  try {
    await dataAutoSyncPromise;
  } finally {
    dataAutoSyncPromise = null;
  }
}

function loadPage(page){
applyRuntimeLayoutFixes();
stopDiscussionLiveRefresh();
if(!isLoggedIn()){
renderLoginScreen();
return;
}

setAuthLockedLayout(false);

/* DASHBOARD */

if(page === "dashboard"){

document.getElementById("content").innerHTML = `
<h2>Command Dashboard</h2>

<div id="welcomeMessage" style="margin-top:2px;margin-bottom:10px;color:#d8f3ff"></div>

<div style="margin-top:10px;border:1px solid rgba(255,255,255,.2);padding:12px;background:rgba(0,0,0,.15)">
  <div style="margin:0 0 12px 0;color:#f3bc40;font-family:'Barlow Condensed','Trebuchet MS',sans-serif;font-size:24px;letter-spacing:.5px;line-height:1.1">FORT WORTH POLICE DEPARTMENT - MISSION STATEMENT</div>
  <p style="margin-top:8px;margin-bottom:10px;line-height:1.45">The Fort Worth Police Department is committed to safeguarding our community through integrity, professionalism, and unwavering service. Our mission is to protect life and property, uphold the law with fairness and respect, and strengthen public trust through transparency and accountability.</p>
  <p style="margin:0;line-height:1.45">We strive to maintain a safe and thriving city by working collaboratively with our residents, embracing innovation, and holding ourselves to the highest standards of conduct. Every member of this department is dedicated to acting with courage, compassion, and honor in the pursuit of justice.</p>
</div>

`;

if (currentUser) {
  const welcome = document.getElementById('welcomeMessage');
  if (welcome) {
    welcome.textContent = 'Welcome ' + formatUserDisplayName(currentUser) + '.';
  }
}

loadDashboardAlerts();
autoSyncOnLoad();

}


/* REPORTS */

if(page === "reports"){

document.getElementById("content").innerHTML = `
<h2>Reports</h2>
<p>Command review center for discipline, cadet evaluations, and internal messages.</p>

<div style="margin-top:12px;display:flex;gap:10px;flex-wrap:wrap;">
  <button id="showAllReportsBtn">All Reports</button>
  <button id="showDisciplineReportsBtn">Discipline Queue</button>
  <button id="showEvaluationReportsBtn">Evaluation Queue</button>
  <button id="showMessageReportsBtn">Internal Messages</button>
</div>

<div style="margin-top:14px;border:1px solid rgba(255,255,255,.2);padding:10px;background:rgba(0,0,0,.15)">
  <b>Link Disciplinary Form Database</b>
  <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;align-items:center">
    <input id="disciplineTabName" type="text" value="disciplinary_forms" style="max-width:220px" placeholder="tab name">
    <input id="disciplineSourceUrl" type="text" style="min-width:260px;flex:1" placeholder="Paste separate disciplinary Google Sheet/CSV link">
    <button id="linkDisciplineSourceBtn">Link Source</button>
  </div>
</div>

<div style="margin-top:10px;border:1px solid rgba(255,255,255,.2);padding:10px;background:rgba(0,0,0,.15)">
  <b>Link Cadet Evaluations Database</b>
  <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;align-items:center">
    <input id="evaluationTabName" type="text" value="cadet_evaluations" style="max-width:220px" placeholder="tab name">
    <input id="evaluationSourceUrl" type="text" style="min-width:260px;flex:1" placeholder="Paste cadet evaluations Google Sheet/CSV link">
    <button id="linkEvaluationSourceBtn">Link Source</button>
  </div>
</div>

<pre id="reportsSummary" style="margin-top:14px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Loading reports summary...</pre>

<div style="margin-top:14px;display:flex;gap:8px;flex-wrap:wrap;align-items:center">
  <label>Type</label>
  <select id="reportTypeFilter">
    <option value="all">All</option>
    <option value="discipline">Discipline</option>
    <option value="evaluation">Evaluation</option>
    <option value="message">Messages</option>
  </select>
  <label>Status</label>
  <select id="reportStatusFilter">
    <option value="all">All</option>
    <option value="pending">Pending</option>
    <option value="approved">Approved</option>
    <option value="denied">Denied</option>
  </select>
  <label>Officer</label>
  <input id="reportOfficerFilter" type="text" placeholder="Search officer name" style="min-width:200px">
  <button id="refreshReportsBtn">Refresh</button>
</div>

<div style="margin-top:10px;overflow:auto">
  <table id="reportsTable">
    <thead>
      <tr>
        <th>Type</th>
        <th>Subject</th>
        <th>Officer</th>
        <th>Date</th>
        <th>Source Tab</th>
        <th>Status</th>
        <th>Approved By</th>
        <th>Approved At</th>
        <th>Actions</th>
      </tr>
    </thead>
    <tbody></tbody>
  </table>
</div>

`;

loadReportsSummary();
loadReportItems();
if (isLoggedIn()) {
  ensureDataTabsSynced()
    .then(() => {
      loadReportsSummary();
      loadReportItems();
    })
    .catch(() => {
      // Keep current data visible if auto-sync fails.
    });
}

const linkBtn = document.getElementById('linkDisciplineSourceBtn');
if (linkBtn) linkBtn.addEventListener('click', linkDisciplinarySource);

const evalLinkBtn = document.getElementById('linkEvaluationSourceBtn');
if (evalLinkBtn) evalLinkBtn.addEventListener('click', linkEvaluationSource);

const savedDisciplineUrl = localStorage.getItem(DISCIPLINE_SOURCE_URL_KEY) || DEFAULT_DISCIPLINE_SOURCE_URL;
const disciplineUrlEl = document.getElementById('disciplineSourceUrl');
if (disciplineUrlEl && savedDisciplineUrl) disciplineUrlEl.value = savedDisciplineUrl;

const savedEvaluationUrl = localStorage.getItem(EVALUATION_SOURCE_URL_KEY) || DEFAULT_EVALUATION_SOURCE_URL;
const evaluationUrlEl = document.getElementById('evaluationSourceUrl');
if (evaluationUrlEl && savedEvaluationUrl) evaluationUrlEl.value = savedEvaluationUrl;

const refreshBtn = document.getElementById('refreshReportsBtn');
if (refreshBtn) refreshBtn.addEventListener('click', () => {
  loadReportsSummary();
  loadReportItems();
});

const allBtn = document.getElementById('showAllReportsBtn');
if (allBtn) allBtn.addEventListener('click', () => {
  const t = document.getElementById('reportTypeFilter');
  const s = document.getElementById('reportStatusFilter');
  if (t) t.value = 'all';
  if (s) s.value = 'all';
  loadReportItems();
});

const disciplineBtn = document.getElementById('showDisciplineReportsBtn');
if (disciplineBtn) disciplineBtn.addEventListener('click', () => {
  const t = document.getElementById('reportTypeFilter');
  const s = document.getElementById('reportStatusFilter');
  if (t) t.value = 'discipline';
  if (s) s.value = 'pending';
  loadReportItems();
});

const evalBtn = document.getElementById('showEvaluationReportsBtn');
if (evalBtn) evalBtn.addEventListener('click', () => {
  const t = document.getElementById('reportTypeFilter');
  const s = document.getElementById('reportStatusFilter');
  if (t) t.value = 'evaluation';
  if (s) s.value = 'pending';
  loadReportItems();
});

const msgBtn = document.getElementById('showMessageReportsBtn');
if (msgBtn) msgBtn.addEventListener('click', () => {
  const t = document.getElementById('reportTypeFilter');
  const s = document.getElementById('reportStatusFilter');
  if (t) t.value = 'message';
  if (s) s.value = 'all';
  loadReportItems();
});

const typeFilter = document.getElementById('reportTypeFilter');
if (typeFilter) typeFilter.addEventListener('change', loadReportItems);

const statusFilter = document.getElementById('reportStatusFilter');
if (statusFilter) statusFilter.addEventListener('change', loadReportItems);

const officerFilter = document.getElementById('reportOfficerFilter');
if (officerFilter) officerFilter.addEventListener('input', loadReportItems);

}


/* MESSAGES */

if(page === "messages"){

document.getElementById("content").innerHTML = `
<h2>Internal Messages</h2>
<p>Secure command inbox with notifications.</p>

<div style="margin-top:12px;border:1px solid rgba(255,255,255,.2);padding:10px;background:rgba(0,0,0,.15)">
  <b>Compose Message</b>
  <div style="margin-top:8px;display:grid;grid-template-columns:1fr;gap:8px;max-width:760px;">
    <label for="msgRecipient">Recipients (Ctrl/Cmd+Click for multiple)</label>
    <select id="msgRecipient" multiple size="8"><option value="">Loading recipients...</option></select>
    <input id="msgSubject" type="text" placeholder="Subject">
    <textarea id="msgBody" rows="5" placeholder="Message body"></textarea>
  </div>
  <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;">
    <button id="sendMessageBtn">Send</button>
    <button id="refreshMessagesBtn">Refresh</button>
    <button id="showInboxBtn">Inbox</button>
    <button id="showSentBtn">Sent</button>
  </div>
</div>

<pre id="messagesStatus" style="margin-top:10px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Loading messages...</pre>

<div style="margin-top:10px;overflow:auto">
  <table id="messagesTable">
    <thead>
      <tr>
        <th>From/To</th>
        <th>Subject</th>
        <th>Preview</th>
        <th>Sent</th>
        <th>Status</th>
        <th>Action</th>
      </tr>
    </thead>
    <tbody></tbody>
  </table>
</div>
`;

const sendBtn = document.getElementById('sendMessageBtn');
if (sendBtn) sendBtn.addEventListener('click', sendInternalMessage);

const refreshBtn = document.getElementById('refreshMessagesBtn');
if (refreshBtn) refreshBtn.addEventListener('click', loadInboxMessages);

const inboxBtn = document.getElementById('showInboxBtn');
if (inboxBtn) inboxBtn.addEventListener('click', loadInboxMessages);

const sentBtn = document.getElementById('showSentBtn');
if (sentBtn) sentBtn.addEventListener('click', loadSentMessages);

loadMessageRecipients();
loadInboxMessages();

}

if(page === "discussions"){

document.getElementById("content").innerHTML = `
<h2>Chat Room</h2>
<p>Open command chat room for all members. Messages refresh automatically.</p>

<div class="discussion-chat-shell" style="margin-top:10px;">
  <div id="discussionMessages" class="discussion-chat-thread">Loading discussions...</div>
  <div class="discussion-chat-compose">
    <textarea id="discussionText" rows="2" placeholder="Type a message and press Enter to send" style="width:100%"></textarea>
    <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;">
      <button id="postDiscussionBtn">Send</button>
      <button id="refreshDiscussionBtn">Refresh</button>
    </div>
  </div>
</div>

<pre id="discussionStatus" style="margin-top:10px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Connecting to discussion feed...</pre>
`;

const postBtn = document.getElementById('postDiscussionBtn');
if (postBtn) postBtn.addEventListener('click', postDiscussionMessage);
const refreshBtn = document.getElementById('refreshDiscussionBtn');
if (refreshBtn) refreshBtn.addEventListener('click', loadDiscussionMessages);
const input = document.getElementById('discussionText');
if (input) {
  input.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      postDiscussionMessage();
    }
  });
}
loadDiscussionMessages();
startDiscussionLiveRefresh();

}

if(page === "adminchat"){

if(!hasAdminAccessClient(currentUser)){
loadPage('dashboard');
return;
}

document.getElementById("content").innerHTML = `
<h2>Admin Chat</h2>
<p>Private internal channel for admins.</p>

<div style="margin-top:10px;border:1px solid rgba(255,255,255,.2);padding:10px;background:rgba(0,0,0,.15)">
  <textarea id="adminChatText" rows="3" placeholder="Send admin-only message" style="width:100%"></textarea>
  <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;">
    <button id="postAdminChatBtn">Send</button>
    <button id="refreshAdminChatBtn">Refresh</button>
  </div>
</div>

<pre id="adminChatStatus" style="margin-top:10px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Loading admin chat...</pre>

<div style="margin-top:10px;overflow:auto">
  <table id="adminChatTable">
    <thead><tr><th>When</th><th>Author</th><th>Message</th><th>Action</th></tr></thead>
    <tbody></tbody>
  </table>
</div>
`;

const postBtn = document.getElementById('postAdminChatBtn');
if (postBtn) postBtn.addEventListener('click', postAdminChatMessage);
const refreshBtn = document.getElementById('refreshAdminChatBtn');
if (refreshBtn) refreshBtn.addEventListener('click', loadAdminChatMessages);
loadAdminChatMessages();

}

if(page === "calendar"){

document.getElementById("content").innerHTML = `
<h2>Event Calendar</h2>
<p>Schedule command events and reminders.</p>

<div style="margin-top:10px;border:1px solid rgba(255,255,255,.2);padding:10px;background:rgba(0,0,0,.15)">
  <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
    <input id="calendarDate" type="date">
    <input id="calendarTime" type="time" value="09:00">
    <select id="calendarEventTimezone" style="min-width:220px"></select>
    <input id="calendarTitle" type="text" placeholder="Event title" style="min-width:220px;flex:1">
  </div>
  <textarea id="calendarNote" rows="2" placeholder="Optional notes" style="width:100%;margin-top:8px"></textarea>
  <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;">
    <button id="calendarAddBtn">Add Event</button>
    <button id="calendarRefreshBtn">Refresh</button>
  </div>
</div>

<pre id="calendarStatus" style="margin-top:10px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Loading events...</pre>

<div style="margin-top:10px;overflow:auto">
  <table id="calendarTable">
    <thead><tr><th>Date</th><th>Time</th><th>Event</th><th>Notes</th><th>By</th><th>Source TZ</th><th>Action</th></tr></thead>
    <tbody></tbody>
  </table>
</div>
`;

initCalendarTimezoneSelectors();

const addBtn = document.getElementById('calendarAddBtn');
if (addBtn) addBtn.addEventListener('click', addCalendarEvent);
const refreshBtn = document.getElementById('calendarRefreshBtn');
if (refreshBtn) refreshBtn.addEventListener('click', loadCalendarEvents);
loadCalendarEvents();

}

if(page === "promotion-recommendations"){

if(!hasLeadershipAccessClient(currentUser)){
loadPage('dashboard');
return;
}

document.getElementById("content").innerHTML = `
<h2>Promotion Recommendations</h2>
<p>Submit a recommendation for High Command review.</p>

<div style="margin-top:10px;border:1px solid rgba(255,255,255,.2);padding:10px;background:rgba(0,0,0,.15)">
  <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
    <select id="promotionOfficerSelect" style="min-width:300px;flex:1"></select>
    <select id="promotionSuggestedRank" style="min-width:220px;">
      <option value="">Suggested Rank...</option>
      <option value="Officer">OFFICER</option>
      <option value="Senior Officer">SENIOR OFFICER</option>
      <option value="Corporal">CORPORAL</option>
      <option value="Sergeant">SERGEANT</option>
    </select>
  </div>
  <textarea id="promotionNotes" rows="4" placeholder="Recommendation notes and reasons" style="width:100%;margin-top:8px"></textarea>
  <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;">
    <button id="submitPromotionRecBtn">Submit Recommendation</button>
    <button id="refreshPromotionRecBtn">Refresh</button>
  </div>
</div>

<pre id="promotionRecStatus" style="margin-top:10px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Loading recommendations...</pre>

<div style="margin-top:10px;overflow:auto">
  <table id="promotionRecTable">
    <thead><tr><th>Officer</th><th>Current Rank</th><th>Suggested Rank</th><th>Submitted By</th><th>Date</th><th>Status</th></tr></thead>
    <tbody></tbody>
  </table>
</div>
`;

const submitBtn = document.getElementById('submitPromotionRecBtn');
if (submitBtn) submitBtn.addEventListener('click', submitPromotionRecommendation);
const refreshBtn = document.getElementById('refreshPromotionRecBtn');
if (refreshBtn) refreshBtn.addEventListener('click', loadPromotionRecommendations);

loadPromotionOfficerOptions();
loadPromotionRecommendations();

}

if(page === "high-command-approval"){

if(!hasLeadershipAccessClient(currentUser)){
loadPage('dashboard');
return;
}

document.getElementById("content").innerHTML = `
<h2>High Command Approval</h2>
<p>Review, approve, deny, or delete promotion recommendations.</p>

<div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
  <label>Status</label>
  <select id="promotionApprovalStatusFilter">
    <option value="under_review">Under Review</option>
    <option value="all">All</option>
    <option value="approved">Approved</option>
    <option value="denied">Denied</option>
  </select>
  <button id="refreshPromotionApprovalBtn">Refresh</button>
</div>

<pre id="promotionApprovalStatus" style="margin-top:10px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Loading approval queue...</pre>

<div style="margin-top:10px;overflow:auto">
  <table id="promotionApprovalTable">
    <thead><tr><th>Officer</th><th>Current</th><th>Suggested</th><th>Notes</th><th>Submitted By</th><th>Date</th><th>Status</th><th>Actions</th></tr></thead>
    <tbody></tbody>
  </table>
</div>
`;

const refreshBtn = document.getElementById('refreshPromotionApprovalBtn');
if (refreshBtn) refreshBtn.addEventListener('click', loadPromotionApprovalQueue);
const statusFilter = document.getElementById('promotionApprovalStatusFilter');
if (statusFilter) statusFilter.addEventListener('change', loadPromotionApprovalQueue);
loadPromotionApprovalQueue();

}

/* FTO */

if(page === "fto"){

if(!hasLeadershipAccessClient(currentUser)){
loadPage('dashboard');
return;
}

document.getElementById("content").innerHTML = `
<h2>Field Training Officer</h2>
<p><b>HEAD OF TRAINING CAPTAIN JASON GREEN (NORA DIVISION)</b></p>

<div style="margin:10px 0;border:1px solid rgba(255,255,255,.2);padding:10px;background:rgba(0,0,0,.15)">
  <b>Add FTO Officer</b>
  <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
    <select id="ftoOfficerSelect" style="min-width:260px;flex:1"></select>
    <button id="addFtoOfficerBtn">Add</button>
    <button id="refreshFtoBtn">Refresh</button>
  </div>
</div>

<pre id="ftoStatus" style="margin-top:10px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Loading FTO roster...</pre>

<div style="margin-top:10px;overflow:auto">
  <table id="ftoTable">
    <thead>
      <tr>
        <th>ID</th>
        <th>Name</th>
        <th>Callsign</th>
        <th>Rank</th>
        <th>Division</th>
        <th>Added By</th>
        <th>Added At</th>
        <th>Action</th>
      </tr>
    </thead>
    <tbody></tbody>
  </table>
</div>
`;

const addBtn = document.getElementById('addFtoOfficerBtn');
if (addBtn) addBtn.addEventListener('click', addFtoOfficer);

const refreshBtn = document.getElementById('refreshFtoBtn');
if (refreshBtn) refreshBtn.addEventListener('click', loadFtoPage);

if (isLoggedIn()) {
  ensureDataTabsSynced()
    .then(() => loadFtoPage())
    .catch(() => loadFtoPage());
} else {
  loadFtoPage();
}

}

/* ADMIN */

if(page === "admin"){

if(!hasAdminAccessClient(currentUser)){
loadPage('dashboard');
return;
}

document.getElementById("content").innerHTML = `
<h2>Admin Access Control</h2>
<p>Grant or revoke admin access for users in Command_Users.</p>

<div style="margin:8px 0">
  <button id="refreshAdminUsersBtn">Refresh</button>
</div>

<pre id="adminStatus" style="margin-top:8px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Loading users...</pre>

<div style="margin-top:10px;overflow:auto">
  <table id="adminUsersTable">
    <thead>
      <tr>
        <th>Name</th>
        <th>Email</th>
        <th>Account</th>
        <th>Role</th>
        <th>Action</th>
      </tr>
    </thead>
    <tbody></tbody>
  </table>
</div>
`;

const refreshBtn = document.getElementById('refreshAdminUsersBtn');
if (refreshBtn) refreshBtn.addEventListener('click', loadAdminUsers);
loadAdminUsers();

}


/* ACCOUNT */

if(page === "account"){

const canAdminReset = hasAdminAccessClient(currentUser);

document.getElementById("content").innerHTML = `
<h2>My Account</h2>
<p>Personal command profile information.</p>
<table>
  <thead>
    <tr><th>Field</th><th>Value</th></tr>
  </thead>
  <tbody>
    <tr><td>Email</td><td>${(currentUser && currentUser.email) || '-'}</td></tr>
    <tr><td>Character Name</td><td>${(currentUser && currentUser.characterName) || '-'}</td></tr>
    <tr><td>Rank</td><td>${(currentUser && currentUser.rank ? currentUser.rank.toUpperCase() : '-')}</td></tr>
    <tr><td>Role</td><td>${(currentUser && currentUser.role ? currentUser.role.toUpperCase() : '-')}</td></tr>
  </tbody>
</table>

<div style="margin-top:16px;border:1px solid rgba(255,255,255,.2);padding:10px;background:rgba(0,0,0,.15)">
  <b>Change Password</b>
  <div style="margin-top:8px;display:grid;grid-template-columns:1fr;gap:6px;max-width:420px;">
    <input id="oldPassword" type="password" placeholder="Old password">
    <input id="newPassword" type="password" placeholder="New password">
    <input id="confirmPassword" type="password" placeholder="Confirm new password">
  </div>
  <button id="changePasswordBtn" style="margin-top:8px;">Update Password</button>
</div>

${canAdminReset ? `
<div style="margin-top:14px;border:1px solid rgba(255,255,255,.2);padding:10px;background:rgba(0,0,0,.15)">
  <b>Admin Password Reset</b>
  <div style="margin-top:8px;display:grid;grid-template-columns:1fr;gap:6px;max-width:420px;">
    <input id="resetEmail" type="email" placeholder="Target account email">
    <input id="resetNewPassword" type="password" placeholder="Temporary new password">
  </div>
  <button id="adminResetPasswordBtn" style="margin-top:8px;">Reset User Password</button>
</div>
` : ''}
<pre id="accountStatus" style="margin-top:14px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">Logged in.</pre>
`;
document.getElementById('changePasswordBtn').addEventListener('click', changePassword);

const adminResetBtn = document.getElementById('adminResetPasswordBtn');
if (adminResetBtn) adminResetBtn.addEventListener('click', adminResetPassword);

}

/* OFFICER ROSTER */

if(page === "roster"){

document.getElementById("content").innerHTML = `
<h2>Officer Roster</h2>
<p id="rosterStatus">Loading roster...</p>
<div style="margin:8px 0">
  <button id="addOfficer">Add Officer</button>
  <button id="refreshRoster">Refresh</button>
  <button id="syncSheets">Sync Now (Manual)</button>
</div>
<table id="rosterTable">
<thead>
<tr>
<th>ID</th>
<th>Name</th>
<th>Callsign</th>
<th>Rank</th>
<th>Division</th>
<th>Status</th>
<th>Activity</th>
<th>Notes</th>
<th>Actions</th>
</tr>
</thead>
<tbody></tbody>
</table>
<pre id="rosterDebug" style="margin-top:8px;color:#800;white-space:pre-wrap"></pre>
`;

loadRoster();
if (isLoggedIn()) {
  ensureDataTabsSynced()
    .then(() => loadRoster())
    .catch(() => {
      // Keep current roster state if auto-sync fails.
    });
}

document.getElementById('refreshRoster').addEventListener('click', loadRoster);
if (isLoggedIn()) {
  document.getElementById('addOfficer').addEventListener('click', () => showOfficerForm());
} else {
  document.getElementById('addOfficer').disabled = true;
  document.getElementById('addOfficer').title = 'Login required';
}
document.getElementById('syncSheets').addEventListener('click', syncGoogleSheets);

}


if(page === "sheettabs"){

document.getElementById("content").innerHTML = `
<h2>Sheet Tabs</h2>
<p>Imported tabs from Google Sheets will appear here.</p>
<div id="sheetTabsList">Loading tabs...</div>
<div id="sheetTabData" style="margin-top:16px"></div>
`;

loadSheetTabs();

}

}


/* ROSTER LOADER */


async function loadRoster(){
  const statusEl = document.getElementById('rosterStatus');
  statusEl.textContent = 'Loading roster...';

  try {
    const response = await fetch('/api/roster');
    const data = await response.json();
    const table = document.querySelector('#rosterTable tbody');
    table.innerHTML = '';
    let count = 0;
    const pick = (obj, keys) => {
      for (const key of keys) {
        const val = String((obj && obj[key]) || '').trim();
        if (val) return val;
      }
      return '';
    };

    (data || []).forEach((item) => {
      const id = pick(item, ['ID', 'id', 'Officer_ID', 'officer_id']);
      let name = pick(item, ['Name', 'name', 'RP_Name', 'rp_name', 'Officer_Name', 'officer_name']);
      let callsign = pick(item, ['Callsign', 'callsign', 'Call_Sign', 'call_sign']);
      const rank = pick(item, ['Rank', 'rank']);
      const division = pick(item, ['Division', 'division', 'Unit', 'unit']);
      const status = pick(item, ['Status', 'status']);
      const activityStatus = pick(item, ['Activity_Status', 'activity_status', 'ActivityStatus', 'activitystatus', 'Activity Status', 'activity status']);
      const notes = pick(item, ['Notes', 'notes', 'Officer_Notes', 'officer_notes', 'Comments', 'comments']);

      if (!name && callsign && /[a-z]/i.test(callsign) && !/\d/.test(callsign)) {
        name = callsign;
        callsign = '';
      }

      const hasIdentity = !!(id || name || callsign);
      if (!hasIdentity) return;

      const displayName = name || '(No Name)';
      const safeIdAttr = String(id || '').replace(/"/g, '&quot;');
      const safeNameAttr = String(name || '').replace(/"/g, '&quot;');
      const safeCallsignAttr = String(callsign || '').replace(/"/g, '&quot;');

      const tr = document.createElement('tr');
      const actionButtons = isLoggedIn()
        ? `<button data-id="${safeIdAttr}" data-name="${safeNameAttr}" data-callsign="${safeCallsignAttr}" onclick="openOfficerProfileFromRow(this)">Profile</button>
            <button onclick="editOfficer('${id}')">Edit</button>
            <button onclick="deleteOfficer('${id}')">Delete</button>`
        : `<button data-id="${safeIdAttr}" data-name="${safeNameAttr}" data-callsign="${safeCallsignAttr}" onclick="openOfficerProfileFromRow(this)">Profile</button>`;
      tr.innerHTML = `
        <td>
          <button
            class="profile-link"
            data-id="${safeIdAttr}"
            data-name="${safeNameAttr}"
            data-callsign="${safeCallsignAttr}"
            onclick="openOfficerProfileFromRow(this)">
            ${id || '(No ID)'}
          </button>
        </td>
        <td>${displayName}</td>
        <td>${callsign}</td>
        <td>${rank ? rank.toUpperCase() : ''}</td>
        <td>${division}</td>
        <td>${escapeHtml(status || '-')}</td>
        <td>${escapeHtml(activityStatus || '-')}</td>
        <td title="${escapeHtml(notes || '')}">${escapeHtml(notes || '-')}</td>
        <td class="actions-cell">
          <div class="row-actions">
            ${actionButtons}
          </div>
        </td>
      `;
      table.appendChild(tr);
      count++;
    });
    statusEl.textContent = 'Roster loaded (' + count + ' officers)';
  } catch (err) {
    console.error('Error loading roster:', err);
    statusEl.textContent = 'Error: ' + err.message;
  }
}

async function openOfficerProfileFromRow(el) {
  const id = String((el && el.dataset && el.dataset.id) || '').trim();
  const name = String((el && el.dataset && el.dataset.name) || '').trim();
  const callsign = String((el && el.dataset && el.dataset.callsign) || '').trim();

  return openOfficerProfile({ id, name, callsign });
}

async function openOfficerProfile(criteria) {
  try {
    const response = await fetch('/api/roster');
    const data = await response.json();
    if (!response.ok) throw new Error('Failed to load roster data');

    const pick = (obj, keys) => {
      for (const key of keys) {
        const val = String((obj && obj[key]) || '').trim();
        if (val) return val;
      }
      return '';
    };

    const id = String((criteria && criteria.id) || '').trim();
    const name = String((criteria && criteria.name) || '').trim();
    const callsign = String((criteria && criteria.callsign) || '').trim();

    let officer = null;
    if (id) {
      officer = (data || []).find(x => {
        const xId = pick(x, ['ID', 'id', 'Officer_ID', 'officer_id']);
        return String(xId).trim() === id;
      });
    }
    if (!officer && name) {
      officer = (data || []).find(x => {
        const xName = pick(x, ['Name', 'name', 'RP_Name', 'rp_name', 'Officer_Name', 'officer_name']);
        return String(xName).trim().toLowerCase() === name.toLowerCase();
      });
    }
    if (!officer && callsign) {
      officer = (data || []).find(x => {
        const xCallsign = pick(x, ['Callsign', 'callsign', 'Call_Sign', 'call_sign']);
        return String(xCallsign).trim().toLowerCase() === callsign.toLowerCase();
      });
    }

    if (!officer) {
      alert('Officer profile not found. Try refreshing the roster.');
      return;
    }

    const enrichedOfficer = await enrichOfficerProfileData(officer);
    renderOfficerProfile(enrichedOfficer);
  } catch (err) {
    alert('Failed to open profile: ' + err.message);
  }
}

function pickOfficerField(obj, keys) {
  for (const key of keys) {
    const val = String((obj && obj[key]) || '').trim();
    if (val) return val;
  }
  return '';
}

function computeColumnsCTFromRecord(record) {
  const output = {};
  const keys = Object.keys(record || {});

  // Columns C through T are zero-based indexes 2..19.
  for (let i = 2; i <= 19; i++) {
    const key = keys[i] || ('Column ' + String.fromCharCode(65 + i));
    output[key] = String((record && record[key]) || '').trim();
  }

  return output;
}

function matchRawRecordToOfficer(raw, officer) {
  const rawId = pickOfficerField(raw, ['ID', 'id', 'Officer_ID', 'officer_id']);
  const rawName = pickOfficerField(raw, ['Name', 'name', 'RP_Name', 'rp_name', 'Officer_Name', 'officer_name']);
  const rawCallsign = pickOfficerField(raw, ['Callsign', 'callsign', 'Call_Sign', 'call_sign']);

  const offId = pickOfficerField(officer, ['ID', 'id', 'Officer_ID', 'officer_id']);
  const offName = pickOfficerField(officer, ['Name', 'name', 'RP_Name', 'rp_name', 'Officer_Name', 'officer_name']).toLowerCase();
  const offCallsign = pickOfficerField(officer, ['Callsign', 'callsign', 'Call_Sign', 'call_sign']).toLowerCase();

  if (offId && rawId && String(rawId).trim() === offId) return true;
  if (offName && rawName && String(rawName).trim().toLowerCase() === offName) return true;
  if (offCallsign && rawCallsign && String(rawCallsign).trim().toLowerCase() === offCallsign) return true;
  return false;
}

async function enrichOfficerProfileData(officer) {
  const hasImported = officer && officer.ImportedFields && Object.keys(officer.ImportedFields).length > 0;
  const hasCT = officer && officer.ColumnsCT && Object.keys(officer.ColumnsCT).length > 0;
  if (hasImported && hasCT) return officer;

  try {
    const response = await fetch('/api/sheets/tab/roster');
    const rows = await response.json();
    if (!response.ok || !Array.isArray(rows)) return officer;

    const rawMatch = rows.find(r => matchRawRecordToOfficer(r, officer));
    if (!rawMatch) return officer;

    const merged = Object.assign({}, officer);
    if (!hasImported) merged.ImportedFields = rawMatch;
    if (!hasCT) merged.ColumnsCT = computeColumnsCTFromRecord(rawMatch);
    return merged;
  } catch (e) {
    return officer;
  }
}

function renderOfficerProfile(officer) {
  const profileId = pickOfficerField(officer, ['ID', 'id', 'Officer_ID', 'officer_id']);
  const profileName = pickOfficerField(officer, ['Name', 'name', 'RP_Name', 'rp_name', 'Officer_Name', 'officer_name']);
  const profileRank = pickOfficerField(officer, ['Rank', 'rank', 'officer_rank']);
  const profileDivision = pickOfficerField(officer, ['Division', 'division', 'Unit', 'unit']);
  const profileStatus = pickOfficerField(officer, ['Status', 'status']);
  const profileActivityStatus = pickOfficerField(officer, ['Activity_Status', 'activity_status', 'ActivityStatus', 'activitystatus', 'Activity Status', 'activity status']);
  const profileFto = pickOfficerField(officer, ['IsFTO', 'is_fto', 'FTO', 'fto']);
  const isFtoActive = ['yes', 'true', '1', 'fto', 'active'].includes(String(profileFto || '').trim().toLowerCase());
  const profileNotes = pickOfficerField(officer, ['Notes', 'notes', 'Officer_Notes', 'officer_notes', 'Comments', 'comments']);
  // Only allow up to Sergeant for promotion
  const rankOptions = [
    'Cadet',
    'Officer',
    'Senior Officer',
    'Corporal',
    'Sergeant'
  ];
  const currentRank = String(profileRank || '').trim();
  const rankOptionSet = new Set(rankOptions.map((r) => r.toLowerCase()));
  if (currentRank && !rankOptionSet.has(currentRank.toLowerCase())) {
    rankOptions.unshift(currentRank);
  }
  const rankOptionsHtml = rankOptions.map((rankLabel) => {
    const selected = currentRank.toLowerCase() === String(rankLabel || '').toLowerCase() ? ' selected' : '';
    return '<option value="' + escapeHtml(rankLabel) + '"' + selected + '>' + escapeHtml(rankLabel.toUpperCase()) + '</option>';
  }).join('');
  const canPromote = !!currentUser && hasLeadershipAccessClient(currentUser);
  const promotionEditor = (canPromote && profileId)
    ? `
    <div style="margin-top:12px;border:1px solid rgba(255,255,255,.2);padding:10px;background:rgba(0,0,0,.15)">
      <b>Promotion Controls</b><br>
      <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;max-width:820px;">
        <label style="min-width:72px;">Rank</label>
        <select id="promotionRankInput" style="min-width:220px;flex:1;">${rankOptionsHtml}</select>
      </div>
      <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;max-width:820px;">
        <label style="min-width:72px;">Division</label>
        <input id="promotionDivisionInput" type="text" value="${escapeHtml(profileDivision || '')}" style="min-width:220px;flex:1;">
      </div>
      <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;max-width:820px;">
        <label style="min-width:72px;">Reason</label>
        <input id="promotionReasonInput" type="text" placeholder="Optional promotion note" style="min-width:220px;flex:1;">
      </div>
      <button id="savePromotionBtn" style="margin-top:8px;">Apply Promotion</button>
      <pre id="promotionStatus" style="margin-top:8px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:8px;border:1px solid rgba(255,255,255,.2)">Ready.</pre>
    </div>`
    : '';
  const notesEditor = (isLoggedIn() && profileId)
    ? `
    <div style="margin-top:12px;border:1px solid rgba(255,255,255,.2);padding:10px;background:rgba(0,0,0,.15)">
      <b>Edit Notes</b><br>
      <textarea id="profileNotesInput" rows="5" style="width:100%;max-width:700px;margin-top:8px;">${escapeHtml(profileNotes || '')}</textarea><br>
      <button id="saveOfficerNotesBtn" style="margin-top:8px;">Save Notes</button>
    </div>`
    : '';

  const imported = (officer && officer.ImportedFields && typeof officer.ImportedFields === 'object')
    ? officer.ImportedFields
    : ((officer && officer.ColumnsCT && typeof officer.ColumnsCT === 'object') ? officer.ColumnsCT : {});
  const importedRows = Object.keys(imported)
    .filter((k) => String(imported[k] || '').trim() !== '')
    .map((k) => [k, imported[k] || '-']);

  const rowsToTable = (rows) => {
    if (!rows.length) return '<p>No data available.</p>';
    let html = '<table><thead><tr><th>Field</th><th>Value</th></tr></thead><tbody>';
    rows.forEach((r) => {
      html += '<tr><td>' + r[0] + '</td><td>' + r[1] + '</td></tr>';
    });
    html += '</tbody></table>';
    return html;
  };

  document.getElementById('content').innerHTML = `
    <h2>Officer Profile</h2>
    <div style="margin:8px 0 16px 0;">
      <button onclick="loadPage('roster')">Back to Roster</button>
    </div>

    <p><b>ID:</b> ${profileId || '-'}</p>
    <p><b>Name:</b> ${profileName || '-'}</p>
    <p><b>Rank:</b> ${escapeHtml(profileRank || '-')}</p>
    <p><b>Division:</b> ${escapeHtml(profileDivision || '-')}</p>
    <p><b>FTO:</b> ${isFtoActive ? '<span style="font-weight:700;color:#ffd166">ACTIVE TRAINER</span>' : '-'}</p>
    <p><b>Status:</b> ${escapeHtml(profileStatus || '-')}</p>
    <p><b>Activity Status:</b> ${escapeHtml(profileActivityStatus || '-')}</p>
    <p><b>Notes:</b></p>
    <pre style="margin-top:6px;white-space:pre-wrap;background:rgba(0,0,0,.2);padding:10px;border:1px solid rgba(255,255,255,.2)">${escapeHtml(profileNotes || 'No notes added.')}</pre>
    ${promotionEditor}
    ${notesEditor}

    <h3 style="margin-top:18px;">All Imported Officer Data</h3>
    ${rowsToTable(importedRows)}
  `;

  if (isLoggedIn() && profileId) {
    const saveBtn = document.getElementById('saveOfficerNotesBtn');
    if (saveBtn) saveBtn.addEventListener('click', () => saveOfficerNotes(profileId));
  }
  if (canPromote && profileId) {
    const promoteBtn = document.getElementById('savePromotionBtn');
    if (promoteBtn) promoteBtn.addEventListener('click', () => promoteOfficerProfile(profileId));
  }
}

async function promoteOfficerProfile(officerId) {
  if (!isLoggedIn()) {
    alert('Login required for promotions.');
    return;
  }

  const id = String(officerId || '').trim();
  if (!id) {
    alert('Cannot promote: missing officer ID.');
    return;
  }

  const rank = String(((document.getElementById('promotionRankInput') || {}).value) || '').trim();
  const division = String(((document.getElementById('promotionDivisionInput') || {}).value) || '').trim();
  const reason = String(((document.getElementById('promotionReasonInput') || {}).value) || '').trim();
  const status = document.getElementById('promotionStatus');

  if (!rank && !division) {
    if (status) status.textContent = 'Enter rank and/or division first.';
    return;
  }

  try {
    const response = await authFetch('/api/roster/' + encodeURIComponent(id) + '/promote', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ rank, division, reason })
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.error || 'Promotion failed');

    if (status) status.textContent = payload.message || 'Promotion updated.';
    await openOfficerProfile({ id });
  } catch (err) {
    if (status) status.textContent = 'Promotion failed: ' + err.message;
  }
}

async function saveOfficerNotes(officerId, explicitNotes) {
  if (!isLoggedIn()) {
    alert('Login required for roster edits.');
    return;
  }

  const id = String(officerId || '').trim();
  if (!id) {
    alert('Cannot save notes: missing officer ID.');
    return;
  }

  const notesValue = (typeof explicitNotes === 'string')
    ? explicitNotes
    : String(((document.getElementById('profileNotesInput') || {}).value) || '');

  try {
    const response = await authFetch('/api/roster/' + encodeURIComponent(id), {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ Notes: notesValue })
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.error || 'Save notes failed');
    await openOfficerProfile({ id });
  } catch (err) {
    alert('Save notes failed: ' + err.message);
  }
}

async function openOfficerNotes(id) {
  if (!isLoggedIn()) {
    alert('Login required for roster edits.');
    return;
  }
  try {
    await openOfficerProfile({ id: String(id || '').trim() });
  } catch (err) {
    alert('Unable to open notes editor: ' + err.message);
  }
}

function showOfficerForm(id = null, data = {}) {
  const statusValue = pickOfficerField(data, ['Status', 'status']);
  const activityStatusValue = pickOfficerField(data, ['Activity_Status', 'activity_status', 'ActivityStatus', 'activitystatus', 'Activity Status', 'activity status']);
  const emailValue = pickOfficerField(data, ['Email', 'email', 'Google_Email', 'google_email', 'Email_Address', 'email_address']);
  const accessRoleValue = normalizeAccessRoleClient(pickOfficerField(data, ['Access_Role', 'access_role', 'Role', 'role', 'RoleOverride', 'roleOverride']));
  const canManageAccess = !!currentUser && hasLeadershipAccessClient(currentUser);

  const formHtml = `
    <div id="officerForm" style="position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:white;color:#111;padding:20px;border:1px solid #ccc;z-index:1000;box-shadow:0 0 10px rgba(0,0,0,0.3);min-width:260px;">
      <h3>${id ? 'Edit' : 'Add'} Officer</h3>
      <label>ID: <input id="formID" value="${data.ID || ''}"></label><br>
      <label>Name: <input id="formName" value="${data.Name || ''}"></label><br>
      <label>Callsign: <input id="formCallsign" value="${data.Callsign || ''}"></label><br>
      <label>Rank: <input id="formRank" value="${data.Rank || ''}"></label><br>
      <label>Division: <input id="formDivision" value="${data.Division || ''}"></label><br>
      <label>Status: <input id="formStatus" value="${escapeHtml(statusValue || '')}"></label><br>
      <label>Activity Status: <input id="formActivityStatus" value="${escapeHtml(activityStatusValue || '')}"></label><br>
      <label>Email: <input id="formEmail" value="${escapeHtml(emailValue || '')}"></label><br>
      ${canManageAccess ? `
      <label>Portal Access Role:
        <select id="formAccessRole">
          <option value="command" ${accessRoleValue === 'command' ? 'selected' : ''}>Command</option>
          <option value="commander" ${accessRoleValue === 'commander' ? 'selected' : ''}>Commander</option>
          <option value="chief" ${accessRoleValue === 'chief' ? 'selected' : ''}>Chief</option>
          <option value="admin" ${accessRoleValue === 'admin' ? 'selected' : ''}>Admin</option>
        </select>
      </label><br>
      <small>Chiefs, commanders, and admins can promote/demote portal access roles.</small><br>` : ''}
      <label>Notes:</label><br>
      <textarea id="formNotes" rows="4" style="width:100%;max-width:380px;">${String(data.Notes || '')}</textarea><br>
      <button id="saveOfficer">Save</button>
      <button onclick="closeForm()">Cancel</button>
    </div>
  `;
  document.body.insertAdjacentHTML('beforeend', formHtml);
  document.getElementById('saveOfficer').addEventListener('click', () => saveOfficer(id));
}

function closeForm() {
  const form = document.getElementById('officerForm');
  if (form) form.remove();
}

async function saveOfficer(id) {
  if (!isLoggedIn()) {
    alert('Login required for roster edits.');
    return;
  }
  const data = {
    ID: document.getElementById('formID').value,
    Name: document.getElementById('formName').value,
    Callsign: document.getElementById('formCallsign').value,
    Rank: document.getElementById('formRank').value,
    Division: document.getElementById('formDivision').value,
    Status: document.getElementById('formStatus').value,
    Activity_Status: document.getElementById('formActivityStatus').value,
    Email: document.getElementById('formEmail').value,
    Notes: document.getElementById('formNotes').value
  };
  try {
    const method = id ? 'PUT' : 'POST';
    const url = id ? `/api/roster/${id}` : '/api/roster';
    const response = await authFetch(url, {
      method: method,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.error || 'Save failed');

    const canManageAccess = !!currentUser && hasLeadershipAccessClient(currentUser);
    if (canManageAccess) {
      const targetEmail = String(data.Email || '').trim();
      const accessRoleEl = document.getElementById('formAccessRole');
      const targetRole = normalizeAccessRoleClient(accessRoleEl ? accessRoleEl.value : 'command');
      if (targetEmail) {
        const accessRes = await authFetch('/api/admin/set-role', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ email: targetEmail, role: targetRole })
        });
        const accessPayload = await accessRes.json().catch(() => ({}));
        if (!accessRes.ok) throw new Error(accessPayload.error || 'Officer saved, but access update failed.');
      }
    }

    closeForm();
    loadRoster();
  } catch (err) {
    alert('Error saving: ' + err.message);
  }
}

async function editOfficer(id) {
  if (!isLoggedIn()) {
    alert('Login required for roster edits.');
    return;
  }
  try {
    const response = await fetch(`/api/roster?id=${id}`);
    const data = await response.json();
    const item = data.find(x => pickOfficerField(x, ['ID', 'id', 'Officer_ID', 'officer_id']) === String(id));
    if (item) {
      showOfficerForm(id, item);
    } else {
      alert('Officer not found');
    }
  } catch (err) {
    alert('Error editing: ' + err.message);
  }
}

async function deleteOfficer(id) {
  if (!isLoggedIn()) {
    alert('Login required for roster edits.');
    return;
  }
  if (!confirm('Delete this officer?')) return;
  try {
    const response = await authFetch(`/api/roster/${id}`, { method: 'DELETE' });
    if (!response.ok) throw new Error('Delete failed');
    loadRoster();
  } catch (err) {
    alert('Error deleting: ' + err.message);
  }
}

async function loadFtoPage() {
  await Promise.all([loadFtoRosterOptions(), loadFtoList()]);
}

async function loadFtoRosterOptions() {
  const select = document.getElementById('ftoOfficerSelect');
  if (!select) return;
  select.innerHTML = '<option value="">Loading roster...</option>';
  try {
    const response = await fetch('/api/roster');
    const data = await response.json();
    if (!response.ok) throw new Error('Failed to load roster');

    const rows = Array.isArray(data) ? data : [];
    const options = rows.map((item) => {
      const id = pickOfficerField(item, ['ID', 'id', 'Officer_ID', 'officer_id']);
      const name = pickOfficerField(item, ['Name', 'name', 'RP_Name', 'rp_name', 'Officer_Name', 'officer_name']);
      const callsign = pickOfficerField(item, ['Callsign', 'callsign', 'Call_Sign', 'call_sign']);
      const rank = pickOfficerField(item, ['Rank', 'rank']);
      const division = pickOfficerField(item, ['Division', 'division', 'Unit', 'unit']);
      if (!id) return null;
      return {
        id,
        label: (rank ? rank + ' ' : '') + (name || '(No Name)') + (callsign ? ' | ' + callsign : '') + (division ? ' | ' + division : '')
      };
    }).filter(Boolean).sort((a, b) => a.label.localeCompare(b.label));

    select.innerHTML = '<option value="">Select officer from roster...</option>';
    options.forEach((o) => {
      const opt = document.createElement('option');
      opt.value = o.id;
      opt.textContent = o.label;
      select.appendChild(opt);
    });
  } catch (err) {
    select.innerHTML = '<option value="">Roster unavailable</option>';
  }
}

async function loadFtoList() {
  const status = document.getElementById('ftoStatus');
  const body = document.querySelector('#ftoTable tbody');
  if (!body) return;

  const currentLoadToken = ++ftoListLoadToken;
  body.innerHTML = '';
  try {
    const response = await authFetch('/api/fto');
    const data = await response.json();
    if (currentLoadToken !== ftoListLoadToken) return;
    if (!response.ok) throw new Error(data.error || 'Failed to load FTO list');

    const items = Array.isArray(data.items) ? data.items : [];
    if (!items.length) {
      if (status) status.textContent = 'No FTO officers added yet.';
      return;
    }

    items.forEach((item) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(item.officerId || '-')}</td>
        <td>${escapeHtml(item.name || '-')}</td>
        <td>${escapeHtml(item.callsign || '-')}</td>
        <td>${escapeHtml(item.rank ? item.rank.toUpperCase() : '-')}</td>
        <td>${escapeHtml(item.division || '-')}</td>
        <td>${escapeHtml(item.addedBy || '-')}</td>
        <td>${escapeHtml(item.addedAt ? new Date(item.addedAt).toLocaleString() : '-')}</td>
        <td><button onclick="removeFtoOfficer('${String(item.officerId || '').replace(/'/g, "\\'")}')">Remove</button></td>
      `;
      body.appendChild(tr);
    });
    if (status) status.textContent = 'FTO roster loaded (' + items.length + ' officers).';
  } catch (err) {
    if (status) status.textContent = 'Error: ' + err.message;
  }
}

async function addFtoOfficer() {
  const select = document.getElementById('ftoOfficerSelect');
  const status = document.getElementById('ftoStatus');
  const officerId = String((select && select.value) || '').trim();
  if (!officerId) {
    if (status) status.textContent = 'Select an officer first.';
    return;
  }
  try {
    const response = await authFetch('/api/fto', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ officerId })
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || 'Failed to add FTO officer');
    if (status) status.textContent = data.message || 'FTO officer added.';
    await loadFtoList();
  } catch (err) {
    if (status) status.textContent = 'Add failed: ' + err.message;
  }
}

async function removeFtoOfficer(officerId) {
  const status = document.getElementById('ftoStatus');
  const id = String(officerId || '').trim();
  if (!id) return;
  if (!confirm('Remove this officer from FTO list?')) return;
  try {
    const response = await authFetch('/api/fto/' + encodeURIComponent(id), { method: 'DELETE' });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || 'Failed to remove FTO officer');
    if (status) status.textContent = data.message || 'FTO officer removed.';
    await loadFtoList();
  } catch (err) {
    if (status) status.textContent = 'Remove failed: ' + err.message;
  }
}

async function loadPromotionOfficerOptions() {
  const select = document.getElementById('promotionOfficerSelect');
  if (!select) return;
  select.innerHTML = '<option value="">Loading roster...</option>';
  try {
    const response = await fetch('/api/roster');
    const data = await response.json();
    if (!response.ok) throw new Error('Failed to load roster');
    const rows = Array.isArray(data) ? data : [];
    const options = rows.map((item) => {
      const id = pickOfficerField(item, ['ID', 'id', 'Officer_ID', 'officer_id']);
      const name = pickOfficerField(item, ['Name', 'name', 'RP_Name', 'rp_name', 'Officer_Name', 'officer_name']);
      const rank = pickOfficerField(item, ['Rank', 'rank']) || '-';
      const callsign = pickOfficerField(item, ['Callsign', 'callsign', 'Call_Sign', 'call_sign']);
      if (!id) return null;
      return {
        id,
        name: name || '(No Name)',
        rank,
        label: (rank ? rank + ' ' : '') + (name || '(No Name)') + (callsign ? ' | ' + callsign : '')
      };
    }).filter(Boolean).sort((a, b) => a.label.localeCompare(b.label));

    select.innerHTML = '<option value="">Select officer...</option>';
    options.forEach((o) => {
      const opt = document.createElement('option');
      opt.value = o.id;
      opt.textContent = o.label;
      opt.dataset.officerName = o.name;
      opt.dataset.currentRank = o.rank;
      select.appendChild(opt);
    });
  } catch (err) {
    select.innerHTML = '<option value="">Roster unavailable</option>';
  }
}

async function submitPromotionRecommendation() {
  const select = document.getElementById('promotionOfficerSelect');
  const rankEl = document.getElementById('promotionSuggestedRank');
  const notesEl = document.getElementById('promotionNotes');
  const status = document.getElementById('promotionRecStatus');
  const officerId = String((select && select.value) || '').trim();
  const option = select && select.options ? select.options[select.selectedIndex] : null;
  const officerName = String((option && option.dataset && option.dataset.officerName) || '').trim();
  const currentRank = String((option && option.dataset && option.dataset.currentRank) || '').trim();
  const suggestedRank = String((rankEl && rankEl.value) || '').trim();
  const notes = String((notesEl && notesEl.value) || '').trim();

  if (!officerId || !officerName || !suggestedRank || !notes) {
    if (status) status.textContent = 'Select an officer, choose a suggested rank, and enter notes.';
    return;
  }

  try {
    const response = await authFetch('/api/promotions/recommendations', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ officerId, officerName, currentRank, suggestedRank, notes })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Submit failed');
    if (notesEl) notesEl.value = '';
    if (status) status.textContent = 'Recommendation submitted and marked under review.';
    await loadPromotionRecommendations();
  } catch (err) {
    if (status) status.textContent = 'Submit failed: ' + err.message;
  }
}

async function loadPromotionRecommendations() {
  const body = document.querySelector('#promotionRecTable tbody');
  const status = document.getElementById('promotionRecStatus');
  if (!body) return;
  body.innerHTML = '<tr><td colspan="6">Loading...</td></tr>';
  try {
    const response = await authFetch('/api/promotions/recommendations?status=all');
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to load recommendations');
    const rows = Array.isArray(data.items) ? data.items : [];
    if (!rows.length) {
      body.innerHTML = '<tr><td colspan="6">No recommendations submitted yet.</td></tr>';
      if (status) status.textContent = 'No recommendations submitted yet.';
      return;
    }
    body.innerHTML = rows.map((r) => {
      const statusClass = 'status-' + sanitizeName(r.status || 'under_review');
      return '<tr>' +
        '<td>' + escapeHtml(r.officerName || '-') + '</td>' +
        '<td>' + escapeHtml((r.currentRank || '-').toUpperCase()) + '</td>' +
        '<td>' + escapeHtml((r.suggestedRank || '-').toUpperCase()) + '</td>' +
        '<td>' + escapeHtml(r.submittedBy || '-') + '</td>' +
        '<td>' + escapeHtml(formatDateTime(r.createdAt || '')) + '</td>' +
        '<td><span class="status-pill ' + statusClass + '">' + escapeHtml(r.status || '-') + '</span></td>' +
      '</tr>';
    }).join('');
    if (status) status.textContent = 'Recommendations: ' + rows.length;
  } catch (err) {
    body.innerHTML = '<tr><td colspan="6">' + escapeHtml(err.message) + '</td></tr>';
    if (status) status.textContent = 'Recommendation load failed: ' + err.message;
  }
}

async function loadPromotionApprovalQueue() {
  const body = document.querySelector('#promotionApprovalTable tbody');
  const status = document.getElementById('promotionApprovalStatus');
  const statusFilter = String((document.getElementById('promotionApprovalStatusFilter') || {}).value || 'under_review').trim();
  if (!body) return;
  body.innerHTML = '<tr><td colspan="8">Loading...</td></tr>';
  try {
    const response = await authFetch('/api/promotions/recommendations?status=' + encodeURIComponent(statusFilter));
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to load approval queue');
    const rows = Array.isArray(data.items) ? data.items : [];
    if (!rows.length) {
      body.innerHTML = '<tr><td colspan="8">No recommendations in this queue.</td></tr>';
      if (status) status.textContent = 'No recommendations in queue.';
      return;
    }
    body.innerHTML = rows.map((r) => {
      const id = String(r.id || '').replace(/'/g, "\\'");
      const statusClass = 'status-' + sanitizeName(r.status || 'under_review');
      return '<tr>' +
        '<td>' + escapeHtml(r.officerName || '-') + '</td>' +
        '<td>' + escapeHtml((r.currentRank || '-').toUpperCase()) + '</td>' +
        '<td>' + escapeHtml((r.suggestedRank || '-').toUpperCase()) + '</td>' +
        '<td title="' + escapeHtml(r.notes || '-') + '">' + escapeHtml(r.notes || '-') + '</td>' +
        '<td>' + escapeHtml(r.submittedBy || '-') + '</td>' +
        '<td>' + escapeHtml(formatDateTime(r.createdAt || '')) + '</td>' +
        '<td><span class="status-pill ' + statusClass + '">' + escapeHtml(r.status || '-') + '</span></td>' +
        '<td class="approval-actions-cell"><div class="approval-actions">' +
          '<button onclick="setPromotionRecommendationStatus(\'' + id + '\',\'approved\')">Approve</button> ' +
          '<button onclick="setPromotionRecommendationStatus(\'' + id + '\',\'denied\')">Deny</button> ' +
          '<button onclick="deletePromotionRecommendation(\'' + id + '\')">Delete</button>' +
        '</div></td>' +
      '</tr>';
    }).join('');
    if (status) status.textContent = 'Queue items: ' + rows.length;
  } catch (err) {
    body.innerHTML = '<tr><td colspan="8">' + escapeHtml(err.message) + '</td></tr>';
    if (status) status.textContent = 'Approval queue load failed: ' + err.message;
  }
}

async function setPromotionRecommendationStatus(id, nextStatus) {
  try {
    const response = await authFetch('/api/promotions/recommendations/' + encodeURIComponent(String(id || '')) + '/status', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ status: nextStatus })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Status update failed');
    await loadPromotionApprovalQueue();
    const recStatus = document.getElementById('promotionRecStatus');
    if (recStatus) await loadPromotionRecommendations();
  } catch (err) {
    alert('Status update failed: ' + err.message);
  }
}

async function deletePromotionRecommendation(id) {
  if (!confirm('Delete this promotion recommendation?')) return;
  try {
    const response = await authFetch('/api/promotions/recommendations/' + encodeURIComponent(String(id || '')), { method: 'DELETE' });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Delete failed');
    await loadPromotionApprovalQueue();
    const recStatus = document.getElementById('promotionRecStatus');
    if (recStatus) await loadPromotionRecommendations();
  } catch (err) {
    alert('Delete failed: ' + err.message);
  }
}

async function createAccount() {
  const email = String(document.getElementById('createEmail').value || '').trim();
  const password = String(document.getElementById('createPassword').value || '');
  const passwordVerify = String(document.getElementById('createPasswordVerify').value || '');
  const status = document.getElementById('accountStatus');

  const looksLikeEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  if (!looksLikeEmail) {
    if (status) status.textContent = 'Create account failed: Enter a valid email address (example: name@domain.com).';
    return;
  }

  if (password !== passwordVerify) {
    if (status) status.textContent = 'Create account failed: Password and Verify Password do not match.';
    return;
  }

  const submitCreate = async () => {
    return fetch('/api/auth/create-account', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password })
    });
  };

  try {
    let response = await submitCreate();
    let data = await response.json();

    // If email lookup fails, trigger a server-side Command_Users ensure and retry once.
    if (!response.ok && /email not found in command_users/i.test(String(data.error || ''))) {
      try {
        await fetch('/api/auth/ensure-command-users', { method: 'POST' });
        response = await submitCreate();
        data = await response.json();
      } catch (e) {
        // Ignore and preserve original create-account error handling.
      }
    }

    if (!response.ok) throw new Error(data.error || 'Create account failed');

    setAuthToken(data.token || '');
    await refreshAuthSession();
    loadPage('dashboard');
    const welcome = 'Account created. Welcome ' + formatUserDisplayName(currentUser) + '.';
    if (status) status.textContent = welcome;
    alert(welcome);
  } catch (err) {
    const base = String(err && err.message || 'Create account failed');
    const hint = /command_users|email not found/i.test(base)
      ? ' If your email is valid, click "Sync Command_Users" and retry. You can also paste a link and click "Link Command_Users".'
      : '';
    if (status) status.textContent = 'Create account failed: ' + base + hint;
  }
}

async function loginAccount() {
  const email = String(document.getElementById('loginEmail').value || '').trim();
  const password = String(document.getElementById('loginPassword').value || '');
  const status = document.getElementById('accountStatus');

  try {
    const response = await fetch('/api/auth/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Login failed');

    setAuthToken(data.token || '');
    await refreshAuthSession();
    loadPage('dashboard');
    const welcome = 'Login successful. Welcome ' + formatUserDisplayName(currentUser) + '.';
    if (status) status.textContent = welcome;
    alert(welcome);
  } catch (err) {
    if (status) status.textContent = 'Login failed: ' + err.message;
  }
}

function findBestField(row, aliases) {
  const keys = Object.keys(row || {});
  const normalized = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  for (const alias of aliases) {
    const wanted = normalized(alias);
    const key = keys.find(k => normalized(k) === wanted);
    if (key) {
      const val = String(row[key] || '').trim();
      if (val) return val;
    }
  }
  return '';
}

function mapAlertRow(source, row) {
  const title = findBestField(row, ['title', 'subject', 'reason', 'type', 'status']) || (source + ' item');
  const person = findBestField(row, ['rp_name', 'name', 'officer_name', 'cadet_name', 'character_name']);
  const detail = findBestField(row, ['notes', 'message', 'comments', 'description', 'summary']);
  const date = findBestField(row, ['date', 'created_date', 'created', 'last_updated', 'timestamp']);

  let line = '[' + source + '] ' + title;
  if (person) line += ' - ' + person;
  if (date) line += ' (' + date + ')';
  if (detail) line += '\n  ' + detail;
  return line;
}

async function loadDashboardAlerts() {
  const box = document.getElementById('dashboardAlerts');

  try {
    const tabsRes = await fetch('/api/sheets/tabs');
    const tabs = await tabsRes.json();
    if (!tabsRes.ok || !Array.isArray(tabs)) throw new Error('Failed to load tab list');

    const sources = [
      { key: 'discipline_records', label: 'Discipline' },
      { key: 'cadet_evaluations', label: 'Cadet Evaluations' },
      { key: 'officer_notes', label: 'Internal Messages' },
      { key: 'internal_messages', label: 'Internal Messages' }
    ];

    const available = sources.filter(s => tabs.includes(s.key));
    if (!available.length) {
      if (box) box.textContent = 'No alert tabs imported yet.';
      setHeaderAlerts([]);
      return;
    }

    const lines = [];
    for (const src of available) {
      const res = await fetch('/api/sheets/tab/' + encodeURIComponent(src.key));
      const rows = await res.json();
      if (!res.ok || !Array.isArray(rows) || !rows.length) continue;

      const alertRows = rows.slice(0, 5);
      if (!alertRows.length) continue;

      alertRows.forEach(r => lines.push(mapAlertRow(src.label, r)));
    }

    if (box) {
      box.textContent = lines.length ? lines.join('\n\n') : ('No alert records found. Tabs checked: ' + available.map(x => x.key).join(', '));
    }
    setHeaderAlerts(lines);
  } catch (err) {
    if (box) box.textContent = 'Alerts unavailable: ' + err.message;
    setHeaderAlerts(['Alerts unavailable: ' + err.message]);
  }
}

async function loadReportsSummary() {
  const box = document.getElementById('reportsSummary');
  if (!box) return;

  try {
    const response = await authFetch('/api/reports/summary');
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to load report summary');

    const summary = data.summary || {};
    const discipline = summary.discipline || {};
    const evaluation = summary.evaluation || {};
    const message = summary.message || {};
    const tabs = Array.isArray(data.configuredTabs) ? data.configuredTabs : [];

    const lines = [
      'Discipline: total ' + (discipline.total || 0) + ' | pending ' + (discipline.pending || 0) + ' | approved ' + (discipline.approved || 0) + ' | denied ' + (discipline.denied || 0),
      'Evaluations: total ' + (evaluation.total || 0) + ' | pending ' + (evaluation.pending || 0) + ' | approved ' + (evaluation.approved || 0) + ' | denied ' + (evaluation.denied || 0),
      'Internal Messages: total ' + (message.total || 0),
      '',
      'Configured report source tabs: ' + (tabs.length ? tabs.join(', ') : 'none')
    ];
    box.textContent = lines.join('\n');
  } catch (err) {
    box.textContent = 'Reports summary unavailable: ' + err.message;
  }
}

function escapeHtml(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatDateTime(value) {
  const raw = String(value || '').trim();
  if (!raw) return '-';
  const d = new Date(raw);
  if (isNaN(d.getTime())) return raw;
  return d.toLocaleString();
}

function showReportDetailsById(reportId) {
  const wantedId = String(reportId || '').trim();
  if (!wantedId) return;

  const item = lastLoadedReportItems.find((x) => String((x && x.id) || '') === wantedId);
  if (!item) return;

  renderReportProfile(item);
}

function reportCanBeApproved(item) {
  const sourceTabText = String((item && item.sourceTab) || '').toLowerCase();
  return (
    item && (
      item.type === 'discipline' ||
      item.type === 'evaluation' ||
      sourceTabText.includes('disciplin') ||
      sourceTabText.includes('eval')
    )
  );
}

function renderReportProfile(item) {
  if (!item) return;

  const rawRow = (item.rawRow && typeof item.rawRow === 'object') ? item.rawRow : {};
  const importedRows = Object.keys(rawRow)
    .filter((k) => String(rawRow[k] || '').trim() !== '')
    .sort()
    .map((k) => [k, rawRow[k] || '-']);
  const rows = importedRows.length ? importedRows : [['(no imported fields)', '-']];

  const rowsToTable = (tableRows) => {
    let html = '<table><thead><tr><th>Field</th><th>Value</th></tr></thead><tbody>';
    tableRows.forEach((r) => {
      html += '<tr><td>' + escapeHtml(r[0]) + '</td><td>' + escapeHtml(r[1]) + '</td></tr>';
    });
    html += '</tbody></table>';
    return html;
  };

  const canApprove = reportCanBeApproved(item);
  const canDelete = !!currentUser && hasLeadershipAccessClient(currentUser);
  const actionsHtml = canApprove
    ? '<button id="reportApproveBtn">Approve</button> ' +
      '<button id="reportDenyBtn">Deny</button> ' +
      '<button id="reportResetBtn">Reset</button>'
    : '<span style="opacity:.85">No approval actions for this report type.</span>';
  const deleteHtml = canDelete ? '<button id="reportDeleteBtn">Delete Report</button>' : '';

  document.getElementById('content').innerHTML = `
    <h2>Report Profile</h2>
    <div style="margin:8px 0 16px 0;display:flex;gap:8px;flex-wrap:wrap;">
      <button id="reportBackBtn">Back to Reports</button>
      ${actionsHtml}
      ${deleteHtml}
    </div>

    <p><b>Type:</b> ${escapeHtml(item.type || '-')}</p>
    <p><b>Subject:</b> ${escapeHtml(item.subject || '-')}</p>
    <p><b>Officer:</b> ${escapeHtml(item.officerName || '-')}</p>
    <p><b>Date:</b> ${escapeHtml(item.reportDate || '-')}</p>
    <p><b>Source Tab:</b> ${escapeHtml(item.sourceTab || '-')}</p>
    <p><b>Status:</b> ${escapeHtml(item.approvalStatus || 'pending')}</p>
    <p><b>Approved By:</b> ${escapeHtml(item.approvedBy || '-')}</p>
    <p><b>Approved At:</b> ${escapeHtml(formatDateTime(item.approvedAt || ''))}</p>

    <h3 style="margin-top:18px;">All Imported Report Data</h3>
    ${rowsToTable(rows)}
  `;

  const backBtn = document.getElementById('reportBackBtn');
  if (backBtn) {
    backBtn.addEventListener('click', () => loadPage('reports'));
  }

  if (canApprove) {
    const approveBtn = document.getElementById('reportApproveBtn');
    const denyBtn = document.getElementById('reportDenyBtn');
    const resetBtn = document.getElementById('reportResetBtn');
    if (approveBtn) approveBtn.addEventListener('click', async () => {
      await setReportApproval(item.id, 'approved');
      loadPage('reports');
    });
    if (denyBtn) denyBtn.addEventListener('click', async () => {
      await setReportApproval(item.id, 'denied');
      loadPage('reports');
    });
    if (resetBtn) resetBtn.addEventListener('click', async () => {
      await setReportApproval(item.id, 'pending');
      loadPage('reports');
    });
  }

  if (canDelete) {
    const deleteBtn = document.getElementById('reportDeleteBtn');
    if (deleteBtn) deleteBtn.addEventListener('click', async () => {
      await deleteReportItem(item.id);
      loadPage('reports');
    });
  }
}

async function loadReportItems() {
  const tableBody = document.querySelector('#reportsTable tbody');
  if (!tableBody) return;

  const typeFilterEl = document.getElementById('reportTypeFilter');
  const statusFilterEl = document.getElementById('reportStatusFilter');
  const officerFilterEl = document.getElementById('reportOfficerFilter');
  const type = String((typeFilterEl && typeFilterEl.value) || 'all').trim();
  const status = String((statusFilterEl && statusFilterEl.value) || 'all').trim();
  const officerFilterText = String((officerFilterEl && officerFilterEl.value) || '').trim().toLowerCase();

  lastLoadedReportItems = [];
  tableBody.innerHTML = '<tr><td colspan="9">Loading reports...</td></tr>';

  try {
    const url = '/api/reports/items?type=' + encodeURIComponent(type) + '&status=' + encodeURIComponent(status);
    const response = await authFetch(url);
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to load report items');

    const items = Array.isArray(data.items) ? data.items : [];
    const filteredItems = officerFilterText
      ? items.filter((item) => String((item && item.officerName) || '').toLowerCase().includes(officerFilterText))
      : items;

    if (!filteredItems.length) {
      tableBody.innerHTML = '<tr><td colspan="9">No reports found for current filters.</td></tr>';
      return;
    }

    lastLoadedReportItems = filteredItems;

    tableBody.innerHTML = filteredItems.map((item) => {
      const canApprove = reportCanBeApproved(item);
      const canDelete = !!currentUser && hasLeadershipAccessClient(currentUser);
      const statusText = escapeHtml(item.approvalStatus || 'pending');
      const actionButtons = canApprove
        ? '<button onclick="setReportApproval(\'' + escapeHtml(item.id) + '\',\'approved\')">Approve</button> ' +
          '<button onclick="setReportApproval(\'' + escapeHtml(item.id) + '\',\'denied\')">Deny</button> ' +
          '<button onclick="setReportApproval(\'' + escapeHtml(item.id) + '\',\'pending\')">Reset</button>'
        : '-';
      const deleteButton = canDelete ? (' <button onclick="deleteReportItem(\'' + escapeHtml(item.id) + '\')">Delete</button>') : '';

      return '<tr>' +
        '<td>' + escapeHtml(item.type) + '</td>' +
        '<td title="' + escapeHtml(item.detail || '') + '"><button type="button" class="report-open-btn" data-report-id="' + escapeHtml(item.id) + '">' + escapeHtml(item.subject || '-') + '</button></td>' +
        '<td>' + escapeHtml(item.officerName || '-') + '</td>' +
        '<td>' + escapeHtml(item.reportDate || '-') + '</td>' +
        '<td>' + escapeHtml(item.sourceTab || '-') + '</td>' +
        '<td>' + statusText + '</td>' +
        '<td>' + escapeHtml(item.approvedBy || '-') + '</td>' +
        '<td>' + escapeHtml(formatDateTime(item.approvedAt || '')) + '</td>' +
        '<td>' + actionButtons + deleteButton + '</td>' +
        '</tr>';
    }).join('');

    const detailButtons = Array.from(tableBody.querySelectorAll('.report-open-btn'));
    detailButtons.forEach((btn) => {
      btn.addEventListener('click', () => {
        const reportId = String(btn.getAttribute('data-report-id') || '').trim();
        showReportDetailsById(reportId);
      });
    });
  } catch (err) {
    tableBody.innerHTML = '<tr><td colspan="9">Unable to load reports: ' + escapeHtml(err.message) + '</td></tr>';
  }
}

async function setReportApproval(reportId, status) {
  try {
    const response = await authFetch('/api/reports/' + encodeURIComponent(reportId) + '/approval', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ status })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to update approval status');

    await loadReportsSummary();
    await loadReportItems();
  } catch (err) {
    alert('Approval update failed: ' + err.message);
  }
}

async function deleteReportItem(reportId) {
  const id = String(reportId || '').trim();
  if (!id) return;
  if (!confirm('Delete this report row from source data?')) return;
  try {
    const response = await authFetch('/api/reports/' + encodeURIComponent(id), { method: 'DELETE' });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Delete failed');
    await loadReportsSummary();
    await loadReportItems();
  } catch (err) {
    alert('Report delete failed: ' + err.message);
  }
}

async function loadDiscussionMessages(options = {}) {
  const opts = options || {};
  const body = document.querySelector('#discussionTable tbody');
  const thread = document.getElementById('discussionMessages');
  const status = document.getElementById('discussionStatus');
  if (!body && !thread) return;

  let wasNearBottom = true;
  if (thread) {
    const distanceToBottom = thread.scrollHeight - thread.scrollTop - thread.clientHeight;
    wasNearBottom = distanceToBottom < 80;
    if (!opts.silent) thread.innerHTML = '<div class="discussion-chat-empty">Loading...</div>';
  } else if (body && !opts.silent) {
    body.innerHTML = '<tr><td colspan="4">Loading...</td></tr>';
  }
  try {
    const response = await authFetch('/api/discussions/messages?limit=250');
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to load discussion board');
    const rows = Array.isArray(data.items) ? data.items : [];
    if (!rows.length) {
      if (thread) {
        thread.innerHTML = '<div class="discussion-chat-empty">No discussion messages yet.</div>';
      } else if (body) {
        body.innerHTML = '<tr><td colspan="4">No discussion messages yet.</td></tr>';
      }
      if (status) status.textContent = 'No discussion messages yet.';
      return;
    }
    if (thread) {
      const me = normalizeEmailClient(currentUser && currentUser.email);
      thread.innerHTML = rows.map((m) => {
        const authorEmail = normalizeEmailClient(m.authorEmail);
        const isMine = !!me && me === authorEmail;
        const canDelete = hasLeadershipAccessClient(currentUser) || isMine;
        const deleteBtn = canDelete
          ? ('<button class="discussion-delete-btn" onclick="deleteDiscussionMessage(\'' + escapeHtml(m.id) + '\')">Delete</button>')
          : '';
        return '<div class="discussion-chat-message ' + (isMine ? 'mine' : 'other') + '">' +
          '<div class="discussion-chat-meta">' +
            '<span>' + escapeHtml(m.author || 'Officer') + '</span>' +
            '<span>' + escapeHtml(formatDateTime(m.createdAt || '')) + '</span>' +
          '</div>' +
          '<div class="discussion-chat-text">' + escapeHtml(m.text || '-') + '</div>' +
          (deleteBtn ? ('<div class="discussion-chat-actions">' + deleteBtn + '</div>') : '') +
        '</div>';
      }).join('');
      if (wasNearBottom || !opts.preserveScroll) {
        thread.scrollTop = thread.scrollHeight;
      }
    } else if (body) {
      body.innerHTML = rows.map((m) => {
        const canDelete = hasLeadershipAccessClient(currentUser) || normalizeEmailClient(currentUser && currentUser.email) === normalizeEmailClient(m.authorEmail);
        const action = canDelete ? ('<button onclick="deleteDiscussionMessage(\'' + escapeHtml(m.id) + '\')">Delete</button>') : '-';
        return '<tr>' +
          '<td>' + escapeHtml(formatDateTime(m.createdAt || '')) + '</td>' +
          '<td>' + escapeHtml(m.author || '-') + '</td>' +
          '<td>' + escapeHtml(m.text || '-') + '</td>' +
          '<td>' + action + '</td>' +
        '</tr>';
      }).join('');
    }
    if (status) status.textContent = 'Discussion messages: ' + rows.length;
  } catch (err) {
    if (thread) {
      thread.innerHTML = '<div class="discussion-chat-empty">' + escapeHtml(err.message) + '</div>';
    } else if (body) {
      body.innerHTML = '<tr><td colspan="4">' + escapeHtml(err.message) + '</td></tr>';
    }
    if (status) status.textContent = 'Discussion load failed: ' + err.message;
  }
}

async function postDiscussionMessage() {
  const input = document.getElementById('discussionText');
  const status = document.getElementById('discussionStatus');
  const text = String((input && input.value) || '').trim();
  if (!text) {
    if (status) status.textContent = 'Enter a message first.';
    return;
  }
  try {
    const response = await authFetch('/api/discussions/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Post failed');
    if (input) input.value = '';
    await loadDiscussionMessages({ preserveScroll: false });
  } catch (err) {
    if (status) status.textContent = 'Post failed: ' + err.message;
  }
}

async function deleteDiscussionMessage(id) {
  try {
    const response = await authFetch('/api/discussions/messages/' + encodeURIComponent(String(id || '')), { method: 'DELETE' });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Delete failed');
    await loadDiscussionMessages({ preserveScroll: true });
  } catch (err) {
    alert('Discussion delete failed: ' + err.message);
  }
}

async function loadAdminChatMessages() {
  const body = document.querySelector('#adminChatTable tbody');
  const status = document.getElementById('adminChatStatus');
  if (!body) return;
  body.innerHTML = '<tr><td colspan="4">Loading...</td></tr>';
  try {
    const response = await authFetch('/api/admin-chat/messages?limit=250');
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to load admin chat');
    const rows = Array.isArray(data.items) ? data.items : [];
    if (!rows.length) {
      body.innerHTML = '<tr><td colspan="4">No admin chat messages yet.</td></tr>';
      if (status) status.textContent = 'No admin chat messages yet.';
      return;
    }
    body.innerHTML = rows.map((m) => {
      return '<tr>' +
        '<td>' + escapeHtml(formatDateTime(m.createdAt || '')) + '</td>' +
        '<td>' + escapeHtml(m.author || '-') + '</td>' +
        '<td>' + escapeHtml(m.text || '-') + '</td>' +
        '<td><button onclick="deleteAdminChatMessage(\'' + escapeHtml(m.id) + '\')">Delete</button></td>' +
      '</tr>';
    }).join('');
    if (status) status.textContent = 'Admin chat messages: ' + rows.length;
  } catch (err) {
    body.innerHTML = '<tr><td colspan="4">' + escapeHtml(err.message) + '</td></tr>';
    if (status) status.textContent = 'Admin chat load failed: ' + err.message;
  }
}

async function postAdminChatMessage() {
  const input = document.getElementById('adminChatText');
  const status = document.getElementById('adminChatStatus');
  const text = String((input && input.value) || '').trim();
  if (!text) {
    if (status) status.textContent = 'Enter a message first.';
    return;
  }
  try {
    const response = await authFetch('/api/admin-chat/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Send failed');
    if (input) input.value = '';
    await loadAdminChatMessages();
  } catch (err) {
    if (status) status.textContent = 'Send failed: ' + err.message;
  }
}

async function deleteAdminChatMessage(id) {
  try {
    const response = await authFetch('/api/admin-chat/messages/' + encodeURIComponent(String(id || '')), { method: 'DELETE' });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Delete failed');
    await loadAdminChatMessages();
  } catch (err) {
    alert('Admin chat delete failed: ' + err.message);
  }
}

async function loadCalendarEvents() {
  const body = document.querySelector('#calendarTable tbody');
  const status = document.getElementById('calendarStatus');
  if (!body) return;
  body.innerHTML = '<tr><td colspan="7">Loading...</td></tr>';
  try {
    const response = await authFetch('/api/calendar/events');
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to load calendar');
    const rows = Array.isArray(data.items) ? data.items : [];
    const viewTz = getCalendarViewTimezone();
    if (!rows.length) {
      body.innerHTML = '<tr><td colspan="7">No events scheduled.</td></tr>';
      if (status) status.textContent = 'No events scheduled.';
      return;
    }
    body.innerHTML = rows.map((ev) => {
      const canDelete = hasLeadershipAccessClient(currentUser) || normalizeEmailClient(currentUser && currentUser.email) === normalizeEmailClient(ev.createdByEmail);
      const action = canDelete ? ('<button onclick="deleteCalendarEvent(\'' + escapeHtml(ev.id) + '\')">Delete</button>') : '-';
      const when = formatCalendarEventForTimezone(ev, viewTz);
      return '<tr>' +
        '<td>' + escapeHtml(when.date) + '</td>' +
        '<td>' + escapeHtml(when.time) + '</td>' +
        '<td>' + escapeHtml(ev.title || '-') + '</td>' +
        '<td>' + escapeHtml(ev.note || '-') + '</td>' +
        '<td>' + escapeHtml(ev.createdBy || '-') + '</td>' +
        '<td>' + escapeHtml(ev.timezone || DEFAULT_CALENDAR_TIMEZONE) + '</td>' +
        '<td>' + action + '</td>' +
      '</tr>';
    }).join('');
    if (status) status.textContent = 'Scheduled events: ' + rows.length;
  } catch (err) {
    body.innerHTML = '<tr><td colspan="7">' + escapeHtml(err.message) + '</td></tr>';
    if (status) status.textContent = 'Calendar load failed: ' + err.message;
  }
}

async function addCalendarEvent() {
  const date = String((document.getElementById('calendarDate') || {}).value || '').trim();
  const time = String((document.getElementById('calendarTime') || {}).value || '').trim();
  const timezone = String((document.getElementById('calendarEventTimezone') || {}).value || '').trim() || DEFAULT_CALENDAR_TIMEZONE;
  const title = String((document.getElementById('calendarTitle') || {}).value || '').trim();
  const note = String((document.getElementById('calendarNote') || {}).value || '').trim();
  const status = document.getElementById('calendarStatus');
  try {
    const response = await authFetch('/api/calendar/events', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ date, time, timezone, title, note })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Add event failed');
    const titleEl = document.getElementById('calendarTitle');
    const noteEl = document.getElementById('calendarNote');
    if (titleEl) titleEl.value = '';
    if (noteEl) noteEl.value = '';
    await loadCalendarEvents();
    await checkCalendarReminders();
  } catch (err) {
    if (status) status.textContent = 'Add event failed: ' + err.message;
  }
}

function getCalendarViewTimezone() {
  let tz = DEFAULT_CALENDAR_TIMEZONE;
  try {
    tz = String(localStorage.getItem(CALENDAR_VIEW_TIMEZONE_KEY) || DEFAULT_CALENDAR_TIMEZONE);
  } catch (e) {
    tz = DEFAULT_CALENDAR_TIMEZONE;
  }
  if (!isValidTimeZoneClient(tz)) return DEFAULT_CALENDAR_TIMEZONE;
  return tz;
}

function initCalendarTimezoneSelectors() {
  const eventTzEl = document.getElementById('calendarEventTimezone');
  const viewTzEl = document.getElementById('calendarViewTimezone');
  const optionsHtml = CALENDAR_TIMEZONE_OPTIONS.map((opt) => (
    '<option value="' + escapeHtml(opt.value) + '">' + escapeHtml(opt.label) + '</option>'
  )).join('');

  if (eventTzEl) {
    eventTzEl.innerHTML = optionsHtml;
    eventTzEl.value = DEFAULT_CALENDAR_TIMEZONE;
  }
  if (viewTzEl) {
    viewTzEl.innerHTML = optionsHtml;
    viewTzEl.value = getCalendarViewTimezone();
  }
}

function isValidTimeZoneClient(timeZone) {
  try {
    Intl.DateTimeFormat('en-US', { timeZone: String(timeZone || '') }).format(new Date());
    return true;
  } catch (e) {
    return false;
  }
}

function formatCalendarEventForTimezone(ev, viewTimeZone) {
  const dateRaw = String(ev && ev.date || '').trim();
  const timeRaw = String(ev && ev.time || '').trim();
  const tz = isValidTimeZoneClient(viewTimeZone) ? viewTimeZone : DEFAULT_CALENDAR_TIMEZONE;
  const utcAt = String(ev && ev.utcAt || '').trim();
  if (!utcAt) return { date: dateRaw || '-', time: timeRaw || '-' };
  const d = new Date(utcAt);
  if (Number.isNaN(d.getTime())) return { date: dateRaw || '-', time: timeRaw || '-' };

  const dateText = new Intl.DateTimeFormat('en-US', {
    timeZone: tz,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).format(d);
  const timeText = new Intl.DateTimeFormat('en-US', {
    timeZone: tz,
    hour: '2-digit',
    minute: '2-digit',
    hour12: true
  }).format(d);
  return { date: dateText, time: timeText };
}

async function deleteCalendarEvent(id) {
  try {
    const response = await authFetch('/api/calendar/events/' + encodeURIComponent(String(id || '')), { method: 'DELETE' });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Delete failed');
    await loadCalendarEvents();
  } catch (err) {
    alert('Calendar delete failed: ' + err.message);
  }
}

async function checkCalendarReminders() {
  if (!isLoggedIn()) return;
  const key = 'fwpd_calendar_reminded_ids';
  let reminded = [];
  try {
    reminded = JSON.parse(sessionStorage.getItem(key) || '[]');
    if (!Array.isArray(reminded)) reminded = [];
  } catch (e) {
    reminded = [];
  }

  try {
    const response = await authFetch('/api/calendar/reminders?windowHours=24');
    const data = await response.json();
    if (!response.ok) return;
    const items = Array.isArray(data.items) ? data.items : [];
    const unseen = items.filter((ev) => !reminded.includes(String(ev.id || '')));
    if (!unseen.length) return;
    const lines = unseen.slice(0, 3).map((ev) => (ev.date || '?') + ' ' + (ev.time || '') + ' - ' + (ev.title || 'Event'));
    alert('Upcoming event reminder:\n' + lines.join('\n'));
    reminded = reminded.concat(unseen.map((ev) => String(ev.id || '')));
    sessionStorage.setItem(key, JSON.stringify(Array.from(new Set(reminded)).slice(-300)));
  } catch (e) {
    // Silent by design.
  }
}

async function linkDisciplinarySource() {
  const tabNameEl = document.getElementById('disciplineTabName');
  const urlEl = document.getElementById('disciplineSourceUrl');
  const tabName = String((tabNameEl && tabNameEl.value) || 'disciplinary_forms').trim();
  const url = String((urlEl && urlEl.value) || '').trim();

  if (!url) {
    alert('Paste the disciplinary source Google Sheet/CSV link first.');
    return;
  }

  try {
    const response = await authFetch('/api/reports/link-disciplinary-source', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tabName, url })
    });
    if (response.status === 401) throw new Error('Session expired. Please log in again, then retry linking.');
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to link disciplinary source');

    try {
      localStorage.setItem(DISCIPLINE_SOURCE_URL_KEY, url);
    } catch (e) {
      // Ignore localStorage write errors.
    }

    alert('Disciplinary source linked as tab "' + tabName + '".');
    await loadReportsSummary();
    await loadReportItems();
    loadSyncStatus();
  } catch (err) {
    alert('Link failed: ' + err.message);
  }
}

async function linkEvaluationSource() {
  const tabNameEl = document.getElementById('evaluationTabName');
  const urlEl = document.getElementById('evaluationSourceUrl');
  const tabName = String((tabNameEl && tabNameEl.value) || 'cadet_evaluations').trim();
  const url = String((urlEl && urlEl.value) || '').trim();

  if (!url) {
    alert('Paste the cadet evaluations source Google Sheet/CSV link first.');
    return;
  }

  try {
    const response = await authFetch('/api/reports/link-evaluation-source', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tabName, url })
    });
    if (response.status === 401) throw new Error('Session expired. Please log in again, then retry linking.');
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to link cadet evaluations source');

    try {
      localStorage.setItem(EVALUATION_SOURCE_URL_KEY, url);
    } catch (e) {
      // Ignore localStorage write errors.
    }

    alert('Cadet evaluations source linked as tab "' + tabName + '".');
    await loadReportsSummary();
    await loadReportItems();
    loadSyncStatus();
  } catch (err) {
    alert('Link failed: ' + err.message);
  }
}

async function changePassword() {
  const oldPassword = String((document.getElementById('oldPassword') || {}).value || '');
  const newPassword = String((document.getElementById('newPassword') || {}).value || '');
  const confirmPassword = String((document.getElementById('confirmPassword') || {}).value || '');
  const status = document.getElementById('accountStatus');

  try {
    const response = await authFetch('/api/auth/change-password', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ oldPassword, newPassword, confirmPassword })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Password update failed');

    if (status) status.textContent = data.message || 'Password updated successfully.';
    const oldEl = document.getElementById('oldPassword');
    const newEl = document.getElementById('newPassword');
    const confirmEl = document.getElementById('confirmPassword');
    if (oldEl) oldEl.value = '';
    if (newEl) newEl.value = '';
    if (confirmEl) confirmEl.value = '';
  } catch (err) {
    if (status) status.textContent = 'Password update failed: ' + err.message;
  }
}

async function adminResetPassword() {
  const email = String((document.getElementById('resetEmail') || {}).value || '').trim();
  const newPassword = String((document.getElementById('resetNewPassword') || {}).value || '');
  const status = document.getElementById('accountStatus');

  try {
    const response = await authFetch('/api/auth/admin-reset-password', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, newPassword })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Admin reset failed');

    if (status) status.textContent = data.message || ('Password reset for ' + email + '.');
    const emailEl = document.getElementById('resetEmail');
    const pwdEl = document.getElementById('resetNewPassword');
    if (emailEl) emailEl.value = '';
    if (pwdEl) pwdEl.value = '';
  } catch (err) {
    if (status) status.textContent = 'Admin reset failed: ' + err.message;
  }
}

function messagePreview(text) {
  const raw = String(text || '').trim();
  if (!raw) return '-';
  return raw.length > 120 ? (raw.slice(0, 117) + '...') : raw;
}

async function loadMessageRecipients() {
  const select = document.getElementById('msgRecipient');
  if (!select) return;

  try {
    const response = await authFetch('/api/messages/users');
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to load recipients');

    const users = Array.isArray(data.users) ? data.users : [];
    if (!users.length) {
      select.innerHTML = '<option value="">No recipients available</option>';
      return;
    }

    select.innerHTML = users.map((u) => {
      const email = escapeHtml(u.email || '');
      const label = escapeHtml(u.displayName || u.email || '');
      return '<option value="' + email + '">' + label + '</option>';
    }).join('');
  } catch (err) {
    select.innerHTML = '<option value="">Failed to load recipients</option>';
  }
}

function renderMessagesTable(items, mode) {
  const tableBody = document.querySelector('#messagesTable tbody');
  const status = document.getElementById('messagesStatus');
  if (!tableBody) return;

  const rows = Array.isArray(items) ? items : [];
  if (!rows.length) {
    tableBody.innerHTML = '<tr><td colspan="6">No messages found.</td></tr>';
    if (status) status.textContent = mode === 'sent' ? 'No sent messages yet.' : 'Inbox is empty.';
    return;
  }

  tableBody.innerHTML = rows.map((m) => {
    const who = mode === 'sent'
      ? ('To: ' + escapeHtml(m.toName || '-'))
      : ('From: ' + escapeHtml(m.fromName || '-'));
    const readStatus = m.readAt ? 'Read' : 'Unread';
    const action = mode === 'sent'
      ? '-'
      : (m.readAt
        ? '<button onclick="markMessageRead(\'' + escapeHtml(m.id) + '\',false)">Mark Unread</button>'
        : '<button onclick="markMessageRead(\'' + escapeHtml(m.id) + '\',true)">Mark Read</button>');

    return '<tr>' +
      '<td>' + who + '</td>' +
      '<td>' + escapeHtml(m.subject || '-') + '</td>' +
      '<td title="' + escapeHtml(m.body || '') + '">' + escapeHtml(messagePreview(m.body || '')) + '</td>' +
      '<td>' + escapeHtml(formatDateTime(m.createdAt || '')) + '</td>' +
      '<td>' + escapeHtml(readStatus) + '</td>' +
      '<td>' + action + '</td>' +
      '</tr>';
  }).join('');

  if (status) {
    status.textContent = (mode === 'sent' ? 'Sent messages' : 'Inbox messages') + ': ' + rows.length;
  }
}

async function loadInboxMessages() {
  const tableBody = document.querySelector('#messagesTable tbody');
  if (tableBody) tableBody.innerHTML = '<tr><td colspan="6">Loading inbox...</td></tr>';

  try {
    const response = await authFetch('/api/messages/inbox?limit=1000');
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to load inbox');
    renderMessagesTable(data.items || [], 'inbox');
    await loadUnreadMessageCount();
  } catch (err) {
    if (tableBody) tableBody.innerHTML = '<tr><td colspan="6">' + escapeHtml(err.message) + '</td></tr>';
    const status = document.getElementById('messagesStatus');
    if (status) status.textContent = 'Inbox failed: ' + err.message;
  }
}

async function loadSentMessages() {
  const tableBody = document.querySelector('#messagesTable tbody');
  if (tableBody) tableBody.innerHTML = '<tr><td colspan="6">Loading sent messages...</td></tr>';

  try {
    const response = await authFetch('/api/messages/sent?limit=1000');
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to load sent messages');
    renderMessagesTable(data.items || [], 'sent');
  } catch (err) {
    if (tableBody) tableBody.innerHTML = '<tr><td colspan="6">' + escapeHtml(err.message) + '</td></tr>';
    const status = document.getElementById('messagesStatus');
    if (status) status.textContent = 'Sent failed: ' + err.message;
  }
}

async function sendInternalMessage() {
  const recipientSelect = document.getElementById('msgRecipient');
  const recipients = Array.from((recipientSelect && recipientSelect.selectedOptions) || [])
    .map((opt) => String(opt.value || '').trim())
    .filter(Boolean);
  const subject = String((document.getElementById('msgSubject') || {}).value || '').trim();
  const body = String((document.getElementById('msgBody') || {}).value || '').trim();
  const status = document.getElementById('messagesStatus');

  if (!recipients.length) {
    if (status) status.textContent = 'Select at least one recipient first.';
    return;
  }

  try {
    const response = await authFetch('/api/messages/send', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ toEmails: recipients, subject, body })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Message send failed');

    const sentCount = Number(data.count || 0);
    const invalidCount = Array.isArray(data.invalidRecipients) ? data.invalidRecipients.length : 0;
    if (status) {
      status.textContent = 'Sent ' + sentCount + ' message(s).' +
        (invalidCount ? (' Skipped ' + invalidCount + ' invalid recipient(s).') : '');
    }
    const subjectEl = document.getElementById('msgSubject');
    const bodyEl = document.getElementById('msgBody');
    const recipientEl = document.getElementById('msgRecipient');
    if (subjectEl) subjectEl.value = '';
    if (bodyEl) bodyEl.value = '';
    if (recipientEl) Array.from(recipientEl.options || []).forEach((opt) => { opt.selected = false; });
    loadSentMessages();
  } catch (err) {
    if (status) status.textContent = 'Send failed: ' + err.message;
  }
}

async function loadAdminUsers() {
  const tbody = document.querySelector('#adminUsersTable tbody');
  const status = document.getElementById('adminStatus');
  if (!tbody) return;
  tbody.innerHTML = '<tr><td colspan="5">Loading...</td></tr>';

  try {
    const response = await authFetch('/api/admin/users');
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to load admin users');

    const users = Array.isArray(data.users) ? data.users : [];
    if (!users.length) {
      tbody.innerHTML = '<tr><td colspan="5">No command users found.</td></tr>';
      if (status) status.textContent = 'No command users available.';
      return;
    }

    tbody.innerHTML = users.map((u) => {
      const email = escapeHtml(u.email || '');
      const name = escapeHtml(u.displayName || u.email || '-');
      const role = escapeHtml(u.role || '-');
      const hasAccount = !!u.hasAccount;
      const isAdmin = !!u.isAdmin;
      const actionBtn = '<button class="admin-toggle-btn" data-email="' + email + '" data-next="' + (isAdmin ? '0' : '1') + '">' + (isAdmin ? 'Revoke Admin' : 'Grant Admin') + '</button>';

      return '<tr>' +
        '<td>' + name + '</td>' +
        '<td>' + email + '</td>' +
        '<td>' + (hasAccount ? 'Yes' : 'No') + '</td>' +
        '<td>' + (role ? role.toUpperCase() : '') + '</td>' +
        '<td>' + actionBtn + '</td>' +
      '</tr>';
    }).join('');

    Array.from(document.querySelectorAll('.admin-toggle-btn')).forEach((btn) => {
      btn.addEventListener('click', async () => {
        const email = String(btn.getAttribute('data-email') || '').trim();
        const nextAdmin = String(btn.getAttribute('data-next') || '') === '1';
        await setAdminAccess(email, nextAdmin);
      });
    });

    const adminCount = users.filter((u) => !!u.isAdmin).length;
    if (status) status.textContent = 'Loaded ' + users.length + ' users. Admin-enabled: ' + adminCount + '.';
  } catch (err) {
    tbody.innerHTML = '<tr><td colspan="5">' + escapeHtml(err.message) + '</td></tr>';
    if (status) status.textContent = 'Admin user load failed: ' + err.message;
  }
}

async function setAdminAccess(email, isAdmin) {
  const status = document.getElementById('adminStatus');
  try {
    const response = await authFetch('/api/admin/set-admin', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, isAdmin: !!isAdmin })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Failed to update admin access');
    if (status) status.textContent = (isAdmin ? 'Granted admin access for ' : 'Revoked admin access for ') + email + '.';
    await loadAdminUsers();
  } catch (err) {
    if (status) status.textContent = 'Admin access update failed: ' + err.message;
  }
}

async function markMessageRead(messageId, read) {
  try {
    const response = await authFetch('/api/messages/' + encodeURIComponent(messageId) + '/read', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ read: !!read })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Message status update failed');
    await loadInboxMessages();
  } catch (err) {
    const status = document.getElementById('messagesStatus');
    if (status) status.textContent = 'Read status update failed: ' + err.message;
  }
}

async function logoutAccount() {
  try {
    await authFetch('/api/auth/logout', { method: 'POST' });
  } catch (e) {
    // Ignore logout errors.
  }
  setAuthToken('');
  stopMessagePolling();
  unreadMessageCount = 0;
  currentUser = null;
  showAuthBanner();
  applyRuntimeLayoutFixes();
  renderLoginScreen('Logged out. Please login to continue.');
}

async function syncGoogleSheets() {
  const isValidGoogleSheetLink = (input) => {
    const value = String(input || '').trim();
    if (!value) return false;
    try {
      const u = new URL(value);
      const host = String(u.hostname || '').toLowerCase();
      const path = String(u.pathname || '').toLowerCase();
      if (!host.includes('docs.google.com')) return false;
      return path.includes('/spreadsheets/');
    } catch (e) {
      return false;
    }
  };

  const rosterURL = String(DEFAULT_ROSTER_SOURCE_URL || '').trim() ||
    prompt('Paste Roster Google Sheets link (tab link or published CSV):', String(DEFAULT_ROSTER_SOURCE_URL || '').trim());
  if (!rosterURL) return;

  if (!isValidGoogleSheetLink(rosterURL)) {
    alert('Invalid roster link. Please paste a Google Sheets link from docs.google.com/spreadsheets/...');
    return;
  }

  const otherTabsInput = prompt(
    'Optional: add other tabs as comma-separated Name|CSV_URL entries.\nExample:\ndivisions|https://...csv,discipline|https://...csv'
  ) || '';

  const tabs = [{ name: 'roster', url: rosterURL.trim() }];

  if (otherTabsInput.trim()) {
    const invalidLinks = [];
    otherTabsInput.split(',').forEach(pair => {
      const parts = pair.split('|');
      if (parts.length >= 2) {
        const name = parts[0].trim();
        const url = parts.slice(1).join('|').trim();
        if (!name || !url) return;
        if (!isValidGoogleSheetLink(url)) {
          invalidLinks.push(name || 'unnamed');
          return;
        }
        tabs.push({ name, url });
      }
    });

    if (invalidLinks.length) {
      alert('Invalid Google Sheets link for tab(s): ' + invalidLinks.join(', ') + '.\nUse docs.google.com/spreadsheets/... links.');
      return;
    }
  }

  try {
    setLocalSyncTabs(tabs);

    await fetch('/api/sheets/config', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tabs, autoSyncOnLoad: true })
    });

    const response = await fetch('/api/sheets/import', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tabs })
    });

    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'Import failed');

    const rosterResult = (result.result || []).find(x => x.name === 'roster');
    const tabSummary = (result.result || []).map(x => {
      if (!x.ok) return `${x.name}: FAILED (${x.error || 'error'})`;
      if (typeof x.rows === 'number') return `${x.name}: OK (${x.rows} rows)`;
      return `${x.name}: OK`;
    }).join(', ');

    let message = 'Import completed. ' + tabSummary;
    if (rosterResult && rosterResult.ok) {
      const parsed = Number(rosterResult.parsedRows || 0);
      const headers = Array.isArray(rosterResult.headers) ? rosterResult.headers.join(', ') : '';
      message += `\nRoster parser: ${parsed} source rows.`;
      if (headers) message += `\nDetected headers: ${headers}`;
    }
    if (rosterResult && rosterResult.ok && Number(rosterResult.rows || 0) === 0) {
      message += '\n\nRoster imported 0 rows. Confirm the roster tab has officer data and a Name/RP_Name or Callsign column.';
    }

    alert(message);
    loadRoster();
    loadSyncStatus();
  } catch (err) {
    alert('Google sync failed: ' + err.message + '\n\nUse a published CSV link or standard sheet link with correct tab gid.');
  }
}

function formatSyncSummary(result) {
  const okRows = (result || []).filter(x => x.ok).map(x => {
    const rows = typeof x.rows === 'number' ? ` (${x.rows} rows)` : '';
    return `${x.name}: OK${rows}`;
  });
  const failRows = (result || []).filter(x => !x.ok).map(x => `${x.name}: FAILED (${x.error || 'error'})`);
  return okRows.concat(failRows).join('\n');
}

async function loadSyncStatus() {
  const box = document.getElementById('syncStatusBox');
  if (!box) return;

  try {
    const response = await fetch('/api/sheets/config');
    const config = await response.json();
    if (!response.ok) throw new Error(config.error || 'Failed to load sync status');

    const lines = [];
    lines.push('Auto-sync on load: ' + (config.autoSyncOnLoad ? 'Enabled' : 'Disabled'));
    lines.push('Saved tabs: ' + ((config.tabs || []).map(t => t.name).join(', ') || 'none'));
    if (config.lastSync && config.lastSync.at) {
      lines.push('Last sync: ' + new Date(config.lastSync.at).toLocaleString());
      if (config.lastSync.summary) {
        lines.push('Summary: ' + config.lastSync.summary.ok + ' OK / ' + config.lastSync.summary.failed + ' failed');
      }
      if (Array.isArray(config.lastSync.result) && config.lastSync.result.length) {
        lines.push('');
        lines.push(formatSyncSummary(config.lastSync.result));
      }
    } else {
      lines.push('Last sync: none yet');
    }

    box.textContent = lines.join('\n');
  } catch (err) {
    box.textContent = 'Sync status unavailable: ' + err.message;
  }
}

async function autoSyncOnLoad() {
  if (sessionStorage.getItem(AUTO_SYNC_SESSION_KEY) === '1') return;

  try {
    const configResponse = await fetch('/api/sheets/config');
    const config = await configResponse.json();
    if (!configResponse.ok) return;

    let tabsToSync = Array.isArray(config.tabs) ? config.tabs : [];
    let autoEnabled = !!config.autoSyncOnLoad;
    const defaultRosterUrl = String(DEFAULT_ROSTER_SOURCE_URL || '').trim();
    const defaultDisciplineUrl = String(DEFAULT_DISCIPLINE_SOURCE_URL || '').trim();
    const defaultEvalUrl = String(DEFAULT_EVALUATION_SOURCE_URL || '').trim();
    const defaultCommandUsersUrl = String(localStorage.getItem(COMMAND_USERS_SOURCE_URL_KEY) || DEFAULT_COMMAND_USERS_SOURCE_URL || '').trim();

    const upsertTab = (name, url) => {
      const safeName = String(name || '').trim().toLowerCase();
      const safeUrl = String(url || '').trim();
      if (!safeName || !safeUrl) return;
      const idx = tabsToSync.findIndex((t) => String((t && t.name) || '').trim().toLowerCase() === safeName);
      if (idx >= 0) tabsToSync[idx] = { name: safeName, url: safeUrl };
      else tabsToSync.push({ name: safeName, url: safeUrl });
    };

    upsertTab('roster', defaultRosterUrl);
    upsertTab('disciplinary_forms', defaultDisciplineUrl);
    upsertTab('cadet_evaluations', defaultEvalUrl);
    upsertTab('command_users', defaultCommandUsersUrl);

    if (tabsToSync.length) {
      autoEnabled = true;
      setLocalSyncTabs(tabsToSync);
      await fetch('/api/sheets/config', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tabs: tabsToSync, autoSyncOnLoad: true })
      });
    }

    // Always include the hard-set roster tab so users do not need to enter URL.
    if (defaultRosterUrl) {
      const hasRoster = tabsToSync.some((t) => String((t && t.name) || '').trim().toLowerCase() === 'roster');
      if (!hasRoster) {
        tabsToSync = tabsToSync.concat([{ name: 'roster', url: defaultRosterUrl }]);
        autoEnabled = true;
        setLocalSyncTabs(tabsToSync);
        await fetch('/api/sheets/config', {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tabs: tabsToSync, autoSyncOnLoad: true })
        });
      }
    }

    // Render free deployments can reset local files; restore from browser cache when available.
    if (!tabsToSync.length) {
      if (defaultRosterUrl) {
        tabsToSync = [{ name: 'roster', url: defaultRosterUrl }];
        autoEnabled = true;
        setLocalSyncTabs(tabsToSync);
        await fetch('/api/sheets/config', {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tabs: tabsToSync, autoSyncOnLoad: true })
        });
      }
    }

    if (!tabsToSync.length) {
      const cachedTabs = getLocalSyncTabs();
      if (cachedTabs.length) {
        tabsToSync = cachedTabs;
        autoEnabled = true;
        await fetch('/api/sheets/config', {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tabs: tabsToSync, autoSyncOnLoad: true })
        });
      }
    }

    if (!autoEnabled || !tabsToSync.length) {
      await loadSyncStatus();
      return;
    }

    sessionStorage.setItem(AUTO_SYNC_SESSION_KEY, '1');
    const syncResponse = await fetch('/api/sheets/sync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tabs: tabsToSync })
    });
    const syncResult = await syncResponse.json();
    if (!syncResponse.ok) throw new Error(syncResult.error || 'Auto-sync failed');
    await loadSyncStatus();
  } catch (err) {
    const box = document.getElementById('syncStatusBox');
    if (box) box.textContent = 'Auto-sync failed: ' + err.message;
  }
}

async function loadSheetTabs() {
  const listEl = document.getElementById('sheetTabsList');
  const dataEl = document.getElementById('sheetTabData');

  try {
    const response = await fetch('/api/sheets/tabs');
    const tabs = await response.json();
    if (!Array.isArray(tabs) || !tabs.length) {
      listEl.textContent = 'No imported tabs yet. Use Officer Roster > Sync Google Sheets.';
      return;
    }

    listEl.innerHTML = '';
    tabs.forEach(name => {
      const btn = document.createElement('button');
      btn.textContent = 'View ' + name;
      btn.style.marginRight = '8px';
      btn.style.marginBottom = '8px';
      btn.addEventListener('click', () => loadSheetTabData(name));
      listEl.appendChild(btn);
    });

    dataEl.textContent = 'Select a tab to preview.';
  } catch (err) {
    listEl.textContent = 'Failed to load tabs: ' + err.message;
  }
}

async function loadSheetTabData(name) {
  const dataEl = document.getElementById('sheetTabData');
  try {
    const response = await fetch('/api/sheets/tab/' + encodeURIComponent(name));
    const rows = await response.json();
    if (!response.ok) throw new Error(rows.error || 'Failed to load tab');

    if (!Array.isArray(rows) || !rows.length) {
      dataEl.textContent = 'No rows found in ' + name;
      return;
    }

    const cols = Object.keys(rows[0]);
    let html = '<h3>' + name + '</h3><table><thead><tr>';
    cols.forEach(c => { html += '<th>' + c + '</th>'; });
    html += '</tr></thead><tbody>';

    rows.slice(0, 100).forEach(r => {
      html += '<tr>';
      cols.forEach(c => { html += '<td>' + (r[c] || '') + '</td>'; });
      html += '</tr>';
    });

    html += '</tbody></table><p>Showing up to first 100 rows.</p>';
    dataEl.innerHTML = html;
  } catch (err) {
    dataEl.textContent = 'Failed to load tab: ' + err.message;
  }
}

if (!getAuthToken()) {
  currentUser = null;
  showAuthBanner();
  renderLoginScreen();
}

refreshAuthSession();