const CoffeeIntelligence = (() => {
  const SESSION_KEY = 'cirs.session.v1';
  const THEME_KEY = 'cirs.theme';
  const POLL_INTERVAL = 60000;

  const state = {
    token: '',
    user: null,
    links: {
      sheet: 'https://docs.google.com/spreadsheets/d/1cZ7iit2zpPsE_gDcJyi2h2UBvVK64TOlMuuBR8xFflg/edit?usp=sharing',
      webApp: 'https://script.google.com/macros/s/AKfycbwwoLfaAVBdrH9l7myWTZ3rvWlvO0NZRi1cwXISK4_2RO1DV5CxpjfBlo2qRF8kMsz_/exec'
    },
    data: {
      uploads: [],
      insights: [],
      activities: [],
      users: [],
      countries: [],
      summary: {}
    },
    page: 'dashboard',
    poller: null,
    busy: false
  };

  const pageMeta = {
    dashboard: ['Workspace overview', 'Dashboard'],
    uploads: ['Research data center', 'Uploads'],
    insights: ['Collaborative notes', 'Insights'],
    countries: ['Market monitoring', 'Countries'],
    activity: ['Research operations', 'Activity'],
    users: ['Administration', 'Users'],
    settings: ['Workspace', 'Settings']
  };

  const contentTypes = [
    ['all', 'All types'],
    ['news', 'News'],
    ['report', 'Report'],
    ['pdf', 'PDF'],
    ['spreadsheet', 'Spreadsheet'],
    ['scientific_article', 'Scientific article'],
    ['reference', 'Reference'],
    ['observation', 'Observation'],
    ['external_link', 'External link'],
    ['google_drive', 'Google Drive link']
  ];

  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];

  const Api = {
    call(name, ...args) {
      return new Promise((resolve, reject) => {
        const requestId = `cirs_${Date.now()}_${Math.random().toString(36).slice(2)}`;
        const frameName = `cirs_frame_${requestId}`;
        const iframe = document.createElement('iframe');
        const form = document.createElement('form');
        const payload = JSON.stringify(args);

        const timeout = window.setTimeout(() => {
          cleanup();
          reject(new Error('The Apps Script API did not respond in time.'));
        }, 30000);

        function cleanup() {
          window.clearTimeout(timeout);
          window.removeEventListener('message', handleMessage);
          form.remove();
          iframe.remove();
        }

        function handleMessage(event) {
          const data = event.data || {};
          if (!data.__cirsResponse || data.requestId !== requestId) return;
          cleanup();
          resolve(data.response);
        }

        window.addEventListener('message', handleMessage);

        iframe.name = frameName;
        iframe.style.display = 'none';
        iframe.setAttribute('aria-hidden', 'true');

        form.method = 'POST';
        form.action = state.links.webApp;
        form.target = frameName;
        form.style.display = 'none';
        appendHidden(form, 'action', name);
        appendHidden(form, 'payload', payload);
        appendHidden(form, 'requestId', requestId);
        appendHidden(form, 'origin', window.location.origin || '*');

        iframe.onerror = () => {
          cleanup();
          reject(new Error('Unable to reach the Apps Script API. Check the Web App deployment permissions.'));
        };

        document.body.appendChild(iframe);
        document.body.appendChild(form);
        form.submit();
      });
    },

    async expect(name, ...args) {
      const result = await Api.call(name, ...args);
      if (!result || result.success === false) {
        throw new Error(result && result.error ? result.error : 'Unexpected backend response.');
      }
      return result;
    }
  };

  const Auth = {
    loadSession() {
      try {
        const saved = JSON.parse(localStorage.getItem(SESSION_KEY) || '{}');
        if (saved.token && saved.user) {
          state.token = saved.token;
          state.user = saved.user;
          return true;
        }
      } catch (error) {
        localStorage.removeItem(SESSION_KEY);
      }
      return false;
    },

    saveSession(token, user) {
      state.token = token;
      state.user = user;
      localStorage.setItem(SESSION_KEY, JSON.stringify({ token, user }));
    },

    clearSession() {
      state.token = '';
      state.user = null;
      localStorage.removeItem(SESSION_KEY);
    },

    async login(event) {
      event.preventDefault();
      const username = $('#loginUsername').value.trim();
      const password = $('#loginPassword').value;
      const button = $('#loginButton');
      const error = $('#loginError');

      error.textContent = '';
      button.classList.add('is-loading');
      button.disabled = true;

      try {
        const result = await Api.expect('login', { username, password });
        Auth.saveSession(result.token, result.user);
        applyBootstrap(result);
        $('#loginForm').reset();
        showApp();
        toast('Signed in successfully.', 'success');
      } catch (err) {
        error.textContent = err.message;
      } finally {
        button.classList.remove('is-loading');
        button.disabled = false;
      }
    },

    async logout() {
      const token = state.token;
      Auth.clearSession();
      stopPolling();
      showLogin();
      if (token) {
        try {
          await Api.call('logout', token);
        } catch (error) {
          console.warn(error);
        }
      }
    }
  };

  function init() {
    bindEvents();
    applyTheme(localStorage.getItem(THEME_KEY) || 'light');

    if (Auth.loadSession()) {
      showApp();
      refreshWorkspace({ quiet: true }).catch(() => {
        Auth.clearSession();
        showLogin();
      });
    } else {
      showLogin();
    }
  }

  function bindEvents() {
    $('#loginForm').addEventListener('submit', Auth.login);
    $('#logoutButton').addEventListener('click', Auth.logout);
    $('#refreshButton').addEventListener('click', () => refreshWorkspace());
    $('#themeToggle').addEventListener('click', toggleTheme);
    $('#mobileMenuButton').addEventListener('click', openSidebar);
    $('#sidebarScrim').addEventListener('click', closeSidebar);

    $$('[data-page]').forEach(button => {
      button.addEventListener('click', () => navigate(button.dataset.page));
    });

    $('#newUploadButton').addEventListener('click', () => openModal('uploadModal'));
    $('#newInsightButton').addEventListener('click', () => openModal('insightModal'));
    $('#newUserButton').addEventListener('click', () => openModal('userModal'));

    $$('.modal-close').forEach(button => {
      button.addEventListener('click', () => closeModal(button.closest('dialog').id));
    });

    $$('dialog.modal').forEach(dialog => {
      dialog.addEventListener('click', event => {
        if (event.target === dialog) closeModal(dialog.id);
      });
    });

    $('#uploadForm').addEventListener('submit', handleUploadSubmit);
    $('#insightForm').addEventListener('submit', handleInsightSubmit);
    $('#commentForm').addEventListener('submit', handleCommentSubmit);
    $('#userForm').addEventListener('submit', handleUserSubmit);
    $('#resetPasswordForm').addEventListener('submit', handlePasswordReset);
    $('#passwordForm').addEventListener('submit', handlePasswordChange);

    ['uploadSearch', 'uploadCountryFilter', 'uploadCategoryFilter', 'uploadTypeFilter', 'uploadAuthorFilter', 'uploadTagFilter', 'uploadFromFilter', 'uploadToFilter']
      .forEach(id => $('#' + id).addEventListener('input', renderUploads));

    ['insightSearch', 'insightCountryFilter', 'insightCategoryFilter']
      .forEach(id => $('#' + id).addEventListener('input', renderInsights));

    $('#insightsList').addEventListener('click', event => {
      const button = event.target.closest('[data-comment-insight]');
      if (button) openComments(button.dataset.commentInsight);
    });

    $('#usersList').addEventListener('click', handleUserAction);
    $('#usersList').addEventListener('change', handleRoleChange);

    document.addEventListener('visibilitychange', () => {
      if (!document.hidden && state.token) refreshWorkspace({ quiet: true });
    });
  }

  function showLogin() {
    $('#loginView').classList.remove('is-hidden');
    $('#appView').classList.add('is-hidden');
    $('#loginUsername').focus();
  }

  function showApp() {
    $('#loginView').classList.add('is-hidden');
    $('#appView').classList.remove('is-hidden');
    syncUserChrome();
    syncPermissions();
    navigate(state.page || 'dashboard');
    startPolling();
  }

  function syncUserChrome() {
    if (!state.user) return;
    $('#userName').textContent = state.user.username || 'User';
    $('#userRole').textContent = state.user.role || 'researcher';
    $('#userInitials').textContent = initials(state.user.username || 'CI');
    $('#settingsUsername').textContent = state.user.username || '-';
    $('#settingsEmail').textContent = state.user.email || '-';
    $('#settingsRole').textContent = state.user.role || '-';
  }

  function syncPermissions() {
    const isAdmin = state.user && state.user.role === 'admin';
    $$('.admin-only').forEach(node => node.classList.toggle('is-hidden', !isAdmin));
    if (!isAdmin && state.page === 'users') navigate('dashboard');
  }

  function navigate(page) {
    if (page === 'users' && (!state.user || state.user.role !== 'admin')) {
      toast('Users is available only to administrators.', 'error');
      page = 'dashboard';
    }

    state.page = page;
    closeSidebar();

    $$('.nav-item').forEach(item => item.classList.toggle('is-active', item.dataset.page === page));
    $$('.page').forEach(panel => panel.classList.toggle('is-active', panel.dataset.pagePanel === page));

    const [eyebrow, title] = pageMeta[page] || pageMeta.dashboard;
    $('#pageEyebrow').textContent = eyebrow;
    $('#pageTitle').textContent = title;

    if (page === 'uploads') renderUploads();
    if (page === 'insights') renderInsights();
    if (page === 'countries') renderCountries();
    if (page === 'activity') renderActivity();
    if (page === 'users') renderUsers();
  }

  async function refreshWorkspace(options = {}) {
    if (!state.token || state.busy) return;
    setLoading(true, options.quiet);
    try {
      const result = await Api.expect('getBootstrap', state.token);
      applyBootstrap(result);
      renderAll();
      if (!options.quiet) toast('Workspace refreshed.', 'success');
    } catch (error) {
      if (!options.quiet) toast(error.message, 'error');
      if (/session|token|expired|unauthorized/i.test(error.message)) {
        Auth.clearSession();
        showLogin();
      }
      throw error;
    } finally {
      setLoading(false, options.quiet);
    }
  }

  function applyBootstrap(result) {
    if (result.user) {
      state.user = result.user;
      localStorage.setItem(SESSION_KEY, JSON.stringify({ token: state.token || result.token, user: state.user }));
    }
    if (result.token) state.token = result.token;
    if (result.links) state.links = { ...state.links, ...result.links };
    if (result.data) {
      state.data = {
        uploads: result.data.uploads || [],
        insights: result.data.insights || [],
        activities: result.data.activities || [],
        users: result.data.users || [],
        countries: result.data.countries || [],
        summary: result.data.summary || {}
      };
    }
    syncUserChrome();
    syncPermissions();
    hydrateDatalists();
    hydrateFilters();
  }

  function renderAll() {
    renderDashboard();
    renderUploads();
    renderInsights();
    renderCountries();
    renderActivity();
    renderUsers();
  }

  function renderDashboard() {
    const summary = state.data.summary || {};
    const stats = [
      ['Total uploads', summary.totalUploads || 0, 'stored sources', 'accent'],
      ['Total insights', summary.totalInsights || 0, 'team notes', 'blue'],
      ['Countries monitored', summary.totalCountries || 0, 'active coverage', 'gold'],
      ['Total users', summary.totalUsers || 0, `${summary.activeUsers || 0} active`, 'rose']
    ];

    $('#statsGrid').innerHTML = stats.map(([label, value, foot, tone]) => `
      <article class="stat-card">
        <div class="stat-top">
          <div>
            <p class="stat-label">${escapeHtml(label)}</p>
            <p class="stat-value">${Number(value).toLocaleString()}</p>
          </div>
          <span class="stat-chip ${tone}">${escapeHtml(foot)}</span>
        </div>
        <div class="stat-foot">${escapeHtml(statFoot(label, summary))}</div>
      </article>
    `).join('');

    $('#overviewCharts').innerHTML = [
      chartTemplate('Uploads by type', summary.uploadsByType || {}),
      chartTemplate('Top categories', summary.topCategories || {}),
      chartTemplate('Research status', summary.uploadsByStatus || {})
    ].join('');

    renderStack('#recentUploads', state.data.uploads.slice(0, 5), upload => ({
      title: upload.title,
      meta: [labelize(upload.type), upload.country, formatDate(upload.date)]
    }));

    renderStack('#recentInsights', state.data.insights.slice(0, 5), insight => ({
      title: insight.title,
      meta: [insight.author, insight.country || 'Global', `${insight.commentCount || 0} comments`]
    }));

    renderTimeline('#dashboardActivity', state.data.activities.slice(0, 6));
    $('#lastRefresh').textContent = `Updated ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`;
  }

  function statFoot(label, summary) {
    if (label === 'Total uploads') return `${summary.highPriorityUploads || 0} high priority sources`;
    if (label === 'Total insights') return `${summary.totalComments || 0} comments in discussions`;
    if (label === 'Countries monitored') return `${summary.coveredCountries || 0} countries with evidence`;
    return 'Admin controlled access';
  }

  function chartTemplate(title, values) {
    const entries = Object.entries(values || {}).sort((a, b) => b[1] - a[1]).slice(0, 6);
    const max = Math.max(1, ...entries.map(([, value]) => Number(value) || 0));
    return `
      <div class="mini-chart">
        <h3>${escapeHtml(title)}</h3>
        ${entries.length ? entries.map(([label, value]) => `
          <div class="bar-row">
            <span title="${escapeHtml(labelize(label))}">${escapeHtml(labelize(label))}</span>
            <div class="bar-track"><div class="bar-fill" style="width:${Math.max(4, (Number(value) || 0) / max * 100)}%"></div></div>
            <strong>${Number(value) || 0}</strong>
          </div>
        `).join('') : '<div class="empty-state">No data yet</div>'}
      </div>
    `;
  }

  function renderUploads() {
    const uploads = filteredUploads();
    const container = $('#uploadsList');
    if (!uploads.length) {
      container.innerHTML = '<div class="empty-state">No uploads match the current filters.</div>';
      return;
    }

    container.innerHTML = uploads.map(upload => `
      <article class="data-card">
        <div>
          <h3 class="card-title">${escapeHtml(upload.title)}</h3>
          ${upload.description ? `<p class="card-description">${escapeHtml(upload.description)}</p>` : ''}
          <div class="badge-row">
            ${badge(upload.country, 'accent')}
            ${badge(upload.category, 'gold')}
            ${badge(labelize(upload.type), 'blue')}
            ${badge(labelize(upload.priority), priorityTone(upload.priority))}
            ${badge(labelize(upload.status), '')}
            ${(upload.tags || []).map(tag => badge('#' + tag, '')).join('')}
          </div>
          <div class="meta-line" style="margin-top:12px">
            <span>${escapeHtml(upload.author || 'Unknown author')}</span>
            <span>${formatDate(upload.date)}</span>
          </div>
        </div>
        <div class="card-actions">
          ${safeLink(upload.gdriveLink, 'Google Drive')}
          ${safeLink(upload.externalLink, 'External link')}
        </div>
      </article>
    `).join('');
  }

  function filteredUploads() {
    const search = $('#uploadSearch').value.trim().toLowerCase();
    const country = $('#uploadCountryFilter').value;
    const category = $('#uploadCategoryFilter').value;
    const type = $('#uploadTypeFilter').value;
    const author = $('#uploadAuthorFilter').value;
    const tag = $('#uploadTagFilter').value.trim().toLowerCase();
    const from = $('#uploadFromFilter').value ? new Date($('#uploadFromFilter').value + 'T00:00:00') : null;
    const to = $('#uploadToFilter').value ? new Date($('#uploadToFilter').value + 'T23:59:59') : null;

    return state.data.uploads.filter(upload => {
      const haystack = [upload.title, upload.description, upload.country, upload.category, upload.author, ...(upload.tags || [])].join(' ').toLowerCase();
      const date = upload.date ? new Date(upload.date) : null;
      return (!search || haystack.includes(search))
        && (!country || upload.country === country)
        && (!category || upload.category === category)
        && (!type || type === 'all' || upload.type === type)
        && (!author || upload.author === author)
        && (!tag || (upload.tags || []).some(value => value.toLowerCase().includes(tag)))
        && (!from || (date && date >= from))
        && (!to || (date && date <= to));
    });
  }

  function renderInsights() {
    const search = $('#insightSearch').value.trim().toLowerCase();
    const country = $('#insightCountryFilter').value;
    const category = $('#insightCategoryFilter').value;
    const insights = state.data.insights.filter(insight => {
      const haystack = [insight.title, insight.content, insight.country, insight.category, insight.author, ...(insight.tags || [])].join(' ').toLowerCase();
      return (!search || haystack.includes(search))
        && (!country || insight.country === country)
        && (!category || insight.category === category);
    });

    const container = $('#insightsList');
    if (!insights.length) {
      container.innerHTML = '<div class="empty-state">No insights match the current filters.</div>';
      return;
    }

    container.innerHTML = insights.map(insight => `
      <article class="insight-card">
        <div class="surface-header">
          <div>
            <h3>${escapeHtml(insight.title)}</h3>
            <div class="meta-line">
              <span>${escapeHtml(insight.author || 'Unknown author')}</span>
              <span>${formatDate(insight.date)}</span>
              <span>${escapeHtml(insight.country || 'Global')}</span>
            </div>
          </div>
          <button class="button button-subtle" data-comment-insight="${escapeAttribute(insight.id)}" type="button">${Number(insight.commentCount || 0)} comments</button>
        </div>
        <p>${escapeHtml(insight.content)}</p>
        <div class="badge-row">
          ${insight.category ? badge(insight.category, 'gold') : ''}
          ${badge(labelize(insight.status || 'discussion'), 'blue')}
          ${badge(labelize(insight.priority || 'medium'), priorityTone(insight.priority))}
          ${(insight.tags || []).map(tag => badge('#' + tag, '')).join('')}
        </div>
      </article>
    `).join('');
  }

  function renderCountries() {
    const container = $('#countriesGrid');
    const countries = state.data.countries || [];
    if (!countries.length) {
      container.innerHTML = '<div class="empty-state">No monitored countries yet.</div>';
      return;
    }

    container.innerHTML = countries.map(country => `
      <article class="country-card">
        <div>
          <p class="eyebrow">${escapeHtml(country.region || 'Coffee market')}</p>
          <h3>${escapeHtml(country.name)}</h3>
        </div>
        <div class="country-stats">
          <div><span>Uploads</span><strong>${country.uploads || 0}</strong></div>
          <div><span>Insights</span><strong>${country.insights || 0}</strong></div>
          <div><span>Categories</span><strong>${(country.topCategories || []).length}</strong></div>
        </div>
        <div class="badge-row">
          ${(country.topCategories || []).slice(0, 5).map(item => badge(`${item.label} ${item.count}`, 'gold')).join('') || badge('No categories yet', '')}
        </div>
      </article>
    `).join('');
  }

  function renderActivity() {
    renderTimeline('#activityTimeline', state.data.activities);
  }

  function renderUsers() {
    const container = $('#usersList');
    if (!state.user || state.user.role !== 'admin') {
      container.innerHTML = '<div class="empty-state">Access denied. Administrator permission is required.</div>';
      return;
    }

    const users = state.data.users || [];
    if (!users.length) {
      container.innerHTML = '<div class="empty-state">No users registered.</div>';
      return;
    }

    container.innerHTML = `
      <table class="data-table">
        <thead>
          <tr>
            <th>Username</th>
            <th>Email</th>
            <th>Role</th>
            <th>Status</th>
            <th>Last login</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>
          ${users.map(user => `
            <tr>
              <td><strong>${escapeHtml(user.username)}</strong></td>
              <td>${escapeHtml(user.email || '-')}</td>
              <td>
                <select class="role-select" data-role-user="${escapeAttribute(user.username)}" ${user.username === state.user.username ? 'disabled' : ''}>
                  <option value="researcher" ${user.role === 'researcher' ? 'selected' : ''}>Researcher</option>
                  <option value="admin" ${user.role === 'admin' ? 'selected' : ''}>Admin</option>
                </select>
              </td>
              <td>${badge(user.active ? 'Active' : 'Inactive', user.active ? 'accent' : 'rose')}</td>
              <td>${formatDate(user.lastLogin)}</td>
              <td>
                <div class="table-actions">
                  <button class="button button-subtle" data-user-action="reset" data-username="${escapeAttribute(user.username)}" type="button">Reset</button>
                  <button class="button button-ghost" data-user-action="toggle" data-active="${user.active ? 'false' : 'true'}" data-username="${escapeAttribute(user.username)}" type="button" ${user.username === state.user.username ? 'disabled' : ''}>${user.active ? 'Deactivate' : 'Activate'}</button>
                </div>
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;
  }

  function renderStack(selector, items, mapper) {
    const container = $(selector);
    if (!items.length) {
      container.innerHTML = '<div class="empty-state">No records yet.</div>';
      return;
    }

    container.innerHTML = items.map(item => {
      const mapped = mapper(item);
      return `
        <div class="stack-item">
          <strong>${escapeHtml(mapped.title || 'Untitled')}</strong>
          <div class="meta-line">${mapped.meta.filter(Boolean).map(value => `<span>${escapeHtml(value)}</span>`).join('')}</div>
        </div>
      `;
    }).join('');
  }

  function renderTimeline(selector, activities) {
    const container = $(selector);
    if (!activities || !activities.length) {
      container.innerHTML = '<div class="empty-state">No activity recorded yet.</div>';
      return;
    }

    container.innerHTML = activities.map(activity => `
      <div class="timeline-item">
        <strong>${escapeHtml(activity.description || activity.type || 'Activity')}</strong>
        <span>${escapeHtml(activity.username || 'System')} - ${formatDateTime(activity.date)}</span>
      </div>
    `).join('');
  }

  async function handleUploadSubmit(event) {
    event.preventDefault();
    await mutate('saveUpload', [{
      title: $('#uploadTitle').value,
      description: $('#uploadDescription').value,
      category: $('#uploadCategory').value,
      country: $('#uploadCountry').value,
      type: $('#uploadType').value,
      priority: $('#uploadPriority').value,
      status: $('#uploadStatus').value,
      tags: splitTags($('#uploadTags').value),
      gdriveLink: $('#uploadDriveLink').value,
      externalLink: $('#uploadExternalLink').value
    }], 'Upload created.');
    $('#uploadForm').reset();
    closeModal('uploadModal');
  }

  async function handleInsightSubmit(event) {
    event.preventDefault();
    await mutate('saveInsight', [{
      title: $('#insightTitle').value,
      content: $('#insightContent').value,
      category: $('#insightCategory').value,
      country: $('#insightCountry').value,
      priority: $('#insightPriority').value,
      status: $('#insightStatus').value,
      tags: splitTags($('#insightTags').value)
    }], 'Insight created.');
    $('#insightForm').reset();
    closeModal('insightModal');
  }

  async function handleCommentSubmit(event) {
    event.preventDefault();
    const insightId = $('#commentForm').dataset.insightId;
    await mutate('saveComment', [insightId, { text: $('#commentText').value }], 'Comment posted.');
    $('#commentText').value = '';
    closeModal('commentsModal');
  }

  async function handleUserSubmit(event) {
    event.preventDefault();
    await mutate('createUser', [{
      username: $('#newUsername').value,
      email: $('#newUserEmail').value,
      password: $('#newUserPassword').value,
      role: $('#newUserRole').value
    }], 'User created.');
    $('#userForm').reset();
    closeModal('userModal');
  }

  async function handleUserAction(event) {
    const button = event.target.closest('[data-user-action]');
    if (!button) return;

    const username = button.dataset.username;
    if (button.dataset.userAction === 'reset') {
      $('#resetPasswordUsername').value = username;
      $('#resetPasswordValue').value = '';
      openModal('resetPasswordModal');
      return;
    }

    if (button.dataset.userAction === 'toggle') {
      const active = button.dataset.active === 'true';
      await mutate('setUserStatus', [username, active], active ? 'User activated.' : 'User deactivated.');
    }
  }

  async function handleRoleChange(event) {
    const select = event.target.closest('[data-role-user]');
    if (!select) return;
    await mutate('updateUserRole', [select.dataset.roleUser, select.value], 'Permission updated.');
  }

  async function handlePasswordReset(event) {
    event.preventDefault();
    await mutate('resetUserPassword', [$('#resetPasswordUsername').value, $('#resetPasswordValue').value], 'Password reset.');
    closeModal('resetPasswordModal');
  }

  async function handlePasswordChange(event) {
    event.preventDefault();
    const currentPassword = $('#currentPassword').value;
    const newPassword = $('#newPassword').value;
    const confirmPassword = $('#confirmPassword').value;

    if (newPassword !== confirmPassword) {
      toast('New passwords do not match.', 'error');
      return;
    }

    await mutate('changePassword', [{ currentPassword, newPassword }], 'Password updated.');
    $('#passwordForm').reset();
  }

  async function mutate(functionName, args, successMessage) {
    setLoading(true, false);
    try {
      const result = await Api.expect(functionName, state.token, ...args);
      applyBootstrap(result);
      renderAll();
      toast(successMessage, 'success');
    } catch (error) {
      toast(error.message, 'error');
      throw error;
    } finally {
      setLoading(false, false);
    }
  }

  function openComments(insightId) {
    const insight = state.data.insights.find(item => item.id === insightId);
    if (!insight) return;

    $('#commentsTitle').textContent = insight.title;
    $('#commentForm').dataset.insightId = insight.id;
    const comments = insight.comments || [];
    $('#commentsList').innerHTML = comments.length ? comments.map(comment => `
      <div class="comment-item">
        <strong>${escapeHtml(comment.author || 'Unknown author')}</strong>
        <span class="muted">${formatDateTime(comment.date)}</span>
        <p>${escapeHtml(comment.text)}</p>
      </div>
    `).join('') : '<div class="empty-state">No comments yet.</div>';

    openModal('commentsModal');
  }

  function hydrateFilters() {
    setSelectOptions('#uploadCountryFilter', [''].concat(unique(state.data.uploads.map(item => item.country))), 'All countries');
    setSelectOptions('#uploadCategoryFilter', [''].concat(unique(state.data.uploads.map(item => item.category))), 'All categories');
    setSelectOptions('#uploadTypeFilter', contentTypes.map(([value]) => value), '', contentTypes);
    setSelectOptions('#uploadAuthorFilter', [''].concat(unique(state.data.uploads.map(item => item.author))), 'All authors');
    setSelectOptions('#insightCountryFilter', [''].concat(unique(state.data.insights.map(item => item.country))), 'All countries');
    setSelectOptions('#insightCategoryFilter', [''].concat(unique(state.data.insights.map(item => item.category))), 'All categories');
  }

  function hydrateDatalists() {
    const countries = unique([
      ...state.data.countries.map(item => item.name),
      ...state.data.uploads.map(item => item.country),
      ...state.data.insights.map(item => item.country)
    ]);
    const categories = unique([
      ...state.data.uploads.map(item => item.category),
      ...state.data.insights.map(item => item.category)
    ]);

    $('#countrySuggestions').innerHTML = countries.map(value => `<option value="${escapeAttribute(value)}"></option>`).join('');
    $('#categorySuggestions').innerHTML = categories.map(value => `<option value="${escapeAttribute(value)}"></option>`).join('');
  }

  function setSelectOptions(selector, values, emptyLabel, pairs) {
    const select = $(selector);
    const current = select.value;
    if (pairs) {
      select.innerHTML = pairs.map(([value, label]) => `<option value="${escapeAttribute(value === 'all' ? '' : value)}">${escapeHtml(label)}</option>`).join('');
    } else {
      select.innerHTML = values.filter(Boolean).length || emptyLabel
        ? values.map((value, index) => `<option value="${escapeAttribute(value)}">${escapeHtml(index === 0 && value === '' ? emptyLabel : value)}</option>`).join('')
        : '';
    }
    if ([...select.options].some(option => option.value === current)) select.value = current;
  }

  function startPolling() {
    stopPolling();
    state.poller = window.setInterval(() => {
      if (!document.hidden && state.token) refreshWorkspace({ quiet: true }).catch(() => {});
    }, POLL_INTERVAL);
  }

  function stopPolling() {
    if (state.poller) window.clearInterval(state.poller);
    state.poller = null;
  }

  function openModal(id) {
    const dialog = $('#' + id);
    if (dialog && !dialog.open) dialog.showModal();
  }

  function closeModal(id) {
    const dialog = $('#' + id);
    if (dialog && dialog.open) dialog.close();
  }

  function openSidebar() {
    $('#sidebar').classList.add('is-open');
    $('#sidebarScrim').classList.add('is-open');
  }

  function closeSidebar() {
    $('#sidebar').classList.remove('is-open');
    $('#sidebarScrim').classList.remove('is-open');
  }

  function setLoading(isLoading, quiet) {
    state.busy = isLoading;
    $('#loadingState').classList.toggle('is-hidden', quiet || !isLoading);
    $$('button[type="submit"], #refreshButton').forEach(button => {
      if (button.id !== 'loginButton') button.disabled = isLoading;
    });
  }

  function toggleTheme() {
    const next = document.body.classList.contains('theme-dark') ? 'light' : 'dark';
    applyTheme(next);
  }

  function applyTheme(theme) {
    const dark = theme === 'dark';
    document.body.classList.toggle('theme-dark', dark);
    localStorage.setItem(THEME_KEY, dark ? 'dark' : 'light');
    $('#themeToggle').textContent = dark ? 'Light mode' : 'Dark mode';
  }

  function toast(message, type = 'info') {
    const node = document.createElement('div');
    node.className = `toast ${type}`;
    node.textContent = message;
    $('#toastRegion').appendChild(node);
    setTimeout(() => node.remove(), 4200);
  }

  function badge(value, tone) {
    if (!value) return '';
    return `<span class="badge ${tone || ''}">${escapeHtml(value)}</span>`;
  }

  function priorityTone(priority) {
    if (priority === 'critical' || priority === 'high') return 'rose';
    if (priority === 'medium') return 'gold';
    return '';
  }

  function safeLink(url, label) {
    if (!url) return '';
    try {
      const parsed = new URL(url);
      if (!['http:', 'https:'].includes(parsed.protocol)) return '';
      return `<a class="link-pill" href="${escapeAttribute(parsed.href)}" target="_blank" rel="noopener">${escapeHtml(label)}</a>`;
    } catch (error) {
      return '';
    }
  }

  function splitTags(value) {
    return String(value || '')
      .split(',')
      .map(tag => tag.trim())
      .filter(Boolean);
  }

  function appendHidden(form, name, value) {
    const input = document.createElement('input');
    input.type = 'hidden';
    input.name = name;
    input.value = value;
    form.appendChild(input);
  }

  function unique(values) {
    return [...new Set(values.map(value => String(value || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  }

  function labelize(value) {
    return String(value || '')
      .replace(/_/g, ' ')
      .replace(/\b\w/g, char => char.toUpperCase());
  }

  function initials(value) {
    const parts = String(value || 'CI').trim().split(/[\s._-]+/).filter(Boolean);
    return (parts[0]?.[0] || 'C').toUpperCase() + (parts[1]?.[0] || parts[0]?.[1] || 'I').toUpperCase();
  }

  function formatDate(value) {
    if (!value) return '-';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return '-';
    return date.toLocaleDateString([], { year: 'numeric', month: 'short', day: '2-digit' });
  }

  function formatDateTime(value) {
    if (!value) return '-';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return '-';
    return date.toLocaleString([], { year: 'numeric', month: 'short', day: '2-digit', hour: '2-digit', minute: '2-digit' });
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function escapeAttribute(value) {
    return escapeHtml(value).replace(/`/g, '&#096;');
  }

  return { init };
})();

document.addEventListener('DOMContentLoaded', CoffeeIntelligence.init);
