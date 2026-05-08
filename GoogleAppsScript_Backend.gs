// Coffee Intelligence Research System
// Google Apps Script API backend for a GitHub Pages frontend.
//
// Deploy as Web App:
// Execute as: Me
// Who has access: Anyone

const SPREADSHEET_ID = '1cZ7iit2zpPsE_gDcJyi2h2UBvVK64TOlMuuBR8xFflg';
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwwoLfaAVBdrH9l7myWTZ3rvWlvO0NZRi1cwXISK4_2RO1DV5CxpjfBlo2qRF8kMsz_/exec';
const SHEETS_URL = 'https://docs.google.com/spreadsheets/d/1cZ7iit2zpPsE_gDcJyi2h2UBvVK64TOlMuuBR8xFflg/edit?usp=sharing';
const SESSION_DAYS = 30;

const TABLES = {
  USERS: {
    name: 'Users',
    headers: ['ID', 'Username', 'Password', 'Email', 'Role', 'Active', 'CreatedAt', 'LastLogin']
  },
  SESSIONS: {
    name: 'Sessions',
    headers: ['Token', 'Username', 'CreatedAt', 'LastSeen', 'ExpiresAt', 'Active']
  },
  UPLOADS: {
    name: 'Uploads',
    headers: ['ID', 'Title', 'Description', 'Category', 'Country', 'Type', 'Priority', 'Status', 'Tags', 'GDrive Link', 'External Link', 'Author', 'Date', 'UpdatedAt']
  },
  INSIGHTS: {
    name: 'Insights',
    headers: ['ID', 'Title', 'Content', 'Category', 'Country', 'Priority', 'Status', 'Tags', 'Author', 'Date', 'UpdatedAt']
  },
  COMMENTS: {
    name: 'Comments',
    headers: ['ID', 'Insight ID', 'Author', 'Text', 'Date']
  },
  ACTIVITY: {
    name: 'Activity',
    headers: ['ID', 'Username', 'Type', 'Description', 'Date']
  },
  COUNTRIES: {
    name: 'Countries',
    headers: ['Country', 'Region']
  }
};

const DEFAULT_COUNTRIES = [
  ['Brazil', 'South America'],
  ['Colombia', 'South America'],
  ['Vietnam', 'Asia'],
  ['Indonesia', 'Asia'],
  ['Ethiopia', 'Africa'],
  ['Honduras', 'Central America'],
  ['India', 'Asia'],
  ['Uganda', 'Africa'],
  ['Peru', 'South America'],
  ['Guatemala', 'Central America'],
  ['Mexico', 'North America'],
  ['Costa Rica', 'Central America'],
  ['Kenya', 'Africa'],
  ['Tanzania', 'Africa'],
  ['Rwanda', 'Africa']
];

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  initializeDatabase_();

  if (!params.action) {
    return jsonOutput_({
      success: true,
      app: 'Coffee Intelligence Research System',
      mode: 'api',
      links: apiLinks_()
    }, params.callback);
  }

  let response;
  try {
    const args = parsePayload_(params.payload);
    response = dispatch_(params.action, args);
  } catch (error) {
    response = {
      success: false,
      error: error && error.message ? error.message : String(error)
    };
  }

  return jsonOutput_(response, params.callback);
}

function doPost(e) {
  const params = e && e.parameter ? e.parameter : {};
  initializeDatabase_();

  let response;
  try {
    const args = parsePayload_(params.payload);
    response = dispatch_(params.action, args);
  } catch (error) {
    response = {
      success: false,
      error: error && error.message ? error.message : String(error)
    };
  }

  return HtmlService
    .createHtmlOutput(postMessageHtml_(response, params.requestId))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function dispatch_(action, args) {
  switch (action) {
    case 'login':
      return login_(args[0] || {});
    case 'logout':
      return logout_(args[0]);
    case 'getBootstrap':
      return withSession_(args[0], function(user) {
        return bootstrap_(user);
      });
    case 'saveUpload':
      return withSession_(args[0], function(user) {
        saveUpload_(user, args[1] || {});
        return bootstrap_(user);
      });
    case 'saveInsight':
      return withSession_(args[0], function(user) {
        saveInsight_(user, args[1] || {});
        return bootstrap_(user);
      });
    case 'saveComment':
      return withSession_(args[0], function(user) {
        saveComment_(user, args[1], args[2] || {});
        return bootstrap_(user);
      });
    case 'createUser':
      return withSession_(args[0], function(user) {
        requireAdmin_(user);
        createUser_(args[1] || {}, user);
        return bootstrap_(user);
      });
    case 'setUserStatus':
      return withSession_(args[0], function(user) {
        requireAdmin_(user);
        setUserStatus_(String(args[1] || ''), args[2] === true, user);
        return bootstrap_(user);
      });
    case 'updateUserRole':
      return withSession_(args[0], function(user) {
        requireAdmin_(user);
        updateUserRole_(String(args[1] || ''), String(args[2] || ''), user);
        return bootstrap_(user);
      });
    case 'resetUserPassword':
      return withSession_(args[0], function(user) {
        requireAdmin_(user);
        resetUserPassword_(String(args[1] || ''), String(args[2] || ''), user);
        return bootstrap_(user);
      });
    case 'changePassword':
      return withSession_(args[0], function(user) {
        changePassword_(user, args[1] || {});
        return bootstrap_(user);
      });
    default:
      throw new Error('Unknown API action: ' + action);
  }
}

function login_(credentials) {
  const username = String(credentials.username || '').trim();
  const password = String(credentials.password || '');

  if (!username || !password) {
    throw new Error('Username and password are required.');
  }

  const userRow = findUserRow_(username);
  if (!userRow) {
    throw new Error('Invalid username or password.');
  }

  const user = userFromRecord_(userRow.record);
  if (!user.active) {
    throw new Error('This user is inactive. Contact an administrator.');
  }

  if (!passwordMatches_(userRow.record.Password, password)) {
    throw new Error('Invalid username or password.');
  }

  if (!String(userRow.record.Password || '').startsWith('sha256:')) {
    updateCell_(userRow.sheet, userRow.row, 'Password', hashPassword_(password));
  }

  const token = createSession_(user.username);
  updateCell_(userRow.sheet, userRow.row, 'LastLogin', new Date().toISOString());
  logActivity_(user.username, 'login', user.username + ' signed in');

  return Object.assign(bootstrap_(user), {
    token: token,
    user: user
  });
}

function logout_(token) {
  if (!token) return { success: true };
  const session = findSessionRow_(token);
  if (session) {
    updateCell_(session.sheet, session.row, 'Active', false);
  }
  return { success: true };
}

function bootstrap_(user) {
  const uploads = getUploads_();
  const insights = getInsights_();
  const comments = getComments_();
  const countries = getCountries_(uploads, insights);
  const activities = getActivities_();
  const users = user.role === 'admin' ? getUsers_() : [];

  const insightComments = comments.reduce(function(acc, comment) {
    const key = comment.insightId;
    if (!acc[key]) acc[key] = [];
    acc[key].push(comment);
    return acc;
  }, {});

  const enrichedInsights = insights.map(function(insight) {
    const threaded = insightComments[insight.id] || [];
    return Object.assign({}, insight, {
      comments: threaded,
      commentCount: threaded.length
    });
  });

  return {
    success: true,
    user: user,
    links: apiLinks_(),
    data: {
      uploads: uploads,
      insights: enrichedInsights,
      activities: activities,
      users: users,
      countries: countries,
      summary: buildSummary_(uploads, enrichedInsights, comments, countries, getUsers_())
    }
  };
}

function saveUpload_(user, data) {
  const title = cleanText_(data.title);
  const category = cleanText_(data.category);
  const country = cleanText_(data.country);
  const type = normalizeChoice_(data.type, ['news', 'report', 'pdf', 'spreadsheet', 'scientific_article', 'reference', 'observation', 'external_link', 'google_drive'], 'reference');

  if (!title || !category || !country) {
    throw new Error('Title, category, and country are required.');
  }

  appendRecord_(TABLES.UPLOADS, {
    ID: Utilities.getUuid(),
    Title: title,
    Description: cleanText_(data.description),
    Category: category,
    Country: country,
    Type: type,
    Priority: normalizeChoice_(data.priority, ['low', 'medium', 'high', 'critical'], 'medium'),
    Status: normalizeChoice_(data.status, ['new', 'reviewing', 'validated', 'archived'], 'new'),
    Tags: normalizeTags_(data.tags).join(', '),
    'GDrive Link': cleanUrl_(data.gdriveLink),
    'External Link': cleanUrl_(data.externalLink),
    Author: user.username,
    Date: new Date().toISOString(),
    UpdatedAt: new Date().toISOString()
  });

  ensureCountry_(country);
  logActivity_(user.username, 'upload_created', user.username + ' created upload: ' + title);
}

function saveInsight_(user, data) {
  const title = cleanText_(data.title);
  const content = cleanText_(data.content);

  if (!title || !content) {
    throw new Error('Title and insight content are required.');
  }

  const country = cleanText_(data.country);
  appendRecord_(TABLES.INSIGHTS, {
    ID: Utilities.getUuid(),
    Title: title,
    Content: content,
    Category: cleanText_(data.category),
    Country: country,
    Priority: normalizeChoice_(data.priority, ['low', 'medium', 'high', 'critical'], 'medium'),
    Status: normalizeChoice_(data.status, ['draft', 'discussion', 'validated', 'archived'], 'discussion'),
    Tags: normalizeTags_(data.tags).join(', '),
    Author: user.username,
    Date: new Date().toISOString(),
    UpdatedAt: new Date().toISOString()
  });

  if (country) ensureCountry_(country);
  logActivity_(user.username, 'insight_created', user.username + ' created insight: ' + title);
}

function saveComment_(user, insightId, data) {
  const text = cleanText_(data.text);
  if (!insightId || !text) {
    throw new Error('Insight and comment text are required.');
  }

  const insight = getInsights_().filter(function(item) {
    return item.id === insightId;
  })[0];

  if (!insight) {
    throw new Error('Insight not found.');
  }

  appendRecord_(TABLES.COMMENTS, {
    ID: Utilities.getUuid(),
    'Insight ID': insightId,
    Author: user.username,
    Text: text,
    Date: new Date().toISOString()
  });

  logActivity_(user.username, 'comment_created', user.username + ' commented on insight: ' + insight.title);
}

function createUser_(data, actor) {
  const username = String(data.username || '').trim();
  const email = String(data.email || '').trim();
  const password = String(data.password || '');
  const role = normalizeChoice_(data.role, ['admin', 'researcher'], 'researcher');

  if (!username || !email || password.length < 6) {
    throw new Error('Username, email, and a password with at least 6 characters are required.');
  }

  if (findUserRow_(username)) {
    throw new Error('A user with this username already exists.');
  }

  appendRecord_(TABLES.USERS, {
    ID: Utilities.getUuid(),
    Username: username,
    Password: hashPassword_(password),
    Email: email,
    Role: role,
    Active: true,
    CreatedAt: new Date().toISOString(),
    LastLogin: ''
  });

  logActivity_(actor.username, 'user_created', actor.username + ' created user: ' + username);
}

function setUserStatus_(username, active, actor) {
  const target = findUserRow_(username);
  if (!target) throw new Error('User not found.');
  if (username === actor.username && active === false) {
    throw new Error('You cannot deactivate your own account.');
  }

  updateCell_(target.sheet, target.row, 'Active', active);
  logActivity_(actor.username, active ? 'user_activated' : 'user_deactivated', actor.username + (active ? ' activated ' : ' deactivated ') + username);
}

function updateUserRole_(username, role, actor) {
  const normalized = normalizeChoice_(role, ['admin', 'researcher'], '');
  if (!normalized) throw new Error('Invalid role.');
  const target = findUserRow_(username);
  if (!target) throw new Error('User not found.');
  if (username === actor.username) {
    throw new Error('You cannot change your own permission.');
  }

  updateCell_(target.sheet, target.row, 'Role', normalized);
  logActivity_(actor.username, 'user_role_changed', actor.username + ' changed ' + username + ' to ' + normalized);
}

function resetUserPassword_(username, password, actor) {
  if (!password || password.length < 6) {
    throw new Error('Password must have at least 6 characters.');
  }
  const target = findUserRow_(username);
  if (!target) throw new Error('User not found.');

  updateCell_(target.sheet, target.row, 'Password', hashPassword_(password));
  logActivity_(actor.username, 'password_reset', actor.username + ' reset password for ' + username);
}

function changePassword_(user, data) {
  const currentPassword = String(data.currentPassword || '');
  const newPassword = String(data.newPassword || '');
  if (newPassword.length < 6) {
    throw new Error('New password must have at least 6 characters.');
  }

  const target = findUserRow_(user.username);
  if (!target || !passwordMatches_(target.record.Password, currentPassword)) {
    throw new Error('Current password is incorrect.');
  }

  updateCell_(target.sheet, target.row, 'Password', hashPassword_(newPassword));
  logActivity_(user.username, 'password_changed', user.username + ' changed their password');
}

function withSession_(token, callback) {
  const user = getSessionUser_(token);
  if (!user) {
    throw new Error('Session expired. Sign in again.');
  }
  return callback(user);
}

function requireAdmin_(user) {
  if (!user || user.role !== 'admin') {
    throw new Error('Administrator permission is required.');
  }
}

function getSessionUser_(token) {
  token = String(token || '').trim();
  if (!token) return null;

  const session = findSessionRow_(token);
  if (!session || !parseBool_(session.record.Active)) return null;

  const expiresAt = new Date(session.record.ExpiresAt);
  if (expiresAt.getTime() && expiresAt.getTime() < Date.now()) {
    updateCell_(session.sheet, session.row, 'Active', false);
    return null;
  }

  const userRow = findUserRow_(session.record.Username);
  if (!userRow) return null;
  const user = userFromRecord_(userRow.record);
  if (!user.active) return null;

  updateCell_(session.sheet, session.row, 'LastSeen', new Date().toISOString());
  updateCell_(session.sheet, session.row, 'ExpiresAt', futureDate_(SESSION_DAYS).toISOString());
  return user;
}

function createSession_(username) {
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  appendRecord_(TABLES.SESSIONS, {
    Token: token,
    Username: username,
    CreatedAt: new Date().toISOString(),
    LastSeen: new Date().toISOString(),
    ExpiresAt: futureDate_(SESSION_DAYS).toISOString(),
    Active: true
  });
  return token;
}

function getUploads_() {
  return records_(TABLES.UPLOADS).map(function(record) {
    return {
      id: String(record.ID || ''),
      title: String(record.Title || ''),
      description: String(record.Description || ''),
      category: String(record.Category || ''),
      country: String(record.Country || ''),
      type: String(record.Type || ''),
      priority: String(record.Priority || 'medium'),
      status: String(record.Status || 'new'),
      tags: splitTags_(record.Tags),
      gdriveLink: String(record['GDrive Link'] || ''),
      externalLink: String(record['External Link'] || ''),
      author: String(record.Author || ''),
      date: normalizeDate_(record.Date),
      updatedAt: normalizeDate_(record.UpdatedAt)
    };
  }).filter(function(upload) {
    return upload.title;
  }).sort(sortByDateDesc_);
}

function getInsights_() {
  return records_(TABLES.INSIGHTS).map(function(record) {
    return {
      id: String(record.ID || ''),
      title: String(record.Title || ''),
      content: String(record.Content || ''),
      category: String(record.Category || ''),
      country: String(record.Country || ''),
      priority: String(record.Priority || 'medium'),
      status: String(record.Status || 'discussion'),
      tags: splitTags_(record.Tags),
      author: String(record.Author || ''),
      date: normalizeDate_(record.Date),
      updatedAt: normalizeDate_(record.UpdatedAt)
    };
  }).filter(function(insight) {
    return insight.id && insight.title;
  }).sort(sortByDateDesc_);
}

function getComments_() {
  return records_(TABLES.COMMENTS).map(function(record) {
    return {
      id: String(record.ID || ''),
      insightId: String(record['Insight ID'] || record['Insight Title'] || ''),
      author: String(record.Author || ''),
      text: String(record.Text || ''),
      date: normalizeDate_(record.Date)
    };
  }).filter(function(comment) {
    return comment.insightId && comment.text;
  }).sort(sortByDateAsc_);
}

function getActivities_() {
  return records_(TABLES.ACTIVITY).map(function(record) {
    return {
      id: String(record.ID || ''),
      username: String(record.Username || ''),
      type: String(record.Type || 'activity'),
      description: String(record.Description || ''),
      date: normalizeDate_(record.Date)
    };
  }).filter(function(activity) {
    return activity.description;
  }).sort(sortByDateDesc_).slice(0, 200);
}

function getUsers_() {
  return records_(TABLES.USERS).map(userFromRecord_).filter(function(user) {
    return user.username;
  }).sort(function(a, b) {
    return a.username.localeCompare(b.username);
  });
}

function getCountries_(uploads, insights) {
  const countryRecords = records_(TABLES.COUNTRIES);
  const byName = {};

  countryRecords.forEach(function(record) {
    const name = String(record.Country || '').trim();
    if (name) {
      byName[name] = {
        name: name,
        region: String(record.Region || 'Global'),
        uploads: 0,
        insights: 0,
        topCategories: []
      };
    }
  });

  uploads.forEach(function(upload) {
    if (!upload.country) return;
    if (!byName[upload.country]) byName[upload.country] = countryBase_(upload.country);
    byName[upload.country].uploads += 1;
  });

  insights.forEach(function(insight) {
    if (!insight.country) return;
    if (!byName[insight.country]) byName[insight.country] = countryBase_(insight.country);
    byName[insight.country].insights += 1;
  });

  Object.keys(byName).forEach(function(name) {
    const categories = {};
    uploads.filter(function(upload) {
      return upload.country === name && upload.category;
    }).forEach(function(upload) {
      categories[upload.category] = (categories[upload.category] || 0) + 1;
    });
    insights.filter(function(insight) {
      return insight.country === name && insight.category;
    }).forEach(function(insight) {
      categories[insight.category] = (categories[insight.category] || 0) + 1;
    });
    byName[name].topCategories = topEntries_(categories, 5);
  });

  return Object.keys(byName).sort().map(function(name) {
    return byName[name];
  });
}

function buildSummary_(uploads, insights, comments, countries, users) {
  const categories = {};
  uploads.forEach(function(upload) {
    if (upload.category) categories[upload.category] = (categories[upload.category] || 0) + 1;
  });
  insights.forEach(function(insight) {
    if (insight.category) categories[insight.category] = (categories[insight.category] || 0) + 1;
  });

  return {
    totalUploads: uploads.length,
    totalInsights: insights.length,
    totalCountries: countries.length,
    coveredCountries: countries.filter(function(country) {
      return country.uploads || country.insights;
    }).length,
    totalUsers: users.length,
    activeUsers: users.filter(function(user) { return user.active; }).length,
    totalComments: comments.length,
    highPriorityUploads: uploads.filter(function(upload) {
      return upload.priority === 'high' || upload.priority === 'critical';
    }).length,
    uploadsByType: countBy_(uploads, 'type'),
    uploadsByStatus: countBy_(uploads, 'status'),
    topCategories: mapEntries_(topEntries_(categories, 8))
  };
}

function initializeDatabase_() {
  Object.keys(TABLES).forEach(function(key) {
    ensureSheet_(TABLES[key]);
  });

  if (records_(TABLES.USERS).length === 0) {
    appendRecord_(TABLES.USERS, {
      ID: Utilities.getUuid(),
      Username: 'admin',
      Password: hashPassword_('admin123'),
      Email: 'admin@coffee-intelligence.local',
      Role: 'admin',
      Active: true,
      CreatedAt: new Date().toISOString(),
      LastLogin: ''
    });
  }

  if (records_(TABLES.COUNTRIES).length === 0) {
    DEFAULT_COUNTRIES.forEach(function(country) {
      appendRecord_(TABLES.COUNTRIES, {
        Country: country[0],
        Region: country[1]
      });
    });
  }
}

function ensureCountry_(countryName) {
  const name = cleanText_(countryName);
  if (!name) return;
  const exists = records_(TABLES.COUNTRIES).some(function(record) {
    return String(record.Country || '').toLowerCase() === name.toLowerCase();
  });
  if (!exists) {
    appendRecord_(TABLES.COUNTRIES, {
      Country: name,
      Region: 'Global'
    });
  }
}

function findUserRow_(username) {
  return findRow_(TABLES.USERS, 'Username', username);
}

function findSessionRow_(token) {
  return findRow_(TABLES.SESSIONS, 'Token', token);
}

function findRow_(table, key, value) {
  const sheet = ensureSheet_(table);
  const headers = getHeaders_(sheet);
  const data = getDataRows_(sheet);
  const colIndex = headers.indexOf(key);
  if (colIndex === -1) return null;

  const expected = String(value || '').toLowerCase();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][colIndex] || '').toLowerCase() === expected) {
      return {
        sheet: sheet,
        row: i + 2,
        record: rowToRecord_(headers, data[i])
      };
    }
  }
  return null;
}

function records_(table) {
  const sheet = ensureSheet_(table);
  const headers = getHeaders_(sheet);
  return getDataRows_(sheet).map(function(row) {
    return rowToRecord_(headers, row);
  });
}

function appendRecord_(table, record) {
  const sheet = ensureSheet_(table);
  const headers = getHeaders_(sheet);
  const row = headers.map(function(header) {
    return record[header] !== undefined ? record[header] : '';
  });
  sheet.appendRow(row);
}

function updateCell_(sheet, row, header, value) {
  const headers = getHeaders_(sheet);
  const col = headers.indexOf(header) + 1;
  if (col <= 0) throw new Error('Missing column: ' + header);
  sheet.getRange(row, col).setValue(value);
}

function ensureSheet_(table) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(table.name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(table.name);
  }

  const lastColumn = sheet.getLastColumn();
  let headers = [];
  if (lastColumn > 0) {
    headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(value) {
      return String(value || '').trim();
    });
  }

  if (!headers.length || headers.every(function(header) { return !header; })) {
    sheet.getRange(1, 1, 1, table.headers.length).setValues([table.headers]);
    return sheet;
  }

  const missing = table.headers.filter(function(header) {
    return headers.indexOf(header) === -1;
  });

  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }

  return sheet;
}

function getHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return [];
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(value) {
    return String(value || '').trim();
  });
}

function getDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
}

function rowToRecord_(headers, row) {
  const record = {};
  headers.forEach(function(header, index) {
    if (header) record[header] = row[index];
  });
  return record;
}

function userFromRecord_(record) {
  return {
    id: String(record.ID || ''),
    username: String(record.Username || '').trim(),
    email: String(record.Email || '').trim(),
    role: normalizeChoice_(record.Role, ['admin', 'researcher'], 'researcher'),
    active: parseBool_(record.Active),
    createdAt: normalizeDate_(record.CreatedAt),
    lastLogin: normalizeDate_(record.LastLogin)
  };
}

function passwordMatches_(stored, password) {
  stored = String(stored || '');
  if (stored.startsWith('sha256:')) {
    return stored === hashPassword_(password);
  }
  return stored === String(password || '');
}

function hashPassword_(password) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password || ''),
    Utilities.Charset.UTF_8
  );
  const hex = digest.map(function(byte) {
    const value = byte < 0 ? byte + 256 : byte;
    return (value < 16 ? '0' : '') + value.toString(16);
  }).join('');
  return 'sha256:' + hex;
}

function logActivity_(username, type, description) {
  appendRecord_(TABLES.ACTIVITY, {
    ID: Utilities.getUuid(),
    Username: username,
    Type: type,
    Description: description,
    Date: new Date().toISOString()
  });
}

function jsonOutput_(payload, callback) {
  const json = JSON.stringify(payload).replace(/<\/script/gi, '<\\/script');
  const callbackName = String(callback || '').trim();

  if (callbackName && /^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callbackName)) {
    return ContentService
      .createTextOutput(callbackName + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function postMessageHtml_(payload, requestId) {
  const message = JSON.stringify({
    __cirsResponse: true,
    requestId: String(requestId || ''),
    response: payload
  }).replace(/</g, '\\u003c').replace(/<\/script/gi, '<\\/script');

  return '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>window.parent.postMessage(' + message + ', "*");</script>' +
    '</body></html>';
}

function parsePayload_(payload) {
  if (!payload) return [];
  const parsed = JSON.parse(payload);
  return Array.isArray(parsed) ? parsed : [parsed];
}

function apiLinks_() {
  return {
    sheet: SHEETS_URL,
    webApp: WEB_APP_URL
  };
}

function cleanText_(value) {
  return String(value || '').trim();
}

function cleanUrl_(value) {
  const url = String(value || '').trim();
  if (!url) return '';
  if (!/^https?:\/\//i.test(url)) {
    throw new Error('Links must start with http:// or https://.');
  }
  return url;
}

function normalizeTags_(tags) {
  if (Array.isArray(tags)) {
    return tags.map(cleanText_).filter(Boolean);
  }
  return splitTags_(tags);
}

function splitTags_(value) {
  if (Array.isArray(value)) return value.map(cleanText_).filter(Boolean);
  return String(value || '').split(',').map(cleanText_).filter(Boolean);
}

function normalizeChoice_(value, allowed, fallback) {
  const normalized = String(value || '').trim().toLowerCase();
  return allowed.indexOf(normalized) >= 0 ? normalized : fallback;
}

function parseBool_(value) {
  if (value === true) return true;
  const text = String(value || '').trim().toLowerCase();
  return text === 'true' || text === 'yes' || text === '1' || text === 'active';
}

function normalizeDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value.toISOString();
  }
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? String(value) : parsed.toISOString();
}

function futureDate_(days) {
  return new Date(Date.now() + days * 24 * 60 * 60 * 1000);
}

function sortByDateDesc_(a, b) {
  return new Date(b.date || 0).getTime() - new Date(a.date || 0).getTime();
}

function sortByDateAsc_(a, b) {
  return new Date(a.date || 0).getTime() - new Date(b.date || 0).getTime();
}

function countBy_(items, key) {
  return items.reduce(function(acc, item) {
    const value = item[key] || 'Unspecified';
    acc[value] = (acc[value] || 0) + 1;
    return acc;
  }, {});
}

function topEntries_(counts, limit) {
  return Object.keys(counts).map(function(label) {
    return {
      label: label,
      count: counts[label]
    };
  }).sort(function(a, b) {
    return b.count - a.count;
  }).slice(0, limit);
}

function mapEntries_(entries) {
  return entries.reduce(function(acc, entry) {
    acc[entry.label] = entry.count;
    return acc;
  }, {});
}

function countryBase_(name) {
  return {
    name: name,
    region: 'Global',
    uploads: 0,
    insights: 0,
    topCategories: []
  };
}

// Run manually from the Apps Script editor if you want to pre-create tabs.
function initializeSpreadsheet() {
  initializeDatabase_();
}
