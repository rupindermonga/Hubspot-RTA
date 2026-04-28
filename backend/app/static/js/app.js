/* Hubspot RTA SPA — Alpine root.

   Two services share one shell: Address Matcher and RTA Status Compare.
   Both follow the same flow: pick file → peek columns → map → run → review → download.
*/

function app() {
  const self = {};

  // ── Top-level state ─────────────────────────────────────────────────
  self.view = 'login';                 // 'login' | 'matcher' | 'compare'
  self.token = localStorage.getItem('hubspot_rta_token') || null;
  self.user = null;
  self.loginForm = { username: '', password: '' };
  self.authError = '';
  self.authLoading = false;

  // ── Auth + helpers ──────────────────────────────────────────────────
  self.api = async function(path, options) {
    const opts = options || {};
    const headers = Object.assign({}, opts.headers || {});
    if (self.token) headers['Authorization'] = 'Bearer ' + self.token;
    return fetch(path, Object.assign({}, opts, { headers: headers }));
  };

  self.init = async function() {
    if (self.token) {
      try {
        const r = await self.api('/api/auth/me');
        if (r.ok) {
          self.user = await r.json();
          self.view = 'matcher';
          return;
        }
      } catch (e) { /* fall through to clear token */ }
      self.token = null;
      localStorage.removeItem('hubspot_rta_token');
    }
  };

  self.login = async function() {
    self.authError = '';
    self.authLoading = true;
    try {
      const r = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(self.loginForm),
      });
      if (!r.ok) {
        const body = await r.json().catch(() => ({}));
        self.authError = body.detail || ('Login failed (HTTP ' + r.status + ')');
        return;
      }
      const data = await r.json();
      self.token = data.access_token;
      localStorage.setItem('hubspot_rta_token', self.token);
      self.user = data.user;
      self.loginForm = { username: '', password: '' };
      self.view = 'matcher';
    } catch (e) {
      self.authError = 'Network error: ' + e.message;
    } finally {
      self.authLoading = false;
    }
  };

  self.logout = function() {
    self.token = null;
    self.user = null;
    localStorage.removeItem('hubspot_rta_token');
    self.view = 'login';
  };

  // Returns CSS class for a status-change row in the compare view
  self.statusChangeClass = function(row) {
    const oldS = (row['Old Status'] || '').toString().toUpperCase().trim();
    const newS = (row['New Status'] || '').toString().toUpperCase().trim();
    if (oldS === 'IN CONSTRUCTION' && newS === 'RTA') return 'status-cell-forward';
    if (oldS === 'RTA' && newS === 'IN CONSTRUCTION') return 'status-cell-regress';
    return '';
  };

  // ── Shared helpers ──────────────────────────────────────────────────
  function pickDefault(columns, preferred) {
    for (let i = 0; i < preferred.length; i++) {
      if (columns.indexOf(preferred[i]) !== -1) return preferred[i];
    }
    return columns[0] || '';
  }

  async function downloadBlob(token, filename, errorTarget) {
    try {
      const r = await self.api('/api/' + errorTarget.service + '/download/' + token);
      if (!r.ok) {
        const body = await r.json().catch(() => ({}));
        errorTarget.target.error = body.detail || ('Download failed (HTTP ' + r.status + ')');
        return;
      }
      const blob = await r.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      errorTarget.target.error = 'Download error: ' + e.message;
    }
  }

  async function peekFile(file, target) {
    const fd = new FormData();
    fd.append('file', file);
    const r = await self.api('/api/' + target.service + '/peek', { method: 'POST', body: fd });
    if (!r.ok) {
      const body = await r.json().catch(() => ({}));
      throw new Error(body.detail || ('Peek failed (HTTP ' + r.status + ')'));
    }
    return r.json();
  }

  // ── Matcher view ────────────────────────────────────────────────────
  self.matcher = {
    service: 'matcher',
    hubFile: null,
    rtaFile: null,
    hubColumns: [],
    rtaColumns: [],
    hubSheets: null,
    rtaSheets: null,
    hubSheet: '',
    rtaSheet: '',
    colMap: {
      hub_street_col: '',
      hub_pc_col: '',
      rta_addr_no_col: '',
      rta_street_col: '',
      rta_locality_col: '',
      rta_pc_col: '',
      rta_status_col: '',
    },
    enableNoPc: false,
    running: false,
    result: null,
    error: '',

    async onHubFile(e) {
      const file = e.target.files[0];
      if (!file) return;
      this.hubFile = file;
      this.error = '';
      try {
        const peek = await peekFile(file, this);
        this.hubColumns = peek.columns;
        this.hubSheets = peek.sheet_names;
        this.hubSheet = (peek.sheet_names && peek.sheet_names[0]) || '';
        this.colMap.hub_street_col = pickDefault(peek.columns, ['Street Address', 'street_address', 'Street']);
        this.colMap.hub_pc_col = pickDefault(peek.columns, ['Postal Code', 'postal_code', 'Zip']);
      } catch (err) {
        this.error = err.message;
      }
    },

    async onRtaFile(e) {
      const file = e.target.files[0];
      if (!file) return;
      this.rtaFile = file;
      this.error = '';
      try {
        const peek = await peekFile(file, this);
        this.rtaColumns = peek.columns;
        this.rtaSheets = peek.sheet_names;
        this.rtaSheet = (peek.sheet_names && peek.sheet_names[0]) || '';
        this.colMap.rta_addr_no_col = pickDefault(peek.columns, ['AddressNo', 'Address Number']);
        this.colMap.rta_street_col = pickDefault(peek.columns, ['StreetName', 'Street Name']);
        this.colMap.rta_locality_col = pickDefault(peek.columns, ['Locality', 'City']);
        this.colMap.rta_pc_col = pickDefault(peek.columns, ['PostalCode', 'Postal Code']);
        // Default RTA Status to last column (matches Streamlit default)
        this.colMap.rta_status_col = pickDefault(peek.columns, ['RTA Status', 'Status']) || peek.columns[peek.columns.length - 1];
      } catch (err) {
        this.error = err.message;
      }
    },

    async run() {
      if (!this.hubFile || !this.rtaFile) return;
      this.running = true;
      this.error = '';
      this.result = null;
      try {
        const fd = new FormData();
        fd.append('hub_file', this.hubFile);
        fd.append('rta_file', this.rtaFile);
        Object.keys(this.colMap).forEach(k => fd.append(k, this.colMap[k]));
        if (this.hubSheet) fd.append('hub_sheet', this.hubSheet);
        if (this.rtaSheet) fd.append('rta_sheet', this.rtaSheet);
        fd.append('enable_no_pc', this.enableNoPc ? 'true' : 'false');

        const r = await self.api('/api/matcher/run', { method: 'POST', body: fd });
        if (!r.ok) {
          const body = await r.json().catch(() => ({}));
          this.error = (typeof body.detail === 'string') ? body.detail : JSON.stringify(body.detail || 'Run failed');
          return;
        }
        this.result = await r.json();
      } catch (e) {
        this.error = 'Network error: ' + e.message;
      } finally {
        this.running = false;
      }
    },

    download() {
      if (!this.result || !this.result.download_token) return;
      downloadBlob(
        this.result.download_token,
        'hubspot_rta_matched_output.xlsx',
        { service: 'matcher', target: this },
      );
    },
  };

  // ── Compare view ────────────────────────────────────────────────────
  self.compare = {
    service: 'compare',
    oldFile: null,
    newFile: null,
    oldColumns: [],
    newColumns: [],
    oldSheets: null,
    newSheets: null,
    oldSheet: '',
    newSheet: '',
    colMap: {
      addr_no_col: '',
      street_col: '',
      locality_col: '',
      pc_col: '',
      status_col: '',
    },
    running: false,
    result: null,
    error: '',

    async onOldFile(e) {
      const file = e.target.files[0];
      if (!file) return;
      this.oldFile = file;
      this.error = '';
      try {
        const peek = await peekFile(file, this);
        this.oldColumns = peek.columns;
        this.oldSheets = peek.sheet_names;
        this.oldSheet = (peek.sheet_names && peek.sheet_names[0]) || '';
        // Set default col map if not yet set (will be re-set when new file loads)
        if (!this.colMap.addr_no_col) {
          this.colMap.addr_no_col = pickDefault(peek.columns, ['AddressNo', 'Address Number']);
          this.colMap.street_col = pickDefault(peek.columns, ['StreetName', 'Street Name']);
          this.colMap.locality_col = pickDefault(peek.columns, ['Locality', 'City']);
          this.colMap.pc_col = pickDefault(peek.columns, ['PostalCode', 'Postal Code']);
          this.colMap.status_col = pickDefault(peek.columns, ['RTA Status', 'Status']) || peek.columns[peek.columns.length - 1];
        }
      } catch (err) {
        this.error = err.message;
      }
    },

    async onNewFile(e) {
      const file = e.target.files[0];
      if (!file) return;
      this.newFile = file;
      this.error = '';
      try {
        const peek = await peekFile(file, this);
        this.newColumns = peek.columns;
        this.newSheets = peek.sheet_names;
        this.newSheet = (peek.sheet_names && peek.sheet_names[0]) || '';
        // Authoritative defaults from the New file (the column dropdowns bind to newColumns)
        this.colMap.addr_no_col = pickDefault(peek.columns, ['AddressNo', 'Address Number']);
        this.colMap.street_col = pickDefault(peek.columns, ['StreetName', 'Street Name']);
        this.colMap.locality_col = pickDefault(peek.columns, ['Locality', 'City']);
        this.colMap.pc_col = pickDefault(peek.columns, ['PostalCode', 'Postal Code']);
        this.colMap.status_col = pickDefault(peek.columns, ['RTA Status', 'Status']) || peek.columns[peek.columns.length - 1];
      } catch (err) {
        this.error = err.message;
      }
    },

    async run() {
      if (!this.oldFile || !this.newFile) return;
      this.running = true;
      this.error = '';
      this.result = null;
      try {
        const fd = new FormData();
        fd.append('old_file', this.oldFile);
        fd.append('new_file', this.newFile);
        Object.keys(this.colMap).forEach(k => fd.append(k, this.colMap[k]));
        if (this.oldSheet) fd.append('old_sheet', this.oldSheet);
        if (this.newSheet) fd.append('new_sheet', this.newSheet);

        const r = await self.api('/api/compare/run', { method: 'POST', body: fd });
        if (!r.ok) {
          const body = await r.json().catch(() => ({}));
          this.error = (typeof body.detail === 'string') ? body.detail : JSON.stringify(body.detail || 'Run failed');
          return;
        }
        this.result = await r.json();
      } catch (e) {
        this.error = 'Network error: ' + e.message;
      } finally {
        this.running = false;
      }
    },

    download() {
      if (!this.result || !this.result.download_token) return;
      downloadBlob(
        this.result.download_token,
        'rta_status_compare_output.xlsx',
        { service: 'compare', target: this },
      );
    },
  };

  return self;
}
