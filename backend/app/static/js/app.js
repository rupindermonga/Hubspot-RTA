/* Hubspot RTA SPA — Alpine root.

   Two services share one shell: Address Matcher and RTA Status Compare.
   Both follow the same flow: pick file → peek columns → map → run → review → download.

   Alpine reactivity note: every mutation goes through `this`, which Alpine
   binds to the proxied scope. Nested matcher/compare objects get a `_root`
   handle in init() so their methods can call back into the root's `api()`
   helper (and the auth token it carries).
*/

function app() {
  return {
    // ── Top-level state ─────────────────────────────────────────────
    // 'landing' (public marketing) | 'login' | 'matcher' | 'compare'
    view: 'landing',
    token: localStorage.getItem('hubspot_rta_token') || null,
    user: null,
    loginForm: { username: '', password: '' },
    authError: '',
    authLoading: false,
    aliases: [],   // populated from /api/matcher/aliases after auth
    datasets: [],  // populated from /api/matcher/datasets after auth — encrypted-at-rest list

    // ── Lifecycle ──────────────────────────────────────────────────
    async init() {
      // Bind nested namespaces back to the root proxy so their methods
      // can reach `this.token`, `this.api`, etc. without closure tricks.
      this.matcher._root = this;
      this.compare._root = this;

      if (this.token) {
        try {
          const r = await this.api('/api/auth/me');
          if (r.ok) {
            this.user = await r.json();
            this.view = 'matcher';
            await this.loadAliases();
            await this.loadDatasets();
            return;
          }
        } catch (e) { /* fall through to clear */ }
        this.token = null;
        localStorage.removeItem('hubspot_rta_token');
      }
    },

    async loadAliases() {
      try {
        const r = await this.api('/api/matcher/aliases');
        if (r.ok) {
          const data = await r.json();
          this.aliases = data.aliases || [];
        }
      } catch (e) { /* non-fatal */ }
    },

    async loadDatasets() {
      try {
        const r = await this.api('/api/matcher/datasets');
        if (r.ok) {
          const data = await r.json();
          this.datasets = data.datasets || [];
        }
      } catch (e) { /* non-fatal */ }
    },

    formatTimestamp(iso) {
      if (!iso) return '';
      const d = new Date(iso);
      return d.toLocaleString();
    },

    // ── Auth + helpers ─────────────────────────────────────────────
    async api(path, options) {
      const opts = options || {};
      const headers = Object.assign({}, opts.headers || {});
      if (this.token) headers['Authorization'] = 'Bearer ' + this.token;
      return fetch(path, Object.assign({}, opts, { headers: headers }));
    },

    async login() {
      this.authError = '';
      this.authLoading = true;
      try {
        const r = await fetch('/api/auth/login', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(this.loginForm),
        });
        if (!r.ok) {
          const body = await r.json().catch(() => ({}));
          this.authError = body.detail || ('Login failed (HTTP ' + r.status + ')');
          return;
        }
        const data = await r.json();
        this.token = data.access_token;
        localStorage.setItem('hubspot_rta_token', this.token);
        this.user = data.user;
        this.loginForm = { username: '', password: '' };
        this.view = 'matcher';
        await this.loadAliases();
        await this.loadDatasets();
      } catch (e) {
        this.authError = 'Network error: ' + e.message;
      } finally {
        this.authLoading = false;
      }
    },

    logout() {
      this.token = null;
      this.user = null;
      this.datasets = [];
      localStorage.removeItem('hubspot_rta_token');
      this.view = 'landing';
    },

    statusChangeClass(row) {
      const oldS = (row['Old Status'] || '').toString().toUpperCase().trim();
      const newS = (row['New Status'] || '').toString().toUpperCase().trim();
      if (oldS === 'IN CONSTRUCTION' && newS === 'RTA') return 'status-cell-forward';
      if (oldS === 'RTA' && newS === 'IN CONSTRUCTION') return 'status-cell-regress';
      return '';
    },

    // ── Matcher view ───────────────────────────────────────────────
    matcher: {
      _root: null,
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
      label: '',
      running: false,
      result: null,
      error: '',
      // Search state
      searchQuery: '',
      searchPostal: '',
      searchScope: 'all',  // 'all' | <dataset_id as string>
      searchRunning: false,
      searchResult: null,
      searchError: '',

      _pickDefault(columns, preferred) {
        for (let i = 0; i < preferred.length; i++) {
          if (columns.indexOf(preferred[i]) !== -1) return preferred[i];
        }
        return columns[0] || '';
      },

      async _peek(file) {
        const fd = new FormData();
        fd.append('file', file);
        const r = await this._root.api('/api/matcher/peek', { method: 'POST', body: fd });
        if (!r.ok) {
          const body = await r.json().catch(() => ({}));
          throw new Error(body.detail || ('Peek failed (HTTP ' + r.status + ')'));
        }
        return r.json();
      },

      async onHubFile(e) {
        const file = e.target.files[0];
        if (!file) return;
        this.hubFile = file;
        this.error = '';
        try {
          const peek = await this._peek(file);
          this.hubColumns = peek.columns;
          this.hubSheets = peek.sheet_names;
          this.hubSheet = (peek.sheet_names && peek.sheet_names[0]) || '';
          this.colMap.hub_street_col = this._pickDefault(peek.columns, ['Street Address', 'street_address', 'Street']);
          this.colMap.hub_pc_col = this._pickDefault(peek.columns, ['Postal Code', 'postal_code', 'Zip']);
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
          const peek = await this._peek(file);
          this.rtaColumns = peek.columns;
          this.rtaSheets = peek.sheet_names;
          this.rtaSheet = (peek.sheet_names && peek.sheet_names[0]) || '';
          this.colMap.rta_addr_no_col = this._pickDefault(peek.columns, ['AddressNo', 'Address Number']);
          this.colMap.rta_street_col = this._pickDefault(peek.columns, ['StreetName', 'Street Name']);
          this.colMap.rta_locality_col = this._pickDefault(peek.columns, ['Locality', 'City']);
          this.colMap.rta_pc_col = this._pickDefault(peek.columns, ['PostalCode', 'Postal Code']);
          this.colMap.rta_status_col = this._pickDefault(peek.columns, ['RTA Status', 'Status']) || peek.columns[peek.columns.length - 1];
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
          if (this.label) fd.append('label', this.label);

          const r = await this._root.api('/api/matcher/run', { method: 'POST', body: fd });
          if (!r.ok) {
            const body = await r.json().catch(() => ({}));
            this.error = (typeof body.detail === 'string') ? body.detail : JSON.stringify(body.detail || 'Run failed');
            return;
          }
          this.result = await r.json();
          await this._root.loadDatasets();
        } catch (e) {
          this.error = 'Network error: ' + e.message;
        } finally {
          this.running = false;
        }
      },

      async selectDataset(id) {
        this.error = '';
        this.result = null;
        try {
          const r = await this._root.api('/api/matcher/datasets/' + id);
          if (!r.ok) {
            const body = await r.json().catch(() => ({}));
            this.error = body.detail || ('Could not load dataset (HTTP ' + r.status + ')');
            return;
          }
          this.result = await r.json();
        } catch (e) {
          this.error = 'Network error: ' + e.message;
        }
      },

      async deleteDataset(id) {
        if (!confirm('Delete this saved dataset? This cannot be undone.')) return;
        try {
          const r = await this._root.api('/api/matcher/datasets/' + id, { method: 'DELETE' });
          if (!r.ok) {
            const body = await r.json().catch(() => ({}));
            this.error = body.detail || ('Delete failed (HTTP ' + r.status + ')');
            return;
          }
          if (this.result && this.result.dataset_id === id) {
            this.result = null;
          }
          await this._root.loadDatasets();
        } catch (e) {
          this.error = 'Network error: ' + e.message;
        }
      },

      async download() {
        const id = this.result && this.result.dataset_id;
        if (!id) return;
        try {
          const r = await this._root.api('/api/matcher/download/' + id);
          if (!r.ok) {
            const body = await r.json().catch(() => ({}));
            this.error = body.detail || ('Download failed (HTTP ' + r.status + ')');
            return;
          }
          const blob = await r.blob();
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = 'hubspot_rta_dataset_' + id + '.xlsx';
          document.body.appendChild(a);
          a.click();
          a.remove();
          URL.revokeObjectURL(url);
        } catch (e) {
          this.error = 'Download error: ' + e.message;
        }
      },

      async runSearch() {
        const q = (this.searchQuery || '').trim();
        if (!q) return;
        this.searchRunning = true;
        this.searchError = '';
        this.searchResult = null;
        try {
          const params = new URLSearchParams({ q: q });
          if (this.searchPostal) params.set('postal', this.searchPostal.trim());
          if (this.searchScope && this.searchScope !== 'all') params.set('dataset_id', this.searchScope);
          const r = await this._root.api('/api/matcher/search?' + params.toString());
          if (!r.ok) {
            const body = await r.json().catch(() => ({}));
            this.searchError = (typeof body.detail === 'string') ? body.detail : JSON.stringify(body.detail || 'Search failed');
            return;
          }
          this.searchResult = await r.json();
        } catch (e) {
          this.searchError = 'Network error: ' + e.message;
        } finally {
          this.searchRunning = false;
        }
      },

      verdictLabel(v) {
        return ({
          'both': 'In Hubspot AND in RTA',
          'hubspot_only': 'Hubspot only — no RTA record',
          'rta_only': 'RTA only — no Hubspot contact',
          'neither': 'Not found in any saved dataset',
        })[v] || v;
      },
      verdictBadgeClass(v) {
        return ({
          'both': 'bg-green-100 text-green-800 border-green-300',
          'hubspot_only': 'bg-blue-100 text-blue-800 border-blue-300',
          'rta_only': 'bg-purple-100 text-purple-800 border-purple-300',
          'neither': 'bg-gray-100 text-gray-700 border-gray-300',
        })[v] || 'bg-gray-100 text-gray-700 border-gray-300';
      },
    },

    // ── Compare view ───────────────────────────────────────────────
    compare: {
      _root: null,
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

      _pickDefault(columns, preferred) {
        for (let i = 0; i < preferred.length; i++) {
          if (columns.indexOf(preferred[i]) !== -1) return preferred[i];
        }
        return columns[0] || '';
      },

      async _peek(file) {
        const fd = new FormData();
        fd.append('file', file);
        const r = await this._root.api('/api/compare/peek', { method: 'POST', body: fd });
        if (!r.ok) {
          const body = await r.json().catch(() => ({}));
          throw new Error(body.detail || ('Peek failed (HTTP ' + r.status + ')'));
        }
        return r.json();
      },

      async onOldFile(e) {
        const file = e.target.files[0];
        if (!file) return;
        this.oldFile = file;
        this.error = '';
        try {
          const peek = await this._peek(file);
          this.oldColumns = peek.columns;
          this.oldSheets = peek.sheet_names;
          this.oldSheet = (peek.sheet_names && peek.sheet_names[0]) || '';
          if (!this.colMap.addr_no_col) {
            this.colMap.addr_no_col = this._pickDefault(peek.columns, ['AddressNo', 'Address Number']);
            this.colMap.street_col = this._pickDefault(peek.columns, ['StreetName', 'Street Name']);
            this.colMap.locality_col = this._pickDefault(peek.columns, ['Locality', 'City']);
            this.colMap.pc_col = this._pickDefault(peek.columns, ['PostalCode', 'Postal Code']);
            this.colMap.status_col = this._pickDefault(peek.columns, ['RTA Status', 'Status']) || peek.columns[peek.columns.length - 1];
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
          const peek = await this._peek(file);
          this.newColumns = peek.columns;
          this.newSheets = peek.sheet_names;
          this.newSheet = (peek.sheet_names && peek.sheet_names[0]) || '';
          this.colMap.addr_no_col = this._pickDefault(peek.columns, ['AddressNo', 'Address Number']);
          this.colMap.street_col = this._pickDefault(peek.columns, ['StreetName', 'Street Name']);
          this.colMap.locality_col = this._pickDefault(peek.columns, ['Locality', 'City']);
          this.colMap.pc_col = this._pickDefault(peek.columns, ['PostalCode', 'Postal Code']);
          this.colMap.status_col = this._pickDefault(peek.columns, ['RTA Status', 'Status']) || peek.columns[peek.columns.length - 1];
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

          const r = await this._root.api('/api/compare/run', { method: 'POST', body: fd });
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

      async download() {
        if (!this.result || !this.result.download_token) return;
        try {
          const r = await this._root.api('/api/compare/download/' + this.result.download_token);
          if (!r.ok) {
            const body = await r.json().catch(() => ({}));
            this.error = body.detail || ('Download failed (HTTP ' + r.status + ')');
            return;
          }
          const blob = await r.blob();
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = 'rta_status_compare_output.xlsx';
          document.body.appendChild(a);
          a.click();
          a.remove();
          URL.revokeObjectURL(url);
        } catch (e) {
          this.error = 'Download error: ' + e.message;
        }
      },
    },
  };
}
