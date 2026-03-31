const { EventEmitter } = require('events');
const sqlite3 = require('sqlite3').verbose();

class ToolingDatabase extends EventEmitter {
  constructor(options = {}) {
    super();
    this.getDb = options.getDb;
    this.resolveDatabasePath = options.resolveDatabasePath;
    this.replacementColumnName = options.replacementColumnName || 'replacement_tooling_id';
    this.replacementColumnType = options.replacementColumnType || 'INTEGER';
    this.supplierMetadataTable = options.supplierMetadataTable || 'supplier_metadata';
    this.silentUpdateFields = options.silentUpdateFields || new Set();
    this.helpers = options.helpers || {};
    this.replacementColumnEnsured = false;
    this.replacementColumnEnsuringPromise = null;
    this.supplierMetadataEnsured = false;
    this.supplierMetadataPromise = null;
    this.revisionTimers = Object.create(null);
    this.dbConnection = null;
  }

  get db() {
    const db = this.dbConnection || (typeof this.getDb === 'function' ? this.getDb() : null);
    if (!db) {
      throw new Error('Database connection not available.');
    }
    return db;
  }

  connect() {
    const dbPath = typeof this.resolveDatabasePath === 'function'
      ? this.resolveDatabasePath()
      : null;

    if (!dbPath) {
      return Promise.reject(new Error('Database path not available.'));
    }

    return new Promise((resolve, reject) => {
      this.dbConnection = new sqlite3.Database(dbPath, (err) => {
        if (err) {
          this.dbConnection = null;
          reject(err);
          return;
        }

        this.configureDatabasePragmas()
          .then(() => this.ensureBaseSchema())
          .then(resolve)
          .catch((schemaError) => {
            reject(schemaError);
          });
      });
    });
  }

  get(sql, params = []) {
    return new Promise((resolve, reject) => {
      this.db.get(sql, params, (err, row) => {
        if (err) {
          reject(err);
          return;
        }
        resolve(row);
      });
    });
  }

  all(sql, params = []) {
    return new Promise((resolve, reject) => {
      this.db.all(sql, params, (err, rows) => {
        if (err) {
          reject(err);
          return;
        }
        resolve(rows || []);
      });
    });
  }

  run(sql, params = []) {
    return new Promise((resolve, reject) => {
      this.db.run(sql, params, function onRun(err) {
        if (err) {
          reject(err);
          return;
        }
        resolve({
          changes: this.changes || 0,
          lastID: this.lastID || null
        });
      });
    });
  }

  emitChange(type, payload = {}) {
    const event = {
      type,
      ...payload,
      timestamp: new Date().toISOString()
    };

    this.emit('change', event);
    this.emit(type, event);
    return event;
  }

  configureDatabasePragmas() {
    return new Promise((resolve, reject) => {
      this.db.serialize(() => {
        this.db.run('PRAGMA busy_timeout = 8000');
        this.db.run('PRAGMA journal_mode = DELETE');
        this.db.run('PRAGMA synchronous = NORMAL', (err) => {
          if (err) {
            reject(err);
          } else {
            resolve();
          }
        });
      });
    });
  }

  async ensureBaseSchema() {
    await this.run(`CREATE TABLE IF NOT EXISTS ferramental (
      id INTEGER PRIMARY KEY AUTOINCREMENT
    )`);
    await this.checkAndAddColumns();
    await this.ensureTodosTable();
    await this.ensureStepHistoryTable();
  }

  async checkAndAddColumns() {
    const columnsToCheck = [
      { name: 'pn', type: 'TEXT' },
      { name: 'pn_description', type: 'TEXT' },
      { name: 'supplier', type: 'TEXT' },
      { name: 'tool_description', type: 'TEXT' },
      { name: 'tool_number_arb', type: 'TEXT' },
      { name: 'asset_number', type: 'TEXT' },
      { name: 'tool_ownership', type: 'TEXT' },
      { name: 'customer', type: 'TEXT' },
      { name: 'tooling_life_qty', type: 'TEXT' },
      { name: 'produced', type: 'TEXT' },
      { name: 'remaining_tooling_life_pcs', type: 'TEXT' },
      { name: 'percent_tooling_life', type: 'TEXT' },
      { name: 'annual_volume_forecast', type: 'TEXT' },
      { name: 'date_remaining_tooling_life', type: 'TEXT' },
      { name: 'date_annual_volume', type: 'TEXT' },
      { name: 'expiration_date', type: 'TEXT' },
      { name: 'finish_due_date', type: 'TEXT' },
      { name: 'amount_brl', type: 'TEXT' },
      { name: 'tool_quantity', type: 'TEXT' },
      { name: 'bailment_agreement_signed', type: 'TEXT' },
      { name: 'tooling_book', type: 'TEXT' },
      { name: 'disposition', type: 'TEXT' },
      { name: 'status', type: 'TEXT' },
      { name: 'steps', type: 'TEXT' },
      { name: 'bu', type: 'TEXT' },
      { name: 'category', type: 'TEXT' },
      { name: 'cummins_responsible', type: 'TEXT' },
      { name: 'comments', type: 'TEXT' },
      { name: 'todos', type: 'TEXT' },
      { name: 'stim_tooling_management', type: 'TEXT' },
      { name: 'invoice', type: 'TEXT' },
      { name: 'vpcr', type: 'TEXT' },
      { name: 'analysis_notes', type: 'TEXT' },
      { name: this.replacementColumnName, type: this.replacementColumnType },
      { name: 'analysis_completed', type: 'INTEGER' }
    ];

    const columns = await this.all('PRAGMA table_info(ferramental)', []);
    const existingColumns = new Set(columns.map((column) => column.name));
    if (existingColumns.has(this.replacementColumnName)) {
      this.replacementColumnEnsured = true;
    }

    const missingColumns = columnsToCheck.filter(({ name }) => !existingColumns.has(name));
    for (const { name, type } of missingColumns) {
      try {
        await this.run(`ALTER TABLE ferramental ADD COLUMN ${name} ${type || 'TEXT'}`);
      } catch (err) {
        if (!/duplicate column/i.test(err.message || '')) {
          throw err;
        }
      }

      if (name === this.replacementColumnName) {
        this.replacementColumnEnsured = true;
      }
    }
  }

  async ensureReplacementColumnExists(force = false) {
    if (force) {
      this.replacementColumnEnsured = false;
      this.replacementColumnEnsuringPromise = null;
    }

    if (this.replacementColumnEnsured) {
      return;
    }

    if (this.replacementColumnEnsuringPromise) {
      return this.replacementColumnEnsuringPromise;
    }

    this.replacementColumnEnsuringPromise = (async () => {
      const columns = await this.all('PRAGMA table_info(ferramental)', []);
      const hasColumn = Array.isArray(columns) && columns.some((column) => column.name === this.replacementColumnName);

      if (hasColumn) {
        this.replacementColumnEnsured = true;
        return;
      }

      await this.run(`ALTER TABLE ferramental ADD COLUMN ${this.replacementColumnName} ${this.replacementColumnType}`);
      this.replacementColumnEnsured = true;
    })();

    try {
      await this.replacementColumnEnsuringPromise;
    } finally {
      this.replacementColumnEnsuringPromise = null;
    }
  }

  isMissingReplacementColumnError(err) {
    if (!err || typeof err.message !== 'string') {
      return false;
    }

    return err.message.includes(`no such column: ${this.replacementColumnName}`);
  }

  async ensureSupplierMetadataTable() {
    if (this.supplierMetadataEnsured) {
      return this.supplierMetadataPromise || Promise.resolve();
    }

    this.supplierMetadataPromise = (async () => {
      await this.run(
        `CREATE TABLE IF NOT EXISTS ${this.supplierMetadataTable} (
          supplier TEXT PRIMARY KEY,
          last_import_timestamp TEXT
        )`
      );

      try {
        await this.run(
          `ALTER TABLE ${this.supplierMetadataTable} ADD COLUMN data_revision INTEGER DEFAULT 0`
        );
      } catch (err) {
        if (!/duplicate column/i.test(err.message || '')) {
          throw err;
        }
      }

      this.supplierMetadataEnsured = true;
    })();

    return this.supplierMetadataPromise;
  }

  async getSupplierImportTimestamp(supplierName) {
    if (!supplierName) {
      return null;
    }

    await this.ensureSupplierMetadataTable();
    const row = await this.get(
      `SELECT last_import_timestamp FROM ${this.supplierMetadataTable} WHERE supplier = ?`,
      [supplierName]
    );
    return row?.last_import_timestamp || null;
  }

  async setSupplierImportTimestamp(supplierName, timestamp) {
    if (!supplierName) {
      return;
    }

    await this.ensureSupplierMetadataTable();
    await this.run(
      `INSERT INTO ${this.supplierMetadataTable} (supplier, last_import_timestamp)
       VALUES (?, ?)
       ON CONFLICT(supplier) DO UPDATE SET last_import_timestamp = excluded.last_import_timestamp`,
      [supplierName, timestamp]
    );
  }

  async getDataRevision(supplierName) {
    if (!supplierName) {
      return 0;
    }

    await this.ensureSupplierMetadataTable();
    const row = await this.get(
      `SELECT data_revision FROM ${this.supplierMetadataTable} WHERE supplier = ?`,
      [supplierName]
    );
    return row?.data_revision || 0;
  }

  async incrementDataRevision(supplierName) {
    if (!supplierName) {
      return;
    }

    await this.ensureSupplierMetadataTable();
    await this.run(
      `INSERT INTO ${this.supplierMetadataTable} (supplier, data_revision)
       VALUES (?, 1)
       ON CONFLICT(supplier) DO UPDATE SET data_revision = COALESCE(data_revision, 0) + 1`,
      [supplierName]
    );
  }

  scheduleRevisionIncrement(supplierName) {
    if (!supplierName) {
      return;
    }

    const normalizedSupplierName = String(supplierName).trim();
    const key = normalizedSupplierName.toLowerCase();
    if (this.revisionTimers[key]) {
      clearTimeout(this.revisionTimers[key]);
    }

    this.revisionTimers[key] = setTimeout(async () => {
      delete this.revisionTimers[key];
      try {
        await this.incrementDataRevision(normalizedSupplierName);
      } catch (err) {
        console.error('[DataRevision] Debounced increment failed:', err);
      }
    }, 3000);
  }

  async ensureStepHistoryTable() {
    try {
      await this.run(
        `CREATE TABLE IF NOT EXISTS step_history (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          tooling_id INTEGER NOT NULL,
          old_step TEXT,
          new_step TEXT NOT NULL,
          changed_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          FOREIGN KEY (tooling_id) REFERENCES ferramental(id) ON DELETE CASCADE
        )`
      );
      console.log('[StepHistory] step_history table ensured');
    } catch (err) {
      console.error('[StepHistory] Error creating step_history table:', err);
      throw err;
    }
  }

  async recordStepChange(toolingId, oldStep, newStep) {
    if (oldStep === newStep) {
      return { recorded: false, reason: 'no_change' };
    }

    try {
      const result = await this.run(
        'INSERT INTO step_history (tooling_id, old_step, new_step) VALUES (?, ?, ?)',
        [toolingId, oldStep || null, newStep]
      );
      console.log(`[StepHistory] Recorded step change for tooling ${toolingId}: ${oldStep} -> ${newStep}`);
      return { recorded: true, id: result.lastID };
    } catch (err) {
      console.error('[StepHistory] Error recording step change:', err);
      throw err;
    }
  }

  async getStepHistory(toolingId) {
    try {
      return await this.all(
        'SELECT * FROM step_history WHERE tooling_id = ? ORDER BY changed_at DESC',
        [toolingId]
      );
    } catch (err) {
      console.error('[StepHistory] Error getting step history:', err);
      throw err;
    }
  }

  async clearStepHistory(toolingId) {
    try {
      const result = await this.run('DELETE FROM step_history WHERE tooling_id = ?', [toolingId]);
      console.log(`[StepHistory] Cleared ${result.changes} entries for tooling ${toolingId}`);
      return { success: true, deleted: result.changes };
    } catch (err) {
      console.error('[StepHistory] Error clearing step history:', err);
      throw err;
    }
  }

  async clearAllStepHistory() {
    await this.ensureStepHistoryTable();

    let historyDeleted = 0;
    try {
      const clearHistoryResult = await this.run('DELETE FROM step_history');
      historyDeleted = clearHistoryResult.changes;
      console.log(`[StepHistory] Cleared ${historyDeleted} entries from step_history`);
    } catch (err) {
      console.error('[StepHistory] Error clearing step_history table:', err);
    }

    try {
      const updateResult = await this.run(
        "UPDATE ferramental SET steps = NULL, analysis_completed = 0, last_update = datetime('now')"
      );
      console.log(`[StepHistory] Cleared steps and analysis_completed from ${updateResult.changes} tooling records`);
      this.emitChange('step-history:cleared-all', {
        historyDeleted,
        toolingsUpdated: updateResult.changes
      });
      return {
        success: true,
        historyDeleted,
        toolingsUpdated: updateResult.changes
      };
    } catch (err) {
      console.error('[StepHistory] Error clearing steps:', err);
      throw err;
    }
  }

  async getStepChangeAverage() {
    const query = `
      SELECT
        AVG(diff_days) AS avg_days,
        COUNT(*) AS intervals_count
      FROM (
        SELECT
          tooling_id,
          (julianday(changed_at) - julianday(prev_changed_at)) AS diff_days
        FROM (
          SELECT
            tooling_id,
            changed_at,
            LAG(changed_at) OVER (
              PARTITION BY tooling_id
              ORDER BY datetime(changed_at) ASC
            ) AS prev_changed_at
          FROM step_history
        ) ranked
        WHERE prev_changed_at IS NOT NULL
      ) intervals
    `;

    try {
      const row = await this.get(query, []);
      return {
        avgDays: row?.avg_days !== null && row?.avg_days !== undefined ? Number(row.avg_days) : null,
        intervalsCount: row?.intervals_count ? Number(row.intervals_count) : 0
      };
    } catch (err) {
      console.error('[StepHistory] Error getting average step change interval:', err);
      throw err;
    }
  }

  async ensureTodosTable() {
    const existingTable = await this.get("SELECT name FROM sqlite_master WHERE type='table' AND name='todos'");
    if (!existingTable) {
      await this.run(
        `CREATE TABLE todos (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          tooling_id INTEGER NOT NULL,
          text TEXT NOT NULL,
          completed INTEGER DEFAULT 0,
          created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          FOREIGN KEY (tooling_id) REFERENCES ferramental(id) ON DELETE CASCADE
        )`
      );
      return;
    }

    const columns = await this.all('PRAGMA table_info(todos)');
    const columnNames = columns.map((column) => column.name);
    const requiredColumns = ['id', 'tooling_id', 'text', 'completed', 'created_at'];
    const missingColumns = requiredColumns.filter((columnName) => !columnNames.includes(columnName));

    if (missingColumns.length === 0) {
      return;
    }

    try {
      await this.run('DROP TABLE IF EXISTS todos_backup');
    } catch (err) {
    }

    await this.run('ALTER TABLE todos RENAME TO todos_backup');
    await this.run(
      `CREATE TABLE todos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tooling_id INTEGER NOT NULL,
        text TEXT NOT NULL,
        completed INTEGER DEFAULT 0,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (tooling_id) REFERENCES ferramental(id) ON DELETE CASCADE
      )`
    );

    try {
      await this.run(
        `INSERT INTO todos (id, tooling_id, text, completed, created_at)
         SELECT id, tooling_id, text, completed, created_at FROM todos_backup`
      );
    } catch (err) {
    }

    try {
      await this.run('DROP TABLE todos_backup');
    } catch (err) {
    }
  }

  async getTodos(toolingId) {
    return this.all('SELECT * FROM todos WHERE tooling_id = ? ORDER BY created_at ASC', [toolingId]);
  }

  async addTodo(toolingId, text) {
    const result = await this.run(
      'INSERT INTO todos (tooling_id, text, completed) VALUES (?, ?, 0)',
      [toolingId, text]
    );

    const todo = { id: result.lastID, tooling_id: toolingId, text, completed: 0 };
    this.emitChange('todo:created', todo);
    return todo;
  }

  async updateTodo(todoId, text, completed) {
    await this.run('UPDATE todos SET text = ?, completed = ? WHERE id = ?', [text, completed, todoId]);
    this.emitChange('todo:updated', { id: Number(todoId), text, completed });
    return { success: true };
  }

  async deleteTodo(todoId) {
    await this.run('DELETE FROM todos WHERE id = ?', [todoId]);
    this.emitChange('todo:deleted', { id: Number(todoId) });
    return { success: true };
  }

  async getSupplierExportRows(supplierName, filteredIds = null) {
    let query;
    let params;

    if (filteredIds && Array.isArray(filteredIds) && filteredIds.length > 0) {
      const placeholders = filteredIds.map(() => '?').join(',');
      query = `
        SELECT
          id,
          pn,
          pn_description,
          tool_description,
          tooling_life_qty,
          produced,
          date_remaining_tooling_life as production_date,
          annual_volume_forecast as forecast,
          date_annual_volume as forecast_date
        FROM ferramental
        WHERE supplier = ? AND id IN (${placeholders})
        ORDER BY id
      `;
      params = [supplierName, ...filteredIds];
    } else {
      query = `
        SELECT
          id,
          pn,
          pn_description,
          tool_description,
          tooling_life_qty,
          produced,
          date_remaining_tooling_life as production_date,
          annual_volume_forecast as forecast,
          date_annual_volume as forecast_date
        FROM ferramental
        WHERE supplier = ?
        ORDER BY id
      `;
      params = [supplierName];
    }

    return this.all(query, params);
  }

  async countSupplierRecordsByName(supplierName) {
    const row = await this.get(
      'SELECT COUNT(*) as count FROM ferramental WHERE LOWER(TRIM(supplier)) = LOWER(TRIM(?))',
      [supplierName]
    );
    return Number(row?.count || 0);
  }

  async createImportedToolingRecord(data) {
    return this.run(
      `INSERT INTO ferramental (
        supplier, pn, pn_description, tool_description, tooling_life_qty, produced,
        date_remaining_tooling_life, annual_volume_forecast,
        date_annual_volume, comments, status
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        data.supplierName,
        data.pn || '',
        data.pnDescription || '',
        data.toolDescription || '',
        data.toolingLifeQty,
        data.producedQty,
        data.productionDateISO,
        data.forecastQty,
        data.forecastDateISO,
        data.comments,
        data.status || 'ACTIVE'
      ]
    );
  }

  async getImportedToolingRecord(id, supplierName) {
    return this.get(
      `SELECT pn,
              pn_description,
              tool_description,
              tooling_life_qty,
              produced,
              date_remaining_tooling_life,
              annual_volume_forecast,
              date_annual_volume,
              comments
         FROM ferramental
        WHERE id = ? AND supplier = ?`,
      [id, supplierName]
    );
  }

  async updateImportedToolingRecord(data) {
    return this.run(
      `UPDATE ferramental
         SET pn = ?,
             pn_description = ?,
             tool_description = ?,
             tooling_life_qty = ?,
             produced = ?,
             date_remaining_tooling_life = ?,
             annual_volume_forecast = ?,
             date_annual_volume = ?,
             comments = ?
       WHERE id = ? AND supplier = ?`,
      [
        data.pn || '',
        data.pnDescription || '',
        data.toolDescription || '',
        data.toolingLifeQty,
        data.producedQty,
        data.productionDateISO,
        data.forecastQty,
        data.forecastDateISO,
        data.comments,
        data.id,
        data.supplierName
      ]
    );
  }

  async getForecastSupplierRows() {
    return this.all(
      `SELECT
         pn,
         supplier,
         annual_volume_forecast as forecast,
         date_annual_volume as forecast_date
       FROM ferramental
       WHERE pn IS NOT NULL AND pn != ''
       GROUP BY pn
       ORDER BY pn`,
      []
    );
  }

  async getToolingIdsByPn(pn) {
    const rows = await this.all('SELECT id FROM ferramental WHERE pn = ?', [pn]);
    return rows.map((row) => row.id);
  }

  async getAllToolingRows() {
    return this.all('SELECT * FROM ferramental ORDER BY id', []);
  }

  async loadTooling() {
    return this.all('SELECT * FROM ferramental ORDER BY id DESC', []);
  }

  async getSuppliersWithStats() {
    const rows = await this.all(
      'SELECT * FROM ferramental WHERE supplier IS NOT NULL AND supplier != \'\'',
      []
    );

    const supplierMap = new Map();
    rows.forEach((item) => {
      const supplierName = String(item.supplier || '').trim();
      if (!supplierName) {
        return;
      }

      if (!supplierMap.has(supplierName)) {
        supplierMap.set(supplierName, {
          supplier: supplierName,
          items: []
        });
      }

      supplierMap.get(supplierName).items.push(item);
    });

    return Array.from(supplierMap.values()).sort((a, b) =>
      a.supplier.localeCompare(b.supplier, 'pt-BR', { sensitivity: 'base' })
    );
  }

  async renameSupplier(currentName, newName) {
    const normalizedCurrent = String(currentName || '').trim();
    const normalizedNew = String(newName || '').trim();

    if (!normalizedCurrent || !normalizedNew) {
      return { success: false, message: 'Invalid supplier name.' };
    }

    if (normalizedCurrent === normalizedNew) {
      return { success: true, supplierName: normalizedNew, updated: 0 };
    }

    try {
      await this.ensureSupplierMetadataTable();
    } catch (err) {
    }

    const existingSupplier = await this.get(
      `SELECT COUNT(1) as count FROM ferramental
       WHERE LOWER(TRIM(supplier)) = LOWER(TRIM(?))
         AND LOWER(TRIM(supplier)) != LOWER(TRIM(?))`,
      [normalizedNew, normalizedCurrent]
    );

    if ((existingSupplier?.count || 0) > 0) {
      return { success: false, message: 'A supplier with this name already exists.' };
    }

    const updateResult = await this.run(
      'UPDATE ferramental SET supplier = ? WHERE TRIM(supplier) = ?',
      [normalizedNew, normalizedCurrent]
    );

    try {
      await this.run(
        `UPDATE ${this.supplierMetadataTable} SET supplier = ? WHERE supplier = ?`,
        [normalizedNew, normalizedCurrent]
      );
    } catch (err) {
    }

    try {
      if (typeof this.helpers.renameAttachmentsFolder === 'function') {
        this.helpers.renameAttachmentsFolder(normalizedCurrent, normalizedNew);
      }
    } catch (renameErr) {
      return {
        success: false,
        message: 'Supplier renamed, but there was an error moving the attachments.'
      };
    }

    if (updateResult.changes > 0) {
      this.emitChange('supplier:renamed', {
        supplierName: normalizedNew,
        previousSupplierName: normalizedCurrent,
        updated: updateResult.changes
      });
    }

    return {
      success: true,
      supplierName: normalizedNew,
      updated: updateResult.changes
    };
  }

  async getToolingBySupplier(supplier) {
    return this.all(
      `SELECT * FROM ferramental
       WHERE supplier = ?
       ORDER BY
         CASE
           WHEN julianday('now') > julianday(expiration_date) THEN 1
           WHEN julianday(expiration_date) - julianday('now') < 365 THEN 2
           ELSE 3
         END,
         expiration_date`,
      [supplier]
    );
  }

  async searchTooling(term) {
    const searchTerm = `%${term}%`;
    return this.all(
      `SELECT * FROM ferramental
       WHERE pn LIKE ?
          OR pn_description LIKE ?
          OR supplier LIKE ?
          OR tool_description LIKE ?
          OR tool_number_arb LIKE ?
          OR asset_number LIKE ?
          OR tool_ownership LIKE ?
          OR customer LIKE ?
          OR tooling_life_qty LIKE ?
          OR produced LIKE ?
          OR annual_volume_forecast LIKE ?
          OR status LIKE ?
          OR steps LIKE ?
          OR bu LIKE ?
          OR category LIKE ?
          OR cummins_responsible LIKE ?
          OR comments LIKE ?
          OR stim_tooling_management LIKE ?
          OR vpcr LIKE ?
          OR disposition LIKE ?
          OR CAST(id AS TEXT) LIKE ?
       ORDER BY id DESC`,
      [
        searchTerm, searchTerm, searchTerm, searchTerm, searchTerm,
        searchTerm, searchTerm, searchTerm, searchTerm, searchTerm,
        searchTerm, searchTerm, searchTerm, searchTerm, searchTerm,
        searchTerm, searchTerm, searchTerm, searchTerm, searchTerm,
        searchTerm
      ]
    );
  }

  async getToolingById(id) {
    return this.get('SELECT * FROM ferramental WHERE id = ?', [id]);
  }

  async getToolingByReplacementId(replacementId) {
    const numericReplacementId = Number(replacementId);
    if (!Number.isFinite(numericReplacementId)) {
      return [];
    }

    return this.all(
      'SELECT * FROM ferramental WHERE replacement_tooling_id = ? ORDER BY id ASC',
      [numericReplacementId]
    );
  }

  async getAllToolingIds() {
    return this.all(
      `SELECT id, pn, pn_description, supplier, tool_description, status, replacement_tooling_id
       FROM ferramental
       WHERE UPPER(TRIM(COALESCE(status, ''))) != 'OBSOLETE'
         AND (replacement_tooling_id IS NULL OR TRIM(CAST(replacement_tooling_id AS TEXT)) = '')
       ORDER BY id ASC`,
      []
    );
  }

  async getIdsWithIncomingLinks(targetIds) {
    if (!Array.isArray(targetIds) || targetIds.length === 0) {
      return [];
    }

    const placeholders = targetIds.map(() => '?').join(',');
    const rows = await this.all(
      `SELECT DISTINCT replacement_tooling_id FROM ferramental
       WHERE replacement_tooling_id IN (${placeholders})
         AND replacement_tooling_id IS NOT NULL
         AND replacement_tooling_id != ''`,
      targetIds
    );

    return rows.map((row) => String(row.replacement_tooling_id));
  }

  async getUniqueResponsibles() {
    const rows = await this.all(
      `SELECT DISTINCT cummins_responsible
       FROM ferramental
       WHERE cummins_responsible IS NOT NULL
         AND TRIM(cummins_responsible) != ''
       ORDER BY cummins_responsible ASC`,
      []
    );

    return rows
      .map((row) => String(row.cummins_responsible || '').trim())
      .filter((name) => name.length > 0);
  }

  async getAnalytics() {
    const rows = await this.all(
      `WITH filtered AS (
         SELECT *
         FROM ferramental
         WHERE supplier IS NOT NULL
           AND TRIM(supplier) != ''
       )
       SELECT
         COUNT(*) as total,
         COUNT(DISTINCT supplier) as suppliers,
         COUNT(DISTINCT cummins_responsible) as responsibles,
         SUM(CASE
           WHEN UPPER(status) = 'ACTIVE' THEN 1
           ELSE 0
         END) as active,
         SUM(CASE
           WHEN ((percent_tooling_life IS NOT NULL AND CAST(percent_tooling_life AS REAL) >= 100.0)
             OR (tooling_life_qty > 0 AND produced >= tooling_life_qty))
             AND NOT (UPPER(status) = 'OBSOLETE' AND replacement_tooling_id IS NOT NULL AND replacement_tooling_id != '')
           THEN 1
           ELSE 0
         END) as expired_total,
         SUM(CASE
           WHEN NOT ((percent_tooling_life IS NOT NULL AND CAST(percent_tooling_life AS REAL) >= 100.0)
                 OR (tooling_life_qty > 0 AND produced >= tooling_life_qty))
             AND expiration_date IS NOT NULL
             AND julianday(expiration_date) - julianday('now') > 0
             AND julianday(expiration_date) - julianday('now') <= 730
             AND NOT (UPPER(status) = 'OBSOLETE' AND replacement_tooling_id IS NOT NULL AND replacement_tooling_id != '')
           THEN 1
           ELSE 0
         END) as expiring_two_years
       FROM filtered`,
      []
    );

    return rows[0];
  }

  async getStepsSummary() {
    const rows = await this.all(
      `SELECT
         steps,
         COUNT(*) as count
       FROM ferramental
       WHERE steps IS NOT NULL
         AND TRIM(steps) != ''
         AND TRIM(steps) != '0'
       GROUP BY steps
       ORDER BY CAST(steps AS INTEGER)`,
      []
    );

    const supplierRows = await this.all(
      `SELECT *
       FROM ferramental
       WHERE steps IS NOT NULL
         AND TRIM(steps) != ''
         AND TRIM(steps) != '0'
         AND supplier IS NOT NULL
         AND TRIM(supplier) != ''
       ORDER BY CAST(steps AS INTEGER), supplier COLLATE NOCASE`,
      []
    );

    const suppliersMap = {};
    (supplierRows || []).forEach((item) => {
      if (!suppliersMap[item.steps]) {
        suppliersMap[item.steps] = {};
      }

      const supplierName = String(item.supplier || '').trim();
      if (!suppliersMap[item.steps][supplierName]) {
        suppliersMap[item.steps][supplierName] = {
          supplier: supplierName,
          items: []
        };
      }

      suppliersMap[item.steps][supplierName].items.push(item);
    });

    return (rows || []).map((row) => ({
      steps: row.steps,
      count: row.count,
      suppliers: Object.values(suppliersMap[row.steps] || {})
    }));
  }

  async getStepSuppliersMetrics(step) {
    return this.all(
      `SELECT supplier,
              COUNT(*) as total,
              SUM(CASE
                     WHEN ((percent_tooling_life IS NOT NULL AND CAST(percent_tooling_life AS REAL) >= 100.0)
                           OR (tooling_life_qty > 0 AND produced >= tooling_life_qty))
                          AND NOT (UPPER(status) = 'OBSOLETE' AND replacement_tooling_id IS NOT NULL AND replacement_tooling_id != '')
                     THEN 1 ELSE 0 END) as expired,
              SUM(CASE
                     WHEN NOT ((percent_tooling_life IS NOT NULL AND CAST(percent_tooling_life AS REAL) >= 100.0)
                           OR (tooling_life_qty > 0 AND produced >= tooling_life_qty))
                          AND expiration_date IS NOT NULL
                          AND julianday(expiration_date) - julianday('now') > 0
                          AND julianday(expiration_date) - julianday('now') <= 730
                          AND NOT (UPPER(status) = 'OBSOLETE' AND replacement_tooling_id IS NOT NULL AND replacement_tooling_id != '')
                     THEN 1 ELSE 0 END) as expiring
       FROM ferramental
       WHERE steps = ?
         AND supplier IS NOT NULL AND TRIM(supplier) != ''
       GROUP BY supplier`,
      [step]
    );
  }

  async updateTooling(id, data, options = {}) {
    try {
      await this.ensureReplacementColumnExists();
    } catch (error) {
    }

    const result = await this.executeToolingUpdate(id, data, 1, {
      skipWhenNoTrackedChanges: false,
      ...options
    });

    if (result?.changes > 0) {
      this.emitChange('tooling:updated', {
        id: Number(id),
        changes: result.changes,
        fields: Object.keys(data || {})
      });
    }

    return result;
  }

  async createTooling(data) {
    const pn = (data?.pn || '').trim();
    const pnDescription = (data?.pn_description || '').trim();
    const supplier = (data?.supplier || '').trim();
    const toolingLife = parseFloat(data?.tooling_life_qty) || 0;
    const produced = parseFloat(data?.produced) || 0;
    const toolDescription = (data?.tool_description || '').trim();
    const productionDateValue = (data?.date_remaining_tooling_life || '').trim();
    const productionDate = productionDateValue.length > 0 ? productionDateValue : null;
    const forecastRaw = data?.annual_volume_forecast;
    const parsedForecast =
      forecastRaw === null || forecastRaw === undefined
        ? null
        : (typeof forecastRaw === 'number' ? forecastRaw : parseFloat(forecastRaw));
    const annualForecast = Number.isFinite(parsedForecast) && parsedForecast > 0 ? parsedForecast : null;
    const forecastDateValue = (data?.date_annual_volume || '').trim();
    const forecastDate = forecastDateValue.length > 0 ? forecastDateValue : null;
    const comments = data?.comments || null;

    if (!pn || !supplier) {
      return { success: false, error: 'PN and supplier are required.' };
    }

    const remaining = toolingLife - produced;
    const percent = toolingLife > 0 ? ((produced / toolingLife) * 100).toFixed(1) : 0;

    const insertResult = await this.run(
      `INSERT INTO ferramental (
         pn,
         pn_description,
         supplier,
         tool_description,
         tooling_life_qty,
         produced,
         remaining_tooling_life_pcs,
         percent_tooling_life,
         annual_volume_forecast,
         date_annual_volume,
         status,
         last_update,
         date_remaining_tooling_life,
         comments
       ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'), ?, ?)`,
      [
        pn,
        pnDescription,
        supplier,
        toolDescription,
        toolingLife,
        produced,
        remaining >= 0 ? remaining : 0,
        percent,
        annualForecast,
        forecastDate,
        'ACTIVE',
        productionDate,
        comments
      ]
    );

    if (supplier) {
      try {
        await this.incrementDataRevision(supplier);
      } catch (revErr) {
        console.error('[DataRevision] Failed to increment on create:', revErr);
      }
    }

    this.emitChange('tooling:created', {
      id: insertResult.lastID,
      supplier
    });

    return { success: true, id: insertResult.lastID };
  }

  async deleteTooling(id) {
    const record = await this.get('SELECT supplier FROM ferramental WHERE id = ?', [id]);
    const supplier = record?.supplier ? String(record.supplier).trim() : '';

    const deleteResult = await this.run('DELETE FROM ferramental WHERE id = ?', [id]);

    if (supplier && deleteResult.changes > 0) {
      try {
        await this.incrementDataRevision(supplier);
      } catch (revErr) {
        console.error('[DataRevision] Failed to increment on delete:', revErr);
      }

      if (typeof this.helpers.cleanupItemAttachments === 'function') {
        try {
          this.helpers.cleanupItemAttachments(supplier, id);
        } catch (cleanupErr) {
          console.error('[Cleanup] Error cleaning up attachments on delete:', cleanupErr);
        }
      }
    }

    if (deleteResult.changes > 0) {
      this.emitChange('tooling:deleted', {
        id: Number(id),
        supplier
      });
    }

    return { success: true, changes: deleteResult.changes };
  }

  async executeToolingUpdate(id, payload, attempt = 1, options = {}) {
    const data = { ...(payload || {}) };
    const skipWhenNoTrackedChanges = options?.skipWhenNoTrackedChanges === true;
    const shouldTrackChangeField = this.helpers.shouldTrackChangeField || (() => false);
    const collectChangeEntries = this.helpers.collectChangeEntries || (() => []);
    const mergeChangeEntriesIntoComments = this.helpers.mergeChangeEntriesIntoComments || ((existingComments) => existingComments);
    const getCurrentTimestampBR = this.helpers.getCurrentTimestampBR || (() => '');

    if (Object.keys(data).length === 0) {
      return { success: true, changes: 0, comments: null, skipped: true };
    }

    if (
      Object.prototype.hasOwnProperty.call(data, 'tooling_life_qty') ||
      Object.prototype.hasOwnProperty.call(data, 'produced')
    ) {
      const toolingLife = parseFloat(data.tooling_life_qty) || 0;
      const produced = parseFloat(data.produced) || 0;
      const remaining = toolingLife - produced;
      const percentUsed = toolingLife > 0 ? ((produced / toolingLife) * 100).toFixed(1) : 0;

      data.remaining_tooling_life_pcs = remaining >= 0 ? remaining : 0;
      data.percent_tooling_life = percentUsed;
    }

    const trackableFields = Object.keys(data).filter((field) => shouldTrackChangeField(field));
    const requiresLookup = trackableFields.length > 0;
    let existingRecord = null;

    if (requiresLookup) {
      const selectFields = new Set(['id', 'comments', 'supplier']);
      trackableFields.forEach((field) => selectFields.add(field));
      const columnList = Array.from(selectFields).join(', ');
      existingRecord = await this.get(`SELECT ${columnList} FROM ferramental WHERE id = ?`, [id]);
    }

    let changeEntries = [];
    if (existingRecord) {
      changeEntries = collectChangeEntries(existingRecord, data);
      if (changeEntries.length > 0) {
        console.debug('[Ferramental][ChangeDebug]', {
          id,
          updates: changeEntries.map((entry) => ({
            field: entry.field,
            before: entry.oldFormatted,
            after: entry.newFormatted
          }))
        });

        const customTimestamp = options?.commentTimestamp;
        data.comments = mergeChangeEntriesIntoComments(
          existingRecord.comments,
          data.comments,
          changeEntries,
          customTimestamp || getCurrentTimestampBR()
        );
      }
    }

    const validFields = Object.keys(data).filter((key) => {
      if (!key || typeof key !== 'string' || key.trim() === '') {
        return false;
      }

      const fieldName = key.trim();
      if (fieldName.startsWith('_') || fieldName.startsWith('$')) {
        return false;
      }

      return data[key] !== undefined;
    });

    if (validFields.length === 0) {
      return {
        success: true,
        changes: 0,
        comments: existingRecord?.comments ?? null,
        skipped: true
      };
    }

    const hasNonTrackableUpdates = validFields.some((field) => !shouldTrackChangeField(field));
    if (
      skipWhenNoTrackedChanges &&
      existingRecord &&
      changeEntries.length === 0 &&
      !hasNonTrackableUpdates
    ) {
      return {
        success: true,
        changes: 0,
        comments: existingRecord?.comments ?? null,
        skipped: true
      };
    }

    const validValues = validFields.map((key) => data[key]);
    const setClause = validFields.map((field) => `${field.trim()} = ?`).join(', ');
    const hasUserEditableFields = Object.keys(payload || {}).some(
      (field) => !this.silentUpdateFields.has(field)
    );
    const query = hasUserEditableFields
      ? `UPDATE ferramental SET ${setClause}, last_update = datetime('now') WHERE id = ?`
      : `UPDATE ferramental SET ${setClause} WHERE id = ?`;

    try {
      const updateResult = await this.run(query, [...validValues, id]);

      if (Object.prototype.hasOwnProperty.call(data, 'steps') && existingRecord) {
        const oldStep = existingRecord.steps || null;
        const newStep = data.steps || null;

        if (oldStep !== newStep) {
          try {
            await this.recordStepChange(id, oldStep, newStep);
          } catch (stepErr) {
            console.error('[StepHistory] Failed to record step change:', stepErr);
          }
        }
      }

      if (updateResult.changes > 0 && hasUserEditableFields) {
        const supplierForRevision = existingRecord?.supplier || data.supplier;
        if (supplierForRevision) {
          this.scheduleRevisionIncrement(String(supplierForRevision).trim());
        }
      }

      return {
        success: true,
        changes: updateResult.changes,
        comments: data.comments ?? existingRecord?.comments ?? null,
        skipped: false,
        lastUpdateModified: hasUserEditableFields
      };
    } catch (err) {
      if (this.isMissingReplacementColumnError(err) && attempt === 1) {
        await this.ensureReplacementColumnExists(true);
        return this.executeToolingUpdate(id, payload, attempt + 1, options);
      }

      throw err;
    }
  }

  async close() {
    Object.values(this.revisionTimers).forEach((timer) => clearTimeout(timer));
    this.revisionTimers = Object.create(null);

    const db = this.dbConnection || (typeof this.getDb === 'function' ? this.getDb() : null);
    if (!db) {
      return;
    }

    await new Promise((resolve, reject) => {
      db.close((err) => {
        if (err) {
          reject(err);
          return;
        }
        resolve();
      });
    });

    this.dbConnection = null;
  }
}

module.exports = {
  ToolingDatabase
};