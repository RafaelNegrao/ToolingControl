const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const sqlite3 = require('sqlite3').verbose();
const fs = require('fs');

let mainWindow;
let db;
const DATABASE_FILE_NAME = 'ferramental_database.db';
const REPLACEMENT_COLUMN_NAME = 'replacement_tooling_id';
const REPLACEMENT_COLUMN_TYPE = 'INTEGER';
let replacementColumnEnsured = false;
let replacementColumnEnsuringPromise = null;

// Obter diretório base do executável
function getAppBaseDir() {
  // Em produção (empacotado), usa o diretório real do executável gerado
  // Em desenvolvimento, usa o diretório do código fonte
  if (process.env.PORTABLE_EXECUTABLE_DIR) {
    return process.env.PORTABLE_EXECUTABLE_DIR;
  }

  if (app && app.isPackaged) {
    return path.dirname(process.execPath);
  }

  if (process.resourcesPath && !process.resourcesPath.endsWith('app.asar')) {
    return path.join(process.resourcesPath, '..');
  }

  return __dirname;
}

function ensureDatabaseFile(baseDir) {
  const targetPath = path.join(baseDir, DATABASE_FILE_NAME);

  if (fs.existsSync(targetPath)) {
    return targetPath;
  }

  try {
    if (!fs.existsSync(baseDir)) {
      fs.mkdirSync(baseDir, { recursive: true });
    }

    const packagedDbPath = path.join(app.getAppPath(), DATABASE_FILE_NAME);
    if (fs.existsSync(packagedDbPath)) {
      fs.copyFileSync(packagedDbPath, targetPath);
      console.log('Banco de dados copiado para o diretório do executável.');
      return targetPath;
    }

    console.warn('Arquivo do banco não encontrado no pacote. Um novo arquivo vazio será criado.');
    fs.writeFileSync(targetPath, '');
    return targetPath;
  } catch (error) {
    console.error('Falha ao preparar o banco de dados:', error);
    return targetPath;
  }
}

// Conectar ao banco de dados
function connectDatabase() {
  const baseDir = getAppBaseDir();
  console.log('Base directory detectado:', baseDir);
  const dbPath = ensureDatabaseFile(baseDir);
  console.log('Caminho do banco de dados:', dbPath);
  
  return new Promise((resolve, reject) => {
    db = new sqlite3.Database(dbPath, (err) => {
      if (err) {
        console.error('Erro ao conectar ao banco de dados:', err);
        reject(err);
        return;
      }

      console.log('Conectado ao banco de dados ferramental_database.db');
      checkAndAddColumns()
        .then(() => ensureTodosTable())
        .then(resolve)
        .catch((schemaError) => {
          console.error('Erro ao ajustar colunas do banco:', schemaError);
          reject(schemaError);
        });
    });
  });
}

// Verificar e adicionar colunas que possam estar faltando
function checkAndAddColumns() {
  const columnsToCheck = [
    { name: 'tool_number_arb', type: 'TEXT' },
    { name: 'tooling_life_qty', type: 'TEXT' },
    { name: 'date_remaining_tooling_life', type: 'TEXT' },
    { name: 'annual_volume_forecast', type: 'TEXT' },
    { name: 'date_annual_volume', type: 'TEXT' },
    { name: 'replacement_tooling_id', type: 'INTEGER' },
    { name: 'bailment_agreement_signed', type: 'TEXT' },
    { name: 'tooling_book', type: 'TEXT' },
    { name: 'disposition', type: 'TEXT' },
    { name: 'finish_due_date', type: 'TEXT' }
  ];
  
  return new Promise((resolve, reject) => {
    db.all('PRAGMA table_info(ferramental)', [], (err, columns) => {
      if (err) {
        reject(err);
        return;
      }

      const existingColumns = new Set(columns.map(col => col.name));
      if (existingColumns.has(REPLACEMENT_COLUMN_NAME)) {
        replacementColumnEnsured = true;
      }

      const missingColumns = columnsToCheck.filter(({ name }) => !existingColumns.has(name));
      if (missingColumns.length === 0) {
        resolve();
        return;
      }

      let completed = 0;
      let hasFailed = false;

      db.serialize(() => {
        missingColumns.forEach(({ name, type }) => {
          const columnType = type || 'TEXT';
          console.log(`Adicionando coluna: ${name}`);
          db.run(`ALTER TABLE ferramental ADD COLUMN ${name} ${columnType}`, (alterErr) => {
            if (hasFailed) {
              return;
            }

            if (alterErr && !/duplicate column/i.test(alterErr.message || '')) {
              hasFailed = true;
              console.error(`Erro ao adicionar coluna ${name}:`, alterErr);
              reject(alterErr);
              return;
            }

            if (name === REPLACEMENT_COLUMN_NAME) {
              replacementColumnEnsured = true;
            }

            completed += 1;
            if (completed === missingColumns.length) {
              resolve();
            }
          });
        });
      });
    });
  });
}

function ensureReplacementColumnExists(force = false) {
  if (force) {
    replacementColumnEnsured = false;
    replacementColumnEnsuringPromise = null;
  }

  if (replacementColumnEnsured) {
    return Promise.resolve();
  }

  if (replacementColumnEnsuringPromise) {
    return replacementColumnEnsuringPromise;
  }

  replacementColumnEnsuringPromise = new Promise((resolve, reject) => {
    db.all('PRAGMA table_info(ferramental)', [], (err, columns) => {
      if (err) {
        replacementColumnEnsuringPromise = null;
        reject(err);
        return;
      }

      const hasColumn = Array.isArray(columns) && columns.some(col => col.name === REPLACEMENT_COLUMN_NAME);
      if (hasColumn) {
        replacementColumnEnsured = true;
        replacementColumnEnsuringPromise = null;
        resolve();
        return;
      }

      db.run(`ALTER TABLE ferramental ADD COLUMN ${REPLACEMENT_COLUMN_NAME} ${REPLACEMENT_COLUMN_TYPE}`, (alterErr) => {
        replacementColumnEnsuringPromise = null;
        if (alterErr && !/duplicate column/i.test(alterErr.message || '')) {
          reject(alterErr);
          return;
        }
        replacementColumnEnsured = true;
        resolve();
      });
    });
  });

  return replacementColumnEnsuringPromise;
}

function isMissingReplacementColumnError(err) {
  if (!err || typeof err.message !== 'string') {
    return false;
  }
  return err.message.includes(`no such column: ${REPLACEMENT_COLUMN_NAME}`);
}

function ensureTodosTable() {
  return new Promise((resolve, reject) => {
    // First check if table exists
    db.get("SELECT name FROM sqlite_master WHERE type='table' AND name='todos'", (err, row) => {
      if (err) {
        console.error('Erro ao verificar tabela todos:', err);
        reject(err);
        return;
      }

      if (!row) {
        // Table doesn't exist, create it
        console.log('Tabela todos não existe, criando...');
        db.run(`
          CREATE TABLE todos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tooling_id INTEGER NOT NULL,
            text TEXT NOT NULL,
            completed INTEGER DEFAULT 0,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (tooling_id) REFERENCES ferramental(id) ON DELETE CASCADE
          )
        `, (createErr) => {
          if (createErr) {
            console.error('Erro ao criar tabela todos:', createErr);
            reject(createErr);
          } else {
            console.log('Tabela todos criada com sucesso');
            resolve();
          }
        });
      } else {
        // Table exists, verify columns
        db.all("PRAGMA table_info(todos)", (pragmaErr, columns) => {
          if (pragmaErr) {
            console.error('Erro ao verificar colunas da tabela todos:', pragmaErr);
            reject(pragmaErr);
            return;
          }

          const columnNames = columns.map(col => col.name);
          const requiredColumns = ['id', 'tooling_id', 'text', 'completed', 'created_at'];
          const missingColumns = requiredColumns.filter(col => !columnNames.includes(col));

          if (missingColumns.length > 0) {
            console.log('Colunas faltando na tabela todos:', missingColumns);
            // Recreate table with all columns
            db.serialize(() => {
              db.run('DROP TABLE IF EXISTS todos_backup', (dropErr) => {
                if (dropErr) console.error('Erro ao limpar backup:', dropErr);
              });
              
              db.run('ALTER TABLE todos RENAME TO todos_backup', (renameErr) => {
                if (renameErr) {
                  console.error('Erro ao renomear tabela todos:', renameErr);
                  reject(renameErr);
                  return;
                }

                db.run(`
                  CREATE TABLE todos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tooling_id INTEGER NOT NULL,
                    text TEXT NOT NULL,
                    completed INTEGER DEFAULT 0,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (tooling_id) REFERENCES ferramental(id) ON DELETE CASCADE
                  )
                `, (createErr) => {
                  if (createErr) {
                    console.error('Erro ao recriar tabela todos:', createErr);
                    reject(createErr);
                    return;
                  }

                  // Try to copy data if possible
                  db.run(`
                    INSERT INTO todos (id, tooling_id, text, completed, created_at)
                    SELECT id, tooling_id, text, completed, created_at FROM todos_backup
                  `, (copyErr) => {
                    if (copyErr) {
                      console.warn('Não foi possível copiar dados antigos:', copyErr);
                    }

                    db.run('DROP TABLE todos_backup', (dropErr) => {
                      if (dropErr) console.error('Erro ao remover backup:', dropErr);
                    });

                    console.log('Tabela todos recriada com sucesso');
                    resolve();
                  });
                });
              });
            });
          } else {
            console.log('Tabela todos verificada - todas as colunas presentes');
            resolve();
          }
        });
      }
    });
  });
}

function executeToolingUpdate(id, payload, attempt = 1) {
  console.log(`[DB Update] ID: ${id}, Payload:`, payload, `Attempt: ${attempt}`);
  return new Promise((resolve, reject) => {
    const data = { ...payload };
    
    // Only recalculate lifecycle if tooling_life_qty or produced are in the payload
    if (data.hasOwnProperty('tooling_life_qty') || data.hasOwnProperty('produced')) {
      const toolingLife = parseFloat(data.tooling_life_qty) || 0;
      const produced = parseFloat(data.produced) || 0;
      const remaining = toolingLife - produced;
      const percentUsed = toolingLife > 0 ? ((produced / toolingLife) * 100).toFixed(1) : 0;

      data.remaining_tooling_life_pcs = remaining >= 0 ? remaining : 0;
      data.percent_tooling_life = percentUsed;
    }

    const allFields = Object.keys(data);
    const allValues = Object.values(data);
    const setClause = allFields.map(field => `${field} = ?`).join(', ');

    const query = `UPDATE ferramental SET ${setClause}, last_update = datetime('now') WHERE id = ?`;
    console.log(`[DB Update] SQL:`, query);
    console.log(`[DB Update] Values:`, [...allValues, id]);

    db.run(
      query,
      [...allValues, id],
      async function(err) {
        if (err && isMissingReplacementColumnError(err) && attempt === 1) {
          console.log('[DB Update] Missing column, retrying...');
          try {
            await ensureReplacementColumnExists(true);
            const retryResult = await executeToolingUpdate(id, payload, attempt + 1);
            resolve(retryResult);
            return;
          } catch (retryError) {
            console.error('[DB Update] Retry failed:', retryError);
            reject(retryError);
            return;
          }
        }

        if (err) {
          console.error('[DB Update] Error:', err);
          reject(err);
        } else {
          console.log(`[DB Update] Success! Changes: ${this.changes}`);
          resolve({ success: true, changes: this.changes });
        }
      }
    );
  });
}

// Handlers IPC para operações do banco
ipcMain.handle('load-tooling', async () => {
  return new Promise((resolve, reject) => {
    db.all('SELECT * FROM ferramental ORDER BY id DESC', [], (err, rows) => {
      if (err) {
        reject(err);
      } else {
        resolve(rows);
      }
    });
  });
});

ipcMain.handle('get-suppliers-with-stats', async () => {
  return new Promise((resolve, reject) => {
    db.all(`
      SELECT 
        supplier,
        COUNT(*) as total,
        SUM(CASE 
          WHEN (percent_tooling_life IS NOT NULL AND CAST(percent_tooling_life AS REAL) >= 100.0)
            OR (tooling_life_qty > 0 AND produced >= tooling_life_qty)
          THEN 1 
          ELSE 0 
        END) as expired,
        SUM(CASE 
          WHEN (percent_tooling_life IS NULL OR CAST(percent_tooling_life AS REAL) < 100.0)
            AND (tooling_life_qty = 0 OR produced < tooling_life_qty)
            AND expiration_date IS NOT NULL 
            AND julianday(expiration_date) - julianday('now') > 0
            AND julianday(expiration_date) - julianday('now') <= 365
          THEN 1 
          ELSE 0 
        END) as warning_1year,
        SUM(CASE 
          WHEN (percent_tooling_life IS NULL OR CAST(percent_tooling_life AS REAL) < 100.0)
            AND (tooling_life_qty = 0 OR produced < tooling_life_qty)
            AND expiration_date IS NOT NULL 
            AND julianday(expiration_date) - julianday('now') > 365
            AND julianday(expiration_date) - julianday('now') <= 730
          THEN 1 
          ELSE 0 
        END) as warning_2years,
        SUM(CASE 
          WHEN expiration_date IS NOT NULL AND julianday(expiration_date) - julianday('now') BETWEEN 730 AND 1825 THEN 1 
          ELSE 0 
        END) as ok_5years,
        SUM(CASE 
          WHEN expiration_date IS NOT NULL AND julianday(expiration_date) - julianday('now') > 1825 THEN 1 
          ELSE 0 
        END) as ok_plus
      FROM ferramental
      WHERE supplier IS NOT NULL AND supplier != ''
      GROUP BY supplier
      ORDER BY supplier
    `, [], (err, rows) => {
      if (err) {
        reject(err);
      } else {
        resolve(rows);
      }
    });
  });
});

ipcMain.handle('get-tooling-by-supplier', async (event, supplier) => {
  return new Promise((resolve, reject) => {
    db.all(`
      SELECT * FROM ferramental 
      WHERE supplier = ?
      ORDER BY 
        CASE 
          WHEN julianday('now') > julianday(expiration_date) THEN 1
          WHEN julianday(expiration_date) - julianday('now') < 365 THEN 2
          ELSE 3
        END,
        expiration_date
    `, [supplier], (err, rows) => {
      if (err) {
        reject(err);
      } else {
        resolve(rows);
      }
    });
  });
});

ipcMain.handle('search-tooling', async (event, term) => {
  return new Promise((resolve, reject) => {
    const searchTerm = `%${term}%`;
    db.all(`
      SELECT * FROM ferramental 
      WHERE pn LIKE ? 
         OR pn_description LIKE ? 
         OR supplier LIKE ?
         OR tool_description LIKE ?
      ORDER BY id DESC
    `, [searchTerm, searchTerm, searchTerm, searchTerm], (err, rows) => {
      if (err) {
        reject(err);
      } else {
        resolve(rows);
      }
    });
  });
});

ipcMain.handle('get-tooling-by-id', async (event, id) => {
  return new Promise((resolve, reject) => {
    db.get('SELECT * FROM ferramental WHERE id = ?', [id], (err, row) => {
      if (err) {
        reject(err);
      } else {
        resolve(row);
      }
    });
  });
});

ipcMain.handle('get-tooling-by-replacement-id', async (event, replacementId) => {
  return new Promise((resolve, reject) => {
    const numericReplacementId = Number(replacementId);
    if (!Number.isFinite(numericReplacementId)) {
      resolve([]);
      return;
    }
    db.all(
      'SELECT * FROM ferramental WHERE replacement_tooling_id = ? ORDER BY id ASC',
      [numericReplacementId],
      (err, rows) => {
        if (err) {
          reject(err);
        } else {
          resolve(rows);
        }
      }
    );
  });
});

ipcMain.handle('get-all-tooling-ids', async () => {
  return new Promise((resolve, reject) => {
    db.all(
      'SELECT id, pn, pn_description, supplier, tool_description FROM ferramental ORDER BY id ASC',
      [],
      (err, rows) => {
        if (err) {
          reject(err);
        } else {
          resolve(rows);
        }
      }
    );
  });
});

ipcMain.handle('get-analytics', async () => {
  return new Promise((resolve, reject) => {
    db.all(`
      WITH filtered AS (
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
          WHEN (percent_tooling_life IS NOT NULL AND CAST(percent_tooling_life AS REAL) >= 100.0)
            OR (tooling_life_qty > 0 AND produced >= tooling_life_qty)
          THEN 1
          ELSE 0
        END) as expired_total,
        SUM(CASE 
          WHEN (percent_tooling_life IS NULL OR CAST(percent_tooling_life AS REAL) < 100.0)
            AND (tooling_life_qty = 0 OR produced < tooling_life_qty)
            AND expiration_date IS NOT NULL 
            AND julianday(expiration_date) - julianday('now') > 0
            AND julianday(expiration_date) - julianday('now') <= 730
          THEN 1
          ELSE 0
        END) as expiring_two_years
      FROM filtered
    `, [], (err, rows) => {
      if (err) {
        reject(err);
      } else {
        resolve(rows[0]);
      }
    });
  });
});

ipcMain.handle('update-tooling', async (event, id, data) => {
  try {
    await ensureReplacementColumnExists();
  } catch (error) {
    console.error('Erro ao garantir coluna de substituição:', error);
  }
  return executeToolingUpdate(id, data);
});

ipcMain.handle('create-tooling', async (event, data) => {
  return new Promise((resolve, reject) => {
    try {
      const pn = (data?.pn || '').trim();
      const supplier = (data?.supplier || '').trim();
      const toolingLife = parseFloat(data?.tooling_life_qty) || 0;
      const produced = parseFloat(data?.produced) || 0;

      if (!pn || !supplier) {
        resolve({ success: false, error: 'PN e fornecedor são obrigatórios.' });
        return;
      }

      const remaining = toolingLife - produced;
      const percent = toolingLife > 0 ? ((produced / toolingLife) * 100).toFixed(1) : 0;

      db.run(
        `INSERT INTO ferramental (
          pn,
          supplier,
          tooling_life_qty,
          produced,
          remaining_tooling_life_pcs,
          percent_tooling_life,
          status,
          last_update,
          date_remaining_tooling_life
        ) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))`,
        [
          pn,
          supplier,
          toolingLife,
          produced,
          remaining >= 0 ? remaining : 0,
          percent,
          'Ativo'
        ],
        function(err) {
          if (err) {
            reject(err);
          } else {
            resolve({ success: true, id: this.lastID });
          }
        }
      );
    } catch (error) {
      reject(error);
    }
  });
});

ipcMain.handle('delete-tooling', async (event, id) => {
  return new Promise((resolve, reject) => {
    db.run('DELETE FROM ferramental WHERE id = ?', [id], function(err) {
      if (err) {
        reject(err);
      } else {
        resolve({ success: true, changes: this.changes });
      }
    });
  });
});

// ===== GERENCIAMENTO DE ANEXOS =====

// Função para obter diretório de anexos
function getAttachmentsDir() {
  const baseDir = getAppBaseDir();
  const attachmentsPath = path.join(baseDir, 'attachments');
  console.log('[attachments] Base dir:', baseDir);
  console.log('[attachments] Target path:', attachmentsPath);

  // Garante que o diretório de anexos existe
  if (!fs.existsSync(attachmentsPath)) {
    try {
      fs.mkdirSync(attachmentsPath, { recursive: true });
      console.log('[attachments] Diretório criado:', attachmentsPath);
    } catch (error) {
      console.error('[attachments] Falha ao criar diretório:', error);
    }
  }

  return attachmentsPath;
}

// Lista anexos de um fornecedor
ipcMain.handle('get-attachments', async (event, supplierName, itemId = null) => {
  const attachmentsDir = getAttachmentsDir();
  let targetDir;
  
  if (itemId) {
    targetDir = path.join(attachmentsDir, sanitizeFileName(supplierName), String(itemId));
    console.log('[get-attachments] Buscando anexos do CARD:', { supplierName, itemId, targetDir });
  } else {
    targetDir = path.join(attachmentsDir, sanitizeFileName(supplierName));
    console.log('[get-attachments] Buscando anexos GERAIS do supplier:', { supplierName, targetDir });
  }
  
  if (!fs.existsSync(targetDir)) {
    console.log('[get-attachments] Diretório não existe:', targetDir);
    return [];
  }
  
  try {
    const allItems = fs.readdirSync(targetDir);
    console.log('[get-attachments] Todos os itens encontrados:', allItems);
    
    // Filtra apenas arquivos verificando o caminho completo
    const files = allItems.filter(itemName => {
      const fullPath = path.join(targetDir, itemName);
      const stats = fs.statSync(fullPath);
      const isFile = stats.isFile();
      console.log('[get-attachments]', itemName, '-> isFile:', isFile, 'isDirectory:', stats.isDirectory());
      return isFile;
    });
    
    console.log('[get-attachments] Total de ARQUIVOS encontrados:', files.length, ':', files);
    
    return files.map(fileName => {
      const filePath = path.join(targetDir, fileName);
      const stats = fs.statSync(filePath);
      return {
        fileName,
        supplierName,
        itemId,
        fileSize: stats.size,
        uploadDate: stats.birthtime
      };
    });
  } catch (error) {
    console.error('[get-attachments] Erro ao listar anexos:', error);
    return [];
  }
});

// Upload de arquivo
ipcMain.handle('upload-attachment', async (event, supplierName, itemId = null) => {
  console.log('[upload-attachment] Solicitação recebida:', { supplierName, itemId });
  const result = await dialog.showOpenDialog({
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Todos os Arquivos', extensions: ['*'] },
      { name: 'Documentos', extensions: ['pdf', 'doc', 'docx', 'xls', 'xlsx'] },
      { name: 'Imagens', extensions: ['jpg', 'jpeg', 'png', 'gif'] }
    ]
  });
  
  if (result.canceled || result.filePaths.length === 0) {
    console.log('[upload-attachment] Operação cancelada pelo usuário.');
    return { success: false };
  }
  
  try {
    const attachmentsDir = getAttachmentsDir();
    let targetDir;
    
    const supplierDir = path.join(attachmentsDir, sanitizeFileName(supplierName));
    const hasCardScope = itemId !== null && itemId !== undefined;
    targetDir = hasCardScope
      ? path.join(supplierDir, String(itemId))
      : supplierDir;
    
    console.log('[upload-attachment] Diretório alvo calculado:', { supplierDir, targetDir, hasCardScope });
    
    if (!fs.existsSync(targetDir)) {
      try {
        fs.mkdirSync(targetDir, { recursive: true });
        console.log('[upload-attachment] Diretório criado:', targetDir);
      } catch (error) {
        console.error('[upload-attachment] Falha ao criar diretório alvo:', error);
        return { success: false, error: 'Não foi possível criar diretório de anexos.' };
      }
    }
    
    const results = result.filePaths.map(sourcePath => {
      try {
        const fileName = path.basename(sourcePath);
        const destPath = path.join(targetDir, fileName);
        fs.copyFileSync(sourcePath, destPath);
        const existsAfterCopy = fs.existsSync(destPath);
        console.log('[upload-attachment] Arquivo copiado:', { sourcePath, destPath, existsAfterCopy });
        return { success: true, fileName };
      } catch (error) {
        console.error('Erro ao copiar arquivo:', error);
        return { success: false, fileName: path.basename(sourcePath), error: error.message };
      }
    });
    
    const hasFailure = results.some(r => r.success !== true);
    const fileCount = results.filter(r => r.success === true).length;
    
    return { 
      success: !hasFailure, 
      results,
      message: fileCount > 1 ? `${fileCount} arquivos anexados` : 'Arquivo anexado'
    };
  } catch (error) {
    console.error('Erro ao fazer upload:', error);
    return { success: false, error: error.message };
  }
});

ipcMain.handle('upload-attachment-from-paths', async (event, supplierName, filePaths, itemId = null) => {
  if (!supplierName || !Array.isArray(filePaths) || filePaths.length === 0) {
    return { success: false, error: 'Nenhum arquivo informado.' };
  }

  try {
    const attachmentsDir = getAttachmentsDir();
    const supplierDir = path.join(attachmentsDir, sanitizeFileName(supplierName));
    const hasCardScope = itemId !== null && itemId !== undefined;
    const targetDir = hasCardScope
      ? path.join(supplierDir, String(itemId))
      : supplierDir;

    if (!fs.existsSync(targetDir)) {
      try {
        fs.mkdirSync(targetDir, { recursive: true });
        console.log('[upload-attachment-from-paths] Diretório criado:', targetDir);
      } catch (error) {
        console.error('[upload-attachment-from-paths] Falha ao criar diretório alvo:', error);
        return { success: false, error: 'Não foi possível criar diretório de anexos.' };
      }
    }

    const results = filePaths.map(sourcePath => {
      if (!sourcePath || !fs.existsSync(sourcePath)) {
        return { success: false, error: 'Arquivo não encontrado.' };
      }

      const stats = fs.statSync(sourcePath);
      if (!stats.isFile()) {
        return { success: false, error: 'Apenas arquivos podem ser anexados.' };
      }

      const fileName = path.basename(sourcePath);
      const destPath = path.join(targetDir, fileName);

      try {
        fs.copyFileSync(sourcePath, destPath);
        return { success: true, fileName };
      } catch (error) {
        console.error('Erro ao copiar arquivo arrastado:', error);
        return { success: false, fileName, error: error.message };
      }
    });

    const hasFailure = results.some(item => item.success !== true);
    const fileCount = results.filter(item => item.success === true).length;

    return {
      success: !hasFailure,
      results,
      message: fileCount > 1 ? `${fileCount} arquivos anexados` : 'Arquivo anexado'
    };
  } catch (error) {
    console.error('Erro ao processar anexos arrastados:', error);
    return { success: false, error: error.message };
  }
});

// Abre arquivo anexado
ipcMain.handle('open-attachment', async (event, supplierName, fileName, itemId = null) => {
  const attachmentsDir = getAttachmentsDir();
  let filePath;
  
  if (itemId) {
    filePath = path.join(attachmentsDir, sanitizeFileName(supplierName), String(itemId), fileName);
  } else {
    filePath = path.join(attachmentsDir, sanitizeFileName(supplierName), fileName);
  }
  
  if (!fs.existsSync(filePath)) {
    throw new Error('Arquivo não encontrado');
  }
  
  try {
    await shell.openPath(filePath);
    return { success: true };
  } catch (error) {
    console.error('Erro ao abrir arquivo:', error);
    throw error;
  }
});

// Exclui arquivo anexado
ipcMain.handle('delete-attachment', async (event, supplierName, fileName, itemId = null) => {
  const attachmentsDir = getAttachmentsDir();
  let filePath;
  
  if (itemId) {
    filePath = path.join(attachmentsDir, sanitizeFileName(supplierName), String(itemId), fileName);
  } else {
    filePath = path.join(attachmentsDir, sanitizeFileName(supplierName), fileName);
  }
  
  if (!fs.existsSync(filePath)) {
    return { success: false, error: 'Arquivo não encontrado' };
  }
  
  try {
    fs.unlinkSync(filePath);
    return { success: true };
  } catch (error) {
    console.error('Erro ao excluir arquivo:', error);
    return { success: false, error: error.message };
  }
});

// Função auxiliar para sanitizar nomes de arquivo/pasta
function sanitizeFileName(name) {
  return name.replace(/[<>:"/\\|?*]/g, '_').trim();
}

function createWindow () {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    frame: false,
    icon: path.join(__dirname, 'ferramentas.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.removeMenu();
  mainWindow.loadFile('index.html');
  
  // Abrir DevTools em desenvolvimento
  // mainWindow.webContents.openDevTools();
}

app.whenReady().then(async () => {
  try {
    await connectDatabase();
  } catch (error) {
    console.error('Falha ao inicializar o banco de dados:', error);
  }

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

// Todos handlers
ipcMain.handle('get-todos', async (event, toolingId) => {
  return new Promise((resolve, reject) => {
    db.all('SELECT * FROM todos WHERE tooling_id = ? ORDER BY created_at ASC', [toolingId], (err, rows) => {
      if (err) {
        console.error('Erro ao buscar todos:', err);
        reject(err);
      } else {
        resolve(rows);
      }
    });
  });
});

ipcMain.handle('add-todo', async (event, toolingId, text) => {
  return new Promise((resolve, reject) => {
    db.run('INSERT INTO todos (tooling_id, text, completed) VALUES (?, ?, 0)', [toolingId, text], function(err) {
      if (err) {
        console.error('Erro ao adicionar todo:', err);
        reject(err);
      } else {
        resolve({ id: this.lastID, tooling_id: toolingId, text, completed: 0 });
      }
    });
  });
});

ipcMain.handle('update-todo', async (event, todoId, text, completed) => {
  return new Promise((resolve, reject) => {
    db.run('UPDATE todos SET text = ?, completed = ? WHERE id = ?', [text, completed, todoId], (err) => {
      if (err) {
        console.error('Erro ao atualizar todo:', err);
        reject(err);
      } else {
        resolve({ success: true });
      }
    });
  });
});

ipcMain.handle('delete-todo', async (event, todoId) => {
  return new Promise((resolve, reject) => {
    db.run('DELETE FROM todos WHERE id = ?', [todoId], (err) => {
      if (err) {
        console.error('Erro ao deletar todo:', err);
        reject(err);
      } else {
        resolve({ success: true });
      }
    });
  });
});

// Window control handlers
ipcMain.on('minimize-window', () => {
  if (mainWindow) {
    mainWindow.minimize();
  }
});

ipcMain.on('maximize-window', () => {
  if (mainWindow) {
    if (mainWindow.isMaximized()) {
      mainWindow.unmaximize();
    } else {
      mainWindow.maximize();
    }
  }
});

ipcMain.on('close-window', () => {
  if (mainWindow) {
    mainWindow.close();
  }
});

app.on('window-all-closed', () => {
  if (db) {
    db.close();
  }
  if (process.platform !== 'darwin') {
    app.quit();
  }

});
