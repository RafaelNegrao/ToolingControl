const { contextBridge, ipcRenderer } = require('electron');

// Expõe APIs seguras para o renderer
contextBridge.exposeInMainWorld('api', {
  // Operações do banco de dados
  loadTooling: () => ipcRenderer.invoke('load-tooling'),
  getSuppliersWithStats: () => ipcRenderer.invoke('get-suppliers-with-stats'),
  getToolingBySupplier: (supplier) => ipcRenderer.invoke('get-tooling-by-supplier', supplier),
  searchTooling: (term) => ipcRenderer.invoke('search-tooling', term),
  getToolingById: (id) => ipcRenderer.invoke('get-tooling-by-id', id),
  getToolingByReplacementId: (replacementId) =>
    ipcRenderer.invoke('get-tooling-by-replacement-id', replacementId),
  getAllToolingIds: () => ipcRenderer.invoke('get-all-tooling-ids'),
  getIdsWithIncomingLinks: (targetIds) => ipcRenderer.invoke('get-ids-with-incoming-links', targetIds),
  getUniqueResponsibles: () => ipcRenderer.invoke('get-unique-responsibles'),
  getAnalytics: () => ipcRenderer.invoke('get-analytics'),
  getStepsSummary: () => ipcRenderer.invoke('get-steps-summary'),
  getStepSuppliersMetrics: (step) => ipcRenderer.invoke('get-step-suppliers-metrics', step),
  getStepChangeAverage: () => ipcRenderer.invoke('get-step-change-average'),
  updateTooling: (id, data) => ipcRenderer.invoke('update-tooling', id, data),
  createTooling: (data) => ipcRenderer.invoke('create-tooling', data),
  deleteTooling: (id) => ipcRenderer.invoke('delete-tooling', id),
  renameSupplier: (currentName, newName) => ipcRenderer.invoke('rename-supplier', currentName, newName),
  
  // Exportação de dados
  exportSupplierData: (supplierName, filteredIds = null) => ipcRenderer.invoke('export-supplier-data', supplierName, filteredIds),
  importSupplierData: (supplierName) => ipcRenderer.invoke('import-supplier-data', supplierName),
  exportEmptyTemplate: () => ipcRenderer.invoke('export-empty-template'),
  importNewSupplier: () => ipcRenderer.invoke('import-new-supplier'),
  exportForecastSupplier: () => ipcRenderer.invoke('export-forecast-supplier'),
  importForecastSupplier: () => ipcRenderer.invoke('import-forecast-supplier'),
  exportForecastManager: () => ipcRenderer.invoke('export-forecast-manager'),
  importForecastManager: () => ipcRenderer.invoke('import-forecast-manager'),
  
  // Gerenciamento de anexos
  getAttachments: (supplierName, itemId = null) => 
    ipcRenderer.invoke('get-attachments', supplierName, itemId),
  getAttachmentsCountBatch: (supplierName, itemIds) =>
    ipcRenderer.invoke('get-attachments-count-batch', supplierName, itemIds),
  uploadAttachment: (supplierName, itemId = null) => 
    ipcRenderer.invoke('upload-attachment', supplierName, itemId),
  uploadAttachmentFromPaths: (supplierName, filePaths, itemId = null) => 
    ipcRenderer.invoke('upload-attachment-from-paths', supplierName, filePaths, itemId),
  openAttachment: (supplierName, fileName, itemId = null) => 
    ipcRenderer.invoke('open-attachment', supplierName, fileName, itemId),
  openAttachmentsFolder: (supplierName) => 
    ipcRenderer.invoke('open-attachments-folder', supplierName),
  deleteAttachment: (supplierName, fileName, itemId = null) => 
    ipcRenderer.invoke('delete-attachment', supplierName, fileName, itemId),
  
  // Todos
  getTodos: (toolingId) => ipcRenderer.invoke('get-todos', toolingId),
  addTodo: (toolingId, text) => ipcRenderer.invoke('add-todo', toolingId, text),
  updateTodo: (todoId, text, completed) => ipcRenderer.invoke('update-todo', todoId, text, completed),
  deleteTodo: (todoId) => ipcRenderer.invoke('delete-todo', todoId),
  
  // Step History
  getStepHistory: (toolingId) => ipcRenderer.invoke('get-step-history', toolingId),
  clearStepHistory: (toolingId) => ipcRenderer.invoke('clear-step-history', toolingId),
  clearAllStepHistory: () => ipcRenderer.invoke('clear-all-step-history'),
  
  // Controles da janela
  minimizeWindow: () => ipcRenderer.send('minimize-window'),
  maximizeWindow: () => ipcRenderer.send('maximize-window'),
  closeWindow: () => ipcRenderer.send('close-window'),
  
  // DevTools
  openDevTools: () => ipcRenderer.send('open-devtools'),
  closeDevTools: () => ipcRenderer.send('close-devtools'),
  
  // Notificações
  onDataUpdated: (callback) => ipcRenderer.on('data-updated', callback)
});
