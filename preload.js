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
  getAnalytics: () => ipcRenderer.invoke('get-analytics'),
  updateTooling: (id, data) => ipcRenderer.invoke('update-tooling', id, data),
  createTooling: (data) => ipcRenderer.invoke('create-tooling', data),
  deleteTooling: (id) => ipcRenderer.invoke('delete-tooling', id),
  
  // Exportação de dados
  exportSupplierData: (supplierName) => ipcRenderer.invoke('export-supplier-data', supplierName),
  importSupplierData: (supplierName) => ipcRenderer.invoke('import-supplier-data', supplierName),
  exportAllData: () => ipcRenderer.invoke('export-all-data'),
  importAllData: () => ipcRenderer.invoke('import-all-data'),
  
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
  deleteAttachment: (supplierName, fileName, itemId = null) => 
    ipcRenderer.invoke('delete-attachment', supplierName, fileName, itemId),
  
  // Todos
  getTodos: (toolingId) => ipcRenderer.invoke('get-todos', toolingId),
  addTodo: (toolingId, text) => ipcRenderer.invoke('add-todo', toolingId, text),
  updateTodo: (todoId, text, completed) => ipcRenderer.invoke('update-todo', todoId, text, completed),
  deleteTodo: (todoId) => ipcRenderer.invoke('delete-todo', todoId),
  
  // Controles da janela
  minimizeWindow: () => ipcRenderer.send('minimize-window'),
  maximizeWindow: () => ipcRenderer.send('maximize-window'),
  closeWindow: () => ipcRenderer.send('close-window'),
  
  // Notificações
  onDataUpdated: (callback) => ipcRenderer.on('data-updated', callback)
});
