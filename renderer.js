// Estado da aplicação
const APP_VERSION = 'v0.1.1';

let currentTab = 'tooling';
let toolingData = [];
let suppliersData = [];
let selectedSupplier = null;
let currentSupplier = null;
let currentSortCriteria = null;
let currentSortOrder = 'asc';
let deleteConfirmState = { id: null, code: '', descriptor: '' };
let addToolingElements = {
  overlay: null,
  form: null,
  pnInput: null,
  supplierInput: null,
  supplierList: null,
  lifeInput: null,
  producedInput: null
};

let attachmentsElements = {
  dropzone: null,
  body: null,
  placeholder: null,
  counterButton: null,
  counterBadge: null,
  modalOverlay: null,
  modalList: null,
  modalEmpty: null,
  modalSupplier: null
};

let replacementTimelineElements = {
  overlay: null,
  list: null,
  empty: null,
  loading: null,
  title: null
};

let currentTimelineRootId = null;
let isReorderingTimeline = false;

let attachmentsDragCounter = 0;
let attachmentsData = [];
let activeSearchTerm = '';
const dateReminderTimers = new Map();
let expirationInfoElements = {
  overlay: null,
  closeButtons: []
};
let productionInfoElements = {
  overlay: null,
  closeButtons: []
};

let expirationFilterEnabled = false;

let replacementIdOptions = [];
let replacementIdOptionsLoaded = false;
let replacementIdOptionsPromise = null;

const STATUS_STORAGE_KEY = 'toolingStatusOptions';
const DEFAULT_STATUS_OPTIONS = [
  'ACTIVE',
  'CONCLUDED',
  'UNDER ANALYSIS',
  'UNDER CONSTRUCTION',
  'EXPIRED',
  'OBSOLETE',
  'RESOURCING'
];

const DEFAULT_REPLACEMENT_PICKER_LABEL = 'Select replacement';

let statusOptions = loadStatusOptionsFromStorage();
let statusSettingsElements = {
  list: null,
  input: null,
  addButton: null
};

const TEXT_INPUT_TYPES = new Set(['text', 'search', '']);
let uppercaseInputHandlerInitialized = false;

function enforceUppercaseValue(element) {
  if (!element || typeof element.value !== 'string') {
    return;
  }

  const { selectionStart, selectionEnd, value } = element;
  const uppercased = value.toUpperCase();

  if (uppercased !== value) {
    element.value = uppercased;
    if (typeof selectionStart === 'number' && typeof selectionEnd === 'number') {
      element.setSelectionRange(selectionStart, selectionEnd);
    }
  }
}

function handleUppercaseInput(event) {
  const target = event.target;
  const isTextInput = target instanceof HTMLInputElement && TEXT_INPUT_TYPES.has(target.type);
  const isTextarea = target instanceof HTMLTextAreaElement;

  if (isTextInput || isTextarea) {
    enforceUppercaseValue(target);
  }
}

function initUppercaseInputListeners() {
  if (uppercaseInputHandlerInitialized) {
    return;
  }

  document.addEventListener('input', handleUppercaseInput, true);
  uppercaseInputHandlerInitialized = true;
}

function initExpirationInfoModal() {
  const overlay = document.getElementById('expirationInfoOverlay');
  const closeButtons = document.querySelectorAll('[data-expiration-info-close]');

  expirationInfoElements = {
    overlay,
    closeButtons
  };

  if (!overlay) {
    return;
  }

  overlay.addEventListener('click', (event) => {
    if (event.target === overlay) {
      closeExpirationInfoModal();
    }
  });

  if (closeButtons && closeButtons.length > 0) {
    closeButtons.forEach((button) => {
      button.addEventListener('click', () => {
        closeExpirationInfoModal();
      });
    });
  }
}

function initProductionInfoModal() {
  const overlay = document.getElementById('productionInfoOverlay');
  const closeButtons = document.querySelectorAll('[data-production-info-close]');

  productionInfoElements = {
    overlay,
    closeButtons
  };

  if (!overlay) {
    return;
  }

  overlay.addEventListener('click', (event) => {
    if (event.target === overlay) {
      closeProductionInfoModal();
    }
  });

  if (closeButtons && closeButtons.length > 0) {
    closeButtons.forEach((button) => {
      button.addEventListener('click', () => {
        closeProductionInfoModal();
      });
    });
  }
}

function openProductionInfoModal(event) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }
  const { overlay, closeButtons } = productionInfoElements;
  if (!overlay) {
    return;
  }
  overlay.classList.add('active');
  requestAnimationFrame(() => {
    closeButtons?.[0]?.focus();
  });
}

function closeProductionInfoModal() {
  const { overlay } = productionInfoElements;
  if (!overlay) {
    return;
  }
  overlay.classList.remove('active');
}

function handleProductionInfoIconKey(event) {
  if (event.key === 'Enter' || event.key === ' ') {
    openProductionInfoModal(event);
  }
}

function openExpirationInfoModal(event) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }
  const { overlay, closeButtons } = expirationInfoElements;
  if (!overlay) {
    return;
  }
  overlay.classList.add('active');
  requestAnimationFrame(() => {
    closeButtons?.[0]?.focus();
  });
}

function closeExpirationInfoModal() {
  const { overlay } = expirationInfoElements;
  if (!overlay) {
    return;
  }
  overlay.classList.remove('active');
}

function handleExpirationInfoIconKey(event) {
  if (event.key === 'Enter' || event.key === ' ') {
    openExpirationInfoModal(event);
  }
}

function handleProductionInfoIconClick(event) {
  openProductionInfoModal(event);
}

async function refreshReplacementIdOptions(force = false) {
  if (replacementIdOptionsPromise && !force) {
    return replacementIdOptionsPromise;
  }

  if (force) {
    replacementIdOptionsLoaded = false;
  }

  replacementIdOptionsPromise = window.api.getAllToolingIds()
    .then((rows) => {
      const normalizedRows = Array.isArray(rows) ? rows : [];
      replacementIdOptions = normalizedRows
        .map((row) => {
          const numericId = Number(row?.id);
          if (!Number.isFinite(numericId)) {
            return null;
          }
          return {
            id: numericId,
            pn: row?.pn || '',
            pn_description: row?.pn_description || '',
            supplier: row?.supplier || '',
            tool_description: row?.tool_description || ''
          };
        })
        .filter(Boolean)
        .sort((a, b) => a.id - b.id);
      replacementIdOptionsLoaded = true;
      return replacementIdOptions;
    })
    .catch((error) => {
      console.error('Erro ao carregar IDs para substituição:', error);
      replacementIdOptions = [];
      replacementIdOptionsLoaded = false;
      return [];
    })
    .finally(() => {
      replacementIdOptionsPromise = null;
    });

  return replacementIdOptionsPromise;
}

async function ensureReplacementIdOptions(force = false) {
  if (!replacementIdOptionsLoaded || force) {
    await refreshReplacementIdOptions(force);
  }
  return replacementIdOptions;
}

function buildReplacementPickerOptionsMarkup(item, cardIndex) {
  const currentNumericId = Number(item.id);
  const options = Array.isArray(replacementIdOptions)
    ? replacementIdOptions.filter(option => option.id !== currentNumericId)
    : [];

  const clearButton = `<button type="button" class="replacement-dropdown-option" onclick="handleReplacementPickerSelect(${cardIndex}, ${item.id}, '')">Clear selection</button>`;

  if (options.length === 0) {
    return clearButton + '<p class="replacement-dropdown-empty">No other cards available</p>';
  }

  const optionButtons = options
    .map((option) => {
      const idText = escapeHtml(`#${option.id}`);
      const pnText = escapeHtml(option.pn || 'N/A');
      const supplierText = escapeHtml(option.supplier || 'N/A');
      const descText = escapeHtml(option.tool_description || 'N/A');
      
      return `<button type="button" class="replacement-dropdown-option" onclick="handleReplacementPickerSelect(${cardIndex}, ${item.id}, ${option.id})">
        <div class="replacement-option-grid">
          <span class="replacement-option-id">${idText}</span>
          <span class="replacement-option-pn">${pnText}</span>
          <span class="replacement-option-supplier">${supplierText}</span>
          <span class="replacement-option-desc">${descText}</span>
        </div>
      </button>`;
    })
    .join('');

  return clearButton + optionButtons;
}

function getReplacementOptionLabel(option) {
  return `#${option.id}`;
}

function getReplacementOptionLabelById(idValue) {
  const numericId = Number(idValue);
  if (!Number.isFinite(numericId)) {
    return null;
  }
  const match = Array.isArray(replacementIdOptions)
    ? replacementIdOptions.find(option => option.id === numericId)
    : null;
  return match ? getReplacementOptionLabel(match) : null;
}

function getToolingItemForCard(card) {
  if (!card) {
    return null;
  }
  const idValue = Number(card.getAttribute('data-item-id'));
  if (!Number.isFinite(idValue)) {
    return null;
  }
  return toolingData.find(item => Number(item.id) === idValue) || null;
}

async function rebuildReplacementPickerOptions(cardIndex) {
  await ensureReplacementIdOptions();
  const card = document.getElementById(`card-${cardIndex}`);
  if (!card) {
    return;
  }
  const list = card.querySelector('[data-replacement-picker-list]');
  if (!list) {
    return;
  }
  const item = getToolingItemForCard(card);
  if (!item) {
    return;
  }
  list.innerHTML = buildReplacementPickerOptionsMarkup(item, cardIndex);
}

// Gerenciamento de abas
document.addEventListener('DOMContentLoaded', async () => {
  initUppercaseInputListeners();

  // Popula versão do app
  const appVersionFooter = document.getElementById('appVersion');
  const appVersionSettings = document.getElementById('appVersionSettings');
  if (appVersionFooter) appVersionFooter.textContent = APP_VERSION;
  if (appVersionSettings) appVersionSettings.textContent = APP_VERSION;

  // Window control buttons
  const minimizeBtn = document.getElementById('minimizeBtn');
  const maximizeBtn = document.getElementById('maximizeBtn');
  const closeBtn = document.getElementById('closeBtn');

  if (minimizeBtn) {
    minimizeBtn.addEventListener('click', () => window.api.minimizeWindow());
  }

  if (maximizeBtn) {
    maximizeBtn.addEventListener('click', () => window.api.maximizeWindow());
  }

  if (closeBtn) {
    closeBtn.addEventListener('click', () => window.api.closeWindow());
  }

  const tabButtons = document.querySelectorAll('.tab-button');
  const tabContents = document.querySelectorAll('.tab-content');
  const sidebar = document.getElementById('sidebar');
  const searchInput = document.getElementById('searchInput');
  const reloadButton = document.getElementById('reloadData');
  const deleteOverlay = document.getElementById('deleteConfirmOverlay');
  const deleteInput = document.getElementById('deleteConfirmInput');
  const addOverlay = document.getElementById('addToolingOverlay');
  const addForm = document.getElementById('addToolingForm');
  const addPN = document.getElementById('addToolingPN');
  const addSupplierInput = document.getElementById('addToolingSupplier');
  const addSupplierList = document.getElementById('addToolingSupplierList');
  const addLife = document.getElementById('addToolingLife');
  const addProduced = document.getElementById('addToolingProduced');
  const attachmentsDropzone = document.getElementById('attachmentsDropzone');
  const attachmentsUploadTrigger = document.getElementById('attachmentsUploadTrigger');
  const attachmentsMessageLabel = document.getElementById('attachmentsMessageLabel');
  const attachmentsCounterButton = document.getElementById('attachmentsCountBtn');
  const attachmentsCounterBadge = document.getElementById('attachmentsCountLabel');
  const attachmentsModalOverlay = document.getElementById('attachmentsModalOverlay');
  const attachmentsModalList = document.getElementById('attachmentsModalList');
  const attachmentsModalEmpty = document.getElementById('attachmentsModalEmpty');
  const attachmentsModalSupplier = document.getElementById('attachmentsModalSupplier');
  const replacementTimelineOverlay = document.getElementById('replacementTimelineOverlay');
  const replacementTimelineList = document.getElementById('replacementTimelineList');
  const replacementTimelineEmpty = document.getElementById('replacementTimelineEmpty');
  const replacementTimelineLoading = document.getElementById('replacementTimelineLoading');
  const replacementTimelineTitle = document.getElementById('replacementTimelineTitle');
  const statusOptionsList = document.getElementById('statusOptionsList');
  const statusOptionInput = document.getElementById('statusOptionInput');
  const addStatusButton = document.getElementById('addStatusButton');
  const supplierSearchInput = document.getElementById('supplierSearchInput');
  const clearSupplierSearchBtn = document.getElementById('clearSupplierSearch');
  const expirationFilterSwitch = document.getElementById('expirationFilterSwitch');
  const supplierMenuBtn = document.getElementById('supplierMenuBtn');
  const supplierFilterOverlay = document.getElementById('supplierFilterOverlay');

  addToolingElements = {
    overlay: addOverlay,
    form: addForm,
    pnInput: addPN,
    supplierInput: addSupplierInput,
    supplierList: addSupplierList,
    lifeInput: addLife,
    producedInput: addProduced
  };

  attachmentsElements = {
    dropzone: attachmentsDropzone,
    uploadButton: attachmentsUploadTrigger,
    messageLabel: attachmentsMessageLabel,
    counterButton: attachmentsCounterButton,
    counterBadge: attachmentsCounterBadge,
    modalOverlay: attachmentsModalOverlay,
    modalList: attachmentsModalList,
    modalEmpty: attachmentsModalEmpty,
    modalSupplier: attachmentsModalSupplier
  };

  replacementTimelineElements = {
    overlay: replacementTimelineOverlay,
    list: replacementTimelineList,
    empty: replacementTimelineEmpty,
    loading: replacementTimelineLoading,
    title: replacementTimelineTitle
  };

  if (replacementTimelineOverlay) {
    replacementTimelineOverlay.addEventListener('click', (event) => {
      if (event.target === replacementTimelineOverlay) {
        closeReplacementTimelineOverlay();
      }
    });
  }

  initStatusSettings({
    list: statusOptionsList,
    input: statusOptionInput,
    addButton: addStatusButton
  });

  initAttachmentsDragAndDrop();
  initExpirationInfoModal();
  initProductionInfoModal();

  if (attachmentsCounterBadge) {
    attachmentsCounterBadge.textContent = '0';
  }

  if (attachmentsCounterButton) {
    attachmentsCounterButton.disabled = true;
  }

  if (attachmentsModalEmpty) {
    attachmentsModalEmpty.style.display = 'block';
  }

  if (attachmentsModalOverlay) {
    attachmentsModalOverlay.addEventListener('click', (e) => {
      if (e.target === attachmentsModalOverlay) {
        closeAttachmentsModal();
      }
    });
  }

  // Função para trocar de aba
  function switchTab(tabName) {
    currentTab = tabName;
    
    // Remove active de todos os botões e conteúdos
    tabButtons.forEach(btn => btn.classList.remove('active'));
    tabContents.forEach(content => content.classList.remove('active'));

    // Adiciona active ao botão e conteúdo correspondentes
    const activeButton = document.querySelector(`[data-tab="${tabName}"]`);
    const activeContent = document.getElementById(`tab-${tabName}`);

    if (activeButton && activeContent) {
      activeButton.classList.add('active');
      activeContent.classList.add('active');
    }

    // Mostra/esconde sidebar apenas na aba Tooling
    if (tabName === 'tooling') {
      sidebar.classList.remove('hidden');
    } else {
      sidebar.classList.add('hidden');
    }

    // Mostra/esconde botões flutuantes apenas na aba Tooling
    const floatingActions = document.querySelector('.floating-actions');
    if (floatingActions) {
      if (tabName === 'tooling') {
        floatingActions.style.display = 'flex';
      } else {
        floatingActions.style.display = 'none';
      }
    }

    // Carrega dados específicos da aba
    if (tabName === 'analytics') {
      loadAnalytics();
    }
  }

  // Event listeners para os botões
  tabButtons.forEach(button => {
    button.addEventListener('click', () => {
      const tabName = button.getAttribute('data-tab');
      switchTab(tabName);
    });
  });

  // Funcionalidade de pesquisa
  if (searchInput) {
    searchInput.addEventListener('input', async (e) => {
      const searchTerm = e.target.value.trim();
      if (searchTerm.length >= 2) {
        activeSearchTerm = searchTerm;
        await searchTooling(searchTerm);
        updateSearchIndicators();
      } else if (searchTerm.length === 0) {
        await clearSearch({ keepOverlayOpen: true });
      }
    });
  }

  // Botão de recarregar dados
  if (reloadButton) {
    reloadButton.addEventListener('click', async () => {
      await loadSuppliers();
      await loadAnalytics();
      if (selectedSupplier) {
        await loadToolingBySupplier(selectedSupplier);
      }
      showNotification('Data reloaded successfully!');
    });
  }

  if (deleteInput) {
    deleteInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        handleDeleteConfirmation();
      }
    });
  }

  if (deleteOverlay) {
    deleteOverlay.addEventListener('click', (e) => {
      if (e.target === deleteOverlay) {
        cancelDeleteTooling();
      }
    });
  }

  if (addForm) {
    addForm.addEventListener('submit', (e) => {
      e.preventDefault();
      submitAddToolingForm();
    });
  }

  if (addOverlay) {
    addOverlay.addEventListener('click', (e) => {
      if (e.target === addOverlay) {
        closeAddToolingModal();
      }
    });
  }

  // Listener para a barra de pesquisa de suppliers
  if (supplierSearchInput) {
    supplierSearchInput.addEventListener('input', async (e) => {
      const searchTerm = e.target.value.trim();
      if (clearSupplierSearchBtn) {
        clearSupplierSearchBtn.style.display = searchTerm ? 'flex' : 'none';
      }
      await filterSuppliersAndTooling(searchTerm);
    });
  }

  // Listener para o switch de filtro de expiração
  if (expirationFilterSwitch) {
    expirationFilterSwitch.addEventListener('change', (e) => {
      toggleExpirationFilter(e.target.checked);
    });
  }

  // Listener para o botão de menu de suppliers
  if (supplierMenuBtn && supplierFilterOverlay) {
    supplierMenuBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      openSupplierFilterOverlay();
    });

    // Fechar overlay ao clicar no fundo
    supplierFilterOverlay.addEventListener('click', (e) => {
      if (e.target === supplierFilterOverlay) {
        closeSupplierFilterOverlay();
      }
    });
  }

  // Inicializa na primeira aba (Tooling)
  switchTab('tooling');
  
  // Event listener global para auto-save ao sair do campo (blur)
  document.addEventListener('focusout', (e) => {
    const target = e.target;
    // Verifica se é um input/textarea/select com data-id e data-field
    if ((target.matches('input[data-id][data-field]') || 
         target.matches('textarea[data-id][data-field]') || 
         target.matches('select[data-id][data-field]')) &&
        !target.classList.contains('calculated')) {
      const id = target.getAttribute('data-id');
      if (id) {
        autoSaveTooling(parseInt(id), true); // true = salvamento imediato
      }
    }
  });
  
  // Carrega dados iniciais
  await refreshReplacementIdOptions();
  await loadSuppliers();
  await loadAnalytics();
  updateSearchIndicators();
});

// Controla o overlay de pesquisa
function toggleSearchOverlay(forceState) {
  const overlay = document.getElementById('searchOverlay');
  const searchInput = document.getElementById('searchInput');
  if (!overlay || !searchInput) return;

  const shouldOpen = typeof forceState === 'boolean'
    ? forceState
    : !overlay.classList.contains('active');

  if (shouldOpen) {
    overlay.classList.add('active');
    searchInput.value = activeSearchTerm;
    setTimeout(() => searchInput.focus(), 100);
  } else {
    overlay.classList.remove('active');
  }

  updateSearchIndicators();
}

function handleFloatingSearchClick() {
  const overlay = document.getElementById('searchOverlay');
  const isActive = overlay?.classList.contains('active');
  toggleSearchOverlay(!isActive);
}

async function handleSearchAction() {
  const overlay = document.getElementById('searchOverlay');
  const isActive = overlay?.classList.contains('active');
  if (activeSearchTerm) {
    await clearSearch({ closeOverlay: true });
  } else if (isActive) {
    toggleSearchOverlay(false);
  }
}

async function clearSearch(options = {}) {
  const { closeOverlay = false, keepOverlayOpen = false } = options;
  const searchInput = document.getElementById('searchInput');

  activeSearchTerm = '';
  if (searchInput) {
    searchInput.value = '';
  }

  if (selectedSupplier) {
    await loadToolingBySupplier(selectedSupplier);
  } else {
    toolingData = [];
    displayTooling([]);
  }

  if (closeOverlay && !keepOverlayOpen) {
    const overlay = document.getElementById('searchOverlay');
    overlay?.classList.remove('active');
  }

  updateSearchIndicators();
}

function updateSearchIndicators() {
  const floatingBtn = document.getElementById('floatingSearchBtn');
  const searchActionBtn = document.getElementById('searchActionBtn');
  const overlay = document.getElementById('searchOverlay');
  const searchInput = document.getElementById('searchInput');

  if (floatingBtn) {
    floatingBtn.classList.toggle('is-filtering', Boolean(activeSearchTerm));
    floatingBtn.setAttribute('title', activeSearchTerm ? 'Search filter active' : 'Search tooling');
  }

  if (searchActionBtn) {
    const icon = searchActionBtn.querySelector('i');
    searchActionBtn.classList.toggle('is-clear', Boolean(activeSearchTerm));
    searchActionBtn.setAttribute('title', activeSearchTerm ? 'Clear search filter' : 'Close search');
    if (icon) {
      icon.className = activeSearchTerm ? 'ph ph-x-circle' : 'ph ph-x';
    }
  }

  if (overlay && !overlay.classList.contains('active') && searchInput && document.activeElement === searchInput) {
    searchInput.blur();
  }
}

// Fecha overlay ao clicar fora
document.addEventListener('click', (e) => {
  const overlay = document.getElementById('searchOverlay');
  if (overlay && e.target === overlay) {
    toggleSearchOverlay(false);
  }
});

document.addEventListener('click', (event) => {
  if (event.target.closest('[data-replacement-picker]')) {
    return;
  }
  closeAllReplacementPickers();
});

// Fecha card de ferramental ao clicar fora
document.addEventListener('click', (event) => {
  const toolingCard = event.target.closest('.tooling-card');
  const expandedCard = document.querySelector('.tooling-card.expanded');
  
  // Se há um card expandido e o clique foi fora dele
  if (expandedCard && !toolingCard) {
    const cardIndex = Array.from(document.querySelectorAll('.tooling-card')).indexOf(expandedCard);
    if (cardIndex !== -1) {
      toggleCard(cardIndex);
    }
  }
});

// Fecha overlay com tecla ESC
document.addEventListener('keydown', (e) => {
  const overlay = document.getElementById('searchOverlay');
  if (overlay && overlay.classList.contains('active') && e.key === 'Escape') {
    toggleSearchOverlay(false);
  }
  const deleteOverlay = document.getElementById('deleteConfirmOverlay');
  if (deleteOverlay && deleteOverlay.classList.contains('active') && e.key === 'Escape') {
    cancelDeleteTooling();
  }
  const addOverlay = document.getElementById('addToolingOverlay');
  if (addOverlay && addOverlay.classList.contains('active') && e.key === 'Escape') {
    closeAddToolingModal();
  }
  const attachmentsOverlay = attachmentsElements.modalOverlay || document.getElementById('attachmentsModalOverlay');
  if (attachmentsOverlay && attachmentsOverlay.classList.contains('active') && e.key === 'Escape') {
    closeAttachmentsModal();
  }
  const expirationOverlay = expirationInfoElements.overlay;
  if (expirationOverlay && expirationOverlay.classList.contains('active') && e.key === 'Escape') {
    closeExpirationInfoModal();
  }
  const productionOverlay = productionInfoElements.overlay;
  if (productionOverlay && productionOverlay.classList.contains('active') && e.key === 'Escape') {
    closeProductionInfoModal();
  }
  const timelineOverlay = replacementTimelineElements.overlay;
  if (timelineOverlay && timelineOverlay.classList.contains('active') && e.key === 'Escape') {
    closeReplacementTimelineOverlay();
  }
  if (e.key === 'Escape') {
    closeAllReplacementPickers();
  }
});

// Carrega fornecedores com estatísticas
async function loadSuppliers() {
  try {
    suppliersData = await window.api.getSuppliersWithStats();
    console.log('Suppliers loaded:', suppliersData);
    applyExpirationFilter();
    populateAddToolingSuppliers();
  } catch (error) {
    console.error('Error loading suppliers:', error);
    showNotification('Error loading suppliers', 'error');
  }
}

function applyExpirationFilter() {
  if (!suppliersData) {
    displaySuppliers([]);
    return;
  }

  if (expirationFilterEnabled) {
    const filteredSuppliers = suppliersData.filter(supplier => {
      const expired = parseInt(supplier.expired) || 0;
      const warning1 = parseInt(supplier.warning_1year) || 0;
      const warning2 = parseInt(supplier.warning_2years) || 0;
      const critical = expired + warning1 + warning2;
      return critical > 0;
    });
    displaySuppliers(filteredSuppliers);
  } else {
    displaySuppliers(suppliersData);
  }
}

function toggleExpirationFilter(enabled) {
  expirationFilterEnabled = enabled;
  const badge = document.getElementById('filterActiveBadge');
  if (badge) {
    badge.style.display = enabled ? 'flex' : 'none';
  }
  applyExpirationFilter();
  // Recarrega os cards do supplier selecionado para aplicar o filtro
  if (selectedSupplier) {
    loadToolingBySupplier(selectedSupplier);
  }
}

function clearExpirationFilter() {
  const filterSwitch = document.getElementById('expirationFilterSwitch');
  if (filterSwitch) {
    filterSwitch.checked = false;
  }
  toggleExpirationFilter(false);
}

function openSupplierFilterOverlay() {
  const overlay = document.getElementById('supplierFilterOverlay');
  if (overlay) {
    overlay.classList.add('active');
  }
}

function closeSupplierFilterOverlay() {
  const overlay = document.getElementById('supplierFilterOverlay');
  if (overlay) {
    overlay.classList.remove('active');
  }
}

// Variável para gerenciar timeout do toast
let toastTimeout = null;

// Função para mostrar mensagens toast
function showToast(message, type = 'success') {
  const toast = document.getElementById('toast');
  const toastMessage = document.getElementById('toastMessage');
  const toastIcon = toast.querySelector('.toast-icon i');
  
  // Limpar timeout anterior se existir
  if (toastTimeout) {
    clearTimeout(toastTimeout);
    toastTimeout = null;
  }
  
  // Remover classe show temporariamente para forçar re-render
  toast.classList.remove('show');
  
  // Aguardar frame seguinte antes de mostrar novamente
  requestAnimationFrame(() => {
    // Definir ícone baseado no tipo
    if (type === 'success') {
      toastIcon.className = 'ph ph-check-circle';
      toast.style.background = 'linear-gradient(135deg, #10b981 0%, #059669 100%)';
    } else if (type === 'error') {
      toastIcon.className = 'ph ph-x-circle';
      toast.style.background = 'linear-gradient(135deg, #ef4444 0%, #dc2626 100%)';
    } else if (type === 'info') {
      toastIcon.className = 'ph ph-info';
      toast.style.background = 'linear-gradient(135deg, #3b82f6 0%, #2563eb 100%)';
    }
    
    toastMessage.textContent = message;
    toast.classList.add('show');
    
    toastTimeout = setTimeout(() => {
      toast.classList.remove('show');
      toastTimeout = null;
    }, 4000);
  });
}

// Habilitar/desabilitar botões de exportar/importar
function updateSupplierDataButtons(enabled) {
  const exportBtn = document.getElementById('exportSupplierBtn');
  const importBtn = document.getElementById('importSupplierBtn');
  
  if (exportBtn) exportBtn.disabled = !enabled;
  if (importBtn) importBtn.disabled = !enabled;
}

// Exportar dados do supplier para Excel
async function exportSupplierData() {
  if (!currentSupplier) {
    showToast('Please select a supplier first', 'error');
    return;
  }

  try {
    showToast('Exporting supplier data...', 'info');
    const result = await window.api.exportSupplierData(currentSupplier);
    
    if (result.success) {
      showToast('Data exported successfully!', 'success');
    } else {
      showToast('Export cancelled', 'info');
    }
  } catch (error) {
    console.error('Error exporting data:', error);
    showToast('Error exporting data. Please try again.', 'error');
  }
}

// Importar dados do supplier a partir de Excel
async function importSupplierData() {
  if (!currentSupplier) {
    showToast('Please select a supplier first', 'error');
    return;
  }

  try {
    showToast('Importing supplier data...', 'info');
    const result = await window.api.importSupplierData(currentSupplier);

    if (!result || !result.success) {
      if (result && result.message === 'Import cancelled') {
        showToast('Import cancelled', 'info');
        return;
      }
      const errorMsg = result?.message || 'Error importing data. Please try again.';
      showToast(errorMsg, 'error');
      return;
    }

    const updated = result.updated ?? 0;
    const created = result.created ?? 0;
    const skipped = result.skipped ?? 0;
    
    let successMsg = 'Import finished: ';
    const parts = [];
    if (updated > 0) parts.push(`${updated} updated`);
    if (created > 0) parts.push(`${created} created`);
    if (skipped > 0) parts.push(`${skipped} skipped`);
    successMsg += parts.join(', ');
    
    showToast(successMsg, 'success');

    await loadToolingBySupplier(currentSupplier);
    await loadSuppliers();
  } catch (error) {
    console.error('Error importing data:', error);
    const fallbackMessage = error?.message || 'Error importing data. Please try again.';
    showToast(fallbackMessage, 'error');
  }
}

// Exibe cards de fornecedores no sidebar
function displaySuppliers(suppliers) {
  const supplierList = document.getElementById('supplierList');
  
  if (!suppliers || suppliers.length === 0) {
    supplierList.innerHTML = '<p style="color: #999; font-size: 12px; text-align: center;">No suppliers found</p>';
    updateStatusBar({ total: 0, expired_total: 0, expiring_two_years: 0 });
    return;
  }

  supplierList.innerHTML = suppliers.map(supplier => {
    const supplierNameRaw = String(supplier.supplier || '');
    const supplierNameForHandler = supplierNameRaw.replace(/'/g, "&#39;");
    const supplierNameHtml = escapeHtml(supplierNameRaw);
    const isActive = selectedSupplier === supplierNameRaw;
    const total = parseInt(supplier.total) || 0;
    const expired = parseInt(supplier.expired) || 0;
    const warning1 = parseInt(supplier.warning_1year) || 0;
    const warning2 = parseInt(supplier.warning_2years) || 0;
    const critical = expired + warning1 + warning2;
    
    return `
        <div class="supplier-card ${isActive ? 'active' : ''}" 
          data-supplier="${supplierNameHtml}" 
          onclick="selectSupplier(event, '${supplierNameForHandler}')">
      <div class="supplier-card-header">
        <i class="ph ph-factory"></i>
        <h4 title="${supplierNameHtml}">${supplierNameHtml}</h4>
      </div>
      <div class="supplier-info">
        <div class="supplier-info-row">
          <span class="info-label">Total Tooling:</span>
          <span class="info-value">${total}</span>
        </div>
        <div class="supplier-info-row">
          <span class="info-label">Expired:</span>
          <span class="info-value ${expired > 0 ? 'expired' : ''}">${expired}</span>
        </div>
        <div class="supplier-info-row">
          <span class="info-label">Expiring within 2 years:</span>
          <span class="info-value ${critical > 0 ? 'critical' : ''}">${critical}</span>
        </div>
      </div>
    </div>
    `;
  }).join('');

  syncStatusBarWithSuppliers();
}

function setActiveSupplierCard(supplierName, explicitElement = null) {
  const cards = document.querySelectorAll('.supplier-card');
  cards.forEach(card => card.classList.remove('active'));
  const normalizedName = typeof supplierName === 'string' ? supplierName.trim() : '';
  if (!normalizedName) {
    return null;
  }
  const fallbackElement = Array.from(cards).find(card => (card.dataset.supplier || '').trim() === normalizedName);
  const targetCard = explicitElement || fallbackElement || null;
  if (targetCard) {
    targetCard.classList.add('active');
  }
  return targetCard;
}

async function handleSupplierSelection(supplierName, { sourceElement = null, forceReload = false } = {}) {
  const normalizedName = typeof supplierName === 'string' ? supplierName.trim() : '';
  if (!normalizedName) {
    return;
  }

  const previousSupplier = selectedSupplier;
  currentSupplier = normalizedName;
  selectedSupplier = normalizedName;
  setActiveSupplierCard(normalizedName, sourceElement);

  const attachmentsContainer = document.getElementById('attachmentsContainer');
  const currentSupplierName = document.getElementById('currentSupplierName');
  if (attachmentsContainer && currentSupplierName) {
    attachmentsContainer.style.display = 'block';
    currentSupplierName.textContent = normalizedName;
  }

  const shouldReload = forceReload || previousSupplier !== normalizedName;
  if (shouldReload) {
    await loadAttachments(normalizedName);
    
    // Verificar buscas ativas (geral e por supplier)
    const searchInput = document.getElementById('searchInput');
    const currentSearchValue = searchInput ? searchInput.value.trim() : '';
    const supplierSearchInput = document.getElementById('supplierSearchInput');
    const supplierSearchTerm = supplierSearchInput ? supplierSearchInput.value.trim() : '';
    const hasGlobalSearch = currentSearchValue.length >= 2;
    const hasSupplierSearch = supplierSearchTerm.length >= 1;
    
    if (hasGlobalSearch) {
      // Sincronizar e reaplicar busca global
      activeSearchTerm = currentSearchValue;
      await searchTooling(activeSearchTerm);
    } else if (hasSupplierSearch) {
      // Reaplicar filtro de suppliers
      await filterSuppliersAndTooling(supplierSearchTerm);
    } else {
      await loadToolingBySupplier(normalizedName);
    }
  }
  
  // Habilitar botões de exportar/importar
  updateSupplierDataButtons(true);
}

// Seleciona fornecedor e exibe ferramentais
async function selectSupplier(evt, supplierName) {
  const cardElement = evt?.currentTarget || evt?.target?.closest('.supplier-card');
  
  // Se clicar no supplier já selecionado, deseleciona
  if (selectedSupplier === supplierName) {
    selectedSupplier = null;
    currentSupplier = null;
    
    // Remove classe active de todos os cards
    const cards = document.querySelectorAll('.supplier-card');
    cards.forEach(card => card.classList.remove('active'));
    
    // Esconde os attachments e limpa a lista de ferramentais
    const attachmentsContainer = document.getElementById('attachmentsContainer');
    if (attachmentsContainer) {
      attachmentsContainer.style.display = 'none';
    }
    
    const toolingList = document.getElementById('toolingList');
    const emptyState = document.getElementById('emptyState');
    if (toolingList) toolingList.style.display = 'none';
    if (emptyState) emptyState.style.display = 'none';
    
    // Desabilitar botões de exportar/importar
    updateSupplierDataButtons(false);
    
    return;
  }
  
  await handleSupplierSelection(supplierName, { sourceElement: cardElement, forceReload: true });
}

function initAttachmentsDragAndDrop() {
  const { dropzone } = attachmentsElements;
  if (!dropzone) {
    return;
  }

  dropzone.addEventListener('dragenter', onAttachmentsDragEnter, false);
  dropzone.addEventListener('dragover', onAttachmentsDragOver, false);
  dropzone.addEventListener('dragleave', onAttachmentsDragLeave, false);
  dropzone.addEventListener('drop', onAttachmentsDrop, false);

  document.addEventListener('dragover', handleWindowDragOver, false);
  document.addEventListener('drop', handleWindowDrop, false);
}

function handleWindowDragOver(event) {
  if (isFileDrag(event)) {
    event.preventDefault();
    event.stopPropagation();
  }
}

function handleWindowDrop(event) {
  if (!isFileDrag(event)) {
    return;
  }

  const { dropzone } = attachmentsElements;
  if (dropzone && dropzone.contains(event.target)) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  attachmentsDragCounter = 0;
  dropzone?.classList.remove('drop-active');
}

function isFileDrag(event) {
  const types = Array.from(event?.dataTransfer?.types || []);
  return types.includes('Files');
}

function onAttachmentsDragEnter(event) {
  if (!isFileDrag(event)) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  attachmentsDragCounter += 1;

  const { dropzone } = attachmentsElements;
  if (!dropzone) {
    return;
  }

  if (event.dataTransfer) {
    event.dataTransfer.dropEffect = currentSupplier ? 'copy' : 'none';
  }

  dropzone.classList.add('drop-active');
}

function onAttachmentsDragOver(event) {
  if (!isFileDrag(event)) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();

  const { dropzone } = attachmentsElements;
  if (!dropzone) {
    return;
  }

  if (event.dataTransfer) {
    event.dataTransfer.dropEffect = currentSupplier ? 'copy' : 'none';
  }

  dropzone.classList.add('drop-active');
}

function onAttachmentsDragLeave(event) {
  if (!isFileDrag(event)) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  attachmentsDragCounter = Math.max(attachmentsDragCounter - 1, 0);

  if (attachmentsDragCounter === 0) {
    attachmentsElements.dropzone?.classList.remove('drop-active');
  }
}

async function onAttachmentsDrop(event) {
  if (!isFileDrag(event)) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  attachmentsDragCounter = 0;
  attachmentsElements.dropzone?.classList.remove('drop-active');

  const files = Array.from(event.dataTransfer?.files || []);
  if (!files.length) {
    return;
  }

  await handleAttachmentFiles(files);
}

async function handleAttachmentFiles(files) {
  if (!files || files.length === 0) {
    return;
  }

  if (!currentSupplier) {
    showNotification('Select a supplier before attaching files.', 'error');
    return;
  }

  const paths = files
    .map(file => file?.path)
    .filter(Boolean);

  if (paths.length === 0) {
    showNotification('Unable to read the dragged files.', 'error');
    return;
  }

  if (typeof window.api.uploadAttachmentFromPaths !== 'function') {
    showNotification('Drag and drop is currently unavailable.', 'error');
    return;
  }

  try {
    const result = await window.api.uploadAttachmentFromPaths(currentSupplier, paths);

    if (!result) {
      showNotification('Unable to attach the files.', 'error');
      return;
    }

    const failures = Array.isArray(result.results)
      ? result.results.filter(item => item?.success !== true)
      : [];

    if (failures.length > 0 || result.success !== true) {
      const errorMessage = failures[0]?.error || result.error || 'Error attaching files.';
      showNotification(errorMessage, 'error');
    } else {
      const successMessage = paths.length > 1
        ? 'Files attached successfully!'
        : 'File attached successfully!';
      showNotification(successMessage);
    }

    await loadAttachments(currentSupplier);
  } catch (error) {
    console.error('Error attaching files via drag and drop:', error);
    showNotification('Error attaching files.', 'error');
  }
}

// Carrega anexos de um fornecedor
async function loadAttachments(supplierName) {
  try {
    const attachments = await window.api.getAttachments(supplierName);
    displayAttachments(attachments);
  } catch (error) {
    console.error('Error loading attachments:', error);
    displayAttachments([]);
  }
}

// Exibe lista de anexos
function displayAttachments(attachments) {
  attachmentsData = Array.isArray(attachments) ? attachments : [];

  const {
    uploadButton,
    messageLabel,
    counterBadge,
    counterButton,
    modalOverlay,
    modalSupplier
  } = attachmentsElements;

  const supplierLabel = currentSupplier || selectedSupplier || '';
  const attachmentsCount = attachmentsData.length;

  if (counterBadge) {
    counterBadge.textContent = attachmentsCount;
  }

  if (counterButton) {
    counterButton.disabled = !supplierLabel;
    counterButton.classList.toggle('has-items', attachmentsCount > 0);
    counterButton.setAttribute('title', supplierLabel ? 'View attachments' : 'Select a supplier first');
  }

  if (uploadButton) {
    const isEnabled = Boolean(supplierLabel);
    uploadButton.disabled = !isEnabled;
    uploadButton.setAttribute('title', isEnabled ? 'Click to attach files' : 'Select a supplier first');
  }

  if (modalSupplier) {
    modalSupplier.textContent = supplierLabel || 'Supplier';
  }

  if (messageLabel) {
    messageLabel.textContent = attachmentsCount === 0
      ? 'Click or drop attachments here.'
      : `Drag new files or click to attach (total: ${attachmentsCount}).`;
  }

  if (modalOverlay && modalOverlay.classList.contains('active')) {
    renderAttachmentsModal();
  }
}

function renderAttachmentsModal() {
  const { modalList, modalEmpty } = attachmentsElements;

  if (!modalList || !modalEmpty) {
    return;
  }

  if (!attachmentsData || attachmentsData.length === 0) {
    modalList.innerHTML = '';
    modalEmpty.style.display = 'block';
    modalList.style.display = 'none';
    return;
  }

  modalEmpty.style.display = 'none';
  modalList.style.display = 'flex';

  modalList.innerHTML = attachmentsData.map(att => `
    <div class="attachment-item">
      <div class="attachment-info">
        <i class="ph ph-file ${getFileIcon(att.fileName)}" style="color: ${getFileColor(att.fileName)}"></i>
        <div class="attachment-details">
          <div class="attachment-name">${att.fileName}</div>
          <div class="attachment-meta">${formatFileSize(att.fileSize)} • ${formatDate(att.uploadDate)}</div>
        </div>
      </div>
      <div class="attachment-actions">
        <button class="btn-attachment" onclick="openAttachment('${att.supplierName}', '${att.fileName}')" title="Open file">
          <i class="ph ph-eye"></i>
        </button>
        <button class="btn-attachment" onclick="deleteAttachment('${att.supplierName}', '${att.fileName}')" title="Delete file">
          <i class="ph ph-trash"></i>
        </button>
      </div>
    </div>
  `).join('');
}

function openAttachmentsModal() {
  const { modalOverlay, modalSupplier } = attachmentsElements;
  const supplierLabel = currentSupplier || selectedSupplier;

  if (!supplierLabel) {
    showNotification('Select a supplier to view attachments.', 'error');
    return;
  }

  if (!modalOverlay) {
    return;
  }

  if (modalSupplier) {
    modalSupplier.textContent = supplierLabel;
  }

  renderAttachmentsModal();
  modalOverlay.classList.add('active');
}

function closeAttachmentsModal() {
  const { modalOverlay } = attachmentsElements;
  if (modalOverlay) {
    modalOverlay.classList.remove('active');
  }
}

// Retorna ícone baseado na extensão do arquivo
function getFileIcon(fileName) {
  const ext = fileName.split('.').pop().toLowerCase();
  const iconMap = {
    pdf: 'ph-file-pdf',
    doc: 'ph-file-doc',
    docx: 'ph-file-doc',
    xls: 'ph-file-xls',
    xlsx: 'ph-file-xls',
    jpg: 'ph-file-image',
    jpeg: 'ph-file-image',
    png: 'ph-file-image',
    zip: 'ph-file-zip',
    rar: 'ph-file-zip'
  };
  return iconMap[ext] || 'ph-file';
}

// Retorna cor baseada na extensão do arquivo
function getFileColor(fileName) {
  const ext = fileName.split('.').pop().toLowerCase();
  const colorMap = {
    pdf: '#c8102e',
    doc: '#2b579a',
    docx: '#2b579a',
    xls: '#217346',
    xlsx: '#217346',
    jpg: '#ff6b6b',
    jpeg: '#ff6b6b',
    png: '#ff6b6b',
    zip: '#ffa502',
    rar: '#ffa502'
  };
  return colorMap[ext] || '#666';
}

// Formata tamanho do arquivo
function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// Upload de arquivo
async function uploadAttachment() {
  if (!currentSupplier) {
    alert('Select a supplier first');
    return;
  }
  
  try {
    const result = await window.api.uploadAttachment(currentSupplier);
    if (result.success) {
      await loadAttachments(currentSupplier);
      // Suporta resposta de múltiplos arquivos
      const message = result.message || 'File attached successfully!';
      showNotification(message);
      
      // Mostra erros individuais se houver
      if (result.results) {
        const failures = result.results.filter(r => !r.success);
        if (failures.length > 0) {
          const errorMsg = failures.map(f => `${f.fileName}: ${f.error}`).join('\n');
          console.error('Upload failures:', errorMsg);
        }
      }
    } else {
      showNotification('Error attaching file(s)', 'error');
    }
  } catch (error) {
    console.error('Error uploading file:', error);
    alert('Error uploading file');
  }
}

// Abre arquivo anexado
async function openAttachment(supplierName, fileName) {
  try {
    await window.api.openAttachment(supplierName, fileName);
  } catch (error) {
    console.error('Error opening file:', error);
    alert('Error opening file');
  }
}

// Exclui arquivo anexado
async function deleteAttachment(supplierName, fileName) {
  if (!confirm(`Are you sure you want to delete the file "${fileName}"?`)) {
    return;
  }
  
  try {
    const result = await window.api.deleteAttachment(supplierName, fileName);
    if (result.success) {
      await loadAttachments(supplierName);
    }
  } catch (error) {
    console.error('Error deleting file:', error);
    alert('Error deleting file');
  }
}

// Abre arquivo anexado do card
async function openCardAttachmentFile(supplierName, fileName, itemId) {
  try {
    await window.api.openAttachment(supplierName, fileName, itemId);
  } catch (error) {
    console.error('Error opening card attachment:', error);
    showNotification('Error opening file', 'error');
  }
}

// Exclui arquivo anexado do card
async function deleteCardAttachmentFile(supplierName, fileName, itemId) {
  if (!confirm(`Are you sure you want to delete the file "${fileName}"?`)) {
    return;
  }
  
  try {
    const result = await window.api.deleteAttachment(supplierName, fileName, itemId);
    if (result.success) {
      showNotification('File deleted successfully!');
      await loadCardAttachments(itemId);
    }
  } catch (error) {
    console.error('Error deleting card attachment:', error);
    showNotification('Error deleting file', 'error');
  }
}

// Carrega ferramentais por fornecedor
async function loadToolingBySupplier(supplier) {
  try {
    // Verificar se há busca ativa no input antes de carregar
    const searchInput = document.getElementById('searchInput');
    const currentSearchValue = searchInput ? searchInput.value.trim() : '';
    // Se há busca ativa no input, usar busca ao invés de carregar todos
    if (currentSearchValue && currentSearchValue.length >= 2) {
      activeSearchTerm = currentSearchValue;
      await searchTooling(currentSearchValue);
      return;
    }

    toolingData = await window.api.getToolingBySupplier(supplier);
    await ensureReplacementIdOptions();
    displayTooling(toolingData);
  } catch (error) {
    console.error('Erro ao carregar ferramentais:', error);
    showNotification('Erro ao carregar ferramentais', 'error');
  }
}

// Busca ferramentais
async function searchTooling(term) {
  try {
    const results = await window.api.searchTooling(term);
    await ensureReplacementIdOptions();
    
    // Filtrar resultados pelo supplier selecionado, se houver
    let filteredResults = results;
    if (selectedSupplier) {
      filteredResults = results.filter(item => {
        const itemSupplier = String(item.supplier || '').trim();
        const selected = String(selectedSupplier || '').trim();
        return itemSupplier === selected;
      });
    }
    
    displayTooling(filteredResults);
    updateSearchIndicators();
  } catch (error) {
    console.error('Erro na busca:', error);
  }
}

// Calcula status de vencimento
function getExpirationStatus(expirationDate) {
  if (!expirationDate) return { class: 'ok', label: 'N/A' };
  
  const now = new Date();
  const expDate = new Date(expirationDate);
  const diffTime = expDate - now;
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  
  if (diffDays < 0) {
    return { class: 'expired', label: 'Vencido' };
  } else if (diffDays <= 365) {
    return { class: 'warning', label: 'Até 1 ano' };
  } else if (diffDays <= 730) {
    return { class: 'warning', label: '1 a 2 anos' };
  } else if (diffDays <= 1825) {
    return { class: 'ok', label: '2 a 5 anos' };
  } else {
    return { class: 'ok', label: 'Mais de 5 anos' };
  }
}

// Calcula o total de vida (apenas inicial, sem revitalização)
function calculateTotalLife(item) {
  const initial = parseLocalizedNumber(item.tooling_life_qty) || 0;
  return initial;
}

// Calcula a vida restante
function calculateRemainingLife(item) {
  const total = parseLocalizedNumber(item.tooling_life_qty) || 0;
  const produced = parseLocalizedNumber(item.produced) || 0;
  const remaining = total - produced;
  return remaining >= 0 ? remaining : 0;
}

// Calcula a porcentagem usada
function calculatePercentageUsed(item) {
  const total = parseLocalizedNumber(item.tooling_life_qty) || 0;
  const produced = parseLocalizedNumber(item.produced) || 0;
  
  if (total === 0) return 0;
  
  const percentage = (produced / total) * 100;
  return Math.round(percentage * 10) / 10; // Arredonda para 1 casa decimal
}

// Recalcula lifecycle quando campos são alterados
function calculateLifecycle(cardIndex) {
  const card = document.getElementById(`card-${cardIndex}`);
  if (!card) return;
  
  const itemId = card.getAttribute('data-item-id');
  const inputs = card.querySelectorAll(`[data-id="${itemId}"]`);
  let item = {};
  
  inputs.forEach(input => {
    const field = input.getAttribute('data-field');
    item[field] = input.value;
  });
  
  // Atualiza campos calculados
  const remainingInput = card.querySelector(`[data-field="remaining_tooling_life_pcs"]`);
  const percentInput = card.querySelector(`[data-field="percent_tooling_life"]`);
  
  const toolingLife = Number(item.tooling_life_qty) || 0;
  const produced = Number(item.produced) || 0;
  const remaining = toolingLife - produced;
  const percent = toolingLife > 0 ? ((produced / toolingLife) * 100).toFixed(1) : '0.0';
  
  if (remainingInput) {
    remainingInput.value = remaining;
    item.remaining_tooling_life_pcs = remaining;
  }
  
  if (percentInput) {
    percentInput.value = percent + '%';
    item.percent_tooling_life = percent;
  }
  
  // Atualiza barra de progresso interna (aba Data)
  const progressFill = card.querySelector('.progress-fill:not([data-progress-fill])');
  const progressLabel = card.querySelector('.progress-label');
  
  if (progressFill) {
    progressFill.style.width = percent + '%';
  }
  
  if (progressLabel) {
    progressLabel.textContent = `${produced} / ${toolingLife} (${percent}%)`;
  }
  
  // Atualiza barra de progresso externa (header)
  const externalProgressFill = card.querySelector('[data-progress-fill]');
  const externalProgressPercent = card.querySelector('[data-progress-percent]');
  
  if (externalProgressFill) {
    externalProgressFill.style.width = percent + '%';
  }
  
  if (externalProgressPercent) {
    externalProgressPercent.textContent = percent + '%';
  }
  
  // Recalcula data de expiração
  calculateExpirationDate(cardIndex, item);
  
  // Salva automaticamente
  autoSaveTooling(itemId);
}

// Calcula data de vencimento baseada no annual forecast
// Calcula a data de validade
function calculateExpirationDate(cardIndex, providedItem = null) {
  const card = document.getElementById(`card-${cardIndex}`);
  if (!card) {
    console.warn(`Card #${cardIndex} não encontrado`);
    return;
  }
  
  const itemId = card.getAttribute('data-item-id');
  let item = providedItem;
  if (!item) {
    item = {};
    const inputs = card.querySelectorAll(`[data-id="${itemId}"]`);
    inputs.forEach(input => {
      const field = input.getAttribute('data-field');
      item[field] = input.value;
    });
  }
  
  // Calcula remaining first se não estiver disponível
  const toolingLife = Number(item.tooling_life_qty) || 0;
  const produced = Number(item.produced) || 0;
  const remaining = toolingLife - produced;
  
  const expirationInput = card.querySelector(`[data-field="expiration_date"]`);
  if (!expirationInput) {
    console.error(`Input expiration_date não encontrado no card ${cardIndex}`);
    return;
  }
  
  const forecast = Number(item.annual_volume_forecast) || 0;
  const productionDateValue = item.date_remaining_tooling_life || '';
  
  const formattedDate = calculateExpirationFromFormula({
    remaining,
    forecast,
    productionDate: productionDateValue
  });

  if (formattedDate) {
    expirationInput.value = formattedDate;
    item.expiration_date = formattedDate;
    console.log(`✓ Expiration atualizado no input: remaining=${remaining}, forecast=${forecast}, productionDate=${productionDateValue}, date=${formattedDate}`);
  } else {
    expirationInput.value = '';
    item.expiration_date = '';
    console.log(`Não foi possível calcular expiration: remaining=${remaining}, forecast=${forecast}, productionDate=${productionDateValue}`);
  }
  
  // Salva automaticamente
  autoSaveTooling(itemId);
}

function handleProducedChange(cardIndex) {
  calculateLifecycle(cardIndex);
  const card = document.getElementById(`card-${cardIndex}`);
  if (card) {
    triggerDateReminder(card, 'date_remaining_tooling_life');
  }
}

function handleForecastChange(cardIndex) {
  calculateExpirationDate(cardIndex);
  const card = document.getElementById(`card-${cardIndex}`);
  if (card) {
    triggerDateReminder(card, 'date_annual_volume');
  }
}

function handleStatusSelectChange(cardIndex, itemId, selectEl) {
  if (selectEl) {
    const value = (selectEl.value || '').trim();
    updateCardHeaderStatus(cardIndex, value);
    updateCardStatusAttribute(cardIndex, value);
  }
  autoSaveTooling(itemId);
}

function updateCardHeaderStatus(cardIndex, statusValue) {
  const card = document.getElementById(`card-${cardIndex}`);
  if (!card) {
    return;
  }
  const display = card.querySelector('[data-card-status]');
  if (display) {
    display.textContent = statusValue || 'N/A';
  }
}

function updateCardStatusAttribute(cardIndex, statusValue) {
  const card = document.getElementById(`card-${cardIndex}`);
  if (!card) {
    return;
  }
  const normalized = (statusValue || '').toString().trim().toLowerCase();
  card.dataset.status = normalized;
  card.classList.toggle('is-obsolete', normalized === 'obsolete');
  syncObsoleteLinkVisibility(card, normalized === 'obsolete');
}

function syncObsoleteLinkVisibility(card, isObsolete) {
  const toggleElements = [
    card.querySelector('[data-obsolete-link]'),
    card.querySelector('[data-replacement-chip]')
  ];

  toggleElements.forEach((element) => {
    if (!element) {
      return;
    }
    if (isObsolete) {
      element.removeAttribute('hidden');
      element.setAttribute('aria-hidden', 'false');
    } else {
      element.setAttribute('hidden', 'true');
      element.setAttribute('aria-hidden', 'true');
    }
  });

  if (!isObsolete) {
    closeReplacementPicker(card);
  }
}

function handleReplacementLinkInput(cardIndex, inputEl) {
  if (!inputEl) {
    return;
  }
  const sanitizedValue = sanitizeReplacementId(inputEl.value);
  inputEl.value = sanitizedValue;
  syncReplacementLinkControls(cardIndex, sanitizedValue);
}

function handleReplacementLinkChange(cardIndex, itemId, inputEl) {
  handleReplacementLinkInput(cardIndex, inputEl);
  autoSaveTooling(itemId);
}

function syncReplacementLinkControls(cardIndex, replacementId) {
  const card = document.getElementById(`card-${cardIndex}`);
  if (!card) {
    return;
  }
  const sanitizedValue = sanitizeReplacementId(replacementId);
  card.dataset.replacementId = sanitizedValue;

  const chip = card.querySelector('[data-replacement-chip]');
  if (chip) {
    chip.classList.toggle('has-link', sanitizedValue !== '');
    const chipButton = chip.querySelector('[data-replacement-chip-button]');
    const chipLabel = chip.querySelector('[data-replacement-chip-label]');
    if (chipButton) {
      chipButton.disabled = sanitizedValue === '';
      chipButton.dataset.targetId = sanitizedValue;
    }
    if (chipLabel) {
      chipLabel.textContent = sanitizedValue ? `#${sanitizedValue}` : 'Link new card';
    }
  }

  const editor = card.querySelector('[data-obsolete-link]');
  if (editor) {
    const openButton = editor.querySelector('[data-replacement-open-btn]');
    if (openButton) {
      openButton.disabled = sanitizedValue === '';
      openButton.dataset.targetId = sanitizedValue;
    }
  }

  const hiddenInput = card.querySelector('[data-field="replacement_tooling_id"]');
  if (hiddenInput && hiddenInput.value !== sanitizedValue) {
    hiddenInput.value = sanitizedValue;
  }

  syncReplacementPickerLabel(card, sanitizedValue);
}

function syncReplacementPickerLabel(card, replacementId) {
  const trigger = card.querySelector('[data-replacement-picker-button]');
  const label = card.querySelector('[data-replacement-picker-label]');
  if (!trigger || !label) {
    return;
  }
  const sanitizedValue = sanitizeReplacementId(replacementId);
  if (!sanitizedValue) {
    label.textContent = DEFAULT_REPLACEMENT_PICKER_LABEL;
    trigger.classList.remove('has-selection');
    return;
  }
  const optionLabel = getReplacementOptionLabelById(sanitizedValue) || `${sanitizedValue}`;
  label.textContent = optionLabel;
  trigger.classList.add('has-selection');
}

async function toggleReplacementPicker(event, cardIndex) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }
  const card = document.getElementById(`card-${cardIndex}`);
  if (!card) {
    return;
  }
  const panel = card.querySelector('[data-replacement-picker-panel]');
  if (!panel) {
    return;
  }
  const isHidden = panel.hasAttribute('hidden');
  closeAllReplacementPickers(card);
  if (isHidden) {
    await rebuildReplacementPickerOptions(cardIndex);
    panel.removeAttribute('hidden');
    const searchInput = panel.querySelector('[data-replacement-picker-search]');
    if (searchInput) {
      searchInput.value = '';
      handleReplacementPickerSearch(cardIndex, '');
      setTimeout(() => searchInput.focus(), 20);
    }
  }
}

function closeReplacementPicker(card) {
  if (!card) {
    return;
  }
  const panel = card.querySelector('[data-replacement-picker-panel]');
  if (panel && !panel.hasAttribute('hidden')) {
    panel.setAttribute('hidden', 'true');
  }
}

function closeAllReplacementPickers(exceptCard = null) {
  document.querySelectorAll('[data-replacement-picker-panel]').forEach((panel) => {
    const card = panel.closest('.tooling-card');
    if (exceptCard && card === exceptCard) {
      return;
    }
    panel.setAttribute('hidden', 'true');
  });
}

function showReplacementTimelineLoading() {
  const { loading, list, empty } = replacementTimelineElements;
  if (loading) {
    loading.style.display = 'block';
  }
  if (list) {
    list.innerHTML = '';
    list.style.display = 'none';
  }
  if (empty) {
    empty.style.display = 'none';
  }
}

async function openReplacementTimelineForCard(cardElement) {
  if (!cardElement) {
    return;
  }
  const startId = cardElement.dataset.itemId;
  await openReplacementTimelineOverlay(startId);
}

async function openReplacementTimelineOverlay(startId) {
  const { overlay, title } = replacementTimelineElements;
  if (!overlay) {
    return;
  }

  const normalizedId = sanitizeReplacementId(startId);
  if (!normalizedId) {
    showNotification('Não foi possível identificar o ferramental selecionado.', 'warning');
    return;
  }

  currentTimelineRootId = normalizedId;
  overlay.classList.add('active');
  if (title) {
    title.textContent = `Linked tooling from #${normalizedId}`;
  }
  showReplacementTimelineLoading();

  try {
    const chain = await buildReplacementTimeline(normalizedId);
    renderReplacementTimeline(chain);
  } catch (error) {
    console.error('Erro ao montar linha temporal de substituições:', error);
    showNotification('Não foi possível carregar a linha temporal de substituições.', 'error');
    renderReplacementTimeline([]);
  }
}

function closeReplacementTimelineOverlay() {
  const { overlay } = replacementTimelineElements;
  currentTimelineRootId = null;
  if (overlay) {
    overlay.classList.remove('active');
  }
}

function renderReplacementTimeline(chain = []) {
  const { list, empty, loading } = replacementTimelineElements;
  if (!list || !empty) {
    return;
  }

  if (loading) {
    loading.style.display = 'none';
  }

  if (!Array.isArray(chain) || chain.length === 0) {
    list.innerHTML = '';
    list.style.display = 'none';
    empty.style.display = 'block';
    return;
  }

  empty.style.display = 'none';
  list.style.display = 'flex';

  list.innerHTML = chain.map((record, index) => {
    // Status based on position: last item is Active, others are Obsolete
    const isLastInChain = index === chain.length - 1;
    const badgeClass = isLastInChain ? 'timeline-status-active' : 'timeline-status-obsolete';
    const label = isLastInChain ? 'Active' : 'Obsolete';
    
    const pn = escapeHtml(record?.pn || 'N/A');
    const description = escapeHtml(record?.tool_description || 'No description available.');
    const supplier = escapeHtml(record?.supplier || '—');
    const toolingId = escapeHtml(String(record?.id || 'N/A'));
    const isCurrent = currentTimelineRootId && String(record?.id) === String(currentTimelineRootId);
    const itemClasses = ['timeline-item'];
    if (isCurrent) {
      itemClasses.push('timeline-item-current');
    }

    return `
      <div class="${itemClasses.join(' ')}" draggable="true" data-record-id="${toolingId}">
        <span class="timeline-item-dot"></span>
        <div class="timeline-item-content">
          <div class="timeline-item-header">
            <span class="timeline-item-id">#${toolingId}</span>
            <span class="timeline-status-badge ${badgeClass}">${label}</span>
          </div>
          <div class="timeline-meta-row">
            <span><span class="timeline-meta-label">PN:</span> ${pn}</span>
            <span><span class="timeline-meta-label">Supplier:</span> ${supplier}</span>
          </div>
          <p class="timeline-item-description">${description}</p>
        </div>
        <button class="timeline-item-action" type="button" onclick="handleTimelineCardNavigate('${toolingId}')" title="Open card">
          <i class="ph ph-arrow-square-out"></i>
        </button>
      </div>
    `;
  }).join('');

  // Attach drag event listeners
  list.querySelectorAll('.timeline-item').forEach(item => {
    item.addEventListener('dragstart', handleTimelineDragStart);
    item.addEventListener('dragover', handleTimelineDragOver);
    item.addEventListener('drop', handleTimelineDrop);
    item.addEventListener('dragend', handleTimelineDragEnd);
    item.addEventListener('dragleave', handleTimelineDragLeave);
  });
}

function getTimelineStatusMeta(statusValue) {
  const normalized = (statusValue || '').toString().trim().toLowerCase();
  if (normalized === 'obsolete') {
    return { badgeClass: 'timeline-status-obsolete', label: 'Obsolete' };
  }
  return { badgeClass: 'timeline-status-active', label: 'Active' };
}

async function buildReplacementTimeline(startId) {
  const startNormalized = sanitizeReplacementId(startId);
  if (!startNormalized) {
    return [];
  }

  const visited = new Set();
  const backward = [];

  // Busca ancestrais
  let parentLookupId = startNormalized;
  while (parentLookupId) {
    const parentRecord = await fetchParentReplacementRecord(parentLookupId);
    if (!parentRecord) {
      break;
    }
    const parentId = sanitizeReplacementId(parentRecord.id);
    if (!parentId || visited.has(parentId)) {
      break;
    }
    backward.push(parentRecord);
    visited.add(parentId);
    parentLookupId = parentId;
  }

  backward.reverse();

  // Busca descendentes (inclui o item de origem)
  const forward = [];
  let currentId = startNormalized;
  while (currentId && !visited.has(currentId)) {
    const record = await fetchToolingRecordById(currentId);
    if (!record) {
      break;
    }
    forward.push(record);
    visited.add(currentId);
    const nextId = sanitizeReplacementId(record.replacement_tooling_id);
    if (!nextId || visited.has(nextId)) {
      break;
    }
    currentId = nextId;
  }

  return [...backward, ...forward];
}

async function fetchToolingRecordById(recordId) {
  const numericId = Number(recordId);
  if (!Number.isFinite(numericId)) {
    return null;
  }

  const localRecord = (toolingData || []).find((item) => Number(item.id) === numericId);
  if (localRecord) {
    return localRecord;
  }

  if (typeof window.api.getToolingById !== 'function') {
    return null;
  }

  try {
    return await window.api.getToolingById(numericId);
  } catch (error) {
    console.error('Erro ao buscar ferramental por ID:', error);
    return null;
  }
}

async function fetchParentReplacementRecord(childId) {
  const normalizedChild = sanitizeReplacementId(childId);
  if (!normalizedChild) {
    return null;
  }

  const localRecord = (toolingData || []).find((item) => sanitizeReplacementId(item?.replacement_tooling_id) === normalizedChild);
  if (localRecord) {
    return localRecord;
  }

  if (typeof window.api.getToolingByReplacementId !== 'function') {
    return null;
  }

  try {
    const parents = await window.api.getToolingByReplacementId(normalizedChild);
    if (Array.isArray(parents) && parents.length > 0) {
      return parents[0];
    }
    return null;
  } catch (error) {
    console.error('Erro ao buscar ferramental predecessor:', error);
    return null;
  }
}

function handleTimelineCardNavigate(itemId) {
  if (!itemId) {
    return;
  }
  closeReplacementTimelineOverlay();
  navigateToLinkedCard(itemId);
}

function updateCardUIAfterReorder(itemId, newStatus, newReplacementId) {
  const card = document.querySelector(`.tooling-card[data-item-id="${itemId}"]`);
  if (!card) {
    console.log(`[UI Update] Card #${itemId} não encontrado na UI`);
    return;
  }
  
  // Update status field
  const statusField = card.querySelector('[data-card-status]');
  if (statusField) {
    statusField.textContent = newStatus || 'N/A';
    console.log(`[UI Update] Card #${itemId} status → '${newStatus}'`);
  }
  
  // Update replacement link hidden input
  const replacementInput = card.querySelector('[data-field="replacement_tooling_id"]');
  if (replacementInput) {
    replacementInput.value = newReplacementId || '';
    console.log(`[UI Update] Card #${itemId} replacement_tooling_id → ${newReplacementId || 'null'}`);
  }
  
  // Update link controls
  const cardIndex = parseInt(card.id.replace('card-', ''), 10);
  if (!isNaN(cardIndex)) {
    syncReplacementLinkControls(cardIndex, newReplacementId || '');
  }
  
  // Update global toolingData
  const dataIndex = toolingData.findIndex(item => String(item.id) === String(itemId));
  if (dataIndex !== -1) {
    toolingData[dataIndex].status = newStatus;
    toolingData[dataIndex].replacement_tooling_id = newReplacementId;
    console.log(`[UI Update] toolingData[${dataIndex}] atualizado`);
  }
}

let timelineDraggedElement = null;

function handleTimelineDragStart(event) {
  timelineDraggedElement = event.currentTarget;
  event.currentTarget.classList.add('timeline-dragging');
  event.dataTransfer.effectAllowed = 'move';
  event.dataTransfer.setData('text/plain', event.currentTarget.dataset.recordId);
}

function handleTimelineDragOver(event) {
  if (event.preventDefault) {
    event.preventDefault();
  }
  event.dataTransfer.dropEffect = 'move';
  
  const targetItem = event.currentTarget;
  if (timelineDraggedElement && timelineDraggedElement !== targetItem) {
    // Remove all existing indicators
    document.querySelectorAll('.timeline-item').forEach(item => {
      item.classList.remove('timeline-drop-before', 'timeline-drop-after');
    });
    
    // Determine if we should insert before or after
    const list = targetItem.parentNode;
    const allItems = Array.from(list.querySelectorAll('.timeline-item'));
    const draggedIndex = allItems.indexOf(timelineDraggedElement);
    const targetIndex = allItems.indexOf(targetItem);
    
    if (draggedIndex < targetIndex) {
      targetItem.classList.add('timeline-drop-after');
    } else {
      targetItem.classList.add('timeline-drop-before');
    }
  }
  return false;
}

function handleTimelineDragLeave(event) {
  // Only remove if leaving the item entirely
  if (!event.currentTarget.contains(event.relatedTarget)) {
    event.currentTarget.classList.remove('timeline-drop-before', 'timeline-drop-after');
  }
}

function handleTimelineDrop(event) {
  if (event.stopPropagation) {
    event.stopPropagation();
  }
  if (event.preventDefault) {
    event.preventDefault();
  }
  
  const targetItem = event.currentTarget;
  targetItem.classList.remove('timeline-drop-before', 'timeline-drop-after');
  
  if (timelineDraggedElement && timelineDraggedElement !== targetItem) {
    const list = targetItem.parentNode;
    const allItems = Array.from(list.querySelectorAll('.timeline-item'));
    const draggedIndex = allItems.indexOf(timelineDraggedElement);
    const targetIndex = allItems.indexOf(targetItem);
    
    if (draggedIndex < targetIndex) {
      list.insertBefore(timelineDraggedElement, targetItem.nextSibling);
    } else {
      list.insertBefore(timelineDraggedElement, targetItem);
    }
    
    updateReplacementChainAfterReorder();
  }
  
  return false;
}

function handleTimelineDragEnd(event) {
  event.currentTarget.classList.remove('timeline-dragging');
  document.querySelectorAll('.timeline-item').forEach(item => {
    item.classList.remove('timeline-drop-before', 'timeline-drop-after');
  });
  timelineDraggedElement = null;
}

async function updateReplacementChainAfterReorder() {
  const list = replacementTimelineElements.list;
  if (!list) {
    console.error('[Timeline] Lista não encontrada');
    return;
  }
  
  const items = Array.from(list.querySelectorAll('.timeline-item'));
  const orderedIds = items.map(item => item.dataset.recordId).filter(Boolean);
  
  console.log('[Timeline] IDs reorganizados:', orderedIds);
  
  if (orderedIds.length === 0) {
    console.warn('[Timeline] Nenhum ID encontrado');
    return;
  }
  
  isReorderingTimeline = true;
  
  try {
    // Update replacement links and status based on new order
    for (let i = 0; i < orderedIds.length; i++) {
      const currentId = orderedIds[i];
      const isLast = i === orderedIds.length - 1;
      
      if (isLast) {
        // Last item: clear status and link
        await window.api.updateTooling(Number(currentId), {
          replacement_tooling_id: null,
          status: ''
        });
        console.log(`[Timeline] Último item #${currentId}: status='', replacement_tooling_id=null`);
      } else {
        // All others: Obsolete + link to next
        await window.api.updateTooling(Number(currentId), {
          replacement_tooling_id: Number(orderedIds[i + 1]),
          status: 'Obsolete'
        });
        console.log(`[Timeline] Item #${currentId}: status='Obsolete', replacement_tooling_id=${orderedIds[i + 1]}`);
      }
    }
    
    showNotification('Cadeia reorganizada e salva com sucesso.', 'success');
    
    // Update cards UI immediately
    console.log('[Timeline] Atualizando UI dos cards...');
    for (let i = 0; i < orderedIds.length; i++) {
      const currentId = orderedIds[i];
      const isLast = i === orderedIds.length - 1;
      const nextId = isLast ? null : Number(orderedIds[i + 1]);
      const newStatus = isLast ? '' : 'Obsolete';
      
      updateCardUIAfterReorder(currentId, newStatus, nextId);
    }
    
    // Reload data to refresh all cards
    console.log('[Timeline] Recarregando dados...');
    if (typeof loadToolingData === 'function') {
      await loadToolingData();
    }
    
    // Wait a bit for data to refresh
    await new Promise(resolve => setTimeout(resolve, 300));
    
    // Rebuild timeline with updated data
    console.log('[Timeline] Reconstruindo timeline...');
    const rootId = orderedIds[0];
    const chain = await buildReplacementTimeline(rootId);
    console.log('[Timeline] Nova cadeia:', chain.map(c => ({ id: c.id, status: c.status })));
    renderReplacementTimeline(chain);
    
  } catch (error) {
    console.error('Erro ao atualizar cadeia de substituição:', error);
    showNotification('Erro ao atualizar a ordem da cadeia.', 'error');
  } finally {
    isReorderingTimeline = false;
  }
}

function handleReplacementPickerSearch(cardIndex, searchValue) {
  const card = document.getElementById(`card-${cardIndex}`);
  if (!card) {
    return;
  }
  const list = card.querySelector('[data-replacement-picker-list]');
  if (!list) {
    return;
  }
  const normalizedQuery = (searchValue || '').trim().toLowerCase();
  const options = list.querySelectorAll('.replacement-dropdown-option');
  options.forEach((option) => {
    const searchText = option.textContent.toLowerCase();
    const matches = !normalizedQuery || searchText.includes(normalizedQuery);
    if (matches) {
      option.removeAttribute('hidden');
    } else {
      option.setAttribute('hidden', 'true');
    }
  });
}

function handleReplacementPickerSelect(cardIndex, itemId, selectedId) {
  const card = document.getElementById(`card-${cardIndex}`);
  if (!card) {
    return;
  }
  const input = card.querySelector('[data-field="replacement_tooling_id"]');
  if (!input) {
    return;
  }
  input.value = sanitizeReplacementId(selectedId);
  handleReplacementLinkChange(cardIndex, itemId, input);
  closeReplacementPicker(card);
}

function handleReplacementLinkButtonClick(event, buttonEl) {
  if (event) {
    event.stopPropagation();
  }
  if (!buttonEl || buttonEl.disabled) {
    return;
  }
  const card = buttonEl.closest('.tooling-card');
  openReplacementTimelineForCard(card);
}

function handleReplacementLinkChipClick(event, buttonEl) {
  if (event) {
    event.stopPropagation();
  }
  if (!buttonEl || buttonEl.disabled) {
    return;
  }
  const card = buttonEl.closest('.tooling-card');
  openReplacementTimelineForCard(card);
}

async function navigateToLinkedCard(targetId) {
  const normalizedId = sanitizeReplacementId(targetId);
  console.log('navigateToLinkedCard:', { targetId, normalizedId });
  
  if (!normalizedId) {
    showNotification('Informe um ID de substituição válido.', 'error');
    return;
  }

  let targetCard = document.querySelector(`.tooling-card[data-item-id="${normalizedId}"]`);
  console.log('targetCard initial:', targetCard);
  
  if (!targetCard) {
    console.log('Card não encontrado, tentando carregar...');
    const cardLoaded = await ensureCardLoadedById(normalizedId);
    console.log('cardLoaded:', cardLoaded);
    
    if (!cardLoaded) {
      showNotification(`O card #${normalizedId} não foi encontrado.`, 'warning');
      return;
    }
    
    // Aguarda um pouco para o DOM ser atualizado
    await new Promise(resolve => setTimeout(resolve, 200));
    targetCard = document.querySelector(`.tooling-card[data-item-id="${normalizedId}"]`);
    console.log('targetCard after load:', targetCard);
  }

  if (!targetCard) {
    showNotification(`Não foi possível exibir o card #${normalizedId}.`, 'warning');
    return;
  }

  const cardIndex = parseInt(targetCard.id.replace('card-', ''), 10);
  console.log('cardIndex:', cardIndex);
  
  // Expand the card if not already expanded
  if (!targetCard.classList.contains('expanded')) {
    // Save any previously expanded cards
    const expandedCards = document.querySelectorAll('.tooling-card.expanded');
    for (const expandedCard of expandedCards) {
      const itemId = expandedCard.getAttribute('data-item-id');
      if (itemId) {
        await saveToolingQuietly(itemId);
      }
    }
    
    // Expand the target card
    targetCard.classList.add('expanded');
    
    // Add blur effect to other cards
    const allCards = document.querySelectorAll('.tooling-card');
    allCards.forEach(c => c.classList.add('has-expanded-sibling'));
    
    // Calculate expiration and load attachments
    calculateExpirationDate(cardIndex);
    const itemId = targetCard.getAttribute('data-item-id');
    if (itemId) {
      await loadCardAttachments(itemId);
    }
  }
  
  ensureCardVisible(targetCard);
  flashCardHighlight(targetCard);
}

async function ensureCardLoadedById(cardId) {
  if (typeof window.api.getToolingById !== 'function') {
    return false;
  }

  const numericId = Number(cardId);
  if (!Number.isFinite(numericId)) {
    return false;
  }

  try {
    const record = await window.api.getToolingById(numericId);
    if (!record) {
      return false;
    }
    const supplierName = String(record.supplier || '').trim();
    if (!supplierName) {
      return false;
    }
    await handleSupplierSelection(supplierName, { forceReload: true });
    return true;
  } catch (error) {
    console.error('Erro ao localizar card vinculado:', error);
    showNotification('Erro ao localizar o card vinculado.', 'error');
    return false;
  }
}

function flashCardHighlight(card) {
  if (!card) {
    return;
  }
  card.classList.add('card-highlight');
  setTimeout(() => {
    card.classList.remove('card-highlight');
  }, 1200);
}

function triggerDateReminder(card, fieldName) {
  const dateInput = card.querySelector(`[data-field="${fieldName}"]`);
  if (!dateInput) {
    return;
  }

  showDateHighlight(dateInput);
}

function showDateHighlight(input) {
  input.classList.add('date-highlight');

  // Cria tooltip se não existir
  let tooltip = input.parentElement.querySelector('.date-reminder-tooltip');
  if (!tooltip) {
    tooltip = document.createElement('div');
    tooltip.className = 'date-reminder-tooltip';
    tooltip.textContent = 'Please update this date';
    input.parentElement.style.position = 'relative';
    input.parentElement.appendChild(tooltip);
  }
  tooltip.classList.add('active');

  if (dateReminderTimers.has(input)) {
    clearTimeout(dateReminderTimers.get(input));
  }

  const timeout = setTimeout(() => {
    input.classList.remove('date-highlight');
    tooltip.classList.remove('active');
    dateReminderTimers.delete(input);
  }, 8000);

  dateReminderTimers.set(input, timeout);
}

// Debounce para salvar automaticamente
let autoSaveTimeouts = {};

function autoSaveTooling(id, immediate = false) {
  // Skip autosave during timeline reordering
  if (isReorderingTimeline) {
    console.log(`[AutoSave] Skipped for #${id} (reordering in progress)`);
    return;
  }
  
  // Cancela timeout anterior se existir
  if (autoSaveTimeouts[id]) {
    clearTimeout(autoSaveTimeouts[id]);
  }
  
  // Se immediate for true, salva imediatamente sem notificação
  if (immediate) {
    saveToolingQuietly(id);
    return;
  }
  
  // Cria novo timeout para salvar após 1 segundo sem edição
  autoSaveTimeouts[id] = setTimeout(async () => {
    await saveToolingQuietly(id);
  }, 1000);
}

// Formata data para exibição
function formatDate(dateString) {
  if (!dateString) return 'N/A';
  try {
    const date = new Date(dateString);
    return date.toLocaleDateString('pt-BR');
  } catch {
    return dateString;
  }
}

// Formata data e hora para exibição
function formatDateTime(dateString) {
  if (!dateString) return 'N/A';
  try {
    let normalized = dateString.trim();
    if (!normalized.includes('T')) {
      normalized = normalized.replace(' ', 'T');
    }
    if (!/[zZ]$/.test(normalized)) {
      normalized += 'Z';
    }
    let date = new Date(normalized);
    if (isNaN(date.getTime())) {
      date = new Date(dateString);
    }
    return date.toLocaleDateString('pt-BR');
  } catch {
    return dateString;
  }
}

function sanitizeReplacementId(value) {
  if (value === null || value === undefined) {
    return '';
  }
  const digitsOnly = String(value).trim().replace(/\D+/g, '');
  if (digitsOnly.length === 0) {
    return '';
  }
  const numeric = parseInt(digitsOnly, 10);
  if (!Number.isFinite(numeric) || numeric <= 0) {
    return '';
  }
  return String(numeric);
}

// Normaliza valores de data vindos do banco (strings, números, datas Excel)
function normalizeExpirationDate(rawValue, itemId) {
  if (rawValue === null || rawValue === undefined) {
    return null;
  }

  const rawString = String(rawValue).trim();
  if (rawString === '') {
    return null;
  }

  const numericValue = Number(rawString);
  const isNumeric = !Number.isNaN(numericValue) && /^-?\d+(\.\d+)?$/.test(rawString);

  if (isNumeric) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const days = Math.floor(numericValue);
    const milliseconds = Math.round((numericValue - days) * 86400000);
    const date = new Date(excelEpoch.getTime() + (days * 86400000) + milliseconds);
    if (!Number.isNaN(date.getTime())) {
      const iso = date.toISOString().split('T')[0];
      console.debug('Data normalizada (excel)', { itemId, rawValue, iso });
      return iso;
    }
  } else {
    const parsed = new Date(rawString);
    if (!Number.isNaN(parsed.getTime())) {
      const iso = parsed.toISOString().split('T')[0];
      console.debug('Data normalizada (string)', { itemId, rawValue, iso });
      return iso;
    }
  }

  console.warn('Data de expiração inválida detectada', { itemId, rawValue });
  return null;
}

// Converte string numérica localizada (pt-BR) para Number
function parseLocalizedNumber(value) {
  if (value === null || value === undefined) {
    return 0;
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }

  let s = String(value).trim();
  if (s === '') {
    return 0;
  }

  // Remove espaços comuns e não separáveis
  s = s.replace(/\s+/g, '').replace(/\u00A0/g, '');

  // Mantém apenas dígitos, separadores e sinal
  s = s.replace(/[^\d.,-]/g, '');
  if (s === '' || s === '-' || s === '+') {
    return 0;
  }

  const thousandsPattern = /^-?\d{1,3}(?:[.,]\d{3})+$/;
  if (thousandsPattern.test(s)) {
    const normalized = Number(s.replace(/[.,]/g, ''));
    return Number.isNaN(normalized) ? 0 : normalized;
  }

  const lastComma = s.lastIndexOf(',');
  const lastDot = s.lastIndexOf('.');
  const separatorIndex = Math.max(lastComma, lastDot);

  if (separatorIndex !== -1) {
    const integerPartRaw = s.slice(0, separatorIndex);
    const fractionalRaw = s.slice(separatorIndex + 1);
    const sign = integerPartRaw.startsWith('-') ? '-' : '';
    const integerPart = integerPartRaw.replace(/[^\d]/g, '') || '0';
    const fractionalPart = fractionalRaw.replace(/[^\d]/g, '');
    if (fractionalPart.length > 0) {
      const normalized = Number(`${sign}${integerPart}.${fractionalPart}`);
      if (!Number.isNaN(normalized)) {
        return normalized;
      }
    }
  }

  const fallback = Number(s.replace(/[^\d-]/g, ''));
  return Number.isNaN(fallback) ? 0 : fallback;
}

function calculateExpirationFromFormula({
  remaining,
  forecast,
  productionDate
}) {
  const baseDate = productionDate ? new Date(productionDate) : new Date();
  if (Number.isNaN(baseDate.getTime())) {
    return null;
  }

  const totalDays = forecast <= 0
    ? Math.round(remaining)
    : Math.round((remaining / forecast) * 365);

  if (Number.isNaN(totalDays)) {
    return null;
  }

  const expirationDate = new Date(baseDate);
  expirationDate.setDate(expirationDate.getDate() + totalDays);
  if (Number.isNaN(expirationDate.getTime())) {
    return null;
  }
  return expirationDate.toISOString().split('T')[0];
}

function generateConfirmationCode(length = 3) {
  const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let code = '';
  for (let i = 0; i < length; i += 1) {
    const index = Math.floor(Math.random() * letters.length);
    code += letters.charAt(index);
  }
  return code;
}

function confirmDeleteTooling(id) {
  const overlay = document.getElementById('deleteConfirmOverlay');
  const descriptionEl = document.getElementById('deleteItemDescription');
  const codeEl = document.getElementById('deleteConfirmCode');
  const inputEl = document.getElementById('deleteConfirmInput');

  if (!overlay || !descriptionEl || !codeEl || !inputEl) {
    console.error('Elementos da confirmação de exclusão não encontrados.');
    return;
  }

  const item = toolingData.find(tool => String(tool.id) === String(id));
  const descriptorParts = [];
  if (item?.pn) descriptorParts.push(item.pn);
  if (item?.tool_description) descriptorParts.push(item.tool_description);
  const descriptor = descriptorParts.join(' - ') || 'este ferramental';

  deleteConfirmState = {
    id,
    code: generateConfirmationCode(),
    descriptor
  };

  descriptionEl.textContent = descriptor;
  codeEl.textContent = deleteConfirmState.code;
  inputEl.value = '';

  overlay.classList.add('active');

  setTimeout(() => {
    inputEl.focus();
  }, 50);
}

function cancelDeleteTooling() {
  const overlay = document.getElementById('deleteConfirmOverlay');
  const inputEl = document.getElementById('deleteConfirmInput');
  if (overlay) {
    overlay.classList.remove('active');
  }
  if (inputEl) {
    inputEl.value = '';
  }
  deleteConfirmState = { id: null, code: '', descriptor: '' };
}

async function handleDeleteConfirmation() {
  const overlay = document.getElementById('deleteConfirmOverlay');
  const inputEl = document.getElementById('deleteConfirmInput');

  if (!inputEl) {
    return;
  }

  const userCode = inputEl.value.trim().toUpperCase();
  if (!deleteConfirmState.id || !deleteConfirmState.code) {
    cancelDeleteTooling();
    return;
  }

  if (userCode !== deleteConfirmState.code) {
    showNotification('Código incorreto. Tente novamente.', 'error');
    inputEl.select();
    return;
  }

  const idToDelete = deleteConfirmState.id;
  cancelDeleteTooling();
  await deleteToolingItem(idToDelete);
  if (overlay) {
    overlay.classList.remove('active');
  }
}

function populateAddToolingSuppliers() {
  const list = addToolingElements.supplierList;
  if (!list) {
    return;
  }

  list.innerHTML = '';

  const suppliers = Array.isArray(suppliersData) ? [...suppliersData] : [];
  const uniqueNames = [...new Set(
    suppliers
      .map(item => String(item?.supplier || '').trim())
      .filter(name => name.length > 0)
  )].sort((a, b) => a.localeCompare(b, 'pt-BR'));

  uniqueNames.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    list.appendChild(option);
  });
}

function openAddToolingModal() {
  const { overlay, pnInput, supplierInput, lifeInput, producedInput } = addToolingElements;
  if (!overlay) {
    return;
  }

  populateAddToolingSuppliers();

  if (pnInput) pnInput.value = '';
  if (supplierInput) {
    // Se há supplier selecionado, preenche o campo
    supplierInput.value = selectedSupplier || '';
    const defaultPlaceholder = supplierInput.getAttribute('data-default-placeholder') || supplierInput.placeholder;
    if (!supplierInput.hasAttribute('data-default-placeholder')) {
      supplierInput.setAttribute('data-default-placeholder', defaultPlaceholder);
    }
    supplierInput.placeholder = selectedSupplier
      ? `Select or type (current: ${selectedSupplier})`
      : defaultPlaceholder;
  }
  if (lifeInput) lifeInput.value = '';
  if (producedInput) producedInput.value = '';

  overlay.classList.add('active');

  setTimeout(() => {
    pnInput?.focus();
  }, 50);
}

function closeAddToolingModal() {
  const { overlay, pnInput, supplierInput, lifeInput, producedInput } = addToolingElements;
  if (overlay) overlay.classList.remove('active');
  if (pnInput) pnInput.value = '';
  if (supplierInput) {
    supplierInput.value = '';
    const defaultPlaceholder = supplierInput.getAttribute('data-default-placeholder');
    if (defaultPlaceholder) {
      supplierInput.placeholder = defaultPlaceholder;
    }
  }
  if (lifeInput) lifeInput.value = '';
  if (producedInput) producedInput.value = '';
}

async function submitAddToolingForm() {
  const { form, pnInput, supplierInput, lifeInput, producedInput } = addToolingElements;

  if (!pnInput || !supplierInput || !lifeInput || !producedInput) {
    return;
  }

  if (form && typeof form.reportValidity === 'function' && !form.reportValidity()) {
    return;
  }

  const pn = pnInput.value.trim();
  const supplier = supplierInput.value.trim();
  const toolingLife = parseLocalizedNumber(lifeInput.value);
  const produced = parseLocalizedNumber(producedInput.value);

  if (!pn) {
    showNotification('Informe o PN do ferramental.', 'error');
    pnInput.focus();
    return;
  }

  if (!supplier) {
    showNotification('Selecione ou cadastre um fornecedor.', 'error');
    supplierInput.focus();
    return;
  }

  if (Number.isNaN(toolingLife) || Number.isNaN(produced)) {
    showNotification('Preencha os valores numéricos corretamente.', 'error');
    return;
  }

  if (toolingLife < 0 || produced < 0) {
    showNotification('Valores numéricos devem ser positivos.', 'error');
    return;
  }

  try {
    const payload = {
      pn,
      supplier,
      tooling_life_qty: toolingLife,
      produced
    };
    const result = await window.api.createTooling(payload);

    if (!result || result.success !== true) {
      showNotification(result?.error || 'Não foi possível criar o ferramental.', 'error');
      return;
    }

    closeAddToolingModal();

    selectedSupplier = supplier;
    currentSupplier = supplier;

    await loadSuppliers();
    await loadAnalytics();
    await refreshReplacementIdOptions(true);
    await loadToolingBySupplier(supplier);
    await loadAttachments(supplier);

    const attachmentsContainer = document.getElementById('attachmentsContainer');
    const currentSupplierName = document.getElementById('currentSupplierName');
    if (attachmentsContainer && currentSupplierName) {
      attachmentsContainer.style.display = 'block';
      currentSupplierName.textContent = supplier;
    }

    showNotification('Ferramental criado com sucesso!');
  } catch (error) {
    console.error('Erro ao criar ferramental:', error);
    showNotification('Erro ao criar ferramental', 'error');
  }
}

async function deleteToolingItem(id) {
  try {
    const result = await window.api.deleteTooling(id);
    if (!result || result.success !== true) {
      showNotification('Não foi possível excluir o ferramental.', 'error');
      return;
    }

    showNotification('Ferramental excluído com sucesso!');
    await loadSuppliers();
    await loadAnalytics();
    await refreshReplacementIdOptions(true);
    if (selectedSupplier) {
      await loadToolingBySupplier(selectedSupplier);
    } else {
      displayTooling([]);
    }
  } catch (error) {
    console.error('Erro ao excluir ferramental:', error);
    showNotification('Erro ao excluir ferramental', 'error');
  }
}

// Exibe ferramentais na interface (cards expansíveis)
async function displayTooling(data) {
  const toolingList = document.getElementById('toolingList');
  const emptyState = document.getElementById('emptyState');

  if (!data || data.length === 0) {
    toolingList.style.display = 'none';
    emptyState.style.display = 'flex';
    return;
  }

  // Aplica filtro de expiração se estiver ativo
  let filteredData = data;
  if (expirationFilterEnabled) {
    filteredData = data.filter(item => {
      const lifeQty = parseLocalizedNumber(item.tooling_life_qty) || 0;
      const produced = parseLocalizedNumber(item.produced) || 0;
      const percentLife = item.percent_tooling_life ? parseFloat(item.percent_tooling_life) : 0;
      
      // Verifica se está expirado
      const isExpired = (percentLife >= 100.0) || (lifeQty > 0 && produced >= lifeQty);
      if (isExpired) return true;
      
      // Verifica se expira em até 2 anos
      if (item.expiration_date) {
        const now = new Date();
        const expDate = new Date(item.expiration_date);
        const daysUntilExpiration = Math.floor((expDate - now) / (1000 * 60 * 60 * 24));
        if (daysUntilExpiration > 0 && daysUntilExpiration <= 730) {
          return true;
        }
      }
      
      return false;
    });
  }

  if (filteredData.length === 0) {
    toolingList.style.display = 'none';
    emptyState.style.display = 'flex';
    return;
  }

  // Ordena automaticamente por % de progresso (maior para menor)
  const sortedData = [...filteredData].sort((a, b) => {
    const lifeA = parseLocalizedNumber(a.tooling_life_qty) || 0;
    const lifeB = parseLocalizedNumber(b.tooling_life_qty) || 0;
    const prodA = parseLocalizedNumber(a.produced) || 0;
    const prodB = parseLocalizedNumber(b.produced) || 0;
    
    const percentA = lifeA > 0 ? ((prodA / lifeA) * 100) : 0;
    const percentB = lifeB > 0 ? ((prodB / lifeB) * 100) : 0;
    
    return percentB - percentA;
  });

  // Carrega contagem de anexos para cada item
  const attachmentCounts = {};
  for (const item of sortedData) {
    try {
      const attachments = await window.api.getAttachments(selectedSupplier, item.id);
      attachmentCounts[item.id] = attachments.length;
    } catch (error) {
      console.error('Erro ao carregar anexos do item', item.id, error);
      attachmentCounts[item.id] = 0;
    }
  }

  // Atualiza o estado global
  toolingData = sortedData;

  // Pre-compute chain membership for all items
  const chainMembership = new Map();
  for (const item of sortedData) {
    const itemId = String(item.id);
    const hasOutgoingLink = sanitizeReplacementId(item.replacement_tooling_id) !== '';
    
    // Check if any item in current data points to this item
    const hasIncomingLinkInData = sortedData.some(t => 
      sanitizeReplacementId(t.replacement_tooling_id) === itemId
    );
    
    // Also check database for incoming links (items from other suppliers)
    let hasIncomingLinkInDB = false;
    try {
      const dbResults = await window.api.getToolingByReplacementId(Number(itemId));
      hasIncomingLinkInDB = dbResults && dbResults.length > 0;
    } catch (error) {
      console.error('Error checking chain membership for', itemId, error);
    }
    
    chainMembership.set(itemId, hasOutgoingLink || hasIncomingLinkInData || hasIncomingLinkInDB);
  }

  emptyState.style.display = 'none';
  toolingList.style.display = 'flex';
  
  toolingList.innerHTML = sortedData.map((item, index) => {
    // Calcula expiration_date se não existir
    const toolingLife = Number(item.tooling_life_qty) || 0;
    const produced = Number(item.produced) || 0;
    const remaining = toolingLife - produced;
    const forecast = Number(item.annual_volume_forecast) || 0;
    const hasForecast = String(item.annual_volume_forecast ?? '').trim() !== '';
    
    let expirationDateValue = normalizeExpirationDate(item.expiration_date, item.id);

    if (!expirationDateValue) {
      const productionDateValue = item.date_remaining_tooling_life || '';
      const calculatedExpiration = calculateExpirationFromFormula({
        remaining,
        forecast,
        productionDate: productionDateValue
      });
      if (calculatedExpiration) {
        expirationDateValue = calculatedExpiration;
        console.debug('Data de expiração recalculada', {
          itemId: item.id,
          forecast,
          remaining,
          productionDate: productionDateValue,
          expirationDateValue
        });
      }
    }

    const expirationInputValue = expirationDateValue || '';
    const expirationDisplay = formatDate(expirationInputValue);

    console.log(`Card ${index} renderizado:`, {
      itemId: item.id,
      expirationInputValue,
      expirationDisplay,
      remaining,
      forecast
    });
    
    const percentUsedValue = toolingLife > 0 ? (produced / toolingLife) * 100 : 0;
    const percentUsed = toolingLife > 0 ? percentUsedValue.toFixed(1) : '0';
    const remainingQty = remaining;
    const lifecycleProgressPercent = Math.min(Math.max(percentUsedValue, 0), 100);
    const lifecycleProgressLabel = `${produced.toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 0})} / ${toolingLife.toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 0})}`;
    const amountBrlValue = (() => {
      const raw = item.amount_brl;
      if (raw === null || raw === undefined) {
        return '';
      }
      const trimmed = String(raw).trim();
      if (trimmed === '') {
        return '';
      }
      const parsed = parseLocalizedNumber(trimmed);
      return Number.isNaN(parsed) ? '' : parsed;
    })();
    const lastUpdateDisplay = formatDateTime(item.last_update);
    const attachmentCount = attachmentCounts[item.id] || 0;
    const statusOptionsMarkup = buildStatusOptionsMarkup(item.status);
    const normalizedStatus = (item.status || '').toString().trim().toLowerCase();
    const isObsolete = normalizedStatus === 'obsolete';
    const replacementIdValue = sanitizeReplacementId(item.replacement_tooling_id);
    const hasReplacementLink = replacementIdValue !== '';
    const replacementPickerOptionsMarkup = buildReplacementPickerOptionsMarkup(item, index);
    const replacementPickerLabel = escapeHtml(
      hasReplacementLink
        ? (getReplacementOptionLabelById(replacementIdValue) || `${replacementIdValue}`)
        : DEFAULT_REPLACEMENT_PICKER_LABEL
    );
    const replacementChipVisibilityAttr = isObsolete ? 'aria-hidden="false"' : 'hidden aria-hidden="true"';
    const replacementEditorVisibilityAttr = isObsolete ? 'aria-hidden="false"' : 'hidden aria-hidden="true"';
    
    // Check if item is part of a replacement chain using pre-computed map
    const isInChain = chainMembership.get(String(item.id)) || false;
    
    // Calcula status de vencimento para ícone
    const expirationStatus = getExpirationStatus(expirationInputValue);
    let statusIconHtml = '';
    let statusIconClass = '';
    
    // Verifica se está expirado por percentual de vida
    const isExpiredByPercent = percentUsedValue >= 100;
    
    // Mostra ícone de status independente de estar em chain
    if (isObsolete) {
      statusIconHtml = '<i class="ph ph-fill ph-archive-box"></i>';
      statusIconClass = 'status-icon-obsolete';
    } else if (isExpiredByPercent || expirationStatus.class === 'expired') {
      statusIconHtml = '<i class="ph ph-fill ph-warning-circle"></i>';
      statusIconClass = 'status-icon-expired';
    } else if (expirationStatus.class === 'warning') {
      statusIconHtml = '<i class="ph ph-fill ph-warning"></i>';
      statusIconClass = 'status-icon-warning';
    }
    
    return `
      <div class="tooling-card" id="card-${index}" data-item-id="${item.id}" data-status="${normalizedStatus}" data-replacement-id="${hasReplacementLink ? replacementIdValue : ''}">
        <div class="tooling-card-header" onclick="toggleCard(${index})">
          <div class="tooling-card-header-top">
            ${statusIconHtml ? `<div class="card-status-icon ${statusIconClass}" title="${isObsolete ? 'Obsolete' : expirationStatus.label}">${statusIconHtml}</div>` : ''}
            <div class="tooling-card-primary">
              <div class="tooling-info-item tooling-info-pn">
                <div class="tooling-info-stack">
                  <div class="tooling-info-pn-row">
                    <span class="tooling-info-id">#${item.id}</span>
                    <span class="tooling-info-separator">•</span>
                    <span class="tooling-info-value highlight">${item.pn || 'N/A'}</span>
                  </div>
                </div>
              </div>
            </div>
            <div class="tooling-card-top-meta">
              <div class="tooling-info-stack tooling-info-last-update">
                <span class="tooling-info-label">Last Update</span>
                <span class="tooling-info-value">${lastUpdateDisplay}</span>
              </div>
              ${isInChain ? `
              <button type="button" class="tooling-chain-indicator" title="View replacement chain" onclick="event.stopPropagation(); openReplacementTimelineForCard(document.getElementById('card-${index}'))">
                <i class="ph ph-git-branch"></i>
              </button>
              ` : ''}
              ${attachmentCount > 0 ? `
              <div class="tooling-attachment-count">
                <i class="ph ph-paperclip"></i>
                <span>${attachmentCount}</span>
              </div>
              ` : ''}
              <div class="tooling-card-expand">
                <i class="ph ph-caret-down"></i>
              </div>
            </div>
          </div>
          <div class="tooling-card-header-divider"></div>
          <div class="tooling-card-main-info">
            <div class="tooling-info-item tooling-info-description">
              <span class="tooling-info-label">Description</span>
              <span class="tooling-info-value">${item.tool_description || 'N/A'}</span>
            </div>
            <div class="tooling-info-item card-collapsible">
              <span class="tooling-info-label">Expiration</span>
              <span class="tooling-info-value">${expirationDisplay}</span>
            </div>
            <div class="tooling-info-item card-collapsible">
              <span class="tooling-info-label">Status</span>
              <span class="tooling-info-value" data-card-status>${item.status || 'N/A'}</span>
            </div>
            <div class="tooling-info-item card-collapsible">
              <span class="tooling-info-label">Steps</span>
              <span class="tooling-info-value">${item.steps || 'N/A'}</span>
            </div>
            <div class="tooling-info-item tooling-info-progress">
              <div class="progress-bar">
                <div class="progress-fill" style="width: ${percentUsed}%" data-progress-fill></div>
              </div>
              <span class="progress-percentage" data-progress-percent>${percentUsed}%</span>
            </div>
          </div>
        </div>
        <div class="tooling-card-body">
          <!-- Abas do Card -->
          <div class="card-tabs">
            <button class="card-tab active" onclick="switchCardTab(${index}, 'data')">
              <i class="ph ph-database"></i>
              <span>Data</span>
            </button>
            <button class="card-tab" onclick="switchCardTab(${index}, 'documentation')">
              <i class="ph ph-folder"></i>
              <span>Documentation</span>
            </button>
            <button class="card-tab" onclick="switchCardTab(${index}, 'comments')">
              <i class="ph ph-chat-circle-text"></i>
              <span>Comments</span>
            </button>
            <button class="card-tab" onclick="switchCardTab(${index}, 'todos')">
              <i class="ph ph-check-square"></i>
              <span>Todos</span>
            </button>
            <button class="card-tab" onclick="switchCardTab(${index}, 'attachments')">
              <i class="ph ph-paperclip"></i>
              <span>Attachments</span>
            </button>
          </div>

          <!-- Aba Data -->
          <div class="card-tab-content active" data-tab="data">
          <div class="details-carousel-wrapper">
            <button class="carousel-nav carousel-nav-prev" onclick="navigateCarousel(${index}, 'prev')" aria-label="Previous">
              <i class="ph ph-caret-left"></i>
            </button>
            <div class="tooling-details-carousel">
              <div class="tooling-details-grid" data-carousel-track>
                <!-- Column 1: Lifecycle -->
                <div class="detail-group detail-grid">
                  <div class="detail-group-title lifecycle-title">
                    <span>Lifecycle</span>
                    <div class="lifecycle-progress">
                      <div class="lifecycle-progress-bar">
                        <div class="lifecycle-progress-fill" style="width: ${lifecycleProgressPercent}%"></div>
                      </div>
                    </div>
                  </div>
                  <div class="detail-item detail-item-full">
                    <span class="detail-label">Tooling Life (Qty)</span>
                    <input type="number" class="detail-input" value="${toolingLife}" data-field="tooling_life_qty" data-id="${item.id}" onchange="calculateLifecycle(${index})" oninput="calculateLifecycle(${index})" min="0" step="1">
                  </div>
                  <div class="detail-item detail-pair">
                    <div class="detail-item">
                      <span class="detail-label">Produced</span>
                      <input type="number" class="detail-input" value="${produced}" data-field="produced" data-id="${item.id}" onchange="handleProducedChange(${index})" oninput="handleProducedChange(${index})" min="0" step="1">
                    </div>
                    <div class="detail-item">
                      <span class="detail-label">
                        Production Date
                        <i class="ph ph-info tooltip-icon" title="Whenever Produced changes, update this date to keep the timeline accurate." role="button" tabindex="0" onclick="openProductionInfoModal(event)" onkeydown="handleProductionInfoIconKey(event)"></i>
                      </span>
                      <input type="date" class="detail-input" value="${item.date_remaining_tooling_life || ''}" data-field="date_remaining_tooling_life" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                    </div>
                  </div>
                  <div class="detail-item">
                    <span class="detail-label">
                      Remaining
                      <i class="ph ph-info tooltip-icon" title="Formula: uses \"Tooling Life (Qty)\" minus \"Produced\"."></i>
                    </span>
                    <input type="text" class="detail-input calculated" value="${remainingQty.toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 0})}" data-field="remaining_tooling_life_pcs" data-id="${item.id}" readonly>
                  </div>
                  <div class="detail-item">
                    <span class="detail-label">
                      % Used
                      <i class="ph ph-info tooltip-icon" title="Formula: (\"Produced\" ÷ \"Tooling Life (Qty)\") × 100."></i>
                    </span>
                    <input type="text" class="detail-input calculated" value="${percentUsed}%" data-field="percent_tooling_life" data-id="${item.id}" readonly>
                  </div>
                  <div class="detail-item detail-pair">
                    <div class="detail-item">
                      <span class="detail-label">Annual Forecast</span>
                      <input type="number" class="detail-input" value="${hasForecast ? forecast : ''}" data-field="annual_volume_forecast" data-id="${item.id}" onchange="handleForecastChange(${index})" min="0" step="1">
                    </div>
                    <div class="detail-item">
                      <span class="detail-label">Forecast Date</span>
                      <input type="date" class="detail-input" value="${item.date_annual_volume || ''}" data-field="date_annual_volume" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                    </div>
                  </div>
                  <div class="detail-item detail-item-full">
                    <span class="detail-label">
                      Expiration (Calculated)
                      <i class="ph ph-info tooltip-icon" title="Formula: today's date + (\"Remaining\" ÷ \"Annual Forecast\") years." role="button" tabindex="0" onclick="openExpirationInfoModal(event)" onkeydown="handleExpirationInfoIconKey(event)"></i>
                    </span>
                    <input type="date" class="detail-input calculated" value="${expirationInputValue}" data-field="expiration_date" data-id="${item.id}" readonly>
                  </div>
                </div>
                
                <!-- Column 2: Identification -->
                <div class="detail-group detail-grid">
                  <div class="detail-group-title">Identification</div>
              <div class="detail-item detail-item-full">
                <span class="detail-label">PN</span>
                <input type="text" class="detail-input" value="${item.pn || ''}" data-field="pn" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item detail-item-full">
                <span class="detail-label">PN Description</span>
                <input type="text" class="detail-input" value="${item.pn_description || ''}" data-field="pn_description" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">BU</span>
                <input type="text" class="detail-input" value="${item.bu || ''}" data-field="bu" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Category</span>
                <input type="text" class="detail-input" value="${item.category || ''}" data-field="category" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item detail-item-full">
                <span class="detail-label">Responsável</span>
                <input type="text" class="detail-input" value="${item.cummins_responsible || ''}" data-field="cummins_responsible" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Status</span>
                <select class="detail-input" data-field="status" data-id="${item.id}" onchange="handleStatusSelectChange(${index}, ${item.id}, this)">
                  ${statusOptionsMarkup}
                </select>
              </div>
              <div class="detail-item">
                <span class="detail-label">Steps</span>
                <select class="detail-input" data-field="steps" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                  <option value="" ${!item.steps ? 'selected' : ''}>Select...</option>
                  <option value="1" ${item.steps === '1' ? 'selected' : ''}>1</option>
                  <option value="2" ${item.steps === '2' ? 'selected' : ''}>2</option>
                  <option value="3" ${item.steps === '3' ? 'selected' : ''}>3</option>
                  <option value="4" ${item.steps === '4' ? 'selected' : ''}>4</option>
                  <option value="5" ${item.steps === '5' ? 'selected' : ''}>5</option>
                  <option value="6" ${item.steps === '6' ? 'selected' : ''}>6</option>
                  <option value="7" ${item.steps === '7' ? 'selected' : ''}>7</option>
                </select>
              </div>
              <div class="detail-item detail-item-full obsolete-link-field ${hasReplacementLink ? 'has-link' : ''}" data-obsolete-link ${replacementEditorVisibilityAttr}>
                <span class="detail-label">Replacement ID</span>
                <div class="replacement-link-field">
                  <div class="replacement-dropdown" data-replacement-picker>
                    <button type="button" class="replacement-dropdown-trigger ${hasReplacementLink ? 'has-selection' : ''}" data-replacement-picker-button onclick="toggleReplacementPicker(event, ${index})">
                      <span data-replacement-picker-label>${replacementPickerLabel}</span>
                      <i class="ph ph-caret-down"></i>
                    </button>
                    <div class="replacement-dropdown-panel" data-replacement-picker-panel hidden>
                      <div class="replacement-dropdown-search">
                        <input type="text" placeholder="Search ID or PN" data-replacement-picker-search oninput="handleReplacementPickerSearch(${index}, this.value)">
                      </div>
                      <div class="replacement-dropdown-list" data-replacement-picker-list>
                        ${replacementPickerOptionsMarkup}
                      </div>
                    </div>
                  </div>
                  <input type="text" class="replacement-hidden-input" value="${hasReplacementLink ? replacementIdValue : ''}" data-field="replacement_tooling_id" data-id="${item.id}" hidden aria-hidden="true">
                  <button type="button" class="btn-link-card" data-replacement-open-btn ${hasReplacementLink ? '' : 'disabled'} data-target-id="${hasReplacementLink ? replacementIdValue : ''}" onclick="handleReplacementLinkButtonClick(event, this)">
                    Go to card
                  </button>
                </div>
                <p class="detail-help">Appears when status is Obsolete. Use the ID displayed on the replacement card.</p>
              </div>
            </div>
            
                
                <!-- Column 3: Supplier & Tooling -->
                <div class="detail-group detail-grid">
                  <div class="detail-group-title">Supplier & Tooling</div>
              <div class="detail-item detail-item-full">
                <span class="detail-label">Tooling Description</span>
                <input type="text" class="detail-input" value="${item.tool_description || ''}" data-field="tool_description" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Supplier</span>
                <input type="text" class="detail-input" value="${item.supplier || ''}" data-field="supplier" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Ownership</span>
                <input type="text" class="detail-input" value="${item.tool_ownership || ''}" data-field="tool_ownership" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Customer</span>
                <input type="text" class="detail-input" value="${item.customer || ''}" data-field="customer" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Tool Number</span>
                <input type="text" class="detail-input" value="${item.tool_number_arb || ''}" data-field="tool_number_arb" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Quantity</span>
                <input type="text" class="detail-input" value="${item.tool_quantity || ''}" data-field="tool_quantity" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Value (BRL)</span>
                <input type="number" class="detail-input" value="${amountBrlValue}" data-field="amount_brl" data-id="${item.id}" onchange="autoSaveTooling(${item.id})" inputmode="decimal" step="0.01" min="0">
                  </div>
                </div>
              </div>
            </div>
            <button class="carousel-nav carousel-nav-next" onclick="navigateCarousel(${index}, 'next')" aria-label="Next">
              <i class="ph ph-caret-right"></i>
            </button>
          </div>
          </div>

          <!-- Aba Documentation -->
          <div class="card-tab-content" data-tab="documentation">
            <div class="documentation-grid">
              <div class="detail-item">
                <span class="detail-label">Bailment Agreement Signed</span>
                <input type="text" class="detail-input" value="${item.bailment_agreement_signed || ''}" data-field="bailment_agreement_signed" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Tooling Book</span>
                <input type="text" class="detail-input" value="${item.tooling_book || ''}" data-field="tooling_book" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Disposition</span>
                <input type="text" class="detail-input" value="${item.disposition || ''}" data-field="disposition" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">VPCR</span>
                <input type="text" class="detail-input" value="${item.vpcr || ''}" data-field="vpcr" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Finish Due Date</span>
                <input type="date" class="detail-input" value="${item.finish_due_date || ''}" data-field="finish_due_date" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Asset Number</span>
                <input type="text" class="detail-input" value="${item.asset_number || ''}" data-field="asset_number" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item full-width">
                <span class="detail-label">STIM</span>
                <input type="text" class="detail-input" value="${item.stim_tooling_management || ''}" data-field="stim_tooling_management" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
            </div>
          </div>

          <!-- Aba Comments -->
          <div class="card-tab-content" data-tab="comments">
            <div class="comments-container">
              <textarea class="comments-textarea" rows="10" data-field="comments" data-id="${item.id}" onchange="autoSaveTooling(${item.id})" placeholder="Add your comments here...">${item.comments || ''}</textarea>
            </div>
          </div>

          <!-- Aba Todos -->
          <div class="card-tab-content" data-tab="todos">
            <div class="todos-container" id="todosContainer-${item.id}">
              <div class="todos-list" id="todosList-${item.id}"></div>
              <button class="btn-add-todo" onclick="addTodoItem(${item.id})">
                <i class="ph ph-plus"></i>
              </button>
            </div>
          </div>

          <!-- Aba Attachments -->
          <div class="card-tab-content" data-tab="attachments">
            <div class="card-attachments" data-card-id="${item.id}">
              <div class="card-attachments-dropzone" onclick="uploadCardAttachment(${item.id})">
                <i class="ph ph-upload-simple"></i>
                <span>Click or drop files here</span>
              </div>
              <div class="card-attachments-list" id="cardAttachments-${item.id}"></div>
            </div>
          </div>

          <div class="card-actions">
            <button class="btn-delete" onclick="confirmDeleteTooling(${item.id})">
              <i class="ph ph-trash"></i>
              Delete
            </button>
            <button class="btn-save" onclick="saveTooling(${item.id})">
              <i class="ph ph-floppy-disk"></i>
              Save Changes
            </button>
          </div>
        </div>
      </div>
    `;
  }).join('');

  sortedData.forEach((item, index) => {
    syncReplacementLinkControls(index, sanitizeReplacementId(item.replacement_tooling_id));
    updateCardStatusAttribute(index, item.status || '');
    // Load todos badge for each card
    updateTodoBadge(item.id);
  });

  setupCardAttachmentDropzones();
}

// Alterna expansão do card
async function toggleAllCards() {
  const cards = document.querySelectorAll('.tooling-card');
  const expandBtn = document.getElementById('floatingExpandBtn');
  const expandedCards = document.querySelectorAll('.tooling-card.expanded');
  const shouldExpand = expandedCards.length === 0;

  for (let i = 0; i < cards.length; i++) {
    const card = cards[i];
    const isExpanded = card.classList.contains('expanded');
    
    if (shouldExpand && !isExpanded) {
      card.classList.add('expanded');
      const cardIndex = parseInt(card.id.replace('card-', ''), 10);
      calculateExpirationDate(cardIndex);
      const itemId = card.getAttribute('data-item-id');
      if (itemId) {
        await loadCardAttachments(itemId);
      }
    } else if (!shouldExpand && isExpanded) {
      const itemId = card.getAttribute('data-item-id');
      if (itemId) {
        await saveToolingQuietly(itemId);
      }
      card.classList.remove('expanded');
    }
  }

  // Remove classe de desfoque quando todos estão expandidos/colapsados
  cards.forEach(c => c.classList.remove('has-expanded-sibling'));

  // Atualiza o ícone do botão
  const icon = expandBtn.querySelector('i');
  if (shouldExpand) {
    icon.className = 'ph ph-arrows-in';
    expandBtn.title = 'Collapse All';
  } else {
    icon.className = 'ph ph-arrows-out';
    expandBtn.title = 'Expand All';
  }
}

function navigateCarousel(cardIndex, direction) {
  const card = document.getElementById(`card-${cardIndex}`);
  if (!card) return;
  
  const track = card.querySelector('[data-carousel-track]');
  const carousel = card.querySelector('.tooling-details-carousel');
  const prevBtn = card.querySelector('.carousel-nav-prev');
  const nextBtn = card.querySelector('.carousel-nav-next');
  if (!track || !carousel) return;
  
  const columns = Array.from(track.children);
  const currentTransform = getComputedStyle(track).transform;
  let currentIndex = 0;
  
  if (currentTransform !== 'none') {
    const matrix = currentTransform.match(/matrix\(([^)]+)\)/);
    if (matrix) {
      const currentX = parseFloat(matrix[1].split(',')[4]) || 0;
      const columnWidth = carousel.offsetWidth;
      currentIndex = Math.round(Math.abs(currentX) / columnWidth);
    }
  }
  
  if (direction === 'next') {
    currentIndex++;
    if (currentIndex >= columns.length) currentIndex = columns.length - 1;
  } else {
    currentIndex--;
    if (currentIndex < 0) currentIndex = 0;
  }
  
  const columnWidth = carousel.offsetWidth;
  const newX = -(currentIndex * columnWidth);
  
  track.style.transform = `translateX(${newX}px)`;
  
  // Update button states
  if (prevBtn) prevBtn.disabled = currentIndex === 0;
  if (nextBtn) nextBtn.disabled = currentIndex === columns.length - 1;
}

async function toggleCard(index) {
  const card = document.getElementById(`card-${index}`);
  const wasExpanded = card.classList.contains('expanded');
  
  // Se está fechando o card, salva antes
  if (wasExpanded) {
    const itemId = card.getAttribute('data-item-id');
    if (itemId) {
      await saveToolingQuietly(itemId);
    }
  }
  
  // Se está abrindo um novo card, salva o card anterior que estava aberto
  if (!wasExpanded) {
    const expandedCards = document.querySelectorAll('.tooling-card.expanded');
    for (const expandedCard of expandedCards) {
      const itemId = expandedCard.getAttribute('data-item-id');
      if (itemId) {
        await saveToolingQuietly(itemId);
      }
    }
  }
  
  card.classList.toggle('expanded');
  
  // Adiciona/remove classe de desfoque nos outros cards
  const allCards = document.querySelectorAll('.tooling-card');
  const hasExpandedCard = document.querySelector('.tooling-card.expanded');
  
  allCards.forEach(c => {
    if (hasExpandedCard) {
      c.classList.add('has-expanded-sibling');
    } else {
      c.classList.remove('has-expanded-sibling');
    }
  });
  
  // Recalcula expiration date ao abrir o card
  if (card.classList.contains('expanded')) {
    calculateExpirationDate(index);
    // Carrega anexos do card
    const itemId = card.getAttribute('data-item-id');
    if (itemId) {
      await loadCardAttachments(itemId);
    }
    ensureCardVisible(card);
  }
}

// Troca de aba dentro do card
function switchCardTab(cardIndex, tabName) {
  const card = document.getElementById(`card-${cardIndex}`);
  if (!card) return;
  
  // Atualiza botões das abas
  const tabs = card.querySelectorAll('.card-tab');
  tabs.forEach(tab => tab.classList.remove('active'));
  const activeTab = Array.from(tabs).find(tab => 
    tab.getAttribute('onclick').includes(`'${tabName}'`)
  );
  if (activeTab) activeTab.classList.add('active');
  
  // Atualiza conteúdo das abas
  const contents = card.querySelectorAll('.card-tab-content');
  contents.forEach(content => content.classList.remove('active'));
  const activeContent = card.querySelector(`.card-tab-content[data-tab="${tabName}"]`);
  if (activeContent) activeContent.classList.add('active');
  
  // Carrega conteúdo específico da aba
  const itemId = card.getAttribute('data-item-id');
  if (itemId) {
    if (tabName === 'attachments') {
      loadCardAttachments(itemId);
    } else if (tabName === 'todos') {
      loadTodos(itemId);
    }
  }
}

function ensureCardVisible(card) {
  if (!card) return;
  requestAnimationFrame(() => {
    const scrollContainer = card.closest('.content-body');
    const scroller = scrollContainer || document.scrollingElement || document.documentElement;
    if (!scroller) {
      card.scrollIntoView({ behavior: 'smooth', block: 'start', inline: 'nearest' });
      return;
    }

    const containerRect = scrollContainer
      ? scrollContainer.getBoundingClientRect()
      : { top: 0, bottom: window.innerHeight };
    const cardRect = card.getBoundingClientRect();
    const topSpacing = 8;
    const desiredTop = containerRect.top + topSpacing;
    const offset = cardRect.top - desiredTop;

    if (Math.abs(offset) < 4) {
      return;
    }

    if (scrollContainer) {
      scrollContainer.scrollBy({ top: offset, behavior: 'smooth' });
    } else {
      window.scrollBy({ top: offset, behavior: 'smooth' });
    }
  });
}

function normalizeItemId(itemId) {
  const numericId = Number(itemId);
  return Number.isNaN(numericId) ? null : numericId;
}

function findToolingItem(itemId) {
  const normalizedId = normalizeItemId(itemId);
  if (normalizedId === null) {
    return { normalizedId: null, toolingItem: null };
  }

  const toolingItem = toolingData.find(t => Number(t.id) === normalizedId) || null;
  return { normalizedId, toolingItem };
}

function setupCardAttachmentDropzones() {
  const dropzones = document.querySelectorAll('.card-attachments-dropzone');

  dropzones.forEach(dropzone => {
    if (!dropzone || dropzone.dataset.ddInitialized === 'true') {
      return;
    }

    const cardWrapper = dropzone.closest('.card-attachments');
    const itemId = cardWrapper?.dataset.cardId;
    if (!itemId) {
      return;
    }

    let dragCounter = 0;

    const handleDragEnter = event => {
      if (!isFileDrag(event)) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      dragCounter += 1;
      dropzone.classList.add('drop-active');
    };

    const handleDragOver = event => {
      if (!isFileDrag(event)) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      if (event.dataTransfer) {
        event.dataTransfer.dropEffect = 'copy';
      }
      dropzone.classList.add('drop-active');
    };

    const handleDragLeave = event => {
      if (!isFileDrag(event)) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      dragCounter = Math.max(dragCounter - 1, 0);
      if (dragCounter === 0) {
        dropzone.classList.remove('drop-active');
      }
    };

    const handleDrop = event => {
      if (!isFileDrag(event)) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      dragCounter = 0;
      dropzone.classList.remove('drop-active');
      const files = Array.from(event.dataTransfer?.files || []);
      if (files.length === 0) {
        return;
      }
      handleCardAttachmentFiles(itemId, files);
    };

    dropzone.addEventListener('dragenter', handleDragEnter, false);
    dropzone.addEventListener('dragover', handleDragOver, false);
    dropzone.addEventListener('dragleave', handleDragLeave, false);
    dropzone.addEventListener('drop', handleDrop, false);

    dropzone.dataset.ddInitialized = 'true';
  });
}

async function handleCardAttachmentFiles(itemId, files) {
  if (!files || files.length === 0) {
    return;
  }

  const { normalizedId, toolingItem } = findToolingItem(itemId);
  if (normalizedId === null || !toolingItem || !toolingItem.supplier) {
    showNotification('Ferramental inválido para anexar arquivos.', 'error');
    return;
  }

  const paths = files
    .map(file => file?.path)
    .filter(Boolean);

  if (paths.length === 0) {
    showNotification('Não foi possível ler os arquivos arrastados.', 'error');
    return;
  }

  try {
    const result = await window.api.uploadAttachmentFromPaths(toolingItem.supplier, paths, normalizedId);

    if (!result) {
      showNotification('Não foi possível anexar os arquivos.', 'error');
      return;
    }

    const failures = Array.isArray(result.results)
      ? result.results.filter(item => item?.success !== true)
      : [];

    if (failures.length > 0 || result.success !== true) {
      const errorMessage = failures[0]?.error || result.error || 'Erro ao anexar arquivos.';
      showNotification(errorMessage, 'error');
    } else {
      const successMessage = paths.length > 1
        ? 'Arquivos anexados com sucesso!'
        : 'Arquivo anexado com sucesso!';
      showNotification(successMessage);
    }

    await loadCardAttachments(normalizedId);
  } catch (error) {
    console.error('Erro ao anexar arquivos ao card:', error);
    showNotification('Erro ao anexar arquivos.', 'error');
  }
}

// Atualiza o contador de anexos no header do card
function updateAttachmentCount(itemId, count) {
  const card = document.querySelector(`[data-item-id="${itemId}"]`);
  if (!card) return;
  
  const topMeta = card.querySelector('.tooling-card-top-meta');
  if (!topMeta) return;
  
  // Remove o contador existente
  const existingCounter = topMeta.querySelector('.tooling-attachment-count');
  if (existingCounter) {
    existingCounter.remove();
  }
  
  // Se tem anexos, adiciona o contador antes do expand button
  if (count > 0) {
    const expandButton = topMeta.querySelector('.tooling-card-expand');
    const counterHtml = `
      <div class="tooling-attachment-count">
        <i class="ph ph-paperclip"></i>
        <span>${count}</span>
      </div>
    `;
    expandButton.insertAdjacentHTML('beforebegin', counterHtml);
  }
}

// Carrega anexos específicos do card
async function loadCardAttachments(itemId) {
  try {
    const { normalizedId, toolingItem } = findToolingItem(itemId);
    if (normalizedId === null) {
      console.warn('loadCardAttachments: itemId inválido', itemId);
      return;
    }

    if (!toolingItem || !toolingItem.supplier) {
      return;
    }

    const attachments = await window.api.getAttachments(toolingItem.supplier, normalizedId);
    
    // Atualiza o contador no header do card
    updateAttachmentCount(normalizedId, attachments.length);
    
    const container = document.getElementById(`cardAttachments-${normalizedId}`);
    
    if (!container) {
      return;
    }

    if (!attachments || attachments.length === 0) {
      container.innerHTML = '<p style="color: #999; font-size: 13px; text-align: center; padding: 20px;">Nenhum anexo disponível</p>';
      return;
    }

    container.innerHTML = attachments.map(att => {
      const fileSize = (att.fileSize / 1024).toFixed(1);
      const uploadDate = new Date(att.uploadDate).toLocaleDateString('pt-BR');
      
      return `
        <div class="card-attachment-item">
          <div class="card-attachment-info">
            <i class="ph ph-file"></i>
            <div class="card-attachment-details">
              <span class="card-attachment-name">${att.fileName}</span>
              <span class="card-attachment-meta">${fileSize} KB • ${uploadDate}</span>
            </div>
          </div>
          <div class="card-attachment-actions">
            <button class="btn-attachment" onclick="openCardAttachmentFile('${att.supplierName}', '${att.fileName}', ${att.itemId})" title="Open">
              <i class="ph ph-eye"></i>
            </button>
            <button class="btn-attachment" onclick="deleteCardAttachmentFile('${att.supplierName}', '${att.fileName}', ${att.itemId})" title="Delete">
              <i class="ph ph-trash"></i>
            </button>
          </div>
        </div>
      `;
    }).join('');
  } catch (error) {
    console.error('Erro ao carregar anexos do card:', error);
  }
}

// Upload de anexo específico do card
async function uploadCardAttachment(itemId) {
  try {
    const { normalizedId, toolingItem } = findToolingItem(itemId);
    if (normalizedId === null) {
      showNotification('Ferramental inválido para anexar arquivos.', 'error');
      return;
    }

    if (!toolingItem || !toolingItem.supplier) {
      showNotification('Fornecedor não encontrado para este ferramental.', 'error');
      return;
    }

    const result = await window.api.uploadAttachment(toolingItem.supplier, normalizedId);
    
    if (result && result.success) {
      showNotification(result.message || 'Arquivo(s) anexado(s) com sucesso!');
      await loadCardAttachments(normalizedId);
    }
  } catch (error) {
    console.error('Erro ao fazer upload de anexo do card:', error);
    showNotification('Erro ao anexar arquivo', 'error');
  }
}

// Salva alterações do ferramental
async function saveTooling(id) {
  try {
    // Coleta todos os dados dos inputs do card
    const inputs = document.querySelectorAll(`[data-id="${id}"]`);
    const data = {};
    
    inputs.forEach(input => {
      const field = input.getAttribute('data-field');
      data[field] = input.value;
    });

    // Normaliza campos numéricos
    const numericFields = ['tooling_life_qty','produced','annual_volume_forecast','remaining_tooling_life_pcs','percent_tooling_life','amount_brl','tool_quantity'];
    numericFields.forEach(f => {
      if (data.hasOwnProperty(f) && data[f] !== '') {
        const parsed = Number(data[f]);
        data[f] = Number.isNaN(parsed) ? 0 : parsed;
      }
    });

    // Envia para o backend
    await window.api.updateTooling(id, data);
    showNotification('Dados salvos com sucesso!');
    
    // Atualiza apenas o item no array local sem recarregar tudo
    const index = toolingData.findIndex(item => item.id === id);
    if (index !== -1) {
      toolingData[index] = { ...toolingData[index], ...data };
      // Recalcula a data de expiração para o card atualizado
      calculateExpirationDate(index);
    }
  } catch (error) {
    console.error('Erro ao salvar:', error);
    showNotification('Erro ao salvar dados', 'error');
  }
}

// Salva sem mostrar notificação (usado ao trocar de card)
async function saveToolingQuietly(id) {
  try {
    // Coleta todos os dados dos inputs do card
    const inputs = document.querySelectorAll(`[data-id="${id}"]`);
    const data = {};
    
    inputs.forEach(input => {
      const field = input.getAttribute('data-field');
      data[field] = input.value;
    });

    // Normaliza campos numéricos
    const numericFields = ['tooling_life_qty','produced','annual_volume_forecast','remaining_tooling_life_pcs','percent_tooling_life','amount_brl','tool_quantity'];
    numericFields.forEach(f => {
      if (data.hasOwnProperty(f) && data[f] !== '') {
        const parsed = Number(data[f]);
        data[f] = Number.isNaN(parsed) ? 0 : parsed;
      }
    });

    // Envia para o backend
    await window.api.updateTooling(id, data);
    
    // Atualiza apenas o item no array local sem recarregar tudo
    const index = toolingData.findIndex(item => item.id === id);
    if (index !== -1) {
      toolingData[index] = { ...toolingData[index], ...data };
    }
    
    // Atualiza a interface: suppliers, barra inferior e card header
    await updateInterfaceAfterSave();
  } catch (error) {
    console.error('Erro ao salvar silenciosamente:', error);
  }
}

// Atualiza suppliers, barra inferior e dados visíveis após salvar
async function updateInterfaceAfterSave() {
  try {
    // Recarrega suppliers com estatísticas atualizadas
    await loadSuppliers();
    
    // Atualiza a barra inferior com analytics
    const analytics = await window.api.getAnalytics();
    if (!Array.isArray(suppliersData) || suppliersData.length === 0) {
      updateStatusBar(analytics);
    } else {
      syncStatusBarWithSuppliers();
    }
  } catch (error) {
    console.error('Erro ao atualizar interface:', error);
  }
}

// Carrega analytics
async function loadAnalytics() {
  try {
    const analytics = await window.api.getAnalytics();
    
    // Métricas principais
    document.getElementById('totalTooling').textContent = analytics.total || 0;
    document.getElementById('totalSuppliers').textContent = analytics.suppliers || 0;
    document.getElementById('totalResponsibles').textContent = analytics.responsibles || 0;
    
    // Calcula dados usando suppliersData (mesma lógica da barra inferior)
    let totalExpired = 0;
    let totalExpiring = 0;
    let totalObsolete = 0;
    
    if (Array.isArray(suppliersData) && suppliersData.length > 0) {
      suppliersData.forEach(supplier => {
        const expired = parseInt(supplier.expired, 10) || 0;
        const warning1 = parseInt(supplier.warning_1year, 10) || 0;
        const warning2 = parseInt(supplier.warning_2years, 10) || 0;
        
        totalExpired += expired;
        // Expiring = expired + warning1 + warning2 (igual à barra inferior)
        totalExpiring += (expired + warning1 + warning2);
      });
    }
    
    // Conta obsoletos de todos os toolings
    const allTooling = await window.api.loadTooling();
    const statusCount = {};
    
    allTooling.forEach(item => {
      const status = String(item.status || 'N/A').trim();
      statusCount[status] = (statusCount[status] || 0) + 1;
      
      if (status.toLowerCase() === 'obsolete') {
        totalObsolete++;
      }
    });
    
    // Atualiza métricas de alerta
    document.getElementById('totalExpired').textContent = totalExpired;
    document.getElementById('totalExpiring').textContent = totalExpiring;
    document.getElementById('totalObsolete').textContent = totalObsolete;
    
    // Distribui por status
    displayStatusDistribution(statusCount);
    
    // Top suppliers
    displayTopSuppliers(suppliersData);
    
    if (!Array.isArray(suppliersData) || suppliersData.length === 0) {
      updateStatusBar(analytics);
    } else {
      syncStatusBarWithSuppliers();
    }
  } catch (error) {
    console.error('Erro ao carregar analytics:', error);
  }
}

function displayStatusDistribution(statusCount) {
  const container = document.getElementById('statusDistribution');
  if (!container) return;
  
  const sortedStatuses = Object.entries(statusCount)
    .sort((a, b) => b[1] - a[1]);
  
  container.innerHTML = sortedStatuses.map(([status, count]) => `
    <div class="status-card">
      <div class="status-card-name">${escapeHtml(status)}</div>
      <div class="status-card-count">${count}</div>
    </div>
  `).join('');
}

function displayTopSuppliers(suppliers) {
  const tableBody = document.querySelector('#topSuppliersTable tbody');
  if (!tableBody || !suppliers) return;
  
  const sortedSuppliers = [...suppliers]
    .sort((a, b) => (parseInt(b.total) || 0) - (parseInt(a.total) || 0))
    .slice(0, 10);
  
  tableBody.innerHTML = sortedSuppliers.map(supplier => {
    const total = parseInt(supplier.total) || 0;
    const expired = parseInt(supplier.expired) || 0;
    const warning1 = parseInt(supplier.warning_1year) || 0;
    const warning2 = parseInt(supplier.warning_2years) || 0;
    const expiring = warning1 + warning2;
    
    return `
      <tr>
        <td>${escapeHtml(supplier.supplier)}</td>
        <td><span class="table-number">${total}</span></td>
        <td><span class="table-number ${expired > 0 ? 'table-number-danger' : ''}">${expired}</span></td>
        <td><span class="table-number ${expiring > 0 ? 'table-number-warning' : ''}">${expiring}</span></td>
      </tr>
    `;
  }).join('');
}

function syncStatusBarWithSuppliers() {
  if (!Array.isArray(suppliersData) || suppliersData.length === 0) {
    return;
  }

  const summary = suppliersData.reduce((acc, supplier) => {
    const total = parseInt(supplier.total, 10) || 0;
    const expired = parseInt(supplier.expired, 10) || 0;
    const warning1 = parseInt(supplier.warning_1year, 10) || 0;
    const warning2 = parseInt(supplier.warning_2years, 10) || 0;

    acc.total += total;
    acc.expired += expired;
    acc.expiring += expired + warning1 + warning2;
    return acc;
  }, { total: 0, expired: 0, expiring: 0 });

  updateStatusBar({
    total: summary.total,
    expired_total: summary.expired,
    expiring_two_years: summary.expiring
  });
}

// Atualiza barra de status
function updateStatusBar(analytics) {
  const statusTotal = document.getElementById('statusTotal');
  const statusExpired = document.getElementById('statusExpired');
  const statusExpiring = document.getElementById('statusExpiring');
  
  if (statusTotal) {
    statusTotal.textContent = analytics.total || 0;
  }
  if (statusExpired) {
    statusExpired.textContent = analytics.expired_total || 0;
  }
  if (statusExpiring) {
    statusExpiring.textContent = analytics.expiring_two_years || 0;
  }
}

class ToastManager {
  constructor() {
    this.container = null;
    this.activeTimeouts = new WeakMap();
  }

  ensureContainer() {
    if (!this.container) {
      this.container = document.createElement('div');
      this.container.className = 'toast-container';
      document.body.appendChild(this.container);
    }
  }

  getVariant(type) {
    const allowed = ['success', 'error', 'warning', 'info'];
    if (!allowed.includes(type)) {
      return 'info';
    }
    return type;
  }

  show(message, type = 'success', duration = 4000) {
    this.ensureContainer();
    const variant = this.getVariant(type);
    const toast = document.createElement('div');
    toast.className = `toast toast-${variant}`;

    const textSpan = document.createElement('span');
    textSpan.className = 'toast-message';
    textSpan.textContent = message;

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'toast-close';
    closeBtn.innerHTML = '<i class="ph ph-x"></i>';
    closeBtn.addEventListener('click', () => this.dismiss(toast));

    toast.appendChild(textSpan);
    toast.appendChild(closeBtn);

    this.container.appendChild(toast);

    const timeout = setTimeout(() => this.dismiss(toast), duration);
    this.activeTimeouts.set(toast, timeout);
  }

  dismiss(toast) {
    if (!toast) return;
    const timeout = this.activeTimeouts.get(toast);
    if (timeout) {
      clearTimeout(timeout);
      this.activeTimeouts.delete(toast);
    }
    toast.classList.add('toast-hide');
    toast.addEventListener('animationend', () => {
      toast.remove();
    }, { once: true });
  }
}

const toastManager = new ToastManager();

// Mostra notificação
function showNotification(message, type = 'success') {
  if (!message) return;
  toastManager.show(message, type);
}

function initStatusSettings(elements = {}) {
  statusSettingsElements = {
    ...statusSettingsElements,
    ...elements
  };

  const { list, input, addButton } = statusSettingsElements;

  if (addButton) {
    addButton.addEventListener('click', handleAddStatusOption);
  }

  if (input) {
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        handleAddStatusOption();
      }
    });
  }

  if (list) {
    list.addEventListener('click', handleStatusListClick);
  }

  renderStatusSettings();
  updateAllStatusSelects();
}

function handleStatusListClick(event) {
  const targetButton = event.target.closest('button[data-status-value]');
  if (!targetButton) {
    return;
  }
  const value = targetButton.getAttribute('data-status-value');
  removeStatusOption(value);
}

function handleAddStatusOption() {
  const input = statusSettingsElements.input;
  if (!input) {
    return;
  }

  const newStatusRaw = input.value.trim();
  if (!newStatusRaw) {
    showNotification('Informe um nome para o status.', 'error');
    return;
  }

  const exists = statusOptions.some(option => option.toLowerCase() === newStatusRaw.toLowerCase());
  if (exists) {
    showNotification('Este status já está na lista.', 'error');
    input.value = '';
    return;
  }

  statusOptions.push(newStatusRaw);
  saveStatusOptionsToStorage();
  renderStatusSettings();
  updateAllStatusSelects();
  input.value = '';
  showNotification('Status adicionado com sucesso!');
}

function removeStatusOption(value) {
  if (!value) {
    return;
  }

  if (statusOptions.length <= 1) {
    showNotification('Mantenha pelo menos um status na lista.', 'error');
    return;
  }

  statusOptions = statusOptions.filter(option => option !== value);
  if (statusOptions.length === 0) {
    statusOptions = [...DEFAULT_STATUS_OPTIONS];
  }

  saveStatusOptionsToStorage();
  renderStatusSettings();
  updateAllStatusSelects();
  showNotification('Status removido.');
}

function renderStatusSettings() {
  const list = statusSettingsElements.list;
  if (!list) {
    return;
  }

  if (!Array.isArray(statusOptions) || statusOptions.length === 0) {
    statusOptions = [...DEFAULT_STATUS_OPTIONS];
  }

  const items = statusOptions.map(status => {
    const safeLabel = escapeHtml(status);
    const isDefault = DEFAULT_STATUS_OPTIONS.includes(status);
    return `
      <li class="status-option-pill ${isDefault ? 'status-option-default' : ''}">
        <span>${safeLabel}</span>
        ${!isDefault ? `<button type="button" class="status-option-remove" data-status-value="${safeLabel}" title="Remover">
          <i class="ph ph-x"></i>
        </button>` : ''}
      </li>
    `;
  });

  list.innerHTML = items.join('') || '<li class="status-option-pill">Nenhum status configurado</li>';
}

function updateAllStatusSelects() {
  const selects = document.querySelectorAll('select.detail-input[data-field="status"]');
  selects.forEach(select => {
    const currentValue = select.value;
    select.innerHTML = buildStatusOptionsMarkup(currentValue);
    if (currentValue) {
      select.value = currentValue;
    }
  });
}

function buildStatusOptionsMarkup(selectedValue) {
  const normalizedSelected = (selectedValue || '').trim();
  const options = Array.isArray(statusOptions) && statusOptions.length > 0
    ? statusOptions
    : [...DEFAULT_STATUS_OPTIONS];

  const normalizedSelectedLower = normalizedSelected.toLowerCase();
  const hasSelected = normalizedSelected
    ? options.some(option => option.toLowerCase() === normalizedSelectedLower)
    : false;

  const optionsHtml = [];
  const blankLabel = '&nbsp;';
  optionsHtml.push(`<option value="" ${normalizedSelected ? '' : 'selected'}>${blankLabel}</option>`);

  options.forEach(option => {
    const safeValue = escapeHtml(option);
    const isSelected = option === normalizedSelected;
    optionsHtml.push(`<option value="${safeValue}" ${isSelected ? 'selected' : ''}>${safeValue}</option>`);
  });

  if (normalizedSelected && !hasSelected) {
    const safeSelected = escapeHtml(normalizedSelected);
    optionsHtml.push(`<option value="${safeSelected}" selected>${safeSelected}</option>`);
  }

  return optionsHtml.join('');
}

function saveStatusOptionsToStorage() {
  try {
    localStorage.setItem(STATUS_STORAGE_KEY, JSON.stringify(statusOptions));
  } catch (error) {
    console.error('Erro ao salvar status personalizados:', error);
  }
}

function loadStatusOptionsFromStorage() {
  try {
    const stored = localStorage.getItem(STATUS_STORAGE_KEY);
    if (stored) {
      const parsed = JSON.parse(stored);
      if (Array.isArray(parsed) && parsed.length > 0) {
        const normalized = parsed
          .map(item => (typeof item === 'string' ? item.trim() : ''))
          .filter(item => item.length > 0);
        
        // Migração: garantir que ACTIVE está na lista se não estiver
        if (normalized.length > 0 && !normalized.includes('ACTIVE')) {
          normalized.unshift('ACTIVE'); // Adiciona no início
          // Salvar a lista atualizada
          localStorage.setItem(STATUS_STORAGE_KEY, JSON.stringify(normalized));
          return Array.from(new Set(normalized));
        }
        
        if (normalized.length > 0) {
          return Array.from(new Set(normalized));
        }
      }
    }
  } catch (error) {
    console.error('Erro ao carregar status personalizados:', error);
  }
  return [...DEFAULT_STATUS_OPTIONS];
}

function escapeHtml(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).replace(/[&<>"']/g, (char) => {
    const map = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    };
    return map[char] || char;
  });
}

// Função para filtrar suppliers e toolings
async function filterSuppliersAndTooling(searchTerm) {
  const normalizedSearch = searchTerm.toLowerCase().trim();
  
  if (!normalizedSearch) {
    // Se não há termo de busca, aplica o filtro de expiração normal
    applyExpirationFilter();
    if (selectedSupplier) {
      await loadToolingBySupplier(selectedSupplier);
    }
    return;
  }

  // Filtra suppliers que correspondem ao termo
  const matchingSuppliers = new Set();
  
  // Busca em todos os toolings
  try {
    const allResults = await window.api.searchTooling(searchTerm);
    
    // Agrupa por supplier
    allResults.forEach(item => {
      const supplierName = String(item.supplier || '').trim();
      if (supplierName) {
        matchingSuppliers.add(supplierName);
      }
    });
    
    // Adiciona suppliers cujo nome contém o termo
    suppliersData.forEach(supplier => {
      const supplierName = String(supplier.supplier || '').toLowerCase();
      if (supplierName.includes(normalizedSearch)) {
        matchingSuppliers.add(supplier.supplier);
      }
    });
    
    // Filtra a lista de suppliers pela busca
    let filteredSuppliers = suppliersData.filter(supplier => 
      matchingSuppliers.has(supplier.supplier)
    );
    
    // Aplica o filtro de expiração se estiver ativo
    if (expirationFilterEnabled) {
      filteredSuppliers = filteredSuppliers.filter(supplier => {
        const expired = parseInt(supplier.expired) || 0;
        const warning1 = parseInt(supplier.warning_1year) || 0;
        const warning2 = parseInt(supplier.warning_2years) || 0;
        const critical = expired + warning1 + warning2;
        return critical > 0;
      });
    }
    
    displaySuppliers(filteredSuppliers);
    
    // Se há um supplier selecionado, filtra seus toolings
    if (selectedSupplier && matchingSuppliers.has(selectedSupplier)) {
      const filteredResults = allResults.filter(item => {
        const itemSupplier = String(item.supplier || '').trim();
        return itemSupplier === selectedSupplier;
      });
      displayTooling(filteredResults);
    }
  } catch (error) {
    console.error('Erro ao filtrar suppliers e toolings:', error);
  }
}

// Função para limpar a pesquisa de suppliers
function clearSupplierSearch() {
  const input = document.getElementById('supplierSearchInput');
  const clearBtn = document.getElementById('clearSupplierSearch');
  
  if (input) {
    input.value = '';
  }
  
  if (clearBtn) {
    clearBtn.style.display = 'none';
  }
  
  filterSuppliersAndTooling('');
}

// ===== TODOS MANAGEMENT =====

async function loadTodos(toolingId) {
  try {
    const todos = await window.api.getTodos(toolingId);
    const todosList = document.getElementById(`todosList-${toolingId}`);
    
    if (!todosList) return;
    
    todosList.innerHTML = todos.map(todo => `
      <div class="todo-item" data-todo-id="${todo.id}">
        <input type="checkbox" class="todo-checkbox" ${todo.completed ? 'checked' : ''} onchange="toggleTodo(${todo.id}, this.checked)">
        <textarea class="todo-text" onblur="updateTodoText(${todo.id}, this.value)" placeholder="Enter todo...">${todo.text}</textarea>
        <button class="todo-delete" onclick="deleteTodo(${todo.id}, ${toolingId})" title="Delete todo">
          <i class="ph ph-trash"></i>
        </button>
      </div>
    `).join('');
    
    // Update badge on card header if there are incomplete todos
    await updateTodoBadge(toolingId);
  } catch (error) {
    console.error('Error loading todos:', error);
  }
}

async function addTodoItem(toolingId) {
  try {
    await window.api.addTodo(toolingId, '');
    await loadTodos(toolingId);
  } catch (error) {
    console.error('Error adding todo:', error);
    showNotification('Error adding todo', 'error');
  }
}

async function toggleTodo(todoId, completed) {
  try {
    const todoItem = document.querySelector(`[data-todo-id="${todoId}"]`);
    const todoText = todoItem ? todoItem.querySelector('.todo-text').value : '';
    await window.api.updateTodo(todoId, todoText, completed ? 1 : 0);
    
    // Update badge for this tooling
    const container = document.querySelector(`[data-todo-id="${todoId}"]`).closest('.todos-container');
    if (container) {
      const toolingId = container.id.replace('todosContainer-', '');
      await updateTodoBadge(toolingId);
    }
  } catch (error) {
    console.error('Error toggling todo:', error);
  }
}

async function updateTodoText(todoId, text) {
  try {
    const todoItem = document.querySelector(`[data-todo-id="${todoId}"]`);
    const checkbox = todoItem ? todoItem.querySelector('.todo-checkbox') : null;
    const completed = checkbox ? (checkbox.checked ? 1 : 0) : 0;
    await window.api.updateTodo(todoId, text, completed);
  } catch (error) {
    console.error('Error updating todo text:', error);
  }
}

async function deleteTodo(todoId, toolingId) {
  try {
    await window.api.deleteTodo(todoId);
    await loadTodos(toolingId);
  } catch (error) {
    console.error('Error deleting todo:', error);
    showNotification('Error deleting todo', 'error');
  }
}

async function updateTodoBadge(toolingId) {
  try {
    const todos = await window.api.getTodos(toolingId);
    const incompleteTodos = todos.filter(t => !t.completed).length;
    const totalTodos = todos.length;
    
    const card = document.querySelector(`.tooling-card[data-item-id="${toolingId}"]`);
    if (!card) return;
    
    // Remove existing badge
    const existingBadge = card.querySelector('.todos-badge');
    if (existingBadge) {
      existingBadge.remove();
    }
    
    // Add new badge if there are incomplete todos
    if (incompleteTodos > 0) {
      const topMeta = card.querySelector('.tooling-card-top-meta');
      if (topMeta) {
        const badge = document.createElement('div');
        badge.className = 'todos-badge';
        badge.innerHTML = `
          <i class="ph ph-check-square"></i>
          <span>${incompleteTodos}/${totalTodos}</span>
        `;
        // Insert before expand button
        const expandBtn = topMeta.querySelector('.tooling-card-expand');
        if (expandBtn) {
          topMeta.insertBefore(badge, expandBtn);
        } else {
          topMeta.appendChild(badge);
        }
      }
    }
  } catch (error) {
    console.error('Error updating todo badge:', error);
  }
}

