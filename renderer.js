// Estado da aplicação
const APP_VERSION = 'v0.1.1';

let currentTab = 'tooling';
let toolingData = [];
let suppliersData = [];
let selectedSupplier = null;
let currentSupplier = null;

// Estado de seleção múltipla para exportação
let selectionModeActive = false;
let selectedToolingIds = new Set();
let currentSortCriteria = null;
let currentSortOrder = 'asc';
let deleteConfirmState = { id: null, code: '', descriptor: '' };
let commentDeleteState = { itemId: null, commentIndex: null };
let commentDeleteElements = {
  overlay: null,
  context: null,
  date: null,
  text: null,
  confirmButton: null
};
let addToolingElements = {
  overlay: null,
  form: null,
  pnInput: null,
  supplierInput: null,
  supplierList: null,
  ownerInput: null,
  ownerList: null,
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

let replacementPickerOverlayState = {
  overlay: null,
  list: null,
  searchInput: null,
  title: null,
  subtitle: null,
  cardIndex: null,
  itemId: null
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
let stepsInfoElements = {
  overlay: null,
  closeButtons: []
};

let expirationFilterEnabled = false;
let stepsFilteredSuppliers = null; // Lista de suppliers filtrados por step (null = sem filtro)
let columnFilters = {}; // Armazena os filtros de coluna ativos
let columnSort = { column: null, direction: null }; // Armazena a ordenação atual (column: 'expiration' ou 'progress', direction: 'asc' ou 'desc')

let replacementIdOptions = [];
let replacementIdOptionsLoaded = false;
let replacementIdOptionsPromise = null;
let supplierFilterRequestId = 0;
let globalSearchRequestId = 0;
let currentToolingRenderToken = 0;
let supplierSearchDebouncedHandler = null;
let globalSearchDebouncedHandler = null;

const ASYNC_METADATA_CONCURRENCY = 6;
const DATA_TAB_CAROUSEL_BREAKPOINT = 1200;

const STATUS_STORAGE_KEY = 'toolingStatusOptions';
const DEFAULT_STATUS_OPTIONS = [
  'ACTIVE',
  'UNDER CONSTRUCTION',
  'OBSOLETE',
  'INACTIVE'
];

const DEFAULT_REPLACEMENT_PICKER_LABEL = 'Select replacement';

// View mode: always 'spreadsheet' with expandable rows
let currentViewMode = 'spreadsheet';

let statusOptions = loadStatusOptionsFromStorage();
let statusSettingsElements = {
  list: null,
  input: null,
  addButton: null
};

const TEXT_INPUT_TYPES = new Set(['text', 'search', '']);
let uppercaseInputHandlerInitialized = false;

// Classe utilitária para calcular métricas de expiração
class ExpirationMetrics {
  /**
   * Calcula métricas de expiração a partir de uma lista de itens de tooling
   * Esta é a ÚNICA função para calcular métricas - usada em TODOS os lugares
   * @param {Array} items - Array de itens de tooling
   * @returns {Object} - { total, expired, expiring }
   */
  static fromItems(items) {
    if (!Array.isArray(items)) {
      return { total: 0, expired: 0, expiring: 0 };
    }
    
    return items.reduce((acc, item) => {
      // Total sempre conta todos os itens
      acc.total += 1;
      
      // Ignora itens com análise concluída apenas para expired e expiring
      if (item.analysis_completed === 1) {
        return acc;
      }
      
      const classification = classifyToolingExpirationState(item);
      if (classification.state === 'expired') {
        acc.expired += 1;
      } else if (classification.state === 'warning') {
        acc.expiring += 1;
      }
      return acc;
    }, { total: 0, expired: 0, expiring: 0 });
  }
  
  /**
   * Verifica se um supplier tem itens críticos (expired ou expiring)
   * @param {Object} supplier - Objeto do supplier com items[]
   * @returns {boolean}
   */
  static hasCriticalItems(supplier) {
    const metrics = this.fromItems(supplier.items || []);
    return (metrics.expired + metrics.expiring) > 0;
  }
}

function resolveToolingExpirationDate(item) {
  if (!item) {
    return '';
  }

  let expirationDateValue = normalizeExpirationDate(item.expiration_date, item.id);
  if (!expirationDateValue) {
    const toolingLife = Number(parseLocalizedNumber(item.tooling_life_qty)) || Number(item.tooling_life_qty) || 0;
    const produced = Number(parseLocalizedNumber(item.produced)) || Number(item.produced) || 0;
    const remaining = toolingLife - produced;
    const forecast = Number(parseLocalizedNumber(item.annual_volume_forecast)) || Number(item.annual_volume_forecast) || 0;
    const productionDateValue = item.date_remaining_tooling_life || '';
    const calculatedExpiration = calculateExpirationFromFormula({
      remaining,
      forecast,
      productionDate: productionDateValue
    });
    if (calculatedExpiration) {
      expirationDateValue = calculatedExpiration;
    }
  }
  return expirationDateValue;
}

function classifyToolingExpirationState(item) {
  if (!item) {
    return { state: 'ok', label: '', expirationDate: '' };
  }

  const normalizedStatus = (item.status || '').toString().trim().toLowerCase();
  const isObsolete = normalizedStatus === 'obsolete';
  const replacementIdValue = sanitizeReplacementId(item.replacement_tooling_id);
  const hasReplacementLink = replacementIdValue !== '';
  const expirationDateValue = resolveToolingExpirationDate(item) || '';
  const expirationStatus = getExpirationStatus(expirationDateValue || '');

  if (isObsolete) {
    return {
      state: hasReplacementLink ? 'obsolete-replaced' : 'obsolete',
      label: hasReplacementLink ? 'Obsolete (replacement linked)' : 'Obsolete',
      expirationDate: expirationDateValue
    };
  }

  const toolingLife = Number(parseLocalizedNumber(item.tooling_life_qty)) || Number(item.tooling_life_qty) || 0;
  const produced = Number(parseLocalizedNumber(item.produced)) || Number(item.produced) || 0;
  const percentUsedValue = toolingLife > 0 ? (produced / toolingLife) * 100 : 0;

  if (percentUsedValue >= 100 || expirationStatus.class === 'expired') {
    return { state: 'expired', label: 'Expirado', expirationDate: expirationDateValue };
  }

  if (expirationStatus.class === 'warning') {
    return { state: 'warning', label: expirationStatus.label, expirationDate: expirationDateValue };
  }

  return { state: 'ok', label: expirationStatus.label, expirationDate: expirationDateValue };
}

function updateSupplierCardMetricsFromItems(supplierName, items) {
  if (!supplierName || !Array.isArray(items)) {
    return;
  }

  const normalizedKey = encodeURIComponent(supplierName.trim().toLowerCase());
  const targetCard = document.querySelector(`.supplier-card[data-supplier-key="${normalizedKey}"]`);

  if (!targetCard) {
    return;
  }

  const metrics = ExpirationMetrics.fromItems(items);
  const totalEl = targetCard.querySelector('[data-metric="total"]');
  const expiredEl = targetCard.querySelector('[data-metric="expired"]');
  const expiringEl = targetCard.querySelector('[data-metric="expiring"]');

  if (totalEl) {
    totalEl.textContent = metrics.total;
  }
  if (expiredEl) {
    expiredEl.textContent = metrics.expired;
    expiredEl.classList.toggle('expired', metrics.expired > 0);
  }
  if (expiringEl) {
    expiringEl.textContent = metrics.expiring;
    expiringEl.classList.toggle('critical', metrics.expiring > 0);
  }
}

/**
 * Atualiza métricas do card do supplier buscando dados frescos do banco
 * Usar esta função para atualizações em tempo real onde toolingData pode estar filtrado
 */
async function refreshSupplierCardMetricsFromDB(supplierName) {
  if (!supplierName) return;
  
  try {
    // Busca dados frescos diretamente do banco
    const items = await window.api.getToolingBySupplier(supplierName);
    
    // Atualiza suppliersData também para manter consistência
    if (Array.isArray(suppliersData)) {
      const supplierEntry = suppliersData.find(s => s.supplier === supplierName);
      if (supplierEntry) {
        supplierEntry.items = items;
      }
    }
    
    updateSupplierCardMetricsFromItems(supplierName, items);
  } catch (e) {
    console.error('Erro ao atualizar métricas do supplier:', e);
  }
}

function debounce(fn, delay = 250) {
  let timeoutId = null;
  function debounced(...args) {
    if (timeoutId) {
      clearTimeout(timeoutId);
    }
    timeoutId = setTimeout(() => {
      timeoutId = null;
      fn(...args);
    }, delay);
  }
  debounced.cancel = () => {
    if (timeoutId) {
      clearTimeout(timeoutId);
      timeoutId = null;
    }
  };
  return debounced;
}

async function runTasksWithLimit(items, limit, task) {
  if (!Array.isArray(items) || items.length === 0) {
    return;
  }

  const workerCount = Math.min(Math.max(limit, 1), items.length);
  let cursor = 0;

  async function worker() {
    while (true) {
      const index = cursor;
      cursor += 1;
      if (index >= items.length) {
        break;
      }
      try {
        await task(items[index], index);
      } catch (error) {
      }
    }
  }

  await Promise.all(Array.from({ length: workerCount }, worker));
}

function isActiveToolingRenderToken(token) {
  return token === currentToolingRenderToken;
}

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

function initStepsInfoModal() {
  const overlay = document.getElementById('stepsInfoOverlay');
  const closeButtons = document.querySelectorAll('[data-steps-info-close]');

  stepsInfoElements = {
    overlay,
    closeButtons
  };

  if (!overlay) {
    return;
  }

  overlay.addEventListener('click', (event) => {
    if (event.target === overlay) {
      closeStepsInfoModal();
    }
  });

  if (closeButtons && closeButtons.length > 0) {
    closeButtons.forEach((button) => {
      button.addEventListener('click', () => {
        closeStepsInfoModal();
      });
    });
  }
}

function openStepsInfoModal(event) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }
  const { overlay, closeButtons } = stepsInfoElements;
  if (!overlay) {
    return;
  }
  overlay.classList.add('active');
  requestAnimationFrame(() => {
    closeButtons?.[0]?.focus();
  });
}

function closeStepsInfoModal() {
  const { overlay } = stepsInfoElements;
  if (!overlay) {
    return;
  }
  overlay.classList.remove('active');
}

function handleStepsInfoIconKey(event) {
  if (event.key === 'Enter' || event.key === ' ') {
    openStepsInfoModal(event);
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

// Função utilitária para encontrar o container do card (normal ou spreadsheet)
function getCardContainer(cardIndex) {
  // Tenta encontrar o card normal primeiro
  let card = document.getElementById(`card-${cardIndex}`);
  
  // Se não encontrar, tenta no spreadsheet expandido
  if (!card) {
    card = document.querySelector(`.spreadsheet-card-container[data-item-index="${cardIndex}"]`);
  }
  
  return card;
}

async function rebuildReplacementPickerOptions(cardIndex) {
  await ensureReplacementIdOptions();
  const card = getCardContainer(cardIndex);
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
  restoreSidebarState();

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
  const addOwnerInput = document.getElementById('addToolingOwner');
  const addOwnerList = document.getElementById('addToolingOwnerList');
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
  const replacementGridCanvas = document.getElementById('replacementGridCanvas');
  const replacementConnectionsCanvas = document.getElementById('replacementConnectionsCanvas');
  const statusOptionsList = document.getElementById('statusOptionsList');
  const statusOptionInput = document.getElementById('statusOptionInput');
  const addStatusButton = document.getElementById('addStatusButton');
  const supplierSearchInput = document.getElementById('supplierSearchInput');
  const clearSupplierSearchBtn = document.getElementById('clearSupplierSearch');
  const expirationFilterSwitch = document.getElementById('expirationFilterSwitch');
  const supplierMenuBtn = document.getElementById('supplierMenuBtn');
  const supplierFilterOverlay = document.getElementById('supplierFilterOverlay');
  const commentDeleteOverlay = document.getElementById('commentDeleteOverlay');
  const commentDeleteContext = document.getElementById('commentDeleteContext');
  const commentDeleteDate = document.getElementById('commentDeleteDate');
  const commentDeleteText = document.getElementById('commentDeleteText');
  const commentDeleteConfirmBtn = document.getElementById('commentDeleteConfirmBtn');

  addToolingElements = {
    overlay: addOverlay,
    form: addForm,
    pnInput: addPN,
    supplierInput: addSupplierInput,
    supplierList: addSupplierList,
    ownerInput: addOwnerInput,
    ownerList: addOwnerList,
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

  commentDeleteElements = {
    overlay: commentDeleteOverlay,
    context: commentDeleteContext,
    date: commentDeleteDate,
    text: commentDeleteText,
    confirmButton: commentDeleteConfirmBtn
  };

  replacementTimelineElements = {
    overlay: replacementTimelineOverlay,
    list: replacementTimelineList,
    empty: replacementTimelineEmpty,
    loading: replacementTimelineLoading,
    title: replacementTimelineTitle,
    gridCanvas: replacementGridCanvas,
    connectionsCanvas: replacementConnectionsCanvas
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
  initStepsInfoModal();
  initSettingsCarouselScrollListener();
  initThousandsMaskBehavior();
  initDevToolsSwitch();

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

  if (commentDeleteOverlay) {
    commentDeleteOverlay.addEventListener('click', (e) => {
      if (e.target === commentDeleteOverlay) {
        closeCommentDeleteModal();
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
    globalSearchDebouncedHandler = debounce((term) => {
      activeSearchTerm = term;
      searchTooling(term);
    }, 250);

    searchInput.addEventListener('input', (e) => {
      const searchTerm = e.target.value.trim();
      if (searchTerm.length >= 2) {
        activeSearchTerm = searchTerm;
        globalSearchDebouncedHandler(searchTerm);
      } else if (searchTerm.length === 0) {
        globalSearchDebouncedHandler?.cancel?.();
        clearSearch({ keepOverlayOpen: true });
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

  // Listener para a barra de pesquisa de suppliers (agora na status bar)
  const statusSupplierSearchInput = document.getElementById('statusSupplierSearchInput');
  if (statusSupplierSearchInput) {
    supplierSearchDebouncedHandler = debounce((term) => {
      filterSuppliersAndTooling(term);
    }, 200);

    statusSupplierSearchInput.addEventListener('input', (e) => {
      const searchTerm = e.target.value.trim();
      const clearBtn = document.getElementById('statusSearchClear');
      if (clearBtn) {
        clearBtn.style.display = searchTerm ? 'flex' : 'none';
      }
      supplierSearchDebouncedHandler(searchTerm);
    });
  }

  // Listener para o switch de filtro de expiração
  if (expirationFilterSwitch) {
    expirationFilterSwitch.addEventListener('change', async (e) => {
      await toggleExpirationFilter(e.target.checked);
      // Atualizar visual do botão de filtro
      const filterBtn = document.getElementById('expirationFilterBtn');
      if (filterBtn) {
        filterBtn.classList.toggle('active', e.target.checked);
      }
    });
  }

  // Listener para o botão de filtro de expiração - abre o modal
  const expirationFilterBtn = document.getElementById('expirationFilterBtn');
  const expirationFilterOverlay = document.getElementById('expirationFilterOverlay');
  if (expirationFilterBtn && expirationFilterOverlay) {
    expirationFilterBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      openExpirationFilterOverlay();
    });

    // Fechar overlay ao clicar no fundo
    expirationFilterOverlay.addEventListener('click', (e) => {
      if (e.target === expirationFilterOverlay) {
        closeExpirationFilterOverlay();
      }
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
      
      // IMPORTANTE: Ignora campos da spreadsheet (linha fechada)
      // pois eles são salvos pela função spreadsheetSave
      const isSpreadsheetField = target.classList.contains('spreadsheet-input') || 
                                  target.classList.contains('spreadsheet-select');
      if (isSpreadsheetField) {
        return; // Não chama autoSaveTooling para campos da spreadsheet
      }
      
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
  globalSearchDebouncedHandler?.cancel?.();

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

// Fecha popup de filtro de comentários ao clicar fora
document.addEventListener('click', (e) => {
  if (!e.target.closest('.comments-filter-wrapper')) {
    document.querySelectorAll('.comments-filter-popup.active').forEach(popup => {
      popup.classList.remove('active');
    });
  }
});

document.addEventListener('click', (event) => {
  if (event.target.closest('[data-replacement-picker]')) {
    return;
  }
  closeAllReplacementPickers();
  closeReplacementPickerOverlay();
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
  const commentOverlay = commentDeleteElements.overlay || document.getElementById('commentDeleteOverlay');
  if (commentOverlay && commentOverlay.classList.contains('active') && e.key === 'Escape') {
    closeCommentDeleteModal();
  }
  const expirationOverlay = expirationInfoElements.overlay;
  if (expirationOverlay && expirationOverlay.classList.contains('active') && e.key === 'Escape') {
    closeExpirationInfoModal();
  }
  const productionOverlay = productionInfoElements.overlay;
  if (productionOverlay && productionOverlay.classList.contains('active') && e.key === 'Escape') {
    closeProductionInfoModal();
  }
  const stepsOverlay = stepsInfoElements.overlay;
  if (stepsOverlay && stepsOverlay.classList.contains('active') && e.key === 'Escape') {
    closeStepsInfoModal();
  }
  const timelineOverlay = replacementTimelineElements.overlay;
  if (timelineOverlay && timelineOverlay.classList.contains('active') && e.key === 'Escape') {
    closeReplacementTimelineOverlay();
  }
  if (e.key === 'Escape') {
    closeAllReplacementPickers();
    closeReplacementPickerOverlay();
    // Fechar barra de pesquisa do status bar
    const statusSearchWrapper = document.getElementById('statusSearchInputWrapper');
    if (statusSearchWrapper && statusSearchWrapper.classList.contains('active')) {
      clearStatusSearch();
    }
  }
});

// Fechar barra de pesquisa ao clicar fora
document.addEventListener('click', (e) => {
  const statusSearchWrapper = document.getElementById('statusSearchInputWrapper');
  const statusSearchBtn = document.getElementById('statusSearchBtn');
  
  if (statusSearchWrapper && statusSearchWrapper.classList.contains('active')) {
    if (!statusSearchWrapper.contains(e.target) && e.target !== statusSearchBtn && !statusSearchBtn.contains(e.target)) {
      const input = document.getElementById('statusSupplierSearchInput');
      if (input && !input.value.trim()) {
        statusSearchWrapper.classList.remove('active');
      }
    }
  }
});

// Carrega fornecedores com estatísticas
async function loadSuppliers() {
  try {
    suppliersData = await window.api.getSuppliersWithStats();
    applyExpirationFilter();
    populateAddToolingSuppliers();
    await loadResponsibles();
  } catch (error) {
    showNotification('Error loading suppliers', 'error');
  }
}

// Carrega lista de responsáveis únicos
let responsiblesData = [];
async function loadResponsibles() {
  try {
    responsiblesData = await window.api.getUniqueResponsibles();
    populateAddToolingOwners();
    if (Array.isArray(toolingData) && toolingData.length > 0) {
      populateCardDataLists(toolingData);
    }
  } catch (error) {
    console.error('Error loading responsibles:', error);
  }
}

function applyExpirationFilter() {
  if (!suppliersData) {
    displaySuppliers([]);
    return;
  }

  if (expirationFilterEnabled) {
    const filteredSuppliers = suppliersData.filter(supplier => {
      return ExpirationMetrics.hasCriticalItems(supplier);
    });
    displaySuppliers(filteredSuppliers);
  } else {
    displaySuppliers(suppliersData);
  }
}

async function toggleExpirationFilter(enabled) {
  expirationFilterEnabled = enabled;
  const badge = document.getElementById('filterActiveBadge');
  if (badge) {
    badge.style.display = enabled ? 'flex' : 'none';
  }
  
  // Recarrega suppliers do banco para atualizar contagens (considera análise concluída)
  try {
    suppliersData = await window.api.getSuppliersWithStats();
  } catch (e) {
    console.error('Erro ao recarregar suppliers:', e);
  }
  
  applyExpirationFilter();
  
  // Recarrega os cards do supplier selecionado para aplicar o filtro
  // A função displayTooling vai recalcular as métricas com os dados originais
  if (selectedSupplier) {
    await loadToolingBySupplier(selectedSupplier);
  }
}

async function clearExpirationFilter() {
  const filterSwitch = document.getElementById('expirationFilterSwitch');
  const filterBtn = document.getElementById('expirationFilterBtn');
  if (filterSwitch) {
    filterSwitch.checked = false;
  }
  if (filterBtn) {
    filterBtn.classList.remove('active');
  }
  await toggleExpirationFilter(false);
}

function openExpirationFilterOverlay() {
  const overlay = document.getElementById('expirationFilterOverlay');
  if (overlay) {
    overlay.classList.add('active');
  }
}

function closeExpirationFilterOverlay() {
  const overlay = document.getElementById('expirationFilterOverlay');
  if (overlay) {
    overlay.classList.remove('active');
  }
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

// Export Annual Volume for Supplier (ID, PN, Supplier, Annual Volume, Annual Volume Date)
async function exportForecastSupplier() {
  try {
    const result = await window.api.exportForecastSupplier();
    if (result.success) {
      showNotification(`Supplier forecast exported successfully to: ${result.filePath}`);
      closeSupplierFilterOverlay();
    } else if (!result.cancelled) {
      showNotification('Failed to export data', 'error');
    }
  } catch (error) {
    showNotification('Error exporting data', 'error');
  }
}

// Import Annual Volume for Supplier (updates Annual Volume and Annual Volume Date by ID)
async function importForecastSupplier() {
  try {
    const result = await window.api.importForecastSupplier();
    if (result.success) {
      showNotification(`Successfully updated ${result.updatedCount} records`);
      closeSupplierFilterOverlay();
      
      // Reload data
      await loadSuppliers();
      await loadAnalytics();
      if (selectedSupplier) {
        await loadToolingBySupplier(selectedSupplier);
      }
    } else if (result.cancelled) {
      // User cancelled
    } else {
      showNotification(result.error || 'Failed to import data', 'error');
    }
  } catch (error) {
    showNotification('Error importing data', 'error');
  }
}

// Export Full Database for Manager
async function exportForecastManager() {
  try {
    const result = await window.api.exportForecastManager();
    if (result.success) {
      showNotification(`Manager database exported successfully to: ${result.filePath}`);
      closeSupplierFilterOverlay();
    } else if (!result.cancelled) {
      showNotification('Failed to export data', 'error');
    }
  } catch (error) {
    showNotification('Error exporting data', 'error');
  }
}

// Import Full Database for Manager
async function importForecastManager() {
  try {
    const result = await window.api.importForecastManager();
    if (result.success) {
      showNotification(`Successfully updated ${result.updatedCount} records`);
      closeSupplierFilterOverlay();
      
      // Reload data
      await loadSuppliers();
      await loadAnalytics();
      if (selectedSupplier) {
        await loadToolingBySupplier(selectedSupplier);
      }
    } else if (result.cancelled) {
      // User cancelled
    } else {
      showNotification(result.error || 'Failed to import data', 'error');
    }
  } catch (error) {
    showNotification('Error importing data', 'error');
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
  const commentBtn = document.getElementById('supplierCommentBtn');
  const selectModeBtn = document.getElementById('toggleSelectModeBtn');
  
  if (exportBtn) exportBtn.disabled = !enabled;
  if (importBtn) importBtn.disabled = !enabled;
  if (commentBtn) commentBtn.disabled = !enabled;
  if (selectModeBtn) selectModeBtn.disabled = !enabled;
  
  // Se desabilitado, desativar modo de seleção
  if (!enabled && selectionModeActive) {
    toggleSelectionMode();
  }
}

// Abrir modal de comentários do supplier
function openSupplierComments() {
  if (!currentSupplier) {
    showNotification('Please select a supplier first', 'error');
    return;
  }
  
  const modal = document.getElementById('supplierCommentsModal');
  const subtitle = document.getElementById('supplierCommentsSubtitle');
  
  if (subtitle) {
    subtitle.textContent = currentSupplier;
  }
  
  // Load existing comments/tasks from localStorage
  loadSupplierCommentsData();
  
  if (modal) {
    modal.classList.add('active');
  }
}

function closeSupplierCommentsModal() {
  const modal = document.getElementById('supplierCommentsModal');
  if (modal) {
    modal.classList.remove('active');
  }
}

function loadSupplierCommentsData() {
  if (!currentSupplier) {
    return;
  }

  const notesTextarea = document.getElementById('supplierNotesText');
  const contactTextarea = document.getElementById('supplierContactText');
  const supplyContinuityInput = document.getElementById('supplyContinuityText');
  const sqieInput = document.getElementById('sqieText');
  const plannerInput = document.getElementById('plannerText');
  const sourcingInput = document.getElementById('sourcingText');
  
  if (!notesTextarea || !contactTextarea || !supplyContinuityInput || !sqieInput || !plannerInput || !sourcingInput) {
    return;
  }

  const storageKey = `supplier_comments_${currentSupplier}`;
  const rawValue = localStorage.getItem(storageKey);
  let savedData = { notes: '', contact: '', supplyContinuity: '', sqie: '', planner: '', sourcing: '' };

  if (rawValue) {
    try {
      const parsed = JSON.parse(rawValue);
      if (typeof parsed === 'string') {
        savedData.notes = parsed;
      } else if (parsed && typeof parsed === 'object') {
        savedData.notes = parsed.notes || '';
        savedData.contact = parsed.contact || '';
        // Migrar campo antigo 'responsible' para 'supplyContinuity' se existir
        savedData.supplyContinuity = parsed.supplyContinuity || parsed.responsible || '';
        savedData.sqie = parsed.sqie || '';
        savedData.planner = parsed.planner || '';
        savedData.sourcing = parsed.sourcing || '';
      }
    } catch (error) {
      savedData.notes = rawValue;
    }
  }

  notesTextarea.value = savedData.notes;
  contactTextarea.value = savedData.contact;
  supplyContinuityInput.value = savedData.supplyContinuity;
  sqieInput.value = savedData.sqie;
  plannerInput.value = savedData.planner;
  sourcingInput.value = savedData.sourcing;
}

function saveSupplierComments() {
  if (!currentSupplier) {
    return;
  }

  const notesTextarea = document.getElementById('supplierNotesText');
  const contactTextarea = document.getElementById('supplierContactText');
  const supplyContinuityInput = document.getElementById('supplyContinuityText');
  const sqieInput = document.getElementById('sqieText');
  const plannerInput = document.getElementById('plannerText');
  const sourcingInput = document.getElementById('sourcingText');
  
  if (!notesTextarea || !contactTextarea || !supplyContinuityInput || !sqieInput || !plannerInput || !sourcingInput) {
    return;
  }

  const storageKey = `supplier_comments_${currentSupplier}`;
  const data = {
    notes: notesTextarea.value || '',
    contact: contactTextarea.value || '',
    supplyContinuity: supplyContinuityInput.value || '',
    sqie: sqieInput.value || '',
    planner: plannerInput.value || '',
    sourcing: sourcingInput.value || ''
  };

  localStorage.setItem(storageKey, JSON.stringify(data));
  showNotification('Supplier comments saved successfully!', 'success');
  closeSupplierCommentsModal();
}

// ===== MODO DE SELEÇÃO MÚLTIPLA PARA EXPORTAÇÃO =====

// Alternar modo de seleção
function toggleSelectionMode() {
  selectionModeActive = !selectionModeActive;
  const selectBtn = document.getElementById('toggleSelectModeBtn');
  const selectionActions = document.getElementById('selectionActions');
  const toolingList = document.getElementById('toolingList');
  const exportBtn = document.getElementById('exportSupplierBtn');
  const spreadsheetCheckboxHeader = document.getElementById('spreadsheetCheckboxHeader');
  
  if (selectionModeActive) {
    // Ativar modo de seleção
    selectedToolingIds.clear();
    if (selectBtn) {
      selectBtn.classList.add('active');
      selectBtn.title = 'Cancel selection';
    }
    if (selectionActions) selectionActions.style.display = 'flex';
    if (toolingList) toolingList.classList.add('selection-mode');
    if (exportBtn) exportBtn.style.display = 'none';
    if (spreadsheetCheckboxHeader) spreadsheetCheckboxHeader.style.display = '';
    updateSelectionCounter();
    
    // Re-renderizar spreadsheet se estiver no modo tabela
    if (currentViewMode === 'spreadsheet') {
      renderSpreadsheetView();
    }
  } else {
    // Desativar modo de seleção
    selectedToolingIds.clear();
    if (selectBtn) {
      selectBtn.classList.remove('active');
      selectBtn.title = 'Select items for export';
    }
    if (selectionActions) selectionActions.style.display = 'none';
    if (toolingList) toolingList.classList.remove('selection-mode');
    if (exportBtn) exportBtn.style.display = '';
    if (spreadsheetCheckboxHeader) spreadsheetCheckboxHeader.style.display = 'none';
    
    // Re-renderizar spreadsheet se estiver no modo tabela
    if (currentViewMode === 'spreadsheet') {
      renderSpreadsheetView();
    }
  }
  
  // Atualizar checkboxes nos cards
  updateCardCheckboxes();
}

// Atualizar checkboxes visuais nos cards
function updateCardCheckboxes() {
  const cards = document.querySelectorAll('.tooling-card');
  
  cards.forEach(card => {
    const itemId = parseInt(card.dataset.itemId);
    let checkbox = card.querySelector('.card-selection-checkbox');
    
    if (selectionModeActive) {
      if (!checkbox) {
        // Criar checkbox
        checkbox = document.createElement('div');
        checkbox.className = 'card-selection-checkbox';
        checkbox.innerHTML = '<i class="ph ph-check"></i>';
        checkbox.onclick = (e) => {
          e.stopPropagation();
          toggleCardSelection(itemId);
        };
        // Inserir dentro do header-top, antes do primeiro elemento
        const headerTop = card.querySelector('.tooling-card-header-top');
        if (headerTop) {
          headerTop.insertBefore(checkbox, headerTop.firstChild);
        }
      }
      // Atualizar estado
      checkbox.classList.toggle('selected', selectedToolingIds.has(itemId));
    } else {
      // Remover checkbox
      if (checkbox) {
        checkbox.remove();
      }
    }
  });
}

// Alternar seleção de um card
function toggleCardSelection(itemId) {
  if (selectedToolingIds.has(itemId)) {
    selectedToolingIds.delete(itemId);
  } else {
    selectedToolingIds.add(itemId);
  }
  
  // Atualizar visual do checkbox
  const card = document.querySelector(`.tooling-card[data-item-id="${itemId}"]`);
  if (card) {
    const checkbox = card.querySelector('.card-selection-checkbox');
    if (checkbox) {
      checkbox.classList.toggle('selected', selectedToolingIds.has(itemId));
    }
  }
  
  updateSelectionCounter();
}

// Selecionar todos os itens visíveis
function selectAllTooling() {
  toolingData.forEach(item => {
    selectedToolingIds.add(item.id);
  });
  updateCardCheckboxes();
  updateSpreadsheetCheckboxes();
  updateSelectionCounter();
}

// Desselecionar todos
function deselectAllTooling() {
  selectedToolingIds.clear();
  updateCardCheckboxes();
  updateSpreadsheetCheckboxes();
  updateSelectionCounter();
}

// ===== FUNÇÕES DE SELEÇÃO PARA SPREADSHEET =====

// Alternar seleção de uma linha da tabela
function toggleSpreadsheetRowSelection(event, itemId) {
  event.stopPropagation();
  
  if (selectedToolingIds.has(itemId)) {
    selectedToolingIds.delete(itemId);
  } else {
    selectedToolingIds.add(itemId);
  }
  
  updateSpreadsheetRowVisual(itemId);
  updateSpreadsheetSelectAllIcon();
  updateSelectionCounter();
}

// Atualizar visual de uma linha específica
function updateSpreadsheetRowVisual(itemId) {
  const row = document.querySelector(`#spreadsheetBody tr[data-id="${itemId}"]`);
  if (row) {
    const isSelected = selectedToolingIds.has(itemId);
    row.classList.toggle('row-selected', isSelected);
    const checkbox = row.querySelector('.spreadsheet-row-checkbox');
    if (checkbox) {
      checkbox.classList.toggle('selected', isSelected);
      const icon = checkbox.querySelector('i');
      if (icon) {
        icon.className = isSelected ? 'ph ph-check-square' : 'ph ph-square';
      }
    }
  }
}

// Toggle selecionar todos na tabela
function toggleSpreadsheetSelectAll() {
  const allSelected = toolingData.every(item => selectedToolingIds.has(item.id));
  
  if (allSelected) {
    // Desselecionar todos
    selectedToolingIds.clear();
  } else {
    // Selecionar todos
    toolingData.forEach(item => {
      selectedToolingIds.add(item.id);
    });
  }
  
  updateSpreadsheetCheckboxes();
  updateCardCheckboxes();
  updateSelectionCounter();
}

// Atualizar ícone de selecionar todos
function updateSpreadsheetSelectAllIcon() {
  const icon = document.getElementById('spreadsheetSelectAllIcon');
  if (icon) {
    const allSelected = toolingData.length > 0 && toolingData.every(item => selectedToolingIds.has(item.id));
    const someSelected = selectedToolingIds.size > 0;
    
    if (allSelected) {
      icon.className = 'ph ph-check-square';
    } else if (someSelected) {
      icon.className = 'ph ph-minus-square';
    } else {
      icon.className = 'ph ph-square';
    }
  }
}

// Atualizar todos os checkboxes da tabela
function updateSpreadsheetCheckboxes() {
  const rows = document.querySelectorAll('#spreadsheetBody tr[data-id]');
  rows.forEach(row => {
    const itemId = parseInt(row.dataset.id);
    if (!isNaN(itemId)) {
      const isSelected = selectedToolingIds.has(itemId);
      row.classList.toggle('row-selected', isSelected);
      const checkbox = row.querySelector('.spreadsheet-row-checkbox');
      if (checkbox) {
        checkbox.classList.toggle('selected', isSelected);
        const icon = checkbox.querySelector('i');
        if (icon) {
          icon.className = isSelected ? 'ph ph-check-square' : 'ph ph-square';
        }
      }
    }
  });
  updateSpreadsheetSelectAllIcon();
}

// Atualizar contador de seleção
function updateSelectionCounter() {
  const counter = document.getElementById('selectionCounter');
  if (counter) {
    const count = selectedToolingIds.size;
    counter.textContent = count === 1 ? '1 item selected' : `${count} items selected`;
  }
}

// Exportar apenas os selecionados
async function exportSelectedItems() {
  if (selectedToolingIds.size === 0) {
    showToast('Select at least one item to export', 'error');
    return;
  }
  
  if (!currentSupplier) {
    showToast('Please select a supplier first', 'error');
    return;
  }

  try {
    showToast('Exporting selected items...', 'info');
    
    const idsToExport = Array.from(selectedToolingIds);
    const result = await window.api.exportSupplierData(currentSupplier, idsToExport);
    
    if (result.success) {
      showToast(`${idsToExport.length} items exported successfully!`, 'success');
      // Desativar modo de seleção após exportar
      toggleSelectionMode();
    } else if (result.cancelled) {
      showToast('Export cancelled', 'info');
    } else {
      showToast('Export cancelled', 'info');
    }
  } catch (error) {
    showToast('Error exporting data. Please try again.', 'error');
  }
}

// Exportar dados do supplier para Excel
async function exportSupplierData() {
  if (!currentSupplier) {
    showToast('Please select a supplier first', 'error');
    return;
  }

  try {
    showToast('Exporting supplier data...', 'info');
    
    // Obter IDs dos cards atualmente visíveis/filtrados
    const filteredIds = toolingData.map(item => item.id);
    
    const result = await window.api.exportSupplierData(currentSupplier, filteredIds);
    
    if (result.success) {
      showToast('Data exported successfully!', 'success');
    } else if (result.cancelled) {
      showToast('Export cancelled', 'info');
    } else {
      showToast('Export cancelled', 'info');
    }
  } catch (error) {
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
    const successMsg = updated === 1
      ? '1 registro atualizado'
      : `${updated} registros atualizados`;
    showToast(successMsg, 'success');

    await loadToolingBySupplier(currentSupplier);
    await loadSuppliers();
  } catch (error) {
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
    const supplierKey = encodeURIComponent(supplierNameRaw.trim().toLowerCase());
    const isActive = selectedSupplier === supplierNameRaw;
    // Usa APENAS ExpirationMetrics.fromItems() para calcular métricas
    const metrics = ExpirationMetrics.fromItems(supplier.items || []);
    
    // Verificar se este supplier deve ser escondido pelo filtro de steps
    const isHiddenByStepsFilter = stepsFilteredSuppliers !== null && !stepsFilteredSuppliers.includes(supplierNameRaw);
    
    return `
        <div class="supplier-card ${isActive ? 'active' : ''}" 
          data-supplier="${supplierNameHtml}" 
          data-supplier-raw="${supplierNameHtml}"
          data-supplier-key="${supplierKey}"
          style="${isHiddenByStepsFilter ? 'display: none;' : ''}"
          onclick="selectSupplier(event, '${supplierNameForHandler}')">
      <div class="supplier-card-header">
        <i class="ph ph-factory"></i>
        <h4 title="${supplierNameHtml}">${supplierNameHtml}</h4>
      </div>
      <div class="supplier-info">
        <div class="supplier-info-row">
          <span class="info-label">Total Tooling:</span>
          <span class="info-value" data-metric="total">${metrics.total}</span>
        </div>
        <div class="supplier-info-row">
          <span class="info-label">Expired:</span>
          <span class="info-value ${metrics.expired > 0 ? 'expired' : ''}" data-metric="expired">${metrics.expired}</span>
        </div>
        <div class="supplier-info-row">
          <span class="info-label">Expiring within 2 years:</span>
          <span class="info-value ${metrics.expiring > 0 ? 'critical' : ''}" data-metric="expiring">${metrics.expiring}</span>
        </div>
      </div>
      <div class="compact-tooltip">
        <div class="tooltip-name">${supplierNameHtml}</div>
        <div class="tooltip-stats">
          <div class="tooltip-stat"><i class="ph ph-wrench"></i> Total: <span>${metrics.total}</span></div>
          <div class="tooltip-stat"><i class="ph ph-warning-circle"></i> Expired: <span>${metrics.expired}</span></div>
          <div class="tooltip-stat"><i class="ph ph-clock-countdown"></i> Expiring 2y: <span>${metrics.expiring}</span></div>
        </div>
      </div>
    </div>
    `;
  }).join('');

  // Adiciona event listeners para posicionar tooltips no modo compacto
  setupCompactTooltips();
  
  syncStatusBarWithSuppliers();
}

// Configura tooltips para o modo compacto (responsivo)
function setupCompactTooltips() {
  const supplierCards = document.querySelectorAll('.supplier-card');
  
  supplierCards.forEach(card => {
    const tooltip = card.querySelector('.compact-tooltip');
    if (!tooltip) return;
    
    card.addEventListener('mouseenter', (e) => {
      const rect = card.getBoundingClientRect();
      tooltip.style.top = `${rect.top}px`;
    });
  });
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

function handleSupplierSelection(supplierName, { sourceElement = null, forceReload = false } = {}) {
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
  
  // Habilitar botões de exportar/importar IMEDIATAMENTE
  updateSupplierDataButtons(true);
  
  // Mostra área de tooling IMEDIATAMENTE (vazia)
  const toolingList = document.getElementById('toolingList');
  const emptyState = document.getElementById('emptyState');
  if (toolingList) {
    toolingList.style.display = 'flex';
    toolingList.innerHTML = '';
  }
  if (emptyState) {
    emptyState.style.display = 'none';
  }

  const shouldReload = forceReload || previousSupplier !== normalizedName;
  if (shouldReload) {
    // Carrega tudo em background DEPOIS da UI responder
    setTimeout(() => {
      loadAttachments(normalizedName).catch(err => {});
      
      // Verifica AMBOS campos de pesquisa
      const searchInput = document.getElementById('searchInput');
      const supplierSearchInput = document.getElementById('supplierSearchInput');
      const globalSearchValue = searchInput ? searchInput.value.trim() : '';
      const supplierSearchValue = supplierSearchInput ? supplierSearchInput.value.trim() : '';
      
      const hasGlobalSearch = globalSearchValue.length >= 2;
      const hasSupplierSearch = supplierSearchValue.length >= 1;
      
      // Pegar o filtro de steps se estiver ativo
      const stepsFilter = document.getElementById('stepsFilter');
      const selectedStep = stepsFilter ? stepsFilter.value : '';
      
      if (hasGlobalSearch) {
        // Usa busca global que já filtra por supplier selecionado
        activeSearchTerm = globalSearchValue;
        searchTooling(activeSearchTerm).catch(err => {});
      } else if (hasSupplierSearch) {
        // Filtra por termo da barra lateral E supplier selecionado
        window.api.searchTooling(supplierSearchValue).then(allResults => {
          let filteredResults = allResults.filter(item => {
            const itemSupplier = String(item.supplier || '').trim();
            return itemSupplier === normalizedName;
          });
          
          // Aplicar filtro de steps se estiver ativo
          if (selectedStep) {
            filteredResults = filteredResults.filter(item => {
              const itemStep = String(item.steps || '').trim();
              return itemStep === selectedStep;
            });
          }
          
          toolingData = filteredResults;
          displayTooling(filteredResults);
        }).catch(err => {});
      } else {
        // Sem nenhuma busca, carrega todos do supplier
        loadToolingBySupplier(normalizedName).catch(err => {});
      }
    }, 0);
  }
}

// Seleciona fornecedor e exibe ferramentais
function selectSupplier(evt, supplierName) {
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
    const spreadsheetContainer = document.getElementById('spreadsheetContainer');
    const emptyState = document.getElementById('emptyState');
    if (toolingList) toolingList.style.display = 'none';
    if (spreadsheetContainer) spreadsheetContainer.style.display = 'none';
    if (emptyState) emptyState.style.display = 'flex';
    
    // Desabilitar botões de exportar/importar
    updateSupplierDataButtons(false);
    
    return;
  }
  
  handleSupplierSelection(supplierName, { sourceElement: cardElement, forceReload: true });
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
    showNotification('Error attaching files.', 'error');
  }
}

// Initialize drag and drop for card attachment dropzones
function initCardAttachmentDragAndDrop(dropzoneElement, itemId) {
  if (!dropzoneElement) return;

  dropzoneElement.addEventListener('dragenter', (e) => handleCardDropzoneDragEnter(e, dropzoneElement), false);
  dropzoneElement.addEventListener('dragover', (e) => handleCardDropzoneDragOver(e, dropzoneElement), false);
  dropzoneElement.addEventListener('dragleave', (e) => handleCardDropzoneDragLeave(e, dropzoneElement), false);
  dropzoneElement.addEventListener('drop', (e) => handleCardDropzoneDrop(e, dropzoneElement, itemId), false);
}

function handleCardDropzoneDragEnter(event, dropzone) {
  if (!isFileDrag(event)) return;
  event.preventDefault();
  event.stopPropagation();
  dropzone.classList.add('drop-active');
}

function handleCardDropzoneDragOver(event, dropzone) {
  if (!isFileDrag(event)) return;
  event.preventDefault();
  event.stopPropagation();
  if (event.dataTransfer) {
    event.dataTransfer.dropEffect = 'copy';
  }
  dropzone.classList.add('drop-active');
}

function handleCardDropzoneDragLeave(event, dropzone) {
  if (!isFileDrag(event)) return;
  event.preventDefault();
  event.stopPropagation();
  
  // Only remove class if leaving the dropzone itself
  if (event.target === dropzone) {
    dropzone.classList.remove('drop-active');
  }
}

async function handleCardDropzoneDrop(event, dropzone, itemId) {
  if (!isFileDrag(event)) return;
  
  event.preventDefault();
  event.stopPropagation();
  dropzone.classList.remove('drop-active');

  const files = Array.from(event.dataTransfer?.files || []);
  if (!files.length) return;

  await handleCardAttachmentFiles(files, itemId);
}

async function handleCardAttachmentFiles(files, itemId) {
  if (!files || files.length === 0) return;
  // Try to get item from toolingData first
  let { normalizedId, toolingItem } = findToolingItem(itemId);
  // If not found in array, try to get from card DOM
  if (normalizedId === null || !toolingItem) {
    const card = document.querySelector(`.tooling-card[data-item-id="${itemId}"]`);
    if (card) {
      // Always normalize the ID from DOM attribute
      normalizedId = normalizeItemId(card.dataset.itemId || itemId);

      // Get supplier from data attribute or fields
      const supplierAttr = card.getAttribute('data-supplier');
      const supplierField = card.querySelector('[data-field="supplier"]');
      const supplier = (supplierField?.value || supplierAttr || '').trim();
      if (supplier) {
        toolingItem = { id: normalizedId, supplier };
      }
    }
  }
  
  if (normalizedId === null || !toolingItem) {
    showNotification('Invalid tooling item.', 'error');
    return;
  }

  if (!toolingItem.supplier) {
    showNotification('Supplier not found for this tooling.', 'error');
    return;
  }
  const paths = files.map(file => file?.path).filter(Boolean);
  if (paths.length === 0) {
    showNotification('Unable to read the dragged files.', 'error');
    return;
  }

  try {
    const result = await window.api.uploadAttachmentFromPaths(toolingItem.supplier, paths, normalizedId);

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

    await loadCardAttachments(normalizedId);
  } catch (error) {
    showNotification('Error attaching files.', 'error');
  }
}

// Carrega anexos de um fornecedor
async function loadAttachments(supplierName) {
  try {
    const attachments = await window.api.getAttachments(supplierName);
    displayAttachments(attachments);
  } catch (error) {
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

async function openAttachmentsFolder() {
  const supplierLabel = currentSupplier || selectedSupplier;
  if (!supplierLabel) {
    showNotification('Select a supplier first', 'error');
    return;
  }
  try {
    await window.api.openAttachmentsFolder(supplierLabel);
  } catch (error) {
    showNotification('Error opening folder', 'error');
  }
}

// Abre a linha expandida do spreadsheet e foca na aba de anexos
function openToolingAttachmentsFromSpreadsheet(itemId) {
  const itemIndex = toolingData.findIndex(t => String(t.id) === String(itemId));
  if (itemIndex === -1) return;
  
  const row = document.querySelector(`tr[data-id="${itemId}"]`);
  if (!row) return;
  
  const isExpanded = row.classList.contains('row-expanded');
  
  // Se não está expandida, expande primeiro
  if (!isExpanded) {
    toggleSpreadsheetRow(itemId, itemIndex);
  }
  
  // Aguarda um pouco para a animação e depois muda para a aba attachments
  setTimeout(() => {
    // Muda para a aba de attachments
    switchCardTab(itemIndex, 'attachments');
  }, 100);
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
    
    // Se foi cancelado, apenas retorna silenciosamente
    if (!result || result.cancelled) {
      return;
    }
    
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
        }
      }
    } else if (result.error) {
      showNotification('Error attaching file(s)', 'error');
    }
  } catch (error) {
    alert('Error uploading file');
  }
}

// Abre arquivo anexado
async function openAttachment(supplierName, fileName) {
  try {
    await window.api.openAttachment(supplierName, fileName);
  } catch (error) {
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
    alert('Error deleting file');
  }
}

// Abre arquivo anexado do card
async function openCardAttachmentFile(supplierName, fileName, itemId) {
  try {
    await window.api.openAttachment(supplierName, fileName, itemId);
  } catch (error) {
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
    showNotification('Error deleting file', 'error');
  }
}

// Carrega ferramentais por fornecedor
// Armazena IDs que têm incoming links de outros suppliers
let externalIncomingLinks = [];

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

    // Mostrar skeleton loading
    showSkeletonLoading();
    
    // Carregar dados
    let data = await window.api.getToolingBySupplier(supplier);
    await ensureReplacementIdOptions();
    
    // Aplicar filtro de steps se estiver ativo
    const stepsFilter = document.getElementById('stepsFilter');
    const selectedStep = stepsFilter ? stepsFilter.value : '';
    if (selectedStep) {
      data = data.filter(item => {
        const itemStep = String(item.steps || '').trim();
        return itemStep === selectedStep;
      });
    }
    
    toolingData = data;
    
    // Busca IDs que têm incoming links de outros suppliers (são apontados por outros)
    const allIds = data.map(item => String(item.id));
    try {
      externalIncomingLinks = await window.api.getIdsWithIncomingLinks(allIds);
    } catch (e) {
      externalIncomingLinks = [];
    }
    
    // Recalcula todas as expiration dates ao carregar
    await recalculateAllExpirationDates();
    
    // Limpa filtros de coluna e ordenação ao trocar de supplier
    columnFilters = {};
    columnSort = { column: null, direction: null };
    
    // Renderizar em chunks para não travar a UI
    // A função displayTooling já atualiza as métricas do supplier card internamente
    await displayToolingInChunks(toolingData);
  } catch (error) {
    showNotification('Erro ao carregar ferramentais', 'error');
    hideSkeletonLoading();
  }
}

function showSkeletonLoading() {
  const toolingList = document.getElementById('toolingList');
  const emptyState = document.getElementById('emptyState');
  
  if (emptyState) emptyState.style.display = 'none';
  if (!toolingList) return;
  
  toolingList.style.display = 'flex';
  toolingList.innerHTML = Array.from({ length: 6 }, () => `
    <div class="skeleton-card">
      <div class="skeleton-header">
        <div class="skeleton-line skeleton-title"></div>
        <div class="skeleton-line skeleton-badge"></div>
      </div>
      <div class="skeleton-body">
        <div class="skeleton-line skeleton-text"></div>
        <div class="skeleton-line skeleton-text short"></div>
        <div class="skeleton-line skeleton-text"></div>
      </div>
    </div>
  `).join('');
}

function hideSkeletonLoading() {
  const toolingList = document.getElementById('toolingList');
  if (toolingList) {
    toolingList.innerHTML = '';
  }
}

async function displayToolingInChunks(data, chunkSize = 10) {
  // Usar a função displayTooling existente, mas de forma não-bloqueante
  hideSkeletonLoading();
  
  // Pequeno delay para permitir que a UI atualize
  await new Promise(resolve => setTimeout(resolve, 0));
  
  // Chamar displayTooling normalmente
  await displayTooling(data);
}

// Busca ferramentais
async function searchTooling(term) {
  const normalizedTerm = typeof term === 'string' ? term.trim() : '';
  if (!normalizedTerm) {
    displayTooling([]);
    updateSearchIndicators();
    return;
  }

  const requestId = ++globalSearchRequestId;

  try {
    const results = await window.api.searchTooling(normalizedTerm);
    if (requestId !== globalSearchRequestId) {
      return;
    }
    await ensureReplacementIdOptions();
    
    let filteredResults = results;
    if (selectedSupplier) {
      filteredResults = results.filter(item => {
        const itemSupplier = String(item.supplier || '').trim();
        const selected = String(selectedSupplier || '').trim();
        return itemSupplier === selected;
      });
    }
    
    // Aplicar filtro de steps se estiver ativo
    const stepsFilter = document.getElementById('stepsFilter');
    const selectedStep = stepsFilter ? stepsFilter.value : '';
    if (selectedStep) {
      filteredResults = filteredResults.filter(item => {
        const itemStep = String(item.steps || '').trim();
        return itemStep === selectedStep;
      });
    }
    
    // A função displayTooling já atualiza as métricas do supplier card internamente
    await displayTooling(filteredResults);
    updateSearchIndicators();
  } catch (error) {
  }
}

// Calcula status de vencimento
function getExpirationStatus(expirationDate) {
  if (!expirationDate) return { class: 'ok', label: '' };
  
  // Normaliza a string de data para formato ISO (YYYY-MM-DD)
  let normalizedDate = String(expirationDate).trim();
  
  // Se estiver no formato DD/MM/YYYY, converte para YYYY-MM-DD
  const ddmmyyyyMatch = normalizedDate.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmmyyyyMatch) {
    const [, day, month, year] = ddmmyyyyMatch;
    normalizedDate = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
  }
  
  const expDate = new Date(normalizedDate);
  
  // Verifica se a data é válida
  if (Number.isNaN(expDate.getTime())) {
    return { class: 'ok', label: '' };
  }
  
  // Normaliza ambas as datas para meia-noite para comparação precisa
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  expDate.setHours(0, 0, 0, 0);
  
  const diffTime = expDate - now;
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  
  if (diffDays < 0) {
    return { class: 'expired', label: 'Expirado' };
  } else if (diffDays <= 730) {
    return { class: 'warning', label: 'Até 2 anos' };
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
  const card = getCardContainer(cardIndex);
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
  
  const toolingLife = parseLocalizedNumber(item.tooling_life_qty) || 0;
  const produced = parseLocalizedNumber(item.produced) || 0;
  const remaining = toolingLife - produced;
  const percentValue = toolingLife > 0 ? (produced / toolingLife) * 100 : 0;
  const percent = toolingLife > 0 ? percentValue.toFixed(1) : '0.0';
  
  if (remainingInput) {
    remainingInput.value = formatIntegerWithSeparators(remaining, { preserveEmpty: true });
    item.remaining_tooling_life_pcs = remaining;
  }
  
  if (percentInput) {
    percentInput.value = percent + '%';
    item.percent_tooling_life = percent;
  }
  
  // Atualiza o toolingData global
  const toolingIndex = toolingData.findIndex(t => String(t.id) === String(itemId));
  if (toolingIndex !== -1) {
    toolingData[toolingIndex].tooling_life_qty = toolingLife;
    toolingData[toolingIndex].produced = produced;
    toolingData[toolingIndex].remaining_tooling_life_pcs = remaining;
    toolingData[toolingIndex].percent_tooling_life = percent;
  }
  
  // Atualiza barra de progresso interna (aba Data)
  const lifecycleProgressFill = card.querySelector('[data-lifecycle-progress-fill]');

  if (lifecycleProgressFill) {
    lifecycleProgressFill.style.width = percent + '%';
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
  
  // Atualiza a barra de progresso na linha da spreadsheet
  updateSpreadsheetProgressBar(itemId);
  
  // Sincroniza tooling_life_qty e produced com a linha da spreadsheet
  syncSpreadsheetFromExpandedCard(itemId, 'tooling_life_qty', toolingLife);
  syncSpreadsheetFromExpandedCard(itemId, 'produced', produced);
  
  // Recalcula data de expiração
  calculateExpirationDate(cardIndex, item);
  
  // Salva automaticamente
  autoSaveTooling(itemId);
}

// Calcula data de vencimento baseada no annual forecast
// Calcula a data de validade
function calculateExpirationDate(cardIndex, providedItem = null, skipSave = false) {
  const card = getCardContainer(cardIndex);
  if (!card) {
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
  const toolingLife = parseLocalizedNumber(item.tooling_life_qty) || 0;
  const produced = parseLocalizedNumber(item.produced) || 0;
  const remaining = toolingLife - produced;
  
  const expirationInput = card.querySelector(`[data-field="expiration_date"]`);
  if (!expirationInput) {
    return;
  }
  
  const forecast = parseLocalizedNumber(item.annual_volume_forecast) || 0;
  const productionDateValue = item.date_remaining_tooling_life || '';
  
  const formattedDate = calculateExpirationFromFormula({
    remaining,
    forecast,
    productionDate: productionDateValue
  });

  if (formattedDate) {
    expirationInput.value = formattedDate;
    item.expiration_date = formattedDate;
    
    // Atualiza o toolingData global
    const toolingIndex = toolingData.findIndex(t => String(t.id) === String(itemId));
    if (toolingIndex !== -1) {
      toolingData[toolingIndex].expiration_date = formattedDate;
    }
    
    // Atualizar header do card em tempo real
    const cardForUpdate = getCardContainer(cardIndex);
    if (cardForUpdate) {
      const expirationDisplay = cardForUpdate.querySelector('[data-card-expiration]');
      if (expirationDisplay) {
        expirationDisplay.textContent = formatDate(formattedDate);
      }
    }
    
    // Atualiza a célula de expiração na linha da spreadsheet
    syncSpreadsheetExpirationCell(itemId);
    
    // Atualiza os ícones de expiração (linha e card)
    updateExpirationIconsForItem(itemId);
  } else {
    expirationInput.value = '';
    item.expiration_date = '';
    
    // Atualiza o toolingData global
    const toolingIndex = toolingData.findIndex(t => String(t.id) === String(itemId));
    if (toolingIndex !== -1) {
      toolingData[toolingIndex].expiration_date = '';
    }
    
    // Limpar header do card
    const cardForClear = getCardContainer(cardIndex);
    if (cardForClear) {
      const expirationDisplay = cardForClear.querySelector('[data-card-expiration]');
      if (expirationDisplay) {
        expirationDisplay.textContent = '';
      }
    }
    
    // Atualiza a célula de expiração na linha da spreadsheet
    syncSpreadsheetExpirationCell(itemId);
    
    // Atualiza os ícones de expiração (linha e card)
    updateExpirationIconsForItem(itemId);
  }
  
  // Salva automaticamente apenas se não for skipSave
  if (!skipSave) {
    autoSaveTooling(itemId);
  }
}

function handleProducedChange(cardIndex) {
  calculateLifecycle(cardIndex);
  const card = getCardContainer(cardIndex);
  if (card) {
    triggerDateReminder(card, 'date_remaining_tooling_life');
  }
}

function handleProductionDateChange(cardIndex) {
  const card = getCardContainer(cardIndex);
  if (!card) return;
  
  const itemId = card.getAttribute('data-item-id');
  const productionDateInput = card.querySelector(`[data-field="date_remaining_tooling_life"]`);
  
  // Atualiza o toolingData global com o novo valor
  const toolingIndex = toolingData.findIndex(t => String(t.id) === String(itemId));
  const dateValue = productionDateInput ? productionDateInput.value || '' : '';
  if (toolingIndex !== -1 && productionDateInput) {
    toolingData[toolingIndex].date_remaining_tooling_life = dateValue;
  }
  
  // Sincroniza date_remaining_tooling_life com a linha da spreadsheet
  syncSpreadsheetFromExpandedCard(itemId, 'date_remaining_tooling_life', dateValue);
  
  // Recalcula a data de expiração
  calculateExpirationDate(cardIndex);
}

// Atualiza a barra de progresso no spreadsheet (linha e card expandido)
function updateSpreadsheetProgressBar(itemId) {
  if (!itemId) return;
  
  // Encontra o item nos dados
  const item = toolingData.find(t => String(t.id) === String(itemId));
  if (!item) return;
  
  // Calcula o percentual
  const toolingLife = Number(parseLocalizedNumber(item.tooling_life_qty)) || Number(item.tooling_life_qty) || 0;
  const produced = Number(parseLocalizedNumber(item.produced)) || Number(item.produced) || 0;
  const percentUsedValue = toolingLife > 0 ? (produced / toolingLife) * 100 : 0;
  const percentUsed = toolingLife > 0 ? percentUsedValue.toFixed(1) : '0';
  
  // Atualiza a barra na linha do spreadsheet
  const row = document.querySelector(`#spreadsheetBody tr[data-id="${itemId}"]`);
  if (row) {
    const progressFill = row.querySelector('.spreadsheet-progress-fill');
    if (progressFill) {
      progressFill.style.width = percentUsed + '%';
    }
  }
  
  // Atualiza a barra no card expandido se estiver aberto
  const expandedCard = document.querySelector(`.spreadsheet-card-container[data-item-id="${itemId}"]`);
  if (expandedCard) {
    // Barra de progresso interna (aba Data)
    const lifecycleProgressFill = expandedCard.querySelector('[data-lifecycle-progress-fill]');
    if (lifecycleProgressFill) {
      lifecycleProgressFill.style.width = percentUsed + '%';
    }
    
    // Barra de progresso externa (header)
    const externalProgressFill = expandedCard.querySelector('[data-progress-fill]');
    const externalProgressPercent = expandedCard.querySelector('[data-progress-percent]');
    
    if (externalProgressFill) {
      externalProgressFill.style.width = percentUsed + '%';
    }
    
    if (externalProgressPercent) {
      externalProgressPercent.textContent = percentUsed + '%';
    }
  }
}

function handleForecastChange(cardIndex) {
  const card = getCardContainer(cardIndex);
  if (!card) return;
  
  const itemId = card.getAttribute('data-item-id');
  const forecastInput = card.querySelector(`[data-field="annual_volume_forecast"]`);
  const forecastDateInput = card.querySelector(`[data-field="date_annual_volume"]`);
  
  // Atualiza o toolingData global com o novo valor do forecast
  const toolingIndex = toolingData.findIndex(t => String(t.id) === String(itemId));
  const forecastValue = forecastInput ? parseLocalizedNumber(forecastInput.value) || 0 : 0;
  if (toolingIndex !== -1 && forecastInput) {
    toolingData[toolingIndex].annual_volume_forecast = forecastValue;
    if (forecastDateInput) {
      toolingData[toolingIndex].date_annual_volume = forecastDateInput.value || '';
    }
  }
  
  // Sincroniza annual_volume_forecast com a linha da spreadsheet
  syncSpreadsheetFromExpandedCard(itemId, 'annual_volume_forecast', forecastValue);
  
  calculateExpirationDate(cardIndex);
  
  if (card) {
    triggerDateReminder(card, 'date_annual_volume');
  }
}

function handleStatusSelectChange(cardIndex, itemId, selectEl) {
  if (selectEl) {
    const value = (selectEl.value || '').trim();
    const card = getCardContainer(cardIndex);
    const previousStatus = card ? (card.dataset.status || '').trim().toLowerCase() : '';
    const newStatus = value.toLowerCase();
    
    // Se estava obsolete e mudou para outro status, limpar replacement_tooling_id
    if (previousStatus === 'obsolete' && newStatus !== 'obsolete') {
      const replacementInput = card.querySelector('[data-field="replacement_tooling_id"]');
      if (replacementInput) {
        replacementInput.value = '';
      }
    }
    
    updateCardHeaderStatus(cardIndex, value);
    updateCardStatusAttribute(cardIndex, value);
    
    // Sincroniza status com a linha da spreadsheet
    syncSpreadsheetFromExpandedCard(itemId, 'status', value);
  }
  autoSaveTooling(itemId, true);
}

function updateCardHeaderStatus(cardIndex, statusValue) {
  const card = getCardContainer(cardIndex);
  if (!card) {
    return;
  }
  const display = card.querySelector('[data-card-status]');
  if (display) {
    display.textContent = statusValue || 'N/A';
  }
}

function updateCardHeaderSteps(cardIndex, stepsValue) {
  const card = getCardContainer(cardIndex);
  if (!card) {
    return;
  }
  const display = card.querySelector('[data-card-steps]');
  if (display) {
    display.textContent = stepsValue || 'N/A';
  }
}

function getStepDescription(stepValue) {
  const descriptions = {
    '1': 'Control Data Update',
    '2': 'Critical Tooling Identification',
    '3': 'Supplier Validation Request',
    '4': 'Critical Tooling Reassessment',
    '5': 'On-Site Technical Analysis',
    '6': 'Technical Confirmation',
    '7': 'Supply Continuity Strategy'
  };
  return descriptions[stepValue] || '';
}

function getStepResponsible(stepValue) {
  const responsibles = {
    '1': 'Supply Continuity',
    '2': 'Supply Continuity',
    '3': 'Supply Continuity',
    '4': 'Supply Continuity',
    '5': 'SQIE',
    '6': 'SQIE',
    '7': 'Sourcing Manager'
  };
  return responsibles[stepValue] || '';
}

function handleStepsSelectChange(cardIndex, itemId, selectEl) {
  if (selectEl) {
    const value = (selectEl.value || '').trim();
    updateCardHeaderSteps(cardIndex, value);
    
    // Sincroniza steps com a linha da spreadsheet
    syncSpreadsheetFromExpandedCard(itemId, 'steps', value);
    
    // Update step description label with description and responsible
    const descriptionLabel = document.getElementById(`stepDescription_${itemId}`);
    if (descriptionLabel) {
      const description = getStepDescription(value);
      const responsible = getStepResponsible(value);
      
      if (description) {
        descriptionLabel.innerHTML = `<div style="color: #8b92a7; line-height: 1.4;">${description}<br><span style="font-size: 0.9em;">Responsible: ${responsible}</span></div>`;
        descriptionLabel.style.display = 'block';
        descriptionLabel.style.textAlign = 'left';
        descriptionLabel.style.marginTop = '4px';
      } else {
        descriptionLabel.textContent = '';
        descriptionLabel.style.display = 'none';
      }
    }
  }
  autoSaveTooling(itemId, true);
}

async function handleAnalysisCompletedChange(itemId, isChecked) {
  try {
    // Atualiza o campo no banco de dados
    const value = isChecked ? 1 : 0;
    await window.api.updateTooling(itemId, { analysis_completed: value });
    
    // Atualiza o item local
    const item = toolingData.find(t => t.id === itemId);
    if (item) {
      item.analysis_completed = value;
      
      // Adiciona comentário automático se foi marcado
      if (isChecked) {
        const now = new Date();
        const dateStr = now.toLocaleDateString('pt-BR', { 
          day: '2-digit', 
          month: '2-digit', 
          year: 'numeric'
        });
        const commentText = 'Analysis completed - tooling reviewed and no revitalization required.';
        const newComment = {
          date: dateStr,
          text: commentText
        };
        
        // Adiciona ao array de comentários
        let comments = [];
        if (item.comments) {
          try {
            comments = typeof item.comments === 'string' ? JSON.parse(item.comments) : item.comments;
          } catch {
            comments = [];
          }
        }
        comments.push(newComment);
        item.comments = JSON.stringify(comments);
        
        // Salva comentário no banco
        await window.api.updateTooling(itemId, { comments: item.comments });
        
        // Atualiza UI dos comentários
        const commentsList = document.getElementById(`commentsList_${itemId}`);
        if (commentsList) {
          commentsList.innerHTML = buildCommentsListHTML(item.comments, itemId);
        }
      }
      
      // Atualiza o ícone de expiração em todos os lugares
      updateExpirationIconsForItem(itemId);
      
      // Atualiza métricas do card do supplier em tempo real (busca dados frescos do banco)
      if (selectedSupplier) {
        refreshSupplierCardMetricsFromDB(selectedSupplier);
      }
    }
    
    showNotification(isChecked ? 'Analysis marked as completed' : 'Analysis marked as pending', 'success');
  } catch (error) {
    console.error('Error updating analysis completed:', error);
    showNotification('Error updating analysis status', 'error');
  }
}

function updateExpirationIconsForItem(itemId) {
  const item = toolingData.find(t => String(t.id) === String(itemId));
  if (!item) return;
  
  const classification = classifyToolingExpirationState(item);
  const isAnalysisCompleted = item.analysis_completed === 1;
  const hasExpirationDate = classification.expirationDate && classification.expirationDate !== '';
  
  // Atualiza ícone na linha do spreadsheet
  const spreadsheetRow = document.querySelector(`#spreadsheetBody tr[data-id="${itemId}"]`);
  if (spreadsheetRow) {
    const expirationCell = spreadsheetRow.querySelector('.spreadsheet-expiration');
    if (expirationCell) {
      const expirationDateDisplay = hasExpirationDate ? formatDate(classification.expirationDate) : '';
      const expirationIconHtml = hasExpirationDate ? getSpreadsheetExpirationIcon(classification.state, isAnalysisCompleted) : '';
      expirationCell.innerHTML = `
        ${expirationIconHtml}
        <span class="expiration-text">${expirationDateDisplay}</span>
      `;
    }
  }
  
  // Atualiza ícone no card expandido (se estiver aberto)
  const cardContainers = document.querySelectorAll(`[data-item-id="${itemId}"]`);
  cardContainers.forEach(container => {
    const expirationInput = container.querySelector('input[data-field="expiration_date"]');
    if (expirationInput) {
      const inputContainer = expirationInput.closest('.input-with-icon');
      if (inputContainer) {
        // Remove ícone antigo
        const oldIcon = inputContainer.querySelector('.input-icon');
        if (oldIcon) {
          oldIcon.remove();
        }
        // Adiciona novo ícone usando a mesma função do card
        const newIconHtml = getCardExpirationIcon(classification.state, isAnalysisCompleted);
        inputContainer.insertAdjacentHTML('afterbegin', newIconHtml);
      }
    }
  });
}

function handlePNChange(cardIndex, itemId, inputEl) {
  if (inputEl) {
    const value = (inputEl.value || '').trim();
    const card = getCardContainer(cardIndex);
    if (card) {
      const pnDisplay = card.querySelector('.tooling-info-value.highlight');
      if (pnDisplay) {
        pnDisplay.textContent = value || 'N/A';
      }
    }
    // Sincroniza pn com a linha da spreadsheet
    syncSpreadsheetFromExpandedCard(itemId, 'pn', value);
  }
  autoSaveTooling(itemId);
}

// Handler genérico para campos de texto do card que sincroniza com a spreadsheet
function handleCardTextFieldChange(itemId, inputEl) {
  if (inputEl) {
    const field = inputEl.dataset.field;
    const value = (inputEl.value || '').trim();
    // Sincroniza com a linha da spreadsheet
    syncSpreadsheetFromExpandedCard(itemId, field, value);
  }
  autoSaveTooling(itemId);
}

function updateCardStatusAttribute(cardIndex, statusValue) {
  const card = getCardContainer(cardIndex);
  if (!card) {
    return;
  }
  const normalized = (statusValue || '').toString().trim().toLowerCase();
  card.dataset.status = normalized;
  card.classList.toggle('is-obsolete', normalized === 'obsolete');
  syncObsoleteLinkVisibility(card, normalized === 'obsolete');
  updateCardStatusIcon(card, normalized);
  enforceChainIndicatorRules(card);

  if (normalized !== 'obsolete') {
    const currentReplacementId = sanitizeReplacementId(card.dataset.replacementId || '');
    if (currentReplacementId) {
      syncReplacementLinkControls(cardIndex, '');
    }
  }
}

function updateCardStatusIcon(card, normalizedStatus) {
  const statusIconContainer = card.querySelector('.card-status-icon');
  const isObsolete = normalizedStatus === 'obsolete';
  
  if (isObsolete) {
    // Show obsolete icon
    if (statusIconContainer) {
      statusIconContainer.className = 'card-status-icon status-icon-obsolete';
      statusIconContainer.innerHTML = '<i class="ph ph-fill ph-archive-box"></i>';
      statusIconContainer.title = 'Obsolete';
    } else {
      // Create icon if it doesn't exist
      const headerTop = card.querySelector('.tooling-card-header-top');
      if (headerTop) {
        const iconDiv = document.createElement('div');
        iconDiv.className = 'card-status-icon status-icon-obsolete';
        iconDiv.innerHTML = '<i class="ph ph-fill ph-archive-box"></i>';
        iconDiv.title = 'Obsolete';
        headerTop.insertBefore(iconDiv, headerTop.firstChild);
      }
    }
  } else {
    // Remove obsolete icon if not obsolete anymore
    if (statusIconContainer && statusIconContainer.classList.contains('status-icon-obsolete')) {
      statusIconContainer.remove();
    }
  }
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
  const card = getCardContainer(cardIndex);
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
  
  // Atualiza o ícone de corrente (chain indicator)
  updateChainIndicatorForCard(card, sanitizedValue);
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

function enforceChainIndicatorRules(card) {
  if (!card) {
    return;
  }
  const chainIndicator = card.querySelector('.tooling-chain-indicator');
  if (!chainIndicator) {
    return;
  }
  const normalizedStatus = (card.dataset.status || '').trim().toLowerCase();
  const replacementIdValue = sanitizeReplacementId(card.dataset.replacementId || '');
  const hasOutgoingChain = normalizedStatus === 'obsolete' && replacementIdValue !== '';
  const hasIncomingChain = card.dataset.hasIncomingChain === 'true' || card.dataset.chainMember === 'true';
  const shouldShow = hasOutgoingChain || hasIncomingChain;
  if (shouldShow) {
    chainIndicator.removeAttribute('hidden');
  } else {
    chainIndicator.setAttribute('hidden', 'true');
  }
}

function updateChainIndicatorForCard(card, replacementId) {
  if (!card) {
    return;
  }
  const sanitizedValue = sanitizeReplacementId(replacementId);
  card.dataset.replacementId = sanitizedValue;
  enforceChainIndicatorRules(card);
}

async function toggleReplacementPicker(event, cardIndex) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }
  await openReplacementPickerOverlay(cardIndex);
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

function ensureReplacementPickerOverlayElements() {
  if (replacementPickerOverlayState.overlay) {
    return true;
  }
  const overlay = document.getElementById('replacementPickerOverlay');
  if (!overlay) {
    return false;
  }

  const list = overlay.querySelector('[data-replacement-overlay-list]');
  const searchInput = overlay.querySelector('[data-replacement-overlay-search]');
  const title = overlay.querySelector('[data-replacement-overlay-title]');
  const subtitle = overlay.querySelector('[data-replacement-overlay-subtitle]');

  replacementPickerOverlayState = {
    overlay,
    list,
    searchInput,
    title,
    subtitle,
    cardIndex: null,
    itemId: null
  };

  if (searchInput) {
    searchInput.addEventListener('input', (event) => {
      handleReplacementPickerOverlaySearch(event.target.value);
    });
  }

  overlay.addEventListener('click', (event) => {
    if (event.target === overlay) {
      closeReplacementPickerOverlay();
    }
  });

  return true;
}

async function openReplacementPickerOverlay(cardIndex) {
  if (!ensureReplacementPickerOverlayElements()) {
    return;
  }

  const card = getCardContainer(cardIndex);
  if (!card) {
    return;
  }

  await ensureReplacementIdOptions();

  const itemId = Number(card.dataset.itemId) || 0;
  replacementPickerOverlayState.cardIndex = cardIndex;
  replacementPickerOverlayState.itemId = itemId;

  if (replacementPickerOverlayState.title) {
    replacementPickerOverlayState.title.textContent = 'Select replacement tooling';
  }
  if (replacementPickerOverlayState.subtitle) {
    const pn = card.querySelector('.tooling-info-value.highlight')?.textContent?.trim();
    replacementPickerOverlayState.subtitle.textContent = pn ? `Card #${itemId} • ${pn}` : `Card #${itemId}`;
  }

  if (replacementPickerOverlayState.list) {
    const stubItem = { id: itemId };
    replacementPickerOverlayState.list.innerHTML = buildReplacementPickerOptionsMarkup(stubItem, cardIndex);
  }

  const overlay = replacementPickerOverlayState.overlay;
  overlay.classList.add('active');

  const searchInput = replacementPickerOverlayState.searchInput;
  if (searchInput) {
    searchInput.value = '';
    setTimeout(() => searchInput.focus(), 40);
  }
}

function closeReplacementPickerOverlay() {
  if (!replacementPickerOverlayState.overlay) {
    return;
  }
  replacementPickerOverlayState.overlay.classList.remove('active');
  replacementPickerOverlayState.cardIndex = null;
  replacementPickerOverlayState.itemId = null;
}

function handleReplacementPickerOverlaySearch(searchValue) {
  if (!replacementPickerOverlayState.list) {
    return;
  }
  const normalizedQuery = (searchValue || '').trim().toLowerCase();
  replacementPickerOverlayState.list.querySelectorAll('.replacement-dropdown-option').forEach((option) => {
    const searchText = option.textContent.toLowerCase();
    const matches = !normalizedQuery || searchText.includes(normalizedQuery);
    if (matches) {
      option.removeAttribute('hidden');
    } else {
      option.setAttribute('hidden', 'true');
    }
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

  // Resetar viewport transform ao abrir
  applyViewportTransform();
  
  // Desenhar grid imediatamente ao abrir
  setTimeout(() => {
    drawReplacementGrid();
  }, 50);

  try {
    const chain = await buildReplacementTimeline(normalizedId);
    renderReplacementTimeline(chain);
  } catch (error) {
    showNotification('Não foi possível carregar a linha temporal de substituições.', 'error');
    renderReplacementTimeline([]);
  }
}

function closeReplacementTimelineOverlay() {
  const { overlay, gridCanvas, connectionsCanvas, list } = replacementTimelineElements;
  currentTimelineRootId = null;
  
  // Reset viewport state
  graphViewportState = {
    scale: 1,
    offsetX: 0,
    offsetY: 0,
    isPanning: false,
    panStartX: 0,
    panStartY: 0,
    draggedNode: null,
    dragStartX: 0,
    dragStartY: 0
  };
  
  if (list) {
    list.style.transform = '';
  }
  
  if (overlay) {
    overlay.classList.remove('active');
  }
  
  // Limpar canvas ao fechar
  if (gridCanvas) {
    const ctx = gridCanvas.getContext('2d');
    ctx.clearRect(0, 0, gridCanvas.width, gridCanvas.height);
  }
  if (connectionsCanvas) {
    const ctx = connectionsCanvas.getContext('2d');
    ctx.clearRect(0, 0, connectionsCanvas.width, connectionsCanvas.height);
  }
}

let graphViewportState = {
  scale: 1,
  offsetX: 0,
  offsetY: 0,
  isPanning: false,
  panStartX: 0,
  panStartY: 0,
  draggedNode: null,
  dragStartX: 0,
  dragStartY: 0
};

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
  list.style.display = 'block';

  // Posicionar cards em layout vertical inicial
  list.innerHTML = chain.map((record, index) => {
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

    // Posição inicial: centralizado verticalmente espaçado
    const initialX = 250;
    const initialY = 50 + (index * 200);

    return `
      <div class="${itemClasses.join(' ')}" 
           data-record-id="${toolingId}" 
           data-node-index="${index}"
           style="left: ${initialX}px; top: ${initialY}px;">
        <div class="timeline-item-content">
          <div class="timeline-item-header">
            <span class="timeline-item-id">#${toolingId}</span>
            <div class="timeline-header-right">
              <span class="timeline-status-badge ${badgeClass}">${label}</span>
              <button class="timeline-item-action" type="button" onclick="handleTimelineCardNavigate(event, ${toolingId})" title="Open card">
                <i class="ph ph-arrow-square-out"></i>
              </button>
            </div>
          </div>
          <div class="timeline-meta-row">
            <span><span class="timeline-meta-label">PN:</span> ${pn}</span>
            <span><span class="timeline-meta-label">Supplier:</span> ${supplier}</span>
          </div>
          <p class="timeline-item-description">${description}</p>
        </div>
      </div>
    `;
  }).join('');

  // Attach node drag listeners
  list.querySelectorAll('.timeline-item').forEach(node => {
    node.addEventListener('mousedown', handleNodeDragStart);
  });

  // Attach viewport pan/zoom listeners
  initGraphViewportControls();

  // Draw grid and connections after render with delay for DOM updates
  setTimeout(() => {
    console.log('Iniciando desenho de grid e conexões');
    drawReplacementGrid();
    drawReplacementConnections(chain);
  }, 200);
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
    return null;
  }
}

function handleTimelineCardNavigate(event, itemId) {
  if (event) {
    event.stopPropagation();
  }
  if (!itemId) {
    return;
  }
  closeReplacementTimelineOverlay();
  navigateToLinkedCard(itemId);
}

function updateCardUIAfterReorder(itemId, newStatus, newReplacementId) {
  const card = document.querySelector(`.tooling-card[data-item-id="${itemId}"]`);
  if (!card) {
    return;
  }
  
  const cardIndex = parseInt(card.id.replace('card-', ''), 10);
  if (isNaN(cardIndex)) {
    return;
  }
  
  // Update status field and visual attributes
  const statusField = card.querySelector('[data-card-status]');
  if (statusField) {
    statusField.textContent = newStatus || 'N/A';
  }
  
  // Update status dropdown/select
  const statusSelect = card.querySelector('select[data-field="status"]');
  if (statusSelect) {
    statusSelect.value = newStatus || '';
  }
  
  // Update card status attributes and visibility
  updateCardStatusAttribute(cardIndex, newStatus);
  
  // Update replacement link hidden input
  const replacementInput = card.querySelector('[data-field="replacement_tooling_id"]');
  if (replacementInput) {
    replacementInput.value = newReplacementId || '';
  }
  
  // Update link controls
  syncReplacementLinkControls(cardIndex, newReplacementId || '');
  
  // Update global toolingData
  const dataIndex = toolingData.findIndex(item => String(item.id) === String(itemId));
  if (dataIndex !== -1) {
    toolingData[dataIndex].status = newStatus;
    toolingData[dataIndex].replacement_tooling_id = newReplacementId;
  }
}

function handleNodeDragStart(event) {
  if (event.target.closest('.timeline-item-action') || event.target.closest('button')) {
    return; // Não iniciar drag se clicar em botões
  }
  
  event.preventDefault();
  const node = event.currentTarget;
  
  graphViewportState.draggedNode = node;
  graphViewportState.dragStartX = event.clientX - parseFloat(node.style.left || 0);
  graphViewportState.dragStartY = event.clientY - parseFloat(node.style.top || 0);
  
  node.classList.add('dragging-node');
  
  document.addEventListener('mousemove', handleNodeDragMove);
  document.addEventListener('mouseup', handleNodeDragEnd);
}

function handleNodeDragMove(event) {
  if (!graphViewportState.draggedNode) return;
  
  event.preventDefault();
  const node = graphViewportState.draggedNode;
  
  const x = event.clientX - graphViewportState.dragStartX;
  const y = event.clientY - graphViewportState.dragStartY;
  
  node.style.left = `${x}px`;
  node.style.top = `${y}px`;
  
  // Redesenhar conexões
  const { list } = replacementTimelineElements;
  if (list) {
    const chain = Array.from(list.querySelectorAll('.timeline-item')).map(n => ({ id: n.dataset.recordId }));
    drawReplacementConnections(chain);
  }
}

function handleNodeDragEnd(event) {
  if (graphViewportState.draggedNode) {
    graphViewportState.draggedNode.classList.remove('dragging-node');
    graphViewportState.draggedNode = null;
  }
  
  document.removeEventListener('mousemove', handleNodeDragMove);
  document.removeEventListener('mouseup', handleNodeDragEnd);
}

function initGraphViewportControls() {
  const viewport = document.getElementById('replacementGraphViewport');
  if (!viewport) return;
  
  // Limpar listeners anteriores
  viewport.removeEventListener('mousedown', handleViewportPanStart);
  viewport.removeEventListener('wheel', handleViewportZoom);
  
  viewport.addEventListener('mousedown', handleViewportPanStart);
  viewport.addEventListener('wheel', handleViewportZoom);
}

function handleViewportPanStart(event) {
  // Apenas pan se clicar no fundo (não em nodes)
  if (event.target.closest('.timeline-item')) return;
  
  event.preventDefault();
  const viewport = event.currentTarget;
  
  graphViewportState.isPanning = true;
  graphViewportState.panStartX = event.clientX - graphViewportState.offsetX;
  graphViewportState.panStartY = event.clientY - graphViewportState.offsetY;
  
  viewport.classList.add('panning');
  
  document.addEventListener('mousemove', handleViewportPanMove);
  document.addEventListener('mouseup', handleViewportPanEnd);
}

function handleViewportPanMove(event) {
  if (!graphViewportState.isPanning) return;
  
  event.preventDefault();
  graphViewportState.offsetX = event.clientX - graphViewportState.panStartX;
  graphViewportState.offsetY = event.clientY - graphViewportState.panStartY;
  
  applyViewportTransform();
  
  // Redesenhar conexões durante pan
  const { list } = replacementTimelineElements;
  if (list) {
    const chain = Array.from(list.querySelectorAll('.timeline-item')).map(n => ({ id: n.dataset.recordId }));
    drawReplacementConnections(chain);
  }
}

function handleViewportPanEnd(event) {
  graphViewportState.isPanning = false;
  const viewport = document.getElementById('replacementGraphViewport');
  if (viewport) {
    viewport.classList.remove('panning');
  }
  
  document.removeEventListener('mousemove', handleViewportPanMove);
  document.removeEventListener('mouseup', handleViewportPanEnd);
}

function handleViewportZoom(event) {
  event.preventDefault();
  
  const delta = -event.deltaY;
  const scaleChange = delta > 0 ? 1.1 : 0.9;
  const newScale = Math.max(0.3, Math.min(3, graphViewportState.scale * scaleChange));
  
  graphViewportState.scale = newScale;
  applyViewportTransform();
  
  // Redesenhar conexões com novo zoom
  const { list } = replacementTimelineElements;
  if (list) {
    const chain = Array.from(list.querySelectorAll('.timeline-item')).map(n => ({ id: n.dataset.recordId }));
    setTimeout(() => {
      drawReplacementGrid();
      drawReplacementConnections(chain);
    }, 10);
  }
}

function applyViewportTransform() {
  const { list, connectionsCanvas } = replacementTimelineElements;
  const transform = `translate(${graphViewportState.offsetX}px, ${graphViewportState.offsetY}px) scale(${graphViewportState.scale})`;
  
  if (list) {
    list.style.transform = transform;
  }
  if (connectionsCanvas) {
    connectionsCanvas.style.transform = transform;
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

function drawReplacementGrid() {
  const { gridCanvas } = replacementTimelineElements;
  if (!gridCanvas) return;

  const viewport = document.getElementById('replacementGraphViewport');
  if (!viewport) return;

  const rect = viewport.getBoundingClientRect();
  const width = rect.width;
  const height = rect.height;
  
  gridCanvas.width = width;
  gridCanvas.height = height;

  const ctx = gridCanvas.getContext('2d');
  const gridSize = 30;
  const gridColor = 'rgba(255, 255, 255, 0.05)';

  ctx.clearRect(0, 0, width, height);
  ctx.strokeStyle = gridColor;
  ctx.lineWidth = 1;

  // Linhas verticais
  for (let x = 0; x < width; x += gridSize) {
    ctx.beginPath();
    ctx.moveTo(x, 0);
    ctx.lineTo(x, height);
    ctx.stroke();
  }

  // Linhas horizontais
  for (let y = 0; y < height; y += gridSize) {
    ctx.beginPath();
    ctx.moveTo(0, y);
    ctx.lineTo(width, y);
    ctx.stroke();
  }
}

function drawReplacementConnections(chain = []) {
  const { connectionsCanvas, list } = replacementTimelineElements;
  if (!connectionsCanvas || !list) {
    console.log('Canvas ou lista não encontrados');
    return;
  }

  const nodes = Array.from(list.querySelectorAll('.timeline-item'));
  console.log(`Desenhando conexões para ${nodes.length} nodes`);
  
  if (nodes.length < 2) {
    console.log('Menos de 2 nodes, não há o que conectar');
    return;
  }

  // Usar tamanho grande o suficiente para acomodar todos os nodes
  const width = 4000;
  const height = 4000;
  
  connectionsCanvas.width = width;
  connectionsCanvas.height = height;

  const ctx = connectionsCanvas.getContext('2d');
  ctx.clearRect(0, 0, width, height);
  
  for (let i = 0; i < nodes.length - 1; i++) {
    const from = nodes[i];
    const to = nodes[i + 1];
    
    const fromLeft = parseFloat(from.style.left || 0);
    const fromTop = parseFloat(from.style.top || 0);
    const toLeft = parseFloat(to.style.left || 0);
    const toTop = parseFloat(to.style.top || 0);
    
    const fromWidth = from.offsetWidth;
    const fromHeight = from.offsetHeight;
    const toWidth = to.offsetWidth;
    
    const fromX = fromLeft + fromWidth / 2;
    const fromY = fromTop + fromHeight;
    const toX = toLeft + toWidth / 2;
    const toY = toTop;

    console.log(`Conectando node ${i} (${fromX}, ${fromY}) -> node ${i+1} (${toX}, ${toY})`);

    // Desenhar linha com curva suave
    ctx.strokeStyle = '#ff6b6b';
    ctx.lineWidth = 4;

    ctx.beginPath();
    ctx.moveTo(fromX, fromY);
    
    const controlOffset = Math.abs(toY - fromY) / 2;
    ctx.bezierCurveTo(
      fromX, fromY + controlOffset,
      toX, toY - controlOffset,
      toX, toY
    );
    
    ctx.stroke();
  }
  
  console.log('Conexões desenhadas com sucesso');
}

async function updateReplacementChainAfterReorder() {
  const list = replacementTimelineElements.list;
  if (!list) {
    return;
  }
  
  const items = Array.from(list.querySelectorAll('.timeline-item'));
  const orderedIds = items.map(item => item.dataset.recordId).filter(Boolean);
  if (orderedIds.length === 0) {
    return;
  }
  
  isReorderingTimeline = true;
  
  try {
    // Update replacement links and status based on new order
    for (let i = 0; i < orderedIds.length; i++) {
      const currentId = orderedIds[i];
      const isLast = i === orderedIds.length - 1;
      
      if (isLast) {
        // Last item: ACTIVE status and no replacement link
        await window.api.updateTooling(Number(currentId), {
          replacement_tooling_id: null,
          status: 'ACTIVE'
        });
      } else {
        // All others: OBSOLETE + link to next
        await window.api.updateTooling(Number(currentId), {
          replacement_tooling_id: Number(orderedIds[i + 1]),
          status: 'OBSOLETE'
        });
      }
    }
    
    showNotification('Cadeia reorganizada e salva com sucesso.', 'success');
    
    // Update cards UI immediately
    for (let i = 0; i < orderedIds.length; i++) {
      const currentId = orderedIds[i];
      const isLast = i === orderedIds.length - 1;
      const nextId = isLast ? null : Number(orderedIds[i + 1]);
      const newStatus = isLast ? 'ACTIVE' : 'OBSOLETE';
      
      updateCardUIAfterReorder(currentId, newStatus, nextId);
    }
    
    // Rebuild timeline with updated data
    const rootId = orderedIds[0];
    const chain = await buildReplacementTimeline(rootId);
    renderReplacementTimeline(chain);
    
  } catch (error) {
    showNotification('Erro ao atualizar a ordem da cadeia.', 'error');
  } finally {
    isReorderingTimeline = false;
  }
}

function handleReplacementPickerSearch(cardIndex, searchValue) {
  const card = getCardContainer(cardIndex);
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
  const card = getCardContainer(cardIndex);
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
  closeReplacementPickerOverlay();
}

function handleReplacementLinkButtonClick(event, buttonEl) {
  if (event) {
    event.stopPropagation();
  }
  if (!buttonEl || buttonEl.disabled) {
    return;
  }
  
  // Obtém o ID do target diretamente do botão
  const targetId = buttonEl.getAttribute('data-target-id');
  if (targetId) {
    navigateToLinkedCard(targetId);
  }
}

function handleReplacementLinkChipClick(event, buttonEl) {
  if (event) {
    event.stopPropagation();
  }
  if (!buttonEl || buttonEl.disabled) {
    return;
  }
  
  // Tenta encontrar o card (normal ou spreadsheet)
  const card = buttonEl.closest('.tooling-card') || buttonEl.closest('.spreadsheet-card-container');
  if (card) {
    openReplacementTimelineForCard(card);
  }
}

async function navigateToLinkedCard(targetId) {
  const normalizedId = sanitizeReplacementId(targetId);
  if (!normalizedId) {
    showNotification('Informe um ID de substituição válido.', 'error');
    return;
  }

  // Verifica se está no modo spreadsheet
  const isSpreadsheetMode = currentViewMode === 'spreadsheet';
  
  if (isSpreadsheetMode) {
    // Modo Spreadsheet: procura pela linha da tabela
    let targetRow = document.querySelector(`tr[data-id="${normalizedId}"]`);
    
    if (!targetRow) {
      const cardLoaded = await ensureCardLoadedById(normalizedId);
      if (!cardLoaded) {
        showNotification(`O card #${normalizedId} não foi encontrado. Verifique se o registro existe.`, 'warning');
        return;
      }
      
      await new Promise(resolve => setTimeout(resolve, 400));
      targetRow = document.querySelector(`tr[data-id="${normalizedId}"]`);
    }

    if (!targetRow) {
      showNotification(`Não foi possível exibir o card #${normalizedId}. O registro pode não existir mais.`, 'warning');
      return;
    }

    // Encontra o índice do item
    const item = toolingData.find(t => String(t.id) === String(normalizedId));
    const itemIndex = toolingData.indexOf(item);
    
    if (itemIndex === -1) {
      showNotification(`Não foi possível encontrar o item #${normalizedId}.`, 'warning');
      return;
    }

    // Expande a linha se não estiver expandida
    const isExpanded = targetRow.classList.contains('row-expanded');
    if (!isExpanded) {
      toggleSpreadsheetRow(normalizedId, itemIndex);
    }

    // Aguarda a linha ser expandida e carrega dados
    setTimeout(() => {
      const detailRow = targetRow.nextElementSibling;
      if (detailRow && detailRow.classList.contains('spreadsheet-detail-row')) {
        // Carrega anexos e calcula expiração
        loadCardAttachments(normalizedId).catch(err => {
        });
        
        // Calcula expiração (se necessário)
        const cardContainer = detailRow.querySelector('.spreadsheet-card-container');
        if (cardContainer) {
          const expirationInput = cardContainer.querySelector('[data-field="expiration_date"]');
          if (expirationInput) {
            calculateExpirationDate(itemIndex, null, true);
          }
        }
        
        ensureSpreadsheetRowVisible(detailRow);
        flashSpreadsheetRowHighlight(detailRow);
      } else {
        ensureSpreadsheetRowVisible(targetRow);
        flashSpreadsheetRowHighlight(targetRow);
      }
    }, 150);
    
  } else {
    // Modo Card: comportamento original
    let targetCard = document.querySelector(`.tooling-card[data-item-id="${normalizedId}"]`);
    if (!targetCard) {
      const cardLoaded = await ensureCardLoadedById(normalizedId);
      if (!cardLoaded) {
        showNotification(`O card #${normalizedId} não foi encontrado. Verifique se o registro existe.`, 'warning');
        return;
      }
      
      await new Promise(resolve => setTimeout(resolve, 400));
      targetCard = document.querySelector(`.tooling-card[data-item-id="${normalizedId}"]`);
      if (!targetCard) {
        await new Promise(resolve => setTimeout(resolve, 400));
        targetCard = document.querySelector(`.tooling-card[data-item-id="${normalizedId}"]`);
      }
    }

    if (!targetCard) {
      showNotification(`Não foi possível exibir o card #${normalizedId}. O registro pode não existir mais.`, 'warning');
      return;
    }

    const cardIndex = parseInt(targetCard.id.replace('card-', ''), 10);
    const expandedCards = document.querySelectorAll('.tooling-card.expanded');
    for (const expandedCard of expandedCards) {
      if (expandedCard !== targetCard) {
        const itemId = expandedCard.getAttribute('data-item-id');
        if (itemId) {
          await saveToolingQuietly(itemId);
        }
        expandedCard.classList.remove('expanded');
      }
    }
    
    const bodyLoaded = targetCard.getAttribute('data-body-loaded') === 'true';
    if (!bodyLoaded) {
      const itemId = targetCard.getAttribute('data-item-id');
      const item = toolingData.find(t => String(t.id) === String(itemId));
      
      if (item) {
        const supplierContext = selectedSupplier || currentSupplier || '';
        const chainMembership = new Map();
        const bodyHTML = buildToolingCardBodyHTML(item, cardIndex, chainMembership, supplierContext);
        
        targetCard.insertAdjacentHTML('beforeend', bodyHTML);
        targetCard.setAttribute('data-body-loaded', 'true');
        applyInitialThousandsMask(targetCard);
        
        const dropzone = targetCard.querySelector('.card-attachments-dropzone');
        if (dropzone && itemId) {
          initCardAttachmentDragAndDrop(dropzone, itemId);
        }
      }
    }
    
    targetCard.classList.add('expanded');
    
    setTimeout(() => {
      calculateExpirationDate(cardIndex, null, true);
      const itemId = targetCard.getAttribute('data-item-id');
      if (itemId) {
        loadCardAttachments(itemId).catch(err => {
        });
      }
    }, 0);
    
    ensureCardVisible(targetCard);
    flashCardHighlight(targetCard);
  }
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

function flashSpreadsheetRowHighlight(row) {
  if (!row) {
    return;
  }
  row.classList.add('row-highlight');
  setTimeout(() => {
    row.classList.remove('row-highlight');
  }, 1200);
}

function ensureSpreadsheetRowVisible(row) {
  if (!row) return;
  
  setTimeout(() => {
    // Encontra a linha principal (não a detail row)
    let mainRow = row;
    if (row.classList.contains('spreadsheet-detail-row')) {
      mainRow = row.previousElementSibling;
    }
    
    if (mainRow) {
      // Calcula a posição ideal considerando o header fixo
      const headerHeight = document.querySelector('.spreadsheet-table thead')?.offsetHeight || 50;
      const container = document.querySelector('.spreadsheet-container');
      if (container) {
        const rowTop = mainRow.offsetTop - headerHeight - 10;
        container.scrollTo({ top: rowTop, behavior: 'smooth' });
      }
    }
  }, 200);
}

function triggerDateReminder(card, fieldName) {
  const dateInput = card.querySelector(`[data-field="${fieldName}"]`);
  if (!dateInput) {
    return;
  }

  showDateHighlight(dateInput);
}

function restoreDateReminders(itemId) {
  const dateFields = ['date_remaining_tooling_life', 'date_annual_volume'];
  
  dateFields.forEach(fieldName => {
    const storageKey = `dateReminder_${itemId}_${fieldName}`;
    const hasReminder = localStorage.getItem(storageKey);
    
    if (hasReminder === 'active') {
      const input = document.querySelector(`[data-id="${itemId}"][data-field="${fieldName}"]`);
      if (input) {
        showDateHighlight(input);
      }
    }
  });
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

  // Salva o estado no localStorage para persistir após fechar/abrir
  const itemId = input.getAttribute('data-id');
  const fieldName = input.getAttribute('data-field');
  if (itemId && fieldName) {
    const storageKey = `dateReminder_${itemId}_${fieldName}`;
    localStorage.setItem(storageKey, 'active');
  }

  // Remove listener anterior se existir
  if (input._dateChangeListener) {
    input.removeEventListener('change', input._dateChangeListener);
  }

  // Adiciona listener para remover o highlight quando o usuário alterar a data
  const changeListener = function() {
    input.classList.remove('date-highlight');
    tooltip.classList.remove('active');
    
    // Remove do localStorage
    if (itemId && fieldName) {
      const storageKey = `dateReminder_${itemId}_${fieldName}`;
      localStorage.removeItem(storageKey);
    }
    
    // Remove o listener após usar
    input.removeEventListener('change', changeListener);
    delete input._dateChangeListener;
  };

  input._dateChangeListener = changeListener;
  input.addEventListener('change', changeListener);
}

// Debounce para salvar automaticamente
let autoSaveTimeouts = {};
const cardSnapshotStore = new Map();
let interfaceRefreshTimerId = null;
let pendingInterfaceRefreshReason = null;
const INTERFACE_REFRESH_DELAY_MS = 800;
const NUMERIC_CARD_FIELDS = [
  'tooling_life_qty',
  'produced',
  'annual_volume_forecast',
  'remaining_tooling_life_pcs',
  'percent_tooling_life',
  'amount_brl',
  'tool_quantity'
];

function getSnapshotKey(id) {
  return String(id ?? '');
}

function collectCardDomValues(id) {
  const snapshotKey = getSnapshotKey(id);
  if (!snapshotKey) {
    return null;
  }
  const elements = document.querySelectorAll(`[data-id="${snapshotKey}"]`);
  if (!elements || elements.length === 0) {
    return null;
  }
  const values = {};
  const fieldPriority = {}; // Rastreia prioridade: 2=card expandido, 1=spreadsheet row, 0=outro
  
  elements.forEach((element) => {
    const field = element.getAttribute('data-field');
    if (!field || typeof field !== 'string') {
      return;
    }
    
    // Validar nome do campo
    const fieldName = field.trim();
    if (fieldName === '' || fieldName.length === 0) {
      return;
    }
    
    // Ignorar campos calculados (como expiration_date)
    if (element.classList.contains('calculated')) {
      return;
    }
    
    // Determinar prioridade do elemento
    // Card expandido (dentro de .spreadsheet-detail-row ou .tooling-card-body) tem prioridade máxima
    const isInExpandedCard = element.closest('.spreadsheet-detail-row') || element.closest('.tooling-card-body');
    const isInSpreadsheetRow = element.closest('tr[data-id]') && element.classList.contains('spreadsheet-input');
    
    let priority = 0;
    if (isInExpandedCard) {
      priority = 2; // Prioridade máxima - card expandido
    } else if (isInSpreadsheetRow) {
      priority = 1; // Prioridade média - linha da spreadsheet
    }
    
    // Só sobrescreve se tiver prioridade maior ou igual
    const currentPriority = fieldPriority[fieldName] ?? -1;
    if (priority >= currentPriority) {
      fieldPriority[fieldName] = priority;
      
      if (element instanceof HTMLInputElement && element.type === 'checkbox') {
        values[fieldName] = element.checked ? '1' : '0';
      } else {
        values[fieldName] = element.value ?? '';
      }
    }
  });
  
  return values;
}

function serializeCardValues(values) {
  if (!values) {
    return '';
  }
  const sortedKeys = Object.keys(values).sort();
  const normalized = sortedKeys.map(key => [key, values[key]]);
  return JSON.stringify(normalized);
}

function normalizeCardPayload(values) {
  if (!values) {
    return null;
  }
  const payload = { ...values };
  NUMERIC_CARD_FIELDS.forEach((field) => {
    if (Object.prototype.hasOwnProperty.call(payload, field) && payload[field] !== '') {
      const parsed = parseLocalizedNumber(payload[field]);
      payload[field] = Number.isNaN(parsed) ? 0 : parsed;
    }
  });
  return payload;
}

function buildCardPayloadFromDom(id) {
  const rawValues = collectCardDomValues(id);
  if (!rawValues) {
    return null;
  }
  
  // Adicionar comentários e expiration_date do toolingData ao payload
  // (expiration_date é calculado e tem classe 'calculated', então é ignorado pelo collectCardDomValues)
  const item = toolingData.find(item => Number(item.id) === Number(id));
  if (item) {
    if (item.comments) {
      rawValues.comments = item.comments;
    }
    if (item.expiration_date !== undefined) {
      rawValues.expiration_date = item.expiration_date;
    }
  }
  
  const serialized = serializeCardValues(rawValues);
  const snapshotKey = getSnapshotKey(id);
  const previousSnapshot = cardSnapshotStore.get(snapshotKey);
  
  // Se não tem snapshot anterior, considera que há mudanças (primeiro salvamento)
  const hasChanges = previousSnapshot === undefined || previousSnapshot !== serialized;
  
  return {
    payload: normalizeCardPayload(rawValues),
    serialized,
    hasChanges
  };
}

function primeCardSnapshots(items) {
  if (!Array.isArray(items) || items.length === 0) {
    return;
  }
  requestAnimationFrame(() => {
    items.forEach((item) => {
      const values = collectCardDomValues(item.id);
      if (!values) {
        return;
      }
      // Adiciona expiration_date ao snapshot (não é coletado do DOM porque tem classe 'calculated')
      if (item.expiration_date !== undefined) {
        values.expiration_date = item.expiration_date;
      }
      // Adiciona comments ao snapshot
      if (item.comments !== undefined) {
        values.comments = item.comments;
      }
      cardSnapshotStore.set(getSnapshotKey(item.id), serializeCardValues(values));
    });
  });
}

function scheduleInterfaceRefresh(reason = 'auto', delay = INTERFACE_REFRESH_DELAY_MS) {
  pendingInterfaceRefreshReason = reason;
  if (interfaceRefreshTimerId) {
    return;
  }
  interfaceRefreshTimerId = setTimeout(async () => {
    interfaceRefreshTimerId = null;
    try {
      await updateInterfaceAfterSave();
    } catch (error) {
    }
  }, Math.max(delay, 0));
}

function autoSaveTooling(id, immediate = false) {
  // Skip autosave during timeline reordering
  if (isReorderingTimeline) {
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
    const normalized = String(dateString).trim();
    const isoMatch = normalized.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$/);
    if (isoMatch) {
      const [, year, month, day] = isoMatch;
      return `${day}/${month}/${year}`;
    }

    const slashMatch = normalized.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);
    if (slashMatch) {
      const [, day, month, year] = slashMatch;
      return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/20${year}`;
    }

    const date = new Date(normalized);
    if (!Number.isNaN(date.getTime())) {
      return date.toLocaleDateString('pt-BR');
    }
    return normalized;
  } catch {
    return dateString;
  }
}

// Formata data e hora para exibição
function formatDateTime(dateString) {
  if (!dateString) return 'N/A';
  try {
    const normalized = String(dateString).trim();
    const isoMatch = normalized.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$/);
    if (isoMatch) {
      const [, year, month, day] = isoMatch;
      return `${day}/${month}/${year}`;
    }

    if (!normalized.includes('T')) {
      const fallbackDate = new Date(normalized);
      if (!Number.isNaN(fallbackDate.getTime())) {
        return fallbackDate.toLocaleDateString('pt-BR');
      }
    }

    let coerced = normalized;
    if (!/[zZ]$/.test(coerced)) {
      coerced += coerced.includes('T') ? 'Z' : 'T00:00:00Z';
    }
    const parsed = new Date(coerced);
    return Number.isNaN(parsed.getTime()) ? normalized : parsed.toLocaleDateString('pt-BR');
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
      return iso;
    }
  } else {
    const parsed = new Date(rawString);
    if (!Number.isNaN(parsed.getTime())) {
      const iso = parsed.toISOString().split('T')[0];
      return iso;
    }
  }
  return null;
}

function computeLocalChainMembership(data, externalIncomingLinks = []) {
  const membership = new Map();
  if (!Array.isArray(data) || data.length === 0) {
    return membership;
  }

  // Cria lookup de incoming links locais (dentro do mesmo supplier)
  const incomingLookup = new Map();
  data.forEach((item) => {
    const targetId = sanitizeReplacementId(item?.replacement_tooling_id);
    if (targetId) {
      incomingLookup.set(targetId, true);
    }
  });
  
  // Adiciona incoming links externos (de outros suppliers)
  if (Array.isArray(externalIncomingLinks)) {
    externalIncomingLinks.forEach(id => {
      incomingLookup.set(String(id), true);
    });
  }

  data.forEach((item) => {
    const itemId = String(item?.id || '').trim();
    if (!itemId) {
      return;
    }
    const hasOutgoingLink = Boolean(sanitizeReplacementId(item?.replacement_tooling_id));
    const hasIncomingLink = incomingLookup.has(itemId);
    membership.set(itemId, hasOutgoingLink || hasIncomingLink);
  });

  return membership;
}

// Versão async que processa em chunks pequenos sem travar a UI
async function computeChainMembershipAsync(data, targetMap, renderToken) {
  if (!Array.isArray(data) || data.length === 0) {
    return;
  }
  
  // Fase 1: Construir lookup de incoming links em chunks
  const incomingLookup = new Map();
  const CHUNK_SIZE = 100; // Processar 100 items por vez
  
  for (let i = 0; i < data.length; i += CHUNK_SIZE) {
    if (renderToken !== currentToolingRenderToken) return; // Cancelado
    
    const chunk = data.slice(i, Math.min(i + CHUNK_SIZE, data.length));
    chunk.forEach((item) => {
      const targetId = sanitizeReplacementId(item?.replacement_tooling_id);
      if (targetId) {
        incomingLookup.set(targetId, true);
      }
    });
    
    // Yield para não travar a UI
    if (i + CHUNK_SIZE < data.length) {
      await new Promise(resolve => setTimeout(resolve, 0));
    }
  }
  
  // Fase 2: Computar membership e atualizar ícones em chunks
  for (let i = 0; i < data.length; i += CHUNK_SIZE) {
    if (renderToken !== currentToolingRenderToken) return; // Cancelado
    
    const chunk = data.slice(i, Math.min(i + CHUNK_SIZE, data.length));
    chunk.forEach((item) => {
      const itemId = String(item?.id || '').trim();
      if (!itemId) return;
      
      const hasOutgoingLink = Boolean(sanitizeReplacementId(item?.replacement_tooling_id));
      const hasIncomingLink = incomingLookup.has(itemId);
      const inChain = hasOutgoingLink || hasIncomingLink;
      
      targetMap.set(itemId, inChain);
      
      // Atualizar ícone imediatamente se o card estiver renderizado
      const card = document.querySelector(`.tooling-card[data-item-id="${itemId}"]`);
      if (card) {
        card.dataset.hasIncomingChain = hasIncomingLink ? 'true' : 'false';
        card.dataset.chainMember = inChain ? 'true' : 'false';
        enforceChainIndicatorRules(card);
      }
    });
    
    // Yield para não travar a UI
    if (i + CHUNK_SIZE < data.length) {
      await new Promise(resolve => setTimeout(resolve, 0));
    }
  }
}

function updateAttachmentCountBadge(toolingId, count, renderToken) {
  if (!isActiveToolingRenderToken(renderToken)) {
    return;
  }
  const badge = document.querySelector(`.tooling-attachment-count[data-item-id="${toolingId}"]`);
  if (!badge) {
    return;
  }
  const valueElement = badge.querySelector('span');
  if (valueElement) {
    valueElement.textContent = String(Number.isFinite(count) ? count : 0);
  }
  if (count > 0) {
    badge.removeAttribute('hidden');
  } else {
    badge.setAttribute('hidden', 'true');
  }
}

async function hydrateAttachmentBadges(items, renderToken, supplierName) {
  if (!Array.isArray(items) || items.length === 0 || !supplierName) {
    return;
  }

  const normalizedItems = items.filter(item => Number.isFinite(Number(item?.id)));
  if (normalizedItems.length === 0) {
    return;
  }

  const supplierSnapshot = supplierName;
  const itemIds = normalizedItems.map(item => item.id);

  try {
    const counts = await window.api.getAttachmentsCountBatch(supplierSnapshot, itemIds);
    
    if (supplierSnapshot !== (selectedSupplier || currentSupplier || '')) {
      return;
    }

    normalizedItems.forEach((item) => {
      const count = counts[item.id] || 0;
      updateAttachmentCountBadge(item.id, count, renderToken);
    });
  } catch (error) {
  }
}

function updateChainIndicatorVisibility(toolingId, hasChain, renderToken) {
  if (!isActiveToolingRenderToken(renderToken)) {
    return;
  }
  const card = document.querySelector(`.tooling-card[data-item-id="${toolingId}"]`);
  if (!card) {
    return;
  }
  card.dataset.hasIncomingChain = hasChain ? 'true' : 'false';
  card.dataset.chainMember = hasChain ? 'true' : (card.dataset.chainMember || 'false');
  enforceChainIndicatorRules(card);
}

async function hydrateChainIndicators(items, renderToken, baseMap = new Map()) {
  if (!Array.isArray(items) || items.length === 0) {
    return;
  }

  const pendingItems = items.filter((item) => {
    const itemId = String(item?.id || '').trim();
    if (!itemId) {
      return false;
    }
    return !baseMap.get(itemId);
  });

  if (pendingItems.length === 0) {
    return;
  }

  const concurrency = Math.max(2, Math.floor(ASYNC_METADATA_CONCURRENCY / 2));

  await runTasksWithLimit(pendingItems, concurrency, async (item) => {
    try {
      const parents = await window.api.getToolingByReplacementId(Number(item.id));
      if (Array.isArray(parents) && parents.length > 0) {
        updateChainIndicatorVisibility(item.id, true, renderToken);
      }
    } catch (error) {
    }
  });
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

const INTEGER_FORMATTER = new Intl.NumberFormat('pt-BR', {
  minimumFractionDigits: 0,
  maximumFractionDigits: 0
});
let thousandsMaskInitialized = false;

function formatIntegerWithSeparators(value, options = {}) {
  const { preserveEmpty = false } = options;
  if (value === null || value === undefined) {
    return preserveEmpty ? '' : '';
  }

  if (preserveEmpty && typeof value === 'string' && value.trim() === '') {
    return '';
  }

  const numericValue = typeof value === 'number'
    ? value
    : parseLocalizedNumber(value);

  if (!Number.isFinite(numericValue)) {
    return preserveEmpty ? '' : '';
  }

  return INTEGER_FORMATTER.format(Math.trunc(numericValue));
}

function extractDigits(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).replace(/\D/g, '');
}

function applyThousandsMaskToInput(input, force = false) {
  if (!input) {
    return;
  }
  if (!force && document.activeElement === input) {
    return;
  }
  const digits = extractDigits(input.value);
  if (!digits) {
    input.value = '';
    return;
  }
  input.value = formatIntegerWithSeparators(digits);
}

function handleThousandsMaskFocus(event) {
  const input = event.target;
  if (!(input instanceof HTMLInputElement) || input.getAttribute('data-mask') !== 'thousands') {
    return;
  }
  const digits = extractDigits(input.value);
  input.value = digits;
  requestAnimationFrame(() => {
    const caret = input.value.length;
    if (typeof input.setSelectionRange === 'function') {
      input.setSelectionRange(caret, caret);
    }
  });
}

function handleThousandsMaskInput(event) {
  const input = event.target;
  if (!(input instanceof HTMLInputElement) || input.getAttribute('data-mask') !== 'thousands') {
    return;
  }
  const digits = extractDigits(input.value);
  if (digits !== input.value) {
    input.value = digits;
    requestAnimationFrame(() => {
      const caret = input.value.length;
      if (typeof input.setSelectionRange === 'function') {
        input.setSelectionRange(caret, caret);
      }
    });
  }
}

function handleThousandsMaskBlur(event) {
  const input = event.target;
  if (!(input instanceof HTMLInputElement) || input.getAttribute('data-mask') !== 'thousands') {
    return;
  }
  applyThousandsMaskToInput(input, true);
}

function applyInitialThousandsMask(root = document) {
  if (!root || typeof root.querySelectorAll !== 'function') {
    return;
  }
  const inputs = root.querySelectorAll('[data-mask="thousands"]');
  inputs.forEach((input) => applyThousandsMaskToInput(input, true));
}

function initThousandsMaskBehavior() {
  if (thousandsMaskInitialized) {
    applyInitialThousandsMask();
    return;
  }
  thousandsMaskInitialized = true;
  document.addEventListener('focusin', handleThousandsMaskFocus, true);
  document.addEventListener('input', handleThousandsMaskInput, true);
  document.addEventListener('focusout', handleThousandsMaskBlur, true);
  applyInitialThousandsMask();
}

// Recalcula todas as expiration dates ao carregar os dados
async function recalculateAllExpirationDates() {
  if (!toolingData || toolingData.length === 0) return;
  
  const updates = [];
  
  for (const item of toolingData) {
    const toolingLife = parseLocalizedNumber(item.tooling_life_qty) || 0;
    const produced = parseLocalizedNumber(item.produced) || 0;
    const remaining = toolingLife - produced;
    const forecast = parseLocalizedNumber(item.annual_volume_forecast) || 0;
    const productionDate = item.date_remaining_tooling_life || '';
    
    const newExpirationDate = calculateExpirationFromFormula({
      remaining,
      forecast,
      productionDate
    });
    
    // Só atualiza se a data calculada for diferente da atual
    const currentExpiration = item.expiration_date || '';
    if (newExpirationDate !== currentExpiration) {
      item.expiration_date = newExpirationDate || '';
      updates.push({
        id: item.id,
        expiration_date: newExpirationDate || ''
      });
    }
  }
  
  // Salva todas as atualizações no banco de dados em background
  if (updates.length > 0) {
    for (const update of updates) {
      try {
        await window.api.updateTooling(update.id, { expiration_date: update.expiration_date });
      } catch (error) {
        console.error('Error updating expiration_date for item', update.id, error);
      }
    }
    console.log(`Recalculated ${updates.length} expiration dates on startup`);
  }
}

function calculateExpirationFromFormula({
  remaining,
  forecast,
  productionDate
}) {
  // Data base é a data de produção (obrigatória)
  if (!productionDate) {
    return null;
  }
  
  // Se não houver forecast (annual volume), não calcula
  if (!forecast || forecast <= 0) {
    return null;
  }
  
  const baseDate = new Date(productionDate);
  if (Number.isNaN(baseDate.getTime())) {
    return null;
  }

  // Fórmula: data_produced + (remaining/annual_volume*365)
  // remaining pode ser negativo (já expirou) - o cálculo vai resultar em data passada
  const totalDays = Math.round((remaining / forecast) * 365);

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

function populateAddToolingOwners() {
  const list = addToolingElements.ownerList;
  if (!list) {
    return;
  }

  list.innerHTML = '';

  // Get unique owners from responsiblesData (loaded from DB)
  const uniqueOwners = Array.isArray(responsiblesData) 
    ? [...responsiblesData].sort((a, b) => a.localeCompare(b, 'pt-BR'))
    : [];

  uniqueOwners.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    list.appendChild(option);
  });
}

function openAddToolingModal() {
  const { overlay, pnInput, supplierInput, ownerInput, lifeInput, producedInput } = addToolingElements;
  if (!overlay) {
    return;
  }

  populateAddToolingSuppliers();
  populateAddToolingOwners();

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
  const { overlay, pnInput, supplierInput, ownerInput, lifeInput, producedInput } = addToolingElements;
  const pnDescriptionInput = document.getElementById('addToolingPNDescription');
  const descriptionInput = document.getElementById('addToolingDescription');
  const forecastInput = document.getElementById('addToolingForecast');
  const productionDateInput = document.getElementById('addToolingProductionDate');
  const forecastDateInput = document.getElementById('addToolingForecastDate');
  
  if (overlay) overlay.classList.remove('active');
  if (pnInput) pnInput.value = '';
  if (pnDescriptionInput) pnDescriptionInput.value = '';
  if (supplierInput) {
    supplierInput.value = '';
    const defaultPlaceholder = supplierInput.getAttribute('data-default-placeholder');
    if (defaultPlaceholder) {
      supplierInput.placeholder = defaultPlaceholder;
    }
  }
  if (ownerInput) ownerInput.value = '';
  if (descriptionInput) descriptionInput.value = '';
  if (lifeInput) lifeInput.value = '';
  if (producedInput) producedInput.value = '';
  if (forecastInput) forecastInput.value = '';
  if (productionDateInput) productionDateInput.value = '';
  if (forecastDateInput) forecastDateInput.value = '';
}

function buildCommentsListHTML(commentsJson, itemId, filterText = null) {
  let comments = [];

  if (commentsJson) {
    try {
      comments = JSON.parse(commentsJson);
      if (!Array.isArray(comments)) {
        comments = [];
      }
    } catch (e) {
      comments = [];
    }
  }

  // Manter índices originais para edição/exclusão antes de filtrar
  const indexedComments = comments.map((comment, idx) => ({ ...comment, originalIndex: idx }));

  // Aplicar filtro por texto se fornecido
  let filteredComments = indexedComments;
  if (filterText && filterText !== 'all') {
    const searchTerm = filterText.toLowerCase();
    filteredComments = indexedComments.filter(comment => {
      const commentText = (comment.text || '').toLowerCase();
      return commentText.includes(searchTerm);
    });
  }

  // Inverter ordem para mostrar comentários mais recentes primeiro (do topo para baixo)
  filteredComments = filteredComments.reverse();

  if (filteredComments.length === 0) {
    const emptyMessage = filterText && filterText !== 'all' 
      ? 'No comments matching this filter'
      : 'No comments yet';
    return `<div class="comments-empty">${emptyMessage}</div>`;
  }

  return filteredComments.map((comment) => {
    const index = comment.originalIndex;
    const date = comment.date || 'N/A';
    const text = escapeHtml(comment.text || '');
    const isInitial = comment.initial === true;
    const isImported = comment.origin === 'import';
    const cardClasses = ['comment-card'];
    if (isInitial) {
      cardClasses.push('comment-initial');
    }
    if (isImported) {
      cardClasses.push('comment-imported');
    }

    let headerRight = '';
    if (isInitial) {
      headerRight = '<span class="comment-initial-badge">Initial</span>';
    } else if (isImported) {
      headerRight = `
        <div class="comment-actions comment-actions--readonly" title="Imported from spreadsheet">
          <i class="ph ph-cloud-arrow-down" aria-hidden="true"></i>
        </div>
      `;
    } else {
      headerRight = `
        <div class="comment-actions">
          <button class="btn-comment-action" onclick="editComment(${itemId}, ${index})" title="Edit">
            <i class="ph ph-pencil-simple"></i>
          </button>
          <button class="btn-comment-action" onclick="openCommentDeleteModal(${itemId}, ${index})" title="Delete">
            <i class="ph ph-trash"></i>
          </button>
        </div>
      `;
    }

    return `
      <div class="${cardClasses.join(' ')}" data-comment-index="${index}">
        <div class="comment-header">
          <span class="comment-date${isImported ? ' comment-date--imported' : ''}">
            <span class="comment-date-text">${escapeHtml(date)}</span>
          </span>
          ${headerRight}
        </div>
        <div class="comment-divider"></div>
        <div class="comment-text" id="commentText_${itemId}_${index}">${text}</div>
      </div>
    `;
  }).join('');
}

function handleCommentKeydown(event, itemId) {
  if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
    event.preventDefault();
    addComment(itemId);
  }
  // Enter normal quebra a linha; Ctrl/Cmd+Enter envia
}

function addComment(itemId) {
  const input = document.getElementById(`commentInput_${itemId}`);
  if (!input) return;
  
  const text = input.value.trim();
  if (!text) return;
  
  const item = toolingData.find(item => Number(item.id) === Number(itemId));
  if (!item) return;
  
  let comments = [];
  if (item.comments) {
    try {
      comments = JSON.parse(item.comments);
      if (!Array.isArray(comments)) {
        comments = [];
      }
    } catch (e) {
      comments = [];
    }
  }
  
  const now = new Date();
  const dateStr = now.toLocaleString('pt-BR', { 
    day: '2-digit', 
    month: '2-digit', 
    year: 'numeric', 
    hour: '2-digit', 
    minute: '2-digit' 
  });
  
  comments.push({
    date: dateStr,
    text: text,
    initial: false
  });
  
  item.comments = JSON.stringify(comments);
  input.value = '';
  
  updateCommentsDisplay(itemId);
  autoSaveTooling(itemId, true);
}

function editComment(itemId, commentIndex) {
  const item = toolingData.find(item => Number(item.id) === Number(itemId));
  if (!item) return;
  
  let comments = [];
  if (item.comments) {
    try {
      comments = JSON.parse(item.comments);
      if (!Array.isArray(comments)) {
        return;
      }
    } catch (e) {
      return;
    }
  }
  
  if (commentIndex < 0 || commentIndex >= comments.length) return;
  if (comments[commentIndex].initial) return;
  if (comments[commentIndex].origin === 'import') {
    showNotification('Imported comments cannot be edited', 'info');
    return;
  }
  
  const textElement = document.getElementById(`commentText_${itemId}_${commentIndex}`);
  if (!textElement) return;
  
  const currentText = comments[commentIndex].text || '';
  const card = textElement.closest('.comment-card');
  
  const editFieldHtml = `<textarea class="comment-edit-input" id="commentEditInput_${itemId}_${commentIndex}" rows="4" onkeydown="handleCommentEditKeydown(event, ${itemId}, ${commentIndex})">${escapeHtml(currentText)}</textarea>`;
  textElement.innerHTML = editFieldHtml;
  
  const editInput = document.getElementById(`commentEditInput_${itemId}_${commentIndex}`);
  if (editInput) {
    editInput.focus();
    editInput.select();
  }
  
  const actions = card.querySelector('.comment-actions');
  if (actions) {
    actions.innerHTML = `
      <button class="btn-comment-action" onclick="saveCommentEdit(${itemId}, ${commentIndex})" title="Save">
        <i class="ph ph-check"></i>
      </button>
      <button class="btn-comment-action" onclick="cancelCommentEdit(${itemId})" title="Cancel">
        <i class="ph ph-x"></i>
      </button>
    `;
  }
}

function handleCommentEditKeydown(event, itemId, commentIndex) {
  if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
    event.preventDefault();
    saveCommentEdit(itemId, commentIndex);
  } else if (event.key === 'Escape') {
    event.preventDefault();
    cancelCommentEdit(itemId);
  }
}

function saveCommentEdit(itemId, commentIndex) {
  const editInput = document.getElementById(`commentEditInput_${itemId}_${commentIndex}`);
  if (!editInput) return;
  
  const newText = editInput.value.trim();
  if (!newText) {
    showNotification('Comment cannot be empty', 'error');
    return;
  }
  
  const item = toolingData.find(item => Number(item.id) === Number(itemId));
  if (!item) return;
  
  let comments = [];
  if (item.comments) {
    try {
      comments = JSON.parse(item.comments);
      if (!Array.isArray(comments)) {
        return;
      }
    } catch (e) {
      return;
    }
  }
  
  if (commentIndex < 0 || commentIndex >= comments.length) return;
  
  // Atualiza o texto e a data do comentário
  const now = new Date();
  const dateStr = now.toLocaleString('pt-BR', { 
    day: '2-digit', 
    month: '2-digit', 
    year: 'numeric', 
    hour: '2-digit', 
    minute: '2-digit' 
  });
  
  comments[commentIndex].text = newText;
  comments[commentIndex].date = dateStr;
  item.comments = JSON.stringify(comments);
  
  updateCommentsDisplay(itemId);
  autoSaveTooling(itemId, true);
}

function cancelCommentEdit(itemId) {
  updateCommentsDisplay(itemId);
}

function getItemCommentsArray(item) {
  if (!item || !item.comments) {
    return [];
  }
  try {
    const parsed = JSON.parse(item.comments);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function ensureCommentDeleteElements() {
  if (!commentDeleteElements.overlay) {
    commentDeleteElements = {
      overlay: document.getElementById('commentDeleteOverlay'),
      context: document.getElementById('commentDeleteContext'),
      date: document.getElementById('commentDeleteDate'),
      text: document.getElementById('commentDeleteText'),
      confirmButton: document.getElementById('commentDeleteConfirmBtn')
    };
  }
  return commentDeleteElements;
}

function resetCommentDeleteState() {
  commentDeleteState = { itemId: null, commentIndex: null };
}

function formatCommentDeleteContext(item) {
  if (!item) {
    return 'Tooling';
  }
  const parts = [];
  if (item.id !== undefined) {
    parts.push(`#${item.id}`);
  }
  if (item.pn) {
    parts.push(item.pn);
  }
  if (item.supplier) {
    parts.push(item.supplier);
  }
  return parts.join(' • ') || 'Tooling';
}

function openCommentDeleteModal(itemId, commentIndex) {
  const item = toolingData.find(entry => Number(entry.id) === Number(itemId));
  if (!item) {
    return;
  }

  const comments = getItemCommentsArray(item);
  if (commentIndex < 0 || commentIndex >= comments.length) {
    return;
  }

  const targetComment = comments[commentIndex];
  if (!targetComment) {
    return;
  }

  if (targetComment.initial) {
    showNotification('Initial comments cannot be deleted', 'info');
    return;
  }

  if (targetComment.origin === 'import') {
    showNotification('Imported comments cannot be deleted', 'info');
    return;
  }

  commentDeleteState = { itemId, commentIndex };
  const elements = ensureCommentDeleteElements();

  if (elements.context) {
    elements.context.textContent = formatCommentDeleteContext(item);
  }
  if (elements.date) {
    elements.date.textContent = targetComment.date || 'No date available';
  }
  if (elements.text) {
    const trimmedText = (targetComment.text || '').trim();
    elements.text.textContent = trimmedText || 'No content in this comment.';
  }

  if (elements.overlay) {
    elements.overlay.classList.add('active');
  }

  requestAnimationFrame(() => {
    elements.confirmButton?.focus();
  });
}

function closeCommentDeleteModal() {
  const elements = ensureCommentDeleteElements();
  elements.overlay?.classList.remove('active');
  if (elements.text) {
    elements.text.textContent = '';
  }
  if (elements.date) {
    elements.date.textContent = '';
  }
  resetCommentDeleteState();
}

function confirmCommentDelete() {
  if (commentDeleteState.itemId === null || commentDeleteState.commentIndex === null) {
    return;
  }
  const deleted = deleteComment(commentDeleteState.itemId, commentDeleteState.commentIndex);
  if (deleted) {
    closeCommentDeleteModal();
  }
}

function deleteComment(itemId, commentIndex) {
  const item = toolingData.find(entry => Number(entry.id) === Number(itemId));
  if (!item) {
    return false;
  }

  const comments = getItemCommentsArray(item);
  if (commentIndex < 0 || commentIndex >= comments.length) {
    return false;
  }

  if (comments[commentIndex]?.initial) {
    showNotification('Initial comments cannot be deleted', 'info');
    return false;
  }

  if (comments[commentIndex]?.origin === 'import') {
    showNotification('Imported comments cannot be deleted', 'info');
    return false;
  }

  comments.splice(commentIndex, 1);
  item.comments = JSON.stringify(comments);

  updateCommentsDisplay(itemId);
  autoSaveTooling(itemId, true);
  showNotification('Comment deleted', 'success');
  return true;
}

function updateCommentsDisplay(itemId) {
  const item = toolingData.find(item => Number(item.id) === Number(itemId));
  if (!item) return;
  
  const commentsList = document.getElementById(`commentsList_${itemId}`);
  if (!commentsList) return;
  
  // Obter o filtro atual, se houver
  const filterBtn = document.getElementById(`commentsFilterBtn_${itemId}`);
  const currentFilter = filterBtn ? filterBtn.dataset.currentFilter : null;
  
  commentsList.innerHTML = buildCommentsListHTML(item.comments || '', itemId, currentFilter);
}

function toggleCommentsFilterPopup(itemId) {
  const popup = document.getElementById(`commentsFilterPopup_${itemId}`);
  if (!popup) return;
  
  // Fechar outros popups abertos
  document.querySelectorAll('.comments-filter-popup.active').forEach(p => {
    if (p.id !== `commentsFilterPopup_${itemId}`) {
      p.classList.remove('active');
    }
  });
  
  popup.classList.toggle('active');
}

function applyCommentsFilter(itemId, filterText) {
  const item = toolingData.find(item => Number(item.id) === Number(itemId));
  if (!item) return;
  
  const commentsList = document.getElementById(`commentsList_${itemId}`);
  if (!commentsList) return;
  
  // Guardar filtro atual no botão
  const filterBtn = document.getElementById(`commentsFilterBtn_${itemId}`);
  if (filterBtn) {
    filterBtn.dataset.currentFilter = filterText;
    // Destacar ícone se filtro ativo
    if (filterText && filterText !== 'all') {
      filterBtn.classList.add('filter-active');
    } else {
      filterBtn.classList.remove('filter-active');
    }
  }
  
  // Fechar popup
  const popup = document.getElementById(`commentsFilterPopup_${itemId}`);
  if (popup) popup.classList.remove('active');
  
  commentsList.innerHTML = buildCommentsListHTML(item.comments || '', itemId, filterText);
}

async function submitAddToolingForm() {
  const { form, pnInput, supplierInput, ownerInput, lifeInput, producedInput } = addToolingElements;
  const pnDescriptionInput = document.getElementById('addToolingPNDescription');
  const descriptionInput = document.getElementById('addToolingDescription');
  const forecastInput = document.getElementById('addToolingForecast');
  const productionDateInput = document.getElementById('addToolingProductionDate');
  const forecastDateInput = document.getElementById('addToolingForecastDate');

  if (!pnInput || !supplierInput || !lifeInput || !producedInput) {
    return;
  }

  if (form && typeof form.reportValidity === 'function' && !form.reportValidity()) {
    return;
  }

  const pn = pnInput.value.trim();
  const pnDescription = pnDescriptionInput ? pnDescriptionInput.value.trim() : '';
  const supplier = supplierInput.value.trim();
  const owner = ownerInput ? ownerInput.value.trim() : '';
  const toolDescription = descriptionInput ? descriptionInput.value.trim() : '';
  const toolingLife = parseLocalizedNumber(lifeInput.value);
  const produced = parseLocalizedNumber(producedInput.value);
  const forecast = forecastInput ? parseLocalizedNumber(forecastInput.value) : 0;
  const productionDate = productionDateInput ? productionDateInput.value : null;
  const forecastDate = forecastDateInput ? forecastDateInput.value : null;

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
    // Criar comentário inicial
    const now = new Date();
    const dateStr = now.toLocaleString('pt-BR', { 
      day: '2-digit', 
      month: '2-digit', 
      year: 'numeric', 
      hour: '2-digit', 
      minute: '2-digit' 
    });
    
    const initialComment = {
      date: dateStr,
      text: `Created with Tooling Life: ${formatIntegerWithSeparators(toolingLife)} pcs`,
      initial: true
    };
    
    const commentsJson = JSON.stringify([initialComment]);
    
    const payload = {
      pn,
      pn_description: pnDescription,
      supplier,
      cummins_responsible: owner || null,
      tool_description: toolDescription,
      tooling_life_qty: toolingLife,
      produced,
      date_remaining_tooling_life: productionDate,
      annual_volume_forecast: forecast > 0 ? forecast : null,
      date_annual_volume: forecastDate,
      status: 'ACTIVE',
      comments: commentsJson
    };
    const result = await window.api.createTooling(payload);

    if (!result || result.success !== true) {
      showNotification(result?.error || 'Não foi possível criar o ferramental.', 'error');
      return;
    }

    closeAddToolingModal();

    selectedSupplier = supplier;
    currentSupplier = supplier;

    // Salvar filtro ativo antes de recarregar
    const supplierSearchInput = document.getElementById('supplierSearchInput');
    const activeFilter = supplierSearchInput ? supplierSearchInput.value.trim() : '';

    await loadSuppliers();
    await loadAnalytics();
    await refreshReplacementIdOptions(true);
    await loadToolingBySupplier(supplier);
    await loadAttachments(supplier);
    
    // Reaplicar filtro se estava ativo
    if (activeFilter.length >= 1) {
      await filterSuppliersAndTooling(activeFilter);
    }

    const attachmentsContainer = document.getElementById('attachmentsContainer');
    const currentSupplierName = document.getElementById('currentSupplierName');
    if (attachmentsContainer && currentSupplierName) {
      attachmentsContainer.style.display = 'block';
      currentSupplierName.textContent = supplier;
    }

    showNotification('Ferramental criado com sucesso!');
  } catch (error) {
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
      // loadToolingBySupplier já atualiza as métricas via displayTooling
      await loadToolingBySupplier(selectedSupplier);
    } else {
      displayTooling([]);
    }
  } catch (error) {
    showNotification('Erro ao excluir ferramental', 'error');
  }
}

// Filtra tooling por Step (global - suppliers e tooling cards)
async function filterToolingByStep() {
  const stepsFilter = document.getElementById('stepsFilter');
  const selectedStep = stepsFilter ? stepsFilter.value : '';
  
  console.log('Filtering by step:', selectedStep);
  
  // Remover mensagem de "no steps" se existir
  const existingMsg = document.getElementById('noStepsMessage');
  if (existingMsg) existingMsg.remove();
  
  if (!selectedStep) {
    // Sem filtro - limpar lista de suppliers filtrados
    stepsFilteredSuppliers = null;
    
    // Mostrar todos os supplier cards e restaurar métricas originais
    const supplierCards = document.querySelectorAll('.supplier-card');
    supplierCards.forEach(card => {
      card.style.display = '';
    });
    
    // Recarregar suppliers para restaurar os contadores originais
    await loadSuppliers();
    
    // Recarregar tooling cards do supplier selecionado (sem filtro)
    if (selectedSupplier) {
      await loadToolingBySupplier(selectedSupplier);
    }
    
    return;
  }
  
  // COM filtro - filtrar suppliers e atualizar contadores
  try {
    // Buscar todos os suppliers com stats
    const allSuppliers = await window.api.getSuppliersWithStats();
    console.log('All suppliers:', allSuppliers.length);
    
    // Para cada supplier, verificar se tem tooling com o step selecionado e calcular métricas
    const suppliersWithStep = [];
    const supplierMetrics = new Map();
    
    for (const supplierObj of allSuppliers) {
      const supplierName = supplierObj.supplier || '';
      const supplierTooling = await window.api.getToolingBySupplier(supplierName);
      
      // Filtrar tooling pelo step selecionado
      const filteredTooling = supplierTooling.filter(item => {
        const itemStep = String(item.steps || '').trim();
        return itemStep === selectedStep;
      });
      
      if (filteredTooling.length > 0) {
        suppliersWithStep.push(supplierName);
        // Calcular métricas apenas dos itens filtrados
        const metrics = ExpirationMetrics.fromItems(filteredTooling);
        supplierMetrics.set(supplierName, metrics);
      }
    }
    
    console.log('Suppliers with step:', suppliersWithStep);
    
    // Salvar lista de suppliers filtrados globalmente
    stepsFilteredSuppliers = suppliersWithStep;
    
    // Atualizar suppliers na sidebar (mostrar apenas os que têm o step e atualizar contadores)
    const supplierCards = document.querySelectorAll('.supplier-card');
    let visibleSuppliers = 0;
    supplierCards.forEach(card => {
      const supplierName = card.dataset.supplier || '';
      if (suppliersWithStep.includes(supplierName)) {
        card.style.display = '';
        visibleSuppliers++;
        
        // Atualizar contadores com valores filtrados
        const metrics = supplierMetrics.get(supplierName);
        if (metrics) {
          const totalEl = card.querySelector('[data-metric="total"]');
          const expiredEl = card.querySelector('[data-metric="expired"]');
          const expiringEl = card.querySelector('[data-metric="expiring"]');
          
          if (totalEl) totalEl.textContent = metrics.total;
          if (expiredEl) {
            expiredEl.textContent = metrics.expired;
            expiredEl.classList.toggle('expired', metrics.expired > 0);
          }
          if (expiringEl) {
            expiringEl.textContent = metrics.expiring;
            expiringEl.classList.toggle('critical', metrics.expiring > 0);
          }
        }
      } else {
        card.style.display = 'none';
      }
    });
    console.log('Visible suppliers:', visibleSuppliers);
    
    // Se tiver supplier selecionado, recarregar os tooling cards (já filtrados)
    if (selectedSupplier) {
      await loadToolingBySupplier(selectedSupplier);
    }
    
  } catch (error) {
    console.error('Error filtering by step:', error);
  }
}

// Exibe ferramentais na interface (cards expansíveis)
async function displayTooling(data) {
  const toolingList = document.getElementById('toolingList');
  const spreadsheetContainer = document.getElementById('spreadsheetContainer');
  const emptyState = document.getElementById('emptyState');

  if (!data || data.length === 0) {
    toolingList.style.display = 'none';
    if (spreadsheetContainer) spreadsheetContainer.style.display = 'none';
    emptyState.style.display = 'flex';
    return;
  }

  // Mostra UI imediatamente, processa em background
  emptyState.style.display = 'none';

  // Defer processing para não travar
  await new Promise(resolve => setTimeout(resolve, 0));

  // Aplica filtro de expiração se estiver ativo
  let filteredData = data;
  if (expirationFilterEnabled) {
    filteredData = data.filter(item => {
      // Ignora itens com análise concluída
      if (item.analysis_completed === 1) {
        return false;
      }
      const classification = classifyToolingExpirationState(item);
      return classification.state === 'expired' || classification.state === 'warning';
    });
  }

  if (filteredData.length === 0) {
    toolingList.style.display = 'none';
    if (spreadsheetContainer) spreadsheetContainer.style.display = 'none';
    emptyState.style.display = 'flex';
    return;
  }

  // Yield antes do sort pesado
  await new Promise(resolve => setTimeout(resolve, 0));

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

  cardSnapshotStore.clear();
  toolingData = sortedData;
  
  // Atualiza métricas do supplier card com os dados ORIGINAIS (não filtrados)
  // para manter o Total correto e recalcular expired/expiring ignorando análise concluída
  if (selectedSupplier) {
    updateSupplierCardMetricsFromItems(selectedSupplier, data);
  }
  
  // Se estiver no modo planilha, renderiza planilha ao invés de cards
  if (currentViewMode === 'spreadsheet') {
    toolingList.style.display = 'none';
    if (spreadsheetContainer) spreadsheetContainer.style.display = 'block';
    renderSpreadsheetView();
    return;
  }
  
  // Modo cards - continua renderização normal
  toolingList.style.display = 'flex';
  if (spreadsheetContainer) spreadsheetContainer.style.display = 'none';
  toolingList.innerHTML = '<div class="loading-cards">Loading cards...</div>';
  
  currentToolingRenderToken += 1;
  const renderToken = currentToolingRenderToken;
  const supplierContext = selectedSupplier || currentSupplier || '';
  
  // Computa chain membership incluindo incoming links de outros suppliers
  const chainMembership = computeLocalChainMembership(sortedData, externalIncomingLinks);
  
  // Limpa loading e começa render
  toolingList.innerHTML = '';
  
  // CHUNK PEQUENO: cada card gera ~500 linhas de HTML!
  // 10 cards = ~5000 linhas por vez, não trava
  const CHUNK_SIZE = 10;
  let currentIndex = 0;
  
  const renderNextChunk = () => {
    if (renderToken !== currentToolingRenderToken) {
      return; // Render foi cancelado
    }
    
    const endIndex = Math.min(currentIndex + CHUNK_SIZE, sortedData.length);
    const chunk = sortedData.slice(currentIndex, endIndex);
    
    // Gera apenas HEADERS (super leve!) - body carrega on-demand
    const htmlChunks = [];
    chunk.forEach((item, relativeIndex) => {
      const index = currentIndex + relativeIndex;
      htmlChunks.push(buildToolingCardHeaderHTML(item, index, chainMembership));
    });
    
    // Insere tudo de uma vez
    toolingList.insertAdjacentHTML('beforeend', htmlChunks.join(''));
    
    currentIndex = endIndex;
    
    if (currentIndex < sortedData.length) {
      setTimeout(renderNextChunk, 0);
    } else {
      // Render completo - AGORA computa chain em background sem travar
      computeChainMembershipAsync(sortedData, chainMembership, renderToken);
      
      // Hidrata dados em background
      setTimeout(() => {
        if (renderToken !== currentToolingRenderToken) return;
        hydrateCardsAfterRender(sortedData, renderToken, supplierContext, chainMembership);
      }, 0);
    }
  };
  
  renderNextChunk();
}

// Toggle do sidebar (abrir/fechar)
function toggleSidebar() {
  const sidebar = document.getElementById('sidebar');
  const toggleIcon = document.getElementById('sidebarToggleIcon');
  
  if (!sidebar) return;
  
  const willCollapse = !sidebar.classList.contains('collapsed');
  
  // Aplica a classe de transição
  sidebar.classList.toggle('collapsed');
  
  // Rotaciona o ícone via CSS transform após a transição
  if (toggleIcon) {
    toggleIcon.style.transform = willCollapse ? 'rotate(180deg)' : 'rotate(0deg)';
  }
  
  // Salva estado no localStorage
  localStorage.setItem('sidebarCollapsed', willCollapse ? 'true' : 'false');
}

// Restaura estado do sidebar ao carregar
function restoreSidebarState() {
  const sidebar = document.getElementById('sidebar');
  const toggleIcon = document.getElementById('sidebarToggleIcon');
  const isCollapsed = localStorage.getItem('sidebarCollapsed') === 'true';
  
  if (sidebar && isCollapsed) {
    sidebar.classList.add('collapsed');
    if (toggleIcon) {
      toggleIcon.style.transform = 'rotate(180deg)';
    }
  }
}

// Alterna entre modo cards e planilha
function setViewMode(mode) {
  // Força sempre o modo planilha com linhas expansíveis
  currentViewMode = 'spreadsheet';
  
  // Atualiza containers
  const toolingList = document.getElementById('toolingList');
  const spreadsheetContainer = document.getElementById('spreadsheetContainer');
  const emptyState = document.getElementById('emptyState');
  const floatingAddBtn = document.getElementById('floatingAddBtn');
  
  // Esconde o botão flutuante de adicionar (usamos a linha de criação na planilha)
  if (floatingAddBtn) {
    floatingAddBtn.style.display = 'none';
  }
  
  if (!toolingData || toolingData.length === 0) {
    if (toolingList) toolingList.style.display = 'none';
    if (spreadsheetContainer) spreadsheetContainer.style.display = 'none';
    if (emptyState) emptyState.style.display = 'flex';
    return;
  }
  
  if (emptyState) emptyState.style.display = 'none';
  
  // Sempre mostra apenas a planilha
  if (toolingList) toolingList.style.display = 'none';
  if (spreadsheetContainer) spreadsheetContainer.style.display = 'block';
  renderSpreadsheetView();
}

// Renderiza a visualização de planilha
function renderSpreadsheetView() {
  const spreadsheetBody = document.getElementById('spreadsheetBody');
  if (!spreadsheetBody || !toolingData) return;
  
  // Aplica filtro de expiração se estiver ativo
  let filteredData = toolingData;
  if (expirationFilterEnabled) {
    filteredData = toolingData.filter(item => {
      const classification = classifyToolingExpirationState(item);
      return classification.state === 'expired' || classification.state === 'warning';
    });
  }
  
  // Aplica filtros de coluna
  filteredData = applyColumnFiltersToData(filteredData);
  
  // Aplica ordenação
  filteredData = applySortToData(filteredData);
  
  // Computa chain membership usando TODOS os dados (não apenas filtrados)
  // e inclui incoming links de outros suppliers
  const chainMembership = computeLocalChainMembership(toolingData, externalIncomingLinks);
  
  // Gera as linhas da planilha
  const rows = filteredData.map(item => {
    const statusClass = getStatusClass(item.status);
    const statusOptionsHtml = buildSpreadsheetStatusOptions(item.status);
    const stepsOptionsHtml = buildSpreadsheetStepsOptions(item.steps);
    
    // Formata valores numéricos para exibição com separador de milhares
    const toolingLifeDisplay = formatNumericForSpreadsheet(item.tooling_life_qty);
    const producedDisplay = formatNumericForSpreadsheet(item.produced);
    const forecastDisplay = formatNumericForSpreadsheet(item.annual_volume_forecast);
    
    // Calcula expiration date e progresso
    const classification = classifyToolingExpirationState(item);
    const hasExpirationDate = classification.expirationDate && classification.expirationDate !== '';
    const expirationDateDisplay = hasExpirationDate ? formatDate(classification.expirationDate) : '';
    const isAnalysisCompleted = item.analysis_completed === 1;
    const expirationIconHtml = hasExpirationDate ? getSpreadsheetExpirationIcon(classification.state, isAnalysisCompleted) : '';
    
    // Calcula progresso
    const toolingLife = Number(parseLocalizedNumber(item.tooling_life_qty)) || Number(item.tooling_life_qty) || 0;
    const produced = Number(parseLocalizedNumber(item.produced)) || Number(item.produced) || 0;
    const percentUsedValue = toolingLife > 0 ? (produced / toolingLife) * 100 : 0;
    const percentUsed = toolingLife > 0 ? percentUsedValue.toFixed(1) : '0';
    
    // Verifica chain membership e anexos
    const membershipKey = String(item.id || '').trim();
    const hasChainMembership = chainMembership?.get(membershipKey) === true;
    const replacementIdValue = sanitizeReplacementId(item.replacement_tooling_id);
    const hasReplacementLink = replacementIdValue !== '';
    
    const isSelected = selectedToolingIds.has(item.id);
    const checkboxHtml = selectionModeActive ? `
      <td class="col-checkbox">
        <div class="spreadsheet-row-checkbox ${isSelected ? 'selected' : ''}" onclick="toggleSpreadsheetRowSelection(event, ${item.id})">
          <i class="ph ${isSelected ? 'ph-check-square' : 'ph-square'}"></i>
        </div>
      </td>` : '';
    
    // Encontra o índice do item para usar nas funções de card
    const itemIndex = toolingData.findIndex(t => t.id === item.id);
    
    // Prepara informações para o tooltip do ID
    const tooltipPnDesc = item.pn_description ? escapeHtml(item.pn_description) : 'N/A';
    const tooltipSupplier = item.supplier ? escapeHtml(item.supplier) : 'N/A';
    const tooltipAssetNumber = item.asset_number ? escapeHtml(item.asset_number) : 'N/A';
    const tooltipProdDate = item.date_remaining_tooling_life ? formatDate(item.date_remaining_tooling_life) : 'N/A';
    const tooltipVolDate = item.date_annual_volume ? formatDate(item.date_annual_volume) : 'N/A';
    const tooltipLastUpdate = item.last_update ? formatDate(item.last_update) : 'N/A';
    
    return `
      <tr data-id="${item.id}" data-item-index="${itemIndex}" class="${isSelected ? 'row-selected' : ''}">
        ${checkboxHtml}
        <td class="col-id id-with-tooltip">
          <span class="id-number">${item.id || ''}</span>
          <div class="id-tooltip">
            <div class="id-tooltip-header">Tooling #${item.id}</div>
            <div class="id-tooltip-item"><span class="tooltip-label">PN Description:</span> ${tooltipPnDesc}</div>
            <div class="id-tooltip-item"><span class="tooltip-label">Supplier:</span> ${tooltipSupplier}</div>
            <div class="id-tooltip-item"><span class="tooltip-label">Asset Number:</span> ${tooltipAssetNumber}</div>
            <div class="id-tooltip-item"><span class="tooltip-label">Production Date:</span> ${tooltipProdDate}</div>
            <div class="id-tooltip-item"><span class="tooltip-label">Volume Date:</span> ${tooltipVolDate}</div>
            <div class="id-tooltip-item"><span class="tooltip-label">Last Update:</span> ${tooltipLastUpdate}</div>
          </div>
        </td>
        <td>
          <input type="text" class="spreadsheet-input" value="${escapeHtml(item.pn || '')}" 
            data-field="pn" data-id="${item.id}">
        </td>
        <td>
          <input type="text" class="spreadsheet-input" value="${escapeHtml(item.tool_description || '')}" 
            data-field="tool_description" data-id="${item.id}">
        </td>
        <td>
          <input type="text" class="spreadsheet-input input-center" inputmode="numeric" data-mask="thousands" 
            value="${toolingLifeDisplay}" data-field="tooling_life_qty" data-id="${item.id}">
        </td>
        <td>
          <input type="text" class="spreadsheet-input input-center" inputmode="numeric" data-mask="thousands" 
            value="${producedDisplay}" data-field="produced" data-id="${item.id}">
        </td>
        <td>
          <input type="text" class="spreadsheet-input input-center" inputmode="numeric" data-mask="thousands" 
            value="${forecastDisplay}" data-field="annual_volume_forecast" data-id="${item.id}">
        </td>
        <td class="spreadsheet-expiration">
          ${expirationIconHtml}
          <span class="expiration-text">${expirationDateDisplay}</span>
        </td>
        <td>
          <select class="spreadsheet-select input-center" data-field="steps" data-id="${item.id}">
            ${stepsOptionsHtml}
          </select>
        </td>
        <td>
          <select class="spreadsheet-select ${statusClass}" data-field="status" data-id="${item.id}">
            ${statusOptionsHtml}
          </select>
        </td>
        <td class="spreadsheet-progress">
          <div class="spreadsheet-progress-bar">
            <div class="spreadsheet-progress-fill" style="width: ${percentUsed}%"></div>
          </div>
        </td>
        <td class="spreadsheet-icons">
          <span class="spreadsheet-icon-chain" ${hasChainMembership ? '' : 'hidden'} title="Replacement chain" onclick="event.stopPropagation(); openReplacementTimelineOverlay(${item.id})">
            <i class="ph ph-git-branch"></i>
          </span>
          <span class="spreadsheet-icon-attachment" data-attachment-icon="${item.id}" hidden title="Has attachments" onclick="event.stopPropagation(); openToolingAttachmentsFromSpreadsheet(${item.id})">
            <i class="ph ph-paperclip"></i>
          </span>
        </td>
        <td class="col-expand">
          <button class="spreadsheet-expand-btn" onclick="toggleSpreadsheetRow(${item.id}, ${itemIndex})" title="Expand details">
            <i class="ph ph-caret-down"></i>
          </button>
        </td>
      </tr>
    `;
  }).join('');
  
  spreadsheetBody.innerHTML = rows;
  
  // Carrega contagem de anexos em background
  loadSpreadsheetAttachmentIcons(filteredData);
  
  // Adiciona linha para novo ferramental no final
  const newRow = document.createElement('tr');
  newRow.className = 'spreadsheet-new-row';
  newRow.id = 'spreadsheetNewRow';
  const emptyCheckboxCell = selectionModeActive ? '<td class="col-checkbox"></td>' : '';
  const emptyExpandCell = '<td class="col-expand"></td>';
  newRow.innerHTML = `
    ${emptyCheckboxCell}
    <td class="col-id new-row-icon"><i class="ph ph-plus-circle"></i></td>
    <td>
      <input type="text" class="spreadsheet-input spreadsheet-new-input" id="newToolingPN" placeholder="PN *">
    </td>
    <td>
      <input type="text" class="spreadsheet-input spreadsheet-new-input" id="newToolingDesc" placeholder="Tooling Description">
    </td>
    <td>
      <input type="text" class="spreadsheet-input spreadsheet-new-input input-right" inputmode="numeric" data-mask="thousands" 
        id="newToolingLife" placeholder="Life *">
    </td>
    <td>
      <input type="text" class="spreadsheet-input spreadsheet-new-input input-right" inputmode="numeric" data-mask="thousands" 
        id="newToolingProduced" placeholder="Produced *">
    </td>
    <td>
      <input type="text" class="spreadsheet-input spreadsheet-new-input input-right" inputmode="numeric" data-mask="thousands" 
        id="newToolingVolume" placeholder="Volume">
    </td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td class="col-expand">
      <button class="spreadsheet-expand-btn spreadsheet-add-btn" onclick="spreadsheetCreateTooling()" title="Create Tooling">
        <i class="ph ph-plus"></i>
      </button>
    </td>
  `;
  spreadsheetBody.appendChild(newRow);
  
  // Aplica máscara de milhares nos inputs numéricos
  const spreadsheetContainer = document.getElementById('spreadsheetContainer');
  if (spreadsheetContainer) {
    applyInitialThousandsMask(spreadsheetContainer);
  }
  
  // Adiciona event listeners para salvar alterações dos campos da spreadsheet
  attachSpreadsheetFieldListeners();
  
  // Atualiza indicadores visuais de filtros e ordenação
  updateFilterButtonIndicators();
  updateSortButtonIndicators();
  
  // Adiciona listeners para os tooltips do ID
  initIdTooltips();
}

// Adiciona event listeners aos campos editáveis da spreadsheet
function attachSpreadsheetFieldListeners() {
  const spreadsheetBody = document.getElementById('spreadsheetBody');
  if (!spreadsheetBody) return;
  
  // Inputs de texto e numéricos
  const inputs = spreadsheetBody.querySelectorAll('input.spreadsheet-input[data-field][data-id]:not(.spreadsheet-new-input)');
  inputs.forEach(input => {
    // Salva ao sair do campo (blur)
    input.addEventListener('blur', async () => {
      await spreadsheetSave(input);
    });
    
    // Salva ao pressionar Enter
    input.addEventListener('keydown', async (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        await spreadsheetSave(input);
        input.blur();
      }
    });
  });
  
  // Selects (salvam imediatamente ao mudar)
  const selects = spreadsheetBody.querySelectorAll('select.spreadsheet-select[data-field][data-id]');
  selects.forEach(select => {
    select.addEventListener('change', async () => {
      await spreadsheetSave(select);
    });
  });
}

// Inicializa tooltips do ID
function initIdTooltips() {
  const idCells = document.querySelectorAll('.id-with-tooltip');
  
  idCells.forEach(cell => {
    cell.addEventListener('mouseenter', (event) => {
      const tooltip = cell.querySelector('.id-tooltip');
      if (!tooltip) return;
      
      // Posiciona o tooltip baseado na posição do mouse
      const rect = cell.getBoundingClientRect();
      tooltip.style.left = `${rect.right + 10}px`;
      tooltip.style.top = `${rect.top}px`;
    });
  });
}

// Estado do popup de filtro
let currentFilterColumn = null;
let currentFilterPopupData = [];

// Aplica filtros de coluna aos dados (suporta múltiplos valores por coluna)
function applyColumnFiltersToData(data) {
  if (!data || Object.keys(columnFilters).length === 0) return data;
  
  return data.filter(item => {
    for (const [column, selectedValues] of Object.entries(columnFilters)) {
      if (!selectedValues || !Array.isArray(selectedValues) || selectedValues.length === 0) continue;
      
      let itemValue = String(item[column] || '').trim();
      
      // Verifica se o valor do item está entre os selecionados
      if (!selectedValues.includes(itemValue)) {
        return false;
      }
    }
    return true;
  });
}

// Atualiza indicadores visuais nos botões de filtro e botão floating de limpar
function updateFilterButtonIndicators() {
  let activeFilterCount = 0;
  
  document.querySelectorAll('.column-filter-btn').forEach(btn => {
    const th = btn.closest('th');
    const column = th?.dataset?.filterable;
    if (column && columnFilters[column] && columnFilters[column].length > 0) {
      btn.classList.add('active');
      activeFilterCount++;
    } else {
      btn.classList.remove('active');
    }
  });
  
  // Atualiza o botão floating de limpar filtros
  const clearFilterBtn = document.getElementById('floatingClearFilterBtn');
  if (clearFilterBtn) {
    clearFilterBtn.style.display = activeFilterCount > 0 ? 'flex' : 'none';
  }
}

// Abre o popup de filtro para uma coluna
function openColumnFilter(event, column) {
  event.stopPropagation();
  
  const popup = document.getElementById('columnFilterPopup');
  const list = document.getElementById('columnFilterList');
  const title = document.getElementById('columnFilterPopupTitle');
  const searchInput = document.getElementById('columnFilterSearch');
  
  if (!popup || !list || !toolingData) return;
  
  currentFilterColumn = column;
  
  // Define título baseado na coluna
  const columnTitles = {
    id: 'Filter by ID',
    pn: 'Filter by PN',
    pn_description: 'Filter by PN Description',
    tool_description: 'Filter by Tooling Description',
    tooling_life_qty: 'Filter by Tooling Life',
    produced: 'Filter by Produced',
    annual_volume_forecast: 'Filter by Annual Volume',
    steps: 'Filter by Steps',
    status: 'Filter by Status'
  };
  title.textContent = columnTitles[column] || 'Filter';
  
  // Aplica todos os outros filtros (exceto o da coluna atual) para obter dados segmentados
  let segmentedData = toolingData;
  
  // Aplica filtro de expiração se estiver ativo
  if (expirationFilterEnabled) {
    segmentedData = segmentedData.filter(item => {
      const classification = classifyToolingExpirationState(item);
      return classification.state === 'expired' || classification.state === 'warning';
    });
  }
  
  // Aplica filtros das outras colunas (não da coluna atual)
  for (const [filterColumn, selectedValues] of Object.entries(columnFilters)) {
    if (filterColumn === column) continue; // Pula a coluna atual
    if (!selectedValues || !Array.isArray(selectedValues) || selectedValues.length === 0) continue;
    
    segmentedData = segmentedData.filter(item => {
      const itemValue = String(item[filterColumn] || '').trim();
      return selectedValues.includes(itemValue);
    });
  }
  
  // Coleta valores únicos da coluna a partir dos dados segmentados
  const uniqueValues = new Set();
  segmentedData.forEach(item => {
    const value = String(item[column] || '').trim();
    if (value) uniqueValues.add(value);
  });
  
  // Ordena valores
  const sortedValues = Array.from(uniqueValues).sort((a, b) => {
    const numA = parseFloat(a.replace(/[^\d.-]/g, ''));
    const numB = parseFloat(b.replace(/[^\d.-]/g, ''));
    if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
    return a.localeCompare(b);
  });
  
  currentFilterPopupData = sortedValues;
  
  // Valores atualmente selecionados para esta coluna
  const selectedValues = columnFilters[column] || [];
  
  // Gera HTML das opções
  let optionsHtml = '';
  sortedValues.forEach((value, index) => {
    const isChecked = selectedValues.includes(value) ? 'checked' : '';
    optionsHtml += `
      <div class="column-filter-option" data-value="${escapeHtml(value)}">
        <input type="checkbox" id="filterOption_${index}" value="${escapeHtml(value)}" ${isChecked}>
        <label for="filterOption_${index}">${escapeHtml(value)}</label>
      </div>
    `;
  });
  
  list.innerHTML = optionsHtml || '<div style="padding: 12px; color: #999; text-align: center;">No values found</div>';
  
  // Limpa busca
  searchInput.value = '';
  
  // Atualiza estado dos botões de ordenação no popup
  updatePopupSortButtons(column);
  
  // Posiciona o popup próximo ao botão
  const btnRect = event.target.closest('.column-filter-btn').getBoundingClientRect();
  popup.style.top = `${btnRect.bottom + 4}px`;
  popup.style.left = `${Math.min(btnRect.left, window.innerWidth - 240)}px`;
  
  // Abre o popup
  popup.classList.add('open');
  
  // Adiciona listener para fechar ao clicar fora
  setTimeout(() => {
    document.addEventListener('click', handleClickOutsideFilterPopup);
  }, 10);
}

// Fecha o popup de filtro
function closeColumnFilter() {
  const popup = document.getElementById('columnFilterPopup');
  if (popup) {
    popup.classList.remove('open');
  }
  currentFilterColumn = null;
  document.removeEventListener('click', handleClickOutsideFilterPopup);
}

// Handler para fechar ao clicar fora
function handleClickOutsideFilterPopup(event) {
  const popup = document.getElementById('columnFilterPopup');
  if (popup && !popup.contains(event.target) && !event.target.closest('.column-filter-btn')) {
    closeColumnFilter();
  }
}

// Atualiza os botões de ordenação no popup
function updatePopupSortButtons(column) {
  const sortBtns = document.querySelectorAll('.column-filter-sort-btn');
  sortBtns.forEach(btn => {
    btn.classList.remove('active');
    const sortDir = btn.dataset.sort;
    if (columnSort.column === column && columnSort.direction === sortDir) {
      btn.classList.add('active');
    }
  });
}

// Define a ordenação a partir do popup
function setColumnSortFromPopup(direction) {
  if (!currentFilterColumn) return;
  
  // Se já está nessa direção, remove a ordenação
  if (columnSort.column === currentFilterColumn && columnSort.direction === direction) {
    columnSort = { column: null, direction: null };
  } else {
    columnSort = { column: currentFilterColumn, direction };
  }
  
  // Atualiza os botões do popup
  updatePopupSortButtons(currentFilterColumn);
  
  // Atualiza indicadores nos headers
  updateSortButtonIndicators();
}

// Filtra as opções pelo texto de busca
function filterColumnOptions() {
  const searchInput = document.getElementById('columnFilterSearch');
  const searchText = (searchInput?.value || '').toLowerCase();
  
  document.querySelectorAll('.column-filter-option').forEach(option => {
    const value = option.dataset.value?.toLowerCase() || '';
    if (value.includes(searchText)) {
      option.classList.remove('hidden');
    } else {
      option.classList.add('hidden');
    }
  });
}

// Seleciona todas as opções visíveis
function selectAllFilterOptions() {
  document.querySelectorAll('.column-filter-option:not(.hidden) input[type="checkbox"]').forEach(cb => {
    cb.checked = true;
  });
}

// Limpa todas as opções
function clearAllFilterOptions() {
  document.querySelectorAll('.column-filter-option input[type="checkbox"]').forEach(cb => {
    cb.checked = false;
  });
}

// Aplica o filtro da coluna atual
function applyColumnFilter() {
  if (!currentFilterColumn) return;
  
  // Coleta valores selecionados
  const selectedValues = [];
  document.querySelectorAll('.column-filter-option input[type="checkbox"]:checked').forEach(cb => {
    selectedValues.push(cb.value);
  });
  
  // Atualiza o filtro
  if (selectedValues.length > 0) {
    columnFilters[currentFilterColumn] = selectedValues;
  } else {
    delete columnFilters[currentFilterColumn];
  }
  
  // Fecha o popup
  closeColumnFilter();
  
  // Re-renderiza a planilha
  renderSpreadsheetView();
}

// Limpa todos os filtros de coluna
function clearColumnFilters() {
  columnFilters = {};
  columnSort = { column: null, direction: null };
  renderSpreadsheetView();
}

// Alterna a ordenação de uma coluna
function toggleColumnSort(event, column) {
  event.stopPropagation();
  
  if (columnSort.column === column) {
    // Se já está ordenando por esta coluna, alterna a direção ou remove
    if (columnSort.direction === 'asc') {
      columnSort.direction = 'desc';
    } else {
      // Remove ordenação
      columnSort = { column: null, direction: null };
    }
  } else {
    // Nova coluna, começa com ascendente
    columnSort = { column, direction: 'asc' };
  }
  
  renderSpreadsheetView();
}

// Aplica ordenação aos dados
function applySortToData(data) {
  if (!data || !columnSort.column) return data;
  
  const sortedData = [...data];
  const column = columnSort.column;
  
  sortedData.sort((a, b) => {
    let valueA, valueB;
    
    if (column === 'expiration') {
      // Ordena por data de expiração
      const classA = classifyToolingExpirationState(a);
      const classB = classifyToolingExpirationState(b);
      const dateA = classA.expirationDate ? new Date(classA.expirationDate) : null;
      const dateB = classB.expirationDate ? new Date(classB.expirationDate) : null;
      
      // Itens sem data vão para o final
      if (!dateA && !dateB) return 0;
      if (!dateA) return 1;
      if (!dateB) return -1;
      
      valueA = dateA.getTime();
      valueB = dateB.getTime();
    } else if (column === 'progress') {
      // Ordena por porcentagem de uso
      const lifeA = Number(parseLocalizedNumber(a.tooling_life_qty)) || Number(a.tooling_life_qty) || 0;
      const prodA = Number(parseLocalizedNumber(a.produced)) || Number(a.produced) || 0;
      const lifeB = Number(parseLocalizedNumber(b.tooling_life_qty)) || Number(b.tooling_life_qty) || 0;
      const prodB = Number(parseLocalizedNumber(b.produced)) || Number(b.produced) || 0;
      
      valueA = lifeA > 0 ? (prodA / lifeA) * 100 : 0;
      valueB = lifeB > 0 ? (prodB / lifeB) * 100 : 0;
    } else if (column === 'id' || column === 'tooling_life_qty' || column === 'produced' || column === 'annual_volume_forecast') {
      // Colunas numéricas
      valueA = Number(parseLocalizedNumber(a[column])) || Number(a[column]) || 0;
      valueB = Number(parseLocalizedNumber(b[column])) || Number(b[column]) || 0;
    } else if (column === 'steps') {
      // Steps são números de 1-7
      valueA = Number(a[column]) || 0;
      valueB = Number(b[column]) || 0;
    } else {
      // Colunas de texto (pn, pn_description, tool_description, status)
      valueA = String(a[column] || '').toLowerCase();
      valueB = String(b[column] || '').toLowerCase();
      
      if (columnSort.direction === 'asc') {
        return valueA.localeCompare(valueB);
      } else {
        return valueB.localeCompare(valueA);
      }
    }
    
    if (columnSort.direction === 'asc') {
      return valueA - valueB;
    } else {
      return valueB - valueA;
    }
  });
  
  return sortedData;
}

// Atualiza indicadores visuais nos botões de ordenação
function updateSortButtonIndicators() {
  document.querySelectorAll('.column-sort-btn').forEach(btn => {
    const th = btn.closest('th');
    const column = th?.dataset?.sortable;
    const icon = btn.querySelector('i');
    
    btn.classList.remove('sort-asc', 'sort-desc');
    
    if (column && columnSort.column === column) {
      if (columnSort.direction === 'asc') {
        btn.classList.add('sort-asc');
        icon.className = 'ph ph-sort-ascending';
      } else {
        btn.classList.add('sort-desc');
        icon.className = 'ph ph-sort-descending';
      }
    } else {
      icon.className = 'ph ph-arrows-down-up';
    }
  });
}

// Carrega ícones de anexos para a planilha
async function loadSpreadsheetAttachmentIcons(data) {
  const supplierContext = selectedSupplier || currentSupplier || '';
  if (!supplierContext) return;
  
  for (const item of data) {
    try {
      const attachments = await window.api.getAttachments(supplierContext, item.id);
      if (attachments && attachments.length > 0) {
        const icon = document.querySelector(`[data-attachment-icon="${item.id}"]`);
        if (icon) {
          icon.removeAttribute('hidden');
        }
      }
    } catch (error) {
      // Silently fail
    }
  }
}

// Toggle da expansão de uma linha da planilha para mostrar detalhes do card
function toggleSpreadsheetRow(itemId, itemIndex) {
  const row = document.querySelector(`tr[data-id="${itemId}"]`);
  if (!row) return;
  
  const expandBtn = row.querySelector('.spreadsheet-expand-btn');
  const isExpanded = row.classList.contains('row-expanded');
  
  // Fecha todas as outras linhas expandidas primeiro
  const allExpandedRows = document.querySelectorAll('.spreadsheet-table tr.row-expanded');
  allExpandedRows.forEach(expandedRow => {
    if (expandedRow !== row) {
      expandedRow.classList.remove('row-expanded');
      const btn = expandedRow.querySelector('.spreadsheet-expand-btn');
      if (btn) btn.innerHTML = '<i class="ph ph-caret-down"></i>';
      
      // Remove a linha de detalhes
      const detailRow = expandedRow.nextElementSibling;
      if (detailRow && detailRow.classList.contains('spreadsheet-detail-row')) {
        // Salva antes de fechar e sincroniza a linha
        const prevItemId = expandedRow.getAttribute('data-id');
        if (prevItemId) {
          Promise.resolve().then(() => {
            saveToolingQuietly(prevItemId);
            syncSpreadsheetRowFromCard(prevItemId);
          });
        }
        detailRow.remove();
      }
    }
  });
  
  if (isExpanded) {
    // Fecha a linha atual
    row.classList.remove('row-expanded');
    if (expandBtn) expandBtn.innerHTML = '<i class="ph ph-caret-down"></i>';
    
    // Remove a linha de detalhes
    const detailRow = row.nextElementSibling;
    if (detailRow && detailRow.classList.contains('spreadsheet-detail-row')) {
      // Salva antes de fechar e sincroniza a linha
      Promise.resolve().then(() => {
        saveToolingQuietly(itemId);
        syncSpreadsheetRowFromCard(itemId);
      });
      detailRow.remove();
    }
  } else {
    // Abre a linha atual
    row.classList.add('row-expanded');
    if (expandBtn) expandBtn.innerHTML = '<i class="ph ph-caret-up"></i>';
    
    // Cria a linha de detalhes com o conteúdo do card
    const item = toolingData.find(t => String(t.id) === String(itemId));
    if (item) {
      const supplierContext = selectedSupplier || currentSupplier || '';
      const chainMembership = new Map();
      const bodyHTML = buildToolingCardBodyHTML(item, itemIndex, chainMembership, supplierContext);
      
      // Conta o número de colunas
      const colCount = row.querySelectorAll('td').length;
      
      // Cria a linha de detalhes
      const detailRow = document.createElement('tr');
      detailRow.className = 'spreadsheet-detail-row';
      detailRow.setAttribute('data-detail-for', itemId);
      detailRow.innerHTML = `
        <td colspan="${colCount}" class="spreadsheet-detail-cell">
          <div class="spreadsheet-card-container" data-item-id="${itemId}" data-item-index="${itemIndex}">
            ${bodyHTML}
          </div>
        </td>
      `;
      
      // Insere após a linha atual
      row.after(detailRow);
      
      // Inicializa o conteúdo do card
      const cardContainer = detailRow.querySelector('.spreadsheet-card-container');
      if (cardContainer) {
        populateCardDataListForSpreadsheet(itemIndex, cardContainer);
        applyInitialThousandsMask(cardContainer);
        
        // Restaura reminders de data persistentes
        restoreDateReminders(itemId);
        
        // Initialize drag and drop for attachment dropzone
        const dropzone = cardContainer.querySelector('.card-attachments-dropzone');
        if (dropzone) {
          initCardAttachmentDragAndDrop(dropzone, itemId);
        }
        
        // Initialize step description and responsible
        const stepSelect = cardContainer.querySelector(`select[data-field="steps"][data-id="${itemId}"]`);
        if (stepSelect) {
          const stepValue = stepSelect.value;
          const descriptionLabel = document.getElementById(`stepDescription_${itemId}`);
          if (descriptionLabel) {
            const description = getStepDescription(stepValue);
            const responsible = getStepResponsible(stepValue);
            
            if (description) {
              descriptionLabel.innerHTML = `<div style="color: #8b92a7; line-height: 1.4;">${description}<br><span style="font-size: 0.9em;">Responsible: ${responsible}</span></div>`;
              descriptionLabel.style.display = 'block';
              descriptionLabel.style.textAlign = 'left';
              descriptionLabel.style.marginTop = '4px';
            } else {
              descriptionLabel.textContent = '';
              descriptionLabel.style.display = 'none';
            }
          }
        }
        
        // Initialize carousel buttons state
        const track = cardContainer.querySelector('[data-carousel-track]');
        const prevBtn = cardContainer.querySelector('.carousel-nav-prev');
        const nextBtn = cardContainer.querySelector('.carousel-nav-next');
        if (track && prevBtn && nextBtn) {
          const columns = Array.from(track.children);
          prevBtn.disabled = true;
          nextBtn.disabled = columns.length <= 1;
          track.style.transform = 'translateX(0)';
          track.dataset.carouselIndex = '0';
        }
        
        // Carrega dados do card em background
        setTimeout(() => {
          calculateExpirationDateForSpreadsheet(itemIndex, cardContainer, true);
          loadCardAttachmentsForSpreadsheet(itemId, cardContainer).catch(err => {});
          
          // Cria snapshot inicial para detectar alterações
          const snapshotKey = getSnapshotKey(itemId);
          const values = collectCardDomValues(itemId);
          if (values) {
            // Adiciona comentários do item ao snapshot
            if (item.comments) {
              values.comments = item.comments;
            }
            cardSnapshotStore.set(snapshotKey, serializeCardValues(values));
          }
          
          // Scroll para mostrar a linha no topo com o card expandido visível
          const detailRow = row.nextElementSibling;
          if (detailRow && detailRow.classList.contains('spreadsheet-detail-row')) {
            // Calcula a posição ideal considerando o header fixo
            const headerHeight = document.querySelector('.spreadsheet-table thead')?.offsetHeight || 50;
            const container = document.querySelector('.spreadsheet-container');
            if (container) {
              const rowTop = row.offsetTop - headerHeight - 10;
              container.scrollTo({ top: rowTop, behavior: 'smooth' });
            }
          }
        }, 50);
      }
    }
  }
}

// Função auxiliar para popular datalist no spreadsheet
function populateCardDataListForSpreadsheet(index, container) {
  const suppliers = getSupplierOptions();
  const responsibles = getResponsibleOptions(toolingData);
  
  const supplierList = container.querySelector(`#supplierList-${index}`);
  if (supplierList) {
    supplierList.innerHTML = suppliers
      .map(name => `<option value="${escapeHtml(name)}"></option>`)
      .join('');
  }

  const ownerList = container.querySelector(`#ownerList-${index}`);
  if (ownerList) {
    ownerList.innerHTML = responsibles
      .map(name => `<option value="${escapeHtml(name)}"></option>`)
      .join('');
  }
}

// Função auxiliar para calcular expiração no spreadsheet
function calculateExpirationDateForSpreadsheet(index, container, skipSave) {
  // Reutiliza a lógica existente adaptada para container
  const item = toolingData[index];
  if (!item) return;
  
  const classification = classifyToolingExpirationState(item);
  const expirationDate = classification.expirationDate || '';
  
  // Atualiza o input de expiração se existir
  const expirationInput = container.querySelector(`input[data-field="expiration_date"][data-id="${item.id}"]`);
  if (expirationInput) {
    expirationInput.value = formatDateForDateInput(expirationDate);
  }
}

// Função auxiliar para carregar anexos no spreadsheet
async function loadCardAttachmentsForSpreadsheet(itemId, container) {
  try {
    const supplierContext = selectedSupplier || currentSupplier || '';
    const attachments = await window.api.getAttachments(supplierContext, itemId);
    const list = container.querySelector('.card-attachments-list');
    const empty = container.querySelector('.card-attachments-empty');
    
    if (!list) return;
    
    if (!attachments || attachments.length === 0) {
      if (empty) empty.style.display = 'flex';
      list.innerHTML = '';
      return;
    }
    
    if (empty) empty.style.display = 'none';
    
    list.innerHTML = attachments.map(att => `
      <div class="card-attachment-item">
        <div class="card-attachment-info">
          <i class="ph ${getFileIcon(att.name)}"></i>
          <div class="card-attachment-details">
            <span class="card-attachment-name" title="${escapeHtml(att.name)}">${escapeHtml(att.name)}</span>
            <span class="card-attachment-meta">${formatFileSize(att.size)} • ${formatDate(att.date)}</span>
          </div>
        </div>
        <div class="card-attachment-actions">
          <button class="card-attachment-btn" onclick="openToolingAttachment('${escapeHtml(att.path)}')" title="Open">
            <i class="ph ph-folder-open"></i>
          </button>
          <button class="card-attachment-btn danger" onclick="deleteToolingAttachment(${itemId}, '${escapeHtml(att.name)}')" title="Delete">
            <i class="ph ph-trash"></i>
          </button>
        </div>
      </div>
    `).join('');
  } catch (error) {
    console.error('Error loading attachments for spreadsheet:', error);
  }
}

// Sincroniza os dados entre o card expandido e a linha da planilha
function syncSpreadsheetRowFromCard(itemId) {
  const row = document.querySelector(`tr[data-id="${itemId}"]`);
  if (!row) return;
  
  const item = toolingData.find(t => String(t.id) === String(itemId));
  if (!item) return;
  
  // Atualiza os inputs na linha
  const fields = ['pn', 'pn_description', 'tool_description', 'tooling_life_qty', 'produced', 'date_remaining_tooling_life', 'annual_volume_forecast', 'date_annual_volume', 'status', 'steps'];
  fields.forEach(field => {
    const input = row.querySelector(`[data-field="${field}"][data-id="${itemId}"]`);
    if (input) {
      const value = item[field];
      if (input.tagName === 'SELECT') {
        input.value = value || '';
        if (field === 'status') {
          input.className = 'spreadsheet-select ' + getStatusClass(value);
        }
      } else if (input.type === 'date') {
        input.value = formatDateForDateInput(value);
      } else if (['tooling_life_qty', 'produced', 'annual_volume_forecast'].includes(field)) {
        input.value = formatNumericForSpreadsheet(value);
      } else {
        input.value = value || '';
      }
    }
  });
  
  // Atualiza a expiração
  const classification = classifyToolingExpirationState(item);
  const expirationCell = row.querySelector('.spreadsheet-expiration');
  if (expirationCell) {
    const hasExpirationDate = classification.expirationDate && classification.expirationDate !== '';
    const expirationDateDisplay = hasExpirationDate ? formatDate(classification.expirationDate) : '';
    const isAnalysisCompleted = item.analysis_completed === 1;
    const expirationIconHtml = hasExpirationDate ? getSpreadsheetExpirationIcon(classification.state, isAnalysisCompleted) : '';
    expirationCell.innerHTML = `
      ${expirationIconHtml}
      <span class="expiration-text">${expirationDateDisplay}</span>
    `;
  }
}

// Sincroniza apenas a célula de expiração na linha da spreadsheet
function syncSpreadsheetExpirationCell(itemId) {
  const row = document.querySelector(`tr[data-id="${itemId}"]`);
  if (!row) return;
  
  const item = toolingData.find(t => String(t.id) === String(itemId));
  if (!item) return;
  
  // Usa diretamente o expiration_date do item (já atualizado)
  const expirationDateValue = item.expiration_date || '';
  const hasExpirationDate = expirationDateValue && expirationDateValue !== '';
  const expirationDateDisplay = hasExpirationDate ? formatDate(expirationDateValue) : '';
  
  // Calcula o estado para ícone
  const expirationStatus = getExpirationStatus(expirationDateValue);
  const toolingLife = Number(parseLocalizedNumber(item.tooling_life_qty)) || 0;
  const produced = Number(parseLocalizedNumber(item.produced)) || 0;
  const percentUsed = toolingLife > 0 ? (produced / toolingLife) * 100 : 0;
  const normalizedStatus = (item.status || '').toString().trim().toLowerCase();
  const isObsolete = normalizedStatus === 'obsolete';
  
  let state = 'ok';
  if (isObsolete) {
    state = 'obsolete';
  } else if (percentUsed >= 100 || expirationStatus.class === 'expired') {
    state = 'expired';
  } else if (expirationStatus.class === 'warning') {
    state = 'warning';
  }
  
  const isAnalysisCompleted = item.analysis_completed === 1;
  const expirationIconHtml = hasExpirationDate ? getSpreadsheetExpirationIcon(state, isAnalysisCompleted) : '';
  
  const expirationCell = row.querySelector('.spreadsheet-expiration');
  if (expirationCell) {
    expirationCell.innerHTML = `
      ${expirationIconHtml}
      <span class="expiration-text">${expirationDateDisplay}</span>
    `;
  }
}

// Recalcula expiration_date a partir da spreadsheet e atualiza tudo
async function recalculateExpirationForSpreadsheet(itemId) {
  const item = toolingData.find(t => String(t.id) === String(itemId));
  if (!item) return;
  
  // Calcula remaining
  const toolingLife = parseLocalizedNumber(item.tooling_life_qty) || 0;
  const produced = parseLocalizedNumber(item.produced) || 0;
  const remaining = toolingLife - produced;
  const forecast = parseLocalizedNumber(item.annual_volume_forecast) || 0;
  const productionDate = item.date_remaining_tooling_life || '';
  
  // Calcula a nova expiration_date
  const formattedDate = calculateExpirationFromFormula({
    remaining,
    forecast,
    productionDate
  });
  
  // Atualiza o toolingData
  item.expiration_date = formattedDate || '';
  
  // Atualiza a célula de expiração na spreadsheet
  syncSpreadsheetExpirationCell(itemId);
  
  // Atualiza os ícones de expiração
  updateExpirationIconsForItem(itemId);
  
  // Se o card expandido estiver aberto, atualiza o input de expiration_date
  const detailRow = document.querySelector(`.spreadsheet-detail-row[data-detail-for="${itemId}"]`);
  if (detailRow) {
    const expirationInput = detailRow.querySelector(`[data-field="expiration_date"][data-id="${itemId}"]`);
    if (expirationInput) {
      expirationInput.value = formattedDate || '';
    }
    // Atualiza também o header do card dentro do detalhe
    const cardHeader = detailRow.querySelector('[data-card-expiration]');
    if (cardHeader) {
      cardHeader.textContent = formattedDate ? formatDate(formattedDate) : '';
    }
  }
  
  // Salva no banco de dados (mesmo se vazio, para limpar valor anterior)
  try {
    await window.api.updateTooling(itemId, { expiration_date: formattedDate || '' });
  } catch (error) {
    console.error('Error saving expiration_date:', error);
  }
}

// Cria novo ferramental a partir da planilha
async function spreadsheetCreateTooling() {
  const pnInput = document.getElementById('newToolingPN');
  const descInput = document.getElementById('newToolingDesc');
  const lifeInput = document.getElementById('newToolingLife');
  const producedInput = document.getElementById('newToolingProduced');
  const prodDateInput = document.getElementById('newToolingProdDate');
  const volumeInput = document.getElementById('newToolingVolume');
  const volDateInput = document.getElementById('newToolingVolDate');
  
  // Validar campos obrigatórios
  const pn = pnInput ? pnInput.value.trim() : '';
  const desc = descInput ? descInput.value.trim() : '';;
  const toolingLife = lifeInput ? parseLocalizedNumber(lifeInput.value) : 0;
  const produced = producedInput ? parseLocalizedNumber(producedInput.value) : 0;
  const prodDate = prodDateInput ? prodDateInput.value : null;
  const volume = volumeInput ? parseLocalizedNumber(volumeInput.value) : 0;
  const volDate = volDateInput ? volDateInput.value : null;
  
  // Verificar campos obrigatórios
  if (!pn) {
    showNotification('Informe o PN do ferramental.', 'error');
    if (pnInput) pnInput.focus();
    return;
  }
  
  if (!currentSupplier) {
    showNotification('Selecione um fornecedor primeiro.', 'error');
    return;
  }
  
  if (toolingLife <= 0) {
    showNotification('Informe a vida útil do ferramental.', 'error');
    if (lifeInput) lifeInput.focus();
    return;
  }
  
  if (produced < 0) {
    showNotification('Valor de produção inválido.', 'error');
    if (producedInput) producedInput.focus();
    return;
  }
  
  try {
    // Criar comentário inicial
    const now = new Date();
    const dateStr = now.toLocaleString('pt-BR', { 
      day: '2-digit', 
      month: '2-digit', 
      year: 'numeric', 
      hour: '2-digit', 
      minute: '2-digit' 
    });
    
    const initialComment = {
      date: dateStr,
      text: `Created with Tooling Life: ${formatIntegerWithSeparators(toolingLife)} pcs`,
      initial: true
    };
    
    const commentsJson = JSON.stringify([initialComment]);
    
    const payload = {
      pn,
      pn_description: null,
      supplier: currentSupplier,
      cummins_responsible: null,
      tool_description: desc,
      tooling_life_qty: toolingLife,
      produced,
      date_remaining_tooling_life: prodDate,
      annual_volume_forecast: volume > 0 ? volume : null,
      date_annual_volume: volDate,
      status: 'ACTIVE',
      comments: commentsJson
    };
    
    const result = await window.api.createTooling(payload);
    
    if (!result || result.success !== true) {
      showNotification(result?.error || 'Não foi possível criar o ferramental.', 'error');
      return;
    }
    
    // Limpar campos
    if (pnInput) pnInput.value = '';
    if (descInput) descInput.value = '';
    if (lifeInput) lifeInput.value = '';
    if (producedInput) producedInput.value = '';
    if (prodDateInput) prodDateInput.value = '';
    if (volumeInput) volumeInput.value = '';
    if (volDateInput) volDateInput.value = '';
    
    // Recarregar dados
    await loadSuppliers();
    await loadAnalytics();
    await refreshReplacementIdOptions(true);
    await loadToolingBySupplier(currentSupplier);
    
    showNotification('Ferramental criado com sucesso!');
    
  } catch (error) {
    console.error('Error creating tooling from spreadsheet:', error);
    showNotification('Erro ao criar ferramental', 'error');
  }
}

// Retorna o ícone HTML para o estado de expiração na planilha
function getSpreadsheetExpirationIcon(state, isAnalysisCompleted = false) {
  // Se a análise foi concluída, mostra ícone diferenciado (clipboard com check)
  if (isAnalysisCompleted && (state === 'expired' || state === 'warning')) {
    return '<i class="ph ph-fill ph-clipboard-text expiration-icon analysis-completed" title="Analysis Completed"></i>';
  }
  
  switch (state) {
    case 'expired':
      return '<i class="ph ph-fill ph-warning-circle expiration-icon expired" title="Expired"></i>';
    case 'warning':
      return '<i class="ph ph-fill ph-warning expiration-icon warning" title="Expiring Soon"></i>';
    case 'obsolete':
    case 'obsolete-replaced':
      return '<i class="ph ph-fill ph-archive-box expiration-icon obsolete" title="Obsolete"></i>';
    default:
      return '<i class="ph ph-check-circle expiration-icon ok" title="OK"></i>';
  }
}

// Ícone de expiração para dentro do input do card (com classe input-icon)
function getCardExpirationIcon(state, isAnalysisCompleted = false) {
  // Se a análise foi concluída, mostra ícone diferenciado (clipboard com check)
  if (isAnalysisCompleted && (state === 'expired' || state === 'warning')) {
    return '<i class="ph ph-fill ph-clipboard-text expiration-icon analysis-completed input-icon" title="Analysis Completed"></i>';
  }
  
  switch (state) {
    case 'expired':
      return '<i class="ph ph-fill ph-warning-circle expiration-icon expired input-icon" title="Expired"></i>';
    case 'warning':
      return '<i class="ph ph-fill ph-warning expiration-icon warning input-icon" title="Expiring Soon"></i>';
    case 'obsolete':
    case 'obsolete-replaced':
      return '<i class="ph ph-fill ph-archive-box expiration-icon obsolete input-icon" title="Obsolete"></i>';
    default:
      return '<i class="ph ph-check-circle expiration-icon ok input-icon" title="OK"></i>';
  }
}

// Constrói opções de status para o select da planilha
function buildSpreadsheetStatusOptions(currentStatus) {
  const normalizedCurrent = (currentStatus || '').toString().trim().toUpperCase();
  let options = '<option value=""></option>';
  
  statusOptions.forEach(status => {
    const selected = status.toUpperCase() === normalizedCurrent ? 'selected' : '';
    options += `<option value="${escapeHtml(status)}" ${selected}>${escapeHtml(status)}</option>`;
  });
  
  return options;
}

// Constrói opções de steps para o select da planilha
function buildSpreadsheetStepsOptions(currentStep) {
  const normalizedCurrent = (currentStep || '').toString().trim();
  const steps = ['', '1', '2', '3', '4', '5', '6', '7'];
  
  return steps.map(step => {
    const selected = step === normalizedCurrent ? 'selected' : '';
    return `<option value="${step}" ${selected}>${step}</option>`;
  }).join('');
}

// Retorna a classe CSS para o status
function getStatusClass(status) {
  const normalized = (status || '').toString().trim().toLowerCase();
  if (normalized === 'active') return 'status-active';
  if (normalized === 'obsolete') return 'status-obsolete';
  if (normalized === 'inactive') return 'status-inactive';
  if (normalized === 'under construction') return 'status-construction';
  return '';
}

// Formata data para exibição no input (formato BR)
function formatDateForInput(dateValue) {
  if (!dateValue) return '';
  // Se já estiver em formato BR (dd/mm/yyyy), retorna como está
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateValue)) return dateValue;
  // Se estiver em formato ISO (yyyy-mm-dd), converte
  if (/^\d{4}-\d{2}-\d{2}/.test(dateValue)) {
    const parts = dateValue.split('-');
    return `${parts[2]}/${parts[1]}/${parts[0]}`;
  }
  return dateValue;
}

// Formata data para input type="date" (formato ISO yyyy-mm-dd)
function formatDateForDateInput(dateValue) {
  if (!dateValue) return '';
  // Se estiver em formato ISO (yyyy-mm-dd), retorna como está
  if (/^\d{4}-\d{2}-\d{2}/.test(dateValue)) {
    return dateValue.substring(0, 10); // Pega apenas yyyy-mm-dd
  }
  // Se estiver em formato BR (dd/mm/yyyy), converte para ISO
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateValue)) {
    const parts = dateValue.split('/');
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }
  return '';
}

// Formata valor numérico para exibição na planilha com separador de milhares
function formatNumericForSpreadsheet(value) {
  if (value === null || value === undefined || value === '') return '';
  const num = parseLocalizedNumber(value);
  if (num === 0 && String(value).trim() !== '0' && String(value).trim() !== '') {
    // Se o valor original não era zero mas parseou como zero, retorna vazio
    return '';
  }
  if (num === 0) return '';
  return formatIntegerWithSeparators(num);
}

// Campos que são numéricos
const numericSpreadsheetFields = ['tooling_life_qty', 'produced', 'annual_volume_forecast'];

// Campos que são datas
const dateSpreadsheetFields = ['date_remaining_tooling_life', 'date_annual_volume'];

// Sincroniza um campo específico da spreadsheet para o card expandido
function syncExpandedCardFromSpreadsheet(itemId, field, value) {
  // Procura o card expandido (dentro de .spreadsheet-detail-row)
  const detailRow = document.querySelector(`.spreadsheet-detail-row[data-detail-for="${itemId}"]`);
  if (!detailRow) return;
  
  // Encontra o input/select correspondente dentro do card expandido
  const targetElement = detailRow.querySelector(`[data-field="${field}"][data-id="${itemId}"]`);
  if (!targetElement) return;
  
  // Atualiza o valor
  if (targetElement.tagName === 'SELECT') {
    targetElement.value = value || '';
  } else if (targetElement.type === 'date') {
    targetElement.value = formatDateForDateInput(value);
  } else if (numericSpreadsheetFields.includes(field)) {
    targetElement.value = formatNumericForSpreadsheet(value);
  } else {
    targetElement.value = value || '';
  }
}

// Sincroniza um campo específico do card expandido para a spreadsheet
function syncSpreadsheetFromExpandedCard(itemId, field, value) {
  const row = document.querySelector(`tr[data-id="${itemId}"]`);
  if (!row) return;
  
  // Encontra o input/select correspondente na linha da spreadsheet
  const targetElement = row.querySelector(`[data-field="${field}"][data-id="${itemId}"].spreadsheet-input, [data-field="${field}"][data-id="${itemId}"].spreadsheet-select`);
  if (!targetElement) return;
  
  // Atualiza o valor
  if (targetElement.tagName === 'SELECT') {
    targetElement.value = value || '';
    if (field === 'status') {
      targetElement.className = 'spreadsheet-select ' + getStatusClass(value);
    }
  } else if (targetElement.type === 'date') {
    targetElement.value = formatDateForDateInput(value);
  } else if (numericSpreadsheetFields.includes(field)) {
    targetElement.value = formatNumericForSpreadsheet(value);
  } else {
    targetElement.value = value || '';
  }
}

// Salva alteração da planilha
async function spreadsheetSave(inputElement) {
  const id = inputElement.dataset.id;
  const field = inputElement.dataset.field;
  let value = inputElement.value;
  
  if (!id || !field) return;
  
  // Processa campos numéricos - remove formatação e converte para número
  if (numericSpreadsheetFields.includes(field)) {
    // Remove a máscara de milhares e converte para número
    const numValue = parseLocalizedNumber(value);
    value = numValue.toString();
  }
  
  // Campos de data já vêm em formato ISO do input type="date"
  // Não precisa de conversão adicional
  
  try {
    const now = new Date().toISOString();
    const updateData = { [field]: value, last_update: now };
    const updateResult = await window.api.updateTooling(id, updateData);
    
    // Atualiza o item no array local
    const item = toolingData.find(t => String(t.id) === String(id));
    if (item) {
      item[field] = value;
      item.last_update = now;
      
      // Atualiza os comentários se foram modificados pelo backend
      if (updateResult?.comments) {
        item.comments = updateResult.comments;
        updateCommentsDisplay(id);
      }
    }
    
    // Atualiza o display do Last Update
    updateLastUpdateDisplay(id, now);
    
    // Atualiza a barra de progresso se mudou tool_life_qty ou produced
    if (field === 'tooling_life_qty' || field === 'produced') {
      updateSpreadsheetProgressBar(id);
      
      // Recalcula o remaining e atualiza no toolingData e card expandido
      const item = toolingData.find(t => String(t.id) === String(id));
      if (item) {
        const toolingLife = parseLocalizedNumber(item.tooling_life_qty) || 0;
        const produced = parseLocalizedNumber(item.produced) || 0;
        const remaining = toolingLife - produced;
        const percentValue = toolingLife > 0 ? (produced / toolingLife) * 100 : 0;
        const percent = toolingLife > 0 ? percentValue.toFixed(1) : '0.0';
        
        // Atualiza toolingData
        item.remaining_tooling_life_pcs = remaining;
        item.percent_tooling_life = percent;
        
        // Sincroniza remaining e percent com o card expandido se estiver aberto
        const detailRow = document.querySelector(`.spreadsheet-detail-row[data-detail-for="${id}"]`);
        if (detailRow) {
          const remainingInput = detailRow.querySelector(`[data-field="remaining_tooling_life_pcs"][data-id="${id}"]`);
          if (remainingInput) {
            remainingInput.value = formatIntegerWithSeparators(remaining, { preserveEmpty: true });
          }
          const percentInput = detailRow.querySelector(`[data-field="percent_tooling_life"][data-id="${id}"]`);
          if (percentInput) {
            percentInput.value = percent + '%';
          }
          // Atualiza barra de progresso interna do card
          const lifecycleProgressFill = detailRow.querySelector('[data-lifecycle-progress-fill]');
          if (lifecycleProgressFill) {
            lifecycleProgressFill.style.width = percent + '%';
          }
          // Atualiza barra de progresso externa (header)
          const externalProgressFill = detailRow.querySelector('[data-progress-fill]');
          const externalProgressPercent = detailRow.querySelector('[data-progress-percent]');
          if (externalProgressFill) {
            externalProgressFill.style.width = percent + '%';
          }
          if (externalProgressPercent) {
            externalProgressPercent.textContent = percent + '%';
          }
        }
      }
    }
    
    // Recalcula expiration_date quando campos relevantes mudam
    const expirationFields = ['tooling_life_qty', 'produced', 'annual_volume_forecast', 'date_remaining_tooling_life'];
    if (expirationFields.includes(field)) {
      await recalculateExpirationForSpreadsheet(id);
    }
    
    // Feedback visual sutil
    inputElement.style.backgroundColor = '#e8f5e9';
    setTimeout(() => {
      inputElement.style.backgroundColor = '';
    }, 500);
    
    // Atualiza classe do select de status se necessário
    if (field === 'status' && inputElement.tagName === 'SELECT') {
      inputElement.className = 'spreadsheet-select ' + getStatusClass(value);
    }
    
    // Sincroniza com o card expandido se existir
    syncExpandedCardFromSpreadsheet(id, field, value);
    
  } catch (error) {
    console.error('Error saving spreadsheet data:', error);
    inputElement.style.backgroundColor = '#ffebee';
    setTimeout(() => {
      inputElement.style.backgroundColor = '';
    }, 1000);
  }
}

// Gera apenas o HEADER do card (versão super leve para lista inicial)
function buildToolingCardHeaderHTML(item, index, chainMembership) {
  const toolingLife = Number(parseLocalizedNumber(item.tooling_life_qty)) || Number(item.tooling_life_qty) || 0;
  const produced = Number(parseLocalizedNumber(item.produced)) || Number(item.produced) || 0;
  const classification = classifyToolingExpirationState(item);
  const expirationDateValue = classification.expirationDate || '';
  const expirationDisplay = formatDate(expirationDateValue || '');
  const percentUsedValue = toolingLife > 0 ? (produced / toolingLife) * 100 : 0;
  const percentUsed = toolingLife > 0 ? percentUsedValue.toFixed(1) : '0';
  const lastUpdateDisplay = formatDateTime(item.last_update);
  const normalizedStatus = (item.status || '').toString().trim().toLowerCase();
  const isObsolete = normalizedStatus === 'obsolete';
  const replacementIdValue = sanitizeReplacementId(item.replacement_tooling_id);
  const hasReplacementLink = replacementIdValue !== '';
  const membershipKey = String(item.id || '').trim();
  const hasChainMembership = chainMembership?.get(membershipKey) === true;
  const shouldShowChainIndicator = hasChainMembership;
  
  let statusIconHtml = '';
  let statusIconClass = '';
  
  if (classification.state === 'obsolete' || classification.state === 'obsolete-replaced') {
    statusIconHtml = '<i class="ph ph-fill ph-archive-box"></i>';
    statusIconClass = 'status-icon-obsolete';
  } else if (classification.state === 'expired') {
    statusIconHtml = '<i class="ph ph-fill ph-warning-circle"></i>';
    statusIconClass = 'status-icon-expired';
  } else if (classification.state === 'warning') {
    statusIconHtml = '<i class="ph ph-fill ph-warning"></i>';
    statusIconClass = 'status-icon-warning';
  }
  
  return `
    <div class="tooling-card" id="card-${index}" data-item-id="${item.id}" data-status="${normalizedStatus}" data-replacement-id="${hasReplacementLink ? replacementIdValue : ''}" data-body-loaded="false" data-supplier="${escapeHtml(item.supplier || '')}" data-chain-member="${hasChainMembership ? 'true' : 'false'}" data-has-incoming-chain="${hasChainMembership && !hasReplacementLink ? 'true' : 'false'}">
      <div class="tooling-card-header" onclick="toggleCard(${index})">
        <div class="tooling-card-header-top">
          ${statusIconHtml ? `<div class="card-status-icon ${statusIconClass}" title="${classification.label}">${statusIconHtml}</div>` : ''}
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
            <button type="button" class="tooling-chain-indicator" data-item-id="${item.id}" ${shouldShowChainIndicator ? '' : 'hidden'} title="View replacement chain" onclick="event.stopPropagation(); openReplacementTimelineForCard(document.getElementById('card-${index}'))">
              <i class="ph ph-git-branch"></i>
            </button>
            <div class="tooling-attachment-count" data-attachment-count data-item-id="${item.id}" hidden>
              <i class="ph ph-paperclip"></i>
              <span>0</span>
            </div>
            <div class="tooling-card-expand">
              <i class="ph ph-caret-down"></i>
            </div>
          </div>
        </div>
        <div class="tooling-card-header-divider"></div>
        <div class="tooling-card-main-info">
          <div class="tooling-info-item tooling-info-description">
            <span class="tooling-info-label">Tooling Description</span>
            <span class="tooling-info-value">${item.tool_description || 'N/A'}</span>
          </div>
          <div class="tooling-info-item tooling-info-description">
            <span class="tooling-info-label">PN Description</span>
            <span class="tooling-info-value">${item.pn_description || 'N/A'}</span>
          </div>
          <div class="tooling-info-item card-collapsible">
            <span class="tooling-info-label">Expiration</span>
            <span class="tooling-info-value" data-card-expiration>${expirationDisplay}</span>
          </div>
          <div class="tooling-info-item card-collapsible">
            <span class="tooling-info-label">Status</span>
            <span class="tooling-info-value" data-card-status>${item.status || 'N/A'}</span>
          </div>
          <div class="tooling-info-item card-collapsible">
            <span class="tooling-info-label">Steps</span>
            <span class="tooling-info-value" data-card-steps>${item.steps || 'N/A'}</span>
          </div>
          <div class="tooling-info-item tooling-info-progress">
            <div class="progress-bar">
              <div class="progress-fill" style="width: ${percentUsed}%" data-progress-fill></div>
            </div>
            <span class="progress-percentage" data-progress-percent>${percentUsed}%</span>
          </div>
        </div>
      </div>
      <!-- Body será carregado dinamicamente ao expandir -->
    </div>
  `;
}

// Gera apenas o BODY do card (chamado on-demand ao expandir)
function buildToolingCardBodyHTML(item, index, chainMembership, supplierContext) {
  const toolingLife = parseLocalizedNumber(item.tooling_life_qty) || 0;
  const produced = parseLocalizedNumber(item.produced) || 0;
  const remaining = toolingLife - produced;
  const forecast = parseLocalizedNumber(item.annual_volume_forecast) || 0;
  const hasForecast = String(item.annual_volume_forecast ?? '').trim() !== '';
  const toolingLifeDisplay = formatIntegerWithSeparators(toolingLife, { preserveEmpty: true });
  const producedDisplay = formatIntegerWithSeparators(produced, { preserveEmpty: true });
  const forecastDisplay = hasForecast ? formatIntegerWithSeparators(forecast, { preserveEmpty: true }) : '';
  
  const classification = classifyToolingExpirationState(item);
  let expirationDateValue = resolveToolingExpirationDate(item);
  const isAnalysisCompleted = item.analysis_completed === 1;
  
  // Usa a função de ícone para dentro do input do card
  const expirationIconHtml = getCardExpirationIcon(classification.state, isAnalysisCompleted);

  const expirationInputValue = expirationDateValue || '';
  const percentUsedValue = toolingLife > 0 ? (produced / toolingLife) * 100 : 0;
  const percentUsed = toolingLife > 0 ? percentUsedValue.toFixed(1) : '0';
  const remainingQty = remaining;
  const remainingDisplay = formatIntegerWithSeparators(remainingQty, { preserveEmpty: true });
  const lifecycleProgressPercent = Math.min(Math.max(percentUsedValue, 0), 100);
  const amountBrlValue = (() => {
    const raw = item.amount_brl;
    if (raw === null || raw === undefined) return '';
    const trimmed = String(raw).trim();
    if (trimmed === '') return '';
    const parsed = parseLocalizedNumber(trimmed);
    return Number.isNaN(parsed) ? '' : parsed;
  })();
  
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
  const replacementEditorVisibilityAttr = isObsolete ? 'aria-hidden="false"' : 'hidden aria-hidden="true"';
  const lastUpdateDisplay = formatDateTime(item.last_update);
  
  // Resto do body vem da função original buildToolingCardHTML...
  // Vou copiar só a parte do <div class="tooling-card-body"> em diante
  return `
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
        <button class="card-tab" onclick="switchCardTab(${index}, 'attachments')">
          <i class="ph ph-paperclip"></i>
          <span>Attachments</span>
        </button>
        <button class="card-tab" onclick="switchCardTab(${index}, 'step-tracking')">
          <i class="ph ph-path"></i>
          <span>Step Tracking</span>
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
                    <div class="lifecycle-progress-fill" style="width: ${lifecycleProgressPercent}%" data-lifecycle-progress-fill></div>
                  </div>
                </div>
              </div>
              <div class="detail-item detail-item-full">
                <span class="detail-label">Tooling Life (Qty)</span>
                <input type="text" class="detail-input" inputmode="numeric" data-mask="thousands" value="${toolingLifeDisplay}" data-field="tooling_life_qty" data-id="${item.id}" onchange="calculateLifecycle(${index})" oninput="calculateLifecycle(${index})">
              </div>
              <div class="detail-item detail-pair">
                <div class="detail-item">
                  <span class="detail-label">Produced</span>
                  <input type="text" class="detail-input" inputmode="numeric" data-mask="thousands" value="${producedDisplay}" data-field="produced" data-id="${item.id}" onchange="handleProducedChange(${index})" oninput="handleProducedChange(${index})">
                </div>
                <div class="detail-item">
                  <span class="detail-label">
                    Prod. Date
                    <i class="ph ph-info tooltip-icon" title="Production Date: Update whenever Produced changes to keep the timeline accurate." role="button" tabindex="0" onclick="openProductionInfoModal(event)" onkeydown="handleProductionInfoIconKey(event)"></i>
                  </span>
                  <input type="date" class="detail-input" value="${item.date_remaining_tooling_life || ''}" data-field="date_remaining_tooling_life" data-id="${item.id}" onchange="handleProductionDateChange(${index})">
                </div>
              </div>
              <div class="detail-item">
                <span class="detail-label">
                  Remaining
                  <i class="ph ph-info tooltip-icon" title="Formula: uses \\"Tooling Life (Qty)\\" minus \\"Produced\\"."></i>
                </span>
                <input type="text" class="detail-input calculated" value="${remainingDisplay}" data-field="remaining_tooling_life_pcs" data-id="${item.id}" readonly>
              </div>
              <div class="detail-item">
                <span class="detail-label">
                  % Used
                  <i class="ph ph-info tooltip-icon" title="Formula: (\\"Produced\\" ÷ \\"Tooling Life (Qty)\\") × 100."></i>
                </span>
                <input type="text" class="detail-input calculated" value="${percentUsed}%" data-field="percent_tooling_life" data-id="${item.id}" readonly>
              </div>
              <div class="detail-item detail-pair">
                <div class="detail-item">
                  <span class="detail-label">Annual Volume</span>
                  <input type="text" class="detail-input" inputmode="numeric" data-mask="thousands" value="${hasForecast ? forecastDisplay : ''}" data-field="annual_volume_forecast" data-id="${item.id}" onchange="handleForecastChange(${index})" oninput="handleForecastChange(${index})">
                </div>
                <div class="detail-item">
                  <span class="detail-label">Volume Date</span>
                  <input type="date" class="detail-input" value="${item.date_annual_volume || ''}" data-field="date_annual_volume" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                </div>
              </div>
              <div class="detail-item detail-item-full">
                <span class="detail-label">
                  Expiration (Calculated)
                  <i class="ph ph-info tooltip-icon" title="Formula: today's date + (\\"Remaining\\" ÷ \\"Annual Volume\\") years." role="button" tabindex="0" onclick="openExpirationInfoModal(event)" onkeydown="handleExpirationInfoIconKey(event)"></i>
                </span>
                <div class="input-with-icon">
                  ${expirationIconHtml}
                  <input type="date" class="detail-input calculated with-icon" value="${expirationInputValue}" data-field="expiration_date" data-id="${item.id}" readonly>
                </div>
              </div>
            </div>
            
            <!-- Column 2: Identification -->
            <div class="detail-group detail-grid">
              <div class="detail-group-title">Identification</div>
              <div class="detail-item detail-item-full">
                <span class="detail-label">PN</span>
                <input type="text" class="detail-input" value="${item.pn || ''}" data-field="pn" data-id="${item.id}" onchange="handlePNChange(${index}, ${item.id}, this)">
              </div>
              <div class="detail-item detail-item-full">
                <span class="detail-label">PN Description</span>
                <input type="text" class="detail-input" value="${item.pn_description || ''}" data-field="pn_description" data-id="${item.id}" onchange="handleCardTextFieldChange(${item.id}, this)">
              </div>
              <div class="detail-item detail-item-full">
                <span class="detail-label">Tooling Description</span>
                <input type="text" class="detail-input" value="${item.tool_description || ''}" data-field="tool_description" data-id="${item.id}" onchange="handleCardTextFieldChange(${item.id}, this)">
              </div>
              <div class="detail-item">
                <span class="detail-label">Status</span>
                <select class="detail-input" data-field="status" data-id="${item.id}" onchange="handleStatusSelectChange(${index}, ${item.id}, this)">
                  ${statusOptionsMarkup}
                </select>
              </div>
              <div class="detail-item">
                <span class="detail-label">
                  Steps
                  <i class="ph ph-info tooltip-icon" title="View all management steps" role="button" tabindex="0" onclick="openStepsInfoModal(event)" onkeydown="handleStepsInfoIconKey(event)"></i>
                </span>
                <select class="detail-input" data-field="steps" data-id="${item.id}" onchange="handleStepsSelectChange(${index}, ${item.id}, this)">
                  <option value="" ${!item.steps ? 'selected' : ''}></option>
                  <option value="1" ${item.steps === '1' ? 'selected' : ''}>1</option>
                  <option value="2" ${item.steps === '2' ? 'selected' : ''}>2</option>
                  <option value="3" ${item.steps === '3' ? 'selected' : ''}>3</option>
                  <option value="4" ${item.steps === '4' ? 'selected' : ''}>4</option>
                  <option value="5" ${item.steps === '5' ? 'selected' : ''}>5</option>
                  <option value="6" ${item.steps === '6' ? 'selected' : ''}>6</option>
                  <option value="7" ${item.steps === '7' ? 'selected' : ''}>7</option>
                </select>
                <span class="detail-sublabel" id="stepDescription_${item.id}" data-step-description></span>
              </div>
              ${(() => {
                const classification = classifyToolingExpirationState(item);
                const showCheckbox = classification.state === 'expired' || classification.state === 'warning';
                if (!showCheckbox) return '';
                const isChecked = item.analysis_completed === 1;
                return `
              <div class="detail-item detail-item-full">
                <label class="analysis-completed-checkbox">
                  <input type="checkbox" ${isChecked ? 'checked' : ''} onchange="handleAnalysisCompletedChange(${item.id}, this.checked)">
                  <span>Análise Concluída</span>
                </label>
              </div>`;
              })()}
              <div class="detail-item detail-item-full obsolete-link-field ${hasReplacementLink ? 'has-link' : ''}" data-obsolete-link ${replacementEditorVisibilityAttr}>
                <span class="detail-label">Replacement Tooling</span>
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
              </div>
            </div>

            <!-- Column 3: Comments -->
            <div class="detail-group detail-grid comments-group">
              <div class="detail-group-title">
                Comments
                <div class="comments-filter-wrapper">
                  <button type="button" class="btn-comments-filter" id="commentsFilterBtn_${item.id}" onclick="toggleCommentsFilterPopup(${item.id})" title="Filter comments">
                    <i class="ph ph-funnel"></i>
                  </button>
                  <div class="comments-filter-popup" id="commentsFilterPopup_${item.id}">
                    <div class="comments-filter-popup-header">Filter by keyword</div>
                    <div class="comments-filter-popup-options">
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'all')">All Comments</button>
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'created')">Created</button>
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'produced')">Produced (qty)</button>
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'tooling life')">Tooling Life (qty)</button>
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'forecast')">Forecast (qty)</button>
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'production date')">Production Date</button>
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'annual volume date')">Annual Volume Date</button>
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'expiration')">Expiration Date</button>
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'status')">Status</button>
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'steps')">Steps</button>
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'disposition')">Disposition</button>
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'category')">Category</button>
                      <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'responsible')">Cummins Responsible</button>
                    </div>
                  </div>
                </div>
              </div>
              <div class="card-comments-container">
                <div class="comments-input-area">
                  <textarea 
                    class="comment-input" 
                    id="commentInput_${item.id}"
                    placeholder="Add a comment..."
                    rows="3"
                    onkeydown="handleCommentKeydown(event, ${item.id})"
                  ></textarea>
                  <button 
                    class="btn-add-comment" 
                    onclick="addComment(${item.id})"
                    title="Add comment"
                  >
                    <i class="ph ph-plus"></i>
                  </button>
                </div>
                <div class="comments-list" id="commentsList_${item.id}">
                  ${buildCommentsListHTML(item.comments || '', item.id)}
                </div>
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
        <div class="tooling-details-grid" style="grid-template-columns: repeat(2, 1fr); gap: 24px;">
          <!-- Column 1: Documentation Fields -->
          <div class="detail-group detail-grid">
            <div class="detail-group-title">Documentation</div>
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
          
          <!-- Column 2: Supplier & Tooling -->
          <div class="detail-group detail-grid">
            <div class="detail-group-title">Supplier & Tooling</div>
            <div class="detail-item detail-pair">
              <div class="detail-item">
                <span class="detail-label">Supplier</span>
                <input type="text" class="detail-input" list="supplierList-${index}" value="${item.supplier || ''}" data-field="supplier" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                <datalist id="supplierList-${index}"></datalist>
              </div>
              <div class="detail-item">
                <span class="detail-label">Responsible</span>
                <input type="text" class="detail-input" list="ownerList-${index}" value="${item.cummins_responsible || ''}" data-field="cummins_responsible" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                <datalist id="ownerList-${index}"></datalist>
              </div>
            </div>
            <div class="detail-item detail-pair">
              <div class="detail-item">
                <span class="detail-label">Ownership</span>
                <input type="text" class="detail-input" value="${item.tool_ownership || ''}" data-field="tool_ownership" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">BU</span>
                <input type="text" class="detail-input" value="${item.bu || ''}" data-field="bu" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
            </div>
            <div class="detail-item detail-pair">
              <div class="detail-item">
                <span class="detail-label">Customer</span>
                <input type="text" class="detail-input" value="${item.customer || ''}" data-field="customer" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
              <div class="detail-item">
                <span class="detail-label">Tool Number</span>
                <input type="text" class="detail-input" value="${item.tool_number_arb || ''}" data-field="tool_number_arb" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
              </div>
            </div>
            <div class="detail-item detail-pair">
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

      <!-- Aba Step Tracking -->
      <div class="card-tab-content" data-tab="step-tracking">
        <div class="step-tracking-container" data-tooling-id="${item.id}">
          <div class="step-tracking-header">
            <span class="step-tracking-current">Current Step: <strong>${item.steps || 'Not set'}</strong></span>
            <button class="step-tracking-clear-btn" onclick="clearStepHistory(${item.id})" title="Clear history">
              <i class="ph ph-trash"></i>
            </button>
          </div>
          <div class="step-tracking-timeline" id="stepTimeline-${item.id}">
            <div class="step-tracking-loading">
              <i class="ph ph-spinner"></i>
              <span>Loading history...</span>
            </div>
          </div>
        </div>
      </div>

      <!-- Footer com Botões -->
      <div class="tooling-card-footer">
        <div class="card-last-update-snick" data-last-update-id="${item.id}">
          <i class="ph ph-clock-clockwise"></i>
          <span>Last update: ${lastUpdateDisplay}</span>
        </div>
        <div class="tooling-card-footer-actions">
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
}

function buildToolingCardHTML(item, index, chainMembership, supplierContext) {
    // Calcula expiration_date se não existir
    const toolingLife = parseLocalizedNumber(item.tooling_life_qty) || 0;
    const produced = parseLocalizedNumber(item.produced) || 0;
    const remaining = toolingLife - produced;
    const forecast = parseLocalizedNumber(item.annual_volume_forecast) || 0;
    const hasForecast = String(item.annual_volume_forecast ?? '').trim() !== '';
    const toolingLifeDisplay = formatIntegerWithSeparators(toolingLife, { preserveEmpty: true });
    const producedDisplay = formatIntegerWithSeparators(produced, { preserveEmpty: true });
    const forecastDisplay = hasForecast ? formatIntegerWithSeparators(forecast, { preserveEmpty: true }) : '';
    
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
      }
    }

    const expirationInputValue = expirationDateValue || '';
    const expirationDisplay = formatDate(expirationInputValue);

    const percentUsedValue = toolingLife > 0 ? (produced / toolingLife) * 100 : 0;
    const percentUsed = toolingLife > 0 ? percentUsedValue.toFixed(1) : '0';
    const remainingQty = remaining;
    const remainingDisplay = formatIntegerWithSeparators(remainingQty, { preserveEmpty: true });
    const lifecycleProgressPercent = Math.min(Math.max(percentUsedValue, 0), 100);
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
    const statusOptionsMarkup = buildStatusOptionsMarkup(item.status);
    const normalizedStatus = (item.status || '').toString().trim().toLowerCase();
    const isObsolete = normalizedStatus === 'obsolete';
    const replacementIdValue = sanitizeReplacementId(item.replacement_tooling_id);
    const hasReplacementLink = replacementIdValue !== '';
    const membershipKey = String(item.id || '').trim();
    const hasChainMembership = chainMembership?.get(membershipKey) === true;
    const shouldShowChainIndicator = hasChainMembership;
    const replacementPickerOptionsMarkup = buildReplacementPickerOptionsMarkup(item, index);
    const replacementPickerLabel = escapeHtml(
      hasReplacementLink
        ? (getReplacementOptionLabelById(replacementIdValue) || `${replacementIdValue}`)
        : DEFAULT_REPLACEMENT_PICKER_LABEL
    );
    const replacementChipVisibilityAttr = isObsolete ? 'aria-hidden="false"' : 'hidden aria-hidden="true"';
    const replacementEditorVisibilityAttr = isObsolete ? 'aria-hidden="false"' : 'hidden aria-hidden="true"';
    
    // Calcula status de vencimento para ícone
    const expirationStatus = getExpirationStatus(expirationInputValue);
    let statusIconHtml = '';
    let statusIconClass = '';
    
    // Verifica se está expirado por percentual de vida
    const isExpiredByPercent = percentUsedValue >= 100;
    const isObsoleteWithReplacement = isObsolete && hasReplacementLink;
    
    // Mostra ícone de status independente de estar em chain
    // Não mostrar ícone de warning/expired se for OBSOLETE com replacement
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
      <div class="tooling-card" id="card-${index}" data-item-id="${item.id}" data-status="${normalizedStatus}" data-replacement-id="${hasReplacementLink ? replacementIdValue : ''}" data-supplier="${escapeHtml(item.supplier || '')}" data-chain-member="${hasChainMembership ? 'true' : 'false'}" data-has-incoming-chain="${hasChainMembership && !hasReplacementLink ? 'true' : 'false'}">
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
              <button type="button" class="tooling-chain-indicator" data-item-id="${item.id}" ${shouldShowChainIndicator ? '' : 'hidden'} title="View replacement chain" onclick="event.stopPropagation(); openReplacementTimelineForCard(document.getElementById('card-${index}'))">
                <i class="ph ph-git-branch"></i>
              </button>
              <div class="tooling-attachment-count" data-attachment-count data-item-id="${item.id}" hidden>
                <i class="ph ph-paperclip"></i>
                <span>0</span>
              </div>
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
              <span class="tooling-info-value" data-card-expiration>${expirationDisplay}</span>
            </div>
            <div class="tooling-info-item card-collapsible">
              <span class="tooling-info-label">Status</span>
              <span class="tooling-info-value" data-card-status>${item.status || 'N/A'}</span>
            </div>
            <div class="tooling-info-item card-collapsible">
              <span class="tooling-info-label">Steps</span>
              <span class="tooling-info-value" data-card-steps>${item.steps || 'N/A'}</span>
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
            <button class="card-tab" onclick="switchCardTab(${index}, 'attachments')">
              <i class="ph ph-paperclip"></i>
              <span>Attachments</span>
            </button>
            <button class="card-tab" onclick="switchCardTab(${index}, 'step-tracking')">
              <i class="ph ph-path"></i>
              <span>Step Tracking</span>
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
                        <div class="lifecycle-progress-fill" style="width: ${lifecycleProgressPercent}%" data-lifecycle-progress-fill></div>
                      </div>
                    </div>
                  </div>
                  <div class="detail-item detail-item-full">
                    <span class="detail-label">Tooling Life (Qty)</span>
                    <input type="text" class="detail-input" inputmode="numeric" data-mask="thousands" value="${toolingLifeDisplay}" data-field="tooling_life_qty" data-id="${item.id}" onchange="calculateLifecycle(${index})" oninput="calculateLifecycle(${index})">
                  </div>
                  <div class="detail-item detail-pair">
                    <div class="detail-item">
                      <span class="detail-label">Produced</span>
                      <input type="text" class="detail-input" inputmode="numeric" data-mask="thousands" value="${producedDisplay}" data-field="produced" data-id="${item.id}" onchange="handleProducedChange(${index})" oninput="handleProducedChange(${index})">
                    </div>
                    <div class="detail-item">
                      <span class="detail-label">
                        Prod. Date
                        <i class="ph ph-info tooltip-icon" title="Production Date: Update whenever Produced changes to keep the timeline accurate." role="button" tabindex="0" onclick="openProductionInfoModal(event)" onkeydown="handleProductionInfoIconKey(event)"></i>
                      </span>
                      <input type="date" class="detail-input" value="${item.date_remaining_tooling_life || ''}" data-field="date_remaining_tooling_life" data-id="${item.id}" onchange="handleProductionDateChange(${index})">
                    </div>
                  </div>
                  <div class="detail-item">
                    <span class="detail-label">
                      Remaining
                      <i class="ph ph-info tooltip-icon" title="Formula: uses \"Tooling Life (Qty)\" minus \"Produced\"."></i>
                    </span>
                    <input type="text" class="detail-input calculated" value="${remainingDisplay}" data-field="remaining_tooling_life_pcs" data-id="${item.id}" readonly>
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
                      <span class="detail-label">Annual Volume</span>
                      <input type="text" class="detail-input" inputmode="numeric" data-mask="thousands" value="${hasForecast ? forecastDisplay : ''}" data-field="annual_volume_forecast" data-id="${item.id}" onchange="handleForecastChange(${index})" oninput="handleForecastChange(${index})">
                    </div>
                    <div class="detail-item">
                      <span class="detail-label">Volume Date</span>
                      <input type="date" class="detail-input" value="${item.date_annual_volume || ''}" data-field="date_annual_volume" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                    </div>
                  </div>
                  <div class="detail-item detail-item-full">
                    <span class="detail-label">
                      Expiration (Calculated)
                      <i class="ph ph-info tooltip-icon" title="Formula: today's date + (\"Remaining\" ÷ \"Annual Volume\") years." role="button" tabindex="0" onclick="openExpirationInfoModal(event)" onkeydown="handleExpirationInfoIconKey(event)"></i>
                    </span>
                    ${hasExpirationIcon ? `
                    <div class="input-with-icon">
                      <i class="${expirationIconClass} input-icon" style="color: ${expirationIconColor}"></i>
                      <input type="date" class="detail-input calculated with-icon" value="${expirationInputValue}" data-field="expiration_date" data-id="${item.id}" readonly>
                    </div>
                    ` : `
                    <input type="date" class="detail-input calculated" value="${expirationInputValue}" data-field="expiration_date" data-id="${item.id}" readonly>
                    `}
                  </div>
                </div>
                
                <!-- Column 2: Identification -->
                <div class="detail-group detail-grid">
                  <div class="detail-group-title">Identification</div>
                  <div class="detail-item detail-item-full">
                    <span class="detail-label">PN</span>
                    <input type="text" class="detail-input" value="${item.pn || ''}" data-field="pn" data-id="${item.id}" onchange="handlePNChange(${index}, ${item.id}, this)">
                  </div>
                  <div class="detail-item detail-item-full">
                    <span class="detail-label">PN Description</span>
                    <input type="text" class="detail-input" value="${item.pn_description || ''}" data-field="pn_description" data-id="${item.id}" onchange="handleCardTextFieldChange(${item.id}, this)">
                  </div>
                  <div class="detail-item detail-item-full">
                    <span class="detail-label">Tooling Description</span>
                    <input type="text" class="detail-input" value="${item.tool_description || ''}" data-field="tool_description" data-id="${item.id}" onchange="handleCardTextFieldChange(${item.id}, this)">
                  </div>
                  <div class="detail-item">
                    <span class="detail-label">Status</span>
                    <select class="detail-input" data-field="status" data-id="${item.id}" onchange="handleStatusSelectChange(${index}, ${item.id}, this)">
                      ${statusOptionsMarkup}
                    </select>
                  </div>
                  <div class="detail-item">
                    <span class="detail-label">Steps</span>
                    <select class="detail-input" data-field="steps" data-id="${item.id}" onchange="handleStepsSelectChange(${index}, ${item.id}, this)">
                      <option value="" ${!item.steps ? 'selected' : ''}></option>
                      <option value="1" ${item.steps === '1' ? 'selected' : ''}>1</option>
                      <option value="2" ${item.steps === '2' ? 'selected' : ''}>2</option>
                      <option value="3" ${item.steps === '3' ? 'selected' : ''}>3</option>
                      <option value="4" ${item.steps === '4' ? 'selected' : ''}>4</option>
                      <option value="5" ${item.steps === '5' ? 'selected' : ''}>5</option>
                      <option value="6" ${item.steps === '6' ? 'selected' : ''}>6</option>
                      <option value="7" ${item.steps === '7' ? 'selected' : ''}>7</option>
                    </select>
                    <span class="detail-sublabel" id="stepDescription_${item.id}" data-step-description></span>
                  </div>
                  ${(() => {
                    const classification = classifyToolingExpirationState(item);
                    const showCheckbox = classification.state === 'expired' || classification.state === 'warning';
                    if (!showCheckbox) return '';
                    const isChecked = item.analysis_completed === 1;
                    return `
                  <div class="detail-item detail-item-full">
                    <label class="analysis-completed-checkbox">
                      <input type="checkbox" ${isChecked ? 'checked' : ''} onchange="handleAnalysisCompletedChange(${item.id}, this.checked)">
                      <span>Análise Concluída</span>
                    </label>
                  </div>`;
                  })()}
                  <div class="detail-item detail-item-full obsolete-link-field ${hasReplacementLink ? 'has-link' : ''}" data-obsolete-link ${replacementEditorVisibilityAttr}>
                    <span class="detail-label">Replacement Tooling</span>
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
                  </div>
                </div>

                <!-- Column 3: Comments -->
                  <div class="detail-group detail-grid comments-group">
                    <div class="detail-group-title">
                      Comments
                      <div class="comments-filter-wrapper">
                        <button type="button" class="btn-comments-filter" id="commentsFilterBtn_${item.id}" onclick="toggleCommentsFilterPopup(${item.id})" title="Filter comments">
                          <i class="ph ph-funnel"></i>
                        </button>
                        <div class="comments-filter-popup" id="commentsFilterPopup_${item.id}">
                          <div class="comments-filter-popup-header">Filter by keyword</div>
                          <div class="comments-filter-popup-options">
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'all')">All Comments</button>
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'created')">Created</button>
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'produced')">Produced (qty)</button>
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'tooling life')">Tooling Life (qty)</button>
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'forecast')">Forecast (qty)</button>
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'production date')">Production Date</button>
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'annual volume date')">Annual Volume Date</button>
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'expiration')">Expiration Date</button>
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'status')">Status</button>
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'steps')">Steps</button>
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'disposition')">Disposition</button>
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'category')">Category</button>
                            <button type="button" class="comments-filter-option" onclick="applyCommentsFilter(${item.id}, 'responsible')">Cummins Responsible</button>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div class="card-comments-container">
                      <div class="comments-input-area">
                        <textarea 
                          class="comment-input" 
                          id="commentInput_${item.id}"
                          placeholder="Add a comment..."
                          rows="3"
                          onkeydown="handleCommentKeydown(event, ${item.id})"
                        ></textarea>
                        <button 
                          class="btn-add-comment" 
                          onclick="addComment(${item.id})"
                          title="Add comment"
                        >
                          <i class="ph ph-plus"></i>
                        </button>
                      </div>
                      <div class="comments-list" id="commentsList_${item.id}">
                        ${buildCommentsListHTML(item.comments || '', item.id)}
                      </div>
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
            <div class="tooling-details-grid" style="grid-template-columns: repeat(2, 1fr); gap: 24px;">
              <!-- Column 1: Documentation Fields -->
              <div class="detail-group detail-grid">
                <div class="detail-group-title">Documentation</div>
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
              
              <!-- Column 2: Supplier & Tooling -->
              <div class="detail-group detail-grid">
                <div class="detail-group-title">Supplier & Tooling</div>
                <div class="detail-item detail-pair">
                  <div class="detail-item">
                    <span class="detail-label">Supplier</span>
                    <input type="text" class="detail-input" list="supplierList-${index}" value="${item.supplier || ''}" data-field="supplier" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                    <datalist id="supplierList-${index}"></datalist>
                  </div>
                  <div class="detail-item">
                    <span class="detail-label">Responsible</span>
                    <input type="text" class="detail-input" list="ownerList-${index}" value="${item.cummins_responsible || ''}" data-field="cummins_responsible" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                    <datalist id="ownerList-${index}"></datalist>
                  </div>
                </div>
                <div class="detail-item detail-pair">
                  <div class="detail-item">
                    <span class="detail-label">Ownership</span>
                    <input type="text" class="detail-input" value="${item.tool_ownership || ''}" data-field="tool_ownership" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                  </div>
                  <div class="detail-item">
                    <span class="detail-label">BU</span>
                    <input type="text" class="detail-input" value="${item.bu || ''}" data-field="bu" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                  </div>
                </div>
                <div class="detail-item detail-pair">
                  <div class="detail-item">
                    <span class="detail-label">Customer</span>
                    <input type="text" class="detail-input" value="${item.customer || ''}" data-field="customer" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                  </div>
                  <div class="detail-item">
                    <span class="detail-label">Tool Number</span>
                    <input type="text" class="detail-input" value="${item.tool_number_arb || ''}" data-field="tool_number_arb" data-id="${item.id}" onchange="autoSaveTooling(${item.id})">
                  </div>
                </div>
                <div class="detail-item detail-pair">
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

          <!-- Aba Step Tracking -->
          <div class="card-tab-content" data-tab="step-tracking">
            <div class="step-tracking-container" data-tooling-id="${item.id}">
              <div class="step-tracking-header">
                <span class="step-tracking-current">Current Step: <strong>${item.steps || 'Not set'}</strong></span>
                <button class="step-tracking-clear-btn" onclick="clearStepHistory(${item.id})" title="Clear history">
                  <i class="ph ph-trash"></i>
                </button>
              </div>
              <div class="step-tracking-timeline" id="stepTimeline-${item.id}">
                <div class="step-tracking-loading">
                  <i class="ph ph-spinner"></i>
                  <span>Loading history...</span>
                </div>
              </div>
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
}

function getSupplierOptions() {
  return Array.isArray(suppliersData)
    ? [...new Set(
        suppliersData
          .map(item => String(item?.supplier || '').trim())
          .filter(name => name.length > 0)
      )].sort((a, b) => a.localeCompare(b, 'pt-BR'))
    : [];
}

function getResponsibleOptions(sourceData = []) {
  if (Array.isArray(responsiblesData) && responsiblesData.length > 0) {
    return [...responsiblesData].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  }

  const fallbackSource = Array.isArray(sourceData) && sourceData.length > 0
    ? sourceData
    : (Array.isArray(toolingData) ? toolingData : []);

  return [...new Set(
    fallbackSource
      .map(item => String(item?.cummins_responsible || '').trim())
      .filter(name => name.length > 0)
  )].sort((a, b) => a.localeCompare(b, 'pt-BR'));
}

function populateCardDataListForIndex(index, supplierOptions, responsibleOptions) {
  const suppliers = Array.isArray(supplierOptions) ? supplierOptions : getSupplierOptions();
  const responsibles = Array.isArray(responsibleOptions)
    ? responsibleOptions
    : getResponsibleOptions();

  const supplierList = document.getElementById(`supplierList-${index}`);
  if (supplierList) {
    supplierList.innerHTML = suppliers
      .map(name => `<option value="${escapeHtml(name)}"></option>`)
      .join('');
  }

  const ownerList = document.getElementById(`ownerList-${index}`);
  if (ownerList) {
    ownerList.innerHTML = responsibles
      .map(name => `<option value="${escapeHtml(name)}"></option>`)
      .join('');
  }
}

function populateCardDataLists(sortedData) {
  if (!Array.isArray(sortedData) || sortedData.length === 0) {
    return;
  }

  const supplierOptions = getSupplierOptions();
  const responsibleOptions = getResponsibleOptions(sortedData);

  sortedData.forEach((_, index) => {
    populateCardDataListForIndex(index, supplierOptions, responsibleOptions);
  });
}

async function hydrateCardsAfterRender(sortedData, renderToken, supplierContext, chainMembership) {
  if (supplierContext) {
    hydrateAttachmentBadges(sortedData, renderToken, supplierContext).catch((error) => {
    });
  }

  hydrateChainIndicators(sortedData, renderToken, chainMembership).catch((error) => {
  });

  sortedData.forEach((item, index) => {
    syncReplacementLinkControls(index, sanitizeReplacementId(item.replacement_tooling_id));
    updateCardStatusAttribute(index, item.status || '');
  });

  // Populate supplier and owner datalists for all cards
  populateCardDataLists(sortedData);

  primeCardSnapshots(sortedData);
  refreshCardCarouselState();
  
  // Atualizar checkboxes se modo de seleção estiver ativo
  if (selectionModeActive) {
    updateCardCheckboxes();
  }
}

// Alterna expansão do card
async function toggleAllCards() {
  const cards = document.querySelectorAll('.tooling-card');
  const expandBtn = document.getElementById('floatingExpandBtn');
  
  if (!expandBtn) {
    return;
  }
  
  if (cards.length === 0) {
    return;
  }

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

  // Atualiza o ícone do botão
  const icon = expandBtn.querySelector('i');
  if (icon) {
    if (shouldExpand) {
      icon.className = 'ph ph-arrows-in';
      expandBtn.title = 'Collapse All';
    } else {
      icon.className = 'ph ph-arrows-out';
      expandBtn.title = 'Expand All';
    }
  }
}

function navigateCarousel(cardIndex, direction) {
  // Tenta encontrar o card normal primeiro
  let card = document.getElementById(`card-${cardIndex}`);
  
  // Se não encontrar, tenta no spreadsheet expandido
  if (!card) {
    card = document.querySelector(`.spreadsheet-card-container[data-item-index="${cardIndex}"]`);
  }
  
  if (!card) return;
  
  const track = card.querySelector('[data-carousel-track]');
  const carousel = card.querySelector('.tooling-details-carousel');
  const prevBtn = card.querySelector('.carousel-nav-prev');
  const nextBtn = card.querySelector('.carousel-nav-next');
  if (!track || !carousel) return;
  
  const columns = Array.from(track.children);
  
  // Get current carousel width (always fresh)
  const carouselWidth = carousel.offsetWidth;
  const gap = 14;
  
  // Use stored index when available (fallback to transform parsing)
  let currentIndex = parseInt(track.dataset.carouselIndex || '0', 10);
  if (Number.isNaN(currentIndex)) currentIndex = 0;
  
  if (!track.dataset.carouselIndex) {
    const currentTransform = getComputedStyle(track).transform;
    if (currentTransform !== 'none') {
      const matrix = currentTransform.match(/matrix\(([^)]+)\)/);
      if (matrix) {
        const currentX = parseFloat(matrix[1].split(',')[4]) || 0;
        currentIndex = Math.round(Math.abs(currentX) / (carouselWidth + gap));
      }
    }
  }
  
  // Navigate
  if (direction === 'next') {
    currentIndex++;
    if (currentIndex >= columns.length) currentIndex = columns.length - 1;
  } else {
    currentIndex--;
    if (currentIndex < 0) currentIndex = 0;
  }
  
  // Calculate new position with current width
  const newX = -(currentIndex * (carouselWidth + gap));
  
  track.style.transform = `translateX(${newX}px)`;
  track.dataset.carouselIndex = String(currentIndex);
  
  // Update button states
  if (prevBtn) prevBtn.disabled = currentIndex === 0;
  if (nextBtn) nextBtn.disabled = currentIndex === columns.length - 1;
}

function refreshCardCarouselState() {
  const isCarouselMode = window.innerWidth <= DATA_TAB_CAROUSEL_BREAKPOINT;
  document.querySelectorAll('.tooling-card').forEach(card => {
    const track = card.querySelector('[data-carousel-track]');
    const carousel = card.querySelector('.tooling-details-carousel');
    const prevBtn = card.querySelector('.carousel-nav-prev');
    const nextBtn = card.querySelector('.carousel-nav-next');
    if (!track || !carousel || !prevBtn || !nextBtn) {
      return;
    }

    const columns = Array.from(track.children);
    const maxIndex = Math.max(0, columns.length - 1);
    const gap = 14;

    if (isCarouselMode) {
      let currentIndex = parseInt(track.dataset.carouselIndex || '0', 10);
      if (Number.isNaN(currentIndex)) {
        currentIndex = 0;
      }
      currentIndex = Math.min(currentIndex, maxIndex);
      track.dataset.carouselIndex = String(currentIndex);

      const columnWidth = carousel.offsetWidth;
      const newX = -(currentIndex * (columnWidth + gap));
      track.style.transform = `translateX(${newX}px)`;

      prevBtn.disabled = currentIndex === 0;
      nextBtn.disabled = maxIndex === 0 || currentIndex === maxIndex;
    } else {
      track.style.transform = 'translateX(0)';
      delete track.dataset.carouselIndex;
      prevBtn.disabled = false;
      nextBtn.disabled = false;
    }
  });
}

// Reset carousel position on window resize
let resizeTimeout;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimeout);
  resizeTimeout = setTimeout(() => {
    refreshCardCarouselState();
  }, 150);
});

document.addEventListener('DOMContentLoaded', () => {
  refreshCardCarouselState();
});

function toggleCard(index) {
  const card = document.getElementById(`card-${index}`);
  const wasExpanded = card.classList.contains('expanded');
  
  // Se está fechando o card, salva em background
  if (wasExpanded) {
    const itemId = card.getAttribute('data-item-id');
    if (itemId) {
      Promise.resolve().then(() => saveToolingQuietly(itemId));
    }
    card.classList.remove('expanded');
  } else {
    // Está abrindo o card
    
    // LAZY LOAD: Carrega body apenas se ainda não foi carregado
    const bodyLoaded = card.getAttribute('data-body-loaded') === 'true';
    if (!bodyLoaded) {
      const itemId = card.getAttribute('data-item-id');
      const item = toolingData.find(t => String(t.id) === String(itemId));
      
      if (item) {
        // Gera o body completo agora
        const supplierContext = selectedSupplier || currentSupplier || '';
        const chainMembership = new Map(); // Já foi computado antes
        const bodyHTML = buildToolingCardBodyHTML(item, index, chainMembership, supplierContext);
        
        // Insere o body no card
        card.insertAdjacentHTML('beforeend', bodyHTML);
        populateCardDataListForIndex(index);
        card.setAttribute('data-body-loaded', 'true');
        applyInitialThousandsMask(card);
        
        // Restaura reminders de data persistentes
        restoreDateReminders(itemId);
        
        // Initialize drag and drop for attachment dropzone
        const dropzone = card.querySelector('.card-attachments-dropzone');
        if (dropzone) {
          initCardAttachmentDragAndDrop(dropzone, itemId);
        }
        
        // Initialize step description
        const stepSelect = card.querySelector(`select[data-field="steps"][data-id="${itemId}"]`);
        if (stepSelect) {
          const stepValue = stepSelect.value;
          const descriptionLabel = document.getElementById(`stepDescription_${itemId}`);
          if (descriptionLabel) {
            const description = getStepDescription(stepValue);
            descriptionLabel.textContent = description;
            descriptionLabel.style.display = description ? 'block' : 'none';
          }
        }
      }
    }
    
    card.classList.add('expanded');
    
    // Initialize carousel buttons state
    const track = card.querySelector('[data-carousel-track]');
    const prevBtn = card.querySelector('.carousel-nav-prev');
    const nextBtn = card.querySelector('.carousel-nav-next');
    if (track && prevBtn && nextBtn) {
      const columns = Array.from(track.children);
      prevBtn.disabled = true; // Start at first column
      nextBtn.disabled = columns.length <= 1; // Disable if only one column
      track.style.transform = 'translateX(0)'; // Reset to first column
      track.dataset.carouselIndex = '0';
    }
    
    // Carrega dados do card em background, sem bloquear
    setTimeout(() => {
      calculateExpirationDate(index, null, true); // skipSave = true
      const itemId = card.getAttribute('data-item-id');
      if (itemId) {
        loadCardAttachments(itemId).catch(err => {
        });
        
        // Cria snapshot inicial para detectar alterações
        const snapshotKey = getSnapshotKey(itemId);
        const values = collectCardDomValues(itemId);
        if (values) {
          // Adiciona comentários do item ao snapshot
          const itemData = toolingData.find(t => String(t.id) === String(itemId));
          if (itemData && itemData.comments) {
            values.comments = itemData.comments;
          }
          cardSnapshotStore.set(snapshotKey, serializeCardValues(values));
        }
      }
    }, 0);
    
    ensureCardVisible(card);
  }
}

// Troca de aba dentro do card
function switchCardTab(cardIndex, tabName) {
  // Tenta encontrar o card normal primeiro
  let card = document.getElementById(`card-${cardIndex}`);
  
  // Se não encontrar, tenta no spreadsheet expandido
  if (!card) {
    card = document.querySelector(`.spreadsheet-card-container[data-item-index="${cardIndex}"]`);
  }
  
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
  if (itemId && tabName === 'attachments') {
    loadCardAttachments(itemId);
  }
  if (itemId && tabName === 'step-tracking') {
    loadStepHistory(itemId);
  }
}

async function loadStepHistory(toolingId) {
  const timeline = document.getElementById(`stepTimeline-${toolingId}`);
  if (!timeline) return;
  
  try {
    const history = await window.api.getStepHistory(toolingId);
    
    if (!history || history.length === 0) {
      timeline.innerHTML = `
        <div class="step-tracking-empty">
          <i class="ph ph-clock-countdown"></i>
          <span>No step changes recorded yet</span>
          <p>Changes to the Step field will appear here as a timeline.</p>
        </div>
      `;
      return;
    }
    
    // Build timeline HTML
    let html = '<div class="step-timeline-list">';
    
    history.forEach((entry, index) => {
      const date = new Date(entry.changed_at);
      const formattedDate = date.toLocaleDateString('pt-BR', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      });
      const formattedTime = date.toLocaleTimeString('pt-BR', {
        hour: '2-digit',
        minute: '2-digit'
      });
      
      const oldStep = entry.old_step || 'Not set';
      const newStep = entry.new_step || 'Not set';
      const isFirst = index === 0;
      
      html += `
        <div class="step-timeline-item ${isFirst ? 'latest' : ''}">
          <div class="step-timeline-marker">
            <div class="step-timeline-dot"></div>
            ${index < history.length - 1 ? '<div class="step-timeline-line"></div>' : ''}
          </div>
          <div class="step-timeline-content">
            <div class="step-timeline-date">
              <i class="ph ph-calendar-blank"></i>
              <span>${formattedDate}</span>
              <span class="step-timeline-time">${formattedTime}</span>
            </div>
            <div class="step-timeline-change">
              <span class="step-from">${escapeHtml(oldStep)}</span>
              <i class="ph ph-arrow-right"></i>
              <span class="step-to">${escapeHtml(newStep)}</span>
            </div>
          </div>
        </div>
      `;
    });
    
    html += '</div>';
    timeline.innerHTML = html;
    
  } catch (error) {
    console.error('[StepHistory] Error loading step history:', error);
    timeline.innerHTML = `
      <div class="step-tracking-error">
        <i class="ph ph-warning-circle"></i>
        <span>Failed to load step history</span>
      </div>
    `;
  }
}

// State for clear step history modal
let clearStepHistoryState = {
  toolingId: null,
  descriptor: ''
};

function clearStepHistory(toolingId) {
  const overlay = document.getElementById('clearStepHistoryOverlay');
  const descriptionEl = document.getElementById('clearStepItemDescription');
  
  if (!overlay || !descriptionEl) return;
  
  const item = toolingData.find(tool => String(tool.id) === String(toolingId));
  const descriptorParts = [];
  if (item?.pn) descriptorParts.push(item.pn);
  if (item?.tool_description) descriptorParts.push(item.tool_description);
  const descriptor = descriptorParts.join(' - ') || 'this tooling';
  
  clearStepHistoryState = {
    toolingId,
    descriptor
  };
  
  descriptionEl.textContent = descriptor;
  overlay.classList.add('active');
}

function cancelClearStepHistory() {
  const overlay = document.getElementById('clearStepHistoryOverlay');
  if (overlay) {
    overlay.classList.remove('active');
  }
  clearStepHistoryState = { toolingId: null, descriptor: '' };
}

async function confirmClearStepHistory() {
  const { toolingId } = clearStepHistoryState;
  
  if (!toolingId) {
    cancelClearStepHistory();
    return;
  }
  
  try {
    await window.api.clearStepHistory(toolingId);
    loadStepHistory(toolingId);
    showToast('Step history cleared', 'success');
  } catch (error) {
    console.error('[StepHistory] Error clearing step history:', error);
    showToast('Failed to clear step history', 'error');
  }
  
  cancelClearStepHistory();
}

// State for clear ALL step history modal
let clearAllStepHistoryState = {
  code: ''
};

function openClearAllStepHistoryModal() {
  const overlay = document.getElementById('clearAllStepHistoryOverlay');
  const codeEl = document.getElementById('clearAllStepsConfirmCode');
  const inputEl = document.getElementById('clearAllStepsConfirmInput');
  
  if (!overlay || !codeEl || !inputEl) {
    console.error('[ClearAllStepHistory] Missing elements!');
    return;
  }
  
  // Close the supplier menu overlay first
  closeSupplierFilterOverlay();
  
  clearAllStepHistoryState.code = generateConfirmationCode();
  codeEl.textContent = clearAllStepHistoryState.code;
  inputEl.value = '';
  
  overlay.classList.add('active');
  inputEl.focus();
}

function cancelClearAllStepHistory() {
  const overlay = document.getElementById('clearAllStepHistoryOverlay');
  const inputEl = document.getElementById('clearAllStepsConfirmInput');
  
  if (overlay) {
    overlay.classList.remove('active');
  }
  if (inputEl) {
    inputEl.value = '';
  }
  clearAllStepHistoryState.code = '';
}

async function handleClearAllStepHistoryConfirmation() {
  const inputEl = document.getElementById('clearAllStepsConfirmInput');
  
  if (!inputEl) {
    cancelClearAllStepHistory();
    return;
  }
  
  const enteredCode = inputEl.value.trim().toUpperCase();
  const expectedCode = clearAllStepHistoryState.code.toUpperCase();
  
  console.log('[ClearAllSteps] Entered:', enteredCode, 'Expected:', expectedCode);
  
  if (enteredCode !== expectedCode) {
    inputEl.classList.add('error');
    inputEl.focus();
    setTimeout(() => inputEl.classList.remove('error'), 500);
    return;
  }
  
  try {
    console.log('[ClearAllSteps] Calling API...');
    const result = await window.api.clearAllStepHistory();
    console.log('[ClearAllSteps] Result:', result);
    showToast('All step checkboxes cleared successfully (steps preserved)', 'success');
    
    // Reload tooling data to reflect the changes
    if (selectedSupplier) {
      await loadToolingBySupplier(selectedSupplier);
    }
  } catch (error) {
    console.error('[StepHistory] Error clearing all step history:', error);
    showToast('Failed to clear all step history', 'error');
  }
  
  cancelClearAllStepHistory();
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
    if (normalizedId === null || !toolingItem || !toolingItem.supplier) {
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
    showNotification('Erro ao anexar arquivo', 'error');
  }
}

// Salva alterações do ferramental
async function saveTooling(id) {
  try {
    const prepared = buildCardPayloadFromDom(id);
    if (!prepared) {
      showNotification('Não foi possível localizar os dados do card.', 'error');
      return;
    }
    if (!prepared.hasChanges) {
      showNotification('Nenhuma alteração detectada.', 'info');
      return;
    }

    const updateResult = await window.api.updateTooling(id, prepared.payload);
    if (updateResult?.comments) {
      prepared.payload.comments = updateResult.comments;
    }
    showNotification('Dados salvos com sucesso!');

    const snapshotKey = getSnapshotKey(id);
    const refreshedSnapshot = collectCardDomValues(id);
    if (refreshedSnapshot) {
      if (prepared.payload.comments !== undefined) {
        refreshedSnapshot.comments = prepared.payload.comments;
      }
      // Adiciona expiration_date ao snapshot (não é coletado do DOM porque tem classe 'calculated')
      if (prepared.payload.expiration_date !== undefined) {
        refreshedSnapshot.expiration_date = prepared.payload.expiration_date;
      }
      cardSnapshotStore.set(snapshotKey, serializeCardValues(refreshedSnapshot));
    } else {
      cardSnapshotStore.set(snapshotKey, prepared.serialized);
    }

    const index = toolingData.findIndex(item => Number(item.id) === Number(id));
    if (index !== -1) {
      toolingData[index] = { ...toolingData[index], ...prepared.payload };
      calculateExpirationDate(index);
      if (updateResult?.comments) {
        updateCommentsDisplay(id);
      }
    }

    scheduleInterfaceRefresh('manual-save');
  } catch (error) {
    showNotification('Erro ao salvar dados', 'error');
  }
}

// Salva sem mostrar notificação (usado ao trocar de card)
async function saveToolingQuietly(id) {
  try {
    const prepared = buildCardPayloadFromDom(id);
    if (!prepared || !prepared.hasChanges) {
      return;
    }

    const updateResult = await window.api.updateTooling(id, prepared.payload);
    if (updateResult?.comments) {
      prepared.payload.comments = updateResult.comments;
    }

    const snapshotKey = getSnapshotKey(id);
    const refreshedSnapshot = collectCardDomValues(id);
    if (refreshedSnapshot) {
      if (prepared.payload.comments !== undefined) {
        refreshedSnapshot.comments = prepared.payload.comments;
      }
      // Adiciona expiration_date ao snapshot (não é coletado do DOM porque tem classe 'calculated')
      if (prepared.payload.expiration_date !== undefined) {
        refreshedSnapshot.expiration_date = prepared.payload.expiration_date;
      }
      cardSnapshotStore.set(snapshotKey, serializeCardValues(refreshedSnapshot));
    } else {
      cardSnapshotStore.set(snapshotKey, prepared.serialized);
    }

    const index = toolingData.findIndex(item => Number(item.id) === Number(id));
    if (index !== -1) {
      // Atualiza o last_update com a data atual
      const now = new Date().toISOString();
      toolingData[index] = { ...toolingData[index], ...prepared.payload, last_update: now };
      
      if (updateResult?.comments) {
        updateCommentsDisplay(id);
      }
      
      // Atualiza o display do Last Update em tempo real
      updateLastUpdateDisplay(id, now);
      
      // Atualiza métricas do card do supplier em tempo real (busca dados frescos do banco)
      if (selectedSupplier) {
        refreshSupplierCardMetricsFromDB(selectedSupplier);
      }
      
      // Sincroniza a linha da spreadsheet com os dados atualizados
      syncSpreadsheetRowFromCard(id);
    }

    scheduleInterfaceRefresh('autosave');
  } catch (error) {
  }
}

// Atualiza o display do Last Update no card
function updateLastUpdateDisplay(id, dateValue) {
  const snickElement = document.querySelector(`[data-last-update-id="${id}"]`);
  if (snickElement) {
    const span = snickElement.querySelector('span');
    if (span) {
      span.textContent = `Last update: ${formatDateTime(dateValue)}`;
    }
  }
}

// Atualiza suppliers, barra inferior e dados visíveis após salvar
async function updateInterfaceAfterSave() {
  try {
    // Recarrega suppliers com estatísticas atualizadas
    const oldSuppliersData = suppliersData;
    suppliersData = await window.api.getSuppliersWithStats();
    
    // Verifica se há filtro de busca de supplier ativo
    const supplierSearchInput = document.getElementById('supplierSearchInput');
    const supplierSearchTerm = supplierSearchInput ? supplierSearchInput.value.trim() : '';
    
    if (supplierSearchTerm.length >= 1) {
      // Aplicar filtro na lista de suppliers e manter toolings filtrados
      const normalizedSearch = supplierSearchTerm.toLowerCase().trim();
      const matchingSuppliers = new Set();
      
      // Busca em todos os toolings para saber quais suppliers têm match
      try {
        const allResults = await window.api.searchTooling(supplierSearchTerm);
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
            return ExpirationMetrics.hasCriticalItems(supplier);
          });
        }
        
        // Atualiza a lista de suppliers
        displaySuppliers(filteredSuppliers);
        
        // Se há um supplier selecionado e ele está nos matching, atualiza seus toolings filtrados
        if (selectedSupplier && matchingSuppliers.has(selectedSupplier)) {
          const filteredResults = allResults.filter(item => {
            const itemSupplier = String(item.supplier || '').trim();
            return itemSupplier === selectedSupplier;
          });
          
          // Atualiza apenas os dados sem recriar os cards (preserva cards expandidos)
          toolingData = filteredResults;
          
          // Atualiza campos dos cards existentes sem recriar tudo
          filteredResults.forEach((item, index) => {
            const card = document.getElementById(`card-${index}`);
            if (card && card.getAttribute('data-item-id') === String(item.id)) {
              // Card já existe com o ID correto, apenas atualiza dados se necessário
              // Não precisa fazer nada pois os inputs já estão sincronizados
            }
          });
        }
      } catch (error) {
      }
    } else {
      // Sem filtro: aplica apenas filtro de expiração
      applyExpirationFilter();
    }
    
    // Atualiza a barra inferior com analytics
    const analytics = await window.api.getAnalytics();
    if (!Array.isArray(suppliersData) || suppliersData.length === 0) {
      updateStatusBar(analytics);
    } else {
      syncStatusBarWithSuppliers();
    }
  } catch (error) {
  }
}

// Carrega analytics
async function loadAnalytics() {
  try {
    const analytics = await window.api.getAnalytics();
    
    // Calcula métricas usando a mesma lógica da barra inferior
    let expiredCount = 0;
    let expiringCount = 0;
    
    if (Array.isArray(suppliersData) && suppliersData.length > 0) {
      const summary = suppliersData.reduce((acc, supplier) => {
        // Usa APENAS ExpirationMetrics.fromItems() para calcular métricas
        const metrics = ExpirationMetrics.fromItems(supplier.items || []);
        acc.expired += metrics.expired;
        acc.expiring += metrics.expiring;
        return acc;
      }, { expired: 0, expiring: 0 });
      
      expiredCount = summary.expired;
      expiringCount = summary.expiring;
    } else {
      expiredCount = analytics.expired_total || 0;
      expiringCount = analytics.expiring_two_years || 0;
    }
    
    // Métricas principais
    document.getElementById('totalTooling').textContent = analytics.total || 0;
    document.getElementById('totalExpired').textContent = expiredCount;
    document.getElementById('totalExpiring').textContent = expiringCount;
    document.getElementById('totalSuppliers').textContent = analytics.suppliers || 0;
    
    // Top suppliers
    displayTopSuppliers(suppliersData);
    
    // Steps summary
    await displayStepsSummary();
    
    if (!Array.isArray(suppliersData) || suppliersData.length === 0) {
      updateStatusBar(analytics);
    } else {
      syncStatusBarWithSuppliers();
    }
  } catch (error) {
  }
}

function displayTopSuppliers(suppliers) {
  const tableBody = document.querySelector('#topSuppliersTable tbody');
  if (!tableBody || !suppliers) return;
  
  const sortedSuppliers = [...suppliers]
    .sort((a, b) => ((b.items || []).length) - ((a.items || []).length))
    .slice(0, 15);
  
  tableBody.innerHTML = sortedSuppliers.map(supplier => {
    // Usa APENAS ExpirationMetrics.fromItems() para calcular métricas
    const metrics = ExpirationMetrics.fromItems(supplier.items || []);
    
    return `
      <tr>
        <td>${escapeHtml(supplier.supplier)}</td>
        <td><span class="table-number">${metrics.total}</span></td>
        <td><span class="table-number table-number--expired">${metrics.expired}</span></td>
        <td><span class="table-number table-number--warning">${metrics.expiring}</span></td>
      </tr>
    `;
  }).join('');
}

async function displayStepsSummary() {
  const tableBody = document.querySelector('#stepsSummaryTable tbody');
  if (!tableBody) return;
  
  // Todos os 7 steps fixos
  const allSteps = ['1', '2', '3', '4', '5', '6', '7'];
  
  // Mapeamento de Action (descrição de cada step)
  const stepActions = {
    '1': 'Control Data Update',
    '2': 'Critical Tooling Identification',
    '3': 'Supplier Validation Request',
    '4': 'Critical Tooling Reassessment',
    '5': 'On-Site Technical Analysis',
    '6': 'Technical Confirmation',
    '7': 'Supply Continuity Strategy'
  };
  
  // Mapeamento de Deadline (prazo de cada step)
  const stepDeadlines = {
    '1': 'September',
    '2': 'September',
    '3': 'October - January',
    '4': 'October - January',
    '5': 'December - March',
    '6': 'December - March',
    '7': 'April - June'
  };
  
  // Mapeamento de Responsible (responsável de cada step)
  const stepResponsibles = {
    '1': 'Supply Continuity',
    '2': 'Supply Continuity',
    '3': 'Supply Continuity',
    '4': 'Supply Continuity',
    '5': 'SQIE',
    '6': 'SQIE',
    '7': 'Sourcing Manager'
  };
  
  // Função para verificar se estamos no prazo
  // Retorna: 'completed' (step deve estar zerado), 'current' (prazo atual), 'upcoming' (futuro)
  function getStepStatus(step, count) {
    const currentMonth = new Date().getMonth(); // 0 = Janeiro, 11 = Dezembro
    
    // Ciclo anual do tooling management:
    // Step 1-2: Julho-Setembro (período de atualização)
    // Step 3-4: Outubro-Janeiro (validação com fornecedor)
    // Step 5-6: Dezembro-Março (análise técnica)  
    // Step 7: Abril-Junho (estratégia de continuidade)
    
    // Para cada step, definir: período ativo e período "atrasado"
    // Mês atual: Novembro (10)
    
    // Usar ordem sequencial no ciclo (Julho = 0 do ciclo, Junho = 11 do ciclo)
    // Converter mês do calendário para mês do ciclo (Julho = mês 6 vira posição 0)
    function toCycleMonth(calendarMonth) {
      // Julho(6) -> 0, Agosto(7) -> 1, ..., Junho(5) -> 11
      return (calendarMonth - 6 + 12) % 12;
    }
    
    const cycleMonth = toCycleMonth(currentMonth);
    
    // Períodos em meses do ciclo (baseado em Julho = 0)
    // Step 1-2: Jul-Set = ciclo 0-2
    // Step 3-4: Out-Jan = ciclo 3-6
    // Step 5-6: Dez-Mar = ciclo 5-8
    // Step 7: Abr-Jun = ciclo 9-11
    const stepCyclePeriods = {
      '1': { start: 0, end: 2 },   // Jul-Set
      '2': { start: 0, end: 2 },   // Jul-Set
      '3': { start: 3, end: 6 },   // Out-Jan
      '4': { start: 3, end: 6 },   // Out-Jan
      '5': { start: 5, end: 8 },   // Dez-Mar
      '6': { start: 5, end: 8 },   // Dez-Mar
      '7': { start: 9, end: 11 }   // Abr-Jun
    };
    
    const period = stepCyclePeriods[step];
    
    const isInPeriod = cycleMonth >= period.start && cycleMonth <= period.end;
    const isPastDeadline = cycleMonth > period.end;
    const isBeforePeriod = cycleMonth < period.start;
    
    // Lógica de status:
    if (isPastDeadline) {
      return count > 0 ? 'behind' : 'completed';
    } else if (isInPeriod) {
      return 'current';
    } else {
      return 'upcoming';
    }
  }
  
  // Função para renderizar o indicador de status
  function renderStatusIndicator(status, count) {
    const statusConfig = {
      'completed': { text: 'DONE', class: 'status-completed' },
      'current': { text: 'NOW', class: 'status-current' },
      'behind': { text: 'LATE', class: 'status-behind' },
      'upcoming': { text: 'ON-GOING', class: 'status-upcoming' }
    };
    
    const config = statusConfig[status] || statusConfig['upcoming'];
    return `<span class="step-status-badge ${config.class}">${config.text}</span>`;
  }
  
  try {
    // Busca dados agregados dos steps
    const stepsData = await window.api.getStepsSummary();
    
    // Cria um mapa para fácil acesso
    const stepsMap = {};
    if (stepsData && stepsData.length > 0) {
      stepsData.forEach(item => {
        stepsMap[item.steps] = item.count;
      });
    }
    
    // Calcula o total para as porcentagens
    const totalWithSteps = allSteps.reduce((sum, step) => sum + (stepsMap[step] || 0), 0);
    
    // Gera array com todos os 7 steps
    const stepsArray = allSteps.map(step => {
      const count = stepsMap[step] || 0;
      return {
        steps: step,
        action: stepActions[step],
        responsible: stepResponsibles[step],
        deadline: stepDeadlines[step],
        count: count,
        percentage: totalWithSteps > 0 ? ((count / totalWithSteps) * 100).toFixed(1) : '0.0',
        status: getStepStatus(step, count)
      };
    });
    
    tableBody.innerHTML = stepsArray.map(item => `
      <tr>
        <td><strong>${escapeHtml(item.steps)}</strong></td>
        <td>${escapeHtml(item.action)}</td>
        <td>${escapeHtml(item.responsible)}</td>
        <td>${escapeHtml(item.deadline)}</td>
        <td><span class="table-number">${item.count}</span></td>
        <td><span class="table-number">${item.percentage}%</span></td>
        <td>${renderStatusIndicator(item.status, item.count)}</td>
      </tr>
    `).join('');
    
  } catch (error) {
    tableBody.innerHTML = '<tr><td colspan="7" style="text-align: center; color: #f44;">Error loading steps data</td></tr>';
  }
}

function syncStatusBarWithSuppliers() {
  if (!Array.isArray(suppliersData) || suppliersData.length === 0) {
    return;
  }

  const summary = suppliersData.reduce((acc, supplier) => {
    // Usa APENAS ExpirationMetrics.fromItems() para calcular métricas
    const metrics = ExpirationMetrics.fromItems(supplier.items || []);

    acc.total += metrics.total;
    acc.expired += metrics.expired;
    acc.expiring += metrics.expiring;
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

    // Trigger animation
    requestAnimationFrame(() => {
      toast.classList.add('show');
    });

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
  const rawTerm = typeof searchTerm === 'string' ? searchTerm : '';
  const normalizedSearch = rawTerm.toLowerCase().trim();
  const requestId = ++supplierFilterRequestId;
  
  if (!normalizedSearch) {
    applyExpirationFilter();
    if (selectedSupplier) {
      await loadToolingBySupplier(selectedSupplier);
    }
    return;
  }

  const matchingSuppliers = new Set();
  
  try {
    const allResults = await window.api.searchTooling(rawTerm);
    if (requestId !== supplierFilterRequestId) {
      return;
    }
    
    allResults.forEach(item => {
      const supplierName = String(item.supplier || '').trim();
      if (supplierName) {
        matchingSuppliers.add(supplierName);
      }
    });
    
    suppliersData.forEach(supplier => {
      const supplierName = String(supplier.supplier || '').toLowerCase();
      if (supplierName.includes(normalizedSearch)) {
        matchingSuppliers.add(supplier.supplier);
      }
    });
    
    let filteredSuppliers = suppliersData.filter(supplier => 
      matchingSuppliers.has(supplier.supplier)
    );
    
    if (expirationFilterEnabled) {
      filteredSuppliers = filteredSuppliers.filter(supplier => {
        return ExpirationMetrics.hasCriticalItems(supplier);
      });
    }
    
    if (requestId !== supplierFilterRequestId) {
      return;
    }

    displaySuppliers(filteredSuppliers);
    
    if (selectedSupplier && matchingSuppliers.has(selectedSupplier)) {
      const filteredResults = allResults.filter(item => {
        const itemSupplier = String(item.supplier || '').trim();
        return itemSupplier === selectedSupplier;
      });
      
      toolingData = filteredResults;
      displayTooling(filteredResults);
    }
  } catch (error) {
  }
}

// Função para expandir/colapsar filtro de steps
function toggleStepsFilter() {
  const dropdown = document.getElementById('stepsFilterDropdown');
  const toggle = document.getElementById('stepsFilterToggle');
  
  if (!dropdown || !toggle) return;
  
  if (dropdown.style.display === 'none') {
    dropdown.style.display = 'block';
    toggle.classList.add('expanded');
  } else {
    dropdown.style.display = 'none';
    toggle.classList.remove('expanded');
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
  supplierSearchDebouncedHandler?.cancel?.();
  
  filterSuppliersAndTooling('');
}

// Status bar search functions
function toggleStatusSearch() {
  const wrapper = document.getElementById('statusSearchInputWrapper');
  const input = document.getElementById('statusSupplierSearchInput');
  const btn = document.getElementById('statusSearchBtn');
  
  if (!wrapper) return;
  
  const isActive = wrapper.classList.contains('active');
  
  if (isActive) {
    wrapper.classList.remove('active');
    input.value = '';
    filterSuppliersAndTooling('');
  } else {
    wrapper.classList.add('active');
    setTimeout(() => input.focus(), 100);
  }
}

function clearStatusSearch() {
  const input = document.getElementById('statusSupplierSearchInput');
  const wrapper = document.getElementById('statusSearchInputWrapper');
  
  if (input) {
    input.value = '';
  }
  
  supplierSearchDebouncedHandler?.cancel?.();
  filterSuppliersAndTooling('');
  
  if (wrapper) {
    wrapper.classList.remove('active');
  }
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
    
  } catch (error) {
  }
}

async function addTodoItem(toolingId) {
  try {
    await window.api.addTodo(toolingId, '');
    await loadTodos(toolingId);
  } catch (error) {
    showNotification('Error adding todo', 'error');
  }
}

async function toggleTodo(todoId, completed) {
  try {
    const todoItem = document.querySelector(`[data-todo-id="${todoId}"]`);
    const todoText = todoItem ? todoItem.querySelector('.todo-text').value : '';
    await window.api.updateTodo(todoId, todoText, completed ? 1 : 0);
  } catch (error) {
  }
}

async function updateTodoText(todoId, text) {
  try {
    const todoItem = document.querySelector(`[data-todo-id="${todoId}"]`);
    const checkbox = todoItem ? todoItem.querySelector('.todo-checkbox') : null;
    const completed = checkbox ? (checkbox.checked ? 1 : 0) : 0;
    await window.api.updateTodo(todoId, text, completed);
  } catch (error) {
  }
}

async function deleteTodo(todoId, toolingId) {
  try {
    await window.api.deleteTodo(todoId);
    await loadTodos(toolingId);
  } catch (error) {
    showNotification('Error deleting todo', 'error');
  }
}

// Function removed - todo badge no longer displayed on card header

// ===== DEVELOPER OPTIONS =====

const DEVTOOLS_ENABLED_KEY = 'devToolsEnabled';

function toggleDevTools(enabled) {
  localStorage.setItem(DEVTOOLS_ENABLED_KEY, enabled ? 'true' : 'false');
  
  if (enabled) {
    // Abre o DevTools
    if (window.api && window.api.openDevTools) {
      window.api.openDevTools();
    }
    showNotification('DevTools opened', 'success');
  } else {
    // Fecha o DevTools
    if (window.api && window.api.closeDevTools) {
      window.api.closeDevTools();
    }
    showNotification('DevTools closed', 'success');
  }
}

function initDevToolsSwitch() {
  const devToolsSwitch = document.getElementById('devToolsSwitch');
  if (devToolsSwitch) {
    const enabled = localStorage.getItem(DEVTOOLS_ENABLED_KEY) === 'true';
    devToolsSwitch.checked = enabled;
    
    // Se estiver habilitado, abre o DevTools ao iniciar
    if (enabled && window.api && window.api.openDevTools) {
      window.api.openDevTools();
    }
  }
}

// ===== SETTINGS CAROUSEL NAVIGATION =====

let currentSettingsCarouselIndex = 0;

function navigateSettingsCarousel(direction) {
  const carousel = document.getElementById('settingsCarousel');
  const cards = carousel ? carousel.querySelectorAll('.settings-card') : [];
  const totalCards = cards.length;
  
  if (totalCards === 0) return;
  
  // Calcular novo índice
  currentSettingsCarouselIndex += direction;
  
  // Circular: voltar ao início ou fim
  if (currentSettingsCarouselIndex < 0) {
    currentSettingsCarouselIndex = totalCards - 1;
  } else if (currentSettingsCarouselIndex >= totalCards) {
    currentSettingsCarouselIndex = 0;
  }
  
  updateSettingsCarouselPosition();
}

function goToSettingsCarouselSlide(index) {
  const carousel = document.getElementById('settingsCarousel');
  const cards = carousel ? carousel.querySelectorAll('.settings-card') : [];
  
  if (index >= 0 && index < cards.length) {
    currentSettingsCarouselIndex = index;
    updateSettingsCarouselPosition();
  }
}

function updateSettingsCarouselPosition() {
  const carousel = document.getElementById('settingsCarousel');
  const indicators = document.querySelectorAll('.carousel-dot');
  const cards = carousel ? carousel.querySelectorAll('.settings-card') : [];
  
  if (!carousel || cards.length === 0) return;
  
  // Verificar se está em modo mobile (carrossel ativo)
  const isMobile = window.innerWidth <= 1200;
  
  if (isMobile) {
    // Scroll para o card ativo
    if (cards[currentSettingsCarouselIndex]) {
      cards[currentSettingsCarouselIndex].scrollIntoView({
        behavior: 'smooth',
        block: 'nearest',
        inline: 'start'
      });
    }
    
    // Atualizar indicadores
    indicators.forEach((dot, index) => {
      if (index === currentSettingsCarouselIndex) {
        dot.classList.add('active');
      } else {
        dot.classList.remove('active');
      }
    });
  }
}

// Detectar scroll manual no carrossel e atualizar indicadores
function initSettingsCarouselScrollListener() {
  const carousel = document.getElementById('settingsCarousel');
  if (!carousel) return;
  
  let scrollTimeout;
  carousel.addEventListener('scroll', () => {
    const isMobile = window.innerWidth <= 1200;
    if (!isMobile) return;
    
    clearTimeout(scrollTimeout);
    scrollTimeout = setTimeout(() => {
      const cards = carousel.querySelectorAll('.settings-card');
      const scrollLeft = carousel.scrollLeft;
      const cardWidth = cards[0] ? cards[0].offsetWidth : 0;
      
      if (cardWidth > 0) {
        const newIndex = Math.round(scrollLeft / cardWidth);
        if (newIndex !== currentSettingsCarouselIndex && newIndex >= 0 && newIndex < cards.length) {
          currentSettingsCarouselIndex = newIndex;
          const indicators = document.querySelectorAll('.carousel-dot');
          indicators.forEach((dot, index) => {
            if (index === currentSettingsCarouselIndex) {
              dot.classList.add('active');
            } else {
              dot.classList.remove('active');
            }
          });
        }
      }
    }, 100);
  });
  
  // Resetar posição quando sair do modo mobile
  window.addEventListener('resize', () => {
    const isMobile = window.innerWidth <= 1200;
    if (!isMobile) {
      currentSettingsCarouselIndex = 0;
      carousel.scrollLeft = 0;
    }
  });
}

