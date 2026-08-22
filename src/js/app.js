/**
 * Main Application Script (Taskpane) - ABNT Assistant
 * Initializes Office.js integration, manages tabs, forms, and UI event listeners.
 */

import { ABNT_CONSTANTS } from './utils/abntRules.js';
import { wordBridge } from './wordApi.js';
import { GeneralFormatter } from './formatters/generalFormatter.js';
import { SelectionFormatter } from './formatters/selectionFormatter.js';
import { HeadingsFormatter } from './formatters/headingsFormatter.js';
import { ReferencesFormatter } from './formatters/referencesFormatter.js';
import { CleanerFormatter } from './formatters/cleanerFormatter.js';
import { AbntLinter } from './auditor/linter.js';
import { AbntAutoFix } from './auditor/autofix.js';
import { CitationGenerator } from './generators/citations.js';
import { ReferenceGenerator } from './generators/references.js';
import { TemplateGenerator } from './generators/templates.js';

// Application state
const appState = {
  selectedFont: ABNT_CONSTANTS.FONTS.ARIAL,
  activeTab: 'tab-home',
  lastAuditResult: null,
  isDarkTheme: false,
};

// Initialize Office.js lifecycle immediately (before DOMContentLoaded)
if (typeof Office !== 'undefined') {
  Office.initialize = function (reason) {
    console.log('Office.initialize fired successfully');
  };

  if (Office.onReady) {
    Office.onReady((info) => {
      console.log('Office.onReady fired for host:', info.host);
      if (info.host === Office.HostType.Word) {
        updateConnectionStatus(true);
        setTimeout(runInitialAudit, 500);
      } else {
        updateConnectionStatus(false);
        setTimeout(runInitialAudit, 500);
      }
    });
  }
}

/**
 * DOM initialization entry point
 */
function initApp() {
  initUI();
  initTheme();
  initEventListeners();
  updateCitationAndRefPreview();

  // Fallback: run in browser/simulator mode when Office is not present
  if (typeof Office === 'undefined' || !Office.context || !Office.context.requirements) {
    updateConnectionStatus(false);
    setTimeout(runInitialAudit, 300);
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initApp);
} else {
  initApp();
}

/**
 * Updates the connection status indicator at the top of the taskpane
 */
function updateConnectionStatus(isWordReal) {
  const pill = document.getElementById('connectionStatusPill');
  const text = document.getElementById('connectionStatusText');

  if (!pill || !text) return;

  if (isWordReal) {
    pill.className = 'status-pill online';
    text.textContent = 'Microsoft Word Conectado';
    console.log('✅ Conexão com Word estabelecida');
  } else {
    pill.className = 'status-pill mock';
    text.textContent = 'Modo Simulador Navegador';
    console.log('⚠️ Operando em modo simulador (sem Word conectado)');
  }
}

/**
 * Initializes tab navigation
 */
function initUI() {
  const tabButtons = document.querySelectorAll('.tab-btn');
  const tabContents = document.querySelectorAll('.tab-content');

  tabButtons.forEach((btn) => {
    btn.addEventListener('click', () => {
      const targetTabId = btn.getAttribute('data-tab');

      tabButtons.forEach((b) => b.classList.remove('active'));
      tabContents.forEach((c) => c.classList.remove('active'));

      btn.classList.add('active');
      const targetContent = document.getElementById(targetTabId);
      if (targetContent) {
        targetContent.classList.add('active');
      }

      appState.activeTab = targetTabId;

      if (targetTabId === 'tab-auditor') {
        runInitialAudit();
      }
    });
  });
}

/**
 * Light / Dark theme toggle
 */
function initTheme() {
  const toggleBtn = document.getElementById('themeToggleBtn');
  if (!toggleBtn) return;
  const html = document.documentElement;

  const savedTheme = localStorage.getItem('abnt_theme');
  if (savedTheme === 'dark') {
    html.setAttribute('data-theme', 'dark');
    appState.isDarkTheme = true;
  }

  toggleBtn.addEventListener('click', () => {
    appState.isDarkTheme = !appState.isDarkTheme;
    const newTheme = appState.isDarkTheme ? 'dark' : 'light';
    html.setAttribute('data-theme', newTheme);
    localStorage.setItem('abnt_theme', newTheme);
  });
}

/**
 * Registers all button and form event listeners
 */
function initEventListeners() {
  // Font selection
  const fontChoiceArial = document.getElementById('fontChoiceArial');
  const fontChoiceTimes = document.getElementById('fontChoiceTimes');

  if (fontChoiceArial && fontChoiceTimes) {
    fontChoiceArial.addEventListener('click', () => {
      fontChoiceArial.classList.add('selected');
      fontChoiceTimes.classList.remove('selected');
      appState.selectedFont = ABNT_CONSTANTS.FONTS.ARIAL;
    });

    fontChoiceTimes.addEventListener('click', () => {
      fontChoiceTimes.classList.add('selected');
      fontChoiceArial.classList.remove('selected');
      appState.selectedFont = ABNT_CONSTANTS.FONTS.TIMES;
    });
  }

  // Full document formatter
  const btnFormatFullDoc = document.getElementById('btnFormatFullDoc');
  if (btnFormatFullDoc) {
    btnFormatFullDoc.addEventListener('click', async () => {
      btnFormatFullDoc.disabled = true;
      btnFormatFullDoc.innerHTML = '<span>⏳ Formatando Documento...</span>';

      try {
        const result = await GeneralFormatter.formatFullDocument({ fontName: appState.selectedFont });
        if (result.success) {
          runInitialAudit();
        } else {
          showToast('Erro ao formatar: ' + (result.error || 'Erro desconhecido'), 'error');
          console.error('Format error:', result);
        }
      } catch (e) {
        showToast('Erro inesperado: ' + e.message, 'error');
        console.error('Unexpected error:', e);
      } finally {
        btnFormatFullDoc.disabled = false;
        btnFormatFullDoc.innerHTML = '<span class="icon-sparkle">✨</span><span>Formatar Documento Inteiro (ABNT)</span>';
      }
    });
  }

  // Margins-only button
  const btnMarginsOnly = document.getElementById('btnMarginsOnly');
  if (btnMarginsOnly) {
    btnMarginsOnly.addEventListener('click', async () => {
      const res = await GeneralFormatter.applyMarginsOnly();
      if (res.success) {
        runInitialAudit();
      } else {
        showToast(res.error, 'error');
      }
    });
  }

  // Quick style buttons (selection formatter)
  const quickBinds = [
    { id: 'btnStyleBody', fn: () => SelectionFormatter.formatBodyText(appState.selectedFont), msg: 'Estilo Corpo de Texto (12pt, 1.5, 1.25cm) aplicado!' },
    { id: 'btnStyleLongQuote', fn: () => SelectionFormatter.formatLongCitation(appState.selectedFont), msg: 'Citação Longa (Recuo 4cm, 10pt, simples) aplicada!' },
    { id: 'btnStyleHeading1', fn: () => HeadingsFormatter.formatHeadingLevel(1, appState.selectedFont), msg: 'Título 1 (Seção Primária) aplicado!' },
    { id: 'btnStyleHeading2', fn: () => HeadingsFormatter.formatHeadingLevel(2, appState.selectedFont), msg: 'Título 2 (Seção Secundária) aplicado!' },
    { id: 'btnStyleHeading3', fn: () => HeadingsFormatter.formatHeadingLevel(3, appState.selectedFont), msg: 'Título 3 (Seção Terciária) aplicado!' },
    { id: 'btnStyleHeading4', fn: () => HeadingsFormatter.formatHeadingLevel(4, appState.selectedFont), msg: 'Título 4 (Seção Quaternária) aplicado!' },
    { id: 'btnStyleCaption', fn: () => SelectionFormatter.formatCaption(appState.selectedFont), msg: 'Legenda (10pt, centralizado) aplicada!' },
    { id: 'btnStyleFootnote', fn: () => SelectionFormatter.formatFootnoteOrSource(appState.selectedFont), msg: 'Fonte / Rodapé (10pt, simples) aplicada!' },
    { id: 'btnStyleReference', fn: () => SelectionFormatter.formatReferenceItem(appState.selectedFont), msg: 'Estilo de Referência (12pt, simples) aplicado!' },
  ];

  quickBinds.forEach(({ id, fn, msg }) => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('click', async () => {
        try {
          const result = await fn();
          if (result && result.success === false) {
            showToast('Erro ao aplicar estilo: ' + (result.error || 'Erro desconhecido'), 'error');
            console.error('Style error:', result);
          }
        } catch (error) {
          showToast('Erro inesperado: ' + error.message, 'error');
          console.error('Unexpected style error:', error);
        }
      });
    }
  });

  const btnFixHeadingNumbers = document.getElementById('btnFixHeadingNumbers');
  if (btnFixHeadingNumbers) {
    btnFixHeadingNumbers.addEventListener('click', async () => {
      const res = await AbntAutoFix.executeFix('fix_heading_dots');
      if (!res.success) {
        showToast(res.message || 'Erro ao ajustar numeração', 'error');
      }
    });
  }

  // Auditor panel
  const btnFixAllIssues = document.getElementById('btnFixAllIssues');
  if (btnFixAllIssues) {
    btnFixAllIssues.addEventListener('click', async () => {
      btnFixAllIssues.disabled = true;
      btnFixAllIssues.innerHTML = '<span>⏳ Corrigindo...</span>';
      try {
        const result = await AbntAutoFix.executeFix('fix_all', { fontName: appState.selectedFont });
        if (result.success) {
          runInitialAudit();
        } else {
          showToast('Erro ao corrigir: ' + (result.message || 'Erro desconhecido'), 'error');
        }
      } catch (error) {
        showToast('Erro inesperado: ' + error.message, 'error');
      } finally {
        btnFixAllIssues.disabled = false;
        btnFixAllIssues.innerHTML = '<span>⚡ Corrigir Todos os Erros</span>';
      }
    });
  }

  const btnRescanAudit = document.getElementById('btnRescanAudit');
  if (btnRescanAudit) btnRescanAudit.addEventListener('click', runInitialAudit);

  const refreshDocBtn = document.getElementById('refreshDocBtn');
  if (refreshDocBtn) refreshDocBtn.addEventListener('click', runInitialAudit);

  // Citations & References
  const refTypeSelect = document.getElementById('refTypeSelect');
  if (refTypeSelect) {
    refTypeSelect.addEventListener('change', () => {
      updateFormVisibility(refTypeSelect.value);
      updateCitationAndRefPreview();
    });
  }

  const inputsToWatch = [
    'inputAuthors', 'inputTitle', 'inputSubtitle', 'inputYear', 'inputPage',
    'inputCity', 'inputPublisher', 'inputJournal', 'inputVolume', 'inputNumber',
    'inputPagesRange', 'inputUrl', 'inputAccessDate'
  ];

  inputsToWatch.forEach((id) => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', updateCitationAndRefPreview);
    }
  });

  const btnInsertCitation = document.getElementById('btnInsertCitation');
  if (btnInsertCitation) {
    btnInsertCitation.addEventListener('click', async () => {
      try {
        const citationText = document.getElementById('previewCitationText')?.textContent || '';
        await wordBridge.insertText(citationText, 'Selection');
      } catch (error) {
        showToast('Erro ao inserir citação: ' + error.message, 'error');
      }
    });
  }

  const btnInsertReference = document.getElementById('btnInsertReference');
  if (btnInsertReference) {
    btnInsertReference.addEventListener('click', async () => {
      try {
        const refText = document.getElementById('previewReferenceText')?.innerText || '';
        await wordBridge.insertText(refText, 'Selection');
      } catch (error) {
        showToast('Erro ao inserir referência: ' + error.message, 'error');
      }
    });
  }

  const btnSortReferences = document.getElementById('btnSortReferences');
  if (btnSortReferences) {
    btnSortReferences.addEventListener('click', async () => {
      const res = await ReferencesFormatter.sortSelectedReferences();
      if (!res.success) {
        showToast(res.message || 'Falha ao ordenar referências.', 'error');
      }
    });
  }

  // Templates
  const templateBtns = document.querySelectorAll('.template-insert-btn');
  templateBtns.forEach((btn) => {
    btn.addEventListener('click', async () => {
      try {
        const templateType = btn.getAttribute('data-template');
        await TemplateGenerator.insertTemplateIntoDocument(templateType);
      } catch (error) {
        showToast('Erro ao inserir modelo: ' + error.message, 'error');
      }
    });
  });

  // Document cleaner
  const btnCleanSpaces = document.getElementById('btnCleanSpaces');
  if (btnCleanSpaces) {
    btnCleanSpaces.addEventListener('click', async () => {
      const res = await CleanerFormatter.cleanDocument();
      if (!res.success) {
        showToast(res.message || 'Erro na higienização', 'error');
      }
      runInitialAudit();
    });
  }

  const btnFixHeadingDotsCleaner = document.getElementById('btnFixHeadingDotsCleaner');
  if (btnFixHeadingDotsCleaner) {
    btnFixHeadingDotsCleaner.addEventListener('click', async () => {
      const res = await AbntAutoFix.executeFix('fix_heading_dots');
      if (!res.success) {
        showToast(res.message || 'Erro ao padronizar títulos', 'error');
      }
      runInitialAudit();
    });
  }
}

/**
 * Shows/hides reference type-specific form fields
 * @param {string} type - Reference type ('book', 'article', 'academic', 'website', 'law')
 */
function updateFormVisibility(type) {
  const bookExtra = document.getElementById('bookExtraFields');
  const articleExtra = document.getElementById('articleExtraFields');
  const websiteExtra = document.getElementById('websiteExtraFields');

  if (bookExtra) bookExtra.style.display = (type === 'book' || type === 'academic') ? 'grid' : 'none';
  if (articleExtra) articleExtra.style.display = type === 'article' ? 'block' : 'none';
  if (websiteExtra) websiteExtra.style.display = type === 'website' ? 'block' : 'none';
}

/**
 * Updates citation and reference preview boxes based on form values
 */
function updateCitationAndRefPreview() {
  const refTypeSelect = document.getElementById('refTypeSelect');
  if (!refTypeSelect) return;

  const type = refTypeSelect.value;
  const authors = document.getElementById('inputAuthors')?.value || 'SILVA, João';
  const title = document.getElementById('inputTitle')?.value || 'Metodologia Científica';
  const subtitle = document.getElementById('inputSubtitle')?.value || '';
  const year = document.getElementById('inputYear')?.value || '2024';
  const page = document.getElementById('inputPage')?.value || '45';

  const parenthetical = CitationGenerator.generateParenthetical(authors, year, page, true);
  const inText = CitationGenerator.generateInText(authors, year, page);

  const previewCitationText = document.getElementById('previewCitationText');
  const previewCitationInText = document.getElementById('previewCitationInText');
  if (previewCitationText) previewCitationText.textContent = parenthetical;
  if (previewCitationInText) previewCitationInText.textContent = inText;

  let refResult = '';
  if (type === 'book') {
    const city = document.getElementById('inputCity')?.value || 'São Paulo';
    const publisher = document.getElementById('inputPublisher')?.value || 'Atlas';
    refResult = ReferenceGenerator.book({ authors, title, subtitle, year, city, publisher });
  } else if (type === 'article') {
    const journalName = document.getElementById('inputJournal')?.value || 'Revista';
    const volume = document.getElementById('inputVolume')?.value || '';
    const number = document.getElementById('inputNumber')?.value || '';
    const pages = document.getElementById('inputPagesRange')?.value || '';
    refResult = ReferenceGenerator.journalArticle({ authors, articleTitle: title, journalName, volume, number, pages, year });
  } else if (type === 'academic') {
    refResult = ReferenceGenerator.academicWork({ authors, title, subtitle, year });
  } else if (type === 'website') {
    const url = document.getElementById('inputUrl')?.value || 'https://exemplo.com';
    const accessDate = document.getElementById('inputAccessDate')?.value || '20 ago. 2024';
    refResult = ReferenceGenerator.website({ authors, title, year, url, accessDate });
  } else if (type === 'law') {
    refResult = ReferenceGenerator.legislation({ lawName: title, summary: subtitle, date: year });
  }

  const previewReferenceText = document.getElementById('previewReferenceText');
  if (previewReferenceText) {
    const formattedHtml = refResult.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    previewReferenceText.innerHTML = formattedHtml;
  }
}

/**
 * Runs the ABNT compliance audit on the current document
 */
async function runInitialAudit() {
  try {
    console.log('🔄 Iniciando verificação do documento...');
    const audit = await AbntLinter.auditDocument();
    if (audit) {
      appState.lastAuditResult = audit;
      renderAuditResults(audit);
      console.log('📋 Resultados da auditoria atualizados na interface');
    }
  } catch (error) {
    console.error('❌ Erro ao executar auditoria:', error);
    showToast('Erro ao analisar documento: ' + error.message, 'error');
  }
}

/**
 * Renders the audit results into the Auditor panel
 * @param {Object} audit - Audit result from AbntLinter.auditDocument()
 */
function renderAuditResults(audit) {
  const scoreVal = document.getElementById('auditScoreVal');
  const scoreCircle = document.getElementById('auditScoreCircle');
  const scoreStatus = document.getElementById('auditScoreStatus');
  const scoreDesc = document.getElementById('auditScoreDesc');

  if (!scoreVal || !scoreCircle) return;

  scoreVal.textContent = `${audit.score}%`;
  scoreCircle.className = `score-circle ${audit.status}`;

  if (audit.status === 'excellent') {
    if (scoreStatus) scoreStatus.textContent = 'Conformidade Excelente';
    if (scoreDesc) scoreDesc.textContent = 'Seu documento atende perfeitamente às normas ABNT!';
  } else if (audit.status === 'good') {
    if (scoreStatus) scoreStatus.textContent = 'Conformidade Boa';
    if (scoreDesc) scoreDesc.textContent = 'Quase tudo perfeito. Pequenos ajustes sugeridos.';
  } else if (audit.status === 'warning') {
    if (scoreStatus) scoreStatus.textContent = 'Atenção Necessária';
    if (scoreDesc) scoreDesc.textContent = 'Foram detectadas divergências de margens ou tipografia.';
  } else {
    if (scoreStatus) scoreStatus.textContent = 'Não Conforme';
    if (scoreDesc) scoreDesc.textContent = 'Múltiplas violações das normas ABNT encontradas.';
  }

  const setMetric = (id, isOk, textOk, textIssue) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.className = isOk ? 'metric-badge ok' : 'metric-badge issue';
    el.textContent = isOk ? `✓ ${textOk}` : `⚠️ ${textIssue}`;
  };

  if (audit.summary) {
    setMetric('metricMargins', audit.summary.marginsOk, 'Margens: 3,3,2,2', 'Margens fora da norma');
    setMetric('metricFonts', audit.summary.fontsOk, 'Fontes: Padronizadas', 'Fontes: Não padronizadas');
    setMetric('metricSpacing', audit.summary.spacingOk, 'Espaçamento: 1.5 OK', 'Espaçamentos incorretos');
    setMetric('metricIndents', audit.summary.indentsOk, 'Recuo: 1.25 cm OK', 'Recuos fora do padrão');
  }

  const container = document.getElementById('issuesListContainer');
  if (!container) return;
  container.innerHTML = '';

  if (!audit.issues || audit.issues.length === 0) {
    container.innerHTML = `
      <div style="padding: 12px; text-align: center; color: var(--accent-emerald); font-weight: 600;">
        🎉 Parabéns! Nenhum erro de formatação ABNT encontrado.
      </div>
    `;
    return;
  }

  audit.issues.forEach((issue) => {
    const card = document.createElement('div');
    card.className = `issue-card ${issue.severity}`;
    card.innerHTML = `
      <div class="issue-header">
        <span class="issue-title">${issue.title}</span>
        <span class="issue-category">${issue.category}</span>
      </div>
      <div class="issue-desc">${issue.description}</div>
      <button class="issue-action" data-fix="${issue.fixAction}">${issue.fixLabel}</button>
    `;

    const btnFix = card.querySelector('.issue-action');
    btnFix.addEventListener('click', async () => {
      btnFix.disabled = true;
      btnFix.textContent = 'Ajustando...';
      const fixResult = await AbntAutoFix.executeFix(issue.fixAction, { fontName: appState.selectedFont });
      if (!fixResult.success) {
        showToast(fixResult.message || 'Erro ao corrigir item', 'error');
      }
      runInitialAudit();
    });

    container.appendChild(card);
  });
}

/**
 * Displays a toast notification
 * @param {string} message - Message to display
 * @param {'info'|'success'|'error'} type - Toast type
 */
function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  if (!container) return;
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = message;

  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => {
      if (toast.parentNode) {
        toast.parentNode.removeChild(toast);
      }
    }, 300);
  }, 5000);
}
