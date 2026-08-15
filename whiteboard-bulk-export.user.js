// ==UserScript==
// @name         whiteboard-bulk-export
// @namespace    https://github.com/wyh2001/whiteboard-bulk-export
// @version      1.0.0
// @description  Export all your boards from Microsoft Whiteboard through simple UI automation.
// @match        https://whiteboard.cloud.microsoft/*
// @run-at       document-idle
// @grant        none
// @license      MIT
// ==/UserScript==

(function () {
  'use strict';

  const CONFIG = {
    elementTimeoutMs: 15000,
    boardLoadTimeoutMs: 30000,
    listRenderTimeoutMs: 600,
    maxRetries: 3,
  };

  const TAG = '[wb-export]';
  const log = (...args) => console.log(TAG, ...args);
  const warn = (...args) => console.warn(TAG, ...args);
  const error = (...args) => console.error(TAG, ...args);

  const UI_TEXT = Object.freeze({
    formatLabel: 'Export format',
    bothFormats: 'PNG and ZIP',
    pngOnly: 'PNG only',
    zipOnly: 'ZIP only',
    startButton: 'Export all whiteboards',
    runningButton: 'Exporting…',
    readyStatus: 'Ready.',
    preparingStatus: 'Reading the whiteboard list…',
    boardStatus: (current, total) => `Exporting whiteboard ${current} of ${total}…`,
    retryStatus: (current, total, attempt) =>
      `Retrying whiteboard ${current} of ${total} (attempt ${attempt})…`,
    emptyStatus: 'No whiteboards found.',
    finishedStatus: (success, total, failed) =>
      failed ? `Finished: ${success} of ${total} succeeded; ${failed} failed.` : `Finished: ${success} of ${total}.`,
    errorStatus: 'Export stopped because of an error.',
  });

  const PANEL_SELECTOR = '[data-wb-export-ui="1"]';
  const FORMAT_SELECTOR = '[data-wb-export-format="1"]';
  const BUTTON_SELECTOR = '[data-wb-export="1"]';
  const STATUS_SELECTOR = '[data-wb-export-status="1"]';
  const FORMAT_ID = 'whiteboard-bulk-export-format';
  const STATUS_ID = 'whiteboard-bulk-export-status';

  let running = false;
  let exportFormat = 'both';
  let exportStatus = UI_TEXT.readyStatus;

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  function isVisible(el) {
    if (!el || el.nodeType !== 1) return false;
    const s = getComputedStyle(el);
    if (s.visibility === 'hidden' || s.display === 'none' || s.opacity === '0') return false;
    const r = el.getBoundingClientRect();
    return r.width > 0 && r.height > 0;
  }

  async function waitFor(label, getter, { timeout = CONFIG.elementTimeoutMs, visible = true } = {}) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
      const el = getter();
      if (el && (!visible || isVisible(el))) return el;
      await sleep(150);
    }
    throw new Error(`Timed out waiting for ${label} (${timeout}ms)`);
  }

  async function waitUntilGone(element, timeout = CONFIG.elementTimeoutMs) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
      if (!element.isConnected || !isVisible(element)) return true;
      await sleep(200);
    }
    return false;
  }

  function sanitizeBoardTitle(title) {
    const sanitized = title.replace(/[\\/:*?"<>|]+/g, '_').replace(/\s+/g, ' ').trim();
    return Array.from(sanitized).slice(0, 80).join('');
  }

  function boardFilenamePrefix(board) {
    const title = sanitizeBoardTitle(board.title) || 'Untitled';
    return `${title}__${board.uuid}`;
  }

  function visibleModals(root = document) {
    return Array.from(root.querySelectorAll('.ms-Modal.is-open.ms-Dialog')).filter(isVisible);
  }

  function isTemplateDialog(dialog) {
    return Boolean(dialog.querySelector('.templatePreviewButton, .useTemplateButton, .startWithBlankCanvasButton'));
  }

  function isExportSuccessDialog(dialog) {
    const text = dialog.textContent;
    return text.includes('Image exported') || text.includes('Zip file exported');
  }

  const FIND = Object.freeze({
    // Gallery
    boardCards: (root = document) => root.querySelectorAll('.boardPreviewOpenBoardButton'),
    boardCardById: (uuid) => document.querySelector(`.boardPreviewOpenBoardButton[id="${uuid}"]`),
    boardGrid: (root = document) => root.querySelector('.scrollableGrid'),
    returnHomeButton: (root = document) => [
      root.querySelector('#homeconnected'),
      root.querySelector('#errorGoToHomeButton'),
    ].find(isVisible),
    errorPageHomeButton: (root = document) => root.querySelector('#errorGoToHomeButton'),
    // Editor menu
    appMenuButton: (root = document) => [
      root.querySelector('#settingsButton'),
      root.querySelector('#AppBarOverflowMenu'),
    ].find(isVisible),
    // Export panel
    exportMenuItem: (root = document) => Array.from(root.querySelectorAll('[data-icon-name="Export"]'))
      .map((icon) => icon.closest('[role="menuitem"]'))
      .find(isVisible),
    exportPanelTitle: (root = document) => root.querySelector('.exportCallout #genericCalloutTitle'),
    imageButton: (root = document) => root.querySelectorAll('.exportCallout .exportOptions button').item(0),
    zipButton: (root = document) => root.querySelectorAll('.exportCallout .exportOptions button').item(1),
    highResRadio: (root = document) => root.querySelector('.ms-Dialog input[type="radio"][id$="-High"]'),
    exportConfirmButton: (root = document) => {
      const radio = FIND.highResRadio(root);
      return radio?.closest('.ms-Dialog')?.querySelector('button.ms-Button--primary');
    },
    successDialog: (root = document) => visibleModals(root).find(isExportSuccessDialog),
    successDialogButton: (root = document) => root.querySelector('button.ms-Button--default'),
    templateDialog: (root = document) => visibleModals(root).find(isTemplateDialog),
    gotItButton: (root = document) => root.querySelector('#gotItButton'),
  });

  const PATTERN = Object.freeze({
    cardId: /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/,
    galleryPath: /\/me\/whiteboards\/?$/,
    boardPath: /\/me\/whiteboards\/[0-9a-fA-F-]{8,}/,
  });

  async function closeDeprecationDialogIfPresent() {
    const button = FIND.gotItButton();
    if (!isVisible(button)) return;

    log('Closing "being retired" deprecation dialog');
    button.click();
    if (!await waitUntilGone(button)) throw new Error('Deprecation dialog did not close in time');
  }

  async function closeTemplateDialogIfPresent() {
    const dialog = FIND.templateDialog();
    if (!dialog) return;

    const closeButton = dialog.querySelector('.ms-Dialog-button--close');
    if (!isVisible(closeButton)) {
      throw new Error('Template dialog is blocking the editor and has no close button');
    }

    log('Closing template dialog');
    closeButton.click();
    if (!await waitUntilGone(dialog)) throw new Error('Template dialog did not close in time');
  }

  async function waitForElementOrDismissTemplate(label, getter, options) {
    const element = await waitFor(
      `${label} or template dialog`,
      () => FIND.templateDialog() || getter(),
      options
    );
    if (!isTemplateDialog(element)) return element;

    await closeTemplateDialogIfPresent();
    return null;
  }

  // Rename generated download anchors before Whiteboard dispatches their click events.
  const RENAMED_DOWNLOAD = Symbol('whiteboardExportRenamed');
  const downloadState = {
    pendingFilename: null,
    renamedCount: 0,
    anchorPrototype: null,
    hadOwnDispatch: false,
    originalDispatch: null,
    interceptedDispatch: null,
  };

  function installDownloadInterceptor() {
    if (downloadState.originalDispatch) return;

    const anchorPrototype = HTMLAnchorElement.prototype;
    const originalDispatch = anchorPrototype.dispatchEvent;
    const interceptedDispatch = function (event) {
      if (
        event?.type === 'click' &&
        this instanceof HTMLAnchorElement &&
        !this[RENAMED_DOWNLOAD] &&
        downloadState.pendingFilename
      ) {
        const href = this.getAttribute('href');
        if (href?.startsWith('data:') || href?.startsWith('blob:')) {
          this[RENAMED_DOWNLOAD] = true;
          this.download = downloadState.pendingFilename;
          downloadState.renamedCount++;
          log('Download renamed:', this.download);
        }
      }
      return originalDispatch.call(this, event);
    };

    downloadState.anchorPrototype = anchorPrototype;
    downloadState.hadOwnDispatch = Object.prototype.hasOwnProperty.call(anchorPrototype, 'dispatchEvent');
    downloadState.originalDispatch = originalDispatch;
    downloadState.interceptedDispatch = interceptedDispatch;
    anchorPrototype.dispatchEvent = interceptedDispatch;

    log('Download interceptor installed');
  }

  function uninstallDownloadInterceptor() {
    downloadState.pendingFilename = null;
    if (!downloadState.originalDispatch) return;
    const anchorPrototype = downloadState.anchorPrototype;
    if (anchorPrototype.dispatchEvent !== downloadState.interceptedDispatch) {
      warn('Download interceptor was replaced by another script; leaving the chain intact');
      return;
    }

    if (downloadState.hadOwnDispatch) {
      anchorPrototype.dispatchEvent = downloadState.originalDispatch;
    } else {
      delete anchorPrototype.dispatchEvent;
    }
    downloadState.anchorPrototype = null;
    downloadState.hadOwnDispatch = false;
    downloadState.originalDispatch = null;
    downloadState.interceptedDispatch = null;
    log('Download interceptor removed');
  }

  function armDownload(prefix, extension) {
    downloadState.pendingFilename = `${prefix}.${extension}`;
    return downloadState.renamedCount;
  }

  function disarmDownload() {
    downloadState.pendingFilename = null;
  }

  function isGalleryPage() {
    return PATTERN.galleryPath.test(location.pathname) ||
      (location.pathname === '/' && Boolean(FIND.boardGrid()));
  }

  function waitForGalleryLoaded() {
    return waitFor('whiteboard gallery', FIND.boardGrid, {
      timeout: CONFIG.boardLoadTimeoutMs,
      visible: false,
    });
  }

  async function returnToGallery() {
    if (isGalleryPage()) {
      await closeDeprecationDialogIfPresent();
      await waitForGalleryLoaded();
      return;
    }

    await closeTemplateDialogIfPresent();
    const homeButton = await waitFor('Home button', FIND.returnHomeButton, {
      timeout: CONFIG.boardLoadTimeoutMs,
    });
    homeButton.click();
    await waitFor(
      'whiteboard gallery after navigation',
      () => isGalleryPage() ? FIND.boardGrid() : null,
      { timeout: CONFIG.boardLoadTimeoutMs, visible: false }
    );
    await closeDeprecationDialogIfPresent();
  }

  function boardTitleFromCard(button) {
    const caption = button.querySelector('.captionText');
    if (!caption) return 'Untitled';

    // Read the direct text node without nested tooltip text.
    const titleNode = Array.from(caption.childNodes)
      .find((node) => node.nodeType === Node.TEXT_NODE && node.textContent.trim());
    return titleNode ? titleNode.textContent.trim() : 'Untitled';
  }

  function collectBoardsMap() {
    const map = new Map();
    for (const button of FIND.boardCards()) {
      const uuid = button.id.trim();
      if (!PATTERN.cardId.test(uuid)) continue;
      if (map.has(uuid)) continue;
      map.set(uuid, { uuid, title: boardTitleFromCard(button) });
    }
    return map;
  }

  function waitForDomCondition(root, condition, {
    timeout = CONFIG.elementTimeoutMs,
    observerOptions = { childList: true, subtree: true },
  } = {}) {
    if (condition()) return Promise.resolve(true);

    return new Promise((resolve) => {
      const observer = new MutationObserver(check);
      const timeoutId = setTimeout(() => finish(false), timeout);

      function finish(matched) {
        clearTimeout(timeoutId);
        observer.disconnect();
        resolve(matched);
      }

      function check() {
        if (condition()) finish(true);
      }

      observer.observe(root, observerOptions);
      check();
    });
  }

  function waitForFirstBoardCard(grid) {
    return waitForDomCondition(grid, () => FIND.boardCards(grid).length > 0, {
      timeout: CONFIG.boardLoadTimeoutMs,
    });
  }

  function addVirtualListLayoutFix() {
    const style = document.createElement('style');
    style.textContent = '.scrollableGrid .ms-List-page{display:flow-root}';
    document.head.appendChild(style);
    return style;
  }

  function renderedListSignature(grid) {
    return Array.from(
      grid.querySelectorAll('.ms-List-cell[data-list-index]'),
      (cell) => cell.getAttribute('data-list-index')
    ).join(',');
  }

  async function scrollToPosition(grid, top) {
    const maxTop = Math.max(0, grid.scrollHeight - grid.clientHeight);
    const targetTop = Math.max(0, Math.min(top, maxTop));
    if (Math.abs(grid.scrollTop - targetTop) < 1) return false;
    const before = renderedListSignature(grid);
    grid.scrollTop = targetTop;
    await waitForDomCondition(
      grid,
      () => renderedListSignature(grid) !== before,
      { timeout: CONFIG.listRenderTimeoutMs }
    );
    return true;
  }

  function scrollForward(grid) {
    const step = Math.max(grid.clientHeight * 2, 240);
    return scrollToPosition(grid, grid.scrollTop + step);
  }

  async function scanBoardList(visit) {
    const grid = await waitFor('whiteboard gallery', FIND.boardGrid, {
      timeout: CONFIG.elementTimeoutMs,
      visible: false,
    });
    if (!await waitForFirstBoardCard(grid)) return null;
    await scrollToPosition(grid, 0);

    // Board tiles float, so flow-root lets Fluent UI measure each page height.
    const layoutFix = addVirtualListLayoutFix();
    let matched = false;
    try {
      while (true) {
        const result = visit();
        if (result) {
          matched = true;
          return result;
        }
        if (!await scrollForward(grid)) return null;
      }
    } finally {
      if (!matched && grid.isConnected) {
        await scrollToPosition(grid, 0);
      }
      layoutFix.remove();
    }
  }

  async function loadAllBoards() {
    const all = new Map();
    await scanBoardList(() => {
      const current = collectBoardsMap();
      current.forEach((board, uuid) => all.set(uuid, board));
      return false;
    });
    return Array.from(all.values());
  }

  async function openBoard(uuid) {
    const before = location.href;
    log('Opening board:', uuid);
    const opened = await scanBoardList(() => {
      const button = FIND.boardCardById(uuid);
      if (isVisible(button)) {
        button.click();
        return true;
      }
      return false;
    });
    if (!opened) throw new Error(`Card not found: ${uuid}`);

    await waitFor(
      'whiteboard editor route',
      () => PATTERN.boardPath.test(location.href) && location.href !== before ? document.documentElement : null,
      { timeout: CONFIG.boardLoadTimeoutMs }
    );
    let editorState = null;
    while (!editorState) {
      editorState = await waitForElementOrDismissTemplate(
        'app menu button or editor error page',
        () => FIND.appMenuButton() || FIND.errorPageHomeButton(),
        { timeout: CONFIG.boardLoadTimeoutMs }
      );
    }
    if (editorState.id === 'errorGoToHomeButton') {
      throw new Error('Whiteboard opened its "Something went wrong" page');
    }
    log('Board loaded');
  }

  async function openExportPanel() {
    while (true) {
      const menuButton = await waitForElementOrDismissTemplate('app menu button', FIND.appMenuButton);
      if (!menuButton) continue;
      menuButton.click();

      const exportItem = await waitForElementOrDismissTemplate('Export menu item', FIND.exportMenuItem);
      if (!exportItem) continue;
      exportItem.click();

      const exportPanel = await waitForElementOrDismissTemplate('export panel', FIND.exportPanelTitle);
      if (!exportPanel) continue;

      log('Export panel opened');
      return;
    }
  }

  async function exportImage(prefix) {
    log('Exporting image...');
    await openExportPanel();

    const imageButton = await waitFor('image export option', FIND.imageButton);
    imageButton.click();

    const highRadio = await waitFor('high-resolution option', FIND.highResRadio, { visible: false });
    const highLabel = highRadio.labels?.[0];
    if (!highLabel) throw new Error('High-resolution option has no associated label');
    highLabel.click();
    await waitFor(
      'selected high-resolution option',
      () => {
        const radio = FIND.highResRadio();
        return radio?.checked ? radio : null;
      },
      { visible: false }
    );

    const confirmButton = await waitFor('export confirmation button', FIND.exportConfirmButton);
    await completeDownload(prefix, 'png', confirmButton);
    log('Image export done');
  }

  async function exportZip(prefix) {
    log('Exporting ZIP...');
    await openExportPanel();

    const zipButton = await waitFor('ZIP export option', FIND.zipButton);
    await completeDownload(prefix, 'zip', zipButton);
    log('ZIP export done');
  }

  async function completeDownload(prefix, extension, button) {
    const downloadMark = armDownload(prefix, extension);
    try {
      button.click();
      const dialog = await waitForSuccessDialog(downloadMark);
      await closeSuccessDialog(dialog);
    } finally {
      disarmDownload();
    }
  }

  async function waitForSuccessDialog(downloadMark) {
    const dialog = await waitFor(
      'download success dialog',
      () => downloadState.renamedCount > downloadMark ? FIND.successDialog() : null,
      { timeout: CONFIG.boardLoadTimeoutMs }
    );
    log('Success dialog appeared');
    return dialog;
  }

  async function closeSuccessDialog(dialog) {
    const closeButton = await waitFor(
      'success dialog close button',
      () => FIND.successDialogButton(dialog)
    );
    closeButton.click();
    if (!await waitUntilGone(dialog)) throw new Error('Success dialog did not close in time');
    log('Success dialog closed');
  }

  async function exportCurrentBoard(prefix, progress, format) {
    if (format !== 'zip' && !progress.image) {
      await exportImage(prefix);
      progress.image = true;
    }
    if (format !== 'png' && !progress.zip) {
      await exportZip(prefix);
      progress.zip = true;
    }
  }

  async function exportBoardWithRetries(board, index, total, format) {
    const prefix = boardFilenamePrefix(board);
    const progress = { image: false, zip: false };
    setExportStatus(UI_TEXT.boardStatus(index + 1, total));
    log(`=== (${index + 1}/${total}) ${prefix} ===`);

    for (let attempt = 1; attempt <= CONFIG.maxRetries; attempt++) {
      if (attempt > 1) {
        setExportStatus(UI_TEXT.retryStatus(index + 1, total, attempt));
        log(`Retry ${attempt}/${CONFIG.maxRetries}: ${prefix} (image=${progress.image}, zip=${progress.zip})`);
      }

      try {
        await openBoard(board.uuid);
        await exportCurrentBoard(prefix, progress, format);
        log(`Done: ${prefix}`);
        return null;
      } catch (caughtError) {
        error(`Attempt ${attempt}/${CONFIG.maxRetries} failed: ${prefix}`, caughtError);
        if (attempt === CONFIG.maxRetries) break;
        await returnToGallery();
      }
    }

    error(`Failed: ${prefix} (image=${progress.image}, zip=${progress.zip})`);
    return prefix;
  }

  async function runBatchExport(format) {
    await closeDeprecationDialogIfPresent();

    log('Scrolling to load all whiteboards...');
    const boards = await loadAllBoards();
    if (!boards.length) {
      setExportStatus(UI_TEXT.emptyStatus);
      alert(UI_TEXT.emptyStatus);
      return;
    }

    log('Found', boards.length, 'whiteboards');

    installDownloadInterceptor();

    // Opening a board moves it to the front. Reverse traversal preserves the
    // captured gallery order after a complete run.
    const exportOrder = boards.slice().reverse();
    const failed = [];
    for (const [index, board] of exportOrder.entries()) {
      const failedPrefix = await exportBoardWithRetries(board, index, boards.length, format);
      if (failedPrefix) failed.push(failedPrefix);

      await returnToGallery();
    }

    const successCount = boards.length - failed.length;
    setExportStatus(UI_TEXT.finishedStatus(successCount, boards.length, failed.length));
    log(`All done. Success ${successCount}, failed ${failed.length}`);
    if (failed.length) log('Failed:', failed);
    const failureList = failed.length ? `\nFailed:\n${failed.join('\n')}` : '';
    alert(`Batch export finished.\nSuccess ${successCount} / ${boards.length}${failureList}`);
  }

  async function main() {
    if (!isGalleryPage()) {
      alert('Please open the whiteboard gallery first (https://whiteboard.cloud.microsoft/me/whiteboards).');
      return;
    }
    if (!confirm('Start exporting all whiteboards?')) return;

    const format = exportFormat;
    running = true;
    syncExportControls();
    setExportStatus(UI_TEXT.preparingStatus);

    try {
      await runBatchExport(format);
    } catch (caughtError) {
      error('Runtime error', caughtError);
      setExportStatus(UI_TEXT.errorStatus);
      alert(`Runtime error: ${caughtError.message}`);
    } finally {
      uninstallDownloadInterceptor();
      running = false;
      syncExportControls();
    }
  }

  function setExportStatus(message) {
    exportStatus = message;
    const panel = document.querySelector(PANEL_SELECTOR);
    const status = panel?.querySelector(STATUS_SELECTOR);
    if (status) {
      status.textContent = message;
    } else if (running || isGalleryPage()) {
      syncExportControls();
    }
  }

  function createExportControls() {
    const panel = document.createElement('div');
    panel.setAttribute('data-wb-export-ui', '1');
    panel.setAttribute('role', 'region');
    panel.setAttribute('aria-label', 'Whiteboard bulk export');
    panel.lang = 'en';
    panel.style.cssText = [
      'position:fixed', 'z-index:2147483647', 'right:16px', 'bottom:16px',
      'width:216px', 'max-width:calc(100vw - 32px)', 'box-sizing:border-box',
      'padding:12px', 'background:#fff', 'color:#1b1a19',
      'border:1px solid #8a8886', 'border-radius:8px',
      'box-shadow:0 2px 8px rgba(0,0,0,.3)', 'font-family:system-ui,sans-serif',
    ].join(';');

    const formatLabel = document.createElement('label');
    formatLabel.htmlFor = FORMAT_ID;
    formatLabel.textContent = UI_TEXT.formatLabel;
    formatLabel.style.cssText = 'display:block;margin-bottom:4px;font-size:13px;font-weight:600';

    const formatSelect = document.createElement('select');
    formatSelect.id = FORMAT_ID;
    formatSelect.setAttribute('data-wb-export-format', '1');
    formatSelect.style.cssText = [
      'width:100%', 'box-sizing:border-box', 'min-height:44px',
      'margin-bottom:10px', 'padding:8px',
      'background:#fff', 'color:#1b1a19', 'border:1px solid #605e5c',
      'border-radius:4px', 'font:inherit',
    ].join(';');
    for (const [value, text] of [
      ['both', UI_TEXT.bothFormats],
      ['png', UI_TEXT.pngOnly],
      ['zip', UI_TEXT.zipOnly],
    ]) {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = text;
      formatSelect.appendChild(option);
    }
    formatSelect.value = exportFormat;
    formatSelect.addEventListener('change', () => {
      exportFormat = formatSelect.value;
    });

    const button = document.createElement('button');
    button.type = 'button';
    button.setAttribute('data-wb-export', '1');
    button.setAttribute('aria-describedby', STATUS_ID);
    button.style.cssText = [
      'width:100%', 'box-sizing:border-box', 'min-height:44px',
      'padding:10px 16px', 'color:#fff',
      'border:1px solid currentColor', 'border-radius:6px', 'font-size:14px',
      'font-weight:600', 'box-shadow:0 2px 4px rgba(0,0,0,.2)',
      'font-family:inherit', 'outline:none',
    ].join(';');
    button.addEventListener('focus', () => {
      button.style.outline = '3px solid CanvasText';
      button.style.outlineOffset = '3px';
    });
    button.addEventListener('blur', () => {
      button.style.outline = 'none';
      button.style.outlineOffset = '0';
    });
    button.addEventListener('click', () => {
      if (!running) void main();
    });

    const status = document.createElement('div');
    status.id = STATUS_ID;
    status.setAttribute('data-wb-export-status', '1');
    status.setAttribute('role', 'status');
    status.setAttribute('aria-live', 'polite');
    status.setAttribute('aria-atomic', 'true');
    status.style.cssText = 'margin-top:8px;font-size:13px;line-height:1.4';

    panel.append(formatLabel, formatSelect, button, status);
    document.body.appendChild(panel);
    log('Export controls injected');
    return panel;
  }

  function syncExportControls() {
    let panel = document.querySelector(PANEL_SELECTOR);
    if (!isGalleryPage() && !running) {
      panel?.remove();
      return;
    }

    let formatSelect = panel?.querySelector(FORMAT_SELECTOR);
    if (
      panel &&
      (!panel.querySelector(BUTTON_SELECTOR) ||
        !formatSelect?.labels.length ||
        !panel.querySelector(STATUS_SELECTOR))
    ) {
      panel.remove();
      panel = null;
    }
    if (!panel) panel = createExportControls();
    formatSelect = panel.querySelector(FORMAT_SELECTOR);
    const button = panel.querySelector(BUTTON_SELECTOR);
    const status = panel.querySelector(STATUS_SELECTOR);
    formatSelect.value = exportFormat;
    formatSelect.disabled = running;
    button.textContent = running ? UI_TEXT.runningButton : UI_TEXT.startButton;
    // Keep the focused button in the tab order while exposing its busy state.
    button.setAttribute('aria-disabled', String(running));
    button.setAttribute('aria-busy', String(running));
    button.style.background = running ? '#605e5c' : '#4f6bed';
    button.style.cursor = running ? 'progress' : 'pointer';
    status.textContent = exportStatus;
  }

  syncExportControls();
  setInterval(syncExportControls, 2000);
})();
