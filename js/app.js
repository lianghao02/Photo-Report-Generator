/* 現場照片清冊生成工具：主程式（App 類別、工作台互動、匯出流程串接）。
 * 由 index.html 的內嵌 <script> 原封不動搬出（連縮排都保留，避免改變跨行樣板字串的內容），
 * 以 <script src> 放在原本的位置，執行順序不變：在 js/validation.js 等模組與三個匯出模組之後載入。
 * 單元測試 tests/unit/app-bridge.js 會從 index.html 與本檔擷取方法；Tailwind 也會掃描本檔的類別名稱。 */

        document.getElementById('year').textContent = new Date().getFullYear();

        // 檢查 docx 函式庫是否已掛載至全域
        function getDocxLib() {
            if (typeof docx !== 'undefined') return docx;
            if (window.docx) return window.docx;
            return null;
        }

        function escapeHtml(value) {
            return String(value ?? '').replace(/[&<>"']/g, char => ({
                '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
            })[char]);
        }

        class PhotoReportApp {
            constructor() {
                this.photos = [];
                this.currentView = 'workspace';
                this.currentIndex = 0;
                this.draggedIndices = [];
                this.stageDrag = null;
                this.pointerCardDrag = null;
                this.suppressCardClickUntil = 0;
                this.lastSelectedIndex = null;
                this.maxRecommendedPhotos = 500;
                this.maxRecommendedBytes = 2 * 1024 * 1024 * 1024;
                this.thumbnailSize = Math.max(90, Math.min(480, this.loadWorkspacePreference('thumbnailSize', 150)));
                this.editorCollapsed = this.loadWorkspacePreference('editorCollapsed', false);
                this.lightboxEditorCollapsed = this.loadWorkspacePreference('lightboxEditorCollapsed', false);
                this.sidePanelWidth = Math.max(280, Math.min(560, this.loadWorkspacePreference('sidePanelWidth', 300)));
                this.panelResizeState = null;
                this.marqueeState = null;
                this.history = [];
                this.historyIndex = -1;
                this.historyLimit = 50;
                this.historyTimer = null;
                this.pendingHistoryLabel = '';
                this.isRestoringHistory = false;
                this.managedUrls = new Set();
                this.processingCancelled = false;
                this.lightboxIndex = null;
                this.savedProjectSignature = '';
                this.savedPhotosByUid = new Map();
                this.currentProjectFileHandle = null;
                this.currentProjectFileName = '';
                this.currentProjectSavedAt = null;
                this.lastLocalEditAt = null;
                this.initialSaveAs = typeof window !== 'undefined' && typeof window.saveAs === 'function' ? window.saveAs : null;
                this.autoDraftDbName = 'PhotoReportAutoDraftDB';
                this.autoDraftStoreName = 'drafts';
                this.autoDraftTimer = null;
                this.lastAutoDraftSavedAt = null;
                this.pendingAutoDraft = null;
                this.hadPhotosInSession = false;
                this.maxAutoDraftPhotos = 120;
                this.maxAutoDraftBytes = 200 * 1024 * 1024;
                this.activeFilter = 'all';
                this.pendingExportAction = null;
                this.pendingSequenceExport = null;
                this.textLimitInputStates = new WeakMap();
                this.textLimitFeedbackTimers = new WeakMap();
                this.initElements();
                this.initEvents();
                this.applyWorkspaceLayout();
                this.setDefaultTodayDate();
                this.updateLayoutDescription();
                this.loadAppVersion();
                this.render();
                this.refreshAllTextLimitUi();
                this.recordHistory('初始狀態');
                this.markProjectSaved();
                this.checkAndOfferAutoDraft();
            }

            initElements() {
                this.fileInput = document.getElementById('fileInput');
                this.folderInput = document.getElementById('folderInput');
                this.workspaceContainer = document.getElementById('workspaceContainer');
                this.appMainLayout = document.getElementById('appMainLayout');
                this.mainPanelSplitter = document.getElementById('mainPanelSplitter');
                this.thumbGridDropZone = document.getElementById('thumbGridDropZone');
                this.thumbGrid = document.getElementById('thumbGrid');
                this.photoPanel = document.getElementById('photoPanel');
                this.editorPanel = document.getElementById('editorPanel');
                this.workspaceLayoutControls = document.getElementById('workspaceLayoutControls');
                this.thumbnailSizeLabel = document.getElementById('thumbnailSizeLabel');
                this.thumbnailSizeRange = document.getElementById('thumbnailSizeRange');
                this.selectionMarquee = document.getElementById('selectionMarquee');
                this.btnToggleEditor = document.getElementById('btnToggleEditor');
                this.gridContainer = document.getElementById('gridContainer');
                this.gridTbody = document.getElementById('gridTbody');
                this.emptyState = document.getElementById('emptyState');
                this.photoCount = document.getElementById('photoCount');
                this.pageEstimate = document.getElementById('pageEstimate');
                this.selectedCountBadge = document.getElementById('selectedCountBadge');
                this.layoutSelect = document.getElementById('layoutSelect');
                this.layoutDescription = document.getElementById('layoutDescription');
                this.layoutMiniPreview = document.getElementById('layoutMiniPreview');
                this.shortcutsModal = document.getElementById('shortcutsModal');
                this.processingStatus = document.getElementById('processingStatus');
                this.processingMessage = document.getElementById('processingMessage');
                this.processingCount = document.getElementById('processingCount');
                this.processingBar = document.getElementById('processingBar');
                this.btnUndo = document.getElementById('btnUndo');
                this.btnRedo = document.getElementById('btnRedo');
                this.btnCancelProcessing = document.getElementById('btnCancelProcessing');
                this.appVersionText = document.getElementById('appVersionText');
                this.appVersionBadge = document.getElementById('appVersionBadge');
                this.btnCheckUpdate = document.getElementById('btnCheckUpdate');
                this.checkUpdateIcon = document.getElementById('checkUpdateIcon');
                this.checkUpdateText = document.getElementById('checkUpdateText');
                this.btnQuickSaveStatus = document.getElementById('btnQuickSaveStatus');
                this.quickSaveStatusIcon = document.getElementById('quickSaveStatusIcon');
                this.quickSaveStatusText = document.getElementById('quickSaveStatusText');
                this.saveProjectMenuText = document.getElementById('saveProjectMenuText');
                this.btnSaveProjectAs = document.getElementById('btnSaveProjectAs');
                this.mergeProjectInput = document.getElementById('mergeProjectInput');
                this.autoDraftBanner = document.getElementById('autoDraftBanner');
                this.autoDraftBannerTitle = document.getElementById('autoDraftBannerTitle');
                this.autoDraftBannerMeta = document.getElementById('autoDraftBannerMeta');
                this.btnRestoreAutoDraft = document.getElementById('btnRestoreAutoDraft');
                this.btnDiscardAutoDraft = document.getElementById('btnDiscardAutoDraft');

                // 右側編輯器元件
                this.editorCard = document.getElementById('editorCard');
                this.editorTitle = document.getElementById('editorTitle');
                this.editorModeBadge = document.getElementById('editorModeBadge');
                this.editSeq = document.getElementById('editSeq');
                this.editSeqLockNotice = document.getElementById('editSeqLockNotice');
                this.editSeqLockText = document.getElementById('editSeqLockText');
                this.btnJumpToParentPhoto = document.getElementById('btnJumpToParentPhoto');
                this.btnUnlinkChildPhoto = document.getElementById('btnUnlinkChildPhoto');
                this.editDate = document.getElementById('editDate');
                this.editTime = document.getElementById('editTime');
                this.defaultLocationInput = document.getElementById('defaultLocation');
                this.defaultLocationCounter = document.getElementById('defaultLocationCounter');
                this.defaultLocationLimitHint = document.getElementById('defaultLocationLimitHint');
                this.editLocation = document.getElementById('editLocation');
                this.editLocationCounter = document.getElementById('editLocationCounter');
                this.editLocationLimitHint = document.getElementById('editLocationLimitHint');
                this.editDesc = document.getElementById('editDesc');
                this.editDescCounter = document.getElementById('editDescCounter');
                this.editDescLimitHint = document.getElementById('editDescLimitHint');
                this.editDateInheritedHint = document.getElementById('editDateInheritedHint');
                this.editLocationInheritedHint = document.getElementById('editLocationInheritedHint');
                this.batchApplyBox = document.getElementById('batchApplyBox');
                this.batchCountNum = document.getElementById('batchCountNum');
                this.batchApplyEmptyValues = document.getElementById('batchApplyEmptyValues');

                // 完整度稽核篩選列
                this.auditFilterBar = document.getElementById('auditFilterBar');
                this.btnToggleAudit = document.getElementById('btnToggleAudit');
                this.auditPanelOpen = false;
                this.auditBadge_all = document.getElementById('auditBadge_all');
                this.auditBadge_missingLocation = document.getElementById('auditBadge_missingLocation');
                this.auditBadge_missingDesc = document.getElementById('auditBadge_missingDesc');
                this.auditBadge_invalidDateTime = document.getElementById('auditBadge_invalidDateTime');
                this.auditBadge_duplicatePhotos = document.getElementById('auditBadge_duplicatePhotos');
                this.auditBadge_overTextLimit = document.getElementById('auditBadge_overTextLimit');
                this.filterNoticeText = document.getElementById('filterNoticeText');

                // 匯出前非阻斷提醒彈窗
                this.exportAuditModal = document.getElementById('exportAuditModal');
                this.exportAuditModalTitle = document.getElementById('exportAuditModalTitle');
                this.exportAuditList = document.getElementById('exportAuditList');
                this.btnAuditViewIssues = document.getElementById('btnAuditViewIssues');
                this.btnAuditProceedExport = document.getElementById('btnAuditProceedExport');

                // 大圖燈箱元件 (鑑識級規格)
                this.imageLightboxModal = document.getElementById('imageLightboxModal');
                this.lightboxViewport = document.getElementById('lightboxViewport');
                this.lightboxStage = document.getElementById('lightboxStage');
                this.lightboxImg = document.getElementById('lightboxImg');
                this.lightboxCaption = document.getElementById('lightboxCaption');
                this.lightboxPosition = document.getElementById('lightboxPosition');
                this.lightboxZoomBadge = document.getElementById('lightboxZoomBadge');
                this.btnLightboxResetZoom = document.getElementById('btnLightboxResetZoom');
                this.btnLightboxRotateCCW = document.getElementById('btnLightboxRotateCCW');
                this.btnLightboxRotateCW = document.getElementById('btnLightboxRotateCW');
                this.btnLightboxPrev = document.getElementById('btnLightboxPrev');
                this.btnLightboxNext = document.getElementById('btnLightboxNext');
                this.lightboxEditorMount = document.getElementById('lightboxEditorMount');
                this.btnLightboxToggleEditor = document.getElementById('btnLightboxToggleEditor');
                this.lightboxToggleEditorText = document.getElementById('lightboxToggleEditorText');
                this.btnLightboxFlipH = document.getElementById('btnLightboxFlipH');
                this.lightboxAnnotationToolbar = document.getElementById('lightboxAnnotationToolbar');
                this.btnLightboxToolPan = document.getElementById('btnLightboxToolPan');
                this.btnLightboxToolRectYellow = document.getElementById('btnLightboxToolRectYellow');
                this.btnLightboxToolRectRed = document.getElementById('btnLightboxToolRectRed');
                this.btnLightboxToolLabel = document.getElementById('btnLightboxToolLabel');
                this.btnLightboxToolCrop = document.getElementById('btnLightboxToolCrop');
                this.lightboxLabelTextInput = document.getElementById('lightboxLabelTextInput');
                this.lightboxLabelColorSelect = document.getElementById('lightboxLabelColorSelect');
                this.lightboxAnnotationCountBadge = document.getElementById('lightboxAnnotationCountBadge');
                this.btnLightboxClearAnnotations = document.getElementById('btnLightboxClearAnnotations');
                this.lightboxAnnotationLayer = document.getElementById('lightboxAnnotationLayer');
                this.lightboxCanvasFrame = document.getElementById('lightboxCanvasFrame');
                this.lightboxLeftSlot = document.getElementById('lightboxLeftSlot');
                this.lightboxSideGap = document.getElementById('lightboxSideGap');
                this.lightboxSideCropArea = document.getElementById('lightboxSideCropArea');
                this.lightboxSideCropCard = document.getElementById('lightboxSideCropCard');
                this.lightboxSideCropCanvas = document.getElementById('lightboxSideCropCanvas');
                this.btnSideCropBorder = document.getElementById('btnSideCropBorder');
                this.lightboxSideLayoutBar = document.getElementById('lightboxSideLayoutBar');
                this.lightboxSideLayoutBtns = Array.from(document.querySelectorAll('[data-lightbox-layout]'));
                this.lightboxSideLayoutStatus = document.getElementById('lightboxSideLayoutStatus');
                this.lightboxBaseCanvasCache = null;
                this.lightboxBaseCanvasPending = null;
                this.compositeRefreshTimers = new WeakMap();

                this.lightboxTool = 'pan';
                this.lightboxAnnotationAction = null;
                this.lightboxCropDraft = null;
                this.isCreatingCropDetail = false;
                this.lightboxZoom = 1.0;
                this.lightboxPanX = 0;
                this.lightboxPanY = 0;
                this.lightboxIsPanning = false;
                this.lightboxPanStart = { x: 0, y: 0 };
                this.lightboxHasRotated = false;
                this.lightboxNeedsRender = false;
                this.lightboxPhotoUid = null;
                this.lightboxBackdropPointerDown = false;
            }

            initEvents() {
                // 大圖燈箱縮放、平移與滾輪互動
                this.initLightboxInteractions();

                // 外部檔案拖入空白狀態或照片舞台；既有卡片拖曳排序不含 Files，不會被攔截。
                this.bindExternalFileDropTarget(this.emptyState);
                this.bindExternalFileDropTarget(this.thumbGrid, this.thumbGridDropZone);

                this.fileInput.addEventListener('change', (e) => {
                    const files = Array.from(e.target.files).filter(f => f.type.startsWith('image/'));
                    if (files.length) this.handleFiles(files).catch(error => this.handleProcessingError(error));
                    this.fileInput.value = '';
                });
                this.folderInput.addEventListener('change', (e) => {
                    const files = Array.from(e.target.files).filter(f => f.type.startsWith('image/'));
                    if (files.length) this.handleFiles(files).catch(error => this.handleProcessingError(error));
                    this.folderInput.value = '';
                });
                document.getElementById('btnAddPhotos').addEventListener('click', () => this.fileInput.click());
                document.getElementById('btnAddFolder').addEventListener('click', () => this.folderInput.click());

                // 按鈕事件
                document.getElementById('btnTodayDate').addEventListener('click', () => this.setDefaultTodayDate());
                document.getElementById('btnToggleSelectAll').addEventListener('click', () => this.toggleSelectAll());
                document.getElementById('btnViewWorkspace').addEventListener('click', () => this.switchView('workspace'));
                document.getElementById('btnViewGrid').addEventListener('click', () => this.switchView('grid'));
                document.getElementById('btnResetThumbnailSize').addEventListener('click', () => this.setThumbnailSize(135));
                this.thumbnailSizeRange.addEventListener('input', (event) => this.setThumbnailSize(event.target.value));
                this.mainPanelSplitter.addEventListener('pointerdown', (event) => this.beginMainPanelResize(event));
                document.addEventListener('pointermove', (event) => this.updateMainPanelResize(event));
                document.addEventListener('pointerup', (event) => this.endMainPanelResize(event));
                document.addEventListener('pointercancel', (event) => this.endMainPanelResize(event, true));
                // 卡片可能在舞台邊界外放開；全域監聽可確實清除「移動中」提示與浮動狀態。
                document.addEventListener('pointermove', (event) => this.updateCardPointerDrag(event));
                document.addEventListener('pointerup', (event) => this.endCardPointerDrag(event));
                document.addEventListener('pointercancel', (event) => this.endCardPointerDrag(event, true));
                window.addEventListener('blur', () => this.cancelCardPointerDrag());
                this.thumbGrid.addEventListener('wheel', (event) => this.handleThumbnailWheel(event), { passive: false });
                // Ctrl＋滾輪離開舞台時不交給瀏覽器縮放整頁；舞台內仍由上方監聽器調整縮圖。
                document.addEventListener('wheel', (event) => {
                    if (event.ctrlKey && !this.thumbGrid.contains(event.target)) event.preventDefault();
                }, { passive: false, capture: true });
                this.thumbGrid.addEventListener('pointerdown', (event) => this.beginMarqueeSelection(event));
                this.thumbGrid.addEventListener('pointermove', (event) => {
                    this.updateStageDrag(event);
                    this.updateMarqueeSelection(event);
                });
                this.thumbGrid.addEventListener('pointerup', (event) => {
                    this.endStageDrag(event);
                    this.endMarqueeSelection(event);
                });
                this.thumbGrid.addEventListener('pointercancel', (event) => {
                    this.endStageDrag(event, true);
                    this.endMarqueeSelection(event);
                });
                this.thumbGrid.addEventListener('dragover', (event) => this.handleStageGridDragOver(event));
                this.thumbGrid.addEventListener('drop', (event) => this.handleStageGridDrop(event));
                document.getElementById('btnAutoArrangeStage').addEventListener('click', () => this.sortPhotosByNaturalFilename({ confirmBeforeSort: true }));
                document.getElementById('btnSortByCaptureTime')?.addEventListener('click', () => this.sortPhotosByCaptureTime({ confirmBeforeSort: true }));
                this.btnToggleEditor.addEventListener('click', () => this.toggleEditor());
                this.btnToggleAudit?.addEventListener('click', () => {
                    this.auditPanelOpen = !this.auditPanelOpen;
                    if (!this.auditPanelOpen && this.activeFilter !== 'all') this.setActiveFilter('all');
                    else this.updateAuditBarUi();
                });
                this.setupMenus();
                document.getElementById('btnClearAll').addEventListener('click', () => this.clearAll());
                document.getElementById('btnBatchRotate').addEventListener('click', () => this.batchRotate(90));
                document.getElementById('btnMoveSelectedBackward').addEventListener('click', () => this.moveSelectedBy(-1));
                document.getElementById('btnMoveSelectedForward').addEventListener('click', () => this.moveSelectedBy(1));
                document.getElementById('btnRemoveSelected').addEventListener('click', () => this.removeSelectedPhotos());
                document.getElementById('btnApplyDefaultsToSelected').addEventListener('click', () => this.applyDefaultsToSelected());
                document.getElementById('btnOpenSequenceModal').addEventListener('click', () => this.openModal('sequenceModal'));
                document.getElementById('btnApplySequence').addEventListener('click', () => this.applySequence());
                document.getElementById('btnParseFilenameTime').addEventListener('click', () => this.parseFilenameTime());
                document.getElementById('btnOpenShortcutsModal')?.addEventListener('click', () => this.openModal('shortcutsModal'));
                document.getElementById('btnExportDocx').addEventListener('click', () => this.confirmExportWithAudit(() => this.exportDocx().catch(e => this.handleExportError(e)), 'Word 清冊'));
                document.getElementById('btnSaveProject').addEventListener('click', () => this.saveProject({ fromUserPicker: true }).catch(error => this.handleProjectError(error, 'save')));
                this.btnSaveProjectAs?.addEventListener('click', () => this.saveProject({ saveAsNew: true, fromUserPicker: true }).catch(error => this.handleProjectError(error, 'save')));
                this.btnQuickSaveStatus?.addEventListener('click', () => this.saveProject({ fromUserPicker: true }).catch(error => this.handleProjectError(error, 'save')));
                this.btnRestoreAutoDraft?.addEventListener('click', () => this.restoreAutoDraft().catch(error => this.handleProjectError(error, 'open')));
                this.btnDiscardAutoDraft?.addEventListener('click', () => this.discardAutoDraft());
                document.getElementById('importProjectInput').addEventListener('change', (e) => this.openProject(e));
                document.getElementById('mergeProjectInput')?.addEventListener('change', (e) => this.mergeProjects(e));
                document.getElementById('btnUndo').addEventListener('click', () => this.undo());
                document.getElementById('btnRedo').addEventListener('click', () => this.redo());
                document.getElementById('btnExportPdf').addEventListener('click', () => this.confirmExportWithAudit(() => this.exportPdf().catch(e => this.handleExportError(e)), 'PDF 清冊'));
                this.layoutSelect.addEventListener('change', () => {
                    this.updateLayoutDescription();
                    this.updatePageEstimate();
                    if (this.isLightboxOpen()) this.syncLightboxAnnotationLayerBounds();
                    this.scheduleHistory('變更清冊版型');
                });
                this.btnCheckUpdate.addEventListener('click', () => this.checkForUpdates());
                this.btnCancelProcessing.addEventListener('click', () => { this.processingCancelled = true; });
                window.addEventListener('beforeunload', (event) => {
                    if (!this.hasUnsavedProjectChanges()) return;
                    // 現代瀏覽器／WebView2 僅顯示其內建離開確認訊息，禁止自訂文案。
                    event.preventDefault();
                    event.returnValue = '';
                });
                ['caseTitle', 'caseDate', 'deptName', 'officerName'].forEach(id => {
                    document.getElementById(id).addEventListener('input', () => this.scheduleHistory('修改案件基本資料'));
                });
                this.bindLimitedTextControl(this.defaultLocationInput, {
                    maxFullWidth: 17,
                    singleLine: true,
                    onAccept: () => {
                        this.refreshAllTextLimitUi();
                        this.photos.forEach((_, idx) => this.refreshThumbCardStatusBadge(idx));
                        this.updateAuditBarUi();
                        this.scheduleHistory('修改案件基本資料');
                    },
                    onFeedback: () => this.refreshDefaultLocationLimitUi()
                });
                this.defaultLocationInput?.addEventListener('blur', () => this.render());

                const caseDateInput = document.getElementById('caseDate');
                caseDateInput.addEventListener('pointerdown', (e) => this.handleMaskedPointerDown(e, caseDateInput));
                caseDateInput.addEventListener('focus', () => this.beginMaskedEntry(caseDateInput));
                caseDateInput.addEventListener('input', (e) => {
                    // 輸入途中不可改寫成含 / 的字串；Windows 觸控鍵盤會把下一鍵插回舊格式位置。
                    this.sanitizeMaskedEntry(e.target, 7);
                    this.markMaskedEntryChanged(e.target);
                });
                caseDateInput.addEventListener('blur', (e) => { e.target.value = this.finishMaskedEntry(e.target, this.formatDateMask); });

                // 右側編輯器即時輸入事件
                this.editSeq.addEventListener('input', (e) => {
                    if (this.editSeq.readOnly) return;
                    this.updateCurrentPhoto('seq', e.target.value);
                });
                this.btnJumpToParentPhoto?.addEventListener('click', () => this.jumpToParentPhoto());
                this.btnUnlinkChildPhoto?.addEventListener('click', () => this.unlinkCurrentChildPhoto());
                this.editTime.addEventListener('pointerdown', (e) => this.handleMaskedPointerDown(e, this.editTime));
                this.editTime.addEventListener('focus', () => this.beginMaskedEntry(this.editTime));
                this.editTime.addEventListener('input', (e) => {
                    const digits = this.sanitizeMaskedEntry(e.target, 6);
                    this.markMaskedEntryChanged(e.target);
                    this.updateCurrentPhoto('time', digits);
                });
                this.editTime.addEventListener('blur', (e) => {
                    const masked = this.finishMaskedEntry(e.target, this.formatTimeMask);
                    e.target.value = masked;
                    this.updateCurrentPhoto('time', masked);
                });
                this.editDate.addEventListener('input', (e) => {
                    const digits = this.sanitizeMaskedEntry(e.target, 7);
                    this.markMaskedEntryChanged(e.target);
                    this.updateCurrentPhoto('date', digits);
                });
                this.editDate.addEventListener('pointerdown', (e) => this.handleMaskedPointerDown(e, this.editDate));
                this.editDate.addEventListener('focus', () => this.beginMaskedEntry(this.editDate));
                this.editDate.addEventListener('blur', (e) => {
                    const masked = this.finishMaskedEntry(e.target, this.formatDateMask);
                    e.target.value = masked;
                    this.updateCurrentPhoto('date', masked);
                });
                this.bindLimitedTextControl(this.editLocation, {
                    maxFullWidth: 17,
                    singleLine: true,
                    onAccept: (val) => this.updateCurrentPhoto('location', val),
                    onFeedback: (draft) => this.refreshEditLocationLimitUi(draft)
                });
                this.bindLimitedTextControl(this.editDesc, {
                    maxFullWidth: 68,
                    singleLine: false,
                    maxLines: 2,
                    lineFullWidth: 34,
                    onAccept: (val) => this.updateCurrentPhoto('desc', val),
                    onFeedback: (draft) => this.refreshEditDescLimitUi(draft)
                });
                this.bindGridTextLimitEvents();

                // 右側向下填滿按鈕
                document.getElementById('btnFillDownDate').addEventListener('click', () => this.fillDownFromIndex(this.currentIndex, 'date'));
                document.getElementById('btnFillDownTime').addEventListener('click', () => this.fillDownFromIndex(this.currentIndex, 'time'));
                document.getElementById('btnFillDownLocation').addEventListener('click', () => this.fillDownFromIndex(this.currentIndex, 'location'));
                document.getElementById('btnFillDownDescAll').addEventListener('click', () => this.fillDownFromIndex(this.currentIndex, 'desc'));
                document.getElementById('btnApplyToSelected').addEventListener('click', () => this.applyEditorToSelected());

                // 上一張 / 下一張 / 刪除
                document.getElementById('btnPrevPhoto').addEventListener('click', () => this.navigatePhoto(-1));
                document.getElementById('btnNextPhoto').addEventListener('click', () => this.navigatePhoto(1));
                document.getElementById('btnDeleteCurrent').addEventListener('click', () => this.removePhoto(this.currentIndex));

                // 鍵盤快速鍵：Ctrl+S 儲存進度；方向鍵切換單張選取；Ctrl＋方向鍵與 Ctrl＋Home／End 重排已選照片；Delete 移除照片。
                window.addEventListener('keydown', (e) => {
                    if ((e.ctrlKey || e.metaKey) && !e.altKey && e.key.toLowerCase() === 's') {
                        e.preventDefault();
                        if (!this.hasBlockingModalOpen()) {
                            this.saveProject({ saveAsNew: Boolean(e.shiftKey), fromUserPicker: true })
                                .catch(error => this.handleProjectError(error, 'save'));
                        }
                        return;
                    }
                    const isCheckboxDeleteShortcut = e.key === 'Delete' &&
                        !e.ctrlKey && !e.altKey && !e.shiftKey && !e.metaKey &&
                        e.target instanceof HTMLInputElement && e.target.type === 'checkbox';
                    if (this.isTypingTarget(e.target) && !isCheckboxDeleteShortcut) {
                        if (e.key === 'Escape' && e.target instanceof HTMLElement) {
                            e.preventDefault();
                            e.target.blur();
                        } else if (e.altKey && (e.key === 'ArrowLeft' || e.key === 'ArrowRight')) {
                            e.preventDefault();
                            if (e.target instanceof HTMLElement) e.target.blur();
                            this.navigatePhoto(e.key === 'ArrowLeft' ? -1 : 1);
                        }
                        return;
                    }
                    const blockingModalOpen = this.hasBlockingModalOpen();
                    const lightboxOpen = !document.getElementById('imageLightboxModal').classList.contains('hidden');
                    if (e.key === 'Escape') {
                        if (lightboxOpen && this.lightboxCropDraft) {
                            this.lightboxCropDraft = null;
                            this.renderLightboxAnnotations();
                            e.preventDefault();
                            return;
                        }
                        ['shortcutsModal', 'sequenceModal', 'updateModal', 'upToDateModal', 'updateErrorModal', 'exportAuditModal', 'imageLightboxModal'].forEach(id => {
                            this.closeModal(id);
                        });
                    }
                    if (lightboxOpen) {
                        if (e.key === 'Escape') {
                            this.closeModal('imageLightboxModal');
                        } else if (e.key === 'Delete' && !e.ctrlKey && !e.altKey && !e.shiftKey && !e.metaKey && !blockingModalOpen) {
                            if (!this.handleDeleteShortcut({ fromLightbox: true })) return;
                        } else if (e.key === 'Enter' && !e.ctrlKey && !e.altKey && !e.shiftKey && this.lightboxCropDraft && !this.isCreatingCropDetail) {
                            this.confirmLightboxCrop();
                        } else if (e.ctrlKey && !e.altKey && e.key.toLowerCase() === 'z') {
                            this.commitLightboxHistoryIfNeeded();
                            if (e.shiftKey) this.redo(); else this.undo();
                        } else if (e.ctrlKey && !e.altKey && e.key.toLowerCase() === 'y') {
                            this.commitLightboxHistoryIfNeeded();
                            this.redo();
                        } else if (e.key === 'ArrowLeft') {
                            this.navigateLightbox(-1);
                        } else if (e.key === 'ArrowRight') {
                            this.navigateLightbox(1);
                        } else if (e.key === 'r' || e.key === 'R') {
                            if (e.shiftKey) {
                                this.rotateLightboxPhoto(-90);
                            } else {
                                this.rotateLightboxPhoto(90);
                            }
                        } else if (!e.ctrlKey && !e.altKey && (e.key === 'h' || e.key === 'H')) {
                            this.toggleLightboxFlipH();
                        } else if (!e.ctrlKey && !e.altKey && (e.key === 'v' || e.key === 'V')) {
                            this.setLightboxTool('pan');
                        } else if (!e.ctrlKey && !e.altKey && (e.key === 'y' || e.key === 'Y')) {
                            this.setLightboxTool('rect-yellow');
                        } else if (!e.ctrlKey && !e.altKey && (e.key === 'm' || e.key === 'M')) {
                            this.setLightboxTool('rect-red');
                        } else if (!e.ctrlKey && !e.altKey && (e.key === 't' || e.key === 'T')) {
                            this.setLightboxTool('label');
                        } else if (!e.ctrlKey && !e.altKey && (e.key === 'c' || e.key === 'C')) {
                            this.setLightboxTool('crop');
                        } else if (e.key === '0') {
                            this.resetLightboxZoom();
                        } else if (e.key === '+' || e.key === '=') {
                            this.zoomLightboxBy(0.2);
                        } else if (e.key === '-' || e.key === '_') {
                            this.zoomLightboxBy(-0.2);
                        } else {
                            return;
                        }
                        e.preventDefault();
                        return;
                    }

                    // 主畫布快捷鍵：按空白鍵或 Enter 鍵放大檢視照片
                    if ((e.key === ' ' || e.key === 'Enter') && !e.ctrlKey && !e.altKey && !e.shiftKey) {
                        if (!e.target.closest('button, a, input, select, textarea') && this.currentIndex >= 0 && this.photos[this.currentIndex]) {
                            e.preventDefault();
                            this.openLightboxByIndex(this.currentIndex);
                            return;
                        }
                    }
                    if (e.ctrlKey && !e.altKey && e.key.toLowerCase() === 'z') {
                        e.preventDefault();
                        if (e.shiftKey) this.redo(); else this.undo();
                        return;
                    }
                    if (e.ctrlKey && !e.altKey && e.key.toLowerCase() === 'y') {
                        e.preventDefault();
                        this.redo();
                        return;
                    }
                    if (e.key === 'Delete' && !e.ctrlKey && !e.altKey && !e.shiftKey && !e.metaKey && !blockingModalOpen) {
                        if (this.handleDeleteShortcut({ fromLightbox: false })) {
                            e.preventDefault();
                            if (isCheckboxDeleteShortcut && e.target instanceof HTMLElement) e.target.blur();
                        }
                        return;
                    }
                    if (e.key === 'Escape') this.closeModal('imageLightboxModal');
                    if (e.altKey && e.key === 'ArrowLeft') {
                        e.preventDefault();
                        this.navigatePhoto(-1);
                    }
                    if (e.altKey && e.key === 'ArrowRight') {
                        e.preventDefault();
                        this.navigatePhoto(1);
                    }
                    if (e.ctrlKey && e.key === 'ArrowLeft') {
                        e.preventDefault();
                        this.moveSelectedBy(-1);
                    }
                    if (e.ctrlKey && e.key === 'ArrowRight') {
                        e.preventDefault();
                        this.moveSelectedBy(1);
                    }
                    if (e.ctrlKey && e.key === 'ArrowUp') {
                        e.preventDefault();
                        this.moveSelectedBy(-this.getGridColumns());
                    }
                    if (e.ctrlKey && e.key === 'ArrowDown') {
                        e.preventDefault();
                        this.moveSelectedBy(this.getGridColumns());
                    }
                    if (!e.ctrlKey && !e.altKey && !e.shiftKey && e.key === 'ArrowLeft') {
                        e.preventDefault();
                        this.selectPhotoByKeyboard(-1);
                    }
                    if (!e.ctrlKey && !e.altKey && !e.shiftKey && e.key === 'ArrowRight') {
                        e.preventDefault();
                        this.selectPhotoByKeyboard(1);
                    }
                    if (!e.ctrlKey && !e.altKey && !e.shiftKey && e.key === 'ArrowUp') {
                        e.preventDefault();
                        this.selectPhotoByKeyboard(-this.getGridColumns());
                    }
                    if (!e.ctrlKey && !e.altKey && !e.shiftKey && e.key === 'ArrowDown') {
                        e.preventDefault();
                        this.selectPhotoByKeyboard(this.getGridColumns());
                    }
                    if (e.ctrlKey && e.key === 'Home') {
                        e.preventDefault();
                        this.moveSelectedToBoundary('start');
                    }
                    if (e.ctrlKey && e.key === 'End') {
                        e.preventDefault();
                        this.moveSelectedToBoundary('end');
                    }
                });

                // 完整度稽核篩選列按鈕
                document.getElementById('filterBtn_all').addEventListener('click', () => this.setActiveFilter('all'));
                document.getElementById('filterBtn_missingLocation').addEventListener('click', () => this.setActiveFilter('missingLocation'));
                document.getElementById('filterBtn_missingDesc').addEventListener('click', () => this.setActiveFilter('missingDesc'));
                document.getElementById('filterBtn_invalidDateTime').addEventListener('click', () => this.setActiveFilter('invalidDateTime'));
                document.getElementById('filterBtn_duplicatePhotos').addEventListener('click', () => this.setActiveFilter('duplicatePhotos'));
                document.getElementById('filterBtn_overTextLimit')?.addEventListener('click', () => this.setActiveFilter('overTextLimit'));

                // 匯出前稽核彈窗按鈕
                document.getElementById('btnAuditRenumber').addEventListener('click', () => {
                    this.pendingSequenceExport = { action: this.pendingExportAction, title: this.pendingExportTitle };
                    this.closeModal('exportAuditModal');
                    this.openModal('sequenceModal');
                });
                this.btnAuditViewIssues.addEventListener('click', () => {
                    this.exportAuditModal.classList.add('hidden');
                    const audit = this.auditPhotosCompleteness();
                    const firstFilter = audit.firstIssueFilter || 'all';
                    this.setActiveFilter(firstFilter);
                });
                this.btnAuditProceedExport.addEventListener('click', () => {
                    this.exportAuditModal.classList.add('hidden');
                    if (this.pendingExportAction) {
                        const action = this.pendingExportAction;
                        this.pendingExportAction = null;
                        action();
                    }
                });
            }

            setDefaultTodayDate() {
                const now = new Date();
                const mYear = now.getFullYear() - 1911;
                const mMonth = now.getMonth() + 1;
                const mDay = now.getDate();
                document.getElementById('caseDate').value = `${mYear}${String(mMonth).padStart(2, '0')}${String(mDay).padStart(2, '0')}`;
                document.getElementById('caseDate').value = this.formatDateMask(document.getElementById('caseDate').value, true);
                this.scheduleHistory('填入今天日期');
            }

            getVisiblePhotoIndices() {
                if (typeof PhotoReportSelection !== 'undefined') {
                    const globalLocation = document.getElementById('defaultLocation')?.value?.trim() || '';
                    if (!this._dupSet && this.activeFilter === 'duplicatePhotos') this._buildDupSet();
                    return PhotoReportSelection.getVisiblePhotoIndices(this.photos, this.activeFilter, globalLocation, this._dupSet);
                }
                const all = this.photos.map((photo, index) => ({ photo, index }));
                if (this.activeFilter === 'all') return all;
                const globalLocation = document.getElementById('defaultLocation')?.value?.trim() || '';
                return all.filter(({ photo }) => {
                    switch (this.activeFilter) {
                        case 'missingLocation':
                            return !photo.location?.trim() && !globalLocation;
                        case 'missingDesc':
                            return !photo.desc?.trim();
                        case 'invalidDateTime': {
                            const d = photo.date?.trim() || '';
                            const t = photo.time?.trim() || '';
                            const dateInvalid = d ? !this.isValidMinguoDate(d) : false;
                            const timeInvalid = !this.isValidTimeFormat(t);
                            return dateInvalid || timeInvalid;
                        }
                        case 'duplicatePhotos': {
                            if (!this._dupSet) this._buildDupSet();
                            return this._dupSet.has(photo.name);
                        }
                        case 'overTextLimit': {
                            const effectiveLoc = photo.location?.trim() || globalLocation;
                            return this.inspectPhotoTextLimits(effectiveLoc, photo.desc?.trim() || '').isOver;
                        }
                        default:
                            return true;
                    }
                });
            }

            async loadAppVersion() {
                const embeddedVersion = this.getCurrentVersion();
                if (this.isLocalFileMode()) {
                    this.appVersionBadge.title = `目前版本 ${embeddedVersion}（直接開啟本機檔案，使用內嵌版本）`;
                    this.checkUpdateText.textContent = '開啟更新頁';
                    this.btnCheckUpdate.title = '直接開啟 GitHub Releases 查看最新版本';
                    return;
                }

                try {
                    const response = await fetch('version.txt', { cache: 'no-store' });
                    if (!response.ok) throw new Error(`version.txt 讀取失敗（${response.status}）`);
                    const version = (await response.text()).trim();
                    if (!version) throw new Error('version.txt 未提供版本號。');
                    this.appVersionText.textContent = version.startsWith('v') ? version : `v${version}`;
                    this.appVersionBadge.title = `目前版本 ${this.getCurrentVersion()}（version.txt）`;
                } catch (error) {
                    // 網路暫時不可用時維持內嵌版本，不在主控台製造無法處理的執行期警告。
                    this.appVersionBadge.title = `目前版本 ${embeddedVersion}（使用內嵌版本）`;
                }
            }

            isLocalFileMode() {
                return window.location.protocol === 'file:';
            }

            getCurrentVersion() {
                return this.appVersionText.textContent.trim() || 'v0.0.0';
            }

            compareVersions(left, right) {
                const toParts = (version) => String(version)
                    .trim()
                    .replace(/^v/i, '')
                    .split(/[+-]/, 1)[0]
                    .split('.')
                    .map(part => Number.parseInt(part, 10) || 0);
                const leftParts = toParts(left);
                const rightParts = toParts(right);
                const length = Math.max(leftParts.length, rightParts.length);
                for (let index = 0; index < length; index++) {
                    const difference = (leftParts[index] || 0) - (rightParts[index] || 0);
                    if (difference) return Math.sign(difference);
                }
                return 0;
            }

            setUpdateButtonLoading(isLoading) {
                this.btnCheckUpdate.disabled = isLoading;
                this.btnCheckUpdate.classList.toggle('opacity-60', isLoading);
                this.btnCheckUpdate.classList.toggle('cursor-wait', isLoading);
                this.btnCheckUpdate.setAttribute('aria-busy', String(isLoading));
                this.checkUpdateIcon.classList.toggle('fa-spin', isLoading);
                this.checkUpdateText.textContent = isLoading ? '檢查中…' : (this.isLocalFileMode() ? '開啟更新頁' : '檢查更新');
            }

            async checkForUpdates() {
                if (this.isLocalFileMode()) {
                    // file:// 無法可靠呼叫跨來源 API；由使用者手勢直接開啟受信任的 Releases 頁面。
                    window.open('https://github.com/lianghao02/Photo-Report-Generator/releases', '_blank', 'noopener,noreferrer');
                    return;
                }
                this.setUpdateButtonLoading(true);
                const controller = new AbortController();
                const timeoutId = window.setTimeout(() => controller.abort(), 12000);
                try {
                    const response = await fetch(
                        'https://api.github.com/repos/lianghao02/Photo-Report-Generator/releases/latest',
                        {
                            headers: { Accept: 'application/vnd.github+json' },
                            signal: controller.signal,
                            cache: 'no-store'
                        }
                    );
                    if (!response.ok) throw new Error(`GitHub 回應 ${response.status}`);
                    const release = await response.json();
                    if (!release?.tag_name || !release?.html_url) throw new Error('GitHub Release 資料格式不完整。');

                    const currentVersion = this.getCurrentVersion();
                    if (this.compareVersions(release.tag_name, currentVersion) > 0) {
                        this.showUpdateAvailable(release, currentVersion);
                    } else {
                        document.getElementById('upToDateVer').textContent = currentVersion;
                        this.openModal('upToDateModal');
                    }
                } catch (error) {
                    console.error('檢查更新失敗：', error);
                    const detail = error.name === 'AbortError'
                        ? '連線逾時，請確認網路狀態後再試一次。'
                        : `暫時無法連線 GitHub 檢查更新（${error.message || '未知錯誤'}）。`;
                    document.getElementById('updateErrorDetail').textContent = detail;
                    this.openModal('updateErrorModal');
                } finally {
                    window.clearTimeout(timeoutId);
                    this.setUpdateButtonLoading(false);
                }
            }

            showUpdateAvailable(release, currentVersion) {
                document.getElementById('updateCurrentVer').textContent = currentVersion;
                document.getElementById('updateLatestVer').textContent = release.tag_name;
                document.getElementById('updateReleaseDate').textContent = release.published_at
                    ? new Date(release.published_at).toLocaleDateString('zh-TW')
                    : '未提供';
                document.getElementById('updateReleaseBody').textContent = release.body?.trim() || '此版本未提供更新說明。';

                const releaseLink = document.getElementById('btnGoToRelease');
                releaseLink.href = release.html_url;
                const assetsContainer = document.getElementById('updateAssetsContainer');
                const assetsList = document.getElementById('updateAssetsList');
                assetsList.replaceChildren();

                const assets = Array.isArray(release.assets) ? release.assets : [];
                assets.forEach(asset => {
                    try {
                        const downloadUrl = new URL(asset.browser_download_url);
                        if (!['https:', 'http:'].includes(downloadUrl.protocol)) return;
                        const link = document.createElement('a');
                        link.href = downloadUrl.href;
                        link.target = '_blank';
                        link.rel = 'noopener noreferrer';
                        link.className = 'btn-secondary px-2.5 py-1.5 rounded-lg text-[11px] font-bold flex items-center gap-1';
                        link.innerHTML = '<i class="fa-solid fa-download"></i>';
                        link.append(document.createTextNode(` ${asset.name || '下載檔案'}`));
                        assetsList.append(link);
                    } catch {
                        // 略過格式不正確的下載網址，保留官方 Release 頁面作為安全退路。
                    }
                });
                assetsContainer.classList.toggle('hidden', assetsList.childElementCount === 0);
                this.openModal('updateModal');
            }

            async handleFiles(files) {
                // Windows 檔案選取／拖放的 FileList 順序不保證一致；每批先依檔名自然排序，
                // 再附加到目前清單尾端，絕不改動既有或手動調整過的照片順序。
                const sortedFiles = [...files].sort((a, b) => this.comparePhotoFileOrder(a, b));
                this.processingCancelled = false;
                const incomingBytes = sortedFiles.reduce((total, file) => total + file.size, 0);
                const totalBytes = this.photos.reduce((total, photo) => total + photo.file.size, 0) + incomingBytes;
                const totalPhotos = this.photos.length + sortedFiles.length;

                if ((totalPhotos > this.maxRecommendedPhotos || totalBytes > this.maxRecommendedBytes) && !confirm(
                    `即將載入 ${totalPhotos} 張照片（約 ${this.formatBytes(totalBytes)}）。\n` +
                    '照片很多時，處理和匯出會比較久，建議分成幾冊。要繼續嗎？'
                )) return;

                let failedCount = 0;
                this.setProgress('正在建立低記憶體縮圖…', 0, sortedFiles.length);
                for (let i = 0; i < sortedFiles.length; i++) {
                    const file = sortedFiles[i];
                    let sourceUrl = '';
                    try {
                        sourceUrl = URL.createObjectURL(file);
                        const thumbnail = await this.createThumbnailUrl(file, sourceUrl);
                        const parsedTime = this.extractTimeFromFilename(file.name);
                        const photoItem = {
                            uid: 'uid_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9),
                            file: file,
                            name: file.name,
                            src: sourceUrl,
                            thumbnailSrc: thumbnail.src,
                            thumbnailAspect: thumbnail.aspect,
                            visualFingerprint: thumbnail.visualFingerprint || '',
                            rotation: 0,
                            flipH: false,
                            parentUid: null,
                            cropRect: null,
                            sideCrop: null,
                            compositeThumbnailSrc: null,
                            _compositeKey: '',
                            annotations: [],
                            exifRotation: await this.readJpegExifRotation(file),
                            seq: String(this.getNextMainSeqNumber()),
                            date: '',
                            time: parsedTime || '12:00',
                            timeAutoDefaulted: !parsedTime,
                            // 留空代表繼承清冊預設地點，避免把預設值誤認為本張覆寫資料。
                            location: '',
                            desc: '',
                            selected: false,
                            stageX: null,
                            stageY: null
                        };
                        this.photos.push(photoItem);
                        this.managedUrls.add(sourceUrl);
                        this.managedUrls.add(photoItem.thumbnailSrc);
                    } catch (error) {
                        if (sourceUrl) URL.revokeObjectURL(sourceUrl);
                        failedCount++;
                        console.warn(`略過無法載入的圖片：${file.name}`, error);
                    }
                    this.setProgress('正在建立低記憶體縮圖…', i + 1, sortedFiles.length);
                    if (this.processingCancelled) {
                        this.render();
                        this.recordHistory('載入照片（已取消）');
                        throw new Error('使用者已取消處理。');
                    }
                    if ((i + 1) % 10 === 0) {
                        this.render();
                        await this.yieldToBrowser();
                    }
                }

                if (this.photos.length > 0 && this.currentIndex >= this.photos.length) {
                    this.currentIndex = 0;
                }
                this.render();
                this.clearProgress();
                this.recordHistory(`載入 ${sortedFiles.length - failedCount} 張照片`);
                if (failedCount) alert(`已載入 ${sortedFiles.length - failedCount} 張照片；${failedCount} 張因無法解碼而略過。`);
            }

            comparePhotoFileOrder(a, b) {
                const aKey = a.webkitRelativePath || a.name || '';
                const bKey = b.webkitRelativePath || b.name || '';
                return aKey.localeCompare(bKey, 'zh-Hant', { numeric: true, sensitivity: 'base' }) ||
                    String(a.name || '').localeCompare(String(b.name || ''), 'zh-Hant', { numeric: true, sensitivity: 'base' });
            }

            async createThumbnailUrl(file, sourceUrl) {
                const maxDimension = 320;
                let source;
                try {
                    source = typeof createImageBitmap === 'function'
                        ? await createImageBitmap(file)
                        : await this.loadImage(sourceUrl);
                    const ratio = Math.min(maxDimension / source.width, maxDimension / source.height, 1);
                    const canvas = document.createElement('canvas');
                    canvas.width = Math.max(1, Math.round(source.width * ratio));
                    canvas.height = Math.max(1, Math.round(source.height * ratio));
                    const context = canvas.getContext('2d');
                    if (!context) throw new Error('無法建立縮圖畫布。');
                    context.drawImage(source, 0, 0, canvas.width, canvas.height);
                    const blob = await new Promise((resolve, reject) => canvas.toBlob(
                        result => result ? resolve(result) : reject(new Error('無法產生縮圖。')),
                        'image/jpeg', 0.8
                    ));
                    return {
                        src: URL.createObjectURL(blob),
                        aspect: canvas.width / canvas.height,
                        visualFingerprint: this.computeVisualFingerprintFromCanvas(canvas)
                    };
                } finally {
                    if (source && typeof source.close === 'function') source.close();
                }
            }

            normalizePhotoFilenameKey(name) {
                const trimmed = String(name || '').trim().toLowerCase();
                return trimmed.replace(/\.(jpe?g|png|webp|gif|bmp|heic)$/i, '').replace(/\s+/g, ' ');
            }

            computeVisualFingerprintFromCanvas(sourceCanvas) {
                if (!sourceCanvas) return '';
                try {
                    const w = 17;
                    const h = 16;
                    const fpCanvas = document.createElement('canvas');
                    fpCanvas.width = w;
                    fpCanvas.height = h;
                    const ctx = fpCanvas.getContext('2d', { willReadFrequently: true });
                    if (!ctx) return '';
                    ctx.fillStyle = '#FFFFFF';
                    ctx.fillRect(0, 0, w, h);
                    ctx.imageSmoothingEnabled = true;
                    ctx.drawImage(sourceCanvas, 0, 0, w, h);
                    const data = ctx.getImageData(0, 0, w, h).data;

                    // 16x16 horizontal dHash（256 bits => 64 位 hex）
                    let hexBits = '';
                    let nibble = 0;
                    let bitIdx = 0;
                    const grayAt = (x, y) => {
                        const i = (y * w + x) * 4;
                        return data[i] * 0.299 + data[i + 1] * 0.587 + data[i + 2] * 0.114;
                    };
                    for (let y = 0; y < h; y++) {
                        for (let x = 0; x < 16; x++) {
                            const bit = grayAt(x, y) > grayAt(x + 1, y) ? 1 : 0;
                            nibble = (nibble << 1) | bit;
                            bitIdx++;
                            if (bitIdx === 4) {
                                hexBits += nibble.toString(16);
                                nibble = 0;
                                bitIdx = 0;
                            }
                        }
                    }

                    // 2x2 象限平均 RGB（4 象限 x RGB => 12 位 hex），輔助區分純色與平緩漸層圖
                    let quadHex = '';
                    const regions = [
                        { x0: 0, x1: 8, y0: 0, y1: 8 },
                        { x0: 9, x1: 17, y0: 0, y1: 8 },
                        { x0: 0, x1: 8, y0: 8, y1: 16 },
                        { x0: 9, x1: 17, y0: 8, y1: 16 }
                    ];
                    regions.forEach(({ x0, x1, y0, y1 }) => {
                        let r = 0, g = 0, b = 0, count = 0;
                        for (let y = y0; y < y1; y++) {
                            for (let x = x0; x < x1; x++) {
                                const i = (y * w + x) * 4;
                                r += data[i];
                                g += data[i + 1];
                                b += data[i + 2];
                                count++;
                            }
                        }
                        if (count > 0) {
                            quadHex += Math.min(15, Math.round((r / count) / 17)).toString(16);
                            quadHex += Math.min(15, Math.round((g / count) / 17)).toString(16);
                            quadHex += Math.min(15, Math.round((b / count) / 17)).toString(16);
                        } else {
                            quadHex += '000';
                        }
                    });
                    return `${hexBits}:${quadHex}`;
                } catch {
                    return '';
                }
            }

            async computeVisualFingerprintFromBlob(blob) {
                if (!blob) return { visualFingerprint: '', aspect: 1 };
                let source = null;
                let tempUrl = '';
                try {
                    if (typeof createImageBitmap === 'function') {
                        source = await createImageBitmap(blob);
                    } else {
                        tempUrl = URL.createObjectURL(blob);
                        source = await this.loadImage(tempUrl);
                    }
                    const w = source.width || source.naturalWidth || 1;
                    const h = source.height || source.naturalHeight || 1;
                    const maxDim = 320;
                    const ratio = Math.min(maxDim / w, maxDim / h, 1);
                    const canvas = document.createElement('canvas');
                    canvas.width = Math.max(1, Math.round(w * ratio));
                    canvas.height = Math.max(1, Math.round(h * ratio));
                    const ctx = canvas.getContext('2d');
                    if (!ctx) return { visualFingerprint: '', aspect: w / h };
                    ctx.drawImage(source, 0, 0, canvas.width, canvas.height);
                    return {
                        visualFingerprint: this.computeVisualFingerprintFromCanvas(canvas),
                        aspect: canvas.width / canvas.height
                    };
                } catch {
                    return { visualFingerprint: '', aspect: 1 };
                } finally {
                    if (source && typeof source.close === 'function') source.close();
                    if (tempUrl) URL.revokeObjectURL(tempUrl);
                }
            }

            compareVisualFingerprints(fpA, fpB) {
                if (!fpA || !fpB) return { match: false, dhashDist: 999, totalQuadDiff: 999, maxQuadDiff: 999, score: 9999 };
                const [hashA = '', quadA = ''] = String(fpA).split(':');
                const [hashB = '', quadB = ''] = String(fpB).split(':');
                if (hashA.length !== 64 || hashB.length !== 64 || quadA.length !== 12 || quadB.length !== 12) {
                    return { match: false, dhashDist: 999, totalQuadDiff: 999, maxQuadDiff: 999, score: 9999 };
                }
                let dhashDist = 0;
                for (let i = 0; i < 64; i++) {
                    const xor = (parseInt(hashA[i], 16) || 0) ^ (parseInt(hashB[i], 16) || 0);
                    dhashDist += ((xor & 1) + ((xor >> 1) & 1) + ((xor >> 2) & 1) + ((xor >> 3) & 1));
                }
                let totalQuadDiff = 0;
                let maxQuadDiff = 0;
                for (let i = 0; i < 12; i++) {
                    const diff = Math.abs((parseInt(quadA[i], 16) || 0) - (parseInt(quadB[i], 16) || 0));
                    totalQuadDiff += diff;
                    if (diff > maxQuadDiff) maxQuadDiff = diff;
                }
                const match = dhashDist <= 28 && maxQuadDiff <= 2 && totalQuadDiff <= 14;
                return { match, dhashDist, totalQuadDiff, maxQuadDiff, score: dhashDist * 2 + totalQuadDiff };
            }

            loadImage(src) {
                return new Promise((resolve, reject) => {
                    const image = new Image();
                    image.onload = () => resolve(image);
                    image.onerror = () => reject(new Error('無法解碼圖片。'));
                    image.src = src;
                });
            }

            setProgress(message, current, total) {
                const percent = total ? Math.round((current / total) * 100) : 0;
                this.processingStatus.classList.remove('hidden');
                this.processingMessage.textContent = message;
                this.processingCount.textContent = `${current} / ${total}`;
                this.processingBar.style.width = `${percent}%`;
                this.btnCancelProcessing.classList.remove('hidden');
            }

            clearProgress() {
                this.processingStatus.classList.add('hidden');
                this.btnCancelProcessing.classList.add('hidden');
                this.processingCancelled = false;
            }

            throwIfProcessingCancelled() {
                if (this.processingCancelled) throw new Error('使用者已取消處理。');
            }

            yieldToBrowser() {
                return new Promise(resolve => requestAnimationFrame(resolve));
            }

            formatBytes(bytes) {
                if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
                if (bytes < 1024 * 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
                return `${(bytes / 1024 / 1024 / 1024).toFixed(1)} GB`;
            }

            async readJpegExifRotation(file) {
                if (file.type !== 'image/jpeg') return 0;
                try {
                    const bytes = new DataView(await file.slice(0, 65536).arrayBuffer());
                    if (bytes.getUint16(0) !== 0xFFD8) return 0;
                    let offset = 2;
                    while (offset + 4 <= bytes.byteLength) {
                        const marker = bytes.getUint16(offset);
                        const length = bytes.getUint16(offset + 2);
                        if (marker === 0xFFE1 && offset + 10 <= bytes.byteLength && bytes.getUint32(offset + 4) === 0x45786966) {
                            const tiff = offset + 10;
                            const littleEndian = bytes.getUint16(tiff) === 0x4949;
                            const ifd = tiff + bytes.getUint32(tiff + 4, littleEndian);
                            const count = bytes.getUint16(ifd, littleEndian);
                            for (let i = 0; i < count; i++) {
                                const entry = ifd + 2 + i * 12;
                                if (entry + 12 <= bytes.byteLength && bytes.getUint16(entry, littleEndian) === 0x0112) {
                                    return ({ 3: 180, 6: 90, 8: 270 })[bytes.getUint16(entry + 8, littleEndian)] || 0;
                                }
                            }
                            return 0;
                        }
                        if ((marker & 0xFF00) !== 0xFF00 || length < 2) break;
                        offset += 2 + length;
                    }
                } catch (error) {
                    console.warn('無法讀取 EXIF 方向，將保留原始方向。', error);
                }
                return 0;
            }

            handleExportError(error) {
                this.clearProgress();
                console.error(error);
                alert(`匯出失敗：${error.message || '圖片無法處理，請確認檔案格式是否受支援。'}`);
            }

            handleProcessingError(error) {
                this.clearProgress();
                if (error?.message === '使用者已取消處理。') {
                    alert('已取消處理，已完成的照片資料會保留。');
                    return;
                }
                console.error(error);
                alert(`處理失敗：${error?.message || '請確認圖片格式是否受支援。'}`);
            }

            extractTimeFromFilename(filename) {
                const m1 = filename.match(/(\d{4})[-_]?(\d{2})[-_]?(\d{2})[-_]?(\d{2})[-_]?(\d{2})[-_]?(\d{2})/);
                if (m1) return `${m1[4]}:${m1[5]}:${m1[6]}`;
                const m2 = filename.match(/[-_](\d{2})(\d{2})(\d{2})[-_.]/);
                if (m2) return `${m2[1]}:${m2[2]}:${m2[3]}`;
                const m3 = filename.match(/[-_](\d{2})(\d{2})[-_.]/);
                if (m3) return `${m3[1]}:${m3[2]}`;
                return null;
            }

            parseFilenameTime() {
                if (!this.photos.length) return alert('請先加入照片。');
                let count = 0;
                this.photos.forEach(p => {
                    const t = this.extractTimeFromFilename(p.name);
                    if (t) {
                        p.time = t;
                        p.timeAutoDefaulted = false;
                        count++;
                    }
                });
                this.render();
                this.recordHistory('解析檔名時間');
                alert(`已從檔名讀到 ${count} 張照片的日期時間。`);
            }

            applyDefaultsToSelected() {
                const selectedPhotos = this.photos.filter(photo => photo.selected);
                if (!selectedPhotos.length) return alert('請先選取照片。');
                const defaultDate = document.getElementById('caseDate').value.trim();
                const defaultLocation = document.getElementById('defaultLocation').value.trim();
                selectedPhotos.forEach(photo => {
                    photo.date = defaultDate;
                    photo.location = defaultLocation;
                });
                this.render();
                this.recordHistory('套用清冊預設至已選照片');
                alert(`已把案件日期與地點寫進 ${selectedPhotos.length} 張照片。`);
            }

            batchRotate(degree) {
                if (!this.photos.length) return;
                this.flushPendingHistory();
                const hasSelected = this.photos.some(p => p.selected);
                this.photos.forEach(p => {
                    if (!hasSelected || p.selected) {
                        const rawDelta = p.flipH ? -degree : degree;
                        this.transformPhotoAnnotationsForRotation(p, degree);
                        p.rotation = (((p.rotation || 0) + rawDelta) % 360 + 360) % 360;
                        if (p.sideCrop) this.scheduleSideCropPreviewRefresh(p);
                    }
                });
                this.render();
                this.recordHistory('批次旋轉照片');
            }

            applySequence() {
                const pendingExport = this.pendingSequenceExport;
                const prefix = document.getElementById('seqPrefix').value;
                const start = parseInt(document.getElementById('seqStart').value, 10) || 1;
                const digits = parseInt(document.getElementById('seqDigits').value, 10) || 1;

                this.normalizeLinkedChildrenOrder();
                let mainIdx = 0;
                this.photos.forEach(p => {
                    if (this.isLinkedCropChild(p)) return;
                    const num = start + mainIdx;
                    const numStr = String(num).padStart(digits, '0');
                    p.seq = `${prefix}${numStr}`;
                    mainIdx++;
                });
                this.syncChildCropSequences();

                this.closeModal('sequenceModal');
                this.render();
                this.recordHistory('重排流水號');
                if (pendingExport?.action) this.confirmExportWithAudit(pendingExport.action, pendingExport.title);
            }

            fillDownFromIndex(index, field) {
                if (!this.photos[index]) return;
                const val = this.photos[index][field];
                const targets = this.photos.slice(index + 1);
                if (!targets.length) return alert('這是最後一張照片，沒有後續照片可填入。');
                const label = { date: '日期', time: '時間', location: '地點', desc: '說明' }[field];
                const overwriteCount = field === 'desc' ? 0 : targets.filter(photo => String(photo[field] || '').trim() && photo[field] !== val).length;
                if (overwriteCount && !confirm(`將第 ${index + 1} 張的${label}套用至後續 ${targets.length} 張照片，其中 ${overwriteCount} 張已有不同內容，將被取代。要繼續嗎？（可按 Ctrl＋Z 復原）`)) return;
                let skipped = 0;
                let updated = 0;
                for (let i = index + 1; i < this.photos.length; i++) {
                    // 跡證說明欄位：若目標格已有資料，跳過不覆寫
                    if (field === 'desc' && this.photos[i][field].trim() !== '') {
                        skipped++;
                        continue;
                    }
                    if (this.photos[i][field] !== val) {
                        this.photos[i][field] = val;
                        if (field === 'time') this.photos[i].timeAutoDefaulted = false;
                        updated++;
                    }
                }
                if (field === 'time' && updated > 0 && this.photos[index]) {
                    this.photos[index].timeAutoDefaulted = false;
                }
                this.render();
                if (updated) this.recordHistory(`向下填滿${field}`);
                alert(`已更新 ${updated} 張照片的${label}。${skipped ? `保留 ${skipped} 張原有說明。` : ''}`);
            }

            fillDownFirst(field) {
                if (!this.photos.length) return;
                this.fillDownFromIndex(0, field);
            }

            // ── 時間遮罩 ────────────────────────────────────────
            // 純數字 → 自動插入冒號：4碼=HH:MM，6碼=HH:MM:SS
            // 輸入時即限制六位數，避免游標位移造成分鐘、秒數重複。
            beginMaskedEntry(input) {
                // 同一輪點擊中，pointerdown 與 focus 都可能進來；只能保存一次原值。
                if (input.dataset.maskEntryPending === 'true') return;
                input.dataset.maskOriginalValue = input.value;
                input.dataset.maskEntryPending = 'true';
                input.value = '';
            }

            // 在瀏覽器定位游標前接手滑鼠／觸控點擊，確保既有值一定被完整覆寫。
            handleMaskedPointerDown(event, input) {
                if (event.button !== undefined && event.button !== 0) return;
                event.preventDefault();
                this.beginMaskedEntry(input);
                input.focus({ preventScroll: true });
            }

            markMaskedEntryChanged(input) {
                input.dataset.maskEntryPending = 'false';
            }

            // 編輯期間只保存數字，等離開欄位後才寫入日期／時間分隔符號。
            // 這可避免 Windows 觸控鍵盤的下一個輸入事件落在剛插入的 / 或 : 前後。
            sanitizeMaskedEntry(input, maxDigits) {
                const digits = input.value.replace(/[^0-9]/g, '').slice(0, maxDigits);
                input.value = digits;
                input.selectionStart = input.selectionEnd = digits.length;
                return digits;
            }

            finishMaskedEntry(input, formatter) {
                const pending = input.dataset.maskEntryPending === 'true';
                const value = pending && !input.value ? (input.dataset.maskOriginalValue || '') : input.value;
                delete input.dataset.maskOriginalValue;
                delete input.dataset.maskEntryPending;
                return formatter.call(this, value, true);
            }

            formatTimeMask(raw, finalise = false) {
                let digits = raw.replace(/[^0-9]/g, '').slice(0, 6);

                let result = '';
                if (digits.length <= 4) {
                    // 尚未完整，只插入第一個冒號
                    result = digits.slice(0, 2);
                    if (digits.length > 2) result += ':' + digits.slice(2, 4);
                    else result = digits;
                } else {
                    // 6 碼：HH:MM:SS
                    result = digits.slice(0, 2) + ':' + digits.slice(2, 4) + ':' + digits.slice(4, 6);
                }
                return result;
            }

            // ── 日期遮罩 ────────────────────────────────────────
            // 民國年 3 碼 + 月 2 碼 + 日 2 碼 = 7 碼數字
            // 輸入 1130826 → 113/08/26
            formatDateMask(raw, finalise = false) {
                const digits = raw.replace(/[^0-9]/g, '').slice(0, 7);

                let result = '';
                if (digits.length === 0) return '';
                result = digits.slice(0, 3);
                if (digits.length > 3) result += '/' + digits.slice(3, 5);
                if (digits.length > 5) result += '/' + digits.slice(5, 7);
                return result;
            }

            // ── 日期與時間格式合法性驗證（委派至 PhotoReportValidation 模組）────
            isValidMinguoDate(raw) {
                if (typeof PhotoReportValidation !== 'undefined') {
                    return PhotoReportValidation.isValidMinguoDate(raw);
                }
                if (!raw || typeof raw !== 'string') return false;
                const digits = raw.replace(/[^0-9]/g, '');
                if (digits.length !== 7) return false;
                const year = parseInt(digits.slice(0, 3), 10);
                const month = parseInt(digits.slice(3, 5), 10);
                const day = parseInt(digits.slice(5, 7), 10);
                if (year < 1 || year > 999 || month < 1 || month > 12) return false;
                const daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
                const adYear = year + 1911;
                const isLeap = (adYear % 4 === 0 && adYear % 100 !== 0) || (adYear % 400 === 0);
                const maxDay = (month === 2 && isLeap) ? 29 : daysInMonth[month - 1];
                return day >= 1 && day <= maxDay;
            }

            isValidTimeFormat(raw) {
                if (typeof PhotoReportValidation !== 'undefined') {
                    return PhotoReportValidation.isValidTimeFormat(raw);
                }
                if (!raw || typeof raw !== 'string') return false;
                const trimmed = raw.trim();
                let h, m, s = 0;
                if (/^\d{2}:\d{2}$/.test(trimmed)) {
                    [h, m] = trimmed.split(':').map(Number);
                } else if (/^\d{2}:\d{2}:\d{2}$/.test(trimmed)) {
                    [h, m, s] = trimmed.split(':').map(Number);
                } else {
                    const digits = trimmed.replace(/[^0-9]/g, '');
                    if (digits.length === 4) {
                        h = parseInt(digits.slice(0, 2), 10);
                        m = parseInt(digits.slice(2, 4), 10);
                    } else if (digits.length === 6) {
                        h = parseInt(digits.slice(0, 2), 10);
                        m = parseInt(digits.slice(2, 4), 10);
                        s = parseInt(digits.slice(4, 6), 10);
                    } else {
                        return false;
                    }
                }
                return h >= 0 && h <= 23 && m >= 0 && m <= 59 && s >= 0 && s <= 59;
            }

            // ── 計算目前畫布的欄數（用於 Ctrl+↑↓ 整列移動）────
            getGridColumns() {
                try {
                    const style = window.getComputedStyle(this.thumbGrid);
                    const cols = style.gridTemplateColumns.split(' ').filter(s => s && s !== '0px').length;
                    return cols > 0 ? cols : 3;
                } catch {
                    return 3;
                }
            }

            // 批次套用右側編輯器內容至已選取照片
            applyEditorToSelected() {
                const selectedList = this.photos.filter(p => p.selected);
                if (selectedList.length === 0) return alert('請先選取要套用的照片。');

                const dateVal = this.editDate.value.trim();
                const timeVal = this.editTime.value.trim();
                const locVal = this.editLocation.value.trim();
                const descVal = this.editDesc.value.trim();
                const applyEmptyValues = this.batchApplyEmptyValues.checked;

                if (!applyEmptyValues && !dateVal && !timeVal && !locVal && !descVal) {
                    return alert('請先填寫要套用的日期、時間、地點或說明；若要清空，請勾選「空白欄位也套用」。');
                }

                selectedList.forEach(p => {
                    if (dateVal || applyEmptyValues) p.date = dateVal;
                    if (timeVal || applyEmptyValues) {
                        p.time = timeVal;
                        p.timeAutoDefaulted = false;
                    }
                    if (locVal || applyEmptyValues) p.location = locVal;
                    if (descVal || applyEmptyValues) p.desc = descVal;
                });

                this.render();
                this.recordHistory('批次套用編輯欄位');
                alert(`已套用到 ${selectedList.length} 張照片。`);
            }

            switchView(view) {
                const nextView = view === 'grid' ? 'grid' : 'workspace';
                this.currentView = nextView;
                this.workspaceContainer.classList.toggle('hidden', nextView !== 'workspace');
                this.gridContainer.classList.toggle('hidden', nextView !== 'grid');
                if (this.previewContainer) this.previewContainer.classList.add('hidden');
                const activeClass = "px-2.5 py-1 rounded text-xs font-bold bg-[#5B5CE2] text-white flex items-center gap-1 shadow-2xs";
                const idleClass = "px-2.5 py-1 rounded text-xs font-bold text-[#6B7280] hover:text-[#20242B] flex items-center gap-1";
                document.getElementById('btnViewWorkspace').className = nextView === 'workspace' ? activeClass : idleClass;
                document.getElementById('btnViewGrid').className = nextView === 'grid' ? activeClass : idleClass;
                const btnViewPreview = document.getElementById('btnViewPreview');
                if (btnViewPreview) btnViewPreview.className = idleClass;
                this.render();
            }

            loadWorkspacePreference(key, fallback) {
                try {
                    const value = localStorage.getItem(`photoReport.${key}`);
                    if (value === null) return fallback;
                    return typeof fallback === 'boolean' ? value === 'true' : Number(value) || fallback;
                } catch {
                    return fallback;
                }
            }

            saveWorkspacePreference(key, value) {
                try {
                    localStorage.setItem(`photoReport.${key}`, String(value));
                } catch {
                    // 隱私模式或受限制環境不儲存版面偏好，仍維持本次操作。
                }
            }

            getCaseData() {
                return Object.fromEntries(['caseTitle', 'caseDate', 'defaultLocation', 'deptName', 'officerName', 'layoutSelect']
                    .map(id => [id, document.getElementById(id).value]));
            }

            applyCaseData(data = {}) {
                ['caseTitle', 'caseDate', 'defaultLocation', 'deptName', 'officerName', 'layoutSelect'].forEach(id => {
                    if (typeof data[id] === 'string') document.getElementById(id).value = data[id];
                });
                this.syncTextLimitInputState(this.defaultLocationInput, this.defaultLocationInput?.value || '');
                this.refreshDefaultLocationLimitUi();
                this.updateLayoutDescription();
                if (this.isLightboxOpen()) this.syncLightboxAnnotationLayerBounds();
            }

            // ── 全形／半形字數計算與排版上限控制（委派至 PhotoReportValidation 模組）──
            getTextHalfWidthUnits(text) {
                if (typeof PhotoReportValidation !== 'undefined' && typeof PhotoReportValidation.getTextHalfWidthUnits === 'function') {
                    return PhotoReportValidation.getTextHalfWidthUnits(text);
                }
                let units = 0;
                for (const ch of String(text ?? '')) {
                    if (ch === '\r' || ch === '\n') continue;
                    const code = ch.codePointAt(0);
                    const isHalf = (code >= 0x0020 && code <= 0x007E) || code === 0x0009 || (code >= 0xFF61 && code <= 0xFF9F);
                    units += isHalf ? 1 : 2;
                }
                return units;
            }

            getTextFullWidthLength(text) {
                return this.getTextHalfWidthUnits(text) / 2;
            }

            formatFullWidthCount(value) {
                if (typeof PhotoReportValidation !== 'undefined' && typeof PhotoReportValidation.formatFullWidthCount === 'function') {
                    return PhotoReportValidation.formatFullWidthCount(value);
                }
                const rounded = Math.round((Number(value) || 0) * 2) / 2;
                return Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(1);
            }

            estimateDescVisualLines(text, lineFullWidth = 34) {
                if (typeof PhotoReportValidation !== 'undefined' && typeof PhotoReportValidation.estimateDescVisualLines === 'function') {
                    return PhotoReportValidation.estimateDescVisualLines(text, lineFullWidth);
                }
                const normalized = String(text ?? '').replace(/\r\n?/g, '\n');
                if (!normalized.length) return 0;
                const segments = normalized.split('\n');
                let totalLines = 0;
                for (let i = 0; i < segments.length; i++) {
                    totalLines += Math.max(1, Math.ceil(this.getTextHalfWidthUnits(segments[i]) / (lineFullWidth * 2)));
                }
                return totalLines;
            }

            truncateToFullWidth(text, maxFullWidth, options = {}) {
                if (typeof PhotoReportValidation !== 'undefined' && typeof PhotoReportValidation.truncateToFullWidth === 'function') {
                    return PhotoReportValidation.truncateToFullWidth(text, maxFullWidth, options);
                }
                const { singleLine = false, maxLines = Infinity, lineFullWidth = 34 } = options;
                const maxUnits = Math.max(0, Math.round((Number(maxFullWidth) || 0) * 2));
                let normalized = String(text ?? '').replace(/\r\n?/g, '\n');
                let truncated = false;
                if (singleLine && normalized.includes('\n')) {
                    normalized = normalized.replace(/\n+/g, '');
                    truncated = true;
                }
                let result = '';
                let usedUnits = 0;
                for (const ch of normalized) {
                    if (ch === '\n') {
                        const candidate = result + '\n';
                        if (Number.isFinite(maxLines) && this.estimateDescVisualLines(candidate, lineFullWidth) > maxLines) {
                            truncated = true;
                            continue;
                        }
                        result = candidate;
                        continue;
                    }
                    const charUnits = this.getTextHalfWidthUnits(ch);
                    if (usedUnits + charUnits > maxUnits) {
                        truncated = true;
                        break;
                    }
                    const candidate = result + ch;
                    if (Number.isFinite(maxLines) && this.estimateDescVisualLines(candidate, lineFullWidth) > maxLines) {
                        truncated = true;
                        break;
                    }
                    result = candidate;
                    usedUnits += charUnits;
                }
                return {
                    text: result,
                    halfWidthUnits: usedUnits,
                    fullWidthLength: usedUnits / 2,
                    visualLines: Number.isFinite(maxLines) ? this.estimateDescVisualLines(result, lineFullWidth) : (result ? 1 : 0),
                    truncated
                };
            }

            inspectPhotoTextLimits(effectiveLocation, desc) {
                if (typeof PhotoReportValidation !== 'undefined' && typeof PhotoReportValidation.inspectTextLimits === 'function') {
                    return PhotoReportValidation.inspectTextLimits(effectiveLocation, desc);
                }
                const locStr = String(effectiveLocation ?? '').trim();
                const descStr = String(desc ?? '').trim();
                const locationFullWidth = this.getTextFullWidthLength(locStr);
                const descFullWidth = this.getTextFullWidthLength(descStr);
                const descLines = this.estimateDescVisualLines(descStr, 34);
                const locationOver = locationFullWidth > 17;
                const descOver = descFullWidth > 68 || descLines > 2;
                return {
                    locationFullWidth,
                    descFullWidth,
                    descLines,
                    locationOver,
                    descOver,
                    isOver: locationOver || descOver
                };
            }

            getTextLimitInputState(input) {
                if (!input) return null;
                if (!this.textLimitInputStates) this.textLimitInputStates = new WeakMap();
                let state = this.textLimitInputStates.get(input);
                if (!state) {
                    state = {
                        isComposing: false,
                        lastCommittedValue: String(input.value ?? ''),
                        truncatedFlash: false
                    };
                    this.textLimitInputStates.set(input, state);
                }
                return state;
            }

            syncTextLimitInputState(input, value) {
                if (!input) return;
                const state = this.getTextLimitInputState(input);
                if (!state) return;
                state.lastCommittedValue = String(value ?? '');
                state.truncatedFlash = false;
            }

            triggerTextLimitTruncatedFlash(input, onClear) {
                if (!input) return;
                const state = this.getTextLimitInputState(input);
                if (state) state.truncatedFlash = true;
                if (!this.textLimitFeedbackTimers) this.textLimitFeedbackTimers = new WeakMap();
                const prevTimer = this.textLimitFeedbackTimers.get(input);
                if (prevTimer) clearTimeout(prevTimer);
                const timer = setTimeout(() => {
                    const curState = this.getTextLimitInputState(input);
                    if (curState) curState.truncatedFlash = false;
                    if (typeof onClear === 'function') onClear();
                }, 2600);
                this.textLimitFeedbackTimers.set(input, timer);
            }

            resolveLimitedTextInput(input, { maxFullWidth, singleLine = false, maxLines = Infinity, lineFullWidth = 34 }) {
                const state = this.getTextLimitInputState(input);
                const prevValue = state ? state.lastCommittedValue : String(input.value ?? '');
                let rawValue = String(input.value ?? '').replace(/\r\n?/g, '\n');
                let strippedNewline = false;
                if (singleLine && rawValue.includes('\n')) {
                    rawValue = rawValue.replace(/\n+/g, '');
                    strippedNewline = true;
                }

                const maxUnits = Math.round((Number(maxFullWidth) || 0) * 2);
                const prevUnits = this.getTextHalfWidthUnits(prevValue);
                const prevLines = Number.isFinite(maxLines) ? this.estimateDescVisualLines(prevValue, lineFullWidth) : 1;
                const prevOver = prevUnits > maxUnits || prevLines > maxLines;

                const rawUnits = this.getTextHalfWidthUnits(rawValue);
                const rawLines = Number.isFinite(maxLines) ? this.estimateDescVisualLines(rawValue, lineFullWidth) : (rawValue ? 1 : 0);
                const rawOver = rawUnits > maxUnits || rawLines > maxLines;

                let finalValue = rawValue;
                let blockedOrTruncated = strippedNewline;

                if (!rawOver) {
                    finalValue = rawValue;
                } else if (prevOver && rawUnits <= prevUnits && rawLines <= prevLines && rawValue.length <= prevValue.length) {
                    // 允許從舊檔帶入的過長文字逐步刪減，不直接暴力截斷使用者原本文字
                    finalValue = rawValue;
                } else if (prevOver) {
                    // 已處於舊資料超長狀態時，不允許再新增字元或換行
                    finalValue = prevValue;
                    blockedOrTruncated = true;
                } else {
                    const res = this.truncateToFullWidth(rawValue, maxFullWidth, {
                        singleLine,
                        maxLines,
                        lineFullWidth
                    });
                    finalValue = res.text;
                    blockedOrTruncated = blockedOrTruncated || res.truncated;
                }

                if (input.value !== finalValue) {
                    const selStart = typeof input.selectionStart === 'number' ? input.selectionStart : finalValue.length;
                    input.value = finalValue;
                    const clampedCursor = Math.max(0, Math.min(selStart, finalValue.length));
                    if (typeof input.setSelectionRange === 'function') {
                        try {
                            input.setSelectionRange(clampedCursor, clampedCursor);
                        } catch {
                            // 略過部分輸入元件不支援 setSelectionRange 的情況
                        }
                    }
                }

                if (state) state.lastCommittedValue = finalValue;
                return { finalValue, blockedOrTruncated };
            }

            bindLimitedTextControl(input, options) {
                if (!input) return;
                this.syncTextLimitInputState(input, input.value || '');
                const commit = (forceAccept = false) => {
                    const state = this.getTextLimitInputState(input);
                    if (!forceAccept && state.lastCommittedValue === String(input.value ?? '') && !state.truncatedFlash) {
                        options.onFeedback?.(input.value);
                        return;
                    }
                    const { finalValue, blockedOrTruncated } = this.resolveLimitedTextInput(input, options);
                    if (blockedOrTruncated) {
                        this.triggerTextLimitTruncatedFlash(input, () => options.onFeedback?.(input.value));
                    } else if (state) {
                        state.truncatedFlash = false;
                    }
                    options.onAccept?.(finalValue);
                    options.onFeedback?.(finalValue);
                };

                input.addEventListener('compositionstart', () => {
                    const state = this.getTextLimitInputState(input);
                    if (state) state.isComposing = true;
                });
                input.addEventListener('compositionend', () => {
                    const state = this.getTextLimitInputState(input);
                    if (state) state.isComposing = false;
                    commit(true);
                });
                input.addEventListener('input', (event) => {
                    const state = this.getTextLimitInputState(input);
                    if (event?.isComposing || state?.isComposing) {
                        options.onFeedback?.(input.value);
                        return;
                    }
                    commit(true);
                });
            }

            bindGridTextLimitEvents() {
                if (!this.gridTbody) return;
                const getGridInputMeta = (target) => {
                    if (!(target instanceof HTMLInputElement)) return null;
                    const field = target.dataset.gridField;
                    const idx = Number(target.dataset.gridIndex);
                    if ((field !== 'location' && field !== 'desc') || !Number.isInteger(idx) || !this.photos[idx]) return null;
                    return {
                        input: target,
                        field,
                        idx,
                        options: field === 'location'
                            ? { maxFullWidth: 17, singleLine: true }
                            : { maxFullWidth: 68, singleLine: false, maxLines: 2, lineFullWidth: 34 }
                    };
                };

                const commitGridInput = (meta) => {
                    const { input, field, idx, options } = meta;
                    const { finalValue, blockedOrTruncated } = this.resolveLimitedTextInput(input, options);
                    const state = this.getTextLimitInputState(input);
                    if (blockedOrTruncated) {
                        this.triggerTextLimitTruncatedFlash(input, () => {
                            input.title = field === 'location'
                                ? '地點最多 17 個全形（2 半形＝1 全形）'
                                : '說明最多 68 個全形・最多 2 行（2 半形＝1 全形）';
                        });
                        input.title = field === 'location'
                            ? '已達上限：地點最多 17 個全形（2 半形＝1 全形）'
                            : '已達上限：說明最多 68 個全形（2 半形＝1 全形），避免變成 3 行';
                    } else if (state) {
                        state.truncatedFlash = false;
                    }

                    const fullLen = this.getTextFullWidthLength(finalValue);
                    const lines = field === 'desc' ? this.estimateDescVisualLines(finalValue, 34) : 1;
                    const isOver = field === 'location'
                        ? fullLen > 17
                        : (fullLen > 68 || lines > 2);
                    input.classList.toggle('input-over-limit', isOver || Boolean(state?.truncatedFlash));
                    this.updatePhotoField(idx, field, finalValue);
                };

                this.gridTbody.addEventListener('compositionstart', (e) => {
                    const meta = getGridInputMeta(e.target);
                    if (!meta) return;
                    const state = this.getTextLimitInputState(meta.input);
                    if (state) state.isComposing = true;
                });
                this.gridTbody.addEventListener('compositionend', (e) => {
                    const meta = getGridInputMeta(e.target);
                    if (!meta) return;
                    const state = this.getTextLimitInputState(meta.input);
                    if (state) state.isComposing = false;
                    commitGridInput(meta);
                });
                this.gridTbody.addEventListener('input', (e) => {
                    const meta = getGridInputMeta(e.target);
                    if (!meta) return;
                    const state = this.getTextLimitInputState(meta.input);
                    if (e?.isComposing || state?.isComposing) return;
                    commitGridInput(meta);
                });
            }

            refreshDefaultLocationLimitUi(draftValue) {
                const input = this.defaultLocationInput;
                const counter = this.defaultLocationCounter;
                const hint = this.defaultLocationLimitHint;
                if (!input || !counter || !hint) return;
                const value = draftValue !== undefined ? String(draftValue) : String(input.value || '');
                const fullLen = this.getTextFullWidthLength(value);
                const state = this.getTextLimitInputState(input);
                const isOver = fullLen > 17;
                const isFlash = Boolean(state?.truncatedFlash);
                const isNear = !isOver && !isFlash && fullLen >= 15;

                counter.textContent = `${this.formatFullWidthCount(fullLen)} / 17 全形`;
                counter.classList.toggle('is-over-limit', isOver || isFlash);
                counter.classList.toggle('is-near-limit', isNear);
                input.classList.toggle('input-over-limit', isOver);

                hint.classList.toggle('is-over-limit', isOver || isFlash);
                hint.classList.toggle('is-near-limit', isNear);
                if (isOver) {
                    hint.textContent = `目前 ${this.formatFullWidthCount(fullLen)} 全形，已超出 17 全形上限（建議精簡）`;
                } else if (isFlash) {
                    hint.textContent = '已達上限：地點最多 17 個全形（超出部分未寫入）';
                } else if (isNear) {
                    hint.textContent = `剩餘 ${this.formatFullWidthCount(17 - fullLen)} 全形（2 半形＝1 全形）`;
                } else {
                    hint.textContent = '最多 17 個全形（2 半形＝1 全形）';
                }
            }

            refreshEditLocationLimitUi(draftValue) {
                const input = this.editLocation;
                const counter = this.editLocationCounter;
                const hint = this.editLocationLimitHint;
                if (!input || !counter || !hint) return;
                const value = draftValue !== undefined ? String(draftValue) : String(input.value || '');
                const fullLen = this.getTextFullWidthLength(value);
                const defaultLoc = String(this.defaultLocationInput?.value || '').trim();
                const inheritedLen = !value.trim() ? this.getTextFullWidthLength(defaultLoc) : 0;
                const state = this.getTextLimitInputState(input);
                const isOver = fullLen > 17;
                const isInheritedOver = !value.trim() && inheritedLen > 17;
                const isFlash = Boolean(state?.truncatedFlash);
                const isNear = !isOver && !isFlash && fullLen >= 15;

                counter.textContent = `${this.formatFullWidthCount(fullLen)} / 17 全形`;
                counter.classList.toggle('is-over-limit', isOver || isFlash || isInheritedOver);
                counter.classList.toggle('is-near-limit', isNear);
                input.classList.toggle('input-over-limit', isOver || isInheritedOver);

                hint.classList.toggle('is-over-limit', isOver || isFlash || isInheritedOver);
                hint.classList.toggle('is-near-limit', isNear);
                if (isOver) {
                    hint.textContent = `目前 ${this.formatFullWidthCount(fullLen)} 全形，已超出 17 全形上限（請精簡）`;
                } else if (isInheritedOver) {
                    hint.textContent = `繼承的案件地點有 ${this.formatFullWidthCount(inheritedLen)} 全形，已超出 17 全形上限`;
                } else if (isFlash) {
                    hint.textContent = '已達上限：地點最多 17 個全形（超出部分未寫入）';
                } else if (isNear) {
                    hint.textContent = `剩餘 ${this.formatFullWidthCount(17 - fullLen)} 全形（2 半形＝1 全形）`;
                } else {
                    hint.textContent = '最多 17 個全形（2 半形＝1 全形）';
                }
            }

            refreshEditDescLimitUi(draftValue) {
                const input = this.editDesc;
                const counter = this.editDescCounter;
                const hint = this.editDescLimitHint;
                if (!input || !counter || !hint) return;
                const value = draftValue !== undefined ? String(draftValue) : String(input.value || '');
                const fullLen = this.getTextFullWidthLength(value);
                const visualLines = this.estimateDescVisualLines(value, 34);
                const state = this.getTextLimitInputState(input);
                const isOver = fullLen > 68 || visualLines > 2;
                const isFlash = Boolean(state?.truncatedFlash);
                const isNear = !isOver && !isFlash && (fullLen >= 62 || (visualLines === 2 && fullLen >= 55));

                counter.textContent = `${this.formatFullWidthCount(fullLen)} / 68 全形`;
                counter.classList.toggle('is-over-limit', isOver || isFlash);
                counter.classList.toggle('is-near-limit', isNear);
                input.classList.toggle('input-over-limit', isOver);

                hint.classList.toggle('is-over-limit', isOver || isFlash);
                hint.classList.toggle('is-near-limit', isNear);
                if (isOver) {
                    hint.textContent = fullLen > 68
                        ? `目前 ${this.formatFullWidthCount(fullLen)} 全形（預估 ${visualLines} 行），超過 68 全形恐變 3 行影響整頁排版`
                        : `目前預估 ${visualLines} 行，超過 2 行恐影響整頁排版（上限 68 全形・最多 2 行）`;
                } else if (isFlash) {
                    hint.textContent = '已達上限（最多 68 全形／2 行），避免變 3 行影響排版';
                } else if (isNear) {
                    hint.textContent = `剩餘 ${this.formatFullWidthCount(68 - fullLen)} 全形（第 ${visualLines || 1}／2 行）`;
                } else {
                    hint.textContent = '最多 68 個全形・最多 2 行（2 半形＝1 全形）';
                }
            }

            refreshAllTextLimitUi() {
                this.refreshDefaultLocationLimitUi();
                this.refreshEditLocationLimitUi();
                this.refreshEditDescLimitUi();
            }

            refreshThumbCardStatusBadge(index) {
                const photo = this.photos[index];
                if (!photo) return;
                const card = document.getElementById(`thumb_${photo.uid}`);
                if (!card) return;
                const defaultLocation = String(this.defaultLocationInput?.value || '').trim();
                const missingFields = [!String(photo.location || '').trim() && !defaultLocation ? '地點' : '', !String(photo.desc || '').trim() ? '說明' : ''].filter(Boolean);
                const overLimit = this.inspectPhotoTextLimits(String(photo.location || '').trim() || defaultLocation, String(photo.desc || '').trim()).isOver;
                const statusIssues = [...missingFields, ...(overLimit ? ['字數過長'] : [])];
                const statusText = statusIssues.length ? `注意：${statusIssues.join('、')}` : '資料完整';
                const badge = card.querySelector('.thumb-status-icon');
                if (badge) {
                    badge.className = `thumb-status-icon text-[13px] ${statusIssues.length ? 'text-amber-700' : 'text-emerald-700'}`;
                    badge.title = statusText;
                    const icon = badge.querySelector('i');
                    if (icon) {
                        icon.className = `fa-solid ${statusIssues.length ? 'fa-circle-exclamation' : 'fa-circle-check'}`;
                    }
                }
                card.title = `第 ${photo.seq} 張｜時間：${photo.time || '未填'}｜${statusText}｜雙擊縮圖放大、雙擊編號列編輯，拖曳可排序`;
            }

            cloneAnnotations(annotations) {
                if (!Array.isArray(annotations)) return [];
                return annotations.map(ann => ({ ...ann }));
            }

            cloneSideCrop(sideCrop) {
                if (!sideCrop || typeof sideCrop !== 'object' || !sideCrop.cropRect) return null;
                return {
                    cropRect: {
                        x: Number(sideCrop.cropRect.x) || 0,
                        y: Number(sideCrop.cropRect.y) || 0,
                        w: Number(sideCrop.cropRect.w) || 0.2,
                        h: Number(sideCrop.cropRect.h) || 0.2
                    },
                    widthRatio: Number(sideCrop.widthRatio) || 0.65,
                    offsetY: Number.isFinite(Number(sideCrop.offsetY)) ? Number(sideCrop.offsetY) : 0.5,
                    borderColor: sideCrop.borderColor === 'red' ? 'red' : sideCrop.borderColor === 'none' ? 'none' : 'yellow',
                    sourceAnnId: sideCrop.sourceAnnId || null
                };
            }

            snapshotState(label = '') {
                return {
                    label,
                    caseData: this.getCaseData(),
                    currentIndex: this.currentIndex,
                    lastSelectedIndex: this.lastSelectedIndex,
                    photos: this.photos.map(photo => ({
                        ...photo,
                        flipH: Boolean(photo.flipH),
                        parentUid: photo.parentUid || null,
                        cropRect: photo.cropRect ? { ...photo.cropRect } : null,
                        sideCrop: this.cloneSideCrop(photo.sideCrop),
                        compositeThumbnailSrc: null,
                        _compositeKey: '',
                        annotations: this.cloneAnnotations(photo.annotations)
                    }))
                };
            }

            historySignature(state) {
                if (typeof PhotoReportHistory !== 'undefined') {
                    return PhotoReportHistory.historySignature(state);
                }
                return JSON.stringify({
                    caseData: state.caseData,
                    currentIndex: state.currentIndex,
                    lastSelectedIndex: state.lastSelectedIndex,
                    photos: state.photos.map(({ uid, rotation, flipH, parentUid, cropRect, sideCrop, annotations, seq, date, time, location, desc, selected, stageX, stageY }) =>
                        ({ uid, rotation, flipH: Boolean(flipH), parentUid: parentUid || null, cropRect: cropRect || null, sideCrop: sideCrop || null, annotations: Array.isArray(annotations) ? annotations : [], seq, date, time, location, desc, selected, stageX, stageY }))
                });
            }

            isExternalFileDrag(event) {
                return Array.from(event.dataTransfer?.types || []).includes('Files');
            }

            bindExternalFileDropTarget(target, overlayTarget = null) {
                const setDragOver = (active) => {
                    target.classList.toggle('external-file-drag-over', active);
                    if (overlayTarget) overlayTarget.classList.toggle('external-file-drag-over', active);
                };
                if (!target) return;
                target.addEventListener('dragenter', (event) => {
                    if (!this.isExternalFileDrag(event)) return;
                    event.preventDefault();
                    setDragOver(true);
                });
                target.addEventListener('dragover', (event) => {
                    if (!this.isExternalFileDrag(event)) return;
                    event.preventDefault();
                    event.dataTransfer.dropEffect = 'copy';
                    setDragOver(true);
                });
                target.addEventListener('dragleave', (event) => {
                    if (!this.isExternalFileDrag(event) || target.contains(event.relatedTarget)) return;
                    setDragOver(false);
                });
                target.addEventListener('drop', async (event) => {
                    if (!this.isExternalFileDrag(event)) return;
                    event.preventDefault();
                    setDragOver(false);
                    const droppedFiles = Array.from(event.dataTransfer?.files || []);
                    const projectFiles = droppedFiles.filter(file => /\.photo-report$/i.test(file.name || ''));
                    if (projectFiles.length === 1) {
                        let fileHandle = null;
                        try {
                            const items = Array.from(event.dataTransfer?.items || []);
                            const firstItem = items.find(item => item.kind === 'file');
                            if (firstItem && typeof firstItem.getAsFileSystemHandle === 'function') {
                                const handle = await firstItem.getAsFileSystemHandle();
                                if (handle && handle.kind === 'file' && /\.photo-report$/i.test(handle.name || '')) {
                                    fileHandle = handle;
                                }
                            }
                        } catch {
                            // 無法取得檔案控制代碼時退回一般檔案開啟
                        }
                        await this.openProjectFromFile(projectFiles[0], fileHandle);
                        return;
                    }
                    if (projectFiles.length >= 2) {
                        await this.mergeProjectFiles(projectFiles);
                        return;
                    }
                    const files = droppedFiles.filter(file => file.type.startsWith('image/'));
                    if (!files.length) return alert('請拖入支援的圖片檔案或 .photo-report 工作進度檔。');
                    this.handleFiles(files).catch(error => this.handleProcessingError(error));
                });
            }

            projectSignature() {
                if (typeof PhotoReportHistory !== 'undefined') {
                    return PhotoReportHistory.projectSignature(this.getCaseData(), this.photos);
                }
                return JSON.stringify({
                    caseData: this.getCaseData(),
                    photos: this.photos.map(({ uid, rotation, flipH, parentUid, cropRect, sideCrop, annotations, seq, date, time, location, desc, stageX, stageY }) =>
                        ({ uid, rotation, flipH: Boolean(flipH), parentUid: parentUid || null, cropRect: cropRect || null, sideCrop: sideCrop || null, annotations: Array.isArray(annotations) ? annotations : [], seq, date, time, location, desc, stageX, stageY }))
                });
            }

            getPhotoContentSignature(photo) {
                if (!photo) return '';
                return JSON.stringify({
                    rotation: Number(photo.rotation) || 0,
                    flipH: Boolean(photo.flipH),
                    annotations: Array.isArray(photo.annotations) ? photo.annotations : [],
                    sideCrop: photo.sideCrop || null,
                    date: photo.date || '',
                    time: photo.time || '',
                    location: photo.location || '',
                    desc: photo.desc || ''
                });
            }

            markProjectSaved(savedTs = null) {
                this.savedProjectSignature = this.projectSignature();
                this.savedPhotosByUid = new Map(
                    this.photos.map(photo => [photo.uid, this.getPhotoContentSignature(photo)])
                );
                if (savedTs === 'clear') {
                    this.currentProjectSavedAt = null;
                } else if (savedTs !== null && Number.isFinite(savedTs)) {
                    this.currentProjectSavedAt = savedTs;
                } else if (this.photos.length > 0 && !this.currentProjectSavedAt) {
                    this.currentProjectSavedAt = Date.now();
                }
                this.lastLocalEditAt = null;
                this.updateProjectSaveStatusUi();
            }

            hasUnsavedProjectChanges() {
                return this.projectSignature() !== this.savedProjectSignature;
            }

            getProjectTimestamp(manifest, file) {
                const manifestTs = manifest?.savedAt ? Date.parse(manifest.savedAt) : NaN;
                if (Number.isFinite(manifestTs) && manifestTs > 0) return manifestTs;
                const fileTs = Number(file?.lastModified);
                if (Number.isFinite(fileTs) && fileTs > 0) return fileTs;
                return 0;
            }

            getCurrentWorkspaceTimestamp() {
                if (this.hasUnsavedProjectChanges()) {
                    return Math.max(this.lastLocalEditAt || 0, this.currentProjectSavedAt || 0, Date.now());
                }
                return this.currentProjectSavedAt || 0;
            }

            updateProjectSaveStatusUi() {
                if (!this.btnQuickSaveStatus || !this.quickSaveStatusText || !this.quickSaveStatusIcon) return;
                const hasPhotos = this.photos.length > 0;
                const unsaved = this.hasUnsavedProjectChanges();
                if (!hasPhotos && !unsaved) {
                    this.btnQuickSaveStatus.classList.add('hidden');
                    if (this.saveProjectMenuText) {
                        this.saveProjectMenuText.textContent = '儲存工作進度（Ctrl+S）';
                    }
                    return;
                }
                this.btnQuickSaveStatus.classList.remove('hidden');
                this.btnQuickSaveStatus.classList.toggle('project-save-badge--unsaved', unsaved);
                this.btnQuickSaveStatus.classList.toggle('project-save-badge--saved', !unsaved);
                if (unsaved) {
                    this.quickSaveStatusIcon.className = 'fa-solid fa-circle-exclamation';
                    const autoDraftHint = this.lastAutoDraftSavedAt ? '・已自動暫存' : '';
                    this.quickSaveStatusText.textContent = `還有修改沒存（Ctrl+S）${autoDraftHint}`;
                    this.btnQuickSaveStatus.title = this.currentProjectFileName
                        ? `點擊直接儲存至「${this.currentProjectFileName}」（快速鍵：Ctrl+S）`
                        : '點擊儲存工作進度檔（快速鍵：Ctrl+S）';
                } else {
                    this.quickSaveStatusIcon.className = 'fa-solid fa-circle-check';
                    this.quickSaveStatusText.textContent = this.currentProjectFileName
                        ? `已儲存：${this.currentProjectFileName}`
                        : '工作進度已儲存';
                    this.btnQuickSaveStatus.title = this.currentProjectFileName
                        ? `目前已儲存至「${this.currentProjectFileName}」（按 Ctrl+S 可再次儲存）`
                        : '目前工作進度已儲存（按 Ctrl+S 可再次儲存）';
                }
                if (this.saveProjectMenuText) {
                    this.saveProjectMenuText.textContent = (this.currentProjectFileHandle && this.currentProjectFileName)
                        ? `儲存至「${this.currentProjectFileName}」（Ctrl+S）`
                        : '儲存工作進度（Ctrl+S）';
                }
            }

            recordHistory(label = '') {
                if (this.isRestoringHistory) return;
                const next = this.snapshotState(label);
                const current = this.history[this.historyIndex];
                if (current && this.historySignature(current) === this.historySignature(next)) {
                    this.updateProjectSaveStatusUi();
                    return;
                }
                this.history.splice(this.historyIndex + 1);
                this.history.push(next);
                if (this.history.length > this.historyLimit) this.history.shift();
                this.historyIndex = this.history.length - 1;
                this.releaseUnusedPhotoUrls();
                this.updateHistoryControls();
                const nonEditLabels = ['初始狀態', '開啟專案', '清空照片', '變更照片選取', '切換目前照片', '框選照片'];
                if (!nonEditLabels.includes(label) && this.hasUnsavedProjectChanges()) {
                    this.lastLocalEditAt = Date.now();
                }
                this.updateProjectSaveStatusUi();
                if (this.photos.length > 0) {
                    this.hadPhotosInSession = true;
                    this.hideAutoDraftBanner();
                    this.scheduleAutoDraftSave();
                }
            }

            scheduleHistory(label) {
                if (this.isRestoringHistory) return;
                if (this.hasUnsavedProjectChanges()) {
                    this.lastLocalEditAt = Date.now();
                }
                this.updateProjectSaveStatusUi();
                if (this.photos.length > 0) {
                    this.scheduleAutoDraftSave();
                }
                clearTimeout(this.historyTimer);
                this.pendingHistoryLabel = label || '';
                this.historyTimer = setTimeout(() => {
                    this.historyTimer = null;
                    const pendingLabel = this.pendingHistoryLabel;
                    this.pendingHistoryLabel = '';
                    this.recordHistory(pendingLabel);
                }, 450);
            }

            flushPendingHistory() {
                if (!this.historyTimer) return;
                clearTimeout(this.historyTimer);
                this.historyTimer = null;
                const pendingLabel = this.pendingHistoryLabel || '修改資料';
                this.pendingHistoryLabel = '';
                this.recordHistory(pendingLabel);
            }

            restoreHistory(index) {
                const state = this.history[index];
                if (!state) return;
                clearTimeout(this.historyTimer);
                this.historyTimer = null;
                this.pendingHistoryLabel = '';
                this.isRestoringHistory = true;
                this.photos = state.photos.map(photo => ({
                    ...photo,
                    flipH: Boolean(photo.flipH),
                    parentUid: photo.parentUid || null,
                    cropRect: photo.cropRect ? { ...photo.cropRect } : null,
                    sideCrop: this.cloneSideCrop(photo.sideCrop),
                    compositeThumbnailSrc: null,
                    _compositeKey: '',
                    annotations: this.cloneAnnotations(photo.annotations)
                }));
                this.currentIndex = state.currentIndex;
                this.lastSelectedIndex = state.lastSelectedIndex;
                this.applyCaseData(state.caseData);
                this.isRestoringHistory = false;
                this.historyIndex = index;
                this.render();
                this.syncAllCompositeThumbnails();
                this.updateHistoryControls();
                if (this.isLightboxOpen() && this.photos.length > 0) {
                    this.lightboxIndex = Math.max(0, Math.min(this.photos.length - 1, this.currentIndex));
                    this.updateLightbox({ preserveZoomIfSamePhoto: true });
                }
            }

            undo() {
                this.flushPendingHistory();
                if (this.historyIndex <= 0) return;
                this.restoreHistory(this.historyIndex - 1);
            }

            redo() {
                this.flushPendingHistory();
                if (this.historyIndex >= this.history.length - 1) return;
                this.restoreHistory(this.historyIndex + 1);
            }

            updateHistoryControls() {
                const undo = document.getElementById('btnUndo');
                const redo = document.getElementById('btnRedo');
                undo.disabled = this.historyIndex <= 0;
                redo.disabled = this.historyIndex >= this.history.length - 1;
            }

            getPhotosPerPage() {
                return this.layoutSelect.value === 'landscape_3' ? 3 : 2;
            }

            sanitizeFilename(rawName, fallback = '現場照片清冊') {
                const cleaned = String(rawName ?? '')
                    .replace(/[\\/:*?"<>|\u0000-\u001F]/g, '_')
                    .replace(/\s+/g, ' ')
                    .trim();
                return cleaned || fallback;
            }

            finalizeActiveInputsBeforeSave() {
                const active = document.activeElement;
                if (active instanceof HTMLInputElement) {
                    if (
                        active.dataset.maskOriginalValue !== undefined ||
                        active.dataset.maskEntryPending !== undefined ||
                        active.classList.contains('lightbox-inline-label-input')
                    ) {
                        active.blur();
                    }
                }
                this.commitLightboxHistoryIfNeeded();
                this.flushPendingHistory();
            }

            async createLitePhotoEntry(photo, maxDimension = 2000, quality = 0.88) {
                const fallback = {
                    blob: photo.file,
                    type: photo.file?.type || 'image/jpeg',
                    exifRotation: photo.exifRotation || 0
                };
                if (!photo || !photo.file) return fallback;
                let source = null;
                let img = null;
                try {
                    img = await this.loadImage(photo.src);
                    let exifRotation = 0;
                    if (photo.exifRotation && typeof createImageBitmap === 'function') {
                        source = await createImageBitmap(photo.file, { imageOrientation: 'none' });
                        exifRotation = photo.exifRotation;
                    } else {
                        source = img;
                    }
                    let w = source.width || img.naturalWidth || img.width;
                    let h = source.height || img.naturalHeight || img.height;
                    if (!w || !h) return fallback;

                    if (w <= maxDimension && h <= maxDimension && !exifRotation && photo.file.size <= 350 * 1024) {
                        return fallback;
                    }

                    const origW = w;
                    const origH = h;
                    if (w > maxDimension || h > maxDimension) {
                        if (w >= h) {
                            h = Math.max(1, Math.round((h * maxDimension) / w));
                            w = maxDimension;
                        } else {
                            w = Math.max(1, Math.round((w * maxDimension) / h));
                            h = maxDimension;
                        }
                    }

                    const rot = ((exifRotation % 360) + 360) % 360;
                    const canvas = document.createElement('canvas');
                    if (rot === 90 || rot === 270) {
                        canvas.width = h;
                        canvas.height = w;
                    } else {
                        canvas.width = w;
                        canvas.height = h;
                    }
                    const ctx = canvas.getContext('2d');
                    if (!ctx) return fallback;
                    ctx.fillStyle = '#FFFFFF';
                    ctx.fillRect(0, 0, canvas.width, canvas.height);
                    ctx.save();
                    ctx.translate(canvas.width / 2, canvas.height / 2);
                    if (rot) ctx.rotate((rot * Math.PI) / 180);
                    ctx.drawImage(source, -w / 2, -h / 2, w, h);
                    ctx.restore();

                    const liteBlob = await new Promise((resolve, reject) => canvas.toBlob(
                        result => result ? resolve(result) : reject(new Error('無法建立輕量照片')),
                        'image/jpeg',
                        quality
                    ));
                    if (liteBlob.size >= photo.file.size && origW <= maxDimension && origH <= maxDimension && !exifRotation) {
                        return fallback;
                    }
                    return {
                        blob: liteBlob,
                        type: 'image/jpeg',
                        exifRotation: 0
                    };
                } catch {
                    return fallback;
                } finally {
                    if (source && source !== img && typeof source.close === 'function') source.close();
                }
            }

            async saveProject(options = {}) {
                if (!this.photos.length) return alert('還沒有照片，不需要儲存。');
                if (typeof JSZip === 'undefined') throw new Error('JSZip 尚未載入，無法儲存專案。');

                const lite = options?.lite !== undefined ? Boolean(options.lite) : true;
                const saveAsNew = Boolean(options?.saveAsNew);
                const fromUserPicker = Boolean(options?.fromUserPicker);

                this.finalizeActiveInputsBeforeSave();
                this.processingCancelled = false;

                const caseData = this.getCaseData();
                const safeTitle = this.sanitizeFilename(caseData.caseTitle, '現場照片清冊');
                const suffix = '_照片清冊專案.photo-report';
                const suggestedName = (!saveAsNew && this.currentProjectFileHandle && this.currentProjectFileName)
                    ? this.currentProjectFileName
                    : `${safeTitle}${suffix}`;

                const isDefaultSaveAs = !this.initialSaveAs || window.saveAs === this.initialSaveAs;
                let targetHandle = (!saveAsNew && isDefaultSaveAs && this.currentProjectFileHandle)
                    ? this.currentProjectFileHandle
                    : null;

                if (!targetHandle && fromUserPicker && isDefaultSaveAs && window.isSecureContext && typeof window.showSaveFilePicker === 'function' && !navigator.webdriver) {
                    try {
                        targetHandle = await window.showSaveFilePicker({
                            suggestedName,
                            types: [{
                                description: '現場照片清冊工作進度檔 (.photo-report)',
                                accept: { 'application/zip': ['.photo-report'] }
                            }]
                        });
                    } catch (pickerError) {
                        if (pickerError?.name === 'AbortError') return false;
                        targetHandle = null;
                    }
                }

                try {
                    const zip = new JSZip();
                    const manifestPhotos = [];
                    const progressText = '正在最佳化圖片並封裝工作進度…';
                    this.setProgress(progressText, 0, this.photos.length);

                    for (let index = 0; index < this.photos.length; index++) {
                        this.throwIfProcessingCancelled();
                        const photo = this.photos[index];
                        const safeArchivePhotoName = this.sanitizeFilename(photo.name, `photo_${index + 1}.jpg`);
                        const archiveName = `${String(index + 1).padStart(4, '0')}_${photo.uid}_${safeArchivePhotoName}`;

                        let entryBlob = photo.file;
                        let entryType = photo.file?.type || 'image/jpeg';
                        let entryExifRotation = photo.exifRotation || 0;

                        if (lite) {
                            const liteEntry = await this.createLitePhotoEntry(photo, 2000, 0.88);
                            this.throwIfProcessingCancelled();
                            entryBlob = liteEntry.blob;
                            entryType = liteEntry.type;
                            entryExifRotation = liteEntry.exifRotation;
                        }

                        let photoFp = photo.visualFingerprint || '';
                        if (!photoFp && photo.file) {
                            const calcFp = await this.computeVisualFingerprintFromBlob(photo.file);
                            photoFp = calcFp.visualFingerprint || '';
                            if (photoFp) photo.visualFingerprint = photoFp;
                        }

                        // 照片本身已是壓縮格式，一律以 STORE 直接封裝，避免重複 DEFLATE 壓縮造成嚴重卡頓
                        zip.file(`images/${archiveName}`, entryBlob, { compression: 'STORE' });
                        manifestPhotos.push({
                            uid: photo.uid,
                            archiveName,
                            name: photo.name,
                            type: entryType,
                            visualFingerprint: photoFp,
                            thumbnailAspect: Number(photo.thumbnailAspect) || 1,
                            rotation: photo.rotation,
                            flipH: Boolean(photo.flipH),
                            parentUid: photo.parentUid || null,
                            cropRect: photo.cropRect ? { ...photo.cropRect } : null,
                            sideCrop: this.cloneSideCrop(photo.sideCrop),
                            annotations: this.cloneAnnotations(photo.annotations),
                            exifRotation: entryExifRotation,
                            seq: photo.seq,
                            date: photo.date,
                            time: photo.time,
                            timeAutoDefaulted: Boolean(photo.timeAutoDefaulted),
                            location: photo.location,
                            desc: photo.desc,
                            selected: photo.selected,
                            stageX: Number.isFinite(photo.stageX) ? photo.stageX : null,
                            stageY: Number.isFinite(photo.stageY) ? photo.stageY : null
                        });
                        this.setProgress(progressText, index + 1, this.photos.length);
                        this.throwIfProcessingCancelled();
                        if ((index + 1) % (lite ? 4 : 10) === 0) await this.yieldToBrowser();
                    }

                    this.throwIfProcessingCancelled();
                    const savedAtIso = new Date().toISOString();
                    const savedAtMs = Date.parse(savedAtIso);
                    zip.file('manifest.json', JSON.stringify({
                        format: 'photo-report-project',
                        version: 1,
                        liteMode: lite,
                        savedAt: savedAtIso,
                        caseData: this.getCaseData(),
                        currentIndex: this.currentIndex,
                        lastSelectedIndex: this.lastSelectedIndex,
                        photos: manifestPhotos
                    }), { compression: 'DEFLATE', compressionOptions: { level: 6 } });

                    const content = await zip.generateAsync({ type: 'blob' });
                    this.throwIfProcessingCancelled();

                    if (targetHandle && typeof targetHandle.createWritable === 'function') {
                        const writable = await targetHandle.createWritable();
                        await writable.write(content);
                        await writable.close();
                        this.currentProjectFileHandle = targetHandle;
                        this.currentProjectFileName = targetHandle.name || suggestedName;
                    } else {
                        saveAs(content, suggestedName);
                        this.currentProjectFileName = suggestedName;
                    }

                    this.markProjectSaved(savedAtMs);
                    this.scheduleAutoDraftSave();
                    return true;
                } finally {
                    this.clearProgress();
                }
            }

            async openProject(event) {
                const file = event?.target?.files?.[0];
                if (event?.target) event.target.value = '';
                if (!file) return false;
                return this.openProjectFromFile(file, null);
            }

            async openProjectFromFile(file, fileHandle = null) {
                if (!file) return false;
                if (typeof JSZip === 'undefined') {
                    this.handleProjectError(new Error('JSZip 尚未載入，無法開啟專案。'), 'open');
                    return false;
                }
                if (this.photos.length && !confirm('開啟後會取代目前的照片與資料，沒存的修改會不見。要繼續嗎？')) return false;

                this.processingCancelled = false;
                const restored = [];
                try {
                    const zip = await JSZip.loadAsync(file);
                    this.throwIfProcessingCancelled();
                    const manifestFile = zip.file('manifest.json');
                    if (!manifestFile) throw new Error('這不是有效的照片清冊專案檔（缺少 manifest.json）。');
                    const manifest = JSON.parse(await manifestFile.async('text'));
                    if (manifest.format !== 'photo-report-project' || !Array.isArray(manifest.photos)) {
                        throw new Error('專案檔格式不支援或內容不完整。');
                    }
                    this.setProgress('正在還原專案照片…', 0, manifest.photos.length);
                    for (let index = 0; index < manifest.photos.length; index++) {
                        this.throwIfProcessingCancelled();
                        const saved = manifest.photos[index];
                        const entry = zip.file(`images/${saved.archiveName}`);
                        if (!entry) throw new Error(`專案檔缺少照片：${saved.name}`);
                        const blob = await entry.async('blob');
                        this.throwIfProcessingCancelled();
                        const imageFile = new File([blob], saved.name, { type: saved.type || blob.type || 'image/jpeg' });
                        const src = URL.createObjectURL(imageFile);
                        let thumbnail;
                        try {
                            thumbnail = await this.createThumbnailUrl(imageFile, src);
                        } catch (thumbErr) {
                            URL.revokeObjectURL(src);
                            throw thumbErr;
                        }
                        restored.push({
                            uid: saved.uid || crypto.randomUUID(),
                            file: imageFile,
                            name: saved.name,
                            src,
                            thumbnailSrc: thumbnail.src,
                            thumbnailAspect: Number(saved.thumbnailAspect) || thumbnail.aspect,
                            visualFingerprint: saved.visualFingerprint || thumbnail.visualFingerprint || '',
                            rotation: Number(saved.rotation) || 0,
                            flipH: Boolean(saved.flipH),
                            parentUid: saved.parentUid || null,
                            cropRect: saved.cropRect && typeof saved.cropRect === 'object' ? {
                                x: Number(saved.cropRect.x) || 0,
                                y: Number(saved.cropRect.y) || 0,
                                w: Number(saved.cropRect.w) || 0.1,
                                h: Number(saved.cropRect.h) || 0.1
                            } : null,
                            sideCrop: this.cloneSideCrop(saved.sideCrop),
                            compositeThumbnailSrc: null,
                            _compositeKey: '',
                            annotations: this.cloneAnnotations(saved.annotations),
                            exifRotation: Number(saved.exifRotation) || 0,
                            seq: String(saved.seq ?? index + 1),
                            date: saved.date || '',
                            time: saved.time || '',
                            timeAutoDefaulted: Boolean(saved.timeAutoDefaulted),
                            location: saved.location || '',
                            desc: saved.desc || '',
                            selected: Boolean(saved.selected),
                            stageX: Number.isFinite(saved.stageX) ? saved.stageX : null,
                            stageY: Number.isFinite(saved.stageY) ? saved.stageY : null
                        });
                        this.setProgress('正在還原專案照片…', index + 1, manifest.photos.length);
                        this.throwIfProcessingCancelled();
                        if ((index + 1) % 5 === 0) await this.yieldToBrowser();
                    }

                    if (this.isLightboxOpen()) {
                        this.lightboxNeedsRender = false;
                        this.closeModal('imageLightboxModal');
                    }
                    this.activeFilter = 'all';
                    this._dupSet = null;
                    this.photos.forEach(photo => this.releasePhotoUrls(photo));
                    this.history = [];
                    this.historyIndex = -1;
                    this.releaseUnusedPhotoUrls();
                    this.photos = restored;
                    this.normalizeLinkedChildrenOrder();
                    restored.forEach(photo => {
                        this.managedUrls.add(photo.src);
                        this.managedUrls.add(photo.thumbnailSrc);
                    });
                    this.applyCaseData(manifest.caseData);
                    this.currentIndex = Math.max(0, Math.min(Number(manifest.currentIndex) || 0, Math.max(0, restored.length - 1)));
                    this.lastSelectedIndex = manifest.lastSelectedIndex ?? this.currentIndex;
                    if (restored.some(photo => !Number.isFinite(photo.stageX) || !Number.isFinite(photo.stageY))) {
                        this.placeUnpositionedPhotos();
                    }
                    this.currentProjectFileHandle = fileHandle || null;
                    this.currentProjectFileName = file.name || '';
                    this.hideAutoDraftBanner();
                    this.recordHistory('開啟專案');
                    const projectTs = this.getProjectTimestamp(manifest, file) || Date.now();
                    this.markProjectSaved(projectTs);
                    this.render();
                    await this.syncAllCompositeThumbnails();
                    alert(`已開啟工作進度，共 ${restored.length} 張照片。`);
                    return true;
                } catch (error) {
                    restored.forEach(item => {
                        if (item.src) URL.revokeObjectURL(item.src);
                        if (item.thumbnailSrc) URL.revokeObjectURL(item.thumbnailSrc);
                    });
                    this.handleProjectError(error, 'open');
                    return false;
                } finally {
                    this.clearProgress();
                }
            }

            async mergeProjects(event) {
                const files = Array.from(event?.target?.files || []);
                if (event?.target) event.target.value = '';
                if (!files.length) return false;
                return this.mergeProjectFiles(files);
            }

            areAnnotationsEqual(annsA, annsB) {
                const listA = Array.isArray(annsA) ? annsA : [];
                const listB = Array.isArray(annsB) ? annsB : [];
                if (listA.length !== listB.length) return false;
                const norm = (a) => ({
                    type: a?.type || '',
                    color: a?.color || 'yellow',
                    x: Math.round((Number(a?.x) || 0) * 1000) / 1000,
                    y: Math.round((Number(a?.y) || 0) * 1000) / 1000,
                    w: Math.round((Number(a?.w) || 0) * 1000) / 1000,
                    h: Math.round((Number(a?.h) || 0) * 1000) / 1000,
                    scale: Math.round((Number(a?.scale) || 1) * 100) / 100,
                    text: String(a?.text || '').trim()
                });
                return JSON.stringify(listA.map(norm)) === JSON.stringify(listB.map(norm));
            }

            areSideCropsEqual(scA, scB) {
                const hasA = Boolean(scA && scA.cropRect);
                const hasB = Boolean(scB && scB.cropRect);
                if (!hasA && !hasB) return true;
                if (!hasA || !hasB) return false;
                const norm = (sc) => ({
                    x: Math.round((Number(sc.cropRect.x) || 0) * 1000) / 1000,
                    y: Math.round((Number(sc.cropRect.y) || 0) * 1000) / 1000,
                    w: Math.round((Number(sc.cropRect.w) || 0) * 1000) / 1000,
                    h: Math.round((Number(sc.cropRect.h) || 0) * 1000) / 1000,
                    widthRatio: Math.round((Number(sc.widthRatio) || 0.65) * 1000) / 1000,
                    offsetY: Math.round((Number.isFinite(Number(sc.offsetY)) ? Number(sc.offsetY) : 0.5) * 1000) / 1000,
                    borderColor: sc.borderColor === 'red' ? 'red' : sc.borderColor === 'none' ? 'none' : 'yellow'
                });
                return JSON.stringify(norm(scA)) === JSON.stringify(norm(scB));
            }

            isCropRectSimilar(rectA, rectB) {
                if (!rectA || !rectB) return false;
                return Math.abs((Number(rectA.x) || 0) - (Number(rectB.x) || 0)) <= 0.025 &&
                    Math.abs((Number(rectA.y) || 0) - (Number(rectB.y) || 0)) <= 0.025 &&
                    Math.abs((Number(rectA.w) || 0) - (Number(rectB.w) || 0)) <= 0.025 &&
                    Math.abs((Number(rectA.h) || 0) - (Number(rectB.h) || 0)) <= 0.025;
            }

            allocateMergedMainSeq(workingPhotos, preferredSeq = '') {
                const existingMain = workingPhotos.filter(p => !this.isLinkedCropChild(p, workingPhotos));
                const usedSeqs = new Set(workingPhotos.map(p => String(p.seq || '').trim()));
                const cleanPreferred = String(preferredSeq || '').trim();
                if (existingMain.length === 0 && cleanPreferred && !usedSeqs.has(cleanPreferred)) {
                    return cleanPreferred;
                }

                let prefix = '';
                let minDigits = 1;
                let maxNum = existingMain.length;
                const lastMain = existingMain[existingMain.length - 1];
                const lastMatch = String(lastMain?.seq || '').match(/^(.*?)(\d+)$/);
                if (lastMatch) {
                    prefix = lastMatch[1];
                    minDigits = lastMatch[2].length;
                }
                existingMain.forEach(p => {
                    const m = String(p.seq || '').match(/^(.*?)(\d+)$/);
                    if (m && m[1] === prefix) {
                        const num = parseInt(m[2], 10);
                        if (Number.isFinite(num) && num > maxNum) maxNum = num;
                    }
                });

                let nextNum = Math.max(1, maxNum + 1);
                let candidate = `${prefix}${String(nextNum).padStart(minDigits, '0')}`;
                while (usedSeqs.has(candidate)) {
                    nextNum++;
                    candidate = `${prefix}${String(nextNum).padStart(minDigits, '0')}`;
                }
                return candidate;
            }

            makeUniqueDetailFilenameInList(parentPhoto, photoList) {
                const origName = String(parentPhoto?.name || 'photo.jpg');
                const dotIdx = origName.lastIndexOf('.');
                const base = dotIdx > 0 ? origName.slice(0, dotIdx) : origName;
                const existingNames = new Set((photoList || []).map(p => p.name));
                let n = 1;
                let candidate = `${base}_特寫${n}.jpg`;
                while (existingNames.has(candidate)) {
                    n++;
                    candidate = `${base}_特寫${n}.jpg`;
                }
                return candidate;
            }

            mergeMatchedPhotoRecord(target, incoming, incomingTs) {
                const targetTs = Number(target._effectiveTs) || 0;
                const preferIncoming = incomingTs > targetTs;
                let complemented = false;
                let overwritten = false;

                ['date', 'location', 'desc'].forEach(field => {
                    const tVal = String(target[field] || '');
                    const iVal = String(incoming[field] || '');
                    const tHas = tVal.trim().length > 0;
                    const iHas = iVal.trim().length > 0;
                    if (!tHas && iHas) {
                        target[field] = iVal;
                        complemented = true;
                    } else if (tHas && iHas && tVal.trim() !== iVal.trim() && preferIncoming) {
                        target[field] = iVal;
                        overwritten = true;
                    }
                });

                const tTime = String(target.time || '').trim();
                const iTime = String(incoming.time || '').trim();
                const tIsRealTime = tTime.length > 0 && !(target.timeAutoDefaulted && tTime === '12:00');
                const iIsRealTime = iTime.length > 0 && !(incoming.timeAutoDefaulted && iTime === '12:00');
                if (!tIsRealTime && iIsRealTime) {
                    target.time = iTime;
                    target.timeAutoDefaulted = false;
                    complemented = true;
                } else if (tIsRealTime && iIsRealTime && tTime !== iTime && preferIncoming) {
                    target.time = iTime;
                    target.timeAutoDefaulted = Boolean(incoming.timeAutoDefaulted);
                    overwritten = true;
                } else if (!tIsRealTime && !iIsRealTime && !tTime && iTime) {
                    target.time = iTime;
                    target.timeAutoDefaulted = Boolean(incoming.timeAutoDefaulted);
                }

                const tRot = (((Number(target.rotation) || 0) % 360) + 360) % 360;
                const iRot = (((Number(incoming.rotation) || 0) % 360) + 360) % 360;
                const tFlip = Boolean(target.flipH);
                const iFlip = Boolean(incoming.flipH);
                const orientSame = (tRot === iRot) && (tFlip === iFlip);
                const tHasOrient = tRot !== 0 || tFlip;
                const iHasOrient = iRot !== 0 || iFlip;

                const tAnns = Array.isArray(target.annotations) ? target.annotations : [];
                const iAnns = Array.isArray(incoming.annotations) ? incoming.annotations : [];
                const tHasAnns = tAnns.length > 0;
                const iHasAnns = iAnns.length > 0;

                const tSide = (target.sideCrop && target.sideCrop.cropRect) ? target.sideCrop : null;
                const iSide = (incoming.sideCrop && incoming.sideCrop.cropRect) ? incoming.sideCrop : null;
                const tHasSide = Boolean(tSide);
                const iHasSide = Boolean(iSide);

                const tHasOverlays = tHasAnns || tHasSide;
                const iHasOverlays = iHasAnns || iHasSide;

                if (orientSame) {
                    if (!tHasAnns && iHasAnns) {
                        target.annotations = this.cloneAnnotations(iAnns);
                        complemented = true;
                    } else if (tHasAnns && iHasAnns && !this.areAnnotationsEqual(tAnns, iAnns) && preferIncoming) {
                        target.annotations = this.cloneAnnotations(iAnns);
                        overwritten = true;
                    }
                    if (!tHasSide && iHasSide) {
                        target.sideCrop = this.cloneSideCrop(iSide);
                        complemented = true;
                    } else if (tHasSide && iHasSide && !this.areSideCropsEqual(tSide, iSide) && preferIncoming) {
                        target.sideCrop = this.cloneSideCrop(iSide);
                        overwritten = true;
                    }
                } else {
                    // 方向不同時，標記與並排綁定在對應方向座標系上，協同切換才不會歪框
                    if (!tHasOverlays && iHasOverlays) {
                        target.rotation = iRot;
                        target.flipH = iFlip;
                        target.annotations = this.cloneAnnotations(iAnns);
                        target.sideCrop = this.cloneSideCrop(iSide);
                        complemented = true;
                    } else if (tHasOverlays && !iHasOverlays) {
                        // 保留目標已有標記或並排的版本與其方向
                    } else if (!tHasOrient && iHasOrient) {
                        target.rotation = iRot;
                        target.flipH = iFlip;
                        complemented = true;
                    } else if (tHasOrient && !iHasOrient) {
                        // 保留目標已調整過的方向
                    } else if (preferIncoming) {
                        target.rotation = iRot;
                        target.flipH = iFlip;
                        if (!this.areAnnotationsEqual(tAnns, iAnns)) {
                            target.annotations = this.cloneAnnotations(iAnns);
                        }
                        if (!this.areSideCropsEqual(tSide, iSide)) {
                            target.sideCrop = this.cloneSideCrop(iSide);
                        }
                        overwritten = true;
                    }
                }

                if (incomingTs > targetTs) {
                    target._effectiveTs = incomingTs;
                }

                return {
                    complemented,
                    overwritten,
                    changed: complemented || overwritten
                };
            }

            async mergeProjectFiles(files) {
                const fileList = Array.from(files || []).filter(Boolean);
                if (!fileList.length) return false;
                if (typeof JSZip === 'undefined') {
                    this.handleProjectError(new Error('JSZip 尚未載入，無法合併進度檔。'), 'merge');
                    return false;
                }

                this.finalizeActiveInputsBeforeSave();
                this.processingCancelled = false;

                const newlyCreatedUrls = new Set();
                try {
                    const parsedSources = [];
                    this.setProgress(`正在讀取 ${fileList.length} 份工作進度檔…`, 0, fileList.length);
                    for (let i = 0; i < fileList.length; i++) {
                        this.throwIfProcessingCancelled();
                        const file = fileList[i];
                        const zip = await JSZip.loadAsync(file);
                        this.throwIfProcessingCancelled();
                        const manifestFile = zip.file('manifest.json');
                        if (!manifestFile) {
                            throw new Error(`「${file.name}」不是有效的照片清冊進度檔（缺少 manifest.json）。`);
                        }
                        const manifest = JSON.parse(await manifestFile.async('text'));
                        if (manifest.format !== 'photo-report-project' || !Array.isArray(manifest.photos)) {
                            throw new Error(`「${file.name}」格式不支援或內容不完整。`);
                        }
                        parsedSources.push({
                            file,
                            zip,
                            manifest,
                            timestamp: this.getProjectTimestamp(manifest, file),
                            originalOrder: i
                        });
                        this.setProgress(`正在讀取 ${fileList.length} 份工作進度檔…`, i + 1, fileList.length);
                    }

                    // 依存檔時間由早到晚排列，確保後面併入較晚存檔的決定性與一致性
                    parsedSources.sort((a, b) => (a.timestamp - b.timestamp) || (a.originalOrder - b.originalOrder));

                    const hadExistingPhotos = this.photos.length > 0;
                    const prevActiveUid = this.photos[this.currentIndex]?.uid || null;
                    const unsavedTimestamp = Math.max(this.lastLocalEditAt || 0, this.currentProjectSavedAt || 0, Date.now());
                    const baseSavedTimestamp = this.currentProjectSavedAt || 0;

                    const workingCaseData = { ...this.getCaseData() };
                    let workingCaseInitialized = hadExistingPhotos;
                    let workingCaseTs = hadExistingPhotos ? this.getCurrentWorkspaceTimestamp() : 0;
                    let caseDataChanged = false;

                    const workingPhotos = [];
                    for (let i = 0; i < this.photos.length; i++) {
                        this.throwIfProcessingCancelled();
                        const photo = this.photos[i];
                        let fp = photo.visualFingerprint || '';
                        let aspect = Number(photo.thumbnailAspect) || 1;
                        if (!fp && photo.file) {
                            const calc = await this.computeVisualFingerprintFromBlob(photo.file);
                            fp = calc.visualFingerprint || '';
                            if (!photo.thumbnailAspect && calc.aspect) aspect = calc.aspect;
                        }
                        const savedSig = this.savedPhotosByUid?.get(photo.uid);
                        const isThisPhotoUnchangedFromSave = Boolean(
                            this.currentProjectSavedAt &&
                            savedSig &&
                            savedSig === this.getPhotoContentSignature(photo)
                        );
                        const effectiveTs = isThisPhotoUnchangedFromSave ? baseSavedTimestamp : unsavedTimestamp;
                        workingPhotos.push({
                            ...photo,
                            flipH: Boolean(photo.flipH),
                            parentUid: photo.parentUid || null,
                            cropRect: photo.cropRect ? { ...photo.cropRect } : null,
                            sideCrop: this.cloneSideCrop(photo.sideCrop),
                            compositeThumbnailSrc: null,
                            _compositeKey: '',
                            annotations: this.cloneAnnotations(photo.annotations),
                            visualFingerprint: fp,
                            thumbnailAspect: aspect,
                            timeAutoDefaulted: Boolean(photo.timeAutoDefaulted),
                            selected: false,
                            _effectiveTs: effectiveTs
                        });
                    }

                    const totalSteps = parsedSources.reduce((sum, src) => sum + src.manifest.photos.length, 0);
                    let completedSteps = 0;
                    let addedCount = 0;
                    const matchedUids = new Set();
                    const complementedUids = new Set();
                    const overwrittenUids = new Set();

                    const createNewWorkingPhoto = async (extSaved, zip, file, extTs, overrideParentUid, isInitialEmptyImport) => {
                        const entry = zip.file(`images/${extSaved.archiveName}`);
                        if (!entry) throw new Error(`進度檔「${file.name}」缺少照片：${extSaved.name}`);
                        const blob = await entry.async('blob');
                        this.throwIfProcessingCancelled();

                        let targetName = extSaved.name || `photo_${workingPhotos.length + 1}.jpg`;
                        if (overrideParentUid && workingPhotos.some(p => p.name === targetName)) {
                            const parentObj = workingPhotos.find(p => p.uid === overrideParentUid);
                            if (parentObj) targetName = this.makeUniqueDetailFilenameInList(parentObj, workingPhotos);
                        }
                        const imageFile = new File([blob], targetName, { type: extSaved.type || blob.type || 'image/jpeg' });
                        const srcUrl = URL.createObjectURL(imageFile);
                        newlyCreatedUrls.add(srcUrl);

                        let thumbnail;
                        try {
                            thumbnail = await this.createThumbnailUrl(imageFile, srcUrl);
                        } catch (err) {
                            URL.revokeObjectURL(srcUrl);
                            newlyCreatedUrls.delete(srcUrl);
                            throw err;
                        }
                        newlyCreatedUrls.add(thumbnail.src);

                        let newUid = extSaved.uid || crypto.randomUUID();
                        if (workingPhotos.some(p => p.uid === newUid)) {
                            newUid = crypto.randomUUID();
                        }

                        let newSeq;
                        if (overrideParentUid) {
                            newSeq = String(extSaved.seq || '');
                        } else if (isInitialEmptyImport && extSaved.seq !== undefined && extSaved.seq !== null && String(extSaved.seq).trim()) {
                            newSeq = String(extSaved.seq);
                        } else {
                            newSeq = this.allocateMergedMainSeq(workingPhotos, extSaved.seq);
                        }

                        return {
                            uid: newUid,
                            file: imageFile,
                            name: targetName,
                            src: srcUrl,
                            thumbnailSrc: thumbnail.src,
                            thumbnailAspect: Number(extSaved.thumbnailAspect) || thumbnail.aspect,
                            visualFingerprint: extSaved.visualFingerprint || thumbnail.visualFingerprint || '',
                            rotation: Number(extSaved.rotation) || 0,
                            flipH: Boolean(extSaved.flipH),
                            parentUid: overrideParentUid || null,
                            cropRect: extSaved.cropRect && typeof extSaved.cropRect === 'object' ? {
                                x: Number(extSaved.cropRect.x) || 0,
                                y: Number(extSaved.cropRect.y) || 0,
                                w: Number(extSaved.cropRect.w) || 0.1,
                                h: Number(extSaved.cropRect.h) || 0.1
                            } : null,
                            sideCrop: this.cloneSideCrop(extSaved.sideCrop),
                            compositeThumbnailSrc: null,
                            _compositeKey: '',
                            annotations: this.cloneAnnotations(extSaved.annotations),
                            exifRotation: Number(extSaved.exifRotation) || 0,
                            seq: newSeq,
                            date: extSaved._materializedDate || '',
                            time: extSaved.time || '',
                            timeAutoDefaulted: Boolean(extSaved.timeAutoDefaulted),
                            location: extSaved._materializedLocation || '',
                            desc: extSaved.desc || '',
                            selected: false,
                            stageX: null,
                            stageY: null,
                            _effectiveTs: extTs
                        };
                    };

                    for (let srcIdx = 0; srcIdx < parsedSources.length; srcIdx++) {
                        this.throwIfProcessingCancelled();
                        const { file, zip, manifest, timestamp: extTs } = parsedSources[srcIdx];
                        const extCase = manifest.caseData || {};
                        const isInitialEmptyImport = !hadExistingPhotos && srcIdx === 0;

                        if (!workingCaseInitialized) {
                            ['caseTitle', 'caseDate', 'defaultLocation', 'deptName', 'officerName', 'layoutSelect'].forEach(k => {
                                if (typeof extCase[k] === 'string') workingCaseData[k] = extCase[k];
                            });
                            workingCaseInitialized = true;
                            workingCaseTs = extTs;
                            caseDataChanged = true;
                        } else {
                            const caseFields = hadExistingPhotos
                                ? ['caseTitle', 'caseDate', 'defaultLocation', 'deptName', 'officerName']
                                : ['caseTitle', 'caseDate', 'defaultLocation', 'deptName', 'officerName', 'layoutSelect'];
                            caseFields.forEach(k => {
                                const curVal = String(workingCaseData[k] || '').trim();
                                const extVal = String(extCase[k] || '').trim();
                                if (!curVal && extVal) {
                                    workingCaseData[k] = extCase[k];
                                    caseDataChanged = true;
                                } else if (!hadExistingPhotos && curVal && extVal && curVal !== extVal && extTs > workingCaseTs) {
                                    // 若較晚檔改寫空白工作台的共用預設日期／地點，暫記既有留空照片的原預設值；未再被較晚檔配對的專屬照片才實體化保留
                                    if (k === 'caseDate') {
                                        workingPhotos.forEach(wp => {
                                            if (!String(wp.date || '').trim() && !wp._fallbackCaseDate) wp._fallbackCaseDate = curVal;
                                        });
                                    } else if (k === 'defaultLocation') {
                                        workingPhotos.forEach(wp => {
                                            if (!String(wp.location || '').trim() && !wp._fallbackDefaultLocation) wp._fallbackDefaultLocation = curVal;
                                        });
                                    }
                                    workingCaseData[k] = extCase[k];
                                    caseDataChanged = true;
                                }
                            });
                            if (extTs > workingCaseTs) workingCaseTs = extTs;
                        }

                        // 更新共用案件資料後再比對預設值，僅對「全新加入」的外部留空照片實體化不同預設值
                        const tgtDefaultDateAfter = String(workingCaseData.caseDate || '').trim();
                        const tgtDefaultLocationAfter = String(workingCaseData.defaultLocation || '').trim();
                        const extDefaultDate = String(extCase.caseDate || '').trim();
                        const extDefaultLocation = String(extCase.defaultLocation || '').trim();

                        const matchedTargetUidsInFile = new Set();
                        const sourceUidToTargetPhoto = new Map();
                        const srcUidSet = new Set(manifest.photos.map(p => p?.uid).filter(Boolean));
                        const extNormalized = manifest.photos.map(saved => {
                            let materializedDate = String(saved.date || '');
                            let materializedLocation = String(saved.location || '');
                            if (!materializedDate.trim() && extDefaultDate && tgtDefaultDateAfter && extDefaultDate !== tgtDefaultDateAfter) {
                                materializedDate = extDefaultDate;
                            }
                            if (!materializedLocation.trim() && extDefaultLocation && tgtDefaultLocationAfter && extDefaultLocation !== tgtDefaultLocationAfter) {
                                materializedLocation = extDefaultLocation;
                            }
                            return {
                                ...saved,
                                _materializedDate: materializedDate,
                                _materializedLocation: materializedLocation
                            };
                        });

                        const isSrcChild = (item) => Boolean(item && item.parentUid && srcUidSet.has(item.parentUid) && item.parentUid !== item.uid);
                        const rootEntries = extNormalized.filter(item => !isSrcChild(item));
                        const childEntries = extNormalized.filter(item => isSrcChild(item));

                        // 1. 先處理主照片（非局部裁切子圖）
                        for (let rIdx = 0; rIdx < rootEntries.length; rIdx++) {
                            this.throwIfProcessingCancelled();
                            const extSaved = rootEntries[rIdx];
                            const entry = zip.file(`images/${extSaved.archiveName}`);
                            if (!entry) throw new Error(`進度檔「${file.name}」缺少照片：${extSaved.name}`);

                            const candidates = workingPhotos.filter(p =>
                                !this.isLinkedCropChild(p, workingPhotos) &&
                                !matchedTargetUidsInFile.has(p.uid)
                            );

                            let matchedTarget = null;
                            if (extSaved.uid) {
                                matchedTarget = candidates.find(p => p.uid === extSaved.uid) || null;
                            }

                            if (!matchedTarget) {
                                const extNormName = this.normalizePhotoFilenameKey(extSaved.name);
                                const sameNameCandidates = candidates.filter(p =>
                                    this.normalizePhotoFilenameKey(p.name) === extNormName
                                );
                                if (sameNameCandidates.length > 0) {
                                    let extFp = extSaved.visualFingerprint || '';
                                    let extAspect = Number(extSaved.thumbnailAspect) || 0;
                                    if (!extFp || !extAspect) {
                                        const blob = await entry.async('blob');
                                        this.throwIfProcessingCancelled();
                                        const calc = await this.computeVisualFingerprintFromBlob(blob);
                                        if (!extFp) extFp = calc.visualFingerprint;
                                        if (!extAspect) extAspect = calc.aspect || 1;
                                    }
                                    let bestScore = Infinity;
                                    sameNameCandidates.forEach(cand => {
                                        if (Math.abs((cand.thumbnailAspect || 1) - (extAspect || 1)) > 0.05) return;
                                        const cmp = this.compareVisualFingerprints(cand.visualFingerprint, extFp);
                                        if (cmp.match && cmp.score < bestScore) {
                                            bestScore = cmp.score;
                                            matchedTarget = cand;
                                        }
                                    });
                                }
                            }

                            if (matchedTarget) {
                                matchedTargetUidsInFile.add(matchedTarget.uid);
                                sourceUidToTargetPhoto.set(extSaved.uid, matchedTarget);
                                matchedUids.add(matchedTarget.uid);
                                const targetTsBefore = Number(matchedTarget._effectiveTs) || 0;
                                // 若較晚的外部檔對此共有照片同樣留空日期／地點，代表繼承新共用預設，解除舊檔暫記預設
                                if (extTs > targetTsBefore) {
                                    if (!String(extSaved.date || '').trim()) delete matchedTarget._fallbackCaseDate;
                                    if (!String(extSaved.location || '').trim()) delete matchedTarget._fallbackDefaultLocation;
                                }
                                // 已配對照片只比對單張明確填寫之日期／地點，不以整份專案預設值強蓋目標單張手寫內容
                                const mergeRes = this.mergeMatchedPhotoRecord(matchedTarget, {
                                    ...extSaved,
                                    date: String(extSaved.date || ''),
                                    location: String(extSaved.location || '')
                                }, extTs);
                                if (mergeRes.complemented) complementedUids.add(matchedTarget.uid);
                                if (mergeRes.overwritten) overwrittenUids.add(matchedTarget.uid);
                            } else {
                                const newPhoto = await createNewWorkingPhoto(extSaved, zip, file, extTs, null, isInitialEmptyImport);
                                workingPhotos.push(newPhoto);
                                matchedTargetUidsInFile.add(newPhoto.uid);
                                if (extSaved.uid) sourceUidToTargetPhoto.set(extSaved.uid, newPhoto);
                                addedCount++;
                            }

                            completedSteps++;
                            this.setProgress(`正在智慧彙整 ${parsedSources.length} 份工作進度…`, completedSteps, Math.max(1, totalSteps));
                            if (completedSteps % 5 === 0) await this.yieldToBrowser();
                        }

                        // 2. 接著處理局部裁切特寫子圖（緊跟對應母圖並支援多層巢狀）
                        let pendingChildren = [...childEntries];
                        let guard = 0;
                        while (pendingChildren.length > 0 && guard <= childEntries.length + 1) {
                            guard++;
                            const nextPending = [];
                            for (let cIdx = 0; cIdx < pendingChildren.length; cIdx++) {
                                this.throwIfProcessingCancelled();
                                const extChild = pendingChildren[cIdx];
                                const targetParent = sourceUidToTargetPhoto.get(extChild.parentUid);
                                if (!targetParent) {
                                    nextPending.push(extChild);
                                    continue;
                                }

                                const entry = zip.file(`images/${extChild.archiveName}`);
                                if (!entry) throw new Error(`進度檔「${file.name}」缺少照片：${extChild.name}`);

                                const childCandidates = workingPhotos.filter(p =>
                                    p.parentUid === targetParent.uid &&
                                    !matchedTargetUidsInFile.has(p.uid)
                                );
                                let matchedChild = extChild.uid
                                    ? childCandidates.find(p => p.uid === extChild.uid) || null
                                    : null;

                                if (!matchedChild && extChild.cropRect) {
                                    matchedChild = childCandidates.find(p => this.isCropRectSimilar(p.cropRect, extChild.cropRect)) || null;
                                }

                                if (matchedChild) {
                                    matchedTargetUidsInFile.add(matchedChild.uid);
                                    sourceUidToTargetPhoto.set(extChild.uid, matchedChild);
                                    matchedUids.add(matchedChild.uid);
                                    const childTsBefore = Number(matchedChild._effectiveTs) || 0;
                                    if (extTs > childTsBefore) {
                                        if (!String(extChild.date || '').trim()) delete matchedChild._fallbackCaseDate;
                                        if (!String(extChild.location || '').trim()) delete matchedChild._fallbackDefaultLocation;
                                    }
                                    const mergeRes = this.mergeMatchedPhotoRecord(matchedChild, {
                                        ...extChild,
                                        date: String(extChild.date || ''),
                                        location: String(extChild.location || '')
                                    }, extTs);
                                    if (mergeRes.complemented) complementedUids.add(matchedChild.uid);
                                    if (mergeRes.overwritten) overwrittenUids.add(matchedChild.uid);
                                } else {
                                    const newChild = await createNewWorkingPhoto(extChild, zip, file, extTs, targetParent.uid, isInitialEmptyImport);
                                    workingPhotos.push(newChild);
                                    matchedTargetUidsInFile.add(newChild.uid);
                                    if (extChild.uid) sourceUidToTargetPhoto.set(extChild.uid, newChild);
                                    addedCount++;
                                }

                                completedSteps++;
                                this.setProgress(`正在智慧彙整 ${parsedSources.length} 份工作進度…`, completedSteps, Math.max(1, totalSteps));
                                if (completedSteps % 5 === 0) await this.yieldToBrowser();
                            }
                            if (nextPending.length === pendingChildren.length) {
                                // 若有異常孤兒或循環相依子圖，降為獨立照片加入
                                for (const orphan of nextPending) {
                                    const orphanPhoto = await createNewWorkingPhoto(orphan, zip, file, extTs, null, isInitialEmptyImport);
                                    workingPhotos.push(orphanPhoto);
                                    if (orphan.uid) sourceUidToTargetPhoto.set(orphan.uid, orphanPhoto);
                                    addedCount++;
                                    completedSteps++;
                                }
                                break;
                            }
                            pendingChildren = nextPending;
                        }
                    }

                    workingPhotos.forEach(p => {
                        if (!String(p.date || '').trim() && p._fallbackCaseDate && p._fallbackCaseDate !== String(workingCaseData.caseDate || '').trim()) {
                            p.date = p._fallbackCaseDate;
                        }
                        if (!String(p.location || '').trim() && p._fallbackDefaultLocation && p._fallbackDefaultLocation !== String(workingCaseData.defaultLocation || '').trim()) {
                            p.location = p._fallbackDefaultLocation;
                        }
                        delete p._fallbackCaseDate;
                        delete p._fallbackDefaultLocation;
                        delete p._effectiveTs;
                    });

                    if (workingPhotos.length === 0 && !hadExistingPhotos) {
                        newlyCreatedUrls.forEach(url => URL.revokeObjectURL(url));
                        alert('所選工作進度檔中沒有任何照片。');
                        return false;
                    }

                    const anyDataChanged = addedCount > 0 || complementedUids.size > 0 || overwrittenUids.size > 0 || caseDataChanged;
                    if (!anyDataChanged && hadExistingPhotos) {
                        newlyCreatedUrls.forEach(url => URL.revokeObjectURL(url));
                        alert(`已完成比對（${parsedSources.length} 份工作進度），目前 ${this.photos.length} 張照片與說明均完全相同，無須重複加入。`);
                        return true;
                    }

                    if (this.isLightboxOpen()) {
                        this.lightboxNeedsRender = false;
                        this.closeModal('imageLightboxModal');
                    }
                    this.activeFilter = 'all';
                    this._dupSet = null;
                    newlyCreatedUrls.forEach(url => this.managedUrls.add(url));
                    this.photos = workingPhotos;
                    this.normalizeLinkedChildrenOrder();
                    this.applyCaseData(workingCaseData);

                    const activeIdx = prevActiveUid ? this.photos.findIndex(p => p.uid === prevActiveUid) : 0;
                    this.currentIndex = activeIdx >= 0 ? activeIdx : 0;
                    if (this.photos[this.currentIndex]) {
                        this.photos.forEach((p, idx) => { p.selected = idx === this.currentIndex; });
                    }
                    this.lastSelectedIndex = this.photos.length ? this.currentIndex : null;

                    // 合併後的整合清冊不再綁定原本單一小檔控制代碼，避免 Ctrl+S 誤覆寫來源檔案
                    this.currentProjectFileHandle = null;
                    this.currentProjectFileName = '';
                    this.currentProjectSavedAt = null;
                    this.lastLocalEditAt = Date.now();
                    this.hideAutoDraftBanner();
                    this.recordHistory('彙整工作進度檔');
                    this.render();
                    await this.syncAllCompositeThumbnails();

                    const totalAfter = this.photos.length;
                    const matchedCount = matchedUids.size;
                    const compCount = complementedUids.size;
                    const overCount = overwrittenUids.size;
                    const details = [];
                    if (hadExistingPhotos) {
                        details.push(`新加入 ${addedCount} 張`);
                    }
                    if (matchedCount > 0) {
                        details.push(`重複照片合併去重 ${matchedCount} 張`);
                    }
                    if (compCount > 0) {
                        details.push(`自動補齊空白欄位 ${compCount} 張`);
                    }
                    if (overCount > 0) {
                        details.push(`較晚版本更新 ${overCount} 張`);
                    }
                    const detailSuffix = details.length ? `（${details.join('、')}）` : '';
                    alert(`已完成彙整 ${parsedSources.length} 份工作進度，目前共 ${totalAfter} 張照片${detailSuffix}。如需覆寫或另存整合檔，請按 Ctrl+S 儲存。`);
                    return true;
                } catch (error) {
                    newlyCreatedUrls.forEach(url => URL.revokeObjectURL(url));
                    this.handleProjectError(error, 'merge');
                    return false;
                } finally {
                    this.clearProgress();
                }
            }

            handleProjectError(error, action = 'open') {
                if (error?.message === '使用者已取消處理。') return;
                console.error(error);
                const prefix = action === 'save'
                    ? '工作進度無法儲存'
                    : action === 'merge'
                        ? '工作進度檔無法合併'
                        : '工作進度檔無法開啟';
                alert(`${prefix}：${error?.message || error}`);
            }

            openAutoDraftDb() {
                return new Promise((resolve, reject) => {
                    if (typeof indexedDB === 'undefined') {
                        return reject(new Error('瀏覽器不支援 IndexedDB'));
                    }
                    const request = indexedDB.open(this.autoDraftDbName, 1);
                    request.onupgradeneeded = () => {
                        const db = request.result;
                        if (!db.objectStoreNames.contains(this.autoDraftStoreName)) {
                            db.createObjectStore(this.autoDraftStoreName);
                        }
                    };
                    request.onsuccess = () => resolve(request.result);
                    request.onerror = () => reject(request.error || new Error('無法開啟自動暫存資料庫'));
                });
            }

            scheduleAutoDraftSave() {
                if (this.isRestoringHistory) return;
                clearTimeout(this.autoDraftTimer);
                if (!this.photos.length) return;
                this.autoDraftTimer = setTimeout(() => {
                    this.autoDraftTimer = null;
                    this.saveAutoDraftNow().catch(() => {});
                }, 1000);
            }

            async saveAutoDraftNow() {
                if (!this.photos.length) return false;
                const totalBytes = this.photos.reduce((sum, p) => sum + (p.file?.size || 0), 0);
                if (this.photos.length > this.maxAutoDraftPhotos || totalBytes > this.maxAutoDraftBytes) {
                    return false;
                }
                const snapshotPhotos = [];
                for (let i = 0; i < this.photos.length; i++) {
                    const photo = this.photos[i];
                    if (!photo || !photo.file) continue;
                    const buffer = await photo.file.arrayBuffer();
                    snapshotPhotos.push({
                        uid: photo.uid,
                        name: photo.name,
                        type: photo.file.type || 'image/jpeg',
                        buffer,
                        rotation: photo.rotation || 0,
                        flipH: Boolean(photo.flipH),
                        parentUid: photo.parentUid || null,
                        cropRect: photo.cropRect ? { ...photo.cropRect } : null,
                        sideCrop: this.cloneSideCrop(photo.sideCrop),
                        annotations: this.cloneAnnotations(photo.annotations),
                        exifRotation: photo.exifRotation || 0,
                        seq: photo.seq,
                        date: photo.date || '',
                        time: photo.time || '',
                        timeAutoDefaulted: Boolean(photo.timeAutoDefaulted),
                        visualFingerprint: photo.visualFingerprint || '',
                        thumbnailAspect: Number(photo.thumbnailAspect) || 1,
                        location: photo.location || '',
                        desc: photo.desc || '',
                        selected: Boolean(photo.selected),
                        stageX: Number.isFinite(photo.stageX) ? photo.stageX : null,
                        stageY: Number.isFinite(photo.stageY) ? photo.stageY : null
                    });
                }
                if (!snapshotPhotos.length) return false;
                const payload = {
                    format: 'photo-report-autodraft',
                    version: 1,
                    savedAt: new Date().toISOString(),
                    caseData: this.getCaseData(),
                    currentIndex: this.currentIndex,
                    lastSelectedIndex: this.lastSelectedIndex,
                    photos: snapshotPhotos
                };
                const db = await this.openAutoDraftDb();
                try {
                    await new Promise((resolve, reject) => {
                        const tx = db.transaction(this.autoDraftStoreName, 'readwrite');
                        tx.objectStore(this.autoDraftStoreName).put(payload, 'latest');
                        tx.oncomplete = () => resolve();
                        tx.onerror = () => reject(tx.error);
                    });
                    this.lastAutoDraftSavedAt = new Date();
                    this.updateProjectSaveStatusUi();
                    return true;
                } finally {
                    db.close();
                }
            }

            async clearAutoDraftDb() {
                clearTimeout(this.autoDraftTimer);
                this.autoDraftTimer = null;
                this.lastAutoDraftSavedAt = null;
                this.pendingAutoDraft = null;
                this.hideAutoDraftBanner();
                try {
                    const db = await this.openAutoDraftDb();
                    try {
                        await new Promise((resolve, reject) => {
                            const tx = db.transaction(this.autoDraftStoreName, 'readwrite');
                            tx.objectStore(this.autoDraftStoreName).delete('latest');
                            tx.oncomplete = () => resolve();
                            tx.onerror = () => reject(tx.error);
                        });
                    } finally {
                        db.close();
                    }
                } catch {
                    // 忽略無權限或不支援 IndexedDB 環境
                }
            }

            hideAutoDraftBanner() {
                this.autoDraftBanner?.classList.add('hidden');
            }

            async checkAndOfferAutoDraft() {
                try {
                    const db = await this.openAutoDraftDb();
                    let draft = null;
                    try {
                        draft = await new Promise((resolve, reject) => {
                            const tx = db.transaction(this.autoDraftStoreName, 'readonly');
                            const req = tx.objectStore(this.autoDraftStoreName).get('latest');
                            req.onsuccess = () => resolve(req.result || null);
                            req.onerror = () => reject(req.error);
                        });
                    } finally {
                        db.close();
                    }
                    if (!draft || draft.format !== 'photo-report-autodraft' || !Array.isArray(draft.photos) || !draft.photos.length) {
                        this.hideAutoDraftBanner();
                        return false;
                    }
                    if (this.photos.length > 0) {
                        this.hideAutoDraftBanner();
                        return false;
                    }
                    this.pendingAutoDraft = draft;
                    if (this.autoDraftBanner && this.autoDraftBannerTitle && this.autoDraftBannerMeta) {
                        const title = draft.caseData?.caseTitle?.trim() || '未命名案件';
                        const savedDate = draft.savedAt ? new Date(draft.savedAt) : null;
                        const timeStr = savedDate && !Number.isNaN(savedDate.getTime())
                            ? savedDate.toLocaleTimeString('zh-TW', { hour: '2-digit', minute: '2-digit' })
                            : '稍早';
                        this.autoDraftBannerTitle.textContent = `發現上次還沒做完的工作進度（案由：${title}・共 ${draft.photos.length} 張照片）`;
                        this.autoDraftBannerMeta.textContent = `自動暫存於 ${timeStr}。點「接續上次進度」可立即還原照片、標記與說明。`;
                        this.autoDraftBanner.classList.remove('hidden');
                    }
                    return true;
                } catch {
                    this.hideAutoDraftBanner();
                    return false;
                }
            }

            async restoreAutoDraft() {
                const draft = this.pendingAutoDraft;
                if (!draft || !Array.isArray(draft.photos) || !draft.photos.length) {
                    this.hideAutoDraftBanner();
                    return false;
                }
                this.processingCancelled = false;
                const restored = [];
                try {
                    this.setProgress('正在還原自動暫存進度…', 0, draft.photos.length);
                    for (let index = 0; index < draft.photos.length; index++) {
                        this.throwIfProcessingCancelled();
                        const saved = draft.photos[index];
                        const blobSource = saved.buffer || saved.blob;
                        if (!blobSource) throw new Error(`草稿缺少照片資料：${saved.name}`);
                        const imageFile = new File([blobSource], saved.name || `photo_${index + 1}.jpg`, { type: saved.type || 'image/jpeg' });
                        const src = URL.createObjectURL(imageFile);
                        let thumbnail;
                        try {
                            thumbnail = await this.createThumbnailUrl(imageFile, src);
                        } catch (thumbErr) {
                            URL.revokeObjectURL(src);
                            throw thumbErr;
                        }
                        restored.push({
                            uid: saved.uid || crypto.randomUUID(),
                            file: imageFile,
                            name: saved.name || `photo_${index + 1}.jpg`,
                            src,
                            thumbnailSrc: thumbnail.src,
                            thumbnailAspect: Number(saved.thumbnailAspect) || thumbnail.aspect,
                            visualFingerprint: saved.visualFingerprint || thumbnail.visualFingerprint || '',
                            rotation: Number(saved.rotation) || 0,
                            flipH: Boolean(saved.flipH),
                            parentUid: saved.parentUid || null,
                            cropRect: saved.cropRect && typeof saved.cropRect === 'object' ? {
                                x: Number(saved.cropRect.x) || 0,
                                y: Number(saved.cropRect.y) || 0,
                                w: Number(saved.cropRect.w) || 0.1,
                                h: Number(saved.cropRect.h) || 0.1
                            } : null,
                            sideCrop: this.cloneSideCrop(saved.sideCrop),
                            compositeThumbnailSrc: null,
                            _compositeKey: '',
                            annotations: this.cloneAnnotations(saved.annotations),
                            exifRotation: Number(saved.exifRotation) || 0,
                            seq: String(saved.seq ?? index + 1),
                            date: saved.date || '',
                            time: saved.time || '',
                            timeAutoDefaulted: Boolean(saved.timeAutoDefaulted),
                            location: saved.location || '',
                            desc: saved.desc || '',
                            selected: Boolean(saved.selected),
                            stageX: Number.isFinite(saved.stageX) ? saved.stageX : null,
                            stageY: Number.isFinite(saved.stageY) ? saved.stageY : null
                        });
                        this.setProgress('正在還原自動暫存進度…', index + 1, draft.photos.length);
                        this.throwIfProcessingCancelled();
                        if ((index + 1) % 5 === 0) await this.yieldToBrowser();
                    }

                    this.activeFilter = 'all';
                    this._dupSet = null;
                    this.photos.forEach(photo => this.releasePhotoUrls(photo));
                    this.history = [];
                    this.historyIndex = -1;
                    this.releaseUnusedPhotoUrls();
                    this.photos = restored;
                    this.normalizeLinkedChildrenOrder();
                    restored.forEach(photo => {
                        this.managedUrls.add(photo.src);
                        this.managedUrls.add(photo.thumbnailSrc);
                    });
                    this.applyCaseData(draft.caseData);
                    this.currentIndex = Math.max(0, Math.min(Number(draft.currentIndex) || 0, Math.max(0, restored.length - 1)));
                    this.lastSelectedIndex = draft.lastSelectedIndex ?? this.currentIndex;
                    if (restored.some(photo => !Number.isFinite(photo.stageX) || !Number.isFinite(photo.stageY))) {
                        this.placeUnpositionedPhotos();
                    }
                    this.currentProjectFileHandle = null;
                    this.currentProjectFileName = '';
                    this.currentProjectSavedAt = null;
                    this.lastLocalEditAt = Date.now();
                    this.savedPhotosByUid = new Map();
                    this.pendingAutoDraft = null;
                    this.lastAutoDraftSavedAt = draft.savedAt ? new Date(draft.savedAt) : new Date();
                    this.hideAutoDraftBanner();
                    this.recordHistory('恢復自動暫存進度');
                    this.render();
                    await this.syncAllCompositeThumbnails();
                    return true;
                } catch (error) {
                    restored.forEach(item => {
                        if (item.src) URL.revokeObjectURL(item.src);
                        if (item.thumbnailSrc) URL.revokeObjectURL(item.thumbnailSrc);
                    });
                    this.handleProjectError(error, 'open');
                    return false;
                } finally {
                    this.clearProgress();
                }
            }

            async discardAutoDraft() {
                await this.clearAutoDraftDb();
                this.updateProjectSaveStatusUi();
            }

            setThumbnailSize(size) {
                this.thumbnailSize = Math.max(90, Math.min(480, Math.round(size)));
                this.saveWorkspacePreference('thumbnailSize', this.thumbnailSize);
                this.applyWorkspaceLayout();
                // 卡片寬度與換列完全由目前尺寸重算，滑桿拖動不必再點擊畫面。
                this.render();
            }

            handleThumbnailWheel(event) {
                if (!event.ctrlKey) return;
                event.preventDefault();
                const direction = event.deltaY < 0 ? 1 : -1;
                this.setThumbnailSize(this.thumbnailSize + direction * 15);
            }

            toggleEditor() {
                this.editorCollapsed = !this.editorCollapsed;
                this.saveWorkspacePreference('editorCollapsed', this.editorCollapsed);
                this.applyWorkspaceLayout();
            }

            applyWorkspaceLayout() {
                this.applyMainPanelWidth();
                this.thumbGrid.style.gridTemplateColumns = `repeat(auto-fill, minmax(min(100%, ${this.thumbnailSize}px), ${this.thumbnailSize}px))`;
                // 清冊編輯以序號閱讀為主：使用標準網格，讓每列由最高卡片決定高度。
                // 不採瀑布流，避免下一張照片跳到前一列的空白處而打亂閱讀順序。
                this.thumbGrid.style.gridAutoRows = 'auto';
                this.thumbGrid.style.gridAutoFlow = 'row';
                this.thumbGrid.style.justifyContent = 'start';
                this.thumbGrid.classList.toggle('thumbnail-density-compact', this.thumbnailSize <= 110);
                this.thumbnailSizeRange.value = String(this.thumbnailSize);
                this.thumbnailSizeLabel.textContent = `${this.thumbnailSize}px`;

                this.workspaceContainer.classList.toggle('lg:grid-cols-12', !this.editorCollapsed);
                this.workspaceContainer.classList.toggle('lg:grid-cols-1', this.editorCollapsed);
                this.photoPanel.classList.toggle('lg:col-span-8', !this.editorCollapsed);
                this.photoPanel.classList.toggle('lg:col-span-1', this.editorCollapsed);
                this.editorPanel.classList.toggle('hidden', this.editorCollapsed);

                const label = this.btnToggleEditor.querySelector('span');
                const icon = this.btnToggleEditor.querySelector('i');
                label.textContent = this.editorCollapsed ? '顯示編輯欄' : '隱藏編輯欄';
                icon.className = this.editorCollapsed ? 'fa-solid fa-pen-to-square' : 'fa-solid fa-table-columns';
                this.btnToggleEditor.setAttribute('aria-pressed', String(this.editorCollapsed));
            }

            applyMainPanelWidth() {
                this.appMainLayout?.style.setProperty('--side-panel-width', `${this.sidePanelWidth}px`);
            }

            beginMainPanelResize(event) {
                if (event.button !== 0 || window.innerWidth < 1024) return;
                this.panelResizeState = { pointerId: event.pointerId };
                this.mainPanelSplitter.setPointerCapture(event.pointerId);
                document.body.classList.add('is-resizing-panels');
                event.preventDefault();
            }

            updateMainPanelResize(event) {
                if (!this.panelResizeState || this.panelResizeState.pointerId !== event.pointerId) return;
                const rect = this.appMainLayout.getBoundingClientRect();
                const maxWidth = Math.min(560, Math.max(320, rect.width - 420));
                this.sidePanelWidth = Math.max(280, Math.min(maxWidth, Math.round(event.clientX - rect.left)));
                this.applyMainPanelWidth();
                event.preventDefault();
            }

            endMainPanelResize(event, cancelled = false) {
                if (!this.panelResizeState || this.panelResizeState.pointerId !== event.pointerId) return;
                if (this.mainPanelSplitter.hasPointerCapture(event.pointerId)) this.mainPanelSplitter.releasePointerCapture(event.pointerId);
                this.panelResizeState = null;
                document.body.classList.remove('is-resizing-panels');
                if (!cancelled) this.saveWorkspacePreference('sidePanelWidth', this.sidePanelWidth);
            }

            getPhotoDisplayLayout(photo) {
                if (photo && photo.sideCrop && photo.compositeThumbnailSrc) {
                    return {
                        safeAspect: 1,
                        imageStyle: 'width:100%;height:100%;transform:translate(-50%, -50%);',
                        src: photo.compositeThumbnailSrc,
                        isComposite: true
                    };
                }
                // 縮圖框固定為正方形：所有卡片等高、每一排對齊；照片以 contain 完整顯示（不裁切）。
                // 正方形旋轉 90° 後仍是同一個方框，因此只需旋轉與水平鏡像，不必依照片比例換算寬高。
                const flipTransform = photo.flipH ? ' scaleX(-1)' : '';
                const imageStyle = `width:100%;height:100%;transform:translate(-50%, -50%)${flipTransform} rotate(${photo.rotation || 0}deg);`;
                return { safeAspect: 1, imageStyle, src: photo.thumbnailSrc, isComposite: false };
            }

            getValidParentPhoto(photo, photos = this.photos) {
                if (!photo || !photo.parentUid || !Array.isArray(photos)) return null;
                const visited = new Set([photo.uid]);
                const parent = photos.find(p => p && p.uid === photo.parentUid) || null;
                if (!parent || visited.has(parent.uid)) return null;
                let walker = parent;
                while (walker && walker.parentUid) {
                    if (visited.has(walker.uid)) return null;
                    visited.add(walker.uid);
                    walker = photos.find(p => p && p.uid === walker.parentUid) || null;
                }
                return parent;
            }

            isLinkedCropChild(photo, photos = this.photos) {
                return Boolean(this.getValidParentPhoto(photo, photos));
            }

            isDescendantOf(candidate, ancestorUid, photos = this.photos) {
                if (!candidate || !ancestorUid || candidate.uid === ancestorUid || !Array.isArray(photos)) return false;
                const visited = new Set([candidate.uid]);
                let parentUid = candidate.parentUid;
                while (parentUid) {
                    if (parentUid === ancestorUid) return true;
                    if (visited.has(parentUid)) return false;
                    visited.add(parentUid);
                    const parent = photos.find(p => p && p.uid === parentUid);
                    parentUid = parent ? parent.parentUid : null;
                }
                return false;
            }

            getNextMainSeqNumber() {
                return this.photos.filter(p => !this.isLinkedCropChild(p)).length + 1;
            }

            syncChildCropSequences() {
                if (!Array.isArray(this.photos) || !this.photos.length) return;
                const uidMap = new Map(this.photos.map(p => [p.uid, p]));
                const childCountByParent = new Map();
                this.photos.forEach(p => {
                    if (!p || !p.parentUid) return;
                    const parent = uidMap.get(p.parentUid);
                    if (!parent || !this.getValidParentPhoto(p, this.photos)) return;
                    const nextSub = (childCountByParent.get(parent.uid) || 0) + 1;
                    childCountByParent.set(parent.uid, nextSub);
                    p.seq = `${parent.seq}-${nextSub}`;
                });
            }

            normalizeLinkedChildrenOrder() {
                if (!Array.isArray(this.photos) || !this.photos.length) return;
                const uidMap = new Map(this.photos.map(p => [p.uid, p]));
                this.photos.forEach(p => {
                    if (!p) return;
                    if (p.parentUid && (!uidMap.has(p.parentUid) || !this.getValidParentPhoto(p, this.photos))) {
                        p.parentUid = null;
                        p.cropRect = null;
                    }
                });

                const activePhoto = this.photos[this.currentIndex] || null;
                const lastSelectedPhoto = this.lastSelectedIndex !== null ? this.photos[this.lastSelectedIndex] || null : null;

                const childrenByParent = new Map();
                const roots = [];
                this.photos.forEach(p => {
                    if (this.isLinkedCropChild(p, this.photos)) {
                        const arr = childrenByParent.get(p.parentUid) || [];
                        arr.push(p);
                        childrenByParent.set(p.parentUid, arr);
                    } else {
                        roots.push(p);
                    }
                });

                const ordered = [];
                const seen = new Set();
                const appendWithDescendants = (node) => {
                    if (!node || seen.has(node.uid)) return;
                    seen.add(node.uid);
                    ordered.push(node);
                    const kids = childrenByParent.get(node.uid);
                    if (kids) {
                        kids.forEach(kid => appendWithDescendants(kid));
                    }
                };
                roots.forEach(root => appendWithDescendants(root));

                this.photos = ordered;
                this.syncChildCropSequences();

                if (activePhoto) {
                    const nextCur = this.photos.indexOf(activePhoto);
                    if (nextCur !== -1) this.currentIndex = nextCur;
                }
                if (lastSelectedPhoto) {
                    const nextLast = this.photos.indexOf(lastSelectedPhoto);
                    this.lastSelectedIndex = nextLast !== -1 ? nextLast : this.currentIndex;
                }
            }

            refreshSequenceBadgesAndGridInputs() {
                this.photos.forEach((p, idx) => {
                    const badge = document.querySelector(`#thumb_${p.uid} .game-seq-badge`);
                    if (badge) badge.textContent = `#${p.seq}`;
                    const row = this.gridTbody?.children[idx];
                    if (row) {
                        const seqInput = row.querySelector('td:nth-child(3) input');
                        if (seqInput && document.activeElement !== seqInput) {
                            seqInput.value = p.seq || '';
                        }
                        const childHint = row.querySelector('.grid-crop-parent-hint');
                        const parentPhoto = this.getValidParentPhoto(p);
                        if (childHint && parentPhoto) {
                            childHint.textContent = `↳ #${parentPhoto.seq} 特寫`;
                        }
                    }
                });
            }

            makeUniqueDetailFilename(parentPhoto) {
                const origName = String(parentPhoto?.name || 'photo.jpg');
                const dotIdx = origName.lastIndexOf('.');
                const base = dotIdx > 0 ? origName.slice(0, dotIdx) : origName;
                const existingNames = new Set(this.photos.map(p => p.name));
                let n = 1;
                let candidate = `${base}_特寫${n}.jpg`;
                while (existingNames.has(candidate)) {
                    n++;
                    candidate = `${base}_特寫${n}.jpg`;
                }
                return candidate;
            }

            jumpToParentPhoto() {
                const cur = this.photos[this.currentIndex];
                const parent = this.getValidParentPhoto(cur);
                if (!parent) return;
                const parentIdx = this.photos.indexOf(parent);
                if (parentIdx < 0) return;
                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                if (this.isLightboxOpen()) {
                    this.lightboxIndex = parentIdx;
                    this.syncLightboxSelectionAndEditor(parentIdx);
                    this.updateLightbox();
                } else {
                    this.photos.forEach((p, i) => { p.selected = i === parentIdx; });
                    this.currentIndex = parentIdx;
                    this.lastSelectedIndex = parentIdx;
                    this.render();
                }
            }

            unlinkCurrentChildPhoto() {
                const cur = this.photos[this.currentIndex];
                if (!cur || !this.isLinkedCropChild(cur)) return;
                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                cur.parentUid = null;
                cur.cropRect = null;
                this.syncChildCropSequences();
                this.render();
                this.recordHistory('解除特寫編號鎖定');
            }

            getPhotoAnnotationBadgeHtml(photo) {
                const annCount = Array.isArray(photo?.annotations) ? photo.annotations.length : 0;
                const flipped = Boolean(photo?.flipH);
                const parentPhoto = this.getValidParentPhoto(photo);
                const hasSideCrop = Boolean(photo?.sideCrop);
                if (!parentPhoto && !hasSideCrop && !flipped && annCount === 0) return '';

                const fullLabels = [];
                const compactTokens = [];
                const tooltipParts = [];
                const badgeFlags = [];

                if (parentPhoto) {
                    fullLabels.push('特寫');
                    compactTokens.push('特');
                    tooltipParts.push(`隨 #${parentPhoto.seq} 連動之局部裁切特寫`);
                    badgeFlags.push('crop');
                }
                if (hasSideCrop) {
                    fullLabels.push('◫並排');
                    compactTokens.push('◫');
                    tooltipParts.push('含右側同格並排局部放大圖（模式 B）');
                    badgeFlags.push('sidecrop');
                }
                if (flipped) {
                    fullLabels.push('⇄翻轉');
                    compactTokens.push('⇄');
                    tooltipParts.push('水平鏡像翻轉');
                    badgeFlags.push('flip');
                }
                if (annCount > 0) {
                    fullLabels.push(`標記${annCount}`);
                    compactTokens.push(`標${annCount}`);
                    tooltipParts.push(`標記 ${annCount} 項`);
                    badgeFlags.push('annotation');
                }

                // 多種標記共存時合併成單一右下角膠囊，避免小縮圖左右互疊並減少遮擋
                const useCompactTokens = fullLabels.length > 1 && (this.thumbnailSize <= 125 || fullLabels.length >= 3);
                const displayText = useCompactTokens ? compactTokens.join('・') : fullLabels.join('・');

                let variantClass = '';
                if (!annCount && !flipped) {
                    if (hasSideCrop && !parentPhoto) variantClass = ' thumb-annotation-badge--sidecrop';
                    else if (parentPhoto && !hasSideCrop) variantClass = ' thumb-annotation-badge--crop';
                }

                return `<span class="thumb-annotation-badge${variantClass}" data-badges="${badgeFlags.join(',')}" title="${escapeHtml(tooltipParts.join('・'))}">${escapeHtml(displayText)}</span>`;
            }

            getStageStep() {
                return this.thumbnailSize + 12;
            }

            estimateStageCardHeight(photo) {
                const { safeAspect } = this.getPhotoDisplayLayout(photo);
                const metaHeight = this.thumbnailSize <= 110 ? 25 : 34;
                return Math.ceil((this.thumbnailSize / safeAspect) + metaHeight);
            }

            getStagePosition(photo, index) {
                const step = this.getStageStep();
                if (Number.isFinite(photo.stageX) && Number.isFinite(photo.stageY)) {
                    return { x: Math.max(0, photo.stageX * step), y: Math.max(0, photo.stageY * step) };
                }
                const availableWidth = Math.max(step, this.thumbGrid.clientWidth - 16);
                const columns = Math.max(1, Math.floor(availableWidth / step));
                return { x: (index % columns) * step, y: Math.floor(index / columns) * step };
            }

            getStageCanvasHeight() {
                const bottom = this.photos.reduce((max, photo, index) => {
                    const position = this.getStagePosition(photo, index);
                    return Math.max(max, position.y + this.estimateStageCardHeight(photo));
                }, 0);
                return Math.max(280, Math.ceil(bottom + 16));
            }

            autoArrangeStage({ record = true, render = true } = {}) {
                if (!this.photos.length) return;
                const step = this.getStageStep();
                const gap = 12;
                const availableWidth = Math.max(this.thumbnailSize, this.thumbGrid.clientWidth - 16, 800);
                let x = 0;
                let y = 0;
                let rowHeight = 0;

                this.photos.forEach(photo => {
                    if (x > 0 && x + this.thumbnailSize > availableWidth) {
                        x = 0;
                        y += rowHeight + gap;
                        rowHeight = 0;
                    }
                    photo.stageX = Math.round((x / step) * 2) / 2;
                    photo.stageY = Math.round((y / step) * 2) / 2;
                    x += this.thumbnailSize + gap;
                    rowHeight = Math.max(rowHeight, this.estimateStageCardHeight(photo));
                });
                this.sortPhotosByStageReadingOrder();
                if (render) this.render();
                if (record) this.recordHistory('自動排列照片舞台');
            }

            sortPhotosByNaturalFilename({ confirmBeforeSort = false } = {}) {
                if (this.photos.length < 2) return;
                if (confirmBeforeSort && !confirm('要依檔名重新排列全部照片嗎？手動排好的順序會被取代。')) return;
                const currentPhoto = this.photos[this.currentIndex];
                this.photos.sort((a, b) => this.comparePhotoFileOrder(a, b));
                this.currentIndex = Math.max(0, this.photos.indexOf(currentPhoto));
                this.lastSelectedIndex = this.currentIndex;
                this.normalizeLinkedChildrenOrder();
                this.render();
                this.recordHistory('依檔名整齊排序');
            }

            getPhotoCaptureSortKeys(photo, defaultDate = '') {
                const rawDate = String(photo?.date || '').trim() || String(defaultDate || '').trim();
                const dateDigits = rawDate.replace(/[^0-9]/g, '');
                let dateKey = '99999999';
                if (dateDigits.length === 7) {
                    dateKey = `0${dateDigits}`;
                } else if (dateDigits.length === 6) {
                    dateKey = `00${dateDigits}`;
                } else if (dateDigits.length === 8) {
                    const y = parseInt(dateDigits.slice(0, 4), 10);
                    if (y >= 1912) {
                        dateKey = `0${String(y - 1911).padStart(3, '0')}${dateDigits.slice(4, 8)}`;
                    } else {
                        dateKey = dateDigits;
                    }
                } else if (dateDigits.length > 0) {
                    dateKey = dateDigits.padStart(8, '0').slice(0, 8);
                }

                const rawTime = String(photo?.time || '').trim();
                const timeDigits = rawTime.replace(/[^0-9]/g, '');
                let timeKey = '999999';
                if (timeDigits.length === 4) {
                    timeKey = `${timeDigits}00`;
                } else if (timeDigits.length >= 6) {
                    timeKey = timeDigits.slice(0, 6);
                } else if (timeDigits.length > 0) {
                    timeKey = timeDigits.padEnd(6, '0').slice(0, 6);
                }

                return { dateKey, timeKey };
            }

            sortPhotosByCaptureTime({ confirmBeforeSort = false } = {}) {
                if (this.photos.length < 2) return;
                this.finalizeActiveInputsBeforeSave();
                if (confirmBeforeSort && !confirm('要依採證日期與時間（由早到晚）重新排列全部照片嗎？手動排好的順序會被取代。')) return;
                const currentPhoto = this.photos[this.currentIndex];
                const defaultDate = document.getElementById('caseDate')?.value?.trim() || '';
                const rootCache = new Map();
                const resolveRoot = (photo) => {
                    if (!photo) return null;
                    if (rootCache.has(photo.uid)) return rootCache.get(photo.uid);
                    let cur = photo;
                    const visited = new Set([cur.uid]);
                    while (this.isLinkedCropChild(cur)) {
                        const parent = this.getValidParentPhoto(cur);
                        if (!parent || visited.has(parent.uid)) break;
                        visited.add(parent.uid);
                        cur = parent;
                    }
                    rootCache.set(photo.uid, cur);
                    return cur;
                };

                this.photos.sort((a, b) => {
                    const rootA = resolveRoot(a) || a;
                    const rootB = resolveRoot(b) || b;
                    if (rootA !== rootB) {
                        const keyA = this.getPhotoCaptureSortKeys(rootA, defaultDate);
                        const keyB = this.getPhotoCaptureSortKeys(rootB, defaultDate);
                        const dateCmp = keyA.dateKey.localeCompare(keyB.dateKey);
                        if (dateCmp !== 0) return dateCmp;
                        const timeCmp = keyA.timeKey.localeCompare(keyB.timeKey);
                        if (timeCmp !== 0) return timeCmp;
                        return this.comparePhotoFileOrder(rootA, rootB);
                    }
                    if (a === rootA && b !== rootB) return -1;
                    if (b === rootB && a !== rootA) return 1;
                    const keyA = this.getPhotoCaptureSortKeys(a, defaultDate);
                    const keyB = this.getPhotoCaptureSortKeys(b, defaultDate);
                    const dateCmp = keyA.dateKey.localeCompare(keyB.dateKey);
                    if (dateCmp !== 0) return dateCmp;
                    const timeCmp = keyA.timeKey.localeCompare(keyB.timeKey);
                    if (timeCmp !== 0) return timeCmp;
                    return this.comparePhotoFileOrder(a, b);
                });

                this.normalizeLinkedChildrenOrder();
                this.currentIndex = Math.max(0, this.photos.indexOf(currentPhoto));
                this.lastSelectedIndex = this.currentIndex;
                this.render();
                this.recordHistory('依採證時間排序');
            }

            placeUnpositionedPhotos() {
                const pending = this.photos.filter(photo => !Number.isFinite(photo.stageX) || !Number.isFinite(photo.stageY));
                if (!pending.length) return;
                if (pending.length === this.photos.length) {
                    this.autoArrangeStage({ record: false, render: false });
                    return;
                }
                const step = this.getStageStep();
                const gap = 12;
                const availableWidth = Math.max(this.thumbnailSize, this.thumbGrid.clientWidth - 16, 800);
                const existingBottom = this.photos.reduce((max, photo, index) => {
                    if (!Number.isFinite(photo.stageX) || !Number.isFinite(photo.stageY)) return max;
                    const position = this.getStagePosition(photo, index);
                    return Math.max(max, position.y + this.estimateStageCardHeight(photo));
                }, 0);
                let x = 0;
                let y = existingBottom ? existingBottom + gap : 0;
                let rowHeight = 0;
                pending.forEach(photo => {
                    if (x > 0 && x + this.thumbnailSize > availableWidth) {
                        x = 0;
                        y += rowHeight + gap;
                        rowHeight = 0;
                    }
                    photo.stageX = Math.round((x / step) * 2) / 2;
                    photo.stageY = Math.round((y / step) * 2) / 2;
                    x += this.thumbnailSize + gap;
                    rowHeight = Math.max(rowHeight, this.estimateStageCardHeight(photo));
                });
                this.sortPhotosByStageReadingOrder();
            }

            sortPhotosByStageReadingOrder() {
                const currentPhoto = this.photos[this.currentIndex];
                this.photos.sort((a, b) => (Number(a.stageY) - Number(b.stageY)) || (Number(a.stageX) - Number(b.stageX)));
                this.currentIndex = Math.max(0, this.photos.indexOf(currentPhoto));
                this.lastSelectedIndex = this.currentIndex;
                this.normalizeLinkedChildrenOrder();
            }

            getStageSlotsInReadingOrder() {
                return [...this.photos]
                    .sort((a, b) => (Number(a.stageY) - Number(b.stageY)) || (Number(a.stageX) - Number(b.stageX)))
                    .map(photo => ({ stageX: photo.stageX, stageY: photo.stageY }));
            }

            applyStageSlotsByPhotoOrder(slots) {
                this.photos.forEach((photo, index) => {
                    const slot = slots[index];
                    if (!slot) return;
                    photo.stageX = slot.stageX;
                    photo.stageY = slot.stageY;
                });
            }

            beginStageDrag(index, event) {
                if (event.button !== 0 || event.target.closest('input, button, .thumbnail-stage')) return;
                const photo = this.photos[index];
                if (!photo) return;
                if (!photo.selected) {
                    this.photos.forEach(item => item.selected = false);
                    photo.selected = true;
                    this.currentIndex = index;
                    this.lastSelectedIndex = index;
                }
                const selectedIndices = this.getSelectedIndicesOrCurrent();
                const initialPositions = new Map(selectedIndices.map(selectedIndex => {
                    const item = this.photos[selectedIndex];
                    return [item.uid, this.getStagePosition(item, selectedIndex)];
                }));
                this.stageDrag = {
                    pointerId: event.pointerId,
                    startX: event.clientX,
                    startY: event.clientY,
                    initialPositions,
                    indices: selectedIndices,
                    moved: false,
                    card: event.currentTarget
                };
                event.currentTarget.setPointerCapture(event.pointerId);
                event.preventDefault();
                event.stopPropagation();
            }

            updateStageDrag(event) {
                const state = this.stageDrag;
                if (!state || event.pointerId !== state.pointerId) return;
                const deltaX = event.clientX - state.startX;
                const deltaY = event.clientY - state.startY;
                state.moved = state.moved || Math.abs(deltaX) > 3 || Math.abs(deltaY) > 3;
                state.indices.forEach(index => {
                    const card = document.getElementById(`thumb_${this.photos[index]?.uid}`);
                    if (card) card.style.transform = `translate(${deltaX}px, ${deltaY}px)`;
                });
                event.preventDefault();
            }

            endStageDrag(event, cancelled = false) {
                const state = this.stageDrag;
                if (!state || event.pointerId !== state.pointerId) return;
                if (state.card?.hasPointerCapture(event.pointerId)) state.card.releasePointerCapture(event.pointerId);
                this.stageDrag = null;
                if (!state.moved || cancelled) {
                    state.indices.forEach(index => document.getElementById(`thumb_${this.photos[index]?.uid}`)?.style.removeProperty('transform'));
                    return;
                }
                const step = this.getStageStep();
                const snap = step / 2;
                const deltaX = event.clientX - state.startX;
                const deltaY = event.clientY - state.startY;
                state.indices.forEach(index => {
                    const photo = this.photos[index];
                    const initial = state.initialPositions.get(photo.uid);
                    photo.stageX = Math.max(0, Math.round((initial.x + deltaX) / snap) * snap / step);
                    photo.stageY = Math.max(0, Math.round((initial.y + deltaY) / snap) * snap / step);
                });
                this.sortPhotosByStageReadingOrder();
                this.render();
                this.recordHistory('自由拖曳照片舞台');
            }

            hasBlockingModalOpen() {
                return ['shortcutsModal', 'sequenceModal', 'updateModal', 'upToDateModal', 'updateErrorModal', 'exportAuditModal']
                    .some(id => {
                        const el = document.getElementById(id);
                        return Boolean(el && !el.classList.contains('hidden'));
                    });
            }

            handleDeleteShortcut({ fromLightbox = false } = {}) {
                if (!this.photos.length || this.hasBlockingModalOpen()) return false;
                if (fromLightbox || this.isLightboxOpen()) {
                    const targetIndex = (this.lightboxIndex !== null && this.photos[this.lightboxIndex])
                        ? this.lightboxIndex
                        : this.currentIndex;
                    if (!this.photos[targetIndex]) return false;
                    this.cancelLightboxInteractions();
                    this.lightboxCropDraft = null;
                    this.currentIndex = targetIndex;
                    return this.removePhoto(targetIndex, { historyLabel: '從清冊移除照片' });
                }
                const indices = this.getSelectedIndicesOrCurrent();
                if (!indices.length) return false;
                if (this.currentView === 'workspace' && this.activeFilter !== 'all') {
                    const visibleSet = new Set(this.getVisiblePhotoIndices().map(v => v.index));
                    if (visibleSet.size === 0 || indices.every(i => !visibleSet.has(i))) return false;
                }
                if (indices.length === 1) {
                    return this.removePhoto(indices[0], { historyLabel: '從清冊移除照片' });
                }
                return this.removeSelectedPhotos();
            }

            removePhoto(index, { historyLabel = '刪除照片', preserveSelectionIfAny = true } = {}) {
                if (!this.photos[index]) return false;
                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                const wasSelected = Boolean(this.photos[index].selected);
                this.photos.splice(index, 1);
                if (index < this.currentIndex) {
                    this.currentIndex = Math.max(0, this.currentIndex - 1);
                } else if (this.currentIndex >= this.photos.length) {
                    this.currentIndex = Math.max(0, this.photos.length - 1);
                }
                this.normalizeLinkedChildrenOrder();
                if (this.photos.length > 0 && !this.isLightboxOpen() && this.activeFilter !== 'all') {
                    const visible = this.getVisiblePhotoIndices();
                    if (visible.length > 0 && !visible.some(v => v.index === this.currentIndex)) {
                        const nextVisible = visible.find(v => v.index >= this.currentIndex) || visible[visible.length - 1];
                        this.currentIndex = nextVisible.index;
                    }
                }
                const anyStillSelected = preserveSelectionIfAny && this.photos.some(p => p.selected);
                if (this.photos.length > 0) {
                    if (wasSelected && !anyStillSelected && this.photos[this.currentIndex]) {
                        this.photos[this.currentIndex].selected = true;
                    }
                    this.lastSelectedIndex = this.currentIndex;
                } else {
                    this.lastSelectedIndex = null;
                }
                this.render();
                this.recordHistory(historyLabel);
                return true;
            }

            removeSelectedPhotos() {
                const selectedIndices = this.getSelectedIndicesOrCurrent();
                if (!selectedIndices.length) return false;
                const noun = selectedIndices.length === 1 ? '這張照片' : `已選取的 ${selectedIndices.length} 張照片`;
                if (!confirm(`要移除${noun}嗎？（不會刪除電腦裡的原檔，可按 Ctrl＋Z 復原）`)) return false;
                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                const selectedSet = new Set(selectedIndices);
                this.photos = this.photos.filter((_, index) => !selectedSet.has(index));
                this.currentIndex = Math.max(0, Math.min(this.currentIndex, this.photos.length - 1));
                this.lastSelectedIndex = this.photos.length ? this.currentIndex : null;
                this.normalizeLinkedChildrenOrder();
                this.render();
                this.recordHistory('移除選取照片');
                return true;
            }

            rotatePhoto(index) {
                const photo = this.photos[index];
                if (!photo) return;
                this.flushPendingHistory();
                const rawDelta = photo.flipH ? -90 : 90;
                this.transformPhotoAnnotationsForRotation(photo, 90);
                photo.rotation = (((photo.rotation || 0) + rawDelta) % 360 + 360) % 360;
                if (photo.sideCrop) this.scheduleSideCropPreviewRefresh(photo);
                this.render();
                this.recordHistory('旋轉照片');
            }

            clearAll() {
                if (!this.photos.length) return;
                if (confirm('要清空清冊中的全部照片嗎？（不會刪除電腦裡的原檔）')) {
                    this.photos = [];
                    this.currentIndex = 0;
                    this.lastSelectedIndex = null;
                    this.activeFilter = 'all';
                    this._dupSet = null;
                    this.currentProjectFileHandle = null;
                    this.currentProjectFileName = '';
                    this.currentProjectSavedAt = null;
                    this.lastLocalEditAt = null;
                    this.savedPhotosByUid = new Map();
                    this.clearAutoDraftDb();
                    this.render();
                    this.recordHistory('清空照片');
                }
            }

            releasePhotoUrls(photo) {
                [photo.src, photo.thumbnailSrc, photo.compositeThumbnailSrc].filter(Boolean).forEach(url => {
                    URL.revokeObjectURL(url);
                    this.managedUrls.delete(url);
                });
            }

            releaseUnusedPhotoUrls() {
                const retained = new Set();
                [...this.photos, ...this.history.flatMap(state => state.photos)].forEach(photo => {
                    if (photo.src) retained.add(photo.src);
                    if (photo.thumbnailSrc) retained.add(photo.thumbnailSrc);
                    if (photo.compositeThumbnailSrc) retained.add(photo.compositeThumbnailSrc);
                });
                this.managedUrls.forEach(url => {
                    if (!retained.has(url)) {
                        URL.revokeObjectURL(url);
                        this.managedUrls.delete(url);
                    }
                });
            }

            updatePageEstimate() {
                const layout = this.layoutSelect.value;
                const count = this.photos.length;
                let perPage = 2;
                if (layout === 'landscape_3') perPage = 3;
                const pages = count === 0 ? 0 : Math.ceil(count / perPage);
                this.photoCount.textContent = count;
                this.pageEstimate.textContent = `（約 ${pages} 頁）`;
            }

            updateLayoutDescription() {
                if (!this.layoutDescription) return;
                const descriptions = {
                    up_down_2: '上下兩張・A4 直式・每頁 2 張',
                    left_right_2: '左右兩張・A4 直式雙欄・每頁 2 張',
                    landscape_3: '橫式三張・A4 橫式・每頁 3 張'
                };
                this.layoutDescription.textContent = descriptions[this.layoutSelect.value] || descriptions.up_down_2;
                this.updateLayoutMiniPreview();
            }

            updateLayoutMiniPreview() {
                const previewEl = this.layoutMiniPreview || document.getElementById('layoutMiniPreview');
                if (!previewEl) return;
                const layout = this.layoutSelect?.value || 'up_down_2';
                let svg = '';
                if (layout === 'up_down_2') {
                    svg = '<svg viewBox="0 0 32 44" class="w-full h-full text-[#5B5CE2]" aria-hidden="true"><rect x="1" y="1" width="30" height="42" rx="2" fill="#FFFFFF" stroke="#C7CBD3" stroke-width="1.2"/><rect x="3.5" y="4" width="25" height="15" rx="1.5" fill="#EEF0FF" stroke="#818CF8" stroke-width="0.8"/><rect x="3.5" y="23" width="25" height="15" rx="1.5" fill="#EEF0FF" stroke="#818CF8" stroke-width="0.8"/><line x1="6" y1="16" x2="26" y2="16" stroke="#A5B4FC" stroke-width="0.8"/><line x1="6" y1="35" x2="26" y2="35" stroke="#A5B4FC" stroke-width="0.8"/></svg>';
                } else if (layout === 'left_right_2') {
                    svg = '<svg viewBox="0 0 32 44" class="w-full h-full text-[#5B5CE2]" aria-hidden="true"><rect x="1" y="1" width="30" height="42" rx="2" fill="#FFFFFF" stroke="#C7CBD3" stroke-width="1.2"/><rect x="3.5" y="4" width="11.5" height="34" rx="1.5" fill="#EEF0FF" stroke="#818CF8" stroke-width="0.8"/><rect x="17" y="4" width="11.5" height="34" rx="1.5" fill="#EEF0FF" stroke="#818CF8" stroke-width="0.8"/></svg>';
                } else if (layout === 'landscape_3') {
                    svg = '<svg viewBox="0 0 44 32" class="w-full h-full text-[#5B5CE2]" aria-hidden="true"><rect x="1" y="1" width="42" height="30" rx="2" fill="#FFFFFF" stroke="#C7CBD3" stroke-width="1.2"/><rect x="3.5" y="4.5" width="11" height="21" rx="1.5" fill="#EEF0FF" stroke="#818CF8" stroke-width="0.8"/><rect x="16.5" y="4.5" width="11" height="21" rx="1.5" fill="#EEF0FF" stroke="#818CF8" stroke-width="0.8"/><rect x="29.5" y="4.5" width="11" height="21" rx="1.5" fill="#EEF0FF" stroke="#818CF8" stroke-width="0.8"/></svg>';
                }
                previewEl.innerHTML = svg;
            }

            // ============================================================
            // 🎯 PaperSwitch 核心：HTML5 Drag & Drop 拖曳排版系統
            // ============================================================
            // 不依賴瀏覽器原生 draggable：Windows 的部分觸控板／卡片子元素會讓 dragstart 不穩定。
            beginCardPointerDrag(index, event) {
                if (event.button !== 0 || this.isTypingTarget(event.target) || event.target.closest('button, a, label, input')) return;
                // 按住 Ctrl、Meta 或 Shift 時為多選／連續選取操作，嚴禁啟動拖曳與覆寫選取狀態
                if (event.ctrlKey || event.metaKey || event.shiftKey) return;
                const card = event.currentTarget;
                if (!(card instanceof HTMLElement) || !this.photos[index]) return;

                if (!this.photos[index].selected) {
                    this.photos.forEach(photo => { photo.selected = false; });
                    this.photos[index].selected = true;
                    this.currentIndex = index;
                    this.lastSelectedIndex = index;
                }
                this.draggedIndices = this.photos
                    .map((photo, photoIndex) => photo.selected ? photoIndex : -1)
                    .filter(photoIndex => photoIndex !== -1);
                this.pointerCardDrag = {
                    pointerId: event.pointerId,
                    sourceCard: card,
                    startX: event.clientX,
                    startY: event.clientY,
                    moved: false,
                    targetCard: null
                };
            }

            updateCardPointerDrag(event) {
                const state = this.pointerCardDrag;
                if (!state || state.pointerId !== event.pointerId) return;

                const deltaX = event.clientX - state.startX;
                const deltaY = event.clientY - state.startY;
                if (!state.moved && Math.hypot(deltaX, deltaY) < 6) return;
                if (!state.moved) {
                    state.moved = true;
                    try {
                        state.sourceCard.setPointerCapture(event.pointerId);
                    } catch (_) {}
                }
                this.draggedIndices.forEach(draggedIndex => {
                    const draggedCard = document.getElementById(`thumb_${this.photos[draggedIndex]?.uid}`);
                    draggedCard?.classList.add('pointer-dragging');
                });
                state.sourceCard.classList.add('pointer-drag-lead');
                state.sourceCard.style.transform = `translate(${deltaX}px, ${deltaY}px)`;
                this.showDragStatus(event.clientX, event.clientY, this.draggedIndices.length);
                this.autoScrollStageWhileDragging(event.clientY);

                // 來源卡片已設為 pointer-events:none，因此可確實命中游標下方的目標卡片。
                const target = document.elementFromPoint(event.clientX, event.clientY)?.closest('.photo-thumb-card');
                if (!target) {
                    this.clearPointerDropTarget(state);
                    return;
                }
                const targetIndex = this.photos.findIndex(photo => `thumb_${photo.uid}` === target.id);
                const targetPhoto = this.photos[targetIndex];
                const isTargetDescendantOfDragged = targetPhoto && this.draggedIndices.some(idx => this.isDescendantOf(targetPhoto, this.photos[idx]?.uid));
                if (targetIndex < 0 || this.draggedIndices.includes(targetIndex) || isTargetDescendantOfDragged) {
                    this.clearPointerDropTarget(state);
                    return;
                }

                if (state.targetCard && state.targetCard !== target) {
                    state.targetCard.classList.remove('pointer-drop-before', 'pointer-drop-after');
                }
                const rect = target.getBoundingClientRect();
                const insertBefore = event.clientX < rect.left + rect.width / 2;
                target.classList.toggle('pointer-drop-before', insertBefore);
                target.classList.toggle('pointer-drop-after', !insertBefore);
                state.targetCard = target;
                this.dragInsertBeforeUid = insertBefore
                    ? this.photos[targetIndex].uid
                    : this.getNextUnmovedPhotoUid(targetIndex);
                this.showInsertOverlay(target, insertBefore);
                event.preventDefault();
            }

            clearPointerDropTarget(state) {
                state.targetCard?.classList.remove('pointer-drop-before', 'pointer-drop-after');
                state.targetCard = null;
                this.dragInsertBeforeUid = null;
                document.getElementById('dragInsertOverlay')?.classList.add('hidden');
            }

            endCardPointerDrag(event, cancelled = false) {
                const state = this.pointerCardDrag;
                if (!state || state.pointerId !== event.pointerId) return;
                try {
                    if (state.sourceCard.hasPointerCapture(event.pointerId)) {
                        state.sourceCard.releasePointerCapture(event.pointerId);
                    }
                } catch (_) {}
                this.pointerCardDrag = null;
                this.clearCardPointerDragVisuals();

                if (!state.moved || cancelled || !state.targetCard) {
                    this.draggedIndices = [];
                    this.dragInsertBeforeUid = null;
                    return;
                }
                this.suppressCardClickUntil = performance.now() + 220;
                this.commitDragSort();
            }

            cancelCardPointerDrag() {
                const state = this.pointerCardDrag;
                if (!state) return;
                try {
                    if (state.sourceCard.hasPointerCapture(state.pointerId)) {
                        state.sourceCard.releasePointerCapture(state.pointerId);
                    }
                } catch (_) {}
                this.pointerCardDrag = null;
                this.draggedIndices = [];
                this.dragInsertBeforeUid = null;
                this.clearCardPointerDragVisuals();
            }

            showDragStatus(clientX, clientY, count) {
                const badge = document.getElementById('dragGhostBadge');
                const text = document.getElementById('dragGhostText');
                if (!badge || !text) return;
                text.textContent = `📄 正在打包移動 ${count} 張`;
                badge.style.left = `${clientX + 16}px`;
                badge.style.top = `${clientY - 44}px`;
            }

            // 拖到舞台區上下緣時自動捲動，讓使用者可連續拖往較後面的照片。
            autoScrollStageWhileDragging(clientY) {
                const rect = this.thumbGrid.getBoundingClientRect();
                const edge = 52;
                const maxStep = 22;
                if (clientY < rect.top + edge) {
                    const ratio = Math.max(0, (clientY - rect.top) / edge);
                    this.thumbGrid.scrollTop -= Math.ceil(maxStep * (1 - ratio));
                } else if (clientY > rect.bottom - edge) {
                    const ratio = Math.max(0, (rect.bottom - clientY) / edge);
                    this.thumbGrid.scrollTop += Math.ceil(maxStep * (1 - ratio));
                }
            }

            showInsertOverlay(targetCard, insertBefore) {
                const overlay = document.getElementById('dragInsertOverlay');
                if (!overlay) return;
                const rect = targetCard.getBoundingClientRect();
                overlay.style.left = `${insertBefore ? rect.left : rect.right}px`;
                overlay.style.top = `${rect.top + rect.height / 2}px`;
                overlay.classList.remove('hidden');
            }

            clearCardPointerDragVisuals() {
                document.querySelectorAll('.photo-thumb-card').forEach(card => {
                    card.style.removeProperty('transform');
                    card.classList.remove('pointer-dragging', 'pointer-drag-lead', 'pointer-drop-before', 'pointer-drop-after');
                });
                const badge = document.getElementById('dragGhostBadge');
                if (badge) {
                    badge.style.left = '-9999px';
                    badge.style.top = '-9999px';
                }
                document.getElementById('dragInsertOverlay')?.classList.add('hidden');
            }

            isTypingTarget(target) {
                return target instanceof HTMLElement && (
                    ['INPUT', 'TEXTAREA', 'SELECT'].includes(target.tagName) || target.isContentEditable
                );
            }

            beginMarqueeSelection(event) {
                if (event.button !== 0 || event.target.closest('.photo-thumb-card')) return;

                const rect = this.thumbGrid.getBoundingClientRect();
                this.marqueeState = {
                    pointerId: event.pointerId,
                    startX: event.clientX - rect.left + this.thumbGrid.scrollLeft,
                    startY: event.clientY - rect.top + this.thumbGrid.scrollTop,
                    additive: event.ctrlKey || event.metaKey,
                    initialSelection: this.photos.map(photo => photo.selected),
                    moved: false
                };
                this.thumbGrid.setPointerCapture(event.pointerId);
                this.selectionMarquee.classList.remove('hidden');
                this.updateMarqueeSelection(event);
            }

            updateMarqueeSelection(event) {
                const state = this.marqueeState;
                if (!state || event.pointerId !== state.pointerId) return;

                const gridRect = this.thumbGrid.getBoundingClientRect();
                const currentX = event.clientX - gridRect.left + this.thumbGrid.scrollLeft;
                const currentY = event.clientY - gridRect.top + this.thumbGrid.scrollTop;
                const left = Math.min(state.startX, currentX);
                const top = Math.min(state.startY, currentY);
                const width = Math.abs(currentX - state.startX);
                const height = Math.abs(currentY - state.startY);
                state.moved ||= width > 4 || height > 4;

                Object.assign(this.selectionMarquee.style, {
                    left: `${left}px`, top: `${top}px`, width: `${width}px`, height: `${height}px`
                });
                if (!state.moved) return;

                const selectionRect = {
                    left: gridRect.left + left - this.thumbGrid.scrollLeft,
                    top: gridRect.top + top - this.thumbGrid.scrollTop,
                    right: gridRect.left + left + width - this.thumbGrid.scrollLeft,
                    bottom: gridRect.top + top + height - this.thumbGrid.scrollTop
                };
                this.photos.forEach((photo, index) => {
                    const card = document.getElementById(`thumb_${photo.uid}`);
                    if (!card) return;
                    const cardRect = card.getBoundingClientRect();
                    const intersects = cardRect.left < selectionRect.right && cardRect.right > selectionRect.left &&
                        cardRect.top < selectionRect.bottom && cardRect.bottom > selectionRect.top;
                    photo.selected = state.additive ? state.initialSelection[index] || intersects : intersects;
                    card.classList.toggle('selected', photo.selected);
                    const checkbox = card.querySelector('input[type="checkbox"]');
                    if (checkbox) checkbox.checked = photo.selected;
                });
                this.updateSelectionUi();
            }

            endMarqueeSelection(event) {
                const state = this.marqueeState;
                if (!state || event.pointerId !== state.pointerId) return;
                if (this.thumbGrid.hasPointerCapture(event.pointerId)) this.thumbGrid.releasePointerCapture(event.pointerId);
                this.selectionMarquee.classList.add('hidden');
                this.marqueeState = null;
                if (!state.moved && !state.additive) this.photos.forEach(photo => photo.selected = false);
                this.render();
                this.recordHistory('框選照片');
            }

            getSelectedIndicesOrCurrent() {
                if (typeof PhotoReportSelection !== 'undefined') {
                    return PhotoReportSelection.getSelectedIndicesOrCurrent(this.photos, this.currentIndex);
                }
                const selectedIndices = this.photos.map((photo, index) => photo.selected ? index : -1).filter(index => index !== -1);
                if (selectedIndices.length) return selectedIndices;
                return this.photos[this.currentIndex] ? [this.currentIndex] : [];
            }

            selectPhotoByKeyboard(delta) {
                if (!this.photos.length || !delta) return;
                const visible = this.getVisiblePhotoIndices();
                if (!visible.length) return;

                let nextIndex;
                if (typeof PhotoReportSelection !== 'undefined') {
                    nextIndex = PhotoReportSelection.computeKeyboardNavIndex({
                        totalPhotos: this.photos.length,
                        currentIndex: this.currentIndex,
                        delta: delta,
                        activeFilter: this.activeFilter,
                        visibleEntries: visible
                    });
                } else if (this.activeFilter !== 'all') {
                    const currentPos = visible.findIndex(({ index }) => index === this.currentIndex);
                    const pos = currentPos < 0 ? 0 : currentPos;
                    const nextPos = Math.max(0, Math.min(pos + delta, visible.length - 1));
                    nextIndex = visible[nextPos].index;
                } else {
                    nextIndex = Math.max(0, Math.min(this.currentIndex + delta, this.photos.length - 1));
                }

                this.photos.forEach((photo, index) => photo.selected = index === nextIndex);
                this.currentIndex = nextIndex;
                this.lastSelectedIndex = nextIndex;
                this.render();
                this.recordHistory('切換目前照片');
                document.getElementById(`thumb_${this.photos[nextIndex].uid}`)?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
            }

            moveSelectedBy(delta) {
                const selectedIndices = this.getSelectedIndicesOrCurrent();
                if (!selectedIndices.length || !delta) return;
                const currentPhoto = this.photos[this.currentIndex];
                const stageSlots = this.getStageSlotsInReadingOrder();
                const selectedSet = new Set(selectedIndices);
                const hasOnlyUnpairedChildSelection = selectedIndices.every(i => {
                    const parent = this.getValidParentPhoto(this.photos[i]);
                    return parent && !selectedSet.has(this.photos.indexOf(parent));
                });

                if (hasOnlyUnpairedChildSelection) {
                    const moving = this.photos.filter((_, i) => selectedSet.has(i));
                    const remaining = this.photos.filter((_, i) => !selectedSet.has(i));
                    const minSelectedIdx = Math.min(...selectedIndices);
                    let anchor = 0;
                    for (let i = 0; i < this.photos.length; i++) {
                        if (i === minSelectedIdx) break;
                        if (!selectedSet.has(i)) anchor++;
                    }
                    const insertAt = Math.max(0, Math.min(anchor + delta, remaining.length));
                    remaining.splice(insertAt, 0, ...moving);
                    this.photos = remaining;
                    this.normalizeLinkedChildrenOrder();
                } else {
                    this.normalizeLinkedChildrenOrder();
                    const blocks = [];
                    this.photos.forEach((p, idx) => {
                        if (!this.isLinkedCropChild(p)) {
                            blocks.push({ items: [p], selected: selectedSet.has(idx) });
                        } else if (blocks.length > 0) {
                            const curBlock = blocks[blocks.length - 1];
                            curBlock.items.push(p);
                            if (selectedSet.has(idx)) curBlock.selected = true;
                        }
                    });
                    const movingBlocks = blocks.filter(b => b.selected);
                    const remainingBlocks = blocks.filter(b => !b.selected);
                    const firstSelectedBlockIdx = blocks.findIndex(b => b.selected);
                    let anchorBlock = 0;
                    for (let i = 0; i < firstSelectedBlockIdx; i++) {
                        if (!blocks[i].selected) anchorBlock++;
                    }
                    const insertBlockAt = Math.max(0, Math.min(anchorBlock + delta, remainingBlocks.length));
                    remainingBlocks.splice(insertBlockAt, 0, ...movingBlocks);
                    this.photos = remainingBlocks.flatMap(b => b.items);
                    this.normalizeLinkedChildrenOrder();
                }

                this.applyStageSlotsByPhotoOrder(stageSlots);
                this.currentIndex = Math.max(0, this.photos.indexOf(currentPhoto));
                this.lastSelectedIndex = this.currentIndex;
                this.render();
                this.recordHistory('移動已選照片');
            }

            moveSelectedToBoundary(boundary) {
                const selectedIndices = this.getSelectedIndicesOrCurrent();
                if (!selectedIndices.length) return;
                const currentPhoto = this.photos[this.currentIndex];
                const stageSlots = this.getStageSlotsInReadingOrder();
                const effectiveSet = new Set(selectedIndices);
                selectedIndices.forEach(idx => {
                    const p = this.photos[idx];
                    if (p && !this.isLinkedCropChild(p)) {
                        this.photos.forEach((item, itemIdx) => {
                            if (this.isDescendantOf(item, p.uid)) effectiveSet.add(itemIdx);
                        });
                    }
                });
                const moving = this.photos.filter((_, index) => effectiveSet.has(index));
                const remaining = this.photos.filter((_, index) => !effectiveSet.has(index));
                this.photos = boundary === 'start' ? [...moving, ...remaining] : [...remaining, ...moving];
                this.normalizeLinkedChildrenOrder();
                this.applyStageSlotsByPhotoOrder(stageSlots);
                this.currentIndex = Math.max(0, this.photos.indexOf(currentPhoto));
                this.lastSelectedIndex = this.currentIndex;
                this.render();
                this.recordHistory('移動已選照片至首尾');
            }

            handleCardDragStart(index, event) {
                if (this.photos[index].selected) {
                    this.draggedIndices = this.photos.map((p, idx) => p.selected ? idx : -1).filter(idx => idx !== -1);
                } else {
                    this.draggedIndices = [index];
                }
                this.dragInsertBeforeUid = this.photos[index]?.uid || null;

                event.dataTransfer.effectAllowed = 'move';
                event.dataTransfer.setData('text/plain', JSON.stringify(this.draggedIndices));

                const ghost = document.getElementById('dragGhostBadge');
                const ghostText = document.getElementById('dragGhostText');
                if (ghost && ghostText) {
                    ghostText.textContent = this.draggedIndices.length > 1
                        ? `已打包 ${this.draggedIndices.length} 張照片移動中`
                        : `移動照片 #${this.photos[index].seq}`;
                    event.dataTransfer.setDragImage(ghost, 30, 20);
                }

                this.createDragPlaceholder(index);
                setTimeout(() => {
                    this.draggedIndices.forEach(idx => {
                        const card = document.getElementById(`thumb_${this.photos[idx]?.uid}`);
                        if (card) {
                            card.classList.add('dragging-stacked');
                            if (idx === index) card.classList.add('drag-lead');
                        }
                    });
                }, 0);
            }

            handleCardDragOver(index, event) {
                event.preventDefault();
                event.stopPropagation();
                event.dataTransfer.dropEffect = 'move';

                if (!this.draggedIndices || this.draggedIndices.length === 0) return;
                const overItem = this.photos[index];
                if (this.draggedIndices.includes(index) || (overItem && this.draggedIndices.some(idx => this.isDescendantOf(overItem, this.photos[idx]?.uid)))) return;

                const card = document.getElementById(`thumb_${this.photos[index].uid}`);
                if (card) {
                    const rect = card.getBoundingClientRect();
                    const midX = rect.left + rect.width / 2;
                    if (event.clientX < midX) {
                        card.classList.add('drag-target-left');
                        card.classList.remove('drag-target-right');
                        this.moveDragPlaceholder(card, false);
                        this.dragInsertBeforeUid = this.photos[index].uid;
                    } else {
                        card.classList.add('drag-target-right');
                        card.classList.remove('drag-target-left');
                        this.moveDragPlaceholder(card, true);
                        this.dragInsertBeforeUid = this.getNextUnmovedPhotoUid(index);
                    }
                }
            }

            handleCardDragLeave(index, event) {
                const card = document.getElementById(`thumb_${this.photos[index].uid}`);
                if (card) {
                    card.classList.remove('drag-target-left', 'drag-target-right');
                }
            }

            handleCardDrop(targetIndex, event) {
                event.preventDefault();
                event.stopPropagation();

                if (!this.draggedIndices || this.draggedIndices.length === 0) {
                    this.handleCardDragEnd();
                    return;
                }

                const dropTarget = this.photos[targetIndex];
                if (this.draggedIndices.includes(targetIndex) || (dropTarget && this.draggedIndices.some(idx => this.isDescendantOf(dropTarget, this.photos[idx]?.uid)))) {
                    this.handleCardDragEnd();
                    return;
                }

                const targetItem = this.photos[targetIndex];
                if (!targetItem) {
                    this.handleCardDragEnd();
                    return;
                }

                const card = document.getElementById(`thumb_${targetItem.uid}`);
                let isAfter = false;
                if (card) {
                    const rect = card.getBoundingClientRect();
                    const midX = rect.left + rect.width / 2;
                    if (event.clientX >= midX) isAfter = true;
                    card.classList.remove('drag-target-left', 'drag-target-right');
                }
                this.dragInsertBeforeUid = isAfter
                    ? this.getNextUnmovedPhotoUid(targetIndex)
                    : targetItem.uid;
                this.commitDragSort();
            }

            createDragPlaceholder(index) {
                this.removeDragPlaceholder();
                const sourceCard = document.getElementById(`thumb_${this.photos[index]?.uid}`);
                if (!sourceCard) return;

                const placeholder = document.createElement('div');
                placeholder.id = 'dragInsertPlaceholder';
                placeholder.className = 'drag-insert-placeholder';
                placeholder.setAttribute('aria-hidden', 'true');
                const count = this.draggedIndices.length;
                placeholder.textContent = count > 1 ? `插入 ${count} 張照片` : '插入位置';
                const sourceHeight = Math.max(76, Math.round(sourceCard.getBoundingClientRect().height));
                placeholder.style.setProperty('--drag-placeholder-height', `${sourceHeight}px`);
                sourceCard.insertAdjacentElement('beforebegin', placeholder);
                this.dragPlaceholder = placeholder;
            }

            moveDragPlaceholder(targetCard, isAfter) {
                if (!this.dragPlaceholder || !targetCard) return;
                targetCard.insertAdjacentElement(isAfter ? 'afterend' : 'beforebegin', this.dragPlaceholder);
            }

            removeDragPlaceholder() {
                this.dragPlaceholder?.remove();
                this.dragPlaceholder = null;
            }

            getNextUnmovedPhotoUid(index) {
                const hovered = this.photos[index];
                const draggingSiblingOfHovered = hovered && this.draggedIndices.some(idx => {
                    const d = this.photos[idx];
                    return d && d.parentUid && (d.parentUid === hovered.uid || d.parentUid === hovered.parentUid);
                });
                const hoveredGroupRootUid = hovered ? (this.isLinkedCropChild(hovered) ? hovered.parentUid : hovered.uid) : null;
                for (let nextIndex = index + 1; nextIndex < this.photos.length; nextIndex++) {
                    if (this.draggedIndices.includes(nextIndex)) continue;
                    const candidate = this.photos[nextIndex];
                    if (!draggingSiblingOfHovered && hoveredGroupRootUid && this.isDescendantOf(candidate, hoveredGroupRootUid)) {
                        continue;
                    }
                    return candidate.uid;
                }
                return null;
            }

            handleStageGridDragOver(event) {
                if (this.isExternalFileDrag(event) || event.target.closest('.photo-thumb-card') || event.target.closest('.drag-insert-placeholder')) return;
                if (!this.draggedIndices?.length) return;
                event.preventDefault();
                event.dataTransfer.dropEffect = 'move';
                this.dragInsertBeforeUid = null;
                const lastUnmovedCard = [...this.thumbGrid.querySelectorAll('.photo-thumb-card')]
                    .filter(card => !card.classList.contains('dragging-stacked'))
                    .at(-1);
                if (lastUnmovedCard) this.moveDragPlaceholder(lastUnmovedCard, true);
            }

            handleStageGridDrop(event) {
                if (this.isExternalFileDrag(event) || event.target.closest('.photo-thumb-card') || event.target.closest('.drag-insert-placeholder')) return;
                if (!this.draggedIndices?.length) return;
                event.preventDefault();
                this.dragInsertBeforeUid = null;
                this.commitDragSort();
            }

            commitDragSort() {
                if (!this.draggedIndices?.length) {
                    this.handleCardDragEnd();
                    return;
                }

                const effectiveSet = new Set(this.draggedIndices);
                this.draggedIndices.forEach(idx => {
                    const draggedPhoto = this.photos[idx];
                    if (draggedPhoto && !this.isLinkedCropChild(draggedPhoto)) {
                        this.photos.forEach((item, itemIdx) => {
                            if (this.isDescendantOf(item, draggedPhoto.uid)) effectiveSet.add(itemIdx);
                        });
                    }
                });

                const leadPhoto = this.photos[this.draggedIndices[0]] || null;
                const movingItems = this.photos.filter((_, idx) => effectiveSet.has(idx));
                const remaining = this.photos.filter((_, idx) => !effectiveSet.has(idx));
                let insertPos = this.dragInsertBeforeUid
                    ? remaining.findIndex(photo => photo.uid === this.dragInsertBeforeUid)
                    : remaining.length;
                if (insertPos < 0) insertPos = remaining.length;

                remaining.splice(insertPos, 0, ...movingItems);
                this.photos = remaining;
                this.normalizeLinkedChildrenOrder();

                this.currentIndex = Math.max(0, leadPhoto ? this.photos.indexOf(leadPhoto) : insertPos);
                this.lastSelectedIndex = this.currentIndex;

                this.handleCardDragEnd();
                this.render();
                this.recordHistory('拖曳排序照片');
            }

            handleCardDragEnd() {
                this.draggedIndices = [];
                this.dragInsertBeforeUid = null;
                this.removeDragPlaceholder();
                document.querySelectorAll('.photo-thumb-card').forEach(c => {
                    c.classList.remove('dragging-stacked', 'drag-lead', 'drag-target-left', 'drag-target-right');
                });
            }

            handleCardClick(index, event) {
                if (performance.now() < this.suppressCardClickUntil) return;
                if (event.target.closest('button, a, label, input')) return;

                if (event.ctrlKey || event.metaKey) {
                    this.photos[index].selected = !this.photos[index].selected;
                    this.lastSelectedIndex = index;
                } else if (event.shiftKey && this.lastSelectedIndex !== null) {
                    const start = Math.min(this.lastSelectedIndex, index);
                    const end = Math.max(this.lastSelectedIndex, index);
                    this.photos.forEach((photo, photoIndex) => photo.selected = photoIndex >= start && photoIndex <= end);
                } else {
                    if (!this.photos[index].selected) {
                        this.photos.forEach((photo, photoIndex) => photo.selected = photoIndex === index);
                    }
                    this.currentIndex = index;
                    this.lastSelectedIndex = index;
                }
                this.currentIndex = index;
                this.render();
                this.recordHistory('變更照片選取');
            }

            handleCardDblClick(index, event) {
                if (event.target.closest('button, a, label, input')) return;
                this.photos.forEach((photo, photoIndex) => {
                    photo.selected = (photoIndex === index);
                });
                this.currentIndex = index;
                this.lastSelectedIndex = index;
                if (this.editorCollapsed) {
                    this.toggleEditor();
                }
                this.render();
                setTimeout(() => {
                    if (this.editDesc) {
                        this.editDesc.focus();
                        const len = this.editDesc.value.length;
                        this.editDesc.setSelectionRange(len, len);
                    }
                }, 50);
            }

            togglePhotoSelect(index, event) {
                event.stopPropagation();
                this.photos[index].selected = !this.photos[index].selected;
                if (this.photos[index].selected) this.currentIndex = index;
                this.lastSelectedIndex = index;
                this.render();
                this.recordHistory('變更照片選取');
            }

            toggleSelectAll(forceState) {
                const targetState = forceState !== undefined ? forceState : !this.photos.every(p => p.selected);
                this.photos.forEach(p => p.selected = targetState);
                this.render();
                this.recordHistory('變更照片選取');
            }

            navigatePhoto(delta) {
                if (this.isLightboxOpen()) {
                    this.navigateLightbox(delta);
                    return;
                }
                // 右側上一張／下一張與 Alt＋方向鍵完全沿用舞台方向鍵流程：
                // 單選目前照片、更新藍框、同步右側欄位並捲動縮圖至可見位置。
                this.selectPhotoByKeyboard(delta);
            }

            updateCurrentPhoto(field, value) {
                const photo = this.photos[this.currentIndex];
                if (!photo) return;
                if (field === 'seq' && this.isLinkedCropChild(photo)) return;
                this.commitLightboxHistoryIfNeeded();
                photo[field] = value;
                if (field === 'time') photo.timeAutoDefaulted = false;
                this.lightboxNeedsRender = true;
                if (field === 'seq') {
                    this.syncChildCropSequences();
                    this.refreshSequenceBadgesAndGridInputs();
                    this.updateEditorTitleAndLightboxCaption();
                } else if (field === 'location' || field === 'desc') {
                    this.refreshThumbCardStatusBadge(this.currentIndex);
                    this.updateAuditBarUi();
                }
                this.scheduleHistory('修改照片欄位');
            }

            // 建立重複檔名 Set（快取於 _dupSet，委派 PhotoReportAudit 模組）
            _buildDupSet() {
                if (typeof PhotoReportAudit !== 'undefined') {
                    this._dupSet = PhotoReportAudit.buildDuplicateNameSet(this.photos);
                    return;
                }
                const nameCount = new Map();
                this.photos.forEach(p => {
                    const n = p.name || '';
                    if (!n) return;
                    nameCount.set(n, (nameCount.get(n) || 0) + 1);
                });
                this._dupSet = new Set([...nameCount.entries()].filter(([, c]) => c > 1).map(([n]) => n));
            }

            // 完整度稽核：委派至 PhotoReportAudit 模組，純讀取不修改 photos
            auditPhotosCompleteness() {
                const globalLocation = document.getElementById('defaultLocation')?.value?.trim() || '';
                this._buildDupSet();
                if (typeof PhotoReportAudit !== 'undefined') {
                    return PhotoReportAudit.auditPhotosCompleteness(this.photos, globalLocation);
                }
                const calcFullWidth = (str) => {
                    let units = 0;
                    for (const ch of String(str ?? '')) {
                        if (ch === '\r' || ch === '\n') continue;
                        const code = ch.codePointAt(0);
                        const isHalf = (code >= 0x0020 && code <= 0x007E) || code === 0x0009 || (code >= 0xFF61 && code <= 0xFF9F);
                        units += isHalf ? 1 : 2;
                    }
                    return units / 2;
                };
                const calcLines = (str) => {
                    const normalized = String(str ?? '').trim().replace(/\r\n?/g, '\n');
                    if (!normalized) return 0;
                    return normalized.split('\n').reduce((sum, seg) => sum + Math.max(1, Math.ceil(calcFullWidth(seg) / 34)), 0);
                };
                let missingLocation = 0, missingDesc = 0, invalidDateTime = 0, duplicatePhotos = 0, overTextLimit = 0;
                this.photos.forEach(photo => {
                    const effectiveLoc = photo.location?.trim() || globalLocation;
                    const descStr = photo.desc?.trim() || '';
                    if (!effectiveLoc) missingLocation++;
                    if (!descStr) missingDesc++;
                    const d = photo.date?.trim() || '';
                    const t = photo.time?.trim() || '';
                    const dateInvalid = d ? !this.isValidMinguoDate(d) : false;
                    const timeInvalid = !this.isValidTimeFormat(t);
                    if (dateInvalid || timeInvalid) invalidDateTime++;
                    if (this._dupSet.has(photo.name)) duplicatePhotos++;
                    if (calcFullWidth(effectiveLoc) > 17 || calcFullWidth(descStr) > 68 || calcLines(descStr) > 2) overTextLimit++;
                });
                const issuesCount = missingLocation + missingDesc + invalidDateTime + duplicatePhotos + overTextLimit;
                let firstIssueFilter = null;
                if (missingDesc > 0) firstIssueFilter = 'missingDesc';
                else if (missingLocation > 0) firstIssueFilter = 'missingLocation';
                else if (invalidDateTime > 0) firstIssueFilter = 'invalidDateTime';
                else if (duplicatePhotos > 0) firstIssueFilter = 'duplicatePhotos';
                else if (overTextLimit > 0) firstIssueFilter = 'overTextLimit';
                return {
                    total: this.photos.length,
                    missingLocation,
                    missingDesc,
                    invalidDateTime,
                    duplicatePhotos,
                    overTextLimit,
                    issuesCount,
                    firstIssueFilter,
                };
            }

            // 切換篩選器（視圖狀態，不進入 Undo/Redo）
            setActiveFilter(filterKey) {
                this.activeFilter = filterKey;
                this._dupSet = null; // 篩選切換時重建 dupSet
                this.render();
            }

            // 漏填檢查列預設收合；手動展開或已在篩選某類問題時才顯示，避免常駐佔用版面。
            applyAuditPanelState(audit) {
                if (!this.auditFilterBar || !this.btnToggleAudit) return;
                const hasPhotos = this.photos.length > 0;
                const open = hasPhotos && (this.auditPanelOpen || this.activeFilter !== 'all');
                this.auditFilterBar.classList.toggle('audit-collapsed', !open);
                this.btnToggleAudit.classList.toggle('hidden', !hasPhotos);
                this.btnToggleAudit.setAttribute('aria-expanded', String(open));
                const issueCount = audit?.issuesCount || 0;
                this.btnToggleAudit.classList.toggle('has-issues', issueCount > 0);
                const summary = issueCount > 0 ? `待檢查 ${issueCount} 項` : '資料檢查通過';
                const label = open ? `收起檢查・${summary}` : summary;
                this.btnToggleAudit.title = '檢查未填地點、未填說明、日期時間異常、同名照片與字數過長；同一張照片可能有多項問題';
                this.btnToggleAudit.replaceChildren();
                const icon = document.createElement('i');
                icon.className = 'fa-solid fa-list-check';
                this.btnToggleAudit.append(icon, document.createTextNode(' ' + label));
            }

            // 頂部「檔案／說明」與工具列「更多」選單：點選項目或點到選單外即收起。
            setupMenus() {
                const menus = Array.from(document.querySelectorAll('details.app-menu'));
                menus.forEach(menu => {
                    menu.addEventListener('toggle', () => {
                        if (menu.open) menus.forEach(other => { if (other !== menu) other.open = false; });
                    });
                    menu.querySelectorAll('.app-menu-item').forEach(item => {
                        item.addEventListener('click', () => setTimeout(() => { menu.open = false; }, 0));
                    });
                });
                document.addEventListener('click', event => {
                    menus.forEach(menu => { if (menu.open && !menu.contains(event.target)) menu.open = false; });
                });
                document.addEventListener('keydown', event => {
                    if (event.key === 'Escape') menus.forEach(menu => { menu.open = false; });
                });
            }

            // 更新篩選列徽章數字與 warning 樣式（委派 PhotoReportAuditUi 模組）
            updateAuditBarUi() {
                if (typeof PhotoReportAuditUi !== 'undefined') {
                    const audit = this.auditPhotosCompleteness();
                    const visibleCount = this.getVisiblePhotoIndices().length;
                    PhotoReportAuditUi.renderAuditBar({
                        filterBar: this.auditFilterBar,
                        totalPhotos: this.photos.length,
                        audit: audit,
                        activeFilter: this.activeFilter,
                        visibleCount: visibleCount,
                        badges: {
                            all: this.auditBadge_all,
                            missingLocation: this.auditBadge_missingLocation,
                            missingDesc: this.auditBadge_missingDesc,
                            invalidDateTime: this.auditBadge_invalidDateTime,
                            duplicatePhotos: this.auditBadge_duplicatePhotos,
                            overTextLimit: this.auditBadge_overTextLimit
                        },
                        filterNoticeText: this.filterNoticeText
                    });
                    this.applyAuditPanelState(audit);
                    return;
                }

                if (!this.auditFilterBar) return;
                const totalPhotos = this.photos.length;
                if (totalPhotos === 0) {
                    this.auditFilterBar.classList.add('hidden');
                    return;
                }
                this.auditFilterBar.classList.remove('hidden');
                const audit = this.auditPhotosCompleteness();
                if (this.auditBadge_all) this.auditBadge_all.textContent = totalPhotos;
                if (this.auditBadge_missingLocation) {
                    this.auditBadge_missingLocation.textContent = audit.missingLocation;
                    this.auditBadge_missingLocation.classList.toggle('has-warning', audit.missingLocation > 0);
                }
                if (this.auditBadge_missingDesc) {
                    this.auditBadge_missingDesc.textContent = audit.missingDesc;
                    this.auditBadge_missingDesc.classList.toggle('has-warning', audit.missingDesc > 0);
                }
                if (this.auditBadge_invalidDateTime) {
                    this.auditBadge_invalidDateTime.textContent = audit.invalidDateTime;
                    this.auditBadge_invalidDateTime.classList.toggle('has-warning', audit.invalidDateTime > 0);
                }
                if (this.auditBadge_duplicatePhotos) {
                    this.auditBadge_duplicatePhotos.textContent = audit.duplicatePhotos;
                    this.auditBadge_duplicatePhotos.classList.toggle('has-warning', audit.duplicatePhotos > 0);
                }
                if (this.auditBadge_overTextLimit) {
                    const overCount = audit.overTextLimit || 0;
                    this.auditBadge_overTextLimit.textContent = overCount;
                    this.auditBadge_overTextLimit.classList.toggle('has-warning', overCount > 0);
                }
                const filterIds = {
                    'all': 'filterBtn_all',
                    'missingLocation': 'filterBtn_missingLocation',
                    'missingDesc': 'filterBtn_missingDesc',
                    'invalidDateTime': 'filterBtn_invalidDateTime',
                    'duplicatePhotos': 'filterBtn_duplicatePhotos',
                    'overTextLimit': 'filterBtn_overTextLimit',
                };
                Object.entries(filterIds).forEach(([key, id]) => {
                    const btn = document.getElementById(id);
                    if (!btn) return;
                    btn.classList.toggle('active', this.activeFilter === key);
                });
                if (this.filterNoticeText) {
                    if (this.activeFilter === 'all') {
                        this.filterNoticeText.classList.add('hidden');
                    } else {
                        const labels = {
                            missingLocation: '未填地點',
                            missingDesc: '未填說明',
                            invalidDateTime: '時間異常',
                            duplicatePhotos: '同名照片',
                            overTextLimit: '字數過長',
                        };
                        const label = labels[this.activeFilter] || '';
                        const count = this.getVisiblePhotoIndices().length;
                        this.filterNoticeText.innerHTML = `<i class="fa-solid fa-filter text-[#5B5CE2]"></i> <span>篩選中：${label}（共 ${count} 張）</span>`;
                        this.filterNoticeText.classList.remove('hidden');
                    }
                }
            }

            hasSequenceOrderMismatch() {
                if (this.photos.length < 2) return false;
                const mainPhotos = [];
                const childCount = new Map();
                for (let i = 0; i < this.photos.length; i++) {
                    const p = this.photos[i];
                    if (this.isLinkedCropChild(p)) {
                        const parent = this.getValidParentPhoto(p);
                        const prev = this.photos[i - 1];
                        if (!parent || !prev || (prev.uid !== parent.uid && !this.isDescendantOf(prev, parent.uid))) {
                            return true;
                        }
                        const subIdx = (childCount.get(parent.uid) || 0) + 1;
                        childCount.set(parent.uid, subIdx);
                        if (String(p.seq || '') !== `${parent.seq}-${subIdx}`) {
                            return true;
                        }
                    } else {
                        mainPhotos.push(p);
                    }
                }
                if (mainPhotos.length < 2) return false;
                const parts = mainPhotos.map(photo => String(photo.seq || '').match(/^(.*?)(\d+)$/));
                // 支援重新編號功能的前綴、起始值與補零；自由文字編號不強制套用數字規則。
                if (parts.some(part => !part)) return false;
                const prefix = parts[0][1];
                const start = Number(parts[0][2]);
                return parts.some((part, index) => part[1] !== prefix || Number(part[2]) !== start + index);
            }

            // 匯出前編號與完整度檢查，保留使用者直接匯出的選擇。
            confirmExportWithAudit(action, title) {
                if (!this.photos.length) {
                    action();
                    return;
                }
                const audit = this.auditPhotosCompleteness();
                const sequenceMismatch = this.hasSequenceOrderMismatch();
                if (audit.issuesCount === 0 && !sequenceMismatch) {
                    action();
                    return;
                }

                document.getElementById('exportSequenceNotice').hidden = !sequenceMismatch;
                document.getElementById('btnAuditRenumber').hidden = !sequenceMismatch;
                this.btnAuditViewIssues.hidden = audit.issuesCount === 0;
                this.exportAuditList.hidden = audit.issuesCount === 0;
                document.getElementById('exportAuditHint').textContent = sequenceMismatch
                    ? '照片編號未依目前順序連續排列。可先重新編號，也可保留原編號匯出。'
                    : '以下資料可再確認；也可以直接匯出。';
                this.btnAuditProceedExport.querySelector('span').textContent = sequenceMismatch ? '保留編號匯出' : '仍要匯出';

                // 有問題時顯示提醒彈窗，暫存 action
                this.pendingExportAction = action;
                this.pendingExportTitle = title;
                const titleEl = this.exportAuditModalTitle || this.exportAuditModal?.querySelector('#exportAuditModalTitle') || this.exportAuditModal?.querySelector('#exportAuditTitle');
                if (typeof PhotoReportModalUi !== 'undefined') {
                    PhotoReportModalUi.showExportAuditPrompt({
                        modal: this.exportAuditModal,
                        titleEl: titleEl,
                        listEl: this.exportAuditList,
                        exportTitle: title,
                        audit: audit
                    });
                    return;
                }

                if (this.exportAuditList) {
                    const items = [];
                    if (audit.missingDesc > 0) items.push(`未填說明：${audit.missingDesc} 張`);
                    if (audit.missingLocation > 0) items.push(`未填地點：${audit.missingLocation} 張`);
                    if (audit.invalidDateTime > 0) items.push(`時間異常：${audit.invalidDateTime} 張`);
                    if (audit.duplicatePhotos > 0) items.push(`同名照片：${audit.duplicatePhotos} 張`);
                    if (audit.overTextLimit > 0) items.push(`地點或說明過長（恐影響排版）：${audit.overTextLimit} 張`);
                    this.exportAuditList.innerHTML = items.map(t => `<li>${t}</li>`).join('');
                }
                if (titleEl) titleEl.textContent = `匯出前確認（${title}）`;
                this.exportAuditModal?.classList.remove('hidden');
            }

            updateSelectionUi() {
                const selectedCount = this.photos.filter(photo => photo.selected).length;
                if (selectedCount > 0) {
                    this.selectedCountBadge.classList.remove('hidden');
                    this.selectedCountBadge.textContent = `已勾選 ${selectedCount} 張`;
                    this.batchApplyBox.classList.remove('hidden');
                    this.batchCountNum.textContent = selectedCount;
                    this.editorModeBadge.textContent = `一次改 ${selectedCount} 張`;
                    this.editorModeBadge.className = 'text-[10px] font-bold px-2 py-0.5 rounded bg-blue-700 text-white whitespace-nowrap';
                } else {
                    this.selectedCountBadge.classList.add('hidden');
                    this.batchApplyBox.classList.add('hidden');
                    this.editorModeBadge.textContent = '單張編輯';
                    this.editorModeBadge.className = 'text-[10px] font-bold px-2 py-0.5 rounded bg-[#5B5CE2]/10 text-[#3730A3] whitespace-nowrap';
                }
                return selectedCount;
            }

            updateToolbarButtons(totalPhotos, selectedCount) {
                const hasPhotos = totalPhotos > 0;
                const hasSelection = selectedCount > 0;

                const setBtnState = (id, enabled) => {
                    const btn = document.getElementById(id);
                    if (btn) {
                        btn.disabled = !enabled;
                        btn.classList.toggle('opacity-40', !enabled);
                        btn.classList.toggle('pointer-events-none', !enabled);
                    }
                };

                setBtnState('btnToggleSelectAll', hasPhotos);
                setBtnState('btnOpenSequenceModal', hasPhotos);
                setBtnState('btnParseFilenameTime', hasPhotos);
                setBtnState('btnBatchRotate', hasPhotos);
                setBtnState('btnAutoArrangeStage', hasPhotos);
                setBtnState('btnSortByCaptureTime', hasPhotos);
                setBtnState('btnClearAll', hasPhotos);
                setBtnState('btnToggleEditor', hasPhotos);
                setBtnState('btnExportPdf', hasPhotos);
                setBtnState('btnExportDocx', hasPhotos);
                document.getElementById('exportEmptyHint').hidden = hasPhotos;

                setBtnState('btnRemoveSelected', hasSelection);
                setBtnState('btnApplyDefaultsToSelected', hasSelection);
                setBtnState('btnMoveSelectedBackward', hasSelection);
                setBtnState('btnMoveSelectedForward', hasSelection);

                document.getElementById('batchRotateText').textContent = hasSelection ? '旋轉已選' : '旋轉全部';
                document.getElementById('btnBatchRotate').title = `將${hasSelection ? `已勾選的 ${selectedCount}` : `全部 ${totalPhotos}`} 張照片順時針旋轉 90°`;
                const applyDefaultsButton = document.getElementById('btnApplyDefaultsToSelected');
                applyDefaultsButton.textContent = hasSelection ? `套用日期與地點至已選 ${selectedCount} 張` : '請先勾選要套用日期與地點的照片';
                applyDefaultsButton.title = hasSelection
                    ? `將案件日期與地點寫入已勾選的 ${selectedCount} 張照片；之後修改案件資料，這些照片不會跟著變`
                    : '勾選照片後，即可套用案件日期與地點';

                // 空白狀態只保留匯入後才需要設定的檢視控制，避免停用按鈕佔滿操作區。
                const photoActionToolbar = document.getElementById('photoActionToolbar');
                const workspaceToolbar = document.getElementById('workspaceToolbar');
                if (photoActionToolbar) photoActionToolbar.classList.toggle('hidden', !hasPhotos);
                document.getElementById('stageStatusBar')?.classList.toggle('hidden', !hasPhotos);
                if (workspaceToolbar) {
                    workspaceToolbar.classList.toggle('pt-1.5', hasPhotos);
                    workspaceToolbar.classList.toggle('border-t', hasPhotos);
                }

                const stageImportBar = document.getElementById('stageImportBar');
                if (stageImportBar) {
                    stageImportBar.classList.toggle('hidden', !hasPhotos);
                }
            }

            // 渲染介面
            render() {
                this.updatePageEstimate();

                const totalPhotos = this.photos.length;
                const selectedCount = this.updateSelectionUi();
                this.updateAuditBarUi();
                const visibleEntries = this.getVisiblePhotoIndices();
                const defaultLocation = document.getElementById('defaultLocation').value.trim();

                document.getElementById('selectAllText').textContent = (totalPhotos > 0 && selectedCount === totalPhotos) ? '全不選' : '全選';
                this.updateToolbarButtons(totalPhotos, selectedCount);
                this.updateProjectSaveStatusUi();

                if (totalPhotos === 0) {
                    if (this.isLightboxOpen()) {
                        this.lightboxNeedsRender = false;
                        this.closeModal('imageLightboxModal');
                    }
                    this.emptyState.classList.remove('hidden');
                    this.workspaceContainer.classList.add('hidden');
                    this.gridContainer.classList.add('hidden');
                    if (this.previewContainer) this.previewContainer.classList.add('hidden');
                    return;
                }

                this.emptyState.classList.add('hidden');
                this.workspaceContainer.classList.toggle('hidden', this.currentView !== 'workspace');
                this.gridContainer.classList.toggle('hidden', this.currentView !== 'grid');
                if (this.previewContainer) this.previewContainer.classList.toggle('hidden', this.currentView !== 'preview');

                if (this.currentIndex >= totalPhotos) this.currentIndex = totalPhotos - 1;
                if (this.currentIndex < 0) this.currentIndex = 0;

                // 更新右側編輯器資料
                this.populateEditorFields();

                // 渲染左側縮圖卡片
                this.thumbGrid.innerHTML = '<div id="selectionMarquee" class="selection-marquee hidden" aria-hidden="true"></div>' + visibleEntries.map(({ photo: p, index: idx }) => {
                    const isSelected = p.selected ? 'selected' : '';
                    const isCurrent = idx === this.currentIndex ? 'active-editing' : '';
                    const isCropChild = this.isLinkedCropChild(p) ? 'photo-thumb-card--crop-child' : '';
                    const missingFields = [!String(p.location || '').trim() && !defaultLocation ? '地點' : '', !String(p.desc || '').trim() ? '說明' : ''].filter(Boolean);
                    const overLimit = this.inspectPhotoTextLimits(String(p.location || '').trim() || defaultLocation, String(p.desc || '').trim()).isOver;
                    const statusIssues = [...missingFields, ...(overLimit ? ['字數過長'] : [])];
                    const statusText = statusIssues.length ? `注意：${statusIssues.join('、')}` : '資料完整';
                    const displayLayout = this.getPhotoDisplayLayout(p);
                    const cardTitle = `第 ${p.seq} 張｜時間：${p.time || '未填'}｜${statusText}｜雙擊縮圖放大、雙擊編號列編輯，拖曳可排序`;

                    return `
                        <div class="photo-thumb-card morandi-card p-1.5 flex flex-col gap-1 bg-white border border-[#C7CBD3]/25 rounded-xl ${isSelected} ${isCurrent} ${isCropChild}"
                             id="thumb_${p.uid}"
                             style="--thumbnail-aspect:${displayLayout.safeAspect};"
                             title="${escapeHtml(cardTitle)}"
                             draggable="false"
                             onclick="window.app.handleCardClick(${idx}, event)"
                             ondblclick="window.app.handleCardDblClick(${idx}, event)"
                             onpointerdown="window.app.beginCardPointerDrag(${idx}, event)">

                            <div class="thumbnail-meta-row flex items-center justify-between min-h-5">
                                <div class="flex min-w-0 items-center gap-1.5">
                                    <input type="checkbox" ${p.selected ? 'checked' : ''} onclick="window.app.togglePhotoSelect(${idx}, event)" class="rounded text-blue-600 cursor-pointer" title="勾選">
                                    <span class="game-seq-badge px-1.5 py-0.5 rounded text-[12px]">#${escapeHtml(p.seq)}</span>
                                    <span class="thumbnail-card-time text-[12px] font-bold text-slate-800 whitespace-nowrap flex items-center gap-0.5" title="${escapeHtml(p.time || '未填時間')}">
                                        <i class="fa-regular fa-clock text-blue-600"></i>${escapeHtml(p.time || '未填時間')}
                                    </span>
                                </div>
                                <span class="thumb-status-icon text-[13px] ${statusIssues.length ? 'text-amber-700' : 'text-emerald-700'}" title="${escapeHtml(statusText)}">
                                    <i class="fa-solid ${statusIssues.length ? 'fa-circle-exclamation' : 'fa-circle-check'}"></i>
                                </span>
                            </div>

                            <div class="thumbnail-stage w-full bg-[#EEF0F3] rounded-lg overflow-hidden relative cursor-zoom-in"
                                 title="雙擊放大檢視"
                                 ondblclick="event.stopPropagation(); window.app.openLightboxByIndex(${idx})">
                                <img src="${displayLayout.src || p.thumbnailSrc}" alt="${escapeHtml(p.name)}" loading="lazy" decoding="async" style="${displayLayout.imageStyle}" class="absolute left-1/2 top-1/2 object-contain pointer-events-none">
                                ${this.getPhotoAnnotationBadgeHtml(p)}
                            </div>

                        </div>
                    `;
                }).join('');
                this.selectionMarquee = document.getElementById('selectionMarquee');
                this.thumbGrid.style.removeProperty('min-height');

                // 渲染試算表視圖 (DataGrid)
                this.gridTbody.innerHTML = this.photos.map((p, idx) => {
                    const locCheck = this.inspectPhotoTextLimits(String(p.location || '').trim(), '');
                    const descCheck = this.inspectPhotoTextLimits('', String(p.desc || '').trim());
                    return `
                    <tr class="hover:bg-[#F7F7F5]/70 ${idx === this.currentIndex ? 'bg-blue-50/50' : ''}"
                        ondblclick="window.app.handleCardDblClick(${idx}, event)">
                        <td class="p-2 text-center">
                            <input type="checkbox" ${p.selected ? 'checked' : ''} onclick="window.app.togglePhotoSelect(${idx}, event)" class="rounded text-[#5B5CE2] cursor-pointer">
                        </td>
                        <td class="p-2 text-center">
                            <img src="${(p.sideCrop && p.compositeThumbnailSrc) ? p.compositeThumbnailSrc : p.thumbnailSrc}" style="transform: ${(p.sideCrop && p.compositeThumbnailSrc) ? 'none' : `${p.flipH ? 'scaleX(-1) ' : ''}rotate(${p.rotation}deg)`};" class="w-10 h-10 object-contain mx-auto rounded border bg-[#EEF0F3] cursor-pointer" onclick="event.stopPropagation(); window.app.openLightboxByIndex(${idx})" title="點擊放大">
                        </td>
                        <td class="p-2">
                            ${(() => {
                                const parentPhoto = this.getValidParentPhoto(p);
                                return parentPhoto
                                    ? `<input type="text" value="${escapeHtml(p.seq)}" readonly class="input-field w-full text-xs font-bold seq-locked" title="隨原圖 #${escapeHtml(parentPhoto.seq)} 自動連動鎖定"><div class="grid-crop-parent-hint text-[10px] text-sky-700 font-bold mt-0.5 whitespace-nowrap">↳ #${escapeHtml(parentPhoto.seq)} 特寫</div>`
                                    : `<input type="text" value="${escapeHtml(p.seq)}" oninput="window.app.updatePhotoField(${idx}, 'seq', this.value)" class="input-field w-full text-xs font-bold text-[#5B5CE2]">`;
                            })()}
                        </td>
                        <td class="p-2">
                            <input type="text" value="${escapeHtml(p.date || '')}" onpointerdown="window.app.handleMaskedPointerDown(event, this)" onfocus="window.app.beginMaskedEntry(this)" oninput="this.value = window.app.sanitizeMaskedEntry(this, 7); window.app.markMaskedEntryChanged(this); window.app.updatePhotoField(${idx}, 'date', this.value)" onblur="this.value = window.app.finishMaskedEntry(this, window.app.formatDateMask); window.app.updatePhotoField(${idx}, 'date', this.value)" class="input-field w-full text-xs font-mono" placeholder="留空＝清冊預設">
                        </td>
                        <td class="p-2">
                            <input type="text" value="${escapeHtml(p.time)}" onpointerdown="window.app.handleMaskedPointerDown(event, this)" onfocus="window.app.beginMaskedEntry(this)" oninput="this.value = window.app.sanitizeMaskedEntry(this, 6); window.app.markMaskedEntryChanged(this); window.app.updatePhotoField(${idx}, 'time', this.value)" onblur="this.value = window.app.finishMaskedEntry(this, window.app.formatTimeMask); window.app.updatePhotoField(${idx}, 'time', this.value)" class="input-field w-full text-xs font-mono" placeholder="輸入 4 或 6 碼">
                        </td>
                        <td class="p-2">
                            <input type="text" data-grid-index="${idx}" data-grid-field="location" value="${escapeHtml(p.location)}" class="input-field w-full text-xs ${locCheck.locationOver ? 'input-over-limit' : ''}" placeholder="留空＝清冊預設" title="地點最多 17 個全形（2 半形＝1 全形）">
                        </td>
                        <td class="p-2">
                            <input type="text" data-grid-index="${idx}" data-grid-field="desc" value="${escapeHtml(p.desc)}" class="input-field w-full text-xs ${descCheck.descOver ? 'input-over-limit' : ''}" placeholder="請填寫現場跡證或狀況說明" title="說明最多 68 個全形・最多 2 行（2 半形＝1 全形）">
                        </td>
                        <td class="p-2 text-[#6B7280] truncate max-w-[110px]" title="${escapeHtml(p.name)}">
                            ${escapeHtml(p.name)}
                        </td>
                        <td class="p-2 text-center">
                            <button type="button" onclick="window.app.rotatePhoto(${idx})" title="旋轉" class="p-1 text-[#5B5CE2]"><i class="fa-solid fa-rotate-right"></i></button>
                            <button type="button" onclick="window.app.removePhoto(${idx})" title="刪除" class="p-1 text-red-500"><i class="fa-regular fa-trash-can"></i></button>
                        </td>
                    </tr>
                `;
                }).join('');
                this.gridTbody.querySelectorAll('input[data-grid-field="location"], input[data-grid-field="desc"]').forEach(gridInput => {
                    this.syncTextLimitInputState(gridInput, gridInput.value || '');
                });
                this.renderPagePreview();
                if (this.isLightboxOpen()) {
                    this.lightboxIndex = this.currentIndex;
                    this.updateLightbox({ preserveZoomIfSamePhoto: true });
                }
            }

            updateEditorTitleAndLightboxCaption() {
                const curPhoto = this.photos[this.currentIndex];
                if (!curPhoto) return;
                const totalPhotos = this.photos.length;
                this.editorTitle.replaceChildren();
                const titleIcon = document.createElement('i');
                titleIcon.className = 'fa-solid fa-pen-to-square text-[#5B5CE2]';
                const titleText = document.createElement('span');
                titleText.textContent = `正在編輯第 ${this.currentIndex + 1}／${totalPhotos} 張（編號 ${curPhoto.seq}）`;
                this.editorTitle.append(titleIcon, document.createTextNode(' '), titleText);
                if (this.lightboxCaption && this.isLightboxOpen() && this.lightboxIndex === this.currentIndex) {
                    const parentPhoto = this.getValidParentPhoto(curPhoto);
                    this.lightboxCaption.textContent = parentPhoto
                        ? `#${curPhoto.seq}（#${parentPhoto.seq} 局部特寫）　${curPhoto.name}`
                        : `#${curPhoto.seq}　${curPhoto.name}`;
                }
            }

            populateEditorFields() {
                const totalPhotos = this.photos.length;
                const curPhoto = this.photos[this.currentIndex];
                if (!curPhoto) return;
                const defaultLocation = document.getElementById('defaultLocation').value.trim();
                this.updateEditorTitleAndLightboxCaption();
                this.editSeq.value = curPhoto.seq || '';
                const parentPhoto = this.getValidParentPhoto(curPhoto);
                if (parentPhoto) {
                    this.editSeq.readOnly = true;
                    this.editSeq.classList.add('seq-locked');
                    this.editSeq.title = `已鎖定為原圖 #${parentPhoto.seq} 的局部特寫，重排或修改原圖編號時會自動連動`;
                    this.editSeqLockNotice?.classList.remove('hidden');
                    if (this.editSeqLockText) {
                        this.editSeqLockText.innerHTML = `<i class="fa-solid fa-link text-sky-700"></i> 隨原圖 #${escapeHtml(parentPhoto.seq)} 自動連動鎖定`;
                    }
                } else {
                    this.editSeq.readOnly = false;
                    this.editSeq.classList.remove('seq-locked');
                    this.editSeq.title = '';
                    this.editSeqLockNotice?.classList.add('hidden');
                }
                this.editDate.value = curPhoto.date || '';
                this.editTime.value = curPhoto.time || '';
                this.editLocation.value = curPhoto.location || '';
                this.editDesc.value = curPhoto.desc || '';
                this.syncTextLimitInputState(this.defaultLocationInput, this.defaultLocationInput?.value || '');
                this.syncTextLimitInputState(this.editLocation, this.editLocation.value);
                this.syncTextLimitInputState(this.editDesc, this.editDesc.value);
                const remaining = totalPhotos - this.currentIndex - 1;
                ['btnFillDownDate', 'btnFillDownTime', 'btnFillDownLocation', 'btnFillDownDescAll'].forEach(id => {
                    const button = document.getElementById(id);
                    button.disabled = remaining === 0;
                    button.title = remaining === 0 ? '沒有後續照片可填入' : id === 'btnFillDownDescAll'
                        ? `填入後續 ${remaining} 張的空白說明，保留已有內容`
                        : `套用至後續 ${remaining} 張，將取代原有內容；覆寫不同內容前會先確認`;
                });
                document.getElementById('fillDownHint').textContent = remaining
                    ? `向下填滿範圍：後續 ${remaining} 張。日期、時間與地點會取代原值；說明只填空白。`
                    : '這是最後一張照片，沒有後續照片可填入。';
                const defaultDate = document.getElementById('caseDate').value.trim();
                this.editDateInheritedHint.textContent = defaultDate
                    ? `空白時用案件日期：${defaultDate}`
                    : '空白時用案件日期（案件日期還沒填）';
                this.editLocationInheritedHint.textContent = defaultLocation
                    ? `空白時用案件地點：${defaultLocation}`
                    : '空白時用案件地點（案件地點還沒填）';
                this.refreshAllTextLimitUi();
            }

            renderPagePreview() {
                const perPage = this.getPhotosPerPage();
                const layout = this.layoutSelect.value;
                const defaultLocation = document.getElementById('defaultLocation').value.trim();
                const gridClass = layout === 'up_down_2' ? 'grid-cols-1 max-w-[420px]' :
                    layout === 'left_right_2' ? 'grid-cols-2' : 'grid-cols-3';
                const pages = [];
                for (let start = 0; start < this.photos.length; start += perPage) {
                    const pagePhotos = this.photos.slice(start, start + perPage);
                    const cells = pagePhotos.map((photo, offset) => {
                        const index = start + offset;
                        return `
                            <button type="button" onclick="window.app.selectPreviewPhoto(${index})" class="text-left min-w-0 border border-[#C7CBD3]/25 rounded-lg p-1.5 bg-white hover:border-[#5B5CE2] transition ${index === this.currentIndex ? 'ring-2 ring-[#5B5CE2]/50' : ''}">
                                <div class="aspect-[4/3] bg-[#EEF0F3] rounded flex items-center justify-center overflow-hidden">
                                    <img src="${(photo.sideCrop && photo.compositeThumbnailSrc) ? photo.compositeThumbnailSrc : photo.thumbnailSrc}" data-preview-uid="${photo.uid}" alt="${escapeHtml(photo.name)}" style="transform: ${(photo.sideCrop && photo.compositeThumbnailSrc) ? 'none' : `${photo.flipH ? 'scaleX(-1) ' : ''}rotate(${photo.rotation}deg)`};" class="w-full h-full object-contain">
                                </div>
                                <div class="mt-1 text-[10px] font-bold text-[#20242B] truncate">#${escapeHtml(photo.seq)} ${escapeHtml(photo.location || defaultLocation || '未填地點')}</div>
                                <div class="text-[10px] text-[#6B7280] truncate">${escapeHtml(photo.desc || '（無說明）')}</div>
                            </button>`;
                    }).join('');
                    pages.push(`
                        <section class="bg-white border-2 border-[#C7CBD3]/25 rounded-xl shadow-sm p-3">
                            <div class="flex items-center justify-between mb-2 text-[11px] text-[#6B7280]">
                                <span class="font-bold text-[#20242B]">第 ${pages.length + 1} 頁</span>
                                <span>${layout === 'up_down_2' ? '直式上下兩張' : layout === 'left_right_2' ? '直式左右兩張' : '橫式三張'}</span>
                            </div>
                            <div class="grid ${gridClass} gap-2 mx-auto">${cells}</div>
                        </section>`);
                }
                if (!this.previewContainer) return;
                this.previewContainer.innerHTML = `
                    <div class="mb-3 text-xs text-[#6B7280] flex items-center gap-1.5"><i class="fa-solid fa-circle-info text-[#5B5CE2]"></i> 分頁預覽僅供確認照片順序與分組；實際表格以 Word 匯出為準。</div>
                    <div class="grid grid-cols-1 xl:grid-cols-2 gap-4">${pages.join('')}</div>`;
            }

            selectPreviewPhoto(index) {
                if (!this.photos[index]) return;
                this.photos.forEach((photo, photoIndex) => photo.selected = photoIndex === index);
                this.currentIndex = index;
                this.lastSelectedIndex = index;
                this.render();
                this.recordHistory('從分頁預覽選取照片');
            }

            updatePhotoField(index, field, value) {
                const photo = this.photos[index];
                if (!photo) return;
                if (field === 'seq' && this.isLinkedCropChild(photo)) return;
                photo[field] = value;
                if (field === 'time') photo.timeAutoDefaulted = false;
                if (field === 'seq') {
                    this.syncChildCropSequences();
                    this.refreshSequenceBadgesAndGridInputs();
                    this.updateEditorTitleAndLightboxCaption();
                } else if (field === 'location' || field === 'desc') {
                    this.refreshThumbCardStatusBadge(index);
                    this.updateAuditBarUi();
                    if (index === this.currentIndex) {
                        const targetEl = field === 'location' ? this.editLocation : this.editDesc;
                        if (targetEl && targetEl.value !== value) {
                            targetEl.value = value;
                            this.syncTextLimitInputState(targetEl, value);
                        }
                        this.refreshAllTextLimitUi();
                    }
                }
                this.scheduleHistory('修改照片欄位');
            }

            initLightboxInteractions() {
                if (!this.lightboxViewport) return;
                this.lightboxViewport.addEventListener('wheel', (e) => this.handleLightboxWheel(e), { passive: false });
                const supportsPointer = typeof window.PointerEvent !== 'undefined';
                if (supportsPointer) {
                    this.lightboxViewport.addEventListener('pointerdown', (e) => this.handleLightboxPointerDown(e));
                } else {
                    this.lightboxViewport.addEventListener('mousedown', (e) => this.handleLightboxPointerDown(e));
                }
                if (this.lightboxAnnotationLayer) {
                    if (supportsPointer) {
                        this.lightboxAnnotationLayer.addEventListener('pointerdown', (e) => this.handleAnnotationLayerPointerDown(e));
                    } else {
                        this.lightboxAnnotationLayer.addEventListener('mousedown', (e) => this.handleAnnotationLayerPointerDown(e));
                    }
                    this.lightboxAnnotationLayer.addEventListener('dblclick', (e) => this.handleAnnotationLayerDblClick(e));
                    this.lightboxAnnotationLayer.addEventListener('click', (e) => this.handleAnnotationLayerClick(e));
                }
                if (this.lightboxSideCropArea) {
                    if (supportsPointer) {
                        this.lightboxSideCropArea.addEventListener('pointerdown', (e) => this.handleSideCropAreaPointerDown(e));
                    } else {
                        this.lightboxSideCropArea.addEventListener('mousedown', (e) => this.handleSideCropAreaPointerDown(e));
                    }
                    this.lightboxSideCropArea.addEventListener('click', (e) => this.handleSideCropAreaClick(e));
                }
                if (this.lightboxImg) {
                    this.lightboxImg.addEventListener('load', () => this.syncLightboxAnnotationLayerBounds());
                    if (typeof ResizeObserver !== 'undefined') {
                        this.lightboxImgResizeObserver = new ResizeObserver(() => {
                            if (this.isLightboxOpen()) this.syncLightboxAnnotationLayerBounds();
                        });
                        this.lightboxImgResizeObserver.observe(this.lightboxImg);
                        if (this.lightboxViewport) this.lightboxImgResizeObserver.observe(this.lightboxViewport);
                    }
                }
                if (this.lightboxLabelTextInput) {
                    this.lightboxLabelTextInput.addEventListener('focus', () => {
                        if (this.lightboxTool !== 'label') this.setLightboxTool('label');
                    });
                    this.lightboxLabelTextInput.addEventListener('input', () => {
                        if (this.lightboxTool !== 'label') this.setLightboxTool('label');
                    });
                    this.lightboxLabelTextInput.addEventListener('keydown', (e) => {
                        if (e.key === 'Enter') {
                            e.preventDefault();
                            this.lightboxLabelTextInput.blur();
                        }
                    });
                }
                if (this.lightboxLabelColorSelect) {
                    this.lightboxLabelColorSelect.addEventListener('change', () => {
                        this.lightboxLabelColorSelect.blur();
                    });
                }
                window.addEventListener('resize', () => {
                    if (this.isLightboxOpen()) this.syncLightboxAnnotationLayerBounds();
                });
                if (supportsPointer) {
                    window.addEventListener('pointermove', (e) => this.handleLightboxPointerMove(e));
                    window.addEventListener('pointerup', (e) => this.handleLightboxPointerUp(e));
                    window.addEventListener('pointercancel', (e) => this.handleLightboxPointerUp(e));
                } else {
                    window.addEventListener('mousemove', (e) => this.handleLightboxPointerMove(e));
                    window.addEventListener('mouseup', (e) => this.handleLightboxPointerUp(e));
                }
                window.addEventListener('blur', () => this.cancelLightboxInteractions());
            }

            handleLightboxBackdropPointerDown(e) {
                this.lightboxBackdropPointerDown = Boolean(e && e.target === this.imageLightboxModal);
            }

            handleLightboxBackdropClick(e) {
                const shouldClose = Boolean(e && e.target === this.imageLightboxModal && this.lightboxBackdropPointerDown);
                this.lightboxBackdropPointerDown = false;
                if (shouldClose) {
                    this.closeModal('imageLightboxModal');
                }
            }

            cancelLightboxInteractions() {
                this.lightboxBackdropPointerDown = false;
                if (this.lightboxAnnotationAction) {
                    const action = this.lightboxAnnotationAction;
                    this.lightboxAnnotationAction = null;
                    const photo = this.photos[this.lightboxIndex];
                    if ((action.mode === 'move-side-crop' || action.mode === 'resize-side-crop') && photo && action.origSideCrop) {
                        photo.sideCrop = this.cloneSideCrop(action.origSideCrop);
                    }
                    if (action.mode === 'draw-crop' && !action.moved) {
                        this.lightboxCropDraft = null;
                    }
                    if (photo && Array.isArray(photo.annotations) && action.mode === 'draw-rect' && !action.moved) {
                        photo.annotations = photo.annotations.filter(item => item.id !== action.annId);
                    }
                    if (this.isLightboxOpen()) this.renderLightboxAnnotations();
                }
                if (this.lightboxIsPanning) {
                    this.lightboxIsPanning = false;
                    this.lightboxViewport?.classList.remove('is-panning');
                    this.lightboxStage?.classList.remove('panning-active');
                }
            }

            setLightboxTool(tool) {
                const allowed = ['pan', 'rect-yellow', 'rect-red', 'label', 'crop'];
                const nextTool = allowed.includes(tool) ? tool : 'pan';
                const leavingCropWithDraft = this.lightboxTool === 'crop' && nextTool !== 'crop' && this.lightboxCropDraft;
                this.lightboxTool = nextTool;
                if (leavingCropWithDraft) {
                    this.lightboxCropDraft = null;
                }
                if (this.lightboxTool === 'rect-yellow' && this.lightboxLabelColorSelect) {
                    this.lightboxLabelColorSelect.value = 'yellow';
                } else if (this.lightboxTool === 'rect-red' && this.lightboxLabelColorSelect) {
                    this.lightboxLabelColorSelect.value = 'red';
                }
                const btnMap = {
                    'pan': this.btnLightboxToolPan,
                    'rect-yellow': this.btnLightboxToolRectYellow,
                    'rect-red': this.btnLightboxToolRectRed,
                    'label': this.btnLightboxToolLabel,
                    'crop': this.btnLightboxToolCrop
                };
                Object.entries(btnMap).forEach(([key, btn]) => {
                    if (btn) btn.classList.toggle('active', key === this.lightboxTool);
                });
                if (this.lightboxAnnotationLayer) {
                    this.lightboxAnnotationLayer.classList.toggle('tool-drawing', this.lightboxTool !== 'pan');
                    if (leavingCropWithDraft && this.isLightboxOpen()) {
                        this.renderLightboxAnnotations();
                    }
                }
            }

            getLightboxImgTransform(photo) {
                const flip = photo?.flipH ? 'scaleX(-1) ' : '';
                return `${flip}rotate(${photo?.rotation || 0}deg)`;
            }

            toggleLightboxFlipH() {
                if (this.lightboxIndex === null || !this.photos[this.lightboxIndex]) return;
                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                const photo = this.photos[this.lightboxIndex];
                photo.flipH = !photo.flipH;
                this.transformPhotoAnnotationsForFlipH(photo);
                if (this.lightboxImg) {
                    this.lightboxImg.style.transform = this.getLightboxImgTransform(photo);
                }
                this.updateLightboxFlipButtonState(photo);
                this.renderLightboxAnnotations();
                this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                this.lightboxNeedsRender = true;
                this.recordHistory('燈箱水平翻轉照片');
            }

            updateLightboxFlipButtonState(photo) {
                if (!this.btnLightboxFlipH) return;
                const flipped = Boolean(photo?.flipH);
                this.btnLightboxFlipH.classList.toggle('is-flipped', flipped);
                this.btnLightboxFlipH.setAttribute('aria-pressed', String(flipped));
            }

            transformNormRectForFlipH(rect) {
                if (!rect) return rect;
                const w = Number(rect.w) || 0.1;
                const x = Math.max(0, Math.min(1 - w, 1 - (Number(rect.x) || 0) - w));
                return { ...rect, x: Math.round(x * 10000) / 10000 };
            }

            transformNormRectForRotation(rect, normDelta) {
                if (!rect || !normDelta) return rect;
                let x = Number(rect.x) || 0;
                let y = Number(rect.y) || 0;
                let w = Number(rect.w) || 0.1;
                let h = Number(rect.h) || 0.1;
                if (normDelta === 90) {
                    const nx = 1 - (y + h);
                    const ny = x;
                    x = nx; y = ny;
                    const tmp = w; w = h; h = tmp;
                } else if (normDelta === 270) {
                    const nx = y;
                    const ny = 1 - (x + w);
                    x = nx; y = ny;
                    const tmp = w; w = h; h = tmp;
                } else if (normDelta === 180) {
                    x = 1 - (x + w);
                    y = 1 - (y + h);
                }
                return {
                    ...rect,
                    x: Math.round(Math.max(0, Math.min(1 - w, x)) * 10000) / 10000,
                    y: Math.round(Math.max(0, Math.min(1 - h, y)) * 10000) / 10000,
                    w: Math.round(w * 10000) / 10000,
                    h: Math.round(h * 10000) / 10000
                };
            }

            transformPhotoAnnotationsForFlipH(photo) {
                if (!photo) return;
                if (Array.isArray(photo.annotations) && photo.annotations.length) {
                    photo.annotations = photo.annotations.map(ann => {
                        if (ann.type === 'rect') return this.transformNormRectForFlipH(ann);
                        if (ann.type === 'label') {
                            const x = Math.max(0.03, Math.min(0.97, 1 - (Number(ann.x) || 0.5)));
                            return { ...ann, x: Math.round(x * 10000) / 10000 };
                        }
                        return { ...ann };
                    });
                }
                if (photo.sideCrop && photo.sideCrop.cropRect) {
                    photo.sideCrop.cropRect = this.transformNormRectForFlipH(photo.sideCrop.cropRect);
                }
            }

            transformPhotoAnnotationsForRotation(photo, deltaDegrees) {
                if (!photo) return;
                const normDelta = ((deltaDegrees % 360) + 360) % 360;
                if (normDelta === 0) return;
                if (Array.isArray(photo.annotations) && photo.annotations.length) {
                    photo.annotations = photo.annotations.map(ann => {
                        if (ann.type === 'rect') return this.transformNormRectForRotation(ann, normDelta);
                        if (ann.type === 'label') {
                            let x = Number(ann.x) || 0.5;
                            let y = Number(ann.y) || 0.5;
                            if (normDelta === 90) {
                                const nx = 1 - y;
                                const ny = x;
                                x = nx; y = ny;
                            } else if (normDelta === 270) {
                                const nx = y;
                                const ny = 1 - x;
                                x = nx; y = ny;
                            } else if (normDelta === 180) {
                                x = 1 - x;
                                y = 1 - y;
                            }
                            return {
                                ...ann,
                                x: Math.round(Math.max(0.03, Math.min(0.97, x)) * 10000) / 10000,
                                y: Math.round(Math.max(0.03, Math.min(0.97, y)) * 10000) / 10000
                            };
                        }
                        return { ...ann };
                    });
                }
                if (photo.sideCrop && photo.sideCrop.cropRect) {
                    photo.sideCrop.cropRect = this.transformNormRectForRotation(photo.sideCrop.cropRect, normDelta);
                }
            }

            getLayoutCellDimensions(layout = this.layoutSelect?.value) {
                const map = {
                    up_down_2: { layout: 'up_down_2', label: '上下兩張', cellW: 535, cellH: 365 },
                    left_right_2: { layout: 'left_right_2', label: '左右兩張', cellW: 315, cellH: 690 },
                    landscape_3: { layout: 'landscape_3', label: '橫式三張', cellW: 335, cellH: 535 }
                };
                return map[layout] || map.up_down_2;
            }

            setLightboxLayout(layout) {
                if (!this.layoutSelect) return;
                const cell = this.getLayoutCellDimensions(layout);
                if (this.layoutSelect.value === cell.layout) {
                    this.syncLightboxAnnotationLayerBounds();
                    return;
                }
                this.layoutSelect.value = cell.layout;
                this.updateLayoutDescription();
                this.updatePageEstimate();
                if (this.currentView === 'preview') this.renderPagePreview();
                this.syncLightboxAnnotationLayerBounds();
                this.scheduleHistory('變更清冊版型');
            }

            syncLightboxSideLayoutUi({ visible = false, activeLayout = this.layoutSelect?.value || 'up_down_2', statusLevel = 'ok', statusText = '' } = {}) {
                if (this.lightboxSideLayoutBar) {
                    const wasHidden = this.lightboxSideLayoutBar.classList.contains('hidden');
                    this.lightboxSideLayoutBar.classList.toggle('hidden', !visible);
                    if (wasHidden === visible && this.isLightboxOpen()) {
                        requestAnimationFrame(() => {
                            if (this.isLightboxOpen()) this.syncLightboxAnnotationLayerBounds();
                        });
                    }
                }
                if (Array.isArray(this.lightboxSideLayoutBtns)) {
                    this.lightboxSideLayoutBtns.forEach(btn => {
                        const isActive = btn.getAttribute('data-lightbox-layout') === activeLayout;
                        btn.classList.toggle('active', isActive);
                        btn.setAttribute('aria-pressed', String(isActive));
                    });
                }
                if (this.lightboxSideLayoutStatus) {
                    this.lightboxSideLayoutStatus.classList.remove(
                        'lightbox-side-layout-status--ok',
                        'lightbox-side-layout-status--fit',
                        'lightbox-side-layout-status--warn'
                    );
                    this.lightboxSideLayoutStatus.classList.add('lightbox-side-layout-status--' + statusLevel);
                    this.lightboxSideLayoutStatus.textContent = statusText;
                }
            }

            computeSideCropGeometry(origW, origH, sideCrop) {
                if (!sideCrop || !sideCrop.cropRect || origW <= 0 || origH <= 0) return null;
                const cw = Math.max(0.02, Math.min(1, Number(sideCrop.cropRect.w) || 0.2));
                const ch = Math.max(0.02, Math.min(1, Number(sideCrop.cropRect.h) || 0.2));
                const cropPixelW = Math.max(2, cw * origW);
                const cropPixelH = Math.max(2, ch * origH);
                const cropAspect = Math.max(0.1, Math.min(10, cropPixelW / cropPixelH));
                const baseAspect = origW / origH;

                const maxWByWidth = 2.5;
                const maxWByHeight = cropAspect / baseAspect;
                const maxWidthRatio = Math.max(0.05, Math.min(maxWByWidth, maxWByHeight));
                const minWidthRatio = Math.min(
                    maxWidthRatio * 0.9,
                    Math.max(0.04, Math.min(0.22, maxWidthRatio * 0.3), 0.22 * maxWByHeight)
                );

                const rawRatio = Number(sideCrop.widthRatio);
                const preferredDefault = baseAspect < 0.85
                    ? Math.min(1.45, maxWidthRatio * 0.86)
                    : Math.min(0.95, maxWidthRatio * 0.82);
                const defaultRatio = Math.min(maxWidthRatio, Math.max(minWidthRatio, preferredDefault));
                const widthRatio = Number.isFinite(rawRatio) && rawRatio > 0
                    ? Math.max(minWidthRatio, Math.min(maxWidthRatio, rawRatio))
                    : defaultRatio;

                const panelW = origW * widthRatio;
                const panelH = Math.min(origH, panelW / cropAspect);
                const gap = Math.min(origW, origH) * 0.028;
                const offsetY = Math.max(0, Math.min(1, Number.isFinite(Number(sideCrop.offsetY)) ? Number(sideCrop.offsetY) : 0.5));
                const panelTop = offsetY * Math.max(0, origH - panelH);

                return {
                    origW,
                    origH,
                    gap,
                    panelW,
                    panelH,
                    panelTop,
                    totalW: origW + gap + panelW,
                    totalH: origH,
                    widthRatio,
                    minWidthRatio,
                    maxWidthRatio,
                    offsetY,
                    cropAspect,
                    borderColor: sideCrop.borderColor === 'red' ? 'red' : sideCrop.borderColor === 'none' ? 'none' : 'yellow'
                };
            }

            getLightboxPhotoCacheKey(photo) {
                return photo ? `${photo.uid}|${photo.rotation || 0}|${Boolean(photo.flipH)}|${photo.exifRotation || 0}` : '';
            }

            async ensureLightboxBaseCanvas(photo) {
                const key = this.getLightboxPhotoCacheKey(photo);
                if (!key) return null;
                if (this.lightboxBaseCanvasCache?.key === key) {
                    return this.lightboxBaseCanvasCache.canvas;
                }
                if (this.lightboxBaseCanvasPending?.key === key) {
                    return this.lightboxBaseCanvasPending.promise;
                }
                const promise = this.renderOrientedCanvas(photo, 2400, { includeAnnotations: false }).then(canvas => {
                    this.lightboxBaseCanvasCache = { key, canvas };
                    return canvas;
                }).finally(() => {
                    if (this.lightboxBaseCanvasPending?.key === key) {
                        this.lightboxBaseCanvasPending = null;
                    }
                });
                this.lightboxBaseCanvasPending = { key, promise };
                return promise;
            }

            paintSideCropCanvas(photo, cardDisplayW, cardDisplayH) {
                if (!photo || !photo.sideCrop || !this.lightboxSideCropCanvas) return;
                const canvas = this.lightboxSideCropCanvas;
                const key = this.getLightboxPhotoCacheKey(photo);
                const drawFromBase = (baseCanvas) => {
                    if (!baseCanvas || this.photos[this.lightboxIndex] !== photo || !photo.sideCrop) return;
                    const dpr = Math.min(2, Math.max(1, (window.devicePixelRatio || 1) * (this.lightboxZoom > 1 ? Math.min(2, this.lightboxZoom) : 1)));
                    const targetW = Math.max(2, Math.round(Math.max(cardDisplayW || 200, 120) * dpr));
                    const targetH = Math.max(2, Math.round(Math.max(cardDisplayH || 200, 120) * dpr));
                    if (canvas.width !== targetW) canvas.width = targetW;
                    if (canvas.height !== targetH) canvas.height = targetH;
                    const ctx = canvas.getContext('2d');
                    if (!ctx) return;
                    const cr = photo.sideCrop.cropRect;
                    const bw = baseCanvas.width;
                    const bh = baseCanvas.height;
                    const sx = Math.max(0, Math.min(bw - 2, Math.round((Number(cr.x) || 0) * bw)));
                    const sy = Math.max(0, Math.min(bh - 2, Math.round((Number(cr.y) || 0) * bh)));
                    const sw = Math.max(2, Math.min(bw - sx, Math.round((Number(cr.w) || 0.2) * bw)));
                    const sh = Math.max(2, Math.min(bh - sy, Math.round((Number(cr.h) || 0.2) * bh)));
                    ctx.fillStyle = '#FFFFFF';
                    ctx.fillRect(0, 0, targetW, targetH);
                    ctx.drawImage(baseCanvas, sx, sy, sw, sh, 0, 0, targetW, targetH);
                    const cw = Math.max(0.02, Math.min(1, Number(cr.w) || 0.2));
                    const ch = Math.max(0.02, Math.min(1, Number(cr.h) || 0.2));
                    const nonSourceAnns = (photo.annotations || []).filter(a => a.id !== photo.sideCrop?.sourceAnnId);
                    if (nonSourceAnns.length > 0) {
                        ctx.save();
                        ctx.beginPath();
                        ctx.rect(0, 0, targetW, targetH);
                        ctx.clip();
                        const virtW = targetW / cw;
                        const virtH = targetH / ch;
                        ctx.translate(-(Number(cr.x) || 0) * virtW, -(Number(cr.y) || 0) * virtH);
                        this.drawAnnotationsOnCanvas(ctx, virtW, virtH, nonSourceAnns);
                        ctx.restore();
                    }
                    canvas.dataset.painted = '1';
                };
                if (!this.lightboxBaseCanvasCache || this.lightboxBaseCanvasCache.key !== key) {
                    delete canvas.dataset.painted;
                }

                if (this.lightboxBaseCanvasCache?.key === key) {
                    drawFromBase(this.lightboxBaseCanvasCache.canvas);
                } else {
                    this.ensureLightboxBaseCanvas(photo)
                        .then(baseCanvas => drawFromBase(baseCanvas))
                        .catch(() => {});
                }
            }

            syncLightboxAnnotationLayerBounds() {
                if (!this.lightboxAnnotationLayer || !this.lightboxImg) return;
                const photo = this.photos[this.lightboxIndex];
                const rot = (((photo?.rotation || 0) % 360) + 360) % 360;
                const isRotated90 = rot === 90 || rot === 270;
                const nw = this.lightboxImg.naturalWidth || 0;
                const nh = this.lightboxImg.naturalHeight || 0;
                const rawVpW = this.lightboxViewport?.clientWidth || 0;
                const sideNavPadding = rawVpW > 560 ? 116 : 24;
                const vpW = Math.max(0, rawVpW - sideNavPadding);
                const vpH = Math.max(0, (this.lightboxViewport?.clientHeight || 0) - 24);

                if (photo && photo.sideCrop && nw > 0 && nh > 0 && vpW > 40 && vpH > 40) {
                    const baseW = isRotated90 ? nh : nw;
                    const baseH = isRotated90 ? nw : nh;
                    const geo = this.computeSideCropGeometry(baseW, baseH, photo.sideCrop);
                    const cell = this.getLayoutCellDimensions();
                    const framePad = 8;
                    const frameScale = Math.min(vpW / cell.cellW, vpH / cell.cellH);
                    const frameW = Math.max(120, Math.round(cell.cellW * frameScale));
                    const frameH = Math.max(120, Math.round(cell.cellH * frameScale));
                    const innerW = Math.max(40, frameW - framePad * 2);
                    const innerH = Math.max(40, frameH - framePad * 2);

                    const fullHGapEst = Math.max(2, Math.round(geo.gap * (innerH / geo.origH)));
                    const estScale = Math.min(Math.max(16, innerW - fullHGapEst) / (geo.origW + geo.panelW), innerH / geo.origH);
                    const gapPx = Math.max(2, Math.round(geo.gap * estScale));
                    const availImgW = Math.max(16, innerW - gapPx);
                    const fitScale = Math.min(availImgW / (geo.origW + geo.panelW), innerH / geo.origH);

                    const leftW = Math.max(4, Math.round(geo.origW * fitScale));
                    const leftH = Math.max(4, Math.round(geo.origH * fitScale));
                    const rawDomW = isRotated90 ? leftH : leftW;
                    const rawDomH = isRotated90 ? leftW : leftH;
                    const rightW = Math.max(4, Math.min(availImgW - leftW, Math.round(geo.panelW * fitScale)));
                    const rightH = Math.max(4, Math.min(leftH, Math.round(geo.panelH * fitScale)));
                    const rightTop = Math.round(geo.offsetY * Math.max(0, leftH - rightH));

                    const fullHLeftW = innerH * (geo.origW / geo.origH);
                    const rawSafeMax = (innerW - fullHGapEst - fullHLeftW) / Math.max(1, fullHLeftW);
                    const isNarrowLayout = rawSafeMax < geo.minWidthRatio + 0.01;
                    const isOverflow = isNarrowLayout || (geo.widthRatio > rawSafeMax + 0.012);
                    const atHeightCap = !isNarrowLayout && rawSafeMax >= geo.maxWidthRatio - 0.01 && geo.widthRatio >= geo.maxWidthRatio - 0.015;
                    const isSnappedRight = !isNarrowLayout && !isOverflow && Math.abs(geo.widthRatio - rawSafeMax) <= 0.025;

                    let statusLevel = 'ok';
                    let statusText = `「${cell.label}」右緣尚有空間（左圖維持滿高）`;
                    if (isNarrowLayout) {
                        statusLevel = 'warn';
                        statusText = cell.layout === 'up_down_2'
                            ? '左圖比例較寬，並排時整組等比縮小'
                            : `「${cell.label}」單欄較窄，並排會縮小圖片（建議改用 -1 特寫）`;
                    } else if (isOverflow) {
                        statusLevel = 'warn';
                        statusText = `超出「${cell.label}」右緣，整組等比縮小（建議縮回或改用 -1）`;
                    } else if (isSnappedRight) {
                        statusLevel = 'fit';
                        statusText = `已貼齊「${cell.label}」右緣（左圖滿高・最適尺寸）`;
                    } else if (atHeightCap) {
                        statusLevel = 'fit';
                        statusText = '右圖已達等高上限（左圖維持滿高）';
                    }

                    if (this.lightboxCanvasFrame) {
                        this.lightboxCanvasFrame.style.width = `${frameW}px`;
                        this.lightboxCanvasFrame.style.height = `${frameH}px`;
                        this.lightboxCanvasFrame.classList.add('has-side-crop');
                        this.lightboxCanvasFrame.classList.toggle('is-side-overflow', isOverflow);
                    }
                    this.syncLightboxSideLayoutUi({ visible: true, activeLayout: cell.layout, statusLevel, statusText });

                    this.lightboxImg.style.width = `${rawDomW}px`;
                    this.lightboxImg.style.height = `${rawDomH}px`;
                    this.lightboxImg.style.maxWidth = 'none';
                    this.lightboxImg.style.maxHeight = 'none';

                    if (this.lightboxLeftSlot) {
                        this.lightboxLeftSlot.style.width = `${leftW}px`;
                        this.lightboxLeftSlot.style.height = `${leftH}px`;
                    }
                    this.lightboxAnnotationLayer.style.width = `${leftW}px`;
                    this.lightboxAnnotationLayer.style.height = `${leftH}px`;

                    if (this.lightboxSideGap) {
                        this.lightboxSideGap.classList.remove('hidden');
                        this.lightboxSideGap.style.width = `${gapPx}px`;
                        this.lightboxSideGap.style.height = `${leftH}px`;
                    }
                    if (this.lightboxSideCropArea) {
                        this.lightboxSideCropArea.classList.remove('hidden');
                        this.lightboxSideCropArea.style.width = `${rightW}px`;
                        this.lightboxSideCropArea.style.height = `${leftH}px`;
                    }
                    if (this.lightboxSideCropCard) {
                        this.lightboxSideCropCard.style.height = `${rightH}px`;
                        this.lightboxSideCropCard.style.top = `${rightTop}px`;
                        this.lightboxSideCropCard.classList.remove('lightbox-side-crop-card--yellow', 'lightbox-side-crop-card--red', 'lightbox-side-crop-card--none');
                        this.lightboxSideCropCard.classList.add(`lightbox-side-crop-card--${geo.borderColor}`);
                        this.lightboxSideCropCard.classList.toggle('is-full-height', (leftH - rightH) <= 2);
                        const frameBorderW = Math.max(2, Math.round(4.5 * (Math.min(leftW + rightW, leftH) / 800)));
                        this.lightboxSideCropCard.style.setProperty('--side-crop-border-w', `${frameBorderW}px`);
                    }
                    if (this.btnSideCropBorder) {
                        this.btnSideCropBorder.textContent = geo.borderColor === 'red' ? '紅框' : geo.borderColor === 'none' ? '無框' : '黃框';
                    }
                    this.paintSideCropCanvas(photo, rightW, rightH);

                    const minSide = Math.max(120, Math.min(leftW, leftH));
                    const scale = minSide / 800;
                    const strokeW = Math.max(2.5, Math.round(5 * scale * 10) / 10);
                    const haloW = Math.max(1.2, Math.round(1.6 * scale * 10) / 10);
                    const fontSize = Math.max(12, Math.round(minSide * 0.036));
                    const padX = Math.max(6, Math.round(fontSize * 0.45));
                    const padY = Math.max(3, Math.round(fontSize * 0.28));
                    const borderW = Math.max(1.5, Math.round(2.5 * scale * 10) / 10);
                    this.lightboxAnnotationLayer.style.setProperty('--ann-stroke-w', `${strokeW}px`);
                    this.lightboxAnnotationLayer.style.setProperty('--ann-halo-w', `${haloW}px`);
                    this.lightboxAnnotationLayer.style.setProperty('--ann-font-size', `${fontSize}px`);
                    this.lightboxAnnotationLayer.style.setProperty('--ann-pad-x', `${padX}px`);
                    this.lightboxAnnotationLayer.style.setProperty('--ann-pad-y', `${padY}px`);
                    this.lightboxAnnotationLayer.style.setProperty('--ann-border-w', `${borderW}px`);
                    return;
                }

                this.syncLightboxSideLayoutUi({ visible: false });
                this.lightboxCanvasFrame?.classList.remove('has-side-crop', 'is-side-overflow');
                if (this.lightboxCanvasFrame) {
                    this.lightboxCanvasFrame.style.width = '';
                    this.lightboxCanvasFrame.style.height = '';
                }
                this.lightboxSideGap?.classList.add('hidden');
                this.lightboxSideCropArea?.classList.add('hidden');
                if (this.lightboxSideCropCanvas) delete this.lightboxSideCropCanvas.dataset.painted;
                if (this.lightboxLeftSlot) {
                    this.lightboxLeftSlot.style.width = '';
                    this.lightboxLeftSlot.style.height = '';
                }

                if (nw > 0 && nh > 0 && vpW > 40 && vpH > 40) {
                    const maxDomW = isRotated90 ? vpH : vpW;
                    const maxDomH = isRotated90 ? vpW : vpH;
                    const fitScale = Math.min(maxDomW / nw, maxDomH / nh);
                    const minTargetScale = Math.max(nw, nh) < 360 ? (360 / Math.max(nw, nh)) : 1;
                    const finalScale = Math.min(fitScale, minTargetScale);
                    const targetW = `${Math.max(1, Math.round(nw * finalScale))}px`;
                    const targetH = `${Math.max(1, Math.round(nh * finalScale))}px`;
                    if (this.lightboxImg.style.width !== targetW) this.lightboxImg.style.width = targetW;
                    if (this.lightboxImg.style.height !== targetH) this.lightboxImg.style.height = targetH;
                    if (this.lightboxImg.style.maxWidth !== 'none') this.lightboxImg.style.maxWidth = 'none';
                    if (this.lightboxImg.style.maxHeight !== 'none') this.lightboxImg.style.maxHeight = 'none';
                }

                const rawW = this.lightboxImg.offsetWidth || this.lightboxImg.clientWidth || 0;
                const rawH = this.lightboxImg.offsetHeight || this.lightboxImg.clientHeight || 0;
                if (rawW <= 0 || rawH <= 0) return;
                const orientedW = isRotated90 ? rawH : rawW;
                const orientedH = isRotated90 ? rawW : rawH;
                this.lightboxAnnotationLayer.style.width = `${orientedW}px`;
                this.lightboxAnnotationLayer.style.height = `${orientedH}px`;

                // 同步預覽與 Canvas 匯出比例（所見即所得）
                const minSide = Math.max(120, Math.min(orientedW, orientedH));
                const scale = minSide / 800;
                const strokeW = Math.max(2.5, Math.round(5 * scale * 10) / 10);
                const haloW = Math.max(1.2, Math.round(1.6 * scale * 10) / 10);
                const fontSize = Math.max(12, Math.round(minSide * 0.036));
                const padX = Math.max(6, Math.round(fontSize * 0.45));
                const padY = Math.max(3, Math.round(fontSize * 0.28));
                const borderW = Math.max(1.5, Math.round(2.5 * scale * 10) / 10);
                this.lightboxAnnotationLayer.style.setProperty('--ann-stroke-w', `${strokeW}px`);
                this.lightboxAnnotationLayer.style.setProperty('--ann-halo-w', `${haloW}px`);
                this.lightboxAnnotationLayer.style.setProperty('--ann-font-size', `${fontSize}px`);
                this.lightboxAnnotationLayer.style.setProperty('--ann-pad-x', `${padX}px`);
                this.lightboxAnnotationLayer.style.setProperty('--ann-pad-y', `${padY}px`);
                this.lightboxAnnotationLayer.style.setProperty('--ann-border-w', `${borderW}px`);
            }

            renderLightboxAnnotations() {
                if (!this.lightboxAnnotationLayer) return;
                const photo = this.photos[this.lightboxIndex];
                if (!photo) {
                    this.lightboxAnnotationLayer.innerHTML = '';
                    return;
                }
                if (!Array.isArray(photo.annotations)) photo.annotations = [];
                this.syncLightboxAnnotationLayerBounds();
                const annotations = photo.annotations;

                if (this.lightboxAnnotationCountBadge) {
                    if (annotations.length > 0 && photo.sideCrop) {
                        this.lightboxAnnotationCountBadge.textContent = `已標記 ${annotations.length} 項＋右側並排`;
                    } else if (photo.sideCrop) {
                        this.lightboxAnnotationCountBadge.textContent = '右側並排中';
                    } else {
                        this.lightboxAnnotationCountBadge.textContent = annotations.length > 0 ? `已標記 ${annotations.length} 項` : '無標記';
                    }
                }
                if (this.btnLightboxClearAnnotations) {
                    this.btnLightboxClearAnnotations.classList.toggle('hidden', annotations.length === 0);
                }

                const annHtml = annotations.map(ann => {
                    const colorClass = ann.color === 'red' ? 'red' : 'yellow';
                    if (ann.type === 'rect') {
                        const left = (Number(ann.x) || 0) * 100;
                        const top = (Number(ann.y) || 0) * 100;
                        const width = (Number(ann.w) || 0.1) * 100;
                        const height = (Number(ann.h) || 0.1) * 100;
                        return `
                            <div class="lightbox-annotation-rect lightbox-annotation-rect--${colorClass}"
                                 data-ann-id="${escapeHtml(ann.id)}"
                                 style="left:${left}%;top:${top}%;width:${width}%;height:${height}%;"
                                 title="拖曳可移動標記框，右下角可調整大小">
                                <button type="button" class="lightbox-annotation-sidecrop-btn" data-ann-sidecrop="${escapeHtml(ann.id)}" title="放到原圖右側同格並排（模式 B）"><i class="fa-solid fa-table-columns pointer-events-none"></i></button>
                                <button type="button" class="lightbox-annotation-crop-btn" data-ann-crop="${escapeHtml(ann.id)}" title="依此標記框裁切局部特寫（插在此張後方）"><i class="fa-solid fa-crop-simple pointer-events-none"></i></button>
                                <button type="button" class="lightbox-annotation-delete" data-ann-delete="${escapeHtml(ann.id)}" title="刪除此標記框">×</button>
                                <span class="lightbox-annotation-resize" data-ann-resize="${escapeHtml(ann.id)}" title="拖曳調整大小"></span>
                            </div>`;
                    }
                    if (ann.type === 'label') {
                        const left = (Number(ann.x) || 0.5) * 100;
                        const top = (Number(ann.y) || 0.5) * 100;
                        const labelScale = Math.max(0.5, Math.min(3.0, Number(ann.scale) || 1));
                        return `
                            <div class="lightbox-annotation-label lightbox-annotation-label--${colorClass}"
                                 data-ann-id="${escapeHtml(ann.id)}"
                                 style="left:${left}%;top:${top}%;--label-scale:${labelScale};"
                                 title="拖曳可移動位置，拖曳右下角圓點可放大縮小，雙擊可修改文字">
                                <span class="lightbox-annotation-label-text">${escapeHtml(ann.text || '')}</span>
                                <button type="button" class="lightbox-annotation-delete" data-ann-delete="${escapeHtml(ann.id)}" title="刪除此標籤">×</button>
                                <span class="lightbox-annotation-resize" data-ann-resize="${escapeHtml(ann.id)}" title="拖曳角落放大縮小標籤"></span>
                            </div>`;
                    }
                    return '';
                }).join('');

                let cropHtml = '';
                if (this.lightboxCropDraft) {
                    const c = this.lightboxCropDraft;
                    const left = (Number(c.x) || 0) * 100;
                    const top = (Number(c.y) || 0) * 100;
                    const width = (Number(c.w) || 0.1) * 100;
                    const height = (Number(c.h) || 0.1) * 100;
                    const showActionsInside = (Number(c.y) + Number(c.h)) > 0.84;
                    const showActionBar = !this.lightboxAnnotationAction || this.lightboxAnnotationAction.mode !== 'draw-crop';
                    const actionsHtml = showActionBar ? `
                        <div class="lightbox-crop-actions ${showActionsInside ? 'lightbox-crop-actions--inside' : ''}">
                            <button type="button" id="btnConfirmLightboxCrop" data-crop-confirm="1" title="建立局部特寫照片並鎖定於此張後方 (Enter)">✓ 插入為局部特寫</button>
                            <button type="button" id="btnSideLightboxCrop" data-crop-side="1" title="把框選局部放到原圖右側同格並排，並自動保留黃框（模式 B）">◫ 放原圖右側</button>
                            <button type="button" id="btnCancelLightboxCrop" data-crop-cancel="1" title="取消選取 (ESC)">取消</button>
                        </div>` : '';
                    cropHtml = `
                        <div class="lightbox-crop-box"
                             data-crop-move="1"
                             style="left:${left}%;top:${top}%;width:${width}%;height:${height}%;"
                             title="拖曳調整裁切位置與大小，確認後點「插入為局部特寫」">
                            ${actionsHtml}
                            <span class="lightbox-annotation-resize" data-crop-resize="1" title="拖曳調整裁切範圍大小"></span>
                        </div>`;
                }

                this.lightboxAnnotationLayer.innerHTML = annHtml + cropHtml;
            }

            updateCropBoxDom() {
                if (!this.lightboxAnnotationLayer || !this.lightboxCropDraft) return;
                const el = this.lightboxAnnotationLayer.querySelector('.lightbox-crop-box');
                if (!el) {
                    this.renderLightboxAnnotations();
                    return;
                }
                const c = this.lightboxCropDraft;
                el.style.left = `${(Number(c.x) || 0) * 100}%`;
                el.style.top = `${(Number(c.y) || 0) * 100}%`;
                el.style.width = `${(Number(c.w) || 0.1) * 100}%`;
                el.style.height = `${(Number(c.h) || 0.1) * 100}%`;
            }

            updateAnnotationElementDom(ann) {
                if (!this.lightboxAnnotationLayer || !ann) return;
                const el = this.lightboxAnnotationLayer.querySelector(`[data-ann-id="${ann.id}"]`);
                if (!el) {
                    this.renderLightboxAnnotations();
                    return;
                }
                if (ann.type === 'rect') {
                    el.style.left = `${(Number(ann.x) || 0) * 100}%`;
                    el.style.top = `${(Number(ann.y) || 0) * 100}%`;
                    el.style.width = `${(Number(ann.w) || 0.1) * 100}%`;
                    el.style.height = `${(Number(ann.h) || 0.1) * 100}%`;
                } else if (ann.type === 'label') {
                    el.style.left = `${(Number(ann.x) || 0.5) * 100}%`;
                    el.style.top = `${(Number(ann.y) || 0.5) * 100}%`;
                    el.style.setProperty('--label-scale', String(Math.max(0.5, Math.min(3.0, Number(ann.scale) || 1))));
                }
                const photo = this.photos[this.lightboxIndex];
                if (photo && photo.sideCrop && ann.id !== photo.sideCrop.sourceAnnId) {
                    const cardRect = this.lightboxSideCropCard?.getBoundingClientRect();
                    this.paintSideCropCanvas(photo, cardRect?.width || 240, cardRect?.height || 240);
                }
            }

            getAnnotationNormalizedPoint(clientX, clientY, syncBounds = false) {
                if (!this.lightboxAnnotationLayer) return { x: 0.5, y: 0.5 };
                if (syncBounds || !this.lightboxAnnotationLayer.style.width) {
                    this.syncLightboxAnnotationLayerBounds();
                }
                const rect = this.lightboxAnnotationLayer.getBoundingClientRect();
                if (!rect.width || !rect.height) return { x: 0.5, y: 0.5 };
                const x = Math.max(0, Math.min(1, (clientX - rect.left) / rect.width));
                const y = Math.max(0, Math.min(1, (clientY - rect.top) / rect.height));
                return { x, y };
            }

            deleteLightboxAnnotation(annId) {
                const photo = this.photos[this.lightboxIndex];
                if (!photo || !Array.isArray(photo.annotations)) return;
                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                const before = photo.annotations.length;
                photo.annotations = photo.annotations.filter(ann => ann.id !== annId);
                if (photo.sideCrop && photo.sideCrop.sourceAnnId === annId) {
                    photo.sideCrop.sourceAnnId = null;
                }
                if (photo.annotations.length !== before) {
                    this.renderLightboxAnnotations();
                    this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                    this.lightboxNeedsRender = true;
                    this.recordHistory('刪除照片標記');
                }
            }

            clearCurrentPhotoAnnotations() {
                const photo = this.photos[this.lightboxIndex];
                if (!photo || !Array.isArray(photo.annotations) || !photo.annotations.length) return;
                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                photo.annotations = [];
                if (photo.sideCrop) {
                    photo.sideCrop.sourceAnnId = null;
                }
                this.renderLightboxAnnotations();
                this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                this.lightboxNeedsRender = true;
                this.recordHistory('清除照片標記');
            }

            openInlineAnnotationLabelInput({ x, y, initialText = '', color = 'yellow', annId = null }) {
                if (!this.lightboxAnnotationLayer) return;
                this.lightboxAnnotationLayer.querySelector('.lightbox-inline-label-editor')?.remove();
                const clampedX = Math.max(0.1, Math.min(0.9, Number(x) || 0.5));
                const clampedY = Math.max(0.08, Math.min(0.92, Number(y) || 0.5));
                const editorEl = document.createElement('div');
                editorEl.className = 'lightbox-inline-label-editor';
                editorEl.style.left = `${clampedX * 100}%`;
                editorEl.style.top = `${clampedY * 100}%`;
                editorEl.innerHTML = `
                    <input type="text" class="lightbox-inline-label-input" placeholder="輸入姓名或文字" value="${escapeHtml(initialText)}">
                    <button type="button" class="lightbox-inline-label-ok">確定</button>
                `;
                ['pointerdown', 'mousedown', 'click', 'dblclick'].forEach(evtName => {
                    editorEl.addEventListener(evtName, (evt) => evt.stopPropagation());
                });
                this.lightboxAnnotationLayer.appendChild(editorEl);
                const inputEl = editorEl.querySelector('input');
                const okBtn = editorEl.querySelector('button');
                let finished = false;

                const commit = ({ fromBlur = false } = {}) => {
                    if (finished) return;
                    finished = true;
                    const trimmed = (inputEl?.value || '').trim();
                    editorEl.remove();
                    const photo = this.photos[this.lightboxIndex];
                    if (!photo) return;
                    if (!Array.isArray(photo.annotations)) photo.annotations = [];
                    if (annId) {
                        const ann = photo.annotations.find(item => item.id === annId);
                        if (!ann || ann.type !== 'label') return;
                        if (!trimmed) {
                            if (!fromBlur) this.deleteLightboxAnnotation(annId);
                            return;
                        }
                        if (ann.text !== trimmed) {
                            this.flushPendingHistory();
                            this.commitLightboxHistoryIfNeeded();
                            ann.text = trimmed;
                            this.renderLightboxAnnotations();
                            this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                            this.lightboxNeedsRender = true;
                            this.recordHistory('修改照片姓名標籤');
                        }
                        return;
                    }
                    if (!trimmed) return;
                    if (this.lightboxLabelTextInput) this.lightboxLabelTextInput.value = trimmed;
                    this.flushPendingHistory();
                    this.commitLightboxHistoryIfNeeded();
                    photo.annotations.push({
                        id: 'ann_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7),
                        type: 'label',
                        color: color === 'red' ? 'red' : 'yellow',
                        x: Math.round(Math.max(0.03, Math.min(0.97, Number(x) || 0.5)) * 10000) / 10000,
                        y: Math.round(Math.max(0.03, Math.min(0.97, Number(y) || 0.5)) * 10000) / 10000,
                        scale: 1,
                        text: trimmed
                    });
                    this.renderLightboxAnnotations();
                    this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                    this.lightboxNeedsRender = true;
                    this.recordHistory('新增照片姓名標籤');
                };

                okBtn?.addEventListener('click', (evt) => {
                    evt.preventDefault();
                    evt.stopPropagation();
                    commit();
                });
                inputEl?.addEventListener('keydown', (evt) => {
                    if (evt.key === 'Enter') {
                        evt.preventDefault();
                        evt.stopPropagation();
                        commit();
                    } else if (evt.key === 'Escape') {
                        evt.preventDefault();
                        evt.stopPropagation();
                        finished = true;
                        editorEl.remove();
                    }
                });
                inputEl?.addEventListener('blur', () => {
                    setTimeout(() => {
                        if (!finished && !editorEl.contains(document.activeElement)) {
                            commit({ fromBlur: true });
                        }
                    }, 0);
                });
                inputEl?.focus();
                inputEl?.select();
            }

            handleAnnotationLayerDblClick(e) {
                const labelEl = e.target.closest('.lightbox-annotation-label');
                if (!labelEl) return;
                e.stopPropagation();
                e.preventDefault();
                const annId = labelEl.getAttribute('data-ann-id');
                const photo = this.photos[this.lightboxIndex];
                const ann = photo?.annotations?.find(item => item.id === annId);
                if (!ann || ann.type !== 'label') return;
                this.openInlineAnnotationLabelInput({
                    x: ann.x,
                    y: ann.y,
                    initialText: ann.text || '',
                    color: ann.color || 'yellow',
                    annId: ann.id
                });
            }

            isNearAnnotationRectBorder(rectEl, clientX, clientY, thresholdPx = 10) {
                if (!rectEl) return false;
                const r = rectEl.getBoundingClientRect();
                if (!r.width || !r.height) return false;
                const minDist = Math.min(
                    Math.abs(clientX - r.left),
                    Math.abs(r.right - clientX),
                    Math.abs(clientY - r.top),
                    Math.abs(r.bottom - clientY)
                );
                return minDist <= thresholdPx;
            }

            handleAnnotationLayerClick(e) {
                const cropConfirmBtn = e.target.closest('[data-crop-confirm]');
                if (cropConfirmBtn) {
                    e.stopPropagation();
                    e.preventDefault();
                    this.confirmLightboxCrop();
                    return;
                }
                const cropCancelBtn = e.target.closest('[data-crop-cancel]');
                if (cropCancelBtn) {
                    e.stopPropagation();
                    e.preventDefault();
                    this.cancelLightboxCrop();
                    return;
                }
                const cropSideBtn = e.target.closest('[data-crop-side]');
                if (cropSideBtn) {
                    e.stopPropagation();
                    e.preventDefault();
                    if (this.lightboxCropDraft) {
                        this.applySideCropFromRect({ ...this.lightboxCropDraft }, { borderColor: 'yellow', addSourceRect: true });
                    }
                    return;
                }
                const annSideBtn = e.target.closest('[data-ann-sidecrop]');
                if (annSideBtn) {
                    e.stopPropagation();
                    e.preventDefault();
                    const annId = annSideBtn.getAttribute('data-ann-sidecrop');
                    const photo = this.photos[this.lightboxIndex];
                    const ann = photo?.annotations?.find(item => item.id === annId);
                    if (ann && ann.type === 'rect') {
                        this.applySideCropFromRect(
                            { x: ann.x, y: ann.y, w: ann.w, h: ann.h },
                            { borderColor: ann.color === 'red' ? 'red' : 'yellow', addSourceRect: false, sourceAnnId: ann.id }
                        );
                    }
                    return;
                }
                const annCropBtn = e.target.closest('[data-ann-crop]');
                if (annCropBtn) {
                    e.stopPropagation();
                    e.preventDefault();
                    const annId = annCropBtn.getAttribute('data-ann-crop');
                    const photo = this.photos[this.lightboxIndex];
                    const ann = photo?.annotations?.find(item => item.id === annId);
                    if (ann && ann.type === 'rect') {
                        this.createDetailCropPhoto(this.lightboxIndex, {
                            x: ann.x, y: ann.y, w: ann.w, h: ann.h
                        }).catch(err => this.handleProcessingError(err));
                    }
                }
            }

            applySideCropFromRect(cropRect, { borderColor = 'yellow', addSourceRect = false, sourceAnnId = null } = {}) {
                const photo = this.photos[this.lightboxIndex];
                if (!photo || !cropRect) return;
                const x = Math.round(Math.max(0, Math.min(0.98, Number(cropRect.x) || 0)) * 10000) / 10000;
                const y = Math.round(Math.max(0, Math.min(0.98, Number(cropRect.y) || 0)) * 10000) / 10000;
                const w = Math.round(Math.max(0.02, Math.min(1 - x, Number(cropRect.w) || 0.2)) * 10000) / 10000;
                const h = Math.round(Math.max(0.02, Math.min(1 - y, Number(cropRect.h) || 0.2)) * 10000) / 10000;
                const cleanRect = { x, y, w, h };

                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                if (!Array.isArray(photo.annotations)) photo.annotations = [];

                let linkedAnnId = sourceAnnId || null;
                if (addSourceRect) {
                    const prevLinkedAnn = photo.sideCrop?.sourceAnnId
                        ? photo.annotations.find(ann => ann.id === photo.sideCrop.sourceAnnId && ann.type === 'rect')
                        : null;
                    if (prevLinkedAnn) {
                        prevLinkedAnn.x = x;
                        prevLinkedAnn.y = y;
                        prevLinkedAnn.w = w;
                        prevLinkedAnn.h = h;
                        linkedAnnId = prevLinkedAnn.id;
                    } else {
                        const existingSimilar = photo.annotations.find(ann =>
                            ann.type === 'rect' &&
                            Math.abs((Number(ann.x) || 0) - x) < 0.015 &&
                            Math.abs((Number(ann.y) || 0) - y) < 0.015 &&
                            Math.abs((Number(ann.w) || 0) - w) < 0.015 &&
                            Math.abs((Number(ann.h) || 0) - h) < 0.015
                        );
                        if (existingSimilar) {
                            linkedAnnId = existingSimilar.id;
                        } else {
                            const newAnnId = 'ann_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);
                            photo.annotations.push({
                                id: newAnnId,
                                type: 'rect',
                                color: borderColor === 'red' ? 'red' : 'yellow',
                                x, y, w, h
                            });
                            linkedAnnId = newAnnId;
                        }
                    }
                }

                const rot = (((photo.rotation || 0) % 360) + 360) % 360;
                const isRot90 = rot === 90 || rot === 270;
                const nw = this.lightboxImg?.naturalWidth || 1200;
                const nh = this.lightboxImg?.naturalHeight || 900;
                const baseW = isRot90 ? nh : nw;
                const baseH = isRot90 ? nw : nh;
                const nextColor = borderColor === 'red' ? 'red' : borderColor === 'none' ? 'none' : 'yellow';
                const defaultGeo = this.computeSideCropGeometry(baseW, baseH, {
                    cropRect: cleanRect,
                    borderColor: nextColor
                });

                photo.sideCrop = {
                    cropRect: cleanRect,
                    widthRatio: defaultGeo ? defaultGeo.widthRatio : 0.65,
                    offsetY: 0.5,
                    borderColor: nextColor,
                    sourceAnnId: linkedAnnId
                };
                this.lightboxCropDraft = null;
                this.setLightboxTool('pan');
                this.renderLightboxAnnotations();
                this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                this.lightboxNeedsRender = true;
                this.recordHistory(addSourceRect ? '框選放到右側並排' : '標記框放到右側並排');
            }

            removeSideCrop() {
                const photo = this.photos[this.lightboxIndex];
                if (!photo || !photo.sideCrop) return;
                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                photo.sideCrop = null;
                photo._compositeKey = '';
                this.renderLightboxAnnotations();
                this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                this.lightboxNeedsRender = true;
                this.recordHistory('移除右側並排特寫');
            }

            cycleSideCropBorder() {
                const photo = this.photos[this.lightboxIndex];
                if (!photo || !photo.sideCrop) return;
                const order = ['yellow', 'red', 'none'];
                const cur = order.includes(photo.sideCrop.borderColor) ? photo.sideCrop.borderColor : 'yellow';
                photo.sideCrop.borderColor = order[(order.indexOf(cur) + 1) % order.length];
                this.syncLightboxAnnotationLayerBounds();
                this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                this.lightboxNeedsRender = true;
                this.recordHistory('切換右側特寫框線');
            }

            handleSideCropAreaClick(e) {
                const delBtn = e.target.closest('[data-side-crop-delete]');
                if (delBtn) {
                    e.stopPropagation();
                    e.preventDefault();
                    this.removeSideCrop();
                    return;
                }
                const borderBtn = e.target.closest('[data-side-crop-border]');
                if (borderBtn) {
                    e.stopPropagation();
                    e.preventDefault();
                    this.cycleSideCropBorder();
                }
            }

            handleSideCropAreaPointerDown(e) {
                if (e.target.closest('[data-side-crop-delete], [data-side-crop-border]')) {
                    e.stopPropagation();
                    return;
                }
                if (e.button !== undefined && e.button !== 0) return;
                const photo = this.photos[this.lightboxIndex];
                if (!photo || !photo.sideCrop) return;
                const nw = this.lightboxImg?.naturalWidth || 1200;
                const nh = this.lightboxImg?.naturalHeight || 900;
                const rot = (((photo.rotation || 0) % 360) + 360) % 360;
                const isRot90 = rot === 90 || rot === 270;
                const baseW = isRot90 ? nh : nw;
                const baseH = isRot90 ? nw : nh;
                const geo = this.computeSideCropGeometry(baseW, baseH, photo.sideCrop);
                if (!geo) return;

                const zoom = this.lightboxZoom || 1;
                const resizeHandle = e.target.closest('[data-side-crop-resize]');
                if (resizeHandle) {
                    e.stopPropagation();
                    e.preventDefault();
                    this.flushPendingHistory();
                    this.commitLightboxHistoryIfNeeded();
                    const leftSlotW = Math.max(80, (this.lightboxLeftSlot?.offsetWidth || 400) * zoom);
                    const rawVpW = this.lightboxViewport?.clientWidth || 0;
                    const sideNavPadding = rawVpW > 560 ? 116 : 24;
                    const vpW = Math.max(120, rawVpW - sideNavPadding);
                    const vpH = Math.max(120, (this.lightboxViewport?.clientHeight || 0) - 24);
                    const cell = this.getLayoutCellDimensions();
                    const frameScale = Math.min(vpW / cell.cellW, vpH / cell.cellH);
                    const innerW = Math.max(40, Math.round(cell.cellW * frameScale) - 16);
                    const innerH = Math.max(40, Math.round(cell.cellH * frameScale) - 16);
                    const fullHGapEst = Math.max(2, Math.round(geo.gap * (innerH / geo.origH)));
                    const fullHLeftW = innerH * (geo.origW / geo.origH);
                    const rawSafeMax = (innerW - fullHGapEst - fullHLeftW) / Math.max(1, fullHLeftW);
                    const snapFitRatio = (rawSafeMax >= geo.minWidthRatio + 0.02 && rawSafeMax <= geo.maxWidthRatio - 0.01)
                        ? Math.round(rawSafeMax * 10000) / 10000
                        : null;
                    this.lightboxAnnotationAction = {
                        mode: 'resize-side-crop',
                        pointerId: e.pointerId,
                        startX: e.clientX,
                        startY: e.clientY,
                        startWidthRatio: geo.widthRatio,
                        minWidthRatio: geo.minWidthRatio,
                        maxWidthRatio: geo.maxWidthRatio,
                        cropAspect: geo.cropAspect,
                        snapFitRatio,
                        approxScale: leftSlotW,
                        origSideCrop: this.cloneSideCrop(photo.sideCrop),
                        moved: false
                    };
                    return;
                }

                const cardEl = e.target.closest('[data-side-crop-move]');
                if (cardEl) {
                    const areaH = (this.lightboxSideCropArea?.clientHeight || 400) * zoom;
                    const cardH = (this.lightboxSideCropCard?.offsetHeight || 200) * zoom;
                    const movableSpanPx = Math.max(0, areaH - cardH);
                    if (movableSpanPx <= 2) {
                        return;
                    }
                    e.stopPropagation();
                    e.preventDefault();
                    this.flushPendingHistory();
                    this.commitLightboxHistoryIfNeeded();
                    this.lightboxAnnotationAction = {
                        mode: 'move-side-crop',
                        pointerId: e.pointerId,
                        startY: e.clientY,
                        startOffsetY: geo.offsetY,
                        movableSpanPx,
                        origSideCrop: this.cloneSideCrop(photo.sideCrop),
                        moved: false
                    };
                }
            }

            confirmLightboxCrop() {
                if (!this.lightboxCropDraft || this.lightboxIndex === null) return;
                const draft = { ...this.lightboxCropDraft };
                this.createDetailCropPhoto(this.lightboxIndex, draft).catch(err => this.handleProcessingError(err));
            }

            cancelLightboxCrop() {
                if (!this.lightboxCropDraft) return;
                this.lightboxCropDraft = null;
                this.renderLightboxAnnotations();
            }

            handleAnnotationLayerPointerDown(e) {
                if (e.target.closest('.lightbox-inline-label-editor')) return;
                if (e.target.closest('[data-crop-confirm], [data-crop-side], [data-crop-cancel], [data-ann-crop], [data-ann-sidecrop]')) {
                    e.stopPropagation();
                    return;
                }
                if (e.button !== undefined && e.button !== 0) return;
                const photo = this.photos[this.lightboxIndex];
                if (!photo) return;
                if (!Array.isArray(photo.annotations)) photo.annotations = [];

                const deleteBtn = e.target.closest('[data-ann-delete]');
                if (deleteBtn) {
                    e.stopPropagation();
                    e.preventDefault();
                    this.deleteLightboxAnnotation(deleteBtn.getAttribute('data-ann-delete'));
                    return;
                }

                const { x: normX, y: normY } = this.getAnnotationNormalizedPoint(e.clientX, e.clientY, true);

                const cropResizeHandle = e.target.closest('[data-crop-resize]');
                if (cropResizeHandle && this.lightboxCropDraft) {
                    e.stopPropagation();
                    e.preventDefault();
                    this.lightboxAnnotationAction = {
                        mode: 'resize-crop',
                        pointerId: e.pointerId,
                        startX: normX,
                        startY: normY,
                        origW: Number(this.lightboxCropDraft.w) || 0.2,
                        origH: Number(this.lightboxCropDraft.h) || 0.2,
                        moved: false
                    };
                    return;
                }

                const resizeHandle = e.target.closest('[data-ann-resize]');
                if (resizeHandle) {
                    const annId = resizeHandle.getAttribute('data-ann-resize');
                    const ann = photo.annotations.find(item => item.id === annId);
                    if (!ann) return;
                    if (ann.type === 'rect') {
                        this.flushPendingHistory();
                        this.commitLightboxHistoryIfNeeded();
                        e.stopPropagation();
                        e.preventDefault();
                        this.lightboxAnnotationAction = {
                            mode: 'resize-rect',
                            pointerId: e.pointerId,
                            annId,
                            startX: normX,
                            startY: normY,
                            origW: Number(ann.w) || 0.1,
                            origH: Number(ann.h) || 0.1,
                            moved: false
                        };
                        return;
                    }
                    if (ann.type === 'label') {
                        const labelEl = resizeHandle.closest('.lightbox-annotation-label');
                        const labelRect = labelEl?.getBoundingClientRect();
                        if (!labelRect) return;
                        const centerX = labelRect.left + labelRect.width / 2;
                        const centerY = labelRect.top + labelRect.height / 2;
                        const refX = e.clientX - centerX;
                        const refY = e.clientY - centerY;
                        const refLenSq = Math.max(16, refX * refX + refY * refY);
                        this.flushPendingHistory();
                        this.commitLightboxHistoryIfNeeded();
                        e.stopPropagation();
                        e.preventDefault();
                        this.lightboxAnnotationAction = {
                            mode: 'resize-label',
                            pointerId: e.pointerId,
                            annId,
                            centerX,
                            centerY,
                            refX,
                            refY,
                            refLenSq,
                            origScale: Math.max(0.5, Math.min(3.0, Number(ann.scale) || 1)),
                            moved: false
                        };
                        return;
                    }
                }

                const existingCropBoxEl = e.target.closest('.lightbox-crop-box[data-crop-move]');
                if (existingCropBoxEl && this.lightboxCropDraft) {
                    e.stopPropagation();
                    e.preventDefault();
                    this.lightboxAnnotationAction = {
                        mode: 'move-crop',
                        pointerId: e.pointerId,
                        startX: normX,
                        startY: normY,
                        origX: Number(this.lightboxCropDraft.x) || 0,
                        origY: Number(this.lightboxCropDraft.y) || 0,
                        origW: Number(this.lightboxCropDraft.w) || 0.2,
                        origH: Number(this.lightboxCropDraft.h) || 0.2,
                        moved: false
                    };
                    return;
                }

                const existingLabelEl = e.target.closest('.lightbox-annotation-label[data-ann-id]');
                const existingRectEl = e.target.closest('.lightbox-annotation-rect[data-ann-id]');
                const canMoveRect = existingRectEl && this.lightboxTool !== 'crop' && (
                    this.lightboxTool === 'pan' ||
                    this.isNearAnnotationRectBorder(existingRectEl, e.clientX, e.clientY, 10)
                );
                const moveTargetEl = (this.lightboxTool !== 'crop' ? existingLabelEl : null) || (canMoveRect ? existingRectEl : null);
                if (moveTargetEl) {
                    const annId = moveTargetEl.getAttribute('data-ann-id');
                    const ann = photo.annotations.find(item => item.id === annId);
                    if (!ann) return;
                    this.flushPendingHistory();
                    this.commitLightboxHistoryIfNeeded();
                    e.stopPropagation();
                    e.preventDefault();
                    this.lightboxAnnotationAction = {
                        mode: 'move',
                        pointerId: e.pointerId,
                        annId,
                        startX: normX,
                        startY: normY,
                        origX: Number(ann.x) || 0,
                        origY: Number(ann.y) || 0,
                        origW: Number(ann.w) || 0,
                        origH: Number(ann.h) || 0,
                        moved: false
                    };
                    return;
                }

                if (this.lightboxTool === 'pan') return;

                if (this.lightboxTool === 'crop') {
                    e.stopPropagation();
                    e.preventDefault();
                    this.lightboxCropDraft = {
                        x: Math.round(normX * 10000) / 10000,
                        y: Math.round(normY * 10000) / 10000,
                        w: 0.01,
                        h: 0.01
                    };
                    this.lightboxAnnotationAction = {
                        mode: 'draw-crop',
                        pointerId: e.pointerId,
                        startX: normX,
                        startY: normY,
                        moved: false
                    };
                    this.renderLightboxAnnotations();
                    return;
                }

                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                e.stopPropagation();
                e.preventDefault();

                if (this.lightboxTool === 'label') {
                    const color = this.lightboxLabelColorSelect?.value === 'red' ? 'red' : 'yellow';
                    const text = this.lightboxLabelTextInput?.value?.trim() || '';
                    if (!text) {
                        this.openInlineAnnotationLabelInput({
                            x: normX,
                            y: normY,
                            initialText: '',
                            color,
                            annId: null
                        });
                        return;
                    }
                    photo.annotations.push({
                        id: 'ann_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7),
                        type: 'label',
                        color,
                        x: Math.round(Math.max(0.03, Math.min(0.97, normX)) * 10000) / 10000,
                        y: Math.round(Math.max(0.03, Math.min(0.97, normY)) * 10000) / 10000,
                        scale: 1,
                        text
                    });
                    this.renderLightboxAnnotations();
                    this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                    this.lightboxNeedsRender = true;
                    this.recordHistory('新增照片姓名標籤');
                    return;
                }

                if (this.lightboxTool === 'rect-yellow' || this.lightboxTool === 'rect-red') {
                    const color = this.lightboxTool === 'rect-red' ? 'red' : 'yellow';
                    const annId = 'ann_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);
                    photo.annotations.push({
                        id: annId,
                        type: 'rect',
                        color,
                        x: Math.round(normX * 10000) / 10000,
                        y: Math.round(normY * 10000) / 10000,
                        w: 0.01,
                        h: 0.01
                    });
                    this.lightboxAnnotationAction = {
                        mode: 'draw-rect',
                        pointerId: e.pointerId,
                        annId,
                        startX: normX,
                        startY: normY,
                        moved: false
                    };
                    this.renderLightboxAnnotations();
                }
            }

            isLightboxOpen() {
                return Boolean(this.imageLightboxModal && !this.imageLightboxModal.classList.contains('hidden'));
            }

            mountEditorToLightbox() {
                if (!this.editorCard || !this.lightboxEditorMount) return;
                if (this.editorCard.parentElement !== this.lightboxEditorMount) {
                    this.lightboxEditorMount.appendChild(this.editorCard);
                }
                this.applyLightboxEditorLayout();
            }

            restoreEditorFromLightbox() {
                if (!this.editorCard || !this.editorPanel) return;
                if (this.editorCard.parentElement !== this.editorPanel) {
                    this.editorPanel.appendChild(this.editorCard);
                }
            }

            applyLightboxEditorLayout() {
                if (!this.lightboxEditorMount) return;
                this.lightboxEditorMount.classList.toggle('hidden', this.lightboxEditorCollapsed);
                if (this.lightboxToggleEditorText) {
                    this.lightboxToggleEditorText.textContent = this.lightboxEditorCollapsed ? '顯示編輯欄' : '隱藏編輯欄';
                }
                if (this.btnLightboxToggleEditor) {
                    this.btnLightboxToggleEditor.setAttribute('aria-expanded', String(!this.lightboxEditorCollapsed));
                }
            }

            toggleLightboxEditor() {
                this.lightboxEditorCollapsed = !this.lightboxEditorCollapsed;
                this.saveWorkspacePreference('lightboxEditorCollapsed', this.lightboxEditorCollapsed);
                this.applyLightboxEditorLayout();
                requestAnimationFrame(() => this.syncLightboxAnnotationLayerBounds());
            }

            syncLightboxSelectionAndEditor(index) {
                if (!this.photos[index]) return;
                if (this.currentIndex !== index) {
                    this.currentIndex = index;
                    const selectedCount = this.photos.filter(p => p.selected).length;
                    if (selectedCount <= 1) {
                        this.photos.forEach((p, idx) => { p.selected = (idx === index); });
                        this.lastSelectedIndex = index;
                    }
                    this.lightboxNeedsRender = true;
                }
                this.updateSelectionUi();
                this.populateEditorFields();
            }

            openLightboxByIndex(index) {
                if (!this.photos[index]) return;
                const wasOpen = this.isLightboxOpen();
                if (this.isTypingTarget(document.activeElement)) {
                    document.activeElement.blur();
                }
                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                if (!wasOpen) {
                    this.setLightboxTool('pan');
                }
                this.lightboxIndex = index;
                this.syncLightboxSelectionAndEditor(index);
                this.mountEditorToLightbox();
                this.openModal('imageLightboxModal');
                this.updateLightbox();
            }

            navigateLightbox(delta) {
                if (!this.photos.length || this.lightboxIndex === null) return;
                if (this.isTypingTarget(document.activeElement)) {
                    document.activeElement.blur();
                }
                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                const newIndex = Math.max(0, Math.min(this.photos.length - 1, this.lightboxIndex + delta));
                if (newIndex === this.lightboxIndex) return;
                this.lightboxIndex = newIndex;
                this.syncLightboxSelectionAndEditor(newIndex);
                this.updateLightbox();
            }

            updateLightbox({ preserveZoomIfSamePhoto = false } = {}) {
                const photo = this.photos[this.lightboxIndex];
                if (!photo) return;
                const isSamePhoto = preserveZoomIfSamePhoto && this.lightboxPhotoUid === photo.uid;
                this.lightboxPhotoUid = photo.uid;
                if (!isSamePhoto) {
                    this.lightboxCropDraft = null;
                    this.resetLightboxZoom();
                }

                if (this.lightboxImg) {
                    this.lightboxImg.style.transform = this.getLightboxImgTransform(photo);
                    const desiredFull = photo.src || photo.thumbnailSrc;
                    if (!(isSamePhoto && this.lightboxImg.getAttribute('src') === desiredFull && this.lightboxImg.complete)) {
                        // 平滑降級：若已具備縮圖先以縮圖即時呈現，避免空白；背景非同步載入原圖無縫置換
                        this.lightboxImg.src = photo.thumbnailSrc || photo.src;
                        if (photo.src && photo.src !== photo.thumbnailSrc) {
                            const targetIndex = this.lightboxIndex;
                            const fullImg = new Image();
                            fullImg.onload = () => {
                                if (this.lightboxIndex === targetIndex && this.lightboxImg) {
                                    this.lightboxImg.src = photo.src;
                                    this.syncLightboxAnnotationLayerBounds();
                                }
                            };
                            fullImg.src = photo.src;
                        }
                    }
                }
                this.updateLightboxFlipButtonState(photo);
                this.renderLightboxAnnotations();
                requestAnimationFrame(() => this.syncLightboxAnnotationLayerBounds());

                if (this.lightboxCaption) {
                    const parentPhoto = this.getValidParentPhoto(photo);
                    this.lightboxCaption.textContent = parentPhoto
                        ? `#${photo.seq}（#${parentPhoto.seq} 局部特寫）　${photo.name}`
                        : `#${photo.seq}　${photo.name}`;
                }
                if (this.lightboxPosition) {
                    this.lightboxPosition.textContent = `${this.lightboxIndex + 1} / ${this.photos.length}`;
                }
                if (this.btnLightboxPrev) {
                    this.btnLightboxPrev.disabled = this.lightboxIndex === 0;
                }
                if (this.btnLightboxNext) {
                    this.btnLightboxNext.disabled = this.lightboxIndex === this.photos.length - 1;
                }
            }

            applyLightboxTransform() {
                if (this.lightboxStage) {
                    this.lightboxStage.style.transform = `translate(${this.lightboxPanX}px, ${this.lightboxPanY}px) scale(${this.lightboxZoom})`;
                }
                if (this.lightboxZoomBadge) {
                    this.lightboxZoomBadge.textContent = `${Math.round(this.lightboxZoom * 100)}%`;
                }
                if (this.lightboxViewport) {
                    if (this.lightboxZoom > 1.01) {
                        this.lightboxViewport.classList.add('is-zoomed');
                    } else {
                        this.lightboxViewport.classList.remove('is-zoomed');
                    }
                }
            }

            resetLightboxZoom() {
                this.lightboxZoom = 1.0;
                this.lightboxPanX = 0;
                this.lightboxPanY = 0;
                this.applyLightboxTransform();
            }

            zoomLightboxBy(step, clientX = null, clientY = null) {
                const oldZoom = this.lightboxZoom;
                let newZoom = Math.max(0.35, Math.min(4.0, Math.round((oldZoom + step) * 100) / 100));
                if (newZoom === oldZoom) return;

                if (clientX !== null && clientY !== null && this.lightboxViewport) {
                    const rect = this.lightboxViewport.getBoundingClientRect();
                    const cursorX = clientX - (rect.left + rect.width / 2);
                    const cursorY = clientY - (rect.top + rect.height / 2);
                    const ratio = newZoom / oldZoom;
                    this.lightboxPanX = cursorX - (cursorX - this.lightboxPanX) * ratio;
                    this.lightboxPanY = cursorY - (cursorY - this.lightboxPanY) * ratio;
                } else {
                    const ratio = newZoom / oldZoom;
                    this.lightboxPanX = this.lightboxPanX * ratio;
                    this.lightboxPanY = this.lightboxPanY * ratio;
                }

                this.lightboxZoom = newZoom;
                if (newZoom <= 1.0) {
                    this.lightboxPanX = 0;
                    this.lightboxPanY = 0;
                }
                this.applyLightboxTransform();
            }

            handleLightboxWheel(e) {
                e.preventDefault();
                e.stopPropagation();
                const deltaSign = Math.sign(e.deltaY);
                const step = deltaSign < 0 ? 0.15 : -0.15;
                this.zoomLightboxBy(step, e.clientX, e.clientY);
            }

            handleLightboxPointerDown(e) {
                if (this.lightboxAnnotationAction) return;
                if (this.lightboxTool !== 'pan') return;
                if (e.button !== 0 || this.lightboxZoom <= 1.0) return;
                this.lightboxIsPanning = true;
                this.lightboxPanStart = {
                    x: e.clientX - this.lightboxPanX,
                    y: e.clientY - this.lightboxPanY
                };
                if (this.lightboxViewport) {
                    this.lightboxViewport.classList.add('is-panning');
                    try {
                        this.lightboxViewport.setPointerCapture(e.pointerId);
                    } catch (_) {}
                }
                if (this.lightboxStage) {
                    this.lightboxStage.classList.add('panning-active');
                }
                e.preventDefault();
            }

            handleLightboxPointerMove(e) {
                if (this.lightboxAnnotationAction) {
                    const action = this.lightboxAnnotationAction;
                    if (action.pointerId !== undefined && e.pointerId !== undefined && action.pointerId !== e.pointerId) return;

                    if (action.mode === 'resize-side-crop' || action.mode === 'move-side-crop') {
                        const photo = this.photos[this.lightboxIndex];
                        if (!photo || !photo.sideCrop) {
                            this.lightboxAnnotationAction = null;
                            return;
                        }
                        if (action.mode === 'resize-side-crop') {
                            const dx = e.clientX - action.startX;
                            const dy = e.clientY - action.startY;
                            const projDelta = (dx * 0.96 + dy * 0.28 * Math.min(1.2, Math.max(0.4, action.cropAspect))) / action.approxScale;
                            let nextRatio = Math.max(action.minWidthRatio, Math.min(action.maxWidthRatio, action.startWidthRatio + projDelta));
                            if (action.snapFitRatio && Math.abs(nextRatio - action.snapFitRatio) < 0.025) {
                                nextRatio = action.snapFitRatio;
                            }
                            nextRatio = Math.round(nextRatio * 10000) / 10000;
                            if (Math.abs(nextRatio - photo.sideCrop.widthRatio) >= 0.002) action.moved = true;
                            photo.sideCrop.widthRatio = nextRatio;
                            this.syncLightboxAnnotationLayerBounds();
                        } else if (action.mode === 'move-side-crop') {
                            if (action.movableSpanPx > 2) {
                                const dy = e.clientY - action.startY;
                                let nextOffset = Math.max(0, Math.min(1, action.startOffsetY + (dy / action.movableSpanPx)));
                                if (nextOffset < 0.035) nextOffset = 0;
                                else if (Math.abs(nextOffset - 0.5) < 0.035) nextOffset = 0.5;
                                else if (nextOffset > 0.965) nextOffset = 1;
                                nextOffset = Math.round(nextOffset * 10000) / 10000;
                                if (Math.abs(nextOffset - photo.sideCrop.offsetY) >= 0.002) action.moved = true;
                                photo.sideCrop.offsetY = nextOffset;
                                this.syncLightboxAnnotationLayerBounds();
                            }
                        }
                        e.preventDefault();
                        return;
                    }

                    const { x: normX, y: normY } = this.getAnnotationNormalizedPoint(e.clientX, e.clientY, false);

                    if (action.mode === 'draw-crop' || action.mode === 'move-crop' || action.mode === 'resize-crop') {
                        const draft = this.lightboxCropDraft;
                        if (!draft) {
                            this.lightboxAnnotationAction = null;
                            return;
                        }
                        if (action.mode === 'draw-crop') {
                            const x1 = Math.max(0, Math.min(1, Math.min(action.startX, normX)));
                            const y1 = Math.max(0, Math.min(1, Math.min(action.startY, normY)));
                            const x2 = Math.max(0, Math.min(1, Math.max(action.startX, normX)));
                            const y2 = Math.max(0, Math.min(1, Math.max(action.startY, normY)));
                            draft.x = Math.round(x1 * 10000) / 10000;
                            draft.y = Math.round(y1 * 10000) / 10000;
                            draft.w = Math.round(Math.max(0.01, x2 - x1) * 10000) / 10000;
                            draft.h = Math.round(Math.max(0.01, y2 - y1) * 10000) / 10000;
                            if ((x2 - x1) >= 0.02 && (y2 - y1) >= 0.02) action.moved = true;
                        } else if (action.mode === 'move-crop') {
                            const dx = normX - action.startX;
                            const dy = normY - action.startY;
                            if (Math.abs(dx) > 0.002 || Math.abs(dy) > 0.002) action.moved = true;
                            draft.x = Math.round(Math.max(0, Math.min(1 - draft.w, action.origX + dx)) * 10000) / 10000;
                            draft.y = Math.round(Math.max(0, Math.min(1 - draft.h, action.origY + dy)) * 10000) / 10000;
                        } else if (action.mode === 'resize-crop') {
                            const dx = normX - action.startX;
                            const dy = normY - action.startY;
                            if (Math.abs(dx) > 0.002 || Math.abs(dy) > 0.002) action.moved = true;
                            draft.w = Math.round(Math.max(0.02, Math.min(1 - draft.x, action.origW + dx)) * 10000) / 10000;
                            draft.h = Math.round(Math.max(0.02, Math.min(1 - draft.y, action.origH + dy)) * 10000) / 10000;
                        }
                        this.updateCropBoxDom();
                        e.preventDefault();
                        return;
                    }

                    const photo = this.photos[this.lightboxIndex];
                    const ann = photo?.annotations?.find(item => item.id === action.annId);
                    if (!ann) {
                        this.lightboxAnnotationAction = null;
                        return;
                    }
                    if (action.mode === 'draw-rect') {
                        const x1 = Math.max(0, Math.min(1, Math.min(action.startX, normX)));
                        const y1 = Math.max(0, Math.min(1, Math.min(action.startY, normY)));
                        const x2 = Math.max(0, Math.min(1, Math.max(action.startX, normX)));
                        const y2 = Math.max(0, Math.min(1, Math.max(action.startY, normY)));
                        ann.x = Math.round(x1 * 10000) / 10000;
                        ann.y = Math.round(y1 * 10000) / 10000;
                        ann.w = Math.round(Math.max(0.01, x2 - x1) * 10000) / 10000;
                        ann.h = Math.round(Math.max(0.01, y2 - y1) * 10000) / 10000;
                        if ((x2 - x1) > 0.012 || (y2 - y1) > 0.012) action.moved = true;
                        this.updateAnnotationElementDom(ann);
                    } else if (action.mode === 'move') {
                        const dx = normX - action.startX;
                        const dy = normY - action.startY;
                        if (Math.abs(dx) > 0.002 || Math.abs(dy) > 0.002) action.moved = true;
                        if (ann.type === 'rect') {
                            ann.x = Math.round(Math.max(0, Math.min(1 - ann.w, action.origX + dx)) * 10000) / 10000;
                            ann.y = Math.round(Math.max(0, Math.min(1 - ann.h, action.origY + dy)) * 10000) / 10000;
                            if (photo.sideCrop && photo.sideCrop.sourceAnnId === ann.id) {
                                photo.sideCrop.cropRect = { x: ann.x, y: ann.y, w: ann.w, h: ann.h };
                                this.syncLightboxAnnotationLayerBounds();
                            }
                        } else if (ann.type === 'label') {
                            ann.x = Math.round(Math.max(0.03, Math.min(0.97, action.origX + dx)) * 10000) / 10000;
                            ann.y = Math.round(Math.max(0.03, Math.min(0.97, action.origY + dy)) * 10000) / 10000;
                        }
                        this.updateAnnotationElementDom(ann);
                    } else if (action.mode === 'resize-rect') {
                        const dx = normX - action.startX;
                        const dy = normY - action.startY;
                        if (Math.abs(dx) > 0.002 || Math.abs(dy) > 0.002) action.moved = true;
                        ann.w = Math.round(Math.max(0.02, Math.min(1 - ann.x, action.origW + dx)) * 10000) / 10000;
                        ann.h = Math.round(Math.max(0.02, Math.min(1 - ann.y, action.origH + dy)) * 10000) / 10000;
                        if (photo.sideCrop && photo.sideCrop.sourceAnnId === ann.id) {
                            photo.sideCrop.cropRect = { x: ann.x, y: ann.y, w: ann.w, h: ann.h };
                            this.syncLightboxAnnotationLayerBounds();
                        }
                        this.updateAnnotationElementDom(ann);
                    } else if (action.mode === 'resize-label') {
                        const curX = e.clientX - action.centerX;
                        const curY = e.clientY - action.centerY;
                        const ratio = ((curX * action.refX) + (curY * action.refY)) / action.refLenSq;
                        const nextScale = Math.round(Math.max(0.5, Math.min(3.0, action.origScale * ratio)) * 100) / 100;
                        if (Math.abs(nextScale - (Number(ann.scale) || 1)) >= 0.01) {
                            action.moved = true;
                        }
                        ann.scale = nextScale;
                        this.updateAnnotationElementDom(ann);
                    }
                    e.preventDefault();
                    return;
                }
                if (!this.lightboxIsPanning) return;
                this.lightboxPanX = e.clientX - this.lightboxPanStart.x;
                this.lightboxPanY = e.clientY - this.lightboxPanStart.y;

                if (this.lightboxViewport) {
                    const maxPanX = (this.lightboxViewport.clientWidth * (this.lightboxZoom - 0.4)) / 2 + 100;
                    const maxPanY = (this.lightboxViewport.clientHeight * (this.lightboxZoom - 0.4)) / 2 + 100;
                    this.lightboxPanX = Math.max(-maxPanX, Math.min(maxPanX, this.lightboxPanX));
                    this.lightboxPanY = Math.max(-maxPanY, Math.min(maxPanY, this.lightboxPanY));
                }
                this.applyLightboxTransform();
            }

            handleLightboxPointerUp(e) {
                if (this.lightboxAnnotationAction) {
                    const action = this.lightboxAnnotationAction;
                    this.lightboxAnnotationAction = null;
                    if (action.mode === 'resize-side-crop' || action.mode === 'move-side-crop') {
                        const photo = this.photos[this.lightboxIndex];
                        if (photo && photo.sideCrop && action.moved) {
                            this.syncLightboxAnnotationLayerBounds();
                            this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                            this.lightboxNeedsRender = true;
                            this.recordHistory(action.mode === 'resize-side-crop' ? '調整右側並排特寫大小' : '微調右側並排特寫位置');
                        }
                        return;
                    }
                    if (action.mode === 'draw-crop' || action.mode === 'move-crop' || action.mode === 'resize-crop') {
                        if (action.mode === 'draw-crop' && (!action.moved || !this.lightboxCropDraft || this.lightboxCropDraft.w < 0.02 || this.lightboxCropDraft.h < 0.02)) {
                            this.lightboxCropDraft = null;
                        }
                        this.renderLightboxAnnotations();
                        return;
                    }
                    const photo = this.photos[this.lightboxIndex];
                    if (photo && Array.isArray(photo.annotations)) {
                        if (action.mode === 'draw-rect' && !action.moved) {
                            photo.annotations = photo.annotations.filter(item => item.id !== action.annId);
                            this.renderLightboxAnnotations();
                            return;
                        }
                        if (action.moved) {
                            this.renderLightboxAnnotations();
                            this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                            this.lightboxNeedsRender = true;
                            this.recordHistory(action.mode === 'draw-rect' ? '新增照片標記框' : '調整照片標記');
                        }
                    }
                    return;
                }
                if (!this.lightboxIsPanning) return;
                this.lightboxIsPanning = false;
                if (this.lightboxViewport) {
                    this.lightboxViewport.classList.remove('is-panning');
                    try {
                        if (e && e.pointerId && this.lightboxViewport.hasPointerCapture(e.pointerId)) {
                            this.lightboxViewport.releasePointerCapture(e.pointerId);
                        }
                    } catch (_) {}
                }
                if (this.lightboxStage) {
                    this.lightboxStage.classList.remove('panning-active');
                }
            }

            rotateLightboxPhoto(deltaDegrees) {
                if (this.lightboxIndex === null || !this.photos[this.lightboxIndex]) return;
                this.flushPendingHistory();
                const photo = this.photos[this.lightboxIndex];
                const rawDelta = photo.flipH ? -deltaDegrees : deltaDegrees;
                this.transformPhotoAnnotationsForRotation(photo, deltaDegrees);
                photo.rotation = (((photo.rotation || 0) + rawDelta) % 360 + 360) % 360;
                if (this.lightboxImg) {
                    this.lightboxImg.style.transform = this.getLightboxImgTransform(photo);
                }
                this.renderLightboxAnnotations();
                this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                this.lightboxHasRotated = true;
            }

            getCompositeThumbSignature(photo) {
                if (!photo || !photo.sideCrop) return '';
                return JSON.stringify({
                    r: photo.rotation || 0,
                    f: Boolean(photo.flipH),
                    e: photo.exifRotation || 0,
                    sc: photo.sideCrop,
                    ann: photo.annotations || []
                });
            }

            async refreshPhotoCompositeThumbnail(photo) {
                if (!photo) return;
                if (!photo.sideCrop) {
                    if (photo.compositeThumbnailSrc) {
                        URL.revokeObjectURL(photo.compositeThumbnailSrc);
                        this.managedUrls.delete(photo.compositeThumbnailSrc);
                        photo.compositeThumbnailSrc = null;
                    }
                    photo._compositeKey = '';
                    this.applyCompositeThumbToDom(photo);
                    return;
                }
                const key = this.getCompositeThumbSignature(photo);
                if (photo._compositeKey === key && photo.compositeThumbnailSrc) {
                    this.applyCompositeThumbToDom(photo);
                    return;
                }
                try {
                    const canvas = await this.renderCompositeCanvas(photo, 320);
                    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.8));
                    if (!blob) return;
                    if (!this.photos.includes(photo) || !photo.sideCrop || this.getCompositeThumbSignature(photo) !== key) return;
                    const nextUrl = URL.createObjectURL(blob);
                    const prevUrl = photo.compositeThumbnailSrc;
                    photo.compositeThumbnailSrc = nextUrl;
                    photo._compositeKey = key;
                    this.managedUrls.add(nextUrl);
                    this.applyCompositeThumbToDom(photo);
                    if (prevUrl && prevUrl !== nextUrl) {
                        URL.revokeObjectURL(prevUrl);
                        this.managedUrls.delete(prevUrl);
                    }
                } catch (_) {}
            }

            applyCompositeThumbToDom(photo) {
                if (!photo) return;
                const idx = this.photos.indexOf(photo);
                const thumbCard = document.getElementById(`thumb_${photo.uid}`);
                const layout = this.getPhotoDisplayLayout(photo);
                if (thumbCard) {
                    const img = thumbCard.querySelector('.thumbnail-stage img');
                    if (img) {
                        if (img.getAttribute('src') !== (layout.src || photo.thumbnailSrc)) {
                            img.src = layout.src || photo.thumbnailSrc;
                        }
                        img.style.cssText = layout.imageStyle;
                    }
                }
                if (idx >= 0 && this.gridTbody && this.gridTbody.children[idx]) {
                    const rowImg = this.gridTbody.children[idx].querySelector('td:nth-child(2) img');
                    if (rowImg) {
                        rowImg.src = layout.isComposite ? layout.src : photo.thumbnailSrc;
                        rowImg.style.transform = layout.isComposite ? 'none' : `${photo.flipH ? 'scaleX(-1) ' : ''}rotate(${photo.rotation || 0}deg)`;
                    }
                }
                const previewImg = this.previewContainer?.querySelector(`img[data-preview-uid="${photo.uid}"]`);
                if (previewImg) {
                    previewImg.src = layout.isComposite ? layout.src : photo.thumbnailSrc;
                    previewImg.style.transform = layout.isComposite ? 'none' : `${photo.flipH ? 'scaleX(-1) ' : ''}rotate(${photo.rotation || 0}deg)`;
                }
            }

            scheduleSideCropPreviewRefresh(photo) {
                if (!photo) return;
                if (!photo.sideCrop) {
                    this.refreshPhotoCompositeThumbnail(photo);
                    return;
                }
                const prevTimer = this.compositeRefreshTimers.get(photo);
                if (prevTimer) clearTimeout(prevTimer);
                const timer = setTimeout(() => {
                    this.compositeRefreshTimers.delete(photo);
                    this.refreshPhotoCompositeThumbnail(photo);
                }, 40);
                this.compositeRefreshTimers.set(photo, timer);
            }

            syncAllCompositeThumbnails() {
                return Promise.all(
                    this.photos.map(p => (p && p.sideCrop) ? this.refreshPhotoCompositeThumbnail(p) : Promise.resolve())
                );
            }

            syncPhotoRotationToCanvas(index, rotation) {
                const photo = this.photos[index];
                if (!photo) return;
                // 更新主畫布卡片縮圖視覺與標記徽章
                const thumbCard = document.getElementById(`thumb_${photo.uid}`);
                if (thumbCard) {
                    const stage = thumbCard.querySelector('.thumbnail-stage');
                    const img = stage?.querySelector('img');
                    if (img) {
                        const layout = this.getPhotoDisplayLayout(photo);
                        const nextSrc = layout.src || photo.thumbnailSrc;
                        if (img.getAttribute('src') !== nextSrc) img.src = nextSrc;
                        img.style.cssText = layout.imageStyle;
                    }
                    if (stage) {
                        stage.querySelectorAll('.thumb-annotation-badge').forEach(el => el.remove());
                        const badgeHtml = this.getPhotoAnnotationBadgeHtml(photo);
                        if (badgeHtml) stage.insertAdjacentHTML('beforeend', badgeHtml);
                    }
                }
                // 更新試算表 DataGrid 縮圖
                if (this.gridTbody && this.gridTbody.children[index]) {
                    const row = this.gridTbody.children[index];
                    const rowImg = row.querySelector('td:nth-child(2) img');
                    if (rowImg) {
                        const isComp = Boolean(photo.sideCrop && photo.compositeThumbnailSrc);
                        const nextRowSrc = isComp ? photo.compositeThumbnailSrc : photo.thumbnailSrc;
                        if (rowImg.getAttribute('src') !== nextRowSrc) rowImg.src = nextRowSrc;
                        rowImg.style.transform = isComp ? 'none' : `${photo.flipH ? 'scaleX(-1) ' : ''}rotate(${rotation}deg)`;
                    }
                }
                this.scheduleSideCropPreviewRefresh(photo);
            }

            commitLightboxHistoryIfNeeded() {
                if (this.lightboxHasRotated) {
                    this.recordHistory('燈箱旋轉照片');
                    this.lightboxHasRotated = false;
                }
            }

            openModal(id) {
                if (typeof PhotoReportModalUi !== 'undefined') {
                    PhotoReportModalUi.openModal(id);
                    return;
                }
                const el = typeof id === 'string' ? document.getElementById(id) : id;
                if (el) el.classList.remove('hidden');
            }

            closeModal(id) {
                if (id === 'sequenceModal') this.pendingSequenceExport = null;
                if (id === 'exportAuditModal') this.pendingExportAction = null;
                if (id === 'imageLightboxModal') {
                    this.lightboxCropDraft = null;
                    this.cancelLightboxInteractions();
                    if (this.isTypingTarget(document.activeElement)) {
                        document.activeElement.blur();
                    }
                    this.flushPendingHistory();
                    this.restoreEditorFromLightbox();
                    this.commitLightboxHistoryIfNeeded();
                    this.lightboxBaseCanvasCache = null;
                }
                if (typeof PhotoReportModalUi !== 'undefined') {
                    PhotoReportModalUi.closeModal(id);
                } else {
                    const el = typeof id === 'string' ? document.getElementById(id) : id;
                    if (el) el.classList.add('hidden');
                }
                if (id === 'imageLightboxModal' && this.lightboxNeedsRender) {
                    this.lightboxNeedsRender = false;
                    this.render();
                }
            }

            drawAnnotationsOnCanvas(ctx, canvasWidth, canvasHeight, annotations) {
                if (!ctx || !Array.isArray(annotations) || !annotations.length) return;
                const minSide = Math.max(240, Math.min(canvasWidth, canvasHeight));
                const scale = minSide / 800;
                const strokeWidth = Math.max(3, Math.round(5 * scale));
                const haloWidth = strokeWidth + Math.max(2, Math.round(2.5 * scale));

                annotations.forEach(ann => {
                    if (!ann) return;
                    const isRed = ann.color === 'red';
                    const mainColor = isRed ? '#EF4444' : '#FFD23F';

                    if (ann.type === 'rect') {
                        const rx = Math.round((Number(ann.x) || 0) * canvasWidth);
                        const ry = Math.round((Number(ann.y) || 0) * canvasHeight);
                        const rw = Math.max(4, Math.round((Number(ann.w) || 0.1) * canvasWidth));
                        const rh = Math.max(4, Math.round((Number(ann.h) || 0.1) * canvasHeight));
                        ctx.save();
                        ctx.lineJoin = 'round';
                        ctx.strokeStyle = 'rgba(15, 23, 42, 0.88)';
                        ctx.lineWidth = haloWidth;
                        ctx.strokeRect(rx, ry, rw, rh);
                        ctx.strokeStyle = mainColor;
                        ctx.lineWidth = strokeWidth;
                        ctx.strokeRect(rx, ry, rw, rh);
                        ctx.restore();
                    } else if (ann.type === 'label' && String(ann.text || '').trim()) {
                        const text = String(ann.text).trim();
                        const labelScale = Math.max(0.5, Math.min(3.0, Number(ann.scale) || 1));
                        const fontSize = Math.max(10, Math.round(minSide * 0.036 * labelScale));
                        ctx.save();
                        ctx.font = `900 ${fontSize}px "Microsoft JhengHei", "PingFang TC", "Noto Sans TC", sans-serif`;
                        const metrics = ctx.measureText(text);
                        const padX = Math.max(4, Math.round(fontSize * 0.45));
                        const padY = Math.max(2, Math.round(fontSize * 0.28));
                        const boxW = Math.ceil(metrics.width) + padX * 2;
                        const boxH = fontSize + padY * 2;
                        let cx = (Number(ann.x) || 0.5) * canvasWidth;
                        let cy = (Number(ann.y) || 0.5) * canvasHeight;
                        cx = Math.max(boxW / 2 + 4, Math.min(canvasWidth - boxW / 2 - 4, cx));
                        cy = Math.max(boxH / 2 + 4, Math.min(canvasHeight - boxH / 2 - 4, cy));
                        const bx = Math.round(cx - boxW / 2);
                        const by = Math.round(cy - boxH / 2);
                        const radius = Math.max(3, Math.round(fontSize * 0.25));

                        ctx.beginPath();
                        if (typeof ctx.roundRect === 'function') {
                            ctx.roundRect(bx, by, boxW, boxH, radius);
                        } else {
                            ctx.rect(bx, by, boxW, boxH);
                        }
                        ctx.fillStyle = mainColor;
                        ctx.fill();
                        ctx.lineWidth = Math.max(1.5, Math.round(2.5 * scale * labelScale));
                        ctx.strokeStyle = '#0F172A';
                        ctx.stroke();

                        ctx.fillStyle = isRed ? '#FFFFFF' : '#0F172A';
                        ctx.textAlign = 'center';
                        ctx.textBaseline = 'middle';
                        ctx.fillText(text, Math.round(cx), Math.round(cy));
                        ctx.restore();
                    }
                });
            }

            async renderOrientedCanvas(photoItem, maxDimension = 1600, { includeAnnotations = true } = {}) {
                return new Promise((resolve, reject) => {
                    const img = new Image();
                    img.onload = async () => {
                        let source = img;
                        let exifRotation = 0;
                        try {
                            if (photoItem.exifRotation && typeof createImageBitmap === 'function') {
                                source = await createImageBitmap(photoItem.file, { imageOrientation: 'none' });
                                exifRotation = photoItem.exifRotation;
                            }
                            const canvas = document.createElement('canvas');
                            let w = source.width;
                            let h = source.height;
                            const rot = (((photoItem.rotation || 0) + exifRotation) % 360 + 360) % 360;

                            if (w > maxDimension || h > maxDimension) {
                                if (w > h) {
                                    h = Math.round((h * maxDimension) / w);
                                    w = maxDimension;
                                } else {
                                    w = Math.round((w * maxDimension) / h);
                                    h = maxDimension;
                                }
                            }

                            if (rot === 90 || rot === 270) {
                                canvas.width = h;
                                canvas.height = w;
                            } else {
                                canvas.width = w;
                                canvas.height = h;
                            }

                            const ctx = canvas.getContext('2d');
                            if (!ctx) return reject(new Error('瀏覽器無法建立影像處理畫布。'));
                            ctx.fillStyle = '#FFFFFF';
                            ctx.fillRect(0, 0, canvas.width, canvas.height);

                            ctx.save();
                            ctx.translate(canvas.width / 2, canvas.height / 2);
                            if (photoItem.flipH) ctx.scale(-1, 1);
                            ctx.rotate((rot * Math.PI) / 180);
                            ctx.drawImage(source, -w / 2, -h / 2, w, h);
                            ctx.restore();

                            if (includeAnnotations && Array.isArray(photoItem.annotations) && photoItem.annotations.length > 0) {
                                this.drawAnnotationsOnCanvas(ctx, canvas.width, canvas.height, photoItem.annotations);
                            }
                            resolve(canvas);
                        } catch (error) {
                            reject(new Error(`無法處理圖片：${photoItem.name}（${error.message}）`));
                        } finally {
                            if (source !== img && typeof source.close === 'function') source.close();
                        }
                    };
                    img.onerror = () => reject(new Error(`無法解碼圖片：${photoItem.name}`));
                    img.src = photoItem.src;
                });
            }

            async renderCompositeCanvas(photoItem, maxDimension = 1600) {
                if (!photoItem || !photoItem.sideCrop || !photoItem.sideCrop.cropRect) {
                    return this.renderOrientedCanvas(photoItem, maxDimension, { includeAnnotations: true });
                }
                const srcMax = Math.max(maxDimension, 2400);
                const baseCanvas = await this.renderOrientedCanvas(photoItem, srcMax, { includeAnnotations: false });
                const hiW = baseCanvas.width;
                const hiH = baseCanvas.height;
                const geo = this.computeSideCropGeometry(hiW, hiH, photoItem.sideCrop);
                if (!geo) return baseCanvas;

                const minTargetScale = Math.max(hiW, hiH) < 360 ? (360 / Math.max(hiW, hiH)) : 1;
                const outScale = Math.min(minTargetScale, maxDimension / Math.max(geo.totalW, geo.totalH));
                const outTotalW = Math.max(8, Math.round(geo.totalW * outScale));
                const outTotalH = Math.max(8, Math.round(geo.totalH * outScale));
                const outOrigW = Math.max(4, Math.round(geo.origW * outScale));
                const outOrigH = outTotalH;
                const outGap = Math.max(2, Math.round(geo.gap * outScale));
                const outPanelW = Math.max(4, outTotalW - outOrigW - outGap);
                const outPanelH = Math.max(4, Math.min(outTotalH, Math.round(geo.panelH * outScale)));
                const outPanelTop = Math.round(geo.offsetY * Math.max(0, outTotalH - outPanelH));

                const outCanvas = document.createElement('canvas');
                outCanvas.width = outTotalW;
                outCanvas.height = outTotalH;
                const outCtx = outCanvas.getContext('2d');
                if (!outCtx) throw new Error('瀏覽器無法建立並排合成畫布。');

                outCtx.fillStyle = '#FFFFFF';
                outCtx.fillRect(0, 0, outTotalW, outTotalH);

                // 左側：原始方向照片＋標記框與姓名標籤
                outCtx.drawImage(baseCanvas, 0, 0, hiW, hiH, 0, 0, outOrigW, outOrigH);
                if (Array.isArray(photoItem.annotations) && photoItem.annotations.length > 0) {
                    outCtx.save();
                    outCtx.beginPath();
                    outCtx.rect(0, 0, outOrigW, outOrigH);
                    outCtx.clip();
                    this.drawAnnotationsOnCanvas(outCtx, outOrigW, outOrigH, photoItem.annotations);
                    outCtx.restore();
                }

                // 右側：框選局部高解析特寫＋範圍內標籤＋外框
                const cr = photoItem.sideCrop.cropRect;
                const sx = Math.max(0, Math.min(hiW - 2, Math.round((Number(cr.x) || 0) * hiW)));
                const sy = Math.max(0, Math.min(hiH - 2, Math.round((Number(cr.y) || 0) * hiH)));
                const sw = Math.max(2, Math.min(hiW - sx, Math.round((Number(cr.w) || 0.2) * hiW)));
                const sh = Math.max(2, Math.min(hiH - sy, Math.round((Number(cr.h) || 0.2) * hiH)));
                const dx = outOrigW + outGap;
                const dy = outPanelTop;
                outCtx.drawImage(baseCanvas, sx, sy, sw, sh, dx, dy, outPanelW, outPanelH);

                const cw = Math.max(0.02, Math.min(1, Number(cr.w) || 0.2));
                const ch = Math.max(0.02, Math.min(1, Number(cr.h) || 0.2));
                const nonSourceAnns = (photoItem.annotations || []).filter(a => a.id !== photoItem.sideCrop?.sourceAnnId);
                if (nonSourceAnns.length > 0) {
                    outCtx.save();
                    outCtx.beginPath();
                    outCtx.rect(dx, dy, outPanelW, outPanelH);
                    outCtx.clip();
                    const virtW = outPanelW / cw;
                    const virtH = outPanelH / ch;
                    outCtx.translate(dx - (Number(cr.x) || 0) * virtW, dy - (Number(cr.y) || 0) * virtH);
                    this.drawAnnotationsOnCanvas(outCtx, virtW, virtH, nonSourceAnns);
                    outCtx.restore();
                }

                if (geo.borderColor !== 'none') {
                    const minSideOut = Math.max(120, Math.min(outTotalW, outTotalH));
                    const frameW = Math.max(2, Math.round(4.5 * (minSideOut / 800)));
                    const haloW = frameW + Math.max(1.5, Math.round(2 * (minSideOut / 800)));
                    const bColor = geo.borderColor === 'red' ? '#EF4444' : '#FFD23F';
                    outCtx.save();
                    outCtx.strokeStyle = 'rgba(15, 23, 42, 0.85)';
                    outCtx.lineWidth = haloW;
                    outCtx.strokeRect(dx + haloW / 2, dy + haloW / 2, Math.max(2, outPanelW - haloW), Math.max(2, outPanelH - haloW));
                    outCtx.strokeStyle = bColor;
                    outCtx.lineWidth = frameW;
                    outCtx.strokeRect(dx + haloW / 2, dy + haloW / 2, Math.max(2, outPanelW - haloW), Math.max(2, outPanelH - haloW));
                    outCtx.restore();
                }

                return outCanvas;
            }

            async renderSideCropStandaloneCanvas(photoItem, maxDimension = 1600) {
                const srcMax = Math.max(maxDimension, 2400);
                const baseCanvas = await this.renderOrientedCanvas(photoItem, srcMax, { includeAnnotations: false });
                const hiW = baseCanvas.width;
                const hiH = baseCanvas.height;
                const geo = this.computeSideCropGeometry(hiW, hiH, photoItem.sideCrop);
                if (!geo) return baseCanvas;

                const cr = photoItem.sideCrop.cropRect;
                const sx = Math.max(0, Math.min(hiW - 2, Math.round((Number(cr.x) || 0) * hiW)));
                const sy = Math.max(0, Math.min(hiH - 2, Math.round((Number(cr.y) || 0) * hiH)));
                const sw = Math.max(2, Math.min(hiW - sx, Math.round((Number(cr.w) || 0.2) * hiW)));
                const sh = Math.max(2, Math.min(hiH - sy, Math.round((Number(cr.h) || 0.2) * hiH)));

                const rawMax = Math.max(sw, sh);
                const scale = rawMax < 480 ? (480 / rawMax) : Math.min(1, maxDimension / rawMax);
                const outW = Math.max(8, Math.round(sw * scale));
                const outH = Math.max(8, Math.round(sh * scale));

                const cropCanvas = document.createElement('canvas');
                cropCanvas.width = outW;
                cropCanvas.height = outH;
                const ctx = cropCanvas.getContext('2d');
                if (!ctx) throw new Error('瀏覽器無法建立局部特寫畫布。');

                ctx.fillStyle = '#FFFFFF';
                ctx.fillRect(0, 0, outW, outH);
                ctx.drawImage(baseCanvas, sx, sy, sw, sh, 0, 0, outW, outH);

                const cw = Math.max(0.02, Math.min(1, Number(cr.w) || 0.2));
                const ch = Math.max(0.02, Math.min(1, Number(cr.h) || 0.2));
                const nonSourceAnns = (photoItem.annotations || []).filter(a => a.id !== photoItem.sideCrop?.sourceAnnId);
                if (nonSourceAnns.length > 0) {
                    ctx.save();
                    ctx.beginPath();
                    ctx.rect(0, 0, outW, outH);
                    ctx.clip();
                    const virtW = outW / cw;
                    const virtH = outH / ch;
                    ctx.translate(-(Number(cr.x) || 0) * virtW, -(Number(cr.y) || 0) * virtH);
                    this.drawAnnotationsOnCanvas(ctx, virtW, virtH, nonSourceAnns);
                    ctx.restore();
                }

                if (geo.borderColor !== 'none') {
                    const minSideOut = Math.max(160, Math.min(outW, outH));
                    const frameW = Math.max(2.5, Math.round(5 * (minSideOut / 520)));
                    const haloW = frameW + Math.max(1.5, Math.round(2 * (minSideOut / 520)));
                    const bColor = geo.borderColor === 'red' ? '#EF4444' : '#FFD23F';
                    ctx.save();
                    ctx.strokeStyle = 'rgba(15, 23, 42, 0.85)';
                    ctx.lineWidth = haloW;
                    ctx.strokeRect(haloW / 2, haloW / 2, Math.max(2, outW - haloW), Math.max(2, outH - haloW));
                    ctx.strokeStyle = bColor;
                    ctx.lineWidth = frameW;
                    ctx.strokeRect(haloW / 2, haloW / 2, Math.max(2, outW - haloW), Math.max(2, outH - haloW));
                    ctx.restore();
                }

                return {
                    canvas: cropCanvas,
                    widthRatio: geo.widthRatio,
                    offsetY: geo.offsetY
                };
            }

            canvasToJpegBlob(canvas, photoName, quality = 0.88) {
                return new Promise((resolve, reject) => {
                    canvas.toBlob((blob) => {
                        if (!blob) return reject(new Error(`無法轉換圖片：${photoName}`));
                        resolve({
                            blob,
                            width: canvas.width,
                            height: canvas.height
                        });
                    }, 'image/jpeg', quality);
                });
            }

            async processImageToBlob(photoItem, maxDimension = 1600, quality = 0.88) {
                if (photoItem && photoItem.sideCrop && photoItem.sideCrop.cropRect) {
                    const [compCanvas, mainCanvas, sideResult] = await Promise.all([
                        this.renderCompositeCanvas(photoItem, maxDimension),
                        this.renderOrientedCanvas(photoItem, maxDimension, { includeAnnotations: true }),
                        this.renderSideCropStandaloneCanvas(photoItem, maxDimension)
                    ]);
                    const [compOut, mainOut, sideOut] = await Promise.all([
                        this.canvasToJpegBlob(compCanvas, photoItem.name, quality),
                        this.canvasToJpegBlob(mainCanvas, photoItem.name, quality),
                        this.canvasToJpegBlob(sideResult.canvas, photoItem.name, Math.min(0.92, quality + 0.02))
                    ]);
                    return {
                        blob: compOut.blob,
                        width: compOut.width,
                        height: compOut.height,
                        mainImage: mainOut,
                        sideImage: {
                            ...sideOut,
                            widthRatio: sideResult.widthRatio,
                            offsetY: sideResult.offsetY
                        }
                    };
                }
                const canvas = await this.renderOrientedCanvas(photoItem, maxDimension, { includeAnnotations: true });
                return this.canvasToJpegBlob(canvas, photoItem.name, quality);
            }

            async createDetailCropPhoto(parentIndex, cropRect) {
                const parentPhoto = this.photos[parentIndex];
                if (!parentPhoto || !cropRect || this.isCreatingCropDetail) return null;
                const x = Math.max(0, Math.min(0.98, Number(cropRect.x) || 0));
                const y = Math.max(0, Math.min(0.98, Number(cropRect.y) || 0));
                const w = Math.max(0.02, Math.min(1 - x, Number(cropRect.w) || 0.2));
                const h = Math.max(0.02, Math.min(1 - y, Number(cropRect.h) || 0.2));

                this.flushPendingHistory();
                this.commitLightboxHistoryIfNeeded();
                this.isCreatingCropDetail = true;
                try {
                    const orientedCanvas = await this.renderOrientedCanvas(parentPhoto, 3200, { includeAnnotations: false });
                    const sx = Math.max(0, Math.min(orientedCanvas.width - 2, Math.round(x * orientedCanvas.width)));
                    const sy = Math.max(0, Math.min(orientedCanvas.height - 2, Math.round(y * orientedCanvas.height)));
                    const sw = Math.max(2, Math.min(orientedCanvas.width - sx, Math.round(w * orientedCanvas.width)));
                    const sh = Math.max(2, Math.min(orientedCanvas.height - sy, Math.round(h * orientedCanvas.height)));

                    const maxOut = 1600;
                    const outScale = Math.min(1, maxOut / Math.max(sw, sh));
                    const outW = Math.max(2, Math.round(sw * outScale));
                    const outH = Math.max(2, Math.round(sh * outScale));

                    const cropCanvas = document.createElement('canvas');
                    cropCanvas.width = outW;
                    cropCanvas.height = outH;
                    const cropCtx = cropCanvas.getContext('2d');
                    if (!cropCtx) throw new Error('瀏覽器無法建立局部裁切畫布。');
                    cropCtx.fillStyle = '#FFFFFF';
                    cropCtx.fillRect(0, 0, outW, outH);
                    cropCtx.drawImage(orientedCanvas, sx, sy, sw, sh, 0, 0, outW, outH);

                    const blob = await new Promise((resolve, reject) => cropCanvas.toBlob(
                        result => result ? resolve(result) : reject(new Error('無法產生局部裁切圖片。')),
                        'image/jpeg',
                        0.92
                    ));
                    const detailName = this.makeUniqueDetailFilename(parentPhoto);
                    const file = new File([blob], detailName, { type: 'image/jpeg' });
                    const sourceUrl = URL.createObjectURL(file);
                    let thumbnail;
                    try {
                        thumbnail = await this.createThumbnailUrl(file, sourceUrl);
                    } catch (thumbErr) {
                        URL.revokeObjectURL(sourceUrl);
                        throw thumbErr;
                    }

                    const detailPhoto = {
                        uid: 'uid_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9),
                        file,
                        name: detailName,
                        src: sourceUrl,
                        thumbnailSrc: thumbnail.src,
                        thumbnailAspect: thumbnail.aspect,
                        visualFingerprint: thumbnail.visualFingerprint || '',
                        rotation: 0,
                        flipH: false,
                        parentUid: parentPhoto.uid,
                        cropRect: {
                            x: Math.round(x * 10000) / 10000,
                            y: Math.round(y * 10000) / 10000,
                            w: Math.round(w * 10000) / 10000,
                            h: Math.round(h * 10000) / 10000
                        },
                        sideCrop: null,
                        compositeThumbnailSrc: null,
                        _compositeKey: '',
                        annotations: [],
                        exifRotation: 0,
                        seq: `${parentPhoto.seq}-1`,
                        date: parentPhoto.date || '',
                        time: parentPhoto.time || '',
                        timeAutoDefaulted: Boolean(parentPhoto.timeAutoDefaulted),
                        location: parentPhoto.location || '',
                        desc: parentPhoto.desc ? `${parentPhoto.desc}（局部特寫）` : '',
                        selected: false,
                        stageX: null,
                        stageY: null
                    };

                    let insertAt = parentIndex + 1;
                    while (insertAt < this.photos.length && this.isDescendantOf(this.photos[insertAt], parentPhoto.uid)) {
                        insertAt++;
                    }
                    this.photos.splice(insertAt, 0, detailPhoto);
                    this.managedUrls.add(sourceUrl);
                    this.managedUrls.add(detailPhoto.thumbnailSrc);
                    this.normalizeLinkedChildrenOrder();

                    const finalIndex = this.photos.indexOf(detailPhoto);
                    this.photos.forEach((p, idx) => { p.selected = idx === finalIndex; });
                    this.currentIndex = finalIndex;
                    this.lastSelectedIndex = finalIndex;
                    this.lightboxCropDraft = null;
                    this.setLightboxTool('pan');
                    this.render();
                    this.recordHistory('建立局部裁切特寫');
                    return detailPhoto;
                } finally {
                    this.isCreatingCropDetail = false;
                }
            }

            // ============================================================
            // 📑 100% 精準對齊原 VBA 範本之 OpenXML 報表生成引擎
            // (全標楷體、頁首頁尾、自然行距、圖片置中、紙張橫直向精確設定)
            // ============================================================
            // 📑 報表匯出委派引擎 (Phase 4 模組化抽離：Word / PDF / Excel)
            // ============================================================
            getExportContext() {
                const caseData = this.getCaseData();
                return {
                    photos: this.photos,
                    caseTitle: caseData.caseTitle || '',
                    caseDate: caseData.caseDate || '',
                    defaultLocation: caseData.defaultLocation || '',
                    deptName: caseData.deptName || '臺南市政府警察局新化分局',
                    globalOfficer: caseData.officerName || '巡官梁家豪',
                    layout: this.layoutSelect.value,
                    processImageToBlob: (photo, maxDim, quality) => this.processImageToBlob(photo, maxDim, quality),
                    setProgress: (msg, cur, total) => this.setProgress(msg, cur, total),
                    clearProgress: () => this.clearProgress(),
                    throwIfCancelled: () => this.throwIfProcessingCancelled(),
                    yieldToBrowser: () => this.yieldToBrowser(),
                    saveAs: typeof saveAs === 'function' ? saveAs : (window.saveAs || null)
                };
            }

            async exportDocx() {
                if (typeof PhotoReportDocxExporter !== 'undefined') {
                    this.processingCancelled = false;
                    return PhotoReportDocxExporter.exportDocx(this.getExportContext());
                }
                alert('Word 匯出模組尚未載入！');
            }

            blobToDataUrl(blob) {
                if (typeof PhotoReportPdfExporter !== 'undefined') {
                    return PhotoReportPdfExporter.blobToDataUrl(blob);
                }
                return new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve(reader.result);
                    reader.onerror = () => reject(new Error('PDF 圖片轉換失敗。'));
                    reader.readAsDataURL(blob);
                });
            }

            async exportPdf() {
                if (typeof PhotoReportPdfExporter !== 'undefined') {
                    this.processingCancelled = false;
                    return PhotoReportPdfExporter.exportPdf(this.getExportContext());
                }
                alert('PDF 匯出模組尚未載入！');
            }

            exportExcel() {
                if (typeof PhotoReportExcelExporter !== 'undefined') {
                    const caseData = this.getCaseData();
                    return PhotoReportExcelExporter.exportExcel({
                        photos: this.photos,
                        caseTitle: caseData.caseTitle || '',
                        caseDate: caseData.caseDate || '',
                        defaultLocation: caseData.defaultLocation || '',
                        globalOfficer: caseData.officerName || '巡官梁家豪'
                    });
                }
                alert('Excel 匯出模組尚未載入！');
            }

            importExcel(e) {
                const file = e.target.files[0];
                if (!file) return;
                if (typeof XLSX === 'undefined') return alert('SheetJS 尚未載入！');

                const reader = new FileReader();
                reader.onload = (evt) => {
                    try {
                        const data = new Uint8Array(evt.target.result);
                        const wb = XLSX.read(data, { type: 'array' });
                        const wsName = wb.SheetNames[0];
                        const ws = wb.Sheets[wsName];
                        const json = XLSX.utils.sheet_to_json(ws, { header: 1 });

                        if (json.length < 2) return alert('Excel 內無有效資料！');

                        const headerRow = json[0];
                        const idxSeq = headerRow.findIndex(h => /編號|序號/i.test(h));
                        const idxName = headerRow.findIndex(h => /檔名|照片/i.test(h));
                        const idxDate = headerRow.findIndex(h => /日期/i.test(h));
                        const idxTime = headerRow.findIndex(h => /時間/i.test(h));
                        const idxLoc = headerRow.findIndex(h => /地點/i.test(h));
                        const idxDesc = headerRow.findIndex(h => /說明|內容/i.test(h));

                        let matchCount = 0;
                        let ambiguousCount = 0;

                        for (let r = 1; r < json.length; r++) {
                            const row = json[r];
                            if (!row || !row.length) continue;

                            const rowSeq = idxSeq !== -1 ? String(row[idxSeq] || '').trim() : '';
                            const rowName = idxName !== -1 ? String(row[idxName] || '').trim() : '';
                            const matches = this.photos.filter(p => (rowName && p.name === rowName) || (rowSeq && p.seq === rowSeq));

                            if (matches.length === 1) {
                                const target = matches[0];
                                if (idxDate !== -1 && row[idxDate]) target.date = String(row[idxDate]).trim();
                                if (idxTime !== -1 && row[idxTime]) target.time = String(row[idxTime]).trim();
                                if (idxLoc !== -1 && row[idxLoc]) target.location = String(row[idxLoc]).trim();
                                if (idxDesc !== -1 && row[idxDesc]) target.desc = String(row[idxDesc]).trim();
                                matchCount++;
                            } else if (matches.length > 1) {
                                ambiguousCount++;
                            }
                        }

                        this.render();
                        const ambiguousMessage = ambiguousCount ? `另有 ${ambiguousCount} 筆因編號或檔名重複而略過。` : '';
                        alert(`已從 Excel 帶入 ${matchCount} 筆說明。${ambiguousMessage}`);
                    } catch (error) {
                        console.error(error);
                        alert('Excel 讀取失敗，請確認檔案未損毀且格式正確。');
                    }
                };
                reader.onerror = () => alert('Excel 檔案讀取失敗。');
                reader.readAsArrayBuffer(file);
                e.target.value = '';
            }

            async exportZip() {
                if (!this.photos.length) return alert('請先加入照片。');
                if (typeof JSZip === 'undefined') return alert('JSZip 尚未載入！');
                this.processingCancelled = false;

                const zip = new JSZip();
                const caseTitle = document.getElementById('caseTitle').value.trim() || '現場照片';

                this.setProgress('正在壓縮照片並建立 ZIP…', 0, this.photos.length);
                for (let i = 0; i < this.photos.length; i++) {
                    const p = this.photos[i];
                    const processed = await this.processImageToBlob(p, 1600, 0.88);
                    const ext = 'jpg';
                    const filename = `${p.seq}_${p.name.replace(/\.[^/.]+$/, '')}.${ext}`;
                    zip.file(filename, processed.blob);
                    this.setProgress('正在壓縮照片並建立 ZIP…', i + 1, this.photos.length);
                    this.throwIfProcessingCancelled();
                    if ((i + 1) % 5 === 0) await this.yieldToBrowser();
                }

                this.setProgress('正在封裝 ZIP 檔案…', this.photos.length, this.photos.length);
                const content = await zip.generateAsync({ type: 'blob' });
                saveAs(content, `${caseTitle}_最佳化照片包.zip`);
                this.clearProgress();
            }
        }

        window.addEventListener('DOMContentLoaded', () => {
            window.app = new PhotoReportApp();
        });
