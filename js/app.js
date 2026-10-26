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
                this.sidePanelWidth = Math.max(280, Math.min(560, this.loadWorkspacePreference('sidePanelWidth', 300)));
                this.panelResizeState = null;
                this.marqueeState = null;
                this.history = [];
                this.historyIndex = -1;
                this.historyLimit = 50;
                this.historyTimer = null;
                this.isRestoringHistory = false;
                this.managedUrls = new Set();
                this.processingCancelled = false;
                this.lightboxIndex = null;
                this.savedProjectSignature = '';
                this.activeFilter = 'all';
                this.pendingExportAction = null;
                this.pendingSequenceExport = null;
                this.initElements();
                this.initEvents();
                this.applyWorkspaceLayout();
                this.setDefaultTodayDate();
                this.updateLayoutDescription();
                this.loadAppVersion();
                this.render();
                this.recordHistory('初始狀態');
                this.markProjectSaved();
            }

            initElements() {
                this.fileInput = document.getElementById('fileInput');
                this.folderInput = document.getElementById('folderInput');
                this.workspaceContainer = document.getElementById('workspaceContainer');
                this.appMainLayout = document.getElementById('appMainLayout');
                this.mainPanelSplitter = document.getElementById('mainPanelSplitter');
                this.thumbGrid = document.getElementById('thumbGrid');
                this.photoPanel = document.getElementById('photoPanel');
                this.editorPanel = document.getElementById('editorPanel');
                this.workspaceLayoutControls = document.getElementById('workspaceLayoutControls');
                this.thumbnailSizeLabel = document.getElementById('thumbnailSizeLabel');
                this.thumbnailSizeRange = document.getElementById('thumbnailSizeRange');
                this.selectionMarquee = document.getElementById('selectionMarquee');
                this.btnToggleEditor = document.getElementById('btnToggleEditor');
                this.gridContainer = document.getElementById('gridContainer');
                this.previewContainer = document.getElementById('previewContainer');
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

                // 右側編輯器元件
                this.editorTitle = document.getElementById('editorTitle');
                this.editorModeBadge = document.getElementById('editorModeBadge');
                this.editSeq = document.getElementById('editSeq');
                this.editDate = document.getElementById('editDate');
                this.editTime = document.getElementById('editTime');
                this.editLocation = document.getElementById('editLocation');
                this.editDesc = document.getElementById('editDesc');
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

                this.lightboxZoom = 1.0;
                this.lightboxPanX = 0;
                this.lightboxPanY = 0;
                this.lightboxIsPanning = false;
                this.lightboxPanStart = { x: 0, y: 0 };
                this.lightboxHasRotated = false;
            }

            initEvents() {
                // 大圖燈箱縮放、平移與滾輪互動
                this.initLightboxInteractions();

                // 外部檔案拖入空白狀態或照片舞台；既有卡片拖曳排序不含 Files，不會被攔截。
                this.bindExternalFileDropTarget(this.emptyState);
                this.bindExternalFileDropTarget(this.thumbGrid);

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
                document.getElementById('btnViewPreview').addEventListener('click', () => this.switchView('preview'));
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
                document.getElementById('btnSaveProject').addEventListener('click', () => this.saveProject().catch(error => this.handleProjectError(error)));
                document.getElementById('importProjectInput').addEventListener('change', (e) => this.openProject(e));
                document.getElementById('btnUndo').addEventListener('click', () => this.undo());
                document.getElementById('btnRedo').addEventListener('click', () => this.redo());
                document.getElementById('btnExportExcel').addEventListener('click', () => this.confirmExportWithAudit(() => this.exportExcel(), 'Excel 清冊'));
                document.getElementById('btnExportZip').addEventListener('click', () => this.exportZip().catch(error => this.handleExportError(error)));
                document.getElementById('btnExportPdf').addEventListener('click', () => this.confirmExportWithAudit(() => this.exportPdf().catch(e => this.handleExportError(e)), 'PDF 清冊'));
                document.getElementById('importExcelInput').addEventListener('change', (e) => this.importExcel(e));
                this.layoutSelect.addEventListener('change', () => {
                    this.updateLayoutDescription();
                    this.updatePageEstimate();
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
                ['caseTitle', 'caseDate', 'defaultLocation', 'deptName', 'officerName'].forEach(id => {
                    document.getElementById(id).addEventListener('input', () => this.scheduleHistory('修改案件基本資料'));
                });

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
                this.editSeq.addEventListener('input', (e) => this.updateCurrentPhoto('seq', e.target.value));
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
                this.editLocation.addEventListener('input', (e) => this.updateCurrentPhoto('location', e.target.value));
                this.editDesc.addEventListener('input', (e) => this.updateCurrentPhoto('desc', e.target.value));

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

                // 鍵盤快速鍵：方向鍵切換單張選取；Ctrl＋方向鍵與 Ctrl＋Home／End 重排已選照片。
                window.addEventListener('keydown', (e) => {
                    if (this.isTypingTarget(e.target)) return;
                    if (e.key === 'Escape') {
                        ['shortcutsModal', 'sequenceModal', 'updateModal', 'upToDateModal', 'updateErrorModal', 'exportAuditModal', 'imageLightboxModal'].forEach(id => {
                            this.closeModal(id);
                        });
                    }
                    const lightboxOpen = !document.getElementById('imageLightboxModal').classList.contains('hidden');
                    if (lightboxOpen) {
                        if (e.key === 'Escape') {
                            this.closeModal('imageLightboxModal');
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
                        const photoItem = {
                            uid: 'uid_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9),
                            file: file,
                            name: file.name,
                            src: sourceUrl,
                            thumbnailSrc: thumbnail.src,
                            thumbnailAspect: thumbnail.aspect,
                            rotation: 0,
                            exifRotation: await this.readJpegExifRotation(file),
                            seq: String(this.photos.length + 1),
                            date: '',
                            time: this.extractTimeFromFilename(file.name) || '12:00',
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
                        aspect: canvas.width / canvas.height
                    };
                } finally {
                    if (source && typeof source.close === 'function') source.close();
                }
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
                const hasSelected = this.photos.some(p => p.selected);
                this.photos.forEach(p => {
                    if (!hasSelected || p.selected) {
                        p.rotation = (p.rotation + degree) % 360;
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

                this.photos.forEach((p, idx) => {
                    const num = start + idx;
                    const numStr = String(num).padStart(digits, '0');
                    p.seq = `${prefix}${numStr}`;
                });

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
                        updated++;
                    }
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
                    if (timeVal || applyEmptyValues) p.time = timeVal;
                    if (locVal || applyEmptyValues) p.location = locVal;
                    if (descVal || applyEmptyValues) p.desc = descVal;
                });

                this.render();
                this.recordHistory('批次套用編輯欄位');
                alert(`已套用到 ${selectedList.length} 張照片。`);
            }

            switchView(view) {
                this.currentView = view;
                this.workspaceContainer.classList.toggle('hidden', view !== 'workspace');
                this.gridContainer.classList.toggle('hidden', view !== 'grid');
                this.previewContainer.classList.toggle('hidden', view !== 'preview');
                const activeClass = "px-2.5 py-1 rounded text-xs font-bold bg-[#5B5CE2] text-white flex items-center gap-1 shadow-2xs";
                const idleClass = "px-2.5 py-1 rounded text-xs font-bold text-[#6B7280] hover:text-[#20242B] flex items-center gap-1";
                document.getElementById('btnViewWorkspace').className = view === 'workspace' ? activeClass : idleClass;
                document.getElementById('btnViewGrid').className = view === 'grid' ? activeClass : idleClass;
                document.getElementById('btnViewPreview').className = view === 'preview' ? activeClass : idleClass;
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
                this.updateLayoutDescription();
            }

            snapshotState(label = '') {
                return {
                    label,
                    caseData: this.getCaseData(),
                    currentIndex: this.currentIndex,
                    lastSelectedIndex: this.lastSelectedIndex,
                    photos: this.photos.map(photo => ({ ...photo }))
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
                    photos: state.photos.map(({ uid, rotation, seq, date, time, location, desc, selected, stageX, stageY }) =>
                        ({ uid, rotation, seq, date, time, location, desc, selected, stageX, stageY }))
                });
            }

            isExternalFileDrag(event) {
                return Array.from(event.dataTransfer?.types || []).includes('Files');
            }

            bindExternalFileDropTarget(target) {
                if (!target) return;
                target.addEventListener('dragenter', (event) => {
                    if (!this.isExternalFileDrag(event)) return;
                    event.preventDefault();
                    target.classList.add('external-file-drag-over');
                });
                target.addEventListener('dragover', (event) => {
                    if (!this.isExternalFileDrag(event)) return;
                    event.preventDefault();
                    event.dataTransfer.dropEffect = 'copy';
                    target.classList.add('external-file-drag-over');
                });
                target.addEventListener('dragleave', (event) => {
                    if (!this.isExternalFileDrag(event) || target.contains(event.relatedTarget)) return;
                    target.classList.remove('external-file-drag-over');
                });
                target.addEventListener('drop', (event) => {
                    if (!this.isExternalFileDrag(event)) return;
                    event.preventDefault();
                    target.classList.remove('external-file-drag-over');
                    const files = Array.from(event.dataTransfer.files).filter(file => file.type.startsWith('image/'));
                    if (!files.length) return alert('請拖入支援的圖片檔案。');
                    this.handleFiles(files).catch(error => this.handleProcessingError(error));
                });
            }

            projectSignature() {
                if (typeof PhotoReportHistory !== 'undefined') {
                    return PhotoReportHistory.projectSignature(this.getCaseData(), this.photos);
                }
                return JSON.stringify({
                    caseData: this.getCaseData(),
                    photos: this.photos.map(({ uid, rotation, seq, date, time, location, desc, stageX, stageY }) =>
                        ({ uid, rotation, seq, date, time, location, desc, stageX, stageY }))
                });
            }

            markProjectSaved() {
                this.savedProjectSignature = this.projectSignature();
            }

            hasUnsavedProjectChanges() {
                return this.projectSignature() !== this.savedProjectSignature;
            }

            recordHistory(label = '') {
                if (this.isRestoringHistory) return;
                const next = this.snapshotState(label);
                const current = this.history[this.historyIndex];
                if (current && this.historySignature(current) === this.historySignature(next)) return;
                this.history.splice(this.historyIndex + 1);
                this.history.push(next);
                if (this.history.length > this.historyLimit) this.history.shift();
                this.historyIndex = this.history.length - 1;
                this.releaseUnusedPhotoUrls();
                this.updateHistoryControls();
            }

            scheduleHistory(label) {
                if (this.isRestoringHistory) return;
                clearTimeout(this.historyTimer);
                this.historyTimer = setTimeout(() => this.recordHistory(label), 450);
            }

            restoreHistory(index) {
                const state = this.history[index];
                if (!state) return;
                clearTimeout(this.historyTimer);
                this.isRestoringHistory = true;
                this.photos = state.photos.map(photo => ({ ...photo }));
                this.currentIndex = state.currentIndex;
                this.lastSelectedIndex = state.lastSelectedIndex;
                this.applyCaseData(state.caseData);
                this.isRestoringHistory = false;
                this.historyIndex = index;
                this.render();
                this.updateHistoryControls();
            }

            undo() {
                if (this.historyIndex <= 0) return;
                this.restoreHistory(this.historyIndex - 1);
            }

            redo() {
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

            async saveProject() {
                if (!this.photos.length) return alert('還沒有照片，不需要儲存。');
                if (typeof JSZip === 'undefined') throw new Error('JSZip 尚未載入，無法儲存專案。');
                const zip = new JSZip();
                const manifestPhotos = [];
                this.setProgress('正在封裝專案檔…', 0, this.photos.length);
                for (let index = 0; index < this.photos.length; index++) {
                    const photo = this.photos[index];
                    const archiveName = `${String(index + 1).padStart(4, '0')}_${photo.uid}_${photo.name}`;
                    zip.file(`images/${archiveName}`, photo.file);
                    manifestPhotos.push({
                        uid: photo.uid, archiveName, name: photo.name, type: photo.file.type,
                        rotation: photo.rotation, exifRotation: photo.exifRotation || 0, seq: photo.seq, date: photo.date, time: photo.time,
                        location: photo.location, desc: photo.desc, selected: photo.selected,
                        stageX: Number.isFinite(photo.stageX) ? photo.stageX : null,
                        stageY: Number.isFinite(photo.stageY) ? photo.stageY : null
                    });
                    this.setProgress('正在封裝專案檔…', index + 1, this.photos.length);
                    if ((index + 1) % 10 === 0) await this.yieldToBrowser();
                }
                zip.file('manifest.json', JSON.stringify({
                    format: 'photo-report-project', version: 1, savedAt: new Date().toISOString(),
                    caseData: this.getCaseData(), currentIndex: this.currentIndex,
                    lastSelectedIndex: this.lastSelectedIndex, photos: manifestPhotos
                }));
                const content = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
                const caseTitle = this.getCaseData().caseTitle || '現場照片清冊';
                saveAs(content, `${caseTitle}_照片清冊專案.photo-report`);
                this.markProjectSaved();
                this.clearProgress();
            }

            async openProject(event) {
                const file = event.target.files?.[0];
                event.target.value = '';
                if (!file) return;
                if (typeof JSZip === 'undefined') return this.handleProjectError(new Error('JSZip 尚未載入，無法開啟專案。'));
                if (this.photos.length && !confirm('開啟後會取代目前的照片與資料，沒存的修改會不見。要繼續嗎？')) return;
                try {
                    const zip = await JSZip.loadAsync(file);
                    const manifestFile = zip.file('manifest.json');
                    if (!manifestFile) throw new Error('這不是有效的照片清冊專案檔（缺少 manifest.json）。');
                    const manifest = JSON.parse(await manifestFile.async('text'));
                    if (manifest.format !== 'photo-report-project' || !Array.isArray(manifest.photos)) {
                        throw new Error('專案檔格式不支援或內容不完整。');
                    }
                    this.setProgress('正在還原專案照片…', 0, manifest.photos.length);
                    const restored = [];
                    for (let index = 0; index < manifest.photos.length; index++) {
                        const saved = manifest.photos[index];
                        const entry = zip.file(`images/${saved.archiveName}`);
                        if (!entry) throw new Error(`專案檔缺少照片：${saved.name}`);
                        const blob = await entry.async('blob');
                        const imageFile = new File([blob], saved.name, { type: saved.type || blob.type || 'image/jpeg' });
                        const src = URL.createObjectURL(imageFile);
                        const thumbnail = await this.createThumbnailUrl(imageFile, src);
                        restored.push({
                            uid: saved.uid || crypto.randomUUID(), file: imageFile, name: saved.name,
                            src, thumbnailSrc: thumbnail.src, thumbnailAspect: thumbnail.aspect,
                            rotation: Number(saved.rotation) || 0, exifRotation: Number(saved.exifRotation) || 0,
                            seq: String(saved.seq ?? index + 1), date: saved.date || '', time: saved.time || '',
                            location: saved.location || '', desc: saved.desc || '', selected: Boolean(saved.selected),
                            stageX: Number.isFinite(saved.stageX) ? saved.stageX : null,
                            stageY: Number.isFinite(saved.stageY) ? saved.stageY : null
                        });
                        this.setProgress('正在還原專案照片…', index + 1, manifest.photos.length);
                        if ((index + 1) % 5 === 0) await this.yieldToBrowser();
                    }
                    this.photos.forEach(photo => this.releasePhotoUrls(photo));
                    this.history = [];
                    this.historyIndex = -1;
                    this.releaseUnusedPhotoUrls();
                    this.photos = restored;
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
                    this.recordHistory('開啟專案');
                    this.markProjectSaved();
                    this.clearProgress();
                    this.render();
                    alert(`已開啟工作進度，共 ${restored.length} 張照片。`);
                } catch (error) {
                    this.clearProgress();
                    this.handleProjectError(error);
                }
            }

            handleProjectError(error) {
                console.error(error);
                alert(`工作進度檔無法開啟：${error?.message || error}`);
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
                icon.className = this.editorCollapsed ? 'fa-solid fa-pen-to-square' : 'fa-solid fa-panel-right';
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
                // 縮圖框固定為正方形：所有卡片等高、每一排對齊；照片以 contain 完整顯示（不裁切）。
                // 正方形旋轉 90° 後仍是同一個方框，因此只需旋轉，不必依照片比例換算寬高。
                const imageStyle = `width:100%;height:100%;transform:translate(-50%, -50%) rotate(${photo.rotation}deg);`;
                return { safeAspect: 1, imageStyle };
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
                this.render();
                this.recordHistory('依檔名整齊排序');
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

            removePhoto(index) {
                if (!this.photos[index]) return;
                this.photos.splice(index, 1);
                if (this.currentIndex >= this.photos.length) {
                    this.currentIndex = Math.max(0, this.photos.length - 1);
                }
                this.render();
                this.recordHistory('刪除照片');
            }

            removeSelectedPhotos() {
                const selectedIndices = this.getSelectedIndicesOrCurrent();
                if (!selectedIndices.length) return;
                const noun = selectedIndices.length === 1 ? '這張照片' : `已選取的 ${selectedIndices.length} 張照片`;
                if (!confirm(`要移除${noun}嗎？（不會刪除電腦裡的原檔，可按 Ctrl＋Z 復原）`)) return;
                const selectedSet = new Set(selectedIndices);
                this.photos = this.photos.filter((_, index) => !selectedSet.has(index));
                this.currentIndex = Math.max(0, Math.min(this.currentIndex, this.photos.length - 1));
                this.lastSelectedIndex = this.photos.length ? this.currentIndex : null;
                this.render();
                this.recordHistory('移除選取照片');
            }

            rotatePhoto(index) {
                this.photos[index].rotation = (this.photos[index].rotation + 90) % 360;
                this.render();
                this.recordHistory('旋轉照片');
            }

            clearAll() {
                if (!this.photos.length) return;
                if (confirm('要清空清冊中的全部照片嗎？（不會刪除電腦裡的原檔）')) {
                    this.photos = [];
                    this.currentIndex = 0;
                    this.render();
                    this.recordHistory('清空照片');
                }
            }

            releasePhotoUrls(photo) {
                [photo.src, photo.thumbnailSrc].filter(Boolean).forEach(url => {
                    URL.revokeObjectURL(url);
                    this.managedUrls.delete(url);
                });
            }

            releaseUnusedPhotoUrls() {
                const retained = new Set();
                [...this.photos, ...this.history.flatMap(state => state.photos)].forEach(photo => {
                    if (photo.src) retained.add(photo.src);
                    if (photo.thumbnailSrc) retained.add(photo.thumbnailSrc);
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
                if (targetIndex < 0 || this.draggedIndices.includes(targetIndex)) {
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

                // 整組同步移動策略：
                // 1. 分離選取組 (moving) 與非選取組 (remaining)
                // 2. 找出選取組在 remaining 中的「錨點」(最小選取索引對應在 remaining 的位置)
                // 3. 以 delta 調整插入位置後重組陣列
                const selectedSet = new Set(selectedIndices);
                const moving = this.photos.filter((_, i) => selectedSet.has(i));
                const remaining = this.photos.filter((_, i) => !selectedSet.has(i));

                // 錨點：選取組在原陣列中最小的索引，delta 決定在 remaining 中偏移幾格
                const minSelectedIdx = Math.min(...selectedIndices);
                // 計算 minSelectedIdx 在 remaining 中的對應位置
                let anchor = 0;
                for (let i = 0; i < this.photos.length; i++) {
                    if (i === minSelectedIdx) break;
                    if (!selectedSet.has(i)) anchor++;
                }
                const insertAt = Math.max(0, Math.min(anchor + delta, remaining.length));
                remaining.splice(insertAt, 0, ...moving);
                this.photos = remaining;
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
                const selectedSet = new Set(selectedIndices);
                const moving = this.photos.filter((_, index) => selectedSet.has(index));
                const remaining = this.photos.filter((_, index) => !selectedSet.has(index));
                this.photos = boundary === 'start' ? [...moving, ...remaining] : [...remaining, ...moving];
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
                if (this.draggedIndices.includes(index)) return;

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

                if (this.draggedIndices.includes(targetIndex)) {
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
                for (let nextIndex = index + 1; nextIndex < this.photos.length; nextIndex++) {
                    if (!this.draggedIndices.includes(nextIndex)) return this.photos[nextIndex].uid;
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

                const movingItems = this.draggedIndices.map(idx => this.photos[idx]);
                const remaining = this.photos.filter((_, idx) => !this.draggedIndices.includes(idx));
                let insertPos = this.dragInsertBeforeUid
                    ? remaining.findIndex(photo => photo.uid === this.dragInsertBeforeUid)
                    : remaining.length;
                if (insertPos < 0) insertPos = remaining.length;

                remaining.splice(insertPos, 0, ...movingItems);
                this.photos = remaining;

                this.currentIndex = insertPos;

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
                // 右側上一張／下一張與 Alt＋方向鍵完全沿用舞台方向鍵流程：
                // 單選目前照片、更新藍框、同步右側欄位並捲動縮圖至可見位置。
                this.selectPhotoByKeyboard(delta);
            }

            updateCurrentPhoto(field, value) {
                if (!this.photos[this.currentIndex]) return;
                this.photos[this.currentIndex][field] = value;
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
                let missingLocation = 0, missingDesc = 0, invalidDateTime = 0, duplicatePhotos = 0;
                this.photos.forEach(photo => {
                    if (!photo.location?.trim() && !globalLocation) missingLocation++;
                    if (!photo.desc?.trim()) missingDesc++;
                    const d = photo.date?.trim() || '';
                    const t = photo.time?.trim() || '';
                    const dateInvalid = d ? !this.isValidMinguoDate(d) : false;
                    const timeInvalid = !this.isValidTimeFormat(t);
                    if (dateInvalid || timeInvalid) invalidDateTime++;
                    if (this._dupSet.has(photo.name)) duplicatePhotos++;
                });
                const issuesCount = missingLocation + missingDesc + invalidDateTime + duplicatePhotos;
                let firstIssueFilter = null;
                if (missingDesc > 0) firstIssueFilter = 'missingDesc';
                else if (missingLocation > 0) firstIssueFilter = 'missingLocation';
                else if (invalidDateTime > 0) firstIssueFilter = 'invalidDateTime';
                else if (duplicatePhotos > 0) firstIssueFilter = 'duplicatePhotos';
                return {
                    total: this.photos.length,
                    missingLocation,
                    missingDesc,
                    invalidDateTime,
                    duplicatePhotos,
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
                this.btnToggleAudit.title = '檢查未填地點、未填說明、日期時間異常與同名照片；同一張照片可能有多項問題';
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
                            duplicatePhotos: this.auditBadge_duplicatePhotos
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
                const filterIds = {
                    'all': 'filterBtn_all',
                    'missingLocation': 'filterBtn_missingLocation',
                    'missingDesc': 'filterBtn_missingDesc',
                    'invalidDateTime': 'filterBtn_invalidDateTime',
                    'duplicatePhotos': 'filterBtn_duplicatePhotos',
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
                const parts = this.photos.map(photo => String(photo.seq || '').match(/^(.*?)(\d+)$/));
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

                if (totalPhotos === 0) {
                    this.emptyState.classList.remove('hidden');
                    this.workspaceContainer.classList.add('hidden');
                    this.gridContainer.classList.add('hidden');
                    this.previewContainer.classList.add('hidden');
                    return;
                }

                this.emptyState.classList.add('hidden');
                this.workspaceContainer.classList.toggle('hidden', this.currentView !== 'workspace');
                this.gridContainer.classList.toggle('hidden', this.currentView !== 'grid');
                this.previewContainer.classList.toggle('hidden', this.currentView !== 'preview');

                if (this.currentIndex >= totalPhotos) this.currentIndex = totalPhotos - 1;
                if (this.currentIndex < 0) this.currentIndex = 0;

                const curPhoto = this.photos[this.currentIndex];

                // 更新右側編輯器資料
                if (curPhoto) {
                    this.editorTitle.replaceChildren();
                    const titleIcon = document.createElement('i');
                    titleIcon.className = 'fa-solid fa-pen-to-square text-[#5B5CE2]';
                    const titleText = document.createElement('span');
                    titleText.textContent = `正在編輯第 ${this.currentIndex + 1}／${totalPhotos} 張（編號 ${curPhoto.seq}）`;
                    this.editorTitle.append(titleIcon, document.createTextNode(' '), titleText);
                    this.editSeq.value = curPhoto.seq || '';
                    this.editDate.value = curPhoto.date || '';
                    this.editTime.value = curPhoto.time || '';
                    this.editLocation.value = curPhoto.location || '';
                    this.editDesc.value = curPhoto.desc || '';
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
                }

                // 渲染左側縮圖卡片
                this.thumbGrid.innerHTML = '<div id="selectionMarquee" class="selection-marquee hidden" aria-hidden="true"></div>' + visibleEntries.map(({ photo: p, index: idx }) => {
                    const isSelected = p.selected ? 'selected' : '';
                    const isCurrent = idx === this.currentIndex ? 'active-editing' : '';
                    const missingFields = [!String(p.location || '').trim() && !defaultLocation ? '地點' : '', !String(p.desc || '').trim() ? '說明' : ''].filter(Boolean);
                    const statusText = missingFields.length ? `待補：${missingFields.join('、')}` : '資料完整';
                    const displayLayout = this.getPhotoDisplayLayout(p);
                    const cardTitle = `第 ${p.seq} 張｜時間：${p.time || '未填'}｜${statusText}｜雙擊縮圖放大、雙擊編號列編輯，拖曳可排序`;

                    return `
                        <div class="photo-thumb-card morandi-card p-1.5 flex flex-col gap-1 bg-white border border-[#C7CBD3]/25 rounded-xl ${isSelected} ${isCurrent}"
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
                                <span class="text-[13px] ${missingFields.length ? 'text-amber-700' : 'text-emerald-700'}" title="${escapeHtml(statusText)}">
                                    <i class="fa-solid ${missingFields.length ? 'fa-circle-exclamation' : 'fa-circle-check'}"></i>
                                </span>
                            </div>

                            <div class="thumbnail-stage w-full bg-[#EEF0F3] rounded-lg overflow-hidden relative cursor-zoom-in"
                                 title="雙擊放大檢視"
                                 ondblclick="event.stopPropagation(); window.app.openLightboxByIndex(${idx})">
                                <img src="${p.thumbnailSrc}" alt="${escapeHtml(p.name)}" loading="lazy" decoding="async" style="${displayLayout.imageStyle}" class="absolute left-1/2 top-1/2 object-contain pointer-events-none">
                            </div>

                        </div>
                    `;
                }).join('');
                this.selectionMarquee = document.getElementById('selectionMarquee');
                this.thumbGrid.style.removeProperty('min-height');

                // 渲染試算表視圖 (DataGrid)
                this.gridTbody.innerHTML = this.photos.map((p, idx) => `
                    <tr class="hover:bg-[#F7F7F5]/70 ${idx === this.currentIndex ? 'bg-blue-50/50' : ''}"
                        ondblclick="window.app.handleCardDblClick(${idx}, event)">
                        <td class="p-2 text-center">
                            <input type="checkbox" ${p.selected ? 'checked' : ''} onclick="window.app.togglePhotoSelect(${idx}, event)" class="rounded text-[#5B5CE2] cursor-pointer">
                        </td>
                        <td class="p-2 text-center">
                            <img src="${p.thumbnailSrc}" style="transform: rotate(${p.rotation}deg);" class="w-10 h-10 object-contain mx-auto rounded border bg-[#EEF0F3] cursor-pointer" onclick="event.stopPropagation(); window.app.openLightboxByIndex(${idx})" title="點擊放大">
                        </td>
                        <td class="p-2">
                            <input type="text" value="${escapeHtml(p.seq)}" oninput="window.app.updatePhotoField(${idx}, 'seq', this.value)" class="input-field w-full text-xs font-bold text-[#5B5CE2]">
                        </td>
                        <td class="p-2">
                            <input type="text" value="${escapeHtml(p.date || '')}" onpointerdown="window.app.handleMaskedPointerDown(event, this)" onfocus="window.app.beginMaskedEntry(this)" oninput="this.value = window.app.sanitizeMaskedEntry(this, 7); window.app.markMaskedEntryChanged(this); window.app.updatePhotoField(${idx}, 'date', this.value)" onblur="this.value = window.app.finishMaskedEntry(this, window.app.formatDateMask); window.app.updatePhotoField(${idx}, 'date', this.value)" class="input-field w-full text-xs font-mono" placeholder="留空＝清冊預設">
                        </td>
                        <td class="p-2">
                            <input type="text" value="${escapeHtml(p.time)}" onpointerdown="window.app.handleMaskedPointerDown(event, this)" onfocus="window.app.beginMaskedEntry(this)" oninput="this.value = window.app.sanitizeMaskedEntry(this, 6); window.app.markMaskedEntryChanged(this); window.app.updatePhotoField(${idx}, 'time', this.value)" onblur="this.value = window.app.finishMaskedEntry(this, window.app.formatTimeMask); window.app.updatePhotoField(${idx}, 'time', this.value)" class="input-field w-full text-xs font-mono" placeholder="輸入 4 或 6 碼">
                        </td>
                        <td class="p-2">
                            <input type="text" value="${escapeHtml(p.location)}" oninput="window.app.updatePhotoField(${idx}, 'location', this.value)" class="input-field w-full text-xs" placeholder="留空＝清冊預設">
                        </td>
                        <td class="p-2">
                            <input type="text" value="${escapeHtml(p.desc)}" oninput="window.app.updatePhotoField(${idx}, 'desc', this.value)" class="input-field w-full text-xs" placeholder="請填寫現場跡證或狀況說明">
                        </td>
                        <td class="p-2 text-[#6B7280] truncate max-w-[110px]" title="${escapeHtml(p.name)}">
                            ${escapeHtml(p.name)}
                        </td>
                        <td class="p-2 text-center">
                            <button type="button" onclick="window.app.rotatePhoto(${idx})" title="旋轉" class="p-1 text-[#5B5CE2]"><i class="fa-solid fa-rotate-right"></i></button>
                            <button type="button" onclick="window.app.removePhoto(${idx})" title="刪除" class="p-1 text-red-500"><i class="fa-regular fa-trash-can"></i></button>
                        </td>
                    </tr>
                `).join('');
                this.renderPagePreview();
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
                                    <img src="${photo.thumbnailSrc}" alt="${escapeHtml(photo.name)}" style="transform: rotate(${photo.rotation}deg);" class="w-full h-full object-contain">
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
                if (!this.photos[index]) return;
                this.photos[index][field] = value;
                this.scheduleHistory('修改照片欄位');
            }

            initLightboxInteractions() {
                if (!this.lightboxViewport) return;
                this.lightboxViewport.addEventListener('wheel', (e) => this.handleLightboxWheel(e), { passive: false });
                this.lightboxViewport.addEventListener('pointerdown', (e) => this.handleLightboxPointerDown(e));
                this.lightboxViewport.addEventListener('mousedown', (e) => this.handleLightboxPointerDown(e));
                window.addEventListener('pointermove', (e) => this.handleLightboxPointerMove(e));
                window.addEventListener('mousemove', (e) => this.handleLightboxPointerMove(e));
                window.addEventListener('pointerup', (e) => this.handleLightboxPointerUp(e));
                window.addEventListener('mouseup', (e) => this.handleLightboxPointerUp(e));
                window.addEventListener('pointercancel', (e) => this.handleLightboxPointerUp(e));
            }

            openLightboxByIndex(index) {
                if (!this.photos[index]) return;
                this.commitLightboxHistoryIfNeeded();
                this.lightboxIndex = index;
                this.updateLightbox();
                this.openModal('imageLightboxModal');
            }

            navigateLightbox(delta) {
                if (!this.photos.length || this.lightboxIndex === null) return;
                this.commitLightboxHistoryIfNeeded();
                const newIndex = Math.max(0, Math.min(this.photos.length - 1, this.lightboxIndex + delta));
                if (newIndex === this.lightboxIndex) return;
                this.lightboxIndex = newIndex;
                this.updateLightbox();
            }

            updateLightbox() {
                const photo = this.photos[this.lightboxIndex];
                if (!photo) return;
                this.resetLightboxZoom();

                if (this.lightboxImg) {
                    this.lightboxImg.style.transform = `rotate(${photo.rotation || 0}deg)`;
                    // 平滑降級：若已具備縮圖先以縮圖即時呈現，避免空白；背景非同步載入原圖無縫置換
                    this.lightboxImg.src = photo.thumbnailSrc || photo.src;
                    if (photo.src && photo.src !== photo.thumbnailSrc) {
                        const targetIndex = this.lightboxIndex;
                        const fullImg = new Image();
                        fullImg.onload = () => {
                            if (this.lightboxIndex === targetIndex && this.lightboxImg) {
                                this.lightboxImg.src = photo.src;
                            }
                        };
                        fullImg.src = photo.src;
                    }
                }

                if (this.lightboxCaption) {
                    this.lightboxCaption.textContent = `#${photo.seq}　${photo.name}`;
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
                const photo = this.photos[this.lightboxIndex];
                photo.rotation = ((photo.rotation || 0) + deltaDegrees) % 360;
                if (photo.rotation < 0) photo.rotation += 360;
                if (this.lightboxImg) {
                    this.lightboxImg.style.transform = `rotate(${photo.rotation}deg)`;
                }
                this.syncPhotoRotationToCanvas(this.lightboxIndex, photo.rotation);
                this.lightboxHasRotated = true;
            }

            syncPhotoRotationToCanvas(index, rotation) {
                const photo = this.photos[index];
                if (!photo) return;
                // 更新主畫布卡片縮圖視覺
                const thumbCard = document.getElementById(`thumb_${photo.uid}`);
                if (thumbCard) {
                    const img = thumbCard.querySelector('.thumbnail-stage img');
                    if (img) {
                        const layout = this.getPhotoDisplayLayout(photo);
                        img.style.cssText = layout.imageStyle;
                    }
                }
                // 更新試算表 DataGrid 縮圖
                if (this.gridTbody && this.gridTbody.children[index]) {
                    const row = this.gridTbody.children[index];
                    const rowImg = row.querySelector('td:nth-child(2) img');
                    if (rowImg) {
                        rowImg.style.transform = `rotate(${rotation}deg)`;
                    }
                }
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
                    this.commitLightboxHistoryIfNeeded();
                }
                if (typeof PhotoReportModalUi !== 'undefined') {
                    PhotoReportModalUi.closeModal(id);
                    return;
                }
                const el = typeof id === 'string' ? document.getElementById(id) : id;
                if (el) el.classList.add('hidden');
            }

            async processImageToBlob(photoItem, maxDimension = 1600, quality = 0.88) {
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
                        const rot = (photoItem.rotation + exifRotation) % 360;

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
                        ctx.rotate((rot * Math.PI) / 180);
                        ctx.drawImage(source, -w / 2, -h / 2, w, h);
                        ctx.restore();

                        canvas.toBlob((blob) => {
                            if (!blob) return reject(new Error(`無法轉換圖片：${photoItem.name}`));
                            resolve({
                                blob: blob,
                                width: canvas.width,
                                height: canvas.height
                            });
                        }, 'image/jpeg', quality);
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
