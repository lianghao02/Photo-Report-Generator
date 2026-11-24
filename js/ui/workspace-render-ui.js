/**
 * workspace-render-ui.js - 工作台縮圖卡片、試算表 (DataGrid) 渲染與全站事件委派模組
 * 徹底移除靜態 HTML 與動態樣板字串中的行內 on* 事件屬性，讓桌面版可移除 script-src-attr 'unsafe-inline'。
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        module.exports = factory();
    } else {
        root.PhotoReportWorkspaceRenderUi = factory();
    }
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
    'use strict';

    function escapeHtml(value) {
        return String(value ?? '').replace(/[&<>"']/g, char => ({
            '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
        })[char]);
    }

    function renderThumbGrid(app, visibleEntries, defaultLocation) {
        if (!app.thumbGrid) return;
        app.thumbGrid.innerHTML = '<div id="selectionMarquee" class="selection-marquee hidden" aria-hidden="true"></div>' + visibleEntries.map(({ photo: p, index: idx }) => {
            const isSelected = p.selected ? 'selected' : '';
            const isCurrent = idx === app.currentIndex ? 'active-editing' : '';
            const isCropChild = app.isLinkedCropChild(p) ? 'photo-thumb-card--crop-child' : '';
            const missingFields = [!String(p.location || '').trim() && !defaultLocation ? '地點' : '', !String(p.desc || '').trim() ? '說明' : ''].filter(Boolean);
            const overLimit = app.inspectPhotoTextLimits(String(p.location || '').trim() || defaultLocation, String(p.desc || '').trim()).isOver;
            const statusIssues = [...missingFields, ...(overLimit ? ['字數過長'] : [])];
            const statusText = statusIssues.length ? `注意：${statusIssues.join('、')}` : '資料完整';
            const displayLayout = app.getPhotoDisplayLayout(p);
            const cardTitle = `第 ${p.seq} 張｜時間：${p.time || '未填'}｜${statusText}｜雙擊縮圖放大、雙擊編號列編輯，拖曳可排序`;

            return `
                <div class="photo-thumb-card morandi-card p-1.5 flex flex-col gap-1 bg-white border border-[#C7CBD3]/25 rounded-xl ${isSelected} ${isCurrent} ${isCropChild}"
                     id="thumb_${escapeHtml(p.uid)}"
                     data-card-index="${idx}"
                     style="--thumbnail-aspect:${displayLayout.safeAspect};"
                     title="${escapeHtml(cardTitle)}"
                     draggable="false">

                    <div class="thumbnail-meta-row flex items-center justify-between min-h-5">
                        <div class="flex min-w-0 items-center gap-1.5">
                            <input type="checkbox" ${p.selected ? 'checked' : ''} data-card-action="toggle-select" data-card-index="${idx}" class="rounded text-blue-600 cursor-pointer" title="勾選">
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
                         data-card-action="open-lightbox"
                         data-card-index="${idx}"
                         title="雙擊放大檢視">
                        <img src="${displayLayout.src || p.thumbnailSrc}" alt="${escapeHtml(p.name)}" loading="lazy" decoding="async" style="${displayLayout.imageStyle}" class="absolute left-1/2 top-1/2 object-contain pointer-events-none">
                        ${app.getPhotoAnnotationBadgeHtml(p)}
                    </div>

                </div>
            `;
        }).join('');
        app.selectionMarquee = document.getElementById('selectionMarquee');
        app.thumbGrid.style.removeProperty('min-height');
    }

    function renderGridTable(app, visibleEntries) {
        if (!app.gridTbody) return;
        const gridLimits = app.getActiveLayoutTextLimits();
        app.gridTbody.innerHTML = visibleEntries.map(({ photo: p, index: idx }) => {
            const locCheck = app.inspectPhotoTextLimits(String(p.location || '').trim(), '');
            const descCheck = app.inspectPhotoTextLimits('', String(p.desc || '').trim());
            const parentPhoto = app.getValidParentPhoto(p);
            const seqCellHtml = parentPhoto
                ? `<input type="text" value="${escapeHtml(p.seq)}" readonly class="input-field w-full text-xs font-bold seq-locked" title="隨原圖 #${escapeHtml(parentPhoto.seq)} 自動連動鎖定"><div class="grid-crop-parent-hint text-[10px] text-sky-700 font-bold mt-0.5 whitespace-nowrap">↳ #${escapeHtml(parentPhoto.seq)} 特寫</div>`
                : `<input type="text" value="${escapeHtml(p.seq)}" data-grid-index="${idx}" data-grid-field="seq" class="input-field w-full text-xs font-bold text-[#5B5CE2]">`;
            const thumbSrc = (p.sideCrop && p.compositeThumbnailSrc) ? p.compositeThumbnailSrc : p.thumbnailSrc;
            const thumbTransform = (p.sideCrop && p.compositeThumbnailSrc) ? 'none' : `${p.flipH ? 'scaleX(-1) ' : ''}rotate(${p.rotation}deg)`;

            return `
            <tr class="hover:bg-[#F7F7F5]/70 ${idx === app.currentIndex ? 'bg-blue-50/50' : ''}" data-grid-row-index="${idx}">
                <td class="p-2 text-center">
                    <input type="checkbox" ${p.selected ? 'checked' : ''} data-grid-action="toggle-select" data-grid-index="${idx}" class="rounded text-[#5B5CE2] cursor-pointer">
                </td>
                <td class="p-2 text-center">
                    <img src="${thumbSrc}" style="transform: ${thumbTransform};" data-grid-action="open-lightbox" data-grid-index="${idx}" class="w-10 h-10 object-contain mx-auto rounded border bg-[#EEF0F3] cursor-pointer" title="點擊放大">
                </td>
                <td class="p-2">
                    ${seqCellHtml}
                </td>
                <td class="p-2">
                    <input type="text" value="${escapeHtml(p.date || '')}" data-grid-index="${idx}" data-grid-field="date" data-mask-type="date" class="input-field w-full text-xs font-mono" placeholder="留空＝清冊預設">
                </td>
                <td class="p-2">
                    <input type="text" value="${escapeHtml(p.time)}" data-grid-index="${idx}" data-grid-field="time" data-mask-type="time" class="input-field w-full text-xs font-mono" placeholder="輸入 4 或 6 碼">
                </td>
                <td class="p-2">
                    <input type="text" data-grid-index="${idx}" data-grid-field="location" value="${escapeHtml(p.location)}" class="input-field w-full text-xs ${locCheck.locationOver ? 'input-over-limit' : ''}" placeholder="留空＝清冊預設" title="地點最多 ${gridLimits.locationMaxFullWidth} 個全形（2 半形＝1 全形）">
                </td>
                <td class="p-2">
                    <input type="text" data-grid-index="${idx}" data-grid-field="desc" value="${escapeHtml(p.desc)}" class="input-field w-full text-xs ${descCheck.descOver ? 'input-over-limit' : ''}" placeholder="請填寫現場跡證或狀況說明" title="說明最多 ${gridLimits.descMaxFullWidth} 個全形・最多 ${gridLimits.descMaxLines} 行（單行約 ${gridLimits.descLineFullWidth} 全形，2 半形＝1 全形）">
                </td>
                <td class="p-2 text-[#6B7280] truncate max-w-[110px]" title="${escapeHtml(p.name)}">
                    ${escapeHtml(p.name)}
                </td>
                <td class="p-2 text-center">
                    <button type="button" data-grid-action="rotate" data-grid-index="${idx}" title="旋轉" class="p-1 text-[#5B5CE2]"><i class="fa-solid fa-rotate-right pointer-events-none"></i></button>
                    <button type="button" data-grid-action="remove" data-grid-index="${idx}" title="刪除" class="p-1 text-red-500"><i class="fa-regular fa-trash-can pointer-events-none"></i></button>
                </td>
            </tr>
            `;
        }).join('');
        app.gridTbody.querySelectorAll('input[data-grid-field="location"], input[data-grid-field="desc"]').forEach(gridInput => {
            app.syncTextLimitInputState(gridInput, gridInput.value || '');
        });
    }

    function bindDelegatedWorkspaceEvents(app) {
        if (app.thumbGrid && !app.thumbGrid.dataset.delegatedBound) {
            app.thumbGrid.dataset.delegatedBound = '1';
            app.thumbGrid.addEventListener('click', (event) => {
                const checkbox = event.target.closest('input[data-card-action="toggle-select"]');
                if (checkbox) {
                    const idx = Number(checkbox.dataset.cardIndex);
                    if (Number.isInteger(idx)) app.togglePhotoSelect(idx, event);
                    return;
                }
                const card = event.target.closest('.photo-thumb-card[data-card-index]');
                if (card) {
                    const idx = Number(card.dataset.cardIndex);
                    if (Number.isInteger(idx)) app.handleCardClick(idx, event);
                }
            });
            app.thumbGrid.addEventListener('dblclick', (event) => {
                const stage = event.target.closest('[data-card-action="open-lightbox"]');
                if (stage) {
                    event.stopPropagation();
                    const idx = Number(stage.dataset.cardIndex);
                    if (Number.isInteger(idx)) app.openLightboxByIndex(idx);
                    return;
                }
                const card = event.target.closest('.photo-thumb-card[data-card-index]');
                if (card) {
                    const idx = Number(card.dataset.cardIndex);
                    if (Number.isInteger(idx)) app.handleCardDblClick(idx, event);
                }
            });
            app.thumbGrid.addEventListener('pointerdown', (event) => {
                const card = event.target.closest('.photo-thumb-card[data-card-index]');
                if (card) {
                    const idx = Number(card.dataset.cardIndex);
                    if (Number.isInteger(idx)) app.beginCardPointerDrag(idx, event);
                }
            });
        }

        if (app.gridTbody && !app.gridTbody.dataset.delegatedBound) {
            app.gridTbody.dataset.delegatedBound = '1';
            app.gridTbody.addEventListener('click', (event) => {
                const actionEl = event.target.closest('[data-grid-action]');
                if (!actionEl) return;
                const action = actionEl.dataset.gridAction;
                const idx = Number(actionEl.dataset.gridIndex);
                if (!Number.isInteger(idx)) return;
                if (action === 'toggle-select') {
                    app.togglePhotoSelect(idx, event);
                } else if (action === 'open-lightbox') {
                    event.stopPropagation();
                    app.openLightboxByIndex(idx);
                } else if (action === 'rotate') {
                    app.rotatePhoto(idx);
                } else if (action === 'remove') {
                    app.removePhoto(idx);
                }
            });
            app.gridTbody.addEventListener('dblclick', (event) => {
                const row = event.target.closest('tr[data-grid-row-index]');
                if (!row) return;
                const idx = Number(row.dataset.gridRowIndex);
                if (Number.isInteger(idx)) app.handleCardDblClick(idx, event);
            });
            app.gridTbody.addEventListener('pointerdown', (event) => {
                const input = event.target.closest('input[data-mask-type]');
                if (input) app.handleMaskedPointerDown(event, input);
            });
            app.gridTbody.addEventListener('focusin', (event) => {
                const input = event.target.closest('input[data-mask-type]');
                if (input) app.beginMaskedEntry(input);
            });
            app.gridTbody.addEventListener('input', (event) => {
                const target = event.target;
                if (!(target instanceof HTMLInputElement)) return;
                const field = target.dataset.gridField;
                const idx = Number(target.dataset.gridIndex);
                if (!Number.isInteger(idx)) return;
                if (field === 'seq') {
                    app.updatePhotoField(idx, 'seq', target.value);
                } else if (field === 'date') {
                    target.value = app.sanitizeMaskedEntry(target, 7);
                    app.markMaskedEntryChanged(target);
                    app.updatePhotoField(idx, 'date', target.value);
                } else if (field === 'time') {
                    target.value = app.sanitizeMaskedEntry(target, 6);
                    app.markMaskedEntryChanged(target);
                    app.updatePhotoField(idx, 'time', target.value);
                }
            });
            app.gridTbody.addEventListener('focusout', (event) => {
                const target = event.target;
                if (!(target instanceof HTMLInputElement)) return;
                const field = target.dataset.gridField;
                const idx = Number(target.dataset.gridIndex);
                if (!Number.isInteger(idx)) return;
                if (field === 'date') {
                    target.value = app.finishMaskedEntry(target, app.formatDateMask);
                    app.updatePhotoField(idx, 'date', target.value);
                } else if (field === 'time') {
                    target.value = app.finishMaskedEntry(target, app.formatTimeMask);
                    app.updatePhotoField(idx, 'time', target.value);
                }
            });
        }
    }

    function bindStaticUiEvents(app) {
        if (document.body?.dataset.staticUiBound) return;
        if (document.body) document.body.dataset.staticUiBound = '1';

        document.getElementById('gridSelectAll')?.addEventListener('click', (event) => {
            app.toggleSelectAll(event.currentTarget.checked);
        });

        document.querySelectorAll('[data-fill-down-first]').forEach(btn => {
            btn.addEventListener('click', () => {
                const field = btn.dataset.fillDownFirst;
                if (field) app.fillDownFirst(field);
            });
        });

        // 統一綁定 Modal 背景點擊關閉與取消／關閉按鈕
        document.querySelectorAll('[data-modal-backdrop]').forEach(backdrop => {
            backdrop.addEventListener('click', (event) => {
                if (event.target !== backdrop) return;
                const modalId = backdrop.dataset.modalBackdrop;
                if (modalId) app.closeModal(modalId);
            });
        });

        document.querySelectorAll('[data-close-modal]').forEach(btn => {
            btn.addEventListener('click', () => {
                const modalId = btn.dataset.closeModal;
                if (modalId) app.closeModal(modalId);
            });
        });

        // 大圖燈箱背景與工具按鈕
        const lightboxBackdrop = document.getElementById('imageLightboxModal');
        lightboxBackdrop?.addEventListener('pointerdown', (event) => {
            if (event.target === lightboxBackdrop) app.handleLightboxBackdropPointerDown(event);
        });
        lightboxBackdrop?.addEventListener('click', (event) => {
            if (event.target === lightboxBackdrop) app.handleLightboxBackdropClick(event);
        });

        document.getElementById('btnLightboxResetZoom')?.addEventListener('click', () => app.resetLightboxZoom());
        document.getElementById('btnLightboxRotateCCW')?.addEventListener('click', () => app.rotateLightboxPhoto(-90));
        document.getElementById('btnLightboxRotateCW')?.addEventListener('click', () => app.rotateLightboxPhoto(90));
        document.getElementById('btnLightboxFlipH')?.addEventListener('click', () => app.toggleLightboxFlipH());
        document.getElementById('btnLightboxToggleEditor')?.addEventListener('click', () => app.toggleLightboxEditor());
        document.getElementById('btnLightboxClearAnnotations')?.addEventListener('click', () => app.clearCurrentPhotoAnnotations());
        document.getElementById('btnLightboxPrev')?.addEventListener('click', () => app.navigateLightbox(-1));
        document.getElementById('btnLightboxNext')?.addEventListener('click', () => app.navigateLightbox(1));

        document.querySelectorAll('[data-lightbox-tool]').forEach(btn => {
            btn.addEventListener('click', () => {
                const tool = btn.dataset.lightboxTool;
                if (tool) app.setLightboxTool(tool);
            });
        });

        document.querySelectorAll('[data-lightbox-layout]').forEach(btn => {
            btn.addEventListener('click', () => {
                const layout = btn.dataset.lightboxLayout;
                if (layout) app.setLightboxLayout(layout);
            });
        });
    }

    return {
        escapeHtml,
        renderThumbGrid,
        renderGridTable,
        bindDelegatedWorkspaceEvents,
        bindStaticUiEvents
    };
});
