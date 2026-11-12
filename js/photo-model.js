(function (root, factory) {
    if (typeof module === 'object' && module.exports) module.exports = factory();
    else root.PhotoReportPhotoModel = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
    'use strict';
    // 可保存資料不包含 File、Object URL、縮圖快取；執行期資源由宿主負責釋放。
    function cloneAnnotations(value) { return Array.isArray(value) ? value.map(item => ({ ...item })) : []; }
    function cloneSideCrop(value) {
        if (!value || typeof value !== 'object' || !value.cropRect) return null;
        return {
            cropRect: { x:Number(value.cropRect.x)||0, y:Number(value.cropRect.y)||0, w:Number(value.cropRect.w)||0.2, h:Number(value.cropRect.h)||0.2 },
            widthRatio:Number(value.widthRatio)||0.65,
            offsetY:Number.isFinite(Number(value.offsetY)) ? Number(value.offsetY) : 0.5,
            borderColor:value.borderColor==='red' ? 'red' : value.borderColor==='none' ? 'none' : 'yellow',
            sourceAnnId:value.sourceAnnId||null
        };
    }
    function toRecord(photo, overrides = {}) {
        return {
            uid:photo.uid, name:photo.name, type:photo.file?.type||photo.type||'image/jpeg',
            visualFingerprint:photo.visualFingerprint||'', thumbnailAspect:Number(photo.thumbnailAspect)||1,
            rotation:photo.rotation||0, flipH:Boolean(photo.flipH), parentUid:photo.parentUid||null,
            cropRect:photo.cropRect ? {...photo.cropRect} : null, sideCrop:cloneSideCrop(photo.sideCrop), annotations:cloneAnnotations(photo.annotations),
            exifRotation:photo.exifRotation||0, seq:photo.seq, date:photo.date||'', time:photo.time||'', timeAutoDefaulted:Boolean(photo.timeAutoDefaulted),
            location:photo.location||'', desc:photo.desc||'', selected:Boolean(photo.selected),
            stageX:Number.isFinite(photo.stageX)?photo.stageX:null, stageY:Number.isFinite(photo.stageY)?photo.stageY:null,
            ...overrides
        };
    }
    function restore(saved, index, runtime) {
        return {
            ...toRecord(saved), uid:saved.uid||runtime.uid, name:saved.name||runtime.name||`photo_${index+1}.jpg`,
            rotation:Number(saved.rotation)||0, exifRotation:Number(saved.exifRotation)||0, seq:String(saved.seq??index+1),
            cropRect:saved.cropRect&&typeof saved.cropRect==='object' ? {x:Number(saved.cropRect.x)||0,y:Number(saved.cropRect.y)||0,w:Number(saved.cropRect.w)||0.1,h:Number(saved.cropRect.h)||0.1}:null,
            ...runtime, thumbnailAspect:Number(saved.thumbnailAspect)||runtime.thumbnailAspect||1,
            visualFingerprint:saved.visualFingerprint||runtime.visualFingerprint||'', compositeThumbnailSrc:null, _compositeKey:''
        };
    }
    function cloneForHistory(photo) {
        return {...photo, flipH:Boolean(photo.flipH), parentUid:photo.parentUid||null, cropRect:photo.cropRect?{...photo.cropRect}:null,
            sideCrop:cloneSideCrop(photo.sideCrop),annotations:cloneAnnotations(photo.annotations),compositeThumbnailSrc:null,_compositeKey:''};
    }
    return {cloneAnnotations,cloneSideCrop,toRecord,restore,cloneForHistory};
});
