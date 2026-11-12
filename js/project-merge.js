(function(root,factory){
    if(typeof module==='object'&&module.exports)module.exports=factory(require('./photo-model'));
    else root.PhotoReportProjectMerge=factory(root.PhotoReportPhotoModel);
})(typeof globalThis!=='undefined'?globalThis:this,function(model){
    'use strict';
    const REPORT_FIELDS = ['seq','date','time','timeAutoDefaulted','location','desc','rotation','flipH','parentUid','cropRect','sideCrop','annotations'];
    function differences(before,after,keys){
        return keys.filter(key=>JSON.stringify(before?.[key]??null)!==JSON.stringify(after?.[key]??null))
            .map(field=>({field,before:JSON.parse(JSON.stringify(before?.[field]??null)),after:JSON.parse(JSON.stringify(after?.[field]??null))}));
    }
    function buildReport({beforePhotos,afterPhotos,beforeCase,afterCase,sources,createdAt=new Date().toISOString()}){
        const beforeByUid=new Map(beforePhotos.map(p=>[p.uid,p]));
        const photos=afterPhotos.flatMap(photo=>{
            const before=beforeByUid.get(photo.uid),after=model.toRecord(photo);
            const changes=differences(before,after,REPORT_FIELDS);
            return !before||changes.length ? [{uid:photo.uid,name:photo.name,seq:photo.seq,kind:before?'updated':'added',changes}] : [];
        });
        return {createdAt,sources:[...sources],beforeCount:beforePhotos.length,afterCount:afterPhotos.length,
            caseChanges:differences(beforeCase,afterCase,['caseTitle','caseDate','defaultLocation','deptName','officerName','layoutSelect']),photos};
    }
    return {
    buildReport,
    cloneAnnotations:model.cloneAnnotations, cloneSideCrop:model.cloneSideCrop,
    getValidParentPhoto(photo, photos = []) {
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
    },
    isLinkedCropChild(photo, photos = []) {
        return Boolean(this.getValidParentPhoto(photo, photos));
    },
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
    },

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
    },

    isCropRectSimilar(rectA, rectB) {
        if (!rectA || !rectB) return false;
        return Math.abs((Number(rectA.x) || 0) - (Number(rectB.x) || 0)) <= 0.025 &&
            Math.abs((Number(rectA.y) || 0) - (Number(rectB.y) || 0)) <= 0.025 &&
            Math.abs((Number(rectA.w) || 0) - (Number(rectB.w) || 0)) <= 0.025 &&
            Math.abs((Number(rectA.h) || 0) - (Number(rectB.h) || 0)) <= 0.025;
    },

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
    },

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
    },

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
    },

    normalizePhotoFilenameKey(name) {
        const trimmed = String(name || '').trim().toLowerCase();
        return trimmed.replace(/\.(jpe?g|png|webp|gif|bmp|heic)$/i, '').replace(/\s+/g, ' ');
    }
    };
});
