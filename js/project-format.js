(function(root,factory){
    if(typeof module==='object'&&module.exports) module.exports=factory(require('./photo-model'));
    else root.PhotoReportProjectFormat=factory(root.PhotoReportPhotoModel);
})(typeof globalThis!=='undefined'?globalThis:this,function(model){
    'use strict';
    const FORMAT='photo-report-project', VERSION=1;
    const CASE_FIELDS=['caseTitle','caseDate','defaultLocation','deptName','officerName','layoutSelect'];
    const TEXT_FIELDS=['name','uid','type','visualFingerprint','parentUid','date','time','location','desc'];
    function object(value){return value&&typeof value==='object'&&!Array.isArray(value);}
    function validateManifest(value,{draft=false}={}){
        if(!object(value)||value.format!==(draft?'photo-report-autodraft':FORMAT)||!Array.isArray(value.photos)) throw new Error('進度檔格式不支援或照片清單不完整。');
        if(value.version!==undefined&&value.version!==VERSION) throw new Error(`進度檔版本 ${value.version} 尚未支援，請使用相容版本的程式開啟。`);
        if(value.caseData!=null&&!object(value.caseData))throw new Error('案件資料格式不正確。');
        for(const key of CASE_FIELDS){if(value.caseData?.[key]!=null&&typeof value.caseData[key]!=='string')throw new Error(`案件欄位「${key}」格式不正確。`);}
        for(const key of ['currentIndex','lastSelectedIndex'])if(value[key]!=null&&!Number.isInteger(Number(value[key])))throw new Error('進度檔的照片選取位置不正確。');
        if(value.savedAt!=null&&typeof value.savedAt!=='string')throw new Error('進度檔的存檔時間格式不正確。');
        if(value.caseData?.layoutSelect&&!['up_down_2','left_right_2','landscape_3'].includes(value.caseData.layoutSelect))throw new Error('進度檔的清冊版型尚未支援。');
        const archiveNames=new Set(),uids=new Set();
        for(let index=0;index<value.photos.length;index++){
            const photo=value.photos[index],label=`第 ${index+1} 張照片`;
            if(!object(photo))throw new Error(`${label}資料不完整。`);
            for(const key of TEXT_FIELDS)if(photo[key]!=null&&typeof photo[key]!=='string')throw new Error(`${label}的「${key}」必須為文字。`);
            if(photo.seq!=null&&!['string','number'].includes(typeof photo.seq))throw new Error(`${label}的編號格式不正確。`);
            if(photo.uid){if(uids.has(photo.uid))throw new Error(`${label}的照片識別碼重複。`);uids.add(photo.uid);}
            if(!draft){
                if(typeof photo.archiveName!=='string'||!photo.archiveName||/[\\/]/.test(photo.archiveName)||photo.archiveName==='..')throw new Error(`${label}的圖片清單路徑不正確。`);
                if(archiveNames.has(photo.archiveName))throw new Error(`${label}的圖片清單重複。`);
                archiveNames.add(photo.archiveName);
            }else if(!photo.buffer&&!photo.blob)throw new Error(`${label}的草稿圖片資料缺少。`);
            if(photo.sideCrop!=null&&(!object(photo.sideCrop)||!object(photo.sideCrop.cropRect)))throw new Error(`${label}的並排特寫資料不正確。`);
            if(photo.annotations!=null&&(!Array.isArray(photo.annotations)||photo.annotations.some(a=>!object(a))))throw new Error(`${label}的標記資料格式不正確。`);
            for(const ann of photo.annotations||[]){
                if(['x','y','w','h','scale'].some(key=>ann[key]!=null&&!Number.isFinite(Number(ann[key]))))throw new Error(`${label}的標記位置不正確。`);
                if(ann.text!=null&&typeof ann.text!=='string')throw new Error(`${label}的標籤文字不正確。`);
            }
            if(photo.sideCrop&&['widthRatio','offsetY'].some(key=>photo.sideCrop[key]!=null&&!Number.isFinite(Number(photo.sideCrop[key]))))throw new Error(`${label}的並排特寫位置不正確。`);
            for(const rect of [photo.cropRect,photo.sideCrop?.cropRect])if(rect!=null&&(!object(rect)||['x','y','w','h'].some(k=>rect[k]!=null&&!Number.isFinite(Number(rect[k])))))throw new Error(`${label}的裁切範圍不正確。`);
            for(const key of ['rotation','exifRotation','stageX','stageY','thumbnailAspect'])if(photo[key]!=null&&!Number.isFinite(Number(photo[key])))throw new Error(`${label}的「${key}」不是有效數值。`);
        }
        return value;
    }
    function createManifest({photos,caseData,currentIndex,lastSelectedIndex,savedAt=new Date().toISOString(),liteMode=true,draft=false}){
        return {format:draft?'photo-report-autodraft':FORMAT,version:VERSION,...(!draft?{liteMode}:{}),savedAt,caseData,currentIndex,lastSelectedIndex,photos};
    }
    async function readArchive(file,{Zip,decodeImage,checkCancelled=()=>{},onProgress=()=>{},yieldToBrowser=async()=>{}}){
        let zip;
        try{zip=await Zip.loadAsync(file,{checkCRC32:true});}catch(error){throw new Error('進度檔壓縮內容損壞或不是有效的 ZIP 檔案。',{cause:error});}
        checkCancelled();
        const manifestEntry=zip.file('manifest.json');if(!manifestEntry)throw new Error('這不是有效的照片清冊進度檔（缺少 manifest.json）。');
        let manifest;
        try{manifest=JSON.parse(await manifestEntry.async('text'));}catch(error){throw new Error('進度檔的 manifest.json 無法讀取。',{cause:error});}
        validateManifest(manifest);
        const blobs=new Map();
        for(let i=0;i<manifest.photos.length;i++){
            checkCancelled();const saved=manifest.photos[i],entry=zip.file(`images/${saved.archiveName}`);
            if(!entry)throw new Error(`進度檔缺少照片：${saved.name||i+1}`);
            const blob=await entry.async('blob');checkCancelled();
            if(!blob.size)throw new Error(`進度檔照片是空檔案：${saved.name||i+1}`);
            try{await decodeImage(blob);}catch(error){throw new Error(`進度檔照片無法解碼：${saved.name||i+1}`,{cause:error});}
            checkCancelled();blobs.set(saved.archiveName,blob);onProgress(i+1,manifest.photos.length);
            if((i+1)%5===0)await yieldToBrowser();
        }
        return{zip,manifest,blobs};
    }
    return{FORMAT,VERSION,CASE_FIELDS,validateManifest,createManifest,readArchive,toRecord:model.toRecord,restore:model.restore};
});
