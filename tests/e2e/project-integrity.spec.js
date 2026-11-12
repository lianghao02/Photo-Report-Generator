const assert=require('node:assert/strict');
const fs=require('node:fs');const path=require('node:path');const {pathToFileURL}=require('node:url');
const JSZip=require('jszip');const {chromium}=require('playwright');const {trackPageErrors}=require('../helpers/page-errors');
async function main(){
    const browser=await chromium.launch({headless:true});
    try{
        const page=await browser.newPage({viewport:{width:1366,height:768}});const clean=trackPageErrors(page);
        await page.goto(pathToFileURL(path.resolve(__dirname,'../../index.html')).href);await page.waitForFunction(()=>!!window.app);
        await page.evaluate(()=>{window.testMessages=[];window.alert=m=>testMessages.push(m);window.confirm=()=>true;});
        const baseline=fs.readFileSync(path.resolve(__dirname,'../fixtures/project-v1.photo-report'));
        const open=buffer=>page.evaluate(async data=>window.app.openProjectFromFile(new File([new Uint8Array(data)],'相容進度.photo-report',{type:'application/zip'})),Array.from(buffer));
        const merge=buffer=>page.evaluate(async data=>window.app.mergeProjectFiles([new File([new Uint8Array(data)],'彙整測試.photo-report',{type:'application/zip'})]),Array.from(buffer));
        assert.equal(await open(baseline),true);
        const initial=await page.evaluate(()=>({photos:app.photos.map(p=>PhotoReportPhotoModel.toRecord(p)),signature:app.projectSignature(),urls:app.photos.map(p=>p.src)}));
        assert.equal(initial.photos.length,2);assert(initial.photos[0].annotations.length);assert(initial.photos[0].sideCrop);assert(initial.photos[0].flipH);assert.equal(initial.photos[1].parentUid,initial.photos[0].uid);
        const alter=async(fn)=>{const zip=await JSZip.loadAsync(baseline);const manifest=JSON.parse(await zip.file('manifest.json').async('string'));await fn(zip,manifest);zip.file('manifest.json',JSON.stringify(manifest));return zip.generateAsync({type:'nodebuffer'});};
        const noVersion=await alter((z,m)=>{delete m.version;});assert.equal(await open(noVersion),true,'省略version的舊格式仍可還原');
        await page.evaluate(()=>{window.temporaryUrls=new Set();const create=URL.createObjectURL.bind(URL),revoke=URL.revokeObjectURL.bind(URL);URL.createObjectURL=blob=>{const url=create(blob);temporaryUrls.add(url);return url;};URL.revokeObjectURL=url=>{temporaryUrls.delete(url);revoke(url);};});
        const snapshot=()=>page.evaluate(()=>({signature:app.projectSignature(),urls:app.photos.map(p=>p.src)}));const protectedState=await snapshot();
        for(const [name,mutate]of[
            ['新版本',(z,m)=>{m.version=99;}],
            ['缺圖',(z,m)=>z.remove('images/'+m.photos[0].archiveName)],
            ['壞圖',(z,m)=>z.file('images/'+m.photos[0].archiveName,'不是圖片')],
            ['壞欄位',(z,m)=>{m.photos[0].desc={bad:true};}],
            ['壞時間',(z,m)=>{m.savedAt={toString:{}};}],
            ['壞標記',(z,m)=>{m.photos[0].annotations[0].x='不是數值';}],
            ['壞版型',(z,m)=>{m.caseData.layoutSelect='unknown';}],
            ['重複UID',(z,m)=>{m.photos[1].uid=m.photos[0].uid;}]
        ]){const bad=await alter(mutate);assert.equal(await open(bad),false,name);assert.deepEqual(await snapshot(),protectedState,name+'不可破壞目前工作');assert.equal(await page.evaluate(()=>temporaryUrls.size),0,name+'不可遺留臨時URL');if(name==='壞圖'){assert.equal(await merge(bad),false,'匹配到既有照片也必須檢查圖片');assert.deepEqual(await snapshot(),protectedState);}}
        assert.equal(await open(Buffer.from('無效ZIP')),false);assert.deepEqual(await snapshot(),protectedState);
        const cancel=await page.evaluate(async data=>{const original=app.loadImage;app.loadImage=async function(url){const result=await original.call(this,url);this.processingCancelled=true;return result;};try{return await app.openProjectFromFile(new File([new Uint8Array(data)],'取消測試.photo-report'));}finally{app.loadImage=original;app.processingCancelled=false;}},Array.from(baseline));assert.equal(cancel,false);assert.deepEqual(await snapshot(),protectedState);assert.equal(await page.evaluate(()=>temporaryUrls.size),0);
        console.log('v1／省略版本相容、版本拒絕、缺圖／壞圖／壞欄位／重複UID、取消及原工作保護通過');
        const saveStates=await page.evaluate(async()=>{
            app.photos[0].desc='存檔狀態驗證';app.recordHistory('狀態驗證');let captured;const originalSave=window.saveAs;window.saveAs=blob=>{captured=blob;};await app.saveProject({lite:false});const download={mode:app.currentProjectSaveMode,text:document.getElementById('quickSaveStatusText').textContent,title:app.btnQuickSaveStatus.title};
            window.saveAs=app.initialSaveAs;let aborted=0;app.photos[0].desc='寫入前未存內容';app.recordHistory('未存內容');app.currentProjectFileHandle={name:'寫入測試.photo-report',createWritable:async()=>({write:async()=>{throw new Error('模擬寫入失敗');},close:async()=>{},abort:async()=>{aborted++;}})};let failed=false;try{await app.saveProject({lite:false});}catch{failed=true;}const afterFailure={failed,aborted,dirty:app.hasUnsavedProjectChanges(),mode:app.currentProjectSaveMode};
            app.currentProjectFileHandle={name:'寫入測試.photo-report',createWritable:async()=>({write:async()=>{},close:async()=>{}})};await app.saveProject({lite:false});const written={mode:app.currentProjectSaveMode,text:document.getElementById('quickSaveStatusText').textContent};
            app.photos[0].desc='取消選檔前修改';app.recordHistory('取消前');const previousTime=app.currentProjectSavedAt;Object.defineProperty(navigator,'webdriver',{configurable:true,get:()=>false});window.showSaveFilePicker=async()=>{throw new DOMException('取消','AbortError');};const cancelled=await app.saveProject({saveAsNew:true,fromUserPicker:true,lite:false});const cancelState={cancelled,dirty:app.hasUnsavedProjectChanges(),sameTime:previousTime===app.currentProjectSavedAt};window.saveAs=originalSave;app.currentProjectFileHandle=null;return{download,afterFailure,written,cancelState,capturedBytes:captured.size};
        });
        assert.equal(saveStates.download.mode,'download');assert.match(saveStates.download.text,/已產生下載檔/);assert.match(saveStates.download.title,/請確認/);assert.equal(saveStates.afterFailure.failed,true);assert.equal(saveStates.afterFailure.aborted,1);assert.equal(saveStates.afterFailure.dirty,true);assert.equal(saveStates.afterFailure.mode,'download');assert.equal(saveStates.written.mode,'written');assert.match(saveStates.written.text,/已寫入檔案/);assert.equal(saveStates.cancelState.cancelled,false);assert(saveStates.cancelState.dirty&&saveStates.cancelState.sameTime);
        console.log('下載／實際寫入狀態、寫入失敗與取消選檔保護通過');

        const duringSave=await page.evaluate(async()=>{
            app.currentProjectFileHandle=null;
            app.photos[0].desc='開始存檔時的說明';app.recordHistory('快照測試');
            const original=app.createLitePhotoEntry;let edited=false,blob;
            window.saveAs=value=>{blob=value;};
            app.createLitePhotoEntry=async function(photo,...args){if(!edited){edited=true;this.photos[0].desc='存檔期間的新修改';this.recordHistory('存檔期間修改');}return original.call(this,photo,...args);};
            try{await app.saveProject();}finally{app.createLitePhotoEntry=original;}
            const zip=await JSZip.loadAsync(blob);const saved=JSON.parse(await zip.file('manifest.json').async('text'));
            return{saved:saved.photos[0].desc,current:app.photos[0].desc,dirty:app.hasUnsavedProjectChanges(),badge:document.getElementById('quickSaveStatusText').textContent};
        });
        assert.equal(duringSave.saved,'開始存檔時的說明');assert.equal(duringSave.current,'存檔期間的新修改');assert(duringSave.dirty);assert.match(duringSave.badge,/還有修改沒存/);
        console.log('儲存期間繼續編輯：檔案保持固定快照，新修改仍標示未存');
        const duringDraft=await page.evaluate(async()=>{
            clearTimeout(app.autoDraftTimer);app.photos[0].desc='草稿開始時的說明';
            const file=app.photos[0].file,original=file.arrayBuffer.bind(file);let edited=false;
            Object.defineProperty(file,'arrayBuffer',{configurable:true,value:async()=>{if(!edited){edited=true;app.photos[0].desc='草稿期間的新修改';}return original();}});
            try{await app.saveAutoDraftNow();}finally{delete file.arrayBuffer;clearTimeout(app.autoDraftTimer);}
            const db=await app.openAutoDraftDb();let payload;
            try{payload=await new Promise((resolve,reject)=>{const req=db.transaction(app.autoDraftStoreName,'readonly').objectStore(app.autoDraftStoreName).get('latest');req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);});}finally{db.close();}
            return{saved:payload.photos[0].desc,current:app.photos[0].desc,hasCurrentDraft:app.lastAutoDraftSignature===app.projectSignature(),badge:document.getElementById('quickSaveStatusText').textContent};
        });
        assert.equal(duringDraft.saved,'草稿開始時的說明');assert.equal(duringDraft.current,'草稿期間的新修改');assert.equal(duringDraft.hasCurrentDraft,false);assert(!duringDraft.badge.includes('已自動暫存'));
        console.log('自動草稿期間繼續編輯：不把舊草稿標為最新修改已暫存');


        assert.equal(await open(baseline),true);
        await page.evaluate(()=>{document.getElementById('caseTitle').value='';app.recordHistory('空白案由');app.markProjectSaved(1000);});
        const incoming=await alter((z,m)=>{m.savedAt='2099-01-01T00:00:00.000Z';m.photos[0].desc='彙整後的新說明';m.photos[0].annotations[0].x=.6;m.photos[0].rotation=90;m.caseData.caseTitle='彙整後案件';});
        assert.equal(await merge(incoming),true);
        const merged=await page.evaluate(()=>({desc:app.photos[0].desc,report:app.lastMergeReport}));assert.equal(merged.desc,'彙整後的新說明');assert(merged.report.photos.some(p=>p.changes.some(c=>c.field==='desc')));assert(merged.report.photos.some(p=>p.changes.some(c=>c.field==='annotations')));assert(merged.report.caseChanges.some(c=>c.field==='caseTitle'));
        await page.locator('header details.app-menu').first().locator('summary').click();await page.click('#btnViewMergeReport');assert(await page.locator('#mergeReportModal').isVisible());assert((await page.locator('#mergeReportBody').innerText()).includes('彙整後的新說明'));
        if(process.env.PHOTO_REPORT_PROJECT_OUTPUT){fs.mkdirSync(process.env.PHOTO_REPORT_PROJECT_OUTPUT,{recursive:true});await page.screenshot({path:path.join(process.env.PHOTO_REPORT_PROJECT_OUTPUT,'merge-details.png')});}
        const countBefore=await page.evaluate(()=>app.photos.length);await page.keyboard.press('Delete');assert.equal(await page.evaluate(()=>app.photos.length),countBefore,'明細開啟時不可刪除照片');await page.keyboard.press('Escape');assert(await page.locator('#mergeReportModal').isHidden());
        await page.evaluate(()=>app.undo());assert.equal(await page.evaluate(()=>app.photos[0].desc),initial.photos[0].desc,'彙整可復原');assert.equal(await page.evaluate(()=>JSON.stringify(app.lastMergeReport)),JSON.stringify(merged.report),'明細為固定的該次紀錄');
        const bad=await alter((z,m)=>{m.version=99;});assert.equal(await merge(bad),false);assert.equal(await page.evaluate(()=>JSON.stringify(app.lastMergeReport)),JSON.stringify(merged.report),'失敗不覆蓋最近成功紀錄');
        assert.equal(await open(baseline),true);assert.equal(await page.evaluate(()=>app.lastMergeReport),null);assert(await page.locator('#btnViewMergeReport').isDisabled());
        console.log('彙整文字／標記／旋轉與案件明細、明細操作防護、Undo、紀錄隔離通過');
        const inputDir=process.env.PHOTO_REPORT_TEST_DIR;
        const photos=inputDir?fs.readdirSync(inputDir).filter(n=>/\.(png|jpe?g)$/i.test(n)).slice(0,4).map(n=>path.join(inputDir,n)):[path.resolve(__dirname,'../fixtures/sample01.jpg'),path.resolve(__dirname,'../fixtures/sample02.jpg')];
        await page.evaluate(()=>app.clearAll());await page.setInputFiles('#fileInput',photos);await page.waitForFunction(n=>app.photos.length===n,photos.length);
        const real=await page.evaluate(async()=>{document.getElementById('caseTitle').value='工作進度功能驗證';app.photos.forEach((p,i)=>{p.desc='測試說明 '+(i+1);});app.recordHistory('照片測試');const before=app.photos.map(p=>({name:p.name,desc:p.desc}));let project;window.saveAs=blob=>{project=blob;};await app.saveProject();const file=new File([project],'測試照片進度.photo-report');const opened=await app.openProjectFromFile(file);return{opened,before,after:app.photos.map(p=>({name:p.name,desc:p.desc})),bytes:project.size};});assert(real.opened);assert.deepEqual(real.after,real.before);
        const output=process.env.PHOTO_REPORT_PROJECT_OUTPUT;
        if(output){
            await page.evaluate(()=>{window.saveAs=(blob,name)=>{window.latestOutput={blob,name};};});
            for(const action of ['saveProject','exportDocx','exportPdf']){const promise=page.waitForEvent('download',{timeout:10000}).catch(()=>null);await page.evaluate(async name=>app[name](),action);if(action==='exportPdf'){const download=await promise;assert(download);await download.saveAs(path.join(output,'實測清冊.pdf'));}else{const file=await page.evaluate(async()=>({name:latestOutput.name,data:Array.from(new Uint8Array(await latestOutput.blob.arrayBuffer()))}));fs.writeFileSync(path.join(output,file.name),Buffer.from(file.data));}}
            await page.screenshot({path:path.join(output,'photo-workflow.png')});
        }
        clean();await page.close();
        console.log(`實際照片／測試素材 ${photos.length} 張：存檔、重開及內容比對通過`);
        const failurePage=await browser.newPage();const errors=trackPageErrors(failurePage);await failurePage.goto('about:blank');const boom=failurePage.waitForEvent('pageerror');await failurePage.evaluate(()=>{setTimeout(()=>{throw new Error('刻意測試錯誤防線');},0);});await boom;assert.throws(errors,/未處理例外/);await failurePage.close();
        const missing=await browser.newPage();const missingErrors=trackPageErrors(missing);await missing.route('**/js/project-format.js',route=>route.abort());const absent=missing.waitForEvent('pageerror');await missing.goto(pathToFileURL(path.resolve(__dirname,'../../index.html')).href);const error=await absent;assert.match(error.message,/必要的/);assert.throws(missingErrors,/未處理例外/);await missing.close();
        console.log('刻意觸發例外與必要模組缺少：測試防線確實能攔截');
    }finally{await browser.close();}
}
main().catch(error=>{console.error(error);process.exitCode=1;});
