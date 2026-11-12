(function(root,factory){
    if(typeof module==='object'&&module.exports)module.exports=factory();
    else root.PhotoReportMergeReportUi=factory();
})(typeof globalThis!=='undefined'?globalThis:this,function(){
    'use strict';
    const LABELS={seq:'照片編號',date:'日期',time:'時間',timeAutoDefaulted:'時間來源',location:'地點',desc:'說明',rotation:'旋轉角度',flipH:'水平翻轉',parentUid:'連動原圖',cropRect:'特寫範圍',sideCrop:'並排特寫',annotations:'標記',caseTitle:'案由',caseDate:'案件日期',defaultLocation:'案件地點',deptName:'製作單位',officerName:'製作人',layoutSelect:'清冊版型'};
    const pct=value=>Math.round(Number(value||0)*1000)/10+'%';
    function rect(value){return `左 ${pct(value.x)}、上 ${pct(value.y)}、寬 ${pct(value.w)}、高 ${pct(value.h)}`;}
    function formatValue(field,value){
        if(value===null||value===undefined||value==='')return '（空白）';
        if(field==='flipH')return value?'已翻轉':'未翻轉';
        if(field==='timeAutoDefaulted')return value?'系統預設時間':'填寫／照片時間';
        if(field==='rotation')return value+'°';
        if(field==='cropRect')return rect(value);
        if(field==='sideCrop')return rect(value.cropRect)+`；寬度 ${pct(value.widthRatio)}、垂直位置 ${pct(value.offsetY)}、框線 ${value.borderColor==='red'?'紅':value.borderColor==='none'?'無':'黃'}`;
        if(field==='annotations')return value.length?value.map((a,i)=>`${i+1}. ${a.type==='label'?'姓名標籤「'+(a.text||'')+'」':'框線'}（${a.color==='red'?'紅':'黃'}）：${rect(a)}${a.scale?'，縮放 '+a.scale:''}`).join('\n'):'（無標記）';
        if(typeof value==='boolean')return value?'是':'否';
        return String(value);
    }
    function element(tag,text,className){const el=document.createElement(tag);if(text!==undefined)el.textContent=text;if(className)el.className=className;return el;}
    function table(changes){
        const scroll=element('div',undefined,'merge-report-table-wrap'),t=element('table');
        const tr=element('tr');for(const title of ['欄位','彙整前','彙整後'])tr.append(element('th',title));const head=element('thead');head.append(tr);t.append(head);
        const body=element('tbody');for(const change of changes){const row=element('tr');row.append(element('th',LABELS[change.field]||change.field),element('td',formatValue(change.field,change.before)),element('td',formatValue(change.field,change.after)));body.append(row);}t.append(body);scroll.append(t);return scroll;
    }
    function render(container,meta,report){
        container.replaceChildren();
        meta.textContent=`${new Date(report.createdAt).toLocaleString('zh-TW')}・${report.sources.length} 份進度・${report.beforeCount} → ${report.afterCount} 張照片`;
        container.append(element('p','來源：'+report.sources.join('、'),'merge-report-sources'));
        if(report.caseChanges.length){container.append(element('h3','共用案件資料'),table(report.caseChanges));}
        if(!report.photos.length&&!report.caseChanges.length)container.append(element('p','本次資料相同，沒有需要變更的內容。'));
        for(const photo of report.photos){const details=element('details');details.open=report.photos.length<=3;details.append(element('summary',`${photo.kind==='added'?'新加入':'更新'}・編號 ${photo.seq}・${photo.name}`));if(photo.changes.length)details.append(table(photo.changes));else details.append(element('p','已加入照片；沿用原有欄位。'));container.append(details);}
    }
    return{render,formatValue};
});
