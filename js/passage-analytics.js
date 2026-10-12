// Read-only passage table. Preferences are local; no passage records are changed.
const PASSAGE_TABLE_KEY = 'steeler_passage_table_v1';
const PASSAGE_TABLE_COLUMNS = [
  ['date','Date','date'], ['route','Route','text'], ['category','Category','text'],
  ['status','Status','text'], ['nights','Nights on board','number'],
  ['nm','Distance (NM)','number'], ['underwayMinutes','Under way (h:mm)','duration'],
  ['engineHours','Engine hours','number'], ['fuel','Fuel used (L)','number'],
  ['fuelPerNm','Litres / NM','number'], ['averageSpeed','Average speed (kn)','number'],
  ['origin','Origin','text'], ['destination','Destination','text']
];
function passageTableView(){
  try {
    const v=JSON.parse(storage.getItem(PASSAGE_TABLE_KEY)||'{}');
    return {sort:PASSAGE_TABLE_COLUMNS.some(c=>c[0]===v.sort)?v.sort:'date', descending:v.descending!==false, filters:v.filters&&typeof v.filters==='object'?v.filters:{}};
  } catch { return {sort:'date',descending:true,filters:{}}; }
}
function passageTableRows(items){
  return items.filter(p=>!isDeletedPassage(p)).map(p=>{
    const n=computePassageCategoryNumbers(p), speed=passageAverageSpeed(p), nights=computeOvernightStats([p]).dates;
    const status=getPassageDashboardStatus(p);
    return {...n,id:p.id,date:getPassageDateValue(p),route:getRouteNames(p).join(' → '),
      category:normalisePassageCategories(p).length?normalisePassageCategories(p):['Uncategorised'],
      status,origin:p.plan?.from||'Unknown',destination:p.plan?.to||'Unknown',nights:nights.length,nightDates:nights,
      underwayMinutes:status==='ERU'?null:n.underwayMinutes,
      averageSpeed:status==='ERU'?null:speed.speed,speedDistance:status==='ERU'?0:speed.distance,speedMinutes:status==='ERU'?0:speed.minutes,
      fuelPerNm:n.economyDistance>0?n.economyFuel/n.economyDistance:null};
  });
}
function filterPassageTable(rows,view){
  return rows.filter(r=>PASSAGE_TABLE_COLUMNS.every(([k,,type])=>{
    const f=view.filters[k];if(!f)return true;
    const value=r[k];
    if(type==='date')return (!Array.isArray(f.values)||f.values.includes(value))&&(!f.from||value>=f.from)&&(!f.to||value<=f.to)&&(!f.month||value.slice(5,7)===f.month);
    if(type==='text')return !Array.isArray(f.values)||(Array.isArray(value)?value:[value]).some(v=>f.values.includes(v));
    if(f.min==null&&f.max==null)return true;
    return value!==null&&Number.isFinite(value)&&(f.min==null||value>=f.min)&&(f.max==null||value<=f.max);
  })).sort((a,b)=>{
    const av=a[view.sort],bv=b[view.sort];
    if(av==null||bv==null)return av==null?(bv==null?0:1):-1;
    const cmp=typeof av==='number'?av-bv:String(av).localeCompare(String(bv),undefined,{numeric:true});
    return (view.descending?-cmp:cmp)||String(a.id).localeCompare(String(b.id));
  });
}
function passageTableTotals(rows){
  const total={nightDates:[...new Set(rows.flatMap(r=>r.nightDates))]};total.nights=total.nightDates.length;
  for(const key of ['nm','underwayMinutes','engineHours','fuel']){
    const values=rows.map(r=>r[key]).filter(v=>v!==null&&Number.isFinite(v));
    total[key]=values.length?values.reduce((a,b)=>a+b,0):null;
  }
  const sum=k=>rows.reduce((n,r)=>n+(r[k]||0),0);
  total.averageSpeed=sum('speedMinutes')>0?sum('speedDistance')*60/sum('speedMinutes'):null;
  total.fuelPerNm=sum('economyDistance')>0?sum('economyFuel')/sum('economyDistance'):null;
  return total;
}
function passageTableValue(row,key,type){
  const v=row[key];if(v==null)return '–';
  if(type==='duration')return `${Math.floor(Math.round(v)/60)}:${String(Math.round(v)%60).padStart(2,'0')}`;
  if(type==='number')return Number(v).toFixed(key==='nights'?0:key==='fuelPerNm'?2:1);
  return Array.isArray(v)?v.join(', '):String(v);
}
function savePassageTableView(view){
  try {storage.setItem(PASSAGE_TABLE_KEY,JSON.stringify(view));}catch(e){warnStorageSaveFailed('analytics preferences',e);}
  renderPassageAnalytics();
}
function openPassageTableFilter(key,rows,view){
  const [,label,type]=PASSAGE_TABLE_COLUMNS.find(c=>c[0]===key), f=view.filters[key]||{};
  const esc=escapeHtml;
  let body='';
  if(type==='number'||type==='duration'){
    body=`<p>Leave either bound empty for no limit.${type==='duration'?' Enter time in minutes.':''}</p><label>Minimum <input id="ptMin" type="number" step="any" value="${esc(String(f.min??''))}"></label><label>Maximum <input id="ptMax" type="number" step="any" value="${esc(String(f.max??''))}"></label>`;
  }else{
    const values=[...new Set(rows.flatMap(r=>Array.isArray(r[key])?r[key]:[r[key]]))].sort();
    const checked=v=>!Array.isArray(f.values)||f.values.includes(v);
    const option=v=>`<label class="pt-filter-option"><input type="checkbox" data-pt-value="${esc(v)}" ${checked(v)?'checked':''}> ${esc(v||'Unknown date')}</label>`;
    body='<div class="pt-filter-actions"><button type="button" id="ptSelectAll">Select all</button><button type="button" id="ptSelectNone">Select none</button></div>';
    if(type==='date'){
      body+=`<div class="pt-date-range"><label>From <input id="ptFrom" type="date" value="${esc(f.from||'')}"></label><label>To <input id="ptTo" type="date" value="${esc(f.to||'')}"></label><label>Month across all years <select id="ptMonth"><option value="">All months</option>${Array.from({length:12},(_,i)=>{const m=String(i+1).padStart(2,'0');return `<option value="${m}" ${f.month===m?'selected':''}>${new Date(2000,i,1).toLocaleString('en-GB',{month:'long'})}</option>`;}).join('')}</select></label></div>`;
      body+='<div class="pt-filter-values">'+[...new Set(values.map(d=>d.slice(0,4)))].map(y=>`<details><summary><input type="checkbox" data-pt-prefix="${esc(y)}" aria-label="Select year ${esc(y)}"> ${esc(y||'Unknown date')}</summary>${[...new Set(values.filter(d=>d.startsWith(y)).map(d=>d.slice(0,7)))].map(m=>`<details class="pt-month"><summary><input type="checkbox" data-pt-prefix="${esc(m)}" aria-label="Select month ${esc(m)}"> ${esc(m||'Unknown date')}</summary>${values.filter(d=>d.slice(0,7)===m).map(option).join('')}</details>`).join('')}</details>`).join('')+'</div>';
    }else body+='<div class="pt-filter-values">'+values.map(option).join('')+'</div>';
  }
  body+='<p id="ptFilterError" role="alert"></p>';
  showModal({title:`Filter ${label}`,bodyHtml:body,okText:'Apply',onOk:()=>{
    let next;
    if(type==='number'||type==='duration'){
      const min=document.getElementById('ptMin'),max=document.getElementById('ptMax');
      if(!min.validity.valid||!max.validity.valid)return false;
      next={min:min.value===''?null:Number(min.value),max:max.value===''?null:Number(max.value)};
      if(next.min!=null&&next.max!=null&&next.min>next.max){document.getElementById('ptFilterError').textContent='Minimum must not exceed maximum.';return false;}
      if(next.min==null&&next.max==null)next=null;
    }else{
      const options=[...modalBody.querySelectorAll('[data-pt-value]')];
      next={};if(options.some(o=>!o.checked))next.values=options.filter(o=>o.checked).map(o=>o.dataset.ptValue);
      if(type==='date'){
        next.from=document.getElementById('ptFrom').value;next.to=document.getElementById('ptTo').value;next.month=document.getElementById('ptMonth').value;
        if(next.from&&next.to&&next.from>next.to){document.getElementById('ptFilterError').textContent='From must be on or before To.';return false;}
      }
      if(!Object.keys(next).some(k=>Array.isArray(next[k])||next[k]))next=null;
    }
    if(next)view.filters[key]=next;else delete view.filters[key];savePassageTableView(view);
  }});
  const updateParents=()=>modalBody.querySelectorAll('[data-pt-prefix]').forEach(parent=>{
    const children=[...modalBody.querySelectorAll('[data-pt-value]')].filter(c=>(parent.dataset.ptPrefix?c.dataset.ptValue.startsWith(parent.dataset.ptPrefix):!c.dataset.ptValue));
    parent.checked=children.every(c=>c.checked);parent.indeterminate=children.some(c=>c.checked)&&!parent.checked;
  });
  modalBody.querySelectorAll('[data-pt-prefix]').forEach(parent=>parent.onchange=()=>{
    modalBody.querySelectorAll('[data-pt-value]').forEach(c=>{if((parent.dataset.ptPrefix?c.dataset.ptValue.startsWith(parent.dataset.ptPrefix):!c.dataset.ptValue))c.checked=parent.checked;});updateParents();
  });
  modalBody.querySelectorAll('[data-pt-value]').forEach(c=>c.onchange=updateParents);
  for(const [id,checked] of [['ptSelectAll',true],['ptSelectNone',false]]){
    const b=document.getElementById(id);if(b)b.onclick=()=>{modalBody.querySelectorAll('[data-pt-value]').forEach(c=>c.checked=checked);updateParents();};
  }
  updateParents();
}
function renderPassageTable(){
  const controls=document.getElementById('passageAnalyticsControls'),box=document.getElementById('passageAnalyticsSummary');if(!controls||!box)return;
  const view=passageTableView(),all=passageTableRows(activePassages()),rows=filterPassageTable(all,view),totals=passageTableTotals(rows);
  const scroll=box.querySelector('.pt-scroll'),left=scroll?.scrollLeft||0,top=scroll?.scrollTop||0;
  controls.innerHTML=`<p role="status">${rows.length} of ${all.length} passages · ${Object.keys(view.filters).length} active column filters</p><button type="button" class="btn btn-secondary" id="ptClear">Clear filters</button>`;
  document.getElementById('ptClear').onclick=()=>savePassageTableView({...view,filters:{}});
  box.innerHTML=`<div class="analytics-table-scroll pt-scroll" tabindex="0" role="region" aria-label="Passage analytics table; scroll for more columns"><table class="analytics-table pt-table"><caption class="sr-only">Passage analytics</caption><thead><tr>${PASSAGE_TABLE_COLUMNS.map(([k,label])=>`<th scope="col" aria-sort="${view.sort===k?(view.descending?'descending':'ascending'):'none'}"><button type="button" data-pt-sort="${k}">${escapeHtml(label)} ${view.sort===k?(view.descending?'↓':'↑'):''}</button><button type="button" data-pt-filter="${k}" aria-label="Filter ${escapeHtml(label)}" aria-pressed="${!!view.filters[k]}">Filter${view.filters[k]?' •':''}</button></th>`).join('')}</tr></thead><tbody>${rows.length?rows.map(r=>`<tr>${PASSAGE_TABLE_COLUMNS.map(([k,,type],i)=>i===0?`<th scope="row">${analyticsRecordLink(r.id,r.date||'Unknown date')}</th>`:`<td>${escapeHtml(passageTableValue(r,k,type))}</td>`).join('')}</tr>`).join(''):`<tr><td colspan="${PASSAGE_TABLE_COLUMNS.length}">No passages match these filters.</td></tr>`}</tbody><tfoot><tr>${PASSAGE_TABLE_COLUMNS.map(([k,,type],i)=>i===0?'<th scope="row">Filtered total</th>':`<td>${type==='text'?(k==='route'?`${rows.length} passages`:''):escapeHtml(passageTableValue(totals,k,type))}</td>`).join('')}</tr></tfoot></table></div>`;
  box.querySelectorAll('[data-pt-sort]').forEach(b=>b.onclick=()=>savePassageTableView({...view,sort:b.dataset.ptSort,descending:view.sort===b.dataset.ptSort?!view.descending:false}));
  box.querySelectorAll('[data-pt-filter]').forEach(b=>b.onclick=()=>openPassageTableFilter(b.dataset.ptFilter,all,view));
  box.querySelector('.pt-scroll').scrollLeft=left;box.querySelector('.pt-scroll').scrollTop=top;
}
