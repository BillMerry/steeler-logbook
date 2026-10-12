const assert=require('node:assert/strict');
(async()=>{
 const browser=await (process.env.BROWSER_ENGINE?require('@playwright/test')[process.env.BROWSER_ENGINE]:require('./dom-harness.cjs')).launch();
 try{
  const context=await browser.newContext({viewport:{width:1024,height:768}}),page=await context.newPage();
  await page.goto('http://127.0.0.1:8765/');
  const result=await page.evaluate(()=>{
   const make=(id,date)=>({id,plan:{date,from:'Lymington',to:'Newtown',categories:['Local','Cruise'],dailySummaries:[{date:'2024-10-06',overnightOnBoard:true}]},entries:[]});
   passages=[make('a','2024-10-06'),make('b','2023-10-07'),make('c','2024-09-08'),{...make('deleted','2024-10-09'),deleted:true}];
   const original=JSON.stringify(passages);
   renderFuelManagementSettings();renderPassageAnalytics();
   const raw=passageTableRows(passages);
   const rows=raw.map((r,i)=>({...r,nm:i===2?null:[10,20][i],fuel:i===2?null:[0,10][i],speedDistance:[10,20,0][i],speedMinutes:[60,60,0][i],economyFuel:[0,10,0][i],economyDistance:[10,20,0][i]}));
   const filtered=filterPassageTable(rows,{sort:'fuel',descending:false,filters:{date:{month:'10'},fuel:{min:0,max:0}}});
   const totals=passageTableTotals(rows);
   const month=filterPassageTable(rows,{sort:'date',descending:true,filters:{date:{month:'10'}}}).map(r=>r.id);
   const sorted=filterPassageTable(rows,{sort:'nm',descending:true,filters:{}}).map(r=>r.id);
   const combined=filterPassageTable(rows,{sort:'date',descending:true,filters:{date:{values:['2024-10-06']},category:{values:['Cruise']}}}).map(r=>r.id);
   document.querySelector('[data-pt-filter="date"]').click();
   document.getElementById('ptSelectNone').click();
   const year=document.querySelector('[data-pt-prefix="2024"]');year.checked=true;year.dispatchEvent(new Event('change'));
   document.getElementById('modalOkBtn').click();
   const saved=passageTableView();
   const count=document.querySelectorAll('.pt-table tbody tr').length;
   document.querySelector('[data-pt-filter="fuel"]').click();
   document.getElementById('ptMin').value='10';document.getElementById('ptMax').value='5';document.getElementById('modalOkBtn').click();
   const invalid=!modalOverlay.classList.contains('hidden')&&document.getElementById('ptFilterError').textContent.includes('Minimum');
   document.getElementById('modalCancelBtn').click();
   return {original:JSON.stringify(passages)===original,filtered:filtered.map(r=>r.id),month,sorted,combined,totals,saved,count,invalid,missing:passageTableValue(rows[2],'nm','number'),zero:passageTableValue(rows[0],'fuel','number'),metricBoxes:document.querySelectorAll('[data-analytics-metric]').length};
  });
  assert.ok(result.original);assert.deepEqual(result.filtered,['a']);assert.deepEqual(result.month,['a','b']);assert.deepEqual(result.sorted,['b','a','c']);assert.deepEqual(result.combined,['a']);
  assert.equal(result.totals.nights,1);assert.equal(result.totals.nm,30);assert.equal(result.totals.fuel,10);assert.equal(result.totals.averageSpeed,15);assert.equal(result.totals.fuelPerNm,1/3);
  assert.equal(result.count,2);assert.ok(result.invalid);assert.equal(result.missing,'–');assert.equal(result.zero,'0.0');assert.equal(result.metricBoxes,0);
  await page.reload();assert.deepEqual(await page.evaluate(()=>passageTableView()),result.saved);
  if(process.env.BROWSER_ENGINE){
   await page.evaluate(()=>{
    passages=Array.from({length:40},(_,i)=>({id:'p'+i,plan:{date:'2024-10-'+String(i%28+1).padStart(2,'0'),from:'Lymington',to:'Newtown Creek',categories:['Local'],dailySummaries:[]},entries:[]}));
    storage.removeItem(PASSAGE_TABLE_KEY);switchToTab('settingsTab');renderFuelManagementSettings();
    document.getElementById('toggleFuelManagementBtn').click();
    document.querySelectorAll('#fuelManagementPanel details').forEach(d=>d.open=true);
    document.getElementById('passageAnalyticsCard').open=true;
   });
   // Keep the real Settings ancestors: their grid sizing must not expand the table.
   await page.evaluate(()=>{document.getElementById('passageAnalyticsCard').scrollIntoView();});
   const measure=await page.evaluate(()=>{
    const sc=document.querySelector('.pt-scroll');sc.scrollLeft=650;sc.scrollTop=150;
    const r=sc.getBoundingClientRect(),cell=sc.querySelector('tbody th').getBoundingClientRect(),head=sc.querySelector('thead th').getBoundingClientRect();
    return {x:Math.abs(cell.left-r.left),y:Math.abs(head.top-r.top),scroll:sc.scrollLeft,width:sc.clientWidth,viewport:innerWidth};
   });
   assert.ok(measure.width<measure.viewport,JSON.stringify(measure));assert.ok(measure.scroll>0,JSON.stringify(measure));assert.ok(measure.x<4,JSON.stringify(measure));assert.ok(measure.y<4,JSON.stringify(measure));
   await page.locator('#passageAnalyticsCard').screenshot({path:'test-results/passage-table-ipad.png'});
  }
  console.log('PASS: passage-table filters, year selection, sorting, weighted totals, unique nights, zero/missing values, persistence and read-only records'+(process.env.BROWSER_ENGINE?', frozen iPad column/header':''));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
