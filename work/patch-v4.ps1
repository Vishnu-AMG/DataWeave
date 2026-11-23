$path = Resolve-Path 'outputs/Excel Column Mapper.html'
$script:content = [IO.File]::ReadAllText($path)
function Replace-Section([string]$start,[string]$end,[string]$replacement) {
  $a = $script:content.IndexOf($start)
  $b = $script:content.IndexOf($end,$a)
  if ($a -lt 0 -or $b -lt 0) { throw "Could not locate section: $start" }
  $script:content = $script:content.Substring(0,$a) + $replacement + $script:content.Substring($b)
}
$bind = @'
function bindStep(){if(step===0){$('#addMore')?.addEventListener('click',()=>$('#sourceInput').click());$('#selectAll')?.addEventListener('click',()=>{datasets.forEach(d=>d.selected=true);render();});$('#selectNone')?.addEventListener('click',()=>{datasets.forEach(d=>d.selected=false);render();});document.querySelectorAll('[data-select]').forEach(el=>el.onchange=()=>{datasets[+el.dataset.select].selected=el.checked;});document.querySelectorAll('[data-remove]').forEach(el=>el.onclick=()=>{datasets.splice(+el.dataset.remove,1);render();});let drop=$('#drop');drop?.addEventListener('click',e=>{if(e.target===drop||e.target.tagName==='STRONG'||e.target.tagName==='SPAN')$('#sourceInput').click();});drop?.addEventListener('dragover',e=>{e.preventDefault();drop.style.borderColor='#4f6bed'});drop?.addEventListener('drop',e=>{e.preventDefault();addFiles(e.dataTransfer.files)});}if(step===1){$('#dsPick')?.addEventListener('change',renderSourcePreview);$('#previewSize')?.addEventListener('change',renderSourcePreview);renderSourcePreview();}if(step===2){$('#mapPick')?.addEventListener('change',renderMapping);renderMapping();}if(step===3){$('#dupeFlag')?.addEventListener('change',e=>dupeFlag=e.target.checked);}if(step===4){$('#exportXlsx')?.addEventListener('click',()=>exportMaster('xlsx'));$('#exportCsv')?.addEventListener('click',()=>exportMaster('csv'));}}
'@
$preview = @'
function renderSourcePreview(){let id=$('#dsPick')?.value,d=datasets.find(x=>x.id===id);if(!d)return;let count=+($('#previewSize')?.value||10);$('#sourcePreview').innerHTML=`<div class="stats"><span class="pill">${d.originalCount} records</span><span class="pill">${d.headers.length} columns</span><span class="pill">${esc(d.file)} · ${esc(d.sheet)}</span></div><div class="tablewrap scroll"><table><thead><tr>${d.headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${d.rows.slice(0,count).map(r=>`<tr>${d.headers.map((_,i)=>`<td>${esc(r[i])}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;}
'@
Replace-Section 'function bindStep(){' 'function renderMapping(){' ($bind + $preview)
$content = $content.Replace('<div id="sourcePreview"></div>','<div class="row" style="margin-top:12px"><label for="previewSize">Preview rows</label><select id="previewSize" class="selector"><option value="10">10 rows</option><option value="5">5 rows</option></select></div><div id="sourcePreview"></div>')
$content = $content.Replace('<input id="dupeFlag" type="checkbox">','<input id="dupeFlag" type="checkbox" ${dupeFlag?''checked'':''}>')
$content = $content.Replace("$('#dupeFlag')?.checked",'dupeFlag')
$content = $content.Replace('d.headers.filter(h=>d.map?.[h]===target);let vals=src.map','d.headers.filter(h=>d.map?.[h]===target).sort((a,b)=>(Number(d.priority?.[a])||d.headers.indexOf(a)+1)-(Number(d.priority?.[b])||d.headers.indexOf(b)+1));let vals=src.map')
$content = $content.Replace('<h2>Download your files</h2><p class="hint">Export the master file in Excel and CSV, plus a separate processing summary and mapping configuration.</p><div class="row"><button class="primary" id="exportAll">Download all outputs</button><button id="exportXlsx">Excel master</button><button id="exportCsv">CSV master</button><button id="exportSummary">Processing summary</button><button id="exportMap">Mapping configuration</button></div>','<h2>Download standardized data</h2><p class="hint">Choose your final appended dataset format.</p><div class="row"><button class="primary" id="exportXlsx">Download Standardized_Master.xlsx</button><button id="exportCsv">Download Standardized_Master.csv</button></div>')
$content = $content.Replace('if(step===2){let bad=template.fields.filter','if(step===2&&!mappingConfirmed)return alert("Confirm the mappings before continuing.");if(step===2){let bad=template.fields.filter')
$content = $content.Replace('if(step<4){step++;render();}else exportMaster(''xlsx'');','if(step<4){step++;if(step===4)recordHistory();render();}')
$content = $content.Replace("dsFrom(file.name,'CSV',cellRows(wb.Sheets[wb.SheetNames[0]]))","dsFrom(file.name,'CSV',cellRows(wb.Sheets[wb.SheetNames[0]]),file.size)")
$content = $content.Replace('dsFrom(file.name,sn,cellRows(wb.Sheets[sn]))','dsFrom(file.name,sn,cellRows(wb.Sheets[sn]),file.size)')
$content = $content.Replace('if(!d)return;d.map=d.map||{};d.priority=d.priority||{};','if(!d)return;mappingConfirmed=!!d.confirmed;d.map=d.map||{};d.priority=d.priority||{};')
$content = $content.Replace('d.map[el.dataset.header]=el.value;mappingConfirmed=false;renderMapping();','d.map[el.dataset.header]=el.value;mappingConfirmed=false;d.confirmed=false;renderMapping();')
$content = $content.Replace('d.priority[el.dataset.priority]=Math.max(1,+el.value||1);mappingConfirmed=false;','d.priority[el.dataset.priority]=Math.max(1,+el.value||1);mappingConfirmed=false;d.confirmed=false;')
$content = $content.Replace('d.map={};d.priority={};mappingConfirmed=false;renderMapping();','d.map={};d.priority={};mappingConfirmed=false;d.confirmed=false;renderMapping();')
$content = $content.Replace("mappingConfirmed=true;$('#confirmMapping')","mappingConfirmed=true;d.confirmed=true;$('#confirmMapping')")
$content = $content.Replace('if(step===2&&!mappingConfirmed)return alert("Confirm the mappings before continuing.");','if(step===2&&!datasets.filter(d=>d.selected).every(d=>d.confirmed))return alert("Confirm the mappings for every selected worksheet before continuing.");')
Replace-Section 'function exportSummary(){' 'function download(blob,name){' ''
[IO.File]::WriteAllText($path,$content,[Text.UTF8Encoding]::new($false))
Write-Output 'Updated browser workflow and exports.'
