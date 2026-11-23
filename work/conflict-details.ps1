$p=Resolve-Path 'outputs/Excel Column Mapper.html'
$script:s=[IO.File]::ReadAllText($p)
function ReplaceText([string]$old,[string]$new){if(-not $script:s.Contains($old)){throw "Missing text: $old"};$script:s=$script:s.Replace($old,$new)}
ReplaceText 'mismatches=[],sourceStats=[]' 'mismatches=[],conflictDetails=[],sourceStats=[]'
ReplaceText 'local.conflicts+=expected.conflicts.length;if(!expected.mapped)continue;' "local.conflicts+=expected.conflicts.length;for(let c of expected.conflicts)conflictDetails.push({file:d.file,sheet:d.sheet,row:row.__sourceRowNumber,column:c.source,field:f.name,priorityValue:String(expected.value??''),conflictingValue:String(c.value??''),actual:String(o[f.name]??''),reason:'A lower-priority source column contains a different non-empty value.'});if(!expected.mapped)continue;"
ReplaceText 'fieldStats,sourceStats,mismatches,unmappedColumns' 'fieldStats,sourceStats,mismatches,conflictDetails,unmappedColumns'
$a=$script:s.IndexOf('<h3>Source worksheet results</h3>')
if($a -lt 0){throw 'Could not find source result heading'}
$block=@'
${v.conflictDetails.length?`<details class="mismatch-details"><summary>Inspect ${fmt(v.conflictDetails.length)} priority conflicts</summary><div class="tablewrap scroll"><table><thead><tr><th>Source file</th><th>Worksheet</th><th>Source row</th><th>Conflicting source column</th><th>Standardized field</th><th>Priority output value</th><th>Conflicting value</th><th>Actual output</th><th>Reason</th></tr></thead><tbody>${v.conflictDetails.map(c=>`<tr><td>${esc(c.file)}</td><td>${esc(c.sheet)}</td><td>${c.row}</td><td>${esc(c.column)}</td><td>${esc(c.field)}</td><td>${esc(c.priorityValue)}</td><td>${esc(c.conflictingValue)}</td><td>${esc(c.actual)}</td><td>${esc(c.reason)}</td></tr>`).join('')}</tbody></table></div></details>`:''}
'@
$script:s=$script:s.Insert($a,$block.Trim())
[IO.File]::WriteAllText($p,$script:s,[Text.UTF8Encoding]::new($false))
Write-Output 'Added inspectable priority conflicts.'
