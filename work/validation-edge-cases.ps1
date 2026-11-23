$p=Resolve-Path 'outputs/Excel Column Mapper.html'
$s=[IO.File]::ReadAllText($p)
$old=@'
recordPct=sourceRecords?Math.max(0,(sourceRecords-Math.min(sourceRecords,recordErrors))/sourceRecords*100):100,fieldPct=expectedValues?correctValues/expectedValues*100:100
'@.Trim()
$new=@'
recordPct=sourceRecords?Math.max(0,(sourceRecords-Math.min(sourceRecords,recordErrors))/sourceRecords*100):null,fieldPct=expectedValues?correctValues/expectedValues*100:null
'@.Trim()
$s=$s.Replace($old,$new)
$old=@'
(sourceConflicts||result.unmapped||missingRequiredRows||sources.some
'@.Trim()
$new=@'
(!sourceRecords||sourceConflicts||result.unmapped||missingRequiredRows||sources.some
'@.Trim()
$s=$s.Replace($old,$new)
$old=@'
pct=x=>`${x.toFixed(1)}%`
'@.Trim()
$new=@'
pct=x=>x==null?'N/A':`${x.toFixed(1)}%`
'@.Trim()
$s=$s.Replace($old,$new)
$old=@'
let expectedValues=fieldStats.reduce
'@.Trim()
$new=@'
for(let o of additionalRows)mismatches.push({file:String(o['Source File Name']||''),sheet:String(o['Source Sheet Name']||''),row:o.__sourceRowNumber||'',column:'(record)',field:'(record)',expected:'No additional output record',actual:'Unmatched output record',reason:'Output record has no matching internal source-row identifier.'});let expectedValues=fieldStats.reduce
'@.Trim()
$s=$s.Replace($old,$new)
[IO.File]::WriteAllText($p,$s,[Text.UTF8Encoding]::new($false))
Write-Output 'Improved N/A accuracy handling and extra-row diagnostics.'
