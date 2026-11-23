$p=Resolve-Path 'outputs/Excel Column Mapper.html'
$s=[IO.File]::ReadAllText($p)
$old=@'
,'Country':['country'],'LinkedIn':['linkedin'
'@.Trim()
$new=@'
,'Country':['country'],'Contact Lead Stage':['lead stage','lead status','status'],'LinkedIn':['linkedin'
'@.Trim()
$s=$s.Replace($old,$new)
$old=@'
if(n==='title'){
'@.Trim()
$new=@'
if(n==='name'){let idx=d?.headers.indexOf(h)??-1,values=idx<0?[]:d.rows.map(r=>String(r[idx]??'').trim()).filter(Boolean),hasEmail=d?.headers.some(x=>normalize(x).includes('email')),hasCompany=d?.headers.some(x=>/company|organization|organisation|business|account/.test(normalize(x))),personLike=values.length&&values.every(v=>{let parts=v.split(/\s+/);return parts.length>=2&&parts.length<=5&&/^[A-ZÀ-ÖØ-Þ]/.test(v)&&!/\b(ltd|limited|inc|llc|plc|corp|corporation|company|group|holdings|agency)\b/i.test(v);});if(personLike&&hasEmail&&hasCompany){let f=personField();return f?{field:f,confidence:'High',reason:'Values look like contact names and the dataset has separate company and email columns.'}:{field:'',confidence:'Low',reason:'Contact-like names found, but the template has no Contact Person field.'};}return {field:'',confidence:'Low',reason:'Name could refer to a contact or an organization. Choose the destination after reviewing the values.'};}if(n==='title'){
'@.Trim()
if(-not $s.Contains($old)){throw 'Title classifier insertion point missing'}
$s=$s.Replace($old,$new)
[IO.File]::WriteAllText($p,$s,[Text.UTF8Encoding]::new($false))
Write-Output 'Added semantic matching for ambiguous Name columns and lead status.'
