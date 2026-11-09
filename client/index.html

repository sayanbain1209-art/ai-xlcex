<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SmartSheet AI</title>
<style>
*{box-sizing:border-box}
body{margin:0;background:#0b1020;color:#f4f6ff;font-family:Arial,sans-serif;padding:16px}
.container{max-width:1100px;margin:auto}
header{text-align:center;padding:20px 0}
h1{font-size:clamp(27px,5vw,34px);color:#72a7ff;margin:8px}
.sub,.small{color:#aab3c8;line-height:1.55}
.card{background:#151e32;border:1px solid #293752;border-radius:16px;padding:20px;margin:16px 0;min-width:0}
h2{font-size:20px;margin:0 0 14px}
textarea,input[type=file]{width:100%;background:#0b1325;color:white;border:1px solid #40516f;border-radius:10px;padding:13px;font-size:16px}
textarea{min-height:170px;resize:vertical}
button{background:#3478f6;color:white;border:0;padding:13px 18px;border-radius:9px;font-size:15px;font-weight:bold;margin:10px 8px 0 0;cursor:pointer}
.secondary{background:#35435c}
#status{color:#91baff;line-height:1.5}
.tablewrap{overflow:auto;margin-top:14px;border:1px solid #35435c;border-radius:10px;max-width:100%}
table{border-collapse:separate;border-spacing:0;width:max-content;min-width:100%;background:#0b1325;font-size:14px;table-layout:auto}
th,td{border-right:1px solid #35435c;border-bottom:1px solid #35435c;padding:12px 16px;text-align:left;vertical-align:middle;white-space:normal;overflow-wrap:anywhere;word-break:normal;line-height:1.45;min-width:105px;max-width:280px}
th{background:#263b60;font-weight:700;position:sticky;top:0;z-index:1}
th:last-child,td:last-child{border-right:0}
tbody tr:last-child td{border-bottom:0}
footer{text-align:center;color:#8390a8;padding:25px}
@media(max-width:600px){body{padding:10px}.card{padding:14px}button{padding:12px 14px}th,td{padding:11px 13px;min-width:100px;max-width:220px}}
</style>
</head>
<body><div class="container">
<header><h1>⚡ SmartSheet AI</h1><p class="sub">AI Powered Automatic Excel Data Entry Platform</p></header>
<section class="card"><h2>📂 Upload Your File</h2><input id="file" type="file" accept=".txt,.csv,.json,.xlsx,.xls"><p class="small">Starter version: TXT, CSV, JSON and Excel. PDF, Word and image OCR require an AI/OCR backend.</p></section>
<section class="card"><h2>✍️ Paste Your Data</h2>
<textarea id="data" placeholder="Paste tabular data, for example:
Employee ID,Employee Name,Department,Basic Salary,HRA,Other Allowance,Deduction
EMP101,Rahul Das,Sales,18000,3000,2000,500
EMP102,Mina Roy,HR,22000,4000,1500,1000"></textarea>
<p class="small">You can paste rows with headings or simple comma-separated records without headings.</p>
<button onclick="analyze()">✨ Analyze Data</button><button class="secondary" onclick="clearAll()">Clear</button><p id="status">Ready to analyze your data.</p></section>
<section class="card"><h2>📊 Automatic Category</h2><p id="category">Waiting for data...</p></section>
<section class="card"><h2>📋 Excel Preview</h2><p class="small" id="hint">Your organized data will appear here.</p>
<div class="tablewrap"><table id="preview"><thead><tr><th>Your data will appear here</th></tr></thead><tbody></tbody></table></div>
<button onclick="autoFitAll()">↔ AutoFit All</button><button class="secondary" onclick="autoFitColumns()">AutoFit Column Width</button><button class="secondary" onclick="autoFitRows()">AutoFit Row Height</button>
<br><button onclick="downloadExcel()">⬇ Download Excel (.xlsx)</button><button class="secondary" onclick="downloadCSV()">Download CSV</button>
</section><footer>SmartSheet AI • Automatic Data Entry</footer></div>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script>
let headers=[],records=[],isPayroll=false;
const $=id=>document.getElementById(id);
function norm(s){return String(s).toLowerCase().replace(/[^a-z0-9]/g,'');}
function parseCSV(text){
 const lines=text.trim().split(/\r?\n/).filter(x=>x.trim());if(!lines.length)return [];
 const delim=lines[0].includes('\t')?'\t':',';
 return lines.map(line=>{let out=[],v='',q=false;for(let i=0;i<line.length;i++){let c=line[i];if(c==='"'){if(q&&line[i+1]==='"'){v+='"';i++;}else q=!q;}else if(c===delim&&!q){out.push(v.trim());v='';}else v+=c;}out.push(v.trim());return out;});
}
function detectCategory(s){s=s.toLowerCase();if(/employee|salary|designation|staff|hra|payroll|emp\d+/.test(s))return 'Employee Data Entry';if(/student|class|roll|school|marks/.test(s))return 'Student Information Entry';if(/patient|hospital|doctor|diagnosis/.test(s))return 'Hospital / Patient Records';if(/invoice|billing|gst/.test(s))return 'Invoice & Billing Entry';if(/product|stock|inventory|quantity/.test(s))return 'Inventory / Product Entry';if(/customer|phone|email|client/.test(s))return 'Customer Database Entry';if(/sales|purchase|order/.test(s))return 'Sales & Purchase Entry';if(/bank|transaction|debit|credit/.test(s))return 'Bank Statement Entry';return 'General Data Entry';}
function analyze(){
 const raw=$('data').value.trim();if(!raw){$('status').textContent='Please paste data or choose a file.';return;}
 try{
  if(raw.startsWith('{')||raw.startsWith('[')){let j=JSON.parse(raw);records=Array.isArray(j)?j:[j];headers=[...new Set(records.flatMap(o=>Object.keys(o)))];}
  else{
   const rows=parseCSV(raw);
   if(rows.length && rows[0].length>1){
    const first=rows[0];
    const looksLikeEmployeeRecord=/^emp\d+$/i.test(first[0]) && first.length>=3;
    if(looksLikeEmployeeRecord){
      headers=first.length===4?['Employee ID','Employee Name','City','Phone']:first.map((_,i)=>['Employee ID','Employee Name','City','Phone'][i]||`Field ${i+1}`);
      records=rows.filter(r=>r.some(v=>v!=='')).map(r=>Object.fromEntries(headers.map((h,i)=>[h,r[i]??''])));
    }else{
      headers=first;records=rows.slice(1).filter(r=>r.some(v=>v!=='')).map(r=>Object.fromEntries(headers.map((h,i)=>[h,r[i]??''])));
    }
   }else{
    const obj={};raw.split(/\r?\n/).forEach(line=>{let m=line.match(/^([^:：]+)\s*[:：]\s*(.*)$/);if(m)obj[m[1].trim()]=m[2].trim();});
    headers=Object.keys(obj);records=Object.keys(obj).length?[obj]:[];
   }
  }
  if(!headers.length||!records.length)throw Error('No data rows found. Check that each record is on a new line and fields are separated by commas.');
  const all=Object.keys(records[0]).map(norm);
  const hasEmployee=all.some(k=>['employeeid','empid','employeename','employee','staffname'].includes(k));
  const hasPay=all.some(k=>['basicsalary','basic','hra','salary','deduction','otherallowance'].includes(k));
  isPayroll=hasEmployee&&hasPay;
  if(isPayroll)addPayrollColumns();
  $('category').textContent=detectCategory(headers.join(' ')+' '+raw);
  $('hint').textContent=isPayroll?'Payroll detected. Gross Salary = Basic + HRA + Other Allowance + Overtime. Net Salary = Gross Salary − Total Deduction. Missing optional amounts are treated as ₹0; check values before export.':'Data arranged with one record per row.';
  render();autoFitAll();
  $('status').textContent=`Analysis complete: ${records.length} record(s), ${headers.length} columns.`;
 }catch(e){$('status').textContent='Could not analyze: '+e.message;}
}
function findKey(obj,names){return Object.keys(obj).find(k=>names.includes(norm(k)));}
function num(v){let n=Number(String(v??'').replace(/[₹,\s]/g,''));return Number.isFinite(n)?n:0;}
function addPayrollColumns(){
 const first=records[0];const basicKey=findKey(first,['basicsalary','basic','basicpay']);
 const hraKey=findKey(first,['hra','houserentallowance']);
 const otherKey=findKey(first,['otherallowance','allowance','otherallowances']);
 const overtimeKey=findKey(first,['overtime','overtimeamount','otamount']);
 const deductionKey=findKey(first,['deduction','totaldeduction','deductions']);
 const grossKey='Gross Salary',totalDedKey='Total Deduction',netKey='Net Salary';
 for(const r of records){const gross=num(r[basicKey])+num(r[hraKey])+num(r[otherKey])+num(r[overtimeKey]);const ded=num(r[deductionKey]);r[grossKey]=gross;r[totalDedKey]=ded;r[netKey]=gross-ded;}
 headers=[...headers.filter(h=>![grossKey,totalDedKey,netKey].includes(h)),grossKey,totalDedKey,netKey];
}
function esc(x){return String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function render(){$('preview').innerHTML='<thead><tr>'+headers.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr></thead><tbody>'+records.map(r=>'<tr>'+headers.map(h=>'<td>'+esc(r[h]??'')+'</td>').join('')+'</tr>').join('')+'</tbody>';}
function autoFitColumns(){
 if(!records.length)return;const table=$('preview'),ctx=document.createElement('canvas').getContext('2d');ctx.font='14px Arial';
 [...table.querySelectorAll('thead th')].forEach((th,i)=>{let max=ctx.measureText(String(headers[i]??'')).width;records.forEach(r=>{max=Math.max(max,ctx.measureText(String(r[headers[i]]??'')).width);});const w=Math.max(105,Math.min(280,Math.ceil(max+38)));th.style.width=w+'px';th.style.minWidth=w+'px';table.querySelectorAll('tbody tr').forEach(tr=>{if(tr.children[i]){tr.children[i].style.width=w+'px';tr.children[i].style.minWidth=w+'px';}});});
}
function autoFitRows(){
 $('preview').querySelectorAll('th,td').forEach(cell=>{const len=(cell.textContent||'').length;cell.style.height='auto';cell.style.whiteSpace='normal';cell.style.overflowWrap='anywhere';cell.style.lineHeight='1.45';cell.style.padding='12px 16px';cell.style.fontSize=len>70?'12px':len>38?'13px':'14px';});
}
function autoFitAll(){autoFitColumns();autoFitRows();}
function csv(){return [headers,...records.map(r=>headers.map(h=>r[h]??''))].map(row=>row.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(',')).join('\r\n');}
function save(blob,name){let a=document.createElement('a'),u=URL.createObjectURL(blob);a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),1000);}
function downloadCSV(){if(!records.length)return $('status').textContent='Analyze data first.';save(new Blob(['\ufeff'+csv()],{type:'text/csv;charset=utf-8'}),'SmartSheet-AI-Data.csv');}
function downloadExcel(){
 if(!records.length)return $('status').textContent='Analyze data first.';if(!window.XLSX)return $('status').textContent='Excel library not loaded. Check internet and retry.';
 autoFitAll();let ws=XLSX.utils.json_to_sheet(records,{header:headers});
 ws['!cols']=headers.map(h=>{const vals=[String(h),...records.map(r=>String(r[h]??''))];const max=vals.reduce((m,v)=>Math.max(m,Math.min(42,v.length)),0);return {wch:Math.max(14,Math.min(42,max+4))};});
 ws['!rows']=[{hpt:26},...records.map(r=>({hpt:Math.min(60,Math.max(24,Math.ceil(Math.max(...headers.map(h=>String(r[h]??'').length),1)/34)*18+8))}))];
 let wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,isPayroll?'Payroll':'SmartSheet Data');XLSX.writeFile(wb,'SmartSheet-AI-Data.xlsx');
}
function clearAll(){headers=[];records=[];isPayroll=false;$('data').value='';$('file').value='';$('category').textContent='Waiting for data...';$('status').textContent='Ready to analyze your data.';$('hint').textContent='Your organized data will appear here.';$('preview').innerHTML='<thead><tr><th>Your data will appear here</th></tr></thead><tbody></tbody>';}
$('file').addEventListener('change',async e=>{let f=e.target.files[0];if(!f)return;try{if(/\.(xlsx|xls)$/i.test(f.name)){if(!window.XLSX)throw Error('Excel reader not loaded.');let wb=XLSX.read(await f.arrayBuffer());let rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});if(!rows.length)throw Error('No data found in workbook.');$('data').value=JSON.stringify(rows,null,2);analyze();}else{$('data').value=await f.text();analyze();}}catch(err){$('status').textContent=err.message;}});
</script></body></html>