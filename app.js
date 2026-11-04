(() => {
  "use strict";
  const REQUIRED = ["记账日期","交易时间","交易地点","交易方式","支出金额","收入金额","余额","对方户名","对方开户行"];
  const OUTPUT = ["交易时间","交易类型","交易对方","商品","收/支","金额(元)","支付方式","当前状态","交易单号","商户单号","备注"];
  const $ = id => document.getElementById(id);
  let converted = null;
  $("fileInput").addEventListener("change", e => { if(e.target.files[0]) processFile(e.target.files[0]); });
  $("resetBtn").addEventListener("click", reset);
  $("exportBtn").addEventListener("click", exportFile);
  const dz=$("dropzone");
  ["dragenter","dragover"].forEach(n=>dz.addEventListener(n,e=>{e.preventDefault();dz.classList.add("drag");}));
  ["dragleave","drop"].forEach(n=>dz.addEventListener(n,e=>{e.preventDefault();dz.classList.remove("drag");}));
  dz.addEventListener("drop",e=>{const f=e.dataTransfer.files[0];if(f)processFile(f);});

  function status(msg,type=""){const el=$("status");el.textContent=msg;el.className="status "+type;}
  function norm(v){return String(v??"").replace(/\s+/g,"").replace(/[：:]/g,"").trim();}
  function findHeader(rows){
    for(let r=0;r<Math.min(rows.length,80);r++){
      const vals=(rows[r]||[]).map(norm);
      const found=REQUIRED.filter(h=>vals.includes(norm(h)));
      if(found.length>=7 && vals.includes("记账日期") && vals.includes("支出金额") && vals.includes("收入金额")) return {row:r,headers:vals};
    }
    throw new Error("未能识别交通银行账单表头。已扫描前80行，请确认原始表头字段。");
  }
  function asText(v){if(v===null||v===undefined)return "";if(v instanceof Date)return formatDate(v);return String(v).trim();}
  function formatDate(d){const p=n=>String(n).padStart(2,"0");return d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate());}
  function parseDate(v){
    if(v instanceof Date)return formatDate(v);
    if(typeof v==="number" && v>20000 && v<80000){const d=XLSX.SSF.parse_date_code(v);return [d.y,String(d.m).padStart(2,"0"),String(d.d).padStart(2,"0")].join("-");}
    const s=asText(v);const m=s.match(/(20\d{2})[年./-](\d{1,2})[月./-](\d{1,2})/);
    if(m)return m[1]+"-"+m[2].padStart(2,"0")+"-"+m[3].padStart(2,"0");
    return "";
  }
  function parseTime(v){
    if(v instanceof Date)return v.toTimeString().slice(0,8);
    if(typeof v==="number" && v>=0 && v<1){const n=Math.round(v*86400);return [Math.floor(n/3600),Math.floor(n%3600/60),n%60].map(x=>String(x).padStart(2,"0")).join(":");}
    const s=asText(v);const m=s.match(/(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
    return m?[m[1].padStart(2,"0"),m[2].padStart(2,"0"),(m[3]||"00").padStart(2,"0")].join(":"):"00:00:00";
  }
  function amount(v){
    if(typeof v==="number")return Math.abs(v);
    const s=String(v??"").replace(/[￥¥,，\s]/g,"");
    if(!s||s==="—"||s==="–"||s==="-")return 0;
    const n=Number(s);return Number.isFinite(n)?Math.abs(n):0;
  }
  function classify(method,place,counterparty,flow){
    const s=(method+" "+place+" "+counterparty).toLowerCase();
    if(/本人转账|本人账户|内部转账|信用卡还款|基金|理财|定期存款|零钱通|充值|提现|银证转账/.test(s))return {type:"中性交易",flow:"中性"};
    if(/退款|退货|冲正/.test(s))return {type:"退款",flow:"收入"};
    if(flow==="支出"){
      if(/预授权确认|消费|支付|pos|二维码|快捷|网银|网上/.test(s))return {type:"商户消费",flow:"支出"};
      if(/转账|汇款|转出/.test(s))return {type:"转账",flow:"中性"};
      return {type:"其他支出",flow:"支出"};
    }
    if(/转账|汇款|转入/.test(s))return {type:"转账",flow:"中性"};
    return {type:"收入",flow:"收入"};
  }
  function parseRows(rows,header,sheetName,filename){
    const idx={};header.headers.forEach((h,i)=>{if(h)idx[h]=i;});
    const get=(row,k)=>idx[norm(k)]===undefined?"":row[idx[norm(k)]];
    const records=[];let nonEmptyRows=0,accepted=0,blankAmount=0,badDate=0,bothSides=0;
    const duplicateKeys=new Map();
    for(let i=header.row+1;i<rows.length;i++){
      const row=rows[i]||[];
      if(!row.some(v=>String(v??"").trim()))continue;
      nonEmptyRows++;
      const date=parseDate(get(row,"记账日期"));
      const time=parseTime(get(row,"交易时间"));
      const place=asText(get(row,"交易地点"));
      const method=asText(get(row,"交易方式"));
      const debit=amount(get(row,"支出金额"));
      const credit=amount(get(row,"收入金额"));
      const counterparty=asText(get(row,"对方户名"));
      const bank=asText(get(row,"对方开户行"));
      if(!date){badDate++;continue;}
      if(debit>0&&credit>0){bothSides++;continue;}
      if(debit===0&&credit===0){blankAmount++;continue;}
      const rawFlow=debit>0?"支出":"收入",money=debit||credit;
      const cls=classify(method,place,counterparty,rawFlow);
      const duplicateKey=[date,time,rawFlow,money.toFixed(2),method,place,counterparty].join("|");
      duplicateKeys.set(duplicateKey,(duplicateKeys.get(duplicateKey)||0)+1);
      const note=[method,bank?("对方开户行："+bank):"",asText(get(row,"余额"))?("交易后余额："+asText(get(row,"余额"))):""].filter(Boolean).join("；");
      const paymentWay="交通银行储蓄卡";
      const state=cls.flow==="支出"?"支付成功":cls.flow==="收入"?"收入已入账":"交易成功";
      records.push({date,time,flow:cls.flow,amount:money,type:cls.type,counterparty,product:place||method||"银行交易",paymentWay,state,note,sourceRow:i+1});
      accepted++;
    }
    records.sort((a,b)=>(a.date+" "+a.time).localeCompare(b.date+" "+b.time));
    if(!records.length)throw new Error("未找到有效交易。请核对账单日期、收入金额和支出金额列。");
    const duplicateCandidates=[...duplicateKeys.values()].reduce((n,c)=>n+Math.max(0,c-1),0);
    const dates=records.map(r=>r.date).sort();
    const income=records.filter(r=>r.flow==="收入"),expense=records.filter(r=>r.flow==="支出"),neutral=records.filter(r=>r.flow==="中性");
    const summary={
      total:records.length,
      incomeCount:income.length,incomeAmount:income.reduce((s,r)=>s+r.amount,0),
      expenseCount:expense.length,expenseAmount:expense.reduce((s,r)=>s+r.amount,0),
      neutralCount:neutral.length,neutralAmount:neutral.reduce((s,r)=>s+r.amount,0),
      start:dates[0],end:dates[dates.length-1]
    };
    return {records,summary,diagnostics:{sheetName,headerRow:header.row+1,nonEmptyRows,accepted,blankAmount,badDate,bothSides,duplicateCandidates,skipped:nonEmptyRows-accepted}};
  }
  function parseBankMetadata(rows){
    const all=rows.slice(0,Math.min(rows.length,8)).flat().map(asText).join(" ");
    const period=all.match(/(20\d{2}[/-]\d{1,2}[/-]\d{1,2})\s*至\s*(20\d{2}[/-]\d{1,2}[/-]\d{1,2})/);
    const name=all.match(/姓名[：:]\s*([^\s]+)/);
    return {periodStart:period?period[1]:"",periodEnd:period?period[2]:"",owner:name?name[1]:""};
  }
  async function processFile(file){
    reset(false);
    if(!/\.xlsx?$/i.test(file.name)){status("文件格式不支持，请选择 .xls 或 .xlsx。","error");return;}
    $("fileInfo").classList.remove("hidden");$("fileInfo").textContent=file.name+" · "+(file.size/1024).toFixed(1)+" KB";
    if(!window.XLSX){status("Excel 解析组件加载失败。请刷新页面后重试。","error");return;}
    status("正在本地读取工作簿及所有有效行…");
    try{
      const data=await file.arrayBuffer();
      const wb=XLSX.read(data,{type:"array",cellDates:true});
      const sheetName=wb.SheetNames.includes("Sheet0")?"Sheet0":wb.SheetNames[0];
      const ws=wb.Sheets[sheetName];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:true});
      const header=findHeader(rows);
      const result=parseRows(rows,header,sheetName,file.name);
      converted={records:result.records,summary:result.summary,diagnostics:result.diagnostics,metadata:parseBankMetadata(rows),filename:file.name.replace(/\.xlsx?$/i,"")};
      $("previewPanel").classList.remove("hidden");$("exportBtn").disabled=false;
      const d=result.diagnostics;
      $("stats").innerHTML=[["有效交易",d.accepted],["原始非空数据行",d.nonEmptyRows],["跳过待核对",d.skipped],["重复候选（保留）",d.duplicateCandidates]].map(x=>'<div class="stat"><b>'+x[1]+'</b><span>'+x[0]+'</span></div>').join("");
      $("previewRows").innerHTML=result.records.slice(0,15).map(r=>"<tr>"+[r.date+" "+r.time,r.counterparty||"—",r.type,r.flow,r.amount.toFixed(2)].map((v,i)=>"<td"+(i===4?' class="num"':"")+">"+escapeHtml(String(v))+"</td>").join("")+"</tr>").join("");
      $("mappingHint").textContent="工作表："+d.sheetName+"；表头第"+d.headerRow+"行；有效交易："+d.accepted+"；原始非空数据行："+d.nonEmptyRows+"；跳过："+d.skipped+"（日期异常 "+d.badDate+"、金额为空/为0 "+d.blankAmount+"、收入支出同时有值 "+d.bothSides+"）；疑似重复 "+d.duplicateCandidates+" 笔，全部保留。预览显示前15笔。";
      status("识别完成。程序不再自动删除疑似重复交易；请核对统计后导出。","success");
    }catch(e){status(e.message||"处理失败，请检查账单文件。","error");}
  }
  function escapeHtml(s){return s.replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
  function moneyText(n){return Number(n||0).toFixed(2);}
  function exportFile(){
    if(!converted||!window.XLSX)return;
    const s=converted.summary,meta=converted.metadata;
    const start=meta.periodStart||s.start,end=meta.periodEnd||s.end;
    const now=new Date(),pad=n=>String(n).padStart(2,"0");
    const exported=now.getFullYear()+"-"+pad(now.getMonth()+1)+"-"+pad(now.getDate())+" "+pad(now.getHours())+":"+pad(now.getMinutes())+":"+pad(now.getSeconds());
    const aoa=[
      ["微信支付账单明细"],
      ["微信昵称：[交通银行账单]"],
      ["起始时间：["+start+" 00:00:00] 终止时间：["+end+" 23:59:59]"],
      ["导出类型：[全部账单]"],
      ["导出时间：["+exported+"]"],
      [],
      ["共"+s.total+"笔记录"],
      ["收入："+s.incomeCount+"笔 "+moneyText(s.incomeAmount)+"元"],
      ["支出："+s.expenseCount+"笔 "+moneyText(s.expenseAmount)+"元"],
      ["中性交易："+s.neutralCount+"笔 "+moneyText(s.neutralAmount)+"元"],
      ["注："],
      ["1. 充值/提现/理财通购买/零钱通存取/信用卡还款等交易，将计入中性交易"],
      ["2. 若交易记录明细无有效内容，则代表该时间段内此账单无交易"],
      ["3. 本明细仅供个人对账使用"],
      ["4. 本账单中所有时间均为UTC+08:00时间"],
      [],
      ["----------------------微信支付账单明细列表--------------------"],
      OUTPUT
    ];
    converted.records.forEach(r=>aoa.push([
      r.date+" "+r.time,r.type,r.counterparty,r.product,r.flow,r.amount,r.paymentWay,r.state,"","",r.note
    ]));
    const ws=XLSX.utils.aoa_to_sheet(aoa);
    // Explicitly assign all summary cells so spreadsheet readers cannot render the
    // preamble as blank rows due to sparse-sheet serialization or stale output.
    aoa.slice(0,17).forEach((row,i)=>{
      if(row && row.length && row[0]!==undefined && row[0]!==null && String(row[0])!==""){
        ws["A"+(i+1)]={t:"s",v:String(row[0])};
      }
    });
    ws["!ref"]="A1:K"+aoa.length;
    ws["!rows"]=Array.from({length:aoa.length},(_,i)=>({hpt:i<17?20:22}));
    ws["!cols"]=[{wch:22},{wch:14},{wch:26},{wch:28},{wch:10},{wch:13},{wch:22},{wch:14},{wch:30},{wch:30},{wch:48}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Sheet1");
    XLSX.writeFile(wb,converted.filename+"_微信账单完整格式.xlsx",{bookType:"xlsx",compression:true});
    status("微信账单格式 XLSX 已生成。","success");
  }
  function reset(clearInput=true){
    converted=null;$("previewPanel").classList.add("hidden");$("exportBtn").disabled=true;
    $("fileInfo").classList.add("hidden");$("fileInfo").textContent="";
    $("stats").innerHTML="";$("previewRows").innerHTML="";
    if(clearInput)$("fileInput").value="";
    status("等待选择文件");
  }
})();