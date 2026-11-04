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
    for(let r=0;r<Math.min(rows.length,30);r++){
      const vals=(rows[r]||[]).map(norm);
      const found=REQUIRED.filter(h=>vals.includes(norm(h)));
      if(found.length>=7 && vals.includes("记账日期") && vals.includes("支出金额") && vals.includes("收入金额")) return {row:r,headers:vals};
    }
    throw new Error("未能识别交通银行账单表头。请确认文件包含“记账日期、支出金额、收入金额”等字段。");
  }
  function asText(v){if(v===null||v===undefined)return "";if(v instanceof Date)return formatDate(v);return String(v).trim();}
  function formatDate(d){const p=n=>String(n).padStart(2,"0");return d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate());}
  function parseDate(v){
    if(v instanceof Date)return formatDate(v);
    if(typeof v==="number" && v>20000 && v<80000){const d=XLSX.SSF.parse_date_code(v);return [d.y,String(d.m).padStart(2,"0"),String(d.d).padStart(2,"0")].join("-");}
    const s=asText(v);const m=s.match(/(20\d{2})[年./-](\d{1,2})[月./-](\d{1,2})/);
    if(m)return m[1]+"-"+m[2].padStart(2,"0")+"-"+m[3].padStart(2,"0");
    return s;
  }
  function parseTime(v){
    if(v instanceof Date)return v.toTimeString().slice(0,8);
    if(typeof v==="number" && v>=0 && v<1){const n=Math.round(v*86400);return [Math.floor(n/3600),Math.floor(n%3600/60),n%60].map(x=>String(x).padStart(2,"0")).join(":");}
    const s=asText(v);const m=s.match(/(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
    return m?[m[1].padStart(2,"0"),m[2].padStart(2,"0"),(m[3]||"00").padStart(2,"0")].join(":"):s;
  }
  function amount(v){if(typeof v==="number")return Math.abs(v);const s=String(v??"").replace(/[￥¥,\s]/g,"");if(!s||s==="—"||s==="-")return 0;const n=Number(s);return Number.isFinite(n)?Math.abs(n):0;}
  function classify(method,place,counterparty){
    const s=(method+" "+place+" "+counterparty).toLowerCase();
    if(/退款|退货|冲正|撤销/.test(s))return "退款";
    if(/转账|转入|转出|汇款|本人账户/.test(s))return "转账";
    if(/取现|现金|atm/.test(s))return "现金";
    if(/利息|工资|代发|入账|存入/.test(s))return "收入";
    if(/消费|支付|pos|二维码|快捷|网银|网上/.test(s))return "商户消费";
    return "其他";
  }
  function ownTransfer(counterparty,method){
    const n=String(counterparty||"").replace(/\s/g,"");
    return !!n && /转账|转入|转出|本人|本行转账|异地交行/.test(String(method||"")) && /(翟维杰|本人|本人账户)/.test(n);
  }
  function processRows(rows,header){
    const idx={};header.headers.forEach((h,i)=>idx[h]=i);
    const get=(row,k)=>idx[norm(k)]===undefined?"":row[idx[norm(k)]];
    const records=[],seen=new Set();let duplicates=0,invalid=0,own=0;
    for(let i=header.row+1;i<rows.length;i++){
      const row=rows[i]||[];if(!row.some(v=>String(v??"").trim()))continue;
      const date=parseDate(get(row,"记账日期")),time=parseTime(get(row,"交易时间"));
      const place=asText(get(row,"交易地点")),method=asText(get(row,"交易方式"));
      const debit=amount(get(row,"支出金额")),credit=amount(get(row,"收入金额"));
      const counterparty=asText(get(row,"对方户名")),bank=asText(get(row,"对方开户行"));
      if(!date || (!debit&&!credit) || (!/^20\d{2}-\d{2}-\d{2}$/.test(date))){invalid++;continue;}
      if(debit>0&&credit>0){invalid++;continue;}
      const flow=debit>0?"支出":"收入",money=debit||credit;
      const key=[date,time,flow,money.toFixed(2),method,place,counterparty].join("|");
      if(seen.has(key)){duplicates++;continue;}seen.add(key);
      const category=classify(method,place,counterparty);
      if(ownTransfer(counterparty,method))own++;
      const product=place||method||"银行交易";
      const note=[method,bank?("对方开户行："+bank):"",get(row,"余额")!==""?("交易后余额："+asText(get(row,"余额"))):"",ownTransfer(counterparty,method)?"疑似本人账户转账，请核实":"",category==="退款"?"银行退款/退货":""].filter(Boolean).join("；");
      records.push([date+" "+time,category,counterparty,product,flow,money,"交通银行借记卡","已完成","","",note]);
    }
    records.sort((a,b)=>a[0].localeCompare(b[0]));
    if(!records.length)throw new Error("没有识别到有效交易。请检查日期、支出金额和收入金额列。");
    return {records,duplicates,invalid,own};
  }
  async function processFile(file){
    reset(false);if(!/\.xlsx?$/i.test(file.name)){status("文件格式不支持，请选择 .xls 或 .xlsx。","error");return;}
    $("fileInfo").classList.remove("hidden");$("fileInfo").textContent=file.name+" · "+(file.size/1024).toFixed(1)+" KB";
    if(!window.XLSX){status("Excel 解析组件加载失败。请刷新页面后重试。","error");return;}
    status("正在本地读取并识别账单…");
    try{
      const data=await file.arrayBuffer();const wb=XLSX.read(data,{type:"array",cellDates:true});
      const sheetName=wb.SheetNames.includes("Sheet0")?"Sheet0":wb.SheetNames[0];
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:"",raw:true});
      const header=findHeader(rows);const result=processRows(rows,header);
      converted={records:result.records,filename:file.name.replace(/\.xlsx?$/i,"")};
      $("previewPanel").classList.remove("hidden");$("exportBtn").disabled=false;
      $("stats").innerHTML=[["有效交易",result.records.length],["排除重复",result.duplicates],["需检查行",result.invalid]].map(x=>'<div class="stat"><b>'+x[1]+'</b><span>'+x[0]+'</span></div>').join("");
      $("previewRows").innerHTML=result.records.slice(0,12).map(r=>"<tr>"+[r[0],r[2]||"—",r[1],r[4],r[5].toFixed(2)].map((v,i)=>"<td"+(i===4?' class="num"':"")+">"+escapeHtml(String(v))+"</td>").join("")+"</tr>").join("");
      $("mappingHint").textContent="识别工作表："+sheetName+"；表头位于第 "+(header.row+1)+" 行。共识别 "+result.records.length+" 笔，排除重复 "+result.duplicates+" 笔，跳过异常/空金额 "+result.invalid+" 行；疑似本人转账 "+result.own+" 笔（已加备注，不自动删除）。预览最多显示 12 笔。";
      status("识别完成。请核对预览后导出。","success");
    }catch(e){status(e.message||"处理失败，请检查账单文件。","error");}
  }
  function escapeHtml(s){return s.replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
  function exportFile(){
    if(!converted||!window.XLSX)return;
    const aoa=Array.from({length:17},()=>[]);aoa.push(OUTPUT);converted.records.forEach(r=>aoa.push(r));
    const ws=XLSX.utils.aoa_to_sheet(aoa);ws["!cols"]=[{wch:21},{wch:13},{wch:24},{wch:24},{wch:9},{wch:12},{wch:19},{wch:12},{wch:24},{wch:24},{wch:42}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Sheet1");
    XLSX.writeFile(wb,converted.filename+"_微信格式.xlsx",{bookType:"xlsx",compression:true});
    status("文件已在本地生成并开始下载。","success");
  }
  function reset(clearInput=true){
    converted=null;$("previewPanel").classList.add("hidden");$("exportBtn").disabled=true;
    $("fileInfo").classList.add("hidden");$("fileInfo").textContent="";
    $("stats").innerHTML="";$("previewRows").innerHTML="";
    if(clearInput)$("fileInput").value="";
    status("等待选择文件");
  }
})();