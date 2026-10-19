/** Render a validated, company-specific workbook spec with the user's BYD expression profile.
 * Usage: node build_workbook.mjs workbook_spec.json [output.xlsx]
 * Dependency: @oai/artifact-tool (provided by the artifact runtime).
 * This is a presentation layer; it does not extract annual reports or infer disclosures.
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const here=path.dirname(fileURLToPath(import.meta.url));
const profile=JSON.parse(await fs.readFile(path.join(here,'../assets/byd-expression-v1.json'),'utf8'));
const input=process.argv[2];if(!input)throw Error('Usage: node build_workbook.mjs workbook_spec.json [output.xlsx]');
const spec=JSON.parse(await fs.readFile(input,'utf8'));const output=path.resolve(process.argv[3]||spec.output);
const wb=Workbook.create();const col=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
for(const d of spec.sheets)wb.worksheets.add(d.name);
const C=profile.notes.colors;
const styles={
 title:{font:{size:14,bold:true,color:'#'+C.header},rowHeight:26},
 subtitle:{font:{size:9,color:'#808080'},rowHeight:17},
 header:{fill:'#'+C.header,font:{color:'#FFFFFF',bold:true},horizontalAlignment:'center',rowHeight:22},
 section:{fill:'#'+C.section,font:{color:'#'+C.header,bold:true},rowHeight:21},
 total:{font:{bold:true}},check:{fill:'#'+C.check,font:{color:'#000000'}},
 assumption:{font:{color:'#'+C.assumption}},ratio:{font:{color:'#000000'}},
 note:{font:{size:9,color:'#555555'},wrapText:true},
 bridge:{font:{bold:true,color:'#000000'}},
 revenueTitle:{font:{size:10,bold:true},rowHeight:21},
 pink:{fill:'#'+profile.revenue.colors.section,font:{bold:true,color:'#000000'},horizontalAlignment:'center',rowHeight:20},
 plainHeader:{font:{bold:true,color:'#000000'},rowHeight:21,borders:{top:{style:'thin',color:'#000000'},bottom:{style:'thin',color:'#000000'}}},
 revenueSummary:{font:{color:'#000000'}},summaryRatio:{font:{color:'#000000'}},
 revenueBridge:{font:{color:'#FF0000',bold:true}},subtotal:{font:{bold:true,color:'#000000'}},
 detailRatioBridge:{font:{color:'#FF0000',bold:true}},detailRatioSubtotal:{font:{color:'#000000',bold:true}},detailRatio:{font:{color:'#000000'}},
 govTitle:{fill:'#'+profile.government.title_color,font:{size:18,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',rowHeight:36},
 govSection:{font:{color:'#'+C.header,bold:true,size:11},rowHeight:24},
 govSummary:{fill:'#'+profile.government.summary_color,font:{bold:true,color:'#000000'}}
};
let formulaCount=0;
for(const d of spec.sheets){
 const sh=wb.worksheets.getItem(d.name),rev=d.layout==='revenue',gov=d.layout==='government',notes=d.layout==='notes';
 const yc=d.yearColumn||4,lc=d.labelColumn||2,sc=d.sourceColumn, N=d.rows.length, nc=Math.max(yc+d.years.length-1,...d.rows.map(r=>r.length),gov?10:13),last=col(nc);
 sh.showGridLines=rev;sh.tabColor=rev?'#FF0000':gov?'#365F91':'#1F4E78';
 sh.getRange(`A1:${last}${N}`).values=d.rows.map(r=>Array.from({length:nc},(_,i)=>typeof r[i]==='string'&&r[i].startsWith('=')?null:r[i]??null));
 sh.getRange(`A1:${last}${N}`).format={font:{name:profile.font,size:10,color:'#000000'},rowHeight:rev?18:15.5,verticalAlignment:'center'};
 sh.getRange(`${col(yc)}1:${last}${N}`).format={columnWidth:rev?14.33:13.83,horizontalAlignment:rev?'center':'right'};
 sh.getRange(`${col(yc)}1:${last}${N}`).setNumberFormat(rev?profile.revenue.amount_format:profile.notes.amount_format);
 if(rev){for(const [a,width]of Object.entries(profile.revenue.widths))if(a!=='year')sh.getRange(`${a}1:${a}${N}`).format.columnWidth=width;
  sh.getRange(`B1:B${N}`).format={font:{size:9},wrapText:true};
 }else if(gov){sh.getRange(`A1:A${N}`).format.columnWidth=45;sh.getRange(`B1:${last}${N}`).format.columnWidth=14.66;}
 else{sh.getRange(`A1:A${N}`).format.columnWidth=2.66;sh.getRange(`B1:B${N}`).format.columnWidth=notes?36.66:48;sh.getRange(`C1:C${N}`).format={columnWidth:notes?16.5:23,font:{color:'#666666',size:9},wrapText:true};}
 for(let rr=0;rr<N;rr++){
  const r=rr+1,row=d.rows[rr],sty=d.styles[String(r)],range=sh.getRange(`${col(lc)}${r}:${last}${r}`);
  for(let i=0;i<row.length;){if(typeof row[i]!=='string'||!row[i].startsWith('=')){i++;continue;}let j=i;while(j+1<row.length&&typeof row[j+1]==='string'&&row[j+1].startsWith('='))j++;
   const fs=row.slice(i,j+1).map(f=>sty==='check'&&!f.startsWith('=IF')?'=ROUND('+f.slice(1)+',8)':f);
   sh.getRange(`${col(i+1)}${r}:${col(j+1)}${r}`).formulas=[fs];formulaCount+=j-i+1;i=j+1;}
  if(sty&&styles[sty])range.format=styles[sty];
  if(['title','subtitle','note','pink','govTitle','govSection'].includes(sty)){
   // Merge to the full year span, so no title is clipped in Excel.
   sh.mergeCells(`${col(lc)}${r}:${last}${r}`);
   if(sty==='note'){const t=String(row[lc-1]||'');range.format.rowHeight=Math.max(25,Math.ceil(t.length/100)*16);}
  }
  if(sty==='total'&&!rev){sh.getRange(`${col(yc)}${r}:${last}${r}`).format.fill='#'+C.total;}
  const numeric=sh.getRange(`${col(yc)}${r}:${last}${r}`);
  if(notes||(!rev&&!gov)){
   if(!['header','title','subtitle','section','note','check','total'].includes(sty)){
    for(let j=yc-1;j<row.length;j++)if(typeof row[j]==='number')sh.getRange(`${col(j+1)}${r}`).format.font.color='#'+C.actual;
   }
   if(sty==='assumption')for(let j=0;j<d.years.length;j++)if(d.years[j]>spec.actualEnd)sh.getRange(`${col(j+yc)}${r}`).format.font.color='#'+C.assumption;
  }
  if(gov&&!['header','govTitle','subtitle','govSection','govSummary','check'].includes(sty))for(let j=yc-1;j<row.length;j++)if(typeof row[j]==='number')sh.getRange(`${col(j+1)}${r}`).format.font.color='#0000FF';
  if(sty==='check'){
   numeric.setNumberFormat('#,##0.00;(#,##0.00);0.00');
   numeric.conditionalFormats.addCustom(`AND(ISNUMBER(${col(yc)}${r}),ABS(${col(yc)}${r})>0.02)`,{fill:'#FCE4D6',font:{color:'#C00000',bold:true}});
  }
  if(sty==='ratio')numeric.setNumberFormat('0.0%');
  if(d.formats[String(r)])numeric.setNumberFormat(d.formats[String(r)]);
  if(['summaryRatio','detailRatioBridge','detailRatioSubtotal','detailRatio'].includes(sty))numeric.setNumberFormat(profile.revenue.ratio_format);
  if(rev){
   const label=String(row[2]||'');
   if(['收入','下面各项加总'].includes(label)){
    sh.getRange(`C${r}`).format={fill:'#FFFF00',font:{bold:true}};
    if(r>=d.detailRange?.[0]&&r<=d.detailRange?.[1])numeric.format={fill:'#FFFF00',font:{bold:true,color:'#000000'}};
   }else if(['revenueSummary','summaryRatio','revenueBridge','detailRatioBridge'].includes(sty)){
    sh.getRange(`C${r}`).format={fill:'#DBE5F1',font:{bold:true,color:'#FF0000'}};
   }
   if(r>=d.detailRange?.[0]&&r<=d.detailRange?.[1]){
    sh.getRange(`C${r}`).format.wrapText=true;
    const lines=Math.max(Math.ceil(label.length/29),Math.ceil(String(row[1]||'').length/17));
    if(lines>1)sh.getRange(`A${r}:${last}${r}`).format.rowHeight=Math.min(58,18*lines);
   }
   if(sty==='assumption')numeric.format={fill:'#FFFF00',font:{color:'#C00000'}};
   if(/数量/.test(label))numeric.setNumberFormat('#,##0');
   if(/工资（万元\/年）/.test(label))numeric.setNumberFormat('#,##0.0');
  }
 }
 if(rev){
  for(const hr of d.headerRows||[])for(let j=0;j<d.years.length;j++)sh.getRange(`${col(yc+j)}${hr}`).values=[[d.years[j]+'年'+(d.years[j]>spec.actualEnd?'E':'')]];
  for(const [start,end]of d.ratioRanges||[])for(let j=yc;j<yc+d.years.length;j++)sh.getRange(`${col(j)}${start}:${col(j)}${end}`).conditionalFormats.add('dataBar',{color:'#638EC6',thresholds:[0,1],gradient:true});
  sh.freezePanes.freezeRows(2);sh.freezePanes.freezeColumns(3);
 }else if(gov){
  for(const hr of d.headerRows||[])for(let j=0;j<d.years.length;j++){const cell=sh.getRange(`${col(yc+j)}${hr}`);cell.values=[[d.years[j]]];cell.setNumberFormat('0');}
  for(const r of d.summaryRows||[6,7,8,9])sh.getRange(`B${r}:${last}${r}`).setNumberFormat('#,##0.00');
  sh.freezePanes.freezeRows(5);sh.freezePanes.freezeColumns(1);
 }else if(d.name!=='目录'){sh.freezePanes.freezeRows(5);sh.freezePanes.freezeColumns(3);}
 console.log('Built',d.name,N);
}
wb.recalculate();
const verification={formulaCount,issues:[],chains:[]};
for(const d of spec.sheets){const sh=wb.worksheets.getItem(d.name),yc=d.yearColumn||4;
 for(const[r,sty]of Object.entries(d.styles))if(sty==='check'){
  const a=sh.getRange(`${col(yc)}${r}:${col(yc+d.years.length-1)}${r}`).values[0];
  a.forEach((v,j)=>{if(typeof v==='number'&&Math.abs(v)>.02)verification.issues.push({sheet:d.name,row:+r,label:d.rows[r-1][(d.labelColumn||2)-1],year:d.years[j],value:v});});
 }
}
for(const t of spec.chainTests||[]){
 const sh=wb.worksheets.getItem(t.inputSheet),out=wb.worksheets.getItem(t.outputSheet),v=sh.getRange(t.input).values[0][0],before=out.getRange(t.output).values[0][0];
 sh.getRange(t.input).values=[[t.testValue??v*1.1]];wb.recalculate();const after=out.getRange(t.output).values[0][0];verification.chains.push({...t,before,after,changed:typeof before==='number'&&typeof after==='number'&&Math.abs(before-after)>1e-6});sh.getRange(t.input).values=[[v]];
}
wb.recalculate();
const formulaErrors=[];
for(const d of spec.sheets){const sh=wb.worksheets.getItem(d.name),nc=Math.max(1,...d.rows.map(r=>r.length));for(const[rowIndex,row]of sh.getRange(`A1:${col(nc)}${d.rows.length}`).values.entries())for(const[columnIndex,v]of row.entries())if(typeof v==='string'&&/^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!)/.test(v))formulaErrors.push({sheet:d.name,cell:col(columnIndex+1)+(rowIndex+1),value:v});}
if(formulaErrors.length)throw Error('Formula errors: '+JSON.stringify(formulaErrors.slice(0,20)));
if(verification.chains.some(t=>!t.changed))throw Error('A forecast chain did not respond to its test input');
const scan=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:50},summary:'Formula errors'});
await fs.mkdir(path.dirname(output),{recursive:true});
await fs.writeFile(output+'.checks.json',JSON.stringify(verification,null,2));await fs.writeFile(output+'.formula-scan.ndjson',scan.ndjson);
if(spec.previewDirectory){await fs.mkdir(spec.previewDirectory,{recursive:true});for(const p of spec.previews||[]){try{const b=await wb.render({sheetName:p.sheet,range:p.range,scale:p.scale||1,format:'png'});await fs.writeFile(path.join(spec.previewDirectory,p.name+'.png'),new Uint8Array(await b.arrayBuffer()));}catch(e){console.log('Preview failed',p.sheet,e.message);}}}
await(await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log('EXPORTED',output,JSON.stringify(verification));
