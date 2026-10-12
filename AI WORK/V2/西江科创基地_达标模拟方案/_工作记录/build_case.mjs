import fs from 'node:fs/promises';
import path from 'node:path';
import {FileBlob,SpreadsheetFile,Workbook} from '@oai/artifact-tool';
const here=path.dirname(new URL(import.meta.url).pathname.replace(/^\/(\w:)/,'$1'));
const dir=decodeURIComponent(here);
const m=JSON.parse(await fs.readFile(path.join(dir,'case_manifest.json'),'utf8'));
if(process.argv.includes('--inspect')){
 const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(m.source));
 const result=await wb.inspect({kind:'region',sheetId:'决策看板',range:'A18:F28',maxChars:3500,tableMaxRows:11,tableMaxCols:6});
 await fs.writeFile(path.join(dir,'原V2看板检查.ndjson'),result.ndjson);
 const view=await wb.render({sheetName:'核心参数输入',range:'A5:D18',scale:1.4,format:'png'});
 await fs.writeFile(path.join(dir,'原V2参数预览.png'),new Uint8Array(await view.arrayBuffer()));
 console.log('已读取V2实际文件并生成局部预览。');
}else{
 // 原生工作簿含复杂公式及图表。只导出本次授权的输入内容，
 // 后续将这些cell payload移植到V2副本，保留原生全部样式和对象。
 const wb=Workbook.create();
 for(const [name,updates] of Object.entries(m.changes)){
  const s=wb.worksheets.add(name);
  for(const [address,value] of Object.entries(updates))s.getRange(address).values=[[value]];
 }
 wb.recalculate();
 const result=await wb.inspect({kind:'region',sheetId:'核心参数输入',range:'A95:D107',maxChars:2000});
 await fs.writeFile(path.join(dir,'新增资料检查.ndjson'),result.ndjson);
 const x=await SpreadsheetFile.exportXlsx(wb);
 await x.save(path.join(dir,'授权输入内容.xlsx'));
 console.log('授权输入内容已通过Artifact Tool写入。');
}
