import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const base=new URL('.',import.meta.url);
const source=new URL('../V2/产业园投资测算_V2_样式调整.xlsx',base);
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(decodeURIComponent(source.pathname.slice(1))));
console.log((await wb.inspect({kind:'sheet',include:'id,name',maxChars:2500})).ndjson);
for(const [sheet,range,name] of [['决策看板','A1:H33','dashboard'],['核心参数输入','A5:D32','inputs'],['利润与现金流','A4:L19','cashflow']]){
 try{
  const blob=await wb.render({sheetName:sheet,range,scale:1.5,format:'png'});
  await fs.writeFile(new URL(name+'.png',base),new Uint8Array(await blob.arrayBuffer()));
  console.log('rendered',name);
 }catch(err){console.log('render_error',name,String(err))}
}
