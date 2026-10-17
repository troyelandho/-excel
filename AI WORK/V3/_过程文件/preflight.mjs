import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile,Workbook} from '@oai/artifact-tool';
const root=new URL('.',import.meta.url);
const path=decodeURIComponent(new URL('../../V2/产业园投资测算_V2_样式调整.xlsx',root).pathname.slice(1));
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(path));
console.log((await w.inspect({kind:'sheet',include:'id,name',maxChars:2000})).ndjson);
const im=await w.render({sheetName:'决策看板',range:'A8:G28',scale:1.2,format:'png'});
await fs.writeFile(new URL('V2_参考.png',root),new Uint8Array(await im.arrayBuffer()));
console.log(w.help('worksheet.protection',{include:'index,examples,notes',maxChars:4000}).ndjson);
