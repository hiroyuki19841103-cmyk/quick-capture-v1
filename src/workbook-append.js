export const WORKBOOK_FILE_NAME="TASK_CONTROL_Master.xlsx";
export const WORKBOOK_MIME="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const NS="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export const QUICK_CAPTURE_TEMPLATE_SHEET="QUICK_CAPTURE_TEMPLATES";
export const DEFAULT_QUICK_CAPTURE_TEMPLATES=[
 {id:"TMP-001",sortOrder:10,name:"確認",body:"【確認】\n対象：\n確認事項：\n期限：",enabled:true,updatedAt:""},
 {id:"TMP-002",sortOrder:20,name:"現場メモ",body:"【現場メモ】\n場所：\n事実：\n気づき：",enabled:true,updatedAt:""},
 {id:"TMP-003",sortOrder:30,name:"思いつき",body:"【思いつき】\n内容：\n次に確認：",enabled:true,updatedAt:""},
 {id:"TMP-004",sortOrder:40,name:"引継ぎ",body:"【引継ぎ】\n相手：\n内容：\n期限：",enabled:true,updatedAt:""}
];

function requireZip(){if(!globalThis.JSZip)throw new Error("Excel同期部品を読み込めませんでした");return globalThis.JSZip;}
function parseXml(text,label){const doc=new DOMParser().parseFromString(text,"application/xml");if(doc.getElementsByTagName("parsererror").length)throw new Error(`${label}を解析できません`);return doc;}
function colName(index){let out="";for(let n=index+1;n>0;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out;}
function colIndex(ref){const letters=String(ref||"").match(/^[A-Z]+/)?.[0]||"";let n=0;for(const ch of letters)n=n*26+ch.charCodeAt(0)-64;return n-1;}
function textOfCell(cell,strings=[]){const type=cell.getAttribute("t")||"";if(type==="inlineStr")return[...cell.getElementsByTagNameNS("*","t")].map(n=>n.textContent||"").join("");const raw=cell.getElementsByTagNameNS("*","v")[0]?.textContent||"";return type==="s"?strings[Number(raw)]||"":raw;}
function sharedStrings(doc){return doc?[...doc.getElementsByTagNameNS("*","si")].map(si=>[...si.getElementsByTagNameNS("*","t")].map(t=>t.textContent||"").join("")):[];}
function rowMap(row,strings){const map=new Map();[...row.getElementsByTagNameNS("*","c")].forEach(cell=>map.set(colIndex(cell.getAttribute("r")),textOfCell(cell,strings)));return map;}
function aliases(record){return new Map([
  ["TaskID",record.id],["ParentID",""],["RootID",record.id],["SortOrder",record.sortOrder],
  ["分類",record.category],["タスク名",record.title],["状態",record.status],["優先度",record.priority],
  ["担当",record.owner],["開始日",record.startDate],["期日",record.dueDate],["進捗率",0],
  ["次の一手",record.nextAction],["阻害要因",""],["メモ",record.notes],["リンクURL",""],
  ["削除","FALSE"],["作成日時",record.now],["更新日時",record.now],["更新元",record.device],["版",1],
  ["想定所要時間（分）",record.expectedMinutes],["想定所要時間(分)",record.expectedMinutes],["想定所要時間",record.expectedMinutes],["ExpectedMinutes",record.expectedMinutes],
  ["次に触る日",""],["今日への配置","自動判定"],["作業分割","分割不可"],
  ["最小作業単位（分）",30],["最小作業単位(分)",30],["MinimumBlockMinutes",30],
  ["フォロー要否","FALSE"],["フォロー担当",""],["次回フォロー日",""],["フォロー想定時間（分）",0],["フォロー内容",""]
 ]);
}
function inlineCell(doc,ref,value){const c=doc.createElementNS(NS,"c");c.setAttribute("r",ref);c.setAttribute("t","inlineStr");const is=doc.createElementNS(NS,"is"),t=doc.createElementNS(NS,"t");const text=String(value??"");if(/^\s|\s$|\n/.test(text))t.setAttributeNS("http://www.w3.org/XML/1998/namespace","xml:space","preserve");t.textContent=text;is.append(t);c.append(is);return c;}
function updateRef(node,lastCol,lastRow){if(!node)return;const current=node.getAttribute("ref")||`A1:${lastCol}${lastRow}`;const start=current.split(":")[0]||"A1";node.setAttribute("ref",`${start}:${lastCol}${lastRow}`);}

async function workbookSheetLocation(zip,sheetName){
 const workbookFile=zip.file("xl/workbook.xml"),relsFile=zip.file("xl/_rels/workbook.xml.rels");if(!workbookFile||!relsFile)return"";
 const workbook=parseXml(await workbookFile.async("text"),"workbook"),sheet=[...workbook.getElementsByTagNameNS("*","sheet")].find(node=>node.getAttribute("name")===sheetName);if(!sheet)return"";
 const relationshipId=sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships","id")||sheet.getAttribute("r:id")||"",rels=parseXml(await relsFile.async("text"),"workbook relationships"),relationship=[...rels.getElementsByTagNameNS("*","Relationship")].find(node=>node.getAttribute("Id")===relationshipId),target=relationship?.getAttribute("Target")||"";
 return target?`xl/${target.replace(/^\/?xl\//,"").replace(/^\//,"")}`:"";
}

export async function readQuickCaptureTemplates(bytes){
 const zip=await requireZip().loadAsync(bytes),location=await workbookSheetLocation(zip,QUICK_CAPTURE_TEMPLATE_SHEET);if(!location)return DEFAULT_QUICK_CAPTURE_TEMPLATES.map(item=>({...item}));
 const sheetFile=zip.file(location);if(!sheetFile)return DEFAULT_QUICK_CAPTURE_TEMPLATES.map(item=>({...item}));
 const doc=parseXml(await sheetFile.async("text"),QUICK_CAPTURE_TEMPLATE_SHEET),shared=zip.file("xl/sharedStrings.xml"),strings=sharedStrings(shared?parseXml(await shared.async("text"),"共有文字列"):null),templates=[];
 [...doc.getElementsByTagNameNS("*","row")].filter(row=>Number(row.getAttribute("r")||0)>1).forEach(row=>{const map=rowMap(row,strings),item={id:String(map.get(0)||"").trim(),sortOrder:Number(map.get(1))||0,name:String(map.get(2)||"").trim(),body:String(map.get(3)||""),enabled:/^(TRUE|1|YES|ON)$/i.test(String(map.get(4)||"")),updatedAt:String(map.get(5)||"")};if(item.id&&item.name&&item.body&&item.enabled)templates.push(item);});
 return templates.sort((a,b)=>a.sortOrder-b.sortOrder||a.name.localeCompare(b.name,"ja"));
}

export async function appendTaskRecord(bytes,record){
 const zip=await requireZip().loadAsync(bytes),sheetFile=zip.file("xl/worksheets/sheet1.xml");
 if(!sheetFile)throw new Error("TASKSシートが見つかりません");
 const doc=parseXml(await sheetFile.async("text"),"TASKSシート"),shared=zip.file("xl/sharedStrings.xml"),strings=sharedStrings(shared?parseXml(await shared.async("text"),"共有文字列"):null),rows=[...doc.getElementsByTagNameNS("*","row")];
 const headerRow=rows.find(row=>Number(row.getAttribute("r"))===1);if(!headerRow)throw new Error("Excel Masterの見出しが見つかりません");
 const headers=rowMap(headerRow,strings),headerCount=Math.max(...headers.keys())+1;
 if(headers.get(0)!=="TaskID")throw new Error("TASK CONTROL用のExcel Masterではありません");
 let maxOrder=0;rows.filter(row=>Number(row.getAttribute("r"))>1).forEach(row=>{const map=rowMap(row,strings),sortIndex=[...headers].find(([,v])=>v==="SortOrder")?.[0];if(sortIndex!=null)maxOrder=Math.max(maxOrder,Number(map.get(sortIndex))||0);});
 record.sortOrder=maxOrder+1;
 const lastRow=Math.max(1,...rows.map(row=>Number(row.getAttribute("r"))||0))+1,lastCol=colName(headerCount-1),values=aliases(record),newRow=doc.createElementNS(NS,"row");newRow.setAttribute("r",String(lastRow));newRow.setAttribute("ht","34");newRow.setAttribute("customHeight","1");
 for(let i=0;i<headerCount;i++)newRow.append(inlineCell(doc,`${colName(i)}${lastRow}`,values.get(String(headers.get(i)||"").trim())??""));
 doc.getElementsByTagNameNS("*","sheetData")[0].append(newRow);updateRef(doc.getElementsByTagNameNS("*","dimension")[0],lastCol,lastRow);updateRef(doc.getElementsByTagNameNS("*","autoFilter")[0],lastCol,lastRow);
 zip.file("xl/worksheets/sheet1.xml",new XMLSerializer().serializeToString(doc));
 const tableFile=zip.file("xl/tables/table1.xml");if(tableFile){const tableDoc=parseXml(await tableFile.async("text"),"Excelテーブル");updateRef(tableDoc.documentElement,lastCol,lastRow);updateRef(tableDoc.getElementsByTagNameNS("*","autoFilter")[0],lastCol,lastRow);zip.file("xl/tables/table1.xml",new XMLSerializer().serializeToString(tableDoc));}
 return zip.generateAsync({type:"uint8array",compression:"DEFLATE",compressionOptions:{level:6}});
}
