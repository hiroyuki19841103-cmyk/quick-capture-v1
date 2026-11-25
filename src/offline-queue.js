const DB = "quick-capture-offline";
const STORE = "outbox";
let opening;
function database() {
  if (!opening) opening = new Promise((resolve, reject) => {
    const request = indexedDB.open(DB, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE, {keyPath:"id"});
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    request.onblocked = () => reject(new Error("保存領域を開けません。別のタブを閉じてください"));
  }).catch(error => { opening = null; throw error; });
  return opening;
}
async function transaction(mode, action) {
  const db = await database();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, mode);
    const req = action(tx.objectStore(STORE));
    tx.oncomplete = () => resolve(req.result);
    tx.onabort = tx.onerror = () => reject(tx.error || new Error("端末内保存に失敗しました"));
  });
}
export const putRecord = record => transaction("readwrite", store => store.put(record));
export const removeRecord = id => transaction("readwrite", store => store.delete(id));
export const listRecords = async () => (await transaction("readonly", store => store.getAll())).sort((a,b) => String(a.now).localeCompare(String(b.now)) || a.id.localeCompare(b.id));
export async function migrateLegacy(key) {
  const raw = localStorage.getItem(key);
  if (!raw) return;
  const records = JSON.parse(raw);
  if (!Array.isArray(records)) throw new Error("旧未送信データを読み込めません");
  const db = await database();
  await new Promise((resolve,reject) => {
    const tx = db.transaction(STORE,"readwrite"), store = tx.objectStore(STORE);
    for (const record of records) {
      const req = store.get(record.id);
      req.onsuccess = () => { if (!req.result) store.put(record); };
    }
    tx.oncomplete = resolve;
    tx.onerror = tx.onabort = () => reject(tx.error);
  });
  localStorage.removeItem(key);
}
