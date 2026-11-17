"use strict";

const autoButton = document.querySelector("#auto-sync");
const autoStatus = document.querySelector("#sync-status");
const automaticCells = new Set();
let automaticScope = "";
let autoTimer;
let autoVersion = 0;
let changingAuto = false;

async function workbookSyncScope(sheetId) {
  const raw = Office.context.document.url;
  if (!raw) throw new Error("Excel bu çalışma kitabının adresini vermedi.");
  const url = new URL(raw);
  const sourceId = [...url.searchParams].find(([key]) => key.toLowerCase() === "sourcedoc")?.[1];
  const guid = /^\{?([\da-f]{8}-[\da-f]{4}-[\da-f]{4}-[\da-f]{4}-[\da-f]{12})\}?$/i.exec(sourceId ?? "");
  const filePath = /\.(?:xlsx|xlsm|xlsb)$/i.test(url.pathname);
  if (url.protocol !== "https:" || (!guid && !filePath)) {
    throw new Error("Çalışma kitabının SharePoint veya OneDrive adresi doğrulanamadı.");
  }
  const documentKey = guid ? guid[1].toLowerCase() : decodeURIComponent(url.pathname).toLowerCase();
  const bytes = new TextEncoder().encode(JSON.stringify([url.origin,documentKey,sheetId]));
  const digest = await crypto.subtle.digest("SHA-256",bytes);
  return Array.from(new Uint8Array(digest),value => value.toString(16).padStart(2,"0")).join("");
}

async function toggleAutomaticTransfer(sheetId) {
  if (changingAuto || (!automaticScope && (busy || !transferable))) return;
  changingAuto = true;
  autoStatus.hidden = false;
  updateTransferButtons();
  try {
    if (automaticScope) {
      const scope = automaticScope;
      automaticScope = "";
      automaticCells.clear();
      autoVersion++;
      window.clearTimeout(autoTimer);
      autoButton.setAttribute("aria-pressed","false");
      autoButton.textContent = "Otomatik aktarımı aç";
      autoStatus.textContent = "Otomatik aktarım kapalı.";
      await askExtension({scope},"sync-disable");
      return;
    }
    if (!Office.context.requirements.isSetSupported("ExcelApi","1.10")) {
      throw new Error("Bu Excel ortamında birleşim yorumları okunamıyor.");
    }
    const scope = await workbookSyncScope(sheetId);
    const response = await askExtension({scope},"sync-enable");
    if (!response?.enabled) throw new Error(response?.error || "Otomatik aktarım açılamadı.");
    automaticScope = scope;
    autoButton.setAttribute("aria-pressed","true");
    autoButton.textContent = "Otomatik aktarımı kapat";
    autoStatus.textContent = "Otomatik aktarım açık. Bundan sonraki tarih girişleri ve silmeler izleniyor.";
  } catch (error) {
    autoStatus.textContent = error.message;
  } finally {
    changingAuto = false;
    updateTransferButtons();
  }
}

function trackAutomaticDate(cell,sheetId) {
  if (!automaticScope) return;
  automaticCells.add(cell);
  autoVersion++;
  scheduleAutomaticTransfer(sheetId);
}

function scheduleAutomaticTransfer(sheetId) {
  window.clearTimeout(autoTimer);
  if (!automaticScope || busy || !automaticCells.size) return;
  autoTimer = window.setTimeout(() => transferAutomaticDates(sheetId),650);
}

async function prepareAutomaticDates(sheetId,cells,groups) {
  const prepared = await prepareExcelBatch(sheetId,cells);
  const rowSources = new Map(prepared.rows.map(row => [
    row,(row.mergeMembers ?? [row]).map(member => MODE_SETTINGS[row.mode].column + member.row),
  ]));
  const allCells = [...new Set([...cells,...groups.flat(),...[...rowSources.values()].flat()])];
  const observations = new Map((await readExcelSyncSources(sheetId,allCells)).map(source => [source.cell,source]));
  const processed = new Set();
  const actions = [];
  for (const [row,members] of rowSources) {
    const changed = cells.filter(cell => members.includes(cell));
    if (!changed.length) continue;
    changed.forEach(cell => processed.add(cell));
    actions.push({row,changed,sources:members.map(cell => observations.get(cell))});
  }
  for (const cell of cells) {
    if (processed.has(cell) || observations.get(cell).date !== null) continue;
    const members = groups.find(group => group.includes(cell));
    if (!members) continue;
    const changed = cells.filter(item => members.includes(item) && !processed.has(item));
    changed.forEach(item => processed.add(item));
    actions.push({row:null,changed,sources:members.map(item => observations.get(item))});
  }
  return {actions,reasons:prepared.reasons};
}

async function transferAutomaticDates(sheetId) {
  if (!automaticScope || !automaticCells.size || busy) return;
  busy = true;
  updateTransferButtons();
  report.hidden = false;
  reportStatus.textContent = "Otomatik aktarım için Excel tarihleri kontrol ediliyor…";
  const scope = automaticScope;
  let cells = [];
  let version;
  try {
    window.clearTimeout(detailsTimer);
    await detailsRead;
    await pendingChange;
    if (scope !== automaticScope) return;
    version = autoVersion;
    cells = [...automaticCells];
    const tracked = await askExtension({scope,cells},"sync-inspect");
    if (!Array.isArray(tracked?.groups)) throw new Error(tracked?.error || "Otomatik kayıtlar okunamadı.");
    const {actions,reasons} = await prepareAutomaticDates(sheetId,cells,tracked.groups);
    await pendingChange;
    if (scope !== automaticScope || version !== autoVersion) return;
    for (const reason of reasons) addAutomaticLine(reason);
    if (!actions.length && !reasons.length) addAutomaticLine(cells.join(", ") + " · Aktarılabilir tarih bulunamadı.");
    for (const operation of actions) {
      if (scope !== automaticScope || version !== autoVersion) break;
      for (const cell of operation.changed) automaticCells.delete(cell);
      const result = await askExtension({scope,...operation},"sync");
      const label = {
        added:"Eklendi",updated:"Güncellendi",cleared:"Silindi",unchanged:"Güncel",retained:"Korundu",
        existing:"Zaten var",ignored:"İşlem yok",error:"Kontrol gerekli",
      }[result?.status] || "Kontrol gerekli";
      addAutomaticLine(operation.changed.join(", ") + " · " + label + " · " +
        (result?.info || result?.error || "Bestseller yanıtı alınamadı."));
      if (["added","updated","cleared","unchanged","existing"].includes(result?.status)) {
        for (const cell of operation.changed) {
          if (!automaticCells.has(cell)) entries.querySelector(`[data-cell="${cell}"]`)?.remove();
        }
      }
      await pendingChange;
    }
    reportStatus.textContent = "Otomatik aktarım sonuçları";
  } catch (error) {
    reportStatus.textContent = error.message;
  } finally {
    if (version === autoVersion) for (const cell of cells) automaticCells.delete(cell);
    busy = false;
    updateTransferButtons();
    scheduleQueueDetails(sheetId);
    scheduleAutomaticTransfer(sheetId);
  }
}

function addAutomaticLine(text) {
  addReportLine(text);
  while (reportRows.children.length > 50) reportRows.firstElementChild.remove();
  reportRows.lastElementChild?.scrollIntoView({block:"nearest"});
}
