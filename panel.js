"use strict";

const dateColumns = {
  P: "Nakliye Varış",
  R: "İthalat Bitiş",
  U: "Depo / Mal Kabul",
  X: "Sevkiyat",
  Z: "Boyner",
};

const statusMessage = document.querySelector("#status");
const entries = document.querySelector("#events");
const setup = document.querySelector("#setup");
const transferSection = document.querySelector("#transfer");
const instructions = document.querySelector("#instructions");
const previewButton = document.querySelector("#preview");
const sendButton = document.querySelector("#send");
const pendingCount = document.querySelector("#pending-count");
const report = document.querySelector("#report");
const reportStatus = document.querySelector("#report-status");
const reportRows = document.querySelector("#report-rows");
const channel = "bestseller-excel-panel-v1";
let pendingChange = Promise.resolve();
let busy = false;
let transferable = false;
let revision = 0;
const clearedCells = new Set();
const isStandalone = window.self === window.top;

if (isStandalone) {
  setup.hidden = false;
  statusMessage.textContent = "HTTPS paneline erişildi. Excel'de açmak için manifesti indir.";
} else if (typeof Office === "undefined") {
  statusMessage.textContent = "Office.js yüklenemedi. Microsoft Office.js adresine erişimi kontrol et.";
} else {
  const timer = window.setTimeout(() => {
    statusMessage.textContent = "Excel bağlantısı zaman aşımına uğradı. Paneli yeniden aç.";
  }, 10000);
  Office.onReady(async ({ host }) => {
    window.clearTimeout(timer);
    if (host !== Office.HostType.Excel) {
      statusMessage.textContent = "Paneli Excel içinde aç.";
      return;
    }
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load(["id", "name"]);
        await context.sync();
        if (!/^(sheet|sayfa)[123]$/i.test(sheet.name)) {
          statusMessage.textContent = "Sheet1, Sheet2 veya Sheet3'e geçip paneli yeniden aç.";
          return;
        }
        const sheetId = sheet.id;
        const sheetName = sheet.name;
        sheet.onChanged.add((event) => {
          pendingChange = pendingChange.then(() => showChange(event, sheetId, sheetName));
          return pendingChange;
        });
        await context.sync();
        transferSection.hidden = false;
        transferable = /^(sheet|sayfa)1$/i.test(sheetName);
        if (!transferable) {
          instructions.textContent = "Burası tarih algılama testi için. Aktarım için Sheet1'de paneli yeniden aç.";
        }
        previewButton.addEventListener("click", (event) => {
          if (event.isTrusted) submitBatch(sheetId, "preview");
        });
        sendButton.addEventListener("click", (event) => {
          if (event.isTrusted) submitBatch(sheetId, "send");
        });
        statusMessage.textContent = sheetName + " izleniyor.";
      });
    } catch (error) {
      statusMessage.textContent = "Excel hatası: " + (error.code || error.message);
    }
  });
}

async function showChange(event, sheetId, sheetName) {
  if (event.source !== Excel.EventSource.local || event.changeType !== Excel.DataChangeType.rangeEdited) return;

  const address = event.address.split("!").pop().replaceAll("$", "").toUpperCase();
  const match = /^([A-Z]{1,3})([1-9]\d*)(?::([A-Z]{1,3})([1-9]\d*))?$/.exec(address);
  if (!match) return;
  const [, firstColumn, firstRow, lastColumn = firstColumn, lastRow = firstRow] = match;
  const columnIndex = (letters) => [...letters].reduce((index, letter) => index * 26 + letter.charCodeAt(0) - 64, 0);
  const left = columnIndex(firstColumn);
  const right = columnIndex(lastColumn);
  const columns = Object.keys(dateColumns).filter((column) => columnIndex(column) >= left && columnIndex(column) <= right);
  if (!columns.length) return;

  try {
    if (firstRow === lastRow && firstColumn === lastColumn &&
        event.details && event.details.valueBefore === event.details.valueAfter) return;

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetId);
      const ranges = columns.map((column) => sheet.getRange(`${column}${firstRow}:${column}${lastRow}`).load(["values", "text"]));
      await context.sync();
      for (const [index, range] of ranges.entries()) {
        const column = columns[index];
        range.values.forEach(([value], offset) => updateDateEntry(
          column + (Number(firstRow) + offset), dateColumns[column], value, range.text[offset][0], sheetName,
        ));
      }
    });
  } catch (error) {
    statusMessage.textContent = "Hücre okunamadı: " + (error.code || error.message);
  }
}

function updateDateEntry(cell, label, value, displayed, sheetName) {
  const existingItem = entries.querySelector(`[data-cell="${cell}"]`);
  if (value === "" || value === null) {
    if (existingItem) {
      existingItem.remove();
      clearedCells.add(cell);
      noteQueueChange();
      statusMessage.textContent = `${sheetName} izleniyor. ${cell} panelden kaldırıldı.`;
    }
    return;
  }

  const dateText = typeof value === "string" ? validTextDate(value) : "";
  const numericDate = typeof value === "number" && Number.isFinite(value) && value > 0 &&
    /^\s*\d{1,4}([./-])\d{1,2}\1\d{1,4}(?:[\sT]|$)/.test(displayed);
  if (!dateText && !numericDate) {
    if (existingItem) {
      existingItem.remove();
      clearedCells.add(cell);
      noteQueueChange();
    }
    statusMessage.textContent = cell + " değişti. Tarihi GG.AA.YYYY SS:DD biçiminde gir.";
    return;
  }
  const entryText = cell + " · " + label + " · " + (dateText || displayed);
  if (existingItem?.textContent === entryText) return;

  statusMessage.textContent = sheetName + " izleniyor. Tarih yakalandı.";
  const item = existingItem ?? document.createElement("li");
  item.dataset.cell = cell;
  item.textContent = entryText;
  entries.prepend(item);
  clearedCells.delete(cell);
  noteQueueChange();
}

function noteQueueChange() {
  revision++;
  if (!busy) report.hidden = true;
  updateTransferButtons();
}

function updateTransferButtons() {
  const count = entries.children.length;
  pendingCount.textContent = String(count);
  previewButton.disabled = sendButton.disabled = busy || !transferable || !count;
}

async function submitBatch(sheetId, action) {
  if (busy || !transferable) return;
  busy = true;
  updateTransferButtons();
  report.hidden = false;
  reportRows.replaceChildren();
  reportStatus.textContent = "Excel'deki son değerler kontrol ediliyor…";
  try {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.10")) {
      throw new Error("Bu Excel ortamında B sütunundaki açıklamalar okunamıyor.");
    }
    await pendingChange;
    if (!entries.children.length) throw new Error("Bekleyen tarih yok.");
    const connection = await askExtension(null, "connect");
    if (!connection.connected) throw new Error(connection.error || "Chrome uzantısı bağlanamadı.");
    await pendingChange;
    const cells = [...entries.children].map((item) => item.dataset.cell);
    if (!cells.length) throw new Error("Bekleyen tarih yok.");
    const snapshot = revision;
    const prepared = await prepareExcelBatch(sheetId, cells, [...clearedCells]);
    await pendingChange;
    if (snapshot !== revision) throw new Error("Hazırlık sırasında tarihler değişti. İşlemi tekrar başlat.");
    for (const cell of prepared.missing) {
      entries.querySelector(`[data-cell="${cell}"]`)?.remove();
      clearedCells.add(cell);
    }
    for (const reason of prepared.reasons) addReportLine(reason);
    const rows = prepared.rows;
    if (!rows.length) throw new Error("Aktarılabilecek geçerli satır bulunamadı.");

    const batch = {protocol:rows.some((row) => row.mergeMembers) ? "bestseller-sheet2-v2" : "bestseller-sheet2-v1",
      id:crypto.randomUUID(), mode:"all", rows};
    reportStatus.textContent = action === "preview"
      ? "Bestseller eşleşmeleri kontrol ediliyor…"
      : "Bestseller gönderim onayını bekliyor. Gerekirse Bestseller sekmesine geç.";
    const reply = await askExtension(batch, action);
    if (reply.error) throw new Error(reply.error);
    reportStatus.textContent = reply.summary;
    for (const result of reply.results ?? []) {
      const mode = MODE_SETTINGS[result.mode];
      const label = {
        preview:"Hazır", added:"Gönderildi", existing:"Zaten var", skipped:"Tekrar",
        mismatch:"Adet uyuşmuyor", error:"Kontrol gerekli",
      }[result.status] ?? result.status;
      addReportLine(`${mode.column}${result.row} · ${result.shipment} / ${result.file} · ${label} · ${result.info}`);
    }
    if (action === "send" && snapshot === revision) {
      const selectedCells = new Set(cells);
      for (const result of reply.results ?? []) {
        if (!["added", "existing", "skipped"].includes(result.status)) continue;
        const row = rows.find((item) => item.row === result.row && item.mode === result.mode);
        if (!row) continue;
        const members = row.mergeMembers ?? [row];
        for (const member of members) {
          const cell = MODE_SETTINGS[row.mode].column + member.row;
          if (selectedCells.has(cell)) entries.querySelector(`[data-cell="${cell}"]`)?.remove();
        }
      }
    }
  } catch (error) {
    reportStatus.textContent = error.message;
  } finally {
    busy = false;
    updateTransferButtons();
  }
}

function addReportLine(text) {
  const line = document.createElement("li");
  line.textContent = text;
  reportRows.append(line);
}

function askExtension(batch, action) {
  const requestId = crypto.randomUUID();
  return new Promise((resolve) => {
    const receive = (event) => {
      const data = event.data;
      if (event.source !== window || event.origin !== location.origin ||
          data?.channel !== channel || data.type !== "response" || data.requestId !== requestId) return;
      window.clearTimeout(timer);
      window.removeEventListener("message", receive);
      resolve(data.response);
    };
    const timer = window.setTimeout(() => {
      window.removeEventListener("message", receive);
      resolve({error:action === "connect"
        ? "Chrome uzantısını yükle ve Excel sekmesini yenile."
        : "Chrome uzantısından yanıt gelmedi. Sonucu Bestseller panelinde kontrol et."});
    }, action === "connect" ? 3000 : 240000);
    window.addEventListener("message", receive);
    window.postMessage({channel, type:"request", requestId, batch, action}, location.origin);
  });
}

function validTextDate(value) {
  const text = value.trim();
  const dayFirst = /^(\d{1,2})([./-])(\d{1,2})\2(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(text);
  const iso = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(text);
  const parts = dayFirst
    ? [dayFirst[4], dayFirst[3], dayFirst[1], ...dayFirst.slice(5)]
    : iso?.slice(1);
  if (!parts) return "";

  const [year, month, day, hour, minute, second] = parts.map((part) => Number(part ?? 0));
  if (year < 1900 || month < 1 || month > 12 || hour > 23 || minute > 59 || second > 59) return "";
  const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
  return day >= 1 && day <= daysInMonth ? text : "";
}
