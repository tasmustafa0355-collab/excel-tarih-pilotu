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
        let pendingChange = Promise.resolve();
        sheet.onChanged.add((event) => {
          pendingChange = pendingChange.then(() => showChange(event, sheetId, sheetName));
          return pendingChange;
        });
        await context.sync();
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
  const match = /^(([PRUXZ])[1-9]\d*)(?::\1)?$/.exec(address);
  if (!match) return;
  const [, cell, column] = match;
  const label = dateColumns[column];

  try {
    const existingItem = entries.querySelector(`[data-cell="${cell}"]`);
    const detail = event.details;
    if (!existingItem && detail && detail.valueBefore === detail.valueAfter) return;

    await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem(sheetId).getRange(cell);
      range.load(["values", "text"]);
      await context.sync();
      const value = range.values[0][0];
      const displayed = range.text[0][0];
      if (value === "" || value === null) {
        if (existingItem) {
          existingItem.remove();
          statusMessage.textContent = `${sheetName} izleniyor. ${cell} panelden kaldırıldı.`;
        }
        return;
      }

      const dateText = typeof value === "string" ? validTextDate(value) : "";
      const numericDate = typeof value === "number" && Number.isFinite(value) && value > 0 &&
        /^\s*\d{1,4}([./-])\d{1,2}\1\d{1,4}(?:[\sT]|$)/.test(displayed);
      if (!dateText && !numericDate) {
        existingItem?.remove();
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
      if (entries.children.length > 15) entries.lastElementChild.remove();
    });
  } catch (error) {
    statusMessage.textContent = "Hücre okunamadı: " + (error.code || error.message);
  }
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
