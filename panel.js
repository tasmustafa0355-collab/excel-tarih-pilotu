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
        if (!/^(sheet|sayfa)[23]$/i.test(sheet.name)) {
          statusMessage.textContent = "Sheet2 veya Sheet3'e geçip paneli kapatıp yeniden aç.";
          return;
        }
        const sheetId = sheet.id;
        const sheetName = sheet.name;
        sheet.onChanged.add((event) => showChange(event, sheetId));
        await context.sync();
        statusMessage.textContent = sheetName + " izleniyor.";
      });
    } catch (error) {
      statusMessage.textContent = "Excel hatası: " + (error.code || error.message);
    }
  });
}

async function showChange(event, sheetId) {
  if (event.source !== Excel.EventSource.local || event.changeType !== Excel.DataChangeType.rangeEdited) return;

  const cell = event.address.split("!").pop().replaceAll("$", "");
  const column = /^([A-Z]+)[1-9]\d*$/.exec(cell)?.[1];
  const label = Object.hasOwn(dateColumns, column) ? dateColumns[column] : "";
  if (!label) return;

  try {
    const detail = event.details;
    if (!detail || detail.valueBefore === detail.valueAfter) return;

    await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem(sheetId).getRange(cell);
      range.load(["values", "text"]);
      await context.sync();
      const value = range.values[0][0];
      const displayed = range.text[0][0];
      const showsDate = /^\s*\d{1,4}([./-])\d{1,2}\1\d{1,4}(?:[ T]|$)/.test(displayed);
      if (value !== detail.valueAfter || typeof value !== "number" || value <= 0 || !showsDate) {
        return;
      }
      const item = document.createElement("li");
      item.textContent = cell + " · " + label + " · " + displayed;
      entries.prepend(item);
      if (entries.children.length > 15) entries.lastElementChild.remove();
    });
  } catch (error) {
    statusMessage.textContent = "Hücre okunamadı: " + (error.code || error.message);
  }
}
