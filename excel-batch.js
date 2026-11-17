"use strict";

const columnModes = Object.fromEntries(Object.entries(MODE_SETTINGS).map(([mode, settings]) => [settings.column, mode]));

async function prepareExcelBatch(sheetId, cells, clearedCells = [], discardedCells = []) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetId);
    const selected = cells.map((cell) => ({cell, row:Number(cell.slice(1)), mode:columnModes[cell[0]]}));
    const selectedRows = new Set(selected.map((item) => item.row));
    const used = sheet.getRange("B:B").getUsedRangeOrNullObject(true);
    const comments = sheet.comments;
    used.load(["rowIndex", "rowCount"]);
    comments.load("items");
    const directRanges = groupNearbyRows([...selectedRows].sort((a,b) => a-b))
      .map((rows) => ({first:rows[0], range:sheet.getRange(`A${rows[0]}:Z${rows.at(-1)}`).load(["values", "text"])}));
    await context.sync();

    const sheetRows = new Map();
    for (const part of directRanges) collectReadRows(sheetRows, part.first, part.range);

    const preview = [];
    const transfer = [];
    const reasons = [];
    const missing = [];
    for (const item of selected) {
      const {cell, row, mode} = item;
      const {values, texts} = sheetRows.get(row);
      const index = MODE_SETTINGS[mode].index;
      const date = parseDate(values[index], texts[index], mode);
      if (!date) {
        if (values[index] === "" || values[index] == null) missing.push(cell);
        else reasons.push(cell + " · Geçerli işlem tarihi yok.");
        continue;
      }
      const a = texts[0].trim(), b = texts[1].trim(), file = texts[4].trim(), note = texts[2].trim();
      const shipment = a + b;
      const pieces = readPieces(values[10]);
      const canTransfer = a && b && file && Number.isSafeInteger(pieces) && pieces >= 0;
      if (canTransfer) transfer.push({uuid:crypto.randomUUID(), row, mode, shipment, file, note, pieces, date});
      preview.push([`${MODE_SETTINGS[mode].label} / ${cell}`, shipment, file, note, pieces, displayDate(date),
        canTransfer ? "" : "Sevkiyat, dosya adı veya K sütunu adedi eksik."]);
    }
    if (!preview.length) return {rows:[], reasons, missing, preview};

    // Read the location of comments first. Only B-column comment contents affect shipment merges.
    const locations = comments.items.map((comment) => ({comment, range:comment.getLocation().load(["columnIndex", "rowIndex"])}));
    if (locations.length) await context.sync();
    const mergeComments = used.isNullObject ? [] : locations.filter(({range}) => range.columnIndex === 1);
    if (mergeComments.length) {
      for (const {comment} of mergeComments) comment.load("content");
      await context.sync();
    }

    const links = mergeComments.filter(({comment}) => parseMergeLink(comment.content));
    let output = transfer;
    if (links.length) {
      let first = used.rowIndex + 1, last = used.rowIndex + used.rowCount;
      for (const row of [...selectedRows, ...links.map(({range}) => range.rowIndex + 1)]) {
        first = Math.min(first, row);
        last = Math.max(last, row);
      }
      for (let offset = first; offset <= last; offset += SCAN_PAGE_SIZE) {
        const range = sheet.getRange(`A${offset}:Z${Math.min(last, offset + SCAN_PAGE_SIZE - 1)}`).load(["values", "text"]);
        await context.sync();
        collectReadRows(sheetRows, offset, range);
      }
      const excluded = new Set([...clearedCells, ...missing]);
      for (const cell of excluded) {
        const entry = sheetRows.get(Number(cell.slice(1)));
        if (!entry) continue;
        const mode = columnModes[cell[0]];
        const index = MODE_SETTINGS[mode].index;
        if (!parseDate(entry.values[index], entry.texts[index], mode)) {
          entry.values[index] = entry.texts[index] = "GEÇERSİZ";
        }
      }
      const source = makeMergeSource(sheetRows, first, last, links);
      const dates = preview.map((line) => parseTextDate(line[5], "depo").slice(0, 10)).sort();
      output = applyCommentMerges(source, dates[0], dates.at(-1), preview, transfer, excluded);
    }

    const discarded = new Set(discardedCells);
    output = output.filter((row) => {
      if (!row.mergeMembers) return true;
      const {column, label} = MODE_SETTINGS[row.mode];
      const removed = row.mergeMembers.filter((member) => discarded.has(column + member.row));
      if (!removed.length) return true;
      const removedLabels = removed.map((member) => column + member.row).join(", ");
      for (const member of row.mergeMembers) {
        const line = preview.find((item) => item[0] === `${label} / ${column}${member.row}`);
        if (line) line[6] = `Birleşim bekliyor: ${removedLabels} gönderim listesinden çıkarıldı.`;
      }
      return false;
    });

    for (const [label, , , , , , reason] of preview) {
      if (reason && !reason.startsWith("Birleşim →")) reasons.push(label + " · " + reason);
    }
    return {rows:output, reasons, missing, preview};
  });
}

function collectReadRows(rows, first, range) {
  range.values.forEach((values,i) => rows.set(first + i, {values, texts:range.text[i]}));
}

async function readExcelSyncSources(sheetId, cells) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetId);
    const byRow = new Map();
    const rowNumbers = [...new Set(cells.map(cell => Number(cell.slice(1))))].sort((a,b) => a-b);
    const parts = groupNearbyRows(rowNumbers)
      .map(rows => ({first:rows[0],range:sheet.getRange(`A${rows[0]}:Z${rows.at(-1)}`).load(["values","text"])}));
    await context.sync();
    for (const part of parts) collectReadRows(byRow,part.first,part.range);
    return cells.map(cell => {
      const {values,texts} = byRow.get(Number(cell.slice(1)));
      const mode = columnModes[cell[0]];
      const {index} = MODE_SETTINGS[mode];
      const value = values[index];
      const pieces = readPieces(values[10]);
      return {
        cell, shipment:texts[0].trim()+texts[1].trim(), note:texts[2].trim(), file:texts[4].trim(),
        pieces:Number.isSafeInteger(pieces) ? pieces : null,
        date:value === "" || value == null ? null : parseDate(value,texts[index],mode) || "invalid",
      };
    });
  });
}

function makeMergeSource(rows, first, last, links) {
  return {
    getComments: () => links.map(({comment,range}) => ({
      getLocation: () => ({getAddress: () => "Sheet1!B" + (range.rowIndex + 1)}),
      getContent: () => comment.content,
    })),
    getRange(address) {
      if (address === "B:B") {
        return {getUsedRange: () => ({getRowIndex: () => first - 1, getRowCount: () => last - first + 1})};
      }
      const [, left, start, right, end] = /^([A-Z])(\d+):([A-Z])(\d+)$/.exec(address);
      const from = left.charCodeAt(0) - 65;
      const to = right.charCodeAt(0) - 64;
      const slice = (key) => Array.from({length:Number(end) - Number(start) + 1},
        (_,i) => rows.get(Number(start) + i)[key].slice(from, to));
      return {getValues: () => slice("values"), getTexts: () => slice("texts")};
    },
  };
}
