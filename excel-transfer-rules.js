"use strict";

// Existing Bestseller SharePoint transfer rules shared with the Office Script.
const SCAN_PAGE_SIZE = 2000;

const MODE_SETTINGS = {
  depo: { column: "U", index: 20, label: "Depo" },
  nakliye: { column: "X", index: 23, label: "Sevkiyat" },
  boyner: { column: "Z", index: 25, label: "Boyner" },
  nakliyeVaris: { column: "P", index: 15, label: "Nakliye Varış" },
  ithalatBitis: { column: "R", index: 17, label: "İthalat Bitiş" }
};


function applyCommentMerges(source, start, end, preview, transfer, clearedCells) {
  if (!preview.length) return transfer;
  const modes= ["depo", "nakliye", "boyner", "nakliyeVaris", "ithalatBitis"];
  const modeRows = new Map ();
  const selectedRows = new Set ();
  for (const mode of modes) modeRows.set(mode, new Map ());
  for (const line of preview) {
    const label = String(line[0]);
    const match = /\d+$/.exec(label);
    const mode = modes.find(key => label.startsWith(MODE_SETTINGS[key].label + " / "));
    if (!mode || !match) continue;
    const row = Number(match[0]);
    modeRows.get(mode)?.set(row, line);
    const day = parseTextDate(String(line[5]), mode).slice(0, 10);
    if (day >= start && day <= end) selectedRows.add(row);
  }
  const links = readDatedMergeLinks(source, start, end, selectedRows, modes, clearedCells);
  if (!links.size) return transfer;

  const wanted = new Set(Array.from(links.values()).map(link => link.key));
  const sourceKeys = new Set ();
  for (const rows of nearby(links.keys())) {
    const texts = source.getRange(`A${rows[0]}:E${rows[rows.length - 1]}`).getTexts();
    for (const row of rows) {
      const text = texts[row - rows[0]];
      sourceKeys.add(commentKey(text[0].trim() + text[1].trim() + text[4].trim()));
    }
  }
  const identities = new Map ();
  const targets = new Map ();
  const duplicates = new Map ();
  const used = source.getRange("B:B").getUsedRange(true);
  if (!used) return transfer;
  const first = used.getRowIndex() + 1;
  const last = first + used.getRowCount() - 1;
  for (let start = first; start <= last; start += SCAN_PAGE_SIZE) {
    const end = Math.min(start + SCAN_PAGE_SIZE - 1, last);
    const texts = source.getRange(`A${start}:E${end}`).getTexts();
    for (let offset = 0; offset < texts.length; offset++) {
      const row = start + offset;
      const text = texts[offset];
      const a = text[0].trim(), b = text[1].trim(), file = text[4].trim(), note = text[2].trim();
      if (!a || !b) continue;
      const shipment = a + b;
      const ship = commentKey(shipment);
      const key = commentKey(shipment + file);
      const targetKey = mergeTitleKey(shipment + file);
      const member= {row, shipment, file, note};
      if (links.has(row)) identities.set(row, member);
      if (file && wanted.has(ship)) indexMember(targets, ship, member);
      if (file && targetKey !== ship && wanted.has(targetKey)) indexMember(targets, targetKey, member);
      const suffix = isInspection(note) ? "-DENETİM" : isChemical(note) ? "-Kimyahane" : "";
      if (file && suffix && !fileRefKey(file).endsWith(fileRefKey(suffix))) {
        const categoryKey = mergeTitleKey(shipment + file + suffix);
        if (wanted.has(categoryKey)) indexMember(targets, categoryKey, member);
      }
      if (sourceKeys.has(key)) indexMember(duplicates, key, member);
    }
  }

  const relevant = new Set ();
  for (const [row, link] of Array.from(links.entries())) {
    const candidates = targets.get(link.key) || [];
    if (!modes.some(mode => modeRows.get(mode)?.has(row) || candidates.some(member => modeRows.get(mode)?.has(member.row)))) continue;
    relevant.add(row);
    for (const member of candidates) relevant.add(member.row);
    const sourceRow = identities.get(row);
    if (sourceRow) {
      const sameSource = duplicates.get(commentKey(sourceRow.shipment + sourceRow.file)) || [];
      for (const member of sameSource) relevant.add(member.row);
    }
  }
  const data = new Map ();
  for (const rows of nearby(relevant.values())) {
    const start = rows[0];
    const {values, texts} = readRows(source, start, rows[rows.length - 1], "Z", "K");
    for (const row of rows) data.set(row, {values:values[row - start], texts:texts[row - start]});
  }

  const suppressed = new Set ();
  const merged= [];
  for (const mode of modes) {
    const lines = modeRows.get(mode);
    if (!lines || !lines.size) continue;
    const groups = new Map ();
    const blocked = new Set ();
    const otherDates = new Map ();
    const mark = (row, message)=> {
      blocked.add(row);
      suppressed.add(`${mode}/${row}`);
      const line = lines.get(row);
      if (line) line[6] = message;
    };
    const dateFor = (row)=> {
      const entry = data.get(row);
      const col = MODE_SETTINGS[mode].index - 10;
      if (!entry || String(entry.values[col]).trim() === "") return "";
      return parseDate(entry.values[col], entry.texts[col], mode) || "INVALID";
    };
    for (const [row, link] of Array.from(links.entries())) {
      const found = targets.get(link.key) || [];
      if (!lines.has(row) && !found.some(member => lines.has(member.row))) continue;
      const sourceRow = identities.get(row);
      const candidates = found.filter(member => member.row !== row);
      const refs = new Set(found.map(member => fileRefKey(member.file)));
      const sourceDate = dateFor(row);
      const dated = candidates.filter(member => {
        const date = dateFor(member.row);
        return !!sourceDate && sourceDate !== "INVALID" && date !== "INVALID" &&
          date.slice(0, 10) === sourceDate.slice(0, 10);
      });
      let target = dated.length === 1 ? dated[0] : undefined;
      // Keep the existing explicit-file rule when one of its two dates is blank.
      if (!link.shipmentOnly && candidates.length === 1 && !target && (!sourceDate || !dateFor(candidates[0].row))) {
        const candidateDate = dateFor(candidates[0].row);
        if (candidateDate !== "INVALID" && sourceDate !== "INVALID") target = candidates[0];
      }
      let reason = "";
      if (link.shipmentOnly && refs.size > 1) reason = "E sütununda birden fazla dosya ref’i var. Manuel kontrol gerekli.";
      else if (!sourceRow || !candidates.length) reason = "Birleşim hedefi veya kaynağı tekil olarak doğrulanamadı. Manuel kontrol gerekli.";
      else if (dated.length > 1) reason = "Aynı işlem tarihinde birden fazla hedef satır var. Manuel kontrol gerekli.";
      else if (!target) reason = "Birleşimde işlem tarihleri farklı, boş veya geçersiz. Manuel kontrol gerekli.";
      if (reason || !sourceRow || !target) {
        mark(row, reason);
        for (const member of candidates) mark(member.row, reason);
        continue;
      }
      const sameSource = duplicates.get(commentKey(sourceRow.shipment + sourceRow.file)) || [];
      const sourceFile = sourceFileKey(sourceRow);
      const sourcePieces = readPieces(data.get(row)?.values[0]);
      const sourceDay = (sourceDate || dateFor(target.row)).slice(0, 10);
      const sourceClashes = sameSource.filter(member => member.row !== row && member.row !== target.row &&
        !candidates.some(candidate => candidate.row === member.row) &&
        sourceFileKey(member) === sourceFile && dateFor(member.row).slice(0, 10) === sourceDay &&
        readPieces(data.get(member.row)?.values[0]) === sourcePieces);
      const sameFile = commentKey(sourceRow.shipment) === commentKey(target.shipment) &&
        sourceFile === sourceFileKey(target);
      if (sourceClashes.length || (sameFile && !isInspection(sourceRow.note))) {
        const otherRows = sourceClashes.length ? sourceClashes : [target];
        const addresses = otherRows.map(member => MODE_SETTINGS[mode].column + member.row).join(", ");
        const message = `Birleşimin kaynak satırı Excel’de tekrarlı (${addresses}). Manuel kontrol gerekli.`;
        for (const member of [sourceRow, target, ...sourceClashes]) mark(member.row, message);
        continue;
      }
      for (const candidate of candidates) {
        if (candidate.row === target.row) continue;
        const others = otherDates.get(target.row) || new Set ();
        others.add(candidate.row);
        otherDates.set(target.row, others);
        const message = `İşlem tarihi birleşim hedefiyle farklı. Eşleşen satır: ${target.row}. Manuel kontrol gerekli.`;
        suppressed.add(`${mode}/${candidate.row}`);
        const line = lines.get(candidate.row);
        if (line && !blocked.has(candidate.row)) line[6] = message;
      }
      const namedTarget = link.targetSuffix && !fileRefKey(target.file).endsWith(fileRefKey(link.targetSuffix))
        ? {...target, file:target.file + link.targetSuffix} : target;
      const group = groups.get(target.row) || {target:namedTarget, sources:[]};
      if (fileRefKey(group.target.file) !== fileRefKey(namedTarget.file)) {
        const message = "Aynı hedef satır için farklı dosya başlıkları belirtilmiş. Manuel kontrol gerekli.";
        for (const member of [target, sourceRow, ...group.sources]) mark(member.row, message);
        continue;
      }
      group.sources.push(sourceRow);
      groups.set(target.row, group);
    }

    for (const group of Array.from(groups.values())) {
      if (!group.sources.some(member => groups.has(member.row))) continue;
      for (const member of [group.target, ...group.sources]) mark(member.row, "Birleşim zinciri var. Manuel kontrol gerekli.");
    }
    const owners = new Map ();
    for (const group of Array.from(groups.values())) {
      const usedKeys = new Set ();
      for (const member of [group.target, ...group.sources]) {
        const sameTarget = member !== group.target && isInspection(member.note) &&
          commentKey(member.shipment) === commentKey(group.target.shipment) &&
          (!member.file || fileRefKey(member.file) === fileRefKey(group.target.file));
        if (sameTarget) continue;
        const key = `${commentKey(member.shipment)}/${fileRefKey(member.file)}`;
        const owner = owners.get(key);
        if (usedKeys.has(key) || (owner && owner !== group)) {
          const message = "Aynı dosya ref’i aynı işlemde birden fazla kez birleşiyor. Manuel kontrol gerekli.";
          for (const candidate of [group.target, ...group.sources]) mark(candidate.row, message);
          if (owner && owner !== group) {
            for (const candidate of [owner.target, ...owner.sources]) mark(candidate.row, message);
          }
        }
        usedKeys.add(key);
        if (!owner) owners.set(key, group);
      }
    }
    for (const group of Array.from(groups.values())) {
      const members = [group.target, ...group.sources];
      if (members.some(member => blocked.has(member.row))) {
        for (const member of members) {
          if (!blocked.has(member.row)) mark(member.row, "Bağlı birleşim satırında manuel kontrol gerekli.");
        }
        continue;
      }
      const mergeMembers = members.map(member => {
        const entry = data.get(member.row);
        return {row:member.row, shipment:member.shipment, file:member.file,
          inspection:isInspection(member.note), chemical:isChemical(member.note) || undefined, pieces:entry ? readPieces(entry.values[0]) : NaN};
      });
      const total = mergeMembers.reduce((sum, member) => sum + member.pieces, 0);
      if (!Number.isSafeInteger(total) || mergeMembers.some(member => !Number.isSafeInteger(member.pieces) || member.pieces < 0)) {
        for (const member of members) mark(member.row, "Birleşimde geçersiz adet var. Manuel kontrol gerekli.");
        continue;
      }
      const target = group.target;
      const date = dateFor(target.row) || dateFor(group.sources[0].row);
      if (!date || !members.some(member => lines.has(member.row))) continue;
      if (members.some(member => dateFor(member.row) && dateFor(member.row).slice(0, 10) !== date.slice(0, 10))) {
        for (const member of members) mark(member.row, "Birleşimde işlem tarihleri farklı. Manuel kontrol gerekli.");
        continue;
      }
      if ((mode === "nakliyeVaris" || mode === "ithalatBitis") &&
          members.some(member => dateFor(member.row) && dateFor(member.row) !== date)) {
        for (const member of members) mark(member.row, "Birleşimde işlem saatleri farklı. Manuel kontrol gerekli.");
        continue;
      }
      const outside = Array.from(otherDates.get(target.row) || []).filter(row => !members.some(member => member.row === row));
      const warning = outside.length ? ` · Diğer tarih, manuel kontrol: ${outside.join(", ")}` : "";
      const summary = `Birleşim → ${target.shipment} / ${target.file} · ${mergeMembers.map(member => member.pieces).join(" + ")} = ${total}${warning}`;
      for (const member of members) {
        suppressed.add(`${mode}/${member.row}`);
        const line = lines.get(member.row);
        if (line) line[6] = summary;
      }
      merged.push({uuid:makeId(), mode, row:target.row, shipment:target.shipment, file:target.file,
        note:target.note, pieces:total, date, mergeMembers});
    }
  }
  return [...transfer.filter(row => !suppressed.has(`${row.mode}/${row.row}`)), ...merged];
}

function readDatedMergeLinks(source, start, end, selectedRows, modes, clearedCells) {
  const comments = new Map ();
  for (const comment of source.getComments()) {
    const address = /!\$?B\$?([1-9]\d*)$/i.exec(comment.getLocation().getAddress());
    if (address) comments.set(Number(address[1]), comment);
  }
  // Use dates already checked during preview. For other commented rows, read
  // only their date columns, in batches, before fetching any comment text.
  const uncheckedRows = Array.from(comments.keys()).filter(row => !selectedRows.has(row)).sort((left, right) => left - right);
  const arrivalIndex = MODE_SETTINGS.nakliyeVaris.index;
  let batchStart = 0;
  while (batchStart < uncheckedRows.length) {
    const first = uncheckedRows[batchStart];
    let batchEnd = batchStart + 1;
    while (batchEnd < uncheckedRows.length && uncheckedRows[batchEnd] < first + SCAN_PAGE_SIZE) batchEnd++;
    const {values, texts} = readRows(source, first, uncheckedRows[batchEnd - 1], "Z", "P");
    for (let index = batchStart; index < batchEnd; index++) {
      const row = uncheckedRows[index];
      const offset = row - first;
      const inRange = modes.some(mode => {
        const column = MODE_SETTINGS[mode].index - arrivalIndex;
        const day = parseDate(values[offset][column], texts[offset][column], mode).slice(0, 10);
        return day >= start && day <= end;
      });
      if (inRange) selectedRows.add(row);
    }
    batchStart = batchEnd;
  }
  const links = new Map ();
  for (const [row, comment] of Array.from(comments.entries())) {
    if (!selectedRows.has(row) && !modes.some((mode) => clearedCells.has(MODE_SETTINGS[mode].column + row))) continue;
    const link = parseMergeLink(comment.getContent());
    if (link) links.set(row, link);
  }
  return links;
}

function nearby(rows) {
  return groupNearbyRows(Array.from(rows).sort((left, right) => left - right));
}

function indexMember(index, key, member) {
  const members = index.get(key) || [];
  members.push(member);
  index.set(key, members);
}

function fileRefKey(value) {
  return value.normalize("NFKC").trim().toLocaleLowerCase("tr-TR").replace(/\s+/g, " ")
    .replace(/\s*[-–—]\s*/g, "-").replace(/-denet[ıi]m$/, "-denetim").replace(/-k[ıi]myahane$/, "-kimyahane");
}

function sourceFileKey(member) {
  const suffix = isInspection(member.note) ? "-DENETİM" : isChemical(member.note) ? "-Kimyahane" : "";
  const file = fileRefKey(member.file).replace(/[ -](?:trans[iı]t|tranz[iı]t)$/, " transit");
  const category = fileRefKey(suffix);
  return category && !file.endsWith(category) ? file + category : file;
}

function isInspection(value) {
  return value.trim().toLocaleLowerCase("tr-TR").replace(/ı/g, "i") === "denetim";
}

function isChemical(value) {
  return value.trim().toLocaleLowerCase("tr-TR").replace(/ı/g, "i") === "kimyahane";
}

function readPieces(raw) {
  return typeof raw === "number" ? raw : /^\d+$/.test(String(raw).trim()) ? Number(raw) : NaN;
}

function commentKey(value) {
  return value.normalize("NFKC").toLocaleLowerCase("tr-TR").replace(/[\s\-–—]/g, "");
}

function mergeTitleKey(value) {
  const key = commentKey(value);
  return key.replace(/denet[ıi]m$/, "denetim").replace(/k[ıi]myahane$/, "kimyahane");
}

function parseMergeLink(content) {
  const note = content.normalize("NFKC").trim().replace(/\s+/g, " ");
  const directed = /^(?:kurumsalda\s+)?(.+?)(?:(?:\s+|['’])(?:de|da|te|ta)|\s+alt[ıi]ndad[ıi]r)[.!]?$/i.exec(note);
  const explicit = directed && /^[a-z]{1,20}[\s-]*\d{1,12}(?:\s*[/–—-]?\s*[\p{L}\p{N}][\p{L}\p{N}\s.\-–—]*)?$/iu.test(directed[1]);
  const target = explicit && directed ? directed[1].replace(/^([a-z]+[\s-]*\d+)\s*\/\s*/i, "$1") : note.replace(/\.$/, "");
  const normalized = fileRefKey(target).replace(/ı/g, "i");
  const targetSuffix = normalized.endsWith("-denetim") ? "-DENETİM" : normalized.endsWith("-kimyahane") ? "-Kimyahane" : "";
  const key = mergeTitleKey(target);
  const shipmentOnly = /^[a-z]{1,20}\d{1,12}$/i.test(key);
  if (explicit || shipmentOnly) return {key, shipmentOnly, targetSuffix};
  return /^kurumsalda\b/i.test(note) ? {key:"", shipmentOnly:false, targetSuffix:""} : undefined;
}


function groupNearbyRows(rows) {
  const groups= [];
  for (const row of rows) {
    const last = groups[groups.length - 1];
    // Batch nearby reads without loading very large gaps between selections.
    if (last && row - last[last.length - 1] <= 12 && row - last[0] < 1000) last.push(row);
    else groups.push([row]);
  }
  return groups;
}

function readRows(
  source, first, last, lastColumn, firstColumn = "A"
) {
  const range = source.getRange(`${firstColumn}${first}:${lastColumn}${last}`);
  return { values: range.getValues(), texts: range.getTexts() };
}

function parseDate(value, displayed, mode) {
  if (typeof value !== "number") return parseTextDate(String(value).trim(), mode);
  if (!Number.isFinite(value) || value < 61 || value >= 100000) return "";
  const asDate = new Date(Date.UTC(1899, 11, 30) + Math.round(value * 86400000));
  const iso = asDate.toISOString();
  const day1904 = new Date(asDate.getTime() + 1462 * 86400000).toISOString().slice(0, 10);
  // A localized display can swap day and month. Use the underlying serial unless
  // the display matches the other Excel date system's exact calendar day.
  const dateFromText = parseTextDate(displayed.trim(), mode);
  const day = dateFromText.slice(0, 10) === day1904 ? day1904 : iso.slice(0, 10);
  return day + "T" + (usesExcelTime(mode) ? iso.slice(11, 16) : "18:00");
}

function usesExcelTime(mode) {
  return mode !== "nakliye" && mode !== "boyner";
}

function displayDate(iso) {
  if (!iso) return "";
  return `${iso.slice(8, 10)}.${iso.slice(5, 7)}.${iso.slice(0, 4)} ${iso.slice(11, 16)}`;
}

function parseTextDate(value, mode) {
  const local = value.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})(?:\s+(\d{1,2})[:.](\d{2})(?::\d{2})?)?$/);
  const iso = value.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::\d{2})?)?$/);
  if (!local && !iso) return "";
  const fields = local || iso;
  const year = Number(fields[local ? 3 : 1]);
  const month = Number(fields[2]);
  const day = Number(fields[local ? 1 : 3]);
  const hour = Number(fields[4] || "0");
  const minute = Number(fields[5] || "0");
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() + 1 !== month ||
      date.getUTCDate() !== day || hour > 23 || minute > 59) return "";
  const pad = (part)=> String(part).padStart(2, "0");
  const clock = usesExcelTime(mode) ? `${pad(hour)}:${pad(minute)}` : "18:00";
  return `${year}-${pad(month)}-${pad(day)}T${clock}`;
}

function safeText(text) {
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function makeId() {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, character => {
    const value = Math.floor(Math.random() * 16);
    return (character === "x" ? value : (value & 3) | 8).toString(16);
  });
}
