import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

const RATE_UPLOAD_COLUMN_COUNT = 18;
const ALLOCATION_UPLOAD_COLUMN_COUNT = 5;
const RATE_VALUE_TARGET_COLUMNS = new Set([1, 2, 3, 4, 5, 11, 12, 13, 14]);
const MAX_REASONABLE_OFFERS = 50;

const RATE_HEADERS = [
  "TOTAL SELL RATE\n(Secret Escapes)",
  "TOTAL SUPPLIER RATE\n(hotel's or tour operator's site)",
  "SINGLE RATE",
  "CHILD RATE",
  "INFANT RATE",
  "START DATE\n(YYYY-MM-DD)",
  "NUMBER OF NIGHTS",
  "MINIMUM NO. OF NIGHTS\n(for flexible hotel stays)",
  "DEPARTURE AIRPORT CODE",
  "DESTINATION AIRPORT CODE",
  "TOTAL DEPOSIT RATE\n(ADULTS)",
  "Deposit amount (per single adult)",
  "CHILD DEPOSIT RATE",
  "INFANT DEPOSIT RATE",
  "FINAL BALANCE DUE DATE",
  "OUTBOUND OVERNIGHT FLIGHT\n(mark TRUE or FALSE)",
  "INBOUND OVERNIGHT FLIGHT\n(mark TRUE or FALSE)",
  "BLACKOUT PER DEPARTURE\n(mark TRUE or FALSE)",
];

const ALLOCATION_HEADERS = [
  "START DATE\n(YYYY-MM-DD)",
  "NO. OF ROOMS ALLOCATED\n(to upload new dates OR replace allotment for existing dates)",
  "ROOMS TO ADD\n(to current allocation)",
  "ROOMS TO REMOVE\n(from current allocation)",
  "BLACK DATE OUT\n(all airports)",
];

export async function buildImportedFiles({ templateBuffer, offerBuffer, offerFileName = "source workbook" }) {
  const templateWb = templateBuffer
    ? await loadTemplateWorkbook(templateBuffer)
    : createBuiltInTemplateWorkbook();

  const sourceWb = loadSourceWorkbook(offerBuffer);
  const ratesTemplate = findRatesTemplate(templateWb);
  const allocationTemplate = findAllocationTemplate(templateWb);

  if (!ratesTemplate) {
    throw new Error("Could not find a rates template tab. Expected 'Rates Template', 'Sheet1', or a sheet with the rates upload headers.");
  }

  if (!allocationTemplate) {
    throw new Error("Could not find an allocation template tab. Expected 'Allocation Template', 'Sheet2', or a sheet with the allocation upload headers.");
  }

  if (sourceWb.getWorksheet("Offer Details")) {
    return buildFilesFromSubmissionOfferForm({ sourceWb, ratesTemplate, allocationTemplate, offerFileName });
  }

  const pairs = getGenericRatesAllocationPairs(sourceWb);
  if (pairs.length > 0) {
    return buildFilesFromGenericPairs({ sourceWb, ratesTemplate, allocationTemplate, pairs, offerFileName });
  }

  throw new Error("Could not recognise the source workbook. Expected Rates/Allocation tabs, or an Offer Details sheet.");
}

async function loadTemplateWorkbook(templateBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);
  return workbook;
}

function createBuiltInTemplateWorkbook() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Offer Importer";
  workbook.created = new Date();
  workbook.modified = new Date();

  createTemplateSheet(workbook, "Rates Template", RATE_HEADERS, RATE_UPLOAD_COLUMN_COUNT, {
    dateColumns: [6, 15],
    widths: [16, 18, 12, 12, 12, 14, 12, 18, 14, 14, 14, 14, 12, 12, 14, 16, 16, 16],
  });

  createTemplateSheet(workbook, "Allocation Template", ALLOCATION_HEADERS, ALLOCATION_UPLOAD_COLUMN_COUNT, {
    dateColumns: [1],
    widths: [14, 22, 16, 16, 16],
  });

  return workbook;
}

function createTemplateSheet(workbook, sheetName, headers, columnCount, { dateColumns, widths }) {
  const sheet = workbook.addWorksheet(sheetName, {
    views: [{ state: "frozen", ySplit: 1 }],
  });

  const headerFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF7" } };
  const styleFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
  const border = {
    top: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
  };

  for (let columnIndex = 1; columnIndex <= columnCount; columnIndex += 1) {
    sheet.getColumn(columnIndex).width = widths[columnIndex - 1] || 14;
  }

  const headerRow = sheet.getRow(1);
  headerRow.height = 58;
  for (let columnIndex = 1; columnIndex <= columnCount; columnIndex += 1) {
    const cell = headerRow.getCell(columnIndex);
    cell.value = headers[columnIndex - 1];
    cell.fill = headerFill;
    cell.border = border;
    cell.font = { bold: true, size: 10 };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  }

  const styleRow = sheet.getRow(2);
  styleRow.height = 18;
  for (let columnIndex = 1; columnIndex <= columnCount; columnIndex += 1) {
    const cell = styleRow.getCell(columnIndex);
    cell.fill = styleFill;
    cell.border = border;
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    if (dateColumns.includes(columnIndex)) {
      cell.numFmt = "yyyy-mm-dd";
    }
  }

  sheet.autoFilter = `${sheet.getColumn(1).letter}1:${sheet.getColumn(columnCount).letter}1`;
  return sheet;
}

class SourceWorkbook {
  constructor(workbook) {
    this.sheetNames = workbook.SheetNames;
    this.worksheets = this.sheetNames.map((name) => new SourceSheet(name, workbook.Sheets[name]));
  }

  getWorksheet(name) {
    return this.worksheets.find((sheet) => sheet.name === name);
  }
}

class SourceSheet {
  constructor(name, sheet) {
    this.name = name;
    this.sheet = sheet;
    this.range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    this.rowCount = this.range.e.r + 1;
    this.columnCount = this.range.e.c + 1;
  }

  getCell(rowNumber, columnNumber) {
    const address = XLSX.utils.encode_cell({ r: rowNumber - 1, c: columnNumber - 1 });
    const cell = this.sheet[address];
    return { value: cell ? normalizeSourceCellValue(cell) : null };
  }
}

function loadSourceWorkbook(buffer) {
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    cellFormula: false,
    raw: true,
  });
  return new SourceWorkbook(workbook);
}

async function buildFilesFromSubmissionOfferForm({ sourceWb, ratesTemplate, allocationTemplate, offerFileName }) {
  const offerDetails = sourceWb.getWorksheet("Offer Details");
  const offerNumbers = getOfferNumbers(sourceWb, offerDetails);
  if (offerNumbers.length === 0) {
    const pairs = getGenericRatesAllocationPairs(sourceWb);
    if (pairs.length > 0) {
      return buildFilesFromGenericPairs({ sourceWb, ratesTemplate, allocationTemplate, pairs, offerFileName });
    }
    throw new Error("No populated offers were found in Offer Details row 22.");
  }

  const files = [];

  for (const offerNumber of offerNumbers) {
    const file = await buildSinglePairFile({
      ratesTemplate,
      allocationTemplate,
      ratesSource: sourceWb.getWorksheet(`Rates (${offerNumber})`),
      allocationSource: sourceWb.getWorksheet(`Allocation (${offerNumber})`),
      outputNumber: offerNumber,
      label: `Offer ${offerNumber}`,
      offerFileName,
    });

    if (file) files.push(file);
  }

  if (files.length === 0) {
    throw new Error(`No Rates/Allocation pairs with values were found in ${offerFileName}.`);
  }

  return {
    files,
    importedLabels: files.map((file) => file.label),
    sourceType: "submission offer form",
  };
}

async function buildFilesFromGenericPairs({ sourceWb, ratesTemplate, allocationTemplate, pairs, offerFileName }) {
  const files = [];

  for (const [index, pair] of pairs.entries()) {
    const outputNumber = index + 1;
    const file = await buildSinglePairFile({
      ratesTemplate,
      allocationTemplate,
      ratesSource: sourceWb.getWorksheet(pair.ratesSheetName),
      allocationSource: sourceWb.getWorksheet(pair.allocationSheetName),
      outputNumber,
      label: pair.label,
      offerFileName,
    });

    if (file) files.push(file);
  }

  if (files.length === 0) {
    throw new Error(`No Rates/Allocation pairs with values were found in ${offerFileName}.`);
  }

  return {
    files,
    importedLabels: files.map((file) => file.label),
    sourceType: "paired rates/allocation workbook",
  };
}

async function buildSinglePairFile({
  ratesTemplate,
  allocationTemplate,
  ratesSource,
  allocationSource,
  outputNumber,
  label,
  offerFileName,
}) {
  const outputWb = createOutputWorkbook();
  let copiedRateRows = 0;
  let copiedAllocationRows = 0;

  if (ratesSource) {
    const ratesSheet = cloneTemplateSheet(outputWb, ratesTemplate, `Rates (${outputNumber})`, RATE_UPLOAD_COLUMN_COUNT);
    copiedRateRows = populateRatesSheet(ratesSheet, ratesSource);
  }

  if (allocationSource) {
    const allocationSheet = cloneTemplateSheet(outputWb, allocationTemplate, `Allocation (${outputNumber})`, ALLOCATION_UPLOAD_COLUMN_COUNT);
    copiedAllocationRows = populateAllocationSheet(allocationSheet, allocationSource);
  }

  if (copiedRateRows === 0 && copiedAllocationRows === 0) {
    return null;
  }

  if (ratesSource && copiedRateRows === 0) {
    return null;
  }

  return {
    filename: buildOutputFilename(offerFileName, label),
    label,
    outputBuffer: await outputWb.xlsx.writeBuffer(),
  };
}

function getGenericRatesAllocationPairs(workbook) {
  const ratesSheets = [];
  const allocationSheets = [];

  workbook.worksheets.forEach((sheet, index) => {
    const ratesLayout = detectRatesLayout(sheet);
    if (ratesLayout) {
      const keyInfo = getPairKey(sheet.name, "rates");
      ratesSheets.push({
        sheetName: sheet.name,
        layout: ratesLayout,
        key: keyInfo.key,
        label: keyInfo.label,
        sortValue: keyInfo.sortValue ?? index + 1,
        index,
      });
      return;
    }

    if (isAllocationLikeSheet(sheet)) {
      const keyInfo = getPairKey(sheet.name, "allocation");
      allocationSheets.push({
        sheetName: sheet.name,
        key: keyInfo.key,
        label: keyInfo.label,
        sortValue: keyInfo.sortValue ?? index + 1,
        index,
      });
    }
  });

  if (ratesSheets.length === 1 && allocationSheets.length === 1) {
    return [{
      ratesSheetName: ratesSheets[0].sheetName,
      allocationSheetName: allocationSheets[0].sheetName,
      label: choosePairLabel(ratesSheets[0], allocationSheets[0]),
      sortValue: ratesSheets[0].index,
    }];
  }

  const allocationsByKey = new Map();
  allocationSheets.forEach((sheet) => {
    if (!allocationsByKey.has(sheet.key)) allocationsByKey.set(sheet.key, []);
    allocationsByKey.get(sheet.key).push(sheet);
  });

  const pairs = [];
  const usedAllocations = new Set();
  const usedRates = new Set();

  ratesSheets.forEach((ratesSheet) => {
    const matchingAllocations = allocationsByKey.get(ratesSheet.key) || [];
    const allocationSheet = matchingAllocations.find((candidate) => !usedAllocations.has(candidate.sheetName));
    if (!allocationSheet) return;

    usedRates.add(ratesSheet.sheetName);
    usedAllocations.add(allocationSheet.sheetName);
    pairs.push({
      ratesSheetName: ratesSheet.sheetName,
      allocationSheetName: allocationSheet.sheetName,
      label: choosePairLabel(ratesSheet, allocationSheet),
      sortValue: ratesSheet.index,
    });
  });

  ratesSheets.forEach((ratesSheet) => {
    if (usedRates.has(ratesSheet.sheetName)) return;

    const allocationSheet = findAdjacentAllocationSheet(ratesSheet, allocationSheets, usedAllocations);
    if (!allocationSheet) return;

    usedRates.add(ratesSheet.sheetName);
    usedAllocations.add(allocationSheet.sheetName);
    pairs.push({
      ratesSheetName: ratesSheet.sheetName,
      allocationSheetName: allocationSheet.sheetName,
      label: choosePairLabel(ratesSheet, allocationSheet),
      sortValue: ratesSheet.index,
    });
  });

  pairRemainingSheetsByOrder({ ratesSheets, allocationSheets, usedRates, usedAllocations, pairs });

  return pairs.sort((a, b) => a.sortValue - b.sortValue);
}

function findAdjacentAllocationSheet(ratesSheet, allocationSheets, usedAllocations) {
  const nextSheet = allocationSheets.find((candidate) => (
    !usedAllocations.has(candidate.sheetName)
    && candidate.index === ratesSheet.index + 1
  ));
  if (nextSheet) return nextSheet;

  return allocationSheets.find((candidate) => (
    !usedAllocations.has(candidate.sheetName)
    && candidate.index === ratesSheet.index - 1
  ));
}

function pairRemainingSheetsByOrder({ ratesSheets, allocationSheets, usedRates, usedAllocations, pairs }) {
  const remainingRates = ratesSheets
    .filter((sheet) => !usedRates.has(sheet.sheetName))
    .sort((a, b) => a.index - b.index);
  const remainingAllocations = allocationSheets
    .filter((sheet) => !usedAllocations.has(sheet.sheetName))
    .sort((a, b) => a.index - b.index);

  const pairCount = Math.min(remainingRates.length, remainingAllocations.length);
  for (let index = 0; index < pairCount; index += 1) {
    const ratesSheet = remainingRates[index];
    const allocationSheet = remainingAllocations[index];
    usedRates.add(ratesSheet.sheetName);
    usedAllocations.add(allocationSheet.sheetName);
    pairs.push({
      ratesSheetName: ratesSheet.sheetName,
      allocationSheetName: allocationSheet.sheetName,
      label: choosePairLabel(ratesSheet, allocationSheet),
      sortValue: ratesSheet.index,
    });
  }
}

function getPairKey(sheetName, kind) {
  const name = normalizeName(sheetName);
  const nightsMatch = /(?:^|\s)(\d+)\s*nts?(?:\s|$)/i.exec(name);
  if (nightsMatch) {
    const nights = Number(nightsMatch[1]);
    return { key: `nights:${nights}`, label: `${nights} Nts`, sortValue: nights };
  }

  const numberMatch = /\((\d+)\)|(?:^|\s)(?:offer|pair)?\s*(\d+)(?:\s|$)/i.exec(name);
  const explicitNumber = numberMatch ? Number(numberMatch[1] || numberMatch[2]) : null;

  let cleaned = name
    .replace(/\((\d+)\)/g, " $1 ")
    .replace(/\b(total|upload|template|worksheet|sheet|sheets|tab|tabs|details|detail|departure|departures|summary|room|rooms|open|opens|available|availability|allotment|cabin|cabins)\b/g, " ")
    .replace(/\b(rates?|rate)\b/g, " ")
    .replace(/\b(all|alloc[a-z]*|aloc[a-z]*|aloc|allocaiton|allot[a-z]*)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!cleaned && explicitNumber !== null) {
    return { key: `number:${explicitNumber}`, label: `Pair ${explicitNumber}`, sortValue: explicitNumber };
  }

  if (/^\d+$/.test(cleaned) && explicitNumber !== null) {
    return { key: `number:${explicitNumber}`, label: `Pair ${explicitNumber}`, sortValue: explicitNumber };
  }

  if (!cleaned) {
    return { key: `${kind}:${name}`, label: titleCase(name || kind) };
  }

  return {
    key: `name:${cleaned}`,
    label: titleCase(cleaned),
    sortValue: explicitNumber,
  };
}

function choosePairLabel(ratesSheet, allocationSheet) {
  if (ratesSheet.label && !/^Rates?$/i.test(ratesSheet.label)) return ratesSheet.label;
  if (allocationSheet.label && !/^Alloc/i.test(allocationSheet.label)) return allocationSheet.label;
  return titleCase(ratesSheet.key.replace(/^[^:]+:/, ""));
}

function detectRatesLayout(sheet) {
  if (
    cleanText(sourceCellValue(sheet, 1, 1)).includes("TOTAL SELL RATE")
    && cleanText(sourceCellValue(sheet, 1, 2)).includes("TOTAL SUPPLIER RATE")
    && cleanText(sourceCellValue(sheet, 1, 6)).includes("START DATE")
  ) {
    return { mode: "sequential", startColumn: 1, dataStartRow: 2 };
  }

  if (
    (
      cleanText(sourceCellValue(sheet, 1, 3)).includes("TOTAL RATE")
      || cleanText(sourceCellValue(sheet, 1, 3)).includes("TOTAL SELL RATE")
    )
    && cleanText(sourceCellValue(sheet, 1, 8)).includes("START DATE")
    && cleanText(sourceCellValue(sheet, 1, 11)).includes("DEPARTURE AIRPORT")
  ) {
    return { mode: "sequential", startColumn: 3, dataStartRow: 2 };
  }

  if (
    cleanText(sourceCellValue(sheet, 2, 3)).includes("PRO ZIMMER ANGEBOTSRATE")
    && cleanText(sourceCellValue(sheet, 2, 4)).includes("PRO ZIMMER VERANSTALTERRATE")
    && cleanText(sourceCellValue(sheet, 2, 9)).includes("START DATE")
    && cleanText(sourceCellValue(sheet, 2, 13)).includes("ABFLUGHAFEN")
  ) {
    return { mode: "germanOfferForm", dataStartRow: 3 };
  }

  if (
    cleanText(sourceCellValue(sheet, 2, 3)).includes("TOTAL OFFER PRICE")
    && cleanText(sourceCellValue(sheet, 2, 4)).includes("TOTAL SELLER RATE")
    && cleanText(sourceCellValue(sheet, 2, 9)).includes("DEPARTURE DATE")
    && cleanText(sourceCellValue(sheet, 2, 10)).includes("NIGHTS")
  ) {
    return { mode: "submissionNoFlight", dataStartRow: 3 };
  }

  return null;
}

function isAllocationLikeSheet(sheet) {
  return isAllocationHeaderRow(sheet, 1) || isAllocationHeaderRow(sheet, 2);
}

function detectAllocationLayout(sheet) {
  if (isAllocationHeaderRow(sheet, 1)) {
    return { dataStartRow: 2 };
  }
  if (isAllocationHeaderRow(sheet, 2)) {
    return { dataStartRow: 3 };
  }
  return { dataStartRow: 2 };
}

function isAllocationHeaderRow(sheet, rowNumber) {
  const startHeader = cleanText(sourceCellValue(sheet, rowNumber, 1));
  const allocationHeader = cleanText(sourceCellValue(sheet, rowNumber, 2));
  return (startHeader.includes("START DATE") || startHeader.includes("DEPARTURE DATE"))
    && (
      allocationHeader.includes("NO. OF ROOMS ALLOCATED")
      || allocationHeader.includes("SHARED ALLOCATION")
      || allocationHeader.includes("ALLOCATION")
    );
}

function getOfferNumbers(sourceWb, offerDetails) {
  const result = getOfferNumbersFromMarkerRow(sourceWb, offerDetails, 22);
  if (result.length > 0) return result;

  for (let rowNumber = 1; rowNumber <= offerDetails.rowCount; rowNumber += 1) {
    const rowLabel = cleanText(sourceCellValue(offerDetails, rowNumber, 1));
    if (rowLabel.includes("OFFER DETAIL")) {
      const rowResult = getOfferNumbersFromMarkerRow(sourceWb, offerDetails, rowNumber);
      if (rowResult.length > 0) return rowResult;
    }
  }

  return [];
}

function getOfferNumbersFromMarkerRow(sourceWb, offerDetails, rowNumber) {
  const result = [];
  const lastColumn = Math.min(Math.max(offerDetails.columnCount || 0, 2), MAX_REASONABLE_OFFERS + 1);

  for (let column = 2; column <= lastColumn; column += 1) {
    const offerNumber = column - 1;
    const title = cleanText(sourceCellValue(offerDetails, rowNumber, column));

    if (
      title.length > 0
      && (
        sourceWb.getWorksheet(`Rates (${offerNumber})`)
        || sourceWb.getWorksheet(`Allocation (${offerNumber})`)
      )
    ) {
      result.push(offerNumber);
    }
  }

  return result;
}

function createOutputWorkbook() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Offer Importer";
  workbook.created = new Date();
  workbook.modified = new Date();
  return workbook;
}

function findRatesTemplate(workbook) {
  return workbook.getWorksheet("Rates Template")
    || workbook.getWorksheet("Sheet1")
    || workbook.worksheets.find((sheet) => cleanText(sheet.getCell(1, 1).value).includes("TOTAL SELL RATE"));
}

function findAllocationTemplate(workbook) {
  return workbook.getWorksheet("Allocation Template")
    || workbook.getWorksheet("Sheet2")
    || workbook.worksheets.find((sheet) => cleanText(sheet.getCell(1, 2).value).includes("NO. OF ROOMS ALLOCATED"));
}

function cloneTemplateSheet(workbook, templateSheet, newName, maxColumns) {
  const sheet = workbook.addWorksheet(newName, {
    properties: { ...templateSheet.properties },
    pageSetup: { ...templateSheet.pageSetup },
    views: templateSheet.views ? structuredCloneSafe(templateSheet.views) : undefined,
  });

  for (let columnIndex = 1; columnIndex <= maxColumns; columnIndex += 1) {
    const sourceColumn = templateSheet.getColumn(columnIndex);
    const targetColumn = sheet.getColumn(columnIndex);
    targetColumn.width = sourceColumn.width;
    targetColumn.hidden = sourceColumn.hidden;
    targetColumn.style = cloneStyle(sourceColumn.style);
  }

  const headerRow = templateSheet.getRow(1);
  const outputHeader = sheet.getRow(1);
  outputHeader.height = headerRow.height;

  for (let columnIndex = 1; columnIndex <= maxColumns; columnIndex += 1) {
    copyCell(headerRow.getCell(columnIndex), outputHeader.getCell(columnIndex));
  }

  const styleRow = templateSheet.getRow(2);
  const outputStyleRow = sheet.getRow(2);
  outputStyleRow.height = styleRow.height;
  for (let columnIndex = 1; columnIndex <= maxColumns; columnIndex += 1) {
    copyStyleOnly(styleRow.getCell(columnIndex), outputStyleRow.getCell(columnIndex));
  }

  copyMerges(templateSheet, sheet, maxColumns);
  sheet.autoFilter = templateSheet.autoFilter || undefined;
  return sheet;
}

function populateRatesSheet(targetSheet, sourceSheet) {
  const layout = detectRatesLayout(sourceSheet);
  let targetRowNumber = 2;
  const lastRow = findLastRatesRow(sourceSheet, layout);
  const expandedRowBuckets = [];

  for (let sourceRowNumber = layout.dataStartRow; sourceRowNumber <= lastRow; sourceRowNumber += 1) {
    if (!ratesRowHasUsableMainRate(sourceSheet, sourceRowNumber, layout)) continue;

    const normalizedValues = [];

    for (let targetColumn = 1; targetColumn <= RATE_UPLOAD_COLUMN_COUNT; targetColumn += 1) {
      const rawValue = getRateSourceValue(sourceSheet, sourceRowNumber, targetColumn, layout);
      normalizedValues.push(normalizeRatesValue(rawValue, targetColumn));
    }

    expandAirportCodeRows(normalizedValues).forEach((rowValues, expansionIndex) => {
      if (!expandedRowBuckets[expansionIndex]) expandedRowBuckets[expansionIndex] = [];
      expandedRowBuckets[expansionIndex].push(rowValues);
    });
  }

  const styleRow = targetSheet.getRow(2);
  for (const rowBucket of expandedRowBuckets) {
    for (const rowValues of rowBucket) {
      writeRatesRow(targetSheet, targetRowNumber, styleRow, rowValues);
      targetRowNumber += 1;
    }
  }

  targetSheet.autoFilter = `A1:R${Math.max(1, targetRowNumber - 1)}`;
  targetSheet.views = [{ state: "frozen", ySplit: 1 }];
  return targetRowNumber - 2;
}

function writeRatesRow(targetSheet, rowNumber, styleRow, rowValues) {
  const targetRow = targetSheet.getRow(rowNumber);
  targetRow.height = styleRow.height;

  for (let index = 0; index < RATE_UPLOAD_COLUMN_COUNT; index += 1) {
    const targetColumn = index + 1;
    const targetCell = targetRow.getCell(targetColumn);
    copyStyleOnly(styleRow.getCell(targetColumn), targetCell);
    targetCell.value = rowValues[index];
    if (targetColumn === 6 || targetColumn === 15) {
      targetCell.numFmt = "yyyy-mm-dd";
    }
  }

  targetRow.commit?.();
}

function expandAirportCodeRows(rowValues) {
  const departureCodes = splitAirportCodes(rowValues[8]);
  const destinationCodes = splitAirportCodes(rowValues[9]);
  const expandedRows = [];

  for (const departureCode of departureCodes) {
    for (const destinationCode of destinationCodes) {
      const expandedValues = [...rowValues];
      expandedValues[8] = departureCode;
      expandedValues[9] = destinationCode;
      expandedRows.push(expandedValues);
    }
  }

  return expandedRows;
}

function splitAirportCodes(value) {
  if (typeof value !== "string" || !value.includes("/")) {
    return [value];
  }

  const codes = value
    .split("/")
    .map((code) => code.trim())
    .filter(Boolean);

  return codes.length > 0 ? codes : [value];
}

function populateAllocationSheet(targetSheet, sourceSheet) {
  const layout = detectAllocationLayout(sourceSheet);
  let targetRowNumber = 2;
  const lastRow = findLastAllocationRow(sourceSheet);

  for (let sourceRowNumber = layout.dataStartRow; sourceRowNumber <= lastRow; sourceRowNumber += 1) {
    const sourceValues = [1, 2, 3, 4, 5].map((columnNumber) => (
      sourceCellValue(sourceSheet, sourceRowNumber, columnNumber)
    ));

    if (!sourceValues.some(hasContent)) continue;

    const targetRow = targetSheet.getRow(targetRowNumber);
    const styleRow = targetSheet.getRow(2);
    targetRow.height = styleRow.height;

    for (let columnIndex = 1; columnIndex <= ALLOCATION_UPLOAD_COLUMN_COUNT; columnIndex += 1) {
      const targetCell = targetRow.getCell(columnIndex);
      copyStyleOnly(styleRow.getCell(columnIndex), targetCell);
      targetCell.value = columnIndex === 1
        ? normalizeDateValue(sourceValues[columnIndex - 1])
        : normalizeAllocationValue(sourceValues[columnIndex - 1]);
    }

    targetRow.getCell(1).numFmt = "yyyy-mm-dd";
    targetRow.commit?.();
    targetRowNumber += 1;
  }

  targetSheet.autoFilter = `A1:E${Math.max(1, targetRowNumber - 1)}`;
  targetSheet.views = [{ state: "frozen", ySplit: 1 }];
  return targetRowNumber - 2;
}

function findLastRatesRow(sheet, layout) {
  const columns = getRatesLastRowColumns(layout);
  return Math.max(
    ...columns.map((columnNumber) => findLastRowInColumn(sheet, columnNumber)),
  );
}

function getRateSourceValue(sheet, rowNumber, targetColumnNumber, layout) {
  const sourceColumn = getRateSourceColumn(targetColumnNumber, layout);
  if (!sourceColumn) return null;
  return sourceCellValue(sheet, rowNumber, sourceColumn);
}

function getRateSourceColumn(targetColumnNumber, layout) {
  if (layout.mode === "sequential") {
    return layout.startColumn + targetColumnNumber - 1;
  }

  if (layout.mode === "submissionNoFlight") {
    const noFlightColumnMap = {
      1: 3,
      2: 4,
      3: 5,
      4: 6,
      5: 7,
      6: 9,
      7: 10,
      8: 11,
      9: 12,
      10: 13,
      11: 17,
      12: null,
      13: null,
      14: null,
      15: 14,
      16: 15,
      17: 16,
      18: null,
    };

    return noFlightColumnMap[targetColumnNumber] || null;
  }

  const germanColumnMap = {
    1: 3,
    2: 4,
    3: 5,
    4: 7,
    5: 6,
    6: 9,
    7: 10,
    8: null,
    9: 13,
    10: 14,
    11: 18,
    12: 19,
    13: 21,
    14: 20,
    15: 15,
    16: 16,
    17: 17,
    18: null,
  };

  return germanColumnMap[targetColumnNumber] || null;
}

function getRatesLastRowColumns(layout) {
  if (layout.mode === "sequential") {
    return [
      layout.startColumn,
      layout.startColumn + 1,
      layout.startColumn + 5,
      layout.startColumn + 8,
      layout.startColumn + 9,
    ];
  }

  return [3, 4, 9, 13, 14];
}

function findLastAllocationRow(sheet) {
  return Math.max(findLastRowInColumn(sheet, 1), findLastRowInColumn(sheet, 2));
}

function findLastRowInColumn(sheet, columnNumber) {
  for (let rowNumber = sheet.rowCount; rowNumber >= 1; rowNumber -= 1) {
    if (hasContent(sourceCellValue(sheet, rowNumber, columnNumber))) {
      return rowNumber;
    }
  }
  return 1;
}

function ratesRowHasUsableMainRate(sheet, rowNumber, layout) {
  return getRatesMainRateColumns(layout).some((columnNumber) => (
    isNonZeroValue(sourceCellValue(sheet, rowNumber, columnNumber))
  ));
}

function getRatesMainRateColumns(layout) {
  if (layout.mode === "sequential") {
    return [layout.startColumn, layout.startColumn + 1];
  }
  return [3, 4];
}

function copyCell(sourceCell, targetCell) {
  targetCell.value = cloneValue(sourceCell.value);
  copyStyleOnly(sourceCell, targetCell);
  if (sourceCell.note) targetCell.note = structuredCloneSafe(sourceCell.note);
}

function copyStyleOnly(sourceCell, targetCell) {
  targetCell.style = cloneStyle(sourceCell.style);
  if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt;
  if (sourceCell.alignment) targetCell.alignment = structuredCloneSafe(sourceCell.alignment);
  if (sourceCell.font) targetCell.font = structuredCloneSafe(sourceCell.font);
  if (sourceCell.fill) targetCell.fill = structuredCloneSafe(sourceCell.fill);
  if (sourceCell.border) targetCell.border = structuredCloneSafe(sourceCell.border);
  if (sourceCell.protection) targetCell.protection = structuredCloneSafe(sourceCell.protection);
}

function copyMerges(sourceSheet, targetSheet, maxColumns) {
  const merges = sourceSheet.model?.merges || [];
  for (const mergeRange of merges) {
    const decoded = decodeRange(mergeRange);
    if (decoded && decoded.right <= maxColumns) {
      targetSheet.mergeCells(mergeRange);
    }
  }
}

function decodeRange(range) {
  const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(range);
  if (!match) return null;
  return {
    left: columnLettersToNumber(match[1]),
    top: Number(match[2]),
    right: columnLettersToNumber(match[3]),
    bottom: Number(match[4]),
  };
}

function columnLettersToNumber(letters) {
  return letters.toUpperCase().split("").reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
}

function sourceCellValue(sheet, rowNumber, columnNumber) {
  return sheet.getCell(rowNumber, columnNumber).value;
}

function normalizeSourceCellValue(cell) {
  if (cell === null || cell === undefined) return null;
  if (cell.t === "z") return null;
  return cell.v ?? null;
}

function normalizeRatesValue(value, targetColumnNumber) {
  if (isBlankLike(value)) {
    return null;
  }
  if (RATE_VALUE_TARGET_COLUMNS.has(targetColumnNumber) && isZeroLike(value)) {
    return null;
  }
  if (targetColumnNumber === 6 || targetColumnNumber === 15) {
    return normalizeDateValue(value);
  }
  if (targetColumnNumber >= 16 && targetColumnNumber <= 18) {
    return normalizeBooleanValue(value);
  }
  return value;
}

function normalizeAllocationValue(value) {
  if (isBlankLike(value)) {
    return null;
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return value;
}

function normalizeDateValue(value) {
  if (value instanceof Date) {
    return createExcelDate(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }

  if (typeof value === "string") {
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value.trim());
    if (match) {
      return createExcelDate(Number(match[1]), Number(match[2]), Number(match[3]));
    }
  }

  return value;
}

function createExcelDate(year, month, day) {
  return new Date(Date.UTC(year, month - 1, day));
}

function normalizeBooleanText(value) {
  const cleaned = value.trim().toUpperCase();
  if (cleaned === "1" || cleaned === "JA" || cleaned === "YES" || cleaned === "TRUE" || cleaned === "WAHR") {
    return "TRUE";
  }
  if (cleaned === "0" || cleaned === "NEIN" || cleaned === "NO" || cleaned === "FALSE" || cleaned === "FALSCH") {
    return "FALSE";
  }
  return value;
}

function normalizeBooleanValue(value) {
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (typeof value === "number" && (value === 0 || value === 1)) {
    return value === 1 ? "TRUE" : "FALSE";
  }
  if (typeof value === "string") {
    return normalizeBooleanText(value);
  }
  return value;
}

function hasContent(value) {
  return !isBlankLike(value);
}

function isBlankLike(value) {
  return value === null || value === undefined || String(value).trim().length === 0;
}

function isNonZeroValue(value) {
  return !isBlankLike(value) && !isZeroLike(value);
}

function isZeroLike(value) {
  if (typeof value === "number") {
    return value === 0;
  }
  if (typeof value === "string") {
    const cleaned = value.trim();
    if (cleaned.length === 0) return false;
    const numericValue = Number(cleaned.replace(/,/g, ""));
    return Number.isFinite(numericValue) && numericValue === 0;
  }
  return false;
}

function cleanText(value) {
  const resolved = value && typeof value === "object" && "result" in value ? value.result : value;
  if (resolved === null || resolved === undefined) return "";
  if (typeof resolved === "object" && "richText" in resolved) {
    return resolved.richText.map((part) => part.text).join("").replace(/\s+/g, " ").trim().toUpperCase();
  }
  return String(resolved).replace(/\s+/g, " ").trim().toUpperCase();
}

function normalizeName(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/\./g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function titleCase(value) {
  return String(value || "")
    .split(/\s+/)
    .filter(Boolean)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}

function buildOutputFilename(offerFileName, label) {
  const baseName = stripExtension(offerFileName);
  return `${sanitizeFileName(label)} upload template - ${sanitizeFileName(baseName)}.xlsx`;
}

function stripExtension(filename) {
  return String(filename).replace(/\.[^.]+$/, "");
}

function sanitizeFileName(value) {
  return String(value)
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function cloneStyle(style) {
  return structuredCloneSafe(style || {});
}

function cloneValue(value) {
  if (value instanceof Date) return new Date(value.getTime());
  if (value && typeof value === "object") return structuredCloneSafe(value);
  return value;
}

function structuredCloneSafe(value) {
  if (value === undefined || value === null) return value;
  if (typeof structuredClone === "function") return structuredClone(value);
  return JSON.parse(JSON.stringify(value));
}
