import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { buildImportedFiles } from "./src/importer.js";

const niagaraPath = "/Users/bogdan/Desktop/Work/Rate Update_Niagara Falls_New York_Miami & Caribbean Cruise_09-06-2026 (2).xls";
const pandasPath = "/Users/bogdan/Desktop/Work/A71190 - RATE AMENDMENTS  - 19686 - 64350 - se-modern-shanghai-chengdu-giant-pandas.xls";

async function buildFilesAndInspect(sourcePath) {
  const result = await buildImportedFiles({
    offerBuffer: await fs.readFile(sourcePath),
    offerFileName: path.basename(sourcePath),
  });

  const files = [];
  for (const file of result.files) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(file.outputBuffer);
    files.push({
      filename: file.filename,
      label: file.label,
      sheetNames: workbook.worksheets.map((sheet) => sheet.name),
      workbook,
    });
  }

  return {
    sourceType: result.sourceType,
    labels: result.importedLabels,
    files,
  };
}

async function buildFilesAndInspectBuffer(offerBuffer, offerFileName) {
  const result = await buildImportedFiles({ offerBuffer, offerFileName });

  const files = [];
  for (const file of result.files) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(file.outputBuffer);
    files.push({
      filename: file.filename,
      label: file.label,
      sheetNames: workbook.worksheets.map((sheet) => sheet.name),
      workbook,
    });
  }

  return {
    sourceType: result.sourceType,
    labels: result.importedLabels,
    files,
  };
}

async function createOutOfOrderCabinWorkbook() {
  const workbook = new ExcelJS.Workbook();
  addRateSheet(workbook, "Eastern - Balcony Cabin", {
    sellRate: 777,
    supplierRate: 888,
    startDate: new Date(2026, 6, 1),
    departure: "LHR",
    destination: "YYZ",
  });
  addRateSheet(workbook, "Eastern - Inside Cabin", {
    sellRate: 111,
    supplierRate: 222,
    startDate: new Date(2026, 6, 2),
    departure: "MAN",
    destination: "YYZ",
  });
  addAllocationSheet(workbook, "Eastern Inside -Allo", {
    startDate: new Date(2026, 6, 2),
    rooms: 11,
  });
  addAllocationSheet(workbook, "Eastern Balcony -Allo", {
    startDate: new Date(2026, 6, 1),
    rooms: 33,
  });
  return workbook.xlsx.writeBuffer();
}

function addRateSheet(workbook, name, { sellRate, supplierRate, startDate, departure, destination }) {
  const sheet = workbook.addWorksheet(name);
  sheet.getCell("A1").value = "TOTAL SELL RATE (Secret Escapes)";
  sheet.getCell("B1").value = "TOTAL SUPPLIER RATE (hotel's or tour operator's site)";
  sheet.getCell("F1").value = "START DATE (YYYY-MM-DD)";
  sheet.getCell("G1").value = "NUMBER OF NIGHTS";
  sheet.getCell("I1").value = "DEPARTURE AIRPORT CODE";
  sheet.getCell("J1").value = "DESTINATION AIRPORT CODE";
  sheet.getCell("P1").value = "OUTBOUND OVERNIGHT FLIGHT";
  sheet.getCell("Q1").value = "INBOUND OVERNIGHT FLIGHT";
  sheet.getCell("A2").value = sellRate;
  sheet.getCell("B2").value = supplierRate;
  sheet.getCell("F2").value = startDate;
  sheet.getCell("G2").value = 10;
  sheet.getCell("I2").value = departure;
  sheet.getCell("J2").value = destination;
  sheet.getCell("P2").value = false;
  sheet.getCell("Q2").value = true;
}

function addAllocationSheet(workbook, name, { startDate, rooms }) {
  const sheet = workbook.addWorksheet(name);
  sheet.getCell("A1").value = "START DATE (YYYY-MM-DD)";
  sheet.getCell("B1").value = "NO. OF ROOMS ALLOCATED";
  sheet.getCell("A2").value = startDate;
  sheet.getCell("B2").value = rooms;
}

async function createSuffixedParenWorkbook() {
  // Mirrors the Mein Schiff format: an Offer Details sheet, rate tabs named
  // `Rates (N) XXXX`, allocation tabs `Allocation (N)`, SHARED allocation header.
  const workbook = new ExcelJS.Workbook();

  const details = workbook.addWorksheet("Offer Details");
  details.getCell("A21").value = "ANGEBOTSDETAILS / offer details";
  details.getCell("B21").value = "ANGEBOT 1 / offer 1";
  details.getCell("C21").value = "ANGEBOT 2 / offer 2";
  details.getCell("B22").value = "Some offer title";
  details.getCell("C22").value = "Some offer title";

  addTotalRateSheet(workbook, "Rates (1) INAO", { total: 3598, supplier: 3997, departure: "FRA" });
  addSharedAllocationSheet(workbook, "Allocation (1)", { rooms: 3 });
  addTotalRateSheet(workbook, "Rates (2) AUAO", { total: 3798, supplier: 4220, departure: "DUS" });
  addSharedAllocationSheet(workbook, "Allocation (2)", { rooms: 2 });

  return workbook.xlsx.writeBuffer();
}

function addTotalRateSheet(workbook, name, { total, supplier, departure }) {
  const sheet = workbook.addWorksheet(name);
  sheet.getCell("C1").value = "TOTAL RATE (AUTOMATIC) (Secret Escapes)";
  sheet.getCell("H1").value = "START DATE (YYYY-MM-DD)";
  sheet.getCell("K1").value = "DEPARTURE AIRPORT CODE";
  sheet.getCell("C2").value = total;
  sheet.getCell("D2").value = supplier;
  sheet.getCell("H2").value = new Date(2026, 10, 17);
  sheet.getCell("I2").value = 8;
  sheet.getCell("K2").value = departure;
  sheet.getCell("L2").value = "LPA";
}

function addSharedAllocationSheet(workbook, name, { rooms }) {
  const sheet = workbook.addWorksheet(name);
  sheet.getCell("A1").value = "ANGEBOT";
  sheet.getCell("A2").value = "Reise-Startdatum start date";
  sheet.getCell("B2").value = "SHARED Allocation";
  sheet.getCell("A3").value = new Date(2026, 10, 17);
  sheet.getCell("B3").value = rooms;
}

async function fileExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

function cellValue(workbook, sheetName, address) {
  return workbook.getWorksheet(sheetName).getCell(address).value;
}

function assertEqual(actual, expected, label) {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  }
}

function assertDate(value, year, month, day, label) {
  if (!(value instanceof Date)) {
    throw new Error(`${label}: expected Date, got ${Object.prototype.toString.call(value)}`);
  }
  assertEqual(value.getFullYear(), year, `${label} year`);
  assertEqual(value.getMonth() + 1, month, `${label} month`);
  assertEqual(value.getDate(), day, `${label} day`);
}

const niagaraFiles = await fileExists(niagaraPath)
  ? await buildFilesAndInspect(niagaraPath)
  : null;

if (niagaraFiles) {
  assertEqual(niagaraFiles.sourceType, "paired rates/allocation workbook", "niagara source type");
  assertEqual(niagaraFiles.files.length, 3, "niagara file count");
  assertEqual(
    niagaraFiles.labels.join("|"),
    "Eastern Inside|Eastern Outside|Eastern Balcony",
    "niagara labels",
  );
  assertEqual(niagaraFiles.files[0].sheetNames.join("|"), "Rates (1)|Allocation (1)", "niagara first tabs");
  assertEqual(niagaraFiles.files[1].sheetNames.join("|"), "Rates (2)|Allocation (2)", "niagara second tabs");
  assertEqual(niagaraFiles.files[2].sheetNames.join("|"), "Rates (3)|Allocation (3)", "niagara third tabs");
  assertEqual(cellValue(niagaraFiles.files[0].workbook, "Rates (1)", "A2"), 3798, "niagara inside A2");
  assertEqual(cellValue(niagaraFiles.files[0].workbook, "Rates (1)", "B2"), 4298, "niagara inside B2");
  assertEqual(cellValue(niagaraFiles.files[0].workbook, "Rates (1)", "I2"), "LHR", "niagara inside departure I2");
  assertEqual(cellValue(niagaraFiles.files[0].workbook, "Rates (1)", "J2"), "YYZ", "niagara inside destination J2");
  assertEqual(cellValue(niagaraFiles.files[0].workbook, "Rates (1)", "P2"), "FALSE", "niagara inside outbound P2");
  assertEqual(cellValue(niagaraFiles.files[0].workbook, "Rates (1)", "Q2"), "TRUE", "niagara inside inbound Q2");
  assertEqual(cellValue(niagaraFiles.files[0].workbook, "Allocation (1)", "B2"), 3, "niagara inside allocation B2");
  assertDate(cellValue(niagaraFiles.files[0].workbook, "Rates (1)", "F2"), 2026, 7, 1, "niagara inside rates F2");
  assertDate(cellValue(niagaraFiles.files[0].workbook, "Allocation (1)", "A2"), 2026, 7, 1, "niagara inside allocation A2");
}

const outOfOrderCabinFiles = await buildFilesAndInspectBuffer(
  await createOutOfOrderCabinWorkbook(),
  "Out of order cabins.xlsx",
);
assertEqual(outOfOrderCabinFiles.files.length, 2, "out-of-order cabin file count");
assertEqual(outOfOrderCabinFiles.labels.join("|"), "Eastern Balcony|Eastern Inside", "out-of-order cabin labels");
assertEqual(cellValue(outOfOrderCabinFiles.files[0].workbook, "Rates (1)", "A2"), 777, "out-of-order balcony rate A2");
assertEqual(cellValue(outOfOrderCabinFiles.files[0].workbook, "Allocation (1)", "B2"), 33, "out-of-order balcony allocation B2");
assertEqual(cellValue(outOfOrderCabinFiles.files[1].workbook, "Rates (2)", "A2"), 111, "out-of-order inside rate A2");
assertEqual(cellValue(outOfOrderCabinFiles.files[1].workbook, "Allocation (2)", "B2"), 11, "out-of-order inside allocation B2");

const suffixedParenFiles = await buildFilesAndInspectBuffer(
  await createSuffixedParenWorkbook(),
  "Mein Schiff format.xlsx",
);
assertEqual(suffixedParenFiles.sourceType, "paired rates/allocation workbook", "suffixed-paren source type");
assertEqual(suffixedParenFiles.files.length, 2, "suffixed-paren file count");
assertEqual(suffixedParenFiles.labels.join("|"), "Inao|Auao", "suffixed-paren labels");
assertEqual(suffixedParenFiles.files[0].sheetNames.join("|"), "Rates (1)|Allocation (1)", "suffixed-paren first tabs");
assertEqual(cellValue(suffixedParenFiles.files[0].workbook, "Rates (1)", "A2"), 3598, "suffixed-paren total rate A2");
assertEqual(cellValue(suffixedParenFiles.files[0].workbook, "Rates (1)", "I2"), "FRA", "suffixed-paren departure I2");
assertEqual(cellValue(suffixedParenFiles.files[0].workbook, "Allocation (1)", "B2"), 3, "suffixed-paren allocation B2");
assertEqual(cellValue(suffixedParenFiles.files[1].workbook, "Rates (2)", "A2"), 3798, "suffixed-paren second total rate A2");

const pandasFiles = await fileExists(pandasPath)
  ? await buildFilesAndInspect(pandasPath)
  : null;

if (pandasFiles) {
  assertEqual(pandasFiles.files.length, 1, "pandas file count");
  assertEqual(cellValue(pandasFiles.files[0].workbook, "Rates (1)", "A2"), 3198, "pandas A2");
  assertEqual(cellValue(pandasFiles.files[0].workbook, "Allocation (1)", "B2"), 5, "pandas allocation B2");
}

console.log(JSON.stringify({
  niagara: {
    count: niagaraFiles?.files.length ?? 0,
    labels: niagaraFiles?.labels ?? [],
    firstFile: niagaraFiles?.files[0]?.filename ?? "sample not found",
  },
  outOfOrderCabins: {
    count: outOfOrderCabinFiles.files.length,
    labels: outOfOrderCabinFiles.labels,
  },
  pandas: {
    count: pandasFiles?.files.length ?? 0,
    labels: pandasFiles?.labels ?? [],
    firstFile: pandasFiles?.files[0]?.filename ?? "sample not found",
  },
}, null, 2));
