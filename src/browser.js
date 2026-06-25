import { buildImportedFiles } from "./importer.js";

const sourceInput = document.querySelector("#source-file");
const sourceNameEl = document.querySelector("#source-name");
const statusEl = document.querySelector("#status");
const dropzone = document.querySelector("#dropzone");
const themeToggle = document.querySelector("#theme-toggle");
const googleSheetUrlInput = document.querySelector("#google-sheet-url");
const googleSheetRunButton = document.querySelector("#google-sheet-run");

let selectedSourceFile = null;
let activeRunId = 0;

initTheme();
initFilePicker();
initGoogleSheetImport();

function initFilePicker() {
  sourceInput.addEventListener("change", () => {
    setSelectedSourceFile(sourceInput.files[0] || null);
  });

  dropzone.addEventListener("click", () => {
    sourceInput.click();
  });

  dropzone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      sourceInput.click();
    }
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.dataset.dragging = "true";
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.dataset.dragging = "false";
    });
  });

  dropzone.addEventListener("drop", (event) => {
    const file = event.dataTransfer?.files?.[0] || null;
    if (!file) return;
    setSelectedSourceFile(file);
  });
}

function setSelectedSourceFile(file) {
  selectedSourceFile = file;
  sourceNameEl.textContent = file ? file.name : "No source workbook selected";
  dropzone.dataset.hasFile = file ? "true" : "false";
  if (file) {
    processSelectedSourceFile(file);
  } else {
    setStatus("Ready. Drop or choose a workbook to start.", "ready");
  }
}

async function processSelectedSourceFile(file) {
  const runId = startRun();
  dropzone.dataset.processing = "true";
  setGoogleSheetControlsDisabled(true);
  setStatus("Working... reading the Excel file.", "busy");

  try {
    const offerBuffer = await file.arrayBuffer();
    if (runId !== activeRunId) return;

    await processOfferBuffer({
      offerBuffer,
      offerFileName: file.name,
      runId,
    });
  } catch (error) {
    if (runId !== activeRunId) return;
    console.error(error);
    setStatus(error.message || String(error), "error");
  } finally {
    if (runId === activeRunId) {
      dropzone.dataset.processing = "false";
      setGoogleSheetControlsDisabled(false);
    }
  }
}

function initGoogleSheetImport() {
  googleSheetRunButton.addEventListener("click", () => {
    processGoogleSheetUrl();
  });

  googleSheetUrlInput.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    processGoogleSheetUrl();
  });
}

async function processGoogleSheetUrl() {
  const runId = startRun();
  dropzone.dataset.processing = "true";
  setGoogleSheetControlsDisabled(true);

  try {
    const request = buildGoogleSheetExportRequest(googleSheetUrlInput.value);
    setStatus("Working... loading the Google Sheet export.", "busy");

    const response = await fetch(request.url, {
      credentials: "omit",
      redirect: "follow",
      referrerPolicy: "no-referrer",
    });
    if (runId !== activeRunId) return;

    if (!response.ok) {
      throw new Error(`Google Sheets export failed with status ${response.status}. Share the Sheet as public or publish it to the web, then try again.`);
    }

    const offerBuffer = await response.arrayBuffer();
    if (runId !== activeRunId) return;

    if (!looksLikeXlsx(offerBuffer)) {
      throw new Error("Google returned a web page instead of an Excel export. Share the Sheet as public or publish it to the web, then try again.");
    }

    await processOfferBuffer({
      offerBuffer,
      offerFileName: request.fileName,
      runId,
    });
  } catch (error) {
    if (runId !== activeRunId) return;
    console.error(error);
    setStatus(getGoogleSheetErrorMessage(error), "error");
  } finally {
    if (runId === activeRunId) {
      dropzone.dataset.processing = "false";
      setGoogleSheetControlsDisabled(false);
    }
  }
}

async function processOfferBuffer({ offerBuffer, offerFileName, runId }) {
  setStatus("Building one Excel file for each populated Rates/Allocation pair.", "busy");
  const result = await buildImportedFiles({
    offerBuffer,
    offerFileName,
  });
  if (runId !== activeRunId) return;

  if (result.files.length === 1) {
    downloadBlob(
      result.files[0].outputBuffer,
      result.files[0].filename,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    showSuccess(`Done. Detected ${result.sourceType}; downloaded ${result.files[0].filename}.`);
    return;
  }

  setStatus(`Downloading ${result.files.length} Excel files. If the browser asks, allow multiple downloads.`, "busy");
  await downloadFiles(result.files, runId);
  if (runId !== activeRunId) return;
  showSuccess(`Done. Detected ${result.sourceType}; downloaded ${result.files.length} files for ${result.importedLabels.join(", ")}.`);
}

function buildGoogleSheetExportRequest(value) {
  const rawValue = value.trim();
  if (!rawValue) {
    throw new Error("Paste a public Google Sheet URL first.");
  }

  let url;
  try {
    url = new URL(rawValue);
  } catch {
    throw new Error("Paste the full Google Sheet link, starting with https://docs.google.com/spreadsheets/.");
  }

  if (url.protocol !== "https:" || url.hostname !== "docs.google.com") {
    throw new Error("For security, this experimental import only accepts https://docs.google.com/spreadsheets/ links.");
  }

  const existingFormat = (url.searchParams.get("format") || url.searchParams.get("output") || "").toLowerCase();
  if (url.pathname.includes("/spreadsheets/") && (url.pathname.includes("/export") || url.pathname.includes("/pub")) && existingFormat === "xlsx") {
    return {
      url: url.toString(),
      fileName: getGoogleSheetFileName(url),
    };
  }

  const publishedMatch = /^\/spreadsheets\/d\/e\/([^/]+)/.exec(url.pathname);
  if (publishedMatch) {
    return {
      url: `https://docs.google.com/spreadsheets/d/e/${publishedMatch[1]}/pub?output=xlsx`,
      fileName: getGoogleSheetFileName(url, publishedMatch[1]),
    };
  }

  const sheetMatch = /^\/spreadsheets\/d\/([^/]+)/.exec(url.pathname);
  if (sheetMatch) {
    return {
      url: `https://docs.google.com/spreadsheets/d/${sheetMatch[1]}/export?format=xlsx`,
      fileName: getGoogleSheetFileName(url, sheetMatch[1]),
    };
  }

  throw new Error("Paste a Google Sheet link from docs.google.com/spreadsheets.");
}

function getGoogleSheetFileName(url, fallbackId = "") {
  const urlTitle = url.searchParams.get("title") || "";
  const idPart = fallbackId ? fallbackId.slice(0, 8) : "export";
  const safeTitle = urlTitle.trim() || `Google Sheet ${idPart}`;
  return `${safeTitle}.xlsx`;
}

function looksLikeXlsx(buffer) {
  const bytes = new Uint8Array(buffer, 0, Math.min(4, buffer.byteLength));
  return bytes[0] === 0x50 && bytes[1] === 0x4b;
}

function getGoogleSheetErrorMessage(error) {
  const message = error.message || String(error);
  if (error instanceof TypeError && /fetch/i.test(message)) {
    return "The browser could not fetch that Google Sheet. Share it as public or publish it to the web, then try again.";
  }
  return message;
}

function setGoogleSheetControlsDisabled(isDisabled) {
  googleSheetUrlInput.disabled = isDisabled;
  googleSheetRunButton.disabled = isDisabled;
}

function startRun() {
  activeRunId += 1;
  return activeRunId;
}

function initTheme() {
  const storedTheme = localStorage.getItem("theme");
  const prefersDark = window.matchMedia?.("(prefers-color-scheme: dark)").matches;
  const initialTheme = storedTheme === "light" || storedTheme === "dark"
    ? storedTheme
    : prefersDark
      ? "dark"
      : "light";

  setTheme(initialTheme);

  themeToggle.addEventListener("click", () => {
    const nextTheme = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
    setTheme(nextTheme);
    localStorage.setItem("theme", nextTheme);
  });
}

function setTheme(theme) {
  document.documentElement.dataset.theme = theme;
  const nextThemeName = theme === "dark" ? "light" : "dark";
  const label = `Switch to ${nextThemeName} mode`;
  themeToggle.setAttribute("aria-label", label);
  themeToggle.title = label;
}

function setStatus(message, type) {
  statusEl.textContent = message;
  statusEl.dataset.type = type;
}

function showSuccess(message) {
  setStatus(message, "success");

  requestAnimationFrame(() => {
    if (typeof window.confetti !== "function") return;

    const bounds = statusEl.getBoundingClientRect();
    const x = clamp((bounds.left + bounds.width / 2) / window.innerWidth, 0.05, 0.95);
    const y = clamp((bounds.top + bounds.height / 2) / window.innerHeight, 0.05, 0.95);

    window.confetti({
      particleCount: 100,
      spread: 70,
      origin: { x, y },
      zIndex: 1000,
      disableForReducedMotion: true,
    });
  });
}

function clamp(value, minimum, maximum) {
  return Math.min(Math.max(value, minimum), maximum);
}

async function downloadFiles(files, runId) {
  for (const file of files) {
    if (runId !== activeRunId) return;
    downloadBlob(
      file.outputBuffer,
      file.filename,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    await wait(250);
  }
}

function wait(milliseconds) {
  return new Promise((resolve) => {
    setTimeout(resolve, milliseconds);
  });
}

function downloadBlob(buffer, filename, type) {
  const blob = new Blob([buffer], { type });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(link.href), 2000);
}
