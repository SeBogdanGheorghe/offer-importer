import { buildImportedFiles } from "./importer.js";

const sourceInput = document.querySelector("#source-file");
const sourceNameEl = document.querySelector("#source-name");
const statusEl = document.querySelector("#status");
const dropzone = document.querySelector("#dropzone");
const themeToggle = document.querySelector("#theme-toggle");

let selectedSourceFile = null;
let activeRunId = 0;

initTheme();
initFilePicker();

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
  const runId = activeRunId + 1;
  activeRunId = runId;
  dropzone.dataset.processing = "true";
  setStatus("Working... reading the Excel file.", "busy");

  try {
    const offerBuffer = await file.arrayBuffer();
    if (runId !== activeRunId) return;

    setStatus("Building one Excel file for each populated Rates/Allocation pair.", "busy");
    const result = await buildImportedFiles({
      offerBuffer,
      offerFileName: file.name,
    });
    if (runId !== activeRunId) return;

    if (result.files.length === 1) {
      downloadBlob(
        result.files[0].outputBuffer,
        result.files[0].filename,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      );
      showSuccess(`Done. Detected ${result.sourceType}; downloaded ${result.files[0].filename}.`);
    } else {
      setStatus(`Downloading ${result.files.length} Excel files. If the browser asks, allow multiple downloads.`, "busy");
      await downloadFiles(result.files, runId);
      if (runId !== activeRunId) return;
      showSuccess(`Done. Detected ${result.sourceType}; downloaded ${result.files.length} files for ${result.importedLabels.join(", ")}.`);
    }
  } catch (error) {
    if (runId !== activeRunId) return;
    console.error(error);
    setStatus(error.message || String(error), "error");
  } finally {
    if (runId === activeRunId) {
      dropzone.dataset.processing = "false";
    }
  }
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
