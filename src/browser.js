import { buildImportedFiles } from "./importer.js";

const sourceInput = document.querySelector("#source-file");
const sourceNameEl = document.querySelector("#source-name");
const generateButton = document.querySelector("#generate");
const statusEl = document.querySelector("#status");
const dropzone = document.querySelector("#dropzone");
const themeToggle = document.querySelector("#theme-toggle");

let selectedSourceFile = null;

initTheme();
initFilePicker();
updateButtonState();

generateButton.addEventListener("click", async () => {
  if (!selectedSourceFile) return;

  setStatus("Working... reading the Excel file.", "busy");
  generateButton.disabled = true;

  try {
    const offerBuffer = await selectedSourceFile.arrayBuffer();

    setStatus("Building one Excel file for each populated Rates/Allocation pair.", "busy");
    const result = await buildImportedFiles({
      offerBuffer,
      offerFileName: selectedSourceFile.name,
    });

    if (result.files.length === 1) {
      downloadBlob(
        result.files[0].outputBuffer,
        result.files[0].filename,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      );
      setStatus(`Done. Detected ${result.sourceType}; downloaded ${result.files[0].filename}.`, "success");
    } else {
      setStatus(`Downloading ${result.files.length} Excel files. If the browser asks, allow multiple downloads.`, "busy");
      await downloadFiles(result.files);
      setStatus(`Done. Detected ${result.sourceType}; downloaded ${result.files.length} files for ${result.importedLabels.join(", ")}.`, "success");
    }
  } catch (error) {
    console.error(error);
    setStatus(error.message || String(error), "error");
  } finally {
    updateButtonState();
  }
});

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
  updateButtonState();
}

function updateButtonState() {
  generateButton.disabled = !selectedSourceFile;
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

async function downloadFiles(files) {
  for (const file of files) {
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
