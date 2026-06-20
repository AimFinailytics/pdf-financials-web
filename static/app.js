const form = document.getElementById("uploadForm");
const fileInput = document.getElementById("fileInput");
const fileList = document.getElementById("fileList");
const dropzone = document.getElementById("dropzone");
const button = document.getElementById("convertButton");
const statusPill = document.getElementById("statusPill");
const resultsPanel = document.getElementById("resultsPanel");
const outputCount = document.getElementById("outputCount");
const downloads = document.getElementById("downloads");
const skipped = document.getElementById("skipped");

function iconRefresh() {
  if (window.lucide) {
    window.lucide.createIcons();
  }
}

function setStatus(text, kind = "") {
  statusPill.textContent = text;
  statusPill.className = `status-pill ${kind}`.trim();
}

function renderFiles() {
  const files = Array.from(fileInput.files || []);
  fileList.innerHTML = "";
  if (!files.length) {
    const empty = document.createElement("span");
    empty.className = "empty";
    empty.textContent = "No PDFs selected";
    fileList.appendChild(empty);
    return;
  }

  for (const file of files) {
    const chip = document.createElement("span");
    chip.className = "file-chip";
    chip.innerHTML = `<i data-lucide="file-text"></i><span>${escapeHtml(file.name)}</span>`;
    fileList.appendChild(chip);
  }
  iconRefresh();
}

function escapeHtml(value) {
  return value.replace(/[&<>"']/g, (char) => {
    const map = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" };
    return map[char];
  });
}

function formatBytes(bytes) {
  if (!bytes) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function renderResults(data) {
  resultsPanel.hidden = false;
  downloads.innerHTML = "";
  outputCount.textContent = `${data.files.length} file${data.files.length === 1 ? "" : "s"}`;

  for (const file of data.files) {
    const row = document.createElement("div");
    row.className = "download-row";
    row.innerHTML = `
      <a href="${file.url}">
        <i data-lucide="${file.name.endsWith(".zip") ? "package" : "file-spreadsheet"}"></i>
        <span>${escapeHtml(file.name)}</span>
      </a>
      <span>${formatBytes(file.size)}</span>
    `;
    downloads.appendChild(row);
  }

  if (data.skipped && data.skipped.length) {
    skipped.hidden = false;
    skipped.innerHTML = data.skipped.map(escapeHtml).join("<br>");
  } else {
    skipped.hidden = true;
    skipped.innerHTML = "";
  }

  iconRefresh();
}

fileInput.addEventListener("change", renderFiles);

dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropzone.classList.add("dragging");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("dragging");
});

dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropzone.classList.remove("dragging");
  fileInput.files = event.dataTransfer.files;
  renderFiles();
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const files = Array.from(fileInput.files || []);
  if (!files.length) {
    setStatus("Select PDFs", "error");
    return;
  }

  const payload = new FormData();
  for (const file of files) {
    payload.append("files", file);
  }

  button.disabled = true;
  setStatus("Processing", "busy");
  resultsPanel.hidden = true;

  try {
    const response = await fetch("/api/convert", {
      method: "POST",
      body: payload,
    });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Conversion failed");
    }
    renderResults(data);
    setStatus("Complete", "done");
  } catch (error) {
    setStatus("Error", "error");
    resultsPanel.hidden = false;
    outputCount.textContent = "";
    downloads.innerHTML = `<div class="skipped">${escapeHtml(error.message)}</div>`;
  } finally {
    button.disabled = false;
  }
});

renderFiles();
window.addEventListener("load", iconRefresh);
