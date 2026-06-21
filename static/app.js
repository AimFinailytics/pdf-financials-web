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
const useAi = document.getElementById("useAi");

const signupOverlay = document.getElementById("signupOverlay");
const signupForm = document.getElementById("signupForm");
const signupEmail = document.getElementById("signupEmail");
const signupName = document.getElementById("signupName");
const signupError = document.getElementById("signupError");

function iconRefresh() {
  if (window.lucide) {
    window.lucide.createIcons();
  }
}

function setStatus(text, kind = "") {
  statusPill.textContent = text;
  statusPill.className = `status-pill ${kind}`.trim();
}

/* ---------- signup gate ---------- */
function showGate() {
  signupOverlay.hidden = false;
  document.body.classList.add("gated");
  iconRefresh();
  setTimeout(() => signupEmail && signupEmail.focus(), 50);
}

function hideGate() {
  signupOverlay.hidden = true;
  document.body.classList.remove("gated");
}

async function checkAccess() {
  // The gate shows by default (markup), so a new visitor always sees it even if
  // this check is slow/fails. Here we only DISMISS it for someone already signed
  // up so returning users skip straight to the converter.
  try {
    const res = await fetch("/api/me");
    const data = await res.json();
    if (data.signed_up) hideGate();
  } catch (_) {
    /* leave the gate up; signing up again just re-captures the email */
  }
}

if (signupForm) {
  signupForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    signupError.hidden = true;
    const email = (signupEmail.value || "").trim();
    if (!email) {
      signupError.textContent = "Please enter your email.";
      signupError.hidden = false;
      return;
    }
    try {
      const res = await fetch("/api/signup", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email, name: (signupName.value || "").trim() }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Signup failed");
      hideGate();
      setStatus("Ready", "");
    } catch (err) {
      signupError.textContent = err.message;
      signupError.hidden = false;
    }
  });
}

/* ---------- file picker ---------- */
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
  payload.append("use_ai", useAi && useAi.checked ? "1" : "0");

  button.disabled = true;
  setStatus("Processing", "busy");
  resultsPanel.hidden = true;

  try {
    const response = await fetch("/api/convert", {
      method: "POST",
      body: payload,
    });
    const data = await response.json();
    if (response.status === 403 && data.code === "signup_required") {
      setStatus("Sign up to continue", "error");
      showGate();
      return;
    }
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
window.addEventListener("load", () => {
  iconRefresh();
  checkAccess();
});
