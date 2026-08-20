const SERVER_URL = "http://127.0.0.1:8765";

const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const fileList = document.getElementById("fileList");
const targetFormat = document.getElementById("targetFormat");
const convertButton = document.getElementById("convertButton");
const clearButton = document.getElementById("clearButton");
const refreshButton = document.getElementById("refreshButton");
const serverStatus = document.getElementById("serverStatus");
const message = document.getElementById("message");

let selectedFiles = [];
let serverReady = false;

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function setMessage(text, isError = false) {
  message.textContent = text;
  message.classList.toggle("error", isError);
}

function updateConvertState() {
  convertButton.disabled = !serverReady || selectedFiles.length === 0;
}

function renderFiles() {
  fileList.textContent = "";

  if (!selectedFiles.length) {
    const item = document.createElement("li");
    item.className = "empty";
    item.textContent = "No files selected";
    fileList.appendChild(item);
    updateConvertState();
    return;
  }

  for (const file of selectedFiles) {
    const item = document.createElement("li");
    const name = document.createElement("span");
    const size = document.createElement("span");

    name.className = "name";
    size.className = "size";
    name.textContent = file.name;
    size.textContent = formatBytes(file.size);

    item.append(name, size);
    fileList.appendChild(item);
  }

  updateConvertState();
}

async function checkServer() {
  serverStatus.textContent = "Checking local helper...";
  serverStatus.className = "status checking";
  serverReady = false;
  updateConvertState();

  try {
    const response = await fetch(`${SERVER_URL}/health`, { cache: "no-store" });
    if (!response.ok) throw new Error("Helper did not respond");

    const body = await response.json();
    serverReady = true;
    serverStatus.textContent = body.officePdf
      ? "Local helper ready, Office export available"
      : "Local helper ready, Office export unavailable";
    serverStatus.className = "status ready";
  } catch (error) {
    serverStatus.textContent = "Start the local helper, then try again";
    serverStatus.className = "status offline";
  }

  updateConvertState();
}

function addFiles(fileCollection) {
  selectedFiles = [...selectedFiles, ...Array.from(fileCollection)];
  setMessage("");
  renderFiles();
}

async function downloadBlob(blob, fallbackName) {
  const url = URL.createObjectURL(blob);
  try {
    await chrome.downloads.download({
      url,
      filename: fallbackName,
      saveAs: true
    });
  } finally {
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
}

function fileNameFromDisposition(disposition) {
  if (!disposition) return "";
  const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8Match) return decodeURIComponent(utf8Match[1].replace(/"/g, ""));

  const plainMatch = disposition.match(/filename="?([^"]+)"?/i);
  return plainMatch ? plainMatch[1] : "";
}

async function convertFiles() {
  if (!selectedFiles.length) return;

  const formData = new FormData();
  formData.append("target", targetFormat.value);
  for (const file of selectedFiles) {
    formData.append("files", file, file.name);
  }

  convertButton.disabled = true;
  setMessage("Converting locally...");

  try {
    const response = await fetch(`${SERVER_URL}/convert`, {
      method: "POST",
      body: formData
    });

    if (!response.ok) {
      let detail = "Conversion failed";
      try {
        const errorBody = await response.json();
        detail = errorBody.error || detail;
      } catch {
        detail = await response.text();
      }
      throw new Error(detail);
    }

    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition");
    const filename = fileNameFromDisposition(disposition)
      || `converted.${selectedFiles.length > 1 ? "zip" : targetFormat.value}`;

    await downloadBlob(blob, filename);
    setMessage("Done. Chrome will ask where to save the converted file.");
  } catch (error) {
    setMessage(error.message, true);
  } finally {
    updateConvertState();
  }
}

fileInput.addEventListener("change", () => {
  addFiles(fileInput.files);
  fileInput.value = "";
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragging");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragging");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragging");
  addFiles(event.dataTransfer.files);
});

clearButton.addEventListener("click", () => {
  selectedFiles = [];
  setMessage("");
  renderFiles();
});

refreshButton.addEventListener("click", checkServer);
convertButton.addEventListener("click", convertFiles);

renderFiles();
checkServer();
