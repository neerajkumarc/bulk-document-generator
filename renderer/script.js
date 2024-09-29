const { ipcRenderer } = require("electron");
const generateBtn = document.getElementById("generate");
const selectTemplateBtn = document.getElementById("selectTemplate");
let selectedFilePath = "";
let nameCount = 0;

selectTemplateBtn.addEventListener("click", () => {
  ipcRenderer.send("select-file");
});

ipcRenderer.on("file-selected", (event, filePath) => {
  selectedFilePath = filePath;
  const selectedFile = document.getElementById("selectedFile");
  selectedFile.textContent = `Selected: ${filePath}`;
  selectedFile.classList.add("fade-in");
});

document.getElementById("names").addEventListener("input", () => {
  const names = document
    .getElementById("names")
    .value.split("\n")
    .filter((name) => name.trim() !== "");
  nameCount = names.length;
  const nameCountElement = document.getElementById("nameCount");
  nameCountElement.textContent = `Names added: ${nameCount}`;
  nameCountElement.classList.add("fade-in");
});

generateBtn.addEventListener("click", () => {
  if (!selectedFilePath) {
    showStatus("Please select a template file first.", "error");
    return;
  }
  if (nameCount === 0) {
    showStatus("Please add names before generating letters.", "error");
    return;
  }
  const names = document
    .getElementById("names")
    .value.split("\n")
    .filter((name) => name.trim() !== "")
    .map((name) => name.replace(/(^\w|\s\w)/g, (m) => m.toUpperCase()).trim());

  ipcRenderer.send("generate-letters", {
    templatePath: selectedFilePath,
    names,
  });
});

ipcRenderer.on("status", (event, message) => {
  document.getElementById("loading").style.display = "block";
  showStatus("Generating letters, please wait...");
  generateBtn.disabled = true;
  selectTemplateBtn.disabled = true;
});

ipcRenderer.on("letters-generated", (event, message) => {
  document.getElementById("loading").style.display = "none";
  showStatus(message, "success");
  generateBtn.disabled = false;
  selectTemplateBtn.disabled = false;
});

ipcRenderer.on("generation-progress", (event, progress) => {
  showStatus(`Generating letters: ${progress.generated}/${progress.total}`);
});

ipcRenderer.on("error", (event, message) => {
  document.getElementById("loading").style.display = "none";
  showStatus(message, "error");
});

function showStatus(message, type = "") {
  const statusElement = document.getElementById("status");
  statusElement.textContent = message;
  statusElement.className = `status ${type} fade-in`;
}
