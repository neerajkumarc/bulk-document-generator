const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const libre = require("libreoffice-convert");
let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 500,
    height: 700,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
  });
  // mainWindow.setMenu(null);
  mainWindow.loadFile(path.join(__dirname, "renderer", "index.html"));
}

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

ipcMain.on("select-file", async (event) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openFile"],
    filters: [{ name: "Word Document", extensions: ["docx"] }],
  });

  if (!result.canceled) {
    event.reply("file-selected", result.filePaths[0]);
  }
});

ipcMain.on("generate-letters", async (event, { templatePath, names }) => {
  try {
    // Open a dialog for selecting where to save the generated files
    const saveDirectory = await dialog.showOpenDialog(mainWindow, {
      properties: ["openDirectory"],
    });

    if (saveDirectory.canceled || !saveDirectory.filePaths.length) {
      event.reply("error", "No directory selected");
      return;
    }

    const outputDir = saveDirectory.filePaths[0];

    // Create 'pdf' and 'docx' folders in the selected directory
    const pdfDir = path.join(outputDir, "pdf");
    const docxDir = path.join(outputDir, "docx");

    // Ensure both directories exist, create them if they don't
    if (!fs.existsSync(pdfDir)) {
      fs.mkdirSync(pdfDir);
    }

    if (!fs.existsSync(docxDir)) {
      fs.mkdirSync(docxDir);
    }

    event.reply("status", "Generating Documents...");
    const content = fs.readFileSync(templatePath, "binary");
    for (let i = 0; i < names.length; i++) {
      const n = names[i];
      const zip = new PizZip(content);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      doc.setData({
        name: n,
      });

      doc.render();

      const buf = doc.getZip().generate({ type: "nodebuffer" });

      const sanitizedName = n.replace(/[^a-z0-9]/gi, "_").toLowerCase();
      const outputDocxPath = path.join(
        docxDir,
        `document_${sanitizedName}.docx`
      );

      // Save DOCX file in the 'docx' folder
      fs.writeFileSync(outputDocxPath, buf);

      // Convert DOCX to PDF using libreoffice-convert
      const outputPdfPath = path.join(pdfDir, `document_${sanitizedName}.pdf`);

      const docxBuffer = fs.readFileSync(outputDocxPath);

      // Convert DOCX to PDF and wait for completion
      await convertToPdf(docxBuffer, outputPdfPath);
      console.log("Saved PDF file to:", outputPdfPath);

      event.reply("generation-progress", {
        generated: i + 1,
        total: names.length,
      });
    }

    event.reply(
      "letters-generated",
      `Documents generated successfully in ${outputDir}!`
    );
  } catch (error) {
    event.reply("error", `Error generating letters: ${error.message}`);
  }
});

// Wrap libre.convert in a Promise to allow async/await
function convertToPdf(inputBuffer, outputPath) {
  return new Promise((resolve, reject) => {
    libre.convert(inputBuffer, ".pdf", undefined, (err, pdfBuffer) => {
      if (err) {
        return reject(err);
      }
      fs.writeFileSync(outputPath, pdfBuffer);
      resolve();
    });
  });
}
