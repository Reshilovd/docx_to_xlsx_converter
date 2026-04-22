const express = require("express");
const cors = require("cors");
const multer = require("multer");
const { execFile } = require("node:child_process");
const fs = require("node:fs/promises");
const fsSync = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { promisify } = require("node:util");

const execFileAsync = promisify(execFile);

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 30 * 1024 * 1024,
  },
});

app.use(cors());

app.get("/api/health", (_req, res) => {
  res.json({ status: "ok", libreOfficePath: resolveLibreOfficeCommand() || null });
});

function getWindowsLibreOfficeCandidates() {
  const candidates = [];
  const programFiles = process.env.ProgramFiles;
  const programFilesX86 = process.env["ProgramFiles(x86)"];
  const localAppData = process.env.LOCALAPPDATA;

  if (programFiles) {
    candidates.push(path.join(programFiles, "LibreOffice", "program", "soffice.exe"));
  }
  if (programFilesX86) {
    candidates.push(path.join(programFilesX86, "LibreOffice", "program", "soffice.exe"));
  }
  if (localAppData) {
    candidates.push(path.join(localAppData, "Programs", "LibreOffice", "program", "soffice.exe"));
  }

  return candidates;
}

function resolveLibreOfficeCommand() {
  if (process.env.LIBREOFFICE_PATH) {
    return process.env.LIBREOFFICE_PATH;
  }

  if (process.platform === "win32") {
    const detected = getWindowsLibreOfficeCandidates().find((candidate) =>
      fsSync.existsSync(candidate)
    );
    if (detected) {
      return detected;
    }
  }

  return "soffice";
}

function buildLibreOfficeNotFoundMessage() {
  const windowsHint =
    process.platform === "win32"
      ? ' Пример: $env:LIBREOFFICE_PATH="C:\\Program Files\\LibreOffice\\program\\soffice.exe"'
      : "";
  return `LibreOffice не найден. Установите LibreOffice и добавьте soffice в PATH или задайте LIBREOFFICE_PATH.${windowsHint}`;
}

async function runLibreOfficeConvert(inputPath, outDir) {
  const command = resolveLibreOfficeCommand();
  const args = [
    "--headless",
    "--convert-to",
    "xlsx",
    "--outdir",
    outDir,
    inputPath,
  ];

  try {
    return await execFileAsync(command, args, { windowsHide: true });
  } catch (error) {
    if (error && error.code === "ENOENT") {
      const wrappedError = new Error(buildLibreOfficeNotFoundMessage());
      wrappedError.code = "LIBREOFFICE_NOT_FOUND";
      throw wrappedError;
    }
    throw error;
  }
}

app.post("/api/convert", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: "Файл не загружен." });
  }

  if (!req.file.originalname.toLowerCase().endsWith(".docx")) {
    return res.status(400).json({ error: "Поддерживаются только .docx файлы." });
  }

  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "docx-to-xlsx-"));
  const sourcePath = path.join(tempDir, "source.docx");
  const outputPath = path.join(tempDir, "source.xlsx");

  try {
    await fs.writeFile(sourcePath, req.file.buffer);
    await runLibreOfficeConvert(sourcePath, tempDir);

    const xlsxBuffer = await fs.readFile(outputPath);
    const downloadName = req.file.originalname.replace(/\.docx$/i, ".xlsx");

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${downloadName}"`);
    res.send(xlsxBuffer);
  } catch (error) {
    const details =
      error.code === "LIBREOFFICE_NOT_FOUND"
        ? error.message
        : error.stderr || error.message || "Неизвестная ошибка.";
    res.status(500).json({
      error: "Ошибка конвертации через LibreOffice.",
      details,
    });
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
});

const port = Number(process.env.PORT || 4000);
app.listen(port, () => {
  // eslint-disable-next-line no-console
  console.log(`Backend started at http://localhost:${port}`);
});
