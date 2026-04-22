import { useState } from "react";
import {
  convertCheckedXlsxToSpansBlob,
  convertDocxFilesToXlsxBlob,
  inspectXlsxColumns,
} from "./docxToXlsx";

function App() {
  const [docxFiles, setDocxFiles] = useState([]);
  const [checkedXlsxFile, setCheckedXlsxFile] = useState(null);
  const [classCodes, setClassCodes] = useState([]);
  const [detectedColumns, setDetectedColumns] = useState(0);
  const [isConvertingDocx, setIsConvertingDocx] = useState(false);
  const [isConvertingSpans, setIsConvertingSpans] = useState(false);
  const [message, setMessage] = useState(
    "Шаг 1: загрузите DOCX и скачайте XLSX для ручной проверки."
  );

  const inferClassCode = (fileName, index) => {
    const base = fileName.replace(/\.docx$/i, "").toLowerCase();
    const suffix = base.match(/[_-]([a-zа-я]{2,5})$/i);
    if (suffix) {
      return suffix[1];
    }
    return `lang_${index + 1}`;
  };

  const onDocxFileChange = (event) => {
    const selectedFiles = Array.from(event.target.files || []);
    setDocxFiles(selectedFiles);
    setClassCodes(selectedFiles.map((file, index) => inferClassCode(file.name, index)));
    setMessage(
      selectedFiles.length > 0
        ? `Выбрано файлов: ${selectedFiles.length}`
        : "Файлы не выбраны."
    );
  };

  const onConvertDocxToColumns = async () => {
    if (docxFiles.length === 0) {
      setMessage("Сначала выберите хотя бы один .docx файл.");
      return;
    }

    if (docxFiles.some((file) => !file.name.toLowerCase().endsWith(".docx"))) {
      setMessage("Поддерживаются только .docx файлы.");
      return;
    }

    setIsConvertingDocx(true);
    setMessage("Шаг 1: формируем XLSX по колонкам...");

    try {
      const blob = await convertDocxFilesToXlsxBlob(docxFiles, { mode: "columns" });
      const outputName =
        docxFiles.length === 1
          ? `${docxFiles[0].name.replace(/\.docx$/i, "")}-columns.xlsx`
          : "combined-columns.xlsx";
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = outputName;
      link.click();
      URL.revokeObjectURL(url);
      setMessage(`Готово: ${outputName}. Проверьте файл и загрузите его во 2-й шаг.`);
    } catch (error) {
      setMessage(`Ошибка конвертации: ${error.message}`);
    } finally {
      setIsConvertingDocx(false);
    }
  };

  const onCheckedXlsxChange = async (event) => {
    const [file] = Array.from(event.target.files || []);
    setCheckedXlsxFile(file || null);
    if (!file) {
      setDetectedColumns(0);
      return;
    }

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setMessage("Для шага 2 нужно загрузить .xlsx файл.");
      setDetectedColumns(0);
      return;
    }

    try {
      const columns = await inspectXlsxColumns(file);
      setDetectedColumns(columns);
      setClassCodes((prev) => {
        const next = [];
        for (let i = 0; i < columns; i += 1) {
          next[i] = (prev[i] || `lang_${i + 1}`).trim();
        }
        return next;
      });
      setMessage(`Шаг 2: найдено колонок ${columns}. Укажите CSS-классы и запускайте склейку.`);
    } catch (error) {
      setMessage(`Ошибка чтения XLSX: ${error.message}`);
      setDetectedColumns(0);
    }
  };

  const onMergeCheckedXlsx = async () => {
    if (!checkedXlsxFile) {
      setMessage("Сначала загрузите проверенный .xlsx файл.");
      return;
    }
    if (!checkedXlsxFile.name.toLowerCase().endsWith(".xlsx")) {
      setMessage("Поддерживается только .xlsx файл на шаге 2.");
      return;
    }
    if (
      detectedColumns === 0 ||
      classCodes.length !== detectedColumns ||
      classCodes.some((code) => !code.trim())
    ) {
      setMessage("Укажите CSS-класс для каждой колонки из XLSX.");
      return;
    }

    setIsConvertingSpans(true);
    setMessage("Шаг 2: склеиваем проверенный XLSX в span...");
    try {
      const blob = await convertCheckedXlsxToSpansBlob(
        checkedXlsxFile,
        classCodes.map((code) => code.trim())
      );
      const outputName = `${checkedXlsxFile.name.replace(/\.xlsx$/i, "")}-spans.xlsx`;
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = outputName;
      link.click();
      URL.revokeObjectURL(url);
      setMessage(`Готово: ${outputName}`);
    } catch (error) {
      setMessage(`Ошибка склейки: ${error.message}`);
    } finally {
      setIsConvertingSpans(false);
    }
  };

  return (
    <main className="container">
      <h1>DOCX to XLSX Converter</h1>
      <p className="description">
        Процесс разделен на 2 шага: сначала DOCX в XLSX (проверка вручную),
        затем загрузка проверенного XLSX и склейка в один HTML `span`-блок.
      </p>

      <div className="card">
        <h2 className="section-title">Шаг 1. DOCX → XLSX (колонки)</h2>
        <input type="file" accept=".docx" multiple onChange={onDocxFileChange} />
        <button type="button" onClick={onConvertDocxToColumns} disabled={isConvertingDocx}>
          {isConvertingDocx ? "Формируем XLSX..." : "Скачать XLSX для проверки"}
        </button>
      </div>

      <div className="card">
        <h2 className="section-title">Шаг 2. Проверенный XLSX → span</h2>
        <input type="file" accept=".xlsx" onChange={onCheckedXlsxChange} />
        {detectedColumns > 0 ? (
          <p className="hint">Найдено колонок: {detectedColumns}</p>
        ) : null}
        {detectedColumns > 0
          ? Array.from({ length: detectedColumns }, (_, index) => (
              <label key={`class-${index}`} className="class-code-row">
                <span>Колонка {index + 1}</span>
                <input
                  type="text"
                  value={classCodes[index] || ""}
                  onChange={(event) => {
                    const next = [...classCodes];
                    next[index] = event.target.value;
                    setClassCodes(next);
                  }}
                  placeholder="Например: ru, kg"
                />
              </label>
            ))
          : null}
        <button type="button" onClick={onMergeCheckedXlsx} disabled={isConvertingSpans}>
          {isConvertingSpans ? "Склеиваем..." : "Склеить в span XLSX"}
        </button>
      </div>

      <p className="status">{message}</p>
    </main>
  );
}

export default App;
