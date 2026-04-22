import { useState } from "react";
import { convertDocxFilesToXlsxBlob } from "./docxToXlsx";

function App() {
  const [files, setFiles] = useState([]);
  const [classCodes, setClassCodes] = useState([]);
  const [outputMode, setOutputMode] = useState("spans");
  const [isConverting, setIsConverting] = useState(false);
  const [message, setMessage] = useState("Выберите один или несколько DOCX файлов.");

  const inferClassCode = (fileName, index) => {
    const base = fileName.replace(/\.docx$/i, "").toLowerCase();
    const suffix = base.match(/[_-]([a-zа-я]{2,5})$/i);
    if (suffix) {
      return suffix[1];
    }
    return `lang_${index + 1}`;
  };

  const onFileChange = (event) => {
    const selectedFiles = Array.from(event.target.files || []);
    setFiles(selectedFiles);
    setClassCodes(selectedFiles.map((file, index) => inferClassCode(file.name, index)));
    setMessage(
      selectedFiles.length > 0
        ? `Выбрано файлов: ${selectedFiles.length}`
        : "Файлы не выбраны."
    );
  };

  const onConvert = async () => {
    if (files.length === 0) {
      setMessage("Сначала выберите хотя бы один .docx файл.");
      return;
    }

    if (files.some((file) => !file.name.toLowerCase().endsWith(".docx"))) {
      setMessage("Поддерживаются только .docx файлы.");
      return;
    }

    if (
      outputMode === "spans" &&
      (classCodes.length !== files.length || classCodes.some((code) => !code.trim()))
    ) {
      setMessage("Укажи CSS-класс для каждого загруженного файла.");
      return;
    }

    setIsConverting(true);
    setMessage("Конвертация в браузере...");

    try {
      const blob = await convertDocxFilesToXlsxBlob(
        files,
        outputMode === "spans"
          ? { mode: "spans", classNames: classCodes.map((code) => code.trim()) }
          : { mode: "columns" }
      );
      const outputName =
        files.length === 1
          ? `${files[0].name.replace(/\.docx$/i, "")}.xlsx`
          : "combined.xlsx";
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = outputName;
      link.click();
      URL.revokeObjectURL(url);
      setMessage(`Готово: ${outputName}`);
    } catch (error) {
      setMessage(`Ошибка конвертации: ${error.message}`);
    } finally {
      setIsConverting(false);
    }
  };

  return (
    <main className="container">
      <h1>DOCX to XLSX Converter</h1>
      <p className="description">
        Можно загрузить несколько идентичных DOCX на разных языках и выбрать режим:
        по отдельным колонкам или схлопнуть в один HTML `span`-блок.
      </p>

      <div className="card">
        <input type="file" accept=".docx" multiple onChange={onFileChange} />
        <div className="mode-row">
          <label>
            <input
              type="radio"
              name="output-mode"
              value="columns"
              checked={outputMode === "columns"}
              onChange={() => setOutputMode("columns")}
            />
            По колонкам (старый формат)
          </label>
          <label>
            <input
              type="radio"
              name="output-mode"
              value="spans"
              checked={outputMode === "spans"}
              onChange={() => setOutputMode("spans")}
            />
            В одну ячейку со span
          </label>
        </div>
        {outputMode === "spans"
          ? files.map((file, index) => (
              <label key={file.name + index} className="class-code-row">
                <span>{file.name}</span>
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
        <button type="button" onClick={onConvert} disabled={isConverting}>
          {isConverting ? "Конвертируем..." : "Конвертировать в XLSX"}
        </button>
      </div>

      <p className="status">{message}</p>
    </main>
  );
}

export default App;
