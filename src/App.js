import { useState } from "react";
import { convertDocxFileToXlsxBlob } from "./docxToXlsx";

function App() {
  const [file, setFile] = useState(null);
  const [isConverting, setIsConverting] = useState(false);
  const [message, setMessage] = useState("Выберите DOCX файл для конвертации.");

  const onFileChange = (event) => {
    const selected = event.target.files?.[0] || null;
    setFile(selected);
    setMessage(selected ? `Выбран файл: ${selected.name}` : "Файл не выбран.");
  };

  const onConvert = async () => {
    if (!file) {
      setMessage("Сначала выберите .docx файл.");
      return;
    }

    if (!file.name.toLowerCase().endsWith(".docx")) {
      setMessage("Поддерживаются только .docx файлы.");
      return;
    }

    setIsConverting(true);
    setMessage("Конвертация в браузере...");

    try {
      const blob = await convertDocxFileToXlsxBlob(file);
      const outputName = `${file.name.replace(/\.docx$/i, "")}.xlsx`;
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
        Конвертация выполняется локально в браузере (без установки LibreOffice).
      </p>

      <div className="card">
        <input type="file" accept=".docx" onChange={onFileChange} />
        <button type="button" onClick={onConvert} disabled={isConverting}>
          {isConverting ? "Конвертируем..." : "Конвертировать в XLSX"}
        </button>
      </div>

      <p className="status">{message}</p>
    </main>
  );
}

export default App;
