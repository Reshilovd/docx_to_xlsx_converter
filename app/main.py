from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from app.converter import convert_docx_to_xlsx

app = FastAPI(
    title="DOCX to XLSX Converter",
    description="Web service for converting DOCX files to XLSX.",
    version="1.0.0",
)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/convert")
async def convert(file: UploadFile = File(...)) -> StreamingResponse:
    if not file.filename:
        raise HTTPException(status_code=400, detail="Filename is required.")

    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are supported.")

    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    try:
        xlsx_bytes = convert_docx_to_xlsx(content)
    except Exception as error:  # pragma: no cover
        raise HTTPException(status_code=500, detail=f"Conversion failed: {error}") from error

    output_name = file.filename.rsplit(".", maxsplit=1)[0] + ".xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{output_name}"'}
    return StreamingResponse(
        content=iter([xlsx_bytes]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )
