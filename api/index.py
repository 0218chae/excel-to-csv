# api/index.py
import io, re, os, zipfile
import pandas as pd
from flask import Flask, request, send_file, abort

app = Flask(__name__)

_invalid = r"[^\w\-\u3131-\u318E\uAC00-\uD7A3\s]"
_space = re.compile(r"\s+")

def safe_name(name: str) -> str:
    name = re.sub(_invalid, "_", name)
    return _space.sub(" ", name).strip() or "sheet"

@app.post("/api/convert")
def convert():
    if "file" not in request.files:
        abort(400, "업로드된 파일이 없습니다.")
    f = request.files["file"]
    if not f.filename:
        abort(400, "파일 이름이 비어 있습니다.")
    if os.path.splitext(f.filename)[1].lower() not in [".xlsx", ".xls"]:
        abort(400, "xlsx/xls만 허용합니다.")

    buf = io.BytesIO(f.read()); buf.seek(0)
    try:
        xls = pd.ExcelFile(buf)
    except Exception as e:
        abort(400, f"엑셀 파일을 여는 중 오류: {e}")

    if not xls.sheet_names:
        abort(400, "시트가 없습니다.")

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        used = {}
        for sheet in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet)
            except Exception as e:
                zf.writestr(f"ERROR_{safe_name(sheet)}.txt", str(e))
                continue
            base = safe_name(sheet)
            n = used.get(base, 0); used[base] = n + 1
            name = f"{base}.csv" if n == 0 else f"{base}_{n+1}.csv"
            zf.writestr(name, df.to_csv(index=False).encode("utf-8-sig"))
    zip_buf.seek(0)
    base = safe_name(os.path.splitext(f.filename)[0]) or "excel"
    return send_file(zip_buf, as_attachment=True,
                     download_name=f"{base}_sheets.zip",
                     mimetype="application/zip")