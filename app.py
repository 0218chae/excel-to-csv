import io
import os
import re
import zipfile
from flask import Flask, render_template, request, send_file, abort
from werkzeug.utils import secure_filename
import pandas as pd

app = Flask(__name__, template_folder=".")

# 업로드 제한: 200KB (로컬 테스트용)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

# 시트명/파일명 안전화 (한글/영문/숫자/밑줄/하이픈/공백 허용)
_invalid_chars = r"[^\w\-\u3131-\u318E\uAC00-\uD7A3\s]"
_space_re = re.compile(r"\s+")


def safe_name(name: str) -> str:
    name = re.sub(_invalid_chars, "_", name)
    name = _space_re.sub(" ", name).strip()
    return name if name else "sheet"


def allowed_file(filename: str) -> bool:
    _, ext = os.path.splitext(filename)
    return ext.lower() in ALLOWED_EXTENSIONS


@app.get("/")
def index():
    # Vercel 구조에 맞춰 index.html이 리포 루트에 있으므로,
    # 로컬 서버에서도 동일 파일을 직접 서빙
    return send_file("index.html")


@app.post("/convert")
def convert():
    if "file" not in request.files:
        abort(400, "업로드된 파일이 없습니다.")

    file = request.files["file"]
    if file.filename == "":
        abort(400, "파일 이름이 비어 있습니다.")

    if not allowed_file(file.filename):
        abort(400, "지원하지 않는 확장자입니다. (.xlsx, .xls만 허용)")

    filename = secure_filename(file.filename)
    base_name, _ = os.path.splitext(filename)
    base_name = safe_name(base_name) or "excel"

    # 파일을 메모리로 읽어 처리 (디스크에 저장하지 않음)
    file_bytes = io.BytesIO(file.read())

    # 서버 측 크기 제한 (200KB) - Flask MAX_CONTENT_LENGTH와 이중 보호
    if file_bytes.seek(0, os.SEEK_END) > 200 * 1024:
        abort(413, "파일이 너무 큽니다. 최대 200KB")
    file_bytes.seek(0)

    if file_bytes.getbuffer().nbytes > 200 * 1024:
        abort(413, "파일이 너무 큽니다. 최대 200KB")

    try:
        xls = pd.ExcelFile(file_bytes)
    except Exception as e:
        abort(400, f"엑셀 파일을 여는 중 오류가 발생했습니다: {e}")

    if not xls.sheet_names:
        abort(400, "시트를 찾지 못했습니다. 엑셀 파일에 시트가 있는지 확인하세요.")

    # ZIP을 메모리에서 생성
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        used = {}
        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name)
            except Exception as e:
                # 읽기 실패 시 빈 텍스트로 대신하지 않고, 에러 로그용 파일을 ZIP에 남김
                err_txt = f'시트 "{sheet_name}" 읽기 실패: {e}'
                zf.writestr(f"ERROR_{safe_name(sheet_name)}.txt", err_txt)
                continue

            safe = safe_name(sheet_name)
            count = used.get(safe, 0)
            used[safe] = count + 1
            out_name = f"{safe}.csv" if count == 0 else f"{safe}_{count+1}.csv"

            # CSV를 UTF-8 BOM(utf-8-sig)로 인코딩 → 엑셀에서 한글 깨짐 방지
            csv_text = df.to_csv(index=False)
            zf.writestr(out_name, csv_text.encode("utf-8-sig"))

    zip_buf.seek(0)
    zip_filename = f"{base_name}_sheets.zip"

    return send_file(
        zip_buf,
        as_attachment=True,
        download_name=zip_filename,
        mimetype="application/zip",
        max_age=0,
    )


if __name__ == "__main__":
    # 로컬 개발용 실행
    app.run(host="127.0.0.1", port=5000, debug=True)