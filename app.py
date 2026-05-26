"""公文格式核稿系统 — FastAPI 后端入口"""

import os
import uuid
import shutil
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
import uvicorn

from parser import parse_document
from formatter import format_document, UPLOAD_DIR, OUTPUT_DIR
from config import FONT_FILES

# 确保目录存在
Path(UPLOAD_DIR).mkdir(parents=True, exist_ok=True)
Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

# 字体目录
FONTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")

app = FastAPI(title="公文格式核稿系统", version="1.0.0")

# 存储已上传文件信息（内存中）
file_store: dict = {}  # file_id → {name, path, paragraphs, issue_count}


# ── 静态文件 ──

@app.get("/")
async def index():
    """返回前端页面"""
    static_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
    return FileResponse(os.path.join(static_dir, "index.html"))


# ── API 端点 ──

@app.post("/api/upload")
async def upload_files(files: list[UploadFile] = File(...)):
    """
    上传 .docx 文件，解析文档结构，返回检测结果。
    支持批量上传。
    """
    results = []

    for file in files:
        if not file.filename.lower().endswith(".docx"):
            results.append({
                "id": None,
                "name": file.filename,
                "error": "仅支持 .docx 格式文件",
                "issue_count": 0,
            })
            continue

        # 保存原始文件（只读不写）
        file_id = str(uuid.uuid4())
        file_path = os.path.join(UPLOAD_DIR, f"{file_id}.docx")

        content = await file.read()
        with open(file_path, "wb") as f:
            f.write(content)

        # 解析文档结构
        try:
            structure = parse_document(file_path)
        except Exception as e:
            results.append({
                "id": file_id,
                "name": file.filename,
                "error": f"文档解析失败：{str(e)}",
                "issue_count": 0,
            })
            continue

        # 存储文件信息
        file_store[file_id] = {
            "name": file.filename,
            "path": file_path,
            "paragraphs": structure["paragraphs"],
            "issue_count": structure["issue_count"],
            "page_issues": structure.get("page_issues", []),
            "corrected_id": None,
        }

        results.append({
            "id": file_id,
            "name": file.filename,
            "paragraphs": structure["paragraphs"],
            "issue_count": structure["issue_count"],
            "page_issues": structure.get("page_issues", []),
        })

    return JSONResponse({"files": results})


@app.post("/api/format")
async def format_files(data: dict):
    """
    一键格式统一。接收文件 ID 列表，返回修正后文件的下载信息。
    """
    file_ids = data.get("file_ids", [])
    if not file_ids:
        raise HTTPException(status_code=400, detail="请提供要处理的文件 ID 列表")

    results = []
    for fid in file_ids:
        if fid not in file_store:
            results.append({"id": fid, "status": "error", "error": "文件不存在"})
            continue

        try:
            corrected_id = format_document(fid)
            file_store[fid]["corrected_id"] = corrected_id
            original_name = file_store[fid]["name"]
            base_name = original_name.rsplit(".", 1)[0]
            results.append({
                "id": fid,
                "status": "corrected",
                "original_name": original_name,
                "corrected_name": f"{base_name}_修正版.docx",
                "download_url": f"/api/download/{fid}",
            })
        except Exception as e:
            results.append({"id": fid, "status": "error", "error": str(e)})

    return JSONResponse({"results": results})


@app.get("/api/download/{file_id}")
async def download_file(file_id: str):
    """下载修正后的 .docx 文件"""
    if file_id not in file_store:
        raise HTTPException(status_code=404, detail="文件不存在")

    corrected_id = file_store[file_id].get("corrected_id")
    if not corrected_id:
        raise HTTPException(status_code=400, detail="该文件尚未进行格式修正")

    output_path = os.path.join(OUTPUT_DIR, f"{corrected_id}.docx")
    if not os.path.exists(output_path):
        raise HTTPException(status_code=404, detail="修正文件不存在，请重新处理")

    original_name = file_store[file_id]["name"]
    base_name = original_name.rsplit(".", 1)[0]
    download_name = f"{base_name}_修正版.docx"

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=download_name,
    )


@app.get("/api/preview/{file_id}")
async def preview_file(file_id: str):
    """获取文件的结构解析预览（原始 vs 修正）"""
    if file_id not in file_store:
        raise HTTPException(status_code=404, detail="文件不存在")

    info = file_store[file_id]
    return JSONResponse({
        "id": file_id,
        "name": info["name"],
        "paragraphs": info["paragraphs"],
        "issue_count": info["issue_count"],
        "page_issues": info.get("page_issues", []),
    })


@app.get("/api/fonts/{filename}")
async def download_font(filename: str):
    """下载字体文件"""
    # 安全检查：防止路径遍历
    safe_name = os.path.basename(filename)
    if safe_name not in FONT_FILES:
        raise HTTPException(status_code=404, detail="字体文件不存在")

    font_path = os.path.join(FONTS_DIR, safe_name)
    if not os.path.exists(font_path):
        raise HTTPException(status_code=404, detail="字体文件不存在")

    return FileResponse(
        font_path,
        media_type="font/ttf",
        filename=safe_name,
    )


@app.get("/api/fonts")
async def list_fonts():
    """列出可下载的字体文件"""
    fonts = []
    for filename, description in FONT_FILES.items():
        filepath = os.path.join(FONTS_DIR, filename)
        available = os.path.exists(filepath)
        fonts.append({
            "filename": filename,
            "description": description,
            "download_url": f"/api/fonts/{filename}",
            "available": available,
        })
    return JSONResponse({"fonts": fonts})


@app.get("/api/health")
async def health():
    """健康检查 + 字体可用性"""
    import ctypes
    from ctypes import wintypes

    # 检查系统字体
    required = ["方正小标宋_GBK", "仿宋_GB2312", "楷体_GB2312", "黑体"]
    font_status = {}
    for font_name in required:
        font_status[font_name] = _check_font_installed(font_name)

    return JSONResponse({
        "status": "ok",
        "fonts": font_status,
        "local_fonts_available": os.path.isdir(FONTS_DIR),
    })


def _check_font_installed(font_name: str) -> bool:
    """检查 Windows 系统是否安装了指定字体"""
    try:
        import ctypes
        from ctypes import wintypes

        # 使用 GDI32 检查字体
        hdc = ctypes.windll.user32.GetDC(0)
        if hdc:
            # 尝试创建字体
            hfont = ctypes.windll.gdi32.CreateFontW(
                16, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, font_name
            )
            if hfont:
                ctypes.windll.gdi32.DeleteObject(hfont)
                ctypes.windll.user32.ReleaseDC(0, hdc)
                return True
            ctypes.windll.user32.ReleaseDC(0, hdc)
        return False
    except Exception:
        return False


# ── 启动 ──

if __name__ == "__main__":
    print("=" * 50)
    print("  公文格式核稿系统 v1.0.0")
    print("  http://localhost:8000")
    print("=" * 50)
    uvicorn.run(app, host="0.0.0.0", port=8000)
