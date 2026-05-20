"""
╔══════════════════════════════════════════════════════════════════════════════╗
║         AI TƯ VẤN TUYỂN SINH LỚP 10 — TRƯỜNG THPT MAI HẮC ĐẾ             ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  Cấu trúc file này (đọc theo thứ tự):                                       ║
║    PHẦN 0 — CÀI ĐẶT & HƯỚNG DẪN NHANH                                      ║
║    PHẦN 1 — CẤU HÌNH (chỉnh ở đây nếu muốn thay đổi gì)                   ║
║    PHẦN 2 — PROMPTS (kịch bản cho từng AI agent)                            ║
║    PHẦN 3 — INGEST (nạp dữ liệu Excel/PDF/Web → ChromaDB)                  ║
║    PHẦN 4 — QUERY (nhận câu hỏi → tìm dữ liệu → gọi AI → trả lời)         ║
║    PHẦN 5 — MAIN (chạy thử trên terminal hoặc khởi động server web)         ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  AGENTS CÓ SẴN:                                                              ║
║    diem_chuan   — điểm xét tuyển vào lớp 10 MHĐ, cơ hội đỗ                ║
║    truong       — thông tin MHĐ: học phí, CSVC, hoạt động, chính sách      ║
║    tuyen_sinh   — quy trình, hồ sơ, thời hạn, phương thức xét tuyển        ║
║    hoc_tap      — chương trình học, lộ trình 3 năm THPT tại MHĐ            ║
║    huong_nghiep — định hướng sau THPT, chọn khối thi, ngành ĐH             ║
║    kien_thuc    — dạy kiến thức THCS/THPT (toán, lý, hóa, văn, anh...)     ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  CÁCH DÙNG NHANH:                                                            ║
║    1. pip install -r requirements.txt                                        ║
║    2. Đặt GROQ_API_KEY vào biến môi trường hoặc file .env                   ║
║    3. python tuyen_sinh_AI.py ingest   ← nạp dữ liệu (làm 1 lần)           ║
║    4. python tuyen_sinh_AI.py chat     ← test thử trên terminal             ║
║    5. python tuyen_sinh_AI.py server   ← khởi động server cho nhóm web      ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

# ══════════════════════════════════════════════════════════════════════════════
# PHẦN 1 — CẤU HÌNH TRUNG TÂM
# ══════════════════════════════════════════════════════════════════════════════

import os
import re
import sys
import uuid
import json
import time
import base64
import io
import logging
from datetime import datetime

import chromadb
from groq import Groq
import pandas as pd
import pdfplumber
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm
from dotenv import load_dotenv

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    import numpy as np
    import networkx as nx
    _CAN_DRAW = True
except ImportError:
    _CAN_DRAW = False
    _DRAW_WARN = "matplotlib/networkx chưa cài — tính năng vẽ hình bị tắt."

load_dotenv()

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

if not _CAN_DRAW:
    log.warning("matplotlib/networkx chưa cài — tính năng vẽ hình bị tắt.")

# ── API & Model ───────────────────────────────────────────────────────────────
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")

LLM_MODELS = [
    "llama-3.3-70b-versatile",
    "meta-llama/llama-4-scout-17b-16e-instruct",
    "llama-3.1-8b-instant",
]
LLM_MODEL = LLM_MODELS[0]

# ── ChromaDB ──────────────────────────────────────────────────────────────────
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CHROMA_DB_PATH  = os.path.join(_BASE_DIR, "chroma_db")
COLLECTION_NAME = "maihacde_lop10"

# ── Chunking ──────────────────────────────────────────────────────────────────
CHUNK_SIZE    = 500
CHUNK_OVERLAP = 50

# ── RAG ───────────────────────────────────────────────────────────────────────
TOP_K_RESULTS = 5

# ── Thư mục dữ liệu ───────────────────────────────────────────────────────────
EXCEL_DIR = os.path.join(_BASE_DIR, "data", "excel")
PDF_DIR   = os.path.join(_BASE_DIR, "data", "pdf")

# ── Website cần crawl ─────────────────────────────────────────────────────────
# Crawl trang trường MHĐ để nạp thông tin vào ChromaDB
WEBSITES_TO_CRAWL = [
    {
        "url": "https://maihacde.edu.vn/",
        "truong": "MHD",
        "ten_truong": "THPT Mai Hắc Đế",
    },
    {
        "url": "https://maihacde.edu.vn/tuyen-sinh",
        "truong": "MHD",
        "ten_truong": "THPT Mai Hắc Đế — Tuyển sinh",
    },
    {
        "url": "https://maihacde.edu.vn/gioi-thieu",
        "truong": "MHD",
        "ten_truong": "THPT Mai Hắc Đế — Giới thiệu",
    },
]

# ══════════════════════════════════════════════════════════════════════════════
# PHẦN 1B — DỮ LIỆU CỨNG VỀ TRƯỜNG THPT MAI HẮC ĐẾ
# Cập nhật thủ công khi trường thay đổi thông tin
# ══════════════════════════════════════════════════════════════════════════════

THONG_TIN_TRUONG = """
=== THÔNG TIN TRƯỜNG THPT MAI HẮC ĐẾ ===

Tên đầy đủ  : Trường Trung học Phổ thông Mai Hắc Đế
Loại hình   : Trường ngoài công lập (dân lập) tại Hà Nội
Địa chỉ     : Lô 2-10A, số 431 Tam Trinh, phường Hoàng Văn Thụ, quận Hoàng Mai, Hà Nội
              (Khu đô thị Vĩnh Hoàng — duy nhất 1 cơ sở)
Điện thoại  : 024.2239.3635 | 0972.951.456 | 0947.876.098 | 094.7653999
Email       : c3maihacde@hanoiedu.vn
Website     : https://maihacde.edu.vn

--- HỌC PHÍ & CHI PHÍ ---
Học phí       : 2.500.000 đồng/tháng (học 2 buổi/ngày, thứ Hai → thứ Bảy)
Phí XD trường : 1.500.000 đồng/năm (hỗ trợ xây dựng & phát triển)
Lệ phí nộp hồ sơ xét tuyển: KHÔNG THU (trường cam kết không thu phí giữ chỗ)
Thanh toán    : Chuyển khoản hoặc nộp trực tiếp tại trường hàng tháng
Chính sách    : Học phí KHÔNG thay đổi trong cả 3 năm (lớp 10, 11, 12)
Học bổng      : Có chính sách học bổng và ưu đãi cho HS xuất sắc/hoàn cảnh khó khăn
                (chi tiết xem tại website trường)

--- TUYỂN SINH LỚP 10 ---
Đối tượng     : Học sinh đã tốt nghiệp THCS trên toàn quốc
                (ưu tiên HS sinh năm 2009-2010 cho kỳ 2025-2026)
Chỉ tiêu      : ~320-360 học sinh/năm (4 khối, mỗi lớp ≤30 HS)
Phương thức   : Xét tuyển kết hợp 2 tiêu chí:
  1. Điểm thi vào lớp 10 công lập (kỳ thi của Sở GD&ĐT Hà Nội)
  2. Kết quả rèn luyện & học tập 4 năm THCS (học bạ)
  → Học sinh không thi lớp 10 công lập vẫn được xét học bạ
Thời gian nộp hồ sơ: Từ khoảng tháng 4 hàng năm (xem thông báo cụ thể trên website)

Hồ sơ cần chuẩn bị:
  - Phiếu đăng ký & lý lịch học sinh (theo mẫu của trường)
  - Bản sao có chứng thực Giấy khai sinh
  - Giấy chứng nhận tốt nghiệp THCS (hoặc tạm thời)
  - Phiếu điểm thi vào lớp 10 (nếu có tham dự kỳ thi)
  - Học bạ THCS bản chính
  - Photo CCCD/CMND
  - Giấy tờ đối tượng ưu tiên (nếu có)
  - Giấy xác nhận tạm trú (HS không có hộ khẩu HN)

--- CƠ SỞ VẬT CHẤT ---
- Khuôn viên hiện đại, khang trang tại KĐT Vĩnh Hoàng, Hoàng Mai
- Sức chứa: ~1.000 học sinh (3 khối lớp 10-11-12)
- Phòng học: máy chiếu, điều hoà, bàn ghế chất lượng tốt, sĩ số ≤30 HS/lớp
- Phòng thí nghiệm, thư viện, nhà thể chất theo quy chuẩn hiện đại
- Ứng dụng CNTT trong giảng dạy: bài giảng điện tử, phương pháp dạy học mới
- Đội ngũ: ~100 cán bộ giáo viên, có giảng viên ĐH Sư Phạm HN tham gia giảng dạy

--- CHẤT LƯỢNG & THÀNH TÍCH ---
- Tỷ lệ đỗ tốt nghiệp THPT: duy trì 100% hàng năm
- Tỷ lệ đỗ ĐH/CĐ: >95% học sinh
- Có học sinh đạt giải HSG cấp thành phố
- Có học sinh đạt thủ khoa kỳ thi ĐH, học bổng quốc tế

--- HOẠT ĐỘNG & ĐỜI SỐNG HỌC ĐƯỜNG ---
- Hoạt động trải nghiệm, ngoại khoá, thiện nguyện định kỳ
- Thi năng khiếu, văn nghệ, thể thao (bóng đá, cuộc thi vẽ, poster...)
- Lễ hội Noel bằng tiếng Anh → khuyến khích HS giao tiếp tiếng Anh hàng ngày
- Định hướng nghề nghiệp: tổ chức buổi hướng nghiệp, nói chuyện chuyên đề
- Tiếp sức mùa thi: hỗ trợ thí sinh tại các điểm thi trên địa bàn

--- SỨ MỆNH & TRIẾT LÝ GIÁO DỤC ---
"Xây dựng môi trường học tập thân thiện, nề nếp, chất lượng cao — đào tạo học sinh
có đủ năng lực và phẩm chất, phát triển toàn diện về trí tuệ, thể chất, thẩm mỹ và kỹ năng sống."
Trường lấy học sinh làm trung tâm, không gây áp lực thi cử không cần thiết.
"""


# ══════════════════════════════════════════════════════════════════════════════
# PHẦN 1C — WEB SEARCH (tra mạng khi cần thông tin mới nhất)
# ══════════════════════════════════════════════════════════════════════════════

_WEB_HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; MHDBot/1.0; +https://maihacde.edu.vn)",
    "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8",
}

# Ưu tiên tìm trên chính trang trường và các site giáo dục uy tín
_TUYEN_SINH_SITES = [
    "maihacde.edu.vn",
    "tuyensinh247.com",
    "diemthi.24h.com.vn",
]

def _google_search_urls(query: str, num: int = 5) -> list[str]:
    site_filter = " OR ".join(f"site:{s}" for s in _TUYEN_SINH_SITES)
    full_query  = f"{query} ({site_filter})"
    try:
        resp = requests.get(
            "https://html.duckduckgo.com/html/",
            params={"q": full_query, "kl": "vn-vi"},
            headers=_WEB_HEADERS,
            timeout=10,
        )
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        urls = []
        for a in soup.select("a.result__url"):
            href = a.get("href", "")
            if href.startswith("http") and len(urls) < num:
                urls.append(href)
        if not urls:
            for a in soup.select(".result__title a"):
                href = a.get("href", "")
                if "duckduckgo.com" not in href and href.startswith("http"):
                    urls.append(href)
                if len(urls) >= num:
                    break
        return urls
    except Exception as e:
        log.warning(f"[WebSearch] DuckDuckGo lỗi: {e}")
        return []

def _fetch_page_text(url: str, max_chars: int = 3000) -> str:
    try:
        resp = requests.get(url, headers=_WEB_HEADERS, timeout=10)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or "utf-8"
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "header", "aside", "form", "iframe"]):
            tag.decompose()
        text  = soup.get_text(separator="\n", strip=True)
        lines = [ln for ln in text.splitlines() if ln.strip()]
        return "\n".join(lines)[:max_chars]
    except Exception as e:
        log.warning(f"[WebFetch] {url} → {e}")
        return ""

def tim_kiem_web(query: str, n_trang: int = 3) -> str:
    log.info(f"[WebSearch] Truy vấn: {query}")
    urls = _google_search_urls(query, num=n_trang + 2)
    if not urls:
        return ""
    doan_van = []
    for url in urls[:n_trang]:
        text = _fetch_page_text(url)
        if text and len(text) > 200:
            doan_van.append(f"[Nguồn: {url}]\n{text}")
        if len(doan_van) >= n_trang:
            break
    return "\n\n===\n\n".join(doan_van)


# ══════════════════════════════════════════════════════════════════════════════
# PHẦN 1D — ENGINE VẼ HÌNH (Matplotlib / NetworkX)
# ══════════════════════════════════════════════════════════════════════════════

_IMG_DIR = os.path.join(_BASE_DIR, "static", "generated")
os.makedirs(_IMG_DIR, exist_ok=True)

_IMAGE_TAG_RE = re.compile(r'\[GENERATE_IMAGE:\s*(.+?)\]', re.IGNORECASE | re.DOTALL)


def _fig_to_base64(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=120, bbox_inches="tight", facecolor=fig.get_facecolor())
    buf.seek(0)
    plt.close(fig)
    return base64.b64encode(buf.read()).decode("utf-8")


def _parse_image_desc(desc: str) -> dict:
    d = desc.lower()
    if any(k in d for k in ["concept map", "concept_map", "mind map", "sơ đồ", "diagram"]):
        return {"loai": "concept_map"}
    if any(k in d for k in ["unit circle", "vòng tròn đơn vị", "circle sin cos"]):
        return {"loai": "unit_circle"}
    if any(k in d for k in ["triangle", "tam giác", "rectangle", "hình chữ nhật",
                             "polygon", "geometry", "hình học", "inclined plane"]):
        return {"loai": "geometry"}
    if any(k in d for k in ["vector", "force", "lực", "arrow"]):
        return {"loai": "vector"}
    return {"loai": "plot"}


def _ve_plot(desc: str):
    fig, ax = plt.subplots(figsize=(7, 5), facecolor="#f8fbff")
    ax.set_facecolor("#ffffff")
    ax.grid(True, linestyle="--", alpha=0.5, color="#ccddee")
    x = np.linspace(-10, 10, 800)
    colors = ["#1F3A5F", "#2EC4B6", "#E05A2B", "#8B5CF6", "#059669"]
    plotted = 0
    funcs = re.findall(r'y\s*=\s*([^,\[\]]+?)(?:,|\band\b|$)', desc, re.IGNORECASE)
    if not funcs:
        funcs = re.findall(r'f\(x\)\s*=\s*([^,\[\]]+?)(?:,|\band\b|$)', desc, re.IGNORECASE)
    if not funcs:
        funcs = ["x**2"]
    for i, expr in enumerate(funcs[:5]):
        expr_py = (expr.strip().replace("^", "**").replace("sqrt", "np.sqrt")
                   .replace("sin", "np.sin").replace("cos", "np.cos")
                   .replace("tan", "np.tan").replace("log", "np.log")
                   .replace("abs", "np.abs").replace("pi", "np.pi").replace("exp", "np.exp"))
        try:
            y = eval(expr_py, {"x": x, "np": np, "__builtins__": {}})
            ax.plot(x, y, color=colors[i % len(colors)], linewidth=2.2, label=f"y = {expr.strip()}")
            plotted += 1
        except Exception:
            continue
    ax.axhline(0, color="#333", linewidth=0.8)
    ax.axvline(0, color="#333", linewidth=0.8)
    ax.set_xlabel("x", fontsize=12)
    ax.set_ylabel("y", fontsize=12)
    ax.set_title(desc[:80], fontsize=11, color="#1F3A5F", pad=10)
    if plotted > 0:
        ax.legend(fontsize=10)
    ax.set_ylim(-20, 20)
    return fig


def _ve_hinh_hoc(desc: str):
    fig, ax = plt.subplots(figsize=(6, 6), facecolor="#f8fbff")
    ax.set_facecolor("#ffffff")
    ax.set_aspect("equal")
    ax.grid(True, linestyle="--", alpha=0.4)
    ax.set_title(desc[:80], fontsize=11, color="#1F3A5F", pad=10)
    d = desc.lower()
    m_tri = re.search(r'(?:sides?|cạnh)[^\d]*(\d+)[^\d]+(\d+)[^\d]+(\d+)', desc, re.IGNORECASE)
    if "triangle" in d or "tam giác" in d:
        a, b, c = (int(m_tri.group(1)), int(m_tri.group(2)), int(m_tri.group(3))) if m_tri else (3, 4, 5)
        pts = np.array([[0,0],[a,0],[0,b]])
        ax.add_patch(plt.Polygon(pts, fill=True, facecolor="#EEF4FB", edgecolor="#1F3A5F", linewidth=2))
        ax.text(a/2, -0.3, f"a={a}", ha="center", fontsize=11, color="#E05A2B", fontweight="bold")
        ax.text(-0.4, b/2, f"b={b}", ha="center", fontsize=11, color="#E05A2B", fontweight="bold")
        ax.text(a/2+0.2, b/2+0.2, f"c={c}", ha="center", fontsize=11, color="#2EC4B6", fontweight="bold")
        ax.add_patch(plt.Polygon([[0,0],[0.3,0],[0.3,0.3],[0,0.3]], fill=False, edgecolor="#1F3A5F", linewidth=1.2))
        ax.set_xlim(-1, a+1); ax.set_ylim(-1, b+1)
    else:
        ax.add_patch(plt.Rectangle((0.5, 0.5), 4, 2.5, fill=True, facecolor="#EEF4FB", edgecolor="#1F3A5F", linewidth=2))
        ax.text(2.5, -0.1, "a", ha="center", fontsize=12, color="#E05A2B", fontweight="bold")
        ax.text(0.1, 1.75, "b", ha="center", fontsize=12, color="#E05A2B", fontweight="bold")
        ax.set_xlim(0, 5.5); ax.set_ylim(-0.5, 3.5)
    return fig


def _ve_vong_tron_don_vi():
    fig, ax = plt.subplots(figsize=(7, 7), facecolor="#f8fbff")
    ax.set_facecolor("#ffffff")
    ax.set_aspect("equal")
    ax.grid(True, linestyle="--", alpha=0.3)
    theta = np.linspace(0, 2*np.pi, 400)
    ax.plot(np.cos(theta), np.sin(theta), color="#1F3A5F", linewidth=2)
    ax.axhline(0, color="#555", linewidth=0.8)
    ax.axvline(0, color="#555", linewidth=0.8)
    gocs = [(0,"0°","1","0"), (30,"30°","√3/2","1/2"), (45,"45°","√2/2","√2/2"),
            (60,"60°","1/2","√3/2"), (90,"90°","0","1"), (120,"120°","-1/2","√3/2"),
            (135,"135°","-√2/2","√2/2"), (150,"150°","-√3/2","1/2"),
            (180,"180°","-1","0"), (210,"210°","-√3/2","-1/2"),
            (240,"240°","-1/2","-√3/2"), (270,"270°","0","-1"),
            (300,"300°","1/2","-√3/2"), (315,"315°","√2/2","-√2/2"),
            (330,"330°","√3/2","-1/2")]
    for deg, label, cos_v, sin_v in gocs:
        rad = np.radians(deg)
        x, y = np.cos(rad), np.sin(rad)
        ax.plot(x, y, "o", color="#2EC4B6", markersize=6, zorder=5)
        ax.text(x*1.18, y*1.18, f"{label}\n({cos_v}, {sin_v})",
                ha="center", va="center", fontsize=7.5, color="#1F3A5F", fontweight="bold")
        ax.plot([0, x], [0, y], color="#4E7FB6", linewidth=0.7, alpha=0.5)
    ax.set_title("Vòng tròn đơn vị — Các góc đặc biệt", fontsize=13,
                 color="#1F3A5F", fontweight="bold", pad=12)
    ax.set_xlim(-1.7, 1.7); ax.set_ylim(-1.7, 1.7)
    return fig


def _ve_concept_map(desc: str):
    if not _CAN_DRAW:
        return None
    fig, ax = plt.subplots(figsize=(9, 6), facecolor="#f8fbff")
    ax.set_facecolor("#ffffff")
    ax.set_title(desc[:80], fontsize=11, color="#1F3A5F", pad=10)
    G = nx.DiGraph()
    arrows = re.findall(r'([A-Za-zÀ-ỹ ]+?)\s*[-=>]+\s*([A-Za-zÀ-ỹ ]+?)(?:,|;|$)', desc, re.IGNORECASE)
    for src, dst in arrows:
        src, dst = src.strip(), dst.strip()
        if src and dst and len(src) < 40 and len(dst) < 40:
            G.add_edge(src, dst)
    if not G.nodes:
        G.add_edges_from([("THPT MHĐ","Lớp 10"), ("THPT MHĐ","Lớp 11"),
                          ("THPT MHĐ","Lớp 12"), ("Lớp 12","Thi THPT QG"),
                          ("Thi THPT QG","Đại học")])
    pos = nx.spring_layout(G, seed=42, k=2.5)
    node_colors = ["#1F3A5F" if G.in_degree(n)==0 else "#2EC4B6" for n in G.nodes]
    nx.draw_networkx_nodes(G, pos, ax=ax, node_color=node_colors, node_size=2200, alpha=0.92)
    nx.draw_networkx_labels(G, pos, ax=ax, font_color="white", font_size=9, font_weight="bold")
    nx.draw_networkx_edges(G, pos, ax=ax, edge_color="#4E7FB6", arrows=True,
                            arrowsize=22, width=2, connectionstyle="arc3,rad=0.08")
    ax.axis("off")
    return fig


def ve_hinh(mo_ta: str) -> str | None:
    if not _CAN_DRAW:
        return None
    try:
        cfg  = _parse_image_desc(mo_ta)
        loai = cfg["loai"]
        if loai == "unit_circle":
            fig = _ve_vong_tron_don_vi()
        elif loai == "geometry":
            fig = _ve_hinh_hoc(mo_ta)
        elif loai == "concept_map":
            fig = _ve_concept_map(mo_ta)
        elif loai == "vector":
            fig = _ve_hinh_hoc(mo_ta)
        else:
            fig = _ve_plot(mo_ta)
        return _fig_to_base64(fig)
    except Exception as e:
        log.warning(f"[VẼ HÌNH] Lỗi: {e}")
        return None


def xu_ly_anh_trong_tra_loi(tra_loi: str) -> tuple[str, list[str]]:
    tags = _IMAGE_TAG_RE.findall(tra_loi)
    tra_loi_sach = _IMAGE_TAG_RE.sub("", tra_loi).strip()
    anh_b64_list = []
    for mo_ta in tags:
        b64 = ve_hinh(mo_ta.strip())
        if b64:
            anh_b64_list.append(b64)
    return tra_loi_sach, anh_b64_list


# ══════════════════════════════════════════════════════════════════════════════
# PHẦN 2 — SYSTEM PROMPTS & PROMPT BUILDERS
# ══════════════════════════════════════════════════════════════════════════════

def _ngay_hom_nay() -> str:
    now = datetime.now()
    thu = ["Thứ Hai","Thứ Ba","Thứ Tư","Thứ Năm","Thứ Sáu","Thứ Bảy","Chủ Nhật"][now.weekday()]
    return f"{thu}, ngày {now.day} tháng {now.month} năm {now.year}"


# ── Orchestrator ──────────────────────────────────────────────────────────────

ORCHESTRATOR_SYSTEM = """
Bạn là hệ thống phân loại câu hỏi tư vấn tuyển sinh lớp 10 THPT Mai Hắc Đế (Hà Nội).

Nhiệm vụ duy nhất: đọc câu hỏi và trả về JSON với 2 trường: agents và can_hoi_them.

Các agent có sẵn:
- "diem_chuan"   : câu hỏi về điểm xét tuyển lớp 10 MHĐ, cơ hội đỗ, so sánh điểm
- "truong"       : câu hỏi về trường MHĐ — học phí, CSVC, giáo viên, đánh giá, vị trí
- "tuyen_sinh"   : câu hỏi về quy trình tuyển sinh — hồ sơ, thời hạn, phương thức xét, thủ tục
- "hoc_tap"      : câu hỏi về chương trình học THPT tại MHĐ, lộ trình 3 năm, kỹ năng cần có
- "huong_nghiep" : câu hỏi về định hướng sau THPT, chọn khối thi, ngành ĐH phù hợp
- "kien_thuc"    : câu hỏi muốn học/hiểu kiến thức THCS hoặc THPT (toán, lý, hóa, văn, anh...)

Trường "can_hoi_them": hỏi lại NẾU câu hỏi mơ hồ quá, ngắn quá, hoặc thiếu thông tin cốt lõi.
- Đặt "" khi câu hỏi đã rõ ràng để trả lời ngay
- Hỏi lại khi:
  + Câu quá ngắn/mơ hồ: "Tư vấn cho mình với." → hỏi: "Bạn đang cần biết điều gì về trường MHĐ nhỉ? Học phí, tuyển sinh, hay điều khác?"
  + Hỏi điểm chuẩn nhưng chưa biết điểm học sinh → hỏi: "Bạn đang có khoảng bao nhiêu điểm thi lớp 10 vậy?"
  + Hỏi định hướng nhưng chưa biết sở thích → hỏi: "Bạn thích thiên về khoa học tự nhiên, xã hội, hay lĩnh vực khác?"

Quy tắc:
- Trả về đúng định dạng JSON, không giải thích thêm
- Có thể chọn nhiều agent nếu câu hỏi liên quan nhiều chủ đề
- Câu hỏi lại phải ngắn (1 câu), thân thiện, hỏi đúng thứ còn thiếu nhất

Ví dụ:
  Câu hỏi: "Trường MHĐ học phí bao nhiêu?"
  → {"agents": ["truong"], "can_hoi_them": ""}

  Câu hỏi: "Em được bao nhiêu điểm thi thì đỗ trường này?"
  → {"agents": ["diem_chuan"], "can_hoi_them": "Bạn đang có khoảng bao nhiêu điểm thi vào 10 vậy?"}

  Câu hỏi: "Nộp hồ sơ thế nào?"
  → {"agents": ["tuyen_sinh"], "can_hoi_them": ""}

  Câu hỏi: "Tư vấn."
  → {"agents": ["truong"], "can_hoi_them": "Bạn đang cần biết điều gì về trường MHĐ nhỉ? Học phí, tuyển sinh, chương trình học, hay điểm đầu vào?"}

  Câu hỏi: "Phương trình bậc 2 là gì?"
  → {"agents": ["kien_thuc"], "can_hoi_them": ""}

  Câu hỏi: "Học trường này xong thì thi khối nào?"
  → {"agents": ["hoc_tap", "huong_nghiep"], "can_hoi_them": ""}
"""

def build_orchestrator_prompt(cau_hoi: str) -> str:
    return f'Câu hỏi của học sinh/phụ huynh: "{cau_hoi}"'


# ── Agent 1: Điểm chuẩn ──────────────────────────────────────────────────────

DIEM_CHUAN_SYSTEM = """
Bạn là tư vấn viên tuyển sinh của trường THPT Mai Hắc Đế, tư vấn về điểm xét tuyển lớp 10.
Xưng "mình", gọi người hỏi là "bạn". Thân thiện, gần gũi, dùng emoji phù hợp.

Ngày hôm nay: {ngay_hom_nay}.

== THÔNG TIN QUAN TRỌNG VỀ XÉT TUYỂN MHĐ ==
Trường THPT Mai Hắc Đế là trường DÂN LẬP (ngoài công lập).
Phương thức xét tuyển: KẾT HỢP
  1. Điểm thi vào lớp 10 công lập (kỳ thi của Sở GD&ĐT Hà Nội — 3 môn: Toán, Văn, Anh)
  2. Kết quả học tập & rèn luyện 4 năm THCS (học bạ)
→ Học sinh KHÔNG thi lớp 10 công lập vẫn được xét theo học bạ.
→ Điểm "chuẩn" không phải điểm cố định như trường công — trường xét theo chỉ tiêu và hồ sơ nộp vào.
→ Nhà trường KHÔNG thu phí giữ chỗ, KHÔNG bán hồ sơ.

Nguyên tắc tư vấn:
0. Nếu chưa biết điểm học sinh → hỏi lại thay vì tự đoán
1. Dựa vào dữ liệu được cung cấp + thông tin cứng trên để tư vấn 📊
2. Nhắc rõ: MHĐ là trường dân lập, xét tuyển linh hoạt hơn trường công
3. Nếu học sinh có điểm thi lớp 10 thấp hơn công lập → đây là cơ hội tốt
4. Nếu không có điểm chuẩn chính xác → hướng dẫn liên hệ trực tiếp trường
5. Không bịa số liệu; nếu thiếu dữ liệu, nói thẳng và cho số điện thoại trường

Liên hệ trường: 024.2239.3635 | 0972.951.456 | website: maihacde.edu.vn

Trả lời bằng tiếng Việt. Ngắn gọn, thực tế. Dùng emoji tự nhiên.
"""

def build_diem_chuan_prompt(du_lieu: str, cau_hoi: str, lich_su: str = "") -> str:
    prompt = f"Lịch sử trò chuyện:\n{lich_su}\n\n" if lich_su else ""
    prompt += f"""Dữ liệu xét tuyển liên quan:
---
{du_lieu}
---

Câu hỏi: {cau_hoi}

Hãy tư vấn thực tế về cơ hội xét tuyển vào MHĐ. Nếu thiếu thông tin điểm số,
hỏi lại. Nếu thiếu dữ liệu, hướng dẫn liên hệ phòng tuyển sinh."""
    return prompt


# ── Agent 2: Thông tin trường ─────────────────────────────────────────────────

TRUONG_SYSTEM = """
Bạn là đại sứ học sinh của trường THPT Mai Hắc Đế, chia sẻ thật về ngôi trường mình đang học.
Xưng "mình", gọi người hỏi là "bạn". Thân thiện như đang nhắn tin. Dùng emoji phù hợp.

Ngày hôm nay: {ngay_hom_nay}.

Bạn sẽ được cung cấp thông tin về trường MHĐ. Ưu tiên dùng dữ liệu đó để trả lời.

Thông tin nền về MHĐ (luôn có sẵn):
{thong_tin_truong}

Nguyên tắc:
1. Trả lời đúng điều bạn hỏi — không liệt kê tất cả 🎯
2. Trung thực: nêu cả ưu điểm lẫn điểm cần cân nhắc (ví dụ: học phí cao hơn trường công)
3. Nhấn mạnh điểm nổi bật: không thu phí giữ chỗ, sĩ số ≤30, học phí ổn định 3 năm
4. Nếu hỏi so sánh với trường khác: phân tích khách quan theo tiêu chí cụ thể
5. Chỉ nói những gì có trong dữ liệu — không bịa thông tin

Trả lời bằng tiếng Việt. Thân thiện, thực tế.
"""

def build_truong_prompt(du_lieu: str, cau_hoi: str, lich_su: str = "") -> str:
    prompt = f"Lịch sử trò chuyện:\n{lich_su}\n\n" if lich_su else ""
    prompt += f"Thông tin bổ sung từ CSDL:\n---\n{du_lieu}\n---\n\nCâu hỏi: {cau_hoi}"
    return prompt


# ── Agent 3: Tuyển sinh ───────────────────────────────────────────────────────

TUYEN_SINH_SYSTEM = """
Bạn là nhân viên phòng tuyển sinh trường THPT Mai Hắc Đế, hướng dẫn quy trình đăng ký rõ ràng, chính xác.
Xưng "mình", gọi người hỏi là "bạn". Dùng emoji phù hợp.

Ngày hôm nay: {ngay_hom_nay}.

Thông tin tuyển sinh cơ bản (luôn có sẵn):
{thong_tin_truong}

Nguyên tắc:
0. Nếu hỏi thời hạn cụ thể mà không có dữ liệu → hướng dẫn xem website hoặc gọi điện
1. Liệt kê hồ sơ cần thiết theo từng bước, rõ ràng ✅
2. Nhắc mạnh: KHÔNG thu phí hồ sơ, KHÔNG thu phí giữ chỗ
3. Nếu hỏi phương thức xét tuyển → giải thích cả 2 phương án (thi + học bạ)
4. Với HS không thi lớp 10 công lập → trấn an: vẫn được xét học bạ
5. Cuối trả lời, nhắc số điện thoại và website trường

Liên hệ: 024.2239.3635 | 0972.951.456 | 0947.876.098 | 094.7653999
Website : maihacde.edu.vn | Email: c3maihacde@hanoiedu.vn

Trả lời bằng tiếng Việt. Rõ ràng, đúng quy trình.
"""

def build_tuyen_sinh_prompt(du_lieu: str, cau_hoi: str, lich_su: str = "") -> str:
    prompt = f"Lịch sử trò chuyện:\n{lich_su}\n\n" if lich_su else ""
    if du_lieu and "Không có dữ liệu" not in du_lieu:
        prompt += f"Thông tin bổ sung:\n---\n{du_lieu}\n---\n\n"
    prompt += f"Câu hỏi về tuyển sinh: {cau_hoi}"
    return prompt


# ── Agent 4: Học tập tại MHĐ ─────────────────────────────────────────────────

HOC_TAP_SYSTEM = """
Bạn là học sinh/cựu học sinh trường THPT Mai Hắc Đế, chia sẻ thật về chương trình học và đời sống học đường.
Xưng "mình", gọi người hỏi là "bạn". Tự nhiên, thực tế.

Ngày hôm nay: {ngay_hom_nay}.

Thông tin về trường MHĐ:
{thong_tin_truong}

Khi tư vấn về học tập tại MHĐ:

1. CHƯƠNG TRÌNH HỌC:
   - Chương trình chuẩn Bộ GD&ĐT (như mọi trường THPT)
   - Đặc điểm MHĐ: sĩ số ≤30 HS/lớp → thầy cô sát sao hơn
   - Ứng dụng CNTT, bài giảng điện tử
   - Tăng cường tiếng Anh (lễ hội, giao tiếp hàng ngày)

2. LỘ TRÌNH 3 NĂM THPT:
   - Lớp 10: nền tảng, làm quen môi trường mới
   - Lớp 11: chuyên sâu, bắt đầu định hướng khối thi
   - Lớp 12: ôn thi THPT Quốc gia, xét tuyển ĐH

3. HOẠT ĐỘNG SONG SONG:
   - Kỹ năng mềm, ngoại khoá, thiện nguyện
   - Hướng nghiệp: buổi nói chuyện chuyên đề, tư vấn ĐH
   - Thể thao, văn nghệ, cuộc thi sáng tạo

Nguyên tắc:
- Thực tế, không lý thuyết suông
- Nếu hỏi cụ thể về môn học → tư vấn theo chương trình THPT phổ thông
- Cuối trả lời hỏi bạn đang cần chuẩn bị gì hoặc đang ở giai đoạn nào

Trả lời bằng tiếng Việt. Cụ thể, dễ áp dụng.
"""

def build_hoc_tap_prompt(du_lieu: str, cau_hoi: str, lich_su: str = "") -> str:
    prompt = f"Lịch sử trò chuyện:\n{lich_su}\n\n" if lich_su else ""
    if du_lieu.strip() and "Không có dữ liệu" not in du_lieu:
        prompt += f"Thông tin bổ sung:\n---\n{du_lieu}\n---\n\n"
    prompt += f"Câu hỏi: {cau_hoi}\n\nHãy tư vấn thực tế về học tập tại MHĐ."
    return prompt


# ── Agent 5: Hướng nghiệp sau THPT ───────────────────────────────────────────

HUONG_NGHIEP_SYSTEM = """
Bạn là người anh/chị đang học ĐH hoặc đi làm, từng học THPT tại Hà Nội,
ngồi tư vấn thật cho em chuẩn bị vào lớp 10.
Xưng "mình", gọi người hỏi là "bạn". Thân thiện, thực tế, có chiều sâu.

Ngày hôm nay: {ngay_hom_nay}.

Tư vấn định hướng theo 4 chiều: ĐAM MÊ, TỐ CHẤT, TIỀM NĂNG, HÀNH ĐỘNG THỬ NGHIỆM.

Khi được hỏi về định hướng, chú ý:
1. CHỌN KHỐI THI THPT:
   - Khối A (Toán-Lý-Hóa): Kỹ thuật, Công nghệ, Y dược
   - Khối B (Toán-Hóa-Sinh): Y, Dược, Nông nghiệp, Sinh học
   - Khối C (Văn-Sử-Địa): Luật, Sư phạm, Báo chí, Nhân văn
   - Khối D (Văn-Toán-Anh): Kinh tế, Ngoại thương, Ngoại ngữ
   - Lớp 10-11 cần học đều tất cả, lớp 12 mới chốt khối

2. NGÀNH NGHỀ XU HƯỚNG 2025-2035:
   - CNTT, AI, khoa học dữ liệu (thiếu nhân lực trầm trọng)
   - Bán dẫn, vi mạch (Việt Nam đang hút đầu tư lớn)
   - Y tế, chăm sóc sức khỏe (dân số già hóa)
   - Logistics, chuỗi cung ứng
   - Năng lượng tái tạo
   - Tài chính số, fintech

3. CHUẨN BỊ TỪ LỚP 10:
   - Học đều các môn, chú ý môn sẽ thi ĐH
   - Tham gia CLB, hoạt động phù hợp sở thích
   - Tiếng Anh chuẩn bị càng sớm càng tốt

Nguyên tắc:
- Không phán xét sở thích
- Nếu phân vân 2 ngành: phân tích "5 năm sau làm gì?"
- Nếu chưa đủ thông tin về học sinh → hỏi thêm trước
- Gợi ý hành động cụ thể có thể làm NGAY

Trả lời bằng tiếng Việt. Thân thiện, thực tế.
"""

def build_huong_nghiep_prompt(du_lieu: str, cau_hoi: str, lich_su: str = "") -> str:
    prompt = f"Lịch sử trò chuyện:\n{lich_su}\n\n" if lich_su else ""
    if du_lieu.strip() and "Không có dữ liệu" not in du_lieu:
        prompt += f"Thông tin tham khảo:\n---\n{du_lieu}\n---\n\n"
    prompt += f"""Học sinh/phụ huynh hỏi: {cau_hoi}

Hãy tư vấn định hướng sau THPT theo 4 chiều: Đam mê, Tố chất, Tiềm năng, Hành động thử nghiệm.
Nếu chưa đủ thông tin về học sinh, hỏi thêm trước."""
    return prompt


# ── Agent 6: Dạy kiến thức ───────────────────────────────────────────────────

KIEN_THUC_SYSTEM = """
Bạn là gia sư dạy kèm cho học sinh THCS/THPT, giải thích kiến thức theo cách dễ hiểu nhất.
Xưng "mình", gọi người hỏi là "bạn". Kiên nhẫn, vui vẻ, không phán xét.

Ngày hôm nay: {ngay_hom_nay}.

== QUY TẮC KÝ HIỆU TOÁN HỌC ==
Luôn dùng ký hiệu ASCII thuần, KHÔNG dùng Unicode toán học:
  - Nhân        : *        (KHÔNG dùng ×)
  - Chia        : /        (KHÔNG dùng ÷)
  - Mũ          : ^        (KHÔNG dùng ², ³)
  - Căn bậc 2   : sqrt()   (KHÔNG dùng √)
  - Phân số     : a/b
  - Pi          : pi       (KHÔNG dùng π)
  - Góc         : goc()    (KHÔNG dùng ∠)

== KHI NÀO TẠO HÌNH ẢNH ==
Thêm [GENERATE_IMAGE: <mô tả tiếng Anh>] ở CUỐI khi cần minh hoạ:
  [GENERATE_IMAGE: plot y=x^2 and y=2*x+1, mark intersections]
  [GENERATE_IMAGE: draw right triangle sides a=3 b=4 c=5]
  [GENERATE_IMAGE: draw unit circle with special angles]
Chỉ thêm khi hình ảnh THỰC SỰ giúp hiểu bài.

== CÁCH GIẢI THÍCH ==
1. Bắt đầu bằng ví dụ thực tế quen thuộc
2. Tăng dần độ sâu: khái niệm → cách dùng → ví dụ → lỗi thường gặp
3. Code mẫu (lập trình): ngắn, có comment từng dòng
4. Hỏi 1 câu nhỏ cuối để kiểm tra hiểu bài
5. Gợi ý học tiếp: sau khái niệm này nên học gì

Trả lời bằng tiếng Việt. Thân thiện như gia sư tốt nhất bạn từng gặp.
"""

def build_kien_thuc_prompt(du_lieu: str, cau_hoi: str, lich_su: str = "") -> str:
    prompt = f"Lịch sử trò chuyện:\n{lich_su}\n\n" if lich_su else ""
    prompt += f"""Bạn hỏi: {cau_hoi}

Hãy giải thích dễ hiểu nhất, dùng ví dụ thực tế.
Nhớ dùng ký hiệu ASCII thuần cho toán học.
Nếu cần hình minh hoạ thì thêm [GENERATE_IMAGE: ...] ở cuối."""
    return prompt


# ── Aggregator ────────────────────────────────────────────────────────────────

AGGREGATOR_SYSTEM = """
Bạn tổng hợp thông tin từ nhiều chuyên gia về trường THPT Mai Hắc Đế thành 1 câu trả lời hoàn chỉnh.
Xưng "mình", gọi người hỏi là "bạn". Viết tự nhiên như đang nhắn tin.

Ngày hôm nay: {ngay_hom_nay}.

Nguyên tắc:
1. Kết hợp thông tin tự nhiên — không copy nguyên xi, không lặp lại
2. Trả lời câu hỏi chính trước, thông tin bổ sung sau
3. Độ dài vừa phải — đủ để trả lời, không dài dòng
4. Kết thúc bằng 1 câu hỏi gợi mở ngắn nếu cần tư vấn thêm
5. Nếu nhiều chuyên gia đều hỏi lại → chỉ hỏi 1 câu quan trọng nhất, không tự đoán

Trả lời bằng tiếng Việt. Thân thiện, tự nhiên.
"""

def build_aggregator_prompt(cau_hoi_goc: str, cac_ket_qua: dict) -> str:
    ten_agent = {
        "diem_chuan":   "Tư vấn điểm xét tuyển",
        "truong":       "Tư vấn thông tin trường",
        "tuyen_sinh":   "Tư vấn quy trình tuyển sinh",
        "hoc_tap":      "Tư vấn chương trình học",
        "huong_nghiep": "Tư vấn định hướng nghề nghiệp",
        "kien_thuc":    "Gia sư kiến thức",
    }
    ket_qua_text = ""
    for agent, ket_qua in cac_ket_qua.items():
        ten = ten_agent.get(agent, agent)
        ket_qua_text += f"--- {ten} ---\n{ket_qua}\n\n"
    return f"""Câu hỏi gốc: "{cau_hoi_goc}"

Thông tin từ các chuyên gia:
{ket_qua_text}
Hãy tổng hợp thành một câu trả lời hoàn chỉnh, tự nhiên."""


# ══════════════════════════════════════════════════════════════════════════════
# PHẦN 3 — INGEST: nạp dữ liệu vào ChromaDB
# ══════════════════════════════════════════════════════════════════════════════

def _khoi_tao_chroma(reset: bool = False):
    from chromadb.utils.embedding_functions import DefaultEmbeddingFunction
    client = chromadb.PersistentClient(path=CHROMA_DB_PATH)
    if reset:
        try:
            client.delete_collection(COLLECTION_NAME)
            log.info("Đã xóa collection cũ.")
        except Exception:
            pass
    collection = client.get_or_create_collection(
        name=COLLECTION_NAME,
        metadata={"hnsw:space": "cosine"},
        embedding_function=DefaultEmbeddingFunction(),
    )
    log.info(f"Collection '{COLLECTION_NAME}' — {collection.count()} documents hiện có.")
    return collection


def _tao_groq_client():
    return Groq(api_key=os.getenv("GROQ_API_KEY", GROQ_API_KEY))


def _embed(texts: list[str]) -> list[list[float]]:
    from chromadb.utils.embedding_functions import DefaultEmbeddingFunction
    ef = DefaultEmbeddingFunction()
    return ef(texts)


def _chunk_text(text: str) -> list[str]:
    text = re.sub(r'\s+', ' ', text).strip()
    if len(text) <= CHUNK_SIZE:
        return [text]
    chunks, start = [], 0
    while start < len(text):
        end = start + CHUNK_SIZE
        if end < len(text):
            boundary = max(
                text.rfind('. ', start, end),
                text.rfind('! ', start, end),
                text.rfind('? ', start, end),
                text.rfind('\n', start, end),
            )
            if boundary > start + CHUNK_SIZE // 2:
                end = boundary + 1
        chunk = text[start:end].strip()
        if chunk:
            chunks.append(chunk)
        start = end - CHUNK_OVERLAP
    return chunks


def _luu_vao_chroma(collection, groq_client, documents, metadatas, ids, ten_nguon):
    if not documents:
        return
    BATCH = 96
    log.info(f"  Đang embed {len(documents)} documents từ '{ten_nguon}'...")
    for i in tqdm(range(0, len(documents), BATCH), desc="  Embed", unit="batch"):
        batch_docs = documents[i:i+BATCH]
        batch_meta = metadatas[i:i+BATCH]
        batch_ids  = ids[i:i+BATCH]
        embeddings = _embed(batch_docs)
        collection.add(documents=batch_docs, embeddings=embeddings,
                       metadatas=batch_meta, ids=batch_ids)
    log.info(f"  Đã lưu {len(documents)} documents từ '{ten_nguon}'.")


def _suy_loai_mhd(filename: str) -> str:
    fn = filename.lower()
    if "diem"    in fn or "diem_chuan" in fn: return "diem_chuan"
    if "tuyen"   in fn or "tuyen_sinh" in fn: return "tuyen_sinh"
    if "hoc_phi" in fn or "hocphi"     in fn: return "hoc_phi"
    if "truong"  in fn or "gioi_thieu" in fn: return "thong_tin_truong"
    if "chuong"  in fn or "hoc_tap"    in fn: return "hoc_tap"
    return "khac"


def nap_thong_tin_truong_cung(collection, groq_client):
    """Nạp dữ liệu cứng về trường MHĐ vào ChromaDB (luôn làm mỗi lần ingest)."""
    log.info("  Nạp dữ liệu cứng về trường MHĐ...")
    chunks     = _chunk_text(THONG_TIN_TRUONG)
    documents  = chunks
    metadatas  = [{"nguon": "hard_coded", "truong": "MHD",
                   "loai": "thong_tin_truong", "nam": 2025, "chunk_idx": i}
                  for i in range(len(chunks))]
    ids        = [str(uuid.uuid4()) for _ in chunks]
    _luu_vao_chroma(collection, groq_client, documents, metadatas, ids, "Dữ liệu cứng MHĐ")


def nap_excel(collection, groq_client):
    if not os.path.exists(EXCEL_DIR):
        log.warning(f"Thư mục {EXCEL_DIR} không tồn tại, bỏ qua.")
        return
    files = [f for f in os.listdir(EXCEL_DIR) if f.endswith(('.xlsx', '.csv', '.xls'))]
    if not files:
        log.warning(f"Không có file Excel/CSV nào trong {EXCEL_DIR}.")
        return
    log.info(f"Tìm thấy {len(files)} file Excel/CSV.")
    for filename in files:
        filepath = os.path.join(EXCEL_DIR, filename)
        df = pd.read_csv(filepath, encoding='utf-8-sig') if filename.endswith('.csv') \
             else pd.read_excel(filepath)
        loai = _suy_loai_mhd(filename)
        documents, metadatas, ids = [], [], []
        for _, row in df.iterrows():
            if "noi_dung" in row and pd.notna(row["noi_dung"]):
                text = str(row["noi_dung"])
            else:
                parts = [f"{col}: {val}" for col, val in row.items()
                         if pd.notna(val) and str(val).strip()]
                text = " | ".join(parts)
            if not text.strip():
                continue
            documents.append(text)
            metadatas.append({
                "nguon": "excel", "file": filename, "loai": loai, "truong": "MHD",
                "nam": int(row.get("nam", 2025)) if "nam" in row else 2025,
            })
            ids.append(str(uuid.uuid4()))
        _luu_vao_chroma(collection, groq_client, documents, metadatas, ids, filename)


def nap_pdf(collection, groq_client):
    if not os.path.exists(PDF_DIR):
        log.warning(f"Thư mục {PDF_DIR} không tồn tại, bỏ qua.")
        return
    files = [f for f in os.listdir(PDF_DIR) if f.endswith('.pdf')]
    if not files:
        log.warning(f"Không có file PDF nào trong {PDF_DIR}.")
        return
    log.info(f"Tìm thấy {len(files)} file PDF.")
    for filename in files:
        filepath = os.path.join(PDF_DIR, filename)
        full_text = ""
        try:
            with pdfplumber.open(filepath) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        full_text += t + "\n"
        except Exception as e:
            log.error(f"  Lỗi đọc PDF {filename}: {e}")
            continue
        if not full_text.strip():
            log.warning(f"  Không trích được văn bản từ {filename}.")
            continue
        chunks    = _chunk_text(full_text)
        loai      = _suy_loai_mhd(filename)
        documents = chunks
        metadatas = [{"nguon": "pdf", "file": filename, "loai": loai, "truong": "MHD",
                      "nam": 2025, "chunk_idx": i}
                     for i in range(len(chunks))]
        ids = [str(uuid.uuid4()) for _ in chunks]
        _luu_vao_chroma(collection, groq_client, documents, metadatas, ids, filename)


def nap_web(collection, groq_client):
    if not WEBSITES_TO_CRAWL:
        log.warning("Chưa cấu hình WEBSITES_TO_CRAWL.")
        return
    headers = {"User-Agent": "Mozilla/5.0 (compatible; MHDBot/1.0)"}
    for site in WEBSITES_TO_CRAWL:
        url, truong = site["url"], site["truong"]
        log.info(f"  Đang crawl: {url}")
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            resp.raise_for_status()
            resp.encoding = 'utf-8'
        except Exception as e:
            log.error(f"  Lỗi crawl {url}: {e}")
            continue
        soup = BeautifulSoup(resp.text, 'html.parser')
        for tag in soup(["script", "style", "nav", "footer", "header", "aside"]):
            tag.decompose()
        main = soup.find("main") or soup.find("article") or soup.find("body")
        text = main.get_text(separator="\n") if main else soup.get_text(separator="\n")
        text = re.sub(r'\n{3,}', '\n\n', text).strip()
        if len(text) < 100:
            log.warning(f"  Trang {url} gần như trống, bỏ qua.")
            continue
        chunks    = _chunk_text(text)
        documents = chunks
        metadatas = [{"nguon": "web", "url": url, "truong": truong,
                      "loai": "thong_tin_truong", "nam": 2025, "chunk_idx": i}
                     for i in range(len(chunks))]
        ids = [str(uuid.uuid4()) for _ in chunks]
        _luu_vao_chroma(collection, groq_client, documents, metadatas, ids, url)
        time.sleep(1)


def chay_ingest():
    print("=" * 60)
    print("  INGEST DỮ LIỆU THPT MAI HẮC ĐẾ → CHROMADB")
    print("=" * 60)
    import sys
    reset = sys.stdin.isatty() and input("\nReset toàn bộ DB cũ? (y/N): ").strip().lower() == 'y'
    collection  = _khoi_tao_chroma(reset=reset)
    groq_client = _tao_groq_client()
    os.makedirs(EXCEL_DIR, exist_ok=True)
    os.makedirs(PDF_DIR,   exist_ok=True)

    print("\n--- Nạp dữ liệu cứng MHĐ ---")
    nap_thong_tin_truong_cung(collection, groq_client)
    print("\n--- Nạp từ Excel/CSV ---")
    nap_excel(collection, groq_client)
    print("\n--- Nạp từ PDF ---")
    nap_pdf(collection, groq_client)
    print("\n--- Crawl website maihacde.edu.vn ---")
    nap_web(collection, groq_client)

    print("\n" + "=" * 60)
    print(f"HOÀN TẤT! Tổng số documents: {collection.count()}")
    print(f"DB lưu tại: {os.path.abspath(CHROMA_DB_PATH)}")
    print("=" * 60)


# ══════════════════════════════════════════════════════════════════════════════
# PHẦN 4 — QUERY: nhận câu hỏi → tìm dữ liệu → gọi AI → trả lời
# ══════════════════════════════════════════════════════════════════════════════

class TuVanTuyenSinh:
    """
    Hệ thống tư vấn tuyển sinh lớp 10 THPT Mai Hắc Đế.

    Cách dùng:
        bot = TuVanTuyenSinh()
        print(bot.hoi("Trường MHĐ học phí bao nhiêu?"))
    """

    def __init__(self):
        log.info("Đang khởi động hệ thống tư vấn MHĐ...")
        from chromadb.utils.embedding_functions import DefaultEmbeddingFunction
        chroma_client = chromadb.PersistentClient(path=CHROMA_DB_PATH)
        self.collection = chroma_client.get_or_create_collection(
            COLLECTION_NAME,
            embedding_function=DefaultEmbeddingFunction(),
        )
        self.groq    = Groq(api_key=os.getenv("GROQ_API_KEY", GROQ_API_KEY))
        self.lich_su = []
        log.info("Sẵn sàng!")

    def hoi(self, cau_hoi: str) -> dict:
        """
        Hàm chính — nhóm web gọi vào đây.
        Trả về: {"tra_loi": str, "anh": list[str]}
        """
        log.info(f"[Câu hỏi] {cau_hoi}")

        agents, can_hoi_them = self._phan_loai(cau_hoi)
        log.info(f"[Agents] {agents}")

        if can_hoi_them and len(self.lich_su) < 4:
            self.lich_su.append({"role": "user",      "content": cau_hoi})
            self.lich_su.append({"role": "assistant",  "content": can_hoi_them})
            self.lich_su = self.lich_su[-20:]
            return {"tra_loi": can_hoi_them, "anh": []}

        ket_qua = {agent: self._chay_agent(agent, cau_hoi) for agent in agents}

        tra_loi_raw = (list(ket_qua.values())[0] if len(ket_qua) == 1
                       else self._tong_hop(cau_hoi, ket_qua))

        tra_loi, anh_list = xu_ly_anh_trong_tra_loi(tra_loi_raw)

        self.lich_su += [{"role": "user",      "content": cau_hoi},
                         {"role": "assistant",  "content": tra_loi}]
        self.lich_su = self.lich_su[-20:]

        return {"tra_loi": tra_loi, "anh": anh_list}

    def hoi_voi_anh(self, cau_hoi: str, image_base64: str, image_type: str = "image/jpeg") -> dict:
        log.info("[Vision] Đang phân tích ảnh...")
        try:
            vision_resp = self.groq.chat.completions.create(
                model="meta-llama/llama-4-scout-17b-16e-instruct",
                max_tokens=2000,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image_url", "image_url": {"url": f"data:{image_type};base64,{image_base64}"}},
                        {"type": "text", "text": (
                            "Đây là ảnh liên quan đến tuyển sinh hoặc học tập. "
                            "Hãy đọc và trích xuất TOÀN BỘ nội dung văn bản trong ảnh. "
                            "Trả lời bằng tiếng Việt. Liệt kê đầy đủ, không tóm tắt."
                        )},
                    ],
                }],
            )
            mo_ta_anh = vision_resp.choices[0].message.content.strip()
        except Exception as e:
            log.warning(f"[Vision] Lỗi: {e}")
            mo_ta_anh = "(Không thể phân tích ảnh)"
        cau_hoi_day_du = f"{cau_hoi}\n\n[Nội dung ảnh đính kèm]: {mo_ta_anh}".strip()
        return self.hoi(cau_hoi_day_du)

    def reset_lich_su(self):
        self.lich_su = []

    # ── Nội bộ ────────────────────────────────────────────────────────────────

    def _goi_llm(self, messages: list[dict], max_tokens: int = 1200) -> str:
        last_error = None
        for model in LLM_MODELS:
            try:
                resp = self.groq.chat.completions.create(
                    model=model, messages=messages, max_tokens=max_tokens)
                if model != LLM_MODELS[0]:
                    log.info(f"[Fallback] Đang dùng: {model}")
                return resp.choices[0].message.content.strip()
            except Exception as e:
                err_str = str(e).lower()
                if any(k in err_str for k in ["rate_limit", "rate limit", "429",
                                               "quota", "capacity", "overloaded",
                                               "tokens per", "requests per",
                                               "decommissioned", "no longer supported"]):
                    log.warning(f"[Fallback] {model} → thử tiếp: {e}")
                    last_error = e
                    time.sleep(0.5)
                    continue
                raise
        raise Exception(f"Tất cả model Groq quá tải. Thử lại sau. (Lỗi: {last_error})")

    def _phan_loai(self, cau_hoi: str) -> tuple[list[str], str]:
        try:
            text = self._goi_llm(
                messages=[
                    {"role": "system", "content": ORCHESTRATOR_SYSTEM},
                    {"role": "user",   "content": build_orchestrator_prompt(cau_hoi)},
                ],
                max_tokens=200,
            )
            text = text.replace("```json", "").replace("```", "").strip()
            data = json.loads(text)
            agents = data.get("agents", ["truong"])
            can_hoi_them = data.get("can_hoi_them", "").strip()
            hop_le = {"diem_chuan", "truong", "tuyen_sinh", "hoc_tap", "huong_nghiep", "kien_thuc"}
            return [a for a in agents if a in hop_le] or ["truong"], can_hoi_them
        except json.JSONDecodeError:
            return ["truong"], ""

    def _tim_du_lieu(self, cau_hoi: str, loai_filter: str = None) -> str:
        query_vec = _embed([cau_hoi])[0]
        try:
            where   = {"loai": loai_filter} if loai_filter else None
            results = self.collection.query(
                query_embeddings=[query_vec], n_results=TOP_K_RESULTS, where=where)
        except Exception:
            results = self.collection.query(
                query_embeddings=[query_vec], n_results=TOP_K_RESULTS)
        docs = results.get("documents", [[]])[0]
        return "\n---\n".join(docs) if docs else ""

    _KHONG_CO_DU_LIEU = "Không có dữ liệu liên quan trong hệ thống."
    _CAN_WEB_SEARCH   = {"diem_chuan", "truong", "tuyen_sinh"}

    def _tim_du_lieu_voi_web(self, ten_agent: str, cau_hoi: str, loai_filter: str | None) -> str:
        du_lieu_local = self._tim_du_lieu(cau_hoi, loai_filter)
        can_web = (
            ten_agent in self._CAN_WEB_SEARCH
            and (not du_lieu_local or len(du_lieu_local) < 300)
        )
        if can_web:
            log.info(f"[WebSearch] ChromaDB thiếu dữ liệu cho '{ten_agent}' → tra maihacde.edu.vn")
            du_lieu_web = tim_kiem_web(f"THPT Mai Hắc Đế {cau_hoi}", n_trang=3)
            if du_lieu_web:
                if du_lieu_local:
                    return f"{du_lieu_local}\n\n--- Bổ sung từ web ---\n{du_lieu_web}"
                return f"[Từ web]\n{du_lieu_web}"
        return du_lieu_local or self._KHONG_CO_DU_LIEU

    def _inject_truong_info(self, system_prompt: str) -> str:
        """Inject thông tin cứng của trường vào system prompt nếu có placeholder."""
        return (system_prompt
                .replace("{thong_tin_truong}", THONG_TIN_TRUONG)
                .replace("{ngay_hom_nay}", _ngay_hom_nay()))

    def _chay_agent(self, ten_agent: str, cau_hoi: str) -> str:
        cau_hinh = {
            "diem_chuan":   (DIEM_CHUAN_SYSTEM,   build_diem_chuan_prompt,   "diem_chuan"),
            "truong":       (TRUONG_SYSTEM,        build_truong_prompt,       "thong_tin_truong"),
            "tuyen_sinh":   (TUYEN_SINH_SYSTEM,    build_tuyen_sinh_prompt,   "tuyen_sinh"),
            "hoc_tap":      (HOC_TAP_SYSTEM,       build_hoc_tap_prompt,      "hoc_tap"),
            "huong_nghiep": (HUONG_NGHIEP_SYSTEM,  build_huong_nghiep_prompt, None),
            "kien_thuc":    (KIEN_THUC_SYSTEM,     build_kien_thuc_prompt,    None),
        }
        system_prompt, build_fn, loai_filter = cau_hinh[ten_agent]
        du_lieu = self._tim_du_lieu_voi_web(ten_agent, cau_hoi, loai_filter)

        lich_su_text = ""
        for msg in self.lich_su[-6:]:
            prefix = "Bạn" if msg["role"] == "user" else "Mình"
            lich_su_text += f"{prefix}: {msg['content']}\n"

        system_final = self._inject_truong_info(system_prompt)

        return self._goi_llm(
            messages=[
                {"role": "system", "content": system_final},
                {"role": "user",   "content": build_fn(du_lieu, cau_hoi, lich_su_text)},
            ],
            max_tokens=1200,
        )

    def _tong_hop(self, cau_hoi_goc: str, cac_ket_qua: dict) -> str:
        aggregator_final = (AGGREGATOR_SYSTEM
                            .replace("{ngay_hom_nay}", _ngay_hom_nay())
                            .replace("{thong_tin_truong}", THONG_TIN_TRUONG))
        return self._goi_llm(
            messages=[
                {"role": "system", "content": aggregator_final},
                {"role": "user",   "content": build_aggregator_prompt(cau_hoi_goc, cac_ket_qua)},
            ],
            max_tokens=1200,
        )


# ══════════════════════════════════════════════════════════════════════════════
# PHẦN 5 — MAIN
# ══════════════════════════════════════════════════════════════════════════════

def chay_chat():
    print("=" * 55)
    print("   AI TƯ VẤN TUYỂN SINH LỚP 10 — THPT MAI HẮC ĐẾ")
    print("   Gõ 'thoat' để thoát | 'moi' để bắt đầu lại")
    print("=" * 55)
    bot = TuVanTuyenSinh()
    while True:
        try:
            cau_hoi = input("\nBạn: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nTạm biệt!")
            break
        if not cau_hoi:
            continue
        if cau_hoi.lower() == "thoat":
            print("Tạm biệt!")
            break
        if cau_hoi.lower() == "moi":
            bot.reset_lich_su()
            print("--- Bắt đầu hội thoại mới ---")
            continue
        ket_qua = bot.hoi(cau_hoi)
        print(f"\nAI: {ket_qua['tra_loi']}")


def chay_server(port: int = 8000):
    from http.server import HTTPServer, BaseHTTPRequestHandler
    sessions: dict[str, TuVanTuyenSinh] = {}

    def lay_bot(sid: str) -> TuVanTuyenSinh:
        if sid not in sessions:
            sessions[sid] = TuVanTuyenSinh()
        return sessions[sid]

    class Handler(BaseHTTPRequestHandler):
        def do_OPTIONS(self):
            self._headers(200)

        def do_POST(self):
            body = self.rfile.read(int(self.headers.get("Content-Length", 0)))
            try:
                data    = json.loads(body)
                sid     = data.get("session_id", "default")
                cau_hoi = data.get("cau_hoi", "").strip()
                if not cau_hoi:
                    self._json({"loi": "Thiếu câu hỏi"}, 400)
                    return
                tra_loi = lay_bot(sid).hoi(cau_hoi)
                self._json({"tra_loi": tra_loi, "session_id": sid})
            except Exception as e:
                self._json({"loi": str(e)}, 500)

        def _headers(self, status):
            self.send_response(status)
            self.send_header("Content-Type",                  "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin",  "*")
            self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type")
            self.end_headers()

        def _json(self, data, status=200):
            self._headers(status)
            self.wfile.write(json.dumps(data, ensure_ascii=False).encode())

        def log_message(self, *args):
            pass

    print(f"Server tại http://localhost:{port}")
    HTTPServer(("", port), Handler).serve_forever()


def _huong_dan():
    print("""
╔══════════════════════════════════════════════════════╗
║   AI TƯ VẤN TUYỂN SINH LỚP 10 — THPT MAI HẮC ĐẾ    ║
╠══════════════════════════════════════════════════════╣
║  python tuyen_sinh_AI.py ingest    Nạp dữ liệu       ║
║  python tuyen_sinh_AI.py chat      Chat thử terminal  ║
║  python tuyen_sinh_AI.py server    Server cho web     ║
╚══════════════════════════════════════════════════════╝

Lần đầu dùng:
  1. pip install chromadb groq pandas pdfplumber
              requests beautifulsoup4 tqdm python-dotenv openpyxl
  2. Tạo .env: GROQ_API_KEY=gsk_...
  3. Đặt file dữ liệu vào data/excel/ hoặc data/pdf/ (nếu có)
  4. python tuyen_sinh_AI.py ingest
  5. python tuyen_sinh_AI.py chat
""")


if __name__ == "__main__":
    lenh = sys.argv[1] if len(sys.argv) > 1 else ""
    if lenh == "ingest":
        chay_ingest()
    elif lenh == "chat":
        chay_chat()
    elif lenh == "server":
        chay_server()
    else:
        _huong_dan()
