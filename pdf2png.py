"""
PDF ↔ Image Converter
모던 GUI 기반 PDF → PNG / 이미지 → PDF 변환 프로그램
"""

import os
import io
import re
import threading
import tkinter as tk
from tkinter import filedialog
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image, ImageTk

import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES


# ── CustomTkinter + TkinterDnD2 통합 클래스 ──────────────────────────────
class TkinterDnDCTk(ctk.CTk, TkinterDnD.DnDWrapper):
    """CustomTkinter 윈도우에 드래그 앤 드롭 지원을 추가"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)


# ── 색상 팔레트 ──────────────────────────────────────────────────────────
COLORS = {
    "bg_dark":      "#0f0f14",
    "bg_card":      "#1a1a24",
    "bg_hover":     "#24243a",
    "accent":       "#6c5ce7",
    "accent_hover": "#7f6ff0",
    "accent_glow":  "#a29bfe",
    "accent2":      "#00b894",   # 이미지→PDF 탭 포인트 색
    "accent2_hover":"#00cec9",
    "text":         "#e8e8f0",
    "text_dim":     "#8888a0",
    "success":      "#00cec9",
    "warning":      "#fdcb6e",
    "danger":       "#ff6b6b",
    "border":       "#2d2d44",
}

# 지원 이미지 확장자
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tiff", ".tif", ".gif"}


# ═══════════════════════════════════════════════════════════════════════════
#  PDF → PNG 탭 클래스
# ═══════════════════════════════════════════════════════════════════════════
class PdfToPngTab:
    def __init__(self, parent_frame):
        self.frame = parent_frame
        self.pdf_files: list[str] = []
        self.dpi_var = ctk.StringVar(value="300")
        self.merge_var = ctk.BooleanVar(value=True)
        self.is_converting = False
        self._build_ui()

    # ── UI 빌드 ──────────────────────────────────────────────────────────
    def _build_ui(self):
        main = self.frame

        # ── 드롭 영역 ───────────────────────────────────────────────────
        self.drop_frame = ctk.CTkFrame(
            main,
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            border_width=2,
            corner_radius=16,
            height=110,
        )
        self.drop_frame.pack(fill="x", pady=(8, 10))
        self.drop_frame.pack_propagate(False)

        drop_inner = ctk.CTkFrame(self.drop_frame, fg_color="transparent")
        drop_inner.place(relx=0.5, rely=0.5, anchor="center")

        self.drop_icon_label = ctk.CTkLabel(
            drop_inner, text="📥",
            font=ctk.CTkFont(size=40),
        )
        self.drop_icon_label.pack()

        self.drop_text_label = ctk.CTkLabel(
            drop_inner,
            text="PDF 파일을 여기에 드래그하거나 클릭하세요",
            font=ctk.CTkFont(size=14),
            text_color=COLORS["text_dim"],
        )
        self.drop_text_label.pack(pady=(4, 0))

        self.drop_sub_label = ctk.CTkLabel(
            drop_inner,
            text="여러 파일을 한 번에 추가할 수 있습니다",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_dim"],
        )
        self.drop_sub_label.pack()

        # 드롭 영역 이벤트
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)
        self.drop_frame.bind("<Button-1>", lambda e: self._browse_files())
        for w in [drop_inner, self.drop_icon_label, self.drop_text_label, self.drop_sub_label]:
            w.bind("<Button-1>", lambda e: self._browse_files())

        # ── 옵션 영역 ───────────────────────────────────────────────────
        options_frame = ctk.CTkFrame(
            main, fg_color=COLORS["bg_card"],
            corner_radius=12, border_color=COLORS["border"], border_width=1,
        )
        options_frame.pack(fill="x", pady=(0, 10))

        options_inner = ctk.CTkFrame(options_frame, fg_color="transparent")
        options_inner.pack(fill="x", padx=20, pady=12)

        # DPI 선택
        dpi_section = ctk.CTkFrame(options_inner, fg_color="transparent")
        dpi_section.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(
            dpi_section, text="🎨  해상도 (DPI)",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLORS["text"],
        ).pack(anchor="w")

        dpi_buttons_frame = ctk.CTkFrame(dpi_section, fg_color="transparent")
        dpi_buttons_frame.pack(anchor="w", pady=(6, 0))

        self.dpi_buttons = {}
        for dpi_val, label in [("150", "150 DPI\n일반"), ("300", "300 DPI\n고화질"), ("600", "600 DPI\n인쇄용")]:
            btn = ctk.CTkButton(
                dpi_buttons_frame,
                text=label,
                width=120, height=50,
                corner_radius=10,
                font=ctk.CTkFont(size=12),
                fg_color=COLORS["accent"] if dpi_val == "300" else COLORS["bg_hover"],
                hover_color=COLORS["accent_hover"],
                command=lambda d=dpi_val: self._set_dpi(d),
            )
            btn.pack(side="left", padx=(0, 8))
            self.dpi_buttons[dpi_val] = btn

        # 병합 옵션
        merge_section = ctk.CTkFrame(options_inner, fg_color="transparent")
        merge_section.pack(fill="x")

        ctk.CTkLabel(
            merge_section, text="📑  페이지 병합",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLORS["text"],
        ).pack(anchor="w")

        merge_desc = ctk.CTkFrame(merge_section, fg_color="transparent")
        merge_desc.pack(fill="x", pady=(6, 0))

        self.merge_switch = ctk.CTkSwitch(
            merge_desc,
            text="모든 페이지를 1장으로 합치기",
            font=ctk.CTkFont(size=12),
            variable=self.merge_var,
            progress_color=COLORS["accent"],
            button_color=COLORS["accent_glow"],
            button_hover_color=COLORS["accent_hover"],
        )
        self.merge_switch.pack(anchor="w")

        ctk.CTkLabel(
            merge_desc,
            text="OFF 시 페이지별 개별 PNG 파일로 저장됩니다",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_dim"],
        ).pack(anchor="w", padx=(54, 0))

        # ── 파일 리스트 ─────────────────────────────────────────────────
        self.file_list_frame = ctk.CTkFrame(
            main, fg_color=COLORS["bg_card"],
            corner_radius=12, border_color=COLORS["border"], border_width=1,
        )
        self.file_list_frame.pack(fill="both", expand=True, pady=(0, 10))

        file_list_header = ctk.CTkFrame(self.file_list_frame, fg_color="transparent")
        file_list_header.pack(fill="x", padx=16, pady=(12, 0))

        self.file_count_label = ctk.CTkLabel(
            file_list_header, text="📋 파일 목록 (0개)",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLORS["text"],
        )
        self.file_count_label.pack(side="left")

        self.clear_btn = ctk.CTkButton(
            file_list_header, text="전체 삭제",
            width=70, height=28,
            corner_radius=8,
            font=ctk.CTkFont(size=11),
            fg_color=COLORS["danger"],
            hover_color="#ff4757",
            command=self._clear_files,
        )
        self.clear_btn.pack(side="right")

        self.file_scroll = ctk.CTkScrollableFrame(
            self.file_list_frame, fg_color="transparent",
            height=120,
        )
        self.file_scroll.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        self.empty_label = ctk.CTkLabel(
            self.file_scroll,
            text="아직 추가된 파일이 없습니다",
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_dim"],
        )
        self.empty_label.pack(pady=20)

        # ── 하단 영역: 변환 버튼 + 프로그레스 ──────────────────────────
        bottom = ctk.CTkFrame(main, fg_color="transparent")
        bottom.pack(fill="x")

        self.progress_bar = ctk.CTkProgressBar(
            bottom,
            progress_color=COLORS["accent"],
            fg_color=COLORS["bg_card"],
            height=6,
            corner_radius=3,
        )
        self.progress_bar.pack(fill="x", pady=(0, 10))
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            bottom, text="",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_dim"],
        )
        self.status_label.pack(fill="x", pady=(0, 8))

        self.convert_btn = ctk.CTkButton(
            bottom,
            text="🚀  변환 시작",
            height=50,
            corner_radius=12,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            command=self._start_convert,
        )
        self.convert_btn.pack(fill="x")

    # ── DPI 선택 ─────────────────────────────────────────────────────────
    def _set_dpi(self, dpi: str):
        self.dpi_var.set(dpi)
        for d, btn in self.dpi_buttons.items():
            if d == dpi:
                btn.configure(fg_color=COLORS["accent"])
            else:
                btn.configure(fg_color=COLORS["bg_hover"])

    # ── 파일 추가 ────────────────────────────────────────────────────────
    def _browse_files(self):
        if self.is_converting:
            return
        files = filedialog.askopenfilenames(
            title="PDF 파일 선택",
            filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")],
        )
        if files:
            self._add_files(list(files))

    def _on_drop(self, event):
        if self.is_converting:
            return
        raw = event.data
        files = []
        if raw.startswith("{"):
            files = re.findall(r'\{(.+?)\}', raw)
        else:
            files = raw.split()

        pdf_files = [f for f in files if f.lower().endswith(".pdf")]
        if pdf_files:
            self._add_files(pdf_files)

    def _add_files(self, filepaths: list[str]):
        for fp in filepaths:
            fp_norm = os.path.normpath(fp)
            if fp_norm not in self.pdf_files and fp_norm.lower().endswith(".pdf"):
                self.pdf_files.append(fp_norm)
        self._refresh_file_list()

    def _remove_file(self, fp: str):
        if fp in self.pdf_files:
            self.pdf_files.remove(fp)
            self._refresh_file_list()

    def _clear_files(self):
        self.pdf_files.clear()
        self._refresh_file_list()

    def _refresh_file_list(self):
        for w in self.file_scroll.winfo_children():
            w.destroy()

        self.file_count_label.configure(text=f"📋 파일 목록 ({len(self.pdf_files)}개)")

        if not self.pdf_files:
            self.empty_label = ctk.CTkLabel(
                self.file_scroll,
                text="아직 추가된 파일이 없습니다",
                font=ctk.CTkFont(size=12),
                text_color=COLORS["text_dim"],
            )
            self.empty_label.pack(pady=20)
            return

        for i, fp in enumerate(self.pdf_files):
            row = ctk.CTkFrame(
                self.file_scroll,
                fg_color=COLORS["bg_hover"] if i % 2 == 0 else "transparent",
                corner_radius=8,
                height=36,
            )
            row.pack(fill="x", pady=1)
            row.pack_propagate(False)

            name = Path(fp).name
            folder = str(Path(fp).parent)
            ctk.CTkLabel(
                row, text=f"  📄 {name}",
                font=ctk.CTkFont(size=12),
                text_color=COLORS["text"],
                anchor="w",
            ).pack(side="left", fill="x", expand=True, padx=(4, 0))

            ctk.CTkLabel(
                row, text=folder,
                font=ctk.CTkFont(size=10),
                text_color=COLORS["text_dim"],
                anchor="e",
            ).pack(side="left", padx=(0, 8))

            ctk.CTkButton(
                row, text="✕", width=28, height=28,
                corner_radius=6,
                font=ctk.CTkFont(size=12),
                fg_color="transparent",
                hover_color=COLORS["danger"],
                text_color=COLORS["text_dim"],
                command=lambda f=fp: self._remove_file(f),
            ).pack(side="right", padx=4)

    # ── 변환 ─────────────────────────────────────────────────────────────
    def _start_convert(self):
        if self.is_converting or not self.pdf_files:
            return
        self.is_converting = True
        self.convert_btn.configure(
            state="disabled", text="⏳  변환 중...",
            fg_color=COLORS["bg_hover"],
        )
        thread = threading.Thread(target=self._do_convert, daemon=True)
        thread.start()

    def _do_convert(self):
        dpi = int(self.dpi_var.get())
        merge = self.merge_var.get()
        total_files = len(self.pdf_files)
        converted = []
        errors = []

        for idx, pdf_path in enumerate(self.pdf_files):
            try:
                self.frame.after(0, self._update_status,
                    f"변환 중... ({idx+1}/{total_files}) {Path(pdf_path).name}",
                    idx / total_files)

                doc = fitz.open(pdf_path)
                page_count = doc.page_count
                images = []

                for pi, page in enumerate(doc):
                    pix = page.get_pixmap(dpi=dpi)
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    images.append(img)

                    progress = (idx + (pi + 1) / page_count) / total_files
                    self.frame.after(0, self._update_progress, progress)

                doc.close()

                base = Path(pdf_path)
                stem = base.stem
                out_dir = base.parent

                if merge or page_count == 1:
                    if page_count == 1:
                        out_img = images[0]
                    else:
                        total_w = max(img.width for img in images)
                        total_h = sum(img.height for img in images)
                        out_img = Image.new("RGB", (total_w, total_h), "white")
                        y = 0
                        for img in images:
                            out_img.paste(img, (0, y))
                            y += img.height

                    out_path = str(out_dir / f"{stem}.png")
                    out_img.save(out_path)
                    converted.append(out_path)
                else:
                    for pi, img in enumerate(images):
                        out_path = str(out_dir / f"{stem}_{pi+1}.png")
                        img.save(out_path)
                        converted.append(out_path)

            except Exception as e:
                errors.append((Path(pdf_path).name, str(e)))

        self.frame.after(0, self._on_convert_done, converted, errors)

    def _update_status(self, text: str, progress: float):
        self.status_label.configure(text=text)
        self.progress_bar.set(progress)

    def _update_progress(self, progress: float):
        self.progress_bar.set(progress)

    def _on_convert_done(self, converted: list[str], errors: list[tuple[str, str]]):
        self.is_converting = False
        self.progress_bar.set(1.0)
        self.convert_btn.configure(
            state="normal", text="🚀  변환 시작",
            fg_color=COLORS["accent"],
        )

        if errors:
            self.status_label.configure(
                text=f"⚠️  {len(errors)}개 파일 오류 발생",
                text_color=COLORS["danger"],
            )
        else:
            self.status_label.configure(
                text=f"✅  변환 완료! {len(converted)}개 PNG 파일 생성",
                text_color=COLORS["success"],
            )

        if converted:
            self._show_result_popup(converted, errors)

    def _show_result_popup(self, converted: list[str], errors: list[tuple[str, str]]):
        popup = ctk.CTkToplevel(self.frame)
        popup.title("변환 결과")
        popup.geometry("500x400")
        popup.configure(fg_color=COLORS["bg_dark"])
        popup.transient(self.frame.winfo_toplevel())
        popup.grab_set()

        ctk.CTkLabel(
            popup, text="✅  변환 완료",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=COLORS["success"],
        ).pack(pady=(20, 4))

        ctk.CTkLabel(
            popup, text=f"{len(converted)}개 PNG 파일이 생성되었습니다",
            font=ctk.CTkFont(size=13),
            text_color=COLORS["text_dim"],
        ).pack(pady=(0, 12))

        scroll = ctk.CTkScrollableFrame(
            popup, fg_color=COLORS["bg_card"],
            corner_radius=10,
        )
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        for fp in converted:
            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", pady=2)

            ctk.CTkLabel(
                row, text=f"📄 {Path(fp).name}",
                font=ctk.CTkFont(size=12),
                text_color=COLORS["text"],
                anchor="w",
            ).pack(side="left", fill="x", expand=True)

            ctk.CTkButton(
                row, text="📂", width=32, height=28,
                corner_radius=6,
                fg_color="transparent",
                hover_color=COLORS["bg_hover"],
                command=lambda f=fp: _open_folder(f),
            ).pack(side="right")

        if errors:
            ctk.CTkLabel(
                popup,
                text=f"⚠️  {len(errors)}개 파일 오류",
                font=ctk.CTkFont(size=12),
                text_color=COLORS["danger"],
            ).pack(pady=(0, 4))

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(fill="x", padx=16, pady=(0, 16))

        if converted:
            ctk.CTkButton(
                btn_frame, text="📂 폴더 열기",
                height=40, corner_radius=10,
                font=ctk.CTkFont(size=13),
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                command=lambda: _open_folder(converted[0]),
            ).pack(side="left", fill="x", expand=True, padx=(0, 4))

        ctk.CTkButton(
            btn_frame, text="닫기",
            height=40, corner_radius=10,
            font=ctk.CTkFont(size=13),
            fg_color=COLORS["bg_hover"],
            hover_color=COLORS["border"],
            command=popup.destroy,
        ).pack(side="right", fill="x", expand=True, padx=(4, 0))


# ═══════════════════════════════════════════════════════════════════════════
#  이미지 → PDF 탭 클래스
# ═══════════════════════════════════════════════════════════════════════════
class ImgToPdfTab:
    def __init__(self, parent_frame):
        self.frame = parent_frame
        self.img_files: list[str] = []
        self.is_converting = False
        self._build_ui()

    # ── UI 빌드 ──────────────────────────────────────────────────────────
    def _build_ui(self):
        main = self.frame

        # ── 드롭 영역 ───────────────────────────────────────────────────
        self.drop_frame = ctk.CTkFrame(
            main,
            fg_color=COLORS["bg_card"],
            border_color=COLORS["accent2"],
            border_width=2,
            corner_radius=16,
            height=110,
        )
        self.drop_frame.pack(fill="x", pady=(8, 10))
        self.drop_frame.pack_propagate(False)

        drop_inner = ctk.CTkFrame(self.drop_frame, fg_color="transparent")
        drop_inner.place(relx=0.5, rely=0.5, anchor="center")

        self.drop_icon_label = ctk.CTkLabel(
            drop_inner, text="🖼️",
            font=ctk.CTkFont(size=40),
        )
        self.drop_icon_label.pack()

        self.drop_text_label = ctk.CTkLabel(
            drop_inner,
            text="이미지 파일을 여기에 드래그하거나 클릭하세요",
            font=ctk.CTkFont(size=14),
            text_color=COLORS["text_dim"],
        )
        self.drop_text_label.pack(pady=(4, 0))

        self.drop_sub_label = ctk.CTkLabel(
            drop_inner,
            text="PNG · JPG · JPEG · BMP · WEBP · TIFF · GIF 지원",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_dim"],
        )
        self.drop_sub_label.pack()

        # 드롭 영역 이벤트
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)
        self.drop_frame.bind("<Button-1>", lambda e: self._browse_files())
        for w in [drop_inner, self.drop_icon_label, self.drop_text_label, self.drop_sub_label]:
            w.bind("<Button-1>", lambda e: self._browse_files())

        # ── 옵션 영역 ───────────────────────────────────────────────────
        options_frame = ctk.CTkFrame(
            main, fg_color=COLORS["bg_card"],
            corner_radius=12, border_color=COLORS["border"], border_width=1,
        )
        options_frame.pack(fill="x", pady=(0, 10))

        options_inner = ctk.CTkFrame(options_frame, fg_color="transparent")
        options_inner.pack(fill="x", padx=20, pady=14)

        # PDF 용지 크기
        size_row = ctk.CTkFrame(options_inner, fg_color="transparent")
        size_row.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            size_row, text="📐  PDF 페이지 크기",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLORS["text"],
        ).pack(anchor="w")

        size_btns = ctk.CTkFrame(size_row, fg_color="transparent")
        size_btns.pack(anchor="w", pady=(6, 0))

        self.page_size_var = ctk.StringVar(value="original")
        self.size_buttons = {}
        for val, label in [("original", "원본 크기\n유지"), ("A4", "A4\n210×297mm"), ("letter", "Letter\n216×279mm")]:
            btn = ctk.CTkButton(
                size_btns,
                text=label,
                width=120, height=50,
                corner_radius=10,
                font=ctk.CTkFont(size=12),
                fg_color=COLORS["accent2"] if val == "original" else COLORS["bg_hover"],
                hover_color=COLORS["accent2_hover"],
                command=lambda v=val: self._set_page_size(v),
            )
            btn.pack(side="left", padx=(0, 8))
            self.size_buttons[val] = btn

        # 여백 옵션
        margin_row = ctk.CTkFrame(options_inner, fg_color="transparent")
        margin_row.pack(fill="x")

        ctk.CTkLabel(
            margin_row, text="📏  여백",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLORS["text"],
        ).pack(anchor="w")

        margin_inner = ctk.CTkFrame(margin_row, fg_color="transparent")
        margin_inner.pack(fill="x", pady=(6, 0))

        self.margin_var = ctk.StringVar(value="none")
        self.margin_buttons = {}
        for val, label in [("none", "여백 없음"), ("small", "작게  10px"), ("medium", "보통  20px"), ("large", "크게  40px")]:
            btn = ctk.CTkButton(
                margin_inner,
                text=label,
                width=100, height=36,
                corner_radius=10,
                font=ctk.CTkFont(size=11),
                fg_color=COLORS["accent2"] if val == "none" else COLORS["bg_hover"],
                hover_color=COLORS["accent2_hover"],
                command=lambda v=val: self._set_margin(v),
            )
            btn.pack(side="left", padx=(0, 6))
            self.margin_buttons[val] = btn

        # ── 파일 리스트 ─────────────────────────────────────────────────
        self.file_list_frame = ctk.CTkFrame(
            main, fg_color=COLORS["bg_card"],
            corner_radius=12, border_color=COLORS["border"], border_width=1,
        )
        self.file_list_frame.pack(fill="both", expand=True, pady=(0, 10))

        file_list_header = ctk.CTkFrame(self.file_list_frame, fg_color="transparent")
        file_list_header.pack(fill="x", padx=16, pady=(12, 0))

        self.file_count_label = ctk.CTkLabel(
            file_list_header, text="🖼️ 이미지 목록 (0개)",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLORS["text"],
        )
        self.file_count_label.pack(side="left")

        self.clear_btn = ctk.CTkButton(
            file_list_header, text="전체 삭제",
            width=70, height=28,
            corner_radius=8,
            font=ctk.CTkFont(size=11),
            fg_color=COLORS["danger"],
            hover_color="#ff4757",
            command=self._clear_files,
        )
        self.clear_btn.pack(side="right")

        # 순서 안내
        ctk.CTkLabel(
            self.file_list_frame,
            text="↕ 파일 순서가 PDF 페이지 순서가 됩니다  |  ▲▼ 버튼으로 순서 변경",
            font=ctk.CTkFont(size=10),
            text_color=COLORS["text_dim"],
        ).pack(anchor="w", padx=16, pady=(2, 0))

        self.file_scroll = ctk.CTkScrollableFrame(
            self.file_list_frame, fg_color="transparent",
            height=140,
        )
        self.file_scroll.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        self.empty_label = ctk.CTkLabel(
            self.file_scroll,
            text="아직 추가된 이미지 파일이 없습니다",
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_dim"],
        )
        self.empty_label.pack(pady=20)

        # ── 하단 영역 ────────────────────────────────────────────────────
        bottom = ctk.CTkFrame(main, fg_color="transparent")
        bottom.pack(fill="x")

        self.progress_bar = ctk.CTkProgressBar(
            bottom,
            progress_color=COLORS["accent2"],
            fg_color=COLORS["bg_card"],
            height=6,
            corner_radius=3,
        )
        self.progress_bar.pack(fill="x", pady=(0, 10))
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            bottom, text="",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_dim"],
        )
        self.status_label.pack(fill="x", pady=(0, 8))

        self.convert_btn = ctk.CTkButton(
            bottom,
            text="📑  PDF로 변환",
            height=50,
            corner_radius=12,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=COLORS["accent2"],
            hover_color=COLORS["accent2_hover"],
            command=self._start_convert,
        )
        self.convert_btn.pack(fill="x")

    # ── 옵션 선택 ─────────────────────────────────────────────────────────
    def _set_page_size(self, val: str):
        self.page_size_var.set(val)
        for v, btn in self.size_buttons.items():
            btn.configure(fg_color=COLORS["accent2"] if v == val else COLORS["bg_hover"])

    def _set_margin(self, val: str):
        self.margin_var.set(val)
        for v, btn in self.margin_buttons.items():
            btn.configure(fg_color=COLORS["accent2"] if v == val else COLORS["bg_hover"])

    # ── 파일 추가 ────────────────────────────────────────────────────────
    def _browse_files(self):
        if self.is_converting:
            return
        files = filedialog.askopenfilenames(
            title="이미지 파일 선택",
            filetypes=[
                ("이미지 파일", "*.png *.jpg *.jpeg *.bmp *.webp *.tiff *.tif *.gif"),
                ("모든 파일", "*.*"),
            ],
        )
        if files:
            self._add_files(list(files))

    def _on_drop(self, event):
        if self.is_converting:
            return
        raw = event.data
        files = []
        if raw.startswith("{"):
            files = re.findall(r'\{(.+?)\}', raw)
        else:
            files = raw.split()

        img_files = [f for f in files if Path(f).suffix.lower() in IMAGE_EXTS]
        if img_files:
            self._add_files(img_files)

    def _add_files(self, filepaths: list[str]):
        for fp in filepaths:
            fp_norm = os.path.normpath(fp)
            if fp_norm not in self.img_files and Path(fp_norm).suffix.lower() in IMAGE_EXTS:
                self.img_files.append(fp_norm)
        self._refresh_file_list()

    def _remove_file(self, fp: str):
        if fp in self.img_files:
            self.img_files.remove(fp)
            self._refresh_file_list()

    def _move_file(self, fp: str, direction: int):
        idx = self.img_files.index(fp)
        new_idx = idx + direction
        if 0 <= new_idx < len(self.img_files):
            self.img_files[idx], self.img_files[new_idx] = self.img_files[new_idx], self.img_files[idx]
            self._refresh_file_list()

    def _clear_files(self):
        self.img_files.clear()
        self._refresh_file_list()

    def _refresh_file_list(self):
        for w in self.file_scroll.winfo_children():
            w.destroy()

        self.file_count_label.configure(text=f"🖼️ 이미지 목록 ({len(self.img_files)}개)")

        if not self.img_files:
            self.empty_label = ctk.CTkLabel(
                self.file_scroll,
                text="아직 추가된 이미지 파일이 없습니다",
                font=ctk.CTkFont(size=12),
                text_color=COLORS["text_dim"],
            )
            self.empty_label.pack(pady=20)
            return

        total = len(self.img_files)
        for i, fp in enumerate(self.img_files):
            row = ctk.CTkFrame(
                self.file_scroll,
                fg_color=COLORS["bg_hover"] if i % 2 == 0 else "transparent",
                corner_radius=8,
                height=38,
            )
            row.pack(fill="x", pady=1)
            row.pack_propagate(False)

            # 페이지 번호
            ctk.CTkLabel(
                row, text=f"  {i+1:02d}",
                font=ctk.CTkFont(size=11),
                text_color=COLORS["accent2"],
                width=28,
            ).pack(side="left")

            # 아이콘 + 파일명
            ext = Path(fp).suffix.lower()
            icon = "🖼️"
            name = Path(fp).name
            ctk.CTkLabel(
                row, text=f"{icon} {name}",
                font=ctk.CTkFont(size=12),
                text_color=COLORS["text"],
                anchor="w",
            ).pack(side="left", fill="x", expand=True, padx=(2, 0))

            # ▼▲ 순서 이동 버튼
            btn_up = ctk.CTkButton(
                row, text="▲", width=26, height=26,
                corner_radius=5,
                font=ctk.CTkFont(size=10),
                fg_color="transparent",
                hover_color=COLORS["bg_card"],
                text_color=COLORS["text_dim"] if i > 0 else COLORS["border"],
                command=lambda f=fp: self._move_file(f, -1),
                state="normal" if i > 0 else "disabled",
            )
            btn_up.pack(side="right", padx=(0, 2))

            btn_down = ctk.CTkButton(
                row, text="▼", width=26, height=26,
                corner_radius=5,
                font=ctk.CTkFont(size=10),
                fg_color="transparent",
                hover_color=COLORS["bg_card"],
                text_color=COLORS["text_dim"] if i < total - 1 else COLORS["border"],
                command=lambda f=fp: self._move_file(f, 1),
                state="normal" if i < total - 1 else "disabled",
            )
            btn_down.pack(side="right", padx=(0, 2))

            # 삭제 버튼
            ctk.CTkButton(
                row, text="✕", width=26, height=26,
                corner_radius=5,
                font=ctk.CTkFont(size=11),
                fg_color="transparent",
                hover_color=COLORS["danger"],
                text_color=COLORS["text_dim"],
                command=lambda f=fp: self._remove_file(f),
            ).pack(side="right", padx=(0, 4))

    # ── 변환 ─────────────────────────────────────────────────────────────
    def _start_convert(self):
        if self.is_converting or not self.img_files:
            return

        # 저장 경로 선택
        first_dir = str(Path(self.img_files[0]).parent)
        first_stem = Path(self.img_files[0]).stem
        out_path = filedialog.asksaveasfilename(
            title="PDF 저장 위치 선택",
            initialdir=first_dir,
            initialfile=f"{first_stem}_output.pdf",
            defaultextension=".pdf",
            filetypes=[("PDF 파일", "*.pdf")],
        )
        if not out_path:
            return

        self.is_converting = True
        self.convert_btn.configure(
            state="disabled", text="⏳  변환 중...",
            fg_color=COLORS["bg_hover"],
        )
        thread = threading.Thread(
            target=self._do_convert,
            args=(out_path,),
            daemon=True,
        )
        thread.start()

    def _do_convert(self, out_path: str):
        page_size = self.page_size_var.get()
        margin_map = {"none": 0, "small": 10, "medium": 20, "large": 40}
        margin = margin_map.get(self.margin_var.get(), 0)
        total = len(self.img_files)
        errors = []

        try:
            pdf_images = []

            for i, fp in enumerate(self.img_files):
                self.frame.after(0, self._update_status,
                    f"처리 중... ({i+1}/{total}) {Path(fp).name}",
                    i / total)
                try:
                    img = Image.open(fp)
                    # RGBA/P 모드 → RGB 변환 (PDF 호환성)
                    if img.mode in ("RGBA", "P", "LA"):
                        background = Image.new("RGB", img.size, (255, 255, 255))
                        if img.mode in ("RGBA", "LA"):
                            background.paste(img, mask=img.split()[-1])
                        else:
                            background.paste(img.convert("RGBA"), mask=img.convert("RGBA").split()[-1])
                        img = background
                    elif img.mode != "RGB":
                        img = img.convert("RGB")

                    if page_size == "A4":
                        # A4: 210×297mm @72dpi → 595×842px
                        a4_w, a4_h = 595, 842
                        m2 = margin * 2
                        avail_w, avail_h = a4_w - m2, a4_h - m2
                        img.thumbnail((avail_w, avail_h), Image.LANCZOS)
                        canvas = Image.new("RGB", (a4_w, a4_h), (255, 255, 255))
                        x = (a4_w - img.width) // 2
                        y = (a4_h - img.height) // 2
                        canvas.paste(img, (x, y))
                        img = canvas
                    elif page_size == "letter":
                        # Letter: 216×279mm @72dpi → 612×792px
                        lt_w, lt_h = 612, 792
                        m2 = margin * 2
                        avail_w, avail_h = lt_w - m2, lt_h - m2
                        img.thumbnail((avail_w, avail_h), Image.LANCZOS)
                        canvas = Image.new("RGB", (lt_w, lt_h), (255, 255, 255))
                        x = (lt_w - img.width) // 2
                        y = (lt_h - img.height) // 2
                        canvas.paste(img, (x, y))
                        img = canvas
                    else:
                        # 원본 크기 + 여백
                        if margin > 0:
                            new_w = img.width + margin * 2
                            new_h = img.height + margin * 2
                            canvas = Image.new("RGB", (new_w, new_h), (255, 255, 255))
                            canvas.paste(img, (margin, margin))
                            img = canvas

                    pdf_images.append(img)
                except Exception as e:
                    errors.append((Path(fp).name, str(e)))

                self.frame.after(0, self._update_progress, (i + 1) / total)

            if not pdf_images:
                raise ValueError("변환 가능한 이미지가 없습니다.")

            # PIL로 PDF 저장
            first_img = pdf_images[0]
            rest = pdf_images[1:]
            first_img.save(
                out_path,
                format="PDF",
                save_all=True,
                append_images=rest,
                resolution=72,
            )

            self.frame.after(0, self._on_convert_done, out_path, errors)

        except Exception as e:
            self.frame.after(0, self._on_convert_error, str(e))

    def _update_status(self, text: str, progress: float):
        self.status_label.configure(text=text, text_color=COLORS["text_dim"])
        self.progress_bar.set(progress)

    def _update_progress(self, progress: float):
        self.progress_bar.set(progress)

    def _on_convert_done(self, out_path: str, errors: list):
        self.is_converting = False
        self.progress_bar.set(1.0)
        self.convert_btn.configure(
            state="normal", text="📑  PDF로 변환",
            fg_color=COLORS["accent2"],
        )

        if errors:
            self.status_label.configure(
                text=f"⚠️  {len(errors)}개 파일 오류 (나머지는 변환 완료)",
                text_color=COLORS["warning"],
            )
        else:
            self.status_label.configure(
                text=f"✅  PDF 생성 완료!  {Path(out_path).name}",
                text_color=COLORS["success"],
            )

        self._show_result_popup(out_path, errors)

    def _on_convert_error(self, err: str):
        self.is_converting = False
        self.progress_bar.set(0)
        self.convert_btn.configure(
            state="normal", text="📑  PDF로 변환",
            fg_color=COLORS["accent2"],
        )
        self.status_label.configure(
            text=f"❌  오류: {err}",
            text_color=COLORS["danger"],
        )

    def _show_result_popup(self, out_path: str, errors: list):
        popup = ctk.CTkToplevel(self.frame)
        popup.title("변환 결과")
        popup.geometry("480x320")
        popup.configure(fg_color=COLORS["bg_dark"])
        popup.transient(self.frame.winfo_toplevel())
        popup.grab_set()

        ctk.CTkLabel(
            popup, text="✅  PDF 생성 완료",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=COLORS["accent2"],
        ).pack(pady=(24, 4))

        ctk.CTkLabel(
            popup, text=Path(out_path).name,
            font=ctk.CTkFont(size=13),
            text_color=COLORS["text"],
        ).pack(pady=(0, 4))

        ctk.CTkLabel(
            popup, text=str(Path(out_path).parent),
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_dim"],
        ).pack(pady=(0, 16))

        if errors:
            err_text = "\n".join(f"• {n}: {e}" for n, e in errors)
            ctk.CTkLabel(
                popup,
                text=f"⚠️  오류 발생 파일:\n{err_text}",
                font=ctk.CTkFont(size=11),
                text_color=COLORS["warning"],
                justify="left",
            ).pack(padx=20, pady=(0, 12))

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=(0, 20))

        ctk.CTkButton(
            btn_frame, text="📂 폴더 열기",
            height=40, corner_radius=10,
            font=ctk.CTkFont(size=13),
            fg_color=COLORS["accent2"],
            hover_color=COLORS["accent2_hover"],
            command=lambda: _open_folder(out_path),
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))

        ctk.CTkButton(
            btn_frame, text="닫기",
            height=40, corner_radius=10,
            font=ctk.CTkFont(size=13),
            fg_color=COLORS["bg_hover"],
            hover_color=COLORS["border"],
            command=popup.destroy,
        ).pack(side="right", fill="x", expand=True, padx=(6, 0))


# ═══════════════════════════════════════════════════════════════════════════
#  PDF 페이지 교체 탭 클래스
# ═══════════════════════════════════════════════════════════════════════════
class PdfPageReplaceTab:
    """PDF의 특정 페이지를 이미지 파일로 대체하는 탭"""

    THUMB_W = 120   # 썸네일 너비
    THUMB_H = 160   # 썸네일 높이
    ACCENT  = "#e17055"   # 오렌지 계열 포인트 컬러
    ACCENT_H = "#d63031"

    def __init__(self, parent_frame):
        self.frame = parent_frame
        self.pdf_path: str | None = None
        self.doc: fitz.Document | None = None
        self.page_count: int = 0
        # {page_index(0-based): image_path}
        self.replacements: dict[int, str] = {}
        self.thumb_images: list = []   # PhotoImage 레퍼런스 보관용
        self.is_saving = False
        self._build_ui()

    # ── UI ──────────────────────────────────────────────────────────────────
    def _build_ui(self):
        main = self.frame

        # ── PDF 드롭 영역 ───────────────────────────────────────────────────
        self.drop_frame = ctk.CTkFrame(
            main,
            fg_color=COLORS["bg_card"],
            border_color=self.ACCENT,
            border_width=2,
            corner_radius=16,
            height=110,
        )
        self.drop_frame.pack(fill="x", pady=(8, 10))
        self.drop_frame.pack_propagate(False)

        drop_inner = ctk.CTkFrame(self.drop_frame, fg_color="transparent")
        drop_inner.place(relx=0.5, rely=0.5, anchor="center")

        self.drop_icon = ctk.CTkLabel(drop_inner, text="📂",
                                      font=ctk.CTkFont(size=40))
        self.drop_icon.pack()

        self.drop_text = ctk.CTkLabel(
            drop_inner,
            text="PDF 파일을 여기에 드래그하거나 클릭하세요",
            font=ctk.CTkFont(size=14),
            text_color=COLORS["text_dim"],
        )
        self.drop_text.pack(pady=(4, 0))

        self.drop_sub = ctk.CTkLabel(
            drop_inner,
            text="교체할 페이지를 선택 후 이미지를 지정합니다",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_dim"],
        )
        self.drop_sub.pack()

        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self._on_drop_pdf)
        self.drop_frame.bind("<Button-1>", lambda e: self._browse_pdf())
        for w in [drop_inner, self.drop_icon, self.drop_text, self.drop_sub]:
            w.bind("<Button-1>", lambda e: self._browse_pdf())

        # ── 교체 목록 안내 ──────────────────────────────────────────────────
        ctrl_bar = ctk.CTkFrame(main, fg_color=COLORS["bg_card"],
                                corner_radius=12, border_color=COLORS["border"],
                                border_width=1)
        ctrl_bar.pack(fill="x", pady=(0, 10))

        ctrl_inner = ctk.CTkFrame(ctrl_bar, fg_color="transparent")
        ctrl_inner.pack(fill="x", padx=16, pady=10)

        self.pdf_name_label = ctk.CTkLabel(
            ctrl_inner,
            text="PDF를 먼저 불러오세요",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLORS["text_dim"],
            anchor="w",
        )
        self.pdf_name_label.pack(side="left", fill="x", expand=True)

        self.replace_count_label = ctk.CTkLabel(
            ctrl_inner,
            text="교체: 0페이지",
            font=ctk.CTkFont(size=12),
            text_color=self.ACCENT,
        )
        self.replace_count_label.pack(side="left", padx=(0, 12))

        self.clear_replace_btn = ctk.CTkButton(
            ctrl_inner, text="교체 초기화",
            width=80, height=30, corner_radius=8,
            font=ctk.CTkFont(size=11),
            fg_color=COLORS["bg_hover"],
            hover_color=COLORS["border"],
            command=self._clear_replacements,
        )
        self.clear_replace_btn.pack(side="left")

        # ── 페이지 그리드 영역 ──────────────────────────────────────────────
        grid_outer = ctk.CTkFrame(main, fg_color=COLORS["bg_card"],
                                  corner_radius=12, border_color=COLORS["border"],
                                  border_width=1)
        grid_outer.pack(fill="both", expand=True, pady=(0, 10))

        grid_hdr = ctk.CTkFrame(grid_outer, fg_color="transparent")
        grid_hdr.pack(fill="x", padx=14, pady=(10, 4))

        ctk.CTkLabel(
            grid_hdr, text="🗂️  페이지 목록",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLORS["text"],
        ).pack(side="left")

        ctk.CTkLabel(
            grid_hdr,
            text="페이지 클릭 → 교체 이미지 선택  |  🔄 클릭 → 교체 해제",
            font=ctk.CTkFont(size=10),
            text_color=COLORS["text_dim"],
        ).pack(side="left", padx=10)

        self.grid_scroll = ctk.CTkScrollableFrame(
            grid_outer, fg_color="transparent",
        )
        self.grid_scroll.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.empty_label = ctk.CTkLabel(
            self.grid_scroll,
            text="PDF를 불러오면 페이지 미리보기가 표시됩니다",
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_dim"],
        )
        self.empty_label.pack(pady=30)

        # ── 하단 영역: 프로그레스바 + 상태 + 저장 버튼 ─────────────────────
        bottom = ctk.CTkFrame(main, fg_color="transparent")
        bottom.pack(fill="x")

        self.progress_bar = ctk.CTkProgressBar(
            bottom,
            progress_color=self.ACCENT,
            fg_color=COLORS["bg_card"],
            height=6,
            corner_radius=3,
        )
        self.progress_bar.pack(fill="x", pady=(0, 10))
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            bottom, text="",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["text_dim"],
        )
        self.status_label.pack(fill="x", pady=(0, 8))

        self.save_btn = ctk.CTkButton(
            bottom,
            text="💾  저장",
            height=50,
            corner_radius=12,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=self.ACCENT,
            hover_color=self.ACCENT_H,
            command=self._start_save,
        )
        self.save_btn.pack(fill="x")

    # ── PDF 열기 ─────────────────────────────────────────────────────────────
    def _browse_pdf(self):
        if self.is_saving:
            return
        fp = filedialog.askopenfilename(
            title="PDF 파일 선택",
            filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")],
        )
        if fp:
            self._load_pdf(fp)

    def _on_drop_pdf(self, event):
        if self.is_saving:
            return
        raw = event.data
        if raw.startswith("{"):
            files = re.findall(r'\{(.+?)\}', raw)
        else:
            files = raw.split()
        pdfs = [f for f in files if f.lower().endswith(".pdf")]
        if pdfs:
            self._load_pdf(pdfs[0])

    def _load_pdf(self, fp: str):
        fp = os.path.normpath(fp)
        if self.doc:
            self.doc.close()
        self.pdf_path = fp
        self.doc = fitz.open(fp)
        self.page_count = self.doc.page_count
        self.replacements.clear()
        self._update_ctrl_bar()
        self.status_label.configure(
            text=f"📄  {Path(fp).name}  ({self.page_count}페이지)",
            text_color=COLORS["text_dim"],
        )
        # 썸네일은 스레드로 생성
        self._clear_grid()
        threading.Thread(target=self._load_thumbnails, daemon=True).start()

    def _clear_grid(self):
        for w in self.grid_scroll.winfo_children():
            w.destroy()
        self.thumb_images.clear()

    # ── 썸네일 로드 (스레드) ────────────────────────────────────────────────
    def _load_thumbnails(self):
        if not self.doc:
            return
        COLS = 4

        # 행 프레임을 미리 생성해 두되 메인 스레드에서 실행
        def _build_grid_frames():
            rows_needed = (self.page_count + COLS - 1) // COLS
            self._row_frames = []
            for r in range(rows_needed):
                rf = ctk.CTkFrame(self.grid_scroll, fg_color="transparent")
                rf.pack(fill="x", pady=4, padx=4)
                self._row_frames.append(rf)

        self.grid_scroll.after(0, _build_grid_frames)
        self.grid_scroll.after(100, self._render_thumbs_batch, 0, COLS)

    def _render_thumbs_batch(self, start_page: int, cols: int):
        """썸네일을 배치 단위로 렌더링 (GUI 응답성 유지)"""
        BATCH = 8
        end_page = min(start_page + BATCH, self.page_count)

        for pi in range(start_page, end_page):
            self._render_one_thumb(pi, cols)

        self.progress_bar.set(end_page / max(self.page_count, 1))

        if end_page < self.page_count:
            self.grid_scroll.after(20, self._render_thumbs_batch, end_page, cols)
        else:
            self.progress_bar.set(0)

    def _render_one_thumb(self, pi: int, cols: int):
        if not self.doc or pi >= self.page_count:
            return

        row_idx = pi // cols
        if row_idx >= len(self._row_frames):
            return
        parent = self._row_frames[row_idx]

        # 미리보기 이미지 생성
        page = self.doc[pi]
        zoom = min(self.THUMB_W / page.rect.width, self.THUMB_H / page.rect.height)
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        img_pil = Image.open(io.BytesIO(pix.tobytes("png")))
        photo = ImageTk.PhotoImage(img_pil)
        self.thumb_images.append(photo)

        # 카드 프레임
        card = ctk.CTkFrame(
            parent,
            fg_color=COLORS["bg_hover"],
            corner_radius=10,
            border_color=COLORS["border"],
            border_width=1,
            width=self.THUMB_W + 20,
        )
        card.pack(side="left", padx=4)
        card.pack_propagate(False)

        # 썸네일 캔버스
        canvas = tk.Canvas(
            card,
            width=self.THUMB_W,
            height=self.THUMB_H,
            bg=COLORS["bg_hover"],
            highlightthickness=0,
            cursor="hand2",
        )
        canvas.pack(padx=4, pady=(6, 2))
        canvas.create_image(self.THUMB_W // 2, self.THUMB_H // 2,
                            anchor="center", image=photo)
        canvas.bind("<Button-1>", lambda e, p=pi: self._select_replacement(p))

        # 페이지 번호 라벨
        page_lbl = ctk.CTkLabel(
            card,
            text=f"P.{pi+1}",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=COLORS["text_dim"],
        )
        page_lbl.pack(pady=(0, 2))

        # 교체 상태 라벨 + 해제 버튼 컨테이너
        info_row = ctk.CTkFrame(card, fg_color="transparent")
        info_row.pack(fill="x", padx=4, pady=(0, 6))

        replace_lbl = ctk.CTkLabel(
            info_row,
            text="─",
            font=ctk.CTkFont(size=9),
            text_color=COLORS["text_dim"],
            wraplength=self.THUMB_W,
        )
        replace_lbl.pack(side="left", fill="x", expand=True)

        clear_btn = ctk.CTkButton(
            info_row, text="🔄",
            width=24, height=20,
            corner_radius=4,
            font=ctk.CTkFont(size=10),
            fg_color="transparent",
            hover_color=COLORS["danger"],
            text_color=COLORS["text_dim"],
            command=lambda p=pi: self._remove_replacement(p),
        )
        clear_btn.pack(side="right")
        clear_btn.pack_forget()   # 초기에는 숨김

        # 카드 참조 저장 (교체 지정 시 UI 업데이트용)
        if not hasattr(self, '_card_widgets'):
            self._card_widgets = {}
        self._card_widgets[pi] = {
            "card": card,
            "canvas": canvas,
            "page_lbl": page_lbl,
            "replace_lbl": replace_lbl,
            "clear_btn": clear_btn,
        }

    # ── 교체 이미지 선택 ────────────────────────────────────────────────────
    def _select_replacement(self, page_idx: int):
        if self.is_saving:
            return
        fp = filedialog.askopenfilename(
            title=f"P.{page_idx+1} 대체 이미지 선택",
            filetypes=[
                ("이미지 파일", "*.png *.jpg *.jpeg *.bmp *.webp *.tiff *.tif"),
                ("모든 파일", "*.*"),
            ],
        )
        if not fp:
            return
        fp = os.path.normpath(fp)
        self.replacements[page_idx] = fp
        self._update_card_ui(page_idx)
        self._update_ctrl_bar()

    def _remove_replacement(self, page_idx: int):
        self.replacements.pop(page_idx, None)
        self._update_card_ui(page_idx)
        self._update_ctrl_bar()

    def _clear_replacements(self):
        idxs = list(self.replacements.keys())
        self.replacements.clear()
        for pi in idxs:
            self._update_card_ui(pi)
        self._update_ctrl_bar()

    def _update_card_ui(self, page_idx: int):
        """카드의 교체 표시 상태를 갱신"""
        if not hasattr(self, '_card_widgets') or page_idx not in self._card_widgets:
            return
        widgets = self._card_widgets[page_idx]
        card: ctk.CTkFrame = widgets["card"]
        canvas: tk.Canvas = widgets["canvas"]
        replace_lbl: ctk.CTkLabel = widgets["replace_lbl"]
        clear_btn: ctk.CTkButton = widgets["clear_btn"]

        if page_idx in self.replacements:
            img_name = Path(self.replacements[page_idx]).name
            replace_lbl.configure(
                text=img_name,
                text_color=self.ACCENT,
            )
            card.configure(border_color=self.ACCENT)
            canvas.configure(bg="#3d1c1c")
            clear_btn.pack(side="right")
        else:
            replace_lbl.configure(text="─", text_color=COLORS["text_dim"])
            card.configure(border_color=COLORS["border"])
            canvas.configure(bg=COLORS["bg_hover"])
            clear_btn.pack_forget()

    def _update_ctrl_bar(self):
        cnt = len(self.replacements)
        self.replace_count_label.configure(text=f"교체: {cnt}페이지")
        if self.pdf_path:
            self.pdf_name_label.configure(
                text=f"📄 {Path(self.pdf_path).name}  ({self.page_count}p)",
                text_color=COLORS["text"],
            )

    # ── 저장 ────────────────────────────────────────────────────────────────
    def _start_save(self):
        if self.is_saving or not self.pdf_path:
            return
        if not self.replacements:
            self.status_label.configure(
                text="⚠️  교체할 페이지를 먼저 선택하세요",
                text_color=COLORS["warning"],
            )
            return

        # 저장 경로 선택
        stem = Path(self.pdf_path).stem
        out_path = filedialog.asksaveasfilename(
            title="저장할 PDF 경로 선택",
            initialdir=str(Path(self.pdf_path).parent),
            initialfile=f"{stem}_replaced.pdf",
            defaultextension=".pdf",
            filetypes=[("PDF 파일", "*.pdf")],
        )
        if not out_path:
            return

        self.is_saving = True
        self.save_btn.configure(state="disabled", text="⏳  저장 중...",
                                 fg_color=COLORS["bg_hover"])
        threading.Thread(target=self._do_save, args=(out_path,), daemon=True).start()

    def _do_save(self, out_path: str):
        try:
            total = self.page_count
            # 소스 문서를 새 문서로 복사 후 페이지 교체
            src = fitz.open(self.pdf_path)
            dst = fitz.open()

            for pi in range(total):
                self.frame.after(0, self._set_progress,
                                 f"처리 중... ({pi+1}/{total})",
                                 (pi + 1) / total)

                if pi in self.replacements:
                    img_path = self.replacements[pi]
                    # 원본 페이지 크기 참조
                    src_page = src[pi]
                    rect = src_page.rect

                    # 새 빈 페이지 추가
                    new_page = dst.new_page(width=rect.width, height=rect.height)

                    # 이미지를 페이지 전체에 삽입
                    img_rect = fitz.Rect(0, 0, rect.width, rect.height)
                    new_page.insert_image(img_rect, filename=img_path)
                else:
                    # 원본 페이지 그대로 복사
                    dst.insert_pdf(src, from_page=pi, to_page=pi)

            dst.save(out_path, garbage=4, deflate=True)
            dst.close()
            src.close()

            self.frame.after(0, self._on_save_done, out_path)
        except Exception as e:
            self.frame.after(0, self._on_save_error, str(e))

    def _set_progress(self, text: str, val: float):
        self.status_label.configure(text=text, text_color=COLORS["text_dim"])
        self.progress_bar.set(val)

    def _on_save_done(self, out_path: str):
        self.is_saving = False
        self.progress_bar.set(1.0)
        self.save_btn.configure(state="normal", text="💾  저장",
                                 fg_color=self.ACCENT)
        self.status_label.configure(
            text=f"✅  저장 완료!  {Path(out_path).name}",
            text_color=COLORS["success"],
        )
        self._show_done_popup(out_path)

    def _on_save_error(self, err: str):
        self.is_saving = False
        self.progress_bar.set(0)
        self.save_btn.configure(state="normal", text="💾  저장",
                                 fg_color=self.ACCENT)
        self.status_label.configure(
            text=f"❌  오류: {err}",
            text_color=COLORS["danger"],
        )

    def _show_done_popup(self, out_path: str):
        popup = ctk.CTkToplevel(self.frame)
        popup.title("저장 완료")
        popup.geometry("420x260")
        popup.configure(fg_color=COLORS["bg_dark"])
        popup.transient(self.frame.winfo_toplevel())
        popup.grab_set()

        ctk.CTkLabel(
            popup, text="✅  저장 완료",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=COLORS["success"],
        ).pack(pady=(24, 4))

        ctk.CTkLabel(
            popup,
            text=f"{len(self.replacements)}개 페이지가 이미지로 교체되었습니다",
            font=ctk.CTkFont(size=13),
            text_color=COLORS["text_dim"],
        ).pack(pady=(0, 4))

        ctk.CTkLabel(
            popup, text=Path(out_path).name,
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text"],
        ).pack()

        ctk.CTkLabel(
            popup, text=str(Path(out_path).parent),
            font=ctk.CTkFont(size=10),
            text_color=COLORS["text_dim"],
        ).pack(pady=(0, 20))

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=(0, 20))

        ctk.CTkButton(
            btn_frame, text="📂 폴더 열기",
            height=40, corner_radius=10,
            font=ctk.CTkFont(size=13),
            fg_color=self.ACCENT,
            hover_color=self.ACCENT_H,
            command=lambda: _open_folder(out_path),
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))

        ctk.CTkButton(
            btn_frame, text="닫기",
            height=40, corner_radius=10,
            font=ctk.CTkFont(size=13),
            fg_color=COLORS["bg_hover"],
            hover_color=COLORS["border"],
            command=popup.destroy,
        ).pack(side="right", fill="x", expand=True, padx=(6, 0))


# ═══════════════════════════════════════════════════════════════════════════
#  공통 유틸
# ═══════════════════════════════════════════════════════════════════════════
def _open_folder(filepath: str):
    folder = os.path.dirname(filepath)
    os.startfile(folder)


# ═══════════════════════════════════════════════════════════════════════════
#  메인 앱
# ═══════════════════════════════════════════════════════════════════════════
class App:
    def __init__(self):
        self.root = TkinterDnDCTk()
        self.root.title("PDF ↔ Image Converter")
        self.root.geometry("720x980")
        self.root.minsize(600, 820)
        self.root.configure(fg_color=COLORS["bg_dark"])

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self._build_ui()

    def _build_ui(self):
        main = ctk.CTkFrame(self.root, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=(16, 16))

        # ── 헤더 ─────────────────────────────────────────────────────────
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(
            header, text="PDF ↔",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=COLORS["text"],
        ).pack(side="left")

        ctk.CTkLabel(
            header, text=" Image",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=COLORS["accent_glow"],
        ).pack(side="left")

        ctk.CTkLabel(
            header, text=" Converter",
            font=ctk.CTkFont(size=28),
            text_color=COLORS["text_dim"],
        ).pack(side="left")

        # ── 탭 뷰 ────────────────────────────────────────────────────────
        self.tabview = ctk.CTkTabview(
            main,
            fg_color=COLORS["bg_card"],
            segmented_button_fg_color=COLORS["bg_dark"],
            segmented_button_selected_color=COLORS["accent"],
            segmented_button_selected_hover_color=COLORS["accent_hover"],
            segmented_button_unselected_color=COLORS["bg_hover"],
            segmented_button_unselected_hover_color=COLORS["bg_card"],
            text_color=COLORS["text"],
            corner_radius=14,
        )
        self.tabview.pack(fill="both", expand=True)

        self.tabview.add("📄  PDF → PNG")
        self.tabview.add("🖼️  이미지 → PDF")
        self.tabview.add("🔄  페이지 교체")

        # 탭 내부 프레임 가져오기
        pdf_tab_frame   = self.tabview.tab("📄  PDF → PNG")
        img_tab_frame   = self.tabview.tab("🖼️  이미지 → PDF")
        repl_tab_frame  = self.tabview.tab("🔄  페이지 교체")

        # 탭 내부에 스크롤 가능한 프레임 추가
        pdf_scroll_outer = ctk.CTkScrollableFrame(
            pdf_tab_frame, fg_color="transparent",
        )
        pdf_scroll_outer.pack(fill="both", expand=True)

        img_scroll_outer = ctk.CTkScrollableFrame(
            img_tab_frame, fg_color="transparent",
        )
        img_scroll_outer.pack(fill="both", expand=True)

        repl_scroll_outer = ctk.CTkScrollableFrame(
            repl_tab_frame, fg_color="transparent",
        )
        repl_scroll_outer.pack(fill="both", expand=True)

        # 각 탭 인스턴스 생성
        self.pdf_tab  = PdfToPngTab(pdf_scroll_outer)
        self.img_tab  = ImgToPdfTab(img_scroll_outer)
        self.repl_tab = PdfPageReplaceTab(repl_scroll_outer)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = App()
    app.run()
