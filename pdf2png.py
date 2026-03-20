"""
PDF to PNG Converter
모던 GUI 기반 PDF → PNG 변환 프로그램
"""

import os
import io
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
    "text":         "#e8e8f0",
    "text_dim":     "#8888a0",
    "success":      "#00cec9",
    "warning":      "#fdcb6e",
    "danger":       "#ff6b6b",
    "border":       "#2d2d44",
}


class PDFtoPNGApp:
    def __init__(self):
        self.root = TkinterDnDCTk()
        self.root.title("PDF → PNG Converter")
        self.root.geometry("720x960")
        self.root.minsize(600, 800)
        self.root.configure(fg_color=COLORS["bg_dark"])

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # 상태 변수
        self.pdf_files: list[str] = []
        self.dpi_var = ctk.StringVar(value="300")
        self.merge_var = ctk.BooleanVar(value=True)
        self.is_converting = False

        self._build_ui()

    # ── UI 빌드 ──────────────────────────────────────────────────────────
    def _build_ui(self):
        # 메인 스크롤 프레임
        main = ctk.CTkFrame(self.root, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=(16, 16))

        # ── 헤더 ─────────────────────────────────────────────────────────
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            header, text="📄  PDF → PNG",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=COLORS["text"],
        ).pack(side="left")

        ctk.CTkLabel(
            header, text="Converter",
            font=ctk.CTkFont(size=28),
            text_color=COLORS["accent_glow"],
        ).pack(side="left", padx=(8, 0))

        # ── 드롭 영역 ───────────────────────────────────────────────────
        self.drop_frame = ctk.CTkFrame(
            main,
            fg_color=COLORS["bg_card"],
            border_color=COLORS["border"],
            border_width=2,
            corner_radius=16,
            height=110,
        )
        self.drop_frame.pack(fill="x", pady=(0, 10))
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
        # 드래그 앤 드롭 경로 파싱 (중괄호 및 공백 처리)
        files = []
        if raw.startswith("{"):
            # {path with spaces} 형식
            import re
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

            # 파일명
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

            # 삭제 버튼
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
                self.root.after(0, self._update_status,
                    f"변환 중... ({idx+1}/{total_files}) {Path(pdf_path).name}",
                    idx / total_files)

                doc = fitz.open(pdf_path)
                page_count = doc.page_count
                images = []

                for pi, page in enumerate(doc):
                    pix = page.get_pixmap(dpi=dpi)
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    images.append(img)

                    # 파일 내 페이지 진행률 반영
                    progress = (idx + (pi + 1) / page_count) / total_files
                    self.root.after(0, self._update_progress, progress)

                doc.close()

                base = Path(pdf_path)
                stem = base.stem
                out_dir = base.parent

                if merge or page_count == 1:
                    # 1장으로 합치기
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
                    # 페이지별 분리
                    for pi, img in enumerate(images):
                        out_path = str(out_dir / f"{stem}_{pi+1}.png")
                        img.save(out_path)
                        converted.append(out_path)

            except Exception as e:
                errors.append((Path(pdf_path).name, str(e)))

        # 완료
        self.root.after(0, self._on_convert_done, converted, errors)

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
            error_text = "\n".join(f"  • {name}: {err}" for name, err in errors)
            self.status_label.configure(
                text=f"⚠️  {len(errors)}개 파일 오류 발생",
                text_color=COLORS["danger"],
            )
        else:
            self.status_label.configure(
                text=f"✅  변환 완료! {len(converted)}개 PNG 파일 생성",
                text_color=COLORS["success"],
            )

        # 결과 파일 폴더 열기 버튼 표시
        if converted:
            # 결과 팝업
            self._show_result_popup(converted, errors)

    def _show_result_popup(self, converted: list[str], errors: list[tuple[str, str]]):
        popup = ctk.CTkToplevel(self.root)
        popup.title("변환 결과")
        popup.geometry("500x400")
        popup.configure(fg_color=COLORS["bg_dark"])
        popup.transient(self.root)
        popup.grab_set()

        # 헤더
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

        # 파일 목록
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
                command=lambda f=fp: self._open_folder(f),
            ).pack(side="right")

        # 에러 표시
        if errors:
            ctk.CTkLabel(
                popup,
                text=f"⚠️  {len(errors)}개 파일 오류",
                font=ctk.CTkFont(size=12),
                text_color=COLORS["danger"],
            ).pack(pady=(0, 4))

        # 버튼
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(fill="x", padx=16, pady=(0, 16))

        if converted:
            # 첫 번째 파일의 폴더 열기
            ctk.CTkButton(
                btn_frame, text="📂 폴더 열기",
                height=40, corner_radius=10,
                font=ctk.CTkFont(size=13),
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                command=lambda: self._open_folder(converted[0]),
            ).pack(side="left", fill="x", expand=True, padx=(0, 4))

        ctk.CTkButton(
            btn_frame, text="닫기",
            height=40, corner_radius=10,
            font=ctk.CTkFont(size=13),
            fg_color=COLORS["bg_hover"],
            hover_color=COLORS["border"],
            command=popup.destroy,
        ).pack(side="right", fill="x", expand=True, padx=(4, 0))

    def _open_folder(self, filepath: str):
        folder = os.path.dirname(filepath)
        os.startfile(folder)

    # ── 실행 ─────────────────────────────────────────────────────────────
    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = PDFtoPNGApp()
    app.run()
