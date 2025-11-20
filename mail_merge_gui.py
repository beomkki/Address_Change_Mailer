#!/usr/bin/env python3
"""메일 머지 실행을 위한 간단한 GUI."""

from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

from generate_mail_merge import run_mail_merge

BASE_DIR = Path(__file__).resolve().parent

MERGE_PROFILES: dict[str, dict[str, object]] = {
    "filing": {
        "label": "출원",
        "template": BASE_DIR / "Filing.docx",
        "excel": BASE_DIR / "Filing_Merge.xlsx",
        "subject": "(«관리번호»«국가코드») New trademark application(s) in «국가명칭»",
    },
    "search": {
        "label": "검색",
        "template": BASE_DIR / "Search.docx",
        "excel": BASE_DIR / "Search_Merge.xlsx",
        "subject": "(«관리번호»«국가코드»-S) Trademark search(es) in «국가명칭»",
    },
}


def _initial_dir(current: str) -> str:
    if current:
        candidate = Path(current)
        if candidate.is_file():
            return str(candidate.parent)
        if candidate.exists():
            return str(candidate)
    return str(BASE_DIR)


def launch_gui() -> None:
    root = tk.Tk()
    root.title("메일 머지 도우미")
    root.resizable(False, False)

    frame = ttk.Frame(root, padding=16)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(1, weight=1)

    profile_var = tk.StringVar(value="filing")

    template_var = tk.StringVar()
    excel_var = tk.StringVar()
    output_default = BASE_DIR / "output-msg"
    output_var = tk.StringVar(value=str(output_default))
    subject_var = tk.StringVar()
    attachment_field_var = tk.StringVar(value="첨부파일")

    def apply_profile(profile: str) -> None:
        data = MERGE_PROFILES.get(profile)
        if not data:
            return
        template_var.set(str(data["template"]))
        excel_var.set(str(data["excel"]))
        subject_var.set(str(data["subject"]))

    def on_profile_change() -> None:
        apply_profile(profile_var.get())

    apply_profile(profile_var.get())

    def browse_template() -> None:
        selected = filedialog.askopenfilename(
            title="샘플 워드 파일 선택",
            initialdir=_initial_dir(template_var.get()),
            filetypes=(("Word 문서", "*.docx"), ("모든 파일", "*.*")),
        )
        if selected:
            template_var.set(selected)

    def browse_excel() -> None:
        selected = filedialog.askopenfilename(
            title="메일머지용 엑셀 선택",
            initialdir=_initial_dir(excel_var.get()),
            filetypes=(("Excel 통합 문서", "*.xlsx"), ("모든 파일", "*.*")),
        )
        if selected:
            excel_var.set(selected)

    def browse_output() -> None:
        selected = filedialog.askdirectory(
            title="MSG 저장 폴더 선택",
            initialdir=_initial_dir(output_var.get()),
        )
        if selected:
            output_var.set(selected)

    def run_merge() -> None:
        excel_path = excel_var.get().strip()
        template_path = template_var.get().strip()
        output_dir = output_var.get().strip()
        subject_template = subject_var.get().strip()
        attachment_field = attachment_field_var.get().strip()

        if not excel_path or not template_path or not output_dir:
            messagebox.showerror("입력 오류", "모든 경로를 선택해 주세요.")
            return

        # 제목 템플릿이 비어있으면 현재 프로필의 기본값 사용
        if not subject_template:
            current_profile = profile_var.get()
            profile_data = MERGE_PROFILES.get(current_profile, {})
            subject_template = str(profile_data.get("subject", ""))

        try:
            generated = run_mail_merge(
                excel_path=excel_path,
                template_path=template_path,
                output_dir=output_dir,
                subject_template=subject_template,
                attachment_field=attachment_field,
            )
        except SystemExit as exc:
            messagebox.showerror("실행 실패", str(exc))
        except Exception as exc:  # pragma: no cover - GUI 사용 시 디버깅 보조
            messagebox.showerror("예기치 못한 오류", str(exc))
        else:
            messagebox.showinfo("완료", f"{generated}건의 MSG 파일을 생성했습니다.")

    ttk.Label(frame, text="메일 유형").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
    profile_frame = ttk.Frame(frame)
    profile_frame.grid(row=0, column=1, columnspan=2, sticky="w", pady=(0, 8))
    for key, data in MERGE_PROFILES.items():
        ttk.Radiobutton(
            profile_frame,
            text=str(data["label"]),
            value=key,
            variable=profile_var,
            command=on_profile_change,
        ).pack(side="left", padx=(0, 12))

    ttk.Label(frame, text="샘플 워드 파일").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
    ttk.Entry(frame, textvariable=template_var, width=50).grid(row=1, column=1, sticky="ew", pady=(0, 8))
    ttk.Button(frame, text="찾기", command=browse_template).grid(row=1, column=2, padx=(8, 0), pady=(0, 8))

    ttk.Label(frame, text="메일머지 엑셀 파일").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
    ttk.Entry(frame, textvariable=excel_var, width=50).grid(row=2, column=1, sticky="ew", pady=(0, 8))
    ttk.Button(frame, text="찾기", command=browse_excel).grid(row=2, column=2, padx=(8, 0), pady=(0, 8))

    ttk.Label(frame, text="이메일 저장 폴더").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
    ttk.Entry(frame, textvariable=output_var, width=50).grid(row=3, column=1, sticky="ew", pady=(0, 8))
    ttk.Button(frame, text="찾기", command=browse_output).grid(row=3, column=2, padx=(8, 0), pady=(0, 8))

    ttk.Label(frame, text="제목 템플릿").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
    ttk.Entry(frame, textvariable=subject_var, width=50).grid(row=4, column=1, sticky="ew", pady=(0, 8))
    ttk.Label(frame, text="(비워두면 기본 템플릿 사용)").grid(row=4, column=2, sticky="w", pady=(0, 8))

    ttk.Label(frame, text="첨부파일 필드명").grid(row=5, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
    ttk.Entry(frame, textvariable=attachment_field_var, width=50).grid(row=5, column=1, sticky="ew", pady=(0, 8))
    ttk.Label(frame, text="(엑셀 컬럼명, 세미콜론으로 구분)").grid(row=5, column=2, sticky="w", pady=(0, 8))

    ttk.Button(frame, text="메일 생성", command=run_merge).grid(row=6, column=0, columnspan=3, sticky="ew")

    root.mainloop()


if __name__ == "__main__":
    launch_gui()
