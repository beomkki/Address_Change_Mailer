#!/usr/bin/env python3
"""주소변경 메일 머지를 위한 간단한 GUI."""

from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

from generate_address_change_mail import run_address_change_mail_merge

BASE_DIR = Path(__file__).resolve().parent


def _initial_dir(current: str) -> str:
    """Get initial directory for file dialog."""
    if current:
        candidate = Path(current)
        if candidate.is_file():
            return str(candidate.parent)
        if candidate.exists():
            return str(candidate)
    return str(BASE_DIR)


def launch_gui() -> None:
    """Launch the address change mail merge GUI."""
    root = tk.Tk()
    root.title("주소변경 메일 머지 도우미")
    root.resizable(False, False)

    frame = ttk.Frame(root, padding=16)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(1, weight=1)

    # Variables for file paths
    marks_var = tk.StringVar(value=str(BASE_DIR / "List of Marks.xlsx"))
    mailing_list_var = tk.StringVar(value=str(BASE_DIR / "메일링 리스트.xlsx"))
    template_var = tk.StringVar(value=str(BASE_DIR / "Address_Change_Mail_Sample.docx"))
    output_default = BASE_DIR / "output-address-change"
    output_var = tk.StringVar(value=str(output_default))

    def browse_marks() -> None:
        """Browse for marks Excel file."""
        selected = filedialog.askopenfilename(
            title="상표 리스트 파일 선택",
            initialdir=_initial_dir(marks_var.get()),
            filetypes=(("Excel 통합 문서", "*.xlsx"), ("모든 파일", "*.*")),
        )
        if selected:
            marks_var.set(selected)

    def browse_mailing_list() -> None:
        """Browse for mailing list Excel file."""
        selected = filedialog.askopenfilename(
            title="메일링 리스트 파일 선택",
            initialdir=_initial_dir(mailing_list_var.get()),
            filetypes=(("Excel 통합 문서", "*.xlsx"), ("모든 파일", "*.*")),
        )
        if selected:
            mailing_list_var.set(selected)

    def browse_template() -> None:
        """Browse for Word template file."""
        selected = filedialog.askopenfilename(
            title="메일 템플릿 파일 선택",
            initialdir=_initial_dir(template_var.get()),
            filetypes=(("Word 문서", "*.docx"), ("모든 파일", "*.*")),
        )
        if selected:
            template_var.set(selected)

    def browse_output() -> None:
        """Browse for output directory."""
        selected = filedialog.askdirectory(
            title="MSG 저장 폴더 선택",
            initialdir=_initial_dir(output_var.get()),
        )
        if selected:
            output_var.set(selected)

    def run_merge() -> None:
        """Run the address change mail merge."""
        marks_path = marks_var.get().strip()
        mailing_list_path = mailing_list_var.get().strip()
        template_path = template_var.get().strip()
        output_dir = output_var.get().strip()

        # Marks, template, and output dir are required
        if not marks_path or not template_path or not output_dir:
            messagebox.showerror("입력 오류", "상표 리스트, 템플릿, 출력 폴더는 필수입니다.")
            return

        # Validate required files exist
        if not Path(marks_path).exists():
            messagebox.showerror("파일 오류", f"상표 리스트 파일을 찾을 수 없습니다:\n{marks_path}")
            return
        if not Path(template_path).exists():
            messagebox.showerror("파일 오류", f"템플릿 파일을 찾을 수 없습니다:\n{template_path}")
            return

        # Mailing list is optional - warn if file doesn't exist
        if mailing_list_path and not Path(mailing_list_path).exists():
            response = messagebox.askyesno(
                "메일링 리스트 없음",
                f"메일링 리스트 파일을 찾을 수 없습니다:\n{mailing_list_path}\n\n"
                "상표 리스트의 수신인 정보만 사용하여 계속하시겠습니까?"
            )
            if not response:
                return

        try:
            generated = run_address_change_mail_merge(
                marks_excel=marks_path,
                mailing_list_excel=mailing_list_path,
                template_path=template_path,
                output_dir=output_dir,
            )
        except SystemExit as exc:
            messagebox.showerror("실행 실패", str(exc))
        except Exception as exc:  # pragma: no cover - GUI 사용 시 디버깅 보조
            messagebox.showerror("예기치 못한 오류", str(exc))
        else:
            messagebox.showinfo("완료", f"{generated}건의 MSG 파일을 생성했습니다.\n\n저장 위치: {output_dir}")

    # GUI Layout
    row = 0

    # Title
    title_label = ttk.Label(frame, text="주소변경 메일 머지", font=("", 14, "bold"))
    title_label.grid(row=row, column=0, columnspan=3, pady=(0, 16))
    row += 1

    # Description
    desc_label = ttk.Label(
        frame,
        text="국가별로 상표를 그룹핑하여 주소변경 안내 메일을 생성합니다.",
        foreground="gray",
    )
    desc_label.grid(row=row, column=0, columnspan=3, pady=(0, 16))
    row += 1

    # Marks Excel file
    ttk.Label(frame, text="상표 리스트 파일").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
    ttk.Entry(frame, textvariable=marks_var, width=50).grid(row=row, column=1, sticky="ew", pady=(0, 8))
    ttk.Button(frame, text="찾기", command=browse_marks).grid(row=row, column=2, padx=(8, 0), pady=(0, 8))
    row += 1

    # Mailing list file
    ttk.Label(frame, text="메일링 리스트 파일").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
    ttk.Entry(frame, textvariable=mailing_list_var, width=50).grid(row=row, column=1, sticky="ew", pady=(0, 8))
    ttk.Button(frame, text="찾기", command=browse_mailing_list).grid(row=row, column=2, padx=(8, 0), pady=(0, 8))
    row += 1

    # Template file
    ttk.Label(frame, text="메일 템플릿 파일").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
    ttk.Entry(frame, textvariable=template_var, width=50).grid(row=row, column=1, sticky="ew", pady=(0, 8))
    ttk.Button(frame, text="찾기", command=browse_template).grid(row=row, column=2, padx=(8, 0), pady=(0, 8))
    row += 1

    # Output directory
    ttk.Label(frame, text="이메일 저장 폴더").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
    ttk.Entry(frame, textvariable=output_var, width=50).grid(row=row, column=1, sticky="ew", pady=(0, 8))
    ttk.Button(frame, text="찾기", command=browse_output).grid(row=row, column=2, padx=(8, 0), pady=(0, 8))
    row += 1

    # Info text
    info_text = ttk.Label(
        frame,
        text="💡 수신인 정보는 상표 리스트 또는 메일링 리스트에서 자동으로 조회됩니다.",
        foreground="blue",
        font=("", 9),
    )
    info_text.grid(row=row, column=0, columnspan=3, pady=(12, 8))
    row += 1

    # Run button
    run_button = ttk.Button(frame, text="메일 생성", command=run_merge)
    run_button.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(8, 0))

    root.mainloop()


if __name__ == "__main__":
    launch_gui()
