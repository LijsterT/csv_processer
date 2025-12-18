# coding: utf-8
"""Personalized CSV Creator GUI application.

This tool allows users to convert Excel files into CSV with customized
formatting options such as custom separators, quoting strategies, encoding,
and line endings. It provides a live preview of the output and persists the
last used configuration.
"""

import datetime as dt
import json
import os
import threading
from dataclasses import dataclass, asdict, field
from itertools import islice
from numbers import Number
from pathlib import Path
from typing import Any, Callable, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

CONFIG_PATH = Path.home() / ".personalized_csv_creator_config.json"
PREVIEW_ROW_LIMIT = 20
SUPPORTED_ENCODINGS = {
    "UTF-8": "utf-8",
    "UTF-8 with BOM": "utf-8-sig",
    "ISO-8859-1": "iso-8859-1",
    "Windows-1252": "cp1252",
}
LINE_ENDING_OPTIONS = {
    "Auto (OS default)": None,
    "Unix (\\n)": "\n",
    "Windows (\\r\\n)": "\r\n",
}
DEFAULT_SEPARATOR = ","


@dataclass
class AppConfig:
    """Serializable configuration for persisting user preferences."""

    last_file: str = ""
    separator_choice: str = ","
    custom_separator: str = ""
    quoting_mode: str = "text"
    quote_choice: str = '"'
    custom_quote: str = ""
    encoding: str = "UTF-8"
    line_ending: str = "Auto (OS default)"
    sheet_name: str = ""
    significant_figures: dict[str, int] = field(default_factory=dict)


class CSVConverterApp:
    """Tkinter application implementing the Personalized CSV Creator."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Personalized CSV Creator")
        self.root.geometry("900x650")
        self.root.minsize(850, 600)

        self.config = self.load_config()
        self.significant_figures: dict[str, int] = dict(self.config.significant_figures)
        self.excel_path: Optional[Path] = Path(self.config.last_file) if self.config.last_file else None
        self.sheet_names: List[str] = []
        self.preview_data: Optional[pd.DataFrame] = None
        self.conversion_thread: Optional[threading.Thread] = None

        self.create_widgets()
        if self.excel_path and self.excel_path.exists():
            self.load_excel_file(self.excel_path)
            self.file_var.set(str(self.excel_path))
            if self.config.sheet_name and self.config.sheet_name in self.sheet_names:
                self.sheet_var.set(self.config.sheet_name)
        else:
            self.excel_path = None

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ------------------------------------------------------------------
    # Configuration
    def load_config(self) -> AppConfig:
        if CONFIG_PATH.exists():
            try:
                with CONFIG_PATH.open("r", encoding="utf-8") as fh:
                    data = json.load(fh)
                return AppConfig(**data)
            except Exception:
                return AppConfig()
        return AppConfig()

    def save_config(self) -> None:
        try:
            with CONFIG_PATH.open("w", encoding="utf-8") as fh:
                json.dump(asdict(self.config), fh, indent=2)
        except Exception:
            # Saving configuration is best-effort; ignore failures.
            pass

    # ------------------------------------------------------------------
    # UI construction
    def create_widgets(self) -> None:
        padding = {"padx": 10, "pady": 5}

        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # File selection
        file_frame = ttk.LabelFrame(main_frame, text="Excel File")
        file_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(file_frame, text="Browse…", command=self.browse_file).pack(side=tk.LEFT, padx=5, pady=5)
        self.file_var = tk.StringVar(value=str(self.excel_path) if self.excel_path else "")
        ttk.Entry(file_frame, textvariable=self.file_var, state="readonly", width=80).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # Sheet selector
        sheet_frame = ttk.Frame(main_frame)
        sheet_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(sheet_frame, text="Sheet:").pack(side=tk.LEFT)
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.sheet_var, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _: self.refresh_preview())

        # Significant figures controls
        sig_fig_frame = ttk.LabelFrame(main_frame, text="Numeric formatting")
        sig_fig_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(
            sig_fig_frame,
            text="Select a column and set significant figures (optional for numeric columns):",
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, **padding)

        self.column_listbox = tk.Listbox(sig_fig_frame, height=6, exportselection=False)
        self.column_listbox.grid(row=1, column=0, rowspan=3, sticky=tk.NSEW, **padding)
        self.column_listbox.bind("<<ListboxSelect>>", self.on_column_select)

        self.sig_fig_var = tk.StringVar()
        ttk.Label(sig_fig_frame, text="Significant figures:").grid(row=1, column=1, sticky=tk.W, **padding)
        self.sig_fig_entry = ttk.Entry(sig_fig_frame, textvariable=self.sig_fig_var, width=10)
        self.sig_fig_entry.grid(row=1, column=2, sticky=tk.W, **padding)

        ttk.Button(sig_fig_frame, text="Apply", command=self.apply_sig_fig).grid(row=2, column=1, sticky=tk.W, **padding)
        ttk.Button(sig_fig_frame, text="Clear", command=self.clear_sig_fig).grid(row=2, column=2, sticky=tk.W, **padding)

        self.sig_fig_status = tk.StringVar(value="No column selected")
        ttk.Label(sig_fig_frame, textvariable=self.sig_fig_status).grid(row=3, column=1, columnspan=2, sticky=tk.W, **padding)

        sig_fig_frame.columnconfigure(0, weight=1)

        # Separator selection
        separator_frame = ttk.LabelFrame(main_frame, text="Separator (Delimiter)")
        separator_frame.pack(fill=tk.X, padx=10, pady=5)

        separator_options = [",", ";", "\\t", "|", ":", "Custom…"]
        separator_default = self.config.separator_choice if self.config.separator_choice in separator_options else ","
        self.separator_var = tk.StringVar(value=separator_default)
        ttk.Label(separator_frame, text="Choose:").grid(row=0, column=0, sticky=tk.W, **padding)
        self.separator_combo = ttk.Combobox(separator_frame, values=separator_options, textvariable=self.separator_var, state="readonly")
        self.separator_combo.grid(row=0, column=1, sticky=tk.W, **padding)
        self.separator_combo.bind("<<ComboboxSelected>>", self.on_separator_change)

        ttk.Label(separator_frame, text="Custom:").grid(row=0, column=2, sticky=tk.W, **padding)
        self.custom_separator_var = tk.StringVar(value=self.config.custom_separator)
        self.custom_separator_entry = ttk.Entry(separator_frame, textvariable=self.custom_separator_var, width=10)
        self.custom_separator_entry.grid(row=0, column=3, sticky=tk.W, **padding)
        self.custom_separator_entry.bind("<KeyRelease>", lambda _: self.refresh_preview())

        # Quoting options
        quoting_frame = ttk.LabelFrame(main_frame, text="Quoting")
        quoting_frame.pack(fill=tk.X, padx=10, pady=5)

        self.quoting_mode_var = tk.StringVar(value=self.config.quoting_mode)
        ttk.Radiobutton(quoting_frame, text="No quoting", value="none", variable=self.quoting_mode_var, command=self.refresh_preview).grid(row=0, column=0, sticky=tk.W, **padding)
        ttk.Radiobutton(quoting_frame, text="Quote text fields", value="text", variable=self.quoting_mode_var, command=self.refresh_preview).grid(row=0, column=1, sticky=tk.W, **padding)
        ttk.Radiobutton(quoting_frame, text="Quote all fields", value="all", variable=self.quoting_mode_var, command=self.refresh_preview).grid(row=0, column=2, sticky=tk.W, **padding)

        quote_char_frame = ttk.Frame(quoting_frame)
        quote_char_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W)
        ttk.Label(quote_char_frame, text="Quote character:").pack(side=tk.LEFT, padx=10, pady=5)

        quote_options = ['"', "'", "`", "Custom…"]
        quote_default = self.config.quote_choice if self.config.quote_choice in quote_options else '"'
        self.quote_choice_var = tk.StringVar(value=quote_default)
        self.quote_combo = ttk.Combobox(quote_char_frame, values=quote_options, textvariable=self.quote_choice_var, width=10, state="readonly")
        self.quote_combo.pack(side=tk.LEFT, padx=5)
        self.quote_combo.bind("<<ComboboxSelected>>", self.on_quote_change)

        ttk.Label(quote_char_frame, text="Custom:").pack(side=tk.LEFT, padx=5)
        self.custom_quote_var = tk.StringVar(value=self.config.custom_quote)
        self.custom_quote_entry = ttk.Entry(quote_char_frame, textvariable=self.custom_quote_var, width=5)
        self.custom_quote_entry.pack(side=tk.LEFT, padx=5)
        self.custom_quote_entry.bind("<KeyRelease>", lambda _: self.refresh_preview())

        # Encoding and line endings
        options_frame = ttk.LabelFrame(main_frame, text="Output Options")
        options_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(options_frame, text="Encoding:").grid(row=0, column=0, sticky=tk.W, **padding)
        encoding_default = self.config.encoding if self.config.encoding in SUPPORTED_ENCODINGS else "UTF-8"
        self.encoding_var = tk.StringVar(value=encoding_default)
        ttk.Combobox(options_frame, values=list(SUPPORTED_ENCODINGS.keys()), textvariable=self.encoding_var, state="readonly").grid(row=0, column=1, sticky=tk.W, **padding)

        ttk.Label(options_frame, text="Line endings:").grid(row=0, column=2, sticky=tk.W, **padding)
        line_ending_default = self.config.line_ending if self.config.line_ending in LINE_ENDING_OPTIONS else "Auto (OS default)"
        self.line_ending_var = tk.StringVar(value=line_ending_default)
        ttk.Combobox(options_frame, values=list(LINE_ENDING_OPTIONS.keys()), textvariable=self.line_ending_var, state="readonly").grid(row=0, column=3, sticky=tk.W, **padding)

        # Preview
        preview_frame = ttk.LabelFrame(main_frame, text="Preview (first 20 rows)")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        preview_container = ttk.Frame(preview_frame)
        preview_container.pack(fill=tk.BOTH, expand=True)

        self.preview_text = tk.Text(preview_container, height=15, wrap=tk.NONE, state=tk.DISABLED, font=("Courier New", 10))
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        preview_scroll_y = ttk.Scrollbar(preview_container, orient=tk.VERTICAL, command=self.preview_text.yview)
        preview_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        preview_scroll_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_text.xview)
        preview_scroll_x.pack(fill=tk.X)
        self.preview_text.configure(yscrollcommand=preview_scroll_y.set, xscrollcommand=preview_scroll_x.set)

        # Action buttons
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(action_frame, text="Refresh Preview", command=self.refresh_preview).pack(side=tk.LEFT)
        self.save_button = ttk.Button(action_frame, text="Save CSV…", command=self.save_csv)
        self.save_button.pack(side=tk.RIGHT)

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, anchor=tk.W)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # Initial state adjustments
        self.update_separator_state()
        self.update_quote_state()
        self.refresh_preview()

    # ------------------------------------------------------------------
    # Helper utilities
    def browse_file(self) -> None:
        filetypes = [
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*"),
        ]
        filename = filedialog.askopenfilename(title="Select Excel file", filetypes=filetypes)
        if filename:
            path = Path(filename)
            self.load_excel_file(path)
            self.file_var.set(str(path))
            self.config.last_file = str(path)
            self.refresh_preview()

    def load_excel_file(self, path: Path) -> None:
        try:
            excel = pd.ExcelFile(path)
            self.sheet_names = excel.sheet_names
            self.sheet_combo.configure(values=self.sheet_names)
            default_sheet = self.sheet_names[0] if self.sheet_names else ""
            if self.config.sheet_name in self.sheet_names:
                default_sheet = self.config.sheet_name
            self.sheet_var.set(default_sheet)
            self.excel_path = path
            self.status_var.set(f"Loaded {path.name}")
        except Exception as exc:
            messagebox.showerror("Error", f"Unable to read Excel file: {exc}")
            self.status_var.set("Failed to load file")
            self.excel_path = None
            self.sheet_names = []
            self.sheet_combo.configure(values=[])
            self.sheet_var.set("")

    def on_separator_change(self, _event: tk.Event) -> None:  # type: ignore[override]
        self.update_separator_state()
        self.refresh_preview()

    def update_separator_state(self) -> None:
        if self.separator_var.get() == "Custom…":
            self.custom_separator_entry.configure(state=tk.NORMAL)
        else:
            self.custom_separator_entry.configure(state=tk.DISABLED)
        # No-op; preview refresh happens via event bindings.

    def on_quote_change(self, _event: tk.Event) -> None:  # type: ignore[override]
        self.update_quote_state()
        self.refresh_preview()

    def update_quote_state(self) -> None:
        if self.quote_choice_var.get() == "Custom…":
            self.custom_quote_entry.configure(state=tk.NORMAL)
        else:
            self.custom_quote_entry.configure(state=tk.DISABLED)

    def get_separator(self) -> Optional[str]:
        value = self.separator_var.get()
        if value == "Custom…":
            custom = self.custom_separator_var.get()
            return custom if custom else None
        if value == "\\t":
            return "\t"
        return value

    def get_quote_char(self) -> Optional[str]:
        value = self.quote_choice_var.get()
        if value == "Custom…":
            custom = self.custom_quote_var.get()
            return custom if custom else None
        return value

    def validate_settings(self, show_dialog: bool = True) -> bool:
        separator = self.get_separator()
        if not separator:
            if show_dialog:
                messagebox.showerror("Invalid separator", "Please provide a non-empty separator.")
            return False
        quote_char = self.get_quote_char()
        if self.quoting_mode_var.get() != "none":
            if not quote_char:
                if show_dialog:
                    messagebox.showerror("Invalid quote", "Please provide a quote character.")
                return False
        if quote_char and len(quote_char) != 1:
            if show_dialog:
                messagebox.showerror("Invalid quote", "Quote character must be a single character.")
            return False
        return True

    # ------------------------------------------------------------------
    # Preview
    def refresh_preview(self) -> None:
        if not self.excel_path:
            return
        if not self.validate_settings(show_dialog=False):
            return
        try:
            sheet_name = self.sheet_var.get() or None
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name, nrows=PREVIEW_ROW_LIMIT)
            self.config.sheet_name = sheet_name or ""
            self.preview_data = df
            self.update_column_list(df.columns)
            self.display_preview(df)
        except Exception as exc:
            messagebox.showerror("Preview error", f"Unable to generate preview: {exc}")
            self.status_var.set("Preview failed")

    def display_preview(self, df: pd.DataFrame) -> None:
        csv_lines = list(self.iter_csv_lines(df, limit=PREVIEW_ROW_LIMIT))
        content = "\n".join(csv_lines)
        self.preview_text.configure(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert(tk.END, content)
        self.preview_text.configure(state=tk.DISABLED)

    # ------------------------------------------------------------------
    # CSV conversion
    def save_csv(self) -> None:
        if not self.excel_path:
            messagebox.showinfo("Select file", "Please choose an Excel file first.")
            return
        if not self.validate_settings():
            return

        default_name = self.excel_path.with_suffix(".csv")
        filename = filedialog.asksaveasfilename(
            title="Save CSV",
            defaultextension=".csv",
            initialfile=default_name.name,
            initialdir=str(default_name.parent),
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")],
        )
        if not filename:
            return

        self.save_button.configure(state=tk.DISABLED)
        self.status_var.set("Converting…")

        thread = threading.Thread(
            target=self.perform_conversion,
            args=(Path(filename),),
            daemon=True,
        )
        self.conversion_thread = thread
        thread.start()

    def perform_conversion(self, output_path: Path) -> None:
        try:
            sheet_name = self.sheet_var.get() or None
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
            total_rows = len(df)
            encoding_label = self.encoding_var.get()
            encoding = SUPPORTED_ENCODINGS.get(encoding_label, "utf-8")
            line_ending_choice = LINE_ENDING_OPTIONS.get(self.line_ending_var.get())
            newline = line_ending_choice if line_ending_choice is not None else os.linesep

            def progress_callback(processed: int, total: int = total_rows) -> None:
                self.root.after(0, lambda p=processed, t=total: self.status_var.set(f"Converting… {p}/{t} rows"))

            lines_iter = self.iter_csv_lines(df, progress_callback=progress_callback)
            with output_path.open("w", encoding=encoding, newline="") as fh:
                for line in lines_iter:
                    fh.write(line)
                    fh.write(newline)
            self.root.after(0, lambda: self.status_var.set(f"Converting… {total_rows}/{total_rows} rows"))
            self.root.after(0, lambda: self.on_conversion_success(output_path))
        except Exception as exc:
            self.root.after(0, lambda: self.on_conversion_error(exc))

    def on_conversion_success(self, output_path: Path) -> None:
        self.status_var.set(f"Saved CSV to {output_path}")
        self.save_button.configure(state=tk.NORMAL)
        messagebox.showinfo("Conversion complete", f"CSV saved to {output_path}")

    def on_conversion_error(self, exc: Exception) -> None:
        self.status_var.set("Conversion failed")
        self.save_button.configure(state=tk.NORMAL)
        messagebox.showerror("Conversion error", f"Failed to save CSV: {exc}")

    # ------------------------------------------------------------------
    # Data transformations
    def iter_csv_lines(
        self,
        df: pd.DataFrame,
        limit: Optional[int] = None,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> Iterable[str]:
        separator = self.get_separator() or DEFAULT_SEPARATOR
        quoting_mode = self.quoting_mode_var.get()
        configured_quote = self.get_quote_char()
        quote_char = configured_quote if configured_quote else '"'

        columns = list(df.columns)
        total_rows = len(df)

        def convert_value(value: Any, column: str) -> Tuple[str, bool]:
            if pd.isna(value):
                return "", False
            if isinstance(value, pd.Timestamp):
                ts = value
                if ts.tzinfo is not None:
                    return ts.isoformat(), False
                if ts.hour == 0 and ts.minute == 0 and ts.second == 0 and ts.microsecond == 0:
                    return ts.date().isoformat(), False
                return ts.isoformat(), False
            if isinstance(value, dt.datetime):
                if value.tzinfo is not None:
                    return value.isoformat(), False
                if value.hour == 0 and value.minute == 0 and value.second == 0 and value.microsecond == 0:
                    return value.date().isoformat(), False
                return value.isoformat(), False
            if isinstance(value, dt.date):
                return value.isoformat(), False
            if isinstance(value, pd.Timedelta):
                return str(value), False
            if isinstance(value, Number) and not isinstance(value, bool):
                return format_number(value, self.significant_figures.get(column)), False
            if isinstance(value, bool):
                return "TRUE" if value else "FALSE", False
            if hasattr(value, "isoformat") and not isinstance(value, str):
                try:
                    return value.isoformat(), False
                except Exception:
                    pass
            return str(value), True

        def needs_quote(text: str, is_text: bool) -> bool:
            if quoting_mode == "all":
                return True
            if quoting_mode == "text" and is_text:
                return True
            if separator and separator in text:
                return True
            if "\n" in text or "\r" in text:
                return True
            if quote_char and quote_char in text:
                return True
            return False

        def escape_text(text: str) -> str:
            return text.replace(quote_char, quote_char * 2)

        def format_cells(cells: Sequence[Tuple[str, bool]]) -> str:
            formatted: List[str] = []
            for text, is_text in cells:
                if needs_quote(text, is_text):
                    formatted.append(f"{quote_char}{escape_text(text)}{quote_char}")
                else:
                    formatted.append(text)
            return separator.join(formatted)

        header_cells = [(str(col), True) for col in columns]
        if progress_callback and limit is None:
            progress_callback(0, total_rows)
        yield format_cells(header_cells)

        row_iter: Iterable[Sequence[Any]] = df.itertuples(index=False, name=None)
        if limit is not None:
            row_iter = islice(row_iter, limit)

        for idx, row in enumerate(row_iter, start=1):
            cells = [convert_value(value, column) for value, column in zip(row, columns)]
            yield format_cells(cells)
            if progress_callback and limit is None and idx % 1000 == 0:
                progress_callback(idx, total_rows)

        if progress_callback and limit is None:
            progress_callback(total_rows, total_rows)

    # ------------------------------------------------------------------
    def on_close(self) -> None:
        self.config.separator_choice = self.separator_var.get()
        self.config.custom_separator = self.custom_separator_var.get()
        self.config.quoting_mode = self.quoting_mode_var.get()
        self.config.quote_choice = self.quote_choice_var.get()
        self.config.custom_quote = self.custom_quote_var.get()
        self.config.encoding = self.encoding_var.get()
        self.config.line_ending = self.line_ending_var.get()
        self.config.sheet_name = self.sheet_var.get()
        self.config.last_file = str(self.excel_path) if self.excel_path else ""
        self.config.significant_figures = self.significant_figures
        self.save_config()
        self.root.destroy()

    # ------------------------------------------------------------------
    # Significant figures helpers
    def update_column_list(self, columns: Sequence[str]) -> None:
        existing_selection = self.column_listbox.curselection()
        selected_value = None
        if existing_selection:
            selected_value = self.column_listbox.get(existing_selection[0])

        self.column_listbox.delete(0, tk.END)
        for col in columns:
            self.column_listbox.insert(tk.END, col)

        # Drop entries for columns that no longer exist
        self.significant_figures = {k: v for k, v in self.significant_figures.items() if k in columns}

        if selected_value in columns:
            idx = columns.index(selected_value)
            self.column_listbox.selection_set(idx)
            self.column_listbox.see(idx)
            self.on_column_select()
        else:
            self.sig_fig_status.set("No column selected")
            self.sig_fig_var.set("")

    def on_column_select(self, _event: Optional[tk.Event] = None) -> None:  # type: ignore[override]
        selection = self.column_listbox.curselection()
        if not selection:
            self.sig_fig_status.set("No column selected")
            return
        column = self.column_listbox.get(selection[0])
        current = self.significant_figures.get(column)
        if current is None:
            self.sig_fig_var.set("")
            self.sig_fig_status.set(f"{column}: no significant figures set")
        else:
            self.sig_fig_var.set(str(current))
            self.sig_fig_status.set(f"{column}: {current} significant figure(s)")

    def apply_sig_fig(self) -> None:
        selection = self.column_listbox.curselection()
        if not selection:
            messagebox.showinfo("Select column", "Please choose a column to update.")
            return
        column = self.column_listbox.get(selection[0])
        value = self.sig_fig_var.get().strip()
        if not value:
            messagebox.showinfo("Missing value", "Enter a positive integer for significant figures or use Clear.")
            return
        try:
            sig_figs = int(value)
            if sig_figs <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid value", "Significant figures must be a positive integer.")
            return
        self.significant_figures[column] = sig_figs
        self.sig_fig_status.set(f"{column}: {sig_figs} significant figure(s)")
        self.refresh_preview()

    def clear_sig_fig(self) -> None:
        selection = self.column_listbox.curselection()
        if not selection:
            messagebox.showinfo("Select column", "Please choose a column to clear.")
            return
        column = self.column_listbox.get(selection[0])
        if column in self.significant_figures:
            del self.significant_figures[column]
        self.sig_fig_var.set("")
        self.sig_fig_status.set(f"{column}: no significant figures set")
        self.refresh_preview()


def format_number(value: Number, sig_figs: Optional[int]) -> str:
    """Format numbers with optional significant figures, trimming trailing decimals."""

    if sig_figs is None:
        return str(value)
    try:
        formatted = format(float(value), f".{sig_figs}g")
    except Exception:
        return str(value)

    if "." in formatted:
        formatted = formatted.rstrip("0").rstrip(".") or "0"
    return formatted


def main() -> None:
    root = tk.Tk()
    app = CSVConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
