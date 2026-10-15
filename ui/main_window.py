"""Главное окно приложения (Tkinter).

Структура:
    Notebook с тремя вкладками:
        - «Генерация» — основной сценарий работы
        - «Ключи»     — управление маркерами (CRUD + импорт из шаблона)
        - «Настройки» — пути, режим, структура папок, поведение
"""

import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import (
    Tk, Toplevel, Frame, Label, Entry, Button, Listbox, Scrollbar,
    StringVar, BooleanVar, IntVar, END, SINGLE, filedialog, messagebox,
    ttk, font as tkfont,
)

from core import config as cfg
from core.excel_reader import (
    read_clients, filter_clients, short_fio, sanitize_folder_name,
    get_excel_columns, Client,
)
from core.docx_generator import (
    generate_docx, list_markers_in_template,
)
from core.converter import convert_to_format, find_soffice
from core.formats import FORMATS, list_titles, by_title
from core.logger import setup_logger, log_run_header, log_run_footer


WINDOW_TITLE = "Генератор документов"
WINDOW_W = 780
WINDOW_H = 800
DEBOUNCE_MS = 200


class MainWindow:
    """Главное окно приложения."""

    def __init__(self):
        self.root = Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(f"{WINDOW_W}x{WINDOW_H}")
        self.root.minsize(720, 700)

        # --- Состояние ---
        self.clients_all: list[Client] = []
        self.clients_filtered: list[Client] = []
        self.selected_client: Client | None = None
        self._filter_job = None
        self._generating = False

        # --- Переменные вкладки «Генерация» ---
        self.var_excel = StringVar()
        self.var_templates = StringVar()
        self.var_output = StringVar()
        self.var_search_fio = StringVar()
        self.var_search_id = StringVar()
        self.var_search_contract = StringVar()
        self.var_search_phone = StringVar()
        self.var_status = StringVar(value="Клиент не выбран")
        self.var_markers_status = StringVar(value="")

        # Форматы — чекбоксы
        self.var_fmt_checks: dict[str, BooleanVar] = {
            code: BooleanVar(value=(code in ("docx", "pdf")))
            for code in FORMATS
        }

        # --- Переменные вкладки «Настройки» ---
        self.var_folder_pattern = StringVar()
        self.var_client_pattern = StringVar()
        self.var_on_missing_marker = StringVar()
        self.var_on_missing_column = StringVar()
        self.var_on_empty_value = StringVar()
        self.var_empty_string_spaces = StringVar(value="5")
        self.var_placeholder_underscores = StringVar(value="10")
        self.var_log_mode = StringVar()

        self._build_ui()
        self._load_all_from_config()

    # =================================================================
    # Построение UI
    # =================================================================

    def _build_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=6)

        self.tab_generate = ttk.Frame(self.notebook)
        self.tab_markers = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_generate, text="Генерация")
        self.notebook.add(self.tab_markers, text="Ключи")
        self.notebook.add(self.tab_settings, text="Настройки")

        self._build_tab_generate()
        self._build_tab_markers()
        self._build_tab_settings()

    # -----------------------------------------------------------------
    # Вкладка «Генерация»
    # -----------------------------------------------------------------

    def _build_tab_generate(self):
        parent = self.tab_generate
        pad = {"padx": 8, "pady": 4}

        # === Блок «Файлы» ===
        frame_files = ttk.LabelFrame(parent, text="Файлы")
        frame_files.pack(fill="x", **pad)

        self._path_row(frame_files, "Excel (clients.xlsx):",
                       self.var_excel, self._on_browse_excel, 0)
        self._path_row(frame_files, "Папка шаблонов:",
                       self.var_templates, self._on_browse_templates, 1)
        self._path_row(frame_files, "Папка сохранения:",
                       self.var_output, self._on_browse_output, 2)

        # === Блок «Поиск» ===
        frame_search = ttk.LabelFrame(parent, text="Поиск клиента")
        frame_search.pack(fill="x", **pad)

        grid = ttk.Frame(frame_search)
        grid.pack(fill="x", padx=6, pady=4)

        ttk.Label(grid, text="ФИО:").grid(row=0, column=0, sticky="w", padx=2)
        ttk.Entry(grid, textvariable=self.var_search_fio, width=25).grid(
            row=0, column=1, sticky="we", padx=2)
        ttk.Label(grid, text="ID:").grid(row=0, column=2, sticky="w", padx=2)
        ttk.Entry(grid, textvariable=self.var_search_id, width=15).grid(
            row=0, column=3, sticky="we", padx=2)

        ttk.Label(grid, text="Договор:").grid(row=1, column=0, sticky="w", padx=2)
        ttk.Entry(grid, textvariable=self.var_search_contract, width=25).grid(
            row=1, column=1, sticky="we", padx=2)
        ttk.Label(grid, text="Телефон:").grid(row=1, column=2, sticky="w", padx=2)
        ttk.Entry(grid, textvariable=self.var_search_phone, width=15).grid(
            row=1, column=3, sticky="we", padx=2)

        grid.columnconfigure(1, weight=1)
        grid.columnconfigure(3, weight=1)

        ttk.Button(frame_search, text="Очистить поиск",
                   command=self._on_clear_search).pack(anchor="e", padx=6, pady=(0, 4))

        # === Список клиентов ===
        frame_list = ttk.LabelFrame(parent, text="Клиенты")
        frame_list.pack(fill="both", expand=True, **pad)

        columns = ("id", "fio", "contract", "phone")
        self.tree = ttk.Treeview(
            frame_list, columns=columns, show="headings",
            selectmode="browse", height=10,
        )
        self.tree.heading("id", text="ID")
        self.tree.heading("fio", text="ФИО")
        self.tree.heading("contract", text="Договор")
        self.tree.heading("phone", text="Телефон")
        self.tree.column("id", width=60, anchor="center")
        self.tree.column("fio", width=300)
        self.tree.column("contract", width=120)
        self.tree.column("phone", width=160)

        vsb = ttk.Scrollbar(frame_list, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=4)
        vsb.pack(side="right", fill="y", padx=(0, 6), pady=4)

        self.tree.bind("<<TreeviewSelect>>", self._on_select_client)
        self.tree.bind("<Double-1>", self._on_double_click)

        # === Статус ===
        frame_status = ttk.Frame(parent)
        frame_status.pack(fill="x", **pad)
        ttk.Label(frame_status, text="Выбран:", font=("", 9, "bold")).pack(side="left")
        ttk.Label(frame_status, textvariable=self.var_status).pack(
            side="left", padx=(6, 0))

        # === Форматы ===
        frame_formats = ttk.LabelFrame(parent, text="Форматы вывода")
        frame_formats.pack(fill="x", **pad)

        fgrid = ttk.Frame(frame_formats)
        fgrid.pack(fill="x", padx=6, pady=4)

        col = 0
        row = 0
        for code, info in FORMATS.items():
            cb = ttk.Checkbutton(
                fgrid, text=info.title, variable=self.var_fmt_checks[code],
            )
            cb.grid(row=row, column=col, sticky="w", padx=6, pady=2)
            col += 1
            if col >= 2:
                col = 0
                row += 1

        # === Кнопки ===
        frame_buttons = ttk.Frame(parent)
        frame_buttons.pack(fill="x", **pad)

        ttk.Button(frame_buttons, text="Очистить",
                   command=self._on_clear_all).pack(side="left")
        ttk.Button(frame_buttons, text="Сгенерировать",
                   command=self._on_generate).pack(side="left", padx=6)
        ttk.Button(frame_buttons, text="Отмена",
                   command=self._on_cancel).pack(side="left")

        self.progress = ttk.Progressbar(frame_buttons, mode="determinate")
        self.progress.pack(side="right", fill="x", expand=True, padx=(12, 0))

    def _path_row(self, parent, label, var, browse_cb, row):
        ttk.Label(parent, text=label).grid(
            row=row, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(parent, textvariable=var).grid(
            row=row, column=1, sticky="we", padx=6, pady=3)
        ttk.Button(parent, text="Обзор…", command=browse_cb).grid(
            row=row, column=2, sticky="e", padx=6, pady=3)
        parent.columnconfigure(1, weight=1)

    # -----------------------------------------------------------------
    # Вкладка «Ключи»
    # -----------------------------------------------------------------

    def _build_tab_markers(self):
        parent = self.tab_markers
        pad = {"padx": 8, "pady": 4}

        info = ttk.Label(
            parent,
            text="Сопоставление маркеров {{...}} в шаблоне и колонок в clients.xlsx",
            font=("", 9, "italic"),
        )
        info.pack(anchor="w", **pad)

        # === Treeview ===
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True, **pad)

        columns = ("marker", "column", "format", "required")
        self.markers_tree = ttk.Treeview(
            frame, columns=columns, show="headings",
            selectmode="browse", height=14,
        )
        self.markers_tree.heading("marker", text="Маркер")
        self.markers_tree.heading("column", text="Колонка Excel")
        self.markers_tree.heading("format", text="Формат")
        self.markers_tree.heading("required", text="Обязательный")
        self.markers_tree.column("marker", width=180)
        self.markers_tree.column("column", width=200)
        self.markers_tree.column("format", width=100, anchor="center")
        self.markers_tree.column("required", width=110, anchor="center")

        vsb = ttk.Scrollbar(frame, orient="vertical",
                            command=self.markers_tree.yview)
        self.markers_tree.configure(yscrollcommand=vsb.set)

        self.markers_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.markers_tree.bind("<Double-1>", lambda e: self._on_marker_edit())

        # === Статус ===
        ttk.Label(parent, textvariable=self.var_markers_status,
                  foreground="red").pack(anchor="w", padx=8)

        # === Кнопки ===
        btns = ttk.Frame(parent)
        btns.pack(fill="x", **pad)

        ttk.Button(btns, text="Добавить",
                   command=self._on_marker_add).pack(side="left")
        ttk.Button(btns, text="Изменить",
                   command=self._on_marker_edit).pack(side="left", padx=4)
        ttk.Button(btns, text="Удалить",
                   command=self._on_marker_delete).pack(side="left")
        ttk.Button(btns, text="Импорт из шаблона…",
                   command=self._on_marker_import).pack(side="left", padx=(20, 0))
        ttk.Button(btns, text="Сохранить",
                   command=self._on_marker_save).pack(side="right")

        self._refresh_markers_tree()

    def _refresh_markers_tree(self):
        for item in self.markers_tree.get_children():
            self.markers_tree.delete(item)
        for m in cfg.get_markers():
            marker = m.get("marker", "")
            column = m.get("column", "")
            fmt = m.get("format", "text")
            required = "✔" if m.get("required") else ""
            self.markers_tree.insert("", END, iid=marker,
                                     values=(marker, column, fmt, required))

    def _on_marker_add(self):
        self._open_marker_dialog(None)

    def _on_marker_edit(self):
        sel = self.markers_tree.selection()
        if not sel:
            messagebox.showinfo("Изменить", "Выберите маркер в списке.")
            return
        marker_name = sel[0]
        for m in cfg.get_markers():
            if m.get("marker") == marker_name:
                self._open_marker_dialog(m)
                return

    def _on_marker_delete(self):
        sel = self.markers_tree.selection()
        if not sel:
            return
        marker_name = sel[0]
        if not messagebox.askyesno("Удалить маркер",
                                   f"Удалить маркер {{{{ {marker_name} }}}}?"):
            return
        markers = [m for m in cfg.get_markers()
                   if m.get("marker") != marker_name]
        cfg.save_markers(markers)
        self._refresh_markers_tree()

    def _on_marker_import(self):
        tpl = filedialog.askopenfilename(
            title="Выберите шаблон .docx",
            filetypes=[("Word", "*.docx *.docm *.dotx *.dotm"),
                       ("Все файлы", "*.*")],
        )
        if not tpl:
            return
        try:
            found = list_markers_in_template(tpl)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать шаблон:\n{e}")
            return

        existing = {m.get("marker", "") for m in cfg.get_markers()}
        added = []
        markers = list(cfg.get_markers())
        for name in sorted(found):
            if name in existing:
                continue
            markers.append({
                "marker": name,
                "column": name,       # по умолчанию — имя маркера
                "format": "text",
                "required": False,
            })
            added.append(name)

        if not added:
            messagebox.showinfo("Импорт", "Новых маркеров не найдено.")
            return

        cfg.save_markers(markers)
        self._refresh_markers_tree()
        messagebox.showinfo(
            "Импорт",
            f"Добавлено маркеров: {len(added)}\n\n" + ", ".join(added),
        )

    def _on_marker_save(self):
        # Маркеры уже сохраняются по каждой операции. Здесь —
        # принудительное обновление Excel-статуса и напоминание.
        self._check_marker_columns()
        messagebox.showinfo("Сохранено", "Маркеры сохранены.")

    def _open_marker_dialog(self, marker: dict | None):
        """Диалог добавления/редактирования маркера."""
        dlg = Toplevel(self.root)
        dlg.title("Редактирование маркера" if marker else "Новый маркер")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.geometry("420x260")

        is_new = marker is None
        m = marker or {"marker": "", "column": "", "format": "text", "required": False}

        var_marker = StringVar(value=m.get("marker", ""))
        var_column = StringVar(value=m.get("column", ""))
        var_format = StringVar(value=m.get("format", "text"))
        var_required = BooleanVar(value=bool(m.get("required")))

        frm = ttk.Frame(dlg)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(frm, text="Маркер (без {{ }}):").grid(
            row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=var_marker, width=30).grid(
            row=0, column=1, sticky="we", pady=4)

        ttk.Label(frm, text="Колонка Excel:").grid(
            row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=var_column, width=30).grid(
            row=1, column=1, sticky="we", pady=4)

        ttk.Label(frm, text="Формат:").grid(
            row=2, column=0, sticky="w", pady=4)
        fmt_combo = ttk.Combobox(
            frm, textvariable=var_format, state="readonly",
            values=["text", "amount", "date", "raw"], width=27,
        )
        fmt_combo.grid(row=2, column=1, sticky="we", pady=4)

        ttk.Checkbutton(frm, text="Обязательный",
                        variable=var_required).grid(
            row=3, column=1, sticky="w", pady=8)

        frm.columnconfigure(1, weight=1)

        def on_ok():
            name = var_marker.get().strip().lstrip("{").rstrip("}")
            if not name:
                messagebox.showwarning("Ошибка", "Имя маркера не может быть пустым.")
                return
            # Проверка на дубликат при создании
            if is_new and any(x.get("marker") == name for x in cfg.get_markers()):
                messagebox.showwarning("Ошибка", f"Маркер {name} уже существует.")
                return
            new_data = {
                "marker": name,
                "column": var_column.get().strip(),
                "format": var_format.get(),
                "required": bool(var_required.get()),
            }
            markers = list(cfg.get_markers())
            if is_new:
                markers.append(new_data)
            else:
                old_name = m.get("marker")
                for i, x in enumerate(markers):
                    if x.get("marker") == old_name:
                        markers[i] = new_data
                        break
            cfg.save_markers(markers)
            self._refresh_markers_tree()
            self._check_marker_columns()
            dlg.destroy()

        btns = ttk.Frame(dlg)
        btns.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(btns, text="OK", command=on_ok).pack(side="right")
        ttk.Button(btns, text="Отмена", command=dlg.destroy).pack(
            side="right", padx=6)

    def _check_marker_columns(self):
        """Проверяет, что все колонки из маркеров есть в Excel.

        Показывает ВСЕ ненайденные колонки (с количеством).
        """
        excel_path = self._get_excel_path()
        if not excel_path or not Path(excel_path).exists():
            self.var_markers_status.set("")
            return

        try:
            columns = set(get_excel_columns(str(excel_path)))
        except Exception as e:
            self.var_markers_status.set(f"⚠ Не удалось прочитать Excel: {e}")
            return

        missing: list[str] = []
        for m in cfg.get_markers():
            col = m.get("column", "").strip()
            if col and col not in columns:
                missing.append(col)

        # Уникальные, отсортированные
        missing_unique = sorted(set(missing))

        if missing_unique:
            n = len(missing_unique)
            cols_str = ", ".join(missing_unique)
            self.var_markers_status.set(
                f"⚠ Колонки не найдены в Excel ({n}): {cols_str}"
            )
        else:
            self.var_markers_status.set("")

    # -----------------------------------------------------------------
    # Вкладка «Настройки»
    # -----------------------------------------------------------------

    def _build_tab_settings(self):
        parent = self.tab_settings
        pad = {"padx": 8, "pady": 4}

        # === Структура папок ===
        f_folder = ttk.LabelFrame(parent, text="Структура папок")
        f_folder.pack(fill="x", **pad)

        ttk.Label(f_folder, text="Шаблон верхних папок:").grid(
            row=0, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(f_folder, textvariable=self.var_folder_pattern, width=40).grid(
            row=0, column=1, sticky="we", padx=6, pady=3)
        ttk.Label(f_folder,
                  text="Доступно: {year} {month} {first_letter}",
                  font=("", 8, "italic")).grid(
            row=1, column=1, sticky="w", padx=6)

        ttk.Label(f_folder, text="Шаблон папки клиента:").grid(
            row=2, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(f_folder, textvariable=self.var_client_pattern, width=40).grid(
            row=2, column=1, sticky="we", padx=6, pady=3)
        ttk.Label(f_folder,
                  text="Доступно: {short_fio} {fio} {date} {contract} {id} {phone}",
                  font=("", 8, "italic")).grid(
            row=3, column=1, sticky="w", padx=6)

        f_folder.columnconfigure(1, weight=1)

        # === Поведение при проблемах — три отдельные группы ===

        # 1. Ненайденный маркер
        f_missing_marker = ttk.LabelFrame(
            parent, text="⚠ Ненайденный маркер (нет в конфиге)"
        )
        f_missing_marker.pack(fill="x", **pad)

        ttk.Radiobutton(
            f_missing_marker, text="Подсветить красным + предупредить",
            variable=self.var_on_missing_marker,
            value="highlight_and_warn",
        ).pack(anchor="w", padx=12, pady=2)

        ttk.Radiobutton(
            f_missing_marker, text="Подсветить красным",
            variable=self.var_on_missing_marker,
            value="highlight",
        ).pack(anchor="w", padx=12, pady=2)

        ttk.Radiobutton(
            f_missing_marker, text="Оставить маркер как есть",
            variable=self.var_on_missing_marker,
            value="keep",
        ).pack(anchor="w", padx=12, pady=2)

        # Прочерк + поле количества
        f_mm_ph = ttk.Frame(f_missing_marker)
        f_mm_ph.pack(anchor="w", fill="x", padx=12, pady=2)
        ttk.Radiobutton(
            f_mm_ph, text="Поставить прочерк для заполнения",
            variable=self.var_on_missing_marker,
            value="placeholder",
        ).pack(side="left")
        ttk.Label(f_mm_ph, text="Количество:").pack(side="left", padx=(12, 4))
        ttk.Entry(
            f_mm_ph, textvariable=self.var_placeholder_underscores, width=5,
        ).pack(side="left")

        # Пробелы + поле количества
        f_mm_sp = ttk.Frame(f_missing_marker)
        f_mm_sp.pack(anchor="w", fill="x", padx=12, pady=2)
        ttk.Radiobutton(
            f_mm_sp, text="Подставить пробелы (пустое место)",
            variable=self.var_on_missing_marker,
            value="empty_string",
        ).pack(side="left")
        ttk.Label(f_mm_sp, text="Количество:").pack(side="left", padx=(12, 4))
        ttk.Entry(
            f_mm_sp, textvariable=self.var_empty_string_spaces, width=5,
        ).pack(side="left")

        # Скрыть маркер
        ttk.Radiobutton(
            f_missing_marker, text="Скрыть маркер (удалить из документа)",
            variable=self.var_on_missing_marker,
            value="hide",
        ).pack(anchor="w", padx=12, pady=2)

        # 2. Ненайденная колонка
        f_missing_column = ttk.LabelFrame(parent, text="⚠ Ненайденная колонка (нет в Excel)")
        f_missing_column.pack(fill="x", **pad)

        ttk.Radiobutton(
            f_missing_column, text="Предупредить в UI",
            variable=self.var_on_missing_column,
            value="warn",
        ).pack(anchor="w", padx=12, pady=2)
        ttk.Radiobutton(
            f_missing_column, text="Молча пропустить",
            variable=self.var_on_missing_column,
            value="silent",
        ).pack(anchor="w", padx=12, pady=2)

        # 3. Пустое поле в Excel
        f_empty = ttk.LabelFrame(parent, text="⚠ Пустое поле в Excel")
        f_empty.pack(fill="x", **pad)

        ttk.Radiobutton(
            f_empty, text="Оставить маркер + подсветить красным",
            variable=self.var_on_empty_value,
            value="keep_marker",
        ).pack(anchor="w", padx=12, pady=2)

        ttk.Radiobutton(
            f_empty, text="Оставить как есть",
            variable=self.var_on_empty_value,
            value="keep",
        ).pack(anchor="w", padx=12, pady=2)

        # Прочерк + поле для количества
        f_ph_row = ttk.Frame(f_empty)
        f_ph_row.pack(anchor="w", fill="x", padx=12, pady=2)
        ttk.Radiobutton(
            f_ph_row, text="Поставить прочерк для заполнения",
            variable=self.var_on_empty_value,
            value="placeholder",
        ).pack(side="left")
        ttk.Label(f_ph_row, text="Количество:").pack(side="left", padx=(12, 4))
        ttk.Entry(
            f_ph_row, textvariable=self.var_placeholder_underscores, width=5,
        ).pack(side="left")

        # Пробелы + поле для количества
        f_sp_row = ttk.Frame(f_empty)
        f_sp_row.pack(anchor="w", fill="x", padx=12, pady=2)
        ttk.Radiobutton(
            f_sp_row, text="Подставить пробелы (пустое место)",
            variable=self.var_on_empty_value,
            value="empty_string",
        ).pack(side="left")
        ttk.Label(f_sp_row, text="Количество:").pack(side="left", padx=(12, 4))
        ttk.Entry(
            f_sp_row, textvariable=self.var_empty_string_spaces, width=5,
        ).pack(side="left")

        # === Логи ===
        f_log = ttk.LabelFrame(parent, text="Логи")
        f_log.pack(fill="x", **pad)

        ttk.Radiobutton(f_log, text="Отдельный файл на каждый запуск (с датой)",
                        variable=self.var_log_mode,
                        value="per_run").pack(anchor="w", padx=6, pady=2)
        ttk.Radiobutton(f_log, text="Один файл generation_log.txt (перезапись)",
                        variable=self.var_log_mode,
                        value="single_file").pack(anchor="w", padx=6, pady=2)

        # === Кнопки ===
        btns = ttk.Frame(parent)
        btns.pack(fill="x", **pad)
        ttk.Button(btns, text="Сохранить настройки",
                   command=self._on_settings_save).pack(side="left")
        ttk.Button(btns, text="Сбросить к дефолтным",
                   command=self._on_settings_reset).pack(side="left", padx=6)
        ttk.Button(btns, text="Открыть папку программы",
                   command=self._on_open_install_dir).pack(side="right")

    def _on_settings_save(self):
        # Парсим числовые поля — если мусор, используем дефолты
        try:
            spaces = int(self.var_empty_string_spaces.get().strip() or "5")
            if spaces < 0:
                spaces = 0
        except ValueError:
            spaces = 5

        try:
            underscores = int(self.var_placeholder_underscores.get().strip() or "10")
            if underscores < 0:
                underscores = 0
        except ValueError:
            underscores = 10

        cfg.save_settings({
            "folder_pattern": self.var_folder_pattern.get().strip() or "{year}/{month}",
            "client_pattern": self.var_client_pattern.get().strip() or "{short_fio}_{date}_{contract}",
            "on_missing_marker": self.var_on_missing_marker.get(),
            "on_missing_column": self.var_on_missing_column.get(),
            "on_empty_value": self.var_on_empty_value.get(),
            "empty_string_spaces": spaces,
            "placeholder_underscores": underscores,
            "log_mode": self.var_log_mode.get(),
        })
        messagebox.showinfo("Сохранено", "Настройки сохранены.")

    def _on_settings_reset(self):
        if not messagebox.askyesno("Сброс", "Сбросить настройки к дефолтным?"):
            return
        cfg.save_settings({})  # пустой → get_settings вернёт дефолты
        self._load_settings_into_ui()
        messagebox.showinfo("Готово", "Настройки сброшены.")

    def _on_open_install_dir(self):
        import subprocess
        install = cfg.get_install_dir()
        try:
            if sys.platform == "win32":
                subprocess.Popen(["explorer", str(install)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(install)])
            else:
                subprocess.Popen(["xdg-open", str(install)])
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть папку:\n{e}")

    # =================================================================
    # Загрузка/сохранение конфига
    # =================================================================

    def _load_all_from_config(self):
        """Читает relative_paths и настройки из config.json."""
        self._load_paths_into_ui()
        self._load_settings_into_ui()
        self._check_marker_columns()

        # Если Excel уже указан — попробуем загрузить клиентов
        excel_path = self._get_excel_path()
        if excel_path and Path(excel_path).exists():
            self.root.after(100, self._reload_clients)

    def _load_paths_into_ui(self):
        """Читает relative_paths и разрешает их относительно install_dir."""
        cfg_data = cfg.load_config()
        rp = cfg_data.get("relative_paths", {})

        def resolve(key: str, default: str) -> str:
            val = rp.get(key, default)
            if not val:
                return ""
            return str(cfg.resolve_path(val))

        self.var_excel.set(resolve("excel_path", "clients.xlsx"))
        self.var_templates.set(resolve("templates_dir", "Шаблоны"))
        self.var_output.set(resolve("output_dir", "Готовые документы"))

    def _load_settings_into_ui(self):
        s = cfg.get_settings()
        self.var_folder_pattern.set(s.get("folder_pattern", "{year}/{month}"))
        self.var_client_pattern.set(s.get("client_pattern", "{short_fio}_{date}_{contract}"))
        self.var_on_missing_marker.set(s.get("on_missing_marker", "highlight_and_warn"))
        self.var_on_missing_column.set(s.get("on_missing_column", "warn"))
        self.var_on_empty_value.set(s.get("on_empty_value", "empty_string"))
        self.var_empty_string_spaces.set(str(s.get("empty_string_spaces", 5)))
        self.var_placeholder_underscores.set(str(s.get("placeholder_underscores", 10)))
        self.var_log_mode.set(s.get("log_mode", "per_run"))

    def _save_paths_to_config(self):
        """Сохраняет текущие пути как relative_paths."""
        rp = {
            "excel_path": cfg.to_relative(self.var_excel.get()),
            "templates_dir": cfg.to_relative(self.var_templates.get()),
            "output_dir": cfg.to_relative(self.var_output.get()),
        }
        cfg.update(relative_paths=rp)

    def _get_excel_path(self) -> str:
        return self.var_excel.get().strip()

    # =================================================================
    # Обзор путей
    # =================================================================

    def _on_browse_excel(self):
        path = filedialog.askopenfilename(
            title="Выберите Excel-файл с клиентами",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Все файлы", "*.*")],
        )
        if path:
            self.var_excel.set(path)
            self._save_paths_to_config()
            self._reload_clients()
            self._check_marker_columns()

    def _on_browse_templates(self):
        path = filedialog.askdirectory(title="Папка с шаблонами")
        if path:
            self.var_templates.set(path)
            self._save_paths_to_config()

    def _on_browse_output(self):
        path = filedialog.askdirectory(title="Папка для сохранения")
        if path:
            self.var_output.set(path)
            self._save_paths_to_config()

    # =================================================================
    # Клиенты
    # =================================================================

    def _reload_clients(self):
        path = self.var_excel.get().strip()
        if not path:
            return
        p = Path(path)
        if not p.exists():
            messagebox.showerror("Ошибка", f"Файл не найден:\n{path}")
            return
        try:
            self.clients_all = read_clients(path)
        except Exception as e:
            messagebox.showerror("Ошибка чтения Excel", str(e))
            return
        self._apply_filter()

    def _apply_filter(self):
        self._filter_job = None
        self.clients_filtered = filter_clients(
            self.clients_all,
            fio=self.var_search_fio.get(),
            id=self.var_search_id.get(),
            contract=self.var_search_contract.get(),
            phone=self.var_search_phone.get(),
        )
        self._refresh_list()

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        for c in self.clients_filtered:
            self.tree.insert("", END, iid=c.id,
                             values=(c.id, c.fio, c.contract_num, c.phone))
        if self.selected_client is not None:
            if self.selected_client.id not in [c.id for c in self.clients_filtered]:
                self.selected_client = None
                self.var_status.set("Клиент не выбран")

    def _on_select_client(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        cid = sel[0]
        for c in self.clients_filtered:
            if c.id == cid:
                self.selected_client = c
                self.var_status.set(f"ID={c.id} | {c.fio} | дог. {c.contract_num}")
                return

    def _on_double_click(self, _event=None):
        self._on_select_client()
        self._on_generate()

    def _on_clear_search(self):
        self.var_search_fio.set("")
        self.var_search_id.set("")
        self.var_search_contract.set("")
        self.var_search_phone.set("")

    def _on_clear_all(self):
        self._on_clear_search()
        self.var_excel.set("")
        self.var_templates.set("")
        self.var_output.set("")
        self.clients_all = []
        self.clients_filtered = []
        self.selected_client = None
        self.var_status.set("Клиент не выбран")
        self._refresh_list()
        for var in self.var_fmt_checks.values():
            var.set(False)

    def _bind_live_search(self):
        for var in (self.var_search_fio, self.var_search_id,
                    self.var_search_contract, self.var_search_phone):
            var.trace_add("write", self._schedule_filter)

    def _schedule_filter(self, *_args):
        if self._filter_job is not None:
            try:
                self.root.after_cancel(self._filter_job)
            except Exception:
                pass
        self._filter_job = self.root.after(DEBOUNCE_MS, self._apply_filter)

    # =================================================================
    # Генерация
    # =================================================================

    def _on_cancel(self):
        if self._generating:
            messagebox.showinfo("Отмена", "Дождитесь завершения текущей генерации.")
        else:
            self.root.destroy()

    def _on_generate(self):
        if self._generating:
            return

        if self.selected_client is None:
            messagebox.showwarning("Внимание", "Сначала выберите клиента.")
            return
        if not self.var_templates.get().strip():
            messagebox.showwarning("Внимание", "Укажите папку шаблонов.")
            return
        if not self.var_output.get().strip():
            messagebox.showwarning("Внимание", "Укажите папку сохранения.")
            return

        selected = [code for code, var in self.var_fmt_checks.items() if var.get()]
        if not selected:
            messagebox.showwarning("Внимание", "Выберите хотя бы один формат.")
            return

        if any(FORMATS[c].needs_libreoffice for c in selected):
            if not find_soffice():
                messagebox.showerror(
                    "LibreOffice не найден",
                    "Для выбранных форматов требуется LibreOffice.\n"
                    "Установите его или выберите только DOCX/TXT.",
                )
                return

        self._generating = True
        self._save_paths_to_config()
        threading.Thread(target=self._generate_worker,
                         args=(self.selected_client, selected),
                         daemon=True).start()

    def _generate_worker(self, client: Client, selected: list[str]):
        templates_dir = Path(self.var_templates.get())
        output_root = Path(self.var_output.get())
        output_root.mkdir(parents=True, exist_ok=True)

        # Логи
        settings = cfg.get_settings()
        log_mode = settings.get("log_mode", "per_run")
        if log_mode == "single_file":
            log_name = "generation_log.txt"
            logs_dir = output_root
        else:
            logs_dir = output_root / "_logs"
            log_name = f"generation_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
        logs_dir.mkdir(parents=True, exist_ok=True)
        log = setup_logger(str(logs_dir), filename=log_name)

        # Формируем путь к папке клиента
        year, month = "", ""
        if client.date and "." in client.date:
            parts = client.date.split(".")
            if len(parts) == 3:
                _, month, year = parts

        folder_pattern = settings.get("folder_pattern", "{year}/{month}")
        top_dir_str = folder_pattern.format(
            year=year, month=month,
            first_letter=(client.fio[:1].upper() if client.fio else "_"),
        ).strip("/").replace("//", "/")
        client_dir = output_root
        if top_dir_str:
            for part in top_dir_str.split("/"):
                if part:
                    client_dir = client_dir / sanitize_folder_name(part)
        client_dir.mkdir(parents=True, exist_ok=True)

        client_pattern = settings.get("client_pattern", "{short_fio}_{date}_{contract}")
        client_folder = client_pattern.format(
            short_fio=short_fio(client.fio),
            date=client.date,
            contract=client.contract_num,
            id=client.id,
            phone=client.phone,
            fio=client.fio,
        )
        client_folder = sanitize_folder_name(client_folder) or f"client_{client.id}"
        client_dir = client_dir / client_folder
        client_dir.mkdir(parents=True, exist_ok=True)

        log_run_header(
            log,
            Excel=self.var_excel.get(),
            Шаблоны=str(templates_dir),
            Куда=str(client_dir),
            Форматы=", ".join(selected),
            Клиент=f"{client.id} — {client.fio}",
        )

        exts = (".docx", ".docm", ".dotx", ".dotm")
        templates = [p for p in templates_dir.iterdir()
                     if p.is_file() and p.suffix.lower() in exts]
        if not templates:
            log.error("Шаблоны не найдены в %s", templates_dir)
            self.root.after(0, lambda: self._finish_generation(0, 0, "Шаблонов нет"))
            return

        log.info("Найдено шаблонов: %d", len(templates))
        for tpl in templates:
            log.info("  • %s", tpl.name)

        markers = cfg.get_markers()
        total = len(templates) * len(selected)
        done = 0
        ok = 0
        fail = 0
        unreplaced_all: list[str] = []

        for tpl in templates:
            docx_name = f"{tpl.stem}.docx"
            docx_path = client_dir / docx_name
            try:
                unreplaced = generate_docx(
                    str(tpl), str(docx_path), client,
                    markers=markers, return_unreplaced=True,
                )
                if unreplaced:
                    unreplaced_all.extend(sorted(unreplaced))
                    log.warning("%s → %s — OK (незаменённых маркеров: %d)",
                                tpl.name, docx_name, len(unreplaced))
                else:
                    log.info("%s → %s — OK", tpl.name, docx_name)
            except Exception as e:
                log.error("%s → DOCX — ОШИБКА: %s", tpl.name, e)
                fail += len(selected)
                done += len(selected)
                self._update_progress(done, total)
                continue

            for code in selected:
                done += 1
                fmt = FORMATS[code]
                try:
                    if code == "docx":
                        ok += 1
                    else:
                        out_path = convert_to_format(
                            str(docx_path), str(client_dir), fmt)
                        log.info("%s → %s — OK", tpl.name, Path(out_path).name)
                        ok += 1
                except Exception as e:
                    log.error("%s → %s — ОШИБКА: %s", tpl.name, code.upper(), e)
                    fail += 1
                self._update_progress(done, total)

        log_run_footer(log, ok=ok, fail=fail)

        # Уникальные незаменённые маркеры
        unique_unreplaced = sorted(set(unreplaced_all))
        self.root.after(0, lambda: self._finish_generation(
            ok, fail, None, unique_unreplaced))

    def _update_progress(self, done: int, total: int):
        pct = int(100 * done / max(total, 1))
        self.root.after(0, lambda: self.progress.configure(value=pct))

    def _finish_generation(self, ok: int, fail: int,
                           note: str | None,
                           unreplaced: list[str] | None = None):
        self._generating = False
        self.progress.configure(value=0)

        if note:
            messagebox.showwarning("Готово", f"{note}\nУспешно: {ok}, ошибок: {fail}")
            return

        msg = (f"Генерация завершена.\n"
               f"Успешно: {ok}, ошибок: {fail}\n\n"
               f"Подробности: {self.var_output.get()}/")
        if unreplaced:
            msg += ("\n\n⚠ Не заменены маркеры:\n"
                    + ", ".join(f"{{{{ {m} }}}}" for m in unreplaced[:10]))
            if len(unreplaced) > 10:
                msg += f"\n... и ещё {len(unreplaced) - 10}"
        messagebox.showinfo("Готово", msg)

    # =================================================================

    def run(self):
        self._bind_live_search()
        self.root.mainloop()
