"""Главное окно приложения (Tkinter)."""

import threading
from pathlib import Path
from tkinter import (
    Tk, Toplevel, Frame, Label, Entry, Button, Listbox, Scrollbar,
    StringVar, IntVar, BooleanVar, END, SINGLE, filedialog, messagebox,
    ttk, font as tkfont,
)

from core import config as cfg
from core.excel_reader import read_clients, filter_clients, Client
from core.docx_generator import generate_docx
from core.converter import convert_to_format, find_soffice
from core.formats import FORMATS, list_titles, by_title
from core.logger import setup_logger, log_run_header, log_run_footer


WINDOW_TITLE = "Генератор документов"
WINDOW_W = 600
WINDOW_H = 750

DEBOUNCE_MS = 200  # задержка перед фильтрацией после ввода


class MainWindow:
    """Главное окно приложения."""

    def __init__(self):
        self.root = Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(f"{WINDOW_W}x{WINDOW_H}")
        self.root.minsize(560, 700)

        # --- Состояние ---
        self.clients_all: list[Client] = []
        self.clients_filtered: list[Client] = []
        self.selected_client: Client | None = None
        self._filter_job = None  # id отложенного вызова after()
        self._generating = False

        # --- Переменные Tkinter ---
        self.var_excel = StringVar()
        self.var_templates = StringVar()
        self.var_output = StringVar()
        self.var_search_fio = StringVar()
        self.var_search_id = StringVar()
        self.var_search_contract = StringVar()
        self.var_search_phone = StringVar()
        self.var_status = StringVar(value="Клиент не выбран")
        self.var_format_titles = list_titles()
        self.var_format = StringVar(value=self.var_format_titles[0])

        # Формат-чекбоксы (по ТЗ можно выбрать несколько форматов сразу)
        self.var_fmt_checks: dict[str, BooleanVar] = {
            code: BooleanVar(value=(code in ("docx", "pdf")))
            for code in FORMATS
        }

        self._build_ui()
        self._load_config_into_ui()
        self._bind_traces()

    # ------------------------------------------------------------------
    # Построение интерфейса
    # ------------------------------------------------------------------

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # ===== Блок «Файлы» =====
        frame_files = ttk.LabelFrame(self.root, text="Файлы")
        frame_files.pack(fill="x", **pad)

        self._path_row(frame_files, "Excel (clients.xlsx):",
                       self.var_excel, self._on_browse_excel, 0)
        self._path_row(frame_files, "Папка шаблонов:",
                       self.var_templates, self._on_browse_templates, 1)
        self._path_row(frame_files, "Папка сохранения:",
                       self.var_output, self._on_browse_output, 2)

        # ===== Блок «Поиск клиента» =====
        frame_search = ttk.LabelFrame(self.root, text="Поиск клиента")
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

        # Кнопка «Очистить поиск»
        ttk.Button(frame_search, text="Очистить поиск",
                   command=self._on_clear_search).pack(anchor="e", padx=6, pady=(0, 4))

        # ===== Список клиентов =====
        frame_list = ttk.LabelFrame(self.root, text="Клиенты")
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
        self.tree.column("fio", width=260)
        self.tree.column("contract", width=110)
        self.tree.column("phone", width=140)

        vsb = ttk.Scrollbar(frame_list, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=4)
        vsb.pack(side="right", fill="y", padx=(0, 6), pady=4)

        self.tree.bind("<<TreeviewSelect>>", self._on_select_client)
        self.tree.bind("<Double-1>", self._on_double_click)

        # ===== Статус выбранного клиента =====
        frame_status = ttk.Frame(self.root)
        frame_status.pack(fill="x", **pad)
        ttk.Label(frame_status, text="Выбран:", font=("", 9, "bold")).pack(side="left")
        ttk.Label(frame_status, textvariable=self.var_status).pack(
            side="left", padx=(6, 0))

        # ===== Блок «Форматы» =====
        frame_formats = ttk.LabelFrame(self.root, text="Форматы вывода")
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

        # ===== Кнопки =====
        frame_buttons = ttk.Frame(self.root)
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

    # ------------------------------------------------------------------
    # Загрузка/сохранение конфига
    # ------------------------------------------------------------------

    def _load_config_into_ui(self):
        c = cfg.load_config()
        self.var_excel.set(c.get("excel_path", ""))
        self.var_templates.set(c.get("templates_dir", ""))
        self.var_output.set(c.get("output_dir", ""))

        # Загружаем Excel сразу, если путь есть
        if self.var_excel.get():
            self.root.after(50, self._reload_clients)

    def _save_config_from_ui(self):
        # Чекбоксы форматов
        selected_formats = [
            code for code, var in self.var_fmt_checks.items() if var.get()
        ]
        cfg.update(
            excel_path=self.var_excel.get(),
            templates_dir=self.var_templates.get(),
            output_dir=self.var_output.get(),
            selected_formats=selected_formats,
        )

    # ------------------------------------------------------------------
    # Привязка «живого» поиска
    # ------------------------------------------------------------------

    def _bind_traces(self):
        for var in (self.var_search_fio, self.var_search_id,
                    self.var_search_contract, self.var_search_phone):
            var.trace_add("write", self._schedule_filter)

    def _schedule_filter(self, *_args):
        # Откладываем фильтрацию — debounce 200 мс
        if self._filter_job is not None:
            self.root.after_cancel(self._filter_job)
        self._filter_job = self.root.after(DEBOUNCE_MS, self._apply_filter)

    # ------------------------------------------------------------------
    # Обзор файлов и папок
    # ------------------------------------------------------------------

    def _on_browse_excel(self):
        path = filedialog.askopenfilename(
            title="Выберите Excel-файл с клиентами",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Все файлы", "*.*")],
        )
        if path:
            self.var_excel.set(path)
            self._reload_clients()
            self._save_config_from_ui()

    def _on_browse_templates(self):
        path = filedialog.askdirectory(title="Папка с шаблонами")
        if path:
            self.var_templates.set(path)
            self._save_config_from_ui()

    def _on_browse_output(self):
        path = filedialog.askdirectory(title="Папка для сохранения")
        if path:
            self.var_output.set(path)
            self._save_config_from_ui()

    # ------------------------------------------------------------------
    # Работа с клиентами
    # ------------------------------------------------------------------

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
        # Сбросим выбор, если клиент исчез из списка
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
        # Двойной клик по строке — сразу генерация
        self._on_select_client()
        self._on_generate()

    def _on_clear_search(self):
        self.var_search_fio.set("")
        self.var_search_id.set("")
        self.var_search_contract.set("")
        self.var_search_phone.set("")
        # фильтрация сработает через trace

    def _on_clear_all(self):
        self.var_search_fio.set("")
        self.var_search_id.set("")
        self.var_search_contract.set("")
        self.var_search_phone.set("")
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

    # ------------------------------------------------------------------
    # Генерация
    # ------------------------------------------------------------------

    def _on_cancel(self):
        if self._generating:
            messagebox.showinfo("Отмена", "Дождитесь завершения текущей генерации.")
        else:
            self.root.destroy()

    def _on_generate(self):
        if self._generating:
            return

        # Валидация
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

        # Проверка LibreOffice
        if any(FORMATS[c].needs_libreoffice for c in selected):
            if not find_soffice():
                messagebox.showerror(
                    "LibreOffice не найден",
                    "Для выбранных форматов требуется LibreOffice.\n"
                    "Установите его или выберите только DOCX/TXT.",
                )
                return

        # Запуск в отдельном потоке, чтобы GUI не подвисал
        self._generating = True
        self._save_config_from_ui()
        threading.Thread(target=self._generate_worker,
                         args=(self.selected_client, selected),
                         daemon=True).start()

    def _generate_worker(self, client: Client, selected: list[str]):
        from core.excel_reader import short_fio, sanitize_folder_name
        from core.config import get_folder_pattern, get_client_pattern
        from datetime import datetime

        templates_dir = Path(self.var_templates.get())
        output_root = Path(self.var_output.get())
        output_root.mkdir(parents=True, exist_ok=True)

        # --- Логи в отдельной папке _logs с датой в имени ---
        logs_dir = output_root / "_logs"
        logs_dir.mkdir(parents=True, exist_ok=True)
        log_name = f"generation_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
        log = setup_logger(str(logs_dir), filename=log_name)

        # --- Формируем путь к папке клиента по шаблонам ---
        # Разбираем дату DD.MM.YYYY → year, month
        year, month = "", ""
        if client.date and "." in client.date:
            parts = client.date.split(".")
            if len(parts) == 3:
                _, month, year = parts

        # Верхние папки: по умолчанию "{year}/{month}" → 2024/01
        folder_pattern = get_folder_pattern()
        top_dir_str = folder_pattern.format(
            year=year,
            month=month,
            first_letter=(client.fio[:1].upper() if client.fio else "_"),
        ).strip("/").replace("//", "/")

        client_dir = output_root
        if top_dir_str:
            for part in top_dir_str.split("/"):
                if part:
                    client_dir = client_dir / sanitize_folder_name(part)
        client_dir.mkdir(parents=True, exist_ok=True)

        # Папка клиента: по умолчанию "{short_fio}_{date}_{contract}"
        client_pattern = get_client_pattern()
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

        # --- Логирование запуска ---
        log_run_header(
            log,
            Excel=self.var_excel.get(),
            Шаблоны=str(templates_dir),
            Куда=str(client_dir),
            Форматы=", ".join(selected),
            Клиент=f"{client.id} — {client.fio}",
        )

        # --- Сканируем шаблоны ---
        exts = (".docx", ".docm", ".dotx", ".dotm")
        templates = [
            p for p in templates_dir.iterdir()
            if p.is_file() and p.suffix.lower() in exts
        ]
        if not templates:
            log.error("Шаблоны не найдены в %s", templates_dir)
            self.root.after(
                0,
                lambda: self._finish_generation(0, 0, "Шаблонов нет"),
            )
            return

        log.info("Найдено шаблонов: %d", len(templates))
        for tpl in templates:
            log.info("  • %s", tpl.name)

        total = len(templates) * len(selected)
        done = 0
        ok = 0
        fail = 0

        for tpl in templates:
            # Имя файла: <Шаблон>.docx (без _ID — папка уже содержит ID/дату)
            docx_name = f"{tpl.stem}.docx"
            docx_path = client_dir / docx_name
            try:
                generate_docx(str(tpl), str(docx_path), client)
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
                        # DOCX уже создан выше — учитываем статистику
                        ok += 1
                    else:
                        out_path = convert_to_format(
                            str(docx_path), str(client_dir), fmt,
                        )
                        log.info("%s → %s — OK", tpl.name, Path(out_path).name)
                        ok += 1
                except Exception as e:
                    log.error("%s → %s — ОШИБКА: %s", tpl.name, code.upper(), e)
                    fail += 1

                self._update_progress(done, total)

        log_run_footer(log, ok=ok, fail=fail)

        self.root.after(0, lambda: self._finish_generation(ok, fail, None))

    def _update_progress(self, done: int, total: int):
        pct = int(100 * done / max(total, 1))
        self.root.after(0, lambda: self.progress.configure(value=pct))

    def _finish_generation(self, ok: int, fail: int, note: str | None):
        self._generating = False
        self.progress.configure(value=0)
        if note:
            messagebox.showwarning("Готово", f"{note}\nУспешно: {ok}, ошибок: {fail}")
        else:
            messagebox.showinfo(
                "Готово",
                f"Генерация завершена.\nУспешно: {ok}, ошибок: {fail}\n"
                f"Подробности: {self.var_output.get()}/generation_log.txt",
            )

    # ------------------------------------------------------------------

    def run(self):
        self.root.mainloop()
