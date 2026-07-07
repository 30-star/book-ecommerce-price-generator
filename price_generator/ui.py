import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .cache import export_price_cache_xlsx, load_price_cache, save_price_cache
from .constants import *
from .paths import default_price_cache_path, default_rules_path, resource_path
from .pricing import normalize_price_cache_key
from .rules import load_delete_rules, save_default_rules
from .sales_analysis_cache import load_sales_analysis_cache, save_sales_analysis_cache
from .spreadsheet import (
    append_cached_prices_csv,
    append_cached_prices_xlsx,
    default_output_path,
    default_price_match_output_path,
    import_sales_analysis_to_cache_csv,
    import_sales_analysis_to_cache_xlsx,
    import_prices_to_cache_csv,
    import_prices_to_cache_xlsx,
    process_csv,
    process_xlsx,
)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("图书电商线上活动价格自动生成器")
        icon_path = resource_path("app.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        self.geometry("780x780")
        self.minsize(720, 740)
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.cache_import_path = tk.StringVar()
        self.combo_cache_import_path = tk.StringVar()
        self.match_input_path = tk.StringVar()
        self.match_output_path = tk.StringVar()
        self.sales_analysis_import_path = tk.StringVar()
        self.rules_path = tk.StringVar(value=default_rules_path())
        self.status = tk.StringVar(value="请选择表格导入价格库，或从价格库匹配售价。")
        self.progress = tk.IntVar(value=0)
        self.shipping_fee_vars = [tk.StringVar(value=str(default)) for _, default in SHIPPING_FEE_RULES]
        self.price_threshold = tk.StringVar(value=str(DEFAULT_PRICE_THRESHOLD))
        self.low_price_margin = tk.StringVar(value=str(DEFAULT_LOW_PRICE_MARGIN))
        self.high_price_margin = tk.StringVar(value=str(DEFAULT_HIGH_PRICE_MARGIN))
        self.combo_price_threshold = tk.StringVar(value=str(DEFAULT_PRICE_THRESHOLD))
        self.combo_low_price_margin = tk.StringVar(value=str(DEFAULT_LOW_PRICE_MARGIN))
        self.combo_high_price_margin = tk.StringVar(value=str(DEFAULT_HIGH_PRICE_MARGIN))
        self.price_match_column = tk.StringVar(value=DEFAULT_PRICE_MATCH_COLUMN_NAME)
        self._build_ui()

    def _build_ui(self):
        frame = ttk.Frame(self, padding=18)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        price_frames = ttk.Frame(frame)
        price_frames.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        price_frames.columnconfigure(0, weight=1)
        price_frames.columnconfigure(1, weight=1)

        normal_group = ttk.LabelFrame(price_frames, text="普通商品", padding=10)
        normal_group.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        normal_group.columnconfigure(0, weight=1)

        combo_group = ttk.LabelFrame(price_frames, text="组合商品", padding=10)
        combo_group.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        combo_group.columnconfigure(0, weight=1)

        price_frame = ttk.LabelFrame(normal_group, text="售价规则", padding=10)
        price_frame.grid(row=0, column=0, sticky="ew")
        ttk.Label(price_frame, text="分界成本价").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(price_frame, textvariable=self.price_threshold, width=8).grid(row=0, column=1, sticky="w", pady=3)
        ttk.Label(price_frame, text="成本<=分界 毛利率").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(price_frame, textvariable=self.low_price_margin, width=8).grid(row=1, column=1, sticky="w", pady=3)
        ttk.Label(price_frame, text="成本>分界 毛利率").grid(row=1, column=2, sticky="w", padx=(16, 8), pady=3)
        ttk.Entry(price_frame, textvariable=self.high_price_margin, width=8).grid(row=1, column=3, sticky="w", pady=3)

        combo_price_frame = ttk.LabelFrame(combo_group, text="组合售价规则", padding=10)
        combo_price_frame.grid(row=0, column=0, sticky="ew")
        ttk.Label(combo_price_frame, text="分界成本价").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(combo_price_frame, textvariable=self.combo_price_threshold, width=8).grid(row=0, column=1, sticky="w", pady=3)
        ttk.Label(combo_price_frame, text="成本<=分界 毛利率").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(combo_price_frame, textvariable=self.combo_low_price_margin, width=8).grid(row=1, column=1, sticky="w", pady=3)
        ttk.Label(combo_price_frame, text="成本>分界 毛利率").grid(row=1, column=2, sticky="w", padx=(16, 8), pady=3)
        ttk.Entry(combo_price_frame, textvariable=self.combo_high_price_margin, width=8).grid(row=1, column=3, sticky="w", pady=3)

        fee_frame = ttk.LabelFrame(frame, text="快递费规则（填写返回数字）", padding=10)
        fee_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        fee_frame.columnconfigure(1, weight=1)
        fee_frame.columnconfigure(3, weight=1)

        for index, (label, _) in enumerate(SHIPPING_FEE_RULES):
            row = index // 2
            column = (index % 2) * 2
            ttk.Label(fee_frame, text=label).grid(row=row, column=column, sticky="w", padx=(0, 8), pady=3)
            ttk.Entry(fee_frame, textvariable=self.shipping_fee_vars[index], width=8).grid(
                row=row, column=column + 1, sticky="w", padx=(0, 18), pady=3
            )

        import_cache_frame = ttk.LabelFrame(normal_group, text="导入单品表格到价格库", padding=10)
        import_cache_frame.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        import_cache_frame.columnconfigure(1, weight=1)
        ttk.Label(import_cache_frame, text="导入表格").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(import_cache_frame, textvariable=self.cache_import_path, width=28).grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=3)
        ttk.Button(import_cache_frame, text="选择", command=self.choose_cache_import).grid(row=0, column=2, sticky="ew", pady=3)
        ttk.Label(
            import_cache_frame,
            text="读取“商品编码”“商品重量”“成本价”，按当前公式计算售价并保存到价格库。",
            foreground="#435064",
            wraplength=320,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 8))
        self.cache_import_button = ttk.Button(import_cache_frame, text="计算并导入价格库", command=self.start_price_cache_import)
        self.cache_import_button.grid(row=2, column=0, columnspan=3, sticky="ew", ipady=6)

        combo_cache_frame = ttk.LabelFrame(combo_group, text="导入组合商品表格到价格库", padding=10)
        combo_cache_frame.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        combo_cache_frame.columnconfigure(1, weight=1)
        ttk.Label(combo_cache_frame, text="导入表格").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(combo_cache_frame, textvariable=self.combo_cache_import_path, width=28).grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=3)
        ttk.Button(combo_cache_frame, text="选择", command=self.choose_combo_cache_import).grid(row=0, column=2, sticky="ew", pady=3)
        ttk.Label(
            combo_cache_frame,
            text="读取“组合商品编码”“组合重量”“组合成本价”，按当前公式计算售价并保存到价格库。",
            foreground="#435064",
            wraplength=320,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 8))
        self.combo_cache_import_button = ttk.Button(
            combo_cache_frame,
            text="计算并导入价格库",
            command=self.start_combo_price_cache_import,
        )
        self.combo_cache_import_button.grid(row=2, column=0, columnspan=3, sticky="ew", ipady=6)

        cache_row = ttk.Frame(frame)
        cache_row.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        cache_row.columnconfigure(0, weight=1, uniform="cache_row")
        cache_row.columnconfigure(1, weight=1, uniform="cache_row")

        cache_frame = ttk.LabelFrame(cache_row, text="从价格库匹配售价", padding=10)
        cache_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        cache_frame.columnconfigure(1, weight=1)
        ttk.Label(cache_frame, text="导入新表格").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(cache_frame, textvariable=self.match_input_path).grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=3)
        ttk.Button(cache_frame, text="选择", command=self.choose_match_input).grid(row=0, column=2, sticky="ew", pady=3)
        ttk.Label(cache_frame, text="固定查找“规格编码”，匹配价格库商品编码，并在新表格末尾新增“售价”列。", foreground="#435064").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(6, 8)
        )
        self.match_button = ttk.Button(cache_frame, text="匹配售价并导出", command=self.start_price_match)
        self.match_button.grid(row=2, column=0, columnspan=2, sticky="ew", ipady=14, padx=(0, 8))
        cache_action_frame = ttk.Frame(cache_frame)
        cache_action_frame.grid(row=2, column=2, sticky="nsew")
        cache_action_frame.columnconfigure(0, weight=1)
        self.export_cache_button = ttk.Button(cache_action_frame, text="导出价格库", command=self.start_export_price_cache)
        self.export_cache_button.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        self.clear_cache_button = ttk.Button(cache_action_frame, text="清空价格库", command=self.start_clear_price_cache)
        self.clear_cache_button.grid(row=1, column=0, sticky="ew")

        sales_analysis_frame = ttk.LabelFrame(cache_row, text="销售利润分析表", padding=10)
        sales_analysis_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        sales_analysis_frame.columnconfigure(1, weight=1)
        ttk.Label(sales_analysis_frame, text="导入表格").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(sales_analysis_frame, textvariable=self.sales_analysis_import_path).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(0, 8),
            pady=3,
        )
        ttk.Button(sales_analysis_frame, text="选择", command=self.choose_sales_analysis_import).grid(row=0, column=2, sticky="ew", pady=3)
        ttk.Label(
            sales_analysis_frame,
            text="读取“内部订单号”“订单重量”“销售金额”“销售成本”，并保存到销售主题分析价格库。",
            foreground="#435064",
            wraplength=320,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 8))
        self.sales_analysis_import_button = ttk.Button(
            sales_analysis_frame,
            text="导入销售主题分析价格库",
            command=self.start_sales_analysis_import,
        )
        self.sales_analysis_import_button.grid(row=2, column=0, columnspan=3, sticky="ew", ipady=14)

        self.progress_bar = ttk.Progressbar(frame, variable=self.progress, maximum=100, mode="determinate")
        self.progress_bar.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(14, 0))

        ttk.Label(frame, textvariable=self.status, foreground="#1f7a8c", wraplength=590).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(10, 0)
        )

    def choose_input(self):
        path = filedialog.askopenfilename(
            title="选择表格文件",
            filetypes=[
                ("表格文件", "*.xlsx *.xlsm *.csv"),
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*"),
            ],
        )

        if not path:
            return

        self.input_path.set(path)
        self.output_path.set(default_output_path(path))

    def choose_output(self):
        initial = self.output_path.get() or default_output_path(self.input_path.get() or "导出.xlsx")
        extension = os.path.splitext(initial)[1].lower()
        filetypes = [("Excel 文件", "*.xlsx")]
        defaultextension = ".xlsx"

        if extension == ".csv":
            filetypes = [("CSV 文件", "*.csv")]
            defaultextension = ".csv"

        path = filedialog.asksaveasfilename(
            title="选择导出位置",
            initialfile=os.path.basename(initial),
            initialdir=os.path.dirname(initial) or os.getcwd(),
            defaultextension=defaultextension,
            filetypes=filetypes,
        )

        if path:
            self.output_path.set(path)

    def choose_cache_import(self):
        path = filedialog.askopenfilename(
            title="选择要导入价格库的表格文件",
            filetypes=[
                ("表格文件", "*.xlsx *.xlsm *.csv"),
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*"),
            ],
        )

        if path:
            self.cache_import_path.set(path)

    def choose_combo_cache_import(self):
        path = filedialog.askopenfilename(
            title="选择要导入组合价格库的表格文件",
            filetypes=[
                ("表格文件", "*.xlsx *.xlsm *.csv"),
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*"),
            ],
        )

        if path:
            self.combo_cache_import_path.set(path)

    def choose_match_input(self):
        path = filedialog.askopenfilename(
            title="选择要匹配售价的表格文件",
            filetypes=[
                ("表格文件", "*.xlsx *.xlsm *.csv"),
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*"),
            ],
        )

        if not path:
            return

        self.match_input_path.set(path)
        self.match_output_path.set(default_price_match_output_path(path))

    def choose_sales_analysis_import(self):
        path = filedialog.askopenfilename(
            title="选择要导入销售主题分析价格库的表格文件",
            filetypes=[
                ("表格文件", "*.xlsx *.xlsm *.csv"),
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*"),
            ],
        )

        if path:
            self.sales_analysis_import_path.set(path)

    def choose_match_output(self):
        initial = self.match_output_path.get() or default_price_match_output_path(self.match_input_path.get() or "匹配售价.xlsx")
        extension = os.path.splitext(initial)[1].lower()
        filetypes = [("Excel 文件", "*.xlsx")]
        defaultextension = ".xlsx"

        if extension == ".csv":
            filetypes = [("CSV 文件", "*.csv")]
            defaultextension = ".csv"

        path = filedialog.asksaveasfilename(
            title="选择匹配售价导出位置",
            initialfile=os.path.basename(initial),
            initialdir=os.path.dirname(initial) or os.getcwd(),
            defaultextension=defaultextension,
            filetypes=filetypes,
        )

        if path:
            self.match_output_path.set(path)

    def choose_rules_file(self):
        path = filedialog.askopenfilename(
            title="选择删除规则文件",
            initialdir=os.path.dirname(self.rules_path.get()) or app_base_dir(),
            filetypes=[
                ("JSON 规则文件", "*.json"),
                ("所有文件", "*.*"),
            ],
        )

        if path:
            self.rules_path.set(path)

    def create_default_rules_file(self):
        path = filedialog.asksaveasfilename(
            title="生成默认删除规则文件",
            initialdir=os.path.dirname(self.rules_path.get()) or app_base_dir(),
            initialfile="rules.json",
            defaultextension=".json",
            filetypes=[("JSON 规则文件", "*.json")],
        )

        if not path:
            return

        try:
            save_default_rules(path)
            self.rules_path.set(path)
            messagebox.showinfo("已生成", f"默认规则文件已生成：\n{path}")
        except Exception as error:
            messagebox.showerror("生成失败", str(error))

    def get_shipping_fees(self):
        fees = []

        for (label, _), variable in zip(SHIPPING_FEE_RULES, self.shipping_fee_vars):
            text = variable.get().strip()
            try:
                fees.append(float(text))
            except ValueError:
                raise ValueError(f"快递费规则“{label}”必须填写数字。")

        return fees

    def get_price_rules_from_vars(self, threshold_var, low_margin_var, high_margin_var, label):
        try:
            threshold = float(threshold_var.get().strip())
            low_margin = float(low_margin_var.get().strip())
            high_margin = float(high_margin_var.get().strip())
        except ValueError:
            raise ValueError(f"{label}里的分界值和毛利率都必须填写数字。")

        if not (0 <= low_margin < 1) or not (0 <= high_margin < 1):
            raise ValueError(f"{label}里的两个毛利率都必须大于等于 0，并且小于 1。")

        return threshold, low_margin, high_margin

    def get_price_rules(self):
        return self.get_price_rules_from_vars(
            self.price_threshold,
            self.low_price_margin,
            self.high_price_margin,
            "普通商品售价规则",
        )

    def get_combo_price_rules(self):
        return self.get_price_rules_from_vars(
            self.combo_price_threshold,
            self.combo_low_price_margin,
            self.combo_high_price_margin,
            "组合售价规则",
        )

    def start_processing(self):
        input_path = self.input_path.get().strip()
        output_path = self.output_path.get().strip()

        if not input_path or not os.path.exists(input_path):
            messagebox.showwarning("缺少文件", "请先选择要导入的表格文件。")
            return

        if not output_path:
            self.output_path.set(default_output_path(input_path))
            output_path = self.output_path.get()

        try:
            shipping_fees = self.get_shipping_fees()
            price_rules = self.get_price_rules()
            delete_rules = load_delete_rules(self.rules_path.get().strip())
            price_match_column = self.price_match_column.get().strip()
            if not price_match_column:
                raise ValueError("售价匹配列名不能为空。")
            price_cache = load_price_cache()
        except ValueError as error:
            messagebox.showwarning("规则错误", str(error))
            return
        except Exception as error:
            messagebox.showwarning("规则错误", f"规则或售价保存文件读取失败：{error}")
            return

        self._set_buttons_state("disabled")
        self.progress_bar.config(mode="indeterminate")
        self.progress_bar.start(10)
        self.progress.set(0)
        self.status.set("正在处理，请稍候...")

        thread = threading.Thread(
            target=self._process_in_thread,
            args=(input_path, output_path, shipping_fees, price_rules, delete_rules, price_match_column, price_cache),
            daemon=True,
        )
        thread.start()

    def start_price_cache_import(self):
        self._start_price_cache_import(
            self.cache_import_path.get().strip(),
            PRODUCT_CODE_COLUMN_NAME,
            WEIGHT_COLUMN_NAME,
            COST_PRICE_COLUMN_NAME,
            "价格库",
            "请先选择要导入价格库的表格文件。",
            self.get_price_rules,
        )

    def start_combo_price_cache_import(self):
        self._start_price_cache_import(
            self.combo_cache_import_path.get().strip(),
            COMBO_PRODUCT_CODE_COLUMN_NAME,
            COMBO_WEIGHT_COLUMN_NAME,
            COMBO_COST_PRICE_COLUMN_NAME,
            "组合价格库",
            "请先选择要导入组合价格库的表格文件。",
            self.get_combo_price_rules,
        )

    def _start_price_cache_import(
        self,
        input_path,
        product_code_column,
        weight_column,
        cost_price_column,
        progress_label,
        missing_file_message,
        price_rules_getter,
    ):
        if not input_path or not os.path.exists(input_path):
            messagebox.showwarning("缺少文件", missing_file_message)
            return

        try:
            shipping_fees = self.get_shipping_fees()
            price_rules = price_rules_getter()
            price_cache = load_price_cache()
        except ValueError as error:
            messagebox.showwarning("规则错误", str(error))
            return
        except Exception as error:
            messagebox.showwarning("读取失败", f"价格库读取失败：{error}")
            return

        self._set_buttons_state("disabled")
        self.progress_bar.config(mode="indeterminate")
        self.progress_bar.start(10)
        self.progress.set(0)
        self.status.set(f"正在计算并导入{progress_label}，请稍候...")

        thread = threading.Thread(
            target=self._import_price_cache_in_thread,
            args=(input_path, shipping_fees, price_rules, price_cache, product_code_column, weight_column, cost_price_column, progress_label),
            daemon=True,
        )
        thread.start()

    def start_price_match(self):
        input_path = self.match_input_path.get().strip()
        output_path = self.match_output_path.get().strip()

        if not input_path or not os.path.exists(input_path):
            messagebox.showwarning("缺少文件", "请先选择要匹配售价的新表格文件。")
            return

        if not output_path:
            self.match_output_path.set(default_price_match_output_path(input_path))
            output_path = self.match_output_path.get()

        try:
            price_cache = load_price_cache()
        except Exception as error:
            messagebox.showwarning("读取失败", f"价格库读取失败：{error}")
            return

        self._set_buttons_state("disabled")
        self.progress_bar.config(mode="indeterminate")
        self.progress_bar.start(10)
        self.progress.set(0)
        self.status.set("正在从价格库匹配售价，请稍候...")

        thread = threading.Thread(
            target=self._match_prices_in_thread,
            args=(input_path, output_path, price_cache),
            daemon=True,
        )
        thread.start()

    def start_sales_analysis_import(self):
        input_path = self.sales_analysis_import_path.get().strip()

        if not input_path or not os.path.exists(input_path):
            messagebox.showwarning("缺少文件", "请先选择要导入销售主题分析价格库的表格文件。")
            return

        try:
            shipping_fees = self.get_shipping_fees()
            sales_cache = load_sales_analysis_cache()
        except ValueError as error:
            messagebox.showwarning("规则错误", str(error))
            return
        except Exception as error:
            messagebox.showwarning("读取失败", f"销售主题分析价格库读取失败：{error}")
            return

        self._set_buttons_state("disabled")
        self.progress_bar.config(mode="indeterminate")
        self.progress_bar.start(10)
        self.progress.set(0)
        self.status.set("正在导入销售主题分析价格库，请稍候...")

        thread = threading.Thread(
            target=self._import_sales_analysis_in_thread,
            args=(input_path, sales_cache, shipping_fees),
            daemon=True,
        )
        thread.start()

    def start_export_price_cache(self):
        path = filedialog.asksaveasfilename(
            title="选择价格库导出位置",
            initialfile="价格库.xlsx",
            initialdir=os.path.expanduser("~/Desktop"),
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )

        if not path:
            return

        try:
            price_cache = load_price_cache()
        except Exception as error:
            messagebox.showwarning("读取失败", f"价格库读取失败：{error}")
            return

        if not price_cache:
            messagebox.showwarning("价格库为空", "当前价格库没有可导出的售价记录。")
            return

        self._set_buttons_state("disabled")
        self.progress_bar.config(mode="indeterminate")
        self.progress_bar.start(10)
        self.progress.set(0)
        self.status.set("正在导出价格库，请稍候...")

        thread = threading.Thread(
            target=self._export_price_cache_in_thread,
            args=(path, price_cache),
            daemon=True,
        )
        thread.start()

    def start_clear_price_cache(self):
        try:
            price_cache = load_price_cache()
        except Exception as error:
            messagebox.showwarning("读取失败", f"价格库读取失败：{error}")
            return

        if not price_cache:
            messagebox.showinfo("价格库为空", "当前价格库没有已保存的售价记录。")
            return

        confirmed = messagebox.askyesno(
            "确认清空价格库",
            f"当前价格库共有 {len(price_cache)} 条售价记录。\n\n确定要全部清空吗？此操作不能撤销。",
        )
        if not confirmed:
            return

        try:
            save_price_cache({})
        except Exception as error:
            messagebox.showerror("清空失败", f"价格库清空失败：{error}")
            return

        self.progress_bar.stop()
        self.progress_bar.config(mode="determinate")
        self.progress.set(100)
        self.status.set("价格库已清空。")
        messagebox.showinfo("清空完成", f"价格库已清空，共删除 {len(price_cache)} 条售价记录。")

    def _match_prices_in_thread(self, input_path, output_path, price_cache):
        try:
            extension = os.path.splitext(input_path)[1].lower()

            if extension == ".csv":
                summary = append_cached_prices_csv(input_path, output_path, self._set_progress, price_cache)
            elif extension in (".xlsx", ".xlsm"):
                summary = append_cached_prices_xlsx(input_path, output_path, self._set_progress, price_cache)
            else:
                raise ValueError("暂不支持该文件格式，请使用 .xlsx、.xlsm 或 .csv。")

            matched = sum(item[1] for item in summary)
            unmatched = sum(item[2] for item in summary)
            self.after(0, lambda: self._finish_match_success(output_path, matched, unmatched))
        except Exception as error:
            message = str(error)
            self.after(0, lambda: self._finish_error(message))

    def _export_price_cache_in_thread(self, output_path, price_cache):
        try:
            exported = export_price_cache_xlsx(output_path, price_cache)
            self.after(0, lambda: self._finish_export_cache_success(output_path, exported))
        except Exception as error:
            message = str(error)
            self.after(0, lambda: self._finish_error(message))

    def _import_sales_analysis_in_thread(self, input_path, sales_cache, shipping_fees):
        try:
            extension = os.path.splitext(input_path)[1].lower()

            if extension == ".csv":
                summary = import_sales_analysis_to_cache_csv(input_path, self._set_progress, sales_cache, shipping_fees)
            elif extension in (".xlsx", ".xlsm"):
                summary = import_sales_analysis_to_cache_xlsx(input_path, self._set_progress, sales_cache, shipping_fees)
            else:
                raise ValueError("暂不支持该文件格式，请使用 .xlsx、.xlsm 或 .csv。")

            save_sales_analysis_cache(sales_cache)
            added = sum(item[1] for item in summary)
            updated = sum(item[2] for item in summary)
            skipped = sum(item[3] for item in summary)
            self.after(0, lambda: self._finish_sales_analysis_import_success(added, updated, skipped))
        except Exception as error:
            message = str(error)
            self.after(0, lambda: self._finish_error(message))

    def _import_price_cache_in_thread(
        self,
        input_path,
        shipping_fees,
        price_rules,
        price_cache,
        product_code_column,
        weight_column,
        cost_price_column,
        progress_label,
    ):
        try:
            extension = os.path.splitext(input_path)[1].lower()

            if extension == ".csv":
                summary = import_prices_to_cache_csv(
                    input_path,
                    self._set_progress,
                    shipping_fees,
                    price_rules,
                    price_cache,
                    product_code_column,
                    weight_column,
                    cost_price_column,
                    progress_label,
                )
            elif extension in (".xlsx", ".xlsm"):
                summary = import_prices_to_cache_xlsx(
                    input_path,
                    self._set_progress,
                    shipping_fees,
                    price_rules,
                    price_cache,
                    product_code_column,
                    weight_column,
                    cost_price_column,
                    progress_label,
                )
            else:
                raise ValueError("暂不支持该文件格式，请使用 .xlsx、.xlsm 或 .csv。")

            save_price_cache(price_cache)
            added = sum(item[1] for item in summary)
            updated = sum(item[2] for item in summary)
            skipped = sum(item[3] for item in summary)
            self.after(0, lambda: self._finish_cache_import_success(added, updated, skipped))
        except Exception as error:
            message = str(error)
            self.after(0, lambda: self._finish_error(message))

    def _process_in_thread(self, input_path, output_path, shipping_fees, price_rules, delete_rules, price_match_column, price_cache):
        try:
            extension = os.path.splitext(input_path)[1].lower()

            if extension == ".csv":
                summary = process_csv(
                    input_path,
                    output_path,
                    "after",
                    self._set_progress,
                    shipping_fees,
                    price_rules,
                    delete_rules,
                    price_match_column,
                    price_cache,
                )
            elif extension in (".xlsx", ".xlsm"):
                summary = process_xlsx(
                    input_path,
                    output_path,
                    "after",
                    self._set_progress,
                    shipping_fees,
                    price_rules,
                    delete_rules,
                    price_match_column,
                    price_cache,
                )
            else:
                raise ValueError("暂不支持该文件格式，请使用 .xlsx、.xlsm 或 .csv。")

            save_price_cache(price_cache)
            deleted = sum(item[2] for item in summary)
            kept = sum(item[1] for item in summary)
            matched = sum(item[3] for item in summary)
            saved = sum(item[4] for item in summary)
            self.after(0, lambda: self._finish_success(output_path, kept, deleted, matched, saved))
        except Exception as error:
            message = str(error)
            self.after(0, lambda: self._finish_error(message))

    def _set_status(self, text):
        self.after(0, lambda: self.status.set(text))

    def _set_progress(self, text, percent=None):
        def update():
            self.status.set(text)
            if percent is not None:
                self.progress_bar.stop()
                self.progress_bar.config(mode="determinate")
                self.progress.set(percent)
            elif str(self.progress_bar.cget("mode")) != "indeterminate":
                self.progress_bar.config(mode="indeterminate")
                self.progress_bar.start(10)

        self.after(0, update)

    def _set_buttons_state(self, state):
        self.cache_import_button.config(state=state)
        self.combo_cache_import_button.config(state=state)
        self.match_button.config(state=state)
        self.export_cache_button.config(state=state)
        self.clear_cache_button.config(state=state)
        self.sales_analysis_import_button.config(state=state)

    def _finish_success(self, output_path, kept, deleted, matched, saved):
        self._set_buttons_state("normal")
        self.progress_bar.stop()
        self.progress_bar.config(mode="determinate")
        self.progress.set(100)
        self.status.set(f"处理完成：保留 {kept} 行，删除 {deleted} 行，匹配保存售价 {matched} 行，新增保存售价 {saved} 行。")
        messagebox.showinfo(
            "处理完成",
            f"已导出：\n{output_path}\n\n匹配保存售价：{matched} 行\n新增保存售价：{saved} 行",
        )

    def _finish_match_success(self, output_path, matched, unmatched):
        self._set_buttons_state("normal")
        self.progress_bar.stop()
        self.progress_bar.config(mode="determinate")
        self.progress.set(100)
        self.status.set(f"售价匹配完成：匹配 {matched} 行，未匹配 {unmatched} 行。")
        messagebox.showinfo(
            "匹配完成",
            f"已导出：\n{output_path}\n\n匹配售价：{matched} 行\n未匹配：{unmatched} 行",
        )

    def _finish_cache_import_success(self, added, updated, skipped):
        self._set_buttons_state("normal")
        self.progress_bar.stop()
        self.progress_bar.config(mode="determinate")
        self.progress.set(100)
        self.status.set(f"价格库导入完成：新增 {added} 条，更新 {updated} 条，跳过 {skipped} 行。")
        messagebox.showinfo(
            "导入完成",
            f"价格库已更新。\n\n新增：{added} 条\n更新：{updated} 条\n跳过：{skipped} 行",
        )

    def _finish_sales_analysis_import_success(self, added, updated, skipped):
        self._set_buttons_state("normal")
        self.progress_bar.stop()
        self.progress_bar.config(mode="determinate")
        self.progress.set(100)
        self.status.set(f"销售主题分析价格库导入完成：新增 {added} 条，更新 {updated} 条，跳过 {skipped} 行。")
        messagebox.showinfo(
            "导入完成",
            f"销售主题分析价格库已更新。\n\n新增：{added} 条\n更新：{updated} 条\n跳过：{skipped} 行",
        )

    def _finish_export_cache_success(self, output_path, exported):
        self._set_buttons_state("normal")
        self.progress_bar.stop()
        self.progress_bar.config(mode="determinate")
        self.progress.set(100)
        self.status.set(f"价格库导出完成：{exported} 条记录。")
        messagebox.showinfo(
            "导出完成",
            f"价格库已导出：\n{output_path}\n\n共导出：{exported} 条记录",
        )

    def _finish_error(self, message):
        self._set_buttons_state("normal")
        self.progress_bar.stop()
        self.progress_bar.config(mode="determinate")
        self.progress.set(0)
        self.status.set(f"处理失败：{message}")
        messagebox.showerror("处理失败", message)
