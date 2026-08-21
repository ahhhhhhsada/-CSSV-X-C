import os
import re
import unicodedata
import pandas as pd
import openpyxl
from tkinter import (
    Tk, filedialog, messagebox, Button, Label, Entry, Frame, ttk,
    StringVar, Radiobutton, IntVar, Text, END
)
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ===================== 全局核心配置 =====================
PART_COL = "Part Number"
NAME_COL = "Part Name"
QTY_COL = "Qty"
LOSS_RATE = 1.2

AVL_CAT_PART = "Number (Catalog Part)"
AVL_MF_PART = "Number (Manufacturer Part)"
AVL_MF_NAME = "Name (Manufacturer)"
AVL_AVAIL = "Availability (Manufacturer Part)"
AVL_SOURCE_STATUS = "Sourcing Status (AML Link)"
AVL_FLAG_COL = "AVL匹配状态"
NO_AVL_TEXT = "【无AVL匹配信息】"
SEARCH_INPUT_COL = "输入料号"

SUPPLIER_SHEET_NAME = "Material Preparation"
SUPPLIER_HEADERS = [
    "Part Number", "MPN", "Manufacturer", "Demand Qty",
    "Purchase Qty", "Unit price", "Total price", "Lead Time"
]

COL_MAPPING = {
    PART_COL: ["part number", "partnumber", "part_no", "part no", "料号", "物料编码", "part num"],
    NAME_COL: ["part name", "partname", "name", "物料名称", "料名"],
    QTY_COL: ["qty", "quantity", "qty.", "数量", "用量", "pcs"]
}

# ===================== 公共工具函数 =====================
def clean_col_name(name):
    return re.sub(r"\s+", "", str(name)).strip().lower()


def standardize_pn(pn):
    """统一料号格式，兼容Excel隐藏空格、零宽字符、全角字符和特殊连字符。"""
    if pd.isna(pn):
        return ""
    value = unicodedata.normalize("NFKC", str(pn))
    value = value.replace("–", "-").replace("—", "-").replace("−", "-").replace("‐", "-")
    value = re.sub(r"[\s\t\n\r\u00A0\u200B-\u200D\uFEFF]+", "", value)
    return value.strip().upper()


def display_width(text):
    return sum(2 if unicodedata.east_asian_width(c) in ("F", "W") else 1 for c in str(text or ""))


def auto_fit_columns(ws, min_w=8, max_w=50, padding=3):
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_width = 0
        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(cell, openpyxl.cell.cell.MergedCell):
                continue
            if cell.value is not None:
                max_width = max(max_width, display_width(cell.value))
        ws.column_dimensions[col_letter].width = max(min_w, min(max_width * 1.1 + padding, max_w))


HEADER_FILL = PatternFill("solid", fgColor="70AD47")
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
ZEBRA_LIGHT_FILL = PatternFill("solid", fgColor="E2EFDA")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")
DATA_FONT = Font(name="微软雅黑", size=10, color="375623")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN_GRAY_SIDE = Side(style="thin", color="D9D9D9")
FULL_BORDER = Border(top=THIN_GRAY_SIDE, bottom=THIN_GRAY_SIDE, left=THIN_GRAY_SIDE, right=THIN_GRAY_SIDE)


def beautify_worksheet(ws):
    if ws.max_row < 1 or ws.max_column < 1:
        return
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}1"
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER_ALIGN
        cell.border = FULL_BORDER
    ws.row_dimensions[1].height = 28
    for row_idx in range(2, ws.max_row + 1):
        row_fill = ZEBRA_LIGHT_FILL if (row_idx - 2) % 2 == 0 else WHITE_FILL
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = DATA_FONT
            cell.fill = row_fill
            cell.border = FULL_BORDER
            cell.alignment = CENTER_ALIGN
        ws.row_dimensions[row_idx].height = 22
    auto_fit_columns(ws)


# ===================== BOM 汇总 =====================
def clean_bom_data(df):
    df_cols = [str(c).strip().lower() for c in df.columns]
    target_cols = {PART_COL: None, NAME_COL: None, QTY_COL: None}
    for target, alias_list in COL_MAPPING.items():
        for idx, col in enumerate(df_cols):
            if col in alias_list or any(alias in col for alias in alias_list):
                target_cols[target] = df.columns[idx]
                break
    clean_df = pd.DataFrame()
    for target_col in [PART_COL, NAME_COL, QTY_COL]:
        original_col = target_cols[target_col]
        clean_df[target_col] = df[original_col] if original_col in df.columns else ""
    clean_df[PART_COL] = clean_df[PART_COL].apply(standardize_pn)
    clean_df[QTY_COL] = pd.to_numeric(clean_df[QTY_COL], errors="coerce").fillna(0)
    return clean_df


def load_boms(file_paths):
    products = {}
    for path in file_paths:
        try:
            name = os.path.splitext(os.path.basename(path))[0]
            products[name] = clean_bom_data(pd.read_excel(path))
        except Exception as exc:
            messagebox.showwarning("文件读取警告", f"{os.path.basename(path)}\n错误：{exc}")
    return products


def calculate(products, production_counts):
    """
    汇总BOM需求，并让“总BOM usage”按照研发原始BOM的Part Number顺序输出。

    多个BOM的排序规则：
    1. 先按用户选择/加载BOM文件的顺序；
    2. 再按每个研发BOM中的原始行顺序；
    3. 重复Part Number采用第一次出现的位置；
    4. 后续BOM中新出现的Part Number依次追加。
    """
    all_details = []
    original_part_order = {}
    next_order = 0

    for prod_name, df in products.items():
        cnt = production_counts.get(prod_name, 0)
        if cnt <= 0:
            continue

        temp = df.copy()
        temp["产品名称"] = prod_name
        temp["生产数量"] = cnt
        temp["无备损总需求"] = (temp[QTY_COL] * cnt).round().astype(int)

        # 在汇总前记录研发BOM的原始料号顺序，避免pivot_table自动按料号排序。
        for part_number in temp[PART_COL].tolist():
            key = standardize_pn(part_number)
            if key and key not in original_part_order:
                original_part_order[key] = next_order
                next_order += 1

        all_details.append(temp)

    if not all_details:
        return None, None, None

    detail_df = pd.concat(all_details, ignore_index=True)
    wide_df = detail_df.pivot_table(
        index=[PART_COL, NAME_COL], columns="产品名称", values="无备损总需求",
        aggfunc="sum", fill_value=0, sort=False
    ).reset_index()
    wide_df["总需求数量"] = wide_df.drop([PART_COL, NAME_COL], axis=1).sum(axis=1).astype(int)
    wide_df[f"备损后总需求(×{LOSS_RATE})"] = (wide_df["总需求数量"] * LOSS_RATE).round().astype(int)

    # 恢复研发初始BOM的Part Number顺序；辅助列不会显示或导出。
    fallback_order = len(original_part_order)
    wide_df["__研发BOM顺序"] = wide_df[PART_COL].apply(
        lambda value: original_part_order.get(standardize_pn(value), fallback_order)
    )
    wide_df = wide_df.sort_values("__研发BOM顺序", kind="stable").drop(
        columns=["__研发BOM顺序"]
    ).reset_index(drop=True)

    dist_df = detail_df.groupby(
        [PART_COL, NAME_COL, "产品名称"], as_index=False, sort=False
    )["无备损总需求"].sum()
    dist_df.rename(columns={"无备损总需求": "单个产品无备损总需求"}, inplace=True)
    dist_df["单个产品无备损总需求"] = dist_df["单个产品无备损总需求"].astype(int)
    return detail_df, dist_df, wide_df


# ===================== AVL =====================
def load_avl(file_path):
    try:
        avl_df = pd.read_excel(file_path, header=0)
        col_map = {clean_col_name(c): c for c in avl_df.columns}
        req_keys = [
            "number(catalogpart)", "number(manufacturerpart)", "name(manufacturer)",
            "availability(manufacturerpart)", "sourcingstatus(amllink)"
        ]
        missing = [key for key in req_keys if key not in col_map]
        if missing:
            messagebox.showerror("AVL字段缺失", f"缺失字段：{','.join(missing)}\n\n请检查AVL表头是否在第1行。")
            return None
        use_df = avl_df[[col_map[key] for key in req_keys]].copy()
        use_df.columns = [AVL_CAT_PART, AVL_MF_PART, AVL_MF_NAME, AVL_AVAIL, AVL_SOURCE_STATUS]
        use_df[AVL_CAT_PART] = use_df[AVL_CAT_PART].apply(standardize_pn)
        use_df[AVL_MF_PART] = use_df[AVL_MF_PART].apply(standardize_pn)
        return use_df.drop_duplicates(keep="first")
    except Exception as exc:
        messagebox.showerror("读取AVL失败", str(exc))
        return None


def merge_avl_info(main_df, avl_df):
    main_df = main_df.copy()
    main_df[PART_COL] = main_df[PART_COL].apply(standardize_pn)
    result = pd.merge(main_df, avl_df, left_on=PART_COL, right_on=AVL_CAT_PART, how="left")
    fill_cols = [AVL_MF_PART, AVL_MF_NAME, AVL_AVAIL, AVL_SOURCE_STATUS]
    for col in fill_cols:
        result[col] = result[col].fillna(NO_AVL_TEXT).apply(
            lambda value: NO_AVL_TEXT if str(value).strip() == "" else value
        )
    result[AVL_FLAG_COL] = result[AVL_MF_PART].apply(
        lambda value: "存在AVL匹配信息" if value != NO_AVL_TEXT else NO_AVL_TEXT
    )
    result.drop(columns=[AVL_CAT_PART], errors="ignore", inplace=True)
    return result[list(main_df.columns) + fill_cols + [AVL_FLAG_COL]]


def parse_part_numbers(raw_text):
    """支持换行、空格、逗号、中英文分号分隔，并按首次输入顺序去重。"""
    values = re.split(r"[\s,，;；]+", raw_text.strip())
    values = [standardize_pn(v) for v in values if standardize_pn(v)]
    return list(dict.fromkeys(values))


def format_part_numbers_for_copy(raw_text, suffix="XX"):
    """
    将批量Part Number转换为竖线拼接格式。
    示例：1CAP005357|1CAP007179|1CAP007951|1CNR005500|XX
    """
    part_numbers = parse_part_numbers(raw_text)
    clean_suffix = standardize_pn(suffix) or "XX"
    if not part_numbers:
        return ""
    return "|".join(part_numbers + [clean_suffix])


def search_avl_parts(part_numbers, avl_df):
    """
    批量搜索AVL：
    1. 同时匹配 Catalog Part 和 Manufacturer Part；
    2. 一个输入料号对应多条AVL时全部保留；
    3. 未匹配输入料号仍保留并标记；
    4. 按用户首次输入顺序输出。
    """
    result_rows = []
    for order, input_part in enumerate(part_numbers, start=1):
        key = standardize_pn(input_part)
        cat_mask = avl_df[AVL_CAT_PART].apply(standardize_pn).eq(key)
        mf_mask = avl_df[AVL_MF_PART].apply(standardize_pn).eq(key)
        matched = avl_df[cat_mask | mf_mask].copy()

        if matched.empty:
            result_rows.append({
                SEARCH_INPUT_COL: input_part,
                "匹配字段": NO_AVL_TEXT,
                AVL_CAT_PART: NO_AVL_TEXT,
                AVL_MF_PART: NO_AVL_TEXT,
                AVL_MF_NAME: NO_AVL_TEXT,
                AVL_AVAIL: NO_AVL_TEXT,
                AVL_SOURCE_STATUS: NO_AVL_TEXT,
                AVL_FLAG_COL: NO_AVL_TEXT,
                "输入顺序": order
            })
            continue

        for _, row in matched.iterrows():
            matched_by = []
            if standardize_pn(row[AVL_CAT_PART]) == key:
                matched_by.append("Catalog Part")
            if standardize_pn(row[AVL_MF_PART]) == key:
                matched_by.append("Manufacturer Part")
            result_rows.append({
                SEARCH_INPUT_COL: input_part,
                "匹配字段": " + ".join(matched_by),
                AVL_CAT_PART: row[AVL_CAT_PART],
                AVL_MF_PART: row[AVL_MF_PART],
                AVL_MF_NAME: row[AVL_MF_NAME],
                AVL_AVAIL: row[AVL_AVAIL],
                AVL_SOURCE_STATUS: row[AVL_SOURCE_STATUS],
                AVL_FLAG_COL: "存在AVL匹配信息",
                "输入顺序": order
            })

    result = pd.DataFrame(result_rows)
    return result.sort_values("输入顺序").drop(columns=["输入顺序"]).reset_index(drop=True)


def safe_sheet_name(name, used_names):
    base = re.sub(r"[\\/*?:\[\]]", "_", str(name))[:31] or "Sheet"
    candidate = base
    index = 1
    while candidate in used_names:
        suffix = f"_{index}"
        candidate = base[:31 - len(suffix)] + suffix
        index += 1
    used_names.add(candidate)
    return candidate


def add_supplier_material_sheet(wb, wide_df):
    """
    在工作簿最后新增供应商Material Preparation模板。
    Part Number取自总BOM usage；Demand Qty取备损后总需求(×1.2)。
    其余供应商填写字段保持为空。
    """
    if SUPPLIER_SHEET_NAME in wb.sheetnames:
        del wb[SUPPLIER_SHEET_NAME]

    ws = wb.create_sheet(SUPPLIER_SHEET_NAME)
    demand_col = f"备损后总需求(×{LOSS_RATE})"
    if PART_COL not in wide_df.columns or demand_col not in wide_df.columns:
        raise ValueError(f"总BOM usage缺少必要字段：{PART_COL} 或 {demand_col}")

    # 完全按照用户模板的8列表头和主要样式创建。
    template_header_fill = PatternFill("solid", fgColor="92D050")
    template_align = Alignment(horizontal="center", vertical="center")
    template_font = Font(name="Calibri", size=11)

    for col_idx, header in enumerate(SUPPLIER_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = template_header_fill
        cell.font = template_font
        cell.alignment = template_align

    # 每个Part Number只输出一次，并保持总BOM usage中的原始顺序。
    material_df = wide_df[[PART_COL, demand_col]].copy()
    material_df[PART_COL] = material_df[PART_COL].apply(standardize_pn)
    material_df = material_df[material_df[PART_COL] != ""]
    material_df = material_df.drop_duplicates(subset=[PART_COL], keep="first").reset_index(drop=True)

    for row_idx, row in material_df.iterrows():
        excel_row = row_idx + 2
        ws.cell(excel_row, 1, row[PART_COL])
        ws.cell(excel_row, 4, int(row[demand_col]) if pd.notna(row[demand_col]) else 0)
        for col_idx in range(1, 9):
            ws.cell(excel_row, col_idx).font = template_font
            ws.cell(excel_row, col_idx).alignment = template_align

    # 保留原模板的冻结窗格、筛选和列宽。
    ws.freeze_panes = "E1"
    ws.auto_filter.ref = "A1:H1"
    template_widths = {
        "A": 18.453125, "B": 23.54296875, "C": 33.453125,
        "D": 13.0, "E": 21.1796875, "F": 16.453125,
        "G": 13.0, "H": 15.6328125
    }
    for col_letter, width in template_widths.items():
        ws.column_dimensions[col_letter].width = width

    # 强制放到最后一个Sheet。
    wb._sheets.remove(ws)
    wb._sheets.append(ws)
    return len(material_df)


def export_frames(path, frames, supplier_source_df=None):
    used = set()
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, df in frames:
            df.to_excel(writer, sheet_name=safe_sheet_name(sheet_name, used), index=False)
    wb = load_workbook(path)
    for ws in wb.worksheets:
        beautify_worksheet(ws)
    supplier_count = 0
    if supplier_source_df is not None:
        supplier_count = add_supplier_material_sheet(wb, supplier_source_df)
    wb.save(path)
    return supplier_count


# ===================== GUI =====================
class BOMAVLApp(Tk):
    def __init__(self):
        super().__init__()
        self.title("采购备料一体化工具（BOM汇总+AVL匹配）")
        self.geometry("1420x980")
        self.files = []
        self.avl_df = None
        self.products = {}
        self.entries = {}
        self.detail = None
        self.dist = None
        self.wide_df = None
        self.avl_search_result = None
        self.bom_search_result = None

        Label(self, text="一、批量选择多个产品研发BOM", font=("微软雅黑", 12, "bold")).pack(pady=4)
        Button(self, text="多选Excel BOM文件", command=self.load_bom_files, width=24, bg="#2196F3", fg="white").pack(pady=2)
        self.bom_status = Label(self, text=f"未加载任何BOM文件 | 备损率{LOSS_RATE}倍", font=("微软雅黑", 11))
        self.bom_status.pack(pady=2)

        Label(self, text="二、选择系统AVL基准表", font=("微软雅黑", 12, "bold")).pack(pady=5)
        self.avl_path_var = StringVar(value="未选择AVL文件")
        Entry(self, textvariable=self.avl_path_var, width=80, font=("微软雅黑", 9)).pack(pady=2)
        Button(self, text="选择AVL表", command=self.load_avl_file, width=24, bg="#009688", fg="white").pack(pady=2)

        self.input_frame = Frame(self)
        self.input_frame.pack(pady=5)
        self.calc_btn = Button(self, text="三、一键汇总计算", command=self.do_calc, state="disabled", width=26,
                               bg="#4CAF50", fg="white", font=("微软雅黑", 11, "bold"))
        self.calc_btn.pack(pady=3)

        export_mode_frame = Frame(self)
        export_mode_frame.pack(pady=3)
        Label(export_mode_frame, text="导出模式：", font=("微软雅黑", 11, "bold")).pack(side="left", padx=5)
        self.export_mode = IntVar(value=2)
        Radiobutton(export_mode_frame, text="仅导出备料Usage", variable=self.export_mode, value=1).pack(side="left", padx=8)
        Radiobutton(export_mode_frame, text="导出Usage+AVL匹配", variable=self.export_mode, value=2).pack(side="left", padx=8)
        self.export_btn = Button(self, text="四、导出Excel", command=self.do_export, state="disabled", width=26,
                                 bg="#FF9800", fg="white", font=("微软雅黑", 11, "bold"))
        self.export_btn.pack(pady=3)

        switch_frame = Frame(self)
        switch_frame.pack(pady=4)
        Button(switch_frame, text="总BOM usage", command=lambda: self.show_table(self.wide_df), width=18, bg="#9C27B0", fg="white").grid(row=0, column=0, padx=4)
        Button(switch_frame, text="分产品明细", command=lambda: self.show_table(self.dist), width=14, bg="#9C27B0", fg="white").grid(row=0, column=1, padx=4)
        Button(switch_frame, text="各产品计算明细", command=lambda: self.show_table(self.detail), width=18, bg="#9C27B0", fg="white").grid(row=0, column=2, padx=4)

        search_frame = Frame(self)
        search_frame.pack(fill="x", padx=10, pady=4)
        left = Frame(search_frame)
        left.pack(side="left", padx=8)
        Label(
            left,
            text="批量搜索总BOM usage（换行/逗号/空格分隔）",
            font=("微软雅黑", 10, "bold")
        ).grid(row=0, column=0, columnspan=3)
        self.search_box = Text(left, font=("Consolas", 10), width=48, height=4)
        self.search_box.grid(row=1, column=0, columnspan=3, pady=2)
        Button(left, text="批量检索BOM", command=self.do_search, width=14, bg="#3F51B5", fg="white").grid(row=2, column=0, padx=3)
        Button(left, text="下载BOM搜索结果", command=self.download_bom_search, width=16, bg="#FF9800", fg="white").grid(row=2, column=1, padx=3)
        Button(left, text="清空", command=lambda: self.search_box.delete("1.0", END), width=10).grid(row=2, column=2, padx=3)

        right = Frame(search_frame)
        right.pack(side="left", padx=25)
        Label(right, text="批量AVL料号搜索（换行/逗号/空格分隔）", font=("微软雅黑", 10, "bold")).grid(row=0, column=0, columnspan=3)
        self.avl_search_text = Text(right, width=58, height=4, font=("Consolas", 10))
        self.avl_search_text.grid(row=1, column=0, columnspan=3, pady=2)
        Button(right, text="批量搜索AVL", command=self.do_avl_search, width=14, bg="#607D8B", fg="white").grid(row=2, column=0, padx=3)
        Button(right, text="下载搜索结果", command=self.download_avl_search, width=14, bg="#FF9800", fg="white").grid(row=2, column=1, padx=3)
        Button(right, text="清空料号", command=lambda: self.avl_search_text.delete("1.0", END), width=12).grid(row=2, column=2, padx=3)

        # Part Number拼接复制工具
        format_frame = Frame(self, bd=1, relief="groove")
        format_frame.pack(fill="x", padx=18, pady=5)
        Label(
            format_frame,
            text="Part Number拼接复制（输出示例：PN1|PN2|PN3|XX）",
            font=("微软雅黑", 10, "bold")
        ).grid(row=0, column=0, columnspan=6, pady=3)

        Label(format_frame, text="输入料号：").grid(row=1, column=0, padx=4, sticky="ne")
        self.format_input_text = Text(format_frame, width=65, height=3, font=("Consolas", 10))
        self.format_input_text.grid(row=1, column=1, columnspan=3, padx=4, pady=2, sticky="w")

        Label(format_frame, text="结尾标识：").grid(row=1, column=4, padx=4, sticky="e")
        self.format_suffix_var = StringVar(value="XX")
        Entry(format_frame, textvariable=self.format_suffix_var, width=10, font=("Consolas", 10)).grid(row=1, column=5, padx=4, sticky="w")

        Button(format_frame, text="生成拼接格式", command=self.generate_copy_format, width=14, bg="#3F51B5", fg="white").grid(row=2, column=0, padx=4, pady=3)
        Button(format_frame, text="复制结果", command=self.copy_formatted_result, width=12, bg="#009688", fg="white").grid(row=2, column=1, padx=4, pady=3, sticky="w")
        Button(format_frame, text="导入总BOM料号", command=self.load_wide_parts_to_formatter, width=15, bg="#795548", fg="white").grid(row=2, column=2, padx=4, pady=3)
        Button(format_frame, text="清空", command=self.clear_formatter, width=10).grid(row=2, column=3, padx=4, pady=3, sticky="w")

        Label(format_frame, text="输出：").grid(row=3, column=0, padx=4, sticky="e")
        self.formatted_result_var = StringVar(value="")
        self.formatted_result_entry = Entry(
            format_frame,
            textvariable=self.formatted_result_var,
            width=125,
            font=("Consolas", 10),
            state="readonly",
            readonlybackground="white"
        )
        self.formatted_result_entry.grid(row=3, column=1, columnspan=5, padx=4, pady=4, sticky="we")
        format_frame.grid_columnconfigure(3, weight=1)

        tree_frame = Frame(self)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=8)
        x_scroll = ttk.Scrollbar(tree_frame, orient="horizontal")
        x_scroll.pack(side="bottom", fill="x")
        y_scroll = ttk.Scrollbar(tree_frame, orient="vertical")
        y_scroll.pack(side="right", fill="y")
        self.tree = ttk.Treeview(tree_frame, show="headings", xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
        self.tree.pack(fill="both", expand=True)
        x_scroll.config(command=self.tree.xview)
        y_scroll.config(command=self.tree.yview)

    def load_bom_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel文件", "*.xlsx *.xls")])
        if not files:
            return
        self.files = list(files)
        self.products = load_boms(self.files)
        self.bom_status.config(text=f"已加载 {len(self.products)} 个产品BOM | 备损率{LOSS_RATE}倍")
        self.build_inputs()
        self.calc_btn.config(state="normal" if self.products else "disabled")

    def load_avl_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xls")])
        if not path:
            return
        self.avl_df = load_avl(path)
        if self.avl_df is not None:
            self.avl_path_var.set(f"已加载AVL：{os.path.basename(path)}，记录数：{len(self.avl_df)}")

    def build_inputs(self):
        for widget in self.input_frame.winfo_children():
            widget.destroy()
        self.entries.clear()
        Label(self.input_frame, text="填写各产品计划生产数量（不生产填0或留空）", font=("微软雅黑", 11, "bold")).grid(row=0, column=0, columnspan=2, pady=4)
        for row, name in enumerate(self.products, start=1):
            Label(self.input_frame, text=f"{name}:").grid(row=row, column=0, sticky="e", padx=6, pady=1)
            entry = Entry(self.input_frame, width=14)
            entry.grid(row=row, column=1, padx=6, pady=1)
            self.entries[name] = entry

    def get_prod_counts(self):
        result = {}
        for name, entry in self.entries.items():
            text = entry.get().strip()
            try:
                value = float(text) if text else 0
                if value < 0:
                    raise ValueError
                result[name] = value
            except ValueError:
                raise ValueError(f"产品“{name}”的生产数量必须是非负数字。")
        return result

    def do_calc(self):
        try:
            self.detail, self.dist, self.wide_df = calculate(self.products, self.get_prod_counts())
        except ValueError as exc:
            messagebox.showerror("输入错误", str(exc))
            return
        if self.wide_df is None:
            messagebox.showwarning("提示", "请至少为一个产品填写大于0的生产数量。")
            return
        self.show_table(self.wide_df)
        self.export_btn.config(state="normal")
        messagebox.showinfo("计算成功", f"汇总计算完成！共 {len(self.wide_df)} 个物料。")

    def show_table(self, df):
        if df is None:
            return
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, minwidth=80, anchor="center")
        for row in df.itertuples(index=False, name=None):
            self.tree.insert("", "end", values=list(row))

    def do_search(self):
        """
        批量搜索总BOM usage：支持换行、空格、逗号和分号分隔；
        按标准化后的Part Number精确匹配，保留输入顺序，并显示未匹配料号。
        """
        if self.wide_df is None:
            messagebox.showwarning("提示", "请先完成BOM汇总计算。")
            return

        raw_text = self.search_box.get("1.0", END)
        part_numbers = parse_part_numbers(raw_text)
        if not part_numbers:
            messagebox.showwarning("提示", "请输入至少一个需要搜索的Part Number。")
            return

        source = self.wide_df.copy()
        source["__标准料号"] = source[PART_COL].apply(standardize_pn)
        rows = []
        matched_inputs = set()

        for order, input_part in enumerate(part_numbers, start=1):
            key = standardize_pn(input_part)
            matched = source[source["__标准料号"] == key].drop(columns=["__标准料号"]).copy()
            if matched.empty:
                empty_row = {col: "" for col in self.wide_df.columns}
                empty_row[PART_COL] = input_part
                empty_row["BOM搜索状态"] = "【总BOM usage中未找到】"
                empty_row["输入顺序"] = order
                rows.append(empty_row)
            else:
                matched_inputs.add(key)
                for _, row in matched.iterrows():
                    item = row.to_dict()
                    item["BOM搜索状态"] = "已找到"
                    item["输入顺序"] = order
                    rows.append(item)

        result = pd.DataFrame(rows)
        ordered_cols = list(self.wide_df.columns) + ["BOM搜索状态", "输入顺序"]
        self.bom_search_result = result[ordered_cols].sort_values("输入顺序").drop(columns=["输入顺序"]).reset_index(drop=True)
        self.show_table(self.bom_search_result)
        messagebox.showinfo(
            "BOM批量搜索完成",
            f"输入料号：{len(part_numbers)} 个\n"
            f"匹配料号：{len(matched_inputs)} 个\n"
            f"未匹配料号：{len(part_numbers) - len(matched_inputs)} 个"
        )

    def download_bom_search(self):
        """下载当前总BOM usage批量搜索结果。"""
        if self.bom_search_result is None:
            self.do_search()
            if self.bom_search_result is None:
                return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel表格", "*.xlsx")],
            initialfile="总BOM_usage批量搜索结果.xlsx"
        )
        if not path:
            return
        try:
            unmatched = self.bom_search_result[
                self.bom_search_result["BOM搜索状态"] == "【总BOM usage中未找到】"
            ].copy()
            export_frames(path, [
                ("BOM搜索结果", self.bom_search_result),
                ("未匹配料号", unmatched)
            ])
            messagebox.showinfo(
                "下载完成",
                f"BOM搜索结果已保存。\n未匹配料号：{len(unmatched)} 个"
            )
        except PermissionError:
            messagebox.showerror("保存失败", "文件正被Excel占用，请关闭后重试。")
        except Exception as exc:
            messagebox.showerror("导出异常", str(exc))

    def do_avl_search(self):
        if self.avl_df is None:
            messagebox.showwarning("提示", "请先加载AVL基准表。")
            return
        # 优先读取右侧批量输入框；右侧为空时自动使用左侧单料号输入框。
        raw_text = self.avl_search_text.get("1.0", END).strip()
        input_source = "右侧批量输入框"
        if not raw_text:
            raw_text = self.search_box.get("1.0", END).strip()
            input_source = "左侧BOM批量输入框"
        part_numbers = parse_part_numbers(raw_text)
        if not part_numbers:
            messagebox.showwarning(
                "提示",
                "请在右侧批量输入框粘贴多个料号，或在左侧输入一个料号。\n"
                "支持换行、空格、逗号和分号分隔。"
            )
            return
        self.avl_search_result = search_avl_parts(part_numbers, self.avl_df)
        self.show_table(self.avl_search_result)
        matched_inputs = self.avl_search_result.loc[
            self.avl_search_result[AVL_FLAG_COL] != NO_AVL_TEXT, SEARCH_INPUT_COL
        ].nunique()
        messagebox.showinfo(
            "AVL搜索完成",
            f"读取位置：{input_source}\n"
            f"输入料号：{len(part_numbers)} 个\n"
            f"匹配料号：{matched_inputs} 个\n"
            f"未匹配料号：{len(part_numbers) - matched_inputs} 个"
        )

    def download_avl_search(self):
        if self.avl_search_result is None:
            self.do_avl_search()
            if self.avl_search_result is None:
                return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel表格", "*.xlsx")], initialfile="AVL批量搜索结果.xlsx")
        if not path:
            return
        try:
            no_match = self.avl_search_result[self.avl_search_result[AVL_FLAG_COL] == NO_AVL_TEXT].copy()
            export_frames(path, [("AVL搜索结果", self.avl_search_result), ("未匹配料号", no_match)])
            messagebox.showinfo("下载完成", f"AVL搜索结果已保存。\n未匹配输入料号：{no_match[SEARCH_INPUT_COL].nunique()} 个")
        except PermissionError:
            messagebox.showerror("保存失败", "文件正被Excel占用，请关闭后重试。")
        except Exception as exc:
            messagebox.showerror("导出异常", str(exc))

    def generate_copy_format(self):
        """把输入的多个Part Number转换成竖线拼接格式。"""
        raw_text = self.format_input_text.get("1.0", END)
        result = format_part_numbers_for_copy(raw_text, self.format_suffix_var.get())
        if not result:
            messagebox.showwarning("提示", "请先输入至少一个Part Number。")
            return
        self.formatted_result_var.set(result)

    def copy_formatted_result(self):
        """一键复制拼接结果到Windows剪贴板。"""
        result = self.formatted_result_var.get().strip()
        if not result:
            self.generate_copy_format()
            result = self.formatted_result_var.get().strip()
        if not result:
            return
        self.clipboard_clear()
        self.clipboard_append(result)
        self.update()
        messagebox.showinfo("复制成功", f"已复制 {result.count('|')} 个分隔符的拼接结果。")

    def load_wide_parts_to_formatter(self):
        """将当前总BOM usage中的全部Part Number导入拼接工具。"""
        if self.wide_df is None or PART_COL not in self.wide_df.columns:
            messagebox.showwarning("提示", "请先完成BOM汇总计算。")
            return
        part_numbers = [
            standardize_pn(value)
            for value in self.wide_df[PART_COL].tolist()
            if standardize_pn(value)
        ]
        part_numbers = list(dict.fromkeys(part_numbers))
        self.format_input_text.delete("1.0", END)
        self.format_input_text.insert("1.0", "\n".join(part_numbers))
        self.generate_copy_format()

    def clear_formatter(self):
        self.format_input_text.delete("1.0", END)
        self.formatted_result_var.set("")

    def do_export(self):
        mode = self.export_mode.get()
        if self.wide_df is None:
            messagebox.showwarning("提示", "请先完成汇总计算。")
            return
        if mode == 2 and self.avl_df is None:
            messagebox.showwarning("提示", "已选择“Usage+AVL匹配”，请先加载AVL文件，或切换为仅导出Usage。")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel表格", "*.xlsx")],
            initialfile="一体化备料汇总.xlsx" if mode == 1 else "一体化备料汇总_带AVL信息.xlsx"
        )
        if not path:
            return
        try:
            frames = []
            if mode == 2:
                wide_full = merge_avl_info(self.wide_df, self.avl_df)
                # 去重后再并入标记，避免一个料号对应多个AVL时放大明细行数
                flag_map = wide_full[[PART_COL, AVL_FLAG_COL]].drop_duplicates(subset=[PART_COL])
                dist_full = self.dist.merge(flag_map, on=PART_COL, how="left")
                detail_full = self.detail.merge(flag_map, on=PART_COL, how="left")
                no_avl_df = wide_full[wide_full[AVL_FLAG_COL] == NO_AVL_TEXT].copy()
                frames.extend([
                    ("总BOM usage", wide_full),
                    ("无AVL物料", no_avl_df),
                    ("分产品明细", dist_full),
                    ("各产品计算明细", detail_full)
                ])
                avl_info = f"\n无AVL物料：{no_avl_df[PART_COL].nunique()} 个，已单独输出到“无AVL物料”Sheet。"
            else:
                wide_full = self.wide_df
                frames.extend([
                    ("总BOM usage", wide_full),
                    ("分产品明细", self.dist),
                    ("各产品计算明细", self.detail)
                ])
                avl_info = ""
            frames.extend((prod_name, df) for prod_name, df in self.products.items())
            supplier_count = export_frames(path, frames, supplier_source_df=self.wide_df)
            messagebox.showinfo(
                "导出完成",
                f"文件保存成功！\n"
                f"总物料数：{wide_full[PART_COL].nunique()}{avl_info}\n"
                f"已在最后一个Sheet生成“{SUPPLIER_SHEET_NAME}”：{supplier_count} 个物料。\n"
                f"Demand Qty已自动取备损后总需求(×{LOSS_RATE})。"
            )
        except PermissionError:
            messagebox.showerror("保存失败", "文件正被Excel占用，请关闭后重试。")
        except Exception as exc:
            messagebox.showerror("导出异常", str(exc))


if __name__ == "__main__":
    BOMAVLApp().mainloop()
