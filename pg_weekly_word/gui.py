from __future__ import annotations

import threading
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, VERTICAL, W, X, Y, Button, Entry, Frame, Label, StringVar, Text, Tk, Toplevel, filedialog, messagebox, ttk

from .core import (
    GROUP_ORDER,
    PROJECT_ROOT,
    GROUP_CONFIG_PATH,
    WeeklyReportError,
    add_group_point,
    city_sheet_names,
    default_city_sheet,
    generate_weekly_report,
    period_display,
    read_city_ledger,
    read_district_base,
    read_groups_excel,
    validate_inputs,
)


GROUP_PLACEHOLDER = f"可选上传 xlsx；留空使用默认 JSON：{GROUP_CONFIG_PATH}"
PLACEHOLDER_COLOR = "#999999"
INPUT_COLOR = "#000000"


class WeeklyReportApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("平谷区垃圾分类周报生成程序")
        self.root.geometry("980x640")
        self.root.minsize(900, 560)
        self.vars = {
            "district": StringVar(),
            "city": StringVar(),
            "template": StringVar(),
            "group": StringVar(value=GROUP_PLACEHOLDER),
            "output": StringVar(),
            "district_period": StringVar(),
            "city_period": StringVar(),
            "previous_district_issues": StringVar(),
            "year": StringVar(value="2026"),
            "sheet": StringVar(),
            "status": StringVar(value="请选择文件后点击校验。"),
            "result": StringVar(),
        }
        self._build()
        self.root.after(100, self.auto_fill_files)

    def _build(self) -> None:
        root_frame = Frame(self.root, padx=8, pady=6)
        root_frame.pack(fill=BOTH, expand=True)

        file_frame = ttk.LabelFrame(root_frame, text="文件选择")
        file_frame.pack(fill=X)
        self._file_row(file_frame, 0, "区级基础台账", "district", [("Excel 97-2003", "*.xls"), ("所有文件", "*.*")], self._after_district_selected)
        self._file_row(file_frame, 1, "周市级检查台账", "city", [("Excel", "*.xlsx *.xlsm"), ("所有文件", "*.*")], self._after_city_selected)
        self._file_row(file_frame, 2, "Jinja 周报模板", "template", [("Word 模板", "*.docx"), ("所有文件", "*.*")])
        self._file_row(file_frame, 3, "分组清单（可选上传）", "group", [("Excel", "*.xlsx"), ("所有文件", "*.*")])
        Button(file_frame, text="新增点位", command=self.open_add_point_dialog).grid(row=3, column=3, padx=4, pady=3)
        self._dir_row(file_frame, 4, "输出目录", "output")
        Button(file_frame, text="自动查找", command=self.auto_fill_files).grid(row=5, column=2, padx=4, pady=3)

        param_frame = ttk.LabelFrame(root_frame, text="参数")
        param_frame.pack(fill=X, pady=(6, 0))
        Label(param_frame, text="年份").grid(row=0, column=0, sticky=W, padx=4, pady=4)
        Entry(param_frame, textvariable=self.vars["year"], width=8).grid(row=0, column=1, sticky=W, padx=4)
        Label(param_frame, text="区级周期").grid(row=0, column=2, sticky=W, padx=4)
        Entry(param_frame, textvariable=self.vars["district_period"], width=16).grid(row=0, column=3, sticky=W, padx=4)
        Label(param_frame, text="市级周期").grid(row=0, column=4, sticky=W, padx=4)
        Entry(param_frame, textvariable=self.vars["city_period"], width=16).grid(row=0, column=5, sticky=W, padx=4)
        Label(param_frame, text="市级工作表").grid(row=0, column=6, sticky=W, padx=4)
        self.sheet_combo = ttk.Combobox(param_frame, textvariable=self.vars["sheet"], width=14, state="readonly")
        self.sheet_combo.grid(row=0, column=7, sticky=W, padx=4)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_city_period())
        Button(param_frame, text="解析", command=self.parse_city_sheets).grid(row=0, column=8, sticky=W, padx=4)
        Label(param_frame, text="上周问题数").grid(row=1, column=0, sticky=W, padx=4)
        Entry(param_frame, textvariable=self.vars["previous_district_issues"], width=8).grid(row=1, column=1, sticky=W, padx=4)
        Label(
            param_frame,
            text=f"周期可留空自动识别；分组清单可选上传 xlsx，不上传则使用 {GROUP_CONFIG_PATH}；上周区级问题总数用于计算对比句。",
            foreground="#555555",
        ).grid(row=2, column=0, columnspan=9, sticky=W, padx=4, pady=(0, 4))

        action_frame = Frame(root_frame)
        action_frame.pack(fill=X, pady=6)
        Button(action_frame, text="校验预览", command=self.validate).pack(side=LEFT, padx=(0, 6))
        Button(action_frame, text="生成周报", command=self.generate).pack(side=LEFT, padx=(0, 6))
        Button(action_frame, text="打开输出目录", command=self.open_output_dir).pack(side=LEFT)
        Label(action_frame, textvariable=self.vars["status"], anchor=W).pack(side=LEFT, padx=16, fill=X, expand=True)

        preview_frame = ttk.LabelFrame(root_frame, text="校验状态与预览")
        preview_frame.pack(fill=BOTH, expand=True)
        self.preview = ttk.Notebook(preview_frame)
        self.preview.pack(fill=BOTH, expand=True, padx=4, pady=4)
        self.summary_text = self._text_tab("统计摘要")
        self.city_text = self._text_tab("市级六大类")
        self.rank_text = self._text_tab("区级排名表")
        self.result_text = self._text_tab("生成结果")

    def _file_row(self, parent: Frame, row: int, label: str, key: str, filetypes: list[tuple[str, str]], callback=None) -> None:
        Label(parent, text=label, width=15, anchor=W).grid(row=row, column=0, sticky=W, padx=4, pady=3)
        entry = Entry(parent, textvariable=self.vars[key])
        entry.grid(row=row, column=1, sticky="we", padx=4, pady=3)
        if key == "group":
            self.group_entry = entry
            self.group_entry.configure(fg=PLACEHOLDER_COLOR)
            self.group_entry.bind("<FocusIn>", self._clear_group_placeholder)
            self.group_entry.bind("<FocusOut>", self._restore_group_placeholder)
        Button(parent, text="选择", command=lambda: self._select_file(key, filetypes, callback)).grid(row=row, column=2, padx=4, pady=3)
        parent.columnconfigure(1, weight=1)

    def _clear_group_placeholder(self, _event=None) -> None:
        if self.vars["group"].get() == GROUP_PLACEHOLDER:
            self.vars["group"].set("")
        self.group_entry.configure(fg=INPUT_COLOR)

    def _restore_group_placeholder(self, _event=None) -> None:
        if not self.vars["group"].get().strip():
            self.vars["group"].set(GROUP_PLACEHOLDER)
            self.group_entry.configure(fg=PLACEHOLDER_COLOR)

    def _dir_row(self, parent: Frame, row: int, label: str, key: str) -> None:
        Label(parent, text=label, width=15, anchor=W).grid(row=row, column=0, sticky=W, padx=4, pady=3)
        Entry(parent, textvariable=self.vars[key]).grid(row=row, column=1, sticky="we", padx=4, pady=3)
        Button(parent, text="选择", command=lambda: self._select_dir(key)).grid(row=row, column=2, padx=4, pady=3)

    def _text_tab(self, title: str):
        frame = Frame(self.preview)
        self.preview.add(frame, text=title)
        yscroll = ttk.Scrollbar(frame, orient=VERTICAL)
        text = Text(frame, wrap="word", yscrollcommand=yscroll.set)
        yscroll.config(command=text.yview)
        yscroll.pack(side=RIGHT, fill=Y)
        text.pack(side=LEFT, fill=BOTH, expand=True)
        return text

    def open_add_point_dialog(self) -> None:
        dialog = Toplevel(self.root)
        dialog.title("新增点位")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        town_var = StringVar()
        group_var = StringVar(value=GROUP_ORDER[0] if GROUP_ORDER else "")

        Label(dialog, text="城乡镇").grid(row=0, column=0, sticky=W, padx=10, pady=(10, 4))
        town_entry = Entry(dialog, textvariable=town_var, width=28)
        town_entry.grid(row=0, column=1, sticky=W, padx=10, pady=(10, 4))
        Label(dialog, text="分组").grid(row=1, column=0, sticky=W, padx=10, pady=4)
        group_combo = ttk.Combobox(dialog, textvariable=group_var, values=GROUP_ORDER, width=25)
        group_combo.grid(row=1, column=1, sticky=W, padx=10, pady=4)

        button_frame = Frame(dialog)
        button_frame.grid(row=2, column=0, columnspan=2, sticky="e", padx=10, pady=(8, 10))
        Button(button_frame, text="取消", command=dialog.destroy).pack(side=LEFT, padx=(0, 6))
        Button(button_frame, text="新增", command=lambda: self._add_group_point(dialog, town_var, group_var)).pack(side=LEFT)

        town_entry.focus_set()
        self.root.wait_window(dialog)

    def _add_group_point(self, dialog: Toplevel, town_var: StringVar, group_var: StringVar) -> None:
        town = town_var.get().strip()
        group = group_var.get().strip()
        try:
            add_group_point(town, group)
        except WeeklyReportError as exc:
            messagebox.showerror("新增失败", str(exc), parent=dialog)
            return
        self.vars["status"].set(f"已新增分组点位：{town}（{group}）。")
        note = ""
        if self._group_path_value():
            note = "\n\n当前界面仍选择了上传分组清单，生成时会优先使用上传文件；如需使用刚新增的后台分组，请清空分组清单上传框。"
        messagebox.showinfo("新增成功", f"已更新后台分组表：{GROUP_CONFIG_PATH}{note}", parent=dialog)
        dialog.destroy()

    def _select_file(self, key: str, filetypes: list[tuple[str, str]], callback=None) -> None:
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            self.vars[key].set(path)
            if key == "group":
                self.group_entry.configure(fg=INPUT_COLOR)
            if callback:
                callback(path)

    def _select_dir(self, key: str) -> None:
        path = filedialog.askdirectory()
        if path:
            self.vars[key].set(path)

    def _after_city_selected(self, path: str) -> None:
        self.parse_city_sheets(path)

    def parse_city_sheets(self, path: str | None = None) -> None:
        path = path or self.vars["city"].get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择周市级检查台账。")
            return
        try:
            names = city_sheet_names(Path(path))
            self.sheet_combo["values"] = names
            default_name = default_city_sheet(Path(path))
            self.vars["sheet"].set(default_name)
            self.refresh_city_period(path, default_name)
            self.vars["status"].set(f"已解析市级工作表，默认选择：{default_name}")
        except Exception as exc:
            messagebox.showerror("市级台账读取失败", str(exc))

    def _after_district_selected(self, path: str) -> None:
        try:
            data = read_district_base(Path(path))
            if data["period_short"]:
                display = period_display(data["period_short"])
                self.vars["district_period"].set(display)
                self.vars["status"].set(f"已根据基础台账填写区级周期：{display}")
        except Exception as exc:
            messagebox.showerror("基础台账读取失败", str(exc))

    def refresh_city_period(self, path: str | None = None, sheet_name: str | None = None) -> None:
        path = path or self.vars["city"].get().strip()
        sheet_name = sheet_name or self.vars["sheet"].get().strip()
        if not path or not sheet_name:
            return
        try:
            data = read_city_ledger(Path(path), sheet_name)
            if data["period_short"]:
                display = period_display(data["period_short"])
                self.vars["city_period"].set(display)
                self.vars["status"].set(f"已根据市级台账 {sheet_name} 填写市级周期：{display}")
        except Exception as exc:
            messagebox.showerror("市级周期识别失败", str(exc))

    def auto_fill_files(self) -> None:
        found: list[str] = []
        template = self._find_template()
        district = self._find_district_base()
        city = self._find_city_ledger()
        output = PROJECT_ROOT / "output"

        if template and not self.vars["template"].get().strip():
            self.vars["template"].set(str(template))
            found.append("模板")
        if district and not self.vars["district"].get().strip():
            self.vars["district"].set(str(district))
            found.append("区级基础台账")
            self._fill_district_period_silent(district)
        if city and not self.vars["city"].get().strip():
            self.vars["city"].set(str(city))
            found.append("市级台账")
            self._fill_city_sheets_silent(city)
        if not self.vars["output"].get().strip():
            output.mkdir(parents=True, exist_ok=True)
            self.vars["output"].set(str(output))
            found.append("输出目录")

        if self.vars["group"].get().strip() == "":
            self.vars["group"].set(GROUP_PLACEHOLDER)
            self.group_entry.configure(fg=PLACEHOLDER_COLOR)

        if found:
            self.vars["status"].set("已自动填入：" + "、".join(found))
        else:
            self.vars["status"].set("未找到可自动填入的新文件，已有选择保持不变。")

    def _find_template(self) -> Path | None:
        preferred = PROJECT_ROOT / "input" / "周报模板.docx"
        if preferred.exists():
            return preferred
        candidates = self._existing_files(["input/*周报模板*.docx", "**/*周报模板*.docx", "**/*template*.docx"])
        return candidates[0] if candidates else None

    def _find_district_base(self) -> Path | None:
        candidates = [
            path for path in self._existing_files(["input/*.xls", "data/**/*.xls", "**/*.xls"])
            if "问题台账" not in path.name and not path.name.startswith("~$")
        ]
        for path in candidates:
            try:
                read_district_base(path)
                return path
            except Exception:
                continue
        return None

    def _find_group_file(self) -> Path | None:
        candidates = self._existing_files(["input/*分组清单*.xlsx", "data/**/*分组清单*.xlsx", "**/*分组清单*.xlsx"])
        for path in candidates:
            try:
                read_groups_excel(path)
                return path
            except Exception:
                continue
        return None

    def _find_city_ledger(self) -> Path | None:
        candidates = [
            path for path in self._existing_files(["input/*.xlsx", "data/**/*.xlsx", "**/*.xlsx"])
            if "分组清单" not in path.name and not path.name.startswith("~$")
        ]
        for path in candidates:
            try:
                sheet = default_city_sheet(path)
                read_city_ledger(path, sheet)
                return path
            except Exception:
                continue
        return None

    def _existing_files(self, patterns: list[str]) -> list[Path]:
        seen: set[Path] = set()
        result: list[Path] = []
        for pattern in patterns:
            for path in PROJECT_ROOT.glob(pattern):
                if path.is_file() and path not in seen and not path.name.startswith("~$"):
                    seen.add(path)
                    result.append(path)
        return sorted(result, key=lambda p: p.stat().st_mtime, reverse=True)

    def _fill_district_period_silent(self, path: Path) -> None:
        try:
            data = read_district_base(path)
            if data["period_short"]:
                self.vars["district_period"].set(period_display(data["period_short"]))
        except Exception:
            pass

    def _fill_city_sheets_silent(self, path: Path) -> None:
        try:
            names = city_sheet_names(path)
            self.sheet_combo["values"] = names
            default_name = default_city_sheet(path)
            self.vars["sheet"].set(default_name)
            data = read_city_ledger(path, default_name)
            if data["period_short"]:
                self.vars["city_period"].set(period_display(data["period_short"]))
        except Exception:
            pass

    def _paths(self) -> dict[str, Path]:
        required = ["district", "city", "template", "output"]
        missing = [key for key in required if not self.vars[key].get().strip()]
        if missing:
            raise WeeklyReportError("请先选择区级基础台账、市级台账、模板和输出目录。分组清单为可选上传。")
        paths = {key: Path(self.vars[key].get().strip()) for key in required}
        group_value = self._group_path_value()
        if group_value:
            paths["group"] = Path(group_value)
        return paths

    def _group_path_value(self) -> str:
        value = self.vars["group"].get().strip()
        if not value or value == GROUP_PLACEHOLDER:
            return ""
        return value

    def validate(self) -> None:
        try:
            paths = self._paths()
            data = validate_inputs(
                paths["district"],
                paths["city"],
                paths.get("group"),
                self.vars["sheet"].get() or None,
                district_period=self.vars["district_period"].get(),
                city_period=self.vars["city_period"].get(),
            )
            self.vars["district_period"].set(self.vars["district_period"].get() or period_display(data["district_period_short"]))
            self.vars["city_period"].set(self.vars["city_period"].get() or period_display(data["city_period_short"]))
            self.vars["status"].set("校验通过。" if not data["missing_groups"] else "校验未通过：存在未分组乡镇。")
            self._fill_validation(data)
        except Exception as exc:
            self.vars["status"].set("校验失败。")
            messagebox.showerror("校验失败", str(exc))

    def _fill_validation(self, data: dict) -> None:
        self._set_text(
            self.summary_text,
            "\n".join(
                [
                    f"市级工作表：{data['city_sheet']}",
                    f"识别到市级周期：{period_display(data['city_period_short'])}",
                    f"识别到区级周期：{period_display(data['district_period_short'])}",
                    f"分组来源：{data['group_source']}",
                    f"市级问题总数：{data['city_issue_total']}",
                    f"市级检查点位数：{data['city_checked_points']}",
                    f"市级日期筛选：周期内 {data['city_filtered_rows']} 条，排除 {data['city_filtered_out_rows']} 条",
                    f"区级问题总数：{data['district_issue_total']}",
                    f"区级检查点位数：{data['district_checked_total']}（小区 {data['district_community_count']}，村居 {data['district_village_count']}）",
                    "未分组乡镇：" + ("、".join(data["missing_groups"]) if data["missing_groups"] else "无"),
                ]
            ),
        )
        self._set_text(
            self.city_text,
            "\n".join(f"{row['name']}：{row['count']}处；{row['description']}" for row in data["city_six_categories"]),
        )
        rank_lines = ["组别\t排名\t乡镇街道\t问题数\t检查小区、村\t问题率"]
        for row in data["rank_preview"]:
            rank_lines.append(
                f"{row['group']}\t{row['rank']}\t{row['town']}\t{row['issue_count']}\t{row['checked_count']}\t{row['issue_rate']}"
            )
        self._set_text(self.rank_text, "\n".join(rank_lines))

    def generate(self) -> None:
        try:
            paths = self._paths()
            previous = self._previous_issue_count()
        except Exception as exc:
            messagebox.showerror("生成失败", str(exc))
            return
        self.vars["status"].set("正在生成，请稍候...")
        thread = threading.Thread(target=self._generate_worker, args=(paths, previous), daemon=True)
        thread.start()

    def _previous_issue_count(self) -> int:
        value = self.vars["previous_district_issues"].get().strip()
        if not value:
            raise WeeklyReportError("请填写上周区级垃圾分类问题总数。")
        if not value.isdigit():
            raise WeeklyReportError("上周区级垃圾分类问题总数必须是非负整数。")
        return int(value)

    def _generate_worker(self, paths: dict[str, Path], previous: int) -> None:
        try:
            result = generate_weekly_report(
                district_base_path=paths["district"],
                city_ledger_path=paths["city"],
                template_path=paths["template"],
                group_path=paths.get("group"),
                output_dir=paths["output"],
                report_year=self.vars["year"].get().strip() or "2026",
                district_period=self.vars["district_period"].get(),
                city_period=self.vars["city_period"].get(),
                city_sheet=self.vars["sheet"].get() or None,
                previous_district_issues=previous,
            )
            message = "\n".join(
                [
                    "生成完成。",
                    f"周报：{result.report_path}",
                    "问题台账和统计摘要已作为临时文件生成并清理。",
                ]
            )
            warnings = result.context.get("warnings") or []
            if warnings:
                message += "\n\n配置提示：\n" + "\n".join(f"- {item}" for item in warnings)
            self.root.after(0, lambda: self._generated(message))
        except Exception as exc:
            error_message = str(exc)
            self.root.after(0, lambda message=error_message: self._failed(message))

    def _generated(self, message: str) -> None:
        self.vars["status"].set("生成完成。")
        self._set_text(self.result_text, message)
        self.preview.select(3)
        messagebox.showinfo("生成完成", message)

    def _failed(self, message: str) -> None:
        self.vars["status"].set("生成失败。")
        messagebox.showerror("生成失败", message)

    def _set_text(self, widget, content: str) -> None:
        widget.delete("1.0", END)
        widget.insert("1.0", content)

    def open_output_dir(self) -> None:
        path = self.vars["output"].get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择输出目录。")
            return
        Path(path).mkdir(parents=True, exist_ok=True)
        import os

        os.startfile(path)


def main() -> None:
    root = Tk()
    WeeklyReportApp(root)
    root.mainloop()
