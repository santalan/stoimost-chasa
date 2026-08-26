"""Windows launcher, Tkinter compatibility shim and robust XLSX reader."""
import hashlib
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

_original_boolean_var = tk.BooleanVar

def _boolean_var_compat(master=None, value=None, name=None):
    if isinstance(master, bool) and value is None:
        value = master
        master = None
    return _original_boolean_var(master=master, value=value, name=name)

tk.BooleanVar = _boolean_var_compat

import run
run.VERSION = '0.2.3'

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
CELL_TEXT = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
PERIOD_RE = re.compile(r"^\s*(\d{1,2})-(\d{4})\s*$")
HOURS_WITH_UNIT_RE = re.compile(r"(-?\d+(?:[.,]\d+)?)\s*[чЧ](?:\.|\b)?")
HOURS_AFTER_SEMI_RE = re.compile(r";\s*(-?\d+(?:[.,]\d+)?)\s*(?:[чЧ])?")


def _col_num(col):
    n = 0
    for ch in col:
        if 'A' <= ch <= 'Z':
            n = n * 26 + ord(ch) - 64
    return n


def _xlsx_rows(path):
    """Read worksheet XML directly and ignore malformed Excel page/print metadata."""
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", NS):
                shared.append("".join((t.text or "") for t in si.iter(CELL_TEXT)))

        sheets = sorted(
            (name for name in z.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)),
            key=lambda x: int(re.search(r"sheet(\d+)\.xml$", x).group(1))
        )
        for sheet_name in sheets:
            root = ET.fromstring(z.read(sheet_name))
            for row in root.findall(".//a:sheetData/a:row", NS):
                vals = {}
                for c in row.findall("a:c", NS):
                    ref = c.attrib.get("r", "")
                    m = re.match(r"([A-Z]+)", ref)
                    if not m:
                        continue
                    col = m.group(1)
                    cell_type = c.attrib.get("t")
                    v = c.find("a:v", NS)
                    if cell_type == "inlineStr":
                        inline = c.find("a:is", NS)
                        value = "".join((t.text or "") for t in inline.iter(CELL_TEXT)) if inline is not None else ""
                    else:
                        value = v.text if v is not None else ""
                        if cell_type == "s" and value != "":
                            try:
                                value = shared[int(value)]
                            except Exception:
                                pass
                    vals[col] = value
                if vals:
                    yield vals


def _find_period(vals):
    preferred = str(vals.get("F", "") or "")
    if PERIOD_RE.match(preferred):
        return run.period(preferred)
    for col, value in sorted(vals.items(), key=lambda kv: _col_num(kv[0])):
        if PERIOD_RE.match(str(value or "")):
            p = run.period(value)
            if p:
                return p
    return None


def _find_amount(vals):
    v = run.fnum(vals.get("O"))
    if v is not None:
        return v
    # Fallback for slightly shifted payroll layouts: search monetary cells around N:P.
    for col in ("N", "P", "M", "Q"):
        v = run.fnum(vals.get(col))
        if v is not None:
            return v
    return None


def _find_hours(vals):
    """Find paid hours in an Оклад row across known old/new layouts."""
    # Strong patterns first: explicit 'ч' or value after semicolon.
    for col in ("K", "J", "L", "I", "H", "G", "M"):
        text = str(vals.get(col, "") or "").strip()
        m = HOURS_WITH_UNIT_RE.search(text)
        if m:
            return float(m.group(1).replace(",", "."))
        m = HOURS_AFTER_SEMI_RE.search(text)
        if m:
            return float(m.group(1).replace(",", "."))

    # Then scan the row for explicit hour markers.
    for value in vals.values():
        text = str(value or "").strip()
        m = HOURS_WITH_UNIT_RE.search(text)
        if m:
            return float(m.group(1).replace(",", "."))

    # Last safe fallback: a standalone numeric J/K/L cell in plausible hour range.
    for col in ("J", "K", "L"):
        text = str(vals.get(col, "") or "").strip().replace(",", ".")
        if re.fullmatch(r"-?\d+(?:\.\d+)?", text):
            value = float(text)
            if -250 <= value <= 250:
                return value
    return 0.0


def robust_parse_files(paths, progress=None):
    out = []
    seen = set()
    warnings = []
    diagnostics = []
    total = len(paths)

    for index, p in enumerate(paths, start=1):
        file_accr_before = len(out)
        file_hours_before = sum(abs(a.get("hours", 0) or 0) for a in out)
        try:
            digest = hashlib.sha256(Path(p).read_bytes()).hexdigest()
            if digest in seen:
                diagnostics.append(f"{Path(p).name}: дубликат пропущен")
                continue
            seen.add(digest)

            role = False
            role_seen = False
            radiologist_seen = False
            for vals in _xlsx_rows(p):
                label = str(vals.get("B", "") or "").strip()
                label_low = label.casefold()
                row_text = " ".join(str(v or "") for v in vals.values()).casefold()

                if label_low == "должность:":
                    role_seen = True
                    radiologist_seen = "рентгенолог" in row_text
                    role = radiologist_seen
                    continue
                if role and label_low == "всего начислено":
                    role = False
                    continue
                if not role or not label:
                    continue

                per = _find_period(vals)
                amt = _find_amount(vals)
                if not per or amt is None:
                    continue

                hrs = _find_hours(vals) if label_low.startswith("оклад") else 0.0
                salary = None
                if label_low == "оклад" and hrs > 0:
                    norm = run.NORM.get(per[0], [None] * 12)[per[1] - 1]
                    if norm:
                        salary = amt * norm / hrs

                out.append({
                    "y": per[0], "m": per[1], "label": label, "amt": amt,
                    "hours": hrs, "cat": run.classify(label), "file": Path(p).name,
                    "salary": salary,
                })

            added = len(out) - file_accr_before
            added_hours = sum(abs(a.get("hours", 0) or 0) for a in out) - file_hours_before
            if not role_seen:
                diagnostics.append(f"{Path(p).name}: строка «Должность» не найдена")
            elif not radiologist_seen:
                diagnostics.append(f"{Path(p).name}: должность не распознана как рентгенолог")
            elif added == 0:
                diagnostics.append(f"{Path(p).name}: начисления не распознаны")
            elif added_hours == 0:
                diagnostics.append(f"{Path(p).name}: начисления найдены, часы в строках «Оклад» не найдены")
        except Exception as exc:
            warnings.append(f"{Path(p).name}: {exc}")
        finally:
            if progress:
                progress(index, total)

    return out, warnings, diagnostics


def _parse_files_compat(paths):
    out, warnings, _ = robust_parse_files(paths)
    return out, warnings

run.parse_files = _parse_files_compat

_original_build = run.App.build

def _build_with_progress(self):
    _original_build(self)
    graph_widget = self.canvas.get_tk_widget()
    parent = graph_widget.master

    # Move status above the graph so it is always visible, even on smaller displays.
    self.status.pack_forget()
    self.status.pack(fill="x", anchor="w", padx=2, pady=(0, 3), before=graph_widget)

    self.progress = ttk.Progressbar(parent, mode="determinate", maximum=100)
    self.progress.pack(fill="x", pady=(0, 5), before=graph_widget)

run.App.build = _build_with_progress


def _calc_with_status(self):
    if not self.paths:
        return messagebox.showinfo(run.APP, "Сначала добавьте расчётные листки.")

    self.status.config(text=f"Обработка {len(self.paths)} файлов…")
    self.progress["value"] = 0
    self.config(cursor="watch")
    self.update_idletasks()

    def progress(done, total):
        self.progress["maximum"] = max(total, 1)
        self.progress["value"] = done
        self.status.config(text=f"Чтение расчётных листков: {done} из {total}…")
        self.update_idletasks()

    try:
        self.accr, warnings, diagnostics = robust_parse_files(self.paths, progress=progress)
        if not self.accr:
            self.rows = []
            self.draw()
            details = "\n".join((warnings + diagnostics)[:10]) or "Подходящих начислений не найдено."
            messagebox.showerror(run.APP, "Не удалось прочитать расчётные листки.\n\n" + details)
            self.status.config(text="Расчёт не выполнен: начисления не распознаны")
            return

        total_hours = sum(abs(a.get("hours", 0) or 0) for a in self.accr)
        if total_hours < 0.01:
            self.rows = []
            self.draw()
            details = "\n".join(diagnostics[:10])
            messagebox.showerror(
                run.APP,
                "Начисления распознаны, но программа не нашла отработанные часы в строках «Оклад».\n\n"
                "Из-за этого стоимость часа рассчитать нельзя.\n\n" + details
            )
            self.status.config(text=f"Найдено {len(self.accr)} начислений, но 0 часов")
            return

        self.status.config(text=f"Расчёт: найдено {len(self.accr)} начислений и {total_hours:.1f} ч…")
        self.update_idletasks()
        self.rows, _ = run.analyze(self.accr, {k: v.get() for k, v in self.mode.items()})
        valid = [r for r in self.rows if r.get("rate") is not None]
        if not valid:
            self.draw()
            messagebox.showerror(
                run.APP,
                "Начисления и часы найдены, но ни для одного месяца не удалось рассчитать стоимость часа.\n"
                "Экспортируйте диагностические данные или пришлите один расчётный лист для проверки формата."
            )
            self.status.config(text=f"Найдено {len(self.accr)} начислений, но нет расчётных месяцев")
            return

        self.draw()
        msg = f"Готово: {len(self.paths)} файлов · {len(self.accr)} начислений · {len(valid)} месяцев"
        if warnings or diagnostics:
            msg += f" · замечаний: {len(warnings) + len(diagnostics)}"
        self.status.config(text=msg)
    except Exception as exc:
        self.status.config(text="Ошибка при построении графика")
        messagebox.showerror(run.APP, f"Не удалось построить график:\n\n{exc}")
    finally:
        self.progress["value"] = 0
        self.config(cursor="")
        self.update_idletasks()

run.App.calc = _calc_with_status
App = run.App


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
