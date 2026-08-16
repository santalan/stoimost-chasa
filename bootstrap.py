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

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
CELL_TEXT = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"


def _xlsx_rows(path):
    """Read sheet1.xml directly, ignoring malformed print/page setup metadata."""
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", NS):
                shared.append("".join((t.text or "") for t in si.iter(CELL_TEXT)))

        root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
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
                        value = shared[int(value)]
                vals[col] = value
            yield vals


def robust_parse_files(paths):
    out = []
    seen = set()
    warnings = []

    for p in paths:
        try:
            digest = hashlib.sha256(Path(p).read_bytes()).hexdigest()
            if digest in seen:
                continue
            seen.add(digest)

            role = False
            for vals in _xlsx_rows(p):
                label = str(vals.get("B", "") or "").strip()
                if label == "Должность:":
                    role_text = " ".join(str(vals.get(c, "") or "") for c in ("E", "F", "G"))
                    role = "Врач-рентгенолог" in role_text
                    continue
                if role and label == "Всего начислено":
                    role = False
                    continue
                if not role or not label:
                    continue

                per = run.period(vals.get("F"))
                amt = run.fnum(vals.get("O"))
                if not per or amt is None:
                    continue

                hrs = 0.0
                if label.lower().startswith("оклад"):
                    for col in ("K", "J"):
                        nums = re.findall(r"-?\d+(?:[.,]\d+)?", str(vals.get(col, "") or ""))
                        if nums:
                            hrs = float(nums[-1].replace(",", "."))
                            break

                salary = None
                if label.lower() == "оклад" and hrs > 0:
                    norm = run.NORM.get(per[0], [None] * 12)[per[1] - 1]
                    if norm:
                        salary = amt * norm / hrs

                out.append({
                    "y": per[0], "m": per[1], "label": label, "amt": amt,
                    "hours": hrs, "cat": run.classify(label), "file": Path(p).name,
                    "salary": salary,
                })
        except Exception as exc:
            warnings.append(f"{Path(p).name}: {exc}")

    return out, warnings


run.parse_files = robust_parse_files

_original_build = run.App.build

def _build_with_progress(self):
    _original_build(self)
    self.progress = ttk.Progressbar(self, mode="indeterminate")
    self.progress.pack(fill="x", padx=10, pady=(0, 6))

run.App.build = _build_with_progress


def _calc_with_status(self):
    if not self.paths:
        return messagebox.showinfo(run.APP, "Сначала добавьте расчётные листки.")

    self.status.config(text=f"Обработка {len(self.paths)} файлов…")
    self.progress.start(12)
    self.config(cursor="watch")
    self.update_idletasks()
    try:
        self.accr, warnings = robust_parse_files(self.paths)
        if not self.accr:
            self.rows = []
            self.draw()
            details = "\n".join(warnings[:8]) if warnings else "Подходящих начислений не найдено."
            messagebox.showerror(run.APP, "Не удалось прочитать расчётные листки.\n\n" + details)
            self.status.config(text="Расчёт не выполнен")
            return

        self.rows, _ = run.analyze(self.accr, {k: v.get() for k, v in self.mode.items()})
        self.draw()
        msg = f"Готово: {len(self.paths)} файлов, {len(self.accr)} начислений"
        if warnings:
            msg += f" · файлов с предупреждениями: {len(warnings)}"
        self.status.config(text=msg)
    except Exception as exc:
        self.status.config(text="Ошибка при построении графика")
        messagebox.showerror(run.APP, f"Не удалось построить график:\n\n{exc}")
    finally:
        self.progress.stop()
        self.config(cursor="")

run.App.calc = _calc_with_status
App = run.App


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
