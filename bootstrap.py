"""Windows launcher and Tkinter compatibility shim for Стоимость часа."""
import tkinter as tk

_original_boolean_var = tk.BooleanVar

def _boolean_var_compat(master=None, value=None, name=None):
    # v0.2 used tk.BooleanVar(False/True). Tkinter treats the first positional
    # argument as master, not value. Preserve that intended behaviour safely.
    if isinstance(master, bool) and value is None:
        value = master
        master = None
    return _original_boolean_var(master=master, value=value, name=name)

tk.BooleanVar = _boolean_var_compat

from run import App


def main():
    App().mainloop()


if __name__ == '__main__':
    main()
