import cdm6calc, cdm7calc, identify, sumIOM, fioParser, chartGUI
import ctypes
import tkinter as tk

ctypes.windll.shcore.SetProcessDpiAwareness(1)

root = tk.Tk()
pane = tk.Frame(root)
pane.pack(fill=tk.BOTH, expand=True)
root.title("CTools v3")
root.geometry("350x400")


def onClick(arg):
    if arg == 1:
        cdm6calc.main()
    elif arg == 2:
        cdm7calc.main()
    elif arg == 3:
        identify.main()
    elif arg == 4:
        sumIOM.main()
    elif arg == 5:
        fioParser.main()
    elif arg == 6:
        chartGUI.gui()


cdm6 = tk.Button(pane, text='CrystalDiskMark 6', command=lambda: onClick(1), bg='#ffe0e0')
cdm7 = tk.Button(pane, text='CrystalDiskMark 7', command=lambda: onClick(2), bg='#fff5e0')
ident = tk.Button(pane, text='Identify', command=lambda: onClick(3), bg='#eaffe0')
iom = tk.Button(pane, text='IOMeter', command=lambda: onClick(4), bg='#e0fffe')
fio = tk.Button(pane, text='FIO Parser', command=lambda: onClick(5), bg='#e3e0ff')
chart = tk.Button(pane, text='Charting Tool', command=lambda: onClick(6), bg='#ffd494')

cdm6.pack(side=tk.TOP, expand=True, fill=tk.BOTH)
cdm7.pack(side=tk.TOP, expand=True, fill=tk.BOTH)
ident.pack(side=tk.TOP, expand=True, fill=tk.BOTH)
iom.pack(side=tk.TOP, expand=True, fill=tk.BOTH)
fio.pack(side=tk.TOP, expand=True, fill=tk.BOTH)
chart.pack(side=tk.TOP, expand=True, fill=tk.BOTH)

root.mainloop()
