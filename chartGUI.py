import tkinter
import ctypes
import sys
from tkinter import filedialog
import tkinter.ttk as ttk
from tkinter_custom_button import TkinterCustomButton


def gui():
    master = tkinter.Tk()
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    perfpath = tkinter.StringVar()
    smartpath = tkinter.StringVar()
    tcpath = tkinter.StringVar()
    pcpath = tkinter.StringVar()
    master.title('Chart GUI')
    width = 970
    height = 700
    screenwidth = master.winfo_screenwidth()
    screenheight = master.winfo_screenheight()
    x = (screenwidth/2) - (width/2)
    y = (screenheight/2) - (height/2)
    master.geometry("%dx%d+%d+%d" % (width, height, x, y))
    master.configure(background='#303E4F')

    def quit_on_click():
        master.quit()
        master.destroy()
        sys.exit()

    def perf_file():
        f = filedialog.askopenfilename(initialdir="/Desktop", title="Select A File",
                                       filetypes=(("Comma Separated Values File", "*.csv"), ("All Files", "*.*")))
        perfpathE.delete(0, tkinter.END)
        perfpathE.insert(0, f)
        return

    def smart_file():
        f = filedialog.askopenfilename(initialdir="/Desktop", title="Select A File",
                                       filetypes=(("Comma Separated Values File", "*.csv"), ("All Files", "*.*")))
        smartpathE.delete(0, tkinter.END)
        smartpathE.insert(0, f)
        return

    def tc_file():
        f = filedialog.askopenfilename(initialdir="/Desktop", title="Select A File",
                                       filetypes=(("Comma Separated Values File", "*.csv"), ("All Files", "*.*")))
        tcpathE.delete(0, tkinter.END)
        tcpathE.insert(0, f)
        return

    def pc_file():
        f = filedialog.askopenfilename(initialdir="/Desktop", title="Select A File",
                                       filetypes=(("Comma Separated Values File", "*.csv"), ("All Files", "*.*")))
        pcpathE.delete(0, tkinter.END)
        pcpathE.insert(0, f)
        return

    def clearall():
        opt.delete(0, tkinter.END)
        voltage.delete(0, tkinter.END)
        channel.delete(0, tkinter.END)
        perfpathE.delete(0, tkinter.END)
        smartpathE.delete(0, tkinter.END)
        tcpathE.delete(0, tkinter.END)
        pcpathE.delete(0, tkinter.END)
        return

    def submitall():
        if str(opt.get()) == '':
            opt.set('Sequential')
            return

        if voltage.get() == '':
            defval = 3.3
            voltage.set(defval)

        if channel.get() == '':
            defval = 101
            channel.set(defval)

        fperfpath = str(perfpathE.get())
        fsmartpath = str(smartpathE.get())
        ftcpath = str(tcpathE.get())
        fpcpath = str(pcpathE.get())
        fchannel = int(channel.get())
        fvoltage = float(voltage.get())

        print('perfpath is:', fperfpath)
        print('smartpath is:', fsmartpath)
        print('tcpath is:', ftcpath)
        print('pcpath is:', fpcpath)
        print('channel is:', fchannel)
        print('voltage is:', fvoltage)

        import combineChart
        combineChart.main(str(opt.get()), str(fperfpath), str(fsmartpath), str(ftcpath), str(fpcpath), int(fchannel),
                          float(fvoltage))

    ttk.Label(master, background='#303E4F', foreground='#ffffff', text="Select Data Type:",
              font=("Calibri Light", 12)).place(x=50, y=50)
    n = tkinter.StringVar()
    opt = ttk.Combobox(master, width=27, textvariable=n, values=('Sequential', 'Random'))
    opt.place(x=200, y=50)
    opt.current()

    ttk.Label(master, background='#303E4F', foreground='#ffffff', text="Select Voltage:",
              font=("Calibri Light", 12)).place(x=500, y=50)
    q = tkinter.StringVar()
    voltage = ttk.Combobox(master, width=27, textvariable=q)
    voltage['values'] = (3.3, 5, 12)
    voltage.place(x=650, y=50)
    voltage.current()

    ttk.Label(master, background='#303E4F', foreground='#ffffff', text="Channel Used:",
              font=("Calibri Light", 12)).place(x=130, y=385)
    ch = tkinter.StringVar()
    channel = ttk.Combobox(master, width=27, textvariable=ch)
    channel['values'] = (101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125)
    channel.place(x=275, y=385)
    channel.current()

    TkinterCustomButton(master=master,
                        bg_color=None,
                        fg_color='#303E4F',
                        border_color="#f57567",
                        text_font=None,
                        text="Performance Data",
                        text_color="#f57567",
                        corner_radius=0,
                        border_width=2,
                        width=200,
                        height=50,
                        hover=False,
                        command=perf_file).place(x=50, y=120)

    TkinterCustomButton(master=master,
                        bg_color=None,
                        fg_color='#303E4F',
                        border_color="#d28aff",
                        text_font=None,
                        text="SMART Temp Data",
                        text_color="#d28aff",
                        corner_radius=0,
                        border_width=2,
                        width=200,
                        height=50,
                        hover=False,
                        command=smart_file).place(x=50, y=220)

    TkinterCustomButton(master=master,
                        bg_color=None,
                        fg_color='#303E4F',
                        border_color="#ffffb0",
                        text_font=None,
                        text="Thermal Coup. Data",
                        text_color="#ffffb0",
                        corner_radius=0,
                        border_width=2,
                        width=200,
                        height=50,
                        hover=False,
                        command=tc_file).place(x=50, y=320)

    TkinterCustomButton(master=master,
                        bg_color=None,
                        fg_color='#303E4F',
                        border_color="#9efdff",
                        text_font=("Century Gothic", 9),
                        text="Power Consumption Data",
                        text_color="#9efdff",
                        corner_radius=0,
                        border_width=2,
                        width=200,
                        height=50,
                        hover=False,
                        command=pc_file).place(x=50, y=460)

    TkinterCustomButton(master=master,
                        bg_color=None,
                        fg_color='#303E4F',
                        border_color="#ff61e9",
                        text="Submit",
                        text_color="#ff61e9",
                        corner_radius=0,
                        border_width=2,
                        width=200,
                        height=50,
                        hover=False,
                        command=submitall).place(x=275, y=560)

    TkinterCustomButton(master=master,
                        bg_color=None,
                        fg_color='#303E4F',
                        border_color="#fdbdff",
                        text="Clear All",
                        text_color="#fdbdff",
                        corner_radius=0,
                        border_width=2,
                        width=200,
                        height=50,
                        hover=False,
                        command=clearall).place(x=500, y=560)

    perfpathE = ttk.Entry(master, width=76, textvariable=perfpath)
    perfpathE.place(x=275, y=130)

    smartpathE = ttk.Entry(master, width=75, textvariable=smartpath)
    smartpathE.place(x=275, y=230)

    tcpathE = ttk.Entry(master, width=75, textvariable=tcpath)
    tcpathE.place(x=275, y=330)

    pcpathE = ttk.Entry(master, width=75, textvariable=pcpath)
    pcpathE.place(x=275, y=470)

    master.protocol("WM_DELETE_WINDOW", quit_on_click)
    master.mainloop()


if __name__ == '__main__':
    gui()
