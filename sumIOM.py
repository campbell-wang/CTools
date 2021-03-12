import os, openpyxl, csv, ctypes
import statistics
from tkinter import *
import tkinter.filedialog


def main():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        root = Tk()
        root.withdraw()
        root.update()
        path = tkinter.filedialog.askopenfilename()
        if path:
            root.withdraw()
            root.update()
            save_dest = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Spreadsheet (.xlsx)", ".xlsx")])
            root.withdraw()
            root.destroy()

            wb = openpyxl.Workbook()
            ws = wb.active
            wb.create_sheet(index=1, title='Worker Average')
            rp = wb[wb.sheetnames[1]]

            writeSums = []
            readSums = []
            iops = 0

            with open(path) as f:
                reader = csv.reader(f, delimiter=',')
                for row in reader:
                    ws.append(row)

            for i in range(15, ws.max_row):  # iterates from row 15 to max row
                if (ws.cell(row=i, column=1)).value == "TimeStamp":
                    readStart = i
                    break

            for i in range(15, readStart):
                if (ws.cell(row=i, column=4)).value != '':  # stop when column D is not empty
                    wset = i
                    iops = 0
                    break
                else:
                    iops += float((ws.cell(row=i, column=8)).value)

                    if ws.cell(row=i, column=3).value == 'Worker 16':  # stop every 16 workers
                        writeSums.append(round(iops, 1))
                        iops = 0

            for i in range(readStart + 1, ws.max_row):
                if (ws.cell(row=i, column=4)).value != '':
                    rset = i
                    break
                else:
                    iops += float((ws.cell(row=i, column=8)).value)

                    if ws.cell(row=i, column=3).value == 'Worker 16':  # stop every 16 workers
                        readSums.append(round(iops, 1))
                        iops = 0

            def template():
                rp.cell(row=4, column=2, value="Write:")
                rp.cell(row=2, column=3, value=ws.cell(row=wset, column=4).value)
                rp.cell(row=4, column=9, value="Read:")
                rp.cell(row=2, column=10, value=ws.cell(row=rset, column=4).value)
                rp.cell(row=2, column=5, value="Write Average")
                rp.cell(row=3, column=5, value="Max")
                rp.cell(row=4, column=5, value="Min")
                rp.cell(row=2, column=12, value="Read Average")
                rp.cell(row=3, column=12, value="Max")
                rp.cell(row=4, column=12, value="Min")
                rp.cell(row=2, column=7, value="kIOPS")
                rp.cell(row=3, column=7, value="kIOPS")
                rp.cell(row=4, column=7, value="kIOPS")
                rp.cell(row=2, column=14, value="kIOPS")
                rp.cell(row=3, column=14, value="kIOPS")
                rp.cell(row=4, column=14, value="kIOPS")

            template()
            a1 = statistics.mean(writeSums)
            a2 = statistics.mean(readSums)
            max1 = max(writeSums)
            min1 = min(writeSums)
            max2 = max(readSums)
            min2 = min(readSums)

            cnt = 0
            cnt2 = 0
            for i in range(5, len(writeSums) + 5):
                rp.cell(row=i, column=3, value="Worker 1")
                rp.cell(row=i, column=4, value=writeSums[cnt])
                cnt += 1

            rp.cell(row=2, column=6, value=round((a1/1000), 1))
            rp.cell(row=3, column=6, value=round((max1/1000), 1))
            rp.cell(row=4, column=6, value=round((min1/1000), 1))

            for i in range(5, len(readSums) + 5):
                rp.cell(row=i, column=10, value="Worker 1")
                rp.cell(row=i, column=11, value=readSums[cnt2])
                cnt2 += 1

            rp.cell(row=2, column=13, value=round((a2/1000), 1))
            rp.cell(row=3, column=13, value=round((max2/1000), 1))
            rp.cell(row=4, column=13, value=round((min2/1000), 1))
            del wb[wb.sheetnames[0]]

            wb.save(save_dest)
        else:
            root.destroy()
    except IndexError:
        print("No file detected. Please try again.")
    finally:
        pass


if __name__ == '__main__':
    main()
