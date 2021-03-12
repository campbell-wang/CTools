import ctypes
import os
import statistics
import tkinter.filedialog
from tkinter import *


def main():
	try:
		ctypes.windll.shcore.SetProcessDpiAwareness(1)
		root = Tk()
		root.withdraw()
		root.update()
		paths = tkinter.filedialog.askopenfilenames()
		if paths:
			root.withdraw()
			root.update()
			save_dest = tkinter.filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text File (.txt)", ".txt")])
			root.withdraw()
			root.destroy()

			iopsnums = []
			mbsnums = []
			tp = []
			tests = []

			for i in range(0, len(paths)):
				with open(str(paths[i]), encoding="utf-16-le") as f:
					lines = f.readlines()
					for line in lines:
						words = line.split()
						tp.append(line.split(':'))
						for n, word in enumerate(words):
							if word in ["IOPS]"]:
								iopsnums.append(float(words[n-1]))
							elif word in ["MB/s"] and words[n-1] != '*':
								mbsnums.append(float(words[n-1]))

			for i in range(tp.index(['\n']) + 1, len(tp)):
				if tp[i] != ['\n']:
					tests.append(tp[i][0])
				else:
					marker = i
					break

			for i in range(marker + 1, len(tp)):
				if tp[i] != ['\n']:
					tests.append(tp[i][0])
				else:
					break

			r_sq_q16_t1 = (mbsnums[0::4])[0::2]
			w_sq_q16_t1 = (mbsnums[0::4])[1::2]
			r_sq_q1_t1 = (mbsnums[1::4])[0::2]
			w_sq_q1_t1 = (mbsnums[1::4])[1::2]
			r_q128t16 = (iopsnums[2::4])[0::2]
			w_q128t16 = (iopsnums[2::4])[1::2]
			r_q1t1 = (iopsnums[3::4])[0::2]
			w_q1t1 = (iopsnums[3::4])[1::2]

			rsq161avg = round(statistics.mean(r_sq_q16_t1), 1)
			rsq161max = round(max(r_sq_q16_t1), 1)
			rsq161min = round(min(r_sq_q16_t1), 1)
			rsq11avg = round(statistics.mean(r_sq_q1_t1), 1)
			rsq11max = round(max(r_sq_q1_t1), 1)
			rsq11min = round(min(r_sq_q1_t1), 1)
			r12816avg = round(statistics.mean(r_q128t16)/1000, 1)
			r12816max = round(max(r_q128t16)/1000, 1)
			r12816min = round(min(r_q128t16)/1000, 1)
			r11avg = round(statistics.mean(r_q1t1)/1000, 1)
			r11max = round(max(r_q1t1)/1000, 1)
			r11min = round(min(r_q1t1)/1000, 1)

			wsq161avg = round(statistics.mean(w_sq_q16_t1), 1)
			wsq161max = round(max(w_sq_q16_t1), 1)
			wsq161min = round(min(w_sq_q16_t1), 1)
			wsq11avg = round(statistics.mean(w_sq_q1_t1), 1)
			wsq11max = round(max(w_sq_q1_t1), 1)
			wsq11min = round(min(w_sq_q1_t1), 1)
			w12816avg = round(statistics.mean(w_q128t16)/1000, 1)
			w12816max = round(max(w_q128t16)/1000, 1)
			w12816min = round(min(w_q128t16)/1000, 1)
			w11avg = round(statistics.mean(w_q1t1)/1000, 1)
			w11max = round(max(w_q1t1)/1000, 1)
			w11min = round(min(w_q1t1)/1000, 1)

			with open(save_dest, "w+") as w:
				w.write("[Read Average, Max, Min]\n")
				w.write(tests[1].strip() + ": " + str(rsq161avg) + " MB/s | Max: " + str(rsq161max) + " MB/s | Min: " + str(rsq161min) + " MB/s\n")
				w.write(tests[2].strip() + ": " + str(rsq11avg) + " MB/s | Max: " + str(rsq11max) + " MB/s | Min: " + str(rsq11min) + " MB/s\n")
				w.write(tests[3].strip() + ": " + str(r12816avg) + " kIOPS | Max: " + str(r12816max) + " kIOPS | Min: " + str(r12816min) + " kIOPS\n")
				w.write(tests[4].strip() + ": " + str(r11avg) + " kIOPS | Max: " + str(r11max) + " kIOPS | Min: " + str(r11min) + " kIOPS\n\n")

				w.write("[Write Average, Max, Min]\n")
				w.write(tests[6].strip() + ": " + str(wsq161avg) + " MB/s | Max: " + str(wsq161max) + " MB/s | Min: " + str(wsq161min) + " MB/s\n")
				w.write(tests[7].strip() + ": " + str(wsq11avg) + " MB/s | Max: " + str(wsq11max) + " MB/s | Min: " + str(wsq11min) + " MB/s\n")
				w.write(tests[8].strip() + ": " + str(w12816avg) + " kIOPS | Max: " + str(w12816max) + " kIOPS | Min: " + str(w12816min) + " kIOPS\n")
				w.write(tests[9].strip() + ": " + str(w11avg) + " kIOPS | Max: " + str(w11max) + " kIOPS | Min: " + str(w11min) + " kIOPS\n\n")

				w.close()
		else:
			root.destroy()
	except IndexError:
		print("No file detected. Please try again.")
	finally:
		pass


if __name__ == '__main__':
	main()
