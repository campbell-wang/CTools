import ctypes
from tkinter import *
import tkinter.filedialog
import statistics


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
						tp.append(line.split(':'))
						words = line.split()
						for n, word in enumerate(words):
							if word in ["IOPS]"]:
								iopsnums.append(float(words[n-1]))
							elif word in ["Random", "Sequential"] and words[n-1] == 'MB/s':
								mbsnums.append(float(words[n-2]))

			for i in range(tp.index(['\n'])+1, len(tp)):
				if tp[i] != ['\n']:
					tests.append(tp[i][0].strip())
				else:
					break

			r_sq_q32t1 = (mbsnums[0::2])
			w_sq_q32t1 = (mbsnums[1::2])
			del mbsnums
			rq32t8 = (iopsnums[0::6])
			wq32t8 = (iopsnums[1::6])
			rq32t1 = (iopsnums[2::6])
			wq32t1 = (iopsnums[3::6])
			rq1t1 = (iopsnums[4::6])
			wq1t1 = (iopsnums[5::6])
			del iopsnums

			aavg = round(statistics.mean(r_sq_q32t1), 1)
			amax = round(max(r_sq_q32t1), 1)
			amin = round(min(r_sq_q32t1), 1)
			bavg = round(statistics.mean(w_sq_q32t1), 1)
			bmax = round(max(w_sq_q32t1), 1)
			bmin = round(min(w_sq_q32t1), 1)
			cavg = statistics.mean(rq32t8)
			cmax = max(rq32t8)
			cmin = min(rq32t8)
			davg = statistics.mean(wq32t8)
			dmax = max(wq32t8)
			dmin = min(wq32t8)
			eavg = statistics.mean(rq32t1)
			emax = max(rq32t1)
			emin = min(rq32t1)
			favg = statistics.mean(wq32t1)
			fmax = max(wq32t1)
			fmin = min(wq32t1)
			gavg = statistics.mean(rq1t1)
			gmax = max(rq1t1)
			gmin = min(rq1t1)
			havg = statistics.mean(wq1t1)
			hmax = max(wq1t1)
			hmin = min(wq1t1)

			tp = [cavg, cmax, cmin, davg, dmax, dmin, eavg, emax, emin, favg, fmax, fmin, gavg, gmax, gmin, havg, hmax, hmin]
			for idx, item in enumerate(tp):
				tp[idx] = round(item/1000, 1)

			with open(save_dest, "w+") as w:
				w.write("[Sequential Read/Write Average, Max, Min]\n")
				w.write(tests[0] + ": " + str(aavg) + " MB/s | Max:  " + str(amax) + " MB/s | Min:  " + str(amin) + " MB/s\n")
				w.write(tests[1] + ": " + str(bavg) + " MB/s | Max:  " + str(bmax) + " MB/s | Min:  " + str(bmin) + " MB/s\n\n")

				w.write("[Write Average, Max, Min]\n")
				w.write(tests[2] + ": " + str(tp[0]) + " kIOPS | Max: " + str(tp[1]) + " kIOPS | Min: " + str(tp[2]) + " kIOPS\n")
				w.write(tests[3] + ": " + str(tp[3]) + " kIOPS | Max: " + str(tp[4]) + " kIOPS | Min: " + str(tp[5]) + " kIOPS\n")
				w.write(tests[4] + ": " + str(tp[6]) + " kIOPS | Max: " + str(tp[7]) + " kIOPS | Min: " + str(tp[8]) + " kIOPS\n")
				w.write(tests[5] + ": " + str(tp[9]) + " kIOPS | Max: " + str(tp[10]) + " kIOPS | Min: " + str(tp[11]) + " kIOPS\n")
				w.write(tests[6] + ": " + str(tp[12]) + " kIOPS | Max: " + str(tp[13]) + " kIOPS | Min: " + str(tp[14]) + " kIOPS\n")
				w.write(tests[7] + ": " + str(tp[15]) + " kIOPS | Max: " + str(tp[16]) + " kIOPS | Min: " + str(tp[17]) + " kIOPS\n\n")

				w.close()
		else:
			root.destroy()

	except IndexError:
		print("No file detected. Please try again.")
	finally:
		pass


if __name__ == '__main__':
	main()
