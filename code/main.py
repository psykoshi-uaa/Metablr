import metablr
import tkinter as tk
from tkinter import messagebox as msg
from tkinter.ttk import Notebook
from tkinter import filedialog

class App(tk.Tk):
	def __init__(self):
		super().__init__()

		self.title("Metablr")
		self.geometry("550x200")
		self.resizable(0, 0)

			# configure entry1 frame
		self.entry1_frame = tk.Frame(self, pady=20)

		self.file1_label = tk.Label(self.entry1_frame, text="File 1:", pady=10, padx=10)
		self.file1_label.pack(side=tk.LEFT, anchor="w")
		self.browse_files1 = tk.Button(self.entry1_frame, text="...", command=lambda:self.browse_files(self.file1_entry))
		self.browse_files1.pack(side=tk.LEFT)
		self.file1_entry = tk.Entry(self.entry1_frame, width=55)
		self.file1_entry.pack(side=tk.LEFT)
		self.reformat_button = tk.Button(self.entry1_frame, text="Reformat", command=lambda:self.reformat_button_pressed(self.file1_entry.get()))
		self.reformat_button.pack(side=tk.RIGHT)

			# configure entry2 frame
		self.entry2_frame = tk.Frame(self, pady=0)

		self.file2_label = tk.Label(self.entry2_frame, text="File 2:", pady=10, padx=10)
		self.file2_label.pack(side=tk.LEFT, anchor="w")
		self.browse_files2 = tk.Button(self.entry2_frame, text="...", command=lambda:self.browse_files(self.file2_entry))
		self.browse_files2.pack(side=tk.LEFT)
		self.file2_entry = tk.Entry(self.entry2_frame, width=55)
		self.file2_entry.pack(side=tk.LEFT)
		self.reformat_button = tk.Button(self.entry2_frame, text="Reformat", command=lambda:self.reformat_button_pressed(self.file2_entry.get()))
		self.reformat_button.pack(side=tk.RIGHT)

			# configure stitch buttons
		self.exit_button_frame = tk.Frame(self, width=200, padx=40, pady=20)
		self.exit_button = tk.Button(self.exit_button_frame, text="Exit", width=10, padx=30, command=self.exit_button_pressed)
		self.exit_button.pack(side=tk.TOP)

		self.stitch_button_frame = tk.Frame(self, width=200, padx=40, pady=20)
		self.stitch_button = tk.Button(self.stitch_button_frame, text="Stitch", width=10, padx=30, command=self.stitch_button_pressed)
		self.stitch_button.pack(side=tk.TOP)

			# pack stitch frames
		self.entry1_frame.pack(side=tk.TOP, anchor="w")
		self.entry2_frame.pack(side=tk.TOP, anchor="w")
		self.exit_button_frame.pack(side=tk.RIGHT)
		self.stitch_button_frame.pack(side=tk.LEFT)


	def browse_files(self, text):
		filename = filedialog.askopenfilename(initialdir = ".", title = "Select a File", filetypes = (("xlsx files", "*.xlsx*"), ("all files", "*.*")))

		text.delete(0, tk.END)
		text.insert(0, filename)
		return


	def save_as(self):
		filename = filedialog.asksaveasfilename(defaultextension=".xlsx")
		if filename is None:
			return		#open error dialog
		return filename

	 
	def stitch_button_pressed(self):
		stitch_args = ["-S", self.file1_entry.get(), self.file2_entry.get()]
		program_log = metablr.Program_Log()
		metablr.program_state((stitch_args), self.save_as(), program_log)
		program_log.print_log()
		return


	def reformat_button_pressed(self, filename):
		reformat_args = ["-R", filename]
		program_log = metablr.Program_Log()
		metablr.program_state((reformat_args), self.save_as(), program_log)
		program_log.print_log()
		return


	def exit_button_pressed(self):
		self.destroy()



if __name__ == "__main__":
	app = App()
	app.mainloop()


