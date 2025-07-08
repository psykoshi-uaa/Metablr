import metablr
import tkinter as tk
from tkinter import messagebox as msg
from tkinter.ttk import Notebook
from tkinter import filedialog

class EventWindow(tk.Tk):
	def __init__(self, label_text, button_text):
		super().__init__()
		self.title("")
		self.geometry("256x100")
		
		self.success_text = tk.Label(self, text=label_text, pady=20)
		self.ok_button = tk.Button(self, text=button_text, width=10, command=self.button_pressed)
		
		self.success_text.pack(side=tk.TOP)
		self.ok_button.pack(side=tk.TOP)
		
		
	def button_pressed(self):
		self.destroy()


		
class App(tk.Tk):
	def __init__(self):
		super().__init__()

		self.title("Metablr")
		#self.geometry("650x380")
		self.resizable(0, 0)

		#configure menu
		#self.menu_bar = tk.Menu(self)
		#self.config(menu=self.menu_bar)
		#self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
		#self.export_menu = tk.Menu(self.menu_bar, tearoff=0)
		#self.menu_bar.add_cascade(label="File", menu=self.file_menu)
		#self.file_menu.add_command(label="Update", command=self.update_program)
		#self.file_menu.add_cascade(label="Export", menu=self.export_menu)
		#self.export_menu.add_command(label="Export as Summary Table", command=self.stitch_button_pressed)
		#self.export_menu.add_command(label="Export as Data Table", command=self.reformat_button_pressed)
		#self.file_menu.add_command(label="Exit", command=self.exit_button_pressed)


		#configure body
		self.filler_frameT = tk.Frame(self, width=25, height=25)
		self.filler_frameT.pack(side=tk.TOP)
		self.filler_frameL = tk.Frame(self, width=25)
		self.filler_frameL.pack(side=tk.LEFT)
		self.filler_frameR = tk.Frame(self, width=25)
		self.filler_frameR.pack(side=tk.RIGHT)

			# configure entry1 frame
		self.entry1_frame = tk.Frame(self, pady=20, relief="sunken", borderwidth=2)

		self.file1_label = tk.Label(self.entry1_frame, text="Positive Mode")
		self.file1_label.pack(side=tk.TOP)

		self.entry1_marginL = tk.Frame(self.entry1_frame, width=25)
		self.entry1_marginL.pack(side=tk.LEFT)
		self.entry1_marginR = tk.Frame(self.entry1_frame, width=25)
		self.entry1_marginR.pack(side=tk.RIGHT)

		self.CD_frame1 = tk.Frame(self.entry1_frame)
		self.file1_CD_label = tk.Label(self.CD_frame1, text="CD Table:")
		self.file1_CD_label.pack(side=tk.TOP, anchor='w')
		self.CD_browse_files1 = tk.Button(self.CD_frame1, text="...", padx=10, command=lambda:self.browse_files(self.file1_CD_entry))
		self.CD_browse_files1.pack(side=tk.LEFT)
		self.file1_CD_entry = tk.Entry(self.CD_frame1, width=80)
		self.file1_CD_entry.pack(side=tk.LEFT)
		self.CD_frame1.pack(side=tk.TOP)

		self.filler_entry1 = tk.Label(self.entry1_frame, pady=5)
		self.filler_entry1.pack(side=tk.TOP)

		self.Inp_frame1 = tk.Frame(self.entry1_frame)
		self.file1_Inp_label = tk.Label(self.Inp_frame1, text="Input File:")
		self.file1_Inp_label.pack(side=tk.TOP, anchor='w')
		self.Inp_browse_files1 = tk.Button(self.Inp_frame1, text="...", padx=10, command=lambda:self.browse_files(self.file1_Inp_entry))
		self.Inp_browse_files1.pack(side=tk.LEFT)
		self.file1_Inp_entry = tk.Entry(self.Inp_frame1, width=80)
		self.file1_Inp_entry.pack(side=tk.LEFT)
		self.Inp_frame1.pack(side=tk.TOP)

		self.filler_entrys = tk.Label(self, pady=5)

			# configure entry2 frame
		self.entry2_frame = tk.Frame(self, pady=20, relief="sunken", borderwidth=2)

		self.file2_label = tk.Label(self.entry2_frame, text="Negative Mode")
		self.file2_label.pack(side=tk.TOP)

		self.entry2_marginL = tk.Frame(self.entry2_frame, width=25)
		self.entry2_marginL.pack(side=tk.LEFT)
		self.entry2_marginR = tk.Frame(self.entry2_frame, width=25)
		self.entry2_marginR.pack(side=tk.RIGHT)

		self.CD_frame2 = tk.Frame(self.entry2_frame)
		self.file2_CD_label = tk.Label(self.CD_frame2, text="CD Table:")
		self.file2_CD_label.pack(side=tk.TOP, anchor='w')
		self.CD_browse_files2 = tk.Button(self.CD_frame2, text="...", padx=10, command=lambda:self.browse_files(self.file2_CD_entry))
		self.CD_browse_files2.pack(side=tk.LEFT)
		self.file2_CD_entry = tk.Entry(self.CD_frame2, width=80)
		self.file2_CD_entry.pack(side=tk.LEFT)
		self.CD_frame2.pack(side=tk.TOP)

		self.filler_entry2 = tk.Label(self.entry2_frame, pady=5)
		self.filler_entry2.pack(side=tk.TOP)

		self.Inp_frame2 = tk.Frame(self.entry2_frame)
		self.file2_Inp_label = tk.Label(self.Inp_frame2, text="Input File:")
		self.file2_Inp_label.pack(side=tk.TOP, anchor='w')
		self.Inp_browse_files2 = tk.Button(self.Inp_frame2, text="...", padx=10, command=lambda:self.browse_files(self.file2_Inp_entry))
		self.Inp_browse_files2.pack(side=tk.LEFT)
		self.file2_Inp_entry = tk.Entry(self.Inp_frame2, width=80)
		self.file2_Inp_entry.pack(side=tk.LEFT)
		self.Inp_frame2.pack(side=tk.TOP)

			# configure stitch buttons
		self.exit_button_frame = tk.Frame(self, width=200, pady=20)
		self.exit_button = tk.Button(self.exit_button_frame, text="Exit", width=10, padx=30, command=self.exit_button_pressed)
		self.exit_button.pack(side=tk.TOP)

		self.stitch_button_frame = tk.Frame(self, width=200, padx=5, pady=20)
		self.stitch_button = tk.Button(self.stitch_button_frame, text="Export Summary", width=10, padx=30, command=self.stitch_button_pressed)
		self.stitch_button.pack(side=tk.TOP)

		self.reformat_button_frame = tk.Frame(self, width=200, padx=5, pady=20)
		self.reformat_button = tk.Button(self.reformat_button_frame, text="Export Data Table", width=10, padx=30, command=self.reformat_button_pressed)
		self.reformat_button.pack(side=tk.TOP)

		self.filler_entrys2 = tk.Label(self, pady=5)

			# pack stitch frames
		self.entry1_frame.pack(side=tk.TOP, anchor="w")
		self.filler_entrys.pack(side=tk.TOP)
		self.entry2_frame.pack(side=tk.TOP, anchor="w")
		self.filler_entrys2.pack(side=tk.TOP)
		self.exit_button_frame.pack(side=tk.RIGHT)
		self.stitch_button_frame.pack(side=tk.LEFT)
		self.reformat_button_frame.pack(side=tk.LEFT)


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
		stitch_args = ["-S", self.file1_CD_entry.get(), self.file2_CD_entry.get(), self.save_as()]
		program_log = metablr.Program_Log()
		metablr.program_state((stitch_args), program_log)
		program_log.print_log()

		self.event_window(program_log, "Success", "OK", "Error: check both xlsx files", "OK")
			
		return


	def reformat_button_pressed(self):
		reformat_args = ["-R", self.file1_CD_entry.get(), self.file1_Inp_entry.get(), self.file2_CD_entry.get(), self.file2_Inp_entry.get(), self.save_as()]
		program_log = metablr.Program_Log()
		metablr.program_state((reformat_args), program_log)
		program_log.print_log()
		
		self.event_window(program_log, "Success", "OK", "Error: ensure CD tables and Input files are xlsx files", "OK")
			
		return
		
		
	def event_window(self, program_log, event_text_success, event_button_success, event_text_error, event_button_error):
		event_text = event_text_success
		event_button = event_button_success
		if (program_log.get_error_count() > 0):
			event_text =  event_text_error
			event_button = event_button_error
		event = EventWindow(event_text, event_button)	
		event.mainloop()

	
	def update_program(self):
		print("update")


	def exit_button_pressed(self):
		self.destroy()



if __name__ == "__main__":
	app = App()
	app.mainloop()
