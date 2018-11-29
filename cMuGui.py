try:
    import os, sys
    import Tkinter as tk        # python3: tkinter
    from tkFileDialog import *  # python3: tkinter.filedialog instead
    from tkMessageBox import askokcancel, showinfo
except:
    print("ExceptionERROR: Missing fundamental packages (required: os, sys, Tkinter).")

try:
    import cIO as cio
    import make_mu as mm
except:
    print("ExceptionERROR: Missing own packages (required: cIO, make_mu).")


class MyGui(tk.Frame):
    def __init__(self, master=None):
        # global variables
        self.max_columnspan = 5
        self.path = os.path.dirname(os.path.abspath(__file__)) + "\\"
        self.path_raster = os.path.dirname(os.path.abspath(__file__)) + "\\geodata\\rasters\\"

        # Create and fill MU object
        self.mu_active = []
        self.mu_workbook = cio.Read(self.path + "mu.xlsx", 0)
        self.mu = self.mu_workbook.read_column("B", 5)
        self.mu.insert(0, "CLEAR ALL")  # required key

        # Construct the Frame object
        ### see lecture slides (inheritance)

        # Pack master Window
        self.pack(expand=True, fill=tk.BOTH)

        # Create a tk.IntVar() that will be controlled by a check button
        self.use_check_button = tk.IntVar()

        # ARRANGE GEOMETRY
        ### create self.dx and self.dy, which define distance holders in x and y-directions (use 8 pixels)

        # width and height of the window
        ww = 450  # controls window width
        wh = 240  # controls window height
        wx = (self.master.winfo_screenwidth() - ww) / 2   # position relative to screen width and ww
        wy = (self.master.winfo_screenheight() - wh) / 2  # position relative to screen height and wh
        self.master.geometry("%dx%d+%d+%d" % (ww, wh, wx, wy))  # set window height and location
        self.master.title("Morphological Unit Maker")  # window title

        # DROP DOWN MENU BAR
        self.mbar = tk.Menu(self)  # create new menubar
        self.master.config(menu=self.mbar)  # attach it to the menubar
        # ADD A CLOSE MENU TO THE MENU BAR
        self.closemenu = tk.Menu(self.mbar, tearoff=0)  # create new menu
        self.mbar.add_cascade(label="Close", menu=self.closemenu)  # attach it to the menubar
        self.closemenu.add_command(label="Credits", command=lambda: self.show_credits())
        self.closemenu.add_command(label="Quit programm", command=lambda: self.myquit())
        # for changing menu entries use :
        # self.closemenu.entryconfig(int(entrynumber), label="New", foreground="other", command=lambda: self.new_f())

        # CHECKBUTTONS
        self.cb_mu = tk.Checkbutton(self, text="Analyze all Morphological Units", onvalue=1, offvalue=0, command=lambda: self.set_mu("flip_list"), variable=self.use_check_button)
        self.cb_mu.grid(sticky=tk.W, row=0, column=0, columnspan=self.max_columnspan, padx=self.xd, pady=self.yd)
        self.cb_mu.select()

        # BUTTONS
        self.b_select = tk.Button(self, text="Add selected", command=lambda: ### complete: call self.set_mu() with argument mu="get_selection"
        self.b_select.grid(sticky=tk.EW, row=1, column=self.max_columnspan - 1, padx=self.xd)

        self.b_show_mu = tk.Button(self, width=12, bg="white", text="Show selected\nMUs", command=lambda: ### complete: call self.shout_list() with the argument the_list=self.mu_active
        self.b_show_mu.grid(sticky=tk.EW, row=2, column=self.max_columnspan - 1, padx=self.xd)

        self.b_change_dir = tk.Button(self, width=12, bg="white", text="Change Raster input directory", command=lambda: ### complete: call self.set_path_raster(), without any argument
        self.b_change_dir.grid(sticky=tk.EW, row=4, column=0, columnspan=self.max_columnspan, padx=self.xd, pady=self.yd)

        self.b_run = tk.Button(self, bg="forest green", text="RUN", command=lambda: ### complete: call mm.main() function with input arguments (1) self.path_raster and (2) self.mu_active
        self.b_run.grid(sticky=tk.EW, row=6, rowspan=2, column=0, columnspan=self.max_columnspan, padx=self.xd, pady=self.yd)

        # LABELS
        self.l_select = tk.Label(self, text="Select a Morphological Unit")  # covers scroll bar
        self.l_select.grid(sticky=tk.W, row=1, rowspan=3, column=0, padx=self.xd, pady=self.yd)

        self.l_r_path = tk.Label(self, text="Current: " + str(self.path_raster))  # covers scroll bar
        self.l_r_path.grid(sticky=tk.W, row=5, column=0, columnspan=self.max_columnspan, padx=self.xd, pady=self.yd)

        # ENTRIES (uncomment for usage)
        # self.entry = tk.Entry(self, width=10, textvariable=self.raster_dir)
        # self.entry.grid(sticky=tk.W, row=2, column=1, padx=self.xd, pady=self.yd)

        # LISTBOX WITH SCROLL BAR
        self.sb_mu = tk.Scrollbar(self, orient=tk.VERTICAL)
        self.sb_mu.grid(sticky=tk.W, row=1, rowspan=3, column=3, padx=0, pady=self.yd)
        self.lb_mu = tk.Listbox(self, height=3, width=15, yscrollcommand=self.sb_mu.set)
        self.fill_listbox()

        # SET START UP SELECTION
        self.set_mu("flip_list")

    def fill_listbox(self):
        # update habitat morphological units listbox
        for e in self.mu:
            self.lb_mu.insert(tk.END, e)
        self.lb_mu.grid(sticky=tk.E, row=1, rowspan=3, column=2, padx=0, pady=self.yd)
        self.sb_mu.config(command=self.lb_mu.yview)

    def myquit(self):
        tk.Frame.quit(self)

    def set_path_raster(self):
        self.path_raster = askdirectory(initialdir=".") + "/"
        self.l_r_path.config(text="Current: " + str(self.path_raster))

    def set_mu(self, mu):
        # mu = STR - morphological unit according to mu.xlsx OR local keyword
        if mu == "flip_list":
            if self.use_check_button.get():
                # deactivate selection menu if checkbutton is selected
                self.l_select["state"] = "disabled"
                self.b_select["state"] = "disabled"
                ### set the state of self.lb_mu to "disabled" -- see the difference
                self.mu_active = self.mu[1:]
            else:
                # activate selection menu if checkbutton is deselected
                self.mu_active = []
                self.l_select["state"] = "normal"
                self.b_select["state"] = "normal"
                ### set the state of self.lb_mu to "normal" -- see the difference

        if mu == "get_selection":
            selected_item = [self.mu[int(item)] for item in self.lb_mu.curselection()][0]
            if not(selected_item.lower() == "clear all"):
                self.mu_active.append(selected_item)
            else:
                self.mu_active = []

    def shout_list(self, the_list):
        msg = "Selected Morphological Units:\n"
        msg = msg + "\n   > " + "\n   > ".join(the_list)
        showinfo("Applied MUs", msg)

    def show_credits(self):
        msg = "GUI Framework from Python lecture (Fall 2018)\nAuthor: Sebastian Schwindt\nInstitute: Pasternack Lab, UC Davis \n\nEmail: sschwindt[at]ucdavis.edu"
        showinfo("Credits", msg)


### enable script to run stand-alone
# if __name__ == "__main__":
#     MyGui().mainloop()
