
from tkinter import Tk, Label, Button, Entry, END, messagebox
from project_class import Project
from data import Data


GREEN = "#71C9CE"
LIGHT_GREEN = "#A6E3E9"
LIGHTEST_GREEN = "#CBF1F5"
BACKGROUND = "#E3FDFD"


BUTTON_TEXT = ("Lucida Sans Unicode", 10)
BUTTON_TEXT_ = ("Lucida Sans Unicode", 10, "bold")
HEAD_TEXT = ("Lucida Sans Unicode", 14)
SMALL_BUTTON_TEXT = ("Lucida Sans Unicode", 6)

last_project = None
time_spent = 0


class AppInterface:
    def __init__(self, project_class: Project, data: Data, config):
        self.project_class = project_class
        self.data = data
        self.config = config

        self.normalize_on_exit = False

        self.window = Tk()
        self.window.title("Work hours")
        self.window.config(padx=8, pady=8, bg=BACKGROUND)


        self.label_set_project = Label(text="Enter the project number:", bg=BACKGROUND, font=HEAD_TEXT)
        self.label_set_project.grid(column=0, row=0, columnspan=4, pady=8, sticky="wn")
        # Buton
        active_project_but = Button(text="Project activate", width=16, height=1, bg="white", font=BUTTON_TEXT ,command=self.activate_project)
        active_project_but.grid(column=2, row=1, columnspan=2, sticky="e")

        self.volne_button = Button(text=self.config["default_projects"][0], width=8, height=1, bg="white", font=BUTTON_TEXT, command=self.quick_fn1)
        self.volne_button.grid(column=0, row=2, sticky="e")

        self.sigdes_button = Button(text=self.config["default_projects"][1], width=8, bg="white", height=1, font=BUTTON_TEXT,command=self.quick_fn2)
        self.sigdes_button.grid(column=1, row=2, sticky="w")

        self.clear = Button(text="Clear text", width=8, height=1, bg="white", font=BUTTON_TEXT, command=self.clear_field)
        self.clear.grid(column=2, row=2, pady=1, sticky="w")

        self.pause = Button(text="Pause", width=8, height=1, bg="white", font=BUTTON_TEXT, command=self.pause_fn)
        self.pause.grid(column=3, row=2, pady=1, sticky="e")
        
        self.normalize_button = Button(self.window, text="norm", width=4, height=1, bg="white", font=SMALL_BUTTON_TEXT, command=self.confirm_normalize)
        self.normalize_button.grid(column=3, row=0, sticky="ne", padx=1, pady=1)


        # Entry
        self.entry = Entry(width=32)
        self.entry.focus()
        self.entry.grid(column=0, columnspan=2, row=1)

        # reloop
        self.window.mainloop()

    def pause_fn(self):
        if self.project_class.project_name == "":
            print("No project active.")
            return
        self.pause_project = self.project_class.project_name
        self.pause_time = self.project_class.stop_time()
        self.pause.config(text="Resume", width=8, height=1, command=self.resume_fn)
        self.project_class.project_pause = True
        self.label_set_project.config(bg=BACKGROUND, text=f"Project '{self.project_class.project_name}' paused.")
        self.window.config(bg=BACKGROUND)
        self.entry.config(bg="white")
      #  self.label_set_project.grid(column=0, row=0)
    

    def resume_fn(self):
        self.project_class.start_time(inc=self.pause_time)
        self.pause.config(text="Pause", width=8, height=1, command=self.pause_fn)
        self.project_class.project_pause = False
        self.label_set_project.config(bg=LIGHT_GREEN, text=f"Project '{self.project_class.project_name}' resumed.")
        self.window.config(bg=LIGHT_GREEN)
        self.entry.config(bg=LIGHTEST_GREEN)
      #  self.label_set_project.grid(column=0, row=0)


    def quick_fn2(self):
        self.entry.delete(0, END)
        self.entry.insert(0, self.config["default_projects"][1])

    def quick_fn1(self):
        self.entry.delete(0, END)
        self.entry.insert(0, self.config["default_projects"][0])

    def init_col_but(self):
        self.sigdes_button.config(bg="white")
        self.volne_button.config(bg="white")

    def cancel_project(self):
        self.data.remove_last()
        self.clear_field()
        self.label_set_project.config(text=f"Canceled: '{self.project_class.project_name}'")
        self.project_class.cancel()

     #   self.label_set_project.grid(column=0, row=0)
        self.window.config(bg=BACKGROUND)
        self.label_set_project.config(bg=BACKGROUND)
        self.entry.config(bg="white")
        self.init_col_but()

        self.clear.config(text="Clear text", command=self.clear_field)
  


    def activate_project(self):
        if self.entry.get() == self.config["default_projects"][0] or self.entry.get() == self.config["default_projects"][1]:
            project = self.entry.get()
        
        elif self.entry.get() == "":
            return
        
        else:
            project = self.entry.get().upper()
        self.clear_field()

            
#            recent_value = pyperclip.paste()
#            while True:
#                tmp_value = pyperclip.paste()
#                if tmp_value != recent_value:
#                    recent_value = tmp_value
#                    project = recent_value
#                    break
#                elif keyboard.is_pressed("x"):
#                    break


        if self.project_class.cancel_var:
            self.project_class.cancel_var = False
            pass
        elif self.project_class.project_name == project:
            self.label_set_project.config(text=f"Project '{project}' already active.")
         #   self.label_set_project.grid(column=0, row=0)
        else:
            if self.project_class.project_pause:
                self.project_class.project_pause = False
                self.data.write_data(self.pause_project, self.pause_time)
                pass
            else:
                try:
                    self.project_class.stop_time()
                    self.data.write_data(self.project_class.project_name, self.project_class.project_time)
                except AttributeError:
                    pass

        self.project_class.is_active = True
        self.project_class.start_time()
        self.project_class.project_name = project


        self.label_set_project.config(text=f"Active project: '{project}'.")
        self.window.config(bg=LIGHT_GREEN)
        self.label_set_project.config(bg=LIGHT_GREEN)
        self.entry.config(bg=LIGHTEST_GREEN)


        if project == self.config["default_projects"][1]:
            self.sigdes_button.config(bg=GREEN)
            self.volne_button.config(bg="white")
        elif project == self.config["default_projects"][0]:
            self.sigdes_button.config(bg="white")
            self.volne_button.config(bg=GREEN)
        else:
            self.init_col_but()

     #   self.label_set_project.grid(column=0, row=0)
        
        # Cancel button
        self.clear.config(text="Cancel", command=lambda: (self.cancel_project() if messagebox.askquestion("Cancel", f"Cancel last {project} project record?") == 'yes' else None))
        self.pause.config(text="Pause", width=8, height=1, command=self.pause_fn)


    def clear_field(self):
        self.label_set_project.config(text="Enter the project number:")
        self.entry.delete(0,END)
        self.init_col_but()
        
    
    def confirm_normalize(self):
        # This function will be called when the normalize button is clicked.
        response = messagebox.askquestion("Normalize hours", f"Normalize month table to {self.config['normalize_hours']} hours per day?")
        if response == 'yes':
            self.normalize_on_exit = True
            self.window.destroy()   # This will close the program after normalizing

