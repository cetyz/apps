# basic template for GUIs in case I forget
# for python 3.x

import tkinter as tk

class MainApplication(tk.Frame):
  def __init__(self, master, *args, **kwargs):
    tk.Frame.__init__(self, master, *args, **kwargs)
    self.master = master
    self.init_window()
    
  def init_window(self):
    self.variables = Variables(self.master)
    self.buttons = Buttons(self.master, self.variables)
    self.body = Body(self.master, self.variables, self.buttons)
    
# use this to put stuff you won't interact with
# like labels
class Body(tk.Frame):
  def __init__(self, master, variables, buttons):
  
    pass

# not necessarily just for buttons
# just use this if interaction is required
# like entry boxes or even dropdown menus
class Buttons(tk.Frame):
  def __init__(self, master, variables):
    
    pass
