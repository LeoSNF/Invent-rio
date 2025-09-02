from tkinter import messagebox
import tkinter as tk
from enum import Enum

class MsgBoxOptions(Enum):
    INFO = 1
    WARNING = 2
    WARNINGOPT = 3
    ERROR = 4

class MsgBox:
    def __init__(self, isToShow:int=1):
        self.isToShow = isToShow
        self.typeMsg = None
        self.bodyMsg = None
   
    def showMsgBox(self, typeMsg:Enum, headMsg:str,  bodyMsg:str):
        '''bodyMsg opções: info, warning ou error'''
        msgOpt = typeMsg.value
        if self.isToShow:
            root = tk.Tk()
            root.wm_attributes("-topmost", 1)
            root.withdraw()
            if msgOpt == 1:
                messagebox.showinfo(headMsg, bodyMsg)
            elif msgOpt == 2:
                messagebox.showwarning(headMsg, bodyMsg)
            elif msgOpt == 3:
                self.__answer = messagebox.askyesno(headMsg, bodyMsg, default="no")
            elif msgOpt == 4:
                messagebox.showerror(headMsg, bodyMsg, parent=root)
            else:
                raise ValueError('Invalid option')
            root.destroy()

    @property
    def answer(self):
        return self.__answer
