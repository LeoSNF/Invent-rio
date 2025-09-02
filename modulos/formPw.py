from tkinter import *

class FormPw:
    def __init__(self, title='Credentials', txtWidth=32):
        fontInfo = ('Courier', 8)
        win = Tk()
        win.title(title)
        win.geometry('280x150')
        try:
            win.iconphoto(True, PhotoImage(file=r'modulos\logo.png'))
        except:
            pass

        frmMain = Frame(win)
        frmMain.grid()

        lblTitle = Label(frmMain, text='\nCredentials:\n', font=fontInfo)
        lblTitle.grid(column=1)

        lblUserId = Label(frmMain, text='Usuário:', font=fontInfo)
        lblUserId.grid(column=0, row=1)

        txtUserId = Entry(frmMain, width=txtWidth)
        txtUserId.grid(column=1,row=1)

        lblUserPW = Label(frmMain, text='Senha:', font=fontInfo)
        lblUserPW.grid(column=0, row=2)

        txtUserPW = Entry(frmMain, width=txtWidth, show='*')
        txtUserPW.grid(column=1,row=2)

        btnOk = Button(frmMain, text='Confirmar', command=lambda: self.btnOk_click(win, txtUserId, txtUserPW))
        btnOk.grid(column=1, row=3, sticky=W, padx=15, pady=10)

        btnCancel = Button(frmMain, text='Cancelar', command=lambda: self.btnCancel_click(win))
        btnCancel.grid(column=1, row=3, sticky=E, padx=30, pady=10)

        win.mainloop()

    def btnOk_click(self, win, txtUserId, txtUserPW):
        self.__userId = txtUserId.get()
        self.__userPw = txtUserPW.get()
        win.destroy()

    def btnCancel_click(self, win):
        win.destroy()

    @property
    def userId(self):
        try: return self.__userId
        except: pass
    
    @property
    def userPw(self):
        try: return self.__userPw
        except: pass