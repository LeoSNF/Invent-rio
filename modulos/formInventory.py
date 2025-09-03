from tkinter import messagebox, filedialog
import tkinter as tk
from PIL import Image, ImageTk
import os
import ttkbootstrap as tb
from enum import Enum

class CommandsCodes(Enum):
    MB52 = 0
    BMBC = 1
    CALCINVT = 2
    EXTRACT_BOTH = 3
    CANCEL = -1

class FormInventory:
    def __init__(self, master: tb.Window):
        self.master = master
        master.title("Sanofi - Controle de Inventário")
        master.geometry("800x600")
        master.resizable(False, False)
        master.update_idletasks()

        width = master.winfo_width()
        height = master.winfo_height()
        x = (master.winfo_screenwidth() // 2) - (width // 2)
        y = (master.winfo_screenheight() // 2) - (height // 2)
        master.geometry(f'{width}x{height}+{x}+{y}')
        
        self.__commandCode = CommandsCodes.CANCEL.value
        # self.root = tb.Window(themename="flatly")
        # self.root.withdraw()
        self.mb52_path = ""
        self.bmbc_path = ""
        self.one_portfolio_path = ""
        self.inventory_path = ""
        self.build_ui_with_grid()

    def build_ui_with_grid(self):
        main_frame = tb.Frame(self.master, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1) 
        main_frame.rowconfigure(3, weight=1) 

        style = tb.Style()
        fonte = ("Arial", 12,)
        style.configure('primary.TButton', font= fonte)
        
        # logo
        try:
            logo_path = os.path.join("Logo", "logosanofi.png")
            temp_image = Image.open(logo_path)
            original_width, original_height = temp_image.size
            target_width = 355
            ratio = target_width / original_width
            target_height = int(original_height * ratio)
            logo_image = temp_image.resize((target_width, target_height), self.getResampleFilter())
            self.logo_photo = ImageTk.PhotoImage(logo_image)
            logo_label = tb.Label(main_frame, image=self.logo_photo)
        except Exception:
            logo_label = tb.Label(main_frame, text="Sanofi", font=("Arial", 28, "bold"), bootstyle="secondary")
        logo_label.grid(row=0, column=0, pady=(10, 10))

        # sap btn
        sap_actions_frame = tb.Frame(main_frame)
        sap_actions_frame.grid(row=1, column=0)
        tb.Button(sap_actions_frame, text="Extrair Ambos", command=lambda: self.setCommandCode(CommandsCodes.EXTRACT_BOTH.value), cursor="hand2", bootstyle="primary", width=15).pack(side=tk.LEFT, padx=10)
        tb.Button(sap_actions_frame, text="Extrair MB52", command=lambda: self.setCommandCode(CommandsCodes.MB52.value), cursor="hand2", bootstyle="primary", width=15).pack(side=tk.LEFT, padx=10)
        tb.Button(sap_actions_frame, text="Extrair BMBC", command=lambda: self.setCommandCode(CommandsCodes.BMBC.value), cursor="hand2", bootstyle="primary", width=15).pack(side=tk.LEFT, padx=10)
        
        tb.Separator(main_frame, bootstyle="secondary").grid(row=2, column=0, sticky="ew", padx=20, pady=20)

        # load file
        files_frame = tb.Frame(main_frame)
        files_frame.grid(row=3, column=0, sticky="nsew", padx=20)
        
        self.mb52_entry, self.mb52_status = self.create_file_input(files_frame, "MB52 (.xlsx)", "Caminho do arquivo...", 'mb52', "Arquivo MB52", ("Excel files", "*.xlsx"))
        self.bmbc_entry, self.bmbc_status = self.create_file_input(files_frame, "BMBC (.xlsx)", "Caminho do arquivo...", 'bmbc', "Arquivo BMBC", ("Excel files", "*.xlsx"))
        self.one_portfolio_entry, self.one_portfolio_status = self.create_file_input(files_frame, "One Portfolio (.xlsb)", "Caminho do arquivo...", 'portfolio', "One Portfolio", ("Excel Binary files", "*.xlsb"))
        self.inventory_entry, self.inventory_status = self.create_file_input(files_frame, "Inventário (.xlsx)", "Caminho do arquivo...", 'inventory', "Inventário", ("Excel files", "*.xlsx"))

        # rodapé
        footer_frame = tb.Frame(main_frame)
        footer_frame.grid(row=4, column=0, sticky="ew", pady=(20, 10))
        self.status_label = tb.Label(footer_frame, text="Carregue os arquivos para processar o inventário.", bootstyle="info")
        self.status_label.pack(pady=5)

        # btn area
        button_container = tb.Frame(footer_frame)
        button_container.pack(pady=5)
        self.calcular_button = tb.Button(button_container, text="Calcular Inventário", command=self.confirm_and_calculate, bootstyle="success", width=20)
        self.limpar_button = tb.Button(button_container, text="Limpar Arquivos", command=self.limpar_tudo, bootstyle="danger", width=20)
    
    def create_file_input(self, parent, button_text, placeholder, keyword_check, file_dialog_title, filetype):
        frame = tb.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        
        entry = tb.Entry(frame)
        
        button = tb.Button(frame, text=button_text, command=lambda e=entry: self._select_file_for_entry(e, keyword_check, file_dialog_title, filetype), bootstyle="secondary", width=20)
        button.pack(side=tk.LEFT, padx=(0, 10))

        def on_focus_in(event):
            if entry.get() == placeholder:
                entry.delete(0, tk.END)
                entry.config(foreground='')
        def on_focus_out(event):
            if not entry.get():
                entry.insert(0, placeholder)
                entry.config(foreground='gray')
        entry.insert(0, placeholder)
        entry.config(foreground='gray')
        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        status_label = tb.Label(frame, text="●", font=("Arial", 16, "bold"), foreground="red")
        status_label.pack(side=tk.LEFT, padx=(5, 0))
        entry.bind("<KeyRelease>", lambda event: self._validate_path_and_set_status(entry, status_label, keyword_check, placeholder))
        
        return entry, status_label

    def _validate_path_and_set_status(self, entry_widget, status_label, keyword_check, placeholder):
        path = entry_widget.get()
        if path and path != placeholder and keyword_check.lower() in path.lower():
            status_label.config(foreground="green")
        else:
            status_label.config(foreground="red")

    def getResampleFilter(self):
        try: return Image.Resampling.LANCZOS
        except AttributeError: return getattr(Image, 'LANCZOS', getattr(Image, 'ANTIALIAS', 1))

    def update_status_label(self):
        paths = [self.mb52_path, self.bmbc_path, self.one_portfolio_path, self.inventory_path]
        placeholders = ["Caminho ou nome do arquivo..."]
        if any(path for path in paths if path and path not in placeholders):
            self.status_label.config(text="Arquivos carregados. Pronto para calcular ou limpar.")
            self.calcular_button.pack(side=tk.LEFT, padx=10)
            self.limpar_button.pack(side=tk.LEFT, padx=10)
        else:
            self.status_label.config(text="Carregue os arquivos para processar o inventário.")
            self.calcular_button.pack_forget()
            self.limpar_button.pack_forget()

    def _select_file_for_entry(self, entry_widget, keyword_check, file_dialog_title, filetype):
        file_path = filedialog.askopenfilename(title=f"Selecionar {file_dialog_title}",filetypes=[filetype])
        status_label_map = {
            self.mb52_entry: self.mb52_status,
            self.bmbc_entry: self.bmbc_status,
            self.one_portfolio_entry: self.one_portfolio_status,
            self.inventory_entry: self.inventory_status
        }
        status_label = status_label_map.get(entry_widget)
        placeholder = "Caminho ou nome do arquivo..."

        if file_path:
            filename = os.path.basename(file_path)
            if keyword_check.lower() in filename.lower():
                entry_widget.config(foreground='')
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, file_path)
            else:
                messagebox.showerror("Arquivo Inválido", f"O nome do arquivo selecionado deve conter a palavra '{keyword_check}'.\n\nArquivo selecionado: {filename}")
        
        if status_label:
            self._validate_path_and_set_status(entry_widget, status_label, keyword_check, placeholder)
        self.update_status_label()

    def limpar_tudo(self):
        entries = [self.mb52_entry, self.bmbc_entry, self.one_portfolio_entry, self.inventory_entry]
        statuses = [self.mb52_status, self.bmbc_status, self.one_portfolio_status, self.inventory_status]

        for entry in entries:
            entry.delete(0, tk.END)
        for status in statuses:
            status.config(foreground="red")
        
        placeholder = "Caminho ou nome do arquivo..."
        for entry in entries:
            entry.insert(0, placeholder)
            entry.config(foreground='gray')

        self.update_status_label()
        messagebox.showinfo("Limpeza", "Todos os campos foram limpos!")

    def setCommandCode(self, commandCode: int):
        self.commandCode = commandCode
        self.master.destroy()

    def confirm_and_calculate(self):
        resposta = messagebox.askyesno(title="Confirmar Ação",  message="Tem certeza que deseja carregar os dados?",default=messagebox.NO )
        if resposta:
            self.setCommandCode(CommandsCodes.CALCINVT.value)

    def get_command_code(self):
        return self.__commandCode

    # @property
    # def mb52_path(self):
    #     return self.mb52_path
    
    # @mb52_path.setter
    # def mb52_path(self, value): 
    #     self.mb52_path = value

    # @property
    # def bmbc_path(self): 
    #     return self.bmbc_path
    
    # @bmbc_path.setter
    # def bmbc_path(self, value): 
    #     self.bmbc_path = value

    # @property
    # def one_portfolio_path(self):
    #     return self.one_portfolio_path
    
    # @one_portfolio_path.setter
    # def one_portfolio_path(self, value): 
    #     self.one_portfolio_path = value

    # @property
    # def inventory_path(self): 
    #     return self.inventory_path
    
    # @inventory_path.setter
    # def inventory_path(self, value): 
    #     self.inventory_path = value

    
