import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import os
from enum import Enum

class CommandsCodes(Enum):
    MB52 = 0
    BMBC = 1
    CALCINVT = 2
    EXTRACT_BOTH = 3
    CANCEL = -1

class FormInventory:
    def __init__(self, master: tk.Tk): 
        self.master = master
        master.title("Sanofi - Controle de Inventário")
        master.geometry("800x500")
        master.resizable(False, False)
        
        master.update_idletasks()
        width = master.winfo_width()
        height = master.winfo_height()
        x = (master.winfo_screenwidth() // 2) - (width // 2)
        y = (master.winfo_screenheight() // 2) - (height // 2)
        master.geometry(f'{width}x{height}+{x}+{y}')

        self.__commandCode = CommandsCodes.CANCEL.value
        
        self.build_ui_with_grid()

    def build_ui_with_grid(self):
        main_frame = ttk.Frame(self.master, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1) 

        style = ttk.Style()
        style.configure('TFrame', background='#E0E0E0')
        style.configure('TLabel', background='#E0E0E0')
        
        fonte_botoes_superiores = ("Arial", 12, "bold")
        fonte_botoes_inferiores = ("Arial", 10)
        cor_roxa_sanofi = "#5A1761"
        cor_roxa_ativa = "#7A2F80"
        cor_sucesso = "#198754"
        cor_perigo = "#DC3545"
        
        style.configure('Secondary.TButton', font=("Arial", 10))
        style.configure('Tall.TEntry', font=('Segoe UI', 11), padding=(5, 7)) 

        header_frame = tk.Frame(main_frame, bg='#E0E0E0') 
        header_frame.grid(row=0, column=0, sticky="ew")

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
            logo_label = tk.Label(header_frame, image=self.logo_photo, bg='#E0E0E0')
        except Exception:
            logo_label = tk.Label(header_frame, text="Sanofi", font=("Arial", 28, "bold"), bg='#E0E0E0')
        
        logo_label.pack(pady=20)

        # sap btn
        sap_actions_frame = ttk.Frame(main_frame)
        sap_actions_frame.grid(row=1, column=0, pady=(20, 0))

        tk.Button(sap_actions_frame, text="Extrair Ambos", command=lambda: self.setCommandCode(CommandsCodes.EXTRACT_BOTH.value), cursor="hand2", width=15, font=fonte_botoes_superiores, bg=cor_roxa_sanofi, fg='white', relief='flat', activebackground=cor_roxa_ativa).pack(side=tk.LEFT, padx=10)
        tk.Button(sap_actions_frame, text="Extrair MB52", command=lambda: self.setCommandCode(CommandsCodes.MB52.value), cursor="hand2", width=15, font=fonte_botoes_superiores, bg=cor_roxa_sanofi, fg='white', relief='flat', activebackground=cor_roxa_ativa).pack(side=tk.LEFT, padx=10)
        tk.Button(sap_actions_frame, text="Extrair BMBC", command=lambda: self.setCommandCode(CommandsCodes.BMBC.value), cursor="hand2", width=15, font=fonte_botoes_superiores, bg=cor_roxa_sanofi, fg='white', relief='flat', activebackground=cor_roxa_ativa).pack(side=tk.LEFT, padx=10)
        
        ttk.Separator(main_frame).grid(row=2, column=0, sticky="ew", padx=20, pady=20)

        files_frame = ttk.Frame(main_frame)
        files_frame.grid(row=3, column=0, sticky="ew", padx=20) 

        self.mb52_entry, self.mb52_status = self.create_file_input(files_frame, "MB52 (.xlsx)", "Caminho do arquivo...", 'mb52', "Arquivo MB52", ("Excel files", "*.xlsx"))
        self.bmbc_entry, self.bmbc_status = self.create_file_input(files_frame, "BMBC (.xlsx)", "Caminho do arquivo...", 'bmbc', "Arquivo BMBC", ("Excel files", "*.xlsx"))
        self.one_portfolio_entry, self.one_portfolio_status = self.create_file_input(files_frame, "One Portfolio (.xlsb)", "Caminho do arquivo...", 'portfolio', "One Portfolio", ("Excel Binary files", "*.xlsb"))
        self.inventory_entry, self.inventory_status = self.create_file_input(files_frame, "Inventário (.xlsx)", "Caminho do arquivo...", 'inventory', "Inventário", ("Excel files", "*.xlsx"))

        # rodapé
        footer_frame = ttk.Frame(main_frame)
        footer_frame.grid(row=4, column=0, sticky="ew", pady=(20, 10))
        self.status_label = ttk.Label(footer_frame, text="Carregue os arquivos para processar o inventário.")
        self.status_label.pack(pady=5)

        # btn area
        button_container = ttk.Frame(footer_frame)
        button_container.pack(pady=5)
        self.calcular_button = tk.Button(button_container, text="Calcular Inventário", command=self.confirm_and_calculate, width=20, font=fonte_botoes_inferiores, bg=cor_sucesso, fg='white', relief='flat', activebackground='#157347')
        self.limpar_button = tk.Button(button_container, text="Limpar Arquivos", command=self.limpar_tudo, width=20, font=fonte_botoes_inferiores, bg=cor_perigo, fg='white', relief='flat', activebackground='#BB2D3B')
    
    def create_file_input(self, parent, button_text, placeholder, keyword_check, file_dialog_title, filetype):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        
        entry = ttk.Entry(frame, style='Tall.TEntry')
        button = ttk.Button(frame, text=button_text, command=lambda e=entry: self._select_file_for_entry(e, keyword_check, file_dialog_title, filetype), style="Secondary.TButton", width=20)
        button.pack(side=tk.LEFT, padx=(0, 10))

        def on_focus_in(event):
            if entry.get() == placeholder:
                entry.delete(0, tk.END)
                entry.configure(foreground='black')
        def on_focus_out(event):
            if not entry.get():
                entry.insert(0, placeholder)
                entry.configure(foreground='gray')
        entry.insert(0, placeholder)
        entry.configure(foreground='gray')
        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        status_frame = ttk.Frame(frame, relief="sunken", borderwidth=1)
        status_frame.pack(side=tk.LEFT, padx=(5, 0))
        status_label = ttk.Label(status_frame, text="●", font=("Arial", 16, "bold"), foreground="red")
        status_label.pack()

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
        paths = [self.mb52_entry.get(), self.bmbc_entry.get(), self.one_portfolio_entry.get(), self.inventory_entry.get()]
        placeholder = "Caminho do arquivo..."
        
        if any(path for path in paths if path and path != placeholder):
            self.status_label.config(text="Arquivos carregados. Pronto para calcular ou limpar.")
            self.calcular_button.pack(side=tk.LEFT, padx=10)
            self.limpar_button.pack(side=tk.LEFT, padx=10)
        else:
            self.status_label.config(text="Carregue os arquivos para processar o inventário.")
            self.calcular_button.pack_forget()
            self.limpar_button.pack_forget()

    def _select_file_for_entry(self, entry_widget, keyword_check, file_dialog_title, filetype):
        file_path = filedialog.askopenfilename(title=f"Selecionar {file_dialog_title}",filetypes=[filetype])
        status_label_map = {self.mb52_entry: self.mb52_status, self.bmbc_entry: self.bmbc_status, self.one_portfolio_entry: self.one_portfolio_status, self.inventory_entry: self.inventory_status}
        status_label = status_label_map.get(entry_widget)
        placeholder = "Caminho do arquivo..."

        if file_path:
            filename = os.path.basename(file_path)
            if keyword_check.lower() in filename.lower():
                entry_widget.configure(foreground='black')
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
        placeholder = "Caminho do arquivo..."

        for entry, status in zip(entries, statuses):
            entry.delete(0, tk.END)
            entry.insert(0, placeholder)
            entry.configure(foreground='gray')
            status.config(foreground="red")
        
        self.update_status_label()
        messagebox.showinfo("Limpeza", "Todos os campos foram limpos!")

    def confirm_and_calculate(self):
        resposta = messagebox.askyesno(title="Confirmar Ação", message="Tem certeza que deseja carregar os dados?",default=messagebox.NO )
        if resposta:
            self.setCommandCode(CommandsCodes.CALCINVT.value)

    def setCommandCode(self, commandCode: int):
        self.__commandCode = commandCode
        self.master.destroy()

    def get_command_code(self):
        return self.__commandCode