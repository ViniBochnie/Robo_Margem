import tkinter as tk
from tkinter import ttk
import threading
from .button_handlers import (handle_margem, handle_cutof, handle_iqt,handle_dtg, handle_filial)
from .logger import Logger

from src import *

class Menu(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Menu")
        self.geometry("800x600")
        
        self.configure(bg="#f0f0f5")
        style = ttk.Style(self)
        style.configure("TButton", padding=(10,8), relief="flat", font=("Segoe", 12))
        style.map("TButton", background=[("active", "#d9d9e5")])
        style.configure("TFrame", background="#f0f0f5")
        style.configure("TLabel", background="#B8E342", font=("Segoe", 14))
        style.configure('Vertical.TScrollbar', gripcount=0, background='#d9d9e5', troughcolor='#f0f0f5', bordercolor='#f0f0f5')

        
        self.logger = Logger()
        self._create_widgets()
        self.Cutof = CutOf(log=self.logger.log)
        self.caminho = None
        
        try:
            self.leitor = Leitor(console=self.logger.log)
            self.setup = self.leitor.Loader()
            self.spu = Sharepoint(self.setup.azure)
            self.logger.log("Inicialização completa.\n")
        except Exception as e:
            self.logger.log(f"Erro ao inicializar: {e}\n")

    def _create_widgets(self):
        # Frame para os botões
        btn_frame = tk.Frame(self)
        # Expande horizontalmente para acomodar botões lado a lado
        btn_frame.pack(pady=15, fill=tk.X)

        # Botões, passando a função de log para o handler
        btn_margem_geral = tk.Button(
            btn_frame,
            text="Gerar Margem Total",
            width=15,
            command=lambda: threading.Thread(
                target=handle_margem,
                args=(self,)
            ).start()  # Executa em uma thread separada para não travar a interface
        )
        
        btn_margem_dtg = tk.Button(
            btn_frame,
            text="Gerar Margem DTG",
            width=15,
            command=lambda: threading.Thread(
                target=handle_dtg,
                args=(self,)
            ).start()  # Executa em uma thread separada para não travar a interface
        )
        
        btn_dtg_iqt = tk.Button(
            btn_frame,
            text="Gerar DTG e IQT",
            width=15,
            command=lambda: threading.Thread(
                target= lambda: (handle_dtg(self), handle_iqt(self))
            ).start()  # Executa em uma thread separada para não travar a interface
        )
        
        btn_filial_iqt = tk.Button(
            btn_frame,
            text="Gerar Filial e IQT",
            width=15,
            command=lambda: threading.Thread(
                target= lambda: (handle_filial(self), handle_iqt(self))
            ).start()  # Executa em uma thread separada para não travar a interface
        )
        
        btn_margem_filial = tk.Button(
            btn_frame,
            text="Gerar Margem Filial",
            width=15,
            command=lambda: threading.Thread(
                target=handle_filial,
                args=(self,)
            ).start()  # Executa em uma thread separada para não travar a interface
        )
        
        btn_cutof = tk.Button(
            btn_frame,
            text="Gerar Cutoff",
            width=10,
            command=lambda: threading.Thread(
                target=handle_cutof,
                args=(self,)
            ).start()  # Executa em uma thread separada para não travar a interface
        )
            
        btn_iqt = tk.Button(
            btn_frame,
            text="Gerar IQT",
            width=10,
            command=lambda: threading.Thread(
                target=handle_iqt,
                args=(self,)
            ).start()  # Executa em uma thread separada para não travar a interface
        )
        
        # Distribui botões horizontalmente
        btn_margem_geral.pack(side=tk.LEFT, padx=5, expand=True)
        btn_margem_dtg.pack(side=tk.LEFT, padx=5, expand=True)
        btn_dtg_iqt.pack(side=tk.LEFT, padx=5, expand=True)
        btn_margem_filial.pack(side=tk.LEFT, padx=5, expand=True)
        btn_filial_iqt.pack(side=tk.LEFT, padx=5, expand=True)
        btn_cutof.pack(side=tk.LEFT, padx=5, expand=True)
        btn_iqt.pack(side=tk.LEFT, padx=5, expand=True)
        
        # Caixa de logs com rolagem
        self.logger.bind_to(self)
