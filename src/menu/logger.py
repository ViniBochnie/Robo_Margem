import tkinter as tk
from tkinter import scrolledtext

class Logger:
    def __init__(self):
        self.log_box = None

    def bind_to(self, parent: tk.Tk):
        # cria e posiciona a ScrolledText
        self.log_box = scrolledtext.ScrolledText(
            parent,
            wrap=tk.WORD,
            height=10,
            state='disabled'
        )
        self.log_box.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def log(self, message: str):
        if not self.log_box:
            return
        self.log_box.configure(state='normal')
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)
        self.log_box.configure(state='disabled')