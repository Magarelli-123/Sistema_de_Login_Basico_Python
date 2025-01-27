from tkinter import Tk, Frame, Label, Button
from constants import CORES

def iniciar_programa():
    """
    Interface principal do programa, aberta após o login.
    """
    # Configuração inicial da janela principal
    janela_principal = Tk()
    janela_principal.title("Programa Principal")
    janela_principal.geometry("800x600")
    janela_principal.configure(bg=CORES["branco_creme"])

    # Frame principal
    frame = Frame(janela_principal, bg=CORES["verde_principal"])
    frame.pack(fill="both", expand=True, padx=20, pady=20)

    # Rótulo de boas-vindas
    Label(
        frame, text="Bem-vindo ao Programa Principal!",
        bg=CORES["verde_principal"], fg=CORES["branco"],
        font=("Arial", 16, "bold")
    ).pack(pady=20)

    # Botão para sair do programa
    Button(
        frame, text="Sair", command=janela_principal.destroy,
        bg=CORES["laranja_suave"], fg=CORES["branco"],
        font=("Arial", 12, "bold")
    ).pack(pady=20)

    # Inicia o loop da interface
    janela_principal.mainloop()
