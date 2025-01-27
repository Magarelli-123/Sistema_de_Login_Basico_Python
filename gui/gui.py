# Importações necessárias
from tkinter import Tk, Frame, Label, Entry, Button, messagebox, IntVar, ttk  # Widgets do Tkinter
from PIL import Image, ImageTk  # Para manipulação de imagens
from constants import CORES  # Arquivo de constantes com cores do sistema
from database import conectar_banco  # Função para conectar ao banco de dados
from utils import carregar_config  # Função para carregar as configurações do config.json
from main_program import iniciar_programa  # Função para iniciar a interface principal após o login

# Carrega as configurações definidas no config.json
config = carregar_config()
image_path = config["image_path"]  # Caminho para as imagens definido no config.json

def iniciar_interface():
    """
    Cria e inicia a interface de login com base no layout do projeto.
    Essa interface inclui as opções de login e registro.
    """
    # Cria a janela principal da interface de login
    janela_loggin = Tk()
    janela_loggin.title("Teste Loggin 1 Acess Panel")  # Define o título da janela
    janela_loggin.geometry("600x300")  # Define o tamanho da janela (largura x altura)
    janela_loggin.configure(background=CORES["branco_creme"])  # Define a cor de fundo
    janela_loggin.resizable(False, False)  # Impede que a janela seja redimensionada
    janela_loggin.attributes("-alpha", 1)  # Define a opacidade da janela
    janela_loggin.iconbitmap(default=f"{image_path}/logo_loggin_icon.ico")  # Define o ícone da janela

    # Carrega a imagem do logo para exibição no lado esquerdo
    logo = Image.open(f"{image_path}/logo_loggin.png")
    logo = logo.resize((150, 150))  # Redimensiona a imagem para um tamanho fixo
    logo = ImageTk.PhotoImage(logo)  # Converte a imagem para um formato compatível com o Tkinter

    # Cria os frames laterais para dividir a interface
    frame_esquerda = Frame(janela_loggin, width=200, height=300, bg=CORES["verde_escuro"], relief="raise")
    frame_esquerda.pack(side="left")  # Posiciona o frame na lateral esquerda

    frame_direita = Frame(janela_loggin, width=395, height=300, bg=CORES["verde_escuro"], relief="raise")
    frame_direita.pack(side="right")  # Posiciona o frame na lateral direita

    # Adiciona o logo no frame esquerdo
    logo_label = Label(frame_esquerda, image=logo, bg=CORES["verde_escuro"])
    logo_label.place(x=20, y=90)  # Posiciona a imagem dentro do frame

    # Adiciona o campo para o usuário no frame direito
    Label(
        frame_direita,
        text="Usuário:",  # Texto do rótulo
        font=("Century Gothic", 20, "bold"),  # Fonte e estilo
        bg=CORES["verde_escuro"],
        fg=CORES["branco_creme"]
    ).place(x=10, y=90)  # Posiciona o rótulo no frame direito

    user_entry = ttk.Entry(frame_direita, width=30)  # Campo de entrada para o usuário
    user_entry.place(x=135, y=102)

    # Adiciona o campo para a senha no frame direito
    Label(
        frame_direita,
        text="Senha:",  # Texto do rótulo
        font=("Century Gothic", 20, "bold"),  # Fonte e estilo
        bg=CORES["verde_escuro"],
        fg=CORES["branco_creme"]
    ).place(x=10, y=143)  # Posiciona o rótulo no frame direito

    senha_entry = ttk.Entry(frame_direita, width=30, show="*")  # Campo de entrada para a senha
    senha_entry.place(x=135, y=155)

    # Função para validar o login do usuário
    def loggin():
        # Conecta ao banco de dados
        conn = conectar_banco()
        if conn:
            try:
                cursor = conn.cursor()
                # Verifica se o usuário e a senha existem no banco
                cursor.execute(
                    "SELECT Usuario, Senha FROM usuarios WHERE Usuario=? AND Senha=?",
                    (user_entry.get(), senha_entry.get())
                )
                verificar_login = cursor.fetchone()

                if verificar_login:
                    # Exibe uma mensagem de sucesso e inicia a interface principal
                    messagebox.showinfo("Login", "Login realizado com sucesso!")
                    janela_loggin.destroy()  # Fecha a janela de login
                    iniciar_programa()  # Inicia a interface principal do programa
                else:
                    # Exibe uma mensagem de erro caso o login falhe
                    messagebox.showerror("Erro", "Usuário ou senha incorretos.")
            except Exception as e:
                # Exibe uma mensagem de erro caso ocorra algum problema
                messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
            finally:
                conn.close()  # Fecha a conexão com o banco
        else:
            # Exibe uma mensagem caso não seja possível conectar ao banco de dados
            messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados.")

    # Função para abrir o formulário de registro
    def register():
        # Remove os botões de login e registro
        login_button.place_forget()
        registrar_button.place_forget()

        # Adiciona os campos para o registro
        nome_label = Label(
            frame_direita,
            text="Nome:",  # Texto do rótulo
            font=("Century Gothic", 20, "bold"),
            bg=CORES["verde_escuro"],
            fg=CORES["branco_creme"]
        )
        nome_label.place(x=10, y=5)
        nome_entry = ttk.Entry(frame_direita, width=30)  # Campo de entrada para o nome
        nome_entry.place(x=135, y=17)

        email_label = Label(
            frame_direita,
            text="Email:",  # Texto do rótulo
            font=("Century Gothic", 20, "bold"),
            bg=CORES["verde_escuro"],
            fg=CORES["branco_creme"]
        )
        email_label.place(x=10, y=45)
        email_entry = ttk.Entry(frame_direita, width=30)  # Campo de entrada para o e-mail
        email_entry.place(x=135, y=57)

        # Função interna para salvar os dados do registro no banco
        def registrar_no_db():
            nome = nome_entry.get()
            email = email_entry.get()
            usuario = user_entry.get()
            senha = senha_entry.get()

            # Verifica se todos os campos estão preenchidos
            if not nome or not email or not usuario or not senha:
                messagebox.showerror("Erro", "Preencha todos os campos.")
                return

            conn = conectar_banco()
            if conn:
                try:
                    cursor = conn.cursor()
                    # Insere os dados no banco
                    cursor.execute(
                        "INSERT INTO usuarios (nome, email, usuario, senha) VALUES (?, ?, ?, ?)",
                        (nome, email, usuario, senha)
                    )
                    conn.commit()
                    messagebox.showinfo("Registro", "Conta criada com sucesso!")
                except Exception as e:
                    messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
                finally:
                    conn.close()

        # Botão para confirmar o registro
        registrar_func = ttk.Button(frame_direita, text="Registrar", width=30, command=registrar_no_db)
        registrar_func.place(x=135, y=225)

        # Função para voltar ao login
        def voltar_loggin():
            nome_label.place_forget()
            nome_entry.place_forget()
            email_label.place_forget()
            email_entry.place_forget()
            registrar_func.place_forget()
            voltar_func_reg.place_forget()

            # Reexibe os botões de login e registro
            login_button.place(x=135, y=225)
            registrar_button.place(x=170, y=260)

        # Botão de voltar ao login
        voltar_func_reg = ttk.Button(frame_direita, text="Voltar", width=20, command=voltar_loggin)
        voltar_func_reg.place(x=170, y=260)

    # Botão de login
    login_button = ttk.Button(frame_direita, text="Login", width=30, command=loggin)
    login_button.place(x=135, y=225)

    # Botão para abrir o registro
    registrar_button = ttk.Button(frame_direita, text="Registrar", width=20, command=register)
    registrar_button.place(x=170, y=260)

    # Inicia o loop da interface gráfica
    janela_loggin.mainloop()
