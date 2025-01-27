from database import inicializar_banco
from gui import iniciar_interface

if __name__ == "__main__":
    # Inicializa o banco de dados criando tabelas, se necessário
    inicializar_banco()

    # Inicia a interface gráfica do sistema
    iniciar_interface()
