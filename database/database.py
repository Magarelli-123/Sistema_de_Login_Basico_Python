import sqlite3
from utils import carregar_config

# Carrega o caminho do banco de dados do config.json
config = carregar_config()
db_path = config["database_path"]

def conectar_banco() -> sqlite3.Connection:
    """
    Conecta ao banco de dados SQLite.
    Retorna a conexão ativa ou None se houver erro.
    """
    try:
        conn = sqlite3.connect(db_path)
        return conn
    except sqlite3.Error as e:
        print(f"Erro ao conectar ao banco de dados: {e}")
        return None

def criar_tabela_usuarios():
    """
    Cria a tabela 'usuarios' no banco de dados, se não existir.
    """
    conn = conectar_banco()
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS usuarios(
                Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                Nome TEXT NOT NULL,
                Email TEXT NOT NULL,
                Usuario TEXT NOT NULL,
                Senha TEXT NOT NULL
            );
            """)
            conn.commit()
        except sqlite3.Error as e:
            print(f"Erro ao criar tabela 'usuarios': {e}")
        finally:
            conn.close()

def criar_tabela_lancamentos():
    """
    Cria a tabela 'vendas' no banco de dados, se não existir.
    """
    conn = conectar_banco()
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("""
            CREATE TABLE vendas(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data DATE,
                canal TEXT,
                venda TEXT,
                cliente TEXT,
                telefone VARCHAR(11),
                hora TIME,
                produto TEXT,
                preco REAL,
                loja TEXT,
                atendente TEXT,
                status TEXT,
                cpf_cnpj VARCHAR(11),
                midia_social TEXT,
                retorno TEXT,
                cidade TEXT,
                carro TEXT,
                observacoes TEXT,
                pos_venda TEXT,
                motivo TEXT,
                dt_nascimento DATE,
                familia_produto TEXT,
                mensagem TEXT,
                enviar TEXT,
                cod TEXT,
                localizacao TEXT,
                ref TEXT,
                endereco TEXT,
                base_troca TEXT,
                prazo TEXT,
                tabela TEXT,
                quem_ira_receber TEXT,
                data_atualizacao DATE,
                hora_atualizacao TIME,
                forma_pgt TEXT
            , user_info VARCHAR(25));
            """)
            conn.commit()
        except sqlite3.Error as e:
            print(f"Erro ao criar tabela 'lancamentos_whatsapp': {e}")
        finally:
            conn.close()

def inicializar_banco():
    """
    Inicializa o banco de dados criando todas as tabelas necessárias.
    """
    criar_tabela_usuarios()
    criar_tabela_lancamentos()
