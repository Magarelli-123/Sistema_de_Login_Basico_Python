import json
import re
import os

def carregar_config() -> dict:
    """
    Carrega as configurações do arquivo config.json.
    Retorna um dicionário com as configurações ajustadas.
    """
    try:
        with open("config.json", "r") as arquivo:
            config = json.load(arquivo)

            # Converte os caminhos para o formato correto do sistema operacional
            config["database_path"] = os.path.normpath(config["database_path"])
            config["image_path"] = os.path.normpath(config["image_path"])
            
            return config
    except FileNotFoundError:
        raise FileNotFoundError("Arquivo config.json não encontrado. Verifique o diretório do projeto.")
    except json.JSONDecodeError:
        raise ValueError("Erro ao decodificar o arquivo config.json. Verifique o formato JSON.")

def validar_email(email: str) -> bool:
    """
    Valida o formato de um e-mail.
    Retorna True se o e-mail for válido, False caso contrário.
    """
    padrao = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    return re.match(padrao, email) is not None

def validar_cpf_cnpj(cpf_cnpj: str) -> bool:
    """
    Valida o formato de um CPF ou CNPJ (apenas o tamanho e se é numérico).
    Retorna True se válido, False caso contrário.
    """
    return cpf_cnpj.isdigit() and len(cpf_cnpj) in [11, 14]
