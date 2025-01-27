# Sistema de Login Básico com Python

## Sobre o Projeto
Esse projeto é um sistema básico de login em python que utiliza interface gráfica e conexão em banco de dados SQL onde, de forma fácil e intuitiva, se pode utilizar para servir de sistema de login para qualquer sistema interno de uma empresa, sem a necessidade de custos adicionais ou bancos de dados em nuvem, funcionando muito bem em servidores de rede interna!

### Funcionalidades Principais
- **Gerenciamento de Usuários**: Login, registro e gerenciamento de perfis de acesso.
- **Interface Gráfica**: Desenvolvida com Tkinter, permitindo uma experiência intuitiva.
- **Integração com Banco de Dados**: Utiliza SQLite para armazenar e gerenciar os dados.

## Tecnologias Utilizadas
- **Linguagem**: Python 3.8+
- **Bibliotecas**:
  - `Tkinter`: Interface gráfica.
  - `sqlite3`: Banco de dados local.
  - `Pillow`: Manipulação de imagens.
 

## Estrutura do Projeto
```
/cadapp/
  ├─ config.json                # Configurações do projeto.
  ├─ database/
  │   └─ database.py          # Conexão e criação de tabelas no banco de dados.
  ├─ gui/
  │   └─ gui.py               # Interface de login e registro.
  ├─ utils/
  │   └─ constants.py        # Constantes como cores e listas predefinidas.
  │   └─ functions.py        # Funções de suporte.
  │   └─ utils.py            # Funções utilitárias (e.g., validação de CPF/CNPJ).
  ├─ tests/
  │   └─ loggin_index.py # Testes de funcionalidades principais.
  ├─ main/
  │   └─ apenas_main.py      # Módulo principal do sistema.
  │   └─ main_program.py     # Controle da interface principal.
  │   └─ main.py             # Inicialização do sistema.
  ├─ README.md                  # Documentação do projeto.
```

## Como Executar o Projeto

### Requisitos
- Python 3.8 ou superior.
- Pip (gerenciador de pacotes do Python).
- As bibliotecas descritas no arquivo `requirements.txt`.

### Instalação
1. Clone o repositório:
   ```bash
   git clone git clone https://github.com/Magarelli-123/sistema_login_basico.git
   ```

2. Navegue até o diretório do projeto:
   ```bash
   cd sistema_login_basico
   ```

3. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

4. Execute o sistema:
   ```bash
   python main.py
   ```

## Configuração
O arquivo `config.json` é usado para definir caminhos e configurações do sistema. Certifique-se de que os caminhos estejam corretos antes de executar o projeto.

### Exemplo de `config.json`:
```json
{
    "database_path": "X:/cadapp/atendimento_wpp_teste1.db",
    "image_path": "X:/cadapp/.venv/cadapp_v.01/main_teste1/imgs"
}
```

## Contribuições
Contribuições são bem-vindas! Sinta-se à vontade para abrir uma *issue* ou enviar um *pull request*.


## Contato
Para mais informações ou dúvidas, entre em contato:
- **E-mail**: magarelli@gmail.com
- **GitHub**: [Seu GitHub](https://github.com/Magarelli-123)

