import os
import re
import tempfile
import json
from PySide6.QtGui import QPalette, QColor
from PySide6.QtCore import Qt

APP_NAME = "AutoMakeCardPSD" # Nome para a pasta de configuração
CONFIG_CAMADAS_FILENAME = "config_camadas_modelos.json"
TEXTO_NENHUM_MODELO = "(nenhum modelo disponível)"
CAMADA_TRATAMENTO = "tratamento"
CAMADA_NOME = "nome"
CAMADA_CONJUGE = "conjuge"
CAMADA_DATA = "data"

def get_app_config_dir():
    """Retorna o caminho para a pasta de configuração da aplicação."""
    # Para Windows: C:\Users\<Usuario>\AppData\Roaming\APP_NAME
    # Para Linux: /home/<Usuario>/.config/APP_NAME
    # Para macOS: /Users/<Usuario>/Library/Application Support/APP_NAME
    path = os.path.join(os.path.expanduser("~"),
                        ".config" if os.name != 'nt' else "AppData/Roaming",
                        APP_NAME)
    os.makedirs(path, exist_ok=True) # Cria a pasta se não existir
    return path

def get_path_config_camadas_json():
    """Retorna o caminho completo para o arquivo JSON de configuração das camadas."""
    return os.path.join(get_app_config_dir(), CONFIG_CAMADAS_FILENAME)


def carregar_configuracoes_camadas_modelos():
    """
    Carrega o dicionário de configurações do arquivo JSON.
    Garante retrocompatibilidade, convertendo o formato antigo para o novo se necessário.
    """
    filepath = get_path_config_camadas_json()
    if not os.path.exists(filepath):
        return {}

    try:
        with open(filepath, "r", encoding="utf-8") as f:
            data = json.load(f)

        # Lógica de migração para retrocompatibilidade
        for nome_modelo, config in data.items():
            # Se a configuração for uma lista (formato antigo)...
            if isinstance(config, list):
                # ...converte para o novo formato de dicionário.
                print(f"INFO: Convertendo configuração do modelo '{nome_modelo}' do formato antigo para o novo.")
                data[nome_modelo] = {
                    "dados_especificos": config,  # A lista antiga vira os "dados_especificos"
                    "regras_texto": {}  # Cria um dicionário vazio para as regras
                }

        return data

    except json.JSONDecodeError:
        print(f"Erro: Arquivo de configuração corrompido ou mal formatado: {filepath}")
        return {}
    except Exception as e:
        print(f"Erro ao carregar configurações de camadas: {e}")
        return {}

def salvar_configuracoes_camadas_modelos(configuracoes_dict):
    """Salva o dicionário de configurações dos modelos (Dados Específicos e Regras de Texto) no arquivo JSON."""
    filepath = get_path_config_camadas_json()
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            # O ensure_ascii=False é importante para salvar caracteres como 'ç' e acentos corretamente.
            json.dump(configuracoes_dict, f, indent=4, ensure_ascii=False)
        print(f"Configurações de modelos salvas em: {filepath}")
        return True
    except Exception as e:
        print(f"Erro ao salvar configurações de modelos: {e}")
        return False

def get_settings_file_path():
        # Caminho para arquivo JSON nos arquivos temporários do sistema
    return os.path.join(tempfile.gettempdir(), "cartao_app_settings.json")

def save_last_output_dir(output_dir):
        # Salva a última pasta de saída utilizada
    settings = {"last_output_dir": output_dir}
    with open(get_settings_file_path(), "w", encoding="utf-8") as f:
            json.dump(settings, f)

def load_last_output_dir():
        # Tenta carregar a última pasta de saída utilizada
    try:
        with open(get_settings_file_path(), "r", encoding="utf-8") as f:
                settings = json.load(f)
        return settings.get("last_output_dir", "")
    except Exception:
        return ""

# Função para aplicar um tema escuro simples
def set_dark_theme(app):
    dark_palette = QPalette()
    dark_palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.Base, QColor(42, 42, 42))
    dark_palette.setColor(QPalette.ColorRole.AlternateBase, QColor(66, 66, 66))
    dark_palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    dark_palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
    dark_palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
    dark_palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)

    dark_palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, QColor(127, 127, 127))
    dark_palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor(127, 127, 127))
    dark_palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.WindowText, QColor(127, 127, 127))

    app.setPalette(dark_palette)
    app.setStyleSheet(
        "QWidget { background-color: %s; color: white; }" % QColor(53, 53, 53).name() +
        "QToolTip { color: #ffffff; background-color: #2a82da; border: 1px solid white; }"
        "QPushButton { background-color: #0078D7; border: none; color: white; padding: 5px; text-align: center; min-height: 20px; border-radius: 3px;}"
        "QPushButton:hover { background-color: #106EBE }"
        "QPushButton:pressed { background-color: #005A9E }"
        "QComboBox { border: 1px solid #555; border-radius: 3px; padding: 1px 18px 1px 3px; min-width: 6em; background-color: #353535; }"
        "QComboBox:editable { background: #353535; }"
        "QComboBox QAbstractItemView { border: 1px solid #555; selection-background-color: #0078D7; background-color: #353535; color: white; }"
        "QTextEdit { background-color: #2A2A2A; color: #F0F0F0; border: 1px solid #555; }"
        "QLabel#PreviewLabel { background-color: #404040; color: black; border-radius: 8px; border: 1px solid #505050; }"
        "QTableWidget { background-color: #2A2A2A; color: #F0F0F0; border: 1px solid #555; gridline-color: #444; }"
        "QHeaderView::section { background-color: #3E3E3E; color: white; padding: 4px; border: 1px solid #555; }"
        "QTableWidget::item:selected { background-color: #0078D7; color: white; }"
    )

#__________________________________________________________________________________________

def obter_caminho_unico(caminho_proposto: str) -> str:
    """
    Verifica se um caminho de arquivo já existe. Se sim, adiciona um sufixo
    numérico (ex: '-2', '-3') até encontrar um caminho que não exista.
    """
    # Se o caminho original não existe, retorna ele imediatamente.
    if not os.path.exists(caminho_proposto):
        return caminho_proposto

    # Separa o caminho em diretório, nome base e extensão.
    # Ex: 'C:/saida/arquivo.png' -> 'C:/saida', 'arquivo', '.png'
    diretorio, nome_completo = os.path.split(caminho_proposto)
    nome_base, extensao = os.path.splitext(nome_completo)

    contador = 2
    while True:
        # Cria um novo nome de arquivo com o contador.
        # Ex: 'arquivo-2.png'
        novo_nome = f"{nome_base}-{contador}{extensao}"

        # Combina com o diretório para ter o caminho completo.
        novo_caminho = os.path.join(diretorio, novo_nome)

        # Se o novo caminho NÃO existe, encontramos um nome único.
        if not os.path.exists(novo_caminho):
            return novo_caminho # Retorna o caminho livre.

        # Se o caminho já existe, incrementa o contador e tenta de novo.
        contador += 1

#_________________________________________________________________________

def formatar_data_para_nome_arquivo(texto_data: str) -> str:
    """
    Converte uma string de data no formato 'DD de mês' para 'MM.DD'.
    Ex: '22 de abril' -> '04.22'.
    A busca pelo mês é case-insensitive.
    Retorna o texto original se a conversão falhar.
    """
    if not isinstance(texto_data, str) or not texto_data.strip():
        return texto_data

    # Dicionário de meses (chaves em minúsculo para busca case-insensitive)
    meses = {
        "janeiro": "01", "fevereiro": "02", "março": "03", "abril": "04",
        "maio": "05", "junho": "06", "julho": "07", "agosto": "08",
        "setembro": "09", "outubro": "10", "novembro": "11", "dezembro": "12"
    }

    texto_lower = texto_data.lower()
    dia_encontrado = None
    mes_encontrado = None

    # 1. Encontrar o dia (o primeiro número na string)
    match_dia = re.search(r'\d+', texto_lower)
    if match_dia:
        # zfill(2) garante que o dia tenha 2 dígitos (ex: '7' vira '07')
        dia_encontrado = match_dia.group(0).zfill(2)

    # 2. Encontrar o mês
    for nome_mes, num_mes in meses.items():
        if nome_mes in texto_lower:
            mes_encontrado = num_mes
            break

    # 3. Se ambos foram encontrados, retorna o formato novo. Senão, o original.
    if dia_encontrado and mes_encontrado:
        return f"{mes_encontrado}.{dia_encontrado}"
    else:
        return texto_data