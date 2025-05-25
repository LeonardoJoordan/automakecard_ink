import os
import tempfile
import json
from PySide6.QtGui import QPalette, QColor
from PySide6.QtCore import Qt

APP_NAME = "AutoMakeCardPSD" # Nome para a pasta de configuração
CONFIG_CAMADAS_FILENAME = "config_camadas_modelos.json"

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
    """Carrega o dicionário de configurações de camadas do arquivo JSON."""
    filepath = get_path_config_camadas_json()
    if not os.path.exists(filepath):
        return {} # Retorna um dicionário vazio se o arquivo não existe
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data
    except json.JSONDecodeError:
        print(f"Erro: Arquivo de configuração de camadas corrompido ou mal formatado: {filepath}")
        return {} # Retorna vazio em caso de erro de leitura/formato
    except Exception as e:
        print(f"Erro ao carregar configurações de camadas: {e}")
        return {}

def salvar_configuracoes_camadas_modelos(configuracoes_dict):
    """Salva o dicionário de configurações de camadas no arquivo JSON."""
    filepath = get_path_config_camadas_json()
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(configuracoes_dict, f, indent=4, ensure_ascii=False)
        print(f"Configurações de camadas salvas em: {filepath}")
        return True
    except Exception as e:
        print(f"Erro ao salvar configurações de camadas: {e}")
        return False

def data_por_extenso(data_str):
    """Recebe uma string no formato DD/MM/AAAA ou DD/MM/AA e retorna 'Ponta Grossa, DD de <mês> de AAAA.'"""
    meses = [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]
    try:
        partes = data_str.strip().split('/')
        if len(partes) != 3:
            return ""
        dia, mes, ano = partes
        if len(ano) == 2:  # Se vier '25', transforma em '2025'
            ano = "20" + ano
        dia = str(int(dia))  # Remove zero à esquerda
        mes_extenso = meses[int(mes) - 1]
        return f"Ponta Grossa, {dia} de {mes_extenso} de {ano}."
    except Exception:
        return ""

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