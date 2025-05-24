import os
import tempfile
import json

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