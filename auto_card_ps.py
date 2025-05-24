import sys
import os
import shutil
import utils
import ps_utils

import tempfile
import json

import win32com.client

import re

import subprocess
from PIL import Image, ImageDraw, ImageFont  # PIL para manipulação de imagem

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTextEdit, QLabel, QComboBox, QPushButton, QFileDialog, QMessageBox,
    QSizePolicy, QTableWidget, QTableWidgetItem, QAbstractItemView,
    QHeaderView
)
from PySide6.QtGui import (
    QFont, QPixmap, QImage, QPalette, QColor, QKeySequence, QIcon, QGuiApplication
)
from PySide6.QtWidgets import QApplication

from PySide6.QtCore import QTranslator, QLocale, QLibraryInfo

from PySide6.QtCore import Qt, QSize, QMimeData, QEvent

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

#def gerar_cartao_photoshop(psd_path, output_path, campos):
#    '''
#    psd_path: caminho do modelo PSD
#    output_path: caminho do PNG de saída
#    campos: dicionário com {'tratamento': str, 'nome': str, 'conjuge': str, 'data': str}
#    '''
#    psApp = win32com.client.Dispatch("Photoshop.Application")
#    psApp.Visible = False  # Coloque True se quiser acompanhar
#   doc = psApp.Open(psd_path)
#
#    for camada_nome, texto in campos.items():
#        try:
#            doc.ArtLayers[camada_nome].TextItem.Contents = texto
#        except Exception as e:
#            print(f"Erro ao alterar camada '{camada_nome}': {e}")
#
#    # Exporta para PNG
#    options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
#    options.Format = 13  # PNG
#    options.PNG8 = False  # PNG-24
#    doc.Export(ExportIn=output_path, ExportAs=2, Options=options)
#
#    doc.Close(2)  # Fecha sem salvar alterações no PSD

class CustomTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.EditKeyPressed | QAbstractItemView.EditTrigger.AnyKeyPressed)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

    def keyPressEvent(self, event: QEvent):
        if event.matches(QKeySequence.StandardKey.Paste):
            self.custom_paste()
        elif event.key() == Qt.Key.Key_Delete or event.key() == Qt.Key.Key_Backspace:
            self.delete_selected_cells_content()
        else:
            super().keyPressEvent(event)

    def delete_selected_cells_content(self):
        for item in self.selectedItems():
            item.setText("")

    def custom_paste(self):
        clipboard = QGuiApplication.clipboard()
        mime_data = clipboard.mimeData()

        if mime_data.hasText():
            text = mime_data.text()
            rows_data = text.strip('\n').split('\n')
            if not rows_data: return

            table_data = []
            for row_str in rows_data:
                table_data.append(row_str.split('\t'))

            start_row = self.currentRow() if self.currentRow() != -1 else 0
            start_col = self.currentColumn() if self.currentColumn() != -1 else 0
            if not self.selectedIndexes():
                start_row, start_col = 0, 0

            num_pasted_rows = len(table_data)
            max_pasted_cols_in_data = 0
            if table_data:
                max_pasted_cols_in_data = max(len(row_content) for row_content in table_data) if table_data[0] else 0

            required_rows = start_row + num_pasted_rows
            if required_rows > self.rowCount():
                self.setRowCount(required_rows)

            for r_idx, row_content in enumerate(table_data):
                current_table_row = start_row + r_idx
                for c_idx, cell_value in enumerate(row_content):
                    current_table_col = start_col + c_idx
                    if current_table_col < self.columnCount():
                        item = self.item(current_table_row, current_table_col)
                        if not item:
                            item = QTableWidgetItem(cell_value)
                            self.setItem(current_table_row, current_table_col, item)
                        else:
                            item.setText(cell_value)
        else:
            paste_event = QEvent(QEvent.Type.KeyPress)
            super().keyPressEvent(paste_event)


class CartaoApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerenciador de Modelos PSD PySide6")
        self.setFixedSize(1000, 550)
        self.log_textbox = QTextEdit()
        self.log_textbox.setReadOnly(True)
        self.log_textbox.append("Log do programa...\n")

        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        # --- Coluna esquerda - Tabela ---
        tabela_container = QWidget()
        tabela_layout = QVBoxLayout(tabela_container)

        self.table_headers = ["tratamento", "nome", "conjuge", "data"]
        self.data_table = CustomTableWidget(self)
        self.data_table.setColumnCount(len(self.table_headers))
        self.data_table.setHorizontalHeaderLabels(self.table_headers)
        self.data_table.setRowCount(15)

        header_height = self.data_table.horizontalHeader().height()
        estimated_row_height = 28
        total_rows_height = 14 * estimated_row_height
        table_target_height = header_height + total_rows_height + self.data_table.horizontalScrollBar().height() + 5
        self.data_table.setFixedHeight(table_target_height)

        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        tabela_layout.addWidget(self.data_table)

        self.output_dir = utils.load_last_output_dir()
        if not self.output_dir or not os.path.exists(self.output_dir):
            self.output_dir = os.path.abspath("cartoes_gerados")
            if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)

        # Botões de gerenciamento da tabela
        table_buttons_layout = QHBoxLayout()
        btn_add_row = QPushButton("Adicionar Linha")
        btn_add_row.clicked.connect(self.add_table_row)
        btn_remove_row = QPushButton("Remover Linha(s) Selecionada(s)")
        btn_remove_row.clicked.connect(self.remove_selected_table_rows)
        btn_clear_table = QPushButton("Limpar Tabela")
        btn_clear_table.clicked.connect(self.clear_table)

        table_buttons_layout.addWidget(btn_add_row)
        table_buttons_layout.addWidget(btn_remove_row)
        table_buttons_layout.addWidget(btn_clear_table)
        tabela_layout.addLayout(table_buttons_layout)

        main_layout.addWidget(tabela_container, 2)

        # --- Coluna direita ---
        direita_frame = QWidget()
        direita_layout = QVBoxLayout(direita_frame)
        direita_frame.setFixedWidth(380)

        #Pre visualização miniatura do arquivo selecionado
        self.preview_label = QLabel("Aguardando seleção de modelo")
        self.preview_label.setObjectName("PreviewLabel")
        self.preview_label.setFixedSize(300, 200)
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setScaledContents(False)
        direita_layout.addWidget(self.preview_label, 0, Qt.AlignmentFlag.AlignHCenter)

        # Botão Gerar Cartões logo abaixo da preview
        self.btn_gerar_cartoes = QPushButton("Gerar Cartões")
        self.btn_gerar_cartoes.setStyleSheet("font-weight: bold; background-color: #1db954;")
        self.btn_gerar_cartoes.clicked.connect(self.gerar_cartoes)
        direita_layout.addWidget(self.btn_gerar_cartoes)

        #selecionar modelo
        self.modelo_combobox = QComboBox()
        self.modelo_combobox.addItem("(nenhum modelo disponível)")
        self.modelo_combobox.currentTextChanged.connect(self.atualizar_preview_modelo)
        direita_layout.addWidget(self.modelo_combobox)


        botoes_modelo_widget = QWidget()
        botoes_modelo_layout = QHBoxLayout(botoes_modelo_widget)
        botoes_modelo_layout.setContentsMargins(0, 0, 0, 0)
        botoes_modelo_layout.setSpacing(5)

        self.btn_adicionar_modelo = QPushButton("Adicionar Modelo")
        self.btn_adicionar_modelo.clicked.connect(self.adicionar_modelo)
        botoes_modelo_layout.addWidget(self.btn_adicionar_modelo)

        self.btn_modificar_modelo = QPushButton("Modificar Modelo")
        self.btn_modificar_modelo.clicked.connect(self.modificar_modelo)
        botoes_modelo_layout.addWidget(self.btn_modificar_modelo)

        self.btn_excluir_modelo = QPushButton("Excluir Modelo")
        self.btn_excluir_modelo.clicked.connect(self.excluir_modelo)
        botoes_modelo_layout.addWidget(self.btn_excluir_modelo)
        direita_layout.addWidget(botoes_modelo_widget)

        self.atualizar_modelos_combobox()
        self._current_pixmap = None

        # caixa de log
        self.log_textbox = QTextEdit()
        self.log_textbox.setReadOnly(True)
        self.log_textbox.append("Log do programa...\n")
        direita_layout.addWidget(self.log_textbox)

        # Layout para mostrar e selecionar pasta de saída
        saida_dir_layout = QHBoxLayout()
        self.saida_dir_label = QLabel()
        self.saida_dir_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.saida_dir_label.setText(self.output_dir)
        self.saida_dir_label.setStyleSheet("color: #A0FFA0; background: #222; padding: 2px 4px; border-radius: 3px;")
        self.btn_selecionar_saida = QPushButton("Selecionar Pasta de Saída")
        self.btn_selecionar_saida.clicked.connect(self.selecionar_pasta_saida)
        saida_dir_layout.addWidget(self.saida_dir_label)
        saida_dir_layout.addWidget(self.btn_selecionar_saida)
        direita_layout.addLayout(saida_dir_layout)

        direita_layout.addStretch(1)
        main_layout.addWidget(direita_frame, 1)

    def gerar_cartoes(self):
        self.log_message("Iniciando geração de cartões...")

        linhas = self.data_table.rowCount()
        colunas = self.data_table.columnCount()
        headers = [self.data_table.horizontalHeaderItem(i).text() for i in range(colunas)]

        total_validos = 0

        for row in range(linhas):
            linha_dados = {}
            linha_vazia = True

            for col in range(colunas):
                item = self.data_table.item(row, col)
                valor = item.text().strip() if item else ""
                if valor:
                    linha_vazia = False
                linha_dados[headers[col]] = valor

            if linha_vazia:
                continue  # pula linhas totalmente vazias

            # Transforma a data da linha, se existir
            data_extenso = utils.data_por_extenso(linha_dados.get('data', ''))
            #self.log_message(f"Cartão {total_validos + 1}: {linha_dados}")
            self.log_message(f"Data por extenso: {data_extenso}")

            total_validos += 1

            psd_modelo_relativo = os.path.join("modelos", self.modelo_combobox.currentText())
            psd_path = os.path.abspath(psd_modelo_relativo)  # Converte para caminho absoluto
            # Gera nome: MM.DD - Nome.png
            data_bruta = linha_dados['data']
            if data_bruta and len(data_bruta) >= 5:
                nome_png = f"{data_bruta[3:5]}.{data_bruta[:2]} - {'nome'}.png"

            else:
                nome_png = f"{linha_dados['nome']}.png"
            output_path = os.path.join(self.output_dir, nome_png)

            campos = {
                'tratamento': linha_dados['tratamento'],
                'nome': linha_dados['nome'],
                'conjuge': linha_dados['conjuge'],
                'data': utils.data_por_extenso(linha_dados['data']) if linha_dados['data'] else "",
            }

            if data_bruta and len(data_bruta) >= 5:
                nome_png = f"{data_bruta[3:5]}.{data_bruta[:2]} - {linha_dados['nome']}.png"
            else:
                nome_png = f"{linha_dados['nome']}.png"
            output_path = os.path.join(self.output_dir, nome_png)

            try:
                ps_utils.gerar_cartao_photoshop(psd_path, output_path, campos)
                self.log_message(f"Cartão salvo: {output_path}")
            except Exception as e:
                self.log_message(f"Erro ao gerar cartão '{nome_png}': {e}")


        self.log_message(
            f"Geração finalizada. {total_validos} cartões preparados")

    #def data_por_extenso(self, data_str):
    #    """Recebe uma string no formato DD/MM/AAAA ou DD/MM/AA e retorna 'Ponta Grossa, DD de <mês> de AAAA.'"""
    #    meses = [
    #        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    #        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    #    ]
    #    try:
    #        partes = data_str.strip().split('/')
    #        if len(partes) != 3:
    #            return ""
    #        dia, mes, ano = partes
    #        if len(ano) == 2:  # Se vier '25', transforma em '2025'
    #            ano = "20" + ano
    #        dia = str(int(dia))  # Remove zero à esquerda
    #        mes_extenso = meses[int(mes) - 1]
    #        return f"Ponta Grossa, {dia} de {mes_extenso} de {ano}."
    #    except Exception:
    #        return ""

    def selecionar_pasta_saida(self):
        pasta = QFileDialog.getExistingDirectory(self, "Escolha a pasta de saída", self.output_dir)
        if pasta:
            self.output_dir = pasta
            self.saida_dir_label.setText(self.output_dir)
            utils.save_last_output_dir(self.output_dir)
            self.log_message(f"Pasta de saída definida para: {self.output_dir}")

#    def get_settings_file_path(self):
#        # Caminho para arquivo JSON nos arquivos temporários do sistema
#        return os.path.join(tempfile.gettempdir(), "cartao_app_settings.json")
#
#    def save_last_output_dir(self, output_dir):
#        # Salva a última pasta de saída utilizada
#        settings = {"last_output_dir": output_dir}
#        with open(utils.get_settings_file_path(), "w", encoding="utf-8") as f:
#            json.dump(settings, f)
#
#    def load_last_output_dir(self):
#        # Tenta carregar a última pasta de saída utilizada
#        try:
#            with open(utils.get_settings_file_path(), "r", encoding="utf-8") as f:
#                settings = json.load(f)
#                return settings.get("last_output_dir", "")
#        except Exception:
#            return ""

    def add_table_row(self):
        current_row_count = self.data_table.rowCount()
        self.data_table.insertRow(current_row_count)

    def remove_selected_table_rows(self):
        selected_rows = sorted(list(set(index.row() for index in self.data_table.selectedIndexes())), reverse=True)
        if not selected_rows:
            QMessageBox.information(self, "Remover Linhas", "Nenhuma linha selecionada para remover.")
            return

        confirm = QMessageBox.question(self, "Confirmar Remoção",
                                       f"Tem certeza que deseja remover {len(selected_rows)} linha(s)?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            for row_index in selected_rows:
                self.data_table.removeRow(row_index)
            self.log_message(f"{len(selected_rows)} linha(s) removida(s) da tabela.")

    def clear_table(self):
        confirm = QMessageBox.question(self, "Limpar Tabela",
                                       "Tem certeza que deseja limpar todos os dados da tabela?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            self.data_table.setRowCount(0)
            self.data_table.setRowCount(20)
            self.log_message("Tabela limpa.")

    def garantir_pasta_modelos(self):
        if not os.path.exists("modelos"):
            os.makedirs("modelos")
            self.log_message("Pasta 'modelos' criada.")

    def log_message(self, message):
        self.log_textbox.append(message)
        self.log_textbox.ensureCursorVisible()

    def atualizar_modelos_combobox(self):
        self.garantir_pasta_modelos()
        try:
            # Lista arquivos .psd
            arquivos = [f for f in os.listdir("modelos") if f.lower().endswith(".psd")]
        except FileNotFoundError:
            self.log_message("Erro: Pasta 'modelos' não encontrada ao listar arquivos.")
            arquivos = []

        self.modelo_combobox.blockSignals(True)
        current_selection = self.modelo_combobox.currentText()
        self.modelo_combobox.clear()

        if arquivos:
            self.modelo_combobox.addItems(arquivos)
            if current_selection in arquivos:
                self.modelo_combobox.setCurrentText(current_selection)
            elif arquivos:
                self.modelo_combobox.setCurrentIndex(0)
            else:
                self.modelo_combobox.addItem("(nenhum modelo disponível)")
        else:
            self.modelo_combobox.addItem("(nenhum modelo disponível)")

        self.modelo_combobox.blockSignals(False)
        self.atualizar_preview_modelo(self.modelo_combobox.currentText())

    def adicionar_modelo(self):
        self.garantir_pasta_modelos()
        arquivo, _ = QFileDialog.getOpenFileName(
            self,
            "Selecione um modelo (.psd)",
            "modelos",
            "Arquivos do Photoshop (*.psd);;Todos os Arquivos (*.*)"
        )
        if arquivo:
            nome_arquivo = os.path.basename(arquivo)
            destino = os.path.join("modelos", nome_arquivo)
            if os.path.exists(destino):
                resp = QMessageBox.question(self, "Substituir modelo",
                                            f"O modelo '{nome_arquivo}' já existe. Deseja substituir?",
                                            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if resp == QMessageBox.StandardButton.No:
                    return
            try:
                shutil.copy2(arquivo, destino)
                self.log_message(f"Modelo '{nome_arquivo}' adicionado/atualizado.")
                self.atualizar_modelos_combobox()
            except Exception as e:
                QMessageBox.critical(self, "Erro ao adicionar", f"Não foi possível copiar o arquivo: {e}")
                self.log_message(f"Erro ao copiar '{nome_arquivo}': {e}")

    def modificar_modelo(self):
        modelo_selecionado = self.modelo_combobox.currentText()
        if not modelo_selecionado or modelo_selecionado == "(nenhum modelo disponível)":
            QMessageBox.information(self, "Modificar Modelo", "Nenhum modelo selecionado para modificar.")
            return

        novo_arquivo, _ = QFileDialog.getOpenFileName(
            self,
            "Selecione o novo arquivo do modelo (.psd)",
            "modelos",
            "Arquivos do Photoshop (*.psd);;Todos os Arquivos (*.*)"
        )
        if novo_arquivo:
            destino = os.path.join("modelos", modelo_selecionado)
            try:
                shutil.copy2(novo_arquivo, destino)
                self.log_message(f"Modelo '{modelo_selecionado}' foi modificado.")
                # Força a regeneração do preview após a modificação
                preview_path = os.path.join("modelos", f"{os.path.splitext(modelo_selecionado)[0]}_preview.png")
                if os.path.exists(preview_path):
                    os.remove(preview_path) # Remove o preview antigo para forçar a criação de um novo
                self.atualizar_preview_modelo(modelo_selecionado)
            except Exception as e:
                QMessageBox.critical(self, "Erro ao modificar", f"Não foi possível substituir o arquivo: {e}")
                self.log_message(f"Erro ao modificar '{modelo_selecionado}': {e}")

    def excluir_modelo(self):
        modelo_selecionado = self.modelo_combobox.currentText()
        if not modelo_selecionado or modelo_selecionado == "(nenhum modelo disponível)":
            QMessageBox.information(self, "Excluir Modelo", "Nenhum modelo selecionado para excluir.")
            return

        resp = QMessageBox.question(self, "Excluir Modelo",
                                    f"Tem certeza que deseja excluir o modelo '{modelo_selecionado}'?",
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if resp == QMessageBox.StandardButton.Yes:
            caminho_psd = os.path.join("modelos", modelo_selecionado)
            preview_path = os.path.join("modelos", f"{os.path.splitext(modelo_selecionado)[0]}_preview.png")
            try:
                if os.path.exists(caminho_psd):
                    os.remove(caminho_psd)
                if os.path.exists(preview_path):
                    os.remove(preview_path)

                self.log_message(f"Modelo '{modelo_selecionado}' foi excluído.")
                self.atualizar_modelos_combobox()
            except Exception as e:
                QMessageBox.critical(self, "Erro ao excluir", f"Erro: {e}")
                self.log_message(f"Erro ao excluir '{modelo_selecionado}': {e}")

    def atualizar_preview_modelo(self, modelo_selecionado=None):
        if modelo_selecionado is None:
            modelo_selecionado = self.modelo_combobox.currentText()

        if not modelo_selecionado or modelo_selecionado == "(nenhum modelo disponível)":
            self.preview_label.clear()
            self.preview_label.setText("Aguardando seleção de modelo")
            self.preview_label.setStyleSheet(
                "background-color: #404040; color: white; border-radius: 8px; border: 1px solid #505050;")
            self._current_pixmap = None
            return

        caminho_psd = os.path.join("modelos", modelo_selecionado)
        # O nome do arquivo de preview será o nome do PSD sem a extensão, mais .png
        caminho_preview = os.path.join("modelos", f"{os.path.splitext(modelo_selecionado)[0]}_preview.png")

        if not os.path.exists(caminho_psd):
            self.preview_label.clear()
            self.preview_label.setText(f"Arquivo PSD '{modelo_selecionado}' não encontrado.")
            self.preview_label.setStyleSheet("background-color: orange; color: black; border-radius: 8px;")
            self.log_message(f"Preview erro: '{caminho_psd}' não existe.")
            self._current_pixmap = None
            return

        # Verifica se o preview existe e se está atualizado (comparando data de modificação)
        if os.path.exists(caminho_preview) and os.path.getmtime(caminho_preview) > os.path.getmtime(caminho_psd):
            try:
                #self.log_message(f"Carregando pré-visualização existente para '{modelo_selecionado}'.")
                pil_img = Image.open(caminho_preview)
                if pil_img.mode != 'RGBA':
                    pil_img = pil_img.convert('RGBA')
                qimage = QImage(pil_img.tobytes("raw", "RGBA"), pil_img.width, pil_img.height,
                                QImage.Format.Format_RGBA8888)
                self._current_pixmap = QPixmap.fromImage(qimage)

                scaled_pixmap = self._current_pixmap.scaledToHeight(self.preview_label.height(),
                                                                    Qt.TransformationMode.SmoothTransformation)
                self.preview_label.setPixmap(scaled_pixmap)
                self.preview_label.setStyleSheet(
                    "background-color: transparent; border: 1px solid gray; border-radius: 8px;")
                return
            except Exception as e:
                self.log_message(
                    f"Erro ao carregar prévia existente ({caminho_preview}): {e}. Tentando gerar novamente.")

        # Se o preview não existe ou está desatualizado, tenta gerá-lo a partir do PSD
        #self.log_message(f"Gerando pré-visualização com Pillow para '{modelo_selecionado}'...")
        #self.preview_label.setText("Gerando prévia...")
        #QApplication.processEvents() # Atualiza a UI para mostrar a mensagem

        try:
            pil_img = Image.open(caminho_psd) # Abre o PSD usando Pillow
            if pil_img.mode != 'RGBA':
                pil_img = pil_img.convert('RGBA')
            pil_img.save(caminho_preview) # Salva como PNG para uso futuro

            qimage = QImage(pil_img.tobytes("raw", "RGBA"), pil_img.width, pil_img.height,
                            QImage.Format.Format_RGBA8888)
            self._current_pixmap = QPixmap.fromImage(qimage)

            scaled_pixmap = self._current_pixmap.scaledToHeight(self.preview_label.height(),
                                                                Qt.TransformationMode.SmoothTransformation)
            self.preview_label.setPixmap(scaled_pixmap)
            self.preview_label.setStyleSheet(
                "background-color: transparent; border: 1px solid gray; border-radius: 8px;")
            #self.log_message(f"Pré-visualização para '{modelo_selecionado}' gerada com sucesso.")

        except Exception as e:
            self.preview_label.clear()
            self.preview_label.setText("Erro ao gerar prévia")
            self.preview_label.setStyleSheet("background-color: orange; color: black; border-radius: 8px;")
            self.log_message(f"Erro inesperado ao gerar prévia para '{modelo_selecionado}': {e}")
            self._current_pixmap = None


if __name__ == "__main__":
    app = QApplication(sys.argv)
    qt_translator = QTranslator()

    translations_path = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
    locale_to_load = QLocale.Portuguese
    if qt_translator.load(locale_to_load, "qtbase", "_", translations_path):
        app.installTranslator(qt_translator)
    else:
        if qt_translator.load("qtbase_pt", "translations"):
            app.installTranslator(qt_translator)
        else:
            print("Falha ao carregar tradução do Qt para Português. Os botões padrão podem aparecer em inglês.")

    set_dark_theme(app)
    window = CartaoApp()
    window.show()
    sys.exit(app.exec())