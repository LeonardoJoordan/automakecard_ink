import sys
import os
import shutil
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

# Função para aplicar um tema escuro simples (opcional, pode ser expandido com QSS)
def set_dark_theme(app):
    # (Código do tema escuro permanece o mesmo da resposta anterior)
    # ... (O código do tema escuro foi omitido aqui para brevidade, mas deve ser incluído) ...
    dark_palette = QPalette()
    dark_palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.Base, QColor(42, 42, 42))  # Usado por QTableWidget background
    dark_palette.setColor(QPalette.ColorRole.AlternateBase, QColor(66, 66, 66))
    dark_palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)  # Cor do texto nos itens da tabela
    dark_palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    dark_palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
    dark_palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))  # Cor de seleção
    dark_palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)  # Cor do texto selecionado

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
        "QTableWidget::item:selected { background-color: #0078D7; color: white; }"  # Seleção na tabela
    )


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
            # Lida com o caso de uma única linha copiada que pode ter espaços mas não tabs
            if not rows_data: return

            table_data = []
            for row_str in rows_data:
                table_data.append(row_str.split('\t'))

            start_row = self.currentRow() if self.currentRow() != -1 else 0
            start_col = self.currentColumn() if self.currentColumn() != -1 else 0
            if not self.selectedIndexes():  # Se nada estiver selecionado, comece em 0,0
                start_row, start_col = 0, 0

            num_pasted_rows = len(table_data)
            max_pasted_cols_in_data = 0
            if table_data:
                max_pasted_cols_in_data = max(len(row_content) for row_content in table_data) if table_data[0] else 0

            # Garantir linhas suficientes
            required_rows = start_row + num_pasted_rows
            if required_rows > self.rowCount():
                self.setRowCount(required_rows)

            # Colar dados
            for r_idx, row_content in enumerate(table_data):
                current_table_row = start_row + r_idx
                for c_idx, cell_value in enumerate(row_content):
                    current_table_col = start_col + c_idx
                    if current_table_col < self.columnCount():  # Apenas colar dentro das colunas existentes
                        item = self.item(current_table_row, current_table_col)
                        if not item:
                            item = QTableWidgetItem(cell_value)
                            self.setItem(current_table_row, current_table_col, item)
                        else:
                            item.setText(cell_value)
        else:
            # Se não for texto, talvez seja uma imagem ou outro formato, deixe o padrão lidar
            # Criando um evento de paste para o handler padrão (pode não ser necessário)
            paste_event = QEvent(QEvent.Type.KeyPress)  # Aproximação
            super().keyPressEvent(paste_event)


class CartaoApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador de Cartões Personalizados PySide6")
        self.setFixedSize(1000, 550)  # Ajustado para acomodar melhor os widgets

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
        self.data_table.setRowCount(15)  # Linhas iniciais

        # --- INÍCIO DO CÓDIGO PARA AJUSTAR A ALTURA DA TABELA ---
        # Pegue a altura do cabeçalho horizontal
        header_height = self.data_table.horizontalHeader().height()

        # Estime a altura de uma linha. Você pode precisar ajustar este valor!
        # Experimente valores entre 25 e 35 para ver o que fica melhor.
        estimated_row_height = 28

        # Calcule a altura total para as 10 linhas
        total_rows_height = 14 * estimated_row_height

        # Calcule a altura alvo para a tabela
        # Adicionamos a altura da scrollbar horizontal (caso ela apareça)
        # e um pequeno padding (5 pixels) para respiro.
        table_target_height = header_height + total_rows_height + self.data_table.horizontalScrollBar().height() + 5

        # Defina a altura fixa da tabela
        self.data_table.setFixedHeight(table_target_height)
        # --- FIM DO CÓDIGO PARA AJUSTAR A ALTURA DA TABELA ---

        # Ajustar largura das colunas para preencher o espaço
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)  # Isso é para largura
        tabela_layout.addWidget(self.data_table)
        # Ajustar largura das colunas para preencher o espaço
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        tabela_layout.addWidget(self.data_table)

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

        main_layout.addWidget(tabela_container, 2)  # Coluna esquerda

        # --- Coluna direita (igual à anterior) ---
        direita_frame = QWidget()
        direita_layout = QVBoxLayout(direita_frame)
        direita_frame.setFixedWidth(380)

        self.preview_label = QLabel("Aguardando seleção de modelo")
        self.preview_label.setObjectName("PreviewLabel")
        self.preview_label.setFixedSize(300, 200)
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setScaledContents(False)
        direita_layout.addWidget(self.preview_label, 0, Qt.AlignmentFlag.AlignHCenter)

        self.modelo_combobox = QComboBox()
        self.modelo_combobox.addItem("(nenhum modelo disponível)")
        self.modelo_combobox.currentTextChanged.connect(self.atualizar_preview_modelo)
        direita_layout.addWidget(self.modelo_combobox)

        self.btn_gerar = QPushButton("Gerar Cartões")
        self.btn_gerar.clicked.connect(self.gerar_cartoes)  # Conectado aqui
        direita_layout.addWidget(self.btn_gerar)

        self.log_textbox = QTextEdit()
        self.log_textbox.setReadOnly(True)
        #self.log_textbox.setFixedHeight(120)
        self.log_textbox.append("Log do programa...\n")
        direita_layout.addWidget(self.log_textbox)

        botoes_modelo_widget = QWidget()
        botoes_modelo_layout = QHBoxLayout(botoes_modelo_widget)
        botoes_modelo_layout.setContentsMargins(0, 0, 0, 0)
        botoes_modelo_layout.setSpacing(5)

        self.btn_adicionar_modelo = QPushButton("Adicionar modelo")
        self.btn_adicionar_modelo.clicked.connect(self.adicionar_modelo)
        botoes_modelo_layout.addWidget(self.btn_adicionar_modelo)

        self.btn_modificar_modelo = QPushButton("Modificar modelo")
        self.btn_modificar_modelo.clicked.connect(self.modificar_modelo)
        botoes_modelo_layout.addWidget(self.btn_modificar_modelo)

        self.btn_excluir_modelo = QPushButton("Excluir modelo")
        self.btn_excluir_modelo.clicked.connect(self.excluir_modelo)
        botoes_modelo_layout.addWidget(self.btn_excluir_modelo)

        direita_layout.addWidget(botoes_modelo_widget)
        direita_layout.addStretch(1)
        main_layout.addWidget(direita_frame, 1)  # Coluna direita

        self.atualizar_modelos_combobox()
        self._current_pixmap = None

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
            self.data_table.setRowCount(0)  # Remove todas as linhas
            self.data_table.setRowCount(20)  # Adiciona 20 linhas vazias de volta (ou como preferir)
            # Alternativamente, para limpar o conteúdo sem remover linhas:
            # for r in range(self.data_table.rowCount()):
            #     for c in range(self.data_table.columnCount()):
            #         item = self.data_table.item(r, c)
            #         if item:
            #             item.setText("")
            self.log_message("Tabela limpa.")

    def garantir_pasta_modelos(self):
        if not os.path.exists("modelos"):
            os.makedirs("modelos")
            self.log_message("Pasta 'modelos' criada.")

    def log_message(self, message):
        self.log_textbox.append(message)
        self.log_textbox.ensureCursorVisible()

    # ... (métodos adicionar_modelo, modificar_modelo, excluir_modelo, atualizar_preview_modelo permanecem os mesmos)
    # ... (O código dos métodos de gerenciamento de modelo e preview foi omitido para brevidade)
    def atualizar_modelos_combobox(self):
        self.garantir_pasta_modelos()
        try:
            arquivos = [f for f in os.listdir("modelos") if f.lower().endswith(".xcf")]
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
            elif arquivos:  # Se a seleção anterior não existe mais, seleciona o primeiro
                self.modelo_combobox.setCurrentIndex(0)
            else:  # Não há arquivos, adiciona placeholder
                self.modelo_combobox.addItem("(nenhum modelo disponível)")
        else:
            self.modelo_combobox.addItem("(nenhum modelo disponível)")

        self.modelo_combobox.blockSignals(False)

        # Força a atualização do preview com base na seleção atual (ou falta dela)
        self.atualizar_preview_modelo(self.modelo_combobox.currentText())

    def adicionar_modelo(self):
        self.garantir_pasta_modelos()
        arquivo, _ = QFileDialog.getOpenFileName(
            self,
            "Selecione um modelo (.xcf)",
            "modelos",
            "Arquivos do GIMP (*.xcf)"
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
            QMessageBox.information(self, "Modificar modelo", "Nenhum modelo selecionado para modificar.")
            return

        novo_arquivo, _ = QFileDialog.getOpenFileName(
            self,
            "Selecione o novo arquivo do modelo (.xcf)",
            "modelos",
            "Arquivos do GIMP (*.xcf)"
        )
        if novo_arquivo:
            destino = os.path.join("modelos", modelo_selecionado)
            try:
                shutil.copy2(novo_arquivo, destino)
                self.log_message(f"Modelo '{modelo_selecionado}' foi modificado.")
                preview_path = os.path.join("modelos", f"{modelo_selecionado}_preview.png")
                if os.path.exists(preview_path):
                    os.remove(preview_path)
                self.atualizar_preview_modelo(modelo_selecionado)
            except Exception as e:
                QMessageBox.critical(self, "Erro ao modificar", f"Não foi possível substituir o arquivo: {e}")
                self.log_message(f"Erro ao modificar '{modelo_selecionado}': {e}")

    def excluir_modelo(self):
        modelo_selecionado = self.modelo_combobox.currentText()
        if not modelo_selecionado or modelo_selecionado == "(nenhum modelo disponível)":
            QMessageBox.information(self, "Excluir modelo", "Nenhum modelo selecionado para excluir.")
            return

        resp = QMessageBox.question(self, "Excluir modelo",
                                    f"Tem certeza que deseja excluir o modelo '{modelo_selecionado}'?",
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if resp == QMessageBox.StandardButton.Yes:
            caminho_xcf = os.path.join("modelos", modelo_selecionado)
            preview_path = os.path.join("modelos", f"{modelo_selecionado}_preview.png")
            try:
                if os.path.exists(caminho_xcf):
                    os.remove(caminho_xcf)
                if os.path.exists(preview_path):
                    os.remove(preview_path)

                self.log_message(f"Modelo '{modelo_selecionado}' foi excluído.")
                self.atualizar_modelos_combobox()
            except Exception as e:
                QMessageBox.critical(self, "Erro ao excluir", f"Erro: {e}")
                self.log_message(f"Erro ao excluir '{modelo_selecionado}': {e}")

    def atualizar_preview_modelo(self, modelo_selecionado=None):
        # (O código deste método permanece o mesmo da resposta anterior)
        # ... (O código do atualizar_preview_modelo foi omitido para brevidade)
        if modelo_selecionado is None:
            modelo_selecionado = self.modelo_combobox.currentText()

        if not modelo_selecionado or modelo_selecionado == "(nenhum modelo disponível)":
            self.preview_label.clear()
            self.preview_label.setText("Aguardando seleção de modelo")
            self.preview_label.setStyleSheet(
                "background-color: #404040; color: white; border-radius: 8px; border: 1px solid #505050;")
            self._current_pixmap = None
            return

        caminho_xcf = os.path.join("modelos", modelo_selecionado)
        caminho_preview = os.path.join("modelos", f"{modelo_selecionado}_preview.png")

        if not os.path.exists(caminho_xcf):
            self.preview_label.clear()
            self.preview_label.setText(f"Arquivo XCF '{modelo_selecionado}' não encontrado.")
            self.preview_label.setStyleSheet("background-color: orange; color: black; border-radius: 8px;")
            self.log_message(f"Preview erro: '{caminho_xcf}' não existe.")
            self._current_pixmap = None
            return

        if os.path.exists(caminho_preview) and os.path.getmtime(caminho_preview) > os.path.getmtime(caminho_xcf):
            try:
                self.log_message(f"Carregando pré-visualização existente para '{modelo_selecionado}'.")
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

        imagemagick_command = shutil.which("magick")
        if not imagemagick_command:
            imagemagick_command = r"C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe"
            if not os.path.exists(imagemagick_command):
                imagemagick_command_fallback = r"C:\Program Files\ImageMagick-7.1.1-Q16\magick.exe"
                if os.path.exists(imagemagick_command_fallback):
                    imagemagick_command = imagemagick_command_fallback
                else:
                    self.preview_label.clear()
                    self.preview_label.setText("ImageMagick não configurado.")
                    self.preview_label.setStyleSheet("background-color: red; color: white; border-radius: 8px;")
                    self.log_message("Erro: Executável do ImageMagick não encontrado.")
                    self._current_pixmap = None
                    return

        comando_convert = [imagemagick_command, caminho_xcf, "-flatten", caminho_preview]
        self.log_message(f"Gerando pré-visualização com ImageMagick para '{modelo_selecionado}'...")
        self.preview_label.setText("Gerando prévia...")
        QApplication.processEvents()

        try:
            result = subprocess.run(comando_convert, check=True, capture_output=True, text=True, shell=False,
                                    errors='replace')
            if result.stdout: self.log_message(f"ImageMagick STDOUT: {result.stdout.strip()}")
            if result.stderr: self.log_message(f"ImageMagick STDERR: {result.stderr.strip()}")

            if not os.path.exists(caminho_preview):
                raise FileNotFoundError(
                    f"ImageMagick executou, mas o arquivo de preview '{caminho_preview}' não foi criado.")

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
            self.log_message(f"Pré-visualização para '{modelo_selecionado}' gerada com sucesso.")

        except FileNotFoundError as e:
            self.preview_label.clear()
            self.preview_label.setText("Erro: ImageMagick falhou.")
            self.preview_label.setStyleSheet("background-color: red; color: white; border-radius: 8px;")
            self.log_message(f"Erro de Arquivo (ImageMagick): {e}")
            self._current_pixmap = None
        except subprocess.CalledProcessError as e:
            self.preview_label.clear()
            self.preview_label.setText("Erro ImageMagick")
            self.preview_label.setStyleSheet("background-color: red; color: white; border-radius: 8px;")
            self.log_message(f"Erro no subprocesso do ImageMagick (código {e.returncode}):")
            if e.stdout: self.log_message(f"STDOUT: {e.stdout.strip()}")
            if e.stderr: self.log_message(f"STDERR: {e.stderr.strip()}")
            self._current_pixmap = None
        except Exception as e:
            self.preview_label.clear()
            self.preview_label.setText("Erro ao gerar prévia")
            self.preview_label.setStyleSheet("background-color: orange; color: black; border-radius: 8px;")
            self.log_message(f"Erro inesperado ao gerar prévia para '{modelo_selecionado}': {e}")
            self._current_pixmap = None

    def gerar_cartoes(self):
        # Obter dados da tabela
        headers = [self.data_table.horizontalHeaderItem(j).text().lower().strip() for j in
                   range(self.data_table.columnCount())]
        dados_para_cartoes = []
        for i in range(self.data_table.rowCount()):
            row_data_list = []
            is_row_empty = True
            for j in range(self.data_table.columnCount()):
                item = self.data_table.item(i, j)
                text = item.text().strip() if item and item.text() else ""
                row_data_list.append(text)
                if text:
                    is_row_empty = False

            # Adiciona a linha apenas se não estiver vazia
            if not is_row_empty:
                dados_para_cartoes.append(row_data_list)

        if not dados_para_cartoes:
            QMessageBox.warning(self, "Gerar Cartões", "Nenhum dado na tabela para gerar cartões.")
            return

        modelo_selecionado = self.modelo_combobox.currentText()
        if modelo_selecionado == "(nenhum modelo disponível)":
            QMessageBox.warning(self, "Gerar Cartões", "Por favor, selecione um modelo de cartão.")
            return

        output_folder = "cartoes_gerados"
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        self.log_message(f"Iniciando geração de cartões com o modelo '{modelo_selecionado}'...")

        caminho_xcf_modelo = os.path.join("modelos", modelo_selecionado)
        # Tenta encontrar 'magick' no PATH primeiro, depois usa caminho fixo
        imagemagick_command = shutil.which("magick")
        if not imagemagick_command:
            imagemagick_command = r"C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe"  # AJUSTE SE NECESSÁRIO
            if not os.path.exists(imagemagick_command):
                self.log_message(
                    "Erro CRÍTICO: ImageMagick não encontrado. Verifique a instalação e o caminho no código.")
                QMessageBox.critical(self, "Erro ImageMagick",
                                     "ImageMagick não encontrado. A geração de cartões foi abortada.")
                return

        for i, row_values in enumerate(dados_para_cartoes):
            row_dict = {headers[j]: (str(row_values[j]) if j < len(row_values) else "") for j in range(len(headers))}

            temp_png_path = os.path.join(output_folder, f"temp_modelo_base_{i}.png")
            output_card_path = os.path.join(output_folder,
                                            f"cartao_{row_dict.get('nome', i + 1).replace(' ', '_')}.png")

            comando_convert_xcf_to_png = [
                imagemagick_command,
                caminho_xcf_modelo,
                "-flatten",
                temp_png_path
            ]

            try:
                subprocess.run(comando_convert_xcf_to_png, check=True, capture_output=True, text=True, errors='replace')

                img = Image.open(temp_png_path)
                draw = ImageDraw.Draw(img)

                font_path = "arial.ttf"  # Tenta carregar Arial
                try:
                    font_nome = ImageFont.truetype(font_path, 40)
                    font_tratamento = ImageFont.truetype(font_path, 30)
                    font_conjuge = ImageFont.truetype(font_path, 30)
                    font_data = ImageFont.truetype(font_path, 25)
                except IOError:
                    self.log_message(f"Aviso: Fonte '{font_path}' não encontrada. Usando fonte padrão.")
                    font_nome = ImageFont.load_default()  # Fallback para fonte padrão
                    font_tratamento = ImageFont.load_default()
                    font_conjuge = ImageFont.load_default()
                    font_data = ImageFont.load_default()

                # Defina as posições X, Y para cada campo de texto no seu modelo
                # Exemplo (AJUSTE ESTAS COORDENADAS):
                pos_tratamento = (50, 100)
                pos_nome = (50, 150)
                pos_conjuge = (50, 200)
                pos_data = (50, 250)
                text_color = (0, 0, 0)  # Preto

                if "tratamento" in row_dict and row_dict["tratamento"]:
                    draw.text(pos_tratamento, row_dict["tratamento"], font=font_tratamento, fill=text_color)
                if "nome" in row_dict and row_dict["nome"]:
                    draw.text(pos_nome, row_dict["nome"], font=font_nome, fill=text_color)
                if "conjuge" in row_dict and row_dict["conjuge"]:  # Adicionado para conjuge
                    draw.text(pos_conjuge, row_dict["conjuge"], font=font_conjuge, fill=text_color)
                if "data" in row_dict and row_dict["data"]:
                    draw.text(pos_data, row_dict["data"], font=font_data, fill=text_color)

                img.save(output_card_path)
                os.remove(temp_png_path)  # Limpa o PNG base temporário

                self.log_message(f"Cartão {i + 1} gerado: {output_card_path}")

            except subprocess.CalledProcessError as e:
                self.log_message(
                    f"Erro ImageMagick ao processar cartão {i + 1} (modelo base): {e.returncode}\nSTDOUT: {e.stdout}\nSTDERR: {e.stderr}")
                if os.path.exists(temp_png_path): os.remove(temp_png_path)  # Tenta limpar
            except FileNotFoundError:  # Para o caso de ImageMagick não ser encontrado no meio do loop (improvável aqui)
                self.log_message("Erro: ImageMagick não encontrado durante a geração. Abortando.")
                if os.path.exists(temp_png_path): os.remove(temp_png_path)
                break
            except Exception as e:
                self.log_message(f"Erro inesperado ao gerar cartão {i + 1} ('{row_dict.get('nome', '')}'): {e}")
                if os.path.exists(temp_png_path): os.remove(temp_png_path)
            QApplication.processEvents()  # Mantém a UI responsiva

        self.log_message("Geração de cartões concluída.")
        QMessageBox.information(self, "Geração Concluída", f"Cartões foram gerados na pasta '{output_folder}'.")

    # NOVA FUNÇÃO para interagir com o GIMP via Python-Fu
    def gerar_cartoes_com_gimp(self):
        self.log_message("Iniciando geração de cartões via GIMP Python-Fu...")
        QApplication.processEvents()  # Para o log aparecer imediatamente

        # 1. Obter dados da tabela
        headers = [self.data_table.horizontalHeaderItem(j).text().lower().strip() for j in
                   range(self.data_table.columnCount())]
        dados_para_cartoes = []
        for i in range(self.data_table.rowCount()):
            row_data_list = []
            is_row_empty = True
            for j in range(self.data_table.columnCount()):
                item = self.data_table.item(i, j)
                text = item.text().strip() if item and item.text() else ""
                row_data_list.append(text)
                if text:
                    is_row_empty = False
            if not is_row_empty:
                dados_para_cartoes.append(row_data_list)

        if not dados_para_cartoes:
            QMessageBox.warning(self, "Gerar Cartões", "Nenhum dado na tabela para gerar cartões.")
            return

        # 2. Obter modelo selecionado
        modelo_selecionado_nome_arquivo = self.modelo_combobox.currentText()
        if modelo_selecionado_nome_arquivo == "(nenhum modelo disponível)":
            QMessageBox.warning(self, "Gerar Cartões", "Por favor, selecione um modelo de cartão.")
            return

        caminho_completo_template_xcf = os.path.abspath(os.path.join("modelos", modelo_selecionado_nome_arquivo))
        if not os.path.exists(caminho_completo_template_xcf):
            QMessageBox.critical(self, "Erro", f"Arquivo de modelo não encontrado: {caminho_completo_template_xcf}")
            self.log_message(f"ERRO: Arquivo de modelo XCF não encontrado em '{caminho_completo_template_xcf}'")
            return

        # 3. Criar pasta de saída
        pasta_saida = "cartoes_gerados"
        if not os.path.exists(pasta_saida):
            try:
                os.makedirs(pasta_saida)
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Não foi possível criar a pasta de saída '{pasta_saida}': {e}")
                self.log_message(f"ERRO: Não foi possível criar a pasta de saída '{pasta_saida}': {e}")
                return

        self.log_message(f"Modelo selecionado: '{modelo_selecionado_nome_arquivo}'")
        QApplication.processEvents()

        # 4. Encontrar o executável do GIMP (Corrigido e completo)
        comando_gimp = shutil.which("gimp-console-3.0")  # Tenta primeiro para GIMP 3.0
        if not comando_gimp:
            self.log_message("gimp-console-3.0 não encontrado, tentando gimp-3.0...")
            comando_gimp = shutil.which("gimp-3.0")  # Outra possibilidade para GIMP 3.0
        if not comando_gimp:
            self.log_message("gimp-3.0 não encontrado, tentando gimp (para GIMP 2.10 ou genérico)...")
            comando_gimp = shutil.which("gimp")  # Fallback para GIMP 2.10 ou genérico
        if not comando_gimp:
            self.log_message("Nenhum GIMP encontrado no PATH, tentando caminho fixo para Windows...")
            # Seu caminho fixo para Windows, talvez precise ser ajustado para GIMP 3.0
            # Para GIMP 3.0, o executável de console pode ainda ser gimp-console-2.10.exe se for uma atualização,
            # ou algo como gimp-console-3.0.exe. Verifique sua instalação.
            # Exemplo: r"C:\Program Files\GIMP 3\bin\gimp-console-3.0.exe" ou similar
            caminho_gimp_fixo_windows = r"C:\Program Files\GIMP 3\bin\gimp-console-3.0.exe"  # AJUSTE ESTE CAMINHO SE NECESSÁRIO

            if os.path.exists(caminho_gimp_fixo_windows):
                comando_gimp = caminho_gimp_fixo_windows
                self.log_message(f"Usando GIMP do caminho fixo: {comando_gimp}")
            else:
                self.log_message(
                    f"ERRO: Executável do GIMP não encontrado no PATH e o caminho fixo '{caminho_gimp_fixo_windows}' também falhou.")
                QMessageBox.critical(self, "Erro GIMP",
                                     "Executável do GIMP não encontrado no PATH ou em caminho fixo.\n"
                                     "Verifique a instalação do GIMP ou ajuste o caminho no código.")
                return  # Aborta se não encontrar o GIMP
        else:
            self.log_message(f"GIMP encontrado em: {comando_gimp}")

        nome_script_gimp = "python_fu_atualizador_cartao_gimp"  # Nome registrado no script Python-Fu

        # 5. Loop para processar cada cartão
        for i, row_values in enumerate(dados_para_cartoes):
            dados_para_este_cartao = {}
            for col_idx, header_name in enumerate(headers):
                dados_para_este_cartao[header_name] = row_values[col_idx] if col_idx < len(row_values) else ""

            dados_para_este_cartao["template_path"] = caminho_completo_template_xcf
            nome_base_arquivo_saida = str(dados_para_este_cartao.get("nome", f"cartao_{i + 1}")).replace(" ",
                                                                                                         "_").replace(
                "/", "-").replace("\\", "-")
            caminho_saida_png = os.path.abspath(os.path.join(pasta_saida, f"{nome_base_arquivo_saida}.png"))
            dados_para_este_cartao["output_filename"] = caminho_saida_png

            caminho_json_temp = os.path.abspath(os.path.join(pasta_saida, f"dados_temp_cartao_{i}.json"))
            try:
                with open(caminho_json_temp, 'w', encoding='utf-8') as f_json:
                    json.dump(dados_para_este_cartao, f_json, ensure_ascii=False, indent=4)
            except Exception as e:
                self.log_message(f"Erro ao criar arquivo JSON temporário '{caminho_json_temp}': {e}")
                continue

            self.log_message(
                f"Processando cartão {i + 1} para: {dados_para_este_cartao.get('nome', 'N/A')}. JSON: {caminho_json_temp}")
            QApplication.processEvents()

            batch_command = f'({nome_script_gimp} RUN-NONINTERACTIVE "{caminho_json_temp}") (gimp-quit 0)'

            try:
                # Para maior robustez, especialmente no Windows com caminhos contendo espaços:
                # Envolver o comando_gimp em aspas se ele contiver espaços.
                # No entanto, passar como lista para subprocess.run é geralmente mais seguro.
                comando_completo = [comando_gimp, "-i", "-b", batch_command]
                self.log_message(f"Executando GIMP com comando: {' '.join(comando_completo)}")

                processo = subprocess.run(comando_completo,
                                          capture_output=True, text=True, check=False,
                                          encoding='utf-8', errors='replace', shell=False)  # shell=False é mais seguro

                if processo.returncode == 0:
                    self.log_message(f"Cartão '{caminho_saida_png}' gerado com sucesso (GIMP stdout abaixo).")
                    if processo.stdout: self.log_message(f"GIMP STDOUT: {processo.stdout.strip()}")
                    # GIMP frequentemente usa stderr para mensagens informativas também
                    if processo.stderr: self.log_message(f"GIMP STDERR: {processo.stderr.strip()}")
                else:
                    self.log_message(f"ERRO ao gerar cartão com GIMP (código de retorno: {processo.returncode}).")
                    self.log_message(f"GIMP STDOUT: {processo.stdout.strip()}")
                    self.log_message(f"GIMP STDERR: {processo.stderr.strip()}")

            except FileNotFoundError:
                self.log_message(
                    f"ERRO CRÍTICO: Executável do GIMP não encontrado em '{comando_gimp}'. A geração foi interrompida.")
                QMessageBox.critical(self, "Erro GIMP", f"Executável do GIMP não encontrado: {comando_gimp}")
                break
            except Exception as e:
                self.log_message(f"ERRO inesperado ao executar o GIMP: {e}")
            finally:
                if os.path.exists(caminho_json_temp):
                    try:
                        os.remove(caminho_json_temp)
                    except Exception as e_del:
                        self.log_message(
                            f"Aviso: não foi possível remover o arquivo temporário {caminho_json_temp}: {e_del}")
            QApplication.processEvents()

        self.log_message("Geração de cartões concluída.")
        QMessageBox.information(self, "Geração Concluída",
                                f"Processo de geração de cartões finalizado. Verifique a pasta '{pasta_saida}' e o log do programa.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    # --- INÍCIO DO CÓDIGO DE TRADUÇÃO ---
    # Cria um objeto QTranslator
    qt_translator = QTranslator()

    # Tenta carregar o arquivo de tradução do Qt para o Português
    # Ele busca nos caminhos padrão de tradução do Qt
    translations_path = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
    locale_to_load = QLocale.Portuguese  # Para Português geral, ou QLocale("pt_BR") para Brasil especificamente

    # O nome do arquivo de tradução para os componentes base do Qt é geralmente "qtbase"
    if qt_translator.load(locale_to_load, "qtbase", "_", translations_path):
        app.installTranslator(qt_translator)
        print(f"Tradução do Qt para Português carregada de: {translations_path}")
    else:
        # Tenta um caminho relativo (útil se você distribuir os arquivos .qm manualmente)
        if qt_translator.load("qtbase_pt", "translations"):  # Procura em uma pasta 'translations'
            app.installTranslator(qt_translator)
            print("Tradução do Qt para Português carregada de: translations/qtbase_pt.qm")
        else:
            print("Falha ao carregar tradução do Qt para Português. Os botões padrão podem aparecer em inglês.")
    # --- FIM DO CÓDIGO DE TRADUÇÃO ---

    set_dark_theme(app)
    window = CartaoApp()
    window.show()
    sys.exit(app.exec())