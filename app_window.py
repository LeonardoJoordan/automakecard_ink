# Imports de bibliotecas padrão do Python
import os
import shutil
import time
import gc # Garbage Collector, para _forcar_fechamento_photoshop
import re
from dialogo_regras_texto import GerenciarRegrasTextoDialog
from collections import defaultdict
# Imports de bibliotecas de terceiros (instaladas com pip)
from PIL import Image # Para manipulação de imagens no preview
import psutil # Para _forcar_fechamento_photoshop
import win32com.client # Para interagir com o Photoshop

# Imports do PySide6
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTextEdit, QLabel, QComboBox, QPushButton, QFileDialog, QMessageBox,
    QTableWidgetItem, QHeaderView
)
from PySide6.QtGui import (
    QFont, QPixmap, QImage, QKeySequence, QIcon
)
from PySide6.QtCore import Qt, QSize, QEvent

# Imports dos módulos do projeto
import utils
import ps_utils
from custom_widgets import CustomTableWidget

from caixa_para_modificar_modelo import CaixaModificarModeloDialog
from dialogo_gerenciar_regras import GerenciarRegrasDialog

# ________________________________________________________________________________________________

# Constantes do Módulo
PASTA_PADRAO_MODELOS = "modelos"
PASTA_PADRAO_SAIDA = "cartoes_gerados"

# A constante para o formato de exportação do Photoshop é melhor definida em ps_utils.py,
# mas podemos mantê-la aqui se for usada apenas neste arquivo. Por hora, vamos deixá-la aqui.
PS_EXPORT_FORMAT_PNG = 13 # PNG


# ________________________________________________________________________________________________

class CartaoApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoMakeCardPSD - Gerador de Cartões")
        self.setFixedSize(1000, 600)  # Ajustei um pouco a altura para acomodar melhor
        self.setWindowIcon(
            QIcon(os.path.join(os.path.dirname(__file__), 'icon.png')))  # Assumindo que você tem um icon.png

        # Inicialização de atributos de dados
        self.configuracoes_modelos = utils.carregar_configuracoes_camadas_modelos()
        self.output_dir = utils.load_last_output_dir()
        if not self.output_dir or not os.path.exists(self.output_dir):
            self.output_dir = os.path.abspath(PASTA_PADRAO_SAIDA)
            if not os.path.exists(self.output_dir):  # Garante que a pasta de saída padrão exista
                os.makedirs(self.output_dir)

        self._current_pixmap = None  # Para armazenar o QPixmap do preview atual
        self.table_headers = []  # Será preenchido dinamicamente

        # Widget principal e layout
        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        # --- Coluna da Esquerda (Tabela de Dados) ---
        esquerda_container = QWidget()
        esquerda_layout = QVBoxLayout(esquerda_container)

        self.data_table = CustomTableWidget(self)  # Usando seu CustomTableWidget
        # As colunas e cabeçalhos serão definidos por _atualizar_tabela_para_modelo
        self.data_table.setColumnCount(0)
        self.data_table.setRowCount(15)  # Número inicial de linhas

        # Ajuste de altura da tabela (pode precisar de refinamento)
        header_height = self.data_table.horizontalHeader().height()
        estimated_row_height = 28
        total_rows_height = 14 * estimated_row_height
        table_target_height = header_height + total_rows_height + self.data_table.horizontalScrollBar().height() + 5
        self.data_table.setFixedHeight(table_target_height)
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        esquerda_layout.addWidget(self.data_table)

        # Botões de gerenciamento da tabela
        # Botão para gerenciar as regras de texto
        self.btn_gerenciar_regras = QPushButton("Gerenciar Regras de Texto")
        # A conexão com a função que abre o diálogo será feita em um passo futuro
        self.btn_gerenciar_regras.clicked.connect(self.abrir_dialogo_gerenciar_regras)
        self.btn_gerenciar_regras.setEnabled(False)  # Começa desabilitado
        esquerda_layout.addWidget(self.btn_gerenciar_regras)

        main_layout.addWidget(esquerda_container, 2)  # Coluna da esquerda ocupa 2/3 da largura

        # --- Coluna da Direita (Controles e Preview) ---
        direita_container = QWidget()
        direita_layout = QVBoxLayout(direita_container)
        direita_container.setFixedWidth(380)  # Largura fixa para a coluna da direita

        # Preview do Modelo
        self.preview_label = QLabel("Selecione um modelo")
        self.preview_label.setObjectName("PreviewLabel")  # Para possível estilização CSS
        self.preview_label.setFixedSize(300, 200)
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setScaledContents(False)  # O escalonamento será feito manualmente
        self.preview_label.setStyleSheet("border: 1px solid gray; border-radius: 5px; background-color: #f0f0f0;")
        direita_layout.addWidget(self.preview_label, 0, Qt.AlignmentFlag.AlignHCenter)

        # Botão Gerar Cartões
        self.btn_gerar_cartoes = QPushButton("Gerar Cartões")
        self.btn_gerar_cartoes.setStyleSheet(
            "font-weight: bold; background-color: #4CAF50; color: white; padding: 8px;")
        self.btn_gerar_cartoes.setFixedHeight(40)
        self.btn_gerar_cartoes.clicked.connect(self.gerar_cartoes)
        self.btn_gerar_cartoes.setEnabled(False)  # Desabilitado até um modelo válido ser selecionado
        direita_layout.addWidget(self.btn_gerar_cartoes)

        # ComboBox para Seleção de Modelo
        label_selecionar_modelo = QLabel("Modelo PSD:")
        direita_layout.addWidget(label_selecionar_modelo)
        self.modelo_combobox = QComboBox()
        self.modelo_combobox.currentTextChanged.connect(self._quando_modelo_mudar)
        direita_layout.addWidget(self.modelo_combobox)

        # Botões de Gerenciamento de Modelos
        botoes_modelo_layout = QHBoxLayout()
        self.btn_adicionar_modelo = QPushButton("Adicionar")
        self.btn_adicionar_modelo.clicked.connect(self.adicionar_modelo)
        self.btn_modificar_modelo = QPushButton("Modificar")
        self.btn_modificar_modelo.clicked.connect(self.modificar_modelo)
        self.btn_excluir_modelo = QPushButton("Excluir")
        self.btn_excluir_modelo.clicked.connect(self.excluir_modelo)

        botoes_modelo_layout.addWidget(self.btn_adicionar_modelo)
        botoes_modelo_layout.addWidget(self.btn_modificar_modelo)
        botoes_modelo_layout.addWidget(self.btn_excluir_modelo)
        direita_layout.addLayout(botoes_modelo_layout)

        # Caixa de Log
        label_log = QLabel("Log de Eventos:")
        direita_layout.addWidget(label_log)
        self.log_textbox = QTextEdit()
        self.log_textbox.setReadOnly(True)
        self.log_textbox.setFixedHeight(100)  # Altura para o log
        self.log_message("Interface iniciada.") # Será chamado quando log_message existir
        direita_layout.addWidget(self.log_textbox)

        # Seleção da Pasta de Saída
        saida_dir_label_desc = QLabel("Pasta de Saída:")
        direita_layout.addWidget(saida_dir_label_desc)

        saida_dir_layout = QHBoxLayout()
        self.saida_dir_label_path = QLabel(self.output_dir)  # Mostra o caminho
        self.saida_dir_label_path.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.saida_dir_label_path.setToolTip(self.output_dir)  # Tooltip para caminhos longos
        self.btn_selecionar_saida = QPushButton("Alterar...")
        self.btn_selecionar_saida.clicked.connect(self.selecionar_pasta_saida)

        saida_dir_layout.addWidget(self.saida_dir_label_path, 1)  # Label ocupa mais espaço
        saida_dir_layout.addWidget(self.btn_selecionar_saida, 0)
        direita_layout.addLayout(saida_dir_layout)

        direita_layout.addStretch(1)  # Empurra tudo para cima
        main_layout.addWidget(direita_container, 1)  # Coluna da direita ocupa 1/3

        # Carregamento inicial dos modelos e log
        self.atualizar_modelos_combobox() # Será chamado quando o método existir
        self.log_message(f"Configurações de {len(self.configuracoes_modelos)} modelos carregadas.")
        self.log_message(f"Pasta de saída padrão: {self.output_dir}")

#________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def log_message(self, message: str):
        """
        Adiciona uma mensagem à caixa de texto de log da interface.

        Args:
            message: A string da mensagem a ser registrada.
        """
        if hasattr(self, 'log_textbox') and self.log_textbox is not None:
            self.log_textbox.append(message)
            self.log_textbox.ensureCursorVisible() # Garante que a última mensagem seja visível
            QApplication.processEvents() # Permite que a UI atualize a mensagem imediatamente
        else:
            # Fallback caso o log_textbox ainda não exista ou tenha sido removido
            print(f"LOG: {message}")

#________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def garantir_pasta_modelos(self):
        """
        Verifica se a pasta padrão de modelos (definida em PASTA_PADRAO_MODELOS)
        existe. Se não existir, tenta criá-la.
        """
        if not os.path.exists(PASTA_PADRAO_MODELOS):
            try:
                os.makedirs(PASTA_PADRAO_MODELOS)
                self.log_message(f"Pasta de modelos '{PASTA_PADRAO_MODELOS}' criada com sucesso.")
            except OSError as e:
                self.log_message(f"ERRO Crítico: Não foi possível criar a pasta de modelos '{PASTA_PADRAO_MODELOS}'. Erro: {e}")
                QMessageBox.critical(self, "Erro de Pasta",
                                     f"Não foi possível criar a pasta de modelos necessária: '{PASTA_PADRAO_MODELOS}'.\n"
                                     f"Verifique as permissões ou crie a pasta manualmente.\nErro: {e}")
                # Poderíamos considerar fechar a aplicação aqui ou desabilitar funcionalidades
                # que dependem desta pasta, mas por ora apenas logamos e avisamos.

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def atualizar_modelos_combobox(self):
        """
        Verifica a pasta de modelos, lista os arquivos .psd disponíveis
        e atualiza o QComboBox de seleção de modelos.
        Também chama _quando_modelo_mudar para atualizar a UI.
        """
        self.garantir_pasta_modelos()  # Garante que a pasta existe

        try:
            arquivos_psd = [f for f in os.listdir(PASTA_PADRAO_MODELOS) if f.lower().endswith(".psd")]
        except FileNotFoundError:
            self.log_message(
                f"ERRO Crítico: Pasta de modelos '{PASTA_PADRAO_MODELOS}' não encontrada ao listar arquivos, mesmo após tentativa de criação.")
            # Isso não deveria acontecer se garantir_pasta_modelos funcionou, mas é uma proteção.
            arquivos_psd = []
        except PermissionError:
            self.log_message(f"ERRO Crítico: Sem permissão para ler a pasta de modelos '{PASTA_PADRAO_MODELOS}'.")
            QMessageBox.critical(self, "Erro de Permissão",
                                 f"Não foi possível ler a pasta de modelos '{PASTA_PADRAO_MODELOS}'.\n"
                                 "Verifique as permissões da pasta.")
            arquivos_psd = []

        self.modelo_combobox.blockSignals(True)  # Bloqueia sinais para evitar chamadas múltiplas a _quando_modelo_mudar

        texto_selecionado_anteriormente = self.modelo_combobox.currentText()
        self.modelo_combobox.clear()

        if arquivos_psd:
            self.modelo_combobox.addItems(sorted(arquivos_psd))  # Adiciona em ordem alfabética
            # Tenta restaurar a seleção anterior
            if texto_selecionado_anteriormente in arquivos_psd:
                self.modelo_combobox.setCurrentText(texto_selecionado_anteriormente)
            elif self.modelo_combobox.count() > 0:
                self.modelo_combobox.setCurrentIndex(0)  # Seleciona o primeiro da lista
            # Se não havia seleção anterior e a lista não está vazia, o primeiro já estará selecionado.
        else:
            self.modelo_combobox.addItem(utils.TEXTO_NENHUM_MODELO)

        self.modelo_combobox.blockSignals(False)  # Libera os sinais

        # Chama explicitamente _quando_modelo_mudar para garantir que a UI (preview, tabela)
        # seja atualizada com base na seleção atual do combobox.
        # Isso é importante mesmo que a seleção não tenha mudado, pois as camadas
        # configuradas para o mesmo modelo podem ter sido alteradas externamente.
        # No entanto, _quando_modelo_mudar ainda não existe. Vamos preparar para quando existir.

        # if hasattr(self, '_quando_modelo_mudar'):
        # self._quando_modelo_mudar(self.modelo_combobox.currentText())
        # else:
        # self.log_message("Aviso: _quando_modelo_mudar ainda não implementado para ser chamado por atualizar_modelos_combobox.")

        # Por enquanto, vamos apenas logar a ação e o que foi selecionado.
        # A atualização da UI (preview, tabela) será tratada quando _quando_modelo_mudar for implementado
        # e a conexão do sinal currentTextChanged for feita no __init__.
        self.log_message(f"ComboBox de modelos atualizado. Selecionado: {self.modelo_combobox.currentText()}")

        # Habilita/desabilita botões de modificar/excluir conforme a seleção
        modelo_valido_selecionado = self.modelo_combobox.currentText() != utils.TEXTO_NENHUM_MODELO
        if hasattr(self, 'btn_modificar_modelo'):  # Verifica se os botões já foram criados no __init__
            self.btn_modificar_modelo.setEnabled(modelo_valido_selecionado)
        if hasattr(self, 'btn_excluir_modelo'):
            self.btn_excluir_modelo.setEnabled(modelo_valido_selecionado)

        # Chama explicitamente _quando_modelo_mudar para garantir que a UI
        # seja sempre sincronizada com o estado final do ComboBox.
        self.log_message(
            f"ComboBox de modelos atualizado. A sincronizar UI para: {self.modelo_combobox.currentText()}")
        self._quando_modelo_mudar(self.modelo_combobox.currentText())

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def _quando_modelo_mudar(self, nome_modelo_selecionado: str):
        """
        Chamado quando o texto do QComboBox de modelos muda.
        Este método coordena a atualização do preview do modelo e da
        estrutura da tabela de dados.

        Args:
            nome_modelo_selecionado: O nome do arquivo do modelo PSD selecionado.
        """
        self.log_message(f"Seleção de modelo alterada para: '{nome_modelo_selecionado}'")

        # Atualiza o preview do modelo
        # Esta chamada será descomentada/adicionada quando 'atualizar_preview_modelo' for implementado
        if hasattr(self, 'atualizar_preview_modelo'):
            self.atualizar_preview_modelo(nome_modelo_selecionado)
        else:
            self.log_message("Aviso: Método 'atualizar_preview_modelo' ainda não implementado.")

        # Atualiza a estrutura da tabela de dados com base no novo modelo
        # Esta chamada será descomentada/adicionada quando '_atualizar_tabela_para_modelo' for implementado
        if hasattr(self, '_atualizar_tabela_para_modelo'):
            self._atualizar_tabela_para_modelo(nome_modelo_selecionado)
        else:
            self.log_message("Aviso: Método '_atualizar_tabela_para_modelo' ainda não implementado.")

        # Habilita ou desabilita o botão "Gerar Cartões"
        # O botão só deve estar habilitado se um modelo válido estiver selecionado
        # e se este modelo tiver camadas configuradas (essa lógica mais fina
        # será feita em _atualizar_tabela_para_modelo).
        modelo_eh_valido = nome_modelo_selecionado != utils.TEXTO_NENHUM_MODELO

        if hasattr(self, 'btn_gerar_cartoes'):
            # A decisão final de habilitar btn_gerar_cartoes será feita em _atualizar_tabela_para_modelo,
            # pois depende das camadas configuradas. Aqui, apenas garantimos que ele esteja desabilitado
            # se nenhum modelo válido for selecionado.
            if not modelo_eh_valido:
                self.btn_gerar_cartoes.setEnabled(False)
            # Se for válido, _atualizar_tabela_para_modelo decidirá.

        # Habilita/desabilita botões de modificar/excluir (redundante com atualizar_modelos_combobox, mas seguro)
        if hasattr(self, 'btn_modificar_modelo'):
            self.btn_modificar_modelo.setEnabled(modelo_eh_valido)
        if hasattr(self, 'btn_excluir_modelo'):
            self.btn_excluir_modelo.setEnabled(modelo_eh_valido)

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def _atualizar_tabela_para_modelo(self, nome_modelo_psd: str):
        """
        Atualiza as colunas da tabela com base nos 'Dados Específicos' configurados
        para o modelo PSD selecionado. Também gerencia o estado dos botões.
        """
        self.data_table.setRowCount(0)

        # Desabilita os botões por padrão se nenhum modelo válido for selecionado
        if not nome_modelo_psd or nome_modelo_psd == utils.TEXTO_NENHUM_MODELO:
            self.table_headers = []
            self.data_table.setColumnCount(0)
            self.data_table.setHorizontalHeaderLabels([])
            self.btn_gerar_cartoes.setEnabled(False)
            self.btn_gerenciar_regras.setEnabled(False)
            self.log_message("Nenhum modelo selecionado. Tabela e botões de ação desabilitados.")
            self.data_table.setRowCount(15)
            return

        # Busca a configuração completa do modelo (que agora é um dicionário)
        config_modelo = self.configuracoes_modelos.get(nome_modelo_psd, {})
        # Pega a lista de 'Dados Específicos' de dentro da configuração
        dados_especificos_configurados = config_modelo.get("dados_especificos", [])

        if not dados_especificos_configurados:
            self.table_headers = []
            self.data_table.setColumnCount(1)
            self.data_table.setHorizontalHeaderLabels(["Status"])
            self.data_table.setRowCount(1)

            mensagem_item = QTableWidgetItem(
                "Modelo não configurado. Use 'Modificar Modelo' para definir os Dados Específicos."
            )
            mensagem_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            mensagem_item.setFlags(mensagem_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
            self.data_table.setItem(0, 0, mensagem_item)

            self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            self.data_table.verticalHeader().setVisible(False)

            self.btn_gerar_cartoes.setEnabled(False)
            # O botão de gerenciar regras deve ficar habilitado, pois o usuário pode querer
            # criar regras mesmo sem dados específicos (embora incomum).
            # Ou podemos desabilitar também para forçar a configuração primeiro. Vamos desabilitar.
            self.btn_gerenciar_regras.setEnabled(False)
            self.log_message(f"Modelo '{nome_modelo_psd}' não possui Dados Específicos configurados.")
        else:
            self.table_headers = dados_especificos_configurados
            self.data_table.setColumnCount(len(self.table_headers))
            self.data_table.setHorizontalHeaderLabels(self.table_headers)

            self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            self.data_table.verticalHeader().setVisible(True)
            self.data_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

            self.data_table.setRowCount(15)

            # Habilita os botões de ação, pois temos um modelo válido e configurado
            self.btn_gerar_cartoes.setEnabled(True)
            self.btn_gerenciar_regras.setEnabled(True)
            self.log_message(
                f"Tabela atualizada para o modelo '{nome_modelo_psd}' com os Dados Específicos: {self.table_headers}."
            )

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def atualizar_preview_modelo(self, nome_modelo_selecionado: str = None):
        """
        Atualiza a QLabel de preview. Carrega uma prévia .png existente se estiver
        atualizada, ou gera uma nova a partir do arquivo .psd usando Pillow.

        Args:
            nome_modelo_selecionado: O nome do arquivo do modelo PSD.
                                     Se None, usa o texto atual do combobox.
        """
        if nome_modelo_selecionado is None:
            nome_modelo_selecionado = self.modelo_combobox.currentText()

        # Limpa o estado do preview antes de tentar carregar um novo
        self._current_pixmap = None
        if hasattr(self, 'preview_label'): # Garante que preview_label existe
            self.preview_label.clear()
            self.preview_label.setText("A carregar pré-visualização...") # Mensagem de carregamento
            self.preview_label.setStyleSheet("border: 1px solid gray; border-radius: 5px; background-color: #e0e0e0; color: #333;")
            QApplication.processEvents() # Força a atualização da UI

        if not nome_modelo_selecionado or nome_modelo_selecionado == utils.TEXTO_NENHUM_MODELO:
            if hasattr(self, 'preview_label'):
                self.preview_label.setText("Nenhum modelo selecionado")
                self.preview_label.setStyleSheet("border: 1px solid gray; border-radius: 5px; background-color: #f0f0f0;")
            return

        caminho_psd = os.path.join(PASTA_PADRAO_MODELOS, nome_modelo_selecionado)
        nome_base_sem_ext = os.path.splitext(nome_modelo_selecionado)[0]
        caminho_preview_png = os.path.join(PASTA_PADRAO_MODELOS, f"{nome_base_sem_ext}_preview.png")

        if not os.path.exists(caminho_psd):
            if hasattr(self, 'preview_label'):
                self.preview_label.setText(f"ERRO: '{nome_modelo_selecionado}' não encontrado.")
                self.preview_label.setStyleSheet("border: 1px solid red; border-radius: 5px; background-color: #ffe0e0; color: red;")
            self.log_message(f"Erro de pré-visualização: Arquivo PSD '{caminho_psd}' não existe.")
            return

        try:
            # Tenta carregar a prévia .png existente se ela for mais nova ou igual ao PSD
            preview_existe_e_atualizado = (
                os.path.exists(caminho_preview_png) and
                os.path.getmtime(caminho_preview_png) >= os.path.getmtime(caminho_psd)
            )

            if preview_existe_e_atualizado:
                pixmap_temp = QPixmap(caminho_preview_png)
                if not pixmap_temp.isNull():
                    self._current_pixmap = pixmap_temp
                    self.log_message(f"Pré-visualização carregada de '{caminho_preview_png}'.")
                    # _display_current_pixmap será chamado no final
                else:
                    self.log_message(f"Aviso: Falha ao carregar QPixmap de '{caminho_preview_png}', mesmo o arquivo existindo.")
                    # Força a regeneração abaixo
                    preview_existe_e_atualizado = False # Para entrar no bloco de geração

            if not preview_existe_e_atualizado:
                self.log_message(f"A gerar nova pré-visualização para '{nome_modelo_selecionado}' a partir do PSD...")
                if hasattr(self, 'preview_label'):
                    self.preview_label.setText("A gerar pré-visualização...")
                QApplication.processEvents()

                with Image.open(caminho_psd) as pil_img:
                    # Converte para RGBA para garantir compatibilidade e transparência
                    if pil_img.mode != 'RGBA':
                        pil_img_convertida = pil_img.convert('RGBA')
                    else:
                        pil_img_convertida = pil_img

                    # Salva o preview em formato PNG
                    pil_img_convertida.save(caminho_preview_png, 'PNG')

                    # Converte para QImage para exibir na UI
                    # É importante usar os dados da imagem convertida (pil_img_convertida)
                    qimage = QImage(pil_img_convertida.tobytes("raw", "RGBA"),
                                    pil_img_convertida.width,
                                    pil_img_convertida.height,
                                    QImage.Format.Format_RGBA8888)
                    self._current_pixmap = QPixmap.fromImage(qimage)
                    self.log_message(f"Nova pré-visualização para '{nome_modelo_selecionado}' gerada e salva em '{caminho_preview_png}'.")

        except FileNotFoundError: # Deve ser pego pelo check os.path.exists(caminho_psd)
            if hasattr(self, 'preview_label'):
                self.preview_label.setText(f"ERRO: PSD '{nome_modelo_selecionado}' não encontrado durante geração.")
            self.log_message(f"ERRO FATAL de pré-visualização: '{caminho_psd}' não encontrado durante tentativa de abertura com Pillow.")
            return # Sai da função
        except Exception as e:
            if hasattr(self, 'preview_label'):
                self.preview_label.setText("Erro ao gerar pré-visualização")
                self.preview_label.setStyleSheet("border: 1px solid orange; border-radius: 5px; background-color: #fff0e0; color: orange;")
            self.log_message(f"ERRO ao gerar/carregar pré-visualização para '{nome_modelo_selecionado}': {e}")
            # Se falhou em gerar, tenta remover um possível arquivo .png corrompido
            if os.path.exists(caminho_preview_png):
                try:
                    os.remove(caminho_preview_png)
                    self.log_message(f"Arquivo de pré-visualização corrompido '{caminho_preview_png}' removido.")
                except OSError as err_remove:
                    self.log_message(f"Aviso: Não foi possível remover o arquivo de pré-visualização corrompido '{caminho_preview_png}'. Erro: {err_remove}")
            return # Sai da função se houve erro

        # Se chegou aqui, _current_pixmap deve ter sido definido (ou é None se tudo falhou antes)
        # A chamada para _display_current_pixmap será feita no próximo passo.
        if hasattr(self, '_display_current_pixmap'):
            self._display_current_pixmap()
        else:
            self.log_message("Aviso: Método '_display_current_pixmap' ainda não implementado para exibir a pré-visualização.")

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def _display_current_pixmap(self):
        """
        Exibe o QPixmap armazenado em self._current_pixmap na QLabel de preview,
        escalonando-o para caber no tamanho da label, mantendo a proporção.
        """
        if hasattr(self, 'preview_label'): # Verifica se preview_label existe
            if self._current_pixmap and not self._current_pixmap.isNull():
                # Escala o pixmap para caber na label, mantendo a proporção
                # e usando transformação suave para melhor qualidade visual.
                scaled_pixmap = self._current_pixmap.scaled(
                    self.preview_label.size(), # Escala para o tamanho atual da label
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation
                )
                self.preview_label.setPixmap(scaled_pixmap)
                # Remove qualquer texto ou estilo de erro/carregamento
                self.preview_label.setStyleSheet("border: 1px solid gray; border-radius: 5px;")
            else:
                # Se _current_pixmap for None ou inválido após as tentativas de carregamento/geração
                self.preview_label.clear()
                self.preview_label.setText("Pré-visualização indisponível")
                self.preview_label.setStyleSheet("border: 1px solid orange; border-radius: 5px; background-color: #fff0e0; color: orange;")
        else:
            self.log_message("Aviso: Tentativa de exibir pixmap, mas preview_label não foi encontrado.")

# ________________________________________________________________________________________________

    def _processar_dados_especificos_configurados(self, nome_arquivo_psd: str, lista_dados_configurados: list[str], regra_nome_arquivo: str):
        """
        Recebe a lista de 'Dados Específicos' e a regra de nome de arquivo,
        atualiza a configuração do modelo e salva o arquivo JSON.
        Preserva as 'Regras de Texto' existentes.
        """
        self.log_message(f"Processando configuração para '{nome_arquivo_psd}'...")
        self.log_message(f"  - Dados Específicos: {lista_dados_configurados}")
        self.log_message(f"  - Regra de Nome de Arquivo: '{regra_nome_arquivo}'")

        if not nome_arquivo_psd:
            self.log_message("ERRO: Tentativa de processar configuração para um nome de arquivo PSD vazio.")
            return

        # Pega a configuração atual do modelo. Se não existir, começa com um dicionário vazio.
        config_atual = self.configuracoes_modelos.get(nome_arquivo_psd, {})

        # Atualiza as chaves com os novos dados recebidos
        config_atual["dados_especificos"] = lista_dados_configurados
        config_atual["regra_nome_arquivo"] = regra_nome_arquivo

        # Garante que a chave 'regras_texto' exista, caso seja um modelo novo.
        if "regras_texto" not in config_atual:
            config_atual["regras_texto"] = {}

        # Coloca a configuração atualizada (ou nova) de volta no dicionário principal
        self.configuracoes_modelos[nome_arquivo_psd] = config_atual

        # Salva o dicionário de configurações inteiro no arquivo JSON
        if utils.salvar_configuracoes_camadas_modelos(self.configuracoes_modelos):
            self.log_message("Arquivo de configuração de modelos salvo com sucesso.")
        else:
            self.log_message("ERRO CRÍTICO ao salvar o arquivo de configuração de modelos.")
            QMessageBox.critical(self, "Erro de Salvamento",
                                 "Não foi possível salvar as configurações no arquivo JSON.")

        # Atualiza a UI se o modelo modificado for o que está selecionado
        if self.modelo_combobox.currentText() == nome_arquivo_psd:
            self._quando_modelo_mudar(nome_arquivo_psd)

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def adicionar_modelo(self):
        """
        Abre um diálogo para o usuário selecionar um novo arquivo PSD.
        Copia o arquivo para a pasta de modelos e, em seguida, abre o
        GerenciarRegrasDialog para que o usuário defina as camadas editáveis.
        """
        self.garantir_pasta_modelos()  # Garante que a pasta de modelos exista

        # Abre o diálogo para selecionar o arquivo PSD de origem
        caminho_origem_psd, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar Modelo PSD para Adicionar",
            "",  # Diretório inicial (pode ser utils.PASTA_PADRAO_MODELOS ou o último usado)
            "Arquivos Photoshop (*.psd);;Todos os Arquivos (*.*)"
        )

        if not caminho_origem_psd:
            self.log_message("Adição de modelo cancelada pelo usuário (nenhum arquivo selecionado).")
            return  # Usuário cancelou a seleção

        nome_base_arquivo = os.path.basename(caminho_origem_psd)
        caminho_destino_psd = os.path.join(PASTA_PADRAO_MODELOS, nome_base_arquivo)

        # Verifica se o modelo já existe na pasta de destino
        if os.path.exists(caminho_destino_psd):
            resposta = QMessageBox.question(self, "Substituir Modelo Existente?",
                                            f"O modelo '{nome_base_arquivo}' já existe na pasta de modelos.\n"
                                            "Deseja substituí-lo?\n\n"
                                            "Se escolher 'Não', a operação será cancelada.",
                                            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                            QMessageBox.StandardButton.No)
            if resposta == QMessageBox.StandardButton.No:
                self.log_message(f"Adição de '{nome_base_arquivo}' cancelada para não substituir o existente.")
                return

        # Tenta copiar o arquivo para a pasta de modelos
        try:
            shutil.copy2(caminho_origem_psd, caminho_destino_psd)
            self.log_message(f"Arquivo '{nome_base_arquivo}' copiado para '{caminho_destino_psd}'.")
        except Exception as e:
            self.log_message(f"ERRO CRÍTICO ao copiar '{caminho_origem_psd}' para '{caminho_destino_psd}': {e}")
            QMessageBox.critical(self, "Erro ao Copiar Arquivo",
                                 f"Não foi possível copiar o arquivo modelo para o destino.\n"
                                 f"Verifique as permissões e o espaço em disco.\nErro: {e}")
            return

        # Se a cópia foi bem-sucedida, abre o diálogo de configuração de camadas.
        # Para um novo modelo, não há camadas existentes para pré-preencher, então passamos uma lista vazia.
        self.log_message(f"A abrir diálogo de configuração de camadas para o novo modelo: '{nome_base_arquivo}'.")
        dialogo_config = GerenciarRegrasDialog(
            psd_filename=nome_base_arquivo,
            camadas_existentes=[],  # Novo modelo, sem camadas pré-definidas
            parent=self
        )

        # Conecta o sinal do diálogo ao nosso método que processa e salva as camadas
        # Usamos uma lambda para passar o nome_base_arquivo corretamente.
        # A conexão é feita aqui, e não globalmente, pois é específica para esta instância do diálogo.
        dialogo_config.configuracaoSalva.connect(
            lambda lista_dados, regra_nome: self._processar_dados_especificos_configurados(
                nome_base_arquivo, lista_dados, regra_nome
            )
        )

        # Executa o diálogo de configuração. O diálogo é modal.
        if dialogo_config.exec():
            # Usuário clicou em "Salvar" no diálogo.
            # O método _processar_camadas_configuradas já foi chamado através do sinal.
            self.log_message(f"Configuração de camadas para '{nome_base_arquivo}' foi definida pelo usuário.")
        else:
            # Usuário clicou em "Cancelar" ou fechou o diálogo.
            # Neste caso, o modelo foi copiado, mas não configurado.
            # Chamamos _processar_camadas_configuradas com uma lista vazia para
            # garantir que, se havia alguma configuração antiga (improvável para um novo, mas seguro),
            # ela seja limpa, e para que o sistema saiba que este modelo existe mas não tem camadas ativas.
            self.log_message(
                f"Configuração de camadas para '{nome_base_arquivo}' cancelada ou não definida pelo usuário.")
            self._processar_camadas_configuradas(nome_base_arquivo, [])

        # Atualiza a lista de modelos no ComboBox para incluir o novo modelo
        self.atualizar_modelos_combobox()
        # Seleciona o modelo recém-adicionado no ComboBox
        self.modelo_combobox.setCurrentText(nome_base_arquivo)

        # _quando_modelo_mudar será chamado automaticamente devido à mudança no currentText do combobox,
        # atualizando assim o preview e a tabela.

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def _excluir_logica_modelo(self, nome_modelo_a_excluir: str) -> bool:
        """
        Lógica interna para excluir os ficheiros de um modelo (PSD e preview)
        e a sua configuração do dicionário self.configuracoes_modelos.
        Também tenta salvar as configurações atualizadas no arquivo JSON.

        Args:
            nome_modelo_a_excluir: O nome do arquivo do modelo PSD a ser excluído.

        Returns:
            True se a exclusão (ficheiros e configuração em memória) for bem-sucedida
            e o salvamento da configuração JSON também. False caso contrário.
        """
        if not nome_modelo_a_excluir or nome_modelo_a_excluir == utils.TEXTO_NENHUM_MODELO:
            self.log_message("Tentativa de excluir um modelo inválido ou não selecionado.")
            return False

        self.log_message(f"A iniciar exclusão lógica para o modelo: '{nome_modelo_a_excluir}'.")

        caminho_psd = os.path.join(PASTA_PADRAO_MODELOS, nome_modelo_a_excluir)
        nome_base_sem_ext = os.path.splitext(nome_modelo_a_excluir)[0]
        caminho_preview_png = os.path.join(PASTA_PADRAO_MODELOS, f"{nome_base_sem_ext}_preview.png")

        arquivos_excluidos_com_sucesso = True # Assumimos sucesso até que algo falhe

        # Tenta excluir o arquivo PSD
        if os.path.exists(caminho_psd):
            try:
                # Antes de excluir, podemos usar os métodos de verificação de arquivo em uso, se desejado.
                # Por exemplo, _aguardar_liberacao_arquivo(caminho_psd)
                # if not self._aguardar_liberacao_arquivo(caminho_psd):
                #     self.log_message(f"ERRO: Arquivo PSD '{caminho_psd}' ainda em uso. Exclusão cancelada.")
                #     QMessageBox.warning(self, "Ficheiro em Uso", f"O ficheiro '{nome_modelo_a_excluir}' parece estar em uso e não pode ser excluído agora.")
                #     return False
                os.remove(caminho_psd)
                self.log_message(f"Arquivo PSD '{caminho_psd}' excluído.")
            except OSError as e:
                self.log_message(f"ERRO ao excluir o arquivo PSD '{caminho_psd}': {e}")
                QMessageBox.warning(self, "Erro ao Excluir Ficheiro",
                                    f"Não foi possível excluir o ficheiro PSD '{nome_modelo_a_excluir}'.\n"
                                    f"Verifique se não está aberto noutro programa.\nErro: {e}")
                arquivos_excluidos_com_sucesso = False # Marca falha na exclusão de arquivos

        # Tenta excluir o arquivo de preview .png
        if os.path.exists(caminho_preview_png):
            try:
                os.remove(caminho_preview_png)
                self.log_message(f"Arquivo de pré-visualização '{caminho_preview_png}' excluído.")
            except OSError as e:
                self.log_message(f"ERRO ao excluir o arquivo de pré-visualização '{caminho_preview_png}': {e}")
                # Não consideramos isto um erro crítico que impeça a remoção da configuração,
                # mas é bom avisar o utilizador se o botão de exclusão foi diretamente clicado.
                # Se chamado internamente, apenas o log pode ser suficiente.
                # QMessageBox.warning(self, "Erro ao Excluir Pré-visualização",
                #                     f"Não foi possível excluir o ficheiro de pré-visualização '{os.path.basename(caminho_preview_png)}'.\n"
                #                     f"Erro: {e}")
                # arquivos_excluidos_com_sucesso = False # Opcional: decidir se falha na preview impede tudo

        # Se a exclusão de algum arquivo físico falhou, podemos decidir parar aqui.
        # Por ora, vamos prosseguir para remover a configuração mesmo que um arquivo físico tenha falhado,
        # mas o retorno final indicará o sucesso geral.
        if not arquivos_excluidos_com_sucesso:
             self.log_message(f"Falha na exclusão de um ou mais ficheiros físicos para '{nome_modelo_a_excluir}'. A configuração ainda será removida da memória e do JSON.")
             # Poderia retornar False aqui se a política for não mexer no JSON se os arquivos não sumiram.

        # Remove a configuração do modelo do dicionário em memória
        configuracao_removida_memoria = False
        if nome_modelo_a_excluir in self.configuracoes_modelos:
            del self.configuracoes_modelos[nome_modelo_a_excluir]
            self.log_message(f"Configuração para '{nome_modelo_a_excluir}' removida da memória.")
            configuracao_removida_memoria = True
        else:
            self.log_message(f"Nenhuma configuração encontrada na memória para '{nome_modelo_a_excluir}' (pode já ter sido removida ou nunca existiu).")
            configuracao_removida_memoria = True # Consideramos sucesso se não havia nada para remover

        # Salva o dicionário de configurações atualizado no arquivo JSON
        json_salvo_com_sucesso = utils.salvar_configuracoes_camadas_modelos(self.configuracoes_modelos)
        if json_salvo_com_sucesso:
            self.log_message("Arquivo de configuração de camadas de modelos salvo com sucesso após exclusão.")
        else:
            self.log_message(f"ERRO CRÍTICO ao salvar o arquivo de configuração de camadas após tentar excluir '{nome_modelo_a_excluir}'.")
            QMessageBox.critical(self, "Erro de Salvamento",
                                 "Não foi possível salvar as atualizações no arquivo de configuração JSON após a exclusão.")
            # Mesmo que o JSON não salve, os arquivos podem ter sido excluídos e a config removida da memória.
            # O estado fica inconsistente.

        # Retorna True se os arquivos foram excluídos (ou não precisavam ser) E a config foi removida da memória E o JSON foi salvo.
        # Ajustar esta lógica conforme a política de erro desejada.
        # Por exemplo, se a exclusão do PSD é mandatória:
        # return arquivos_excluidos_com_sucesso and configuracao_removida_memoria and json_salvo_com_sucesso
        # Por ora, vamos ser um pouco mais permissivos: se a config em memória foi atualizada e o JSON salvou,
        # consideramos a "operação lógica" um sucesso, mesmo que um arquivo tenha falhado em ser excluído.
        # O método que chama pode usar o retorno para decidir se atualiza a UI.
        return configuracao_removida_memoria and json_salvo_com_sucesso

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def excluir_modelo(self):
        """
        Slot para o botão 'Excluir Modelo'. Obtém o modelo selecionado,
        pede confirmação ao utilizador e, se confirmado, chama a lógica
        interna de exclusão (_excluir_logica_modelo).
        Atualiza a interface gráfica (ComboBox de modelos) após a exclusão.
        """
        modelo_selecionado = self.modelo_combobox.currentText()

        if not modelo_selecionado or modelo_selecionado == utils.TEXTO_NENHUM_MODELO:
            QMessageBox.information(self, "Excluir Modelo",
                                    "Nenhum modelo selecionado para excluir.")
            return

        # Pergunta de confirmação mais detalhada
        confirmacao = QMessageBox.question(self, "Confirmar Exclusão",
                                           f"Tem a certeza que deseja excluir permanentemente o modelo '{modelo_selecionado}'?\n\n"
                                           "Isto irá remover o ficheiro PSD, a sua pré-visualização e todas as configurações de camadas associadas.\n"
                                           "Esta ação não pode ser desfeita.",
                                           QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                           QMessageBox.StandardButton.No)  # Botão 'Não' como padrão

        if confirmacao == QMessageBox.StandardButton.Yes:
            self.log_message(f"Utilizador confirmou a exclusão do modelo: '{modelo_selecionado}'.")

            # Chama a lógica interna de exclusão
            sucesso_logico = self._excluir_logica_modelo(modelo_selecionado)

            if sucesso_logico:
                self.log_message(
                    f"Exclusão lógica do modelo '{modelo_selecionado}' concluída com sucesso (configuração e JSON). A atualizar UI.")
                # Se a exclusão lógica (config e JSON) foi bem-sucedida, atualiza o ComboBox.
                # A atualização do ComboBox irá, por sua vez, chamar _quando_modelo_mudar,
                # que atualizará o preview e a tabela para o novo estado (provavelmente "nenhum modelo" ou o próximo da lista).
                self.atualizar_modelos_combobox()
                QMessageBox.information(self, "Modelo Excluído",
                                        f"O modelo '{modelo_selecionado}' foi excluído com sucesso.")
            else:
                # _excluir_logica_modelo já deve ter mostrado um QMessageBox.critical/warning
                # sobre a falha na exclusão de ficheiros ou no salvamento do JSON.
                self.log_message(
                    f"Falha na operação de exclusão lógica para '{modelo_selecionado}'. A UI pode não ter sido totalmente atualizada.")
                # Mesmo com falha, tentamos atualizar o combobox para refletir o estado mais próximo do real.
                self.atualizar_modelos_combobox()
        else:
            self.log_message(f"Exclusão do modelo '{modelo_selecionado}' cancelada pelo utilizador.")

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def modificar_modelo(self):
        """
        Inicia o fluxo para modificar um modelo existente usando um diálogo personalizado.
        """
        modelo_selecionado = self.modelo_combobox.currentText()

        if not modelo_selecionado or modelo_selecionado == utils.TEXTO_NENHUM_MODELO:
            QMessageBox.information(self, "Modificar Modelo", "Nenhum modelo selecionado para modificar.")
            return

        self.log_message(f"A abrir diálogo de decisão de modificação para: '{modelo_selecionado}'.")

        # Cria e executa o nosso novo diálogo personalizado
        dialogo = CaixaModificarModeloDialog(modelo_selecionado, self)

        # O diálogo só continua se o utilizador clicar num dos botões de ação (não no 'Cancelar')
        if dialogo.exec():
            # Verifica a escolha guardada no diálogo
            if dialogo.escolha == 'camadas':
                self.log_message(
                    f"Opção escolhida: Apenas alterar configuração de camadas para '{modelo_selecionado}'.")
                self._handle_alterar_apenas_camadas(modelo_selecionado)
            elif dialogo.escolha == 'arquivo':
                self.log_message(
                    f"Opção escolhida: Substituir ficheiro PSD e reconfigurar para '{modelo_selecionado}'.")
                self._handle_substituir_psd_e_reconfigurar(modelo_selecionado)
        else:
            # O utilizador clicou em 'Cancelar' ou fechou a janela
            self.log_message("Modificação de modelo cancelada pelo utilizador.")
# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def _handle_alterar_apenas_camadas(self, nome_modelo_selecionado: str):
        """
        Lida com a alteração apenas da configuração de um modelo existente.
        """
        self.log_message(f"A iniciar modificação de configuração para o modelo: '{nome_modelo_selecionado}'.")

        # Busca a configuração completa do modelo, incluindo a regra de nome de arquivo
        config_modelo = self.configuracoes_modelos.get(nome_modelo_selecionado, {})
        dados_especificos_atuais = config_modelo.get("dados_especificos", [])
        regra_nome_arquivo_atual = config_modelo.get("regra_nome_arquivo", "")  # Pega a regra

        self.log_message(f"Dados Específicos atuais: {dados_especificos_atuais}")
        self.log_message(f"Regra de Nome de Arquivo atual: '{regra_nome_arquivo_atual}'")

        # Cria o diálogo passando TODOS os dados existentes
        dialogo_config = GerenciarRegrasDialog(
            psd_filename=nome_modelo_selecionado,
            camadas_existentes=dados_especificos_atuais,
            regra_nome_arquivo_existente=regra_nome_arquivo_atual,  # Passa a regra para o diálogo
            parent=self
        )

        # Conecta o sinal para receber os DOIS argumentos (lista e string)
        dialogo_config.configuracaoSalva.connect(
            lambda novos_dados, nova_regra: self._processar_dados_especificos_configurados(
                nome_modelo_selecionado, novos_dados, nova_regra
            )
        )

        if dialogo_config.exec():
            self.log_message(f"Configuração para '{nome_modelo_selecionado}' foi salva pelo utilizador.")
        else:
            self.log_message(f"Modificação de configuração para '{nome_modelo_selecionado}' cancelada.")

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def _handle_substituir_psd_e_reconfigurar(self, nome_modelo_existente: str):
        """
        Lida com a Opção B do 'modificar_modelo':
        1. Guarda as camadas do modelo existente como sugestão.
        2. Pede ao utilizador para selecionar um novo ficheiro PSD.
        3. Exclui o modelo antigo (ficheiros e configuração).
        4. Copia o novo ficheiro PSD para a pasta de modelos.
        5. Abre o GerenciarRegrasDialog para o novo ficheiro, pré-preenchido com as camadas sugeridas.
        6. Processa a configuração salva.
        7. Atualiza a UI.

        Args:
            nome_modelo_existente: O nome do ficheiro do modelo PSD a ser substituído.
        """
        self.log_message(f"A iniciar substituição do FICHEIRO PSD e reconfiguração para o modelo: '{nome_modelo_existente}'.")

        # 1. Guarda a configuração de camadas atual do modelo existente como sugestão
        camadas_sugeridas = self.configuracoes_modelos.get(nome_modelo_existente, [])
        self.log_message(f"Camadas sugeridas (do modelo antigo '{nome_modelo_existente}'): {camadas_sugeridas}")

        # 2. Pede ao utilizador para selecionar o novo ficheiro PSD
        caminho_novo_psd_origem, _ = QFileDialog.getOpenFileName(
            self,
            f"Selecione o NOVO ficheiro PSD para substituir '{nome_modelo_existente}'",
            "",  # Diretório inicial
            "Ficheiros Photoshop (*.psd);;Todos os Ficheiros (*.*)"
        )

        if not caminho_novo_psd_origem:
            self.log_message("Seleção de novo ficheiro PSD cancelada. Operação de substituição abortada.")
            return # Utilizador cancelou

        nome_base_novo_psd = os.path.basename(caminho_novo_psd_origem)
        caminho_destino_novo_psd = os.path.join(PASTA_PADRAO_MODELOS, nome_base_novo_psd)

        # Verifica se o novo ficheiro escolhido resultaria num conflito de nome
        # (se o novo nome é diferente do antigo, mas já existe na pasta de modelos)
        if nome_base_novo_psd != nome_modelo_existente and os.path.exists(caminho_destino_novo_psd):
            resposta_conflito = QMessageBox.question(self, "Conflito de Nome",
                                                 f"Já existe um modelo com o nome '{nome_base_novo_psd}' na pasta de modelos.\n"
                                                 "Deseja substituí-lo (e ao ficheiro PSD associado)?",
                                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                                 QMessageBox.StandardButton.No)
            if resposta_conflito == QMessageBox.StandardButton.No:
                self.log_message(f"Operação de substituição abortada devido a conflito de nome com '{nome_base_novo_psd}' e recusa em substituir.")
                return

        # 3. Exclui o modelo antigo (ficheiros e configuração)
        # _excluir_logica_modelo retorna True em sucesso (lógico), False em falha.
        self.log_message(f"A tentar excluir o modelo antigo '{nome_modelo_existente}' antes de adicionar o novo.")
        if not self._excluir_logica_modelo(nome_modelo_existente):
            # _excluir_logica_modelo já deve ter mostrado uma mensagem de erro.
            self.log_message(f"Falha ao excluir o modelo antigo '{nome_modelo_existente}'. Operação de substituição abortada.")
            # É importante atualizar o combobox aqui, pois o modelo antigo pode ter sido parcialmente removido
            # ou a sua configuração pode estar inconsistente.
            self.atualizar_modelos_combobox()
            return

        self.log_message(f"Modelo antigo '{nome_modelo_existente}' excluído com sucesso (ou já não existia logicamente).")

        # 4. Copia o novo ficheiro PSD para a pasta de modelos
        try:
            shutil.copy2(caminho_novo_psd_origem, caminho_destino_novo_psd)
            self.log_message(f"Novo ficheiro PSD '{nome_base_novo_psd}' copiado para '{caminho_destino_novo_psd}'.")
        except Exception as e:
            self.log_message(f"ERRO CRÍTICO ao copiar o novo ficheiro PSD '{caminho_novo_psd_origem}' para o destino: {e}")
            QMessageBox.critical(self, "Erro ao Copiar Novo Ficheiro",
                                 f"Não foi possível copiar o novo ficheiro PSD.\nErro: {e}")
            # Mesmo que a cópia falhe, o modelo antigo foi excluído. Atualiza a UI.
            self.atualizar_modelos_combobox()
            return

        # 5. Abre o GerenciarRegrasDialog para o novo ficheiro, com as camadas antigas como sugestão
        self.log_message(f"A abrir diálogo de configuração de camadas para o novo ficheiro PSD: '{nome_base_novo_psd}'.")
        dialogo_config_novo = GerenciarRegrasDialog(
            psd_filename=nome_base_novo_psd,
            camadas_existentes=camadas_sugeridas, # Usa as camadas do modelo antigo como sugestão
            parent=self
        )
        dialogo_config_novo.configuracaoSalva.connect(
            lambda novos_dados, nova_regra: self._processar_dados_especificos_configurados(
                nome_base_novo_psd, novos_dados, nova_regra
            )
        )

        # 6. Processa o resultado do diálogo
        if dialogo_config_novo.exec():
            self.log_message(f"Configuração de camadas para o novo ficheiro '{nome_base_novo_psd}' foi salva pelo utilizador.")
        else:
            self.log_message(f"Configuração de camadas para '{nome_base_novo_psd}' cancelada ou não definida. O ficheiro PSD foi adicionado sem camadas ativas.")
            self._processar_camadas_configuradas(nome_base_novo_psd, []) # Salva uma config vazia

        # 7. Atualiza a UI: ComboBox e seleciona o novo modelo
        self.log_message("A atualizar ComboBox de modelos e a selecionar o novo modelo.")
        self.atualizar_modelos_combobox()
        self.modelo_combobox.setCurrentText(nome_base_novo_psd)
        # A mudança no currentText do combobox deve acionar _quando_modelo_mudar, atualizando preview e tabela.

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def add_table_row(self):
        """
        Adiciona uma nova linha vazia ao final da tabela de dados.
        A linha terá o número de colunas atualmente definido na tabela.
        """
        if not hasattr(self, 'data_table'):
            self.log_message("ERRO: Tentativa de adicionar linha, mas data_table não existe.")
            return

        # Verifica se há colunas definidas na tabela.
        # Não adiciona linha se a tabela não estiver configurada (ex: nenhum modelo selecionado ou configurado)
        if self.data_table.columnCount() == 0:
            self.log_message(
                "Aviso: Tentativa de adicionar linha, mas a tabela não tem colunas definidas (nenhum modelo configurado?).")
            QMessageBox.information(self, "Adicionar Linha",
                                    "Não é possível adicionar uma linha pois a tabela não está configurada.\n"
                                    "Selecione um modelo e configure as suas camadas primeiro.")
            return

        current_row_count = self.data_table.rowCount()
        self.data_table.insertRow(current_row_count)
        self.log_message(f"Nova linha adicionada à tabela. Total de linhas: {self.data_table.rowCount()}")

        # Opcional: rolar para a nova linha se a tabela tiver muitas linhas
        # self.data_table.scrollToBottom()

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def remove_selected_table_rows(self):
        """
        Remove todas as linhas que estão atualmente selecionadas na tabela de dados.
        Pede confirmação ao utilizador antes de remover.
        """
        if not hasattr(self, 'data_table'):
            self.log_message("ERRO: Tentativa de remover linhas, mas data_table não existe.")
            return

        # Obtém os índices únicos das linhas selecionadas
        # O utilizador pode selecionar múltiplas células na mesma linha,
        # então usamos set() para obter apenas os índices de linha únicos.
        selected_indexes = self.data_table.selectedIndexes()
        if not selected_indexes:
            QMessageBox.information(self, "Remover Linhas", "Nenhuma linha selecionada para remover.")
            return

        # Extrai os índices das linhas e ordena-os em ordem decrescente
        # para evitar problemas ao remover múltiplas linhas (remover de baixo para cima).
        unique_selected_rows = sorted(list(set(index.row() for index in selected_indexes)), reverse=True)

        if not unique_selected_rows:  # Segurança extra, embora selected_indexes já deva cobrir
            QMessageBox.information(self, "Remover Linhas", "Nenhuma linha efetivamente selecionada para remover.")
            return

        num_linhas_a_remover = len(unique_selected_rows)
        plural_s = "s" if num_linhas_a_remover > 1 else ""

        confirmacao = QMessageBox.question(self, "Confirmar Remoção",
                                           f"Tem a certeza que deseja remover a(s) {num_linhas_a_remover} linha(s) selecionada(s)?",
                                           QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                           QMessageBox.StandardButton.No)

        if confirmacao == QMessageBox.StandardButton.Yes:
            for row_index in unique_selected_rows:
                self.data_table.removeRow(row_index)
            self.log_message(f"{num_linhas_a_remover} linha(s) removida(s) da tabela.")
        else:
            self.log_message("Remoção de linha(s) cancelada pelo utilizador.")

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def clear_table(self):
        """
        Remove todas as linhas e dados da tabela, redefinindo-a para um
        estado com um número padrão de linhas vazias.
        Pede confirmação ao utilizador antes de limpar.
        """
        if not hasattr(self, 'data_table'):
            self.log_message("ERRO: Tentativa de limpar tabela, mas data_table não existe.")
            return

        if self.data_table.rowCount() == 0 and self.data_table.columnCount() == 0:
            QMessageBox.information(self, "Limpar Tabela", "A tabela já está vazia ou não configurada.")
            return

        if self.data_table.rowCount() == 0:
            # Se não há linhas, mas há colunas, significa que está vazia mas configurada.
            # Podemos permitir limpar para resetar para o número padrão de linhas.
            pass  # Permite prosseguir para a confirmação.

        confirmacao = QMessageBox.question(self, "Confirmar Limpeza",
                                           "Tem a certeza que deseja limpar todos os dados da tabela?\n"
                                           "Todas as informações inseridas nas linhas serão perdidas.",
                                           QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                           QMessageBox.StandardButton.No)

        if confirmacao == QMessageBox.StandardButton.Yes:
            # Remove todas as linhas existentes
            self.data_table.setRowCount(0)

            # Readiciona um número padrão de linhas vazias,
            # mas apenas se a tabela tiver colunas (ou seja, estiver configurada por um modelo)
            if self.data_table.columnCount() > 0:
                self.data_table.setRowCount(
                    15)  # Ou o seu número padrão de linhas definido no __init__ ou _atualizar_tabela_para_modelo
                self.log_message("Tabela limpa. Todas as linhas removidas e 15 linhas vazias adicionadas.")
            else:
                # Se não há colunas, a tabela permanece com 0 linhas (estado não configurado)
                self.log_message("Tabela limpa. Todas as linhas removidas. Tabela não está configurada com colunas.")

            QMessageBox.information(self, "Tabela Limpa", "Todos os dados da tabela foram removidos.")
        else:
            self.log_message("Limpeza da tabela cancelada pelo utilizador.")

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def selecionar_pasta_saida(self):
        """
        Abre um diálogo QFileDialog para o utilizador selecionar uma pasta
        onde os cartões gerados serão guardados.
        Atualiza o atributo self.output_dir e a label na interface.
        Guarda a última pasta selecionada para uso futuro.
        """
        if not hasattr(self, 'output_dir'):
            # Fallback caso self.output_dir não tenha sido inicializado (improvável)
            diretorio_inicial = os.path.expanduser("~")  # Pasta home do utilizador
        else:
            diretorio_inicial = self.output_dir

        nova_pasta_saida = QFileDialog.getExistingDirectory(
            self,
            "Escolha a Pasta de Saída para os Cartões Gerados",
            diretorio_inicial
        )

        if nova_pasta_saida:  # Se o utilizador selecionou uma pasta e não cancelou
            self.output_dir = os.path.abspath(nova_pasta_saida)  # Garante caminho absoluto
            if hasattr(self, 'saida_dir_label_path'):
                self.saida_dir_label_path.setText(self.output_dir)
                self.saida_dir_label_path.setToolTip(self.output_dir)  # Atualiza tooltip para caminhos longos

            # Guarda a última pasta selecionada usando a função de utils.py
            utils.save_last_output_dir(self.output_dir)
            self.log_message(f"Pasta de saída definida para: {self.output_dir}")
        else:
            self.log_message("Seleção de pasta de saída cancelada pelo utilizador.")

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def gerar_cartoes(self):
        """
        Orquestra a geração dos cartões. Versão com a lógica de regras corrigida
        para priorizar a aplicação da regra sobre o valor literal da tabela.
        """
        self.log_message("A iniciar processo de geração de cartões...")
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

        # --- Validações Iniciais (sem alterações) ---
        modelo_selecionado_nome = self.modelo_combobox.currentText()
        if not modelo_selecionado_nome or modelo_selecionado_nome == utils.TEXTO_NENHUM_MODELO:
            QMessageBox.warning(self, "Geração de Cartões", "Por favor, selecione um modelo PSD primeiro.")
            QApplication.restoreOverrideCursor()
            return

        caminho_psd_modelo = os.path.abspath(os.path.join(PASTA_PADRAO_MODELOS, modelo_selecionado_nome))
        if not os.path.exists(caminho_psd_modelo):
            QMessageBox.critical(self, "Erro de Arquivo",
                                 f"O ficheiro modelo PSD '{modelo_selecionado_nome}' não foi encontrado.")
            QApplication.restoreOverrideCursor()
            return

        config_modelo = self.configuracoes_modelos.get(modelo_selecionado_nome, {})
        regras_texto = config_modelo.get("regras_texto", {})
        dados_especificos_configurados = config_modelo.get("dados_especificos", [])
        if not dados_especificos_configurados:
            QMessageBox.warning(self, "Geração de Cartões",
                                f"O modelo '{modelo_selecionado_nome}' não possui Dados Específicos configurados.")
            QApplication.restoreOverrideCursor()
            return

        if not self.output_dir or not os.path.exists(self.output_dir):
            self.selecionar_pasta_saida()
            if not self.output_dir or not os.path.exists(self.output_dir):
                QApplication.restoreOverrideCursor()
                return

        # --- Recolha de Dados da Tabela (sem alterações) ---
        dados_para_geracao = []
        for num_linha in range(self.data_table.rowCount()):
            dados_linha_atual = {}
            linha_contem_dados = False
            for num_coluna, nome_header in enumerate(self.table_headers):
                item_tabela = self.data_table.item(num_linha, num_coluna)
                valor_celula = item_tabela.text().strip() if item_tabela else ""
                dados_linha_atual[nome_header] = valor_celula
                if valor_celula:
                    linha_contem_dados = True
            if linha_contem_dados:
                dados_para_geracao.append(dados_linha_atual)

        if not dados_para_geracao:
            QMessageBox.information(self, "Geração de Cartões",
                                    "A tabela está vazia ou não contém dados para processar.")
            QApplication.restoreOverrideCursor()
            return

        self.log_message(f"Encontrados {len(dados_para_geracao)} cartões para gerar.")

        # --- Interação com Photoshop (LÓGICA DE REGRAS CORRIGIDA) ---
        ps_app = None
        doc_modelo = None
        cartoes_gerados_count = 0
        erros_na_geracao = 0

        try:
            self.log_message("A tentar conectar-se ao Photoshop...")
            ps_app = win32com.client.Dispatch("Photoshop.Application")
            ps_app.Visible = False
            opcoes_exportacao = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
            opcoes_exportacao.Format = 13  # PNG
            opcoes_exportacao.PNG8 = False
            doc_modelo = ps_app.Open(caminho_psd_modelo)

            for i, dados_cartao_atual in enumerate(dados_para_geracao):
                self.log_message(f"A processar cartão {i + 1}/{len(dados_para_geracao)}...")
                QApplication.processEvents()

                campos_finais_para_psd = {}
                # Copiamos os dados da linha para um dicionário temporário que podemos modificar
                dados_a_processar = dados_cartao_atual.copy()

                # 1. PROCESSAR AS REGRAS PRIMEIRO
                if regras_texto:
                    for camada_alvo, regra in regras_texto.items():
                        # Encontra todos os placeholders, ex: {nome}, {conjuge}
                        placeholders = re.findall(r'\{([^{}]+)\}', regra)

                        rule_has_valid_data = False
                        if not placeholders:
                            rule_has_valid_data = True
                        else:
                            # Verifica se pelo menos um placeholder pode ser preenchido com dados reais
                            for placeholder in placeholders:
                                if placeholder in dados_a_processar and dados_a_processar[placeholder]:
                                    rule_has_valid_data = True
                                    break

                        if rule_has_valid_data:
                            # Se a regra é válida, formata o texto e o define como o valor final para esta camada
                            texto_final = regra.format_map(defaultdict(str, dados_a_processar))
                            campos_finais_para_psd[camada_alvo] = texto_final
                            # Remove o campo do dicionário temporário para não ser adicionado novamente
                            if camada_alvo in dados_a_processar:
                                del dados_a_processar[camada_alvo]

                # 2. ADICIONAR OS DADOS RESTANTES (que não tinham regras ou cujas regras eram inválidas)
                for campo_restante, valor_restante in dados_a_processar.items():
                    campos_finais_para_psd[campo_restante] = valor_restante

                # Define o nome do ficheiro de saída
                nome_base_ficheiro = dados_cartao_atual.get("nome", f"cartao_{i + 1}").replace(" ", "_")
                caminho_saida_completo = os.path.join(self.output_dir, f"{nome_base_ficheiro}.png")

                try:
                    ps_utils.gerar_cartao_photoshop(
                        psApp=ps_app,
                        doc=doc_modelo,
                        output_path=caminho_saida_completo,
                        campos=campos_finais_para_psd,
                        export_options_obj=opcoes_exportacao
                    )
                    self.log_message(f"Cartão salvo com sucesso: {caminho_saida_completo}")
                    cartoes_gerados_count += 1
                except Exception as e_ps_util:
                    self.log_message(f"ERRO ao gerar cartão para dados: {dados_cartao_atual}. Erro: {e_ps_util}")
                    erros_na_geracao += 1

        except Exception as e_geral:
            self.log_message(f"ERRO GERAL durante a geração de cartões: {e_geral}")
            QMessageBox.critical(self, "Erro na Geração", f"Ocorreu um erro inesperado:\n{e_geral}")
        finally:
            if doc_modelo is not None:
                doc_modelo.Close(2)
            if ps_app is not None:
                ps_app = None
                gc.collect()

            QApplication.restoreOverrideCursor()
            self.log_message("Processo de geração de cartões finalizado.")

            if cartoes_gerados_count > 0:
                mensagem_final = f"{cartoes_gerados_count} cartões gerados com sucesso."
                if erros_na_geracao > 0:
                    mensagem_final += f"\n\nOcorreram {erros_na_geracao} erros. Verifique o log para mais detalhes."
                QMessageBox.information(self, "Geração Concluída", mensagem_final)
            elif erros_na_geracao > 0:
                QMessageBox.warning(self, "Geração Falhou",
                                    f"Nenhum cartão foi gerado com sucesso. Ocorreram {erros_na_geracao} erros.")

# ________________________________________________________________________________________________

    def _sanitizar_nome_arquivo(self, nome: str) -> str:
        """Remove caracteres inválidos para nomes de arquivo."""
        # Remove caracteres inválidos no Windows e outros sistemas
        caracteres_invalidos = r'[\\/:"*?<>|]'
        nome_limpo = re.sub(caracteres_invalidos, '', nome)
        # Substitui múltiplos espaços por um único underscore
        nome_limpo = re.sub(r'\s+', '_', nome_limpo)
        return nome_limpo

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def _forcar_fechamento_photoshop(self):
        """
        Tenta forçar o fechamento de instâncias do Photoshop.
        Primeiro, tenta fechar graciosamente através de COM.
        Depois, procura processos do Photoshop e tenta encerrá-los.
        Este é um método mais "agressivo" e deve ser usado com cautela.
        """
        self.log_message("A tentar forçar o fechamento de instâncias do Photoshop...")

        # Tenta fechar via COM primeiro (mais gracioso)
        try:
            # Verifica se temos uma instância COM ativa na nossa aplicação
            # (embora em gerar_cartoes já limpemos ps_app e doc_modelo no finally)
            if hasattr(self, 'psApp') and self.psApp is not None:
                try:
                    if hasattr(self, 'doc_modelo') and self.doc_modelo is not None:
                        self.doc_modelo.Close(2)  # psDoNotSaveChanges
                        self.doc_modelo = None
                    self.psApp.Quit()
                    self.psApp = None
                    self.log_message("Instância COM do Photoshop (self.psApp) fechada.")
                except Exception as e_com_self:
                    self.log_message(f"Aviso: Erro ao fechar self.psApp via COM: {e_com_self}")

            # Tenta obter uma instância ativa do Photoshop no sistema e fechá-la
            try:
                ps_ativo = win32com.client.GetActiveObject("Photoshop.Application")
                if ps_ativo:
                    self.log_message("Instância ativa do Photoshop encontrada no sistema.")
                    # Fecha todos os documentos abertos sem salvar
                    while ps_ativo.Documents.Count > 0:
                        ps_ativo.Documents.Item(1).Close(2)  # psDoNotSaveChanges (índice é 1-based)
                    ps_ativo.Quit()
                    self.log_message("Instância ativa do Photoshop no sistema foi instruída a fechar.")
                    # Aguarda um pouco para o Photoshop processar o Quit
                    time.sleep(2)
            except Exception as e_com_getactive:
                self.log_message(
                    f"Nenhuma instância ativa do Photoshop encontrada via GetActiveObject ou erro ao tentar fechar: {e_com_getactive}")

        except Exception as e_com_geral:
            self.log_message(f"Erro durante a tentativa de fechamento via COM: {e_com_geral}")

        # Agora, procura processos do Photoshop e tenta encerrá-los
        # Isto é mais arriscado pois pode levar à perda de trabalho não salvo
        # se o utilizador tiver o Photoshop aberto manualmente com trabalho em progresso.
        processos_photoshop_encontrados = []
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and 'photoshop' in proc.info['name'].lower():
                    processos_photoshop_encontrados.append(proc)
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue

        if processos_photoshop_encontrados:
            self.log_message(f"Encontrados {len(processos_photoshop_encontrados)} processo(s) com 'photoshop' no nome.")
            for proc_ps in processos_photoshop_encontrados:
                try:
                    self.log_message(
                        f"A tentar terminar o processo do Photoshop: PID {proc_ps.pid}, Nome: {proc_ps.name()}")
                    # Poderia pedir confirmação ao utilizador aqui antes de terminar processos
                    # resp_term = QMessageBox.warning(self, "Terminar Processo?",
                    #                                 f"O processo '{proc_ps.name()}' (PID: {proc_ps.pid}) parece ser do Photoshop.\n"
                    #                                 "Deseja tentar terminá-lo? Isto pode causar perda de trabalho não salvo.",
                    #                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
                    # if resp_term == QMessageBox.StandardButton.Yes:
                    #     proc_ps.terminate() # Tenta terminar graciosamente primeiro
                    #     proc_ps.wait(timeout=3) # Espera um pouco
                    #     if proc_ps.is_running():
                    #         proc_ps.kill() # Força o encerramento
                    #     self.log_message(f"Processo PID {proc_ps.pid} instruído a terminar.")
                    # else:
                    #     self.log_message(f"Término do processo PID {proc_ps.pid} cancelado pelo utilizador (ou lógica).")

                    # Por agora, vamos apenas logar, sem terminar automaticamente para segurança.
                    # A terminação de processos deve ser feita com muito cuidado.
                    self.log_message(
                        f"AVISO: Processo do Photoshop encontrado (PID {proc_ps.pid}). A terminação manual pode ser necessária se houver bloqueios.")

                except psutil.NoSuchProcess:
                    self.log_message(f"Processo PID {proc_ps.pid} já não existia ao tentar terminar.")
                except Exception as e_proc:
                    self.log_message(f"Erro ao tentar interagir com o processo PID {proc_ps.pid}: {e_proc}")
        else:
            self.log_message("Nenhum processo com 'photoshop' no nome encontrado em execução (via psutil).")

        gc.collect()  # Força a recolha de lixo
        self.log_message("Tentativa de forçar fechamento do Photoshop concluída.")

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def _arquivo_esta_em_uso(self, caminho_arquivo: str) -> bool:
        """
        Verifica se um ficheiro está aparentemente em uso por outro processo.
        Tenta abrir o ficheiro em modo de escrita exclusiva. Se falhar,
        assume que está em uso.

        Args:
            caminho_arquivo: O caminho completo para o ficheiro a ser verificado.

        Returns:
            True se o ficheiro parecer estar em uso, False caso contrário
            (ou se o ficheiro não existir).
        """
        if not os.path.exists(caminho_arquivo):
            # Se o ficheiro não existe, não pode estar em uso.
            return False

        try:
            # Tenta abrir o ficheiro em modo de leitura e escrita binária.
            # Se o ficheiro estiver bloqueado por outro processo, isto deve levantar uma exceção.
            # O modo 'r+b' tenta abrir para leitura e escrita.
            # Poderíamos usar 'a' (append) que é menos intrusivo, mas 'r+b' é um bom teste.
            with open(caminho_arquivo, 'r+b') as f:
                # Se conseguimos abrir, não está (obviamente) bloqueado de forma exclusiva.
                # No entanto, esta verificação não é 100% garantida em todos os OS
                # e para todos os tipos de bloqueios, mas é uma boa heurística.
                pass # Apenas abrimos e fechamos.
            return False # Conseguiu abrir, então não está em uso exclusivo.
        except (IOError, OSError, PermissionError) as e:
            # Comuns exceções se o ficheiro estiver bloqueado (ex: PermissionError no Windows)
            self.log_message(f"Aviso: Ficheiro '{caminho_arquivo}' parece estar em uso. Exceção: {e}")
            return True # Falhou ao abrir, assume que está em uso.
        except Exception as e_inesperada:
            # Captura outras exceções inesperadas durante a tentativa de abertura.
            self.log_message(f"Aviso: Exceção inesperada ao verificar se ficheiro '{caminho_arquivo}' está em uso: {e_inesperada}")
            return True # Por segurança, assume que está em uso.

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def _aguardar_liberacao_arquivo(self, caminho_arquivo: str, max_tentativas: int = 10, delay_segundos: float = 0.5) -> bool:
        """
        Aguarda um ficheiro ser libertado por outros processos, verificando
        repetidamente se o ficheiro ainda está em uso.

        Args:
            caminho_arquivo: O caminho completo para o ficheiro a ser verificado.
            max_tentativas: O número máximo de vezes que a verificação será feita.
            delay_segundos: O tempo em segundos a aguardar entre cada tentativa.

        Returns:
            True se o ficheiro for libertado dentro do número de tentativas,
            False caso contrário (ou se o ficheiro não existir inicialmente).
        """
        if not os.path.exists(caminho_arquivo):
            self.log_message(f"Aviso em _aguardar_liberacao_arquivo: Ficheiro '{caminho_arquivo}' não existe, portanto não está bloqueado.")
            return True # Se não existe, considera-se "libertado" para operações como escrita.

        self.log_message(f"A aguardar libertação do ficheiro: '{os.path.basename(caminho_arquivo)}'. Máximo de {max_tentativas} tentativas.")

        for tentativa in range(max_tentativas):
            if not self._arquivo_esta_em_uso(caminho_arquivo):
                if tentativa > 0: # Só loga se precisou de mais de uma tentativa
                    self.log_message(f"Ficheiro '{os.path.basename(caminho_arquivo)}' libertado após {tentativa + 1} tentativa(s).")
                else:
                    self.log_message(f"Ficheiro '{os.path.basename(caminho_arquivo)}' está livre na primeira verificação.")
                return True # Ficheiro está livre

            self.log_message(f"Ficheiro '{os.path.basename(caminho_arquivo)}' ainda em uso. Tentativa {tentativa + 1}/{max_tentativas}. A aguardar {delay_segundos}s...")
            QApplication.processEvents() # Permite que a UI não congele durante a espera
            time.sleep(delay_segundos)

        self.log_message(f"AVISO: Ficheiro '{os.path.basename(caminho_arquivo)}' ainda parece estar em uso após {max_tentativas} tentativas.")
        return False # Ficheiro não foi libertado

# ________________________________________________________________________________________________

    # Este método deve estar dentro da classe CartaoApp

    def _copiar_arquivo_seguro(self, caminho_origem: str, caminho_destino: str, max_tentativas_copia: int = 3,
                               delay_tentativa_seg: float = 1.0) -> bool:
        """
        Copia um ficheiro da origem para o destino de forma mais segura, incluindo:
        - Verificação se o ficheiro de origem existe.
        - Tentativa de aguardar a libertação do ficheiro de destino (se existir).
        - Múltiplas tentativas de cópia em caso de falha.
        - Criação de um backup temporário do ficheiro de destino antes de sobrescrever.
        - Tentativa de restaurar o backup em caso de falha na cópia.

        Args:
            caminho_origem: O caminho completo do ficheiro de origem.
            caminho_destino: O caminho completo do ficheiro de destino.
            max_tentativas_copia: Número máximo de tentativas para a operação de cópia.
            delay_tentativa_seg: Delay em segundos entre as tentativas de cópia.

        Returns:
            True se a cópia for bem-sucedida, False caso contrário.
        """
        self.log_message(
            f"A iniciar cópia segura de '{os.path.basename(caminho_origem)}' para '{os.path.basename(caminho_destino)}'.")

        if not os.path.exists(caminho_origem):
            self.log_message(f"ERRO em _copiar_arquivo_seguro: Ficheiro de origem '{caminho_origem}' não existe.")
            QMessageBox.critical(self, "Erro de Cópia", f"O ficheiro de origem não foi encontrado:\n{caminho_origem}")
            return False

        # Se o ficheiro de destino já existe, tenta aguardar a sua libertação
        if os.path.exists(caminho_destino):
            self.log_message(
                f"Ficheiro de destino '{os.path.basename(caminho_destino)}' existe. A verificar se está livre...")
            if not self._aguardar_liberacao_arquivo(caminho_destino, max_tentativas=5, delay_segundos=0.5):
                self.log_message(
                    f"ERRO em _copiar_arquivo_seguro: Ficheiro de destino '{caminho_destino}' ainda em uso após tentativas. Cópia abortada.")
                QMessageBox.warning(self, "Ficheiro em Uso",
                                    f"O ficheiro de destino '{os.path.basename(caminho_destino)}' parece estar em uso e não pode ser sobrescrito agora.")
                return False
            self.log_message(
                f"Ficheiro de destino '{os.path.basename(caminho_destino)}' está livre para ser sobrescrito.")

        caminho_backup = None
        for tentativa in range(max_tentativas_copia):
            self.log_message(f"Tentativa de cópia {tentativa + 1}/{max_tentativas_copia}...")
            QApplication.processEvents()
            try:
                # 1. Faz backup do ficheiro de destino original, se existir
                if os.path.exists(caminho_destino):
                    caminho_backup = caminho_destino + ".bak_copia_segura"
                    if os.path.exists(caminho_backup):  # Remove backup antigo se existir
                        try:
                            os.remove(caminho_backup)
                        except OSError:
                            pass  # Ignora se não conseguir remover backup antigo

                    shutil.copy2(caminho_destino, caminho_backup)
                    self.log_message(f"Backup do destino original criado em: '{os.path.basename(caminho_backup)}'")

                # 2. Realiza a cópia
                shutil.copy2(caminho_origem, caminho_destino)

                # 3. Verifica se a cópia foi bem-sucedida (existe e tem tamanho > 0)
                if os.path.exists(caminho_destino) and os.path.getsize(caminho_destino) > 0:
                    self.log_message(
                        f"Cópia de '{os.path.basename(caminho_origem)}' para '{os.path.basename(caminho_destino)}' bem-sucedida na tentativa {tentativa + 1}.")
                    # Remove o backup se a cópia foi bem-sucedida
                    if caminho_backup and os.path.exists(caminho_backup):
                        try:
                            os.remove(caminho_backup)
                            self.log_message(f"Backup '{os.path.basename(caminho_backup)}' removido com sucesso.")
                        except OSError as e_rm_bak:
                            self.log_message(
                                f"Aviso: Não foi possível remover o ficheiro de backup '{os.path.basename(caminho_backup)}'. Erro: {e_rm_bak}")
                    return True  # Sucesso

            except PermissionError as e_perm:
                self.log_message(f"ERRO de Permissão na tentativa {tentativa + 1} de cópia: {e_perm}")
            except IOError as e_io:
                self.log_message(f"ERRO de I/O na tentativa {tentativa + 1} de cópia: {e_io}")
            except Exception as e_geral:
                self.log_message(f"ERRO Inesperado na tentativa {tentativa + 1} de cópia: {e_geral}")

            # Se a tentativa falhou e não é a última, aguarda antes de tentar novamente
            if tentativa < max_tentativas_copia - 1:
                self.log_message(f"A aguardar {delay_tentativa_seg}s antes da próxima tentativa de cópia...")
                time.sleep(delay_tentativa_seg)
            else:  # Última tentativa falhou
                self.log_message(f"Todas as {max_tentativas_copia} tentativas de cópia falharam.")
                # Tenta restaurar o backup, se um foi feito
                if caminho_backup and os.path.exists(caminho_backup):
                    self.log_message(
                        f"A tentar restaurar o backup '{os.path.basename(caminho_backup)}' para '{os.path.basename(caminho_destino)}'...")
                    try:
                        # Garante que o destino (potencialmente corrompido) não impeça o move
                        if os.path.exists(caminho_destino):
                            os.remove(caminho_destino)
                        shutil.move(caminho_backup, caminho_destino)
                        self.log_message("Backup restaurado com sucesso.")
                    except Exception as e_restore:
                        self.log_message(
                            f"ERRO CRÍTICO: Não foi possível restaurar o backup '{os.path.basename(caminho_backup)}'. O ficheiro de destino '{os.path.basename(caminho_destino)}' pode estar corrompido ou ausente. Erro: {e_restore}")
                        QMessageBox.critical(self, "Falha Crítica na Cópia",
                                             f"Falha ao copiar o ficheiro e também ao restaurar o backup.\n"
                                             f"O ficheiro '{os.path.basename(caminho_destino)}' pode estar num estado inconsistente.")
                elif os.path.exists(caminho_destino) and os.path.getsize(
                        caminho_destino) == 0:  # Se a cópia criou um ficheiro vazio
                    try:
                        os.remove(caminho_destino)
                        self.log_message(
                            f"Ficheiro de destino vazio '{os.path.basename(caminho_destino)}' removido após falha na cópia.")
                    except OSError:
                        pass

                QMessageBox.warning(self, "Falha na Cópia",
                                    f"Não foi possível copiar o ficheiro '{os.path.basename(caminho_origem)}' para '{os.path.basename(caminho_destino)}' após {max_tentativas_copia} tentativas.")
                return False  # Todas as tentativas falharam

        return False  # Segurança, não deveria ser alcançado se a lógica do loop estiver correta

# ________________________________________________________________________________________________

    # Em app_window.py

    def abrir_dialogo_gerenciar_regras(self):
        """
        Abre o diálogo para o usuário criar e editar as Regras de Texto.
        Esta versão NÃO abre o Photoshop.
        """
        modelo_selecionado = self.modelo_combobox.currentText()
        if not modelo_selecionado or modelo_selecionado == utils.TEXTO_NENHUM_MODELO:
            QMessageBox.warning(self, "Ação Inválida", "Selecione um modelo para gerenciar suas regras.")
            return

        config_modelo = self.configuracoes_modelos.get(modelo_selecionado, {})
        dados_especificos = config_modelo.get("dados_especificos", [])
        regras_atuais = config_modelo.get("regras_texto", {})

        if not dados_especificos:
            QMessageBox.information(self, "Dados Não Configurados",
                                    "Configure primeiro os 'Dados Específicos' do modelo usando o botão 'Modificar'.\n"
                                    "As regras de texto usam esses dados como variáveis.")
            return

        self.log_message(f"Abrindo gerenciador de regras para o modelo: {modelo_selecionado}")

        # Chama o diálogo passando os dados da tabela (variáveis) e as regras atuais
        dialogo = GerenciarRegrasTextoDialog(
            dados_especificos_disponiveis=dados_especificos,
            regras_atuais=regras_atuais,
            parent=self
        )

        dialogo.regrasSalvas.connect(
            lambda regras: self._processar_regras_salvas(modelo_selecionado, regras)
        )

        dialogo.exec()

# ________________________________________________________________________________________________

    def _processar_regras_salvas(self, nome_modelo, novas_regras):
        """Recebe as regras do diálogo e as salva na configuração."""
        if nome_modelo in self.configuracoes_modelos:
            self.configuracoes_modelos[nome_modelo]['regras_texto'] = novas_regras
            self.log_message(f"Regras de texto para o modelo '{nome_modelo}' atualizadas.")
            utils.salvar_configuracoes_camadas_modelos(self.configuracoes_modelos)
        else:
            self.log_message(f"ERRO: Tentativa de salvar regras para modelo não encontrado: {nome_modelo}")