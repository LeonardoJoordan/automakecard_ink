import win32com.client # Necessário para psApp = win32com.client.Dispatch(...) em gerar_cartoes
import psutil
import gc
import time
import shutil
import os
# Imports de bibliotecas de terceiros (instaladas com pip)
from PIL import Image # Usado em atualizar_preview_modelo

# Imports do PySide6
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, # QApplication é usado para processEvents
    QTextEdit, QLabel, QComboBox, QPushButton, QFileDialog, QMessageBox,
    QSizePolicy, QTableWidgetItem, QAbstractItemView, QHeaderView
)
from PySide6.QtGui import (
    QFont, QPixmap, QImage, QPalette, QColor, QKeySequence, QIcon, QGuiApplication
)
from PySide6.QtCore import Qt, QSize, QMimeData, QEvent # Removido QTranslator, QLocale, QLibraryInfo daqui

# Imports dos nossos próprios módulos
import utils
import ps_utils
from custom_widgets import CustomTableWidget
from dialogo_config_camadas import ConfigCamadasDialog

PASTA_PADRAO_MODELOS = "modelos"
PASTA_PADRAO_SAIDA = "cartoes_gerados"
TEXTO_NENHUM_MODELO = "(nenhum modelo disponível)"
CAMADA_TRATAMENTO = "tratamento"
CAMADA_NOME = "nome"
CAMADA_CONJUGE = "conjuge"
CAMADA_DATA = "data"
PS_EXPORT_FORMAT_PNG = 13

class CartaoApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerenciador de Modelos PSD PySide6")
        self.setFixedSize(1000, 550)
        self.log_textbox = QTextEdit()
        self.log_textbox.setReadOnly(True)
        self.log_textbox.append("Log do programa...\n")
        self.configuracoes_modelos = utils.carregar_configuracoes_camadas_modelos()
        self.log_message(f"Configurações de camadas de {len(self.configuracoes_modelos)} modelos carregadas.")
        # selecionar modelo
        #self.modelo_combobox = QComboBox()
        #self.modelo_combobox.addItem(TEXTO_NENHUM_MODELO)  # Usando constante!
        # ANTES: self.modelo_combobox.currentTextChanged.connect(self.atualizar_preview_modelo)
        # DEPOIS:
        #self.modelo_combobox.currentTextChanged.connect(
        #    self._quando_modelo_mudar)  # Conecta a um novo método "gerenciador"
        #direita_layout.addWidget(self.modelo_combobox)

        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        # --- Coluna esquerda - Tabela ---
        tabela_container = QWidget()
        tabela_layout = QVBoxLayout(tabela_container)

        self.table_headers = [CAMADA_TRATAMENTO, CAMADA_NOME, CAMADA_CONJUGE, CAMADA_DATA]
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
            self.output_dir = os.path.abspath(PASTA_PADRAO_SAIDA)
            if not os.path.exists(PASTA_PADRAO_MODELOS):
                os.makedirs(PASTA_PADRAO_MODELOS)

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
        self.modelo_combobox.addItem(TEXTO_NENHUM_MODELO)
        self.modelo_combobox.currentTextChanged.connect(self._quando_modelo_mudar)
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

    def _quando_modelo_mudar(self, nome_modelo_selecionado):
        """Chamado quando o texto do combobox de modelos muda."""
        self.atualizar_preview_modelo(nome_modelo_selecionado)
        self._atualizar_tabela_para_modelo(nome_modelo_selecionado)

    def _atualizar_tabela_para_modelo(self, nome_modelo_psd):
        """
        Atualiza as colunas e cabeçalhos da tabela de dados
        com base nas camadas configuradas para o modelo PSD selecionado.
        """
        if not nome_modelo_psd or nome_modelo_psd == TEXTO_NENHUM_MODELO:
            # Limpa a tabela e desabilita o botão se nenhum modelo válido estiver selecionado
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(0)
            self.table_headers = []  # Limpa os cabeçalhos internos
            self.btn_gerar_cartoes.setEnabled(False)
            # Poderíamos adicionar uma mensagem na área da tabela aqui, se desejado.
            # Ex: self.data_table.setStatusTip("Selecione um modelo e configure suas camadas.")
            self.log_message("Nenhum modelo selecionado ou modelo inválido. Tabela limpa.")
            return

        # Busca as camadas configuradas para este modelo no nosso dicionário
        # self.configuracoes_modelos é carregado no __init__
        camadas_configuradas = self.configuracoes_modelos.get(nome_modelo_psd, [])

        if not camadas_configuradas:
            # Modelo existe, mas não tem camadas configuradas (ou a configuração é uma lista vazia)
            self.data_table.setRowCount(1)  # Uma linha para a mensagem
            self.data_table.setColumnCount(1)  # Uma coluna para a mensagem
            self.table_headers = []  # Sem cabeçalhos reais

            mensagem_item = QTableWidgetItem(
                "Este modelo não tem camadas configuradas. Use 'Modificar Modelo' para defini-las.")
            mensagem_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            mensagem_item.setFlags(mensagem_item.flags() ^ Qt.ItemFlag.ItemIsEditable)  # Torna não editável
            self.data_table.setItem(0, 0, mensagem_item)
            self.data_table.horizontalHeader().setVisible(False)  # Esconde cabeçalho
            self.data_table.verticalHeader().setVisible(False)  # Esconde cabeçalho vertical

            # Esticar a célula da mensagem para ocupar toda a tabela
            self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            self.data_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

            self.btn_gerar_cartoes.setEnabled(False)  # Não pode gerar sem camadas
            self.log_message(f"Modelo '{nome_modelo_psd}' não possui camadas configuradas.")
        else:
            # Temos camadas configuradas! Vamos montar a tabela.
            self.table_headers = camadas_configuradas  # Atualiza os cabeçalhos da instância

            num_colunas = len(self.table_headers)
            self.data_table.setColumnCount(num_colunas)
            self.data_table.setHorizontalHeaderLabels(self.table_headers)

            # Restaurar visibilidade dos cabeçalhos e modo de redimensionamento
            self.data_table.horizontalHeader().setVisible(True)
            self.data_table.verticalHeader().setVisible(True)
            self.data_table.horizontalHeader().setSectionResizeMode(
                QHeaderView.ResizeMode.Stretch)  # Ou como estava antes
            self.data_table.verticalHeader().setSectionResizeMode(
                QHeaderView.ResizeMode.Interactive)  # Ou como estava antes

            # Define um número padrão de linhas ou mantém as existentes se a estrutura for compatível
            # Por simplicidade, vamos resetar para um número padrão de linhas.
            # Você pode querer uma lógica mais sofisticada aqui no futuro para preservar dados.
            self.data_table.setRowCount(0)  # Limpa linhas antigas
            self.data_table.setRowCount(15)  # Adiciona 15 linhas padrão (ou use uma constante)

            self.btn_gerar_cartoes.setEnabled(True)
            self.log_message(
                f"Tabela atualizada para o modelo '{nome_modelo_psd}' com as colunas: {self.table_headers}")

    # Dentro da classe CartaoApp, em app_window.py

    def gerar_cartoes(self):
        self.log_message("Iniciando geração de cartões...")
        linhas = self.data_table.rowCount()
        colunas = self.data_table.columnCount()
        headers = [self.data_table.horizontalHeaderItem(i).text() for i in range(colunas)]

        # ... (pegar headers, etc.) ...

        psApp = None
        doc = None  # Variável para o documento
        try:
            self.log_message("Conectando ao Photoshop...")
            QApplication.processEvents()
            psApp = win32com.client.Dispatch("Photoshop.Application")
            psApp.Visible = False


            # Pega o caminho do modelo UMA VEZ
            psd_modelo_selecionado = self.modelo_combobox.currentText()
            if not psd_modelo_selecionado or psd_modelo_selecionado == TEXTO_NENHUM_MODELO:
                self.log_message("ERRO: Nenhum modelo PSD selecionado.")
                QMessageBox.warning(self, "Seleção de Modelo", "Por favor, selecione um modelo PSD.")
                return

            psd_path_modelo_unico = os.path.abspath(os.path.join(PASTA_PADRAO_MODELOS, psd_modelo_selecionado))
            if not os.path.exists(psd_path_modelo_unico):
                self.log_message(f"ERRO: Modelo PSD não encontrado: {psd_path_modelo_unico}")
                QMessageBox.critical(self, "Erro de Arquivo",
                                     f"O arquivo modelo PSD não foi encontrado:\n{psd_path_modelo_unico}")
                return

            self.log_message(f"Abrindo modelo PSD: {psd_path_modelo_unico}...")
            QApplication.processEvents()
            doc = psApp.Open(psd_path_modelo_unico)  # Abrir o documento UMA VEZ
            export_options_obj = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
            export_options_obj.Format = PS_EXPORT_FORMAT_PNG  # PNG
            export_options_obj.PNG8 = False  # PNG-24


            total_validos = 0
            for row in range(linhas):
                # PASSO A: Inicializa as variáveis para CADA linha
                linha_dados = {}  # Dicionário para guardar os dados da linha atual
                linha_vazia = True  # Assume que a linha está vazia até provar o contrário

                # PASSO B: Loop INTERNO para ler cada coluna da linha atual
                for col in range(colunas):
                    item = self.data_table.item(row, col)
                    valor = item.text().strip() if item else ""
                    if valor:  # Se encontrar qualquer valor na linha
                        linha_vazia = False  # Marca que a linha NÃO está vazia
                    linha_dados[headers[col]] = valor  # Guarda o valor (ou string vazia)

                # PASSO C: Pula para a próxima linha se esta estiver completamente vazia
                if linha_vazia:
                    continue

                # PASSO D: Se a linha não está vazia, incrementa e processa
                total_validos += 1

                self.log_message(
                    f"Processando cartão {total_validos} para: {linha_dados.get('nome', '(nome não encontrado)')}")  # Adicionado .get() para segurança
                QApplication.processEvents()

                # Gera nome do arquivo PNG
                data_bruta = linha_dados.get('data', '')
                nome_base = linha_dados.get('nome', f'cartao_{total_validos}')
                if data_bruta and len(data_bruta) >= 5:
                    nome_png = f"{data_bruta[3:5]}.{data_bruta[:2]} - {nome_base}.png"
                else:
                    nome_png = f"{nome_base}.png"
                output_path = os.path.join(self.output_dir, nome_png)

                campos_psd = {}
                for nome_camada_header in headers:  # 'headers' agora vem da configuração do modelo
                    # A chave em linha_dados é o nome_camada_header
                    campos_psd[nome_camada_header] = linha_dados.get(nome_camada_header, '')

                # CASO ESPECIAL PARA A DATA (se ela ainda precisar de formatação por extenso)
                # Se uma das suas camadas configuradas se chamar, por exemplo, "data_evento_psd"
                # e você quer que o valor para ela seja a data por extenso vinda da coluna "data" da tabela:
                #
                # Supondo que nos headers você tenha um nome de camada que representa a data (ex: "data_formatar")
                # E no PSD você tem uma camada que vai receber a data formatada (ex: "TEXTO_DATA_EVENTO_PSD")
                #
                # Esta parte precisará de uma lógica mais específica baseada em como você quer
                # mapear os dados da tabela para os campos que vão para o ps_utils.
                # Por simplicidade agora, vamos assumir que todos os headers da tabela
                # correspondem diretamente a uma camada no PSD que recebe aquele texto.

                # Se você ainda tem uma coluna "data" na tabela e quer formatá-la para uma camada específica
                # que foi configurada pelo usuário (ex: o usuário configurou uma camada chamada "DataCompleta"):
                if 'data' in linha_dados:  # Se a coluna 'data' (original) ainda está sendo usada na tabela
                    data_bruta_para_formatar = linha_dados.get('data', '')
                    data_formatada = utils.data_por_extenso(data_bruta_para_formatar)

                    # Agora, como saber em qual camada do PSD colocar essa data_formatada?
                    # Se o usuário configurou uma camada chamada, por exemplo, CAMADA_DATA_FORMATADA_PSD
                    # e essa CAMADA_DATA_FORMATADA_PSD é um dos 'headers' da tabela atual:
                    # if CAMADA_DATA_FORMATADA_PSD in campos_psd:
                    #    campos_psd[CAMADA_DATA_FORMATADA_PSD] = data_formatada
                    # Ou, se uma das camadas configuradas pelo usuário se chamar literalmente "data":
                    if utils.CAMADA_DATA in campos_psd:  # Se CAMADA_DATA foi um dos headers configurados
                        campos_psd[utils.CAMADA_DATA] = data_formatada

                # Se as constantes CAMADA_TRATAMENTO, etc. ainda são relevantes como
                # NOMES DE CAMADAS REAIS no PSD que o usuário pode configurar, então o
                # dicionário 'campos_psd' acima já estará correto.
                # A questão é: os 'headers' da tabela SÃO os nomes das camadas do PSD.

                ps_utils.gerar_cartao_photoshop(psApp, doc, output_path, campos_psd, export_options_obj)
                self.log_message(f"Cartão salvo: {output_path}")

            self.log_message(f"Geração finalizada. {total_validos} cartões preparados.")

        except Exception as e:
            self.log_message(f"ERRO GERAL: Ocorreu um problema ao gerar os cartões. {e}")
            QMessageBox.critical(self, "Erro na Geração", f"Ocorreu um erro: {e}")
        finally:
            if doc is not None:
                doc.Close(2)  # Fechar o documento no final de tudo (sem salvar alterações)
                doc = None  # Limpar referência
            if psApp is not None:
                psApp.Quit() # Descomente se quiser fechar o Photoshop
                psApp = None
                self.log_message("Conexão com o Photoshop finalizada.")


    def selecionar_pasta_saida(self):
        pasta = QFileDialog.getExistingDirectory(self, "Escolha a pasta de saída", self.output_dir)
        if pasta:
            self.output_dir = pasta
            self.saida_dir_label.setText(self.output_dir)
            utils.save_last_output_dir(self.output_dir)
            self.log_message(f"Pasta de saída definida para: {self.output_dir}")


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
        if not os.path.exists(PASTA_PADRAO_MODELOS):
            os.makedirs(PASTA_PADRAO_MODELOS)
            self.log_message("Pasta '{PASTA_PADRAO_MODELOS}' criada.")

    def log_message(self, message):
        self.log_textbox.append(message)
        self.log_textbox.ensureCursorVisible()

    def atualizar_modelos_combobox(self):
        self.garantir_pasta_modelos()
        try:
            # Lista arquivos .psd
            arquivos = [f for f in os.listdir(PASTA_PADRAO_MODELOS) if f.lower().endswith(".psd")]
        except FileNotFoundError:
            self.log_message("Erro: Pasta '{PASTA_PADRAO_MODELOS}' não encontrada ao listar arquivos.")
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
                self.modelo_combobox.addItem(TEXTO_NENHUM_MODELO)
        else:
            self.modelo_combobox.addItem(TEXTO_NENHUM_MODELO)

        self.modelo_combobox.blockSignals(False)
        self._quando_modelo_mudar(self.modelo_combobox.currentText())

    def adicionar_modelo(self):
        self.garantir_pasta_modelos()
        arquivo_psd_selecionado, _ = QFileDialog.getOpenFileName(
            self,
            "Selecione um modelo (.psd)",
            PASTA_PADRAO_MODELOS,  # Usando a constante, ótimo!
            "Arquivos do Photoshop (*.psd);;Todos os Arquivos (*.*)"
        )

        if arquivo_psd_selecionado:  # Se o usuário selecionou um arquivo
            nome_arquivo_base = os.path.basename(arquivo_psd_selecionado)
            caminho_destino = os.path.join(PASTA_PADRAO_MODELOS, nome_arquivo_base)

            # Verifica se o modelo já existe e pergunta sobre substituir (como antes)
            if os.path.exists(caminho_destino):
                resp = QMessageBox.question(self, "Substituir modelo",
                                            f"O modelo '{nome_arquivo_base}' já existe. Deseja substituir?",
                                            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if resp == QMessageBox.StandardButton.No:
                    return  # Usuário não quis substituir, então não fazemos mais nada

            # --- INÍCIO DAS NOVAS LINHAS PARA O DIÁLOGO DE CONFIGURAÇÃO ---

            # Passo 2 que você não compreendeu: Criar uma instância do diálogo
            # Passamos 'nome_arquivo_base' para o título e 'self' para que o diálogo seja "filho" da janela principal.
            # Como é um NOVO modelo, passamos camadas_existentes=None (ou uma lista vazia).
            dialogo_config = ConfigCamadasDialog(nome_arquivo_base, camadas_existentes=None, parent=self)

            # Passo 3 que você não compreendeu: Conectar o sinal do diálogo ao nosso novo método
            # Quando o diálogo emitir 'configuracaoSalva', nosso método '_processar_camadas_configuradas' será chamado.
            # Usamos uma função lambda para poder passar 'nome_arquivo_base' também para o nosso método.
            dialogo_config.configuracaoSalva.connect(
                lambda lista_camadas: self._processar_camadas_configuradas(nome_arquivo_base, lista_camadas)
            )

            # Passo 4 que você não compreendeu: Executar o diálogo e verificar se o usuário salvou
            # dialog.exec() mostra o diálogo e espera o usuário interagir.
            # Ele retorna True (ou um valor específico) se o usuário clicou em "Salvar/OK".
            if dialogo_config.exec():
                # Usuário clicou em "Salvar" no diálogo de configuração.
                # O sinal 'configuracaoSalva' já foi emitido e nosso método _processar_camadas_configuradas
                # já foi chamado e imprimiu a lista (se não estiver vazia após o aviso).
                self.log_message(f"Configuração de camadas para '{nome_arquivo_base}' definida.")
            else:
                # Usuário clicou em "Cancelar" no diálogo ou fechou a janela.
                # Conforme sua regra: o modelo será adicionado, mas ficará "não configurado".
                # O método _processar_camadas_configuradas NÃO será chamado se ele cancelou.
                # Precisamos explicitamente "registrar" que este modelo tem uma configuração vazia.
                self._processar_camadas_configuradas(nome_arquivo_base, [])  # Passa uma lista vazia
                self.log_message(f"Configuração de camadas para '{nome_arquivo_base}' não definida pelo usuário.")

            # --- FIM DAS NOVAS LINHAS PARA O DIÁLOGO DE CONFIGURAÇÃO ---

            # O resto do código continua como antes: copiar o arquivo e atualizar
            try:
                shutil.copy2(arquivo_psd_selecionado, caminho_destino)
                self.log_message(f"Modelo '{nome_arquivo_base}' adicionado/atualizado.")
                self.atualizar_modelos_combobox()  # Atualiza a lista de modelos
            except Exception as e:
                QMessageBox.critical(self, "Erro ao adicionar", f"Não foi possível copiar o arquivo: {e}")
                self.log_message(f"Erro ao copiar '{nome_arquivo_base}': {e}")

    def _processar_camadas_configuradas(self, nome_arquivo_psd, lista_camadas_configuradas):
        self.log_message(f"Salvando configuração de camadas para '{nome_arquivo_psd}': {lista_camadas_configuradas}")

        if not lista_camadas_configuradas:  # Se o usuário salvou sem camadas (ou cancelou e passamos lista vazia)
            # Se o modelo já existia e agora não tem camadas, removemos a configuração dele.
            # Se era um modelo novo e não configurou, não fazemos nada aqui ainda,
            # pois ele não estaria em self.configuracoes_modelos.
            # Podemos refinar isso se necessário, mas por ora, se a lista for vazia,
            # vamos garantir que ele não tenha uma configuração inválida.
            if nome_arquivo_psd in self.configuracoes_modelos:
                del self.configuracoes_modelos[nome_arquivo_psd]
                self.log_message(
                    f"Configuração de camadas removida para '{nome_arquivo_psd}' pois nenhuma camada foi definida.")
        else:
            self.configuracoes_modelos[nome_arquivo_psd] = lista_camadas_configuradas

        # Salva o dicionário inteiro de configurações no arquivo JSON
        if utils.salvar_configuracoes_camadas_modelos(self.configuracoes_modelos):
            self.log_message("Arquivo de configuração de camadas salvo com sucesso.")
        else:
            self.log_message("ERRO ao salvar arquivo de configuração de camadas.")
            QMessageBox.critical(self, "Erro de Salvamento", "Não foi possível salvar as configurações de camadas.")

    def modificar_modelo(self):
        modelo_selecionado = self.modelo_combobox.currentText()
        if not modelo_selecionado or modelo_selecionado == TEXTO_NENHUM_MODELO:
            QMessageBox.information(self, "Modificar Modelo", "Nenhum modelo selecionado para modificar.")
            return

        camadas_atuais_para_edicao = self.configuracoes_modelos.get(modelo_selecionado, [])
        self.log_message(f"Modificando modelo '{modelo_selecionado}'. Camadas atuais: {camadas_atuais_para_edicao}")

        # Abrir o diálogo para o usuário selecionar o arquivo PSD
        novo_arquivo_psd_path, _ = QFileDialog.getOpenFileName(
            self,
            f"Selecionar NOVO arquivo para o modelo '{modelo_selecionado}' (ou o mesmo para apenas editar camadas)",
            PASTA_PADRAO_MODELOS,
            "Arquivos do Photoshop (*.psd);;Todos os Arquivos (*.*)"
        )

        if not novo_arquivo_psd_path:  # Se o usuário cancelou a seleção do arquivo
            return

        # Configurar o diálogo de camadas
        dialogo_config = ConfigCamadasDialog(
            modelo_selecionado,
            camadas_existentes=camadas_atuais_para_edicao,
            parent=self
        )

        dialogo_config.configuracaoSalva.connect(
            lambda lista_camadas: self._processar_camadas_configuradas(modelo_selecionado, lista_camadas)
        )

        # Executar o diálogo de configuração
        if dialogo_config.exec():
            self.log_message(f"Configuração de camadas para '{modelo_selecionado}' foi atualizada.")
        else:
            self.log_message(
                f"Modificação da configuração de camadas para '{modelo_selecionado}' cancelada pelo usuário.")
            return  # Se cancelou, não precisa copiar o arquivo

        # Garantir que qualquer instância do Photoshop seja fechada antes de copiar
        self._forcar_fechamento_photoshop()

        caminho_destino_modelo = os.path.join(PASTA_PADRAO_MODELOS, modelo_selecionado)

        # Aguardar liberação completa do arquivo
        if not self._aguardar_liberacao_arquivo(caminho_destino_modelo):
            QMessageBox.critical(self, "Erro ao modificar",
                                 "O arquivo ainda está sendo usado por outro processo. "
                                 "Feche o Photoshop manualmente e tente novamente.")
            return

        # Definir o caminho de destino ANTES de tentar copiar


        # Copiar o arquivo de forma segura
        if not self._copiar_arquivo_seguro(novo_arquivo_psd_path, caminho_destino_modelo):
            QMessageBox.critical(self, "Erro ao modificar",
                                 "Não foi possível substituir o arquivo. Verifique se:\n"
                                 "• O arquivo não está aberto no Photoshop\n"
                                 "• Você tem permissões de escrita na pasta\n"
                                 "• Há espaço suficiente em disco")
            self.log_message(f"ERRO: Falha ao copiar arquivo para '{modelo_selecionado}'")
            return

        # Se chegou aqui, a cópia foi bem-sucedida
        self.log_message(f"Conteúdo do arquivo PSD para o modelo '{modelo_selecionado}' foi atualizado.")

        # Forçar regeneração do preview
        try:
            preview_path = os.path.join(PASTA_PADRAO_MODELOS,
                                        f"{os.path.splitext(modelo_selecionado)[0]}_preview.png")
            if os.path.exists(preview_path):
                os.remove(preview_path)

            self.atualizar_modelos_combobox()
            self.modelo_combobox.setCurrentText(modelo_selecionado)

        except Exception as e:
            self.log_message(f"Aviso: Erro ao atualizar preview para '{modelo_selecionado}': {e}")
            # Não é um erro crítico, apenas um aviso

    def excluir_modelo(self):
        modelo_selecionado = self.modelo_combobox.currentText()
        if not modelo_selecionado or modelo_selecionado == TEXTO_NENHUM_MODELO:
            QMessageBox.information(self, "Excluir Modelo", "Nenhum modelo selecionado para excluir.")
            return

        resp = QMessageBox.question(self, "Excluir Modelo",
                                    f"Tem certeza que deseja excluir o modelo '{modelo_selecionado}'?",
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if resp == QMessageBox.StandardButton.Yes:
            caminho_psd = os.path.join(PASTA_PADRAO_MODELOS, modelo_selecionado)
            preview_path = os.path.join(PASTA_PADRAO_MODELOS, f"{os.path.splitext(modelo_selecionado)[0]}_preview.png")
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

    def _forcar_fechamento_photoshop(self):
        """Força o fechamento de todas as instâncias conhecidas do Photoshop"""
        import psutil

        try:
            # Primeiro: tentar fechar instâncias conhecidas da classe
            if hasattr(self, "doc") and self.doc is not None:
                try:
                    self.doc.Close(2)  # Fecha sem salvar
                except:
                    pass
                self.doc = None

            if hasattr(self, "psApp") and self.psApp is not None:
                try:
                    self.psApp.Quit()
                except:
                    pass
                self.psApp = None

            # Segundo: buscar processos do Photoshop que podem estar rodando
            processos_photoshop = []
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if proc.info['name'] and 'photoshop' in proc.info['name'].lower():
                        processos_photoshop.append(proc)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

            # Se encontrou processos do Photoshop, tentar conectar e fechar
            if processos_photoshop:
                self.log_message(f"Encontrados {len(processos_photoshop)} processos do Photoshop ativos")
                try:
                    # Tenta conectar a uma instância existente
                    ps_app = win32com.client.GetActiveObject("Photoshop.Application")
                    if ps_app:
                        # Fecha todos os documentos abertos
                        while ps_app.Documents.Count > 0:
                            ps_app.Documents[0].Close(2)  # Fecha sem salvar
                        ps_app.Quit()
                        self.log_message("Photoshop existente foi fechado com sucesso")
                except:
                    self.log_message("Não foi possível conectar ao Photoshop existente")

            # Força garbage collection
            gc.collect()

            # Aguarda o sistema liberar recursos
            time.sleep(1.5)

            self.log_message("Cleanup do Photoshop concluído")

        except Exception as e:
            self.log_message(f"Erro durante cleanup do Photoshop: {e}")

    def _arquivo_esta_em_uso(self, caminho_arquivo):
        """Verifica se um arquivo está sendo usado por outro processo"""
        if not os.path.exists(caminho_arquivo):
            return False

        try:
            # Tenta abrir o arquivo em modo de escrita exclusiva
            with open(caminho_arquivo, 'r+b'):
                return False
        except (IOError, OSError, PermissionError):
            return True

    def _aguardar_liberacao_arquivo(self, caminho_arquivo, max_tentativas=15, delay_inicial=0.2):
        """
        Aguarda um arquivo ser liberado por outros processos

        Args:
            caminho_arquivo: Caminho do arquivo para verificar
            max_tentativas: Número máximo de tentativas
            delay_inicial: Delay inicial em segundos

        Returns:
            bool: True se o arquivo foi liberado, False caso contrário
        """
        if not os.path.exists(caminho_arquivo):
            return True  # Se não existe, não está em uso

        self.log_message(f"Verificando se arquivo está liberado: {os.path.basename(caminho_arquivo)}")

        for tentativa in range(max_tentativas):
            if not self._arquivo_esta_em_uso(caminho_arquivo):
                if tentativa > 0:
                    self.log_message(f"Arquivo liberado após {tentativa + 1} tentativas")
                return True

            # Delay progressivo: 0.2s, 0.4s, 0.6s, etc.
            delay = delay_inicial * (tentativa + 1)
            self.log_message(f"Arquivo em uso - tentativa {tentativa + 1}/{max_tentativas} (aguardando {delay:.1f}s)")
            time.sleep(delay)

        self.log_message(f"AVISO: Arquivo ainda em uso após {max_tentativas} tentativas")
        return False

    def _copiar_arquivo_seguro(self, origem, destino, max_tentativas=5):
        """
        Copia um arquivo com múltiplas tentativas e verificações de segurança

        Args:
            origem: Caminho do arquivo de origem
            destino: Caminho do arquivo de destino
            max_tentativas: Número máximo de tentativas

        Returns:
            bool: True se a cópia foi bem-sucedida, False caso contrário
        """
        if not os.path.exists(origem):
            self.log_message(f"ERRO: Arquivo de origem não existe: {origem}")
            return False

        # Verifica se o arquivo de destino está liberado
        if not self._aguardar_liberacao_arquivo(destino):
            self.log_message("ERRO: Arquivo de destino ainda está em uso")
            return False

        # Tenta a cópia com múltiplas tentativas
        for tentativa in range(max_tentativas):
            try:
                # Fazer backup do arquivo original se existir
                if os.path.exists(destino):
                    backup_path = destino + ".backup"
                    shutil.copy2(destino, backup_path)

                # Realizar a cópia
                shutil.copy2(origem, destino)

                # Verificar se a cópia foi bem-sucedida
                if os.path.exists(destino) and os.path.getsize(destino) > 0:
                    # Remove o backup se tudo deu certo
                    backup_path = destino + ".backup"
                    if os.path.exists(backup_path):
                        os.remove(backup_path)

                    if tentativa > 0:
                        self.log_message(f"Cópia bem-sucedida na tentativa {tentativa + 1}")
                    return True

            except PermissionError as e:
                delay = (tentativa + 1) * 0.5  # 0.5s, 1s, 1.5s, etc.
                self.log_message(f"Tentativa {tentativa + 1}/{max_tentativas} falhou: {e}")

                if tentativa < max_tentativas - 1:
                    self.log_message(f"Aguardando {delay}s antes da próxima tentativa...")
                    time.sleep(delay)

            except Exception as e:
                self.log_message(f"Erro inesperado na cópia: {e}")
                break

        # Se chegou aqui, a cópia falhou
        # Tentar restaurar backup se existir
        backup_path = destino + ".backup"
        if os.path.exists(backup_path):
            try:
                shutil.move(backup_path, destino)
                self.log_message("Backup restaurado após falha na cópia")
            except:
                self.log_message("AVISO: Não foi possível restaurar backup")

        return False

    def atualizar_preview_modelo(self, modelo_selecionado=None):
        if modelo_selecionado is None:
            modelo_selecionado = self.modelo_combobox.currentText()

        if not modelo_selecionado or modelo_selecionado == TEXTO_NENHUM_MODELO:
            self.preview_label.clear()
            self.preview_label.setText("Aguardando seleção de modelo")
            self.preview_label.setStyleSheet(
                "background-color: #404040; color: white; border-radius: 8px; border: 1px solid #505050;")
            self._current_pixmap = None
            return

        caminho_psd = os.path.join(PASTA_PADRAO_MODELOS, modelo_selecionado)
        # O nome do arquivo de preview será o nome do PSD sem a extensão, mais .png
        caminho_preview = os.path.join(PASTA_PADRAO_MODELOS, f"{os.path.splitext(modelo_selecionado)[0]}_preview.png")

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
                # <<< A MUDANÇA IMPORTANTE ESTÁ AQUI >>>
                with Image.open(caminho_psd) as pil_img_psd:  # Abre o PSD DENTRO do 'with'
                    if pil_img_psd.mode != 'RGBA':
                        pil_img_convertida = pil_img_psd.convert('RGBA')
                    else:
                        pil_img_convertida = pil_img_psd  # Se já for RGBA, apenas usa a referência

                    # Cria uma cópia em memória para salvar o preview PNG,
                    # permitindo que o arquivo PSD original seja fechado pelo 'with'
                    imagem_para_salvar_preview = pil_img_convertida.copy()
                    imagem_para_salvar_preview.save(caminho_preview)  # Salva o preview como PNG

                    # Cria o QImage a partir da imagem convertida (que ainda está em memória)
                    qimage = QImage(pil_img_convertida.tobytes("raw", "RGBA"),
                                    pil_img_convertida.width, pil_img_convertida.height,
                                    QImage.Format.Format_RGBA8888)
                # Neste ponto, ao sair do bloco 'with', o arquivo PSD em 'caminho_psd' é fechado pelo Pillow.

                self._current_pixmap = QPixmap.fromImage(qimage)

                scaled_pixmap = self._current_pixmap.scaledToHeight(self.preview_label.height(),
                                                                    Qt.TransformationMode.SmoothTransformation)
                self.preview_label.setPixmap(scaled_pixmap)
                self.preview_label.setStyleSheet(
                    "background-color: transparent; border: 1px solid gray; border-radius: 8px;")
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
