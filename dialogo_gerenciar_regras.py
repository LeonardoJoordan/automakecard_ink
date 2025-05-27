from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QScrollArea, QWidget, QFrame, QDialogButtonBox, QMessageBox,
    QSizePolicy
)
from PySide6.QtCore import Qt, Signal
from functools import partial


class GerenciarRegrasDialog(QDialog):
    configuracaoSalva = Signal(list)

    def __init__(self, psd_filename, camadas_existentes=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Configurar Dados para: {psd_filename}")
        self.setFixedSize(500, 350)
        self.setObjectName(f"GerenciarRegrasDialog_{psd_filename.replace('.', '_')}")

        self.psd_filename = psd_filename
        self.linhas_de_camada_widgets = []

        main_layout = QVBoxLayout(self)
        instruction_label = QLabel(
            "Adicione os 'Dados Específicos' que virarão as colunas da tabela.\n"
            "Inclua nomes de camadas do Photoshop e também 'dados virtuais' (ex: nome da mãe).")
        main_layout.addWidget(instruction_label)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area_widget_contents = QWidget()
        self.layout_camadas = QVBoxLayout(self.scroll_area_widget_contents)
        self.layout_camadas.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.scroll_area.setWidget(self.scroll_area_widget_contents)
        main_layout.addWidget(self.scroll_area)

        self.btn_add_camada = QPushButton("+ Adicionar Dado Específico")  # Nome do botão atualizado
        self.btn_add_camada.clicked.connect(self.adicionar_linha_camada_vazia)
        main_layout.addWidget(self.btn_add_camada)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        main_layout.addWidget(line)

        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        main_layout.addWidget(self.button_box)

        if camadas_existentes:
            for nome_camada in camadas_existentes:
                self.adicionar_linha_camada(nome_camada_texto=nome_camada)
        else:
            self.adicionar_linha_camada()

        self._atualizar_estado_botoes_mover()

        self.setLayout(main_layout)

    def abrir_dialogo_gerenciar_regras(self):
        """
        Prepara os dados e abre o diálogo para gerenciar as Regras de Texto.
        """
        modelo_selecionado = self.modelo_combobox.currentText()
        if not modelo_selecionado or modelo_selecionado == utils.TEXTO_NENHUM_MODELO:
            QMessageBox.warning(self, "Ação Inválida", "Selecione um modelo para gerenciar suas regras.")
            return

        self.log_message(f"Abrindo gerenciador de regras para o modelo: {modelo_selecionado}")

        caminho_psd = os.path.abspath(os.path.join(PASTA_PADRAO_MODELOS, modelo_selecionado))
        if not os.path.exists(caminho_psd):
            QMessageBox.critical(self, "Erro de Arquivo",
                                 f"O arquivo do modelo '{modelo_selecionado}' não foi encontrado.")
            return

        camadas_reais_no_psd = ps_utils.listar_camadas_de_texto(caminho_psd)
        if not camadas_reais_no_psd:
            QMessageBox.information(self, "Nenhuma Camada de Texto",
                                    "Não foi encontrada nenhuma camada de texto no arquivo PSD para criar regras.")
            return

        config_modelo = self.configuracoes_modelos.get(modelo_selecionado, {})
        regras_atuais = config_modelo.get("regras_texto", {})

        # Chama o NOVO e CORRETO diálogo
        dialogo = GerenciarRegrasTextoDialog(
            camadas_psd=camadas_reais_no_psd,
            regras_atuais=regras_atuais,
            parent=self
        )

        # Conecta o sinal 'regrasSalvas' (que agora existe) à nossa função de processamento
        dialogo.regrasSalvas.connect(
            lambda regras: self._processar_regras_salvas(modelo_selecionado, regras)
        )

        dialogo.exec()

    def adicionar_linha_camada_vazia(self):
        self.adicionar_linha_camada(nome_camada_texto="")
        self._atualizar_estado_botoes_mover()

    def adicionar_linha_camada(self, nome_camada_texto=""):
        linha_widget = QWidget()
        linha_layout = QHBoxLayout(linha_widget)
        linha_layout.setContentsMargins(0, 0, 0, 0)

        edit_nome_camada = QLineEdit(nome_camada_texto)
        edit_nome_camada.setPlaceholderText("Nome do Dado Específico")  # Texto do placeholder atualizado
        linha_layout.addWidget(edit_nome_camada)

        btn_excluir_linha = QPushButton("X")
        btn_excluir_linha.setFixedSize(30, 30)
        btn_excluir_linha.setStyleSheet(
            "QPushButton { color: white; background-color: #e63946; font-weight: bold; border-radius: 5px; }")
        btn_excluir_linha.setToolTip("Remover este campo")  # Tooltip atualizado
        btn_excluir_linha.clicked.connect(
            partial(self._remover_linha_camada, linha_widget, edit_nome_camada)
        )
        linha_layout.addWidget(btn_excluir_linha)

        btn_mover_cima = QPushButton("↑")
        btn_mover_cima.setFixedSize(30, 30)
        btn_mover_cima.setToolTip("Mover para cima")
        btn_mover_cima.clicked.connect(partial(self._mover_linha, linha_widget, direcao=-1))
        linha_layout.addWidget(btn_mover_cima)

        btn_mover_baixo = QPushButton("↓")
        btn_mover_baixo.setFixedSize(30, 30)
        btn_mover_baixo.setToolTip("Mover para baixo")
        btn_mover_baixo.clicked.connect(partial(self._mover_linha, linha_widget, direcao=1))
        linha_layout.addWidget(btn_mover_baixo)

        self.layout_camadas.addWidget(linha_widget)
        self.linhas_de_camada_widgets.append(
            {"widget": linha_widget, "edit": edit_nome_camada, "up_btn": btn_mover_cima, "down_btn": btn_mover_baixo}
        )

    def _remover_linha_camada(self, linha_widget_a_remover, edit_camada_a_remover):
        if len(self.linhas_de_camada_widgets) > 1:
            item_para_remover = None
            for item_dict in self.linhas_de_camada_widgets:
                if item_dict["edit"] == edit_camada_a_remover:
                    item_para_remover = item_dict
                    break

            if item_para_remover:
                self.linhas_de_camada_widgets.remove(item_para_remover)

            self.layout_camadas.removeWidget(linha_widget_a_remover)
            linha_widget_a_remover.deleteLater()
        else:
            edit_camada_a_remover.clear()

    def _mover_linha(self, linha_widget_a_mover, direcao):
        idx_atual = -1
        dados_linha_atual = None
        for i, item_dict in enumerate(self.linhas_de_camada_widgets):
            if item_dict["widget"] == linha_widget_a_mover:
                idx_atual = i
                dados_linha_atual = item_dict
                break

        if idx_atual == -1:
            return

        idx_novo = idx_atual + direcao

        if 0 <= idx_novo < len(self.linhas_de_camada_widgets):
            self.linhas_de_camada_widgets.pop(idx_atual)
            self.linhas_de_camada_widgets.insert(idx_novo, dados_linha_atual)

            self.layout_camadas.removeWidget(linha_widget_a_mover)
            self.layout_camadas.insertWidget(idx_novo, linha_widget_a_mover)

            self._atualizar_estado_botoes_mover()

    def _atualizar_estado_botoes_mover(self):
        num_linhas = len(self.linhas_de_camada_widgets)
        for i, item_dict in enumerate(self.linhas_de_camada_widgets):
            item_dict["up_btn"].setEnabled(i > 0)
            item_dict["down_btn"].setEnabled(i < num_linhas - 1)

    def accept(self):
        nomes_camadas = []
        for item_dict in self.linhas_de_camada_widgets:
            nome = item_dict["edit"].text().strip()
            if nome:
                if nome in nomes_camadas:
                    QMessageBox.warning(self, "Nome Duplicado",
                                        f"O nome '{nome}' já foi adicionado. Use nomes únicos.")
                    return
                nomes_camadas.append(nome)

        if not nomes_camadas:
            reply = QMessageBox.question(self, "Nenhum Dado Definido",
                                         "Nenhum 'Dado Específico' foi definido. Deseja salvar mesmo assim (o modelo ficará não configurado)?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                return

        self.configuracaoSalva.emit(nomes_camadas)
        super().accept()