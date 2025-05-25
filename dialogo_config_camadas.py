from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QScrollArea, QWidget, QFrame, QDialogButtonBox, QMessageBox,
    QSizePolicy # Adicionado para spacers
)
from PySide6.QtCore import Qt, Signal
from functools import partial  # Vamos usar isso para os botões X


class ConfigCamadasDialog(QDialog):
    configuracaoSalva = Signal(list)

    def __init__(self, psd_filename, camadas_existentes=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Configurar Camadas para: {psd_filename}")
        self.setFixedSize(500, 350)
        self.setObjectName(f"ConfigCamadasDialog_{psd_filename.replace('.', '_')}")  # Nome de objeto único

        self.psd_filename = psd_filename
        # Esta lista agora vai guardar tuplas: (QHBoxLayout da linha, QLineEdit do nome)
        self.linhas_de_camada_widgets = []

        main_layout = QVBoxLayout(self)
        instruction_label = QLabel(
            "Quais camadas seu modelo possui?\nPreencha o nome da camada exatamente como está no Photoshop.")
        main_layout.addWidget(instruction_label)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area_widget_contents = QWidget()
        self.layout_camadas = QVBoxLayout(self.scroll_area_widget_contents)
        self.layout_camadas.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.scroll_area.setWidget(self.scroll_area_widget_contents)
        main_layout.addWidget(self.scroll_area)

        self.btn_add_camada = QPushButton("+ Adicionar Camada")
        self.btn_add_camada.clicked.connect(self.adicionar_linha_camada_vazia)  # Conectado a um novo método wrapper
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
            self.adicionar_linha_camada()  # Adiciona a primeira linha em branco

        self._atualizar_estado_botoes_mover()  # Chama para definir o estado inicial dos botões

        self.setLayout(main_layout)

    def adicionar_linha_camada_vazia(self):
        self.adicionar_linha_camada(nome_camada_texto="")
        self._atualizar_estado_botoes_mover() # Atualiza após adicionar

    def adicionar_linha_camada(self, nome_camada_texto=""):
        linha_widget = QWidget()
        linha_layout = QHBoxLayout(linha_widget)
        linha_layout.setContentsMargins(0,0,0,0)

        edit_nome_camada = QLineEdit(nome_camada_texto)
        edit_nome_camada.setPlaceholderText("Nome da Camada no PSD")
        linha_layout.addWidget(edit_nome_camada)

        btn_excluir_linha = QPushButton("X")
        btn_excluir_linha.setFixedSize(30, 30)
        btn_excluir_linha.setStyleSheet("QPushButton { color: white; background-color: #e63946; font-weight: bold; border-radius: 5px; }")
        btn_excluir_linha.setToolTip("Remover esta camada")
        btn_excluir_linha.clicked.connect(
            partial(self._remover_linha_camada, linha_widget, edit_nome_camada)
        )
        linha_layout.addWidget(btn_excluir_linha)

        # --- NOVOS BOTÕES MOVER ---
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
        # --- FIM NOVOS BOTÕES MOVER ---

        self.layout_camadas.addWidget(linha_widget)
        # Agora guardamos os botões de mover também para poder habilitá-los/desabilitá-los
        self.linhas_de_camada_widgets.append(
            {"widget": linha_widget, "edit": edit_nome_camada, "up_btn": btn_mover_cima, "down_btn": btn_mover_baixo}
        )
        # Não chamamos _atualizar_estado_botoes_mover aqui, pois quem chama (adicionar_linha_camada_vazia ou o __init__) já o faz.

    def _remover_linha_camada(self, linha_widget_a_remover, edit_camada_a_remover):
        if len(self.linhas_de_camada_widgets) > 1:
            item_para_remover = None
            for item_dict in self.linhas_de_camada_widgets:
                if item_dict["edit"] == edit_camada_a_remover:
                    item_para_remover = item_dict
                    break

            if item_para_remover:
                self.linhas_de_camada_widgets.remove(item_para_remover)

            # Remove o widget da linha do layout e o deleta para liberar memória
            self.layout_camadas.removeWidget(linha_widget_a_remover)
            linha_widget_a_remover.deleteLater()
        else:
            # Se é a última linha, apenas limpa o texto do QLineEdit, como você pediu.
            edit_camada_a_remover.clear()

    def _mover_linha(self, linha_widget_a_mover, direcao):
        # Encontra o índice atual do widget da linha e seu dicionário de dados
        idx_atual = -1
        dados_linha_atual = None
        for i, item_dict in enumerate(self.linhas_de_camada_widgets):
            if item_dict["widget"] == linha_widget_a_mover:
                idx_atual = i
                dados_linha_atual = item_dict
                break

        if idx_atual == -1:  # Não deveria acontecer
            return

        idx_novo = idx_atual + direcao

        # Verifica se o novo índice é válido
        if 0 <= idx_novo < len(self.linhas_de_camada_widgets):
            # Remove da lista de dados e insere na nova posição
            self.linhas_de_camada_widgets.pop(idx_atual)
            self.linhas_de_camada_widgets.insert(idx_novo, dados_linha_atual)

            # Remove do layout e insere na nova posição
            self.layout_camadas.removeWidget(linha_widget_a_mover)
            self.layout_camadas.insertWidget(idx_novo, linha_widget_a_mover)

            self._atualizar_estado_botoes_mover()  # Atualiza o estado de todos os botões


    def _atualizar_estado_botoes_mover(self):
        num_linhas = len(self.linhas_de_camada_widgets)
        for i, item_dict in enumerate(self.linhas_de_camada_widgets):
            item_dict["up_btn"].setEnabled(i > 0)  # Habilita "Cima" se não for o primeiro
            item_dict["down_btn"].setEnabled(i < num_linhas - 1)  # Habilita "Baixo" se não for o último

    def accept(self):
        nomes_camadas = []
        for item_dict in self.linhas_de_camada_widgets:  # Agora é uma lista de dicionários
            nome = item_dict["edit"].text().strip()
            # ... (resto do accept como antes) ...
            if nome:
                if nome in nomes_camadas:
                    QMessageBox.warning(self, "Nome Duplicado",
                                        f"O nome de camada '{nome}' já foi adicionado. Use nomes únicos.")
                    return
                nomes_camadas.append(nome)

        if not nomes_camadas:
            reply = QMessageBox.question(self, "Nenhuma Camada Definida",
                                         "Nenhuma camada foi definida...",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                return

        self.configuracaoSalva.emit(nomes_camadas)
        super().accept()

    # O método reject() não precisa de mudanças por enquanto.
    # O método get_configuracao_camadas() também não é essencial se usarmos o sinal.