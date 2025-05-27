# dialogo_regras_texto.py
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QScrollArea, QWidget, QDialogButtonBox, QMessageBox, QFrame
)
from PySide6.QtCore import Qt, Signal
from functools import partial

class GerenciarRegrasTextoDialog(QDialog):
    regrasSalvas = Signal(dict)

    def __init__(self, dados_especificos_disponiveis, regras_atuais, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Gerenciar Regras de Texto")
        self.setMinimumSize(700, 500)

        self.linhas_widgets = []
        main_layout = QVBoxLayout(self)

        # Mostra as variáveis que o usuário pode usar
        variaveis_formatadas = ", ".join([f"{{{d}}}" for d in dados_especificos_disponiveis])
        info_label = QLabel(f"<b>Variáveis disponíveis:</b> {variaveis_formatadas}")
        info_label.setWordWrap(True)
        main_layout.addWidget(info_label)

        # Cabeçalho
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("<b>Nome da Camada no Photoshop (Alvo)</b>"))
        header_layout.addWidget(QLabel("<b>Conteúdo da Camada (Regra)</b>"))
        header_layout.addSpacing(40) # Espaço para o botão de remover
        main_layout.addLayout(header_layout)

        # Área de Rolagem para as regras
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_widget = QWidget()
        self.regras_layout = QVBoxLayout(scroll_widget)
        self.regras_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        scroll_area.setWidget(scroll_widget)
        main_layout.addWidget(scroll_area)

        # Carrega as regras que já existem
        if regras_atuais:
            for camada, regra in regras_atuais.items():
                self.adicionar_linha_regra(camada, regra)

        # Botão para adicionar nova regra
        self.btn_add_regra = QPushButton("+ Adicionar Nova Regra")
        # Linha corrigida:
        self.btn_add_regra.clicked.connect(lambda: self.adicionar_linha_regra())
        main_layout.addWidget(self.btn_add_regra)

        # Botões de Salvar e Cancelar
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)

    def adicionar_linha_regra(self, nome_camada="", regra=""):
        linha_widget = QWidget()
        linha_layout = QHBoxLayout(linha_widget)
        linha_layout.setContentsMargins(0,0,0,0)

        edit_camada_alvo = QLineEdit(nome_camada)
        edit_camada_alvo.setPlaceholderText("Ex: Nome Completo")

        edit_regra_texto = QLineEdit(regra)
        edit_regra_texto.setPlaceholderText("Ex: {nome} {sobrenome}")

        btn_remover = QPushButton("X")
        btn_remover.setFixedSize(30, 30)
        btn_remover.setStyleSheet("color: white; background-color: #e63946; border-radius: 5px;")
        btn_remover.setToolTip("Remover esta regra")
        btn_remover.clicked.connect(partial(self.remover_linha, linha_widget))

        linha_layout.addWidget(edit_camada_alvo)
        linha_layout.addWidget(edit_regra_texto)
        linha_layout.addWidget(btn_remover)

        self.regras_layout.addWidget(linha_widget)
        self.linhas_widgets.append({
            'widget': linha_widget,
            'alvo': edit_camada_alvo,
            'regra': edit_regra_texto
        })

    def remover_linha(self, linha_widget):
        # Encontra o item para remover da lista de controle
        item_para_remover = next((item for item in self.linhas_widgets if item['widget'] == linha_widget), None)
        if item_para_remover:
            self.linhas_widgets.remove(item_para_remover)

        # Remove o widget do layout e o deleta
        linha_widget.deleteLater()

    def accept(self):
        novas_regras = {}
        nomes_alvo = []
        for linha in self.linhas_widgets:
            alvo = linha['alvo'].text().strip()
            regra = linha['regra'].text().strip()

            if alvo and regra: # Apenas salva se ambos os campos estiverem preenchidos
                if alvo in nomes_alvo:
                    QMessageBox.warning(self, "Alvo Duplicado", f"A camada alvo '{alvo}' foi definida mais de uma vez. Use nomes de camada únicos.")
                    return
                novas_regras[alvo] = regra
                nomes_alvo.append(alvo)

        self.regrasSalvas.emit(novas_regras)
        super().accept()