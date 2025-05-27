from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QTextEdit, QPushButton, QScrollArea, QWidget, QDialogButtonBox
)
from PySide6.QtCore import Qt, Signal
import sys

class GerenciarRegrasTextoDialog(QDialog):
    regrasSalvas = Signal(dict)

    def __init__(self, placeholders, regras_atuais, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Gerenciar Regras de Texto")
        self.setMinimumSize(850, 600)
        self.placeholders = placeholders

        main_layout = QVBoxLayout(self)

        # Variáveis disponíveis
        variaveis_formatadas = ", ".join([f"{{{d}}}" for d in self.placeholders])
        info_label = QLabel(f"<b>Variáveis disponíveis para usar nas regras:</b><br>{variaveis_formatadas}")
        info_label.setWordWrap(True)
        main_layout.addWidget(info_label)

        # Cabeçalho
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("<b>Camada Alvo (Placeholder)</b>"), 1)
        header_layout.addWidget(QLabel("<b>Conteúdo / Regra de Formatação</b>"), 5)
        main_layout.addLayout(header_layout)

        # Área de rolagem
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_widget = QWidget()
        self.regras_layout = QVBoxLayout(scroll_widget)
        self.regras_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        scroll_area.setWidget(scroll_widget)
        main_layout.addWidget(scroll_area)

        self.linhas_widgets = {}
        for nome in self.placeholders:
            regra = regras_atuais.get(nome, "")
            self.adicionar_linha_grid(nome, regra)

        # Botões de ação
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)

    def adicionar_linha_grid(self, nome, regra):
        linha_widget = QWidget()
        linha_layout = QHBoxLayout(linha_widget)
        linha_layout.setContentsMargins(0, 0, 0, 0)

        # Placeholder (label como botão estilizado)
        label_placeholder = QPushButton(nome)
        label_placeholder.setEnabled(False)
        label_placeholder.setStyleSheet(
            "color: #FFF; background-color: #282828; border: 1px solid #333;"
            "font-size: 16px; border-radius: 6px; min-width: 120px; min-height: 40px;"
        )

        # Campo de edição
        edit_regra = QTextEdit(regra)
        edit_regra.setPlaceholderText(f"Digite a regra para {{{nome}}}...")
        edit_regra.setFixedHeight(90)

        # Botão de excluir (limpar)
        btn_excluir = QPushButton("X")
        btn_excluir.setFixedSize(28, 28)
        btn_excluir.setStyleSheet("color: white; background-color: #e63946; border-radius: 7px;")
        btn_excluir.setToolTip("Excluir esta regra")
        btn_excluir.clicked.connect(lambda: edit_regra.setPlainText(""))

        linha_layout.addWidget(label_placeholder, 1)
        linha_layout.addWidget(edit_regra, 5)
        linha_layout.addWidget(btn_excluir)

        self.regras_layout.addWidget(linha_widget)
        self.linhas_widgets[nome] = {'widget': linha_widget, 'edit': edit_regra}

    def accept(self):
        novas_regras = {}
        for nome, refs in self.linhas_widgets.items():
            texto = refs['edit'].toPlainText().strip()
            if texto:
                novas_regras[nome] = texto
        self.regrasSalvas.emit(novas_regras)
        super().accept()

# ====== MAIN PARA TESTAR ======
if __name__ == "__main__":
    app = QApplication(sys.argv)
    placeholders = ["Nome", "Conjuge", "Data", "Endereço", "país", "loja", "carro"]
    regras_iniciais = {
        "Nome": "Sr(a). {Nome}",
        "Conjuge": "Casado(a) com {Conjuge} em {Data}",
        "Data": ""
    }
    dlg = GerenciarRegrasTextoDialog(placeholders, regras_iniciais)
    dlg.exec()
