# dialogo_regras_texto.py (Layout do Usuário Adaptado)
from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QTextEdit, QPushButton, QScrollArea, QWidget, QDialogButtonBox
)
from PySide6.QtCore import Qt, Signal
import sys  # Necessário para o bloco de teste __main__
# Bloco corrigido
from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QTextEdit, QPushButton, QScrollArea, QWidget, QDialogButtonBox,
    QSizePolicy  # <-- ESTA É A LINHA QUE FALTAVA
)


class GerenciarRegrasTextoDialog(QDialog):
    regrasSalvas = Signal(dict)

    def __init__(self, dados_especificos_disponiveis, regras_atuais, parent=None):  # Parâmetro renomeado
        super().__init__(parent)
        self.setWindowTitle("Gerenciar Regras de Texto")
        self.setMinimumSize(850, 600)
        self.dados_especificos = dados_especificos_disponiveis  # Usando o nome consistente

        main_layout = QVBoxLayout(self)

        # Variáveis disponíveis
        variaveis_formatadas = ", ".join([f"{{{d}}}" for d in self.dados_especificos])
        info_label = QLabel(f"<b>Variáveis disponíveis para usar nas regras:</b><br>{variaveis_formatadas}")
        info_label.setWordWrap(True)
        main_layout.addWidget(info_label)

        # Cabeçalho
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("<b>Camada Alvo (Dado Específico)</b>"), 1)
        header_layout.addWidget(QLabel("<b>Conteúdo / Regra de Formatação</b>"), 5)
        header_layout.addSpacing(35)  # Espaço para o botão de limpar
        main_layout.addLayout(header_layout)

        # Área de rolagem
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_widget = QWidget()
        self.regras_layout = QVBoxLayout(scroll_widget)
        self.regras_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        scroll_area.setWidget(scroll_widget)
        main_layout.addWidget(scroll_area)

        self.linhas_widgets = {}  # Armazena referências aos widgets de cada linha
        # Cria uma linha para cada "Dado Específico"
        for nome_dado in self.dados_especificos:
            regra_existente = regras_atuais.get(nome_dado, "")
            self.adicionar_linha_grid(nome_dado, regra_existente)

        # Botões de ação
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)

    def adicionar_linha_grid(self, nome_dado, regra_texto):
        linha_widget = QWidget()  # Widget container para a linha
        linha_layout = QHBoxLayout(linha_widget)
        linha_layout.setContentsMargins(0, 0, 0, 0)  # Sem margens internas

        # "Dado Específico" (label estilizado como botão não clicável)
        label_dado_especifico = QPushButton(nome_dado)
        label_dado_especifico.setEnabled(False)  # Não clicável
        label_dado_especifico.setStyleSheet(
            "color: #FFFFFF; background-color: #33373B; border: 1px solid #4A4F54;"
            "font-size: 14px; border-radius: 5px; padding: 5px; min-height: 30px;"
            "text-align: center;"  # Garante que o texto do botão esteja centralizado
        )
        # Ajusta a política de tamanho para que o botão não expanda demais
        label_dado_especifico.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)

        # Campo de edição para a regra de texto
        edit_regra = QTextEdit(regra_texto)
        edit_regra.setPlaceholderText(f"Digite a regra para {{{nome_dado}}} (ou deixe em branco)")
        edit_regra.setFixedHeight(70)  # Altura para visualização de múltiplas linhas

        # Botão para limpar (excluir) o conteúdo da regra
        btn_limpar_regra = QPushButton("X")
        btn_limpar_regra.setFixedSize(28, 28)  # Tamanho fixo para o botão
        btn_limpar_regra.setStyleSheet(
            "color: white; background-color: #C94444; border-radius: 5px; font-weight: bold;"
        )
        btn_limpar_regra.setToolTip(f"Limpar a regra para '{nome_dado}'")
        # Conecta o clique do botão para limpar o QTextEdit correspondente
        btn_limpar_regra.clicked.connect(lambda: edit_regra.setPlainText(""))

        # Adiciona os widgets ao layout da linha com proporções
        linha_layout.addWidget(label_dado_especifico, 2)  # Proporção 2 para o label
        linha_layout.addWidget(edit_regra, 5)  # Proporção 5 para o campo de texto
        linha_layout.addWidget(btn_limpar_regra)  # Botão de limpar ocupa espaço natural

        # Adiciona a linha completa ao layout principal de regras
        self.regras_layout.addWidget(linha_widget)
        # Guarda a referência ao QTextEdit para fácil acesso ao salvar
        self.linhas_widgets[nome_dado] = {'widget': linha_widget, 'edit': edit_regra}

    def accept(self):
        novas_regras = {}
        # Itera sobre os "Dados Específicos" para os quais criamos linhas
        for nome_dado, refs in self.linhas_widgets.items():
            texto_regra = refs['edit'].toPlainText().strip()
            if texto_regra:  # Salva a regra apenas se ela não estiver vazia
                novas_regras[nome_dado] = texto_regra

        self.regrasSalvas.emit(novas_regras)
        super().accept()


# ====== BLOCO DE TESTE (MANTIDO DO SEU EXEMPLO) ======
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Simula os dados que seriam passados pelo app_window.py
    dados_para_teste = ["Nome", "Conjuge", "Data", "Endereço", "País", "Loja", "Carro"]
    regras_iniciais_teste = {
        "Nome": "Sr(a). {Nome}",
        "Conjuge": "Casado(a) com {Conjuge} em {Data}",
        "Data": ""  # Exemplo de regra vazia que não será salva
    }

    dlg = GerenciarRegrasTextoDialog(dados_para_teste, regras_iniciais_teste)

    # Para ver o que seria salvo:
    # def mostrar_regras_salvas(regras):
    # print("Regras a serem salvas:", regras)
    # dlg.regrasSalvas.connect(mostrar_regras_salvas)

    if dlg.exec():
        print("Diálogo fechado com 'Salvar'.")
        # O sinal 'regrasSalvas' já teria sido emitido e processado
        # pela função conectada em um app real.
    else:
        print("Diálogo fechado com 'Cancelar' ou 'X'.")

    sys.exit(app.exec())
