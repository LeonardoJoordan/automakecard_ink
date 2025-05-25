from PySide6.QtWidgets import (
    QDialog, QHBoxLayout, QPushButton, QSizePolicy, QApplication
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont


class CaixaModificarModeloDialog(QDialog):
    def __init__(self, nome_modelo_atual, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Modificar: {nome_modelo_atual}")
        self.setModal(True)
        self.setFixedSize(410, 230)

        # Variável para armazenar a escolha do utilizador
        self.escolha = None

        # Layout principal
        layout = QHBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(12, 24, 12, 12)

        # Estilo dos botões
        estilo_botao = """
        QPushButton {
            background-color: #444;
            color: white;
            font-size: 16px;
            font-weight: bold;
            border-radius: 16px;
            border: 2px solid #666;
            padding: 20px 10px;
            min-height: 100px;
        }
        QPushButton:hover {
            background-color: #5c5c5c;
            border-color: #7b7b7b;
        }
        QPushButton:pressed {
            background-color: #333;
            border-color: #555;
        }
        """

        # Botão para alterar apenas camadas
        self.btn_alterar_camadas = QPushButton("Alterar\nCamadas")
        self.btn_alterar_camadas.setStyleSheet(estilo_botao)
        self.btn_alterar_camadas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.btn_alterar_camadas.setToolTip("Apenas modificar a configuração das camadas do modelo existente")

        # Botão para alterar arquivo
        self.btn_alterar_arquivo = QPushButton("Alterar\nArquivo")
        self.btn_alterar_arquivo.setStyleSheet(estilo_botao)
        self.btn_alterar_arquivo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.btn_alterar_arquivo.setToolTip("Substituir o ficheiro PSD e reconfigurar o modelo")

        # Adiciona os botões ao layout
        layout.addWidget(self.btn_alterar_camadas)
        layout.addWidget(self.btn_alterar_arquivo)

        # Estilo do diálogo
        self.setStyleSheet("""
            QDialog {
                background-color: #232323;
                border-radius: 12px;
                border: 1px solid #444;
            }
        """)

        # Conecta os sinais
        self.btn_alterar_camadas.clicked.connect(self.on_alterar_camadas)
        self.btn_alterar_arquivo.clicked.connect(self.on_alterar_arquivo)

    def on_alterar_camadas(self):
        """Define a escolha como 'camadas' e aceita o diálogo"""
        self.escolha = 'camadas'
        self.accept()

    def on_alterar_arquivo(self):
        """Define a escolha como 'arquivo' e aceita o diálogo"""
        self.escolha = 'arquivo'
        self.accept()

    def exec(self):
        """
        Sobrescreve o método exec para garantir que escolha seja None se cancelado
        """
        self.escolha = None
        return super().exec()


# TESTE ISOLADO (pode manter para testar o visual do diálogo separadamente):
if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)

    # Para testar, passamos um nome de modelo fictício:
    dialogo = CaixaModificarModeloDialog("exemplo_modelo.psd")

    if dialogo.exec():
        print(f"Escolha do utilizador: {dialogo.escolha}")
        if dialogo.escolha == 'camadas':
            print("Utilizador escolheu: Alterar Camadas")
        elif dialogo.escolha == 'arquivo':
            print("Utilizador escolheu: Alterar Arquivo")
    else:
        print("Diálogo cancelado ou fechado pelo utilizador")

    app.quit()