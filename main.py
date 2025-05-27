import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTranslator, QLocale, QLibraryInfo
from app_window import CartaoApp  # A classe da nossa janela principal.
from utils import set_dark_theme  # A função para o tema escuro.
from dialogo_gerenciar_regras import GerenciarRegrasDialog
from dialogo_regras_texto import GerenciarRegrasTextoDialog

# Ponto de partida oficial do programa
if __name__ == "__main__":
    # Cria a aplicação
    app = QApplication(sys.argv)

    # Configura a tradução do Qt
    qt_translator = QTranslator()
    translations_path = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
    locale_to_load = QLocale.Portuguese
    if qt_translator.load(locale_to_load, "qtbase", "_", translations_path):
        app.installTranslator(qt_translator)
    else:
        if qt_translator.load("qtbase_pt", "translations"):
            app.installTranslator(qt_translator)
        else:
            print("Falha ao carregar tradução do Qt para Português.")

    # Monta o programa, passo a passo
    set_dark_theme(app)  # Pinta o fundo
    window = CartaoApp()  # Constrói a janela
    window.show()  # Mostra a janela

    # Inicia o loop de eventos e encerra o programa quando a janela for fechada
    sys.exit(app.exec())