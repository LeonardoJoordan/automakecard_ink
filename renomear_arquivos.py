import os
import shutil
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import sys

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

def get_base_path():
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    else:
        return os.path.abspath(".")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Renomear e Copiar Imagens v2.1") # Versão incrementada
        self.geometry("550x350")
        self.resizable(False, False)

        # --- Configurar Ícone (Opcional) ---
        # try:
        #     icon_path = os.path.join(get_base_path(), "icon.ico")
        #     if os.path.exists(icon_path):
        #         self.iconbitmap(icon_path)
        # except Exception as e:
        #     print(f"Erro ao definir ícone: {e}")
        # -----------------------------------

        self.pasta_origem = ctk.StringVar()
        self.pasta_destino = ctk.StringVar()

        # --- Layout ---
        self.label_origem = ctk.CTkLabel(self, text="1. Selecione a pasta principal (Origem):")
        self.label_origem.pack(pady=(15, 0), padx=20, anchor="w")

        self.frame_origem = ctk.CTkFrame(self)
        self.frame_origem.pack(pady=5, padx=20, fill="x")

        self.entry_origem = ctk.CTkEntry(self.frame_origem, textvariable=self.pasta_origem, state="readonly") # Readonly para evitar digitação direta
        self.entry_origem.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.botao_selecionar_origem = ctk.CTkButton(self.frame_origem, text="Selecionar", command=self.selecionar_pasta_origem, width=100)
        self.botao_selecionar_origem.pack(side="left")

        self.label_destino = ctk.CTkLabel(self, text="2. Selecione a pasta de Destino:")
        self.label_destino.pack(pady=(10, 0), padx=20, anchor="w")

        self.frame_destino = ctk.CTkFrame(self)
        self.frame_destino.pack(pady=5, padx=20, fill="x")

        self.entry_destino = ctk.CTkEntry(self.frame_destino, textvariable=self.pasta_destino, state="readonly") # Readonly
        self.entry_destino.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.botao_selecionar_destino = ctk.CTkButton(self.frame_destino, text="Selecionar", command=self.selecionar_pasta_destino, width=100)
        self.botao_selecionar_destino.pack(side="left")

        self.label_status = ctk.CTkLabel(self, text="")
        self.label_status.pack(pady=(15,5))

        self.botao_iniciar = ctk.CTkButton(self, text="Iniciar Cópia", command=self.iniciar_processamento_wrapper)
        self.botao_iniciar.pack(pady=5)

        self.botao_sair = ctk.CTkButton(self, text="Sair", command=self.destroy, fg_color="grey")
        self.botao_sair.pack(pady=(5, 15))

    def selecionar_pasta_origem(self):
        """Abre diálogo para selecionar pasta de ORIGEM."""
        # Guarda o estado atual para restaurar se o diálogo for cancelado
        pasta_anterior = self.pasta_origem.get()
        pasta = filedialog.askdirectory(title="Selecione a Pasta de Origem", initialdir=pasta_anterior if pasta_anterior else None)
        if pasta: # Somente atualiza se uma pasta foi selecionada
            self.pasta_origem.set(pasta)
            # self.atualizar_status("") # <-- REMOVIDO

    def selecionar_pasta_destino(self):
        """Abre diálogo para selecionar pasta de DESTINO."""
        pasta_anterior = self.pasta_destino.get()
        pasta = filedialog.askdirectory(title="Selecione a Pasta de Destino", initialdir=pasta_anterior if pasta_anterior else None)
        if pasta: # Somente atualiza se uma pasta foi selecionada
            self.pasta_destino.set(pasta)
            # self.atualizar_status("") # <-- REMOVIDO

    def mostrar_mensagem(self, titulo, mensagem, tipo="info"):
        """Função auxiliar para mostrar messagebox na thread principal."""
        # Garante que a messagebox apareça sobre a janela principal
        self.lift()
        self.focus_force()
        if tipo == "erro":
            messagebox.showerror(titulo, mensagem, parent=self)
        elif tipo == "aviso":
            messagebox.showwarning(titulo, mensagem, parent=self)
        else:
            messagebox.showinfo(titulo, mensagem, parent=self)
        self.focus_force() # Tenta garantir que a janela mantenha o foco

    def atualizar_status(self, texto, reabilitar_botao=False):
        """Função auxiliar para atualizar a GUI (label e botões) na thread principal."""
        self.label_status.configure(text=texto)
        if reabilitar_botao:
            # Reabilita botões
            self.botao_iniciar.configure(state="normal", text="Iniciar Cópia")
            self.botao_selecionar_origem.configure(state="normal")
            self.botao_selecionar_destino.configure(state="normal")
            self.entry_origem.configure(state="readonly") # Mantém readonly
            self.entry_destino.configure(state="readonly") # Mantém readonly
        else:
            # Desabilita botões durante o processamento
             self.botao_iniciar.configure(state="disabled", text="Processando...")
             self.botao_selecionar_origem.configure(state="disabled")
             self.botao_selecionar_destino.configure(state="disabled")
             # Pode ser útil desabilitar as entries também visualmente
             self.entry_origem.configure(state="disabled")
             self.entry_destino.configure(state="disabled")


    def iniciar_processamento_wrapper(self):
        """Função chamada pelo botão que valida e inicia a thread."""
        origem = self.pasta_origem.get()
        destino = self.pasta_destino.get()

        # Validação
        if not origem or not destino:
            self.mostrar_mensagem("Erro de Entrada", "Por favor, selecione a pasta de Origem e a pasta de Destino.", "erro")
            return

        if os.path.abspath(origem) == os.path.abspath(destino):
             self.mostrar_mensagem("Erro de Lógica", "A pasta de Origem e a pasta de Destino não podem ser a mesma.", "erro")
             return

        # Verifica se a origem é uma pasta válida
        if not os.path.isdir(origem):
            self.mostrar_mensagem("Erro de Origem", f"A pasta de origem selecionada não é válida ou não existe:\n{origem}", "erro")
            return

        # Verifica se o destino é uma pasta válida (ou pode ser criada)
        # É um pouco mais complexo garantir que o *caminho pai* exista se o destino não existe ainda.
        # Simplificação: Apenas verifica se o destino existe e *não* é uma pasta.
        if os.path.exists(destino) and not os.path.isdir(destino):
             self.mostrar_mensagem("Erro de Destino", f"Já existe um ARQUIVO com o nome da pasta de destino selecionada:\n{destino}\n\nPor favor, escolha outro local ou nome.", "erro")
             return


        # Desabilitar botões e mostrar status ANTES de iniciar a thread
        self.atualizar_status("Iniciando processamento...", reabilitar_botao=False)

        # Criar e iniciar a thread
        thread = threading.Thread(target=self.processar_arquivos, args=(origem, destino), daemon=True)
        thread.start()

    def processar_arquivos(self, origem, destino_base):
        """Função que executa o trabalho pesado em uma thread separada."""
        try:
            # Garante que a pasta de destino exista (cria se necessário)
            # Esta linha pode gerar um OSError se não houver permissão ou o caminho for inválido
            os.makedirs(destino_base, exist_ok=True)

            # Atualiza status via self.after para rodar na thread principal
            self.after(0, self.atualizar_status, "Verificando pastas e copiando arquivos...", False) # Mantém desabilitado

            arquivos_copiados = 0
            arquivos_ignorados = 0
            erros_copia = 0
            destino_abs = os.path.abspath(destino_base) # Cache do caminho absoluto

            # Itera pela árvore de diretórios da origem
            for root, dirs, files in os.walk(origem, topdown=True):
                root_abs = os.path.abspath(root)

                # Impede os.walk de entrar na pasta de destino SE ela for subpasta da origem
                # Usar list comprehension para modificar dirs[:] in-place
                dirs[:] = [d for d in dirs if os.path.abspath(os.path.join(root, d)) != destino_abs]

                # Pula a iteração se o diretório atual for a própria pasta de destino
                # (Segurança extra, a linha acima deve prevenir a entrada)
                if root_abs == destino_abs:
                    continue

                nome_pasta_origem = os.path.basename(root) # Nome da pasta pai na origem

                for file in files:
                    # Verifica extensão em minúsculas
                    if file.lower().endswith(('.jpg', '.jpeg', '.png')):
                        try:
                            # Monta o novo nome: pastaPaiOrigem_nomeArquivoOriginal.ext
                            novo_nome = f"{nome_pasta_origem}_{file}"
                            caminho_origem_completo = os.path.join(root, file)
                            caminho_destino_base_arq = os.path.join(destino_base, novo_nome)
                            caminho_destino_final = caminho_destino_base_arq # Assume não haver colisão inicialmente

                            # Lógica para evitar sobrescrever arquivos existentes no destino
                            contador = 1
                            # Separa nome base e extensão para adicionar o contador corretamente
                            base_nome, extensao = os.path.splitext(novo_nome)
                            # Verifica se o caminho final já existe
                            while os.path.exists(caminho_destino_final):
                                # Cria novo nome com contador: nomeBase_contador.ext
                                novo_nome_contador = f"{base_nome}_{contador}{extensao}"
                                caminho_destino_final = os.path.join(destino_base, novo_nome_contador)
                                contador += 1
                            # --- Fim da lógica de colisão ---

                            # Copia o arquivo preservando metadados (data de modificação, etc.)
                            shutil.copy2(caminho_origem_completo, caminho_destino_final)
                            arquivos_copiados += 1
                        except Exception as e: # Captura qualquer erro durante a cópia individual
                            erros_copia += 1
                            # Log detalhado no console é importante para depuração
                            print(f"ERRO ao copiar '{caminho_origem_completo}' para '{caminho_destino_final}': {e}")
                            # Você pode adicionar uma lista de erros para mostrar ao usuário no final, se preferir
                    else:
                        arquivos_ignorados += 1 # Conta arquivos que não são das extensões desejadas

            # --- Processamento Concluído ---
            # Monta a mensagem final de forma mais detalhada
            msg_final = f"Processo Concluído!\n\n"
            msg_final += f"  - Imagens copiadas: {arquivos_copiados}\n"
            if arquivos_ignorados > 0:
                msg_final += f"  - Arquivos ignorados (não imagem): {arquivos_ignorados}\n"
            if erros_copia > 0:
                msg_final += f"  - Erros durante a cópia: {erros_copia} (verifique o console/terminal para detalhes)\n"
                tipo_msg = "aviso" # Mostra como aviso se houve erros
            else:
                 tipo_msg = "info" # Mostra como informação se tudo correu bem

            # Agenda a atualização da GUI na thread principal
            # Reabilita os botões
            self.after(0, self.atualizar_status, "Processo finalizado.", True)
            # Mostra a mensagem de resultado (com um pequeno delay para garantir que a label atualizou)
            self.after(50, self.mostrar_mensagem, "Resultado da Cópia", msg_final, tipo_msg)

        except OSError as e: # Erro específico de criação de diretório/permissão na pasta de destino
             print(f"Erro Crítico de Sistema de Arquivos (OSError): {e}")
             # Tenta reabilitar botões e mostrar erro na GUI
             self.after(0, self.atualizar_status, f"Erro ao acessar/criar pasta de destino!", True)
             self.after(10, self.mostrar_mensagem, "Erro Crítico", f"Não foi possível criar ou acessar a pasta de destino:\n'{destino_base}'\n\nVerifique as permissões ou se o caminho é válido.\n\nDetalhe: {e}", "erro")
        except Exception as e:
            # Captura qualquer outro erro inesperado durante o processo principal
            print(f"Erro Geral Inesperado: {e}")
             # Tenta reabilitar botões e mostrar erro na GUI
            self.after(0, self.atualizar_status, "Ocorreu um erro inesperado.", True)
            self.after(10, self.mostrar_mensagem, "Erro Inesperado", f"Ocorreu um erro inesperado durante o processamento:\n{e}", "erro")


if __name__ == "__main__":
    app = App()
    app.mainloop()