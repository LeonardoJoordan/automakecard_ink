import os
import time
import pyautogui
import pygetwindow as gw # Adicionado
import customtkinter as ctk
from tkinter import filedialog, messagebox # Adicionado messagebox
import threading
import subprocess # Alternativa a os.system

# --- Configurações Iniciais ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")
pyautogui.PAUSE = 0.2 # Pequena pausa global entre comandos pyautogui
pyautogui.FAILSAFE = True # Mover mouse pro canto superior esquerdo para parar

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Automação Photoshop v1.1")
        self.geometry("600x450") # Aumentei um pouco altura
        self.resizable(False, False)

        self.pasta_imagens = ""
        self.processando = False # Flag para controlar estado
        self.pausado = False
        self.cancelado = False

        # --- Caminho do Photoshop (Verifique se está correto!) ---
        self.caminho_photoshop = r"C:\Program Files\Adobe\Adobe Photoshop 2024\Photoshop.exe"
        # --------------------------------------------------------

        # --- Widgets da GUI ---
        self.label_instrucao = ctk.CTkLabel(self, text="1. Verifique se o Photoshop está aberto.")
        self.label_instrucao.pack(pady=(10, 0))

        self.label_instrucao2 = ctk.CTkLabel(self, text="2. Selecione a pasta com as imagens:")
        self.label_instrucao2.pack(pady=(5, 5))

        # Frame para seleção de pasta
        self.frame_pasta = ctk.CTkFrame(self)
        self.frame_pasta.pack(pady=5, padx=20, fill="x")
        self.entry_pasta = ctk.CTkEntry(self.frame_pasta, placeholder_text="Nenhuma pasta selecionada", state="readonly", width=400)
        self.entry_pasta.pack(side="left", fill="x", expand=True, padx=(0,5))
        self.botao_selecionar = ctk.CTkButton(self.frame_pasta, text="Selecionar Pasta", command=self.selecionar_pasta, width=120)
        self.botao_selecionar.pack(side="left")

        # Frame para botões de controle
        self.frame_botoes = ctk.CTkFrame(self)
        self.frame_botoes.pack(pady=10)

        self.botao_iniciar = ctk.CTkButton(self.frame_botoes, text="Iniciar Automação", command=self.iniciar_thread)
        self.botao_iniciar.grid(row=0, column=0, padx=5, pady=5)

        self.botao_pausar = ctk.CTkButton(self.frame_botoes, text="Pausar", command=self.toggle_pausa, state="disabled")
        self.botao_pausar.grid(row=0, column=1, padx=5, pady=5)

        self.botao_cancelar = ctk.CTkButton(self.frame_botoes, text="Cancelar", command=self.cancelar_processo, state="disabled")
        self.botao_cancelar.grid(row=0, column=2, padx=5, pady=5)

        # Log Textbox
        self.log_status = ctk.CTkTextbox(self, width=550, height=150, state="disabled") # Começa desabilitado para escrita manual
        self.log_status.pack(pady=10, padx=20, fill="x")

    # --- Funções Auxiliares ---
    def mostrar_mensagem(self, titulo, mensagem, tipo="info"):
        """Mostra messagebox de forma segura."""
        self.lift() # Traz a janela principal para frente
        if tipo == "erro":
            messagebox.showerror(titulo, mensagem, parent=self)
        elif tipo == "aviso":
            messagebox.showwarning(titulo, mensagem, parent=self)
        else:
            messagebox.showinfo(titulo, mensagem, parent=self)

    def log_via_after(self, message):
        """Envia mensagem para o log de forma segura."""
        if self.log_status:
            self.after(0, lambda: self._insert_log(message))

    def _insert_log(self, message):
        """Método interno para inserir no log (chamado via after)."""
        if self.log_status:
            self.log_status.configure(state="normal") # Habilita escrita
            self.log_status.insert("end", message + "\n")
            self.log_status.see("end") # Rola para o final
            self.log_status.configure(state="disabled") # Desabilita escrita

    def atualizar_estado_gui(self, processando):
        """Atualiza o estado dos botões."""
        self.processando = processando
        if processando:
            self.botao_iniciar.configure(state="disabled")
            self.botao_selecionar.configure(state="disabled")
            self.entry_pasta.configure(state="disabled")
            self.botao_pausar.configure(state="normal", text="Pausar" if not self.pausado else "Retomar")
            self.botao_cancelar.configure(state="normal")
        else:
            self.botao_iniciar.configure(state="normal")
            self.botao_selecionar.configure(state="normal")
            self.entry_pasta.configure(state="readonly")
            self.botao_pausar.configure(state="disabled", text="Pausar")
            self.botao_cancelar.configure(state="disabled")
            self.pausado = False # Reseta pausa ao finalizar

    # --- Funções de Controle ---
    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com as imagens")
        if pasta:
            self.pasta_imagens = pasta
            # Atualiza Entry em vez de Label
            self.entry_pasta.configure(state="normal")
            self.entry_pasta.delete(0, "end")
            self.entry_pasta.insert(0, pasta)
            self.entry_pasta.configure(state="readonly")
            self.log_via_after(f"Pasta selecionada: {pasta}")

    def toggle_pausa(self):
        if not self.processando: return # Não faz nada se não estiver processando
        self.pausado = not self.pausado
        texto_botao_pausa = "Retomar" if self.pausado else "Pausar"
        texto_status = "Processo pausado." if self.pausado else "Processo retomado."
        self.botao_pausar.configure(text=texto_botao_pausa)
        self.log_via_after(texto_status)

    def cancelar_processo(self):
        if self.processando and not self.cancelado:
            self.cancelado = True
            self.pausado = False # Libera de possível pausa
            self.log_via_after("Cancelamento solicitado. Aguardando fim da imagem atual...")
            self.botao_cancelar.configure(state="disabled") # Desabilita para não clicar de novo
            self.botao_pausar.configure(state="disabled") # Desabilita pausa

    def iniciar_thread(self):
        """Valida e inicia a thread de automação."""
        if not self.pasta_imagens:
            self.mostrar_mensagem("Erro", "Nenhuma pasta de imagens foi selecionada.", tipo="erro")
            return
        if not os.path.isdir(self.pasta_imagens):
             self.mostrar_mensagem("Erro", f"A pasta selecionada não é válida:\n{self.pasta_imagens}", tipo="erro")
             return
        # Opcional: Verificar se Photoshop.exe existe
        # if not os.path.exists(self.caminho_photoshop):
        #     self.mostrar_mensagem("Erro", f"Executável do Photoshop não encontrado em:\n{self.caminho_photoshop}", tipo="erro")
        #     return

        # Verificar se o Photoshop parece estar aberto (Requer pygetwindow)
        try:
            if not gw.getWindowsWithTitle("Adobe Photoshop"):
                 resposta = messagebox.askyesno("Photoshop Aberto?", "O Photoshop não parece estar aberto.\n\nDeseja continuar mesmo assim?\n(A automação provavelmente falhará)", parent=self)
                 if not resposta:
                     return # Cancela se o usuário disser não
        except Exception as e:
            self.log_via_after(f"Aviso: Não foi possível verificar se o Photoshop está aberto ({e}).")


        self.cancelado = False
        self.pausado = False
        self.atualizar_estado_gui(True) # Atualiza botões para estado "processando"
        self.log_status.configure(state="normal")
        self.log_status.delete("1.0", "end") # Limpa log
        self.log_status.configure(state="disabled")
        self.log_via_after("Iniciando automação...")

        # Inicia a thread
        thread = threading.Thread(target=self.iniciar_automacao, daemon=True)
        thread.start()

    def ativar_janela_photoshop(self):
        """Tenta encontrar e ativar a janela do Photoshop. Retorna True se sucesso."""
        try:
            photoshop_windows = gw.getWindowsWithTitle("Adobe Photoshop")
            if photoshop_windows:
                ps_window = photoshop_windows[0]
                # Traz para frente e ativa
                if ps_window.isMinimized:
                    ps_window.restore()
                ps_window.activate()
                time.sleep(0.5) # Pequena pausa para garantir ativação
                return True
            else:
                self.log_via_after("AVISO: Janela do Photoshop não encontrada.")
                return False
        except Exception as e:
            # Pode falhar se não tiver permissão, etc.
            self.log_via_after(f"Erro ao tentar ativar janela PS: {e}")
            return False

    # --- Função Principal de Automação (executa na thread) ---
    def iniciar_automacao(self):
        resultado_final_sucesso = True # Assume sucesso inicialmente
        try:
            arquivos = [f for f in os.listdir(self.pasta_imagens) if f.lower().endswith((".jpg", ".jpeg", ".png"))]

            if not arquivos:
                self.log_via_after("Nenhuma imagem (.jpg, .jpeg, .png) encontrada na pasta.")
                resultado_final_sucesso = False # Termina sem sucesso
                return # Sai da função

            total_arquivos = len(arquivos)
            self.log_via_after(f"Encontradas {total_arquivos} imagens para processar.")

            for i, arquivo in enumerate(arquivos, start=1):
                # --- Checagem de Pausa e Cancelamento ---
                while self.pausado and not self.cancelado:
                    time.sleep(0.5) # Espera enquanto pausado

                if self.cancelado:
                    self.log_via_after("Processo cancelado pelo usuário.")
                    resultado_final_sucesso = False
                    break # Sai do loop FOR

                # --- Processamento da Imagem Atual ---
                caminho_arquivo = os.path.join(self.pasta_imagens, arquivo)
                self.log_via_after(f"--- Processando Imagem {i}/{total_arquivos}: {arquivo} ---")

                try:
                    # 1. Ativar Photoshop e Abrir Imagem (usando Ctrl+O)
                    if not self.ativar_janela_photoshop():
                        raise Exception("Falha ao ativar janela do Photoshop.") # Pula para o except

                    pyautogui.hotkey('ctrl', 'o')
                    time.sleep(1.0) # Espera diálogo abrir (AJUSTAR SE NECESSÁRIO)
                    pyautogui.write(caminho_arquivo)
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    self.log_via_after("   - Aguardando imagem carregar...")
                    # !! TEMPO CRÍTICO !! Ajustar conforme necessário!
                    time.sleep(10) # <<<<<<< AJUSTAR TEMPO DE CARGA DA IMAGEM >>>>>>>

                    # 2. Reativar Janela e Rodar Ação (F11)
                    if not self.ativar_janela_photoshop():
                         raise Exception("Falha ao reativar PS antes de rodar Ação.")

                    self.log_via_after("   - Executando Ação (F11)...")
                    pyautogui.press("f11")
                    # !! TEMPO CRÍTICO !! Ajustar conforme necessário!
                    time.sleep(15) # <<<<<<< AJUSTAR TEMPO DE EXECUÇÃO DA AÇÃO F11 >>>>>>>

                    # 3. Reativar Janela e Fechar Sem Salvar (Ctrl+W -> N)
                    if not self.ativar_janela_photoshop():
                        raise Exception("Falha ao reativar PS antes de fechar.")

                    self.log_via_after("   - Fechando imagem sem salvar (Ctrl+W -> N)...")
                    pyautogui.hotkey("ctrl", "w")
                    time.sleep(1.0) # Espera diálogo de salvar (AJUSTAR SE NECESSÁRIO)

                    # !! ATENÇÃO AO IDIOMA DO PHOTOSHOP !!
                    # 'n' geralmente é para "Não" ou "No". Pode mudar!
                    pyautogui.press("n")
                    time.sleep(1.5) # Pausa antes da próxima imagem (AJUSTAR SE NECESSÁRIO)

                    self.log_via_after(f"   - Imagem {arquivo} processada.")

                except Exception as e_img:
                    self.log_via_after(f"ERRO ao processar '{arquivo}': {e_img}")
                    self.log_via_after("   - Tentando fechar possíveis janelas/diálogos abertos...")
                    # Tenta fechar qualquer diálogo ou a imagem de forma segura
                    if self.ativar_janela_photoshop():
                        pyautogui.press('esc') # Tenta fechar diálogos
                        time.sleep(0.5)
                        pyautogui.hotkey('ctrl', 'w') # Tenta fechar a imagem
                        time.sleep(1)
                        pyautogui.press('n') # Tenta dizer não salvar
                        time.sleep(1)
                    resultado_final_sucesso = False # Marca que houve erro
                    # Decidir se quer continuar ou parar tudo
                    # break # Descomente para parar tudo no primeiro erro

            # --- Fim do Loop ---

        except Exception as e_main:
            # Erro geral fora do loop de arquivos (ex: listar diretório)
            self.log_via_after(f"ERRO GERAL na automação: {e_main}")
            resultado_final_sucesso = False
        finally:
            # --- Finalização (executa sempre, mesmo com erro ou cancelamento) ---
            # Usa self.after para chamar a função de volta na thread principal da GUI
            self.after(0, self.atualizar_estado_gui, False) # False = não está mais processando
            final_msg = "Automação concluída." if resultado_final_sucesso and not self.cancelado else "Automação finalizada com avisos/erros." if not self.cancelado else "Automação interrompida."
            self.log_via_after(f"\n==== {final_msg} ====")
            if resultado_final_sucesso and not self.cancelado :
                 self.after(10, self.mostrar_mensagem, "Sucesso", "Automação concluída com sucesso!", "info")
            elif not resultado_final_sucesso and not self.cancelado:
                 self.after(10, self.mostrar_mensagem, "Atenção", "Automação finalizada, mas ocorreram erros.\nVerifique o log para detalhes.", "aviso")


if __name__ == "__main__":
    # Verifica se o caminho do Photoshop existe (opcional, mas útil)
    photoshop_path = r"C:\Program Files\Adobe\Adobe Photoshop 2024\Photoshop.exe"
    if not os.path.exists(photoshop_path):
         messagebox.showwarning("Photoshop não encontrado", f"O caminho configurado para o Photoshop não foi encontrado:\n{photoshop_path}\n\nA automação pode não funcionar corretamente ao tentar abrir arquivos.")

    app = App()
    # Sobrescreve o caminho se necessário (ex: se carregado de um config)
    app.caminho_photoshop = photoshop_path
    app.mainloop()