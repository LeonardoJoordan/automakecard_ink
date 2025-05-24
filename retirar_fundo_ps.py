import os
import time
import pyautogui
import pygetwindow as gw
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import subprocess
import sys
try:
    import pyperclip # Tenta importar para usar a área de transferência
except ImportError:
    pyperclip = None # Define como None se não estiver instalado
    print("AVISO: Biblioteca 'pyperclip' não encontrada. "
          "O script usará pyautogui.write(), que pode ser menos confiável "
          "para colar caminhos/nomes. Instale com: pip install pyperclip")

# --- Configurações Iniciais ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")
pyautogui.PAUSE = 0.2
pyautogui.FAILSAFE = True

# Função para obter caminho base (útil para PyInstaller)
def get_base_path():
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    else:
        return os.path.abspath(".")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Automação Photoshop v1.4 - Ativação Otimizada") # Mantendo v1.4
        self.geometry("600x450")
        self.resizable(False, False)

        self.pasta_imagens = ""
        self.processando = False
        self.pausado = False
        self.cancelado = False
        self.caminho_photoshop = r"C:\Program Files\Adobe\Adobe Photoshop 2024\Photoshop.exe"

        # --- Widgets da GUI (sem alterações) ---
        self.label_instrucao = ctk.CTkLabel(self, text="1. Verifique se o Photoshop está aberto.")
        self.label_instrucao.pack(pady=(10, 0))
        self.label_instrucao2 = ctk.CTkLabel(self, text="2. Selecione a pasta com as imagens:")
        self.label_instrucao2.pack(pady=(5, 5))
        self.frame_pasta = ctk.CTkFrame(self)
        self.frame_pasta.pack(pady=5, padx=20, fill="x")
        self.entry_pasta = ctk.CTkEntry(self.frame_pasta, placeholder_text="Nenhuma pasta selecionada", state="readonly", width=400)
        self.entry_pasta.pack(side="left", fill="x", expand=True, padx=(0,5))
        self.botao_selecionar = ctk.CTkButton(self.frame_pasta, text="Selecionar Pasta", command=self.selecionar_pasta, width=120)
        self.botao_selecionar.pack(side="left")
        self.frame_botoes = ctk.CTkFrame(self)
        self.frame_botoes.pack(pady=10)
        self.botao_iniciar = ctk.CTkButton(self.frame_botoes, text="Iniciar Automação", command=self.iniciar_thread)
        self.botao_iniciar.grid(row=0, column=0, padx=5, pady=5)
        self.botao_pausar = ctk.CTkButton(self.frame_botoes, text="Pausar", command=self.toggle_pausa, state="disabled")
        self.botao_pausar.grid(row=0, column=1, padx=5, pady=5)
        self.botao_cancelar = ctk.CTkButton(self.frame_botoes, text="Cancelar", command=self.cancelar_processo, state="disabled")
        self.botao_cancelar.grid(row=0, column=2, padx=5, pady=5)
        self.log_status = ctk.CTkTextbox(self, width=550, height=150, state="disabled")
        self.log_status.pack(pady=10, padx=20, fill="x")

    # --- Funções Auxiliares (sem alterações) ---
    def mostrar_mensagem(self, titulo, mensagem, tipo="info"):
        self.lift()
        if tipo == "erro": messagebox.showerror(titulo, mensagem, parent=self)
        elif tipo == "aviso": messagebox.showwarning(titulo, mensagem, parent=self)
        else: messagebox.showinfo(titulo, mensagem, parent=self)

    def log_via_after(self, message):
        if self.log_status: self.after(0, lambda: self._insert_log(message))

    def _insert_log(self, message):
        try:
            if self.log_status:
                self.log_status.configure(state="normal")
                self.log_status.insert("end", message + "\n")
                self.log_status.see("end")
                self.log_status.configure(state="disabled")
        except Exception as e: print(f"Erro ao inserir no log: {e}")

    def atualizar_estado_gui(self, processando):
        self.processando = processando
        state_normal = "normal"; state_disabled = "disabled"
        if processando:
            self.botao_iniciar.configure(state=state_disabled)
            self.botao_selecionar.configure(state=state_disabled)
            self.entry_pasta.configure(state=state_disabled)
            self.botao_pausar.configure(state=state_normal, text="Pausar" if not self.pausado else "Retomar")
            self.botao_cancelar.configure(state=state_normal)
        else:
            self.botao_iniciar.configure(state=state_normal)
            self.botao_selecionar.configure(state=state_normal)
            self.entry_pasta.configure(state="readonly")
            self.botao_pausar.configure(state=state_disabled, text="Pausar")
            self.botao_cancelar.configure(state=state_disabled)
            self.pausado = False

    # --- Funções de Controle (sem alterações) ---
    def selecionar_pasta(self):
        pasta_inicial = self.pasta_imagens if self.pasta_imagens else get_base_path()
        pasta = filedialog.askdirectory(title="Selecione a pasta com as imagens", initialdir=pasta_inicial)
        if pasta:
            self.pasta_imagens = pasta
            self.entry_pasta.configure(state="normal")
            self.entry_pasta.delete(0, "end"); self.entry_pasta.insert(0, pasta)
            self.entry_pasta.configure(state="readonly")
            self.log_via_after(f"Pasta selecionada: {pasta}")

    def toggle_pausa(self):
        if not self.processando: return
        self.pausado = not self.pausado
        texto_botao = "Retomar" if self.pausado else "Pausar"
        texto_status = "Processo pausado." if self.pausado else "Processo retomado."
        self.botao_pausar.configure(text=texto_botao)
        self.log_via_after(texto_status)

    def cancelar_processo(self):
        if self.processando and not self.cancelado:
            self.cancelado = True
            self.pausado = False
            self.log_via_after("Cancelamento solicitado. Aguardando fim da imagem atual...")
            self.botao_cancelar.configure(state="disabled")
            self.botao_pausar.configure(state="disabled", text="Pausar")

    def iniciar_thread(self):
        if not self.pasta_imagens or not os.path.isdir(self.pasta_imagens):
             self.mostrar_mensagem("Erro", "Selecione uma pasta de imagens válida.", tipo="erro"); return
        try:
            if not gw.getWindowsWithTitle("Adobe Photoshop"):
                 if not messagebox.askyesno("Photoshop Aberto?", "Photoshop não detectado. Continuar?", parent=self): return
        except Exception as e: self.log_via_after(f"Aviso: Verificação do Photoshop falhou ({e}).")

        self.cancelado = False; self.pausado = False
        self.atualizar_estado_gui(True)
        self.log_status.configure(state="normal"); self.log_status.delete("1.0", "end"); self.log_status.configure(state="disabled")
        self.log_via_after("Iniciando automação...")
        thread = threading.Thread(target=self.iniciar_automacao, daemon=True)
        thread.start()

    def ativar_janela_photoshop(self):
        # Função mantida, principalmente para uso ANTES do Ctrl+O
        try:
            photoshop_windows = gw.getWindowsWithTitle("Adobe Photoshop")
            if photoshop_windows:
                ps_window = photoshop_windows[0]
                if ps_window.isMinimized: ps_window.restore()
                ps_window.activate()
                time.sleep(0.5)
                return True
            else:
                # A mensagem de aviso só será mostrada se chamada antes do Ctrl+O e falhar
                self.log_via_after("AVISO: Janela Photoshop não encontrada (antes do Ctrl+O).")
                return False
        except Exception as e:
            self.log_via_after(f"Erro ao ativar janela PS: {e}")
            return False

    # --- Função Principal de Automação (COM ATIVAÇÃO OTIMIZADA) ---
    def iniciar_automacao(self):
        resultado_final_sucesso = True
        primeiro_arquivo = True
        try:
            arquivos = [f for f in os.listdir(self.pasta_imagens)
                        if f.lower().endswith((".jpg", ".jpeg", ".png"))]
            if not arquivos:
                self.log_via_after("Nenhuma imagem encontrada."); resultado_final_sucesso = False; return

            total_arquivos = len(arquivos)
            self.log_via_after(f"Encontradas {total_arquivos} imagens.")

            for i, arquivo in enumerate(arquivos, start=1):
                while self.pausado and not self.cancelado: time.sleep(0.5)
                if self.cancelado: self.log_via_after("Processo cancelado."); resultado_final_sucesso = False; break

                caminho_arquivo = os.path.join(self.pasta_imagens, arquivo)
                self.log_via_after(f"--- Processando {i}/{total_arquivos}: {arquivo} ---")

                try:
                    # 1. ATIVAR PS E ABRIR ARQUIVO (Única ativação necessária aqui)
                    if not self.ativar_janela_photoshop():
                        # Se falhar aqui, não tem como continuar
                        raise Exception("Falha ao ativar janela PS antes de abrir arquivo.")

                    pyautogui.hotkey('ctrl', 'o')
                    self.log_via_after("   - Diálogo 'Abrir' acionado...")
                    time.sleep(2.0) # <<<< AJUSTAR TEMPO ABERTURA DIÁLOGO >>>>

                    pasta = os.path.dirname(caminho_arquivo)
                    nome_arquivo = os.path.basename(caminho_arquivo)

                    if primeiro_arquivo:
                        # Lógica para o primeiro arquivo (Ctrl+L, Tabs, etc.)...
                        self.log_via_after("   - Primeiro arquivo: Navegando pasta e focando nome...")
                        pyautogui.hotkey('ctrl', 'l')
                        time.sleep(0.5)
                        if pyperclip: pyperclip.copy(pasta); time.sleep(0.2); pyautogui.hotkey('ctrl', 'v')
                        else: pyautogui.write(pasta)
                        time.sleep(0.5); pyautogui.press('enter')
                        self.log_via_after(f"   - Pasta '{pasta}' acessada.")
                        time.sleep(2.0) # <<<< AJUSTAR TEMPO CARGA PASTA >>>>

                        self.log_via_after("   - Pressionando TAB x7 para focar nome...") # Ajustado para 7 no código anterior
                        for _ in range(7):
                            pyautogui.press('tab')
                            time.sleep(0.1)
                        self.log_via_after("   - Foco no campo 'Nome' (esperado).")
                        time.sleep(0.3)

                        self.log_via_after(f"   - Inserindo nome: {nome_arquivo}...")
                        if pyperclip: pyperclip.copy(nome_arquivo); time.sleep(0.2); pyautogui.hotkey('ctrl', 'v')
                        else: pyautogui.write(nome_arquivo)
                        time.sleep(1)
                        pyautogui.press('enter')
                        primeiro_arquivo = False # Marca que o primeiro já foi
                    else:
                        # Lógica para arquivos subsequentes...
                        self.log_via_after("   - Arquivo subsequente: Inserindo nome direto...")
                        if pyperclip: pyperclip.copy(nome_arquivo); time.sleep(0.2); pyautogui.hotkey('ctrl', 'v')
                        else: pyautogui.write(nome_arquivo)
                        time.sleep(1)
                        pyautogui.press('enter')

                    # 2. AGUARDAR CARGA DA IMAGEM
                    self.log_via_after("   - Aguardando imagem carregar...")
                    time.sleep(8) # <<<<<<< AJUSTAR TEMPO DE CARGA DA IMAGEM >>>>>>>

                    # 3. EXECUTAR AÇÃO (F11) - SEM reativar a janela
                    self.log_via_after("   - Executando Ação (F11)...")
                    pyautogui.press("f11")
                    time.sleep(25) # <<<<<<< AJUSTAR TEMPO DA AÇÃO F11 >>>>>>>

                    # 4. FECHAR SEM SALVAR (Ctrl+W -> N) - SEM reativar a janela
                    self.log_via_after("   - Fechando sem salvar (Ctrl+W -> N)...")
                    pyautogui.hotkey("ctrl", "w")
                    time.sleep(2.0) # <<<< AJUSTAR ESPERA DIÁLOGO SALVAR >>>>
                    pyautogui.press("n") # Confirma "Não"
                    time.sleep(3) # <<<< AJUSTAR PAUSA ENTRE IMAGENS >>>>

                    self.log_via_after(f"   - Imagem {arquivo} processada.")

                # --- Tratamento de Erro por Imagem ---
                except Exception as e_img:
                    self.log_via_after(f"ERRO ao processar '{arquivo}': {e_img}")
                    self.log_via_after("   - Tentando fechar diálogos/janelas...")
                    # Tenta reativar AQUI porque um erro pode ter mudado o foco
                    if self.ativar_janela_photoshop():
                         pyautogui.press('esc', presses=2, interval=0.3)
                         time.sleep(0.5); pyautogui.hotkey('ctrl', 'w')
                         time.sleep(1); pyautogui.press('n'); time.sleep(1)
                    resultado_final_sucesso = False
                    # break # Descomentar para parar no primeiro erro

            # --- Fim do Loop FOR ---
        except Exception as e_main:
            self.log_via_after(f"ERRO GERAL: {e_main}")
            resultado_final_sucesso = False
        finally:
            # --- Finalização ---
            self.after(0, self.atualizar_estado_gui, False)
            msg_final = "interrompida." if self.cancelado else "concluída com sucesso!" if resultado_final_sucesso else "finalizada com erros."
            tipo_msg = "aviso" if self.cancelado or not resultado_final_sucesso else "info"
            self.log_via_after(f"\n==== Automação {msg_final} ====")
            self.after(50, self.mostrar_mensagem, "Fim do Processo", f"Automação {msg_final}", tipo_msg)

# --- Ponto de Entrada Principal ---
if __name__ == "__main__":
    app = App()
    app.mainloop()