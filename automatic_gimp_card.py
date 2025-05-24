import customtkinter as ctk
import os
import shutil

from PIL import Image, ImageTk
import subprocess # Mantenha o subprocess, pois vamos usá-lo para ImageMagick
from tkinter import filedialog, messagebox
import tkinter as tk
from tksheet import Sheet

from PIL import ImageDraw, ImageFont



from PySide6.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QMessageBox
from PySide6.QtCore import Qt
from PySide6.QtGui import QKeySequence

import sys

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class CartaoApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gerador de Cartões Personalizados")
        self.geometry("1000x510")
        self.resizable(False, False)

        # Frame principal em 2 colunas
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Coluna esquerda - tabela
        tabela_frame = ctk.CTkFrame(main_frame, width=600)
        tabela_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

        self.textbox_dados = ctk.CTkTextbox(tabela_frame, font=("Consolas", 14), wrap="none", height=440)
        self.textbox_dados.pack(fill="both", expand=True, padx=5, pady=5)
        self.textbox_dados.insert("end", "Cole aqui os dados")  # Exemplo de cabeçalho

        # Coluna direita
        direita_frame = ctk.CTkFrame(main_frame, width=400)
        direita_frame.pack(side="left", fill="y")

        # Pré-visualização do modelo
        self.preview_label = ctk.CTkLabel(
            direita_frame,
            text="Aguardando seleção de modelo",
            width=300, height=200, anchor="center",
            fg_color="lightgray", text_color="black",
            corner_radius=8)
        self.preview_label.pack(pady=10, padx=10)

        # Combobox para modelos
        self.modelo_combobox = ctk.CTkComboBox(direita_frame, values=["(nenhum modelo disponível)"])
        self.modelo_combobox.pack(pady=5, padx=10, fill="x")
        self.modelo_combobox.bind("<<ComboboxSelected>>", self.atualizar_preview_modelo)

        # Botão Gerar Cartões
        self.btn_gerar = ctk.CTkButton(direita_frame, text="Gerar Cartões")
        self.btn_gerar.pack(pady=10, padx=10, fill="x")

        # Log do programa
        self.log_textbox = ctk.CTkTextbox(direita_frame, height=120)
        self.log_textbox.insert("end", "Log do programa...\n")
        self.log_textbox.configure(state="disabled")
        self.log_textbox.pack(pady=10, padx=10, fill="x")

        # Botões de Gerenciamento de Modelo
        botoes_modelo_frame = ctk.CTkFrame(direita_frame)
        botoes_modelo_frame.pack(pady=5, padx=10, fill="x")

        self.btn_adicionar_modelo = ctk.CTkButton(
            botoes_modelo_frame,
            text="Adicionar modelo",
            command=self.adicionar_modelo
        )

        self.btn_adicionar_modelo.pack(side="left", padx=5, expand=True)

        self.btn_modificar_modelo = ctk.CTkButton(
            botoes_modelo_frame,
            text="Modificar modelo",
            command=self.modificar_modelo
        )

        self.btn_modificar_modelo.pack(side="left", padx=5, expand=True)

        self.btn_excluir_modelo = ctk.CTkButton(
            botoes_modelo_frame,
            text="Excluir modelo",
            command=self.excluir_modelo
        )

        self.btn_excluir_modelo.pack(side="left", padx=5, expand=True)

        # Espaço extra para respiro visual abaixo dos botões de modelos
        ctk.CTkLabel(direita_frame, text="").pack(pady=8)

        #Chama a função para carregar os modelos e exibir a pré-visualização inicial
        self.atualizar_modelos_combobox()

    def garantir_pasta_modelos(self):
        if not os.path.exists("modelos"):
            os.makedirs("modelos")

    def atualizar_modelos_combobox(self):
        self.garantir_pasta_modelos()
        arquivos = [f for f in os.listdir("modelos") if f.lower().endswith(".xcf")]
        if arquivos:
            self.modelo_combobox.configure(values=arquivos)
            self.modelo_combobox.set(arquivos[0])
            self.atualizar_preview_modelo()
        else:
            self.modelo_combobox.configure(values=["(nenhum modelo disponível)"])
            self.modelo_combobox.set("(nenhum modelo disponível)")
            self.atualizar_preview_modelo()

    def adicionar_modelo(self):
        self.garantir_pasta_modelos()
        arquivo = filedialog.askopenfilename(
            title="Selecione um modelo (.xcf)",
            filetypes=[("Arquivos do GIMP", "*.xcf")]
        )
        if arquivo:
            nome_arquivo = os.path.basename(arquivo)
            destino = os.path.join("modelos", nome_arquivo)
            # Se já existe, pergunta se deseja substituir
            if os.path.exists(destino):
                resp = messagebox.askyesno("Substituir modelo",
                                           f"O modelo '{nome_arquivo}' já existe. Deseja substituir?")
                if not resp:
                    return
            shutil.copy2(arquivo, destino)
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert("end", f"Modelo '{nome_arquivo}' adicionado/atualizado.\n")
            self.log_textbox.configure(state="disabled")
            self.log_textbox.see("end")
            self.atualizar_modelos_combobox()

    def modificar_modelo(self):
        modelo_selecionado = self.modelo_combobox.get()
        if modelo_selecionado == "(nenhum modelo disponível)":
            messagebox.showinfo("Modificar modelo", "Nenhum modelo selecionado para modificar.")
            return
        novo_arquivo = filedialog.askopenfilename(
            title="Selecione o novo arquivo do modelo (.xcf)",
            filetypes=[("Arquivos do GIMP", "*.xcf")]
        )
        if novo_arquivo:
            destino = os.path.join("modelos", modelo_selecionado)
            shutil.copy2(novo_arquivo, destino)
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert("end", f"Modelo '{modelo_selecionado}' foi modificado.\n")
            self.log_textbox.configure(state="disabled")
            self.log_textbox.see("end")
            self.atualizar_modelos_combobox()

    def excluir_modelo(self):
        modelo_selecionado = self.modelo_combobox.get()
        if modelo_selecionado == "(nenhum modelo disponível)":
            messagebox.showinfo("Excluir modelo", "Nenhum modelo selecionado para excluir.")
            return
        resp = messagebox.askyesno("Excluir modelo", f"Tem certeza que deseja excluir o modelo '{modelo_selecionado}'?")
        if resp:
            caminho = os.path.join("modelos", modelo_selecionado)
            try:
                os.remove(caminho)
                # Remove também a pré-visualização, se existir
                preview_path = os.path.join("modelos", f"{modelo_selecionado}_preview.png")
                if os.path.exists(preview_path):
                    os.remove(preview_path)

                self.log_textbox.configure(state="normal")
                self.log_textbox.insert("end", f"Modelo '{modelo_selecionado}' foi excluído.\n")
                self.log_textbox.configure(state="disabled")
                self.log_textbox.see("end")
                self.atualizar_modelos_combobox()
            except Exception as e:
                messagebox.showerror("Erro ao excluir", f"Erro: {e}")

    def atualizar_preview_modelo(self, event=None):
        modelo = self.modelo_combobox.get()
        if not modelo or modelo == "(nenhum modelo disponível)":
            self.preview_label.configure(image=None, text="Aguardando seleção de modelo", fg_color="lightgray")
            return

        caminho_xcf = os.path.join("modelos", modelo)
        caminho_preview = os.path.join("modelos", f"{modelo}_preview.png")

        # --- NOVA LÓGICA COM IMAGEMAGICK ---

        # 1. Verificação de pré-visualização existente para otimização
        # Se a pré-visualização PNG já existe e é mais recente que o arquivo XCF,
        # simplesmente a carregamos, economizando tempo e recursos.
        if os.path.exists(caminho_preview) and os.path.getmtime(caminho_preview) > os.path.getmtime(caminho_xcf):
            try:
                self.log_textbox.configure(state="normal")
                self.log_textbox.insert("end", f"Carregando pré-visualização existente para '{modelo}'.\n")
                self.log_textbox.configure(state="disabled")
                self.log_textbox.see("end")

                img = Image.open(caminho_preview)
                # Redimensiona a imagem para a altura base, mantendo a proporção
                h_base = 200
                proporcao = h_base / img.height
                w_novo = int(img.width * proporcao)
                # NÃO use img.resize() aqui, CTkImage fará o redimensionamento.
                # NÃO use ImageTk.PhotoImage.
                # Use CTkImage para compatibilidade com HighDPI e para evitar o UserWarning
                self._ctk_img_preview = ctk.CTkImage(light_image=img,
                                                     dark_image=img,
                                                     size=(w_novo, h_base))
                self.preview_label.configure(image=self._ctk_img_preview, text="", fg_color="transparent")
                return  # Saímos da função pois a pré-visualização já foi carregada
            except Exception as e:
                self.log_textbox.configure(state="normal")
                self.log_textbox.insert("end",
                                        f"Erro ao carregar prévia existente: {e}\nTentando gerar novamente com ImageMagick...\n")
                self.log_textbox.configure(state="disabled")
                self.log_textbox.see("end")
                # Se houver erro ao carregar a prévia existente, o fluxo continua para gerar uma nova.

        # 2. Caminho para o executável do ImageMagick
        # Apenas 'magick' é geralmente suficiente se estiver no PATH do sistema.
        # Se não estiver no PATH, você precisaria do caminho completo:
        # imagemagick_path = r"C:\Program Files\ImageMagick-7.1.1-Q16\magick.exe" # Ajuste o caminho se necessário
        # Usamos 'magick' diretamente pois confiamos que o instalador o adicionou ao PATH.
        imagemagick_command = r"C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe"

        # 3. Construção do comando ImageMagick
        # O comando é super simples: 'magick [arquivo_de_entrada] [arquivo_de_saida]'
        # Podemos adicionar -flatten para garantir que as camadas sejam mescladas,
        # embora para XCF o ImageMagick já faça isso por padrão na conversão.
        comando_convert = [
            imagemagick_command,
            caminho_xcf,
            "-flatten",  # Garante que todas as camadas visíveis sejam mescladas em uma única imagem
            caminho_preview
        ]

        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", f"Gerando pré-visualização com ImageMagick para '{modelo}'...\n")
        self.log_textbox.configure(state="disabled")
        self.log_textbox.see("end")

        # 4. Execução do comando via subprocess
        try:
            # check=True fará com que um CalledProcessError seja levantado se o comando retornar um código de erro.
            # capture_output=True e text=True são bons para depuração, caso algo dê errado.
            result = subprocess.run(comando_convert, check=True, capture_output=True, text=True, shell=False)

            if result.stdout:
                self.log_textbox.configure(state="normal")
                self.log_textbox.insert("end", f"ImageMagick STDOUT: {result.stdout}\n")
                self.log_textbox.configure(state="disabled")
            if result.stderr:
                self.log_textbox.configure(state="normal")
                self.log_textbox.insert("end", f"ImageMagick STDERR: {result.stderr}\n")
                self.log_textbox.configure(state="disabled")

            # 5. Carregar e exibir a imagem gerada
            img = Image.open(caminho_preview)
            # Redimensiona a imagem para a altura base, mantendo a proporção
            h_base = 200
            proporcao = h_base / img.height
            w_novo = int(img.width * proporcao)
            img = img.resize((w_novo, h_base), Image.LANCZOS)
            # Use CTkImage para compatibilidade com HighDPI e para evitar o UserWarning
            self._ctk_img_preview = ctk.CTkImage(light_image=img,
                                                 dark_image=img,
                                                 size=(w_novo, h_base))
            self.preview_label.configure(image=self._ctk_img_preview, text="", fg_color="white")

            self.log_textbox.configure(state="normal")
            self.log_textbox.insert("end", f"Pré-visualização para '{modelo}' gerada com sucesso.\n")
            self.log_textbox.configure(state="disabled")
            self.log_textbox.see("end")

        except FileNotFoundError:
            # Este erro acontece se 'magick' não for encontrado (não está no PATH ou o nome está errado)
            self.preview_label.configure(image=None, text="ImageMagick não encontrado", fg_color="red")
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert("end", "Erro: Executável do ImageMagick 'magick' não encontrado.\n")
            self.log_textbox.insert("end", "Verifique se está instalado e no PATH do sistema.\n")
            self.log_textbox.configure(state="disabled")
            self.log_textbox.see("end")
        except subprocess.CalledProcessError as e:
            # Este erro acontece se o comando 'magick' for executado, mas retornar um erro (ex: arquivo XCF corrompido)
            self.preview_label.configure(image=None, text="Erro ImageMagick", fg_color="red")
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert("end", f"Erro no subprocesso do ImageMagick: {e.returncode}\n")
            self.log_textbox.insert("end", f"STDOUT: {e.stdout}\n")
            self.log_textbox.insert("end", f"STDERR: {e.stderr}\n")
            self.log_textbox.configure(state="disabled")
            self.log_textbox.see("end")
        except Exception as e:
            # Outros erros, como problemas ao abrir o PNG gerado pelo PIL
            self.preview_label.configure(image=None, text="Erro ao gerar prévia", fg_color="orange")
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert("end", f"Erro inesperado ao gerar prévia: {e}\n")
            self.log_textbox.configure(state="disabled")
            self.log_textbox.see("end")


# Restante do código da classe CartaoApp e o if __name__ == "__main__":
if __name__ == "__main__":
    app = CartaoApp()
    app.mainloop()

