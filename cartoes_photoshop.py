import os
import pandas as pd
import time
from win32com.client import Dispatch
import sys # Import sys to check Python architecture

# --- Verificação de Arquitetura (Importante para COM) ---
# A biblioteca pywin32 (win32com) precisa corresponder à arquitetura do Python (32 ou 64 bits)
# e à arquitetura do Photoshop instalado. Se houver incompatibilidade, podem ocorrer erros.
print(f"Arquitetura Python: {'64-bit' if sys.maxsize > 2**32 else '32-bit'}")
# Verifique manualmente se a arquitetura do Photoshop corresponde.

# --- Caminhos ---
caminho_template = r"C:\\Users\\leona\\Downloads\\cartoes_python.psd"  # Ajuste conforme o nome do seu template
caminho_saida = r"C:\\Users\\leona\\Downloads\\cartoes"
planilha_dados = r"C:\\Users\\leona\\Downloads\\dados_aniversario.xlsx"  # Ajuste conforme o nome da sua planilha

# --- Processamento ---
try:
    # Verifica se a pasta de saída existe, se não, cria
    os.makedirs(caminho_saida, exist_ok=True)
    print(f"Pasta de saída '{caminho_saida}' verificada/criada.")

    # Lendo a planilha
    # Certifique-se de que 'openpyxl' está instalado (pip install openpyxl)
    print(f"Lendo a planilha: {planilha_dados}")
    df = pd.read_excel(planilha_dados)
    print(f"Planilha lida com sucesso. {len(df)} registros encontrados.")

    # Conectando ao Photoshop
    print("Conectando ao Photoshop...")
    try:
        app = Dispatch("Photoshop.Application")
        app.Visible = True # Mantenha True para ver o processo, False para rodar em segundo plano
        print("Conexão com Photoshop estabelecida.")
    except Exception as e:
        print(f"Erro ao conectar/iniciar o Photoshop: {e}")
        print("Verifique se o Photoshop está instalado e se as permissões COM estão corretas.")
        sys.exit(1) # Sai do script se não conseguir conectar

    # Itera sobre as linhas da planilha
    for idx, row in df.iterrows():
        print(f"\nProcessando registro {idx + 1}/{len(df)}: {row['Nome']}")
        doc = None # Inicializa doc como None para o bloco finally
        try:
            # Abrir o template
            print(f"Abrindo template: {caminho_template}")
            doc = app.Open(caminho_template)
            # Pequena pausa para garantir que o documento carregou
            # Pode ser ajustado ou removido se não for necessário
            time.sleep(1)
            print("Template aberto.")

            # Atualizar camadas de texto
            camadas_atualizadas = 0
            for layer in doc.ArtLayers:
                # Verifica se é uma camada de texto antes de acessar TextItem
                if layer.Kind == 2: # 2 corresponde a psTextLayer
                    layer_name_lower = layer.Name.lower()
                    if layer_name_lower == "tratamento":
                        layer.TextItem.contents = str(row["Tratamento"]) # Converte para string por segurança
                        print(f"  - Camada 'Tratamento' atualizada para: {row['Tratamento']}")
                        camadas_atualizadas += 1
                    elif layer_name_lower == "nome":
                        layer.TextItem.contents = str(row["Nome"])
                        print(f"  - Camada 'Nome' atualizada para: {row['Nome']}")
                        camadas_atualizadas += 1
                    elif layer_name_lower == "conjuge":
                        # Trata valores nulos ou vazios na planilha
                        conjuge_val = row["Cônjuge"]
                        if pd.isna(conjuge_val) or str(conjuge_val).strip() == "":
                            layer.TextItem.contents = "" # Define como vazio no Photoshop
                            print("  - Camada 'Conjuge' definida como vazia (valor nulo/vazio na planilha).")
                        else:
                            layer.TextItem.contents = str(conjuge_val)
                            print(f"  - Camada 'Conjuge' atualizada para: {conjuge_val}")
                        camadas_atualizadas += 1
                    elif layer_name_lower == "data":
                        # Formata a data se necessário (exemplo: converter para string)
                        data_val = row["Data"]
                        if isinstance(data_val, pd.Timestamp):
                             # Formata como DD/MM/AAAA (ajuste o formato se precisar)
                            layer.TextItem.contents = data_val.strftime('%d/%m/%Y')
                            print(f"  - Camada 'Data' atualizada para: {data_val.strftime('%d/%m/%Y')}")
                        else:
                            layer.TextItem.contents = str(data_val) # Converte para string
                            print(f"  - Camada 'Data' atualizada para: {data_val}")
                        camadas_atualizadas += 1

            if camadas_atualizadas < 4:
                 print(f"  - Aviso: Apenas {camadas_atualizadas} de 4 camadas esperadas foram encontradas/atualizadas.")
                 print("     Verifique os nomes das camadas no template PSD: 'Tratamento', 'Nome', 'Conjuge', 'Data'.")


            # Exportar como PNG
            options = Dispatch("Photoshop.ExportOptionsSaveForWeb")
            options.Format = 13  # 13 = PNG-24
            options.PNG8 = False # Garante que não é PNG-8
            options.Interlaced = False # Opcional: Define se o PNG será entrelaçado

            # Cria um nome de arquivo seguro (remove caracteres inválidos)
            nome_seguro = "".join(c for c in row['Nome'] if c.isalnum() or c in (' ', '_')).rstrip()
            nome_arquivo = os.path.join(caminho_saida, f"{nome_seguro}.png")

            print(f"Exportando para: {nome_arquivo}")
            # ExportAs=2 significa psExportDocument
            doc.Export(ExportIn=nome_arquivo, ExportAs=2, Options=options)
            print(f"Cartão gerado para: {row['Nome']}")

        except Exception as e:
            # Captura erros durante o processamento de um único cartão
            print(f"Erro ao gerar cartão para {row.get('Nome', 'Nome Desconhecido')} (Registro {idx + 1}): {e}")
            # Você pode querer adicionar mais detalhes do erro aqui, como o traceback
            # import traceback
            # print(traceback.format_exc())
        finally:
            # Garante que o documento seja fechado mesmo se ocorrer um erro
            if doc:
                try:
                    # 2 = psDoNotSaveChanges
                    doc.Close(2)
                    print("Documento do template fechado sem salvar.")
                except Exception as close_err:
                    print(f"Erro ao fechar o documento do template: {close_err}")

    print("\nProcessamento de todos os registros concluído!")

except FileNotFoundError:
    print(f"Erro: A planilha '{planilha_dados}' não foi encontrada. Verifique o caminho.")
except ImportError as e:
    print(f"Erro de importação: {e}")
    print("Certifique-se de que as bibliotecas 'pandas' e 'openpyxl' estão instaladas.")
    print("Use: pip install pandas openpyxl pywin32")
except Exception as e:
    # Captura outros erros gerais (ex: erro ao criar pasta, erro inesperado na leitura do excel)
    print(f"Ocorreu um erro inesperado no script: {e}")
    # import traceback
    # print(traceback.format_exc()) # Descomente para ver o traceback completo

finally:
    # Opcional: Desconectar do Photoshop no final (geralmente não é estritamente necessário com Dispatch)
    # Se você quiser ter certeza que a aplicação foi liberada:
    # try:
    #     if 'app' in locals() and app:
    #         # app.Quit() # Descomente se quiser fechar o Photoshop ao final
    #         del app
    #         print("Recursos do Photoshop liberados.")
    # except Exception as quit_err:
    #     print(f"Erro ao tentar liberar/fechar o Photoshop: {quit_err}")
    pass

print("Script finalizado.")