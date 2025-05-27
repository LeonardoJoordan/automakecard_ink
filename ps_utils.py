# ps_utils.py
import win32com.client

PS_EXPORT_TYPE_SAVE_FOR_WEB = 2

# A função agora recebe 'psApp' como um argumento!
def gerar_cartao_photoshop(psApp, doc, output_path, campos, export_options_obj):
    """
    Recebe uma instância do app Photoshop (psApp) e um documento aberto (doc),
    modifica o documento e exporta um cartão.
    """

    # O documento já está aberto, apenas modificamos
    for camada_nome, texto in campos.items():
        try:
            doc.ArtLayers[camada_nome].TextItem.Contents = texto
        except Exception as e:
            print(f"Aviso: Erro ao alterar camada '{camada_nome}': {e}")

    # Exporta para PNG
    #options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
    #options.Format = 13  # PNG
    #options.PNG8 = False  # PNG-24
    doc.Export(ExportIn=output_path, ExportAs=PS_EXPORT_TYPE_SAVE_FOR_WEB, Options=export_options_obj)

#--------------------------------------------------------------------------------------------------------

def listar_camadas_de_texto(psd_path):
    """
    Abre um arquivo PSD de forma invisível, lista os nomes de todas as camadas de texto
    e fecha o arquivo. Retorna uma lista com os nomes.
    """
    camadas_de_texto = []
    doc = None
    try:
        # Tenta conectar ao Photoshop. Pode falhar se não estiver instalado/aberto.
        psApp = win32com.client.Dispatch("Photoshop.Application")

        # O argumento 'True' no final abre o documento como "read-only" (apenas leitura)
        doc = psApp.Open(psd_path, None, True)

        for layer in doc.ArtLayers:
            # layer.Kind == 2 significa que é uma camada de texto (TextLayer)
            if layer.Kind == 2:
                camadas_de_texto.append(layer.Name)

        return camadas_de_texto

    except Exception as e:
        print(f"ERRO ao listar camadas de texto do arquivo '{psd_path}': {e}")
        # Retorna uma lista vazia em caso de qualquer erro
        return []

    finally:
        # Este bloco SEMPRE será executado, mesmo que um erro aconteça.
        # Garante que o documento aberto de forma invisível seja fechado.
        if doc:
            doc.Close(2)  # 2 = psDoNotSaveChanges
