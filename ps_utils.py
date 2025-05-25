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


