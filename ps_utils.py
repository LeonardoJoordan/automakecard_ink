import win32com.client

def gerar_cartao_photoshop(psd_path, output_path, campos):
    '''
    psd_path: caminho do modelo PSD
    output_path: caminho do PNG de saída
    campos: dicionário com {'tratamento': str, 'nome': str, 'conjuge': str, 'data': str}
    '''
    psApp = win32com.client.Dispatch("Photoshop.Application")
    psApp.Visible = False  # Coloque True se quiser acompanhar
    doc = psApp.Open(psd_path)

    for camada_nome, texto in campos.items():
        try:
            doc.ArtLayers[camada_nome].TextItem.Contents = texto
        except Exception as e:
            print(f"Erro ao alterar camada '{camada_nome}': {e}")

    # Exporta para PNG
    options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
    options.Format = 13  # PNG
    options.PNG8 = False  # PNG-24
    doc.Export(ExportIn=output_path, ExportAs=2, Options=options)

    doc.Close(2)  # Fecha sem salvar alterações no PSD