#!/usr/bin/env python
# -*- coding: utf-8 -*-

from gimpfu import *
import json
import os

def plugin_atualiza_e_exporta_cartao(caminho_arquivo_json_dados):
    try:
        # Carrega os dados do arquivo JSON
        with open(caminho_arquivo_json_dados, 'r', encoding='utf-8') as f:
            dados_cartao = json.load(f)

        caminho_template_xcf = dados_cartao.get("template_path")
        caminho_saida_cartao = dados_cartao.get("output_filename")

        # --- IMPORTANTE: Mapeamento dos seus dados para os NOMES DAS CAMADAS DE TEXTO no GIMP ---
        # Você PRECISA verificar os nomes exatos das suas camadas de texto no arquivo XCF
        # e atualizar os valores à direita neste dicionário.
        # A chave (ex: "nome") deve corresponder ao cabeçalho da sua tabela no app Python.
        mapeamento_camadas_texto = {
            "tratamento": "TextoTratamento",  # Ex: Se sua camada no GIMP para 'tratamento' se chama 'TextoTratamento'
            "nome": "TextoNomePrincipal",     # Ex: Se sua camada no GIMP para 'nome' se chama 'TextoNomePrincipal'
            "conjuge": "TextoConjuge",        # Ex: Se sua camada no GIMP para 'conjuge' se chama 'TextoConjuge'
            "data": "TextoDataEvento",         # Ex: Se sua camada no GIMP para 'data' se chama 'TextoDataEvento'
            # Adicione aqui outros campos/camadas se tiver, por exemplo:
            # "campo_extra_da_tabela": "NomeDaCamadaDeTextoNoGIMP"
        }
        # ------------------------------------------------------------------------------------

        if not caminho_template_xcf or not caminho_saida_cartao:
            gimp.message("ERRO: Caminho do template ou nome do arquivo de saída não fornecido no JSON.")
            return

        # Carrega a imagem (template XCF)
        imagem = pdb.gimp_file_load(caminho_template_xcf, os.path.basename(caminho_template_xcf))
        if not imagem:
            gimp.message(f"ERRO: Não foi possível carregar o template XCF: {caminho_template_xcf}")
            return

        # Itera sobre os dados do cartão e atualiza as camadas de texto correspondentes
        for chave_dado, valor_texto in dados_cartao.items():
            if chave_dado in mapeamento_camadas_texto:
                nome_camada_gimp = mapeamento_camadas_texto[chave_dado]
                camada_texto = pdb.gimp_image_get_layer_by_name(imagem, nome_camada_gimp)

                if camada_texto and pdb.gimp_item_is_text_layer(camada_texto):
                    pdb.gimp_text_layer_set_text(camada_texto, str(valor_texto)) # Converte para string por segurança
                    # gimp.message(f"Camada '{nome_camada_gimp}' atualizada para: '{valor_texto}'") # Para debug
                # else:
                    # gimp.message(f"AVISO: Camada de texto '{nome_camada_gimp}' não encontrada ou não é uma camada de texto no XCF.") # Para debug

        # Para exportar como PNG, geralmente mesclamos as camadas visíveis.
        # É mais seguro duplicar a imagem, mesclar a duplicata e exportar, para não alterar o XCF original.
        # Mas para simplificar, vamos exportar diretamente. O GIMP tentará mesclar camadas visíveis.
        # Se você quiser garantir uma imagem "achatada":
        # drawable_export = pdb.gimp_image_merge_visible_layers(imagem, CLIP_TO_IMAGE)
        # Se não, use None para o drawable se a função de exportação permitir (como file_png_save2)

        drawable_export = None # Para file_png_save2, None usa a imagem inteira e mescla visíveis

        pdb.file_png_save2(imagem, drawable_export, caminho_saida_cartao, os.path.basename(caminho_saida_cartao),
                           0,  # interlace (0=None)
                           9,  # compression (0-9)
                           1,  # bkgd (save background color)
                           1,  # gama (save gamma)
                           1,  # offs (save layer offset)
                           1,  # phys (save resolution)
                           1)  # time (save timestamp)

        gimp.message(f"Cartão gerado com sucesso: {caminho_saida_cartao}")

        # Fecha a imagem do GIMP SEM salvar as alterações no arquivo XCF original
        pdb.gimp_image_delete(imagem)

    except Exception as e:
        # Em caso de erro, mostra uma mensagem no console de erros do GIMP
        error_message = f"ERRO no plugin Python-Fu 'atualizador_cartao_gimp': {str(e)}"
        gimp.message(error_message)
        # Para debug mais avançado, você pode logar o traceback completo em um arquivo
        # import traceback
        # debug_file_path = os.path.join(os.path.expanduser("~"), "gimp_plugin_error_log.txt")
        # with open(debug_file_path, "a", encoding="utf-8") as err_file:
        #     err_file.write(error_message + "\n")
        #     err_file.write(traceback.format_exc() + "\n---\n")

# Registra o plugin no GIMP
register(
    "python_fu_atualizador_cartao_gimp",  # Nome interno do plugin (usado na chamada batch)
    "Atualizador de Cartões via JSON",
    "Atualiza camadas de texto em um XCF com base em dados de um arquivo JSON e exporta como PNG.",
    "Seu Nome/Gemini",
    "Seu Nome/Gemini",
    "2025",
    "Atualizar Cartão via JSON...", # Rótulo do menu (não usaremos diretamente, mas é bom ter)
    "",  # Tipos de imagem que aceita (deixe vazio para não aparecer no menu de imagem)
    [
        (PF_STRING, "caminho_arquivo_json_dados", "Caminho do arquivo JSON com os dados do cartão", "")
    ],
    [], # Resultados de saída
    plugin_atualiza_e_exporta_cartao, # Função principal
    menu="<Toolbox>/File/Create" # Opcional: onde apareceria no menu do GIMP
)

main()