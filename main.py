import PyPDF2

def extrair_texto_pdfs(lista_de_pdfs, caminho_saida):
    texto_total = ""

    for arquivo in lista_de_pdfs:
        try:
            with open(arquivo, 'rb') as pdf_file:
                pdf = PyPDF2.PdfReader(pdf_file)
                for pagina in pdf.pages:
                    texto = pagina.extract_text()
                    if texto:
                        texto_total += texto + "\n"
        except Exception as e:
            print(f"Erro ao processar {arquivo}: {e}")

    try:
        with open(caminho_saida, 'w', encoding='utf-8') as arquivo_saida:
            arquivo_saida.write(texto_total)
        return True, f"Texto extraído com sucesso para:\n{caminho_saida}"
    except Exception as e:
        return False, f"Erro ao salvar o arquivo: {e}"
