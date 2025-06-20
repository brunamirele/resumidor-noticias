import os
from docx import Document
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate
import requests
#from langchain.callbacks import get_openai_callback

def buscar_link_google(titulo, veiculo):
    api_key = os.getenv("GOOGLE_API_KEY")
    cx = os.getenv("GOOGLE_CX")
    query = f'"{titulo}" +{veiculo}'
    url = f"https://www.googleapis.com/customsearch/v1?key={api_key}&cx={cx}&q={query}&dateRestrict=d3"

    try:
        res = requests.get(url)
        data = res.json()
        if "items" in data and data["items"]:
            return data["items"][0]["link"]
    except Exception as e:
        print(f"[Erro na busca Google] {e}")
    return "[link não encontrado]"

# === 1. Funções para converter e extrair notícias ===

'''def converter_doc_para_docx(caminho_doc):
    pythoncom.CoInitialize()
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(caminho_doc)
    caminho_docx = caminho_doc + "x"
    doc.SaveAs(caminho_docx, FileFormat=16)  # 16 = formato .docx
    doc.Close()
    word.Quit()
    return caminho_docx'''

def extrair_noticias_heading1(caminho_docx):
    doc = Document(caminho_docx)
    noticias = {}
    atual = []
    contador = 0
    em_noticia = False

    for p in doc.paragraphs:
        estilo = p.style.name.lower()
        texto = p.text.strip()
        if not texto:
            continue

        if estilo == "heading 1":
            if atual:
                contador += 1
                noticias[f"noticia{contador}"] = "\n".join(atual).strip()
                atual = []
            em_noticia = True
            atual.append(texto)
        elif em_noticia:
            atual.append(texto)

    if atual:
        contador += 1
        noticias[f"noticia{contador}"] = "\n".join(atual).strip()

    print(f"Total de notícias encontradas: {len(noticias)}")
    return noticias

'''def processar_arquivo(caminho_doc):
    if not caminho_doc.lower().endswith('.doc'):
        raise ValueError("O arquivo precisa ser .doc")
    print("Convertendo para .docx...")
    caminho_docx = converter_doc_para_docx(caminho_doc)
    print("Extraindo notícias baseadas em Heading 1...")
    noticias = extrair_noticias_heading1(caminho_docx)
    os.remove(caminho_docx)  # limpa .docx temporário
    return noticias'''

def processar_arquivo(caminho_docx):
    if not caminho_docx.lower().endswith('.docx'):
        raise ValueError("O arquivo precisa ser .docx")
    
    print("Extraindo notícias baseadas em Heading 1...")
    noticias = extrair_noticias_heading1(caminho_docx)
    
    return noticias


# === 2. Inicializar ambiente e chain de resumo ===

load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")

prompt = ChatPromptTemplate.from_template(
    "Resuma a notícia: {noticia} em até 100 palavras (não ultrapasse 100 palavras)."
)
chain = prompt | ChatOpenAI() | StrOutputParser()

# === 3. Salvar resumos em arquivo .docx ===

from docx import Document
from docx.shared import Pt

'''def exportar_resumos_para_word(noticias_dict, resumos_dict, caminho_saida='resumos.docx'):
    doc = Document()
    doc.add_heading('Resumos das Notícias', level=1)

    for i in range(1, len(noticias_dict) + 1):
        noticia_key = f'noticia{i}'
        resumo_key = f'resumo{i}'

        noticia = noticias_dict.get(noticia_key, '')
        resumo = resumos_dict.get(resumo_key, '[Resumo não disponível]')

        # Pegamos o primeiro parágrafo da notícia como título (linha Heading 1 original)
        titulo = noticia.split('\n')[0]

        doc.add_heading(f'{i:02d}. {titulo}', level=2)

        p = doc.add_paragraph(resumo)
        p.style.font.size = Pt(11)

    doc.save(caminho_saida)
    print(f"\nArquivo Word exportado com sucesso para: {os.path.abspath(caminho_saida)}")'''

def exportar_resumos_para_word(noticias_dict, resumos_dict, caminho_saida='resumos.docx'):
    doc = Document()
    doc.add_heading('Resumos das Notícias', level=1)

    for i in range(1, len(noticias_dict) + 1):
        noticia_key = f'noticia{i}'
        resumo_key = f'resumo{i}'

        noticia = noticias_dict.get(noticia_key, '')
        resumo = resumos_dict.get(resumo_key, '[Resumo não disponível]')

        linhas = noticia.split('\n')
        titulo = linhas[0] if len(linhas) > 0 else '[Título não encontrado]'
        veiculo = linhas[2] if len(linhas) > 2 else '[Veículo não identificado]'

        link = buscar_link_google(titulo, veiculo)

        doc.add_heading(f'{i:02d}. {titulo}', level=2)
        doc.add_paragraph(f"🔗 {link}")
        p = doc.add_paragraph(resumo)
        p.style.font.size = Pt(11)

    doc.save(caminho_saida)
    print(f"\nArquivo Word exportado com sucesso para: {os.path.abspath(caminho_saida)}")

# === 4. Executar todo o processo: leitura, resumo e exibição ===

def resumir_noticias(noticias_dict):
    resumos = {}
    for chave, noticia in noticias_dict.items():
        try:
            resumo = chain.invoke({'noticia': noticia})
            resumos[chave.replace("noticia", "resumo")] = resumo
        except Exception as e:
            resumos[chave.replace("noticia", "resumo")] = f"[ERRO] {e}"
    return resumos
