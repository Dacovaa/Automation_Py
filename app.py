import pandas as pd
import re
from difflib import SequenceMatcher
import unicodedata

def remover_acentos(texto):
    return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')

def extrair_palavras_chave(texto):
    if not isinstance(texto, str):
        texto = str(texto) if texto is not None else ""
        
    texto = texto.lower()
    texto = remover_acentos(texto)
    texto = re.sub(r'[^a-z0-9\s]', '', texto)
    termos_irrelevantes = [
        "forma", "farmaceutica", "apresentacao", "via", "administracao", "oral", 
        "capsula", "comprimido", "revestido", "injeção", "suspensao", "liquido",
        "injetavel", "ampola", "frascos", "doses", "drageia", "solucao", "substancia", 
        "farmaceutica", "miligramas", "", "", "humano", "uso", "especificações", 
        "especificacoes", "concentracao", "gerais", "revestido","unidade","intramuscular",
        "intravenosa", "medicamentos", "de", "principioativo","principio","ativo","principioconcentracao1",
        "controlados", "controlado", "concentrao","concentracao", "cp", "com", "em", "contendo", "embalado",
        "mgml", "frasco", "comprimidos", "blister"
    ]
    
    palavras = texto.split()
    palavras_relevantes = [palavra for palavra in palavras if palavra not in termos_irrelevantes]
    
    palavras_relevantes = palavras_relevantes[:3]
    
    return " ".join(palavras_relevantes)

def similaridade(a, b):
    return SequenceMatcher(None, a, b).ratio()

def verificar_dosagem(dosagem_destino, dosagem_cotacao):
    if dosagem_destino and dosagem_cotacao:
        return dosagem_destino == dosagem_cotacao
    return True  # Se uma das dosagens for ausente, não considerar a diferença como erro

planilha_cotacao = pd.read_excel(r'\\Recdist2\rede\PLANILHA COTAÇÕES PREGÕES\TABELA LANÇAMENTOS PREGÕES_10_04_22.xlsx')
planilha_destino = pd.read_excel(r'\\Recdist2\rede\PLANILHA COTAÇÕES PREGÕES\TABELA DESTINO AUTOMAÇÃO.xlsx')

limite_similaridade = 0.80

itens_correspondidos = 0


colunas_para_converter = ['UNID.', 'FABRICANTE', 'EMBALAGEM', 'ANVISA']
planilha_destino[colunas_para_converter] = planilha_destino[colunas_para_converter].astype('object')


for i, descricao_destino in enumerate(planilha_destino['DESCRIÇÃO']):
    descricao_destino_tratada = extrair_palavras_chave(descricao_destino)
    melhor_similaridade = 0
    melhor_match = None
    melhor_correspondencia = None

    dosagem_destino = re.search(r'(\d+)', descricao_destino.lower()) #extrair dosagem dest
    dosagem_destino = dosagem_destino.group(0) if dosagem_destino else None

    for j, descricao_cotacao in enumerate(planilha_cotacao['DESCRIÇÃO']):
        descricao_cotacao_tratada = extrair_palavras_chave(descricao_cotacao)

        descricao_cotacao = str(descricao_cotacao) if descricao_cotacao is not None else ""
        dosagem_cotacao = re.search(r'(\d+)', descricao_cotacao.lower()) #extrair dosagem cot
        dosagem_cotacao = dosagem_cotacao.group(0) if dosagem_cotacao else None
        
        if not verificar_dosagem(dosagem_destino, dosagem_cotacao):
            continue
        
        sim = similaridade(descricao_destino_tratada, descricao_cotacao_tratada)
        
        if sim > melhor_similaridade:
            melhor_similaridade = sim
            melhor_match = descricao_cotacao
            melhor_correspondencia = planilha_cotacao.iloc[j]

    if melhor_similaridade >= limite_similaridade:
        planilha_destino.at[i, 'Correspondente'] = str(melhor_match)
        planilha_destino.at[i, 'SIMILARIDADE'] = str(melhor_similaridade)
        planilha_destino.at[i, 'DESCRIÇÃO TRATADA'] = str(descricao_destino_tratada) + " " + str(dosagem_destino)
        planilha_destino.at[i, 'UNID.'] = str(melhor_correspondencia['UNID.'])
        planilha_destino.at[i, 'FABRICANTE'] = str(melhor_correspondencia['FABRICANTE'])
        planilha_destino.at[i, 'EMBALAGEM'] = str(melhor_correspondencia['EMBALAGEM'])
        planilha_destino.at[i, 'ANVISA'] = str(melhor_correspondencia['ANVISA'])
        planilha_destino.at[i, 'CUSTO'] = float(melhor_correspondencia['CUSTO'])
        planilha_destino.at[i, 'MÍN. SP'] = float(melhor_correspondencia['MÍN. SP'])
        itens_correspondidos += 1
    else:
        planilha_destino.at[i, 'Correspondente'] = "Nenhuma correspondência"
        planilha_destino.at[i, 'SIMILARIDADE'] = str(melhor_similaridade)
        planilha_destino.at[i, 'DESCRIÇÃO TRATADA'] = str(descricao_destino_tratada) + " " + str(dosagem_destino)
        
planilha_destino.to_excel(r'\\Recdist2\rede\PLANILHA COTAÇÕES PREGÕES\TABELA DESTINO AUTOMAÇÃO COM CORRESPONDENTES.xlsx', index=False)

print(f"Total de itens correspondidos: {itens_correspondidos} de {len(planilha_destino)}")
