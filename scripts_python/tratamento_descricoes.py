# Projeto desenvolvido utilizando Python no Excel (função PY).
# Fiz este código em Python no Excel para automatizar a separação de informações que anteriormente ficavam concentradas em uma única célula. 
# O código identifica e organiza os dados em diferentes colunas, como NF, Fornecedor, Documento, e Descrição, facilitando o tratamento e a análise das informações.

import re
import pandas as pd

src = xl("B2:B194", headers=False).iloc[:, 0].fillna('').astype(str)

marcadores = r'Desembolso Diverso|Compra para Estoque|Compra para Uso Direto|Simples Faturamento|Angela fez|historico|FAZENDA|pagamento'

def separar(texto):
    texto = str(texto).strip()
    nf = ''
    fornecedor = ''
    doc = ''
    fornecedor2 = ''
    descricao = ''

    inicio = re.match(r'^\s*([^-]+?)\s*-\s*(.*)$', texto)
    restante = inicio.group(2).strip() if inicio else texto

    segunda_parte = re.match(r'^(.+?)\s*-\s*(.*)$', restante)
    if segunda_parte:
        nf = segunda_parte.group(1).strip()
        bloco_fornecedor = segunda_parte.group(2).strip()
    else:
        bloco_fornecedor = restante

    partes_fornecedor = re.split(r'(?=' + marcadores + r')', bloco_fornecedor, maxsplit=1, flags=re.I)
    fornecedor = partes_fornecedor[0].strip()

    achou_doc = re.search(r'\bDoc\.?\s*([^\s:]+)', texto, flags=re.I)
    if achou_doc:
        doc = achou_doc.group(1).strip()

    achou_fornecedor2 = re.search(r'\bDoc\.?\s*[^\s:]+\s+de\s+(.+?)\s*:', texto, flags=re.I)
    if achou_fornecedor2:
        fornecedor2 = achou_fornecedor2.group(1).strip()
        descricao = texto[achou_fornecedor2.end():].strip()
    else:
        pos_dois_pontos = texto.find(':')
        if pos_dois_pontos >= 0:
            descricao = texto[pos_dois_pontos + 1:].strip()
        elif len(partes_fornecedor) > 1:
            descricao = partes_fornecedor[1].strip()

    return [nf, fornecedor, doc, fornecedor2, descricao]

pd.DataFrame([separar(x) for x in src], columns=["NF", "Fornecedor", "Documento", "Fornecedor 2", "Descrição"]) 
