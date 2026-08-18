import pandas as pd
import numpy as np
import re

# 1. Configuração e Leitura
caminho = r"C:\estruturação razao com nf\janei-2026.xlsx"
df = pd.read_excel(caminho, header=3)

# 2. Limpeza Inicial
df = df.dropna(how='all', axis=1)

cols_map = {}
for col in df.columns:
    if "Unnamed" in str(col):
        cols_map[col] = "Natureza_Saldo"
df.rename(columns=cols_map, inplace=True)

# 3. Conta Contábil
df['Conta_Contabil'] = df['Emp.'].where(df['Data'] == 'Conta:')
df['Conta_Contabil'] = df['Conta_Contabil'].ffill()

divisao = df['Conta_Contabil'].str.split('-', n=1, expand=True)
df['Numero_Conta'] = divisao[0].str.strip()
df['Nome_Conta'] = divisao[1].str.strip()

# 4. Filtro
df_limpo = df[df['Data'].astype(str).str.startswith('20', na=False)].copy()

# ==============================================================================
# 5. EXTRAÇÃO DO NÚMERO DA NF (CORRIGIDO)
# ==============================================================================

def extrair_nf(texto):
    if pd.isna(texto):
        return None

    texto_str = str(texto)

    # EXPLICAÇÃO DO PADRÃO NOVO:
    # (?: ... ) -> Grupo de palavras possíveis (NF, NFSe, OS, REF, etc)
    # [.:\s-]* -> Aceita qualquer combinação de pontos, dois pontos, espaços ou traços depois da palavra
    # (\d+)     -> Captura os dígitos que vierem depois

    padrao = r"(?:NFSe|NFS-e|NF|SERVICOREF|REF|NUNESREF|Cupom|Doc| NFSe:|NFS-e:|NF:|Nf Debito: |SERVICOREF |-REF |NF.|NUNESREF | REF|NUREF | NF|NF. DEV.: | - REF |NF -|Nf. Transf. |Nf.Dev: |Nf De Cortesia : |Nf De Cortesia:|Nf.|Nf-E: |Nf. Dev.|NF DEBITO: |Nfs-E:|NF. DEV.  |NF.DEV: |Nf. Dev.  |Nf. Dev.: |Nf |Nf Dev.: |NF.DEV.: |NF DEV.|Nf - |Nf Dev.|Nf.Dev.: |Nfse:|Nf: |NF\s|- NF)[.:\s-]*(\d+)"

    # flags=re.IGNORECASE -> Ignora se é maiúscula ou minúscula
    resultado = re.search(padrao, texto_str, flags=re.IGNORECASE)

    if resultado:
        return resultado.group(1)
    return None

# Aplica a função
df_limpo["Número_NF"] = df_limpo["Histórico"].apply(extrair_nf)

# Verifica se funcionou (mostra linhas que não ficaram vazias)
print(f"NFs encontradas: {df_limpo['Número_NF'].notnull().sum()}")
# ==============================================================================

# 6. Tratamento de Números
cols_numericas = ['Débito', 'Crédito', 'Saldo']
for col in cols_numericas:
    if col in df_limpo.columns:
        df_limpo[col] = pd.to_numeric(df_limpo[col], errors='coerce').fillna(0)

# 7. Saldos
df_limpo['Saldo_Movimento'] = df_limpo['Crédito'] - df_limpo['Débito']

# 8. Organizar
cols_ordem = ['Data', 'Numero_Conta', 'Nome_Conta', 'Histórico', 'Número_NF', 'Débito', 'Crédito', 'Saldo', 'Saldo_Movimento']
cols_restantes = [c for c in df_limpo.columns if c not in cols_ordem]
df_limpo = df_limpo[cols_ordem + cols_restantes]

# 9. Salvar
print(df_limpo[['Histórico', 'Número_NF']].head()) # Mostra para conferir
arquivo_saida = r"C:\estruturação razao com nf\Razao_Final_Com_NF_Corrigido.xlsx"
df_limpo.to_excel(arquivo_saida, index=False)
