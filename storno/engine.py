import pandas as pd
import numpy as np
import os
import json
from datetime import datetime

# Regras Fiscal - RICMS-TO
REGRAS = {
    "Laticínios": {"prefixos": ["0401", "0402", "0403", "0404", "0405", "0406"], "percentual": 0.40},
    "Arroz": {"prefixos": ["1006"], "percentual": 0.65, "macro": "Cesta Básica"},
    "Feijão": {"prefixos": ["0713"], "percentual": 0.65, "macro": "Cesta Básica"},
    "Farinha de Mandioca": {"prefixos": ["110620"], "percentual": 0.65, "macro": "Cesta Básica"},
    "Fubá de Milho": {"prefixos": ["110220"], "percentual": 0.65, "macro": "Cesta Básica"},
    "Açúcar Cristal": {"prefixos": ["170114", "170199", "1701"], "percentual": 0.65, "macro": "Cesta Básica"},
    "Óleo de Soja": {"prefixos": ["1507"], "percentual": 0.65, "macro": "Cesta Básica"},
    "Café": {"prefixos": ["0901"], "percentual": 0.65, "macro": "Cesta Básica"},
    "Sal": {"prefixos": ["2501"], "percentual": 0.65, "macro": "Cesta Básica"},
}

def limpa_ncm(ncm):
    if pd.isna(ncm): return ""
    v = str(ncm).strip()
    if v.endswith(".0"): 
        v = v[:-2]
    v = v.replace(".", "").replace("-", "")
    if v.isdigit() and len(v) < 8:
        v = v.zfill(8)
    return v

def processar_planilha(filepath, output_dir):
    try:
        # 1. Leitura robusta
        if filepath.lower().endswith('.csv'):
            # Tenta vários separadores comuns em arquivos fiscais
            for sep in [',', ';', '\t']:
                try:
                    df = pd.read_csv(filepath, sep=sep, dtype=str, encoding='utf-8')
                    if len(df.columns) > 1: break
                except: continue
        else:
            df = pd.read_excel(filepath, dtype=str)

        # 2. Limpeza profunda dos nomes das colunas (remove espaços, aspas e quebras de linha)
        df.columns = df.columns.str.strip().str.replace('"', '').str.replace("'", "")

        # 3. Mapeamento Flexível (Aceita 'Numero' ou 'Numero_Nota', etc)
        mapeamento = {
            "Numero_Nota": ["Numero", "Numero_Nota", "No.", "Num"],
            "Data_Nota": ["Dt_Emissao", "Data_Nota", "Data"],
            "Fornecedor": ["Rz_Emit", "Fornecedor", "Emitente"],
            "Produto": ["Produto", "Descricao", "Nome_Produto"],
            "NCM": ["NCM", "NCM_SH"],
            "Credito_ICMS": ["Valor_ICMS", "Credito_ICMS", "Vlr_ICMS"]
        }

        colunas_finais = {}
        for destino, origens in mapeamento.items():
            encontrou = False
            for o in origens:
                if o in df.columns:
                    colunas_finais[o] = destino
                    encontrou = True
                    break
            if not encontrou:
                return {"sucesso": False, "erro": f"Coluna essencial não encontrada. Verifique se existe algo como: {origens}"}

        df = df.rename(columns=colunas_finais)

        # 4. Tratamento de valores numéricos
        df['Credito_ICMS'] = pd.to_numeric(df['Credito_ICMS'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        
        # Opcional: tentar converter Valor_Produto ou Valor_Unitario se precisarmos
        for col in ['Valor_Produto', 'Valor_Unitario', 'Quantidade', 'Valor_Total_Nota']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        
        # 5. Classificação e Cálculo
        df['NCM_Limpo'] = df['NCM'].apply(limpa_ncm)
        
        def classificar(ncm):
            for cat, reg in REGRAS.items():
                for pref in reg["prefixos"]:
                    if ncm.startswith(pref): 
                        return cat, reg["percentual"], reg.get("macro", "Laticínios")
            return "Outros", 0.0, "Outros"

        df['Classificacao'] = df['NCM_Limpo'].apply(classificar)
        df['Categoria'] = df['Classificacao'].apply(lambda x: x[0])
        df['Percentual_Estorno'] = df['Classificacao'].apply(lambda x: x[1])
        df['Macro_Categoria'] = df['Classificacao'].apply(lambda x: x[2])
        df['Valor_Estorno'] = df['Credito_ICMS'] * df['Percentual_Estorno']
        df['Credito_Aproveitavel'] = df['Credito_ICMS'] - df['Valor_Estorno']

        df_estorno = df[df['Categoria'] != "Outros"].copy()

        if df_estorno.empty:
            return {"sucesso": False, "erro": "Nenhum produto gera estorno conforme os NCMs da planilha."}

        # 6. Agrupamento por Nota
        if 'Numero_Nota' in df_estorno.columns:
            df_nota = df_estorno.groupby("Numero_Nota").agg({
                "Data_Nota": "first",
                "Fornecedor": "first",
                "Credito_ICMS": "sum",
                "Valor_Estorno": "sum"
            }).reset_index()
        else:
            df_nota = pd.DataFrame() # Fallback case

        # 7. Aggregações Top e Macros
        df_macro = df_estorno.groupby("Macro_Categoria")["Valor_Estorno"].sum().reset_index()
        macro_dict = {row['Macro_Categoria']: float(row['Valor_Estorno']) for _, row in df_macro.iterrows()}
        total_laticinios = macro_dict.get("Laticínios", 0.0)
        total_cesta = macro_dict.get("Cesta Básica", 0.0)

        df_produtos = df_estorno.groupby("Produto")["Valor_Estorno"].sum().reset_index()
        df_produtos = df_produtos.sort_values(by="Valor_Estorno", ascending=False).head(10)

        if 'Fornecedor' in df_estorno.columns:
            df_forn = df_estorno.groupby("Fornecedor")["Valor_Estorno"].sum().reset_index()
            df_forn = df_forn.sort_values(by="Valor_Estorno", ascending=False).head(10)
        else:
            df_forn = pd.DataFrame()

        # 8. Geração do Arquivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"resultado_estorno_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_estorno.to_excel(writer, sheet_name="Por Produto", index=False)
            if not df_nota.empty:
                df_nota.to_excel(writer, sheet_name="Por Nota", index=False)
            
        return {
            "sucesso": True,
            "resumo": {
                "total_linhas": int(len(df)),
                "linhas_com_estorno": int(len(df_estorno)),
                "total_estorno": float(df_estorno['Valor_Estorno'].sum()),
                "total_laticinios": total_laticinios,
                "total_cesta_basica": total_cesta
            },
            "graficos": {
                "top_produtos": json.loads(df_produtos.to_json(orient='records')),
                "top_fornecedores": json.loads(df_forn.to_json(orient='records')) if not df_forn.empty else []
            },
            "preview_produto": json.loads(df_estorno.head(50).to_json(orient='records')),
            "preview_nota": json.loads(df_nota.head(50).to_json(orient='records')) if not df_nota.empty else [],
            "arquivo_saida": output_filename
        }

    except Exception as e:
        return {"sucesso": False, "erro": str(e)}
