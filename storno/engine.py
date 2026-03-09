import pandas as pd
import numpy as np
import os
from datetime import datetime

# Regras Fiscal - RICMS-TO (CST 020)
REGRAS = {
    "Laticínios": {"prefixos": ["0401", "0402", "0403", "0404", "0405", "0406"], "percentual": 0.40},
    "Arroz": {"prefixos": ["1006"], "percentual": 0.65},
    "Feijão": {"prefixos": ["0713"], "percentual": 0.65},
    "Farinha de Trigo": {"prefixos": ["1101"], "percentual": 0.65},
    "Açúcar": {"prefixos": ["1701"], "percentual": 0.65},
    "Óleo de Soja": {"prefixos": ["1507"], "percentual": 0.65},
    "Café": {"prefixos": ["0901"], "percentual": 0.65},
    "Macarrão": {"prefixos": ["1902"], "percentual": 0.65},
}

def limpa_ncm(ncm):
    if pd.isna(ncm):
        return ""
    return str(ncm).replace(".", "").replace("-", "").strip()

def classificar_produto(ncm):
    ncm_str = limpa_ncm(ncm)
    for categoria, regra in REGRAS.items():
        for prefixo in regra["prefixos"]:
            if ncm_str.startswith(prefixo):
                return categoria, regra["percentual"]
    return "Outros", 0.0

def processar_planilha(filepath, output_dir):
    try:
        # Carregar arquivo
        if filepath.lower().endswith('.csv'):
            df = pd.read_csv(filepath, dtype={'NCM': str})
        elif filepath.lower().endswith('.xlsx'):
            df = pd.read_excel(filepath, dtype={'NCM': str})
        else:
            return {"sucesso": False, "erro": "Formato não suportado. Envie XLSX ou CSV."}
            
        colunas_obrigatorias = [
            "Numero_Nota", "Data_Nota", "Fornecedor", 
            "Produto", "NCM", "Credito_ICMS"
        ]
        
        # Validar colunas
        for col in colunas_obrigatorias:
            if col not in df.columns:
                return {"sucesso": False, "erro": f"Coluna obrigatória não encontrada: {col}"}
                
        if df.empty:
            return {"sucesso": False, "erro": "A planilha enviada está vazia."}
            
        # Preparar NCM e Credito_ICMS
        df['NCM_Limpo'] = df['NCM'].apply(limpa_ncm)
        df['Credito_ICMS'] = pd.to_numeric(df['Credito_ICMS'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        
        # Aplicar regras
        df['Classificacao'] = df['NCM_Limpo'].apply(classificar_produto)
        df['Categoria'] = df['Classificacao'].apply(lambda x: x[0])
        df['Percentual_Estorno'] = df['Classificacao'].apply(lambda x: x[1])
        
        # Calcular valores
        df['Valor_Estorno'] = df['Credito_ICMS'] * df['Percentual_Estorno']
        df['Credito_Aproveitavel'] = df['Credito_ICMS'] - df['Valor_Estorno']
        
        # Máscara para estorno
        df_estorno = df[df['Categoria'] != "Outros"].copy()
        
        if df_estorno.empty:
            return {"sucesso": False, "erro": "Nenhum produto na planilha gera estorno (NCM não corresponde às regras)."}
        
        # Resumo final
        resumo_processamento = {
            "total_linhas": len(df),
            "linhas_com_estorno": len(df_estorno),
            "total_estorno": float(df_estorno['Valor_Estorno'].sum())
        }
        
        # ──────── ABA 1: Estorno por Produto ────────
        aba1 = df_estorno[[
            "Numero_Nota", "Data_Nota", "Fornecedor", "Produto", "NCM", 
            "Categoria", "Credito_ICMS", "Percentual_Estorno", 
            "Valor_Estorno", "Credito_Aproveitavel"
        ]].copy()
        
        # ──────── ABA 2: Resumo por Nota ────────
        aba2 = df.groupby(["Numero_Nota", "Data_Nota", "Fornecedor"]).agg(
            Total_Credito_Nota=('Credito_ICMS', 'sum'),
            Total_Estorno_Nota=('Valor_Estorno', 'sum'),
            Total_Aproveitavel_Nota=('Credito_Aproveitavel', 'sum')
        ).reset_index()
        
        # Adicionar Qtd_Produtos_Com_Estorno em Aba 2
        qtd_estorno = df_estorno.groupby(["Numero_Nota", "Data_Nota", "Fornecedor"]).size().reset_index(name='Qtd_Produtos_Com_Estorno')
        aba2 = pd.merge(aba2, qtd_estorno, on=["Numero_Nota", "Data_Nota", "Fornecedor"], how='left')
        aba2['Qtd_Produtos_Com_Estorno'] = aba2['Qtd_Produtos_Com_Estorno'].fillna(0).astype(int)
        
        # ──────── ABA 3: Totais Gerais ────────
        aba3 = df_estorno.groupby("Categoria").agg(
            Qtd_Produtos=('Produto', 'count'),
            Total_Credito=('Credito_ICMS', 'sum'),
            Total_Estorno=('Valor_Estorno', 'sum'),
            Total_Aproveitavel=('Credito_Aproveitavel', 'sum')
        ).reset_index()
        
        linha_total = pd.DataFrame([{
            "Categoria": "TOTAL",
            "Qtd_Produtos": aba3['Qtd_Produtos'].sum(),
            "Total_Credito": aba3['Total_Credito'].sum(),
            "Total_Estorno": aba3['Total_Estorno'].sum(),
            "Total_Aproveitavel": aba3['Total_Aproveitavel'].sum()
        }])
        aba3 = pd.concat([aba3, linha_total], ignore_index=True)
        
        # Formatar Aba 1 para preview no JSON
        preview_df = aba1.head(10).copy()
        preview_df['Percentual_Estorno'] = preview_df['Percentual_Estorno'].apply(lambda x: f"{x*100:.0f}%")
        preview = preview_df.to_dict(orient='records')
        
        # Salvar as abas no Excel
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"resultado_estorno_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        # Formatação de campos para o Excel final
        aba1['Percentual_Estorno'] = aba1['Percentual_Estorno'].apply(lambda x: f"{x*100:.0f}%")
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            aba1.to_excel(writer, sheet_name="Estorno por Produto", index=False)
            aba2.to_excel(writer, sheet_name="Resumo por Nota", index=False)
            aba3.to_excel(writer, sheet_name="Totais Gerais", index=False)
            
        return {
            "sucesso": True,
            "resumo": resumo_processamento,
            "preview": preview,
            "arquivo_saida": output_filename
        }
        
    except Exception as e:
        return {"sucesso": False, "erro": f"Erro ao processar: {str(e)}"}
