import pandas as pd

def gerar():
    data = {
        "Numero_Nota": [
            "1001", "1001", "1001", "1001", "1001",
            "1002", "1002", "1002", "1002", "1002",
            "1003", "1003", "1003", "1003", "1003"
        ],
        "Data_Nota": [
            "01/10/2023", "01/10/2023", "01/10/2023", "01/10/2023", "01/10/2023",
            "02/10/2023", "02/10/2023", "02/10/2023", "02/10/2023", "02/10/2023",
            "03/10/2023", "03/10/2023", "03/10/2023", "03/10/2023", "03/10/2023"
        ],
        "Fornecedor": [
            "Atacadão XYZ", "Atacadão XYZ", "Atacadão XYZ", "Atacadão XYZ", "Atacadão XYZ",
            "Distribuidora Alfa", "Distribuidora Alfa", "Distribuidora Alfa", "Distribuidora Alfa", "Distribuidora Alfa",
            "Comércio Beta", "Comércio Beta", "Comércio Beta", "Comércio Beta", "Comércio Beta"
        ],
        "Produto": [
            "Arroz Branco 5kg", # Estorno
            "Feijão Carioca 1kg", # Estorno
            "Leite UHT Integral", # Estorno
            "Refrigerante Cola 2L",  # Sem estorno
            "Farinha de Trigo 1kg", # Estorno
            
            "Café Torrado 500g", # Estorno
            "Açúcar Cristal 2kg", # Estorno
            "Queijo Mussarela", # Estorno
            "Óleo de Soja 900ml", # Estorno
            "Biscoito Recheado", # Sem estorno
            
            "Macarrão Espaguete 500g", # Estorno 
            "Manteiga Extra 200g", # Estorno
            "Iogurte Morango", # Estorno
            "Arroz Parboilizado 5kg", # Estorno
            "Sabão em Pó 1kg" # Sem estorno
        ],
        "NCM": [
            "1006.30.21",
            "0713.33.19",
            "0401.20.10",
            "2202.10.00",
            "1101.00.10",
            
            "0901.21.00",
            "1701.14.00",
            "0406.90.10",
            "1507.90.11",
            "1905.31.00",
            
            "1902.11.00",
            "0405.10.00",
            "0403.20.00",
            "1006.30.11",
            "3402.20.00"
        ],
        "Credito_ICMS": [
            25.50,
            12.30,
            40.00,
            18.00, # Sem estorno
            15.00,
            
            30.20,
            20.00,
            55.00,
            22.40,
            10.50, # Sem estorno
            
            14.80,
            8.90,
            12.00,
            28.00,
            25.00 # Sem estorno
        ]
    }
    
    df = pd.DataFrame(data)
    df.to_excel("planilha_exemplo.xlsx", index=False)
    print("Planilha de exemplo gerada com sucesso: planilha_exemplo.xlsx")

if __name__ == "__main__":
    gerar()
