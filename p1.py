import pandas as pd
import numpy as np
from scipy.optimize import curve_fit
import matplotlib.pyplot as plt
import os

# --- 1. Carregar e Preparar os Dados ---
# Verifique se o arquivo Excel existe no diretório atual
# **ATENÇÃO: Renomeie o arquivo para o nome original .xlsx, se necessário**
# Por exemplo, se o nome do seu arquivo Excel for realmente "Apendice_II_Banco_de_dados_monitoramento.xlsx"
file_name = "Apendice_II_Banco_de_dados_monitoramento.xlsx"
sheet_name = "BD_Geral" # Nome da aba que contém os dados relevantes

if not os.path.exists(file_name):
    print(f"Erro: O arquivo '{file_name}' não foi encontrado no diretório atual.")
    print("Por favor, verifique se o nome do arquivo está correto (incluindo a extensão .xlsx) e se ele está no mesmo local do script.")
else:
    try:
        # **Alterado de pd.read_csv para pd.read_excel**
        df = pd.read_excel(file_name, sheet_name=sheet_name)

        # Renomear as colunas para facilitar o acesso (ajuste conforme o seu Excel)
        # Assumindo que as colunas relevantes são 'PPV (mm/s)', 'Distância (m)' e 'Carga Maxima por Espera (kg)'
        # Por favor, verifique os nomes exatos das colunas no seu arquivo Excel e ajuste abaixo se necessário.
        df_columns = df.columns.tolist()
        print(f"Colunas disponíveis no Excel (aba '{sheet_name}'): {df_columns}")

        # Mapeamento de colunas para nomes padronizados
        # **Ajustado para usar 'Q' para a carga, conforme solicitado.**
        column_mapping = {
            'PPV (mm/s)': 'PPV',
            'Distancia (m)': 'Distancia',
            'Carga Maxima por Espera (kg)': 'Q' # Renomeado para 'Q'
        }

        # Tentativa de padronizar os nomes das colunas.
        # Se os nomes das colunas no seu Excel forem diferentes, você precisará ajustar este mapeamento.
        for old_name, new_name in column_mapping.items():
            if old_name in df.columns:
                df.rename(columns={old_name: new_name}, inplace=True)
            else:
                print(f"Aviso: Coluna '{old_name}' não encontrada. Verifique o nome da coluna no seu Excel e o mapeamento no script.")

        # Verificar se as colunas necessárias estão presentes após o renomeamento
        # **Ajustado para usar 'Q'**
        required_columns = ['PPV', 'Distancia', 'Q']
        if not all(col in df.columns for col in required_columns):
            print("Erro: Nem todas as colunas necessárias (PPV, Distancia, Q) foram encontradas.")
            print("Verifique os nomes das colunas no seu arquivo Excel e o mapeamento no script.")
        else:
            # Remover linhas com valores nulos nas colunas de interesse
            df_cleaned = df.dropna(subset=['PPV', 'Distancia', 'Q']).copy()

            # Garantir que os dados são numéricos
            df_cleaned['PPV'] = pd.to_numeric(df_cleaned['PPV'], errors='coerce')
            df_cleaned['Distancia'] = pd.to_numeric(df_cleaned['Distancia'], errors='coerce')
            df_cleaned['Q'] = pd.to_numeric(df_cleaned['Q'], errors='coerce')

            # Remover linhas onde a conversão para numérico resultou em NaN
            df_cleaned.dropna(subset=['PPV', 'Distancia', 'Q'], inplace=True)

            # Filtrar dados para evitar logaritmos de zero ou negativos
            # **Usando 'Q'**
            df_filtered = df_cleaned[(df_cleaned['PPV'] > 0) &
                                     (df_cleaned['Distancia'] > 0) &
                                     (df_cleaned['Q'] > 0)].copy()

            if df_filtered.empty:
                print("Não há dados válidos para análise após a filtragem.")
            else:
                # Calcular o Fator de Distância Escalonada (SD)
                # **Usando 'Q'**
                df_filtered['SD'] = df_filtered['Distancia'] / np.sqrt(df_filtered['Q'])

                # Aplicar logaritmo natural para linearizar a equação
                df_filtered['ln_PPV'] = np.log(df_filtered['PPV'])
                df_filtered['ln_SD'] = np.log(df_filtered['SD'])

                # --- 2. Realizar a Regressão Linear ---
                # Definir a função linear para o ajuste
                # Agora A = ln(K) e B = alpha
                def linear_function(x, a, b):
                    return a + b * x

                # Realizar o ajuste da curva (regressão linear)
                # xdata = ln_SD, ydata = ln_PPV
                popt, pcov = curve_fit(linear_function, df_filtered['ln_SD'], df_filtered['ln_PPV'])

                # popt contém os parâmetros ótimos (a e b)
                A = popt[0]  # Onde A = ln(K)
                B = popt[1]  # Onde B = alpha

                # --- 3. Calcular K e alpha ---
                K = np.exp(A)
                alpha = B

                print("\n--- Resultados da Análise ---")
                print(f"O valor de K é: {K:.4f}")
                print(f"O valor de alpha é: {alpha:.4f}")

                # --- 4. Visualizar os Resultados ---
                plt.figure(figsize=(10, 6))
                plt.scatter(df_filtered['ln_SD'], df_filtered['ln_PPV'], label='Dados Originais (ln)')
                plt.plot(df_filtered['ln_SD'], linear_function(df_filtered['ln_SD'], A, B), color='red', label=f'Ajuste Linear: ln(PPV) = {A:.2f} + {B:.2f} * ln(SD)')
                plt.xlabel('ln(Fator de Distância Escalonada - SD)')
                plt.ylabel('ln(PPV)')
                plt.title('Regressão Linearizada para Determinação de K e alpha')
                plt.legend()
                plt.grid(True)
                plt.show()

                # Opcional: Plotar a curva original com os valores K e alpha encontrados
                plt.figure(figsize=(10, 6))
                plt.scatter(df_filtered['SD'], df_filtered['PPV'], label='Dados Originais')
                x_fit = np.linspace(df_filtered['SD'].min(), df_filtered['SD'].max(), 100)
                ppv_fit = K * (x_fit**(alpha))
                plt.plot(x_fit, ppv_fit, color='red', label=f'Ajuste da Curva PPV = {K:.2f} * SD^({alpha:.2f})')
                plt.xlabel('Fator de Distância Escalonada (SD)')
                plt.ylabel('PPV (mm/s)')
                plt.title('Ajuste da Curva PPV Original')
                plt.legend()
                plt.grid(True)
                plt.show()

    except Exception as e:
        print(f"Ocorreu um erro ao processar o arquivo Excel: {e}")
        print("Verifique se o arquivo está no formato Excel correto (.xlsx) e se a aba e as colunas esperadas existem.")
        print("Você pode precisar instalar a engine 'openpyxl' para ler arquivos .xlsx: pip install openpyxl")