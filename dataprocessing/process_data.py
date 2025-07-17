import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import time
import datetime
import os
import threading
import subprocess
import gspread
from google.oauth2 import service_account
from openpyxl import load_workbook
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

scopes = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

json_file = Path("dataprocessing", "credentials.json") # <-------------- Change to actual location


class DataProcessing:

    def login(self):
        credentials = service_account.Credentials.from_service_account_file(json_file)
        scoped_credentials = credentials.with_scopes(scopes)
        gc = gspread.authorize(scoped_credentials)
        return gc

    # Function that performs the main code execution and updates status
    def execute_code(self, file_path):
        print("Starting data processing")
        # Function to load an Excel file with specific parameters
        def load_excel(file_param):
            file_path = file_param['file_path']
            kwargs = file_param.get('kwargs', {})

            teste = os.path.splitext(file_path)
            df = None # Initialize df to avoid potential UnboundLocalError

            if teste[1] == '.xlsx':
                df = pd.read_excel(file_path, **kwargs)
                # Convert to string BEFORE applying .str.replace()
                if 'INV(000)' in df.columns:
                    df['INV(000)'] = df['INV(000)'].astype(str)
                if 'Inserção' in df.columns:
                    df['Inserção'] = df['Inserção'].astype(str)
                
                # Now apply replacements safely
                if 'INV(000)' in df.columns:
                    df['INV(000)'] = df['INV(000)'].str.replace('.', '', regex=False)
                    df['INV(000)'] = df['INV(000)'].str.replace(',', '.', regex=False)
                if 'Inserção' in df.columns:
                    df['Inserção'] = df['Inserção'].str.replace('.', '', regex=False)
                    df['Inserção'] = df['Inserção'].str.replace(',', '.', regex=False)
                print(df)

            elif teste[1] == '.xls':
                #read the hmtl(xls) file:
                dfs = pd.read_html(file_path, header=None)
                df = dfs[1]
                new_header = df.iloc[-1]
                print(new_header)
                df = df[:-1]
                df.columns = new_header
                df = df.reset_index(drop=True)

                # pre data treatment:
                # For .xls, pd.read_html often brings data in as strings,
                # but it's safer to ensure it, especially if 'INV(000)' or 'Inserção'
                # might sometimes be numeric for some reason.
                if 'INV(000)' in df.columns:
                    df['INV(000)'] = df['INV(000)'].astype(str)
                if 'Inserção' in df.columns:
                    df['Inserção'] = df['Inserção'].astype(str)

                if 'INV(000)' in df.columns:
                    df['INV(000)'] = df['INV(000)'].str.replace('.', '', regex=False)
                    df['INV(000)'] = df['INV(000)'].str.replace(',', '.', regex=False)
                if 'Inserção' in df.columns:
                    df['Inserção'] = df['Inserção'].str.replace('.', '', regex=False)
                    df['Inserção'] = df['Inserção'].str.replace(',', '.', regex=False)
                print(df)
                
            return df


        file_param = {
            'file_path': file_path,
            'kwargs': {} 
        }

        # 2. Call your custom load_excel function directly.
        df = load_excel(file_param)

        gc = self.login()
        planilha = gc.open("COBERTURA")
        aba = planilha.worksheet("COB")
        dados = aba.get_all_records()
        Cdf = pd.DataFrame(dados)

        time.sleep(2)
        
        print("Perfoming processing...")

        time.sleep(3)

        df = df.rename(columns={'Emissora Radio' : 'Emissora TV'})

        # Strip whitespace while preserving NaN values
        df['Emissora TV'] = df['Emissora TV'].str.strip()
        df['Anunciante'] = df['Anunciante'].str.strip()
        df['Marca'] = df['Marca'].str.strip()
        df['Agência'] = df['Agência'].str.strip()
        df['Marca'] = df['Marca'].str.strip()

        #correcting emissora that came with wrong data
        mask = df['Emissora TV'] == 'RECORD'
        df.loc[mask, 'Emissora TV'] = 'RECORD TV'

        time.sleep(3)

        df['INV(000)'] = df['INV(000)'].astype(np.float32)
        df['Inserção'] = df['Inserção'].astype(np.int16)
        df['Ano-Mês'] = df['Ano-Mês'].astype(np.int32)
        df['Praça'] = df['Praça'].astype('category')
        df['Emissora TV'] = df['Emissora TV'].astype('category')
        df['Categoria'] = df['Categoria'].astype('category')
        df['Tipo Veiculação'] = df['Tipo Veiculação'].astype('category')

        time.sleep(3)

        conditions = (
            df['Anunciante'].str.contains(r'\bGLOBO\b', case=False, regex=True) |
            df['Anunciante'].str.contains('BANDEIRANTES') |
            df['Anunciante'].str.contains('RECORD') |
            df['Anunciante'].str.contains('GAZETA') |
            df['Anunciante'].str.contains('CNT') |
            df['Anunciante'].str.contains('JOVEM PAN') |
            df['Anunciante'].str.contains('RADIO') |
            df['Anunciante'].str.contains('MASSA') |
            df['Anunciante'].str.contains('TELEVISAO') |
            df['Anunciante'].str.contains('TV') |
            df['Anunciante'].str.contains('SBT') |
            df['Anunciante'].str.contains('RIC') |
            (df['Marca'].str.contains('CARTOLA', na=False)) |
            df['Anunciante'].str.contains('{DESCONHECIDO}') |
            df['Anunciante'].str.contains('TSE') |
            df['Categoria'].str.contains('CAMPANHAS BENEFICIENTES SOCIAIS') |
            df['Categoria'].str.contains('CAMPANHAS PARTIDARIAS') |
            (df['Marca'].str.contains('BEECON')) |
            (df['Marca'].str.contains('TOPVIEW'))
        )

        df.drop(df[conditions].index, inplace=True)

        #function to extract a value from the row (finding the city name from the prefecture name)
        def extract_value(row):
            start = 'PREF MUN '
            end = ' (GMP)'
            value = row['Anunciante']
            if start in value and end in value:
                extracted_value = value.split(start)[1].split(end)[0]
                return extracted_value
            else:
                return row['Cidade Autorização']

        #apply the city name from the correct prefecture to the city column
        df['Cidade Autorização'] = df.apply(extract_value, axis=1)

        #rename the column from the gov est pr to match it's location
        df['UF Autorização'] = np.where(df['Anunciante'].str.contains('GOV EST PR (GEP)', 
                                                                    regex=False), 'PARANA', df['UF Autorização'])
        df['Cidade Autorização'] = np.where(df['Anunciante'].str.contains('GOV EST PR (GEP)', 
                                                                        regex=False), 'CURITIBA', df['Cidade Autorização'])

        #correcting some clients that came with wrong data
        mask = df['Marca'] == 'MUFFATAO'
        df.loc[mask, 'Anunciante'] = 'PEDRO MUFFATO & CIA LTDA'

        clients_maringa = (
            (df['Marca'] == 'CANCAO') |
            df['Marca'].str.contains('AMIGAO', na=False) |
            df['Marca'].str.contains('ZAELI', na=False) |
            ((df['Marca'].str.contains('IMPACTO', na=False)) & (df['Praça'].str.contains('MARINGA', na=False))) |
            ((df['Marca'].str.contains('SOLUCIONADOR', na=False)) & (df['Praça'].str.contains('MARINGA', na=False))) |
            df['Marca'].str.contains('COAMO', na=False) |
            df['Marca'].str.contains('UNICESUMAR', na=False) |
            df['Marca'].str.contains('ORAL SIN', na=False)
        )
        clients_curitiba = (
            df['Marca'].str.contains('PEDROSO', na=False) |
            (df['Anunciante'] == 'ALTHAIA') |
            (df['Marca'].str.contains('ORAL UNIC', na=False)) & (df['Praça'].str.contains('CURITIBA', na=False)) |
            (df['Anunciante'] == 'LIGGA TELECOM') |
            df['Marca'].str.contains('CARRERA CARNEIRO', na=False) |
            (df['Marca'].str.contains('ORAL SIN', na=False)) & (df['Praça'].str.contains('CURITIBA', na=False))
        )
        clients_cascavel = (
            (df['Marca'].str.contains('IMPACTO PRIME', na=False)) & (df['Praça'].str.contains('CASCAVEL', na=False)) |    
            (df['Marca'].str.contains('SUPERGASBRAS ENERGIA', na=False)) & (df['Praça'].str.contains('CASCAVEL', na=False)) |
            df['Marca'].str.contains('ARMAZEM DA MARIA', na=False) |
            (df['Marca'].str.contains('BLUEFIT', na=False)) & (df['Praça'].str.contains('CASCAVEL', na=False)) |
            df['Anunciante'].str.contains('CBS EMPREENDIMENTOS IMOBILIARIOS', na=False) |
            (df['Marca'].str.contains('UMUPREV', na=False)) & (df['Praça'].str.contains('CASCAVEL', na=False)) |
            (df['Marca'].str.contains('ORAL UNIC', na=False)) & (df['Praça'].str.contains('CASCAVEL', na=False)) |
            df['Marca'].str.contains('SHOPPING CHINA', na=False) |
            df['Marca'].str.contains('ODONTO SAN', na=False) |
            (df['Anunciante'].str.contains('CIA BEAL DE ALIMENTOS', na=False)) & (df['Praça'].str.contains('CASCAVEL', na=False)) |
            df['Anunciante'].str.contains('IND E COM DE LATICINIOS PEREIRA', na=False) |
            (df['Marca'] == 'MUFFATAO') |
            (df['Marca'] == 'ITAIPU BINACIONAL') |
            (df['Marca'] == 'FOZ TINTAS') |
            (df['Marca'] == 'UNIPRIME') |
            (df['Anunciante'] == 'FRIMESA')
        )
        clients_londrina = (
            df['Anunciante'].str.contains('SUPER MUFFATO', na=False) |
            df['Marca'].str.contains('SOLUCAO', na=False)
        )

        # Rename the city column that commonly mismatches its real location
        df.loc[clients_maringa, ['Cidade Autorização', 'UF Autorização']] = ['MARINGA', 'PARANA']
        df.loc[clients_curitiba, ['Cidade Autorização', 'UF Autorização']] = ['CURITIBA', 'PARANA']
        df.loc[clients_londrina, ['Cidade Autorização', 'UF Autorização']] = ['LONDRINA', 'PARANA']
        df.loc[clients_cascavel, ['Cidade Autorização', 'UF Autorização']] = ['CASCAVEL', 'PARANA']

        mask = df['Marca'] == 'MAX' 
        df.loc[mask, 'Agência'] = 'NOBRE PROPAGANDA'

        #creates a few columns
        df['Vl Tab (000)'] = df['INV(000)'] * 1000
        df['Desconto'] = 0
        df['Valor Líquido Projetado'] = 0
        df['Cobertura'] = None
        df['Região'] = None
        df['Mercado'] = None
        df['Data'] = pd.to_datetime(df['Ano-Mês'], format='%Y%m')

        df.reset_index(inplace=True, drop=True)

        print("Creating and processing columns ALMOST complete")

        #function to determine coverage:
        def determine_coverage(row, Cdf):
            cidade_autorizacao = row['Cidade Autorização'].upper()
            
            if 'COBERTURA' in Cdf.columns and pd.notna(Cdf.loc[Cdf['Municipio'] == cidade_autorizacao, 'COBERTURA']).any():
                return Cdf.loc[Cdf['Municipio'] == cidade_autorizacao, 'COBERTURA'].values[0]
            else:
                return 'IMPORT'
            
        #function to determine our coverage region:
        cidade_region_map = dict(zip(Cdf['Municipio'], Cdf['Região']))

        planilha = gc.open("MERCADO TV CLIENTES AJUSTES DIRETORIA")
        aba = planilha.worksheet("CURITIBA")
        dados = aba.get_all_records()
        market_cwb = pd.DataFrame(dados)

        market_cwb['ANUNCIANTE'] = market_cwb['ANUNCIANTE'].str.strip()

        print(market_cwb)

        #function to determine our market:
        def set_market(row):
            UF = row['UF Autorização']
            anunciante = row['Anunciante'].upper()
            cidade_autorizacao = row['Cidade Autorização'].upper()

            if UF == 'PARANA' and 'PREF' not in anunciante and anunciante not in market_cwb['ANUNCIANTE'].str.upper().values:
                return 'LOCAL'
            elif UF != 'PARANA':
                return 'IMPORT'

            if 'PREF' in anunciante and (cidade_autorizacao in ['CURITIBA', 'MARINGA', 'CASCAVEL', 
                                                                'TOLEDO', 'FOZ DO IGUACU', 'LONDRINA']):
                return 'PREF SEDE'
            elif 'PREF' in anunciante:
                return 'PREF'
            elif 'GOV' in anunciante and 'FEDERAL' not in anunciante:
                return 'GOVERNO'
            elif '(GEP)' in anunciante and ('ASSEMBLEIA' not in anunciante):
                return 'GOVERNO'
            elif 'ASSEMBLEIA' in anunciante:
                return 'ASSEMBLEIA'
            
            # Handle the market_cwb case more robustly
            matches = market_cwb[market_cwb['ANUNCIANTE'].str.upper() == anunciante]
            if not matches.empty:
                # Take the first match if there are multiple
                return matches.iloc[0]['MERCADO']

        #fill some of our Columns:
        df['Cobertura'] = df.apply(determine_coverage, args=(Cdf,), axis=1)
        df['Região'] = df['Cidade Autorização'].map(cidade_region_map).fillna('IMPORT')
        df['Mercado'] = df.apply(set_market, axis=1)

        print("almost...")

        df['Mercado'] = df['Mercado'].astype('category')
        df['Cobertura'] = df['Cobertura'].astype('category')
        df['Região'] = df['Região'].astype('category')

        print("Applying discount on column")
        
        def discount_giver(row):
            emissora = row['Emissora TV']
            coverage = row['Cobertura']
            praca = row['Praça']
            anunciante = row['Anunciante']
            
            #SBT discounts
            if 'SBT' in emissora and 'MERCHANDISING' not in praca:
                return 0.94
            elif 'SBT' in emissora and 'MERCHANDISING' in praca:
                return 0.93
            #BAND discounts
            elif 'BANDEIRANTES' in emissora and 'MERCHANDISING' not in praca and (praca in ['LONDRINA', 'MARINGA', 'CURITIBA']):
                return 0.95
            elif 'BANDEIRANTES' in emissora and 'MERCHANDISING' in praca and (praca in ['LONDRINA', 'MARINGA', 'CURITIBA']):
                return 0.95
            elif 'BANDEIRANTES' in emissora and 'MERCHANDISING' not in praca and 'CASCAVEL' in praca:
                return 0.95
            elif 'BANDEIRANTES' in emissora and 'MERCHANDISING' in praca and 'CASCAVEL' in praca:
                return 0.95
            #CNT discounts
            elif 'CNT' in emissora and 'CURITIBA' in praca:
                return 0.9
            #GLOBO discounts
            elif 'GLOBO' in emissora and 'MERCHANDISING' not in praca and (praca in ['MARINGA', 'LONDRINA', 'FOZ DO IGUACU']):
                return 0.3
            elif 'GLOBO' in emissora and 'MERCHANDISING' in praca and (praca in ['MARINGA', 'LONDRINA', 'FOZ DO IGUACU']):
                return 0.29
            elif 'GLOBO' in emissora and 'MERCHANDISING' not in praca and (praca in ['PARANAVAI', 'PONTA GROSSA', 'GUARAPUAVA']):
                return 0.4
            elif 'GLOBO' in emissora and 'MERCHANDISING' not in praca and (praca in ['CURITIBA', 'CASCAVEL']):
                return 0.3
            elif 'GLOBO' in emissora and 'MERCHANDISING' in praca and (praca in ['CURITIBA', 'CASCAVEL']):
                return 0.24
            
            #Import discounts
            elif 'BANDEIRANTES' in emissora and 'IMPORT' in coverage:
                return 0.95
            elif 'SBT' in emissora and 'IMPORT' in coverage:
                return 0.9
            elif 'GLOBO' in emissora and 'IMPORT' in coverage:
                return 0.1
            elif 'CNT' in emissora and 'IMPORT' in coverage:
                return 0.96
            elif 'RECORD' in emissora and 'IMPORT' in coverage:
                return 0.82
            
            #Specific Discounts
            #SBT discounts:
            elif 'SBT' in emissora and 'CONDOR' in anunciante and 'MERCHANDISING' not in praca:
                return 0.89
            elif 'SBT' in emissora and 'CONDOR' in anunciante and 'MERCHANDISING' in praca:
                return 0.88
            elif 'SBT' in emissora and 'SUPER MUFFATO' in anunciante and 'MERCHANDISING' not in praca:
                return 0.89
            elif 'SBT' in emissora and 'SUPER MUFFATO' in anunciante and 'MERCHANDISING' in praca:
                return 0.88
            elif 'SBT' in emissora and 'ALIANCA' in anunciante and 'MERCHANDISING' not in praca:
                return 0.9950
            elif 'SBT' in emissora and 'ALIANCA' in anunciante and 'MERCHANDISING' in praca:
                return 0.93
            elif 'SBT' in emissora and 'RIO VERDE' in anunciante:
                return 0.98
            elif 'SBT' in emissora and 'ODONTO EXCELLENCE' in anunciante:
                return 0.96
            #BAND discounts
            elif 'BANDEIRANTES' in emissora and 'MALUCELLLI' in anunciante:
                return 0.98
            elif 'BANDEIRANTES' in emissora and 'PONTO DE VISAO' in anunciante:
                return 0.9850
            elif 'BANDEIRANTES' in emissora and 'O SOLUCIONADOR' in anunciante:
                return 0.98
            elif 'BANDEIRANTES' in emissora and 'SUPER MUFFATO' in anunciante:
                return 0.97
            #GLOBO discounts
            elif 'GLOBO' in emissora and 'COORITIBA FOOT BALL CLUB' in anunciante:
                return 0.5
            elif 'GLOBO' in emissora and 'PONTO DE VISÃO' in anunciante:
                return 0.7
            elif 'GLOBO' in emissora and 'KURTEN' in anunciante:
                return 0.55
            elif 'GLOBO' in emissora and 'JOCKEY PLAZA SHOP' in anunciante:
                return 0.5
            #GOV discounts
            elif 'CNT' not in emissora and 'RECORD' not in emissora and 'GOV' in anunciante:
                return 0.13
            elif 'CNT' in emissora and 'GOV' in anunciante:
                return 0.50
            #ASSEM discounts
            elif 'CNT' not in emissora and 'RECORD' not in emissora and 'ASSEMBLEIA' in anunciante:
                return 0.13
            elif 'CNT' in emissora and 'ASSEMBLEIA' in anunciante:
                return 0.50
            #PREF discounts
            elif 'GLOBO' in emissora and 'PREF MUN CURITIBA (GMP)' in anunciante:
                return 0.15
            elif 'BANDEIRANTES' in emissora and 'PREF MUN CURITIBA (GMP)' in anunciante:
                return 0.20
            elif 'SBT' in emissora and 'PREF MUN CURITIBA (GMP)' in anunciante:
                return 0.20
            elif 'CNT' in emissora and 'PREF MUN CURITIBA (GMP)' in anunciante:
                return 0.55
            else:
                return 0

        df['Desconto'] = df.apply(discount_giver, axis=1)

        # Apply some discounts
        mask1 = (df['Desconto'] != 0) & (df['Agência'] == '{DIRETO}')  # Check for direct agency
        mask2 = (df['Desconto'] != 0) & (df['Agência'] != '{DIRETO}')  # Check for non-direct agency

        df['Valor Líquido Projetado'] = 0  # Initialize the column with zeros

        df.loc[mask1, 'Valor Líquido Projetado'] = df['Vl Tab (000)'] * (1 - df['Desconto'])
        df.loc[mask2, 'Valor Líquido Projetado'] = df['Vl Tab (000)'] * (1 - df['Desconto']) * (1 - 0.2)

        
        # print(df.dtypes)
        # print(df)

        print("Making absent clients report")
        #Create the Pivot Table.            
        Adf = pd.pivot_table(df,values='Vl Tab (000)', index=['Anunciante', 'Marca', 'Agência', 'Categoria', 'UF Autorização', 'Cidade Autorização', 'Cobertura', 'Mercado'],columns=['Emissora TV'], aggfunc='sum', fill_value=0, observed=True)            
        Adf['Total'] = Adf.sum(axis=1)            
        Adf = Adf.reset_index()             
        #Creates the dataframes filtered by 'Cobertura'            
        # status_elem.print("Criando Relatório PR")            
        Adf_PR = Adf[Adf['Cobertura'] != 'IMPORT']            
        print("PR report done")   
        Adf_CWB = Adf[Adf['Cobertura'] == 'CTBA']            
        print("CWB report done")          
        Adf_LON = Adf[Adf['Cobertura'] == 'LON']            
        print("LON report done")            
        Adf_OES = Adf[Adf['Cobertura'] == 'OESTE']            
        print("OESTE report done")           
        Adf_MAR = Adf[Adf['Cobertura'] == 'MAR']            
        #Creates a List to drop the Column Cobertura from these DFs            
        dataframes = [Adf, Adf_PR, Adf_CWB, Adf_LON, Adf_OES, Adf_MAR]            
        #Drops the Column from the list            
        dataframes= [Adf.drop(columns='Cobertura', inplace=True) for Adf in dataframes]          
        # Create or get the file path            
        print("Finishing XLSX report")             
        file_path = Path("data", "processed", f"{datetime.datetime.now().strftime('%B')}_.xlsx")            
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:            
            # Save each DataFrame to a separate worksheet                
            Adf.to_excel(writer, sheet_name='GERAL', index=False)               
            Adf_PR.to_excel(writer, sheet_name='PARANÁ', index=False)                
            Adf_CWB.to_excel(writer, sheet_name='CURITIBA', index=False)               
            Adf_LON.to_excel(writer, sheet_name='LONDRINA', index=False)                
            Adf_OES.to_excel(writer, sheet_name='OESTE', index=False)               
            Adf_MAR.to_excel(writer, sheet_name='MARINGÁ', index=False)                
            df.to_excel(writer, sheet_name='MENSAL', index=False)          
            print("Data processing done")

