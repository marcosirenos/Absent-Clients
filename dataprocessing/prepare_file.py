import pandas as pd
import html5lib
import pyexcel as pe

class FilePreparation:
    def convert_file(self, source):
        df_list = pd.read_html(source, encoding='utf-8')
        df = df_list[1]
        new_header = df.tail(1).values[0]
        df = df.iloc[:-1]
        df.columns = new_header
        df = df.reset_index(drop=True)
        df.to_excel(r"C:\Users\guilherme.oliveira\Documents\GitHub\Absent-Clients\data\raw\MonitorFlexExportacao.xlsx", engine='openpyxl', index=False)
    
