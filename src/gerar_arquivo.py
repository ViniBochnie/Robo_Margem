from datetime import date
import pandas as pd
import os

def GerarArquivo(resultado: dict,setup):
    hoje = date.today()
    caminho_margem = getattr(setup.settings,'Saida') + f'{hoje.day}.{hoje.month}.{hoje.year}' + getattr(setup.settings,'Formato')
    
    existe = os.path.exists(caminho_margem)
    writer_args ={
        "engine": 'openpyxl',
        "mode": 'a' if existe else 'w',
        ** ({"if_sheet_exists": 'overlay'} if existe else {})
    }
    
    with pd.ExcelWriter(caminho_margem, **writer_args) as writer:
        for aba, df in resultado.items():
            df['Data de término'] = pd.to_datetime(df['Data de término'], format='%d/%m/%Y').dt.strftime('%d/%m/%Y')
            df.to_excel(writer, sheet_name=aba.upper(), index=False)
    return caminho_margem