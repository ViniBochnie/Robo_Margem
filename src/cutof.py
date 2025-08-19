import pandas as pd

class CutOf:
    def __init__(self, log):
        self.log = log
        
    def GerarCutof(self,caminho_margem,setup):
        all_sheets = pd.read_excel(caminho_margem, sheet_name=None, engine='openpyxl')
        dfCutof = pd.concat(all_sheets.values(),axis=0,ignore_index=True)
        
        for cli in getattr(setup.cutof,'Cliente'):
            self.log(f'Gerando Cutof para o cliente: {cli}')
            caminho_cutof=getattr(setup.cutof,'Saida') + getattr(setup.cutof,'Arquivo') + f' {cli}' + getattr(setup.cutof,'Formato')
            cutof= dfCutof[dfCutof[getattr(setup.cutof,'Coluna_Cliente')].str.contains(cli,na=False)]

            col_index_map = {col_name: i for i, col_name in enumerate(cutof.columns)}
            col=col_index_map["PREVISÃO DE ENTREGA"]
            cutof.iloc[:,col] = pd.to_datetime(cutof["PREVISÃO DE ENTREGA"], format='%d/%m/%Y').dt.strftime('%d/%m/%Y')
            
            cutof.to_excel(caminho_cutof,index=False)
