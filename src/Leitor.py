import pandas as pd
from datetime import date
from os.path import exists as ex
import os
from .Config import Settings

class Leitor:
     def __init__(self, console = None):
          self.log = console or print  # Use console if provided, otherwise default to print
     
     def Loader(self):
          try:
               settings= Settings.Ler()
          except:
               self.log(
                    'ERRO AO LER CONFIGURAÇÕES...'
               )

          return settings

     def Carregar(self,settings,base):
          lista = []
          lista.append(base)
          analistas=getattr(settings,'Analista')
          for filial in getattr(settings,base):
               for an in analistas:
                    arqiuivo=getattr(settings,base)[filial]['Arquivo'] + '_' + an
                    caminho = os.path.join('Temp',arqiuivo + getattr(settings,'Formato'))
                    if ex(caminho):

                         temp_df = pd.read_excel(caminho,nrows=0) #cabecalho
                         codex/add-logging-for-column-order-mismatch
                         temp_cols = list(temp_df.columns)
                         expected = getattr(settings,'Coluna_padrao')
                         temp_cols = [c.strip().lower() for c in temp_df.columns]
                         expected = [c.strip().lower() for c in settings.Coluna_padrao]
                         main

                         if temp_cols == expected:
                              self.log(f'CARREGANDO {filial.upper()}')
                              locals()[filial] = pd.read_excel(caminho)
                              locals()[filial] = locals()[filial].drop(locals()[filial].index[-1])
                              self.log(str(locals()[filial].shape[0]) + ' LINHAS CARREGADAS')
                              locals()[filial].Name = filial
                              lista.append(locals()[filial])
                              break
                         else:
                              self.log(f'Colunas fora de ordem no arquivo {filial}. Corrija e tente novamente...')
                              self.log(f"Esperado: {expected}")
                              self.log(f"Encontrado: {temp_cols}")
                    else:
                         continue

          return lista

     def CriarColuna(self,lista,settings):
          base = lista[0]
          for item in lista[1:]:
               self.log(f'CRIANDO COLUNA PARA {item.Name.upper()}')

               for col in settings.Colunas:
                    pos = settings.Colunas[col]['Posicao']
                    nome= settings.Colunas[col]['Nome']

                    escreva = getattr(settings,base)[item.Name][col]
                    item.insert(pos,nome,escreva,True)

          return lista

     def Unir(self,lista):
          self.log(f'UNINDO {lista[0].upper()}')
          lista.pop(0)
          df = pd.concat(lista, ignore_index=True,axis=0)

          return df

     def FormataBase(self,df,settings):
          self.log(f'ADICIONANDO DATA')
          hoje = date.today()
          for rename in getattr(settings,'Renomear'):
               valor = getattr(settings,'Renomear')[rename]
               df.rename(columns={rename:valor},inplace=True)

          df['Data'] = f'{hoje.day}/{hoje.month}/{hoje.year}'

          return df

     def Cte(self,df,settings):
          df[getattr(settings,'Coluna_Cte')] = df[getattr(settings,'Coluna_Cte')].str[getattr(settings,'I_Cte'):getattr(settings,'F_Cte')]
          return df