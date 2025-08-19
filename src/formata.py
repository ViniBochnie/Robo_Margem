import src.Excel as excel

def Formatar(caminho_margem,setup, log):
    if getattr(setup.settings,'Formatar_Excel'):
        log(f'EDITANDO ARQUIVO')
        excel.PersMargem(getattr(setup.settings,'Modelo'),caminho_margem)

    log('PROCESSO FINALIZADO....')