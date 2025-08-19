import shutil
import glob
from src import *
import os
from time import sleep
from tkinter import filedialog

def handle_margem(app):
    try:
        locais = ["Filial", "Dtg"]
        app.logger.log("Baixando arquivos da margem.")
        app.spu.download(
            sharepoint_path=app.setup.settings.DOWNLOAD_PATH,
            biblioteca=app.setup.azure.BIBLIOTECA,
            files = locais
        )
        app.logger.log("Arquivos baixados com sucesso. Iniciando o processamento dos arquivos.")
        
        resultado = {}
        for local in locais:
            lc = app.leitor.Carregar(app.setup.settings, local)
            lc = app.leitor.CriarColuna(lc, app.setup.settings)
            lc = app.leitor.Unir(lc)
            lc = app.leitor.FormataBase(lc,app.setup.settings)
            lc = app.leitor.Cte(lc, app.setup.settings)
            resultado[local] = lc
        app.caminho = GerarArquivo(resultado, app.setup)
        Formatar(app.caminho, app.setup, app.logger.log)
        limpar_pasta_temp(app)
        #app.spu.upload(
        #    app.caminho,
        #    app.setup.settings.SHAREPOINT_PATH,
        #    app.setup.azure.BIBLIOTECA
        #)
        app.logger.log("Arquivo enviados com sucesso.\n")
    except Exception as e:
        app.logger.log(f"Erro ao processar os arquivos: {e}\n")

def handle_filial(app):
    try:
        local = "Filial"
        files = list(getattr(app.setup.settings, local).keys())
        files = list(set(f.replace('_', '').upper() for f in files))
        app.logger.log("Baixando arquivos da margem.")
        app.spu.download(
            sharepoint_path=app.setup.settings.DOWNLOAD_PATH,
            biblioteca=app.setup.azure.BIBLIOTECA,
            files=files
        )
        app.logger.log("Arquivos baixados com sucesso. Iniciando o processamento dos arquivos.")
        
        resultado = {}
        lc = app.leitor.Carregar(app.setup.settings, local)
        lc = app.leitor.CriarColuna(lc, app.setup.settings)
        lc = app.leitor.Unir(lc)
        lc = app.leitor.FormataBase(lc,app.setup.settings)
        lc = app.leitor.Cte(lc, app.setup.settings)
        resultado[local] = lc
        app.caminho = GerarArquivo(resultado, app.setup)
        Formatar(app.caminho, app.setup, app.logger.log)
        limpar_pasta_temp(app)
        #app.spu.upload(
        #    app.caminho,
        #    app.setup.settings.SHAREPOINT_PATH,
        #    app.setup.azure.BIBLIOTECA
        #)
        app.logger.log("Arquivo enviados com sucesso.\n")
    except Exception as e:
        app.logger.log(f"Erro ao processar os arquivos: {e}\n")

def handle_dtg(app):
    try:
        local = "Dtg"
        files = list(getattr(app.setup.settings, local).keys())
        files = list(set(f.replace('_', '').upper() for f in files))
        app.logger.log("Baixando arquivos da margem.")
        app.spu.download(
            sharepoint_path=app.setup.settings.DOWNLOAD_PATH,
            biblioteca=app.setup.azure.BIBLIOTECA,
            files=files
        )
        app.logger.log("Arquivos baixados com sucesso. Iniciando o processamento dos arquivos.")
        
        lc = app.leitor.Carregar(app.setup.settings, local)
        lc = app.leitor.CriarColuna(lc, app.setup.settings)
        lc = app.leitor.Unir(lc)
        lc = app.leitor.FormataBase(lc, app.setup.settings)
        lc = app.leitor.Cte(lc, app.setup.settings)
        resultado = {local: lc}
        app.caminho = GerarArquivo(resultado, app.setup)
        Formatar(app.caminho, app.setup, app.logger.log)
        limpar_pasta_temp(app)
        #app.spu.upload(
        #    app.caminho,
        #    app.setup.settings.SHAREPOINT_PATH,
        #    app.setup.azure.BIBLIOTECA
        #)
        app.logger.log("Arquivo enviados com sucesso.\n")
    except Exception as e:
        app.logger.log(f"Erro ao processar os arquivos: {e}\n")

def handle_cutof(app):
    if getattr(app.setup.cutof,'GerarCutof'):
        if not getattr(app, 'caminho', None):
            file_path = filedialog.askopenfilename(
                title="Selecione o arquivo de margem",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if file_path:
                app.caminho = file_path
                app.logger.log(f"Arquivo de margem selecionado: {app.caminho}")
            else:
                app.logger.log("Nenhum arquivo selecionado. Abortando a geração do Cutoff.")
                return
        app.Cutof.GerarCutof(app.caminho, app.setup)
    
def handle_iqt(app):
    app.logger.log("Gerando IQT.")
    
    if  not getattr(app, 'caminho', None):
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo de margem",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            app.caminho = file_path
            app.logger.log(f"Arquivo de margem selecionado: {app.caminho}")
        else:
            app.logger.log("Nenhum arquivo selecionado. Abortando a geração do IQT.")
            return
    
    if getattr(app.setup.iqt,'GerarIqt'):
        arquivos = GerarIqt(app)
        app.logger.log("IQT gerado com sucesso.")
        
        for a in arquivos:
            #app.spu.upload(
            #    a,
            #    app.setup.iqt.SHAREPOINT_PATH,
            #    app.setup.azure.BIBLIOTECA
            #)
            sleep(2)  # Pequena pausa para evitar sobrecarga no SharePoint
            
def limpar_pasta_temp(app):
    """Remove todos os arquivos da pasta Temp."""
    temp_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'Temp')
    temp_path = os.path.normpath(temp_path)
    for file_path in glob.glob(os.path.join(temp_path, '*')):
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            app.logger.log(f"Erro ao remover {file_path}: {e}")