import pythoncom
import googlemaps
import sqlite3
import os
from datetime import date, datetime
from win32com.client import DispatchEx


def create_distance_table():
    conn = sqlite3.connect("distancias.db")
    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS distancias (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        origem TEXT NOT NULL,
        destino TEXT NOT NULL,
        distancia REAL
    );
    """)

    conn.commit()
    conn.close()

def update_distances(sheet, setup):
    row = 2
    
    conn = sqlite3.connect("distancias.db")
    cursor = conn.cursor()
    
    gmaps = googlemaps.Client(key=setup.gmaps.api_key)

    latlong_destino_col = 'AI'
    cidade_origem_col = 'V'
    km_column = 'AN'

    while sheet.Cells(row, "AK").Value != None:
        km = sheet.Cells(row, km_column).Value
        cidade = sheet.Cells(row, cidade_origem_col).Value
        latlong_origem = getattr(setup.ceps, cidade.strip().upper().replace(' ', '_') , None)
        
        lat_o = latlong_origem['latitude']
        long_0 = latlong_origem['longitude']
        
        origem = f'{lat_o},{long_0}'
        
        if km == -2146826246:
            latlong_destino = sheet.Cells(row, latlong_destino_col).Value
            content = latlong_destino.strip().lstrip('(').rstrip(')')
            parts = content.split(',')
            lat_d = float(parts[0]+'.'+parts[1])
            long_d = float(parts[2]+'.'+parts[3])
            
            destino = f'{lat_d},{long_d}'

            if all([origem, destino]):
                cursor.execute("SELECT distancia FROM distancias WHERE origem = ? AND destino = ?", (origem,destino))
                distance = cursor.fetchone()

                if distance is None:
                    try:
                        matrix = gmaps.distance_matrix(origins=origem, destinations=destino, mode='driving',units='metric')
                        elem = matrix['rows'][0]['elements'][0]
                        if elem.get('status') == 'OK':
                            distance = round(elem['distance']['value'] / 1000,2)
                    except Exception as e:
                        print(e)
                        distance = None

                    if distance != None:
                        cursor.execute("INSERT INTO distancias (origem, destino, distancia) VALUES (?, ?,?)", (origem, destino, distance))
                        conn.commit()

                else:
                    sheet.Cells(row, km_column).Value = distance[0]

        row += 1

    conn.close()

def GerarIqt(app) -> list[str]:
    margem = app.caminho
    setup = app.setup
    create_distance_table()  # Cria a tabela de distâncias se não existir
    
    pythoncom.CoInitialize()  # Inicializa o COM para o thread atual
    excel=DispatchEx('Excel.Application')
    excel.Visible=setup.iqt.Exibir_Excel
    excel.DisplayAlerts=False
    
    #abre margem
    bMargem = excel.Workbooks.Open(margem)
    
    lista= [bMargem.Worksheets(l).Name for l in range(1,bMargem.Worksheets.Count + 1)]
    saida = []
    for l in lista:
        sMargem = bMargem.Worksheets(l)

        #abre template
        bTemp = excel.Workbooks.Open(getattr(setup.settings,'Modelo') + f'modeloIqt-{l}.xlsx')
        
        sTemp = bTemp.Worksheets(getattr(setup.iqt,'Planilha'))
        
        #deletar todas as linhas da tabela
        tabela=getattr(setup.iqt,'Tabela')
        rows=sTemp.Range(tabela).Rows.Count
        if rows>1:
            sTemp.Activate
            sTemp.Range(tabela).Select()
            excel.Selection.EntireRow.Delete()
        
        #copia a margem
        sMargem.Activate()
        sMargem.Range("A1").AutoFilter(Field=5,Criteria1="8TOMBO",Criteria2="9TOMBO",Operator=2)
        sMargem.AutoFilter.Range.Select()
        excel.Selection.Copy(Destination=sTemp.Range("A1"))
    
        #data tombo
        data = date.today()
        data=datetime.strptime(f'{data.year}-{data.month}-{data.day}','%Y-%m-%d',)
        inicio=datetime.strptime('1900-01-01','%Y-%m-%d')
        sTemp.Range(f"{tabela}[DATA TOMBO]").Value = abs((data-inicio).days + 2)
        sTemp.Range(f"{tabela}[DATA TOMBO]").NumberFormat = "[$-pt-BR]d-mmm;@"
        
        #atualiza km
        update_distances(sTemp, setup)
    
        #salvar arquivo na pasta Base
        out_dir = os.path.join(setup.iqt.Saida, "Base")
        os.makedirs(out_dir, exist_ok=True)

        # 8) Monta o nome completo com extensão
        hoje = date.today()
        nome = f"IQT {l} {hoje.day}-{hoje.month}-{hoje.year}.xlsx"
        out_path = os.path.join(out_dir, nome)
        out_path = os.path.normpath(out_path)
        
        exists = os.path.exists(out_path)
        if not exists:
            # 9) Salva especificando Filename e FileFormat (51 = .xlsx)
            #OBS: aqui usamos DispatchEx, então chamamos sobre o objeto de Workbook bTemp
            bTemp.SaveAs(Filename=out_path, FileFormat=51)

            saida.append(out_path)
    
    bMargem.Close()
    excel.Application.Quit()
    return saida
