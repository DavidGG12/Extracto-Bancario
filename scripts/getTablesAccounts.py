import pandas as pd
import re
from datetime import datetime, time
from Helper.Db import Db

df = pd.read_excel("INBURSA EJEMPLO EXTRACTO BANCARIO M.N.xlsx", engine="openpyxl")

columns = ["Fecha", "Referencia", "Referencia_Ext", "Referencia_Leyenda", "Referencia_Numerica", "Concepto", "Movimiento", "Cargo", "Abono", "Saldo", "Ordenante", "RFC_Ordenante"] 
account = r"No. Cuenta:"
end = r"MOVIMIENTOS:"
cuenta = []
end = []

for row_idx, row in df.iterrows():
    cellDic = {}
    
    for col_idx, value in row.items():
        
        if re.match(account, str(value)):
            
            col_word = chr(65 + df.columns.get_loc(col_idx))
            cell = f"{col_word}{row_idx + 2}"
            
            cellDic["Column"] = col_word
            cellDic["Row"] = row_idx + 2
            cellDic["Account"] = value.split("No. Cuenta:")[-1].split("|")[0].strip()
            
            cuenta.append(cellDic)
        if re.match(account, str(value)):
            end.append(row_idx)

con = Db()

for i in range(len(cuenta) - 1):
    date:str = datetime.now()
    dateFormat = date.strftime("%d%m%Y_%H%M")

    dicActual = cuenta[i]
    dicSiguiente = cuenta[i+1]
    dfParcial = df.iloc[(dicActual["Row"]+1):(dicSiguiente["Row"]-6)].copy()
    dfParcial.columns = columns
    dfParcial["Cuenta"] = dicActual["Account"]
    dfParcial.to_sql("Tbl_Tesoreria_Temp", con=con.connectionDb(), if_exists="append", index=False, chunksize=100)

dfResult = con.storedProcedure("pa_Tesoreria_InDatos", None)
print(dfResult)


