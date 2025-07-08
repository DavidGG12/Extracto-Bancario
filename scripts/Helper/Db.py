import os
from sqlalchemy import create_engine, text
import pandas as pd
import pyodbc

class Db:
    def __init__(self):
        self.__svrDb = "10.128.10.19"
        self.__nmDb = "dbTesoreria"
        self.__userDb = "sa"
        self.__pwdDb = "SiFiAd1019"
        # self.__svrDb = os.getenv("SvrDb")
        # self.__nmDb = os.getenv("NmDb")
        # self.__userDb = os.getenv("UserDb")
        # self.__pwdDb = os.getenv("PwdDb")

    def connectionDbAlchemy(self):
        strConnection = (
            f"mssql+pyodbc://{self.__userDb}:{self.__pwdDb}@{self.__svrDb}/{self.__nmDb}"
            "?driver=ODBC+Driver+17+for+SQL+Server&charset=utf8&MultipleActiveResultSets=True"
        )
        return create_engine(strConnection)
    
    def connectionDbPyodbc(self):
        strConnection = "DRIVER={ODBC Driver 17 for SQL Server};" +f"SERVER={self.__svrDb};" +f"DATABASE={self.__nmDb};" + f"UID={self.__userDb};" + f"PWD={self.__pwdDb};" 
        return pyodbc.connect(strConnection, autocommit=True)

    def storedProcedure(self, nameProcedure:str, parameters:dict):
        engine = self.connectionDbPyodbc()
        cursor = engine.cursor()

        paramPlaceHolder = ", ".join(["?" for _ in parameters]) if parameters else ""
        sql = f"EXEC {nameProcedure} {paramPlaceHolder}"

        cleanedParams = []
        if parameters:
            for value in parameters.values():
                if isinstance(value, str):
                    cleanedParams.append(value.replace("'", "") if value else "")
                else:
                    cleanedParams.append(value)
        
        cursor.execute(sql, cleanedParams)
        
        try:
            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()
            return pd.DataFrame.from_records(rows, columns=columns)
        except pyodbc.ProgrammingError:
            return pd.DataFrame()
