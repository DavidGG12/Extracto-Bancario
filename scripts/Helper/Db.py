import os
from sqlalchemy import create_engine, text
import pandas as pd

class Db:
    def __init__(self):
        self.__svrDb = os.getenv("SvrDb")
        self.__nmDb = os.getenv("NmDb")
        self.__userDb = os.getenv("UserDb")
        self.__pwdDb = os.getenv("PwdDb")

    def connectionDb(self):
        strConnection = (
            f"mssql+pyodbc://{self.__userDb}:{self.__pwdDb}@{self.__svrDb}/{self.__nmDb}"
            "?driver=ODBC+Driver+17+for+SQL+Server"
        )
        return create_engine(strConnection)
    
    
    def storedProcedure(self, nameProcedure:str, parameters:dict):
        engine = self.connectionDb()

        paramPlaceHolder = ", ".join([f":{key}" for key in parameters]) if parameters else ""
        sql = text(f"EXEC {nameProcedure} {paramPlaceHolder}")

        cleanedParams = {}
        if parameters:
            for key, value in parameters.items():
                if isinstance(value, str):
                    cleanedParams[key] = value.replace("'", "") if value else ""
                else:
                    cleanedParams[key] = value
        
        with engine.connect() as connection:
            result = connection.execute(sql, **cleanedParams)
            df = pd.DataFrame(result.fetchall(), columns=result.keys())

        return df
