import os
from sqlalchemy import create_engine

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