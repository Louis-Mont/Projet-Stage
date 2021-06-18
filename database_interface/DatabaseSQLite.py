from database_interface.Database import Database
import sqlite3


class DatabaseSQLite(Database):
    def __init__(self):
        super().__init__()

    def connect(self, path):
        """
        Connects to a DSN
        :param path: The path where the file is
        :type path: str
        :return: Logs,True|False if connection is successful or not
        """
        try:
            self.DB = sqlite3.connect(path)
            return "Connection successful", True
        except sqlite3.Error as Er:
            return f"{Er.args[1]}", False
