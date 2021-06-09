from . import Database
import sqlite3


class DatabaseSQLite(Database.Database):
    def __init__(self):
        super().__init__()

    def connect(self, path):
        """
        Connects to a DSN
        :param path: The path where the file is
        :type path: str
        :return: True if the connection is successful, False if it isn't
        """
        try:
            self.DB = sqlite3.connect(f"DSN={path};")
            return True
        except sqlite3.Error:
            return False
