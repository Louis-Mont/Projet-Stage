from utils import dict_to_str, kcv, l_to_str
from abc import ABC, abstractmethod


class Database(ABC):
    def __init__(self):
        self.DB = None

    @abstractmethod
    def connect(self, *args):
        pass

    def __del__(self):
        if self.DB is not None:
            self.DB.close()

    def df(self, table, n_conditions=None):
        """
        Delete all content from the table not in conditions
        :param table: The table you're deleting from
        :type table: str
        :param n_conditions: The no conditions
        :type n_conditions: tuple
        """
        if self.DB is not None:
            self._df(table, n_conditions)
            self.DB.commit()
            return True
        else:
            return False

    def _df(self, table, n_conditions):
        """
        :type table: str
        :type n_conditions: tuple
        """
        cur_ps = self.DB.cursor()
        if n_conditions is None:
            cur_ps.execute(f"DELETE FROM {table}")
        else:
            cur_ps.execute(f"DELETE FROM {table} WHERE {n_conditions[0]} NOT IN ({l_to_str(n_conditions[1])})")

    def ii(self, table, values, conditions=None):
        """
        Inserts into the table the cols using cur or Update if there is conditions
        :param table: The table you're inserting into
        :type table: str
        :param values: The keys represent the titles and cols the cols inserted into said columns
        :type values: dict
        :param conditions: The names of the columns you're checking in n,0 and their cols in n,1
        :type conditions: dict
        :return: False if there is no connection, True if the Insert Into/Update is successful
        """
        if self.DB is not None:
            self._ii(table, values, conditions)
            self.DB.commit()
            return True
        else:
            return False

    def _ii(self, table: str, values: dict, conditions: dict):
        """
        :type table: str
        :type values: dict
        :type conditions: dict
        """
        cur = self.DB.cursor()
        if conditions is None:
            keys, values = dict_to_str(values)
            cur.execute(
                f"INSERT INTO {table} ({keys}) VALUES ({values})")
        else:
            # Update
            u_values = values.copy()
            for c in conditions:
                try:
                    u_values.pop(c)
                except KeyError:
                    pass
            # debug = f"UPDATE {table} SET {kcv(u_values, '=')} WHERE {kcv(conditions, '=', ' AND ')}"
            cur.execute(
                f"UPDATE {table} SET {kcv(u_values, '=')} WHERE {kcv(conditions, '=', ' AND ')}")

    def sf(self, table, cols, distinct=False, conditions=None):
        """
        Selects from the table the cols, distinct or not, depending the conditions
        :param table: The table you're selecting from
        :type table: str
        :param cols: The columns where the cols are
        :type cols: list
        :param distinct:
        :type distinct: bool
        :param conditions: The names of the columns you're checking in n,0 and their cols in n,1
        :type conditions: dict
        :return: False if there is no connection, True if the Insert Into/Update is successful
        """
        if self.DB is not None:
            return self._sf(table, cols, distinct, conditions)
        else:
            return False

    def _sf(self, table: str, cols: list, distinct: bool, conditions: dict):
        """
        :type table: str
        :type cols: list
        :type distinct: bool
        :type conditions: dict
        :rtype: str
        """
        cur = self.DB.cursor()
        where = "" if conditions is None else f" WHERE {kcv(conditions, '=', 'AND')}"
        cur.execute(
            f"SELECT {'DISTINCT ' if distinct else ''}{l_to_str(cols)} FROM {table}{where}")
        return cur.fetchall()
