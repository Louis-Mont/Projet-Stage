from tkinter import Tk
from user_interface.Extract import Extract
from user_interface.Structure import Structure
from database_interface.DatabaseSQLite import DatabaseSQLite


def get_rec(db):
    """
    :param db: the sqlite database you're looking into
    :type db: DatabaseSQLite
    :return: List of Recycleries
    :rtype: list
    """
    return [row[0] for row in db.sf('Organisation', ['Recyclerie'])]


def get_cat(db):
    """
    :param db: the sqlite database you're looking into
    :type db: DatabaseSQLite
    :return: List of Categories
    :rtype: list
    """
    return [row[0] for row in db.sf('Catégorie', ['Catégorie'])]


def get_origin(db):
    """
    :param db: the sqlite database you're looking into
    :type db: DatabaseSQLite
    :return: List of Origins
    :rtype: list
    """
    return [row[0] for row in db.sf('Arrivage', ['origine'], True) if row[0] != 0]


def get_min_y(db):
    """
    :param db: the sqlite database you're looking into
    :type db: DatabaseSQLite
    :return: minimum year in the DB
    :rtype: int
    """
    return db.sf('Arrivage', ["Min(date)"])[0][0][:4]


def next_window(frame, gdr_db):
    """
    :type frame: Tk
    :type gdr_db: DatabaseSQLite
    """
    frame.destroy()
    min_year = get_min_y(gdr_db)
    ext = Extract(frame, min_year, None, None, get_cat(gdr_db), get_origin(gdr_db))

    frame.mainloop()


if __name__ == '__main__':
    path = "finale.db"
    # only one frame or mainloop will open other Windows
    frame = Tk()

    # DB Init
    db_sql = DatabaseSQLite()
    cr = db_sql.connect(path)
    if not cr[1]:
        raise FileNotFoundError(cr[0])
    print(cr[0])

    struct = Structure(frame, lambda: next_window(frame, db_sql), get_rec(db_sql))

    frame.mainloop()
