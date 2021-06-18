from tkinter import Tk

from user_interface.Structure import Structure
from database_interface.DatabaseSQLite import DatabaseSQLite


def get_rec(db):
    """
    :param db: the sqlite database you're looking into
    :return: List of Recycleries
    """
    cur = db.cursor()
    cur.execute("SELECT Recyclerie FROM Organisation")
    return [row[0] for row in cur.fetchall()]


def next_window():
    pass


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

    struct = Structure(frame, lambda: next_window(), get_rec(db_sql.DB))

    frame.mainloop()
