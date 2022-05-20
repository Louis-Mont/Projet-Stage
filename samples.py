from database_interface.DatabaseSQLite import DatabaseSQLite
import json

if __name__ == '__main__':
    path = "finale.db"

    # DB Init
    db_sql = DatabaseSQLite()
    cr = db_sql.connect(path)
    if not cr[1]:
        raise FileNotFoundError(cr[0])
    print(cr[0])

    rq = db_sql.DB.cursor().execute("SELECT poids_total FROM Arrivage ORDER BY poids_total DESC").fetchall()[:50]

    j = json.dumps(rq)

    print(j)
