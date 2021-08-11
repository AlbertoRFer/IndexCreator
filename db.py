import sqlite3


class Archivo:
    def __init__(self, name):
        """Connects to the database.

        The name of the database is passed as a parameter.
        """
        self.conn = sqlite3.connect(name)

    def __del__(self):
        self.conn.close()

    def create_table(self, table):
        """Creates a new table in the database.

        The name of the table is passed as a parameter.
        Returns true if the table is created correctly, false otherwise.
        """
        exp = """
        CREATE TABLE %s
        (num INT PRIMARY KEY,
         princCompareciente TEXT SECONDARY KEY,
         otrosComparecientes TEXT,
         nombreEscritura TEXT)
        """ % (table,)

        try:
            curs = self.conn.cursor()
            curs.execute(exp)
            return True
        except Exception:
            return False

    def delete_table(self, table):
        """Deletes a table from the database.

        The name of the table is passed as a parameter.
        Returns true if the table is deleted correctly, false otherwise.
        """
        exp = """
        DROP TABLE %s
        """ % (table,)

        try:
            curs = self.conn.cursor()
            curs.execute(exp)
            return True
        except Exception:
            return False

    def view_table(self, table):
        """Returns the content of a table in the database.

        The name of the table is passed as a parameter.
        Returns a list where every element corresponds to a row of the table.
        If an error occurs returns None.
        """
        exp = """SELECT * FROM %s""" % (table,)

        try:
            cur = self.conn.cursor()
            cur.execute(exp)
            return cur.fetchall()
        except Exception:
            return None

    def insert_data(self, table, num, p_name, o_name, n_escr):
        """Inserts a new row in a table.

        The name of the table and all the necessary data are passed
        as parameters.
        Returns true if the data is inserted correctly, false otherwise.
        """
        exp = """INSERT INTO %s
        (num, princCompareciente, otrosComparecientes, nombreEscritura)
        VALUES(?, ?, ?, ?)""" % (table,)

        try:
            cur = self.conn.cursor()
            cur.execute(exp, (num, p_name, o_name, n_escr))

            self.conn.commit()
            return True
        except Exception:
            return False

    def doc_by_number(self, table, number):
        """Gets a row of table filtered by the number of the document.

        The name of the table and the number of the document are passed
        as parameters
        Returns the row of the table that match with the number of the
        document. If an error occurs return None.
        """
        exp = """SELECT * FROM %s WHERE num = ?""" % (table,)

        try:
            cur = self.conn.cursor()
            cur.execute(exp, (number,))

            return cur.fetchall()
        except Exception:
            return None

    def max_number(self, table):
        """Return the maximum number amongst the documents stored in a table.

        The name of the table is passed as a parameter.
        Returns the maximum number amongst the documents stored in a table.
        If an error occurs returns None.
        """
        exp = """SELECT MAX(num) FROM %s""" % table

        # try:
        cur = self.conn.cursor()
        cur.execute(exp)

        return cur.fetchone()
        # except Exception:
        #     return None

    def view_tables(self):
        """Returns the names of all the tables in the database.

        Returns a list of names of the database's tables.
        If an error occurs returns None.
        """
        exp = """SELECT name FROM sqlite_master WHERE type='table'"""

        try:
            cur = self.conn.cursor()
            cur.execute(exp)

            return cur.fetchall()
        except Exception:
            return None

    def get_table_by_letter(self, protocol, letter, number_ini, number_fin):
        exp = """SELECT * 
                FROM %s 
                WHERE princCompareciente 
                LIKE '%s%%' 
                AND %d <= num 
                AND num <= %d""" \
                % (protocol, letter, number_ini, number_fin)

        try:
            cur = self.conn.cursor()
            cur.execute(exp)

            return cur.fetchall()
        except Exception:
            return None

    def delete_data(self, table, number):
        exp = """DELETE FROM %s WHERE num = ?""" % (table,)

        try:
            cur = self.conn.cursor()
            cur.execute(exp, (number,))
            self.conn.commit()

            return True
        except Exception:
            return None

    def search(self, table, field, string):
        sel = """SELECT * 
                 FROM %s 
                 WHERE %s 
                 LIKE '%%%s%%'""" % (table, field, string)

        cur = self.conn.cursor()
        cur.execute(sel)
        return cur.fetchall()


if __name__ == "__main__":
    miArchivo = Archivo('Protocolo.db')
    # miArchivo.insert_data("Protocolo2017", 1, "El Yo", "", "TEST")
    # print(miArchivo.get_table_by_letter("Protocolo2017", "E", 0, 1000))
    # print(miArchivo.delete_data("Protocolo2017", 1))
    # print(miArchivo.get_table_by_letter("Protocolo2017", "E", 0, 1000))
    print(miArchivo.search("Ba"))