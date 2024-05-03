import sqlite3
from typing import Literal, overload


class BetterSqlite3:
    """
    A wrapper for the sqlite3 module that provides a more intuitive interface.
    """

    def __init__(self, db_file: str):
        """
        Initialize a BetterSqlite3 object.

        Parameters:
            db_file (str): The path to the SQLite database file.

        Returns:
            None

        Example:
            >>> db = BetterSqlite3("mydatabase.db")
        """
        self.db_file = db_file
        self.conn = sqlite3.connect(db_file)
        self.cursor = self.conn.cursor()

    def tableAdd(
        self,
        table_name: str,
        columns: dict[str, Literal["TEXT", "INTEGER", "REAL", "BLOB", "NUMERIC"]],
    ) -> None:
        """
        Add a table to the database.

        Parameters:
            table_name (str): The name of the table to be created.
            columns (dict): A dictionary containing the column names and data types.

        Returns:
            None

        Example:
            >>> columns = {"id": "INTEGER", "name": "TEXT", "age": "INTEGER"}
            >>> db = BetterSqlite3("mydatabase.db")
            >>> db.addTable("people", columns)
        """
        columns_str = ", ".join([f"{k} {v}" for k, v in columns.items()])
        self.cursor.execute(f"CREATE TABLE IF NOT EXISTS {table_name} ({columns_str})")
        self.conn.commit()
        return None

    def dataInsert(self, table_name: str, data: dict) -> None:
        """
        Insert data into a table.

        Parameters:
            table_name (str): The name of the table to insert data into.
            data (dict): A dictionary containing the data to be inserted.

        Returns:
            None

        Example:
            >>> data = {"id": 1, "name": "John", "age": 30}
            >>> db = BetterSqlite3("mydatabase.db")
            >>> db.insert("people", data)
        """
        columns = ", ".join(data.keys())
        placeholders = ", ".join(["?" for _ in data.keys()])
        values = tuple(data.values())
        self.cursor.execute(
            f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})", values
        )
        self.conn.commit()
        return None

    @overload
    def dataSelect(self, table_name: str, where: str, columns: list[str] = ["*"]):
        """
        Select data from a table. Only return records that meet specific criteria.

        Parameters:
            table_name (str): The name of the table to select data from.
            where (str): A SQL condition to specify which rows to select.
            columns (list): A list of column names to select. Defaults to ["*"].

        Returns:
            A list of dictionaries containing the selected data.

        Example:
            >>> db = BetterSqlite3("mydatabase.db")
            >>> db.tableAdd("people", {"id": "INTEGER", "name": "TEXT", "age": "INTEGER"})
            >>> db.dataInsert("people", {"id": 1, "name": "John", "age": 30})
            >>> db.dataInsert("people", {"id": 2, "name": "Jane", "age": 25})
            >>> db.dataSelect("people", "id == 2", ["id", "name"])
            [{"id": 2, "name": "Jane"}]
        """
        columns_str = ", ".join(columns)
        self.cursor.execute(f"SELECT {columns_str} FROM {table_name} WHERE {where}")
        return self.cursor.fetchall()

    def dataSelect(self, table_name: str, columns: list[str] = ["*"]) -> list[dict]:
        """
        Select data from a table.

        Parameters:
            table_name (str): The name of the table to select data from.
            columns (list): A list of column names to select. Defaults to ["*"].

        Returns:
            A list of dictionaries containing the selected data.

        Example:
            >>> db = BetterSqlite3("mydatabase.db")
            >>> data = db.select("people", ["id", "name"])
            >>> print(data)
            [{"id": 1, "name": "John"}, {"id": 2, "name": "Jane"}]
        """
        columns_str = ", ".join(columns)
        self.cursor.execute(f"SELECT {columns_str} FROM {table_name}")
        return self.cursor.fetchall()

    def dataUpdate(self, table_name: str, data: dict, condition: str) -> None:
        """
        Update data in a table.

        Parameters:
            table_name (str): The name of the table to update data in.
            data (dict): A dictionary containing the data to be updated.
            condition (str): A SQL condition to specify which rows to update.

        Returns:
            None

        Example:
            >>> data = {"age": 31}
            >>> condition = "id = 1"
            >>> db = BetterSqlite3("mydatabase.db")
            >>> db.update("people", data, condition)
        """
        set_str = ", ".join([f"{k} = ?" for k in data.keys()])
        values = tuple(data.values()) + (condition,)
        self.cursor.execute(
            f"UPDATE {table_name} SET {set_str} WHERE {condition}", values
        )
        self.conn.commit()
        return None

    def dataDelete(self, table_name: str, condition: str) -> None:
        """
        Delete data from a table.

        Parameters:
            table_name (str): The name of the table to delete data from.
            condition (str): A SQL condition to specify which rows to delete.

        Returns:
            None

        Example:
            >>> condition = "id = 1"
            >>> db = BetterSqlite3("mydatabase.db")
            >>> db.delete("people", condition)
        """
        self.cursor.execute(f"DELETE FROM {table_name} WHERE {condition}")
        self.conn.commit()
        return None

    def dbClose(self) -> None:
        """
        Close the database connection.

        Parameters:
            None

        Returns:
            None

        Example:
            >>> db = BetterSqlite3("mydatabase.db")
            >>> db.close()
        """
        self.conn.close()
        return None

    def dbOpen(self) -> None:
        """
        Open the database connection.

        Parameters:
            None

        Returns:
            None

        Example:
            >>> db = BetterSqlite3("mydatabase.db")
            >>> db.open()
        """
        self.conn = sqlite3.connect(self.db_file)
        self.cursor = self.conn.cursor()
        return None

    def __del__(self):
        self.conn.close()
