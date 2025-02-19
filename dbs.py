import pyodbc

class SQLServerConnector:
    def __init__(self, server, database, username, password):
        self.server = server
        self.database = database
        self.username = username
        self.password = password
        self.connection = None
        self.cursor = None
        self.connect()

    def connect(self):
        try:
            self.connection = pyodbc.connect(
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={self.server};"
                f"DATABASE={self.database};"
                f"UID={self.username};"
                f"PWD={self.password}"
            )
            self.cursor = self.connection.cursor()
            print("Database connection established.")
        except Exception as e:
            print(f"Error connecting to database: {e}")
            raise

    def insert_row(self, table_name, row):
        try:
            placeholders = ", ".join(["?"] * len(row))
            sql = f"INSERT INTO {table_name} VALUES ({placeholders})"
            self.cursor.execute(sql, row)
            self.connection.commit()
            print(f"Row inserted into {table_name}.")
        except Exception as e:
            print(f"Error inserting row: {e}")
            self.connection.rollback()

    def close(self):
        """Close the database connection."""
        if self.cursor:
            self.cursor.close()
        if self.connection:
            self.connection.close()
        print("Database connection closed.")

# Example Usage
if __name__ == "__main__":
    server = "your_server"
    database = "your_database"
    username = "your_username"
    password = "your_password"
    table_name = "your_table"

    row = (1, 'John Doe', 30)

    db_connector = SQLServerConnector(server, database, username, password)
    try:
        db_connector.insert_row(table_name, row)
        # Insert more rows as needed
        db_connector.insert_row(table_name, (2, 'Jane Smith', 25))
    finally:
        db_connector.close()
