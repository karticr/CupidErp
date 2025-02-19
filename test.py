import pyodbc

class SQLServerConnector:
    def __init__(self):
        """
        Initialize the SQLServerConnector with the connection string for CupidHRNew at 192.168.1.45.
        """
        self.connection_string = (
            "DRIVER={ODBC Driver 17 for SQL Server};"
            "Server=192.168.1.45;"
            "Database=CupidHRNew;"
            "Trusted_Connection=yes;"
        )
        self.connection = None
        self.cursor = None
        self.connect()

    def connect(self):
        """Establish connection to the SQL Server using the connection string."""
        try:
            self.connection = pyodbc.connect(self.connection_string)
            self.cursor = self.connection.cursor()
            print("Database connection established.")
        except Exception as e:
            print(f"Error connecting to database: {e}")
            raise

    def insert_row(self, table_name, date, time, machine, pass_count, reject_count):
        """
        Insert a single row into the specified table.

        :param table_name: Name of the table to insert into
        :param date: Date of the record (DATE)
        :param time: Time of the record (TIME)
        :param machine: Machine name or ID (VARCHAR)
        :param pass_count: Count of passed items (INT)
        :param reject_count: Count of rejected items (INT)
        """
        try:
            sql = f"INSERT INTO {table_name} (Date, Time, Machine, PassCount, RejectCount) VALUES (?, ?, ?, ?, ?)"
            self.cursor.execute(sql, (date, time, machine, pass_count, reject_count))
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
    # Initialize the connector
    db_connector = SQLServerConnector()
    table_name = "Tbl_ET_Record_Auto_Fetch"

    # Example row to insert
    date = "2024-12-06"
    time = "14:30:00"
    machine = "ET03-L1"
    pass_count = 100
    reject_count = 5

    try:
        db_connector.insert_row(table_name, date, time, machine, pass_count, reject_count)
        # Insert more rows as needed
        db_connector.insert_row(table_name, "2024-12-07", "14:45:00", "Machine B", 120, 10)
        print("done  ")
    finally:
        db_connector.close()
