class DataNotAvailable(Exception):
    def __init__(self, table_name):
        self.table_name = table_name
        self.message = 'No new data available for ' + table_name


