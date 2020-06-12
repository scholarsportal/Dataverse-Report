import psycopg2

from metrics.config import config


class DatabaseConnection(object):
    """ Database Connection instance. """

    def __init__(self):
        self._params = config.read('postgresql')
        self.connection = None

    def __enter__(self):
        # print('Connecting to postgres database...')
        self.connection = psycopg2.connect(**self._params)
        return self.connection

    def __exit__(self, type, value, traceback):
        if traceback is None:
            self.connection.commit()
        else:
            self.connection.rollback()
        self.connection.close()
        # print('Database connection closed.')
