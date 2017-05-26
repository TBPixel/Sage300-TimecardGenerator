import logging


class Data( object ):

    def __init__( self ):
        self._data = {}



    #      #
    # DATA #
    #      #
    def add( self, key, data ):
        """
        add( 'name', 'John' )

        Store a new piece of data in the Employee's data

        :param key: str
        :param data: any
        :return: any
        """
        assert not self.data_exists( key ), 'Key \'{0}\' already exists in data'.format( key )

        logging.info( 'Adding data at key \'{0}\' of type {1}'.format( key, type( data ) ) )

        self._data[key] = data

        return self._data[key]


    def data_exists( self, key ):
        """
        data_exists( 'name' )

        Returns if the given key exists in the data dictionary

        :param key: str
        :return: bool
        """
        assert type( key ) == str, 'Key expected type {0}, got {1}'.format( str, type( key ) )

        return key in self._data


    def get( self, key ):
        """
        get( 'name' )

        Retrieves a piece of data from dictionary by key

        :param key: str 
        :return: any
        """
        assert self.data_exists( key ), 'Key \'{0}\' does not exists in data'.format( key )

        logging.info( 'Retrieving data at key \'{0}\''.format( key ) )

        return self._data[key]


    def get_all( self ):
        """
        Retrieves the data dictionary on the instance
        :return: dict
        """
        return self._data


    def remove( self, key ):
        """
        remove( 'name' )

        Removes a piece of data from dictionary by key

        :param key: str
        """
        assert self.data_exists( key ), 'Key \'{0}\' does not exists in data'.format( key )

        logging.info( 'Removing data at key \'{0}\''.format( key ) )

        del self._data[key]
