import logging

from timecardgenerator import components


class Employee( object ):

    def __init__( self, id, coordinates ):

        self._id = id

        self.coordinates = []
        self.coordinates.append( coordinates )
        self.hours = []
        self.data = components.Data()

        self._set_default_data()

        logging.info( 'Employee \'{0}\' instantiated'.format( self._id ) )


    def add_hours( self, date, duration ):
        """
        Appends new hours to the employee

        :param date: datetime.date
        :param duration: int
        """
        self.hours.append({
            'date': date,
            'hours': duration
        })


    def get_hours_sum( self ):
        sum = 0

        for shift in self.hours:
            sum += self.get_shift_sum(shift['hours'])

        return sum


    def get_shift_sum( self, hours ):
        ( h, m, s ) = hours.strftime( "%H:%M:%S" ).split( ':' )
        return int( h ) + ( int( m ) / 60 )


    def get_shift_count( self ):
        return len( self.hours )


    def get_id( self ):
        """
        Returns the employee's id

        :return: str 
        """
        return self._id


    def _set_default_data( self ):
        self.data.add(
            'AV',
            {
                'key': 'DAYS',
                'value': 1
            }
        )