from collections import namedtuple


#              #
# Named Tuples #
#              #
Geometry = namedtuple( 'Geometry', 'width height' )
Grid     = namedtuple( 'Grid', 'rows columns' )

SheetDates  = namedtuple( 'SheetWeek', 'Sunday Monday Tuesday Wednesday Thursday Friday Saturday' )
WeekDay     = namedtuple( 'WeekDay', 'date coordinate' )
Coordinate  = namedtuple( 'Coordinate', 'column row' )
Coordinates = namedtuple( 'Coordinates', 'start end' )
