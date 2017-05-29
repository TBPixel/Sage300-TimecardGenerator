import os
import logging
import configparser
import tkinter
import pyodbc
import openpyxl.utils
import arrow
import datetime
import time

# Utilities
from timecardgenerator import components, helpers, models, tuples



# Application Instance
class TimecardGenerator( object ):

    gui = components.GUI(
        title   = 'Sage 300 Timecard Generator',
        geometry= tuples.Geometry( width=320, height=256 ),
        grid    = tuples.Grid( rows=2, columns=1 )
    )


    def __init__( self ):
        # Public Properties
        self.config      = None
        self.spreadsheet = None
        self.dates       = []
        self.employees   = {}


    # __main__ #
    def run( self ):
        """
        Application Entry Point in __main__
        """

        # Read Configuration
        self.config = configparser.ConfigParser()
        config_file = os.path.abspath(
            os.path.join(
                os.path.dirname( __file__ ), '../{0}'.format( 'config/user.ini' )
            )
        )
        self.config.read( config_file )

        logging.info( 'Opened configuration file {0}'.format( config_file ) )

        # Configure GUI
        self._create_gui()

        # Connect to Database
        self._db_connect()

        logging.info( 'Sage 300 Timecard Generator initialized' )

        # Initialize Tkinter Mainloop
        self.gui.mainloop()



    def generate( self ):
        """
        Generate the importable spreadsheet template
        """

        # Retrieve Spreadsheet Data
        # Dates, Employees, Employee Hours etc...
        logging.info( 'Retrieving spreadsheet data...' )
        self.get_dates()
        self.get_employees()
        self.get_hours()


        # Remove employees without hours
        logging.info( 'Ignoring employees without hours this pay-period...' )
        no_hours = []
        for id, employee in self.employees.items():
            if ( len( employee.hours ) == 0 ):
                no_hours.append( id )

        for id in no_hours:
            logging.info( 'Removing employee with id \'{0}\' without hours'.format( id ) )
            del self.employees[id]


        # Retrieve Employee Database data
        logging.info( 'Retrieving Employee database data...' )
        self.query_employee_data()


        # Generate Timesheet
        logging.info( 'Generating Timecard \'{0}\''.format( self.gui.get_widget( 'field_timecard' ).get() ) )
        self.spreadsheet.generate( employees=self.employees )



    #     #
    # GUI #
    #     #
    def open_spreadsheet( self ):
        """
        Opens the file dialog prompt to select a spreadsheet
        """

        # Ensures no active spreadsheet is set
        if not ( self.spreadsheet == None ): self.close_spreadsheet()

        # Opens file selection dialogue
        file = self.gui.load_file(
            filetypes=(
                ( 'Excel files', '*.xlsx' ),
                ( 'All files', '*.*' )
            )
        )

        # Do not open spreadsheet if no file selected
        if ( file == None  ) or ( file == '' ):
            logging.info( 'Canceled' )
            return

        logging.info( 'Opening spreadsheet file {0}'.format( file ) )

        self.spreadsheet = models.Spreadsheet( file )


    def close_spreadsheet( self ):
        """
        Clears the currently active spreadsheet
        """
        if ( self.spreadsheet == None ): return

        logging.info( 'Closing file {0}'.format( self.spreadsheet.file ) )

        self.spreadsheet = None


    def _create_gui( self ):
        """
        Instantiates the GUI
        """
        logging.info( 'Creating GUI' )

        # Create File Menu #
        menu = self.gui.get_menu()

        file_menu = tkinter.Menu( menu, tearoff=0 )
        file_menu.add_command( label='Open Spreadsheet',  command=self.open_spreadsheet )
        file_menu.add_command( label='Close Spreadsheet', command=self.close_spreadsheet )
        file_menu.add_command( label='Run',               command=self.generate )
        file_menu.add_separator()
        file_menu.add_command( label='Quit',              command=self.gui.get_root().quit )

        menu.add_cascade( label='File', menu=file_menu )


        # Create Containers #
        container = self.gui.get_container()

        body = self.gui.add_widget(
            'frame_body',
            tkinter.Frame( container, width=320, height=224, bg='white' )
        )

        footer = self.gui.add_widget(
            'frame_footer',
            tkinter.Frame( container, width=320, height=32, bg='white' )
        )

        # Configure Responsive Containers
        self.gui.configure_grid( frame=body, grid=tuples.Grid( rows=1, columns=2 ) )
        self.gui.configure_grid( frame=footer )


        # Add Inputs #
        field_timecard_label = self.gui.add_widget(
            'field_timecard_label',
            self._create_styled_label( body, text='Timecard:' )
        )

        field_timecard = self.gui.add_widget(
            'field_timecard',
            self._create_styled_entry( body )
        )

        button_run = self.gui.add_widget(
            'button_run',
            self._create_styled_button( footer, text='Run', command=self.generate )
        )


        # Configure Grid #
        # Containers
        body.grid( row=0, column=0, sticky='nsew' )
        footer.grid( row=1, column=0, sticky='nsew' )

        # Inputs
        pad  = {
            'padx': 4,
            'pady': 4
        }
        ipad = {
            'ipadx': 4,
            'ipady': 4
        }
        field_timecard_label.grid( row=0, column=0, sticky='new', **pad )
        field_timecard.grid( row=0, column=1, sticky='new', **ipad )

        button_run.grid( row=0, column=0, sticky='nsew', **ipad )


    def _create_styled_label( self, frame, text ):
        return tkinter.Label(
            frame,
            text=text,
            background='white',
            borderwidth=0,
            font='Arial',
            foreground='#222222'
        )


    def _create_styled_entry( self, frame, textvariable=None ):
        return tkinter.Entry(
            frame,
            textvariable=textvariable,
            background='white',
            borderwidth=1,
            highlightbackground='#444444',
            highlightthickness=2,
            highlightcolor='#3E80B4',
            font='Arial',
            foreground='#222222',
            relief=tkinter.FLAT
        )


    def _create_styled_button( self, frame, text, command ):
        return tkinter.Button(
            frame,
            text=text,
            command=command,
            activebackground='#3E80B4',
            activeforeground='#FFFFFF',
            background='#70A2CA',
            borderwidth=0,
            disabledforeground='#FFFFFF',
            font='Arial',
            foreground='#FFFFFF',
            relief=tkinter.FLAT
        )


    #          #
    # Database #
    #          #
    def _db_connect( self ):
        """
        Connects to the available database in config
        """

        logging.info( 'Connecting to Database...' )

        driver   = self.config.get( 'DB', 'DRIVER' )
        server   = self.config.get( 'DB', 'SERVER' )
        database = self.config.get( 'DB', 'DATABASE' )
        username = self.config.get( 'DB', 'UID' )
        password = self.config.get( 'DB', 'PWD' )

        # Database Connection
        self._dbConnect = pyodbc.connect( 'DRIVER={0};SERVER={1};DATABASE={2};UID={3};PWD={4}'.format(
            driver,
            server,
            database,
            username,
            password
            )
        )
        self._db = self._dbConnect.cursor()


    def query_employee_data( self ):

        for id, employee in self.employees.items():

            # Retrieves database data for employees by their id
            self._db.execute("""
            SELECT
                employee.OTSCHED,
                detail.CATEGORY,
                dist.EARNDED,
                dist.DISTCODE,
                dist.EXPACCT,
                dist.OTACCT
            FROM CPEMPL employee, CPEMPD detail, CPDIST dist
            WHERE employee.EMPLOYEE = detail.EMPLOYEE
            AND detail.EARNDED		= dist.EARNDED
            AND dist.DISTCODE		= detail.DISTCODE
            AND dist.AUDTUSER		= employee.AUDTUSER
            AND dist.EARNDED		= ?
            AND employee.EMPLOYEE	= ?
            """,
            'HRLY',
            id )

            data = self._db.fetchone()

            if ( data == None ):
                self.gui.show_error(
                    title='Database Query Error!',
                    message='Could not locate employee with id \'{0}\' in DB\n\n Will skip employee to continue...'.format( id )
                )
                continue

            coordinates = [ 'Y',        'E',       'F',       'BB',       'T',       'V'  ]
            keys        = [ 'OTSCHED', 'CATEGORY', 'EARNDED', 'DISTCODE', 'EXPACCT', 'OTACCT' ]

            for i, value in enumerate( data ):
                # Strip whitespace
                if ( type( value ) == str ):
                    value = value.replace( ' ', '' )
                    value = value.replace( '-', '' )

                    try:
                        value = int( value )
                    except:
                        pass

                # Add Employee data
                employee.data.add(
                    coordinates[i],
                    {
                        'key': keys[i],
                        'value': value
                    }
                )



    #             #
    # SPREADSHEET #
    #             #
    def _get_employee_coordinate( self, coordinate ):
        """
        Return possible coordinate of employee given a coordinate row to check
        
        :param coordinate: tuples.Coordinate
        :return: tuples.Coordinate
        """
        sheet_min = openpyxl.utils.coordinate_from_string( self.spreadsheet.min_coordinate() )
        sheet_min = tuples.Coordinate( column=sheet_min[0], row=sheet_min[1] )

        return tuples.Coordinate( column=sheet_min.column, row=coordinate.row )


    def _get_employee_from_coordinate( self, coordinate ):
        """
        Return employee found at given coordinate
        
        :param coordinate: tuples.Coordinate
        :return: models.Employee
        """
        coord = self._get_employee_coordinate( coordinate )
        key   = self.spreadsheet.get_cell( '{0}{1}'.format( coord.column, coord.row ) )

        if not key in self.employees:
            return

        return self.employees[key]


    def _is_datetime( self, data ):
        """
        Return if passed data is datetime object
        
        :param data: any
        :return: boolean
        """
        # Datetime object
        if ( not type( data ) == datetime.datetime ):
            return False

        # Datetime string
        if ( type( data ) == str ):
            try:
                datetime.datetime.strptime( data, '%x' )
            except:
                return False

        return True


    def _is_time( self, data ):
        """
        Return if passed data is time object
        
        :param data: data
        :return: boolean
        """
        # Type Checking
        if data == None or type( data ) == int or type( data ) == bool:
            return False

        # string or datetime object
        if type( data ) == str or type( data ) == datetime.datetime:
            try:
                date = arrow.get( data )
            except:
                return False

            excel = arrow.get( 1899, 12, 31 )

            if date > excel:
                return False

        return True


    def _is_greater_than_zero_time( self, data ):
        """
        Return if passed data is time greater than 00:00

        :param data: data
        :return: boolean
        """
        # Type Checking
        if data == None or type( data ) == int or type( data ) == bool:
            return False

        # string or datetime object
        if type( data ) == str or type( data ) == datetime.datetime or type( data ) == datetime.time:
            try:
                date = arrow.get( data )
            except:
                date = arrow.get( datetime.time.strftime( data, '%Y-%m-%d %H:%M:%S' ) )

            zero = arrow.get( 1899, 12, 30 )

            if date <= zero:
                return False

        return True

    def _is_end_of_week( self, week: dict, weekday: arrow ):
        """
        Return if passed weekday is in the given week
        
        :param week: dict
        :param weekday: arrow
        :return: boolean
        """
        if ( not weekday in week ):
            return False

        return True


    def _is_employee_row( self, coordinate ):
        """
        Return if given coordinate exists within employees coordinates list
        
        :param coordinate: tuples.Coordinate
        :return: boolean
        """
        sheet_coordinate = self.spreadsheet.min_coordinate()
        coord            = tuples.Coordinate( column=sheet_coordinate[0], row=coordinate.row )
        employee         = self._get_employee_from_coordinate( coord )

        if not employee:
            return False

        if not coord in employee.coordinates:
            return False

        return True


    def get_dates( self ):
        """
        Retrieves dates from the spreadsheet

        :returns: namedtuple
        """
        assert not ( self.spreadsheet == None ), 'Spreadsheet must be set before employees can be retrieved'

        logging.info( 'Retrieving Dates...' )

        weeks = [{}]

        def loop( cell, coordinate ):

            # Data is datetime and not datetime.time
            if self._is_datetime( cell.value ) and ( not self._is_time( cell.value ) ):

                day     = arrow.get( cell.value )
                weekday = day.format( 'dddd' )

                # Creates a new week when an existing day has been found in the latest week
                if self._is_end_of_week( weeks[-1], weekday ):
                    weeks.append({})

                # Append new weekday to week dict
                weeks[-1][weekday] = tuples.WeekDay( date=day, coordinate=coordinate )

        self.spreadsheet.loop_sheet(
            tuples.Coordinates( start=self.spreadsheet.min_coordinate(), end=self.spreadsheet.max_coordinate() ),
            loop
        )

        # Extracts the weeks into namedtuples
        for i, week in enumerate( weeks ):
            weeks[i] = tuples.SheetDates( **week )

        logging.info( 'Retrieved workdays from spreadsheet' )

        self.dates = weeks


    def get_employees( self ):
        """
        Retrieves employees from spreadsheet
        """
        assert not ( self.spreadsheet == None ), 'Spreadsheet must be set before employees can be retrieved'

        logging.info( 'Retrieving Employees...' )

        employees = {}
        periodend = self.dates[-1][-1].date.date()
        payperiod = self.gui.get_widget( 'field_timecard' ).get()

        def loop( cell, coordinate ):
            if not cell.value in employees:

                new_employee = models.Employee( id=cell.value, coordinates=coordinate )
                new_employee.data.add( 'A', {
                        'key': 'id',
                        'value': cell.value
                    }
                )
                new_employee.data.add( 'B', {
                        'key': 'periodend',
                        'value': periodend
                    }
                )
                new_employee.data.add( 'C', {
                        'key': 'payperiod',
                        'value': payperiod
                    }
                )
                employees[cell.value] = new_employee
            else:
                employees[cell.value].coordinates.append( coordinate )

        # Get starting coordinate of loop
        start = self.spreadsheet.min_coordinate()

        # Modify end to encompass only the first column
        end = openpyxl.utils.coordinate_from_string( self.spreadsheet.max_coordinate() )
        end = '{0}{1}'.format( start[0], end[1] )

        self.spreadsheet.loop_sheet(
            tuples.Coordinates( start=start, end=end ),
            loop
        )

        logging.info( 'Retreived employees from Spreadsheet' )

        self.employees = employees


    def get_hours( self ):
        """
        Acquires each employees clocked hours from the spreadsheet
        """
        logging.info( 'Retrieving Employee Hours...' )

        def loop( cell, coordinate ):

            if self._is_time( cell.value ) and self._is_greater_than_zero_time( cell.value ) and self._is_employee_row( coordinate ):
                # Get Date
                date     = None
                employee = self._get_employee_from_coordinate( coordinate )

                # Loop weeks
                for week in self.dates:
                    # Loop days of week
                    for day in week:
                        # Ensure same column and find the oldest parental row of dates
                        if ( day.coordinate.column == coordinate.column ) and ( day.coordinate.row < coordinate.row ):
                            date = day

                # Append new hours to employee
                employee.add_hours( date, cell.value )

        self.spreadsheet.loop_sheet(
            tuples.Coordinates( start=self.spreadsheet.min_coordinate(), end=self.spreadsheet.max_coordinate() ),
            loop
        )

        logging.info( 'Retrieved employee hours' )
