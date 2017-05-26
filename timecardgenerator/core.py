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

                    print( value )

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
    def get_dates( self ):
        """
        Retrieves dates from the spreadsheet

        :returns: namedtuple
        """
        assert not ( self.spreadsheet == None ), 'Spreadsheet must be set before employees can be retrieved'

        weeks = [{}]

        def loop( cell, coordinate ):

            if ( type( cell.value ) == datetime.datetime ):

                day = arrow.get( cell.value )

                if ( day > arrow.get( time.gmtime(0) ).replace( day=+1 ) ):

                    weekday = day.format( 'dddd' )

                    # Creates a new week when an existing day has been found in the latest week
                    if ( weekday in weeks[-1] ):
                        weeks.append({})

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

        employees = {}
        periodend = self.dates[-1][-1].date.date()
        payperiod = self.gui.get_widget( 'field_timecard' ).get()

        def loop( cell, coordinate ):
            if not ( cell.value == None ):

                if not ( cell.value in employees ):
                    new_employee = models.Employee( id=cell.value, coordinates=cell.coordinate )
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
                    employees[cell.value].coordinates.append( cell.coordinate )

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
        employees = self.employees

        def loop( cell, coordinate ):

            if not ( cell.value == None ) and ( type( cell.value ) == datetime.time ):
                #
                employee_coordinate = openpyxl.utils.coordinate_from_string( self.spreadsheet.min_coordinate() )
                employee_coordinate = '{0}{1}'.format( employee_coordinate[0], coordinate[1] )
                employee_key        = self.spreadsheet.get_cell( employee_coordinate )

                if ( employee_coordinate in employees[employee_key].coordinates ):
                    # Get Date
                    date = None

                    for week in self.dates:
                        for day in week:
                            if ( day.coordinate.column == coordinate.column ) and ( day.coordinate.row < coordinate.row ):
                                date = day

                    # Append new hours to employee
                    self.employees[employee_key].add_hours( date, cell.value )

        #
        self.spreadsheet.loop_sheet(
            tuples.Coordinates( start=self.spreadsheet.min_coordinate(), end=self.spreadsheet.max_coordinate() ),
            loop
        )

        logging.info( 'Retrieved employee hours' )
