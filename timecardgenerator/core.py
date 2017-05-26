import os
import logging
import configparser
import tkinter
import pyodbc

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
        # self._db_connect()

        logging.info( 'Sage 300 Timecard Generator initialized' )

        # Initialize Tkinter Mainloop
        self.gui.mainloop()



    def generate( self ):
        """
        Generate the importable spreadsheet template
        """

        # Retrieve Spreadsheet Data
        # Dates, Employees, Employee Hours etc...


        # Remove employees without hours


        # Retrieve Employee Database data


        # Generate Timesheet
        timecard_name = self.gui.get_widget( 'field_timecard' ).get()



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
            self.create_styled_label( body, text='Timecard:' )
        )

        field_timecard = self.gui.add_widget(
            'field_timecard',
            self.create_styled_entry( body )
        )

        button_run = self.gui.add_widget(
            'button_run',
            self.create_styled_button( footer, text='Run', command=self.generate )
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


    def create_styled_label( self, frame, text ):
        return tkinter.Label(
            frame,
            text=text,
            background='white',
            borderwidth=0,
            font='Arial',
            foreground='#222222'
        )


    def create_styled_entry( self, frame, textvariable=None ):
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


    def create_styled_button( self, frame, text, command ):
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
