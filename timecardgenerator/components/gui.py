# Packages
import logging
import tkinter
from tkinter import filedialog
from tkinter import messagebox
from timecardgenerator import tuples



class GUI( object ):

    _tkinter_classlist = dict( [ ( name, cls ) for name, cls in tkinter.__dict__.items() if isinstance( cls, type ) ] )


    def __init__( self, title='', geometry=tuples.Geometry( width=0, height=0 ), grid=tuples.Grid( rows=1, columns=1 ) ):

        # Intialize Base Tkinter Object
        self._tk = tkinter.Tk()
        self._tk.title( title )
        self._tk.geometry( '{0}x{1}'.format( geometry.width, geometry.height ) )
        self._tk.rowconfigure( 0, weight=1 )
        self._tk.columnconfigure( 0, weight=1 )

        # Initialize Tkinter Menu
        self._menu = tkinter.Menu( self._tk )

        # Add to root
        self._tk.config( menu=self._menu )

        # Configure Response Container
        self._container = tkinter.Frame( self._tk, width=geometry.width, height=geometry.height )
        self._container.grid( row=0,column=0, sticky='nsew' )
        self.configure_grid( self._container, grid )

        # Declare dictionary for widget creation
        self._widgets = {}

        logging.info( 'GUI instantiated' )


    def destroy( self ):
        """
        Clears all attributes of the instance
        """

        # Clear Widgets
        for key in list( self._widgets.keys() ):
            self.remove_widget( key )

        # Clear Container and Root
        self._menu      = None
        self._container = None
        self._tk        = None

        logging.info( 'GUI Destroyed' )


    def get_container( self ):
        """
        Retrieves the tkinter container Frame

        :return: tkinter.Frame
        """
        logging.info( 'Retrieving GUI Container' )
        return self._container


    def get_menu( self ):
        """
        Retrieves the tkinter Menu

        :return: tkinter.Menu 
        """
        logging.info( 'Retrieving GUI Menu' )
        return self._menu


    def get_root( self ):
        """
        Retrieves the tkinter root

        :return: tkinter.Tk
        """
        logging.info( 'Retrieving GUI root' )
        return self._tk



    #         #
    # tkinter #
    #         #
    def mainloop( self ):
        """
        Initiates the tkinter mainloop
        """
        self._tk.mainloop()


    def load_file( self, filetypes=( ( 'All files', '*.*' ) ) ):
        """        
        Opens the file dialogue prompt

        :return: str
        """
        file = filedialog.askopenfilename( filetypes=filetypes )

        return file


    def show_alert( self, title='Alert', message='' ):
        messagebox.showinfo( title, message )


    def show_error( self, title='Error', message='' ):
        messagebox.showerror( title, message )



    #        #
    # FRAMES #
    #        #
    def configure_grid( self, frame, grid=tuples.Grid( rows=1, columns=1 ) ):

        for row in range( grid.rows ):
            frame.rowconfigure( row, weight=1 )

        for column in range( grid.columns ):
            frame.columnconfigure( column, weight=1 )



    #         #
    # WIDGETS #
    #         #
    def add_widget( self, key, widget ):
        """
        add_widget( 'button', tkinter.Button( container, text='Yes', command=callback ) )

        Stores a new tkinter widget in dictionary for deletion/modification later

        :param key: str
        :param widget: tkinter widget
        :return: tkinter widget
        """
        assert not self.widget_exists( key ), 'Key {0} already exists in GUI widgets'.format( key )
        assert self.validate_widget( widget ), 'Widget must be a valid tkinter widget'

        logging.info( 'Adding widget at key {0} of type {1}'.format( key, type( widget ) ) )

        self._widgets[key] = widget

        return self._widgets[key]


    def widget_exists( self, key ):
        """
        widget_exists( 'button' )

        Returns if the given key exists in the widgets dictionary

        :param key: str
        :return: bool
        """
        assert type( key ) == str, 'Key expected type {0}, got {1}'.format( str, type( key ) )

        return key in self._widgets


    def get_widget( self, key ):
        """
        get_widget( 'button' )

        Retrieves a tkinter widget from dictionary by key

        :param key: str 
        :return: tkinter widget
        """
        assert self.widget_exists( key ), 'Key {0} does not exist in GUI widgets'.format( key )

        logging.info( 'Retrieving widget at key {0}'.format( key ) )

        return self._widgets[key]


    def remove_widget( self, key ):
        """
        remove_widget( 'button' )

        Removes a tkinter widget from dictionary by key

        :param key: str 
        """
        assert self.widget_exists( key ), 'Key {0} does not exist in GUI widgets'.format( key )

        logging.info( 'Removing widget at key {0}'.format( key ) )

        del self._widgets[key]


    def validate_widget( self, widget ):
        """
        validate_widget( tkinter.Button( container, text='Yes', command=callback ) )

        Validates that the passed widget is a valid tkinter widget

        :param widget: tkinter widget 
        :return: bool
        """
        for _, classname in self._tkinter_classlist.items():

            if ( type( widget ) == classname ):
                return True

        return False
