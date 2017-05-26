import logging
import openpyxl


class Spreadsheet( object ):

    def __init__( self, file ):
        self.file = file
