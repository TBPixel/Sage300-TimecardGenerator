import sys

from timecardgenerator import TimecardGenerator

def main( args=None ):
    """The main routine."""
    if args is None:
        args = sys.argv[1:]

    # Run Application
    app = TimecardGenerator()

    app.run()


if __name__ == '__main__':
    main()
