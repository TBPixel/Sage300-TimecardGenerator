from setuptools import setup, find_packages


with open( 'README.md' ) as f:
    readme = f.read()

with open( 'LICENSE' ) as f:
    license = f.read()

setup(
    name='Sage300-TimecardGenerator',
    version='0.1.0',
    description='A Sage 300 Timecard Generator built in Python',
    long_description=readme,
    author='Tony Barry',
    author_email='TonyMPBarry@gmail.com',
    url='https://github.com/TBPixel/Sage300-TimecardGenerator',
    license=license,
    packages=find_packages( exclude=( 'tests', 'docs' ) ),
    entry_points={
          'gui_scripts': [
              'TimecardGenerator = timecardgenerator.__main__:main'
          ]
    }
)
