iams2rf
==========

Tools for converting IAMS records to Researcher Format.

## Installation

From GitHub:

    git clone https://github.com/victoriamorris/iams2rf
    cd iams2rf

To install as a Python package:

    python setup.py install
    
To create stand-alone executable (.exe) files for individual scripts:

    python setup.py py2exe
    
Executable files will be created in the folder iams2rf\dist, and should be copied to an executable path.
    
## Usage

### Running scripts

The following scripts can be run from anywhere, once the package is installed

#### snapshot2sql

Converting the IAMS Published Snapshot to an SQL database:

    Usage: snapshot2sql -i IAMS_SNAPSHOT_PATH -d DB_PATH [OPTIONS] 

    Convert IAMS_SNAPSHOT_PATH to an SQL database at DB_PATH.

    Options:
      --debug   Debug mode.
      --help    Show help message and exit.


#### sql2rf

Searching for records within an SQL database created using snapshot2sql
and converting to Researcher Format:

    Usage: sql2rf -d DB_PATH -r REQUEST_PATH -o OUTPUT_FOLDER [OPTIONS]

    Search DB_PATH for records meeting criteria in REQUEST_PATH.

    If REQUEST_PATH is not specified you will be given the option to set parameters for the output.
    Depending upon the parameters set in REQUEST_PATH, or input by the user,
    some or all of the following files will be created:
        * records_IAMS.csv
        * names_IAMS.csv
        * titles_IAMS.csv
        * topics_IAMS.csv

    Options:
      -o        OUTPUT_FOLDER to save output files.
      --debug   Debug mode.
      --help    Show help message and exit.
