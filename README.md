![reform_logo](https://cloud.githubusercontent.com/assets/25346275/22784940/00e54ccc-eeca-11e6-80cd-4d003e7e5361.png)

# marc2rf
Tools for converting MARC records to Researcher Format 

## Requirements

Requires the regex module from https://bitbucket.org/mrabarnett/mrab-regex. The built-in re module is not sufficient.

## Installation

From GitHub:

    git clone https://github.com/victoriamorris/marc2rf
    cd marc2rf

To install as a Python package:

    python setup.py install
    
To create stand-alone executable (.exe) files for individual scripts:

    python setup.py py2exe
    
Executable files will be created in the folder marc2rf\dist, and should be copied to an executable path.

## Usage

### Running scripts

The following scripts can be run from anywhere, once the package is installed:

#### write_rf_config

MARC record selection for Researcher Format.
This utility prepares config files for selection of MARC records to convert to Researcher Format.
    
    Usage: write_rf_config -r REQUEST_PATH [OPTIONS]
    
    Write config files for criteria in REQUEST_PATH.

    Options:
      -o        OUTPUT_FOLDER to save output files.
      --debug   Debug mode.
      --help    Show help message and exit.

#### researcherFormat

MARC record conversion for Researcher Format.
This utility transforms a file of MARC records to Researcher Format.
    
    Usage: researcherFormat -i MARC_PATH -r REQUEST_PATH [OPTIONS]
    
    Convert MARC_PATH to Researcher Format with parameters set in REQUEST_PATH.
    
    If REQUEST_PATH is not specified you will be given the option to set parameters for the output.
    Depending upon the parameters set in REQUEST_PATH, or input by the user, 
    some or all of the following files will be created:
        * records.csv
        * names.csv
        * titles.csv
        * topics.csv    
        * classification.csv
        
    Options:
    
    At most one of ...
      -d       Default transformation.
      -b       Default transformation for BNB records.
      -c       Default transformation for Music records.
      -e       Default transformation for ESTC records.
      -f       Default transformation for FRBRization.
      -m       Use MARC fields instead of column headings.
      -n       Default transformation for Newspaper records.
    
    Any of ...    
      -o        OUTPUT_FOLDER to save output files.
      --debug   Debug mode.
      --help    Show help message and exit.       
    
    Output files differ for FRBRization and MARC field options.


### Notes
 
The file specified in REQUEST_PATH must be an Outlook message submitted via the online form http://www.mappamorris.co.uk/researcherFormat/RFdatasetrequest.php and saved in the format 'Outlook Message Format - Unicode (*.msg)'

MARC input files must have .lex file extensions.
