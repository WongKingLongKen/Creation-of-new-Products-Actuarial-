# premium-holiday-table-conversion-v1.0.py
'''
This python script transform the downloaded data from as400 into prophet-readable .fac table format.
'''
import logging
import os
import traceback
import csv
import pandas as pd

def get_filenames():
    '''
    This function returns the data source downloaded from as400
        Args:
            None
        Returns:
            a list of file names that are downloaded from as400 (List)
    '''
    return [
        'premium-holiday.csv',
        'PH policy list MC.csv',
        'Reinstate.csv',
        'Reinstate MC.csv',
    ]

def get_dtype():
    '''
    This function gives a dictionary of corresponding data type for data from as400
        Args:
            None
        Returns:
            data type dictionary (Dict): dictionary with - 
                key: names of columns needed from source
                values: corresponding data type
    '''
    return {
        'CHDRNUM': 'str',
        'Historical no. of months': 'Int64',
        'Current PH Start Date': 'Int64',
        'Current PH End Date': 'Int64',
        'Lapse Start Date': 'Int64',
        'Lapse End Date': 'Int64',
    }

def get_output_cols():
    '''
    This function provides a dictionary that the column we need for final output;
    and corresponding renameed column label.
    '''
    return {
        'CHDRNUM': 'POL_NUMBER',
        'Historical no. of months': 'PAST_PH_M',
        'Current PH Start Date': 'CUR_PH_START',
        'Current PH End Date': 'CUR_PH_END',
    }

def to_fac(df, filename):
    '''
    This function adds the column at the beginning of tables with values = "*" and
    outputs the table as '.fac'.
        Args:
            df (DataFrame): table of the premium holiday
        returns:
            as table output in .fac extension ready for prophet to read
    '''
    logging.info('Writing to %s ...', filename)
    df.insert(0, '!2', '*') # insert a column with header "!2" and value "*" 
    df.to_csv(filename, index=None, sep=',', quoting=csv.QUOTE_NONE)

def read_and_validate_csv(filename):
    '''
    This function reads data as csv into dataframe if possible.
    Warning message would be returned if file is not found.
        Args:

    '''
    try:
        logging.info('Reading from %s...', filename)
        return pd.read_csv(filename, dtype=get_dtype())
    except FileNotFoundError as file_not_found_err:
        logging.info(traceback.format_exc())
        logging.warning(file_not_found_err)

def main():
    '''
    This main function controls the major work flow of the program.
    The main purpose and logic of this python script are included in this function.
    '''
    df = pd.concat([
        read_and_validate_csv(filename) for filename in get_filenames()
    ]).fillna(0)
    df = df[get_output_cols().keys()].rename(columns=get_output_cols())
    to_fac(df, 'PREM_HOL_INFO.fac')
    logging.info('Completed')

def init_logger():
    '''
    This function initialize the logging config, e.g. logginf file name and formats
    '''
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s %(levelname)-8s %(name)-20s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        filename='debug.log',
        filemode='a',
    )
    # define a Handler which writes INFO messages or higher to the sys.stderr
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    # set a format which is simpler for console use
    formatter = logging.Formatter(
        '%(asctime)s %(levelname)-8s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )
    # tell the handler to use this format
    console.setFormatter(formatter)
    # add the handler to the root logger
    logging.getLogger('').addHandler(console)

# allow this script to be run from another active folder from cmd/PowerShell
# any relative file path will then be relative to the folder containing this python script instead of your current folder in cmd/PowerShell
SCRIPT_FOLDER = os.path.dirname(__file__)
os.chdir(SCRIPT_FOLDER)
logger = logging.getLogger(__name__)
file_path = os.path.realpath(__file__)
init_logger()
logger.info('This script is running from %s', file_path)

try:
    main()
except Exception as err:
    logger.info(traceback.format_exc())
    logger.fatal(err)
