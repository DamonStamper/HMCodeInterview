try:
    import logging
    import openpyxl
except:
    raise Exception("Could not load required python libraries. Please run 'pip install -r requirements.txt' then try again.")

#Set logging options.
logger = logging.getLogger("__main__")

def saveData(data, output_filename):
    """Interface for saving data"""
    logger.debug('Calling saveData')
    saveDataAsCSV(data, output_filename)

def saveDataAsCSV(data, output_filename):
    """Implementation for saving data to CSV"""
    logger.debug('Calling saveDataAsCSV')
    data.to_csv(output_filename, index = False)
    logger.debug(f'Data saved as CSV at location "{output_filename}"')