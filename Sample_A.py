workbook_name = 'Sample A.xlsx'
output_filename = 'Sample A - Output.csv'
worksheet_name = 'Stop Loss'
output_worksheet_name = 'Sample A - Output'

logging_level = 'DEBUG'
try:
    import logging

    import pandas
except:
    raise Exception("Could not load required python libraries. Please run 'pip install -r requirements.txt' then try again.")

#Set logging options.
logging.basicConfig(format='%(asctime)s %(message)s')
logger = logging.getLogger(__name__)
logging_level = logging_level.upper()
loglevels = {
    'CRITICAL' : logging.CRITICAL,
    'ERROR' : logging.ERROR,
    'WARNING' : logging.WARNING,
    'INFO' : logging.INFO,
    'DEBUG' : logging.DEBUG
}
level = loglevels[logging_level]
logger.setLevel(level)

def main():
    logger.debug('Calling main')

    data = getWorkbook(workbook_name)
    saveData(data)











def getWorkbook(workbook_name):
    logger.debug('Calling getWorkbook')
    workbook = pandas.read_excel(workbook_name, header=5) # This sets the column labels and removes the header(first 5 rows)
    return workbook

def saveData(data): # non-OOP adapter pattern
    logger.debug('Calling saveData')
    saveDataAsCSV(data)

def saveDataAsCSV(data):
    logger.debug('Calling saveDataAsCSV')
    data.to_csv(output_filename, index = False)
    logger.debug(f'Data saved as CSV at location "{output_filename}"')

# def RemoveHeader(workbook):
#     logger.debug('Calling RemoveHeader')
#     logger.debug(f"first dataframe before removing header:\n\n {workbook.iloc[[0]]}")
#     # workbook = workbook.drop(index=workbook.index[[0, 1, 2, 3]]) # starting at index 0, remove 5 columns
#     logger.debug(f"first dataframe after removing header:\n\n {workbook.iloc[[0]]}")
#     return workbook

main()