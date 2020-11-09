workbook_name = 'Sample A.xlsx'
output_workbook_name = 'Sample A - Output.csv'
worksheet_name = 'Stop Loss'
output_worksheet_name = 'Sample A - Output'

logging_level = 'DEBUG'
try:
    import csv
    import logging

    import openpyxl
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

logger.info('Loading data')
logger.debug(f'Loading workbook "{workbook_name}"')
workbook_object = openpyxl.load_workbook(filename = workbook_name)
logger.debug(f'Loaded workbook, {workbook_object}')

logger.debug(f'Selecting worksheet_name "{worksheet_name}"')
worksheet_object = workbook_object[worksheet_name]
logger.debug(f'Selected worksheet_name, "{worksheet_object}"')
logger.debug(f"A1 contents:\n {worksheet_object['A1'].value}")

logger.info('Writing data')
logger.debug(f'Opening {output_workbook_name} in write mode.')

try:
    with open(output_workbook_name, 'w', newline='') as outputfile:
        logger.debug(f'Opened {output_workbook_name} in write mode.')
        output = csv.writer(outputfile, delimiter=' ', quotechar='|', quoting=csv.QUOTE_MINIMAL)

        fieldnames = ['first_name', 'last_name']
        logger.debug('Creating DictWriter')
        writer = csv.DictWriter(outputfile, fieldnames=fieldnames)
        logger.debug('Created DictWriter')
        logger.debug('Writing data')
        writer.writeheader()
        writer.writerow({'first_name': 'Baked', 'last_name': 'Beans'})
        writer.writerow({'first_name': 'Lovely', 'last_name': 'Spam'})
        logger.debug('Wrote data')
except Exception as e:
    if isinstance(e, PermissionError):
        print("PermissionError. Please ensure that the file is not currently open.")
    else:
        raise