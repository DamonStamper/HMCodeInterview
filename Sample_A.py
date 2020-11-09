input_filename = 'Sample A.xlsx'
output_filename = 'Sample A - Output.csv'

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

    data = getData(input_filename)
    data = cleanData(data)
    saveData(data)











def getData(input_filename): # non-OOP adapter pattern
    logger.debug('Calling getData')
    data = getDataFromExcel(input_filename)
    return data

def getDataFromExcel(input_filename):
    logger.debug('Calling getDataFromExcel')
    workbook = pandas.read_excel(input_filename, header=5) # This sets the column labels and removes the header(first 5 rows)
    logger.debug(f'Column labels:\n {workbook.columns.tolist()}')
    return workbook

def cleanData(data):
    logger.debug('Calling cleanData')
    data = removeBlankRows(data)
    data = FillInMissingData(data)
    return data

def removeBlankRows(data):
    logger.debug('Calling removeBlankRows')
    data = data.dropna(subset = ["Claim Number"])
    return data

def FillInMissingData(data):
    logger.debug('Calling FillInMissingData')
    data = data.ffill() # Using ffill to propagate data from "top" rows to "below" rows because input data was designed for human readers in that they would assume missing data on "child" rows could be found on "parent" rows.
    return data

def saveData(data):
    logger.debug('Calling saveData')
    data = formatDataForSaving(data) # Doing this at save time since there is some data loss (rounding numbers) and may cause unexpected results otherwise.
    saveDataAsCSV(data)

def formatDataForSaving(data):
    # data = data.style.format({'label': '${0:,.2f}'})
    # data['Allowance'] = data['Allowance'].replace( '[\$,)]','', regex=True ).replace( '[(]','-',   regex=True )
    # data['Allowance'] = data['Allowance'].apply("${:.2f}")
    format_mapping={
        'Allowance': '${:,.2f}',
        'Paid\nAmount': '${:,.2f}'
        }
    for key, value in format_mapping.items():
        data[key] = data[key].apply(value.format)
    return data

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