input_filename = 'Sample B.xlsx'
output_filename = 'Sample B - Output.csv'

columnInfo = {
    'name': 'A4', 
    'IDNumber': 'A5', 
    'contractPeriod': 'I3',
    'contractBasis': 'I4',
    'Paid Claims Date': 'I5',
    'type': 'I6'
    }

logging_level = 'DEBUG'
try:
    import logging
    import platform

    import openpyxl
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
    # data = cleanData(data)
    saveData(data)

def getData(input_filename): # non-OOP adapter pattern
    logger.debug('Calling getData')
    data = getDataFromExcel(input_filename)
    return data

def getDataFromExcel(input_filename):
    logger.debug('Calling getDataFromExcel')
    EnrollmentInformation = pandas.read_excel(input_filename, sheet_name = 'Enrollment Information', header=0, dtype=object)
    Claims = pandas.read_excel(input_filename, sheet_name = 'Claims 02-13-20', header=9, dtype=str, converters= {'DTE_DISP':pandas.to_datetime, 'DTE_SRVC_BEG':pandas.to_datetime, 'DTE_SRVC_END':pandas.to_datetime, 'PAT_ID': int, 'YTD Total Amount ':int, 'Reimbursement \nAmt. Requested':int})
    logger.info('Please note that expected output calls for "YTD Total Amount" and "Reimbursement \nAmt. Requested" fields to be a integers (which means no decimal values). However this is currency which means that we are not dealing in whole units. I am going to make a judgement call and allow these fields to be floats/have decimal values.')
    additonalColumnsDataframe = claimsExtraInfo()
    dataframe = mergeDataframes(EnrollmentInformation, additonalColumnsDataframe)
    dataframe = mergeDataframes(dataframe, Claims)
    dataframe = fillDataframeDesiredData(dataframe)
    dataframe = removePaddedZeros(dataframe, ["CODE 1"])
    dataframe = fillDataframeFromTo(EnrollmentInformation, dataframe)


    dataframe = setDateTimeColumns(dataframe)
    # Arrange columns in specific order
    cols = (EnrollmentInformation.columns.tolist() + additonalColumnsDataframe.columns.tolist() + Claims.columns.tolist())
    dataframe = dataframe[cols]
    return dataframe

def removePaddedZeros(dataframe, columns):
    logger.debug('Calling removePaddedZeros')
    dataframe[columns] = dataframe[columns].apply(pandas.to_numeric)
    return dataframe

def fillDataframeDesiredData(dataframe):
    # for key in columnInfo:
    #     dataframe.ffill
    # dataframe = dataframe.ffill(list(columnInfo.keys()))
    cols = list(columnInfo.keys())
    logger.debug(f'cols to ffill():\n{cols}')
    logger.debug(f'dataframe before fillDataframeDesiredData:\n{dataframe}')
    dataframe.loc[:,cols] = dataframe.loc[:,cols].ffill()
    logger.debug(f'dataframe after fillDataframeDesiredData:\n{dataframe}')
    # dataframe = dataframe.ffill(list(columnInfo.keys()))
    return dataframe

def fillDataframeFromTo(fromDataframe, toDataframe):
    # cols = ['X', 'Y']
    cols = fromDataframe.columns.tolist()
    toDataframe.loc[:,cols] = toDataframe.loc[:,cols].ffill()
    return toDataframe

def dateFix(input):
    """Format the input variable to a non 0-padded 'month/day/year hour:minute' format"""
    # TODO: Is running platform.system() performant given that this function is called repeatedly via map? system() means this is a method, but is it just an accessor method or something that does logic that doesn't need repeated ROWS*3 times?
    iterant = pandas.to_datetime(input, errors='ignore')
    try:
        # iterant = iterant.strftime('%m/%d/%Y %H:%M')
        if platform.system() != 'Windows':
            iterant = iterant.strftime('%-m/%-d/%Y %-H:%M')
        else:
            iterant = iterant.strftime('%#m/%#d/%Y %#H:%M')
    except Exception:
        pass
    return iterant

def setDateTimeColumns(data):
    datetimeColumns = ('DTE_DISP','DTE_SRVC_BEG','DTE_SRVC_END')
    for column in datetimeColumns:
        logger.debug(f'Converting datetime in {column}')
        data[column] = list(map(dateFix, data[column]))
        # data[column]
    return data

def claimsExtraInfo():
    wb = openpyxl.load_workbook(input_filename)
    for s in range(len(wb.sheetnames)):
        # logger.debug(f'Looking at sheet {s.title}')
        if 'Claims' in  wb.sheetnames[s]:
            wb.active = s
            ws = wb.active
            logger.debug(ws.title)

    # Kindof a reverse interpolation? Replace the values in the dict with the appropriate, cleaned, data in the spreadsheet thus creating a key value pair representing column name and column value.
    for columnName, location in columnInfo.items():
        columnInfo[columnName] = (ws[location].value).split(':')[1].strip()

    dataframe = pandas.DataFrame([columnInfo], columns=columnInfo.keys())
    # logger.debug(f'\n{dataframe}')
    return dataframe

def mergeDataframes(dataframe1, dataframe2):
    #  clumn_addendum_value = "\n".join(column_addendum_values)
    # mergedDataframe = pandas.merge(dataframe1, dataframe2)
    # mergedDataframe = dataframe1.join(dataframe2[dataframe2.columns])
    mergedDataframe = dataframe2.join(dataframe1[dataframe1.columns])
    # mergedDataframe = mergedDataframe.ffill()
    # logger.debug(mergedDataframe)
    return mergedDataframe

def saveData(data):
    logger.debug('Calling saveData')
    saveDataAsCSV(data)

def saveDataAsCSV(data):
    logger.debug('Calling saveDataAsCSV')
    data.to_csv(output_filename, index = False)
    logger.debug(f'Data saved as CSV at location "{output_filename}"')

main()