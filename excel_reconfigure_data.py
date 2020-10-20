# purpose of this script is to read a string value from excel cell,
# parse it
# write the parsed value back to file
# rpw = read parse write
def rpw(path, SourceWS, SourceColumn, DestinationSheet,
        DestinationColumn):

    """
    This method takes input string from excel cell
    (expected format package1.package2.objectOfInterest),
    splits the string by '.',
    write the last value (objectOfInterest) to new sheet,
    saves the xls file.

    Compatibility issue detected with file containing
    automatically generated tables.

    :param path: relative path to input file
    :param SourceWS: identification of worksheet with source data
    :param SourceColumn: identification of column with source data
    :param DestinationSheet: identification of destination sheet
    :param DestinationColumn: identification of destination column
    :return:
    """
    import openpyxl

    # load workbood
    wb = openpyxl.load_workbook(path)

    # define worksheet containing source data
    SourceWS = wb[SourceWS]

    # define target worksheet
    targetWS = wb[DestinationSheet]

    # starting row
    row = 2


    while True:

        # contains coordinates of cell (like 'A2')
        sourceCellCoordinates = SourceColumn + str(row)

        # object representing the cell
        sourceCell = SourceWS[sourceCellCoordinates]

        # value stored in excel in the cell
        SourceCellValue = sourceCell.value

        # if empty cell -> end
        if not SourceCellValue:
            break

        # parse
        # splits the cell value based on '.' (create list of strings) and saves the last item in the list
        ParsedString = SourceCellValue.split('.')[-1]

        # write
        # define coordinates of target cell
        targetCellCoordinates = DestinationColumn + str(row)

        # create target cell object
        targetCell = targetWS[targetCellCoordinates]

        # assign content to target cell
        targetCell.value = ParsedString

        # incerease the row number
        row = int(row) + 1

        # switch the new row to string
        row = str(row)

    # save changes to the file
    wb.save(path)

# call the method
rpw(r'C:\Users\XXXXXXXX\Desktop\test.xlsx',
    'inputsheet',
    'inputcolumn',
    'outputsheet',
    'outputcolumn')
