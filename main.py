import collections
import pandas
import numpy
import itertools

def main():

    def TestingOfSheets():
        # Check if at least two columns
        assert len(currentDataframe.columns) >= 2, "The sheet %s has not two but %s columns" % (
            currentDataframe, len(currentDataframe.columns))
        # Check if value is included
        assert strValueHeader in currentDataframe.columns.values, "The sheet %s does not contain column %s" % (
            currentDataframe, strValueHeader)
        # Assert that values are in last column
        assert currentDataframe.columns.get_loc(strValueHeader) == len(
            currentDataframe.columns) - 1, "The last columns of the sheets %s are not column %s" % (
            currentDataframe, strValueHeader)
        # Check if values sum up to 1
        #assert abs(currentDataframe[strValueHeader].sum() - 1) <= float(0.01) , "The column %s of sheet %s does not sum to 1" % (
        #    currentDataframe, strValueHeader)
        # Check if at least one neu attribute is included
        newColumn = False
        for colHeader in currentDataframe.columns.values:
            if not colHeader in resultDataframe:
                newColumn = True
                # print(currentDataframe.columns.values)
        assert newColumn, "No new column in the sheet was found"
        # In jedem Blatt gibt es also MIND. eine Spalte mit neuen Charakteristika
        # TODO: Sicherstellen, dass jede Überschrift in einer Tabelle einzigartig ist
        # TODO: Sicherstellen, dass in jeder Liste nur eine neue Charakteristik dazukommt
        # TODO: Prüfen, ob in jedem Sheet jeder Eintrage einmalig ist - und nicht z.B. zweimal "EFH + Kohle" vorkommt

    def LadenDerSheets():
        # Loading Spreadsheet
        xlsxFilterFile = pandas.ExcelFile(strFilename)
        listOfSheetNames = xlsxFilterFile.sheet_names
        for sheetName in listOfSheetNames:
            dfSheet = xlsxFilterFile.parse(sheetName)
            dictOfSheets[sheetName] = dfSheet
        # TODO: Es sollte nicht notwendig sein, relative Werte in den Excel-Sheets einzugeben. Die Konvertierung kann
        # intern stattfinden. Es sollten keine relativen Werte, sondern absolute Werte ausgegeben werden

    def InsertNewFilterWithoutIntersection():
        numberOfRows = len(previousResultDataframe.index) * len(currentDataframe.index)
        newResultDataframe = pandas.DataFrame(columns=setOfNewDataframeHeaders, index=numpy.zeros(numberOfRows))
        indexNewResultDataframe = 0
        for indexPreviousResultDataframe, rowPreviousResultDataframe in previousResultDataframe.iterrows():
            valuePrevious = rowPreviousResultDataframe[strValueHeader]
            for indexCurrentDataframe, rowCurrentDataframe in currentDataframe.iterrows():
                valueCurrent = rowCurrentDataframe[strValueHeader]
                # Copy entries from previous Dataframe
                for header in setOfPreviousDataframeHeaders:
                    # newResultDataframe[header][indexNewResultDataframe] = rowPreviousResultDataframe[header]
                    icol = newResultDataframe.columns.get_loc(header)
                    newResultDataframe.set_value(index=indexNewResultDataframe, col=icol, takeable=True, value=rowPreviousResultDataframe[header])
                # Copy entries from current Dataframe
                for header in setOfNewHeaders:
                    #newResultDataframe[header][indexNewResultDataframe] = rowCurrentDataframe[header]
                    icol = newResultDataframe.columns.get_loc(header)
                    newResultDataframe.set_value(index=indexNewResultDataframe, col=icol, takeable=True, value=rowCurrentDataframe[header])
                # Copy new value
                newValue = valuePrevious * valueCurrent
                # newResultDataframe[strValueHeader][indexNewResultDataframe] = newValue
                icol = newResultDataframe.columns.get_loc(strValueHeader)
                newResultDataframe.set_value(index=indexNewResultDataframe, col=icol, takeable=True, value=newValue)
                indexNewResultDataframe += 1
        return newResultDataframe

    def InsertNewFilterWithIntersection():
        listOfNewAttributes = list()
        for header in setOfNewHeaders:
            listOfNewAttributes.append(list(set(currentDataframe[header])))
        # Transposiong the list
        listOfNewAttributes = list(map(list, zip(*listOfNewAttributes)))
        newAttributesDataframe = pandas.DataFrame(listOfNewAttributes, columns=setOfNewHeaders)

        numberOfVariationsOfIntersectingColumns = 1
        for header in setOfIntersectingHeaders:
            setOfItems = set(currentDataframe[header])
            numberOfVariationsOfIntersectingColumns += len(setOfItems)

        lengthOfCurrentDataframe = len(currentDataframe.index)
        numberOfNewVariants = len(listOfNewAttributes)
        numberOfRowsWhichAreAddedPerEntry = numberOfNewVariants # lengthOfCurrentDataframe / numberOfVariationsOfIntersectingColumns
        numberOfRows = len(previousResultDataframe.index) * numberOfRowsWhichAreAddedPerEntry
        newResultDataframe = pandas.DataFrame(columns=setOfNewDataframeHeaders, index=numpy.zeros(int(numberOfRows)))

        adjustingValuesRequired = CheckIfValuesNeedToBeAdjusted()
        if adjustingValuesRequired:
            pass #AdjustValuesOfCurrentDataframe()

        indexNewResultDataframe = 0
        index = 0
        for indexPreviousResultDataframe, rowPreviousResultDataframe in previousResultDataframe.iterrows():
            index += 1
            valuePrevious = rowPreviousResultDataframe[strValueHeader]
            for indexOfNewAttributeDataframe, rowOfNewAttributeDataframe in newAttributesDataframe.iterrows():
                indexNewResultDataframe = (index-1) * numberOfRowsWhichAreAddedPerEntry + indexOfNewAttributeDataframe
                # Copy entries from previous Dataframe
                for header in setOfPreviousDataframeHeaders:
                    ivalue = rowPreviousResultDataframe[header]
                    icol = newResultDataframe.columns.get_loc(header)
                    newResultDataframe.set_value(index=indexNewResultDataframe, col = icol, takeable=True, value = ivalue)
                # Inserting Variations
                for header in newAttributesDataframe:
                    ivalue = newAttributesDataframe[header][indexOfNewAttributeDataframe]
                    icol = newResultDataframe.columns.get_loc(header)
                    newResultDataframe.set_value(index=indexNewResultDataframe, col = icol, takeable=True, value = ivalue)
                # Now the correct values have to be readout
                sumValueCurrent = 0
                numberOfColumns = len(setOfIntersectingHeaders)
                if numberOfColumns == 1:
                    headerSearchedFor_1 = setOfIntersectingHeaders[0]
                    valueSearchedFor_1 =rowPreviousResultDataframe[headerSearchedFor_1]
                    sumValueCurrent = currentDataframe.loc[currentDataframe[headerSearchedFor_1] == valueSearchedFor_1, strValueHeader].sum()
                    newHeader = next(iter(setOfNewHeaders))
                    newAttribute = rowOfNewAttributeDataframe[newHeader]
                    valueCurrent = currentDataframe.loc[(currentDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                                                        (currentDataframe[newHeader] == newAttribute),
                                                        strValueHeader].sum()
                elif numberOfColumns == 2:
                    headerSearchedFor_1 = setOfIntersectingHeaders[0]
                    valueSearchedFor_1 = rowPreviousResultDataframe[headerSearchedFor_1]
                    headerSearchedFor_2 = setOfIntersectingHeaders[1]
                    valueSearchedFor_2 = rowPreviousResultDataframe[headerSearchedFor_2]
                    sumValueCurrent = currentDataframe.loc[(currentDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                                             (currentDataframe[headerSearchedFor_2] == valueSearchedFor_2),
                                             strValueHeader].sum()
                    newHeader = next(iter(setOfNewHeaders))
                    newAttribute = rowOfNewAttributeDataframe[newHeader]
                    valueCurrent = currentDataframe.loc[(currentDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                                                (currentDataframe[headerSearchedFor_2] == valueSearchedFor_2) &
                                                (currentDataframe[newHeader] == newAttribute),
                                             strValueHeader].sum()
                #elif numberOfColumns == 3:
                #    sumvalue = currentDataframe.loc[currentDataframe[0] == row[0], currentDataframe[1] == row[1], currentDataframe[3] == row[3], strValueHeader].sum()
                else:
                    valueCurrent = 0

                if valueCurrent == 0:
                    #print(valueCurrent)
                    pass
                if sumValueCurrent == 0:
                    print(sumValueCurrent)
                finalValueCurrent = valueCurrent/sumValueCurrent
                newValue = valuePrevious * finalValueCurrent
                #newResultDataframe[strValueHeader][indexNewResultDataframe] = newValue
                icol = newResultDataframe.columns.get_loc(strValueHeader)
                newResultDataframe.set_value(index=indexNewResultDataframe, col=icol, takeable=True, value=newValue)
                indexNewResultDataframe += 1
        return newResultDataframe

    def CheckIfValuesNeedToBeAdjusted():
        valuesNeedToBeAdjusted = False
        #If there is more than one attribute which defines distribution, the values in current sheet have to be matched
        numberOfIntersectingAttributes = len(setOfIntersectingHeaders)
        if numberOfIntersectingAttributes == 1:
            valuesNeedToBeAdjusted = False
            return valuesNeedToBeAdjusted
        elif numberOfIntersectingAttributes == 2:
            for indexOfPreviousDataframe, rowOfPreviousDataframe in previousResultDataframe.iterrows():
                headerSearchedFor_1 = setOfIntersectingHeaders[0]
                valueSearchedFor_1 = rowOfPreviousDataframe[headerSearchedFor_1]
                headerSearchedFor_2 = setOfIntersectingHeaders[1]
                valueSearchedFor_2 = rowOfPreviousDataframe[headerSearchedFor_2]
                targetValueInPreviousDataframe = previousResultDataframe.loc[(previousResultDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                                                (previousResultDataframe[headerSearchedFor_2] == valueSearchedFor_2), strValueHeader].sum()
                actualValueInCurrentDataframe = currentDataframe.loc[(currentDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                                                (currentDataframe[headerSearchedFor_2] == valueSearchedFor_2), strValueHeader].sum()
                if abs(targetValueInPreviousDataframe - actualValueInCurrentDataframe) > 0.0000001:
                    valuesNeedToBeAdjusted = True
                    return valuesNeedToBeAdjusted
        elif numberOfIntersectingAttributes == 3:
            for indexOfCurrentDataframe, rowOfCurrentDataframe in currentDataframe.iterrows():
                targetValueInPreviousDataframe = rowOfCurrentDataframe[strValueHeader]
                headerSearchedFor_1 = setOfIntersectingHeaders[0]
                valueSearchedFor_1 = rowOfCurrentDataframe[headerSearchedFor_1]
                headerSearchedFor_2 = setOfIntersectingHeaders[1]
                valueSearchedFor_2 = rowOfCurrentDataframe[headerSearchedFor_2]
                headerSearchedFor_3 = setOfIntersectingHeaders[2]
                valueSearchedFor_3 = rowOfCurrentDataframe[headerSearchedFor_3]
                actualValueInCurrentDataframe = previousResultDataframe.loc[(previousResultDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                                                                     (previousResultDataframe[headerSearchedFor_2] == valueSearchedFor_2) &
                                                                     (previousResultDataframe[headerSearchedFor_3] == valueSearchedFor_3),
                                                                     strValueHeader].sum()
                actualValueInCurrentDataframe = currentDataframe.loc[(currentDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                                                                     (currentDataframe[headerSearchedFor_2] == valueSearchedFor_2) &
                                                                     (currentDataframe[headerSearchedFor_3] == valueSearchedFor_3),
                                                                     strValueHeader].sum()
                if abs(targetValueInPreviousDataframe - actualValueInCurrentDataframe) > 0.0000001:
                    valuesNeedToBeAdjusted = True
                    return valuesNeedToBeAdjusted
        elif numberOfIntersectingAttributes > 3:
            raise ValueError('There are more than 3 columns intersecting, this is not yet implemented')
        return valuesNeedToBeAdjusted

    def AdjustValuesOfCurrentDataframe():
        valuesNeedToBeAdjusted = False
        # If there is more than one attribute which defines distribution, the values in current sheet have to be matched
        numberOfIntersectingAttributes = len(setOfIntersectingHeaders)
        if numberOfIntersectingAttributes == 1:
            raise ValueError('There should be no adjustment required...')
        elif numberOfIntersectingAttributes == 2:
            for indexOfPreviousDataframe, rowOfPreviousDataframe in previousResultDataframe.iterrows():
                headerSearchedFor_1 = setOfIntersectingHeaders[0]
                valueSearchedFor_1 = rowOfPreviousDataframe[headerSearchedFor_1]
                headerSearchedFor_2 = setOfIntersectingHeaders[1]
                valueSearchedFor_2 = rowOfPreviousDataframe[headerSearchedFor_2]
                targetValueInPreviousDataframe = previousResultDataframe.loc[
                    (previousResultDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                    (previousResultDataframe[headerSearchedFor_2] == valueSearchedFor_2), strValueHeader].sum()
                actualValueInCurrentDataframe = currentDataframe.loc[
                    (currentDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                    (currentDataframe[headerSearchedFor_2] == valueSearchedFor_2), strValueHeader].sum()
                correctionFactor = targetValueInPreviousDataframe / actualValueInCurrentDataframe
                for indexOfCurrentDataframe, rowOfCurrentDataframe in currentDataframe.iterrows():
                    if ((rowOfCurrentDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                        (rowOfCurrentDataframe[headerSearchedFor_2] == valueSearchedFor_2)):
                        rowOfCurrentDataframe[strValueHeader] = rowOfCurrentDataframe[strValueHeader] * correctionFactor
        elif numberOfIntersectingAttributes == 3:
            for indexOfCurrentDataframe, rowOfCurrentDataframe in currentDataframe.iterrows():
                targetValueInPreviousDataframe = rowOfCurrentDataframe[strValueHeader]
                headerSearchedFor_1 = setOfIntersectingHeaders[0]
                valueSearchedFor_1 = rowOfCurrentDataframe[headerSearchedFor_1]
                headerSearchedFor_2 = setOfIntersectingHeaders[1]
                valueSearchedFor_2 = rowOfCurrentDataframe[headerSearchedFor_2]
                headerSearchedFor_3 = setOfIntersectingHeaders[2]
                valueSearchedFor_3 = rowOfCurrentDataframe[headerSearchedFor_3]
                actualValueInCurrentDataframe = previousResultDataframe.loc[
                    (previousResultDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                    (previousResultDataframe[headerSearchedFor_2] == valueSearchedFor_2) &
                    (previousResultDataframe[headerSearchedFor_3] == valueSearchedFor_3),
                    strValueHeader].sum()
                actualValueInCurrentDataframe = currentDataframe.loc[
                    (currentDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                    (currentDataframe[headerSearchedFor_2] == valueSearchedFor_2) &
                    (currentDataframe[headerSearchedFor_3] == valueSearchedFor_3),
                    strValueHeader].sum()
                correctionFactor = targetValueInPreviousDataframe / actualValueInCurrentDataframe
                for indexOfCurrentDataframe, rowOfCurrentDataframe in currentDataframe.iterrows():
                    if ((currentDataframe[headerSearchedFor_1] == valueSearchedFor_1) &
                        (currentDataframe[headerSearchedFor_2] == valueSearchedFor_2) &
                        (currentDataframe[headerSearchedFor_3] == valueSearchedFor_3)):
                        rowOfCurrentDataframe[strValueHeader] = rowOfCurrentDataframe[strValueHeader] * correctionFactor
        elif numberOfIntersectingAttributes > 3:
            raise ValueError('There are more than 3 columns intersecting, this is not yet implemented')

    # ==================================================================================================================
    # Definition of files
    # ==================================================================================================================
    strFilename = 'StatistischeVerteilungC.xlsx'  # 'filter1.xlsx'
    strFilenameCSV = 'output.csv'
    global strValueHeader
    strValueHeader = 'value'
    # Collection/Set of spreadsheets
    dictOfSheets = dict()
    LadenDerSheets()
    # Result Dataframe
    resultDataframe = {}
    for sheets in dictOfSheets:
        currentDataframe = dictOfSheets[sheets]
        TestingOfSheets()

    # ==================================================================================================================
    # Working with sheets
    # ==================================================================================================================
    for sheets in dictOfSheets:
        currentDataframe = dictOfSheets[sheets]
        if not len(resultDataframe):
            # resultDataframe is empty
            resultDataframe = currentDataframe
        else:
            previousResultDataframe = resultDataframe
            # Determining the dimensions
            setOfIntersectingHeaders = set(previousResultDataframe.columns).intersection(set(currentDataframe.columns))
            setOfIntersectingHeaders = list(filter(strValueHeader.__ne__, setOfIntersectingHeaders))
            setOfPreviousDataframeHeaders = list(filter(strValueHeader.__ne__, previousResultDataframe.columns))
            setOfNewHeaders = set(currentDataframe.columns) - set(previousResultDataframe.columns)
            setOfNewDataframeHeaders = set(filter(strValueHeader.__ne__, set(previousResultDataframe.columns))).union(setOfNewHeaders)
            setOfNewDataframeHeaders.add(strValueHeader)
            dictOfAttributesOfIntersectingHeaders = {}
            for header in setOfIntersectingHeaders:
                dictOfAttributesOfIntersectingHeaders[header]=set(currentDataframe[header])

            if len(setOfIntersectingHeaders) == 0:
                newResultDataframe = InsertNewFilterWithoutIntersection()
            else:
                newResultDataframe = InsertNewFilterWithIntersection()

            resultDataframe = newResultDataframe
    resultDataframe.to_csv(strFilenameCSV) #, header=1)
    print(resultDataframe)
    print("The sum of all relative values is {sum}".format(sum = resultDataframe[strValueHeader].sum()))

if __name__ == "__main__":
    main()