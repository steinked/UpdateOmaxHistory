Imports System.IO, System.Text.RegularExpressions

Module Module1

    Dim wsMaterial, wsMaterialCategory, wsMachineability, wsThickness, wsPierces,
        wsCutModel, wsEtchSpeed, wsEstTime, wsEstCost, wsEstAbrasive, wsTPLen,
        wsCutLen, wsTTSCutting, wsTTSTraversing, wsTTSRelayCycles, wsTTSEtching As Integer

    Sub Main()
        'Start parsing log files
        'The history files are stored in the path selected
        'The Excel file to put the data in is also selected by the user

        Dim curFileInfo As IO.FileInfo
        Dim files As FileSystemInfo()
        Dim foundFiles As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
        Dim verFieldIndex, textFieldIndex, dateTimeField, wsMachName, wsFolderName, wsFN, wsDate,
            wsLoadTime, wsCStart, wsCStop, wsCFinish, wsFinishCol, wsTimeCutting,
            wsLoadToCut, wsDurationLoaded, numFiles, firstRow As Integer
        Dim LogFilePath, ExcelFileName As String
        Dim ver18, pathStarted, partFnInd, cutStartInd, pathFinInd, cutPauseInd, cutStopInd, dryRunInd As String
        Dim beginParse, foundFileName, prevString, cutStarted, cutStopped, pathFinished As String
        Dim objExcelApp As Object
        Dim wb, ws As Object
        Dim newFileLoop As Boolean
        'wsCurRow is the current row of the data store worksheet
        'It is incremented each time we start to parse another cut (aka path)
        Dim wsCurRow, afterCutStopLineCount As Integer
        'lastDate used when we are appending data to the spreadsheet.  Is last file date
        'in the spreadsheet.
        Dim lastDate As Date

        'Declare variables use in calculation.  First one is timeCutting - 
        'Time the machine is spent cutting between one start-stop cycle
        'cStartTime is to hold value to use later to calculate cutting time.
        Dim cStartTime, savedLoadTime As Date

        objExcelApp = CreateObject("Excel.Application")
        'wb = objExcelApp.Workbooks(SpreadsheetNameBox.Text)

        numFiles = 0
        verFieldIndex = 1
        textFieldIndex = 2
        dateTimeField = 0
        LogFilePath = "C:\Users\Public\OMAX_Corporation\AllUserData\History\history subset\subset"
        ExcelFileName = "C:\Users\Dan\Documents\autoHistory.xlsx"
        ver18 = "ver 18.0"
        pathStarted = "Path started with the following setup:"
        partFnInd = "Part File Name: "
        dryRunInd = "Dry Run:"
        cutStartInd = "inches from Abs Home: "
        pathFinInd = "Path Finished."
        cutStopInd = "Path stopped"
        cutPauseInd = "Path paused"
        beginParse = "beginParse"
        foundFileName = "foundFileName"
        cutStarted = "cutStarted"
        cutStopped = "cutStopped"
        pathFinished = "pathFinished"
        wsMachName = 1
        wsFolderName = 2
        wsFN = 3
        wsDate = 4
        wsLoadTime = 5
        wsCStart = 6
        wsCStop = 7
        wsCFinish = 8
        wsFinishCol = 9
        wsTimeCutting = 10
        wsLoadToCut = 11
        'not actually used - calculated in spreadsheet.
        wsDurationLoaded = 12

        'initialize field indexes (declared in Public Class above)
        initFields()

        wb = objExcelApp.Workbooks.Open(ExcelFileName, ReadOnly:=False)
        wb.activate
        ws = wb.activesheet
        ' add line to make ws the active sheet
        objExcelApp.Visible = True


        'Start adding data on the first empty ROW
        wsCurRow = ws.Range("A1").CurrentRegion.Rows.Count
            firstRow = wsCurRow + 1
            lastDate = CDate(ws.Cells(wsCurRow, wsDate).value)

        'THIS BEGINS THE MAJOR LOOP - going through each history file
        'First step is to get a list of all history files in the history file directory
        'and use the Order By to order them by the last time they were written to.

        'WAS files = New DirectoryInfo(LogFilePath.Text).GetFileSystemInfos("?=-Omax-1*")
        'changed because some omax2 files starting 10/23/2015 are H=-Aalto and H=-Czar, etc


        'Need to get files since the last date in the Excel spreadhseet
        files = New DirectoryInfo(LogFilePath).GetFileSystemInfos("?=-*")
        Dim theFiles = From file In files Order By file.LastWriteTime
                       Where file.LastWriteTime.Date > lastDate.Date Select file.FullName

        'WAS files = New DirectoryInfo(LogFilePath.Text).GetFileSystemInfos("?=-Omax-1*")
        'changed because some omax2 files starting 10/23/2015 are H=-Aalto and H=-Czar, etc
        foundFiles = My.Computer.FileSystem.GetFiles(
            LogFilePath, FileIO.SearchOption.SearchTopLevelOnly, "?=-*")

        For Each foundFile As String In theFiles
            curFileInfo = My.Computer.FileSystem.GetFileInfo(foundFile)

            newFileLoop = True
            numFiles = numFiles + 1

            'The nex big chunk of code parses through each individual file
            Using MyReader As New FileIO.TextFieldParser(foundFile)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters("|")

                Dim currentRow As String()
                Dim state As String : state = beginParse
                Dim myDateTime As String
                Dim myDateValue As Date
                Dim newtime As New Date
                Dim timeDiff As TimeSpan

                While Not MyReader.EndOfData
                    Dim currentField As String

                    Try
                        currentRow = MyReader.ReadFields()
                        currentField = currentRow(textFieldIndex)
                    Catch ex As FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try

                    'Use ParseDateTime to remove the colon so date.Parse can read date/time
                    myDateTime = ParseDateTime(currentRow(dateTimeField))
                    myDateValue = Date.Parse(myDateTime)

                    If (myDateValue.Date > lastDate.Date) Then

                        'This is the state machine.  Search for different
                        'string depending on the state we are in
                        Select Case state
                            Case beginParse
                                'Reset to 0.  This is used in the rare case we have a 
                                'start->manual stop and another start-> manual stop in a 
                                'row without loading a new file.  The last cut on file 
                                '5005A01.ORD is an example of this.
                                afterCutStopLineCount = 0

                                If MyRegExMatcher(partFnInd, currentField) Then

                                    state = foundFileName

                                    wsCurRow = wsCurRow + 1

                                    ws.cells(wsCurRow, wsFN).value =
                                          Strings.Right(currentField, currentField.Length -
                                                           currentField.LastIndexOf(":") - 1)
                                    ws.cells(wsCurRow, wsFolderName).value = foundFile

                                    ws.cells(wsCurRow, wsLoadTime).value = myDateValue
                                    ws.cells(wsCurRow, wsDate).value = myDateValue.Date

                                    savedLoadTime = myDateValue
                                End If
                            Case foundFileName
                                'We are capturing a LOT of data in this section.
                                'First things like material, machineability, etc.
                                'Finally when we find the cutStartInd we change state
                                Dim strToConvert As String
                                Dim myDec As Decimal

                                If MyRegExMatcher("Material:", currentField) Then
                                    ws.cells(wsCurRow, wsMaterial).value = getFieldInfo(wsMaterial, currentField) 'parse value of this field
                                ElseIf MyRegExMatcher("MaterialCategory:", currentField) Then
                                    ws.cells(wsCurRow, wsMaterialCategory).value = getFieldInfo(wsMaterialCategory, currentField) 'parse value of this field
                                    'The space after the : below is important
                                    'On some OMAX-2 files they have <Enter Custom Ceramic/Carbide Machineability:>
                                    'and we DON'T want to try to capture in that scenario
                                ElseIf MyRegExMatcher("Machineability: ", currentField) Then
                                    'If Not MyRegExMatcher("Machineability:>", currentField) Then
                                    'ws.cells(wsCurRow, wsMachineability).value = CDec(getFieldInfo(wsMachineability, currentField)) 'parse value of this field
                                    'End If
                                ElseIf MyRegExMatcher("TiltKLSLockThickness:", currentField) Then
                                    'do nothing - just need to go to next line
                                ElseIf MyRegExMatcher("Thickness:", currentField) Then
                                    ws.cells(wsCurRow, wsThickness).value = CDec(getFieldInfo(wsThickness, currentField)) 'parse value of this field
                                ElseIf MyRegExMatcher("Pierces:", currentField) Then
                                    ws.cells(wsCurRow, wsPierces).value = CInt(getFieldInfo(wsPierces, currentField)) 'parse value of this field
                                ElseIf MyRegExMatcher("Cutting model used:", currentField)
                                    ws.cells(wsCurRow, wsCutModel).value = CInt(getFieldInfo(wsCutModel, currentField)) 'parse value of this field
                                ElseIf MyRegExMatcher("EtchSpeed:", currentField)
                                    ws.cells(wsCurRow, wsEtchSpeed).value = CInt(getFieldInfo(wsEtchSpeed, currentField)) 'parse value of this field
                                ElseIf MyRegExMatcher("Estimated time", currentField)
                                    strToConvert = getFieldInfo(wsEstTime, currentField)
                                    myDec = CDec(strToConvert)
                                    ws.cells(wsCurRow, wsEstTime).value = Date.FromOADate(myDec / 1440.0) 'parse value of this field
                                ElseIf MyRegExMatcher("Estimated cost", currentField)
                                    ws.cells(wsCurRow, wsEstCost).value = getFieldInfo(wsEstCost, currentField) 'parse value of this field
                                ElseIf MyRegExMatcher("Estimated abrasive", currentField)
                                    ws.cells(wsCurRow, wsEstAbrasive).value = getFieldInfo(wsEstAbrasive, currentField) 'parse value of this field
                                ElseIf MyRegExMatcher("Length of tool", currentField)
                                    ws.cells(wsCurRow, wsTPLen).value = getFieldInfo(wsTPLen, currentField) 'parse value of this field
                                ElseIf MyRegExMatcher("Length of cutting", currentField)
                                    ws.cells(wsCurRow, wsCutLen).value = getFieldInfo(wsCutLen, currentField) 'parse value of this field
                                ElseIf MyRegExMatcher("spent cutting", currentField)
                                    strToConvert = getFieldInfo(wsTTSCutting, currentField)
                                    myDec = CDec(strToConvert)
                                    ws.cells(wsCurRow, wsTTSCutting).value = Date.FromOADate(myDec / 1440.0)
                                ElseIf MyRegExMatcher("spent etching", currentField)
                                    strToConvert = getFieldInfo(wsTTSEtching, currentField)
                                    myDec = CDec(strToConvert)
                                    ws.cells(wsCurRow, wsTTSEtching).value = Date.FromOADate(myDec / 1440.0)
                                ElseIf MyRegExMatcher("spent travers", currentField)
                                    strToConvert = getFieldInfo(wsTTSTraversing, currentField)
                                    myDec = CDec(strToConvert)
                                    ws.cells(wsCurRow, wsTTSTraversing).value = Date.FromOADate(myDec / 1440.0)
                                ElseIf MyRegExMatcher("cycling", currentField)
                                    strToConvert = getFieldInfo(wsTTSRelayCycles, currentField)
                                    myDec = CDec(strToConvert)
                                    ws.cells(wsCurRow, wsTTSRelayCycles).value = Date.FromOADate(myDec / 1440.0)
                                End If

                                If MyRegExMatcher(dryRunInd, currentField) Then
                                    'This is a dry run - ignore it
                                    'Clear all fields and begin parse all over
                                    ws.range(ws.cells(wsCurRow, wsFolderName), ws.cells(wsCurRow, wsTTSRelayCycles)).ClearContents
                                    state = beginParse

                                    'reduce the row count because the beginParse case will 
                                    'increment it
                                    wsCurRow = wsCurRow - 1
                                ElseIf MyRegExMatcher(cutStartInd, currentField) Then
                                    state = cutStarted

                                    myDateTime = ParseDateTime(currentRow(dateTimeField))
                                    myDateValue = Date.Parse(myDateTime)
                                    ws.cells(wsCurRow, wsCStart).value = myDateValue

                                    'Save this to use in calculating different from start to stop (cutting time)
                                    cStartTime = myDateValue

                                End If
                            Case cutStarted

                                afterCutStopLineCount = 0

                                'Initial IF is to handle 6:24:35PM -> 5:39:06 transition in file
                                'H=-Omax-1-SD Design-Inner Field Tile-5217IF06 dated 6/19/14
                                'Has and 'inches from abd home: and then 20 days later loaded.
                                'Never had one of the typical finishes
                                If MyRegExMatcher(pathStarted, currentField) Then
                                    'This is an error.  Dump filename, etc.  Clear current row and start again.

                                    'reset so we can look for the next row
                                    state = beginParse

                                ElseIf MyRegExMatcher(cutStopInd, currentField) Or
                                    MyRegExMatcher(cutPauseInd, currentField) Then

                                    'Use ParseDateTime to remove the colon so date.Parse can read date/time
                                    myDateTime = ParseDateTime(currentRow(dateTimeField))
                                    myDateValue = Date.Parse(myDateTime)

                                    state = cutStopped

                                    ws.cells(wsCurRow, wsCStop).value = myDateValue

                                    'Calculate time from cut start to cut stop
                                    newtime = New Date
                                    'timeDiff = myDateValue - cStartTime
                                    'if added because file Oma-1-SD Design-Inner Field Tile-5217IF06 crashes here
                                    timeDiff = myDateValue - Date.FromOADate(ws.cells(wsCurRow, wsCStart).value)
                                    If (timeDiff.Days > 1) Then

                                        'reset so we can look for the next row
                                        state = beginParse
                                        wsCurRow = wsCurRow + 1

                                    Else

                                        newtime = newtime.Add(timeDiff)
                                        ws.cells(wsCurRow, wsTimeCutting).value = newtime

                                        'Calculate time from when this file was opened to cut finished
                                        newtime = New Date
                                        timeDiff = myDateValue - Date.FromOADate(ws.cells(wsCurRow, wsLoadTime).value)
                                        newtime = newtime.Add(timeDiff)
                                        ws.cells(wsCurRow, wsLoadToCut).value = newtime
                                    End If


                                ElseIf MyRegExMatcher(pathFinInd, currentField) Then
                                    'Use ParseDateTime to remove the colon so date.Parse can read date/time
                                    myDateTime = ParseDateTime(currentRow(dateTimeField))
                                    myDateValue = Date.Parse(myDateTime)

                                    state = pathFinished

                                    ws.cells(wsCurRow, wsCStop).value = myDateValue

                                    'Calculate time from cut start to cut stop
                                    newtime = New Date
                                    'timeDiff = myDateValue - cStartTime
                                    timeDiff = myDateValue - Date.FromOADate(ws.cells(wsCurRow, wsCStart).value)
                                    'timeDiff = TimeSerial(myDateValue.Hour, myDateValue.Minute, myDateValue.Second) - TimeSerial(cStartTime.Hour, cStartTime.Minute, cStartTime.Second)
                                    newtime = newtime.Add(timeDiff)
                                    ws.cells(wsCurRow, wsTimeCutting).value = newtime

                                    'Calculate time from when this file was opened to cut finished
                                    newtime = New Date
                                    timeDiff = myDateValue - Date.FromOADate(ws.cells(wsCurRow, wsLoadTime).value)
                                    newtime = newtime.Add(timeDiff)
                                    ws.cells(wsCurRow, wsLoadToCut).value = newtime

                                    'Also capturing Finish time - probably redundant, but just in case
                                    'Could be useful in determining how often a part goes from start to
                                    'finish vs. non-finished cuts - usually indicating problem
                                    'Use ParseDateTime to remove the colon so date.Parse can read date/time
                                    ws.cells(wsCurRow, wsCFinish).value = myDateValue
                                    ws.cells(wsCurRow, wsFinishCol).value = "Finish"
                                End If

                            Case cutStopped
                                'This is very similar to pathFinished below.  We found either cutStopInd
                                'or cutPauseInd.  If we find "Inches from Abs Home" in the next three
                                'lines, handle it.  Otherwise state becomes beginParse again
                                If afterCutStopLineCount < 4 Then
                                    'if we find the cutStartInd (inches from Abs Home) create new row, copy
                                    'from previous and look for the next cut stop indicator
                                    If MyRegExMatcher(cutStartInd, currentField) Then
                                        state = cutStarted

                                        'Use ParseDateTime to remove the colon so date.Parse can read date/time
                                        myDateTime = ParseDateTime(currentRow(dateTimeField))
                                        myDateValue = Date.Parse(myDateTime)

                                        'Capture data from this row to new row. 
                                        ws.range(ws.cells(wsCurRow + 1, wsFolderName), ws.cells(wsCurRow + 1, wsTTSRelayCycles)).value =
                                            ws.range(ws.cells(wsCurRow, wsFolderName), ws.cells(wsCurRow, wsTTSRelayCycles)).value
                                        'except clear out the FINISH-related fields since this isn't a path finish
                                        ws.cells(wsCurRow + 1, wsCFinish).ClearContents
                                        ws.cells(wsCurRow + 1, wsFinishCol).ClearContents

                                        ws.cells(wsCurRow + 1, wsDate).value = myDateValue.Date
                                        ws.cells(wsCurRow + 1, wsLoadTime).value = myDateValue
                                        ws.cells(wsCurRow + 1, wsCStart).value = myDateValue
                                        wsCurRow = wsCurRow + 1
                                        cStartTime = myDateValue

                                        'since we are starting a new entry, save new load time
                                        savedLoadTime = myDateValue
                                    End If
                                Else
                                    state = beginParse
                                End If
                                afterCutStopLineCount = afterCutStopLineCount + 1
                            Case pathFinished
                                'Capture data from this row to new row. 
                                ws.range(ws.cells(wsCurRow + 1, wsFolderName), ws.cells(wsCurRow + 1, wsTTSRelayCycles)).value =
                                    ws.range(ws.cells(wsCurRow, wsFolderName), ws.cells(wsCurRow, wsTTSRelayCycles)).value
                                'Except clear out the "finish"-related fields
                                ws.cells(wsCurRow + 1, wsCFinish).ClearContents
                                ws.cells(wsCurRow + 1, wsFinishCol).ClearContents

                                If MyRegExMatcher(cutStartInd, currentField) Then

                                    'ONLY capture start time if this is a new cut  
                                    'Log file will look like:
                                    'inches from Abs Home: 
                                    'Path Finished. It took 7.3400 min.
                                    'inches from Abs Home: 
                                    'Path Finished. It took 7.3410 min.
                                    state = cutStarted
                                    'Use ParseDateTime to remove the colon so date.Parse can read date/time
                                    myDateTime = ParseDateTime(currentRow(dateTimeField))
                                    myDateValue = Date.Parse(myDateTime)

                                    ws.cells(wsCurRow + 1, wsDate).value = myDateValue.Date
                                    ws.cells(wsCurRow + 1, wsCStart).value = myDateValue
                                    ws.cells(wsCurRow + 1, wsLoadTime).value = myDateValue
                                    wsCurRow = wsCurRow + 1
                                    cStartTime = myDateValue

                                    'since we are starting a new entry, save new load time
                                    savedLoadTime = myDateValue
                                Else
                                    state = beginParse
                                End If
                        End Select
                        prevString = currentField
                    End If
                End While
                MyReader.Close()
            End Using
        Next



        wb.Close(SaveChanges:=True)
        objExcelApp.quit
    End Sub

    Function MyRegExMatcher(findStr As String, searchInStr As String) As Boolean
        ' Instantiate the regular expression object. 
        Dim m As Match = Regex.Match(searchInStr, findStr)
        Return m.Success
    End Function
    Function ParseDateTime(dtStr As String) As String
        Dim lStr, rStr As String
        Dim colonIdx As Integer
        colonIdx = dtStr.LastIndexOf(":")

        lStr = Strings.Left(dtStr, colonIdx)

        rStr = Strings.Right(dtStr, dtStr.Length - colonIdx - 1)
        dtStr = lStr & " " & rStr

        Return dtStr
    End Function
    Function getFieldInfo(fieldIdx As Integer, myStr As String) As String
        'this is to capture the data such as material, material thickness, etc
        Dim colonIdx As Integer
        Dim afterColon As String

        colonIdx = myStr.LastIndexOf(":")
        Select Case fieldIdx
            'these don't need anything special - just capture text after the ":"
            Case wsMaterial, wsMaterialCategory, wsCutModel, wsEtchSpeed,
                 wsEstCost, wsTPLen, wsCutLen
                Return Strings.Right(myStr, myStr.Length - colonIdx - 1)
            Case wsMachineability
                If myStr.EndsWith(")") Then
                    Return "0"
                Else
                    Return Strings.Right(myStr, myStr.Length - colonIdx - 1)
                End If
            Case wsThickness
                afterColon = Strings.Right(myStr, myStr.Length - colonIdx - 1)
                colonIdx = afterColon.IndexOf("i")
                If colonIdx = -1 Then
                    Return afterColon
                Else
                    Return Strings.Left(afterColon, afterColon.IndexOf("i") - 1)
                End If
            Case wsPierces
                afterColon = Strings.Right(myStr, myStr.Length - colonIdx - 1)
                Return Strings.Left(afterColon, afterColon.IndexOf("(") - 1)
            Case wsEstAbrasive
                afterColon = Strings.Right(myStr, myStr.Length - colonIdx - 1)
                Return Strings.Left(afterColon, afterColon.IndexOf("L") - 1)
            'These fields end in min. or hours. - remove the min and return the text
            Case wsEstTime, wsTTSCutting, wsTTSTraversing, wsTTSRelayCycles, wsTTSEtching
                afterColon = Strings.Right(myStr, myStr.Length - colonIdx - 1)

                If MyRegExMatcher("min", afterColon) Then
                    Return Strings.Left(afterColon, afterColon.IndexOf("m") - 1)
                Else
                    Return Strings.Left(afterColon, afterColon.IndexOf("h") - 1)
                End If
        End Select
        Return Strings.Left(myStr, colonIdx)
    End Function

    Private Sub initFields()
        wsMaterial = 13
        wsMaterialCategory = 14
        wsMachineability = 15
        wsThickness = 16
        wsPierces = 17
        wsCutModel = 18
        wsEtchSpeed = 19
        wsEstTime = 20
        wsEstCost = 21
        wsEstAbrasive = 22
        wsTPLen = 23
        wsCutLen = 24
        wsTTSCutting = 25
        wsTTSEtching = 26
        wsTTSTraversing = 27
        wsTTSRelayCycles = 28
    End Sub

End Module
