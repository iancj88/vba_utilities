Option Explicit
Option Base 1
Public Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long

''==============================================================================
'' Program:     FindColNum
'' Desc:        Reads a header row to find the column number
''              associated with the input string
'' Called by:   General Utility
'' Call:        FindColNum(strSearchTerm, vntArrSource, optional intRowHeader)
'' Arguments:   strSearchTerm   --  The exact search term
''              vntArrSource    --  A 2D array containing the header row
'               intRowHeader    --  The row number containing the header text
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date        Programmer       Change
'' 1/17/16    Ian C Johnson     Written
''==============================================================================
Public Function FindColNum(strSearchTerm As String, vntArrSource As Variant, _
                                        Optional ByVal intRowHeader As Integer = 1) As Long
    Dim lngResult As Long, cntI As Long
    lngResult = -1
    ''Loop through the specified row in the 2D array
    For cntI = LBound(vntArrSource, 2) To UBound(vntArrSource, 2)
        ''If an array value matches the search string
        If vntArrSource(intRowHeader, cntI) = strSearchTerm Then
            ''set the result to the counter and stop looping
            lngResult = cntI
            Exit For
        End If
    Next cntI
    
    ''Error Check for debug help
    If lngResult = -1 Then
        Debug.Print strSearchTerm & " Not Found on Row: " & intRowHeader
    End If
    
    FindColNum = lngResult
End Function

''==============================================================================
'' Program:     FindRowNum
'' Desc:        loops through a specified column until a match is made with
''              the string search term. Returns the row number.
'' Called by:   General Utility
'' Call:        FindRowNum(searchTerm, sourceArray, Optional headerCol)
'' Arguments:   strSearchTerm   --  The exact search term
''              vntArrSource    --  A 2D array containing the header column
''              lngColHeader    --  The column number containing the header text
'' Comments:    Written quickly. No error-checking.
'' Changes----------------------------------------------------------------------
'' Date        Programmer       Change
'' 1/17/16    Ian C Johnson     Written
''==============================================================================
Public Function FindRowNum(strSearchTerm As String, vntArrSource() As Variant, _
                                        Optional lngColHeader As Long = 1) As Long
    Dim lngResult As Long, cntI As Integer
    lngResult = -1
    
    ''Loop through the specified row in the 2D array
    For cntI = LBound(vntArrSource, 1) To UBound(vntArrSource, 1)
        ''If an array value matches the search string
        If vntArrSource(cntI, lngColHeader) = strSearchTerm Then
            ''set the result to the counter and stop looping
            lngResult = cntI
            Exit For
        End If
    Next cntI
    
    ''Error Check for debug help
    If lngResult = -1 Then
        MsgBox ("Search term: " & strSearchTerm & "Not Found in " _
                & "Row: " & intRowHeader)
        Debug.Print strSearchTerm; intRowHeader
    End If
    
    FindRowNum = lngResult
End Function

''==============================================================================
'' Program:     ShtToVarArr
'' Desc:        Writes the complete range of cells containing data on a sheet
''              into a 2-dimensional variable array. Array Indexes are equivalent
''              to row and column numbers.
'' Called by:   General Utility
'' Call:        ShtToVarArr(shtIn)
'' Arguments:   shtIn           --  The sheet containing data to be moved into an array
'' Comments:    Written quickly. No error-checking.
'' Changes----------------------------------------------------------------------
'' Date        Programmer       Change
'' 1/17/16    Ian C Johnson     Written
''==============================================================================
Public Function ShtToVarArr(shtIn As Worksheet) As Variant
    Dim vntArrDimensions As Variant '   2x2 array which define the sheet's data range
    Dim strRngAddress As String     '   String value that holds the range address
    strRngAddress = ""
	
	'' Remove the filters from the input sheet
	TurnFilterOff(shtIn)
    
    ''Determine how large the array needs to be
    vntArrDimensions = ShtDimensionsTo2DArr(shtIn)

    ''Find a string value for the range containing all the data
    strRngAddress = Range(Cells(vntArrDimensions(1)(1), vntArrDimensions(1)(2)), _
                    Cells(vntArrDimensions(2)(1), vntArrDimensions(2)(2))).Address
    
    ''Fill the array with the range
    ShtToVarArr = shtIn.Range(strRngAddress).Value
End Function

''==============================================================================
'' Program:     ShtDimensionsTo2DArr
'' Desc:        Determines the innermost and outermost cell containing data
''              on the input sheet. Returns a 2D array with the row,col indexes
''              of these cells.
'' Called by:   General Utility
'' Call:        ShtDimensionsTo2DArr(shtIn)
'' Arguments:   shtIn           --  The sheet whose dimensions will be measured
''              boolUseCellTL   --  optional, use inner cell as top-left aka Cell(1,1)
'' Comments:    Written quickly. No error-checking.
'' Changes----------------------------------------------------------------------
'' Date        Programmer       Change
'' 1/17/16    Ian C Johnson     Written
''==============================================================================
Function ShtDimensionsTo2DArr(shtIn As Worksheet, Optional boolUseCellTL = True) As Variant
    Dim rngFirstCell As Variant, rngLastCell As Variant
    On Error GoTo ErrorHandler
    With shtIn
        ''Search for the outermost aka 'last' cell containing data
        Set rngLastCell = .Cells(.Cells.Find(What:="*", SearchOrder:=xlRows, _
            SearchDirection:=xlPrevious, LookIn:=xlValues).Row, _
            .Cells.Find(What:="*", SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious, LookIn:=xlValues).Column)
            
        ''Search for the innermost cell, typically Cell(1,1)
        '' If it should use the default range for the inner cell
        If boolUseCellTL = True Then
            '' default is set to the range .Cells(1,1)
            Set rngFirstCell = .Cells(1, 1)
        Else
            ''Search for the first cell containing data. If used,
            ''index no longer corresponds to row numbers
            Set rngFirstCell = .Cells(.Cells.Find(What:="*", After:=rngLastCell, SearchOrder:=xlRows, _
                SearchDirection:=xlNext, LookIn:=xlValues).Row, _
                .Cells.Find(What:="*", After:=rngLastCell, SearchOrder:=xlByColumns, _
                SearchDirection:=xlNext, LookIn:=xlValues).Column)
        End If
    End With

    ''Store row and column numbers in temporary arrays

    Dim vntArrtemp(2) As Variant
    Dim vntArrCellInner(2) As Variant
    Dim vntArrCellOuter(2) As Variant
    vntArrCellInner(1) = rngFirstCell.Row
    vntArrCellInner(2) = rngFirstCell.Column
    vntArrCellOuter(1) = rngLastCell.Row
    vntArrCellOuter(2) = rngLastCell.Column
    vntArrtemp(1) = vntArrCellInner
    vntArrtemp(2) = vntArrCellOuter

    ShtDimensionsTo2DArr = vntArrtemp

    ''Free memory from temporary arrays
    Erase vntArrCellInner
    Erase vntArrCellOuter
    Erase vntArrtemp
    Exit Function
ErrorHandler:
    Debug.Print "Nothing found on " & shtIn.Name
    ShtDimensionsTo2DArr = Array(Array(1, 1), Array(1, 1))
End Function

''==============================================================================
'' Program:     MaxRow
'' Desc:        Determines the last row containing data on the input worksheet
'' Called by:   General Utility
'' Call:        MaxRow(sht As Worksheet)
'' Arguments:   shtIn           --  The sheet whose highest row number containing
''                                  data will be found
'' Comments:    Written quickly. No error-checking.
'' Changes----------------------------------------------------------------------
'' Date        Programmer       Change
'' 1/17/16    Ian C Johnson     Written
''==============================================================================
Function MaxRow(shtIn As Worksheet) As Long
    Dim lastRow As Long
    
    'Search backwards for data. Return row number on first instance that it is found.
    With shtIn
        lastRow = .Cells.Find(What:="*", SearchOrder:=xlRows, _
            SearchDirection:=xlPrevious, LookIn:=xlValues).Row
    End With
    
    MaxRow = lastRow
End Function

''==============================================================================
'' Program:     MaxCol
'' Desc:        Determines the last row containing data on the input worksheet
'' Called by:   General Utility
'' Call:        MaxCol(sht As Worksheet)
'' Arguments:   shtIn           --  The sheet whose highest row number containing
''                                  data will be found
'' Comments:    Written quickly. No error-checking.
'' Changes----------------------------------------------------------------------
'' Date        Programmer       Change
'' 1/17/16    Ian C Johnson     Written
''==============================================================================
Function maxCol(shtIn As Worksheet) As Long
    Dim lastCol As Long
    With shtIn
        lastCol = .Cells.Find(What:="*", SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious, LookIn:=xlValues).Column
    End With
    maxCol = lastCol
End Function

''==============================================================================
'' Program:     ShtExists
'' Desc:        Primarily used for error handling. Returns true if a worksheet
''              exists with the given input string name.
'' Called by:   General Utility
'' Call:        ShtExists(SName,wbToCheck)
'' Arguments:   strName         --  The name of the sheet whose existence
''                                  will be tested.
''              wbToCheck       --  The workbook which will be checked for their
''                                  sheet. Defaults to ThisWorkbook
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date        Programmer       Change
'' 1/17/16    Ian C Johnson     Written
''==============================================================================
Function ShtExists(strName As String, Optional ByVal wbToCheck As Workbook) As Boolean
    On Error Resume Next

    ''Default to the calling workbook
    If wbToCheck Is Nothing Then Set wbToCheck = ThisWorkbook
    
    ''The Len function returns the length of the sheet name string.
    '' if the sheet name does not equal 0, then Cbool returns true.
    ShtExists = CBool(Len(wbToCheck.Sheets(strName).Name))
End Function

''==============================================================================
'' Program:     BkExists
'' Desc:        Primarily used for error handling. Returns true if a file exists
''                  in the given directory or [default] workbook directory
'' Called by:   General Utility
'' Call:        BkExists(strFileName, Optional strFileDir)
'' Arguments:   strFileName     --  The name of the file whose existence
''                                  will be tested
''              strFileDir      --  Optional, the directory containing the file.
''                                  Defaults to ThisWorkbook directory
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Function BkExists(strFileName As String, Optional ByVal strFileDir As String) As Boolean
    On Error GoTo 0

    ''Set default directory to the directory of the workbook containing this code
    If strFileDir = vbNullString Then strFileDir = ThisWorkbook.Path & "\"

    ''If path doesn't end in '\' then add it.
    If Right(strFileDir, 1) <> "\" Then strFileDir = strFileDir & "\"
    
    '' use the dir function to test whether the filepath with file name returns a directory
    If Len(Dir(strFileDir & strFileName)) <> 0 Then
        BkExists = True
    Else
        BkExists = False
    End If
End Function

''==============================================================================
'' Program:     BkIsOpen
'' Desc:        Checks whether a book is open given its name
'' Called by:   General Utility
'' Call:        BkIsOpen(strBkName)
'' Arguments:   strBkName       --  the name of the book which will be checked
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Function BkIsOpen(strBkName As String) As Boolean
    On Error Resume Next

    ''If the workbook is nothing, it is not open. Return false.
    BkIsOpen = Not (Application.Workbooks(strBkName) Is Nothing)
End Function


''==============================================================================
'' Program:     BkOpen
'' Desc:        Checks whether a book is open given its name
'' Called by:   General Utility
'' Call:        BkIsOpen(strBkName)
'' Arguments:   strFName       --  the name of the book which will be checked
''								   without the .xlsx or .xlsm suffix
''				strFPath	   --  the path to the directory containing the file
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Public Function BkOpen(strFName As String, strFPath As String) As Workbook
    Dim wbOpened As Workbook
    ''if the book exists, else exit, and is not open,
    ''  then open it, else do nothing
    If BkExists(strFName & ".xlsx", strFPath) Then
        If Not BkIsOpen(strFName) Then
            Set wbOpened = Workbooks.Open(strFPath & strFName)
        Else
            Set wbOpened = Workbooks(strFName)
        End If
            
    Else
        Debug.Print "File not found: " & strFPath & strFName
        Exit Function
    End If
    
    Set BkOpen = wbOpened
End Function

''==============================================================================
'' Program:     ReturnFiscalYear
'' Desc:        Given a calendar date, calculate the Montana State University
''              fiscal year. Return the year as an integer.
'' Called by:   General Utility
'' Call:        ReturnFiscalYear(dteIn)
'' Arguments:   dteIn           --  the date for which the fiscal year will be
''                                  calculated
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
'' 3/1/16       Ian C Johnson       Using DateParts function, can now handle any
''                                  year
''==============================================================================
Public Function ReturnFiscalYear(dteIn As Date) As Integer
	
    Dim intQuarter As Integer
    Dim intYear As Integer
    
    intQuarter = DatePart("q", dteIn)
    intYear = DatePart("yyyy", dteIn)
    
    If intQuarter > 2 Then
        intYear = intYear + 1
    End If

    ReturnFiscalYear = intYear

End Function

''==============================================================================
'' Program:     GetFileNameFromUser
'' Desc:        Prompts the user to enter a file name. Typically before saving.
'' Called by:   General Utility
'' Call:        GetFileNameFromUser(Optional strNameDefault)
'' Arguments:   strNameDefault  --  Optional  default string to be used if
''                                  invalid input is give
'' Comments:	I'm not sure if this works 
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Public Function GetFileNameFromUser(Optional strNameDefault As String) As String
    
    '' Set string to be used as a default file name
    If strNameDefault = Nothing Then strNameDefault = "Default"
    
    Dim strFName As String
    ''Prompt user to enter a filename
    strFName = InputBox(Prompt:="Set Workbook File Name:", _
                        Title:="Workbook File Name", _
                        Default:="file name here")

    ''If the user enters a blank value or leaves the prompt use the default value
    If strFName = "file name here" Or strFName = vbNullString Then
       strFName = strNameDefault
    End If

    GetFileNameFromUser = strFName
End Function

''==============================================================================
'' Program:     BkOpenFromTemplate
'' Desc:        Opens a new workbook from a given template and then saves as a
''              new name. Useful for creating a new file from a template. The
''				template must exist the same directory as the workbook containing
''				the vb module calling this script.
'' Called by:   General Utility
'' Call:        BkOpenFromTemplate(strNewName, Optional sTemplateName, Optional bdateStamp)
'' Arguments:   strNewName      --  The name of the new file to be created from 
''									the template.
''              sTemplateName   --  optional, specify the name of the template file 
''									if it  differs from "template". the .xlsx suffix 
''                                  is not needed.
''              bdateStamp      --  optional, boolean to specify whether to add a
''									date stamp to the new file's name. useful for
''									differentiating multple versions and/or datasets
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Public Function BkOpenFromTemplate(strNewName As String, Optional sTemplateName As String, _
                                    Optional bdateStamp As Boolean) As Workbook

    Dim bkNew As Workbook
    If bTimeStamp = Null Then bTimeStamp = False
    
    
    Dim strDate As String
    If bdateStamp = True Then
        strDate = "_" & Year(Date) & "-" & Month(Date) & "-" & Day(Date)
    Else
        strDate = vbNullString
    End If
    If sTemplateName = "" Then
        Set bkNew = Workbooks.Add(ThisWorkbook.Path & "\template.xlsx")
    Else
        Set bkNew = Workbooks.Add(ThisWorkbook.Path & "\" & sTemplateName & ".xlsx")
    End If
    
    'Debug.Print ThisWorkbook.Path
    'Debug.Print strNewName & strDate
    bkNew.SaveAs (ThisWorkbook.Path & "/" & strNewName & strDate & ".xlsx")

    bkNew.Save
    Set BkOpenFromTemplate = bkNew
    
End Function

''==============================================================================
'' Program:     BkAddNew
'' Desc:        Creates a new workbook and saves it. If no file name used in
''              the arguments, it will prompt the user to enter one. Saves the
''              workbook in teh calling workbooks directory.
'' Called by:   General Utility
'' Call:        BkAddNew(Optional strFilePath, Optional strBkName)
'' Arguments:   strFilePath     --  Optional file path for the new book. Will
''                                  use ThisWorkbook path if none is give.
''                                  invalid input is give
''              strBkName       --  Optional name for the new workbook. Will
''                                  prompt user for name if none is given via
''                                  argument
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Public Function BkAddNew(Optional ByVal strFilePath As String, _
                            Optional strBkName As String) As Workbook
    ''Set default file path to be the same as ThisWorkbook
    Dim strNewFileFolder As String
    Dim strNewFileType As String
    Dim bkNew As Workbook
    

    strNewFileType = "Excel Files 1997-2003 (*.xls), *.xls," & _
               "Excel Files 2007 (*.xlsx), *.xlsx," & _
               "All files (*.*), *.*"
               
    If strFilePath = "" Then
        strFilePath = Application.GetSaveAsFilename( _
                        InitialFileName:=strNewFileFolder, _
                        fileFilter:=strNewFileType)
    
    End If
    
    Set bkNew = Workbooks.Add
    Application.DisplayAlerts = False
    bkNew.SaveAs Filename:=strFilePath
    Application.DisplayAlerts = True
    Set BkAddNew = bkNew
End Function

''==============================================================================
'' Program:     ShtAddNew
'' Desc:        Creates a new Sheet in the specified workbook with the specified
''              name. If fEraseSht1 is true, then simply rename Sheet 1. This is
''              Useful when using a newly created workbook.
'' Called by:   General Utility
'' Call:        ShtAddNew(bkEffected, strShtName, Optional fEraseSht1)
'' Arguments:   bkEffected      --  The book in which the sheet will be created
''              strShtName      --  The name of the new worksheet
''              fEraseSht1      --  boolean to determine whether to use the
''                                  the default 'Sheet1' as the new sheet. Will
''                                  erase that sheet and then rename it. CAUTION
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 3/1/16       Ian C Johnson       Written
''==============================================================================
Public Function ShtAddNew(bkEffected As Workbook, strShtName As String, _
                        Optional fEraseSht1 As Boolean = False) As Worksheet
    If fEraseSht1 = False Then
        bkEffected.Worksheets.Add().Name = strShtName
    Else
        Call ClearSht(bkEffected.Worksheets("Sheet1"))
        bkEffected.Worksheets("Sheet1").Name = strShtName
    End If
End Function

''==============================================================================
'' Program:     CellToStrRng
'' Desc:        Takes two cells and finds name of the range if they were connected.
''              Outputs the range as a string value.
'' Called by:   General Utility
'' Call:        CellToStrRng(vntArrInnerCell, vntArrOuterCell)
'' Arguments:   vntArrInnerCell --  Variant Array used to designate the inner i.e.
''                                  smallest row,col cell. 1st value is the row
''                                  2nd Value is the column
''              vntArrOuterCell --  See explanation for InnerCell array except this
''                                  array represents the outer cell with the
''                                  largest row,col values.
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Public Function CellToStrRng(vntArrInnerCell As Variant, vntArrOuterCell As Variant) As String
    CellToStrRng = Range(Cells(vntArrInnerCell(1), vntArrInnerCell(2)), _
                    Cells(vntArrOuterCell(1), vntArrOuterCell(2))).Address
End Function

''==============================================================================
'' Program:     UniqueItems
'' Desc:        Loops through an array to returns an array of unique values or
''              the number of unique items.
'' Called by:   General Utility
'' Call:        UniqueItems(vntArrIn, Count)
'' Arguments:   vntArrIn		--  the input array to be looped through
''              Count			--  optional, true to return a count of unique
''									items rather than the items themselves
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Function UniqueItems(vntArrIn As Variant, Optional Count As Variant) As Variant
	Dim Unique() As Variant ''array that holds the unique items
	Dim Element As Variant  ''value contained in the input array
	Dim NumUnique As Long   ''number of unique items
	Dim i As Integer        ''counter variable
	Dim FoundMatch As Boolean ''designates whether the item is already contained in
							''unique Array.

	''If 2nd argument is missing, assign default value
	If Count = vbNullString Then Count = True
	''Counter for number of unique elements
	NumUnique = 0
	
	''Loop thru the input array
	For Each Element In vntArrIn
		FoundMatch = False
		
		''Has item been added yet?
		For i = 1 To NumUnique
			If Element = Unique(i) Then
				FoundMatch = True
				Exit For ''(exit loop)
			End If
		Next i
		
		''If not in list, add the item to unique list
		If Not FoundMatch And Not IsEmpty(Element) Then
			NumUnique = NumUnique + 1
			ReDim Preserve Unique(NumUnique)
			Unique(NumUnique) = Element
		End If
	Next Element
	
	''Assign a value to the function
	If Count Then UniqueItems = NumUnique Else UniqueItems = Unique
End Function

''==============================================================================
'' Program:     ClearSht
'' Desc:        Clears a sheet of all values. WARNING: Cannot be undone.
'' Called by:   General Utility
'' Call:        ClearSht(shtIn)
'' Arguments:   shtIn       --      The Sheet which will be cleared of values
'' Comments:    Use with caution. Does not delete formatting.
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Sub ClearSht(shtIn As Worksheet)
    Dim tempArray() As Variant
    Dim rngStr As String
    ''Find the dimensions of the data on the sheet
    tempArray = ShtDimensionsTo2DArr(shtIn)
    
    ''Turn the dimensions into a string range address
    rngStr = CellToStrRng(tempArray(1), tempArray(2))
    
    ''Clear the sheet using the address derived above
    shtIn.Range(rngStr).Clear

    Erase tempArray
End Sub

''==============================================================================
'' Program:     ArrToDict
'' Desc:        Takes an 2 dimensional array of data and returns a dictionary
''              of keys and values from two respective columns. Depends on the 
''				Microsoft Scripting Runtime found Tools -> References
'' Called by:   General Utility
'' Call:        ArrToDict(vntArrIn, headerRow, strKey, strVal)
'' Arguments:   vntArrIn    --      The array containing the data from which to
''                                  to derive the dictionary
''              headerRow   --      The row index number which contains header
''                                  info to match input strKey and strVal
''              strKey      --      The name of the column which will be used
''                                  as keys in the output dict.
''              strVal      --      The name of the column which will be used
''                                  as values in the output dict.
'' Comments:    If a value already exists in the dictionary, a message is
''              printed to the console, and the new value is not added.
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
'' 3/16/16      Ian C Johnson       Can handle already existing key in dict
''==============================================================================
Public Function ArrToDict(vntArrIn As Variant, _
                            headerRow As Integer, _
                            strKey As String, strVal As String) As Dictionary

    Dim colKey As Long, colVal As Long
    colKey = FindColNum(strKey, vntArrIn, headerRow)
    colVal = FindColNum(strVal, vntArrIn, headerRow)
    
    Dim tempDict As Dictionary
    Set tempDict = New Dictionary
    Dim irow As Long
    
    For irow = LBound(vntArrIn, 1) To UBound(vntArrIn, 1)
        If vntArrIn(irow, colKey) <> "" And vntArrIn(irow, colVal) <> "" Then
            If tempDict.Exists(Trim(vntArrIn(irow, colKey))) Then
                'Debug.Print "The value exists in the dictionary. Key: "; vntArrIn(iRow, colKey)
            Else
                tempDict.Add Trim(vntArrIn(irow, colKey)), Trim(vntArrIn(irow, colVal))
            End If
        End If
    Next irow
    
    Set ArrToDict = tempDict
    Set tempDict = Nothing
End Function

''==============================================================================
'' Program:     NumberOfDimensions
'' Desc:        Returns the number of dimensions of a given array. Primarily for
''				error checking
'' Called by:   General Utility
'' Call:        NumberOfDimensions(arrIn)
'' Arguments:   arrIn       --      The array to whose # of dimensions will
''                                  be measured
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Public Function NumberOfDimensions(arrIn As Variant) As Integer
    Dim ErrorCheck As Boolean
    Dim FinalDimension As Long
    ''Setup Error handler to return final dimension of the array
    On Error GoTo FinalDimension
    
    ''Arrays may have up to 60k dimensions
    Dim DimNum As Long
    For DimNum = 1 To 60000
        ''Check if the DimNum throughs an error
        ErrorCheck = LBound(arrIn, DimNum)
    Next DimNum
    Exit Function
    
    ''Return the final dimension minus one (the one right before the error)
FinalDimension:
        NumberOfDimensions = DimNum - 1
End Function

''==============================================================================
'' Program:     StrReplaceDictVal
'' Desc:        Replaces a every instance of a specific dictionary value
''              with another value.
'' Called by:   General Utility
'' Call:        StrReplaceDictVal(dictIn, strOldVal, strNewVal)
'' Arguments:   dictIn      --      The Dictionary which will be updated
''              strOldVal   --      The value which will be replaced
''              strNewVal   --      The new value which will take the old
''                                  values place
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Public Function StrReplaceDictVal(dictIn As Dictionary, strOldVal As String, _
                                strNewVal As String) As Dictionary
    Dim strKey As Variant
    Dim dictOut As Dictionary
    Set dictOut = New Dictionary
    ''loop through the dictionary.
    For Each strKey In dictIn.Keys()
        ''Anytime the item in the dictionary matches the input value,
        ''replace it with the new item
        If dictIn(strKey) = strOldVal Then
            dictOut.Add strKey, strNewVal
        Else
            dictOut.Add strKey, dictIn(strKey)
        End If
    Next strKey
    Set dictIn = Nothing
    Set StrReplaceDictVal = dictOut
End Function

''==============================================================================
'' Program:     ReturnVntFilePaths
'' Desc:        Gets filepaths to specific files from the user. Returns an
''              array with these values
'' Called by:   General Utility
'' Call:        ReturnVntFilePaths()
'' Comments:	Requires further testing. may not work properly.
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
Public Function ReturnVntFilePaths() As Variant()
	
    ''Use the filedialog object to browse for files
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim filePaths() As Variant
    
    On Error GoTo ErrorHandler
    
    ''initialize and show the file dialog window
    With fd
        .InitialFileName = ActiveWorkbook.Path & Application.PathSeparator
        .AllowMultiSelect = True
        .ButtonName = "Select File(s) to Load"
        .Title = "Load File(s)"
        .Show
            ''If one or more files have been selected, add them to the array
            If fd.SelectedItems(1) <> vbNullString Then
                Dim i As Integer
                i = 0
                Dim vntSelectedItem As Variant
                For Each vntSelectedItem In .SelectedItems
                    i = i + 1
                    ''make the array big enough for the new filepath
                    ReDim Preserve filePaths(i)
                    filePaths(i) = vntSelectedItem
                Next vntSelectedItem
            End If
    End With
    
    ReturnVntFilePaths = filePaths

    ''Clear memory before exiting
    Erase filePaths
    Set fd = Nothing
    Exit Function

ErrorHandler:
    Set fd = Nothing
    Erase filePaths
    MsgBox "Error " & Err & ": " & Error(Err)

End Function

''==============================================================================
'' Program:     LoadRptToVnt
'' Desc:        Takes a csv or delimited text file and outputs into a 2D array
'' Called by:   General Utility
'' Call:        LoadRptLinesToVnt(strFilePath, strDelim, Optional ignoreShtData)
'' Arguments:   strFilePath     --      the file path to the file to be split
''              strDelim        --      The deliminiter used to split each field
'' Comments:	May not work properly.....
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Public Function LoadRptLinesToVnt(strFilePath As String, _
                                strDelim As String) As Variant

    Dim vntArrFileData() As Variant
    Dim strLineFromFile As String
    Dim LineItems() As Variant
    
    Open strFilePath For Input As #1
    Dim i As Integer
    i = 0
    Do Until EOF(1)
        Line Input #1, strLineFromFile
        LineItems = Split(strLineFromFile, strDelim)
        i = i + 1
        ReDim Preserve vntArrFileData(i)
        vntArrFileData(i) = LineItems
    Loop
    
    Close #1
    LoadRptToVnt = vntArrFileData
    Erase vntArrFileData
End Function

''==============================================================================
'' Program:     SpeedSettingsOn
'' Desc:        Sets certain application settings to speed up vba code runtime
'' Called by:   General Utility
'' Call:        SpeedSettingsOn(speedUp)
'' Arguments:   speedUp         --      boolean, true to turn on the speed settings
'' Comments:	Be sure to set this to FALSE at the end of the module. Otherwise,
''				use an error handler to turn this back on if something goes wrong.
''				It's a PITA finding all these settings in gui console.
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 1/17/16      Ian C Johnson       Written
''==============================================================================
Public Sub SpeedSettingsOn(speedUp As Boolean)
     If speedUp = True Then
        '' Turn off select application features/settings
        Application.ScreenUpdating = False
        Application.DisplayStatusBar = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        ActiveSheet.DisplayPageBreaks = False 'note this is a sheet-level setting
    End If
    If speedUp = False Then
        '' Turn on select application features/settings
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        ActiveSheet.DisplayPageBreaks = False 'note this is a sheet-level setting
        Exit Sub
    End If
End Sub

''==============================================================================
'' Program:     FindMin
'' Desc:        Returns the smallest value from a 1D array. Returns as double.
'' Called by:   General Utility
'' Call:        FindMin(vntArr1DIn, fIgnoreBlank)
'' Arguments:   vntArr1DIn          --  the 1D array containing numerical or
''                                      date values.
''              fIgnoreBlank        --  Whether the test should be applied to
''                                      blank or empty values.
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 3/1/16       Ian C Johnson       Written
''==============================================================================
Public Function FindMin(vntArr1DIn As Variant, fIgnoreBlank As Boolean) As Double
    Dim dblMinVal As Double
    Dim Value As Variant
    
    dblMinVal = MaxDouble()
    
    For Each Value In vntArr1DIn
        If TypeName(Value) <> "String" And TypeName(Value) <> "Empty" Then
            If fIgnoreBlank = True Then
                If Value <> "" Then
                    If Value <> CDate(0) Then
                        If CDbl(Value) < dblMinVal Then dblMinVal = Value
                    End If
                End If
            Else
                If CDbl(Value) < dblMinVal Then dblMinVal = Value
            End If
        End If
    Next Value
    
    FindMin = dblMinVal
End Function

''==============================================================================
'' Program:     FindMax
'' Desc:        Returns the largest value from a 1D array. Returns as double.
'' Called by:   General Utility
'' Call:        FindMax(vntArr1DIn, fIgnoreBlank)
'' Arguments:   vntArr1DIn          --  the 1D array containing numerical or
''                                      date values.
''              fIgnoreBlank        --  Whether the test should be applied to
''                                      blank or empty values.
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 3/1/16       Ian C Johnson       Written
''==============================================================================
Public Function FindMax(vntArr1DIn As Variant, fIgnoreBlank As Boolean) As Double
    Dim dblMaxVal As Double
    Dim Value As Variant
    
    dblMaxVal = MinDouble()
    
    For Each Value In vntArr1DIn
        If TypeName(Value) <> "String" And TypeName(Value) <> "Empty" Then
            If fIgnoreBlank = True Then
                If Value <> "" Then
                    If Value <> CDate(0) Then
                        If CDbl(Value) > dblMaxVal Then dblMaxVal = Value
                    End If
                End If
            Else
                If CDbl(Value) > dblMaxVal Then dblMaxVal = Value
            End If
        End If
    Next Value
    FindMax = dblMaxVal
End Function

''==============================================================================
'' Program:     ArrToRangeStr
'' Desc:        Returns a string value of a range appropriately sized for the
''              given array. Optional values allow row or column offset.
'' Called by:   General Utility
'' Call:        ArrToRangeStr(vntArrIn, fIgnoreBlank)
'' Arguments:   vntArrIn            --  the array which will be written to the
''                                      book
''              intOffsetRow        --  if the row value of the range should be
''                                      offset from the top
''              intOffsetCol        --  value of the column offset from the left
'' Comments:	This enables quickly writing an array to a sheet. much easier to
''				set a range to the values of an array than to create looping
''				structures through the x,y dimensions of the array and sheet.
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 3/16/16       Ian C Johnson       Written
''==============================================================================
Public Function ArrToRangeStr(vntArrIn As Variant, _
                                Optional intOffsetRow As Integer = 0, _
                                Optional intOffsetCol As Integer = 0) As String
    Dim lngArrMaxX As Long, lngArrMaxY As Long
    
    ''Ensure that the number of dimensions is less than or equal to 2
    If NumberOfDimensions(vntArrIn) < 2 Then
        Debug.Print "ArrToRangeStr handed array with more than 2D"
        ArrToRangeStr = vbNullString
        Exit Function
    End If
    
    ''Get numerical value of max row and col. Offset if necessary
    lngArrMaxX = UBound(vntArrIn, 2) + intOffsetCol
    lngArrMaxY = UBound(vntArrIn, 1) + intOffsetRow
    
    ''test if it is base 0. If so, add 1 because row/columns are base 1
    If LBound(vntArrIn, 2) = 0 Then lngArrMaxX = lngArrMaxX + 1
    If LBound(vntArrIn, 1) = 0 Then lngArrMaxY = lngArrMaxY + 1
    
    ''derive range address given offset and max X,Y values
    ArrToRangeStr = Range(Cells(1 + intOffsetRow, 1 + intOffsetCol), _
                          Cells(lngArrMaxY, lngArrMaxX)).Address
End Function

''==============================================================================
'' Program:     SetSourceFileStr
'' Desc:        Used to set the location of any source files used in the project
''              Change the strFilePath variable here so that dictionaries and
''              other objects initialized using outside data look to the correct
''              location when loading.
'' Called by:   General Utility
'' Call:        SetSourceFileStr()
'' Comments:	If you have some common directory for lookup tables, set it here
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 3/21/16      Ian C Johnson      Written
''==============================================================================
Private Function SetSourceFileStr() As String
    Dim strFilePath As String
    ''Define the location of excel files containg dictionary tables
    strFilePath = "C:\VBA_Source\Tables\"
    SetSourceFileStr = strFilePath
End Function

''==============================================================================
'' Program:     GetDeptOrgDict
'' Desc:        Returns a dictionary of the Depts as keys and Organizations as
''              items. This useful whenever doiong rollups to higher level
''              organizations as the finance numbers do not always coincide with
''              their true hierarchy.
'' Arguements:  iLvl    Integer where 1 is the VP level org and 2 is the EMR Org
'' Called by:   General Utility
'' Call:        GetDeptOrgDict()
'' Comments:	This is an example of programmatically opening and reading
''				another excel file to create a lookup dictionary
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 3/21/16      Ian C Johnson       Written
''==============================================================================
Public Function GetDeptOrgDict(Optional iLvl As Integer = 2) As Dictionary
    Dim strFilePath As String
    strFilePath = SetSourceFileStr()
    Dim strFileName As String
    strFileName = "TableDept"
    Dim strWsName As String
    strWsName = "TableDept"
	
    ''strKey and strItem identifies the org and
    '' dept columns by header name to be used as keys and values
    Dim strKey As String
    Dim strItem As String
    
    strKey = "Sub Dept"
    If iLvl = 2 Then
        strItem = "Dept Hierarchy Lvl2" ''EMR Orgs
    ElseIf iLvl = 1 Then
        strItem = "Dept Hierarchy Lvl1" ''VP Orgs
    Else
        Debug.Print "GetDeptOrgDict handed invalid parameter = " & iLvl
    End If
    
    ''Open TableDeptOrg File
    Dim wbDictValues As Workbook
    Set wbDictValues = BkOpen(strFileName, strFilePath)
    
    'Load table into array
    Dim vntArrTable() As Variant
    vntArrTable = ShtToVarArr(wbDictValues.Worksheets(strWsName))
   
    ''Pull the dept and org into a dictionary
    Dim dictOut As Dictionary
    Dim iHeaderRow As Integer
    iHeaderRow = 2
    Set dictOut = ArrToDict(vntArrTable, iHeaderRow, strKey, strItem)
    
    ''Close workbook and return dictionary
    wbDictValues.Close
    Set GetDeptOrgDict = dictOut
    
    ''Free up some memory
    Set dictOut = Nothing
    Erase vntArrTable
    Set wbDictValues = Nothing
End Function

''==============================================================================
'' Program:     IncrementDictItem
'' Desc:        A script to be used for counting key values. If a key value does
''				not exist, it is added with a value of the increment (1 in most
''				cases). Otherwise, if the key already exists in the passed
''				dictionary, it increments the value by one. 
'' Called by:   General Utility
'' Call:        IncrementDictItem(dictIn, strKey, Opt iIncrmnt)
'' Arguments:   dictIn              --  the dictionary to be incremented
''              strKey              --  the Key value to be incremented
''              iIncrmnt            --  optional increment value, default is 1
'' Comments:	The values of the passed dictionary must be double or integer(!)
''				It is easiest to hand this function an empty dictionary at the 
''				start of the counting loop(s)
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 3/22/16      Ian C Johnson       Written
''==============================================================================
Public Function IncrementDictItem(dictIn As Dictionary, _
                                    ByVal key As Variant, _
                                    Optional dblIncrmnt As Double = 1) As Dictionary
    ''If the item exists, increment it, if not, add it to the dict
    ''  with the increment as the initial item
    If dictIn.Exists(key) Then
        dictIn(key) = dictIn(key) + dblIncrmnt
    Else
        dictIn.Add key, dblIncrmnt
    End If
    Set IncrementDictItem = dictIn

End Function

''==============================================================================
'' Program:     AbbreviateGIDS
'' Desc:        takes a column of an array, truncates the GIDs, and formats
''              properly
'' Called by:   General Utility
'' Call:        IncrementDictItem(vntArrIn, Opt strColGID, Opt lngColHeader)
'' Arguments:   vntArrIn            --  the array containing GIDs values in one
''                                      column
''              strColGID           --  optional string GID column name
''              lngColHeader        --  optional header row value
'' Comments:	Due to the security policies regarding 9 digit gids, any document
''				containing names and gids must abbreviate the gids. 
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 3/22/16      Ian C Johnson       Written
''==============================================================================
Public Function AbbreviateGIDS(vntArrIn As Variant, Optional strColGID As String = "GID", _
                                Optional lngColHeader As Long = 1) As Variant
    Dim irow As Long
    Dim lngColGID As Long
    Dim strGID As String
    lngColGID = FindColNum(strColGID, vntArrIn, lngColHeader)
    
    For irow = lngColHeader + 1 To UBound(vntArrIn, 1)
        strGID = vntArrIn(irow, lngColGID)
        If strGID <> "" Then
            vntArrIn(irow, lngColGID) = "-" & Right(strGID, 4)
        End If
    Next irow
    
    AbbreviateGIDS = vntArrIn

End Function

''==============================================================================
'' Program:     CreateHeaderDict
'' Desc:        Create a dictionary object and fill it with column names and 
''				integer index. Column names are the key and column number is the
''				value. Useful for mapping arrays of data
'' Called by:   General Utility
'' Call:        CreateHeaderDict(vArrIn, iRowHeader)
'' Arguments:   vArrIn              --  the array of data from which the 
''										dictionary will be derived								
''              iRowHeader          --  the row number containing column names
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 3/22/16      Ian C Johnson       Written
''==============================================================================
Public Function CreateHeaderDict(vArrIn() As Variant, iRowHeader As Integer) As Dictionary
    Dim dictOut As Dictionary
    Set dictOut = New Dictionary
    
    Dim icol As Integer
    For icol = LBound(vArrIn, 2) To UBound(vArrIn, 2)
        If vArrIn(iRowHeader, icol) <> vbNullString And Not dictOut.Exists(vArrIn(iRowHeader, icol)) Then
            dictOut.Add vArrIn(iRowHeader, icol), icol
        End If
    Next icol
    Set CreateHeaderDict = dictOut
End Function

''==============================================================================
'' Program:     ShellSortNumbers
'' Desc:        Use the shell-sort algorithm to sort an array of numbers
'' Called by:   General Utility
'' Call:        ShellSortNumbers(varray)
'' Arguments:   varray	            --  the 1D array to be sorted
'' Comments:	
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 3/22/16      VBPro	  	        downloaded from http://www.vbcode.com/asp/showsn.asp?theID=568
''==============================================================================
Public Sub ShellSortNumbers(vArray As Variant)
  Dim lLoop1 As Long
  Dim lHold As Long
  Dim lHValue As Long
  Dim lTemp As Long


  lHValue = LBound(vArray)
  Do
    lHValue = 3 * lHValue + 1
  Loop Until lHValue > UBound(vArray)
  Do
    lHValue = lHValue / 3
    For lLoop1 = lHValue + LBound(vArray) To UBound(vArray)
      lTemp = vArray(lLoop1)
      lHold = lLoop1
      Do While vArray(lHold - lHValue) > lTemp
        vArray(lHold) = vArray(lHold - lHValue)
        lHold = lHold - lHValue
        If lHold < lHValue Then Exit Do
      Loop
      vArray(lHold) = lTemp
    Next lLoop1
  Loop Until lHValue = LBound(vArray)
End Sub


''==============================================================================
'' Program:     BrowseForFolder
'' Desc:        allow a user to select a system folder and return that folder's path
'' Called by:   General Utility
'' Call:        BrowseForFolder(OpenAt)
'' Arguments:   OpenAt	            --  An optional folderpath to open the original folder 
'' Comments:	
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 3/22/16      VBPro	  	        modified from 
									http://stackoverflow.com/questions/19372319/vba-folder-picker-set-where-to-start#19373904
''==============================================================================
Public Function BrowseForFolder(Optional OpenAt As String) As Variant
     'Function purpose:  To Browser for a user selected folder.
     'If the "OpenAt" path is provided, open the browser at that directory
     'NOTE:  If invalid, it will open at the Desktop level

    Dim ShellApp As Object
    
    If Right(OpenAt, 1) <> "\" Then
        OpenAt = OpenAt & "\"
    End If
	
     'Create a file browser window at the default folder
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Select Folder for Reports", 0, OpenAt)

     'Set the folder to that selected.  (On error in case cancelled)
    On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
    On Error GoTo 0

     'Destroy the Shell Application
    Set ShellApp = Nothing

     'Check for invalid or non-entries and send to the Invalid error
     'handler if found
     'Valid selections can begin L: (where L is a letter) or
     '\\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
    Case Is = ":"
        If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select

    Exit Function

Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFolder = False
End Function

''==============================================================================
'' Program:     GetUniqueJCATDict
'' Desc:        Returns a dictionary of the GID-POSN Keys and JCAT Assignation
''              as items. THis will allow a particular position to be linked with
''              a JCAT, but may not allow if the person changes.
'' Called by:   General Utility
'' Call:        GetUniqueJCATDict()
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 8/2/16       Ian C Johnson       Written
''==============================================================================
Public Function GetJCATDict(Optional bGetUniqueAsKeyUserInputAsItem = False, _
                            Optional bGetJCATCodeAsKeyJCATTitleAsItem = False) As Dictionary
    Dim strFilePath As String
    strFilePath = SetSourceFileStr()
    
    Dim strFileName As String
    strFileName = "TableJCAT"
    Dim strWbName As String
    
        
    Dim strKey As String
    Dim strItem As String
    
    ''Determine sheet, key, and item columns depending on input parameters
    If bGetUniqueAsKeyUserInputAsItem Then
        strWbName = "TableCurrJCAT"
        strKey = "Unique"
        strItem = "Entered JCAT"
    Else
        strWbName = "TableJCATXwalk"
        strKey = "JCAT"
        strItem = "Minor Code/Description"
    End If
    
    ''Open TableDeptOrg File
    Dim wbDictValues As Workbook
    Set wbDictValues = BkOpen(strFileName, strFilePath)
    
    'Load table into array
    Dim vntArrTable() As Variant
    vntArrTable = ShtToVarArr(wbDictValues.Worksheets(strWbName))
   

   'Create the unique key and define item column

    
''''May want to implement a way to creat uniquekey dynamically
    'Dim iColkey1 as Integer, iColkey2 as integer
    'iColKey1 = FindColNum(vntArrTable,"GID",1)
    'iColKey2 = FindColNum(vntArrTable,"POSN",1)
    
    ''Pull the dept and org into a dictionary
    Dim dictOut As Dictionary
    Dim iHeaderRow As Integer
    iHeaderRow = 1
    Set dictOut = ArrToDict(vntArrTable, iHeaderRow, strKey, strItem)
    
    ''Close workbook and return dictionary
    wbDictValues.Close
    Set GetJCATDict = dictOut
    
    ''Free up some memory
    Set dictOut = Nothing
    Erase vntArrTable
    Set wbDictValues = Nothing
End Function

''==============================================================================
'' Program:     ComputeMedianFromCntDict
'' Desc:        Compute the key of median value from a dictionary containing 
''				numbers as values.
'' Called by:   General Utility
'' Call:        ComputeMedianFromCntDict(dictIn)
'' Arguments:   dictIn            --  the dictionary of numbers as values
'' Comments:
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 8/2/16       Ian C Johnson       Written
''==============================================================================
Public Function ComputeMedianFromCntDict(dictIn As Dictionary) As Double
    Dim TempArr() As Variant
    ''Write values to array
    Dim itemCnt As Long
    Dim vkey As Variant, lngKey As Long
    For Each vkey In dictIn.Keys()
        itemCnt = itemCnt + dictIn(vkey)
    Next vkey
    
    Dim itemKeyCnt As Long, i As Long
    
    ReDim TempArr(1 To dictItemCnt)
    For Each vkey In dictIn.Keys()
        lngKey = vkey
        itemKeyCnt = dictIn(vkey)
        i = 1
        Do While itemKeyCnt > 0
            TempArr(i) = lngKey
            i = i + 1
            itemKeyCnt = itemKeyCnt - 1
        Loop
    Next vkey
    
    TempArr = CombSort(TempArr, 1, False)
    
    Dim median As Double
    Dim medianIndx As Long
    
    medianIndx = (dictItemCnt + 1) / 2
    median = TempArr(medianIndx)
    ComputeMedianFromCntDict = median
End Function

''==============================================================================
'' Program:     LoadIndexFromAllEEReport
'' Desc:       	This won't work without the funding source object. THis should 
''				be moved out of the general utility file.
'' Called by:   General Utility
'' Call:        LoadIndexFromAllEEReport(wsIn, iHeaderRow)
'' Arguments:   wsIn            	--  the worksheet with the all employee data
''				iHeaderRow			--	headerrow if it differs from the all employee
''										report's default of 13.
'' Comments:	
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 8/2/16       Ian C Johnson       Written
''==============================================================================
Public Function LoadIndexFromAllEEReport(wsIn As Worksheet, Optional iHeaderRow As Integer = 13) As Dictionary
    Dim vaAllEE() As Variant
    vaAllEE = ShtToVarArr(wsIn)
    Dim icolGID As Integer, icolPOSN As Integer, iColSuffx As Integer
    Dim iColIndx As Integer, iColPercent As Integer

    icolGID = FindColNum("GID", vaAllEE, iHeaderRow)
    icolPOSN = FindColNum("Position Number", vaAllEE, iHeaderRow)
    iColSuffx = FindColNum("Suffix", vaAllEE, iHeaderRow)
    iColIndx = FindColNum("Index", vaAllEE, iHeaderRow)
    iColPercent = FindColNum("Percent", vaAllEE, iHeaderRow)
    
    Dim dictJobsFundSrc As Dictionary
    Set dictJobsFundSrc = New Dictionary
    Dim irow As Integer
    Dim uniqueKey As String
    Dim oFundSources As FundingSources
    Dim indx As String
    Dim percent As Double
    For irow = iHeaderRow + 1 To UBound(vaAllEE, 1)
        ''Create the key
       indx = vaAllEE(irow, iColIndx)
       percent = vaAllEE(irow, iColPercent)
       
        uniqueKey = vaAllEE(irow, icolGID) & vaAllEE(irow, icolPOSN) & vaAllEE(irow, iColSuffx)
        If dictJobsFundSrc.Exists(uniqueKey) Then
            Set oFundSources = dictJobsFundSrc(uniqueKey)
        Else
            Set oFundSources = New FundingSources
        End If
        Call oFundSources.addFund(indx, percent)
        Set dictJobsFundSrc.Item(uniqueKey) = oFundSources
    Next irow
    Set LoadIndexFromAllEEReport = dictJobsFundSrc
End Function

''==============================================================================
'' Program:     CopyArrRowToSht
'' Desc:        Simply loop through one row of data and write it to a sheet.
'' Called by:   General Utility
'' Call:        CopyArrRowToSht(wsOut vaDataIn, sheetRow, arrayRow)
'' Arguments:   wsOut            	--  the worhseet to write the data to
''				vaDataIn		 	--	the array containing the data to write
''				sheetRow			--	the row on the sheet which will be overwritten
''				arrayRow			--	the row in the array containing the data 
''										to be written.
'' Comments:	Useful for selectively writing rows that may match some logical
''				criteria.
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 8/2/16       Ian C Johnson       Written
''==============================================================================
Public Sub CopyArrRowToSht(wsOut As Worksheet, vaDataIn As Variant, sheetRow As Integer, arrayRow As Integer)
    Dim icol As Integer
    For icol = LBound(vaDataIn, 2) To UBound(vaDataIn, 2)
        wsOut.Cells(sheetRow, icol) = vaDataIn(arrayRow, icol)
    Next icol
End Sub

''==============================================================================
'' Program:     PadGID
'' Desc:        take a GID string and add necessary zeros if they are missing.
''				Commonly occurs when a gid value is loaded as an integer rather
''				than a string. Returns a string containing the full GID
'' Called by:   General Utility
'' Call:        PadGID(sGID)
'' Arguments:   sGID	            --  a full or partial gid of string type
'' Comments:	because gids are in the format "-[0-9]{8}", zeros may be dropped
''				between the dash and first non-zero number.
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 3/22/16      Ian C Johnson       created
''==============================================================================
Public Function PadGID(sGID As String) As String
    Dim currLen As Integer
    
    Do While Len(sGID) < 9
        currLen = Len(sGID)
        sGID = Right(sGID, currLen - 1)
        sGID = "-0" & sGID
    Loop
    
    PadGID = sGID
End Function

''==============================================================================
'' Program:     TurnFilterOff
'' Desc:        removes AutoFilter if one exists on the sheet
'' Called by:   ShtToVarArr
'' Call:        TurnFilterOff(wsIn)
'' Arguments:   dictIn            --  the worksheet to have filters removed.
'' Comments:	filters generally screw things up/lead to incomplete datasets 
''				when loading a sheet into an array. It is safest to always call
''				this before loading into an array
'' Changes----------------------------------------------------------------------
'' Date         Programmer          Change
'' 8/2/16       Ian C Johnson       Written
''==============================================================================
Sub TurnFilterOff(wsIn)
    'removes AutoFilter if one exists
     wsIn.AutoFilterMode = False
End Sub



