Attribute VB_Name = "ModuleOldFormat"


Private myFiles() As String
Private Fnum As Long

Sub RDB_Merge_Data_BrowseNume()
    UserForm1.Hide
    
    Dim myFiles As Variant
    Dim myCountOfFiles As Long
    Dim oApp As Object
    Dim oFolder As Variant
    Dim myValueName As String
    
    Set oApp = CreateObject("Shell.Application")

    'Browse to the folder
    Set oFolder = oApp.BrowseForFolder(0, "Select folder", 512)
    If Not oFolder Is Nothing Then

        myCountOfFiles = Get_File_Names( _
                         MyPath:=oFolder.Self.Path, _
                         SubFolders:=False, _
                         ExtStr:="*.xl*", _
                         myReturnedFiles:=myFiles)

        If myCountOfFiles = 0 Then
            MsgBox "No files that match the ExtStr in this folder"
            Exit Sub
        End If
        myValueName = InputBox("Numele sheet-ului de unde se va copia informatia:")
        Get_Data _
                FileNameInA:=True, _
                PasteAsValues:=True, _
                SourceShName:=myValueName, _
                SourceShIndex:=0, _
                SourceRng:="A1:IV65536", _
                StartCell:="", _
                myReturnedFiles:=myFiles

    End If

End Sub


Sub RDB_Merge_Data_BrowseIndex()
    UserForm1.Hide
    
    Dim myFiles As Variant
    Dim myCountOfFiles As Long
    Dim oApp As Object
    Dim oFolder As Variant
    Dim myValueIndex As Integer
    
    Set oApp = CreateObject("Shell.Application")

    'Browse to the folder
    Set oFolder = oApp.BrowseForFolder(0, "Select folder", 512)
    If Not oFolder Is Nothing Then

        myCountOfFiles = Get_File_Names( _
                         MyPath:=oFolder.Self.Path, _
                         SubFolders:=False, _
                         ExtStr:="*.xl*", _
                         myReturnedFiles:=myFiles)

        If myCountOfFiles = 0 Then
            MsgBox "No files that match the ExtStr in this folder"
            Exit Sub
        End If
        myValueIndex = InputBox("Numarul sheet-ului de unde se va copia informatia:")
        Get_Data _
                FileNameInA:=True, _
                PasteAsValues:=True, _
                SourceShName:="", _
                SourceShIndex:=myValueIndex, _
                SourceRng:="A1:IV65536", _
                StartCell:="", _
                myReturnedFiles:=myFiles

    End If

End Sub

Function Get_File_Names(MyPath As String, SubFolders As Boolean, _
                        ExtStr As String, myReturnedFiles As Variant) As Long

    Dim Fso_Obj As Object, RootFolder As Object
    Dim SubFolderInRoot As Object, file As Object

    'Add a slash at the end if the user forget it
    If Right(MyPath, 1) <> "\" Then
        MyPath = MyPath & "\"
    End If

    'Create FileSystemObject object
    Set Fso_Obj = CreateObject("Scripting.FileSystemObject")

    Erase myFiles()
    Fnum = 0

    'Test if the folder exist and set RootFolder
    If Fso_Obj.FolderExists(MyPath) = False Then
        Exit Function
    End If
    Set RootFolder = Fso_Obj.GetFolder(MyPath)

    'Fill the array(myFiles)with the list of Excel files in the folder(s)
    'Loop through the files in the RootFolder
    For Each file In RootFolder.Files
        If LCase(file.Name) Like LCase(ExtStr) Then
            Fnum = Fnum + 1
            ReDim Preserve myFiles(1 To Fnum)
            myFiles(Fnum) = MyPath & file.Name
        End If
    Next file

    'Loop through the files in the Sub Folders if SubFolders = True
    If SubFolders Then
        Call ListFilesInSubfolders(OfFolder:=RootFolder, FileExt:=ExtStr)
    End If

    myReturnedFiles = myFiles
    Get_File_Names = Fnum
End Function


Function ListFilesInSubfolders(OfFolder As Object, FileExt As String)



    Dim SubFolder As Object
    Dim fileInSubfolder As Object

    For Each SubFolder In OfFolder.SubFolders
        ListFilesInSubfolders OfFolder:=SubFolder, FileExt:=FileExt

        For Each fileInSubfolder In SubFolder.Files
            If LCase(fileInSubfolder.Name) Like LCase(FileExt) Then
                Fnum = Fnum + 1
                ReDim Preserve myFiles(1 To Fnum)
                myFiles(Fnum) = SubFolder & "\" & fileInSubfolder.Name
            End If
        Next fileInSubfolder

    Next SubFolder
End Function


Function RDB_Last(choice As Integer, rng As Range)

' 1 = last row
' 2 = last column
' 3 = last cell
    Dim lrw As Long
    Dim lcol As Integer

    Select Case choice

    Case 1:
        On Error Resume Next
        RDB_Last = rng.Find(What:="*", _
                            After:=rng.Cells(1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
        On Error GoTo 0

    Case 2:
        On Error Resume Next
        RDB_Last = rng.Find(What:="*", _
                            After:=rng.Cells(1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
        On Error GoTo 0

    Case 3:
        On Error Resume Next
        lrw = rng.Find(What:="*", _
                       After:=rng.Cells(1), _
                       LookAt:=xlPart, _
                       LookIn:=xlFormulas, _
                       SearchOrder:=xlByRows, _
                       SearchDirection:=xlPrevious, _
                       MatchCase:=False).Row
        On Error GoTo 0

        On Error Resume Next
        lcol = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        LookAt:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

        On Error Resume Next
        RDB_Last = rng.Parent.Cells(lrw, lcol).Address(False, False)
        If Err.Number > 0 Then
            RDB_Last = rng.Cells(1).Address(False, False)
            Err.Clear
        End If
        On Error GoTo 0

    End Select
End Function









' Note: You not have to change the macro below, you only
' edit and run the RDB_Merge_Data above.


Function Get_Data(FileNameInA As Boolean, PasteAsValues As Boolean, SourceShName As String, _
             SourceShIndex As Integer, SourceRng As String, StartCell As String, myReturnedFiles As Variant)
    Dim SourceRcount As Long
    Dim SourceRange As Range, destrange As Range
    Dim mybook As Workbook, BaseWks As Worksheet
    Dim rnum As Long, CalcMode As Long
    Dim SourceSh As Variant
    Dim sh As Worksheet
    Dim i As Long

    'Change ScreenUpdating, Calculation and EnableEvents
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    'Add a new workbook with one sheet named "Combine Sheet"
    Set BaseWks = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    BaseWks.Name = "Combine Sheet"

    'Set start row for the Data
    rnum = 1

    'Check if we use a named sheet or the index
    If SourceShName = "" Then
        SourceSh = SourceShIndex
    Else
        SourceSh = SourceShName
    End If

    'Loop through all files in the array(myFiles)
    For i = LBound(myReturnedFiles) To UBound(myReturnedFiles)
        Set mybook = Nothing
        On Error Resume Next
        Set mybook = Workbooks.Open(myReturnedFiles(i))
        On Error GoTo 0

        If Not mybook Is Nothing Then

            If LCase(SourceShName) <> "all" Then

                'Set SourceRange and check if it is a valid range
                On Error Resume Next

                If StartCell <> "" Then
                    With mybook.Sheets(SourceSh)
                        Set SourceRange = .Range(StartCell & ":" & RDB_Last(3, .Cells))
                        'Test if the row of the last cell >= then the row of the StartCell
                        If RDB_Last(1, .Cells) < .Range(StartCell).Row Then
                            Set SourceRange = Nothing
                        End If
                    End With
                Else
                    With mybook.Sheets(SourceSh)
                        Set SourceRange = Application.Intersect(.UsedRange, .Range(SourceRng))
                    End With
                End If

                If Err.Number > 0 Then
                    Err.Clear
                    Set SourceRange = Nothing
                Else
                    'if SourceRange use all columns then skip this file
                    If SourceRange.Columns.Count >= BaseWks.Columns.Count Then
                        Set SourceRange = Nothing
                    End If
                End If

                On Error GoTo 0

                If Not SourceRange Is Nothing Then

                    'Check if there enough rows to paste the data
                    SourceRcount = SourceRange.Rows.Count
                    If rnum + SourceRcount >= BaseWks.Rows.Count Then
                        MsgBox "Sorry there are not enough rows in the sheet to paste"
                        mybook.Close savechanges:=False
                        BaseWks.Parent.Close savechanges:=False
                        GoTo ExitTheSub
                    End If

                    'Set the destination cell
                    If FileNameInA = True Then
                        Set destrange = BaseWks.Range("B" & rnum)
                        With SourceRange
                            BaseWks.Cells(rnum, "A"). _
                                    Resize(.Rows.Count).Value = myReturnedFiles(i)
                        End With
                    Else
                        Set destrange = BaseWks.Range("A" & rnum)
                    End If

                    'Copy/paste the data
                    If PasteAsValues = True Then
                        With SourceRange
                            Set destrange = destrange. _
                                            Resize(.Rows.Count, .Columns.Count)
                        End With
                        destrange.Value = SourceRange.Value
                    Else
                        SourceRange.Copy destrange
                    End If

                    rnum = rnum + SourceRcount
                End If

                'Close the workbook without saving
                mybook.Close savechanges:=False

            Else

                'Loop through all sheets in mybook
                For Each sh In mybook.Worksheets

                    'Set SourceRange and check if it is a valid range
                    On Error Resume Next

                    If StartCell <> "" Then
                        With sh
                            Set SourceRange = .Range(StartCell & ":" & RDB_Last(3, .Cells))
                            If RDB_Last(1, .Cells) < .Range(StartCell).Row Then
                                Set SourceRange = Nothing
                            End If
                        End With
                    Else
                        With sh
                            Set SourceRange = Application.Intersect(.UsedRange, .Range(SourceRng))
                        End With
                    End If

                    If Err.Number > 0 Then
                        Err.Clear
                        Set SourceRange = Nothing
                    Else
                        'if SourceRange use almost all columns then skip this file
                        If SourceRange.Columns.Count > BaseWks.Columns.Count - 2 Then
                            Set SourceRange = Nothing
                        End If
                    End If
                    On Error GoTo 0

                    If Not SourceRange Is Nothing Then

                        'Check if there enough rows to paste the data
                        SourceRcount = SourceRange.Rows.Count
                        If rnum + SourceRcount >= BaseWks.Rows.Count Then
                            MsgBox "Sorry there are not enough rows in the sheet to paste"
                            mybook.Close savechanges:=False
                            BaseWks.Parent.Close savechanges:=False
                            GoTo ExitTheSub
                        End If

                        'Set the destination cell
                        If FileNameInA = True Then
                            Set destrange = BaseWks.Range("C" & rnum)
                            With SourceRange
                                BaseWks.Cells(rnum, "A"). _
                                        Resize(.Rows.Count).Value = myReturnedFiles(i)
                                BaseWks.Cells(rnum, "B"). _
                                        Resize(.Rows.Count).Value = sh.Name
                            End With
                        Else
                            Set destrange = BaseWks.Range("A" & rnum)
                        End If

                        'Copy/paste the data
                        If PasteAsValues = True Then
                            With SourceRange
                                Set destrange = destrange. _
                                                Resize(.Rows.Count, .Columns.Count)
                            End With
                            destrange.Value = SourceRange.Value
                        Else
                            SourceRange.Copy destrange
                        End If

                        rnum = rnum + SourceRcount
                    End If

                Next sh

                'Close the workbook without saving
                mybook.Close savechanges:=False
            End If
        End If

        'Open the next workbook
    Next i

    'Set the column width in the new workbook
    BaseWks.Columns.AutoFit

ExitTheSub:
    'Restore ScreenUpdating, Calculation and EnableEvents
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Function


Function RDB_Filter_Data()
    Dim myFiles As Variant
    Dim myCountOfFiles As Long

    myCountOfFiles = Get_File_Names( _
                     MyPath:="C:\Users\Ron\test", _
                     SubFolders:=False, _
                     ExtStr:="*.xl*", _
                     myReturnedFiles:=myFiles)


    If myCountOfFiles = 0 Then
        MsgBox "No files that match the ExtStr in this folder"
        Exit Function
    End If

    Get_Filter _
            FileNameInA:=True, _
            SourceShName:="", _
            SourceShIndex:=1, _
            FilterRng:="A1:D" & Rows.Count, _
            FilterField:=1, _
            FilterValue:="ron", _
            myReturnedFiles:=myFiles

End Function


' Note: You not have to change the macro below, you only
' edit and run the RDB_Filter_Data above.


Function Get_Filter(FileNameInA As Boolean, SourceShName As String, _
               SourceShIndex As Integer, FilterRng As String, FilterField As Integer, _
               FilterValue As String, myReturnedFiles As Variant)
    Dim SourceRange As Range, destrange As Range
    Dim mybook As Workbook, BaseWks As Worksheet
    Dim rnum As Long, CalcMode As Long
    Dim SourceSh As Variant
    Dim rng As Range
    Dim RwCount As Long
    Dim i As Long

    'Change ScreenUpdating, Calculation and EnableEvents
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    'Add a new workbook with one sheet named "Combine Sheet"
    Set BaseWks = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    BaseWks.Name = "Combine Sheet"

    'Set start row for the Data
    rnum = 1

    'Check if we use a named sheet or the index
    If SourceShName = "" Then
        SourceSh = SourceShIndex
    Else
        SourceSh = SourceShName
    End If

    'Loop through all files in the array(myFiles)
    For i = LBound(myReturnedFiles) To UBound(myReturnedFiles)
        Set mybook = Nothing
        On Error Resume Next
        Set mybook = Workbooks.Open(myReturnedFiles(i))
        On Error GoTo 0

        If Not mybook Is Nothing Then

            'Set SourceRange and check if it is a valid range
            On Error Resume Next

            With mybook.Sheets(SourceSh)
                Set SourceRange = Application.Intersect(.UsedRange, .Range(FilterRng))
            End With

            If Err.Number > 0 Then
                Err.Clear
                Set SourceRange = Nothing
            Else
                'if SourceRange use all columns then skip this file
                If SourceRange.Columns.Count >= BaseWks.Columns.Count Then
                    Set SourceRange = Nothing
                End If
            End If
            On Error GoTo 0

            If Not SourceRange Is Nothing Then

                'Find the last row in BaseWks
                rnum = RDB_Last(1, BaseWks.Cells) + 1

                With SourceRange.Parent
                    Set rng = Nothing

                    'Firstly, remove the AutoFilter
                    .AutoFilterMode = False

                    'Filter the range on the FilterField column
                    SourceRange.AutoFilter Field:=FilterField, _
                                           Criteria1:=FilterValue

                    With .AutoFilter.Range
                        'Check if there are results after you use AutoFilter
                        RwCount = .Columns(1).Cells. _
                                  SpecialCells(xlCellTypeVisible).Cells.Count - 1

                        If RwCount = 0 Then
                            'There is no data, only the header
                        Else
                            ' Set a range without the Header row
                            Set rng = .Resize(.Rows.Count - 1, .Columns.Count). _
                                      Offset(1, 0).SpecialCells(xlCellTypeVisible)

                            If FileNameInA = True Then
                                'Copy the range and the file name
                                If rnum + RwCount < BaseWks.Rows.Count Then
                                    BaseWks.Cells(rnum, "A").Resize(RwCount).Value _
                                          = mybook.Name
                                    rng.Copy BaseWks.Cells(rnum, "B")
                                End If
                            Else
                                'Copy the range
                                If rnum + RwCount < BaseWks.Rows.Count Then
                                    rng.Copy BaseWks.Cells(rnum, "A")
                                End If
                            End If
                        End If
                    End With

                    'Remove the AutoFilter
                    .AutoFilterMode = False

                End With
            End If

            'Close the workbook without saving
            mybook.Close savechanges:=False
        End If

        'Open the next workbook
    Next i

    'Set the column width in the new workbook
    BaseWks.Columns.AutoFit

    'Restore ScreenUpdating, Calculation and EnableEvents
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Function



Function RDB_Copy_Sheet()
    Dim myFiles As Variant
    Dim myCountOfFiles As Long

    myCountOfFiles = Get_File_Names( _
                     MyPath:="C:\Users\Ron\test", _
                     SubFolders:=False, _
                     ExtStr:="*.xl*", _
                     myReturnedFiles:=myFiles)

    If myCountOfFiles = 0 Then
        MsgBox "No files that match the ExtStr in this folder"
        Exit Function
    End If

    Get_Sheet _
            PasteAsValues:=True, _
            SourceShName:="", _
            SourceShIndex:=1, _
            myReturnedFiles:=myFiles

End Function



' Note: You not have to change the macro below, you only
' edit and run the RDB_Copy_Sheet above.


Sub Get_Sheet(PasteAsValues As Boolean, SourceShName As String, _
              SourceShIndex As Integer, myReturnedFiles As Variant)
    Dim mybook As Workbook, BaseWks As Worksheet
    Dim CalcMode As Long
    Dim SourceSh As Variant
    Dim sh As Worksheet
    Dim i As Long

    'Change ScreenUpdating, Calculation and EnableEvents
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    On Error GoTo ExitTheSub

    'Add a new workbook with one sheet
    Set BaseWks = Workbooks.Add(xlWBATWorksheet).Worksheets(1)


    'Check if we use a named sheet or the index
    If SourceShName = "" Then
        SourceSh = SourceShIndex
    Else
        SourceSh = SourceShName
    End If

    'Loop through all files in the array(myFiles)
    For i = LBound(myReturnedFiles) To UBound(myReturnedFiles)
        Set mybook = Nothing
        On Error Resume Next
        Set mybook = Workbooks.Open(myReturnedFiles(i))
        On Error GoTo 0

        If Not mybook Is Nothing Then

            'Set sh and check if it is a valid
            On Error Resume Next
            Set sh = mybook.Sheets(SourceSh)

            If Err.Number > 0 Then
                Err.Clear
                Set sh = Nothing
            End If
            On Error GoTo 0

            If Not sh Is Nothing Then
                sh.Copy After:=BaseWks.Parent.Sheets(BaseWks.Parent.Sheets.Count)

                On Error Resume Next
                ActiveSheet.Name = mybook.Name
                On Error GoTo 0

                If PasteAsValues = True Then
                    With ActiveSheet.UsedRange
                        .Value = .Value
                    End With
                End If

            End If
            'Close the workbook without saving
            mybook.Close savechanges:=False
        End If

        'Open the next workbook
    Next i

    ' delete the first sheet in the workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    BaseWks.Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

ExitTheSub:
    'Restore ScreenUpdating, Calculation and EnableEvents
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub








