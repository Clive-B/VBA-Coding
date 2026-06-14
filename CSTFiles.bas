Attribute VB_Name = "CSTFiles"
' Add this at the top of your module (before any procedures)
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub ProcessCSTFiles()
    Dim fso As Object, folder As Object, file As Object
    Dim townDateDict As Object
    Dim townDateKey As Variant
    Dim filePat As Variant
    Dim filePath As String
    Dim mocFiles As Object, mtcFiles As Object
    Dim townDateKeys As Variant
    Dim fileList As Object
    Dim workbookPath As String
    Dim closeworkBook As String
    Dim myArray As Variant
    Dim element As Variant
    Dim oWSHShell As Object
    Dim folderPath As String
    Dim Trends As String
    Dim KPI As Scripting.TextStream
    Dim lastMTCFile As String
    Dim Finish As Boolean
    Dim Count As Integer
    Dim isFirstCycleMOC As Boolean
    Dim isFirstCycleMTC As Boolean
    Dim lastFileInFolder As String
    Dim isLastElement As Boolean ' New variable to track last array element
    Dim callType As String
    Dim openedWorkbook As Workbook

    Count = 0
    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop") & "\"
    Set oWSHShell = Nothing
    Trends = Sheets("Option").Range("AO1").Value
    myArray = Array("MTN", "TELECEL", "AT")
    
    For Each element In myArray
        ' Determine if this is the last element in the array
        isLastElement = (element = myArray(UBound(myArray))) ' New logic
        
        folderPath = GetDesktop & "QoS Automation\" & Trends & "\Telcos\" & element & "\CST\LOGS"
        workbookPath = GetDesktop & "QoS Automation\Templates\Voice Template.xlsm"
        
        ' Get last file in current folder
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(folderPath) Then GoTo NextElement

        Set folder = fso.GetFolder(folderPath)
        lastFileInFolder = GetLastFileInFolder(folder)
        Set townDateDict = CreateObject("Scripting.Dictionary")
        closeworkBook = NextworkbookPath()
        Sleep 3000
        Set openedWorkbook = Workbooks.Open(workbookPath)

        ' File processing
        For Each file In folder.files
            Dim town As String, fileDate As String, key As String
            town = ExtractTown(file.Name)
            fileDate = ExtractDate(file.Name)
            If town <> "" And fileDate <> "" Then
                key = town & "_" & fileDate
                If Not townDateDict.Exists(key) Then
                    Set townDateDict(key) = New Collection
                End If
                townDateDict(key).Add file.Path
            End If
        Next file

        townDateKeys = townDateDict.Keys
        isFirstCycleMOC = True
        isFirstCycleMTC = True

        For Each townDateKey In townDateKeys
            Set mocFiles = New Collection
            Set mtcFiles = New Collection
            Set fileList = townDateDict(townDateKey)

            For Each filePat In fileList
                callType = GetCstCallType(CStr(filePat), fso)
                If callType = "MOC" Then
                    mocFiles.Add filePat
                ElseIf callType = "MTC" Then
                    mtcFiles.Add filePat
                End If
            Next filePat
            
            If mtcFiles.Count > 0 Then
                lastMTCFile = mtcFiles(mtcFiles.Count)
            Else
                lastMTCFile = ""
            End If

            ' Process MOC files with LastFileInFolder and LastArrayElement
            For Each filePat In mocFiles
                Finish = False
                workbookPath = openedWorkbook.FullName
                Dim isLastMOC As Boolean
                isLastMOC = (filePat = lastFileInFolder)
                ' Pass isLastElement to ProcessFile
                Call ProcessFile(filePat, element, Trends, workbookPath, Finish, isLastMOC, isLastElement)
            Next filePat

            ' Process MTC files with LastFileInFolder and LastArrayElement
            For Each filePat In mtcFiles
                Finish = (filePat = lastMTCFile)
                workbookPath = openedWorkbook.FullName
                Dim isLastMTC As Boolean
                isLastMTC = (filePat = lastFileInFolder)
                ' Pass isLastElement to ProcessFile
                Call ProcessFile(filePat, element, Trends, workbookPath, Finish, isLastMTC, isLastElement)
            Next filePat
        Next townDateKey
NextElement:
    Next element

    Set fso = Nothing
    Set folder = Nothing
    Set townDateDict = Nothing
    Set mocFiles = Nothing
    Set mtcFiles = Nothing
    Set fileList = Nothing
End Sub
Function ExtractTown(fileName As String) As String
    Dim parts() As String
    parts = Split(fileName, " ")
    If UBound(parts) >= 2 Then
        ExtractTown = parts(2) ' Adjust index based on filename structure
    Else
        ExtractTown = ""
    End If
End Function

Function ExtractDate(fileName As String) As String
    Dim parts() As String
    parts = Split(fileName, " ")
    If UBound(parts) >= 0 Then
        ExtractDate = parts(0) ' Adjust index based on filename structure
    Else
        ExtractDate = ""
    End If
End Function
Private Function FindCstWorkbookPath() As String
    Dim wb As Workbook
    Dim workbookPath As String
    
    ' Initialize the result as an empty string
    workbookPath = ""
    
    ' Loop through all open workbooks
    For Each wb In Workbooks
        ' Exclude "QoS.xlsm" and find the other workbook
        If wb.Name <> "QoS.xlsm" Then
            workbookPath = wb.FullName ' Get the full path of the other workbook
            Exit For ' Exit the loop once the other workbook is found
        End If
    Next wb
    
    ' Return the path of the other workbook (empty string if not found)
    FindCstWorkbookPath = workbookPath
End Function
Function NextworkbookPath() As String
    Dim wb As Workbook
    Dim closeworkBook As String
    
    ' Initialize the result as an empty string
    closeworkBook = ""
    
    ' Loop through all open workbooks
    For Each wb In Workbooks
        ' Exclude "QoS.xlsm" and find the other workbook
        If wb.Name <> "QoS.xlsm" And wb.Name <> "Voice Template.xlsm" Then
            closeworkBook = wb.FullName ' Get the full path of the other workbook
            wb.Close SaveChanges:=False
            Exit For ' Exit the loop once the other workbook is found
        End If
    Next wb
    
End Function
Function GetLastFileInFolder(folder As Object) As String
    Dim lastFile As Object
    Dim file As Object
    Dim lastFileName As String
    
    ' Initialize with empty string
    lastFileName = ""
    
    ' Loop through all files in the folder
    For Each file In folder.files
        ' Compare filenames alphabetically
        If file.Name > lastFileName Then
            lastFileName = file.Name
            Set lastFile = file
        End If
    Next file
    
    ' Return the path of the last file (or empty if folder is empty)
    If Not lastFile Is Nothing Then
        GetLastFileInFolder = lastFile.Path
    Else
        GetLastFileInFolder = ""
    End If
End Function
Private Function GetCstCallType(ByVal filePat As String, ByVal fso As Object) As String
    Dim KPI As Object
    Dim ArryLine As String
    Dim CurrentLine() As String
    Dim Keep() As String

    On Error GoTo CleanUp
    Set KPI = fso.OpenTextFile(filePat)
    Do Until KPI.AtEndOfStream
        ArryLine = KPI.ReadLine
        CurrentLine = Split(ArryLine, ",")
        If UBound(CurrentLine) >= 3 Then
            If CurrentLine(0) = "#DL" Then
                Keep = Split(CurrentLine(3), " ")
                If UBound(Keep) >= 1 Then
                    GetCstCallType = UCase$(Replace(Keep(1), """", ""))
                    Exit Do
                End If
            End If
        End If
    Loop

CleanUp:
    On Error Resume Next
    If Not KPI Is Nothing Then KPI.Close
End Function

' Updated ProcessFile with new parameter
Sub ProcessFile(filePat, element, Trends, workbookPath As String, Finish As Boolean, Last As Boolean, isLastArrayElement As Boolean)
    Dim macroName As String
    macroName = "CSTCOVMOS"
    Fil = filePat
    DirPath = Trends
    
    ' Use isLastArrayElement for final actions
    If isLastArrayElement Then
        ' Add any logic needed for the last telco here
        ' Example: Save workbooks, trigger final reports, etc.
    End If
    
    ' Existing macro call
    Application.Run "'" & workbookPath & "'!" & macroName, Fil, element, DirPath, Finish, Last, isLastArrayElement
End Sub
