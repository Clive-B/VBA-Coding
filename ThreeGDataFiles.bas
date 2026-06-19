Attribute VB_Name = "ThreeGDataFiles"
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Module-level dictionaries to track latest files by town and day
Private latest3GFilesByTownDay As Object
Private latest4GFilesByTownDay As Object
Public NameTown As String
'Public TownName As String
Public Function GetPublicVar() As String
    GetPublicVar = NameTown
End Function
Sub ProcessDataFiles()
    Dim fso As Object, folder As Object, file As Object
    Dim townDateDict As Object, locDict As Object
    Dim townDateKey As Variant, locKey As Variant
    Dim element As Variant, fileItem As Variant
    Dim workbookPath As String, GetDesktop As String, folderPath As String
    Dim oWSHShell As Object, trends As String
    Dim KPI As Object, Keep() As String, ArryLine As String, CurrentLine() As String
    Dim T3GFiles As Collection, T4GFiles As Collection
    Dim isLastElement As Boolean, hasDuplicates As Boolean
    Dim MyArray As Variant, loc As String, lastLocFile, DupBreak As String
    Dim isLastInLocation As Boolean, filePath As String, filePat As String
    Dim passNumber As Integer, locCount As Long, currentLocIndex As Long
    Dim isLastLocationInTownDate As Boolean
    Dim locations As Variant
    Dim key As String, town As String, fileDate As String
    Dim locationName As String
    Dim FinalTime3G As Boolean, FinalTime4G As Boolean
    Dim fileDict As Object
    Dim townDayKey As String
    Dim fileList As Collection
    Dim hasMultipleTimes As Boolean

    ' Initialize file tracking
    InitializeFileTracking
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop") & "\"
    Set oWSHShell = Nothing
    trends = Sheets("Option").Range("AO1").value
    MyArray = Array("MTN", "TELECEL", "AT")
    
    For Each element In MyArray
        isLastElement = (element = MyArray(UBound(MyArray)))
        folderPath = GetDesktop & "QoS Automation\" & trends & "\Telcos\" & element & "\DATA\LOGS"
        workbookPath = GetDesktop & "QoS Automation\Templates\3G Data Template.xlsm"
        
        Set folder = fso.GetFolder(folderPath)
        Set townDateDict = CreateObject("Scripting.Dictionary")
        Set fileList = New Collection
        
        ' First, collect all files in order
        For Each file In folder.files
            fileList.Add file.Path
        Next file
        
        NextworkbookPath
        Sleep 3000
        Workbooks.Open workbookPath

        ' Group files by town + date in their natural order
        For Each fileItem In fileList
            Dim fileName As String
            fileName = fso.GetFileName(fileItem)
            town = ExtractTown(fileName)
            locationName = ExtractLocation(fileName)
            fileDate = ExtractDate(fileName)
            
            If town <> "" And fileDate <> "" Then
                key = town & "_" & fileDate
                If Not townDateDict.Exists(key) Then
                    townDateDict.Add key, CreateObject("Scripting.Dictionary")
                End If
                If townDateDict(key).Exists(locationName) Then
                    townDateDict(key)(locationName).Add fileItem
                Else
                    townDateDict(key).Add locationName, New Collection
                    townDateDict(key)(locationName).Add fileItem
                End If
            End If
        Next fileItem

        For passNumber = 1 To 2
            For Each townDateKey In townDateDict.Keys
                Set T3GFiles = New Collection
                Set T4GFiles = New Collection
                hasDuplicates = (townDateDict(townDateKey).count > 1)
                
                ' Extract town and date from the key
                Dim townParts() As String
                townParts = Split(townDateKey, "_")
                town = townParts(0)
                fileDate = townParts(1)
                townDayKey = town & "_" & fileDate
                
                ' Get all locations for this town-date
                locations = townDateDict(townDateKey).Keys
                locCount = townDateDict(townDateKey).count
                
                ' Process each location in order
                currentLocIndex = 0
                For Each locKey In locations
                    currentLocIndex = currentLocIndex + 1
                    isLastLocationInTownDate = (currentLocIndex = locCount)
                    
                    ' Create a dictionary to track files by technology
                    Set fileDict = CreateObject("Scripting.Dictionary")
                    
                    ' Process files in their natural order
                    For Each fileItem In townDateDict(townDateKey)(locKey)
                        filePath = CStr(fileItem)
                        Set KPI = fso.OpenTextFile(filePath)
                        Do Until KPI.AtEndOfStream
                            ArryLine = KPI.ReadLine
                            CurrentLine = Split(ArryLine, ",")
                            If UBound(CurrentLine) >= 3 And CurrentLine(0) = "#DL" Then
                                Keep = Split(CurrentLine(3), " ")
                                If UBound(Keep) >= 1 Then
                                    Keep(1) = Replace(Keep(1), """", "")
                                End If
                            End If
                        Loop
                        KPI.Close

                        If UBound(Keep) >= 1 Then
                            If Keep(1) = "3G" Then
                                T3GFiles.Add filePath
                                If Not fileDict.Exists("3G") Then
                                    fileDict.Add "3G", New Collection
                                End If
                                fileDict("3G").Add filePath
                            ElseIf Keep(1) = "4G" Then
                                T4GFiles.Add filePath
                                If Not fileDict.Exists("4G") Then
                                    fileDict.Add "4G", New Collection
                                End If
                                fileDict("4G").Add filePath
                            End If
                        End If
                    Next fileItem

                    ' Process 3G files for location in their natural order
                    If Not fileDict Is Nothing And fileDict.Exists("3G") Then
                        ' Find the latest file without sorting
                        lastLocFile = FindLatestFile(fileDict("3G"), fso, "3G", townDayKey)
                        
                        For Each fileItem In fileDict("3G")
                            filePath = fileItem
                            isLastInLocation = (filePath = lastLocFile)
                            workbookPath = FindworkbookPath()
                            FinalTime3G = IsLastFileForTownDay(filePath, fso, "3G")
                            ' Process each file in its original order
                            NameTown = town
                            fileName = fso.GetFileName(fileItem)
                            locationName = ExtractLocation(fileName)
                            DupBreak = fileDate & " " & NameTown & " " & locationName
                            ProcessFile filePath, DupBreak, fileName, locationName, fileDate, trends, workbookPath, False, _
                                      isLastLocationInTownDate, isLastElement, hasDuplicates, _
                                      isLastInLocation, FinalTime3G, passNumber, locKey, CStr(locKey)
                        Next fileItem
                    End If

                    ' Process 4G files for location in their natural order
                    If Not fileDict Is Nothing And fileDict.Exists("4G") Then
                        ' Find the latest file without sorting
                        lastLocFile = FindLatestFile(fileDict("4G"), fso, "4G", townDayKey)
                        
                        For Each fileItem In fileDict("4G")
                            filePath = fileItem
                            isLastInLocation = (filePath = lastLocFile)
                            workbookPath = FindworkbookPath()
                            FinalTime4G = IsLastFileForTownDay(filePath, fso, "4G")
                            ' Process each file in its original order
'                            ProcessFile filePath, Trends, workbookPath, False, _
'                                      isLastLocationInTownDate, isLastElement, hasDuplicates, _
'                                      isLastInLocation, FinalTime4G, passNumber, locKey, CStr(locKey)
                        Next fileItem
                    End If
                    
                    Set fileDict = Nothing
                Next locKey
            Next townDateKey
        Next passNumber
        
        ' Cleanup between elements
        NextworkbookPath
        Sleep 3000
        Workbooks.Open workbookPath
        Set folder = Nothing
        Set townDateDict = Nothing
        Set fileList = Nothing
    Next element
    
    On Error Resume Next ' Handle errors if workbook isn't found
    Set wb = Workbooks(GetFileName(workbookPath))
    On Error GoTo 0
    
    ' If workbook found, close it without saving changes
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=True
    End If
    
    Call Process4GDataFiles
    Set fso = Nothing
End Sub
Sub ProcessFile( _
    filePath As String, _
    DupBreak As String, _
    fileName As String, _
    locationName As String, _
    fileDate As String, _
    trends As String, _
    workbookPath As String, _
    Finish As Boolean, _
    isLastLocationInTownDate As Boolean, _
    isLastElement As Boolean, _
    hasDuplicates As Boolean, _
    isLastInLocation As Boolean, _
    FinalTime As Boolean, _
    passNumber As Integer, _
    locKey As Variant, _
    ByVal location As String)
    
    Dim macroName As String
    Dim fil As String
    Dim dirPath As String
    macroName = "ThreeGData"
    fil = filePath
    dirPath = trends
    LogfileName = fileName
'    Fil As String, DupBreak As String, LogfileName As String, locationName As String
    
    Application.Run "'" & workbookPath & "'!" & macroName, _
    fil, DupBreak, LogfileName, locationName, fileDate, dirPath, Finish, isLastLocationInTownDate, isLastElement, isLastInLocation, FinalTime, passNumber, location, locKey
End Sub
Function ExtractTown(fileName As String) As String
    Dim parts() As String
    parts = Split(fileName, " ")
    If UBound(parts) >= 2 Then
        ExtractTown = parts(2)
    Else
        ExtractTown = ""
    End If
End Function

Function ExtractLocation(fileName As String) As String
    Dim parts() As String, i As Long
    parts = Split(fileName, " ")
    For i = 3 To UBound(parts)
        If parts(i) = "DATA" Then Exit For
        ExtractLocation = ExtractLocation & parts(i) & " "
    Next
    ExtractLocation = Trim(ExtractLocation)
End Function

Function ExtractDate(fileName As String) As String
    Dim parts() As String
    parts = Split(fileName, " ")
    If UBound(parts) >= 0 Then
        ExtractDate = parts(0)
    Else
        ExtractDate = ""
    End If
End Function

Function ExtractTime(fileName As String) As String
    Dim parts() As String
    parts = Split(fileName, " ")
    If UBound(parts) >= 1 Then
        ExtractTime = parts(1)
    Else
        ExtractTime = "000000"
    End If
End Function

Sub NextworkbookPath()
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name <> "QoS.xlsm" And wb.Name <> "3G Data Template.xlsm" Then
            wb.Close SaveChanges:=False
            Exit For
        End If
    Next
End Sub

Function FindworkbookPath() As String
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name <> "QoS.xlsm" Then
            FindworkbookPath = wb.fullName
            Exit Function
        End If
    Next
    FindworkbookPath = ""
End Function

Sub InitializeFileTracking()
    Set latest3GFilesByTownDay = CreateObject("Scripting.Dictionary")
    Set latest4GFilesByTownDay = CreateObject("Scripting.Dictionary")
End Sub
Function FindLatestFile(files As Collection, fso As Object, techType As String, townDayKey As String) As String
    Dim filePath As Variant
    Dim fileName As String
    Dim maxTime As String
    Dim fileTime As String
    Dim latestFile As String
    
    maxTime = "000000"
    latestFile = ""
    
    For Each filePath In files
        fileName = fso.GetFileName(filePath)
        fileTime = ExtractTime(fileName)
        
        If fileTime > maxTime Then
            maxTime = fileTime
            latestFile = filePath
        End If
    Next filePath
    
    ' Store the latest file for this town-day and technology
    If techType = "3G" Then
        latest3GFilesByTownDay(townDayKey) = latestFile
    ElseIf techType = "4G" Then
        latest4GFilesByTownDay(townDayKey) = latestFile
    End If
    
    FindLatestFile = latestFile
End Function
Function IsLastFileForTownDay(filePath As String, fso As Object, techType As String) As Boolean
    Dim fileName As String
    Dim town As String
    Dim fileDate As String
    Dim townDayKey As String
    
    ' Extract information from the filename
    fileName = fso.GetFileName(filePath)
    town = ExtractTown(fileName)
    fileDate = ExtractDate(fileName)
    
    ' Create a key for this town and day combination
    townDayKey = town & "_" & fileDate
    
    ' Check if we've processed this town-day combination
    If techType = "3G" Then
        If latest3GFilesByTownDay.Exists(townDayKey) Then
            IsLastFileForTownDay = (filePath = latest3GFilesByTownDay(townDayKey))
        Else
            IsLastFileForTownDay = False
        End If
    ElseIf techType = "4G" Then
        If latest4GFilesByTownDay.Exists(townDayKey) Then
            IsLastFileForTownDay = (filePath = latest4GFilesByTownDay(townDayKey))
        Else
            IsLastFileForTownDay = False
        End If
    End If
End Function
