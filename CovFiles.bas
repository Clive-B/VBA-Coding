Attribute VB_Name = "CovFiles"
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Module-level dictionaries to track latest files by town and day
Private latest3GFilesByTownDay As Object
Private latest4GFilesByTownDay As Object
Public NameTown, NetworkCurr, keepOne, GetDay As String
'Public TownName As String
Public Function GetPublicVari() As String
    GetPublicVari = NameTown
End Function
Public Function GetNetwork() As String
    GetNetwork = NetworkCurr
End Function
Public Function GetKeepOne() As String
    GetKeepOne = keepOne
End Function
Public Function GetDayy() As String
    GetDayy = GetDay
End Function
Sub ProcessCovFiles()
    Dim fso As Object, folder As Object, file As Object
    Dim townDateDict As Object, locDict As Object
    Dim townDateKey As Variant, locKey As Variant
    Dim element As Variant, fileItem As Variant
    Dim workbookPath As String, GetDesktop As String, folderPath As String
    Dim oWSHShell As Object, Trends As String
    Dim KPI As Object, Keep() As String, ArryLine As String, CurrentLine() As String
    Dim Runn() As String
    Dim Slow() As String
    Dim T3GFiles As Collection, T4GFiles As Collection
    Dim isLastElement As Boolean, hasDuplicates As Boolean
    Dim myArray As Variant, loc As String, lastLocFile As String
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
    Dim wb As Workbook
    Dim openedWorkbook As Workbook
    Dim techType As String

    ' Initialize file tracking
    InitializeFileTracking
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop") & "\"
    Set oWSHShell = Nothing
    Trends = Sheets("Option").Range("AO1").Value
    myArray = Array("MTN", "TELECEL", "AT")
    
    For Each element In myArray
        isLastElement = (element = myArray(UBound(myArray)))
        folderPath = GetDesktop & "QoS Automation\" & Trends & "\Telcos\" & element & "\COVERAGE\LOGS"
        workbookPath = GetDesktop & "QoS Automation\Templates\RSCP Template.xlsm"
        
        NetworkCurr = element
        If Not fso.FolderExists(folderPath) Then GoTo NextElement

        Set folder = fso.GetFolder(folderPath)
        Set townDateDict = CreateObject("Scripting.Dictionary")
        Set fileList = New Collection
        
        ' First, collect all files in order
        For Each file In folder.files
            fileList.Add file.Path
        Next file
        
        NextworkbookPath
        Sleep 3000
        Set openedWorkbook = Workbooks.Open(workbookPath)

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

        For passNumber = 1 To 1
            For Each townDateKey In townDateDict.Keys
                Set T3GFiles = New Collection
                Set T4GFiles = New Collection
                hasDuplicates = (townDateDict(townDateKey).Count > 1)
                
                ' Extract town and date from the key
                Dim townParts() As String
                townParts = Split(townDateKey, "_")
                town = townParts(0)
                fileDate = townParts(1)
                townDayKey = town & "_" & fileDate
                
                ' Get all locations for this town-date
                locations = townDateDict(townDateKey).Keys
                locCount = townDateDict(townDateKey).Count
                
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
                        techType = GetCoverageTech(filePath, fso)

                        If techType <> "" Then
                            If techType = "3G" Then
                                keepOne = techType
                                T3GFiles.Add filePath
                                If Not fileDict.Exists("3G") Then
                                    fileDict.Add "3G", New Collection
                                End If
                                fileDict("3G").Add filePath
                            ElseIf techType = "4G" Then
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
                            workbookPath = openedWorkbook.FullName
                            Runn = Split(CStr(locKey), "CST")
                            Slow = Split(Runn(1), ".")
                            GetDay = Slow(0)
'                            TownName = Split(AA, " ")
                            FinalTime3G = IsLastFileForTownDay(filePath, fso, "3G")
                            NameTown = town
                            ' Process each file in its original order
                            ProcessFile filePath, Trends, workbookPath, False, _
                                      isLastLocationInTownDate, isLastElement, hasDuplicates, _
                                      isLastInLocation, FinalTime3G, passNumber, townDayKey, CStr(locKey)
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
'        Workbooks.Open workbookPath
        Set folder = Nothing
        Set townDateDict = Nothing
        Set fileList = Nothing
NextElement:
    Next element
    
    On Error Resume Next ' Handle errors if workbook isn't found
    Set wb = Workbooks(GetFileName(workbookPath))
    On Error GoTo 0
    
    ' If workbook found, close it without saving changes
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=True
    End If
    Call Process4GCovFiles
    Set fso = Nothing
End Sub
Sub Process4GCovFiles()

    Dim fso As Object, folder As Object, file As Object
    Dim townDateDict As Object, locDict As Object
    Dim townDateKey As Variant, locKey As Variant
    Dim element As Variant, fileItem As Variant
    Dim workbookPath As String, GetDesktop As String, folderPath As String
    Dim oWSHShell As Object, Trends As String
    Dim KPI As Object, Keep() As String, ArryLine As String, CurrentLine() As String
    Dim Runn() As String
    Dim Slow() As String
    Dim T3GFiles As Collection, T4GFiles As Collection
    Dim isLastElement As Boolean, hasDuplicates As Boolean
    Dim myArray As Variant, loc As String, lastLocFile As String
    Dim isLastInLocation As Boolean, filePath As String, filePat As String
    Dim passNumber As Integer, locCount As Long, currentLocIndex As Long
    Dim isLastLocationInTownDate As Boolean
    Dim locations As Variant
    Dim key As String, town As String, fileDate As String
    Dim locationName As String
    Dim FinalTime3G As Boolean, FinalTime4G As Boolean
    Dim fileDict As Object
    Dim townDayKey As String
    Dim Trance As String
    Dim fileList As Collection
    Dim openedWorkbook As Workbook
    Dim techType As String

    ' Initialize file tracking
    InitializeFileTracking
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop") & "\"
    Set oWSHShell = Nothing
    Trends = Sheets("Option").Range("AO1").Value
    myArray = Array("MTN", "TELECEL")
    
    For Each element In myArray
        isLastElement = (element = myArray(UBound(myArray)))
        folderPath = GetDesktop & "QoS Automation\" & Trends & "\Telcos\" & element & "\COVERAGE\LOGS"
        workbookPath = GetDesktop & "QoS Automation\Templates\RSRP Template.xlsm"
        
        NetworkCurr = element
        If Not fso.FolderExists(folderPath) Then GoTo Next4GElement

        Set folder = fso.GetFolder(folderPath)
        Set townDateDict = CreateObject("Scripting.Dictionary")
        Set fileList = New Collection
        
        ' First, collect all files in order
        For Each file In folder.files
            fileList.Add file.Path
        Next file
        
        NextworkbookPath
        Sleep 3000
        Set openedWorkbook = Workbooks.Open(workbookPath)

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

        For passNumber = 1 To 1
            For Each townDateKey In townDateDict.Keys
                Set T3GFiles = New Collection
                Set T4GFiles = New Collection
                hasDuplicates = (townDateDict(townDateKey).Count > 1)
                
                ' Extract town and date from the key
                Dim townParts() As String
                townParts = Split(townDateKey, "_")
                town = townParts(0)
                fileDate = townParts(1)
                townDayKey = town & "_" & fileDate
                
                ' Get all locations for this town-date
                locations = townDateDict(townDateKey).Keys
                locCount = townDateDict(townDateKey).Count
                
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
                        techType = GetCoverageTech(filePath, fso)

                        If techType <> "" Then
                            If techType = "3G" Then
                                T3GFiles.Add filePath
                                If Not fileDict.Exists("3G") Then
                                    fileDict.Add "3G", New Collection
                                End If
                                fileDict("3G").Add filePath
                            ElseIf techType = "4G" Then
                                keepOne = techType
                                Trance = ExtractFileName(filePath)
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
                            Runn = Split(CStr(locKey), "CST")
                            Slow = Split(Runn(1), ".")
                            GetDay = Slow(0)
                            FinalTime3G = IsLastFileForTownDay(filePath, fso, "3G")
                            NameTown = town
                            ' Process each file in its original order
'                            ProcessFile filePath, Trends, workbookPath, False, _
'                                      isLastLocationInTownDate, isLastElement, hasDuplicates, _
'                                      isLastInLocation, FinalTime3G, passNumber, townDayKey, CStr(locKey)
                        Next fileItem
                    End If

                    ' Process 4G files for location in their natural order
                    If Not fileDict Is Nothing And fileDict.Exists("4G") Then
                        ' Find the latest file without sorting
                        lastLocFile = FindLatestFile(fileDict("4G"), fso, "4G", townDayKey)
                        
                        For Each fileItem In fileDict("4G")
                            filePath = fileItem
                            isLastInLocation = (filePath = lastLocFile)
                            workbookPath = openedWorkbook.FullName
                            Runn = Split(CStr(locKey), "CST")
                            Slow = Split(Runn(1), ".")
                            GetDay = Slow(0)
                            FinalTime4G = IsLastFileForTownDay(filePath, fso, "4G")
                            NameTown = town
                            ' Process each file in its original order
                            Process4GFile filePath, Trends, workbookPath, False, _
                                      isLastLocationInTownDate, isLastElement, hasDuplicates, _
                                      isLastInLocation, FinalTime4G, passNumber, townDayKey, CStr(locKey)
                        Next fileItem
                    End If
                    
                    Set fileDict = Nothing
                Next locKey
            Next townDateKey
        Next passNumber
        
        ' Cleanup between elements
        NextworkbookPath
        Sleep 3000
'        Workbooks.Open workbookPath
        Set folder = Nothing
        Set townDateDict = Nothing
        Set fileList = Nothing
Next4GElement:
    Next element

    Set fso = Nothing
End Sub
Sub ProcessFile( _
    filePath As String, _
    Trends As String, _
    workbookPath As String, _
    Finish As Boolean, _
    isLastLocationInTownDate As Boolean, _
    isLastElement As Boolean, _
    hasDuplicates As Boolean, _
    isLastInLocation As Boolean, _
    FinalTime As Boolean, _
    passNumber As Integer, _
    townDayKey As Variant, _
    ByVal location As String)
    
    Dim macroName As String
    Dim Fil As String
    Dim DirPath As String
    macroName = "ThreeGCov"
    Fil = filePath
    DirPath = Trends

    ' Corrected parameter order: townDayKey before location
    Application.Run "'" & workbookPath & "'!" & macroName, _
        Fil, DirPath, Finish, isLastLocationInTownDate, isLastElement, isLastInLocation, FinalTime, passNumber, townDayKey, location
End Sub
Sub Process4GFile( _
    filePath As String, _
    Trends As String, _
    workbookPath As String, _
    Finish As Boolean, _
    isLastLocationInTownDate As Boolean, _
    isLastElement As Boolean, _
    hasDuplicates As Boolean, _
    isLastInLocation As Boolean, _
    FinalTime As Boolean, _
    passNumber As Integer, _
    townDayKey As Variant, _
    ByVal location As String)
    
    Dim macroName As String
    Dim Fil As String
    Dim DirPath As String
    macroName = "FourGCov"
    Fil = filePath
    DirPath = Trends
    
    Application.Run "'" & workbookPath & "'!" & macroName, _
    Fil, DirPath, Finish, isLastLocationInTownDate, isLastElement, isLastInLocation, FinalTime, passNumber, townDayKey, location
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

Private Function GetCoverageTech(ByVal filePath As String, ByVal fso As Object) As String
    Dim KPI As Object
    Dim ArryLine As String
    Dim CurrentLine() As String
    Dim Keep() As String

    On Error GoTo CleanUp
    Set KPI = fso.OpenTextFile(filePath)
    Do Until KPI.AtEndOfStream
        ArryLine = KPI.ReadLine
        CurrentLine = Split(ArryLine, ",")
        If UBound(CurrentLine) >= 3 Then
            If CurrentLine(0) = "#DL" Then
                Keep = Split(CurrentLine(3), " ")
                If UBound(Keep) >= 1 Then
                    GetCoverageTech = UCase$(Replace(Keep(1), """", ""))
                    Exit Do
                End If
            End If
        End If
    Loop

CleanUp:
    On Error Resume Next
    If Not KPI Is Nothing Then KPI.Close
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
            FindworkbookPath = wb.FullName
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

Function ExtractFileName(filePath As String) As String
    Dim i As Long
    Dim fileName As String
    
    ' Replace any forward slashes with backslashes for consistency
    filePath = Replace(filePath, "/", "\")
    
    ' Find the last backslash position
    i = InStrRev(filePath, "\")
    
    ' Extract everything after the last backslash
    If i > 0 Then
        fileName = Mid(filePath, i + 1)
    Else
        fileName = filePath ' No path separators found
    End If
    
    ExtractFileName = fileName
End Function
' Helper function to extract filename from path
Function GetFileName(fullPath As String) As String
    GetFileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
End Function
