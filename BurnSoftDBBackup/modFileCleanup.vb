''' <summary>
''' Module File Clean up
''' </summary>
Module ModFileCleanup
    ''' <summary>
    ''' The string file Working File
    ''' </summary>
    Dim _strFile As String
    ''' <summary>
    ''' The string date created, Used to get the Last Time a File was modified /Created
    ''' </summary>
    Dim _strDateCreated
    ''' <summary>
    ''' The string date old Used to get the value of todaydate minus tthe DaysOld Constant
    ''' </summary>
' ReSharper disable once NotAccessedField.Local
    Dim _strDateOld
    ''' <summary>
    ''' The current file
    ''' </summary>
    Dim _currentFile
    ''' <summary>
    ''' The string date count, Current Working Directory
    ''' </summary>
    Dim _strDateCount
    ''' <summary>
    ''' The days2 old, To Count files that arex days old.  MOstly used for reporting
    ''' </summary>
    Dim _days2Old
    ''' <summary>
    ''' The string file type, Grab the Array from DayOld
    ''' </summary>
    Dim _strFileType
    ''' <summary>
    ''' The string file arr, File Type Array
    ''' </summary>
    Dim _strFileArr
    ''' <summary>
    ''' Array of Files Deleted
    ''' </summary>
    Dim _arrayFiles
    ''' <summary>
    ''' The split array
    ''' </summary>
    Dim _splitArray
    ''' <summary>
    ''' The string net iq
    ''' </summary>
    Dim _strNetIq
    ''' <summary>
    ''' The is enabled
    ''' </summary>
    Dim _isEnabled As Boolean
    ''' <summary>
    ''' The days old
    ''' </summary>
    Public DaysOld As Integer
    ''' <summary>
    ''' The root directory
    ''' </summary>
    Public RootDirectory As String
    ''' <summary>
    ''' The do over
    ''' </summary>
    Public DoOver As Boolean
    ''' <summary>
    ''' 
    ''' </summary>
    Const FileType = "bak"
    ''' <summary>
    ''' 
    ''' </summary>
    Const Maxfilesleft = 6
    ''' <summary>
    ''' simulate
    ''' </summary>
    Const Simulate = False
    ''' <summary>
    ''' Do messages
    ''' </summary>
    Const DoMsg = False
    ''' <summary>
    ''' Get the date of last modified form the file that was passed
    ''' </summary>
    ''' <param name="strFile"></param>
    ''' <returns></returns>
    Private Function ShowFileInfo(ByVal strFile)
        Dim fso, f
        fso = CreateObject("Scripting.FileSystemObject")
        f = fso.GetFile(strFile)
        Return FormatDateTime(f.DateLastModified)
    End Function
    ''' <summary>
    ''' delete the the selected file, but check to make sure there are mote than 6 left in the structure
    ''' if there are 5 backups and this was one of those 5, then it will abort the delete
    ''' </summary>
    ''' <param name="strFile"></param>
    ''' <returns></returns>
    Private Function DeleteOldFile(ByVal strFile)
        Dim bAns As Boolean = False
        Dim fso
        Dim strDeletedFile
        fso = CreateObject("Scripting.FileSystemObject")
        strDeletedFile = fso.GetFile(strFile)
        DeleteOldFile = bAns
        If FileCount <= Maxfilesleft Then Exit Function
        If Not Simulate Then strDeletedFile.Delete()
        bAns = True
        Return bAns
    End Function
    ''' <summary>
    ''' get the date created to determin if the file is old enough to delete, then it will
    ''' put it in a string array
    ''' </summary>
    ''' <param name="strFileCreated"></param>
    ''' <returns></returns>
    Private Function GetCreatedDate(ByVal strFileCreated) As String
        Dim sAns As String
        Dim getCurrentDate As String
        Dim currentMonth As String
        _strFile = RootDirectory & strFileCreated
        _strDateCreated = ShowFileInfo(_strFile)
        getCurrentDate = FormatDateTime(_strDateCreated, 2)
        _strDateCreated = getCurrentDate
        getCurrentDate = Now.Date
        _days2Old = 0 - DaysOld
        _strDateOld = DateAdd("d", _days2Old, getCurrentDate)
        currentMonth = DateDiff("d", _strDateCreated, getCurrentDate)
        If currentMonth >= DaysOld Then
            If DeleteOldFile(_strFile) Then
                _strDateCount = _strDateCount + 1
                _arrayFiles = _arrayFiles + "," + _strFile
            End If
        End If
        Return sAns
    End Function
    ''' <summary>
    ''' get the files from the root directory to get the files that should be deleted
    ''' </summary>
    ''' <returns></returns>
    Private Function GetFileList()
        Try
            Dim fsof, fi, flf, sf, fc
            Dim strFileSplit
            sf = ""
            fsof = CreateObject("Scripting.FileSystemObject")
            fi = fsof.GetFolder(RootDirectory)
            fc = fi.Files
            For Each flf In fc
                sf &= flf.Name
                strFileSplit = Split(sf, ".")
                If strFileSplit(1) = _strFileType Then
                    _currentFile = sf
                    Call GetCreatedDate(_currentFile)
                End If
                sf = ""
            Next
        Catch ex As Exception
            Dim intErr As Integer
            intErr = Err.Number
            Call UpdateLog(Err.Number & "::" & ex.Message.ToString & "(" & RootDirectory & ")", "modfilecleanup", "GetFileList")
            Select Case intErr
                Case 76
                    MsgBox("Unable to Find Target Path, Please try again!.")
                    DoOver = True
                Case Else
                    MsgBox("Unknown error has occured. Please check your path to make sure it is vaild!")
                    DoOver = True
            End Select
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' count the files in the rootdirectory
    ''' </summary>
    ''' <returns></returns>
    Private Function FileCount() As Integer
        Dim fsof, fi, flf, sf, fc
        Dim strFileSplit
        Dim i : i = 0
        sf = ""
        fsof = CreateObject("Scripting.FileSystemObject")
        fi = fsof.GetFolder(RootDirectory)
        fc = fi.Files
        For Each flf In fc
            sf = sf & flf.Name
            strFileSplit = Split(sf, ".")
            If strFileSplit(1) = _strFileType Then i = i + 1
            sf = ""
        Next
        Return i
    End Function
    ''' <summary>
    ''' get the setttings for this app for the app that this application supports
    ''' </summary>
    Private Sub GetSettings()
        Dim bsReg As New BurnSoft.GlobalClasses.BsRegistry
        bsReg.DefaultRegPath = RegKey
        RootDirectory = FormatDirectory(bsReg.GetLastWorkingDir)
        DaysOld = bsReg.GetDaysOld
        _isEnabled = bsReg.UseTracking
    End Sub
    ''' <summary>
    ''' start deleting files for cleanup
    ''' </summary>
    Public Sub DoDelete()
        _strFileArr = Split(FileType, ",")
        _strFileType = UBound(_strFileArr)
        Call GetSettings()
        If Not _isEnabled Then Exit Sub
        _strDateCount = 0
        _strFileType = FileType
        Call GetFileList()
        If DoMsg Then
            If _strDateCount = 0 Then
                MsgBox("0 Files were deleted!")
            Else
                _splitArray = Replace(_arrayFiles, ",", Chr(10) & Chr(13))
                _strNetIq = _strDateCount & " Files were deleted!" & Chr(10) & Chr(13) & "The Following Files Where Deleted:" & Chr(10) & Chr(13) & _splitArray
                MsgBox(_strNetIq)
            End If
        End If
    End Sub
End Module
