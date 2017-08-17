Module modFileCleanup
    Dim strFile As String         '-Working File
    Dim strDateCreated  '-Used to get the Last Time a File was modified /Created
    Dim strDateOld      '-Used to get the value of todaydate minus tthe DaysOld Constant
    Dim CurrentFile     '-Working File
    Dim MyFolderList    '-Used to Split the folders in an array
    Dim x               '-Count Folder List array
    Dim strWorkingDir   '-Current Working Directory
    Dim strDateCount    '- To Count files that arex days old.  MOstly used for reporting
    Dim Days2Old        '-Used to Convert the Constant Daysold into a negative
    Dim strFileType     '-Grab the Array from DayOld
    Dim strFileArr      '-File Type Array
    Dim ArrayFiles      '-Array of Files Deleted
    Dim SplitArray
    Dim strNetIQ
    Dim IsEnabled As Boolean
    Public DaysOld As Integer
    Public RootDirectory As String
    Public DoOver As Boolean
    Const FileType = "bak"
    Const DeleteAllFiles = "n" 'DeleteAllFiles without DateCheck
    Const MAXFILESLEFT = 6
    Const SIMULATE = False
    Const DoMSG = False
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
        If FileCount <= MAXFILESLEFT Then Exit Function
        If Not SIMULATE Then strDeletedFile.Delete()
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
        Dim sAns As String = ""
        Dim GetCurrentDate As String
        Dim CurrentMonth As String
        strFile = RootDirectory & strFileCreated
        strDateCreated = ShowFileInfo(strFile)
        GetCurrentDate = FormatDateTime(strDateCreated, 2)
        strDateCreated = GetCurrentDate
        GetCurrentDate = Now.Date
        Days2Old = 0 - DaysOld
        strDateOld = DateAdd("d", Days2Old, GetCurrentDate)
        CurrentMonth = DateDiff("d", strDateCreated, GetCurrentDate)
        If CurrentMonth >= DaysOld Then
            If DeleteOldFile(strFile) Then
                strDateCount = strDateCount + 1
                ArrayFiles = ArrayFiles + "," + strFile
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
                If strFileSplit(1) = strFileType Then
                    CurrentFile = sf
                    Call GetCreatedDate(CurrentFile)
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
            If strFileSplit(1) = strFileType Then i = i + 1
            sf = ""
        Next
        Return i
    End Function
    ''' <summary>
    ''' get the setttings for this app for the app that this application supports
    ''' </summary>
    Private Sub GetSettings()
        Dim BSReg As New BurnSoft.GlobalClasses.BSRegistry
        BSReg.DefaultRegPath = RegKey
        RootDirectory = FormatDirectory(BSReg.GetLastWorkingDir)
        DaysOld = BSReg.GetDaysOld
        IsEnabled = BSReg.UseTracking
    End Sub
    ''' <summary>
    ''' start deleting files for cleanup
    ''' </summary>
    Public Sub DoDelete()
        strFileArr = Split(FileType, ",")
        strFileType = UBound(strFileArr)
        Call GetSettings()
        If Not IsEnabled Then Exit Sub
        strDateCount = 0
        strFileType = FileType
        Call GetFileList()
        If DoMSG Then
            If strDateCount = 0 Then
                MsgBox("0 Files were deleted!")
            Else
                SplitArray = Replace(ArrayFiles, ",", Chr(10) & Chr(13))
                strNetIQ = strDateCount & " Files were deleted!" & Chr(10) & Chr(13) & "The Following Files Where Deleted:" & Chr(10) & Chr(13) & SplitArray
                MsgBox(strNetIQ)
            End If
        End If
    End Sub
End Module
