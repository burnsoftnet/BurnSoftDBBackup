Imports System.IO
Imports System.Text
Imports Microsoft.Win32
''' <summary>
''' The GlobalClasses namespace.
''' </summary>
' ReSharper disable once CheckNamespace
Namespace BurnSoft.GlobalClasses
    ''' <summary>
    ''' Class BSRegistry.
    ''' </summary>
    Public Class BsRegistry
        ''' <summary>
        ''' The default reg path
        ''' </summary>
        Public DefaultRegPath As String
#Region "Registry General Functions and Subs"
        ''' <summary>
        ''' Creates the sub key.
        ''' </summary>
        ''' <param name="strValue">The string value.</param>
        Private Sub CreateSubKey(ByVal strValue As String)
            Registry.CurrentUser.CreateSubKey(strValue)
        End Sub
        ''' <summary>
        ''' Regs the sub key exists.
        ''' </summary>
        ''' <param name="strValue">The string value.</param>
        ''' <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        Private Function RegSubKeyExists(ByVal strValue As String) As Boolean
            Dim bAns As Boolean
            Try
                Dim myReg As RegistryKey
                myReg = Registry.CurrentUser.OpenSubKey(strValue, True)
                If myReg Is Nothing Then
                    bAns = False
                Else
                    bAns = True
                End If
            Catch ex As Exception
                bAns = False
            End Try
            Return bAns
        End Function
        ''' <summary>
        ''' Gets the reg sub key value.
        ''' </summary>
        ''' <param name="strKey">The string key.</param>
        ''' <param name="strValue">The string value.</param>
        ''' <param name="strDefault">The string default.</param>
        ''' <returns>System.String.</returns>
        Private Function GetRegSubKeyValue(ByVal strKey As String, ByVal strValue As String, ByVal strDefault As String) As String
            Dim sAns As String
            Dim myReg As RegistryKey
            Try
                If RegSubKeyExists(strKey) Then
                    myReg = Registry.CurrentUser.OpenSubKey(strKey, True)
                    If Len(myReg.GetValue(strValue)) > 0 Then
                        sAns = myReg.GetValue(strValue)
                    Else
                        myReg.SetValue(strValue, strDefault)
                        sAns = strDefault
                    End If
                Else
                    Call CreateSubKey(strKey)
                    myReg = Registry.CurrentUser.OpenSubKey(strKey, True)
                    myReg.SetValue(strValue, strDefault)
                    sAns = strDefault
                End If
            Catch ex As Exception
                sAns = strDefault
            End Try
            Return sAns
        End Function
        ''' <summary>
        ''' Settingses the exists.
        ''' </summary>
        ''' <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
' ReSharper disable once UnusedMember.Local
        Private Function SettingsExists() As Boolean
            Dim bAns As Boolean
            Dim myReg As RegistryKey
            Dim strValue As String = DefaultRegPath & "\Settings"
            On Error Resume Next
            myReg = Registry.CurrentUser.OpenSubKey(strValue, True)
            If myReg Is Nothing Then
                bAns = False
            Else
                bAns = True
            End If
            Return bAns
        End Function
        ''' <summary>
        ''' Gets the last working dir.
        ''' </summary>
        ''' <returns>System.String.</returns>
        Public Function GetLastWorkingDir() As String
            Dim sAns As String
            Dim strValue As String = DefaultRegPath & "\Settings"
            sAns = GetRegSubKeyValue(strValue, "LastPath", "")
            Return sAns
        End Function
        ''' <summary>
        ''' Gets the database path.
        ''' </summary>
        ''' <returns>System.String.</returns>
        Public Function GetDbPath() As String
            Dim sAns As String
            sAns = GetRegSubKeyValue(DefaultRegPath, "Database", DBName)
            Return sAns
        End Function
        ''' <summary>
        ''' Gets the days old.
        ''' </summary>
        ''' <returns>System.Int64.</returns>
        Public Function GetDaysOld() As Long
            Dim lAns As Long
            lAns = CLng(GetRegSubKeyValue(DefaultRegPath & "\Settings", "TrackHistoryDays", 30))
            Return lAns
        End Function
        ''' <summary>
        ''' Uses the tracking.
        ''' </summary>
        ''' <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        Public Function UseTracking() As Boolean
            Dim bAns As Boolean
            bAns = CBool(GetRegSubKeyValue(DefaultRegPath & "\Settings", "TrackHistory", True))
            Return bAns
        End Function
        ''' <summary>
        ''' Gets the application path.
        ''' </summary>
        ''' <returns>System.Object.</returns>
        Public Function GetApplicationPath()
            Dim sAns As String
            sAns = GetRegSubKeyValue(DefaultRegPath, "Path", Application.StartupPath) & "\"
            Return sAns
        End Function
        ''' <summary>
        ''' Saves the reg setting.
        ''' </summary>
        ''' <param name="sKey">The s key.</param>
        ''' <param name="sValue">The s value.</param>
        Public Sub SaveRegSetting(ByVal sKey As String, ByVal sValue As String)
            Dim myReg As RegistryKey
            Dim strValue As String = DefaultRegPath & "Settings"
            If Not RegSubKeyExists(strValue) Then Call CreateSubKey(strValue)
            myReg = Registry.CurrentUser.CreateSubKey(strValue, RegistryKeyPermissionCheck.Default)
            myReg.SetValue(sKey, sValue)
            myReg.Close()
        End Sub
#End Region
    End Class
    ''' <summary>
    ''' Class BSFileSystem.
    ''' </summary>
    Public Class BsFileSystem
        ''' <summary>
        ''' Logs the file.
        ''' </summary>
        ''' <param name="strPath">The string path.</param>
        ''' <param name="strMessage">The string message.</param>
        Public Sub LogFile(ByVal strPath As String, ByVal strMessage As String)
            Dim sendMessage As String = DateTime.Now & vbTab & strMessage
            Call AppendToFile(strPath, sendMessage)
        End Sub
        ''' <summary>
        ''' Deletes the file.
        ''' </summary>
        ''' <param name="strPath">The string path.</param>
        Public Sub DeleteFile(ByVal strPath As String)
            If File.Exists(strPath) Then
                File.Delete(strPath)
            End If
        End Sub
        ''' <summary>
        ''' Files the exists.
        ''' </summary>
        ''' <param name="strPath">The string path.</param>
        ''' <returns>System.Object.</returns>
        Public Function FileExists(ByVal strPath As String)
            Return File.Exists(strPath)
        End Function
        Private Sub CreateFile(ByVal strPath As String)
            If File.Exists(strPath) = False Then
                Dim fs As New FileStream(strPath, FileMode.Append, FileAccess.Write, FileShare.Write)
                fs.Close()
            End If
        End Sub
        ''' <summary>
        ''' Appends to file.
        ''' </summary>
        ''' <param name="strPath">The string path.</param>
        ''' <param name="strNewLine">The string new line.</param>
        Private Sub AppendToFile(ByVal strPath As String, ByVal strNewLine As String)
            If File.Exists(strPath) = False Then Call CreateFile(strPath)
            Dim sw As New StreamWriter(strPath, True, Encoding.ASCII)
            sw.WriteLine(strNewLine)
            sw.Close()
        End Sub
        ''' <summary>
        ''' Outs the put to file.
        ''' </summary>
        ''' <param name="strPath">The string path.</param>
        ''' <param name="strNewLine">The string new line.</param>
        Public Sub OutPutToFile(ByVal strPath As String, ByVal strNewLine As String)
            Call AppendToFile(strPath, strNewLine)
        End Sub
        ''' <summary>
        ''' Moves the file.
        ''' </summary>
        ''' <param name="strFrom">The string from.</param>
        ''' <param name="strTo">The string to.</param>
        Public Sub MoveFile(ByVal strFrom As String, ByVal strTo As String)
            If File.Exists(strFrom) Then
                File.Move(strFrom, strTo)
            End If
        End Sub
        ''' <summary>
        ''' Copies the file.
        ''' </summary>
        ''' <param name="strFrom">The string from.</param>
        ''' <param name="strTo">The string to.</param>
        Public Sub CopyFile(ByVal strFrom As String, ByVal strTo As String)
            If File.Exists(strFrom) Then
                File.Copy(strFrom, strTo)
            End If
        End Sub
        ''' <summary>
        ''' Creates the directory.
        ''' </summary>
        ''' <param name="strPath">The string path.</param>
        Public Sub CreateDirectory(ByVal strPath As String)
            If Directory.Exists(strPath) Then
                Directory.CreateDirectory(strPath)
            End If
        End Sub
        ''' <summary>
        ''' Directories the exists.
        ''' </summary>
        ''' <param name="strPath">The string path.</param>
        ''' <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        Public Function DirectoryExists(ByVal strPath As String) As Boolean
            Return Directory.Exists(strPath)
        End Function
        ''' <summary>
        ''' Deletes the directory.
        ''' </summary>
        ''' <param name="strPath">The string path.</param>
        Public Sub DeleteDirectory(ByVal strPath As String)
            If Directory.Exists(strPath) Then
                Directory.Delete(strPath)
            End If
        End Sub
        ''' <summary>
        ''' Moves the directory.
        ''' </summary>
        ''' <param name="strFrom">The string from.</param>
        ''' <param name="strTo">The string to.</param>
        Public Sub MoveDirectory(ByVal strFrom As String, ByVal strTo As String)
            If Directory.Exists(strFrom) Then
                Directory.Move(strFrom, strTo)
            End If
        End Sub
        ''' <summary>
        ''' Renames the file.
        ''' </summary>
        ''' <param name="strFrom">The string from.</param>
        ''' <param name="strTo">The string to.</param>
        Public Sub RenameFile(ByVal strFrom As String, ByVal strTo As String)
            File.Move(strFrom, strTo)
        End Sub
        ''' <summary>
        ''' Gets the path of file.
        ''' </summary>
        ''' <param name="strFile">The string file.</param>
        ''' <returns>System.String.</returns>
        Public Function GetPathOfFile(ByVal strFile As String) As String
            Dim sAns As String
            sAns = Path.GetDirectoryName(strFile)
            Return sAns
        End Function
        ''' <summary>
        ''' Gets the ext of file.
        ''' </summary>
        ''' <param name="strFile">The string file.</param>
        ''' <returns>System.String.</returns>
        Public Function GetExtOfFile(ByVal strFile As String) As String
            Dim sAns As String
            sAns = Path.GetExtension(strFile)
            Return sAns
        End Function
        ''' <summary>
        ''' Gets the name of file.
        ''' </summary>
        ''' <param name="strFile">The string file.</param>
        ''' <returns>System.String.</returns>
        Public Function GetNameOfFile(ByVal strFile As String) As String
            Dim sAns As String
            sAns = Path.GetFileName(strFile)
            Return sAns
        End Function
        ''' <summary>
        ''' Files the has extension.
        ''' </summary>
        ''' <param name="strFile">The string file.</param>
        ''' <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        Public Function FileHasExtension(ByVal strFile As String) As Boolean
            Dim bAns As Boolean
            bAns = Path.HasExtension(strFile)
            Return bAns
        End Function
        ''' <summary>
        ''' Gets the name of file wo ext.
        ''' </summary>
        ''' <param name="strFile">The string file.</param>
        ''' <returns>System.String.</returns>
        Public Function GetNameOfFileWoExt(ByVal strFile As String) As String
            Dim sAns As String
            sAns = Path.GetFileNameWithoutExtension(strFile)
            Return sAns
        End Function
    End Class
End Namespace