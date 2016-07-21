Imports System
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Win32
Namespace BurnSoft.GlobalClasses
    Public Class BSRegistry
        Public DefaultRegPath As String
#Region "Registry General Functions and Subs"
        Private Sub CreateSubKey(ByVal strValue As String)
            Microsoft.Win32.Registry.CurrentUser.CreateSubKey(strValue)
        End Sub
        Private Function RegSubKeyExists(ByVal strValue As String) As Boolean
            Dim bAns As Boolean = False
            Try
                Dim MyReg As RegistryKey
                MyReg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(strValue, True)
                If MyReg Is Nothing Then
                    bAns = False
                Else
                    bAns = True
                End If
            Catch ex As Exception
                bAns = False
            End Try
            Return bAns
        End Function
        Private Function GetRegSubKeyValue(ByVal strKey As String, ByVal strValue As String, ByVal strDefault As String) As String
            Dim sAns As String = ""
            Dim strMsg As String = ""
            Dim MyReg As RegistryKey
            Try
                If RegSubKeyExists(strKey) Then
                    MyReg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(strKey, True)
                    If Len(MyReg.GetValue(strValue)) > 0 Then
                        sAns = MyReg.GetValue(strValue)
                    Else
                        MyReg.SetValue(strValue, strDefault)
                        sAns = strDefault
                    End If
                Else
                    Call CreateSubKey(strKey)
                    MyReg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(strKey, True)
                    MyReg.SetValue(strValue, strDefault)
                    sAns = strDefault
                End If
            Catch ex As Exception
                sAns = strDefault
            End Try
            Return sAns
        End Function
        Private Function SettingsExists() As Boolean
            Dim bAns As Boolean = False
            Dim MyReg As RegistryKey
            Dim strValue As String = DefaultRegPath & "\Settings"
            On Error Resume Next
            MyReg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(strValue, True)
            If MyReg Is Nothing Then
                bAns = False
            Else
                bAns = True
            End If
            Return bAns
        End Function
        Public Function GetLastWorkingDir() As String
            Dim sAns As String = ""
            Dim strValue As String = DefaultRegPath & "\Settings"
            sAns = GetRegSubKeyValue(strValue, "LastPath", "")
            Return sAns
        End Function
        Public Function GetDBPath() As String
            Dim sAns As String = ""
            sAns = GetRegSubKeyValue(DefaultRegPath, "Database", DBName)
            Return sAns
        End Function
        Public Function GetDaysOld() As Long
            Dim lAns As Long = 0
            lAns = CLng(GetRegSubKeyValue(DefaultRegPath & "\Settings", "TrackHistoryDays", 30))
            Return lAns
        End Function
        Public Function UseTracking() As Boolean
            Dim bAns As Boolean = True
            bAns = CBool(GetRegSubKeyValue(DefaultRegPath & "\Settings", "TrackHistory", True))
            Return bAns
        End Function
        Public Function GetApplicationPath()
            Dim sAns As String = ""
            sAns = GetRegSubKeyValue(DefaultRegPath, "Path", Application.StartupPath) & "\"
            Return sAns
        End Function
        Public Sub SaveRegSetting(ByVal sKey As String, ByVal sValue As String)
            Dim MyReg As RegistryKey
            Dim strValue As String = DefaultRegPath & "Settings"
            If Not RegSubKeyExists(strValue) Then Call CreateSubKey(strValue)
            MyReg = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(strValue, RegistryKeyPermissionCheck.Default)
            MyReg.SetValue(sKey, sValue)
            MyReg.Close()
        End Sub
#End Region
    End Class
    Public Class BSFileSystem
        Public Sub LogFile(ByVal strPath As String, ByVal strMessage As String)
            Dim SendMessage As String = DateTime.Now & vbTab & strMessage
            Call AppendToFile(strPath, SendMessage)
        End Sub
        Public Sub DeleteFile(ByVal strPath As String)
            If File.Exists(strPath) Then
                File.Delete(strPath)
            End If
        End Sub
        Public Function FileExists(ByVal strPath As String)
            Return File.Exists(strPath)
        End Function
        Private Sub CreateFile(ByVal strPath As String)
            If File.Exists(strPath) = False Then
                Dim fs As New FileStream(strPath, FileMode.Append, FileAccess.Write, FileShare.Write)
                fs.Close()
            End If
        End Sub
        Private Sub AppendToFile(ByVal strPath As String, ByVal strNewLine As String)
            If File.Exists(strPath) = False Then Call CreateFile(strPath)
            Dim sw As New StreamWriter(strPath, True, Encoding.ASCII)
            sw.WriteLine(strNewLine)
            sw.Close()
        End Sub
        Public Sub OutPutToFile(ByVal strPath As String, ByVal strNewLine As String)
            Call AppendToFile(strPath, strNewLine)
        End Sub
        Public Sub MoveFile(ByVal strFrom As String, ByVal strTo As String)
            If File.Exists(strFrom) Then
                File.Move(strFrom, strTo)
            End If
        End Sub
        Public Sub CopyFile(ByVal strFrom As String, ByVal strTo As String)
            If File.Exists(strFrom) Then
                File.Copy(strFrom, strTo)
            End If
        End Sub
        Public Sub CreateDirectory(ByVal strPath As String)
            If Directory.Exists(strPath) Then
                Directory.CreateDirectory(strPath)
            End If
        End Sub
        Public Function DirectoryExists(ByVal strPath As String) As Boolean
            Return Directory.Exists(strPath)
        End Function
        Public Sub DeleteDirectory(ByVal strPath As String)
            If Directory.Exists(strPath) Then
                Directory.Delete(strPath)
            End If
        End Sub
        Public Sub MoveDirectory(ByVal strFrom As String, ByVal strTo As String)
            If Directory.Exists(strFrom) Then
                Directory.Move(strFrom, strTo)
            End If
        End Sub
        Public Sub RenameFile(ByVal strFrom As String, ByVal strTo As String)
            File.Move(strFrom, strTo)
        End Sub
        Public Function GetPathOfFile(ByVal strFile As String) As String
            Dim sAns As String = ""
            sAns = Path.GetDirectoryName(strFile)
            Return sAns
        End Function
        Public Function GetExtOfFile(ByVal strFile As String) As String
            Dim sAns As String = ""
            sAns = Path.GetExtension(strFile)
            Return sAns
        End Function
        Public Function GetNameOfFile(ByVal strFile As String) As String
            Dim sAns As String = ""
            sAns = Path.GetFileName(strFile)
            Return sAns
        End Function
        Public Function FileHasExtension(ByVal strFile As String) As Boolean
            Dim bAns As Boolean = False
            bAns = Path.HasExtension(strFile)
            Return bAns
        End Function
        Public Function GetNameOfFileWOExt(ByVal strFile As String) As String
            Dim sAns As String = ""
            sAns = Path.GetFileNameWithoutExtension(strFile)
            Return sAns
        End Function
    End Class
End Namespace