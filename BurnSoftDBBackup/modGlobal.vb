Imports System.Configuration
Module modGlobal
    Public AppName As String
    Public MainAppName As String
    Public DBName As String
    Public RegKey As String
    Public CheckProcess As Boolean
    Public MyLogFile As String
    Public AppPath As String
    Public AppABV As String
    Public DBLastLoc As String
    Public DoAutoBackup As Boolean
    Public Sub SetINIT()
        AppName = System.Configuration.ConfigurationManager.AppSettings("AppName")
        MainAppName = System.Configuration.ConfigurationManager.AppSettings("MainAppName")
        DBName = System.Configuration.ConfigurationManager.AppSettings("DBName")
        RegKey = System.Configuration.ConfigurationManager.AppSettings("RegKey")
        CheckProcess = CBool(System.Configuration.ConfigurationManager.AppSettings("CheckProcess"))
        MyLogFile = System.Configuration.ConfigurationManager.AppSettings("LogFilename")
        AppABV = System.Configuration.ConfigurationManager.AppSettings("AppABV")
        DoAutoBackup = GetCommand("auto", "Bool")
        Dim Obj As New BurnSoft.GlobalClasses.BSRegistry
        Obj.DefaultRegPath = RegKey
        DBLastLoc = Obj.GetDBPath
    End Sub
    Public Function GetCommand(ByVal strLookFor As String, ByVal strType As String, Optional ByRef DidExist As Boolean = False) As String
        Dim sAns As String = ""
        DidExist = False
        Select Case LCase(strType)
            Case "string"
                sAns = ""
            Case "bool"
                sAns = "false"
            Case "int"
                sAns = 0
            Case Else
                sAns = ""
        End Select
        Dim cmdLine() As String = System.Environment.GetCommandLineArgs
        Dim i As Integer = 0
        Dim intCount As Integer = cmdLine.Length
        Dim strValue As String = ""
        If intCount > 1 Then
            For i = 1 To intCount - 1
                strValue = cmdLine(i)
                strValue = Replace(strValue, "/", "")
                strValue = Replace(strValue, "--", "")
                Dim strSplit() As String = Split(strValue, "=")
                Dim intLBound As Integer = LBound(strSplit)
                Dim intUBound As Integer = UBound(strSplit)
                If LCase(strSplit(intLBound)) = LCase(strLookFor) Then
                    If intUBound <> 0 Then
                        sAns = strSplit(intUBound)
                    Else
                        sAns = "true"
                    End If
                    DidExist = True
                    Exit For
                End If
            Next
        End If
        Return LCase(sAns)
    End Function

    Public Function NewFileName() As String
        Dim sAns As String
        sAns = AppABV & "_" & Format(Now(), "yyyyMMdd") & "_" & Format(Now(), "hmmss") & ".bak"
        Return sAns
    End Function
    Public Sub UpdateLog(ByVal sMsg As String, ByVal sForm As String, ByVal sPro As String)
        Dim Obj As New BurnSoft.GlobalClasses.BSFileSystem
        Dim Msg As String = sMsg & "::" & sForm & "." & sPro
        Call Obj.LogFile(MyLogFile, Msg)
    End Sub
    Public Function FormatDirectory(ByVal sValue As String) As String
        Dim sAns As String = sValue
        If Len(sValue) <> 0 Then
            Dim iCount As Integer = Len(sValue)
            Dim LastChar As String = Mid(sValue, iCount)
            If LastChar <> "\" Then
                sAns = sValue & "\"
            Else
                sAns = sValue
            End If
        End If
        Return sAns
    End Function
End Module
