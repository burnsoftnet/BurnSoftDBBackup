Public Class frmMain
    Const BACKUPMODE = "win"  'win,old
    Sub DoWinBackup()
        Try
            btnBackup.Enabled = False
            btnCancel.Enabled = False
            tmrAutoBack.Enabled = False
            Dim strSource As String = DBLastLoc
            Dim sFile As String = NewFileName()
            Dim sDestFile As String = FormatDirectory(lblPath.Text) & sFile
            My.Computer.FileSystem.CopyFile(strSource, sDestFile, FileIO.UIOption.AllDialogs, FileIO.UICancelOption.ThrowException)
            Dim Obj As New BurnSoft.GlobalClasses.BSRegistry
            Obj.DefaultRegPath = RegKey
            Obj.SaveRegSetting("LastPath", lblPath.Text)
            Obj.SaveRegSetting("LastFile", sFile)
            Obj.SaveRegSetting("Successful", Now)
            'ProgressBar1.Visible = False
            If Not DoAutoBackup Then MsgBox("Backup Completed Successfully!" & Chr(10) & "Backup File Name: " & sFile, MsgBoxStyle.OkOnly)
            Global.System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            Dim ObjFS As New BurnSoft.GlobalClasses.BSFileSystem
            Dim strform As String = "frmMain"
            Dim strProcedure As String = "DoWinBackup"
            Dim sMessage As String = strform & "." & strProcedure & "::" & Err.Number & "::" & ex.Message.ToString()
            ObjFS.LogFile(MyLogFile, sMessage)
            Dim mAns As String
            Select Case Err.Number
                Case 76
                    mAns = MsgBox("Unable to find Destination " & lblPath.Text, MsgBoxStyle.AbortRetryIgnore)
                    Select Case mAns
                        Case vbRetry
                            Call DoBackup()
                        Case vbAbort
                            Me.Close()
                        Case vbIgnore
                            Call NewPath()
                            Call DoWinBackup()
                    End Select
                Case 53
                    mAns = MsgBox("Unable to find Source Database " & DBLastLoc, MsgBoxStyle.RetryCancel)
                    If mAns = vbRetry Then Call DoBackup()
                    If mAns = vbCancel Then Me.Close()
                Case 57
                    mAns = MsgBox("Destination File is currently in Use.", MsgBoxStyle.RetryCancel)
                    If mAns = vbRetry Then Call DoBackup()
                    If mAns = vbCancel Then Me.Close()
                Case 5
                    MsgBox("Operation Canceled per your request.")
                Case Else
                    MsgBox("An error occurred while backing up your database.  Please try again!")
            End Select
            btnBackup.Enabled = True
            btnCancel.Enabled = True
        End Try
    End Sub
    Sub DoBackup()
        Try
            tmrAutoBack.Enabled = False
            ProgressBar1.Visible = True
            Dim strSource As String = DBLastLoc
            Dim sFile As String = NewFileName()
            Dim sDestFile As String = FormatDirectory(lblPath.Text) & sFile
            Dim fil As New IO.FileInfo(strSource)
            Dim strmIn As IO.FileStream = fil.OpenRead
            Dim strmOut As IO.FileStream = IO.File.Create(sDestFile)
            ProgressBar1.Maximum = strmIn.Length
            ProgressBar1.Minimum = 0
            Dim iValue As Long = strmIn.Position
            Dim MyI As Long = 0

            Do Until strmOut.Length = strmIn.Length
                strmOut.WriteByte(strmIn.ReadByte)
                iValue = strmIn.Position
                ProgressBar1.Value = iValue
                ProgressBar1.Refresh()
                Me.Refresh()
            Loop
            strmOut.Close()
            strmIn.Close()
            strmOut.Dispose()
            strmIn.Dispose()
            fil = Nothing
            Dim Obj As New BurnSoft.GlobalClasses.BSRegistry
            Obj.DefaultRegPath = RegKey
            Obj.SaveRegSetting("LastPath", lblPath.Text)
            Obj.SaveRegSetting("LastFile", sFile)
            Obj.SaveRegSetting("Successful", Now)
            ProgressBar1.Visible = False
            If Not DoAutoBackup Then MsgBox("Backup Completed Successfully!", MsgBoxStyle.OkOnly)
            Global.System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            Dim ObjFS As New BurnSoft.GlobalClasses.BSFileSystem
            Dim strform As String = "frmMain"
            Dim strProcedure As String = "DoBackup"
            Dim sMessage As String = strform & "." & strProcedure & "::" & Err.Number & "::" & ex.Message.ToString()
            ObjFS.LogFile(MyLogFile, sMessage)
            Dim mAns As String
            Select Case Err.Number
                Case 76
                    mAns = MsgBox("Unable to find Destination " & lblPath.Text, MsgBoxStyle.AbortRetryIgnore)
                    Select Case mAns
                        Case vbRetry
                            Call DoBackup()
                        Case vbAbort
                            Me.Close()
                        Case vbIgnore
                            Call NewPath()
                            Call DoBackup()
                    End Select
                Case 53
                    mAns = MsgBox("Unable to find Source Database " & DBLastLoc, MsgBoxStyle.RetryCancel)
                    If mAns = vbRetry Then Call DoBackup()
                    If mAns = vbCancel Then Me.Close()
                Case 57
                    mAns = MsgBox("Destination File is currently in Use.", MsgBoxStyle.RetryCancel)
                    If mAns = vbRetry Then Call DoBackup()
                    If mAns = vbCancel Then Me.Close()
                Case Else
                    MsgBox("An error occurred while backing up your database.  Please try again!")
            End Select
        End Try
    End Sub
    Sub NewPath()
        FolderBrowserDialog1.ShowDialog()
        If Len(FolderBrowserDialog1.SelectedPath) > 0 Then
            lblPath.Text = FormatDirectory(FolderBrowserDialog1.SelectedPath)
            btnBackup.Enabled = True
        End If
    End Sub
    Private Sub btnPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPath.Click
        Call NewPath()
    End Sub
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call SetINIT()
        Dim Obj As New BurnSoft.GlobalClasses.BSRegistry
        Obj.DefaultRegPath = RegKey
        Me.Text = "BurnSoft " & MainAppName & "DB Backup"
        lblPath.Text = Obj.GetLastWorkingDir
        If Len(lblPath.Text) > 0 Then
            btnBackup.Enabled = True
        Else
            btnBackup.Enabled = False
        End If
        Call DoDelete()
        tmrAutoBack.Enabled = DoAutoBackup
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Global.System.Windows.Forms.Application.Exit()
    End Sub
    Sub RunBackers()
        Select Case BACKUPMODE
            Case "win"
                Call DoWinBackup()
            Case "old"
                Call DoBackup()
            Case Else
                Call DoWinBackup()
        End Select
    End Sub
    Private Sub btnBackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBackup.Click
        Call RunBackers()
    End Sub
    Private Sub tmrAutoBack_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrAutoBack.Tick
        btnBackup.Enabled = False
        Call RunBackers()
        tmrAutoBack.Enabled = False
        Global.System.Windows.Forms.Application.Exit()
    End Sub
End Class
