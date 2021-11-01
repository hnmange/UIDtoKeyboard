Imports System.IO
Imports System.Timers
Imports Microsoft.Win32
Imports PCSC
Imports PCSC.Iso7816
Imports PCSC.Monitoring

Public Class frmMain

    Private Shared ReadOnly _contextFactory As IContextFactory = ContextFactory.Instance
    Private _hContext As ISCardContext
    Dim readerName As String
    Dim readingMode As String
    Dim isstart As Boolean = False
    Dim logFile As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "UIDToKeyboard", "Logs", "Log_" + DateTime.Now.ToString("yyyy_MM_dd") + ".txt")
    Shared _timer As System.Timers.Timer
    Function loadReaderList()
        Dim readerList As String()
        Try
            cbxReaderList.DataSource = Nothing

            _hContext = _contextFactory.Establish(SCardScope.System)
            readerList = _hContext.GetReaders()
            _hContext.Release()

            If readerList.Length > 0 Then
                cbxReaderList.DataSource = readerList
                'Else
                '    MessageBox.Show("No card reader detected!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If

            Return True
        Catch ex As Exceptions.PCSCException
            'MessageBox.Show("Error: getReaderList() : " & ex.Message & " (" & ex.SCardError.ToString() & ")")
            LogError("Error:" + ex.Message)
            Return False
        End Try
    End Function

    Dim monitor

    Private Sub startMonitor()
        Try
            Dim monitorFactory As MonitorFactory = MonitorFactory.Instance
            monitor = monitorFactory.Create(SCardScope.System)
            AttachToAllEvents(monitor)
            monitor.Start(cbxReaderList.Text)

            readerName = cbxReaderList.Text
            readingMode = txtReadingMode.Text
            isstart = True
        Catch ex As Exception
            LogError("Error:" + ex.Message)
        End Try

    End Sub

    Private Sub AttachToAllEvents(monitor As ISCardMonitor)
        AddHandler monitor.CardInserted, AddressOf cardInit
    End Sub

    Sub cardInit(eventName As SCardMonitor, unknown As CardStatusEventArgs)
        If readingMode = 1 OrElse readingMode = 2 Then
            SendUID4Byte()
        ElseIf readingMode = 3 OrElse readingMode = 4 Then
            SendUID7Byte()
        End If
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtReadingMode.Text = 1
        Start()
        AddApplicationToStartup()
        ToolStripMenuItem1.Checked = True
    End Sub

    Private Sub btnRefreshReader_Click(sender As Object, e As EventArgs) Handles btnRefreshReader.Click
        loadReaderList()
    End Sub

    Private Sub btnStartMonitor_Click(sender As Object, e As EventArgs) Handles btnStartMonitor.Click
        Try
            If txtReadingMode.Text <> 1 AndAlso txtReadingMode.Text <> 2 AndAlso txtReadingMode.Text <> 3 AndAlso txtReadingMode.Text <> 4 Then
                MessageBox.Show("Error: Reading mode not macth the preset.")
            Else
                If isstart = True Then
                    monitor.Cancel()
                End If
                startMonitor()
                isstart = True
            End If
        Catch ex As Exception
            LogError("Error:" + ex.Message)
        End Try
    End Sub

    Private Sub btnStopMonitor_Click(sender As Object, e As EventArgs) Handles btnStopMonitor.Click
        If isstart = True Then
            monitor.Cancel()
        End If
    End Sub

    Function SendUID4Byte()
        Try
            Using context = _contextFactory.Establish(SCardScope.System)
                Using rfidReader = context.ConnectReader(readerName, SCardShareMode.Shared, SCardProtocol.Any)
                    Using rfidReader.Transaction(SCardReaderDisposition.Leave)

                        Dim apdu As Byte() = {&HFF, &HCA, &H0, &H0, &H4}
                        Dim sendPci = SCardPCI.GetPci(rfidReader.Protocol)
                        Dim receivePci = New SCardPCI()

                        Dim receiveBuffer = New Byte(255) {}
                        Dim command = apdu.ToArray()
                        Dim bytesReceived = rfidReader.Transmit(sendPci, command, command.Length, receivePci, receiveBuffer, receiveBuffer.Length)
                        Dim responseApdu = New ResponseApdu(receiveBuffer, bytesReceived, IsoCase.Case2Short, rfidReader.Protocol)

                        If readingMode = 1 Then
                            Dim uid As String = BitConverter.ToString(responseApdu.GetData())
                            uid = uid.Replace("-", "")

                            SendKeys.SendWait(uid + "{ENTER}")
                        ElseIf readingMode = 2 Then
                            Dim uid As Byte() = New Byte(3) {}
                            Dim revuid As Byte() = New Byte(3) {}
                            Array.Copy(responseApdu.GetData(), uid, 4)
                            Array.Copy(uid, revuid, 4)
                            Array.Reverse(revuid, 0, 4)

                            Dim uid2 As String = BitConverter.ToString(revuid)
                            uid2 = uid2.Replace("-", "")

                            SendKeys.SendWait(uid2 + "{ENTER}")
                        End If
                    End Using
                End Using
            End Using
        Catch
            LogError("Error while SendUID4Byte:")
        End Try

        Return True
    End Function

    Function SendUID7Byte()
        Try
            Using context = _contextFactory.Establish(SCardScope.System)
                Using rfidReader = context.ConnectReader(readerName, SCardShareMode.Shared, SCardProtocol.Any)
                    Using rfidReader.Transaction(SCardReaderDisposition.Leave)

                        Dim apdu As Byte() = {&HFF, &HCA, &H0, &H0, &H7}
                        Dim sendPci = SCardPCI.GetPci(rfidReader.Protocol)
                        Dim receivePci = New SCardPCI()

                        Dim receiveBuffer = New Byte(255) {}
                        Dim command = apdu.ToArray()
                        Dim bytesReceived = rfidReader.Transmit(sendPci, command, command.Length, receivePci, receiveBuffer, receiveBuffer.Length)
                        Dim responseApdu = New ResponseApdu(receiveBuffer, bytesReceived, IsoCase.Case2Short, rfidReader.Protocol)

                        If readingMode = 3 Then
                            Dim uid As String = BitConverter.ToString(responseApdu.GetData())
                            uid = uid.Replace("-", "")

                            SendKeys.SendWait(uid + "{ENTER}")
                        ElseIf readingMode = 4 Then
                            Dim uid As Byte() = New Byte(6) {}
                            Dim revuid As Byte() = New Byte(6) {}
                            Array.Copy(responseApdu.GetData(), uid, 7)
                            Array.Copy(uid, revuid, 7)
                            Array.Reverse(revuid, 0, 7)

                            Dim uid2 As String = BitConverter.ToString(revuid)
                            uid2 = uid2.Replace("-", "")

                            SendKeys.SendWait(uid2 + "{ENTER}")
                        End If
                    End Using
                End Using
            End Using
        Catch
            LogError("Error while SendUID4Byte:")
        End Try

        Return True
    End Function

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Opacity = 1
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub frmMain_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            NotifyIcon1.Visible = True
            NotifyIcon1.ShowBalloonTip(2000)
            Me.Opacity = 0
        End If
    End Sub

    Private Sub LogError(content As String)
        Directory.CreateDirectory(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "UIDToKeyboard", "Logs"))
        File.AppendAllText(logFile, DateTime.Now.ToString() + ":" + content + Environment.NewLine)
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        If ToolStripMenuItem1.Checked Then
            AddApplicationToStartup()
        Else
            RemoveApplicationFromStartup()
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Environment.Exit(0)
    End Sub

    Public Sub AddApplicationToStartup()
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
                key.SetValue("UIDtoKeyboard", """" & Application.ExecutablePath & """")
            End Using
        Catch ex As Exception
            LogError("Error while adding to startup:" + ex.Message)
        End Try

    End Sub

    Public Sub RemoveApplicationFromStartup()
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
                key.DeleteValue("UIDtoKeyboard", False)
            End Using
        Catch ex As Exception
            LogError("Error while removing from startup:" + ex.Message)
        End Try
    End Sub

    Public Sub Start()
        _timer = New System.Timers.Timer(30000)
        AddHandler _timer.Elapsed, New ElapsedEventHandler(AddressOf Handler)
        _timer.Enabled = True
    End Sub

    Public Sub Handler(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        Threading.Thread.Sleep(2000)
        loadReaderList()
        Threading.Thread.Sleep(2000)
        startMonitor()
        If isstart Then
            _timer.Enabled = False
        End If
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.WindowsShutDown Then
            Environment.Exit(0)
        Else
            NotifyIcon1.Visible = True
            NotifyIcon1.ShowBalloonTip(2000)
            Me.Opacity = 0
            e.Cancel = True
        End If
    End Sub
End Class
