Imports System.Resources
Imports Microsoft.Win32

Public Class Analoguhr
    Inherits System.Windows.Forms.Form
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As IntPtr) As Boolean
    Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As IntPtr, ByRef wow64Process As Boolean) As Boolean
    Private Declare Function GetCurrentProcess Lib "kernel32" () As IntPtr

    Private Function Is64BitOS() As Boolean
        Dim isWow64 As Boolean = False
        If IsWow64Process(GetCurrentProcess(), isWow64) Then
            Return isWow64 ' True = 64-Bit-Windows
        Else
            Return False ' API nicht verfügbar oder Fehler
        End If
    End Function

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    Private components As System.ComponentModel.IContainer
    Friend WithEvents NotifyUhr As NotifyIcon
    Friend WithEvents ctxTray As ContextMenu
    Friend WithEvents mnuTray1 As MenuItem
    Friend WithEvents mnuTray2 As MenuItem
    Friend WithEvents mnuTray3 As MenuItem
    Friend WithEvents mnuTray4 As MenuItem
    Friend WithEvents mnuTray5 As MenuItem
    Friend WithEvents mnuTray6 As MenuItem
    Friend WithEvents mnuTray7 As MenuItem
    Friend WithEvents mnuTray8 As MenuItem
    Friend WithEvents mnuTray9 As MenuItem
    Friend WithEvents mnuTray10 As MenuItem
    Friend WithEvents mnuTray11 As MenuItem
    Friend WithEvents mnuTray12 As MenuItem
    Friend WithEvents mnuTray13 As MenuItem
    Friend WithEvents mnuTray14 As MenuItem
    Friend WithEvents mnuTray15 As MenuItem
    Friend WithEvents mnuTray16 As MenuItem

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As ResourceManager = New ResourceManager(GetType(Analoguhr))
        Me.NotifyUhr = New NotifyIcon(Me.components)
        Me.ctxTray = New ContextMenu()

        Me.mnuTray1 = New MenuItem("&Exit", AddressOf mnuTray1_Click)
        Me.mnuTray2 = New MenuItem("&About", AddressOf mnuTray2_Click)
        Me.mnuTray3 = New MenuItem()
        AddHandler Me.mnuTray3.DrawItem, AddressOf mnuTray3_DrawItem
        AddHandler Me.mnuTray3.MeasureItem, AddressOf mnuTray3_MeasureItem
        Me.mnuTray3.OwnerDraw = True
        Me.mnuTray4 = New MenuItem("&Sound", AddressOf mnuTray4_Click)
        Me.mnuTray5 = New MenuItem("&Services", AddressOf mnuTray5_Click)
        Me.mnuTray6 = New MenuItem("&Windows Explorer", AddressOf mnuTray6_Click)
        Me.mnuTray7 = New MenuItem("&Windows Task Manager", AddressOf mnuTray7_Click)
        Me.mnuTray8 = New MenuItem("&Windows Certificates", AddressOf mnuTray8_Click)
        Me.mnuTray9 = New MenuItem("&Monitor Properties", AddressOf mnuTray9_Click)
        Me.mnuTray10 = New MenuItem("&Control Panel", AddressOf mnuTray10_Click)
        Me.mnuTray11 = New MenuItem("&Programs and Features", AddressOf mnuTray11_Click)
        Me.mnuTray12 = New MenuItem("&Network Connections", AddressOf mnuTray12_Click)
        Me.mnuTray13 = New MenuItem("&Power Options", AddressOf mnuTray13_Click)
        Me.mnuTray14 = New MenuItem("&Computer Management", AddressOf mnuTray14_Click)
        Me.mnuTray15 = New MenuItem("&Device Manager", AddressOf mnuTray15_Click)
        Me.mnuTray16 = New MenuItem("&Change Date/Time", AddressOf mnuTray16_Click)

        Me.ctxTray.MenuItems.Add(Me.mnuTray16)
        Me.ctxTray.MenuItems.Add(Me.mnuTray15)
        Me.ctxTray.MenuItems.Add(Me.mnuTray14)
        Me.ctxTray.MenuItems.Add(Me.mnuTray13)
        Me.ctxTray.MenuItems.Add(Me.mnuTray12)
        Me.ctxTray.MenuItems.Add(Me.mnuTray11)
        Me.ctxTray.MenuItems.Add(Me.mnuTray10)
        Me.ctxTray.MenuItems.Add(Me.mnuTray9)
        Me.ctxTray.MenuItems.Add(Me.mnuTray8)
        Me.ctxTray.MenuItems.Add(Me.mnuTray7)
        Me.ctxTray.MenuItems.Add(Me.mnuTray6)
        Me.ctxTray.MenuItems.Add(Me.mnuTray5)
        Me.ctxTray.MenuItems.Add(Me.mnuTray4)
        Me.ctxTray.MenuItems.Add(Me.mnuTray3)
        Me.ctxTray.MenuItems.Add(Me.mnuTray2)
        Me.ctxTray.MenuItems.Add(Me.mnuTray1)

        Me.NotifyUhr.Icon = CType(resources.GetObject("NotifyUhr.Icon"), Icon)
        Me.NotifyUhr.Text = "Analoguhr"
        Me.NotifyUhr.Visible = True
        '
        'Analoguhr
        '
        Me.AutoScaleBaseSize = New Size(5, 13)
        Me.ClientSize = New Size(292, 77)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Analoguhr"
        Me.Opacity = 0.0R
        Me.ShowInTaskbar = False
        Me.Text = "Analoguhr"
        Me.ResumeLayout(False)
    End Sub

#End Region

    Private Sub mnuTray3_MeasureItem(ByVal sender As Object, ByVal e As MeasureItemEventArgs)
        e.ItemHeight = 2
        e.ItemWidth = 146
    End Sub

    Private Sub mnuTray3_DrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)
        Dim asm As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
        Dim stream As System.IO.Stream = asm.GetManifestResourceStream("Analoguhr.Trennlinie.bmp")
        Dim trennBild As Image = Image.FromStream(stream)
        e.Graphics.DrawImage(trennBild, e.Bounds.Left, e.Bounds.Top)
        stream.Close()
    End Sub

    Private Sub mnuTray16_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL timedate.cpl,,0")
    End Sub

    Private Sub mnuTray15_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")
            Dim productName As String = key.GetValue("ProductName").ToString()
            Dim buildNumber As String = key.GetValue("CurrentBuildNumber").ToString()

            'Messagebox.Show("Windows-Version: " & productName & vbCrLf & "Build: " & buildNumber)

            ' Überprüfe, ob es eine bestimmte Windows-Version ist
            If productName.IndexOf("Windows XP") >= 0 Then
                'Messagebox.Show("Windows XP recognized – Registry")
                Process.Start("devmgmt.msc")

                'Windows Vista
            ElseIf buildNumber.IndexOf("6002") >= 0 Then
                'Messagebox.Show("Windows Vista recognized – Registry")
                Process.Start("devmgmt.msc")

            Else
                'Messagebox.Show("all outher Windows recognized – Registry: " & productName)
                Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL hdwwiz.cpl")
            End If

        Catch ex As Exception
            'Messagebox.Show("Registry-Access failed: " & ex.Message)
        End Try
    End Sub

    Private Sub mnuTray14_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("compmgmt.msc")
    End Sub

    Private Sub mnuTray13_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL powercfg.cpl")
    End Sub

    Private Sub mnuTray12_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL ncpa.cpl")
    End Sub

    Private Sub mnuTray11_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL appwiz.cpl")
    End Sub

    Private Sub mnuTray10_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("control.exe")
    End Sub

    Private Sub mnuTray9_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")
            Dim productName As String = key.GetValue("ProductName").ToString()
            Dim buildNumber As String = key.GetValue("CurrentBuildNumber").ToString()
            Dim is64Bit As Boolean = Is64BitOS() ' ? ersetzt Environment.Is64BitProcess
            Dim rundllPath As String

            'Messagebox.Show("Windows-Version: " & productName & vbCrLf & "Build: " & buildNumber)

            If is64Bit Then
                'Messagebox.Show("Sysnative are used, is this Right ?: " & CStr(is64Bit))
                rundllPath = "C:\Windows\Sysnative\rundll32.exe" ' Zugriff auf echtes System32 aus 32-Bit-App
            Else
                'Messagebox.Show("System32 are used, is this Right ?: " & CStr(is64Bit))
                rundllPath = "C:\Windows\System32\rundll32.exe"
            End If

            ' Überprüfe, ob es eine bestimmte Windows-Version ist
            If productName.IndexOf("Windows 11") >= 0 Then
                'Messagebox.Show("Windows 11 recognized – Registry")
                'Process.Start("ms-settings:display")
                'Process.Start("Monitor.exe")
                Process.Start(rundllPath, "display.dll,ShowAdapterSettings 0")

            ElseIf productName.IndexOf("Windows 10") >= 0 Then
                'Messagebox.Show("Windows 10 recognized – Registry")
                'Process.Start("ms-settings:display")
                'Process.Start("Monitor.exe")
                Process.Start(rundllPath, "display.dll,ShowAdapterSettings 0")

            ElseIf productName.IndexOf("Windows 8.1") >= 0 Then
                'Messagebox.Show("Windows 8.1 recognized – Registry")
                'Process.Start("control.exe", "desk.cpl")
                'Process.Start("Monitor.exe")
                Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL desk.cpl,,3")

            ElseIf productName.IndexOf("Windows 7") >= 0 Then
                'Messagebox.Show("Windows 7 recognized – Registry")
                'Process.Start("control.exe", "desk.cpl")
                'Process.Start("Monitor.exe")
                Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL desk.cpl,,3")

            ElseIf buildNumber.IndexOf("6002") >= 0 Then
                'Messagebox.Show("Windows Vista recognized – Registry")
                'Process.Start("control.exe", "desk.cpl")
                'Process.Start("Monitor.exe")
                Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL desk.cpl,,3")

            ElseIf productName.IndexOf("Windows XP") >= 0 Then
                'Messagebox.Show("Windows XP recognized – Registry")
                'Process.Start("control.exe", "desk.cpl")
                'Process.Start("Monitor.exe")
                Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL desk.cpl,,3")

            Else
                'Messagebox.Show("Unknow or nicht unsupported Windows-Version: " & productName)
            End If

        Catch ex As Exception
            'Messagebox.Show("Error on Start: " & ex.Message)
            'Messagebox.Show("Registry-Access failed: " & ex.Message)
        End Try
    End Sub

    Private Sub mnuTray8_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("certmgr.msc")
    End Sub

    Private Sub mnuTray7_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("taskmgr.exe")
    End Sub

    Private Sub mnuTray6_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("explorer.exe")
    End Sub

    Private Sub mnuTray5_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("services.msc")
    End Sub

    Private Sub mnuTray4_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL mmsys.cpl")
    End Sub

    Private Sub mnuTray3_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Nichts
    End Sub

    Private Sub mnuTray2_Click(ByVal sender As Object, ByVal e As EventArgs)
        Messagebox.Show("Analoguhr v1.6.0" & Environment.NewLine & "(C) Your Name" & Environment.NewLine & "08.August.2025", "Information", MessageboxButtons.OK, MessageboxIcon.Information)
    End Sub

    Private m_CloseOk As Boolean = False
    Private Sub mnuTray1_Click(ByVal sender As Object, ByVal e As EventArgs)
        m_CloseOk = True
        Me.Close()
        Application.Exit()
        End
    End Sub

    ' ? Linksklick und Rechtsklick sauber behandelt
    Private Sub NotifyUhr_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyUhr.MouseUp
        'If e.Button = MouseButtons.Left Then
        If e.Button = MouseButtons.Left OrElse e.Button = MouseButtons.Right Then
            SetForegroundWindow(Me.Handle)
            ctxTray.Show(Me, Me.PointToClient(Control.MousePosition))
        End If
    End Sub
End Class