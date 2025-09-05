Imports System.Windows.Forms
Imports System.Drawing
Imports System.Reflection
Imports BCOM = MicroStationDGN
''' <summary>
''' Professional About dialog for SSM MicroStation V8i Library
''' Displays version information, credits, and system details
''' </summary>
''' <remarks>
''' This VBA-style form integrates with MicroStation V8i using modal dialog events
''' and provides professional branding and information display
''' </remarks>
Public Class AboutDialog
    Inherits Form
    Private ustation As BCOM.Application
    Private WithEvents lblTitle As Label
    Private WithEvents lblVersion As Label
    Private WithEvents lblDescription As Label
    Private WithEvents lblCopyright As Label
    Private WithEvents lblSystemInfo As Label
    Private WithEvents btnOK As Button
    Private WithEvents btnSystemDetails As Button
    Private WithEvents picLogo As PictureBox
    ''' <summary>
    ''' Initializes the About dialog with professional styling and branding
    ''' </summary>
    ''' <param name="msApp">Reference to MicroStation Application object</param>
    Public Sub New(msApp As BCOM.Application)
        ustation = msApp
        InitializeComponent()
        LoadVersionInformation()
        SetupProfessionalStyling()
    End Sub
    ''' <summary>
    ''' Sets up the form layout and controls with professional appearance
    ''' </summary>
    Private Sub InitializeComponent()
        Me.Text = "About SSM MicroStation V8i Tools"
        Me.Size = New Size(480, 360)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.BackColor = Color.White
        ' Logo placeholder
        picLogo = New PictureBox()
        picLogo.Size = New Size(64, 64)
        picLogo.Location = New Point(20, 20)
        picLogo.BackColor = Color.LightBlue
        picLogo.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(picLogo)
        ' Title
        lblTitle = New Label()
        lblTitle.Text = "SSM MicroStation V8i Tools"
        lblTitle.Font = New Font("Segoe UI", 16, FontStyle.Bold)
        lblTitle.ForeColor = Color.DarkBlue
        lblTitle.Location = New Point(100, 20)
        lblTitle.Size = New Size(350, 30)
        Me.Controls.Add(lblTitle)
        ' Version
        lblVersion = New Label()
        lblVersion.Text = "Version 1.0.0"
        lblVersion.Font = New Font("Segoe UI", 10, FontStyle.Regular)
        lblVersion.Location = New Point(100, 55)
        lblVersion.Size = New Size(350, 20)
        Me.Controls.Add(lblVersion)
        ' Description
        lblDescription = New Label()
        lblDescription.Text = "Professional 2D/3D CAD tools and geometry library for MicroStation V8i." & vbCrLf & 
                             "Provides comprehensive automation and extension capabilities."
        lblDescription.Font = New Font("Segoe UI", 9, FontStyle.Regular)
        lblDescription.Location = New Point(20, 110)
        lblDescription.Size = New Size(430, 40)
        Me.Controls.Add(lblDescription)
        ' Copyright
        lblCopyright = New Label()
        lblCopyright.Text = "© 2024 SSM Engineering. All rights reserved."
        lblCopyright.Font = New Font("Segoe UI", 8, FontStyle.Regular)
        lblCopyright.Location = New Point(20, 165)
        lblCopyright.Size = New Size(430, 15)
        Me.Controls.Add(lblCopyright)
        ' System Info
        lblSystemInfo = New Label()
        lblSystemInfo.Text = "Loading system information..."
        lblSystemInfo.Font = New Font("Consolas", 8, FontStyle.Regular)
        lblSystemInfo.Location = New Point(20, 190)
        lblSystemInfo.Size = New Size(430, 80)
        lblSystemInfo.BorderStyle = BorderStyle.FixedSingle
        lblSystemInfo.BackColor = Color.LightGray
        Me.Controls.Add(lblSystemInfo)
        ' OK Button
        btnOK = New Button()
        btnOK.Text = "OK"
        btnOK.Size = New Size(80, 30)
        btnOK.Location = New Point(290, 290)
        btnOK.DialogResult = DialogResult.OK
        Me.AcceptButton = btnOK
        Me.Controls.Add(btnOK)
        ' System Details Button
        btnSystemDetails = New Button()
        btnSystemDetails.Text = "System Details"
        btnSystemDetails.Size = New Size(100, 30)
        btnSystemDetails.Location = New Point(180, 290)
        Me.Controls.Add(btnSystemDetails)
    End Sub
    ''' <summary>
    ''' Loads and displays version information from assembly metadata
    ''' </summary>
    Private Sub LoadVersionInformation()
        Try
            Dim assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim version = assembly.GetName().Version
            lblVersion.Text = $"Version {version.Major}.{version.Minor}.{version.Build}"
            ' Load MicroStation version info
            If ustation IsNot Nothing Then
                lblSystemInfo.Text = $"MicroStation V8i Integration Library" & vbCrLf &
                                   $"Assembly: {assembly.GetName().Name}" & vbCrLf &
                                   $"Version: {version}" & vbCrLf &
                                   $"Build Date: {System.IO.File.GetCreationTime(assembly.Location):yyyy-MM-dd}" & vbCrLf &
                                   $".NET Framework: {Environment.Version}" & vbCrLf &
                                   $"OS: {Environment.OSVersion}"
            End If
        Catch ex As Exception
            lblSystemInfo.Text = "Error loading version information: " & ex.Message
        End Try
    End Sub
    ''' <summary>
    ''' Applies professional styling and visual enhancements
    ''' </summary>
    Private Sub SetupProfessionalStyling()
        ' Add gradient background effect (simplified for VB.NET compatibility)
        Me.BackColor = Color.FromArgb(240, 248, 255) ' Light blue background
        ' Style the logo area
        picLogo.BackColor = Color.SteelBlue
        picLogo.ForeColor = Color.White
        ' Add visual separators
        Dim separator As New Label()
        separator.BackColor = Color.LightGray
        separator.Height = 1
        separator.Width = 430
        separator.Location = New Point(20, 185)
        Me.Controls.Add(separator)
    End Sub
    ''' <summary>
    ''' Handles system details button click - shows extended system information
    ''' </summary>
    Private Sub btnSystemDetails_Click(sender As Object, e As EventArgs) Handles btnSystemDetails.Click
        Try
            Dim detailsForm As New SystemDetailsDialog(ustation)
            detailsForm.ShowDialog(Me)
        Catch ex As Exception
            MessageBox.Show($"Error displaying system details: {ex.Message}", "Error", 
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
    ''' <summary>
    ''' Shows the About dialog as a modal dialog integrated with MicroStation
    ''' </summary>
    ''' <returns>DialogResult indicating user's choice</returns>
    Public Function ShowAboutDialog() As DialogResult
        Try
            Return Me.ShowDialog()
        Catch ex As Exception
            MessageBox.Show($"Error displaying About dialog: {ex.Message}", "Error", 
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return DialogResult.Cancel
        End Try
    End Function
    ''' <summary>
    ''' Handles MicroStation dialog opened events
    ''' </summary>
    Private Sub OnMicroStationDialogOpened(DialogBoxName As String, ByRef DialogResult As BCOM.MsdDialogBoxResult)
        ' Handle integration with MicroStation's dialog system
        If DialogBoxName.Contains("About") Then
            DialogResult = BCOM.MsdDialogBoxResult.msdDialogBoxResultOK
        End If
    End Sub
    ''' <summary>
    ''' Handles MicroStation dialog closed events
    ''' </summary>
    Private Sub OnMicroStationDialogClosed(DialogBoxName As String, DialogResult As BCOM.MsdDialogBoxResult)
        ' Clean up resources when MicroStation dialog closes
        ' This ensures proper integration with MicroStation's window management
    End Sub
    ''' <summary>
    ''' Clean up resources when dialog is disposed
    ''' </summary>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing Then
                ' Clean up resources
                If ustation IsNot Nothing Then
                    ustation = Nothing
                End If
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
End Class
''' <summary>
''' Extended system details dialog showing comprehensive system information
''' </summary>
Public Class SystemDetailsDialog
    Inherits Form
    Private ustation As BCOM.Application
    Private WithEvents txtDetails As TextBox
    Private WithEvents btnClose As Button
    Private WithEvents btnCopy As Button
    ''' <summary>
    ''' Initializes the system details dialog
    ''' </summary>
    ''' <param name="msApp">Reference to MicroStation Application object</param>
    Public Sub New(msApp As BCOM.Application)
        ustation = msApp
        InitializeComponent()
        LoadSystemDetails()
    End Sub
    ''' <summary>
    ''' Sets up the system details dialog layout and controls
    ''' </summary>
    Private Sub InitializeComponent()
        Me.Text = "System Details - SSM MicroStation V8i Tools"
        Me.Size = New Size(600, 450)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.Sizable
        ' Details text box
        txtDetails = New TextBox()
        txtDetails.Multiline = True
        txtDetails.ReadOnly = True
        txtDetails.ScrollBars = ScrollBars.Both
        txtDetails.Font = New Font("Consolas", 9)
        txtDetails.Location = New Point(10, 10)
        txtDetails.Size = New Size(560, 350)
        txtDetails.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Me.Controls.Add(txtDetails)
        ' Copy button
        btnCopy = New Button()
        btnCopy.Text = "Copy to Clipboard"
        btnCopy.Size = New Size(120, 30)
        btnCopy.Location = New Point(350, 370)
        btnCopy.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        Me.Controls.Add(btnCopy)
        ' Close button
        btnClose = New Button()
        btnClose.Text = "Close"
        btnClose.Size = New Size(80, 30)
        btnClose.Location = New Point(480, 370)
        btnClose.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        btnClose.DialogResult = DialogResult.OK
        Me.AcceptButton = btnClose
        Me.Controls.Add(btnClose)
    End Sub
    ''' <summary>
    ''' Loads comprehensive system information including MicroStation details
    ''' </summary>
    Private Sub LoadSystemDetails()
        Try
            Dim details As New System.Text.StringBuilder()
            details.AppendLine("=== SSM MicroStation V8i Tools - System Details ===")
            details.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            details.AppendLine()
            ' Assembly Information
            details.AppendLine("=== Assembly Information ===")
            Dim assembly = System.Reflection.Assembly.GetExecutingAssembly()
            details.AppendLine($"Assembly Name: {assembly.GetName().Name}")
            details.AppendLine($"Version: {assembly.GetName().Version}")
            details.AppendLine($"Location: {assembly.Location}")
            details.AppendLine($"Build Date: {System.IO.File.GetCreationTime(assembly.Location):yyyy-MM-dd HH:mm:ss}")
            details.AppendLine()
            ' System Environment
            details.AppendLine("=== System Environment ===")
            details.AppendLine($"Operating System: {Environment.OSVersion}")
            details.AppendLine($".NET Framework: {Environment.Version}")
            details.AppendLine($"Machine Name: {Environment.MachineName}")
            details.AppendLine($"User Name: {Environment.UserName}")
            details.AppendLine($"Current Directory: {Environment.CurrentDirectory}")
            details.AppendLine($"System Directory: {Environment.SystemDirectory}")
            details.AppendLine($"Processor Count: {Environment.ProcessorCount}")
            details.AppendLine($"Working Set: {Environment.WorkingSet:N0} bytes")
            details.AppendLine()
            ' MicroStation Information
            If ustation IsNot Nothing Then
                details.AppendLine("=== MicroStation V8i Information ===")
                Try
                    details.AppendLine($"Application Object: Connected")
                    details.AppendLine($"Version: MicroStation V8i")
                    details.AppendLine($"Visible: {ustation.Visible}")
                    If ustation.ActiveDesignFile IsNot Nothing Then
                        details.AppendLine($"Active Design File: {ustation.ActiveDesignFile.Name}")
                        details.AppendLine($"Design File Path: {ustation.ActiveDesignFile.Path}")
                    Else
                        details.AppendLine("Active Design File: None")
                    End If
                Catch ex As Exception
                    details.AppendLine($"Error accessing MicroStation info: {ex.Message}")
                End Try
            Else
                details.AppendLine("=== MicroStation V8i Information ===")
                details.AppendLine("MicroStation Application: Not Connected")
            End If
            details.AppendLine()
            ' COM References
            details.AppendLine("=== COM References ===")
            details.AppendLine("MicroStationDGN: Available")
            details.AppendLine("Bentley MicroStation DGN 8.9 Object Library: Referenced")
            details.AppendLine()
            txtDetails.Text = details.ToString()
        Catch ex As Exception
            txtDetails.Text = $"Error loading system details: {ex.Message}"
        End Try
    End Sub
    ''' <summary>
    ''' Copies system details to clipboard
    ''' </summary>
    Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
        Try
            Clipboard.SetText(txtDetails.Text)
            MessageBox.Show("System details copied to clipboard.", "Information", 
                          MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"Error copying to clipboard: {ex.Message}", "Error", 
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
End Class