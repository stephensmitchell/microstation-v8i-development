Imports System.Windows.Forms
Imports BCOM = MicroStationDGN
''' <summary>
''' Professional UI Manager for SSM MicroStation V8i Library
''' Coordinates and manages all user interface components including dialogs, panels, and windows
''' Provides centralized UI management and integration with MicroStation's interface system
''' </summary>
''' <remarks>
''' This manager handles the lifecycle and interaction of all professional UI components:
''' - About dialog and system information
''' - Help system panel with integrated documentation  
''' - Settings and configuration windows
''' - Status and notification systems
''' - Menu and toolbar integration
''' </remarks>
Public Class UIManager
    Implements IDisposable
    Private ustation As BCOM.Application
    Private helpPanel As HelpSystemPanel
    Private aboutDialog As AboutDialog
    Private _isInitialized As Boolean = False
    Private disposedValue As Boolean = False
    ' UI State Management
    Private _helpPanelVisible As Boolean = False
    Private _aboutDialogOpen As Boolean = False
    ''' <summary>
    ''' Initializes the UI Manager with MicroStation integration
    ''' </summary>
    ''' <param name="msApp">Reference to MicroStation Application object</param>
    Public Sub New(msApp As BCOM.Application)
        If msApp Is Nothing Then
            Throw New ArgumentNullException(NameOf(msApp), "MicroStation Application object cannot be null")
        End If
        ustation = msApp
        InitializeUI()
    End Sub
    ''' <summary>
    ''' Gets the current initialization state of the UI Manager
    ''' </summary>
    ''' <returns>True if UI components are successfully initialized</returns>
    Public ReadOnly Property IsInitialized As Boolean
        Get
            Return _isInitialized
        End Get
    End Property
    ''' <summary>
    ''' Gets the visibility state of the help panel
    ''' </summary>
    ''' <returns>True if help panel is currently visible</returns>
    Public ReadOnly Property IsHelpPanelVisible As Boolean
        Get
            Return _helpPanelVisible
        End Get
    End Property
    ''' <summary>
    ''' Gets the state of the about dialog
    ''' </summary>
    ''' <returns>True if about dialog is currently open</returns>
    Public ReadOnly Property IsAboutDialogOpen As Boolean
        Get
            Return _aboutDialogOpen
        End Get
    End Property
    ''' <summary>
    ''' Initializes all UI components and establishes MicroStation integration
    ''' </summary>
    Private Sub InitializeUI()
        Try
            ' Initialize Help System Panel
            InitializeHelpPanel()
            ' Initialize About Dialog (created on-demand for performance)
            ' aboutDialog will be created when first needed
            ' Register MicroStation menu integration
            RegisterMenuItems()
            ' Setup UI event handlers
            SetupEventHandlers()
            _isInitialized = True
        Catch ex As Exception
            Throw New Exception($"Failed to initialize UI Manager: {ex.Message}", ex)
        End Try
    End Sub
    ''' <summary>
    ''' Initializes the integrated help system panel
    ''' </summary>
    Private Sub InitializeHelpPanel()
        Try
            helpPanel = New HelpSystemPanel(ustation)
            ' Help panel creates its own MicroStation window integration
        Catch ex As Exception
            Throw New Exception($"Failed to initialize help panel: {ex.Message}", ex)
        End Try
    End Sub
    ''' <summary>
    ''' Registers menu items and toolbar buttons in MicroStation's interface
    ''' </summary>
    Private Sub RegisterMenuItems()
        Try
            ' This would typically involve MicroStation's menu system
            ' For now, we'll prepare the infrastructure for menu integration
            ' Note: Full menu integration would require additional MicroStation APIs
            ' that may not be directly available through COM interop
            ' This could be enhanced with MDL or custom menu creation
        Catch ex As Exception
            ' Log error but don't fail initialization
            System.Diagnostics.Debug.WriteLine($"Menu registration warning: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Sets up event handlers for UI coordination and MicroStation integration
    ''' </summary>
    Private Sub SetupEventHandlers()
        Try
            ' Setup application-level event handlers if available
            ' This ensures proper cleanup and state management
            AddHandler Application.ApplicationExit, AddressOf OnApplicationExit
        Catch ex As Exception
            ' Log error but don't fail initialization
            System.Diagnostics.Debug.WriteLine($"Event handler setup warning: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Shows the integrated help panel
    ''' Creates the panel if it doesn't exist, or brings it to front if it does
    ''' </summary>
    Public Sub ShowHelpPanel()
        Try
            If Not _isInitialized Then
                Throw New InvalidOperationException("UI Manager is not initialized")
            End If
            If helpPanel Is Nothing Then
                InitializeHelpPanel()
            End If
            helpPanel.ShowHelpPanel()
            _helpPanelVisible = True
        Catch ex As Exception
            ShowErrorMessage("Error Displaying Help", 
                           $"Unable to display help panel: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Hides the help panel
    ''' </summary>
    Public Sub HideHelpPanel()
        Try
            If helpPanel IsNot Nothing Then
                helpPanel.HideHelpPanel()
                _helpPanelVisible = False
            End If
        Catch ex As Exception
            ' Handle errors gracefully - hiding shouldn't cause critical failures
            System.Diagnostics.Debug.WriteLine($"Error hiding help panel: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Toggles the visibility of the help panel
    ''' </summary>
    Public Sub ToggleHelpPanel()
        If _helpPanelVisible Then
            HideHelpPanel()
        Else
            ShowHelpPanel()
        End If
    End Sub
    ''' <summary>
    ''' Shows the About dialog with version and system information
    ''' </summary>
    ''' <returns>DialogResult from the About dialog</returns>
    Public Function ShowAboutDialog() As DialogResult
        Try
            If Not _isInitialized Then
                Throw New InvalidOperationException("UI Manager is not initialized")
            End If
            ' Create About dialog on-demand for better performance
            If aboutDialog Is Nothing Then
                aboutDialog = New AboutDialog(ustation)
            End If
            _aboutDialogOpen = True
            Dim result As DialogResult = aboutDialog.ShowAboutDialog()
            _aboutDialogOpen = False
            Return result
        Catch ex As Exception
            _aboutDialogOpen = False
            ShowErrorMessage("Error Displaying About Dialog", 
                           $"Unable to display about dialog: {ex.Message}")
            Return DialogResult.Cancel
        End Try
    End Function
    ''' <summary>
    ''' Shows a professional status message in MicroStation's status bar
    ''' </summary>
    ''' <param name="message">Status message to display</param>
    ''' <param name="area">Status bar area (optional)</param>
    Public Sub ShowStatusMessage(message As String, Optional area As BCOM.MsdStatusBarArea = BCOM.MsdStatusBarArea.msdStatusBarAreaMiddle)
        Try
            If ustation IsNot Nothing Then
                ustation.ShowTempMessage(area, message)
            End If
        Catch ex As Exception
            ' Fallback to system message if MicroStation status fails
            System.Diagnostics.Debug.WriteLine($"Status message: {message}")
        End Try
    End Sub
    ''' <summary>
    ''' Shows a professional error message using MicroStation's message system
    ''' </summary>
    ''' <param name="title">Error title</param>
    ''' <param name="message">Error message details</param>
    ''' <param name="priority">Message priority level</param>
    Public Sub ShowErrorMessage(title As String, message As String, 
                               Optional priority As BCOM.MsdMessageCenterPriority = BCOM.MsdMessageCenterPriority.msdMessageCenterPriorityError)
        Try
            If ustation IsNot Nothing Then
                ' Use MicroStation's integrated message system
                ustation.ShowMessage($"{title}: {message}", "", priority, True)
            Else
                ' Fallback to Windows Forms message box
                MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            ' Last resort fallback
            MessageBox.Show($"{title}: {message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' Shows a professional information message using MicroStation's message system
    ''' </summary>
    ''' <param name="title">Message title</param>
    ''' <param name="message">Message details</param>
    ''' <param name="priority">Message priority level</param>
    Public Sub ShowInformationMessage(title As String, message As String, 
                                    Optional priority As BCOM.MsdMessageCenterPriority = BCOM.MsdMessageCenterPriority.msdMessageCenterPriorityInfo)
        Try
            If ustation IsNot Nothing Then
                ustation.ShowMessage($"{title}: {message}", "", priority)
            Else
                MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub
    ''' <summary>
    ''' Shows a professional confirmation dialog using MicroStation's message system
    ''' </summary>
    ''' <param name="title">Dialog title</param>
    ''' <param name="message">Confirmation message</param>
    ''' <returns>True if user confirms, False otherwise</returns>
    Public Function ShowConfirmationDialog(title As String, message As String) As Boolean
        Try
            ' For confirmation dialogs, Windows Forms provides better control
            Dim result As DialogResult = MessageBox.Show(message, title, 
                                                       MessageBoxButtons.YesNo, 
                                                       MessageBoxIcon.Question, 
                                                       MessageBoxDefaultButton.Button2)
            Return result = DialogResult.Yes
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Displays context-sensitive help for a specific topic
    ''' </summary>
    ''' <param name="topic">Help topic to display</param>
    Public Sub ShowContextHelp(topic As String)
        Try
            ShowHelpPanel()
            ' Note: Additional logic would be needed to navigate to specific topic
            ' This could be implemented by extending the HelpSystemPanel
        Catch ex As Exception
            ShowErrorMessage("Context Help Error", 
                           $"Unable to display help for topic '{topic}': {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Gets comprehensive system information for diagnostics
    ''' </summary>
    ''' <returns>Formatted system information string</returns>
    Public Function GetSystemInformation() As String
        Try
            Dim info As New System.Text.StringBuilder()
            info.AppendLine("=== SSM MicroStation V8i Tools - System Information ===")
            info.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            info.AppendLine()
            info.AppendLine("UI Manager Status:")
            info.AppendLine($"  Initialized: {_isInitialized}")
            info.AppendLine($"  Help Panel Visible: {_helpPanelVisible}")
            info.AppendLine($"  About Dialog Open: {_aboutDialogOpen}")
            info.AppendLine()
            info.AppendLine("MicroStation Integration:")
            If ustation IsNot Nothing Then
                info.AppendLine($"  Connection: Active")
                info.AppendLine($"  Visible: {ustation.Visible}")
            Else
                info.AppendLine($"  Connection: Not Available")
            End If
            info.AppendLine()
            info.AppendLine($"System Environment:")
            info.AppendLine($"  OS: {Environment.OSVersion}")
            info.AppendLine($"  .NET Framework: {Environment.Version}")
            info.AppendLine($"  Working Directory: {Environment.CurrentDirectory}")
            Return info.ToString()
        Catch ex As Exception
            Return $"Error retrieving system information: {ex.Message}"
        End Try
    End Function
    ''' <summary>
    ''' Handles application exit to ensure proper cleanup
    ''' </summary>
    Private Sub OnApplicationExit(sender As Object, e As EventArgs)
        Dispose()
    End Sub
    ''' <summary>
    ''' Disposes of all UI resources and cleans up MicroStation integration
    ''' </summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    ''' <summary>
    ''' Protected dispose method for proper cleanup
    ''' </summary>
    ''' <param name="disposing">True if disposing managed resources</param>
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                Try
                    ' Hide and dispose help panel
                    If helpPanel IsNot Nothing Then
                        helpPanel.HideHelpPanel()
                        helpPanel.Dispose()
                        helpPanel = Nothing
                    End If
                    ' Dispose about dialog
                    If aboutDialog IsNot Nothing Then
                        aboutDialog.Dispose()
                        aboutDialog = Nothing
                    End If
                    ' Remove event handlers
                    RemoveHandler Application.ApplicationExit, AddressOf OnApplicationExit
                    ' Release MicroStation reference
                    If ustation IsNot Nothing Then
                        ' Note: We don't dispose of ustation as it's managed externally
                        ustation = Nothing
                    End If
                    _isInitialized = False
                Catch ex As Exception
                    ' Log cleanup errors but don't throw during dispose
                    System.Diagnostics.Debug.WriteLine($"Error during UIManager dispose: {ex.Message}")
                End Try
            End If
            disposedValue = True
        End If
    End Sub
    ''' <summary>
    ''' Finalizer to ensure cleanup if Dispose wasn't called
    ''' </summary>
    Protected Overrides Sub Finalize()
        Dispose(False)
        MyBase.Finalize()
    End Sub
End Class