Imports System.Windows.Forms
Imports BCOM = MicroStationDGN
''' <summary>
''' Comprehensive integration example demonstrating professional UI components
''' Shows how to integrate About dialog, Help system, and UI Manager with MicroStation V8i
''' </summary>
''' <remarks>
''' This example provides a complete working demonstration of:
''' - Professional About dialog with system information
''' - Integrated help system with dockable panel
''' - UI Manager for centralized interface coordination
''' - MicroStation COM integration and event handling
''' - Error handling and resource management
''' </remarks>
Public Class UIIntegrationExample
    Private ustation As BCOM.Application
    Private uiManager As UIManager
    Private isConnected As Boolean = False
    ''' <summary>
    ''' Main entry point for the UI integration demonstration
    ''' </summary>
    ''' <remarks>
    ''' This method demonstrates the complete workflow for integrating
    ''' professional UI components with MicroStation V8i
    ''' </remarks>
    Public Shared Sub RunUIDemo()
        Dim demo As New UIIntegrationExample()
        demo.ExecuteDemo()
    End Sub
    ''' <summary>
    ''' Executes the complete UI integration demonstration
    ''' </summary>
    Public Sub ExecuteDemo()
        Try
            ' Step 1: Connect to MicroStation
            Console.WriteLine("=== SSM MicroStation V8i UI Integration Demo ===")
            Console.WriteLine("Step 1: Connecting to MicroStation...")
            ConnectToMicroStation()
            If Not isConnected Then
                Console.WriteLine("Demo cannot continue without MicroStation connection.")
                Return
            End If
            ' Step 2: Initialize UI Manager
            Console.WriteLine("Step 2: Initializing UI Manager...")
            InitializeUIManager()
            ' Step 3: Demonstrate UI Components
            Console.WriteLine("Step 3: Demonstrating UI Components...")
            DemonstrateUIComponents()
            ' Step 4: Interactive Demo
            Console.WriteLine("Step 4: Starting Interactive Demo...")
            RunInteractiveDemo()
        Catch ex As Exception
            Console.WriteLine($"Demo Error: {ex.Message}")
        Finally
            ' Cleanup
            CleanupDemo()
        End Try
    End Sub
    ''' <summary>
    ''' Establishes connection to MicroStation V8i using the MsCom utility
    ''' </summary>
    Private Sub ConnectToMicroStation()
        Try
            ' Use the MsCom utility from our library for professional connection handling
            ustation = MsCom.Attach(True)
            isConnected = True
            Console.WriteLine("✓ Successfully connected to MicroStation V8i")
            Console.WriteLine($"  Version: MicroStation V8i")
            Console.WriteLine($"  Visible: {ustation.Visible}")
            If ustation.ActiveDesignFile IsNot Nothing Then
                Console.WriteLine($"  Active File: {ustation.ActiveDesignFile.Name}")
            End If
        Catch ex As Exception
            Console.WriteLine($"✗ Failed to connect to MicroStation: {ex.Message}")
            Console.WriteLine("  Please ensure MicroStation V8i is running and try again.")
            isConnected = False
        End Try
    End Sub
    ''' <summary>
    ''' Initializes the UI Manager with all professional components
    ''' </summary>
    Private Sub InitializeUIManager()
        Try
            uiManager = New UIManager(ustation)
            If uiManager.IsInitialized Then
                Console.WriteLine("✓ UI Manager initialized successfully")
                Console.WriteLine("  - Help system panel ready")
                Console.WriteLine("  - About dialog prepared")
                Console.WriteLine("  - MicroStation integration active")
            Else
                Throw New Exception("UI Manager failed to initialize")
            End If
        Catch ex As Exception
            Console.WriteLine($"✗ Failed to initialize UI Manager: {ex.Message}")
            Throw
        End Try
    End Sub
    ''' <summary>
    ''' Demonstrates all UI components with automated testing
    ''' </summary>
    Private Sub DemonstrateUIComponents()
        Try
            ' Demonstrate status messages
            Console.WriteLine("Demonstrating status messages...")
            uiManager.ShowStatusMessage("SSM Tools: UI Demo in progress")
            System.Threading.Thread.Sleep(1000)
            ' Demonstrate information message
            Console.WriteLine("Demonstrating information messages...")
            uiManager.ShowInformationMessage("Demo Information", 
                "This demonstrates the professional messaging system integration.")
            System.Threading.Thread.Sleep(2000)
            ' Test system information
            Console.WriteLine("Retrieving system information...")
            Dim sysInfo As String = uiManager.GetSystemInformation()
            Console.WriteLine("System information retrieved successfully")
            Console.WriteLine("First few lines:")
            Dim lines As String() = sysInfo.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            For i As Integer = 0 To Math.Min(4, lines.Length - 1)
                Console.WriteLine($"  {lines(i)}")
            Next
        Catch ex As Exception
            Console.WriteLine($"Error in UI demonstration: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Runs an interactive demonstration allowing user to test all features
    ''' </summary>
    Private Sub RunInteractiveDemo()
        Try
            Console.WriteLine()
            Console.WriteLine("=== Interactive UI Demo Menu ===")
            Console.WriteLine("Choose an option to test:")
            Console.WriteLine("1. Show About Dialog")
            Console.WriteLine("2. Show Help Panel")
            Console.WriteLine("3. Hide Help Panel")
            Console.WriteLine("4. Toggle Help Panel")
            Console.WriteLine("5. Test Confirmation Dialog")
            Console.WriteLine("6. Display System Information")
            Console.WriteLine("7. Test Error Handling")
            Console.WriteLine("8. Show Status Messages")
            Console.WriteLine("9. Test Context Help")
            Console.WriteLine("0. Exit Demo")
            Console.WriteLine()
            Dim running As Boolean = True
            While running
                Console.Write("Enter your choice (0-9): ")
                Dim input As String = Console.ReadLine()
                Select Case input
                    Case "1"
                        TestAboutDialog()
                    Case "2"
                        TestShowHelpPanel()
                    Case "3"
                        TestHideHelpPanel()
                    Case "4"
                        TestToggleHelpPanel()
                    Case "5"
                        TestConfirmationDialog()
                    Case "6"
                        TestSystemInformation()
                    Case "7"
                        TestErrorHandling()
                    Case "8"
                        TestStatusMessages()
                    Case "9"
                        TestContextHelp()
                    Case "0"
                        running = False
                        Console.WriteLine("Exiting interactive demo...")
                    Case Else
                        Console.WriteLine("Invalid option. Please enter 0-9.")
                End Select
                If running Then
                    Console.WriteLine()
                    Console.WriteLine("Press any key to continue...")
                    Console.ReadKey()
                    Console.WriteLine()
                End If
            End While
        Catch ex As Exception
            Console.WriteLine($"Error in interactive demo: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Tests the About dialog functionality
    ''' </summary>
    Private Sub TestAboutDialog()
        Try
            Console.WriteLine("Testing About Dialog...")
            Console.WriteLine("Opening About dialog (check MicroStation for modal dialog)...")
            Dim result As DialogResult = uiManager.ShowAboutDialog()
            Console.WriteLine($"About dialog closed with result: {result}")
        Catch ex As Exception
            Console.WriteLine($"Error testing About dialog: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Tests showing the help panel
    ''' </summary>
    Private Sub TestShowHelpPanel()
        Try
            Console.WriteLine("Testing Show Help Panel...")
            uiManager.ShowHelpPanel()
            Console.WriteLine("Help panel should now be visible in MicroStation")
            Console.WriteLine($"Help panel visible: {uiManager.IsHelpPanelVisible}")
        Catch ex As Exception
            Console.WriteLine($"Error showing help panel: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Tests hiding the help panel
    ''' </summary>
    Private Sub TestHideHelpPanel()
        Try
            Console.WriteLine("Testing Hide Help Panel...")
            uiManager.HideHelpPanel()
            Console.WriteLine("Help panel should now be hidden")
            Console.WriteLine($"Help panel visible: {uiManager.IsHelpPanelVisible}")
        Catch ex As Exception
            Console.WriteLine($"Error hiding help panel: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Tests toggling the help panel visibility
    ''' </summary>
    Private Sub TestToggleHelpPanel()
        Try
            Console.WriteLine("Testing Toggle Help Panel...")
            Console.WriteLine($"Current state: {uiManager.IsHelpPanelVisible}")
            uiManager.ToggleHelpPanel()
            Console.WriteLine($"New state: {uiManager.IsHelpPanelVisible}")
        Catch ex As Exception
            Console.WriteLine($"Error toggling help panel: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Tests the confirmation dialog functionality
    ''' </summary>
    Private Sub TestConfirmationDialog()
        Try
            Console.WriteLine("Testing Confirmation Dialog...")
            Dim confirmed As Boolean = uiManager.ShowConfirmationDialog(
                "Demo Confirmation",
                "This is a test confirmation dialog." & vbCrLf & 
                "Click Yes to confirm or No to cancel.")
            Console.WriteLine($"User confirmation result: {confirmed}")
        Catch ex As Exception
            Console.WriteLine($"Error testing confirmation dialog: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Tests system information display
    ''' </summary>
    Private Sub TestSystemInformation()
        Try
            Console.WriteLine("Testing System Information...")
            Dim sysInfo As String = uiManager.GetSystemInformation()
            Console.WriteLine("=== System Information ===")
            Console.WriteLine(sysInfo)
            Console.WriteLine("=== End System Information ===")
        Catch ex As Exception
            Console.WriteLine($"Error retrieving system information: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Tests error handling and error message display
    ''' </summary>
    Private Sub TestErrorHandling()
        Try
            Console.WriteLine("Testing Error Handling...")
            uiManager.ShowErrorMessage("Demo Error", 
                "This is a demonstration error message." & vbCrLf & 
                "It shows how errors are displayed professionally in MicroStation.")
            Console.WriteLine("Error message displayed successfully")
        Catch ex As Exception
            Console.WriteLine($"Error testing error handling: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Tests various status message types
    ''' </summary>
    Private Sub TestStatusMessages()
        Try
            Console.WriteLine("Testing Status Messages...")
            Console.WriteLine("Showing info status...")
            uiManager.ShowStatusMessage("Information: Demo status message")
            System.Threading.Thread.Sleep(2000)
            Console.WriteLine("Showing prompt status...")
            uiManager.ShowStatusMessage("Prompt: Ready for input", BCOM.MsdStatusBarArea.msdStatusBarAreaLeft)
            System.Threading.Thread.Sleep(2000)
            Console.WriteLine("Showing command status...")
            uiManager.ShowStatusMessage("Command: Demo active", BCOM.MsdStatusBarArea.msdStatusBarAreaMiddle)
            System.Threading.Thread.Sleep(2000)
            Console.WriteLine("Status message test completed")
        Catch ex As Exception
            Console.WriteLine($"Error testing status messages: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Tests context-sensitive help functionality
    ''' </summary>
    Private Sub TestContextHelp()
        Try
            Console.WriteLine("Testing Context Help...")
            uiManager.ShowContextHelp("Getting Started")
            Console.WriteLine("Context help displayed - help panel should show 'Getting Started' topic")
        Catch ex As Exception
            Console.WriteLine($"Error testing context help: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Comprehensive test of all UI components working together
    ''' </summary>
    Private Sub TestComprehensiveIntegration()
        Try
            Console.WriteLine("=== Comprehensive Integration Test ===")
            ' Test sequence demonstrating professional workflow
            uiManager.ShowStatusMessage("Starting comprehensive UI test...")
            System.Threading.Thread.Sleep(1000)
            ' Show information about the test
            uiManager.ShowInformationMessage("Integration Test", 
                "This test demonstrates all UI components working together professionally.")
            ' Show help panel
            uiManager.ShowHelpPanel()
            System.Threading.Thread.Sleep(2000)
            ' Show about dialog
            uiManager.ShowAboutDialog()
            ' Hide help panel
            uiManager.HideHelpPanel()
            ' Final status
            uiManager.ShowStatusMessage("Comprehensive integration test completed successfully")
            Console.WriteLine("✓ Comprehensive integration test completed")
        Catch ex As Exception
            Console.WriteLine($"✗ Comprehensive integration test failed: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Cleans up resources and connections
    ''' </summary>
    Private Sub CleanupDemo()
        Try
            Console.WriteLine("Cleaning up demo resources...")
            If uiManager IsNot Nothing Then
                uiManager.Dispose()
                uiManager = Nothing
            End If
            If ustation IsNot Nothing Then
                MsCom.ReleaseComObject(ustation)
                ustation = Nothing
            End If
            Console.WriteLine("✓ Cleanup completed successfully")
        Catch ex As Exception
            Console.WriteLine($"Error during cleanup: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Entry point for VBA integration - can be called directly from MicroStation VBA
    ''' </summary>
    ''' <remarks>
    ''' This method provides a simplified entry point for VBA users who want to
    ''' integrate the professional UI components into their existing workflows
    ''' </remarks>
    Public Shared Sub InitializeSSMUI()
        Try
            ' Connect to current MicroStation session
            Dim ustation As BCOM.Application = MsCom.Attach(True)
            ' Initialize UI Manager
            Dim uiManager As New UIManager(ustation)
            ' Show welcome message
            uiManager.ShowInformationMessage("SSM Tools Initialized", 
                "Professional UI components are now available." & vbCrLf & 
                "Access Help panel and About dialog through the UI Manager.")
            ' Optionally show help panel
            uiManager.ShowHelpPanel()
        Catch ex As Exception
            MessageBox.Show($"Error initializing SSM UI: {ex.Message}", "Initialization Error", 
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class