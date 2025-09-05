Imports System.Windows.Forms
Imports System.Drawing
Imports System.Diagnostics
Imports BCOM = MicroStationDGN
''' <summary>
''' Professional Help System panel for SSM MicroStation V8i Library
''' Provides integrated help documentation, tutorials, and quick access links
''' Implements as a dockable MicroStation panel using native window APIs
''' </summary>
''' <remarks>
''' This panel integrates with MicroStation's docking system and provides
''' professional help navigation, online resources, and context-sensitive help
''' </remarks>
Public Class HelpSystemPanel
    Inherits UserControl
    Private ustation As BCOM.Application
    Private mdlLibrary As BCOM._MdlLibrary
    Private msWindow As Integer = 0
    Private isInitialized As Boolean = False
    ' UI Controls
    Private WithEvents tvHelpTopics As TreeView
    Private WithEvents wbHelpContent As WebBrowser
    Private WithEvents pnlNavigation As Panel
    Private WithEvents pnlContent As Panel
    Private WithEvents splitter As Splitter
    Private WithEvents btnHome As Button
    Private WithEvents btnBack As Button
    Private WithEvents btnForward As Button
    Private WithEvents btnRefresh As Button
    Private WithEvents btnPrint As Button
    Private WithEvents lblTitle As Label
    ''' <summary>
    ''' Initializes the Help System panel with MicroStation integration
    ''' </summary>
    ''' <param name="msApp">Reference to MicroStation Application object</param>
    Public Sub New(msApp As BCOM.Application)
        ustation = msApp
        mdlLibrary = ustation.MdlLibrary
        InitializeComponent()
        SetupHelpTopics()
        CreateMicroStationWindow()
    End Sub
    ''' <summary>
    ''' Sets up the panel layout and controls with professional appearance
    ''' </summary>
    Private Sub InitializeComponent()
        Me.Size = New Size(400, 600)
        Me.BackColor = Color.FromArgb(240, 240, 240)
        ' Title bar
        lblTitle = New Label()
        lblTitle.Text = "SSM MicroStation V8i Tools - Help"
        lblTitle.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblTitle.BackColor = Color.FromArgb(0, 120, 215)
        lblTitle.ForeColor = Color.White
        lblTitle.TextAlign = ContentAlignment.MiddleLeft
        lblTitle.Padding = New Padding(10, 0, 0, 0)
        lblTitle.Dock = DockStyle.Top
        lblTitle.Height = 30
        Me.Controls.Add(lblTitle)
        ' Navigation toolbar
        Dim toolStrip As New ToolStrip()
        toolStrip.BackColor = Color.FromArgb(250, 250, 250)
        toolStrip.GripStyle = ToolStripGripStyle.Hidden
        btnBack = New Button()
        btnBack.Text = "◄"
        btnBack.Size = New Size(30, 25)
        btnBack.FlatStyle = FlatStyle.Flat
        btnBack.BackColor = Color.Transparent
        btnBack.Enabled = False
        btnForward = New Button()
        btnForward.Text = "►"
        btnForward.Size = New Size(30, 25)
        btnForward.FlatStyle = FlatStyle.Flat
        btnForward.BackColor = Color.Transparent
        btnForward.Enabled = False
        btnHome = New Button()
        btnHome.Text = "🏠"
        btnHome.Size = New Size(30, 25)
        btnHome.FlatStyle = FlatStyle.Flat
        btnHome.BackColor = Color.Transparent
        btnRefresh = New Button()
        btnRefresh.Text = "⟳"
        btnRefresh.Size = New Size(30, 25)
        btnRefresh.FlatStyle = FlatStyle.Flat
        btnRefresh.BackColor = Color.Transparent
        btnPrint = New Button()
        btnPrint.Text = "🖨"
        btnPrint.Size = New Size(30, 25)
        btnPrint.FlatStyle = FlatStyle.Flat
        btnPrint.BackColor = Color.Transparent
        Dim toolbarPanel As New Panel()
        toolbarPanel.Height = 35
        toolbarPanel.Dock = DockStyle.Top
        toolbarPanel.BackColor = Color.FromArgb(250, 250, 250)
        toolbarPanel.Padding = New Padding(5)
        toolbarPanel.Controls.AddRange({btnBack, btnForward, btnHome, btnRefresh, btnPrint})
        btnForward.Left = btnBack.Right + 2
        btnHome.Left = btnForward.Right + 10
        btnRefresh.Left = btnHome.Right + 2
        btnPrint.Left = btnRefresh.Right + 2
        Me.Controls.Add(toolbarPanel)
        ' Navigation panel
        pnlNavigation = New Panel()
        pnlNavigation.Width = 180
        pnlNavigation.Dock = DockStyle.Left
        pnlNavigation.BackColor = Color.White
        pnlNavigation.BorderStyle = BorderStyle.FixedSingle
        ' Help topics tree
        tvHelpTopics = New TreeView()
        tvHelpTopics.Dock = DockStyle.Fill
        tvHelpTopics.BorderStyle = BorderStyle.None
        tvHelpTopics.ShowLines = True
        tvHelpTopics.ShowPlusMinus = True
        tvHelpTopics.ShowRootLines = True
        tvHelpTopics.FullRowSelect = True
        tvHelpTopics.HideSelection = False
        pnlNavigation.Controls.Add(tvHelpTopics)
        Me.Controls.Add(pnlNavigation)
        ' Splitter
        splitter = New Splitter()
        splitter.Dock = DockStyle.Left
        splitter.Width = 3
        splitter.BackColor = Color.Gray
        Me.Controls.Add(splitter)
        ' Content panel
        pnlContent = New Panel()
        pnlContent.Dock = DockStyle.Fill
        pnlContent.BackColor = Color.White
        ' Web browser for content
        wbHelpContent = New WebBrowser()
        wbHelpContent.Dock = DockStyle.Fill
        wbHelpContent.ScriptErrorsSuppressed = True
        pnlContent.Controls.Add(wbHelpContent)
        Me.Controls.Add(pnlContent)
        ' Load home page
        LoadHomePage()
    End Sub
    ''' <summary>
    ''' Sets up the help topics tree with organized content structure
    ''' </summary>
    Private Sub SetupHelpTopics()
        Try
            tvHelpTopics.Nodes.Clear()
            ' Getting Started
            Dim nodeGettingStarted = tvHelpTopics.Nodes.Add("Getting Started")
            nodeGettingStarted.Nodes.Add("Installation Guide")
            nodeGettingStarted.Nodes.Add("Quick Start Tutorial")
            nodeGettingStarted.Nodes.Add("System Requirements")
            nodeGettingStarted.Nodes.Add("First Steps")
            ' API Reference
            Dim nodeApiRef = tvHelpTopics.Nodes.Add("API Reference")
            nodeApiRef.Nodes.Add("Geometry Library")
            nodeApiRef.Nodes.Add("2D Drawing Functions")
            nodeApiRef.Nodes.Add("3D Modeling Functions")
            nodeApiRef.Nodes.Add("Export Functions")
            nodeApiRef.Nodes.Add("Utility Functions")
            ' User Guide
            Dim nodeUserGuide = tvHelpTopics.Nodes.Add("User Guide")
            nodeUserGuide.Nodes.Add("Creating Geometry")
            nodeUserGuide.Nodes.Add("Exporting Models")
            nodeUserGuide.Nodes.Add("Advanced Features")
            nodeUserGuide.Nodes.Add("Best Practices")
            nodeUserGuide.Nodes.Add("Troubleshooting")
            ' Examples & Tutorials
            Dim nodeExamples = tvHelpTopics.Nodes.Add("Examples & Tutorials")
            nodeExamples.Nodes.Add("Basic 2D Drawing")
            nodeExamples.Nodes.Add("3D Solid Modeling")
            nodeExamples.Nodes.Add("Complex Assemblies")
            nodeExamples.Nodes.Add("Automation Scripts")
            nodeExamples.Nodes.Add("Integration Examples")
            ' Technical Support
            Dim nodeSupport = tvHelpTopics.Nodes.Add("Technical Support")
            nodeSupport.Nodes.Add("FAQ")
            nodeSupport.Nodes.Add("Known Issues")
            nodeSupport.Nodes.Add("Contact Support")
            nodeSupport.Nodes.Add("Report Bug")
            nodeSupport.Nodes.Add("Feature Request")
            ' Online Resources
            Dim nodeOnline = tvHelpTopics.Nodes.Add("Online Resources")
            nodeOnline.Nodes.Add("Documentation Website")
            nodeOnline.Nodes.Add("Video Tutorials")
            nodeOnline.Nodes.Add("Community Forum")
            nodeOnline.Nodes.Add("Download Updates")
            nodeOnline.Nodes.Add("Release Notes")
            ' Expand Getting Started by default
            nodeGettingStarted.Expand()
        Catch ex As Exception
            ' Handle errors gracefully
            MessageBox.Show($"Error setting up help topics: {ex.Message}", "Error", 
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
    ''' <summary>
    ''' Creates and integrates the panel with MicroStation's native window system
    ''' </summary>
    Private Sub CreateMicroStationWindow()
        Try
            If mdlLibrary Is Nothing Then
                Return
            End If
            ' Initialize MicroStation native window system
            Dim initResult = mdlLibrary.mdlNativeWindow_initialize("SSMHelpPanel")
            If initResult <> 0 Then
                Throw New Exception("Failed to initialize MicroStation native window system")
            End If
            ' Create dockable window
            Dim windowResult = mdlLibrary.mdlNativeWindow_getDockableWindow(
                msWindow,
                1001, ' Unique window ID
                "SSM Tools - Help",
                0, ' Hook function (not used in .NET)
                0, ' Interest events (not used in .NET)
                "SSM.Help.Position" ' User preference alias for window position
            )
            If windowResult = 0 AndAlso msWindow <> 0 Then
                ' Set this UserControl as the window content
                Dim contentResult = mdlLibrary.mdlNativeWindow_setAsContent(
                    Me.Handle.ToInt32(),
                    0, ' Dialog pointer (not used)
                    msWindow
                )
                If contentResult = 0 Then
                    ' Attach content and size to fit
                    mdlLibrary.mdlNativeWindow_attachContent(msWindow, 1)
                    ' Add to MicroStation's window list for proper management
                    mdlLibrary.mdlNativeWindow_addToWindowList(msWindow)
                    isInitialized = True
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"Error creating MicroStation window: {ex.Message}", "Error", 
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' Loads the home page content with welcome information and navigation
    ''' </summary>
    Private Sub LoadHomePage()
        Dim homeContent As String = GenerateHomePageHTML()
        wbHelpContent.DocumentText = homeContent
    End Sub
    ''' <summary>
    ''' Generates HTML content for the help system home page
    ''' </summary>
    ''' <returns>HTML string for home page</returns>
    Private Function GenerateHomePageHTML() As String
        Return $"
<!DOCTYPE html>
<html>
<head>
    <title>SSM MicroStation V8i Tools - Help</title>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background-color: #fafafa; }}
        .header {{ background: linear-gradient(135deg, #0078d4, #106ebe); color: white; padding: 20px; margin: -20px -20px 20px -20px; }}
        .header h1 {{ margin: 0; font-size: 24px; }}
        .header p {{ margin: 5px 0 0 0; opacity: 0.9; }}
        .section {{ background: white; padding: 20px; margin-bottom: 20px; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
        .section h2 {{ color: #0078d4; margin-top: 0; }}
        .link-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-top: 15px; }}
        .quick-link {{ background: #f0f8ff; padding: 15px; border-radius: 5px; text-decoration: none; color: #0066cc; border-left: 4px solid #0078d4; }}
        .quick-link:hover {{ background: #e6f2ff; }}
        .quick-link h3 {{ margin: 0 0 5px 0; font-size: 16px; }}
        .quick-link p {{ margin: 0; font-size: 14px; color: #666; }}
        .version-info {{ background: #f8f9fa; border: 1px solid #dee2e6; padding: 15px; border-radius: 5px; margin-top: 20px; }}
    </style>
</head>
<body>
    <div class='header'>
        <h1>Welcome to SSM MicroStation V8i Tools</h1>
        <p>Professional 2D/3D CAD automation and geometry library</p>
    </div>
    <div class='section'>
        <h2>Getting Started</h2>
        <p>Welcome to the comprehensive help system for SSM MicroStation V8i Tools. This integrated help panel provides access to documentation, tutorials, and support resources.</p>
        <div class='link-grid'>
            <a href='#' class='quick-link' onclick='parent.selectHelpTopic(""Installation Guide"")'>
                <h3>📦 Installation Guide</h3>
                <p>Step-by-step installation instructions</p>
            </a>
            <a href='#' class='quick-link' onclick='parent.selectHelpTopic(""Quick Start Tutorial"")'>
                <h3>🚀 Quick Start</h3>
                <p>Get up and running in minutes</p>
            </a>
            <a href='#' class='quick-link' onclick='parent.selectHelpTopic(""API Reference"")'>
                <h3>📖 API Reference</h3>
                <p>Complete function documentation</p>
            </a>
            <a href='#' class='quick-link' onclick='parent.selectHelpTopic(""Examples & Tutorials"")'>
                <h3>💡 Examples</h3>
                <p>Code examples and tutorials</p>
            </a>
        </div>
    </div>
    <div class='section'>
        <h2>Key Features</h2>
        <ul>
            <li><strong>Comprehensive 2D/3D Geometry Library</strong> - Create all standard CAD geometry types</li>
            <li><strong>Multi-format Export</strong> - Export to STL, STEP, IGES, DXF, PDF, and more</li>
            <li><strong>MicroStation Integration</strong> - Native integration with MicroStation V8i</li>
            <li><strong>Professional Tools</strong> - Advanced automation and workflow tools</li>
            <li><strong>.NET Framework Support</strong> - Modern .NET development capabilities</li>
        </ul>
    </div>
    <div class='section'>
        <h2>Support & Resources</h2>
        <p>Need help? Check out these resources:</p>
        <ul>
            <li><a href='#' onclick='parent.openOnlineResource(""docs"")'>📄 Online Documentation</a></li>
            <li><a href='#' onclick='parent.openOnlineResource(""forum"")'>💬 Community Forum</a></li>
            <li><a href='#' onclick='parent.openOnlineResource(""support"")'>🔧 Technical Support</a></li>
            <li><a href='#' onclick='parent.selectHelpTopic(""FAQ"")'>❓ Frequently Asked Questions</a></li>
        </ul>
    </div>
    <div class='version-info'>
        <strong>Version:</strong> 1.0.0 | <strong>Build:</strong> {DateTime.Now:yyyyMMdd} | <strong>Framework:</strong> .NET 4.8.1<br>
        <strong>MicroStation:</strong> V8i | <strong>Last Updated:</strong> {DateTime.Now:MMMM yyyy}
    </div>
</body>
</html>"
    End Function
    ''' <summary>
    ''' Handles help topic selection from the tree view
    ''' </summary>
    Private Sub tvHelpTopics_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvHelpTopics.AfterSelect
        If e.Node IsNot Nothing Then
            LoadHelpContent(e.Node.Text)
        End If
    End Sub
    ''' <summary>
    ''' Loads help content based on the selected topic
    ''' </summary>
    ''' <param name="topic">The help topic to load</param>
    Private Sub LoadHelpContent(topic As String)
        Try
            Dim content As String = GenerateTopicContent(topic)
            wbHelpContent.DocumentText = content
        Catch ex As Exception
            wbHelpContent.DocumentText = $"<html><body><h2>Error Loading Content</h2><p>Error loading help topic '{topic}': {ex.Message}</p></body></html>"
        End Try
    End Sub
    ''' <summary>
    ''' Generates HTML content for specific help topics
    ''' </summary>
    ''' <param name="topic">The help topic name</param>
    ''' <returns>HTML content for the topic</returns>
    Private Function GenerateTopicContent(topic As String) As String
        Select Case topic.ToLower()
            Case "installation guide"
                Return GenerateInstallationGuide()
            Case "api reference"
                Return GenerateAPIReference()
            Case "quick start tutorial"
                Return GenerateQuickStartTutorial()
            Case "faq"
                Return GenerateFAQ()
            Case Else
                Return GenerateGenericTopicContent(topic)
        End Select
    End Function
    ''' <summary>
    ''' Generates installation guide content
    ''' </summary>
    Private Function GenerateInstallationGuide() As String
        Return $"
<!DOCTYPE html>
<html>
<head>
    <title>Installation Guide</title>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; line-height: 1.6; }}
        .header {{ background: #0078d4; color: white; padding: 15px; margin: -20px -20px 20px -20px; }}
        .step {{ background: #f8f9fa; border-left: 4px solid #0078d4; padding: 15px; margin: 15px 0; }}
        .code {{ background: #f4f4f4; padding: 10px; border-radius: 3px; font-family: 'Consolas', monospace; }}
    </style>
</head>
<body>
    <div class='header'>
        <h1>📦 Installation Guide</h1>
        <p>Complete setup instructions for SSM MicroStation V8i Tools</p>
    </div>
    <div class='step'>
        <h3>Step 1: System Requirements</h3>
        <ul>
            <li>MicroStation V8i (08.11.xx.xx or later)</li>
            <li>.NET Framework 4.8.1 or later</li>
            <li>Windows 10/11 or Windows Server 2016+</li>
            <li>Administrator privileges for installation</li>
        </ul>
    </div>
    <div class='step'>
        <h3>Step 2: Download and Extract</h3>
        <p>1. Download the SSM.Msv8i.Tools package</p>
        <p>2. Extract to your preferred installation directory</p>
        <p>3. Ensure all files are unblocked (Right-click → Properties → Unblock)</p>
    </div>
    <div class='step'>
        <h3>Step 3: Register COM Components</h3>
        <p>Run as Administrator:</p>
        <div class='code'>
regasm SSM.Msv8i.Lib.dll /codebase<br>
regasm SSM.Msv8i.Tools.dll /codebase
        </div>
    </div>
    <div class='step'>
        <h3>Step 4: Configure MicroStation</h3>
        <p>1. Open MicroStation V8i</p>
        <p>2. Go to Utilities → Macro → VBA</p>
        <p>3. Load the SSM tools VBA project</p>
        <p>4. Verify all references are loaded correctly</p>
    </div>
    <div class='step'>
        <h3>Step 5: Verification</h3>
        <p>Test the installation by running a simple geometry creation script.</p>
    </div>
</body>
</html>"
    End Function
    ''' <summary>
    ''' Generates API reference content
    ''' </summary>
    Private Function GenerateAPIReference() As String
        Return $"
<!DOCTYPE html>
<html>
<head>
    <title>API Reference</title>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; line-height: 1.6; }}
        .header {{ background: #0078d4; color: white; padding: 15px; margin: -20px -20px 20px -20px; }}
        .api-section {{ background: white; border: 1px solid #dee2e6; margin: 15px 0; padding: 20px; border-radius: 5px; }}
        .method {{ background: #f8f9fa; padding: 10px; margin: 10px 0; border-radius: 3px; }}
        .code {{ background: #f4f4f4; padding: 10px; border-radius: 3px; font-family: 'Consolas', monospace; font-size: 12px; }}
    </style>
</head>
<body>
    <div class='header'>
        <h1>📖 API Reference</h1>
        <p>Complete function documentation for SSM MicroStation V8i Tools</p>
    </div>
    <div class='api-section'>
        <h2>Geometry Creation Functions</h2>
        <div class='method'>
            <h3>CreateLine</h3>
            <p>Creates a line element between two points.</p>
            <div class='code'>
Public Function CreateLine(startPoint As Point3d, endPoint As Point3d, Optional template As _Element = Nothing) As LineElement
            </div>
            <p><strong>Parameters:</strong></p>
            <ul>
                <li><code>startPoint</code> - Starting point of the line</li>
                <li><code>endPoint</code> - Ending point of the line</li>
                <li><code>template</code> - Optional template element for attributes</li>
            </ul>
        </div>
        <div class='method'>
            <h3>CreateCircle</h3>
            <p>Creates a circle element with specified center and radius.</p>
            <div class='code'>
Public Function CreateCircle(center As Point3d, radius As Double, Optional template As _Element = Nothing) As EllipseElement
            </div>
        </div>
        <div class='method'>
            <h3>CreateArc</h3>
            <p>Creates an arc element with center, radius, and angle parameters.</p>
            <div class='code'>
Public Function CreateArc(center As Point3d, radius As Double, startAngle As Double, sweepAngle As Double, Optional template As _Element = Nothing) As ArcElement
            </div>
        </div>
    </div>
    <div class='api-section'>
        <h2>Export Functions</h2>
        <p>Functions for exporting geometry to various CAD formats.</p>
        <div class='method'>
            <h3>ExportToSTL</h3>
            <p>Exports selected geometry to STL format for 3D printing.</p>
            <div class='code'>
Public Function ExportToSTL(filePath As String, Optional selection As ElementEnumerator = Nothing) As Boolean
            </div>
        </div>
    </div>
</body>
</html>"
    End Function
    ''' <summary>
    ''' Generates quick start tutorial content
    ''' </summary>
    Private Function GenerateQuickStartTutorial() As String
        Return "<!-- Quick Start Tutorial HTML content here -->"
    End Function
    ''' <summary>
    ''' Generates FAQ content
    ''' </summary>
    Private Function GenerateFAQ() As String
        Return "<!-- FAQ HTML content here -->"
    End Function
    ''' <summary>
    ''' Generates generic topic content for topics without specific implementations
    ''' </summary>
    ''' <param name="topic">Topic name</param>
    Private Function GenerateGenericTopicContent(topic As String) As String
        Return $"
<!DOCTYPE html>
<html>
<head>
    <title>{topic}</title>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; line-height: 1.6; }}
        .header {{ background: #0078d4; color: white; padding: 15px; margin: -20px -20px 20px -20px; }}
    </style>
</head>
<body>
    <div class='header'>
        <h1>{topic}</h1>
    </div>
    <p>Documentation for <strong>{topic}</strong> is currently being developed.</p>
    <p>Please check back later or contact support for more information.</p>
</body>
</html>"
    End Function
    ''' <summary>
    ''' Shows the help panel by making the MicroStation window visible
    ''' </summary>
    Public Sub ShowHelpPanel()
        Try
            If isInitialized AndAlso msWindow <> 0 Then
                ' Restore window if minimized and bring to front
                mdlLibrary.mdlNativeWindow_restore(msWindow)
                Me.Show()
            Else
                ' Try to recreate the window if it wasn't initialized
                CreateMicroStationWindow()
                If isInitialized Then
                    Me.Show()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"Error showing help panel: {ex.Message}", "Error", 
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' Hides the help panel
    ''' </summary>
    Public Sub HideHelpPanel()
        Try
            If isInitialized AndAlso msWindow <> 0 Then
                mdlLibrary.mdlNativeWindow_minimize(msWindow)
            End If
            Me.Hide()
        Catch ex As Exception
            ' Handle errors gracefully
        End Try
    End Sub
    ''' <summary>
    ''' Toolbar button event handlers
    ''' </summary>
    Private Sub btnHome_Click(sender As Object, e As EventArgs) Handles btnHome.Click
        LoadHomePage()
        tvHelpTopics.SelectedNode = Nothing
    End Sub
    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        If tvHelpTopics.SelectedNode IsNot Nothing Then
            LoadHelpContent(tvHelpTopics.SelectedNode.Text)
        Else
            LoadHomePage()
        End If
    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            wbHelpContent.ShowPrintDialog()
        Catch ex As Exception
            MessageBox.Show($"Printing error: {ex.Message}", "Print Error", 
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
    ''' <summary>
    ''' Cleanup resources when the panel is disposed
    ''' </summary>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso isInitialized AndAlso msWindow <> 0 Then
                ' Remove from window list and destroy MicroStation window
                mdlLibrary.mdlNativeWindow_removeFromWindowList(msWindow)
                mdlLibrary.mdlNativeWindow_destroyMSWindow(msWindow, True)
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
End Class