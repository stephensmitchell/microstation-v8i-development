Imports System
Imports System.Runtime.InteropServices
Imports MicroStationDGN
''' <summary>
''' Main program demonstrating the MicroStation V8i Geometry Library
''' This program shows how to use all the geometry creation and export functions
''' </summary>
Module Program
    Sub Main()
        Console.WriteLine("MicroStation V8i Geometry Library Demo")
        Console.WriteLine("=====================================")
        Try
            ' Get MicroStation Application
            Dim ustation As Application = GetMicroStationApplication()
            If ustation Is Nothing Then
                Console.WriteLine("Could not connect to MicroStation. Please ensure MicroStation V8i is running.")
                Console.ReadKey()
                Return
            End If
            Console.WriteLine("Connected to MicroStation successfully!")
            Console.WriteLine()
            ' Create geometry library instance
            Dim geometryLib As New GeometryLibrary(ustation)
            ' Create test application instance  
            Dim testApp As New GeometryLibraryTestApp(ustation)
            ' Show menu
            ShowMenu(geometryLib, testApp)
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
            Console.WriteLine("Stack Trace: " & ex.StackTrace)
        End Try
        Console.WriteLine()
        Console.WriteLine("Press any key to exit...")
        Console.ReadKey()
    End Sub
    ''' <summary>
    ''' Gets a reference to the running MicroStation application
    ''' </summary>
    Private Function GetMicroStationApplication() As Application
        Try
            ' Try to get existing MicroStation instance
            Return CType(Marshal.GetActiveObject("MicroStationDGN.Application"), Application)
        Catch ex As Exception
            Console.WriteLine("Could not get MicroStation application: " & ex.Message)
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' Shows the interactive menu
    ''' </summary>
    Private Sub ShowMenu(geometryLib As GeometryLibrary, testApp As GeometryLibraryTestApp)
        Dim exitRequested As Boolean = False
        While Not exitRequested
            Console.Clear()
            Console.WriteLine("MicroStation V8i Geometry Library - Main Menu")
            Console.WriteLine("============================================")
            Console.WriteLine()
            Console.WriteLine("1. Run All Tests")
            Console.WriteLine("2. Create Geometry Showcase")
            Console.WriteLine("3. Create Basic 2D Shapes")
            Console.WriteLine("4. Create Complex Shapes") 
            Console.WriteLine("5. Create 3D Geometry")
            Console.WriteLine("6. Test Export Functions")
            Console.WriteLine("7. Show Library Information")
            Console.WriteLine("0. Exit")
            Console.WriteLine()
            Console.Write("Select an option (0-7): ")
            Dim input As String = Console.ReadLine()
            Console.WriteLine()
            Select Case input
                Case "1"
                    Console.WriteLine("Running all tests...")
                    testApp.RunAllTests()
                Case "2"
                    Console.WriteLine("Creating geometry showcase...")
                    testApp.CreateGeometryShowcase()
                    Console.WriteLine("Geometry showcase created in active model!")
                Case "3"
                    CreateBasic2DShapes(geometryLib)
                Case "4"
                    CreateComplexShapes(geometryLib)
                Case "5"
                    Create3DGeometry(geometryLib)
                Case "6"
                    TestExportFunctions(geometryLib)
                Case "7"
                    ShowLibraryInformation()
                Case "0"
                    exitRequested = True
                Case Else
                    Console.WriteLine("Invalid option. Please try again.")
            End Select
            If Not exitRequested Then
                Console.WriteLine()
                Console.WriteLine("Press any key to return to menu...")
                Console.ReadKey()
            End If
        End While
    End Sub
    ''' <summary>
    ''' Creates basic 2D shapes demonstration
    ''' </summary>
    Private Sub CreateBasic2DShapes(geometryLib As GeometryLibrary)
        Console.WriteLine("Creating basic 2D shapes...")
        ' Create a line
        Dim line As LineElement = geometryLib.CreateLineFromCoordinates(0, 0, 100, 0)
        geometryLib.AddElementToModel(line)
        Console.WriteLine("✓ Created horizontal line")
        ' Create a circle
        Dim circle As EllipseElement = geometryLib.CreateCircleFromCoordinates(150, 50, 30)
        geometryLib.AddElementToModel(circle)
        Console.WriteLine("✓ Created circle")
        ' Create a rectangle
        Dim rect As ShapeElement = geometryLib.CreateRectangleFromCoordinates(200, 20, 280, 80)
        geometryLib.AddElementToModel(rect)
        Console.WriteLine("✓ Created rectangle")
        ' Create an arc
        Dim arc As EllipseElement = geometryLib.CreateArcFromCoordinates(350, 50, 25, 0, Math.PI)
        geometryLib.AddElementToModel(arc)
        Console.WriteLine("✓ Created arc")
        ' Add text labels
        Dim lineLabel As TextElement = geometryLib.CreateTextFromCoordinates("Line", 20, -20)
        Dim circleLabel As TextElement = geometryLib.CreateTextFromCoordinates("Circle", 130, 10)
        Dim rectLabel As TextElement = geometryLib.CreateTextFromCoordinates("Rectangle", 220, 0)
        Dim arcLabel As TextElement = geometryLib.CreateTextFromCoordinates("Arc", 335, 10)
        geometryLib.AddElementsToModel(New _Element() {lineLabel, circleLabel, rectLabel, arcLabel})
        Console.WriteLine("✓ Added text labels")
        Console.WriteLine("Basic 2D shapes created successfully!")
    End Sub
    ''' <summary>
    ''' Creates complex shapes demonstration
    ''' </summary>
    Private Sub CreateComplexShapes(geometryLib As GeometryLibrary)
        Console.WriteLine("Creating complex shapes...")
        ' Create regular polygons
        Dim pentagon As ComplexShapeElement = geometryLib.CreateRegularPolygon(
            geometryLib.CreatePoint3d(50, 150, 0), 30, 5)
        geometryLib.AddElementToModel(pentagon)
        Console.WriteLine("✓ Created pentagon")
        Dim hexagon As ComplexShapeElement = geometryLib.CreateRegularPolygon(
            geometryLib.CreatePoint3d(150, 150, 0), 30, 6)
        geometryLib.AddElementToModel(hexagon)
        Console.WriteLine("✓ Created hexagon")
        Dim octagon As ComplexShapeElement = geometryLib.CreateRegularPolygon(
            geometryLib.CreatePoint3d(250, 150, 0), 30, 8)
        geometryLib.AddElementToModel(octagon)
        Console.WriteLine("✓ Created octagon")
        ' Create a star shape
        Dim starPoints(9) As Point3d
        Dim centerX As Double = 350
        Dim centerY As Double = 150
        For i As Integer = 0 To 9
            Dim angle As Double = i * Math.PI / 5
            Dim radius As Double = If(i Mod 2 = 0, 30, 15)
            starPoints(i) = geometryLib.CreatePoint3d(
                centerX + radius * Math.Cos(angle),
                centerY + radius * Math.Sin(angle),
                0)
        Next
        Dim star As ShapeElement = geometryLib.CreatePolygon(starPoints)
        geometryLib.AddElementToModel(star)
        Console.WriteLine("✓ Created star shape")
        Console.WriteLine("Complex shapes created successfully!")
    End Sub
    ''' <summary>
    ''' Creates 3D geometry demonstration
    ''' </summary>
    Private Sub Create3DGeometry(geometryLib As GeometryLibrary)
        Console.WriteLine("Creating 3D geometry...")
        ' Create 3D lines
        Dim line3d1 As LineElement = geometryLib.CreateLine3D(
            geometryLib.CreatePoint3d(0, 250, 0),
            geometryLib.CreatePoint3d(50, 250, 25))
        geometryLib.AddElementToModel(line3d1)
        Dim line3d2 As LineElement = geometryLib.CreateLine3D(
            geometryLib.CreatePoint3d(50, 250, 25),
            geometryLib.CreatePoint3d(100, 250, 0))
        geometryLib.AddElementToModel(line3d2)
        Console.WriteLine("✓ Created 3D lines")
        ' Create 3D points at different elevations
        For i As Integer = 0 To 5
            Dim point3d As PointStringElement = geometryLib.CreatePointFromCoordinates(
                150 + i * 20, 250, i * 8)
            geometryLib.AddElementToModel(point3d)
        Next
        Console.WriteLine("✓ Created 3D points at different elevations")
        Console.WriteLine("3D geometry created successfully!")
    End Sub
    ''' <summary>
    ''' Tests export functions
    ''' </summary>
    Private Sub TestExportFunctions(geometryLib As GeometryLibrary)
        Console.WriteLine("Testing export functions...")
        Try
            ' Get ustation reference from geometry library
            Dim ustationApp As Application = CType(Marshal.GetActiveObject("MicroStationDGN.Application"), Application)
            Dim exporter As New GeometryExporter(ustationApp)
            Dim options As GeometryExporter.ExportOptions = exporter.CreateDefaultExportOptions()
            ' Test directory
            Dim testDir As String = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), "GeometryLibraryTest")
            If Not System.IO.Directory.Exists(testDir) Then
                System.IO.Directory.CreateDirectory(testDir)
            End If
            Console.WriteLine("Export test directory: " & testDir)
            ' Test single format export
            Dim testFile As String = System.IO.Path.Combine(testDir, "test_geometry.dxf")
            Dim success As Boolean = exporter.ExportModel(testFile, GeometryExporter.ExportFormat.DXF, options)
            If success Then
                Console.WriteLine("✓ DXF export test successful")
            Else
                Console.WriteLine("✗ DXF export test failed")
            End If
            ' Show export statistics
            Dim stats = exporter.GetExportStatistics()
            Console.WriteLine("Export Statistics:")
            For Each kvp In stats
                Console.WriteLine("  " & kvp.Key & ": " & kvp.Value.ToString())
            Next
        Catch ex As Exception
            Console.WriteLine("Export test error: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Shows information about the geometry library
    ''' </summary>
    Private Sub ShowLibraryInformation()
        Console.WriteLine("MicroStation V8i Geometry Library Information")
        Console.WriteLine("===========================================")
        Console.WriteLine()
        Console.WriteLine("Library Features:")
        Console.WriteLine("• Complete 2D geometry creation (lines, circles, arcs, polygons)")
        Console.WriteLine("• Complex shape support (regular polygons, custom shapes)")
        Console.WriteLine("• 3D geometry support (3D lines, points, surfaces)")
        Console.WriteLine("• Curve and spline support (B-splines, interpolation curves)")
        Console.WriteLine("• Text and annotation support")
        Console.WriteLine("• Dimension creation")
        Console.WriteLine("• Element property management")
        Console.WriteLine("• Geometric calculations and utilities")
        Console.WriteLine("• Comprehensive export functionality")
        Console.WriteLine("• Multiple export formats (STL, STEP, IGES, DXF, etc.)")
        Console.WriteLine()
        Console.WriteLine("API Compatibility:")
        Console.WriteLine("• Built for MicroStation V8i (Version 8)")
        Console.WriteLine("• Uses Bentley MicroStationDGN COM API")
        Console.WriteLine("• .NET Framework 4.8.1 compatible")
        Console.WriteLine("• VB.NET implementation")
        Console.WriteLine()
        Console.WriteLine("Classes:")
        Console.WriteLine("• GeometryLibrary - Main geometry creation functions")
        Console.WriteLine("• GeometryExporter - Export functionality")
        Console.WriteLine("• GeometryLibraryTestApp - Comprehensive test suite")
        Console.WriteLine()
    End Sub
End Module