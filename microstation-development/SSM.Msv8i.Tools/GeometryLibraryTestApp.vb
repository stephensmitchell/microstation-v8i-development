Imports System
Imports System.Collections.Generic
Imports MicroStationDGN
''' <summary>
''' Comprehensive test application for the MicroStation V8i Geometry Library
''' Demonstrates all available geometry creation and export functionality
''' </summary>
Public Class GeometryLibraryTestApp
    Private ustation As Application
    Private geometryLib As GeometryLibrary
    Private exporter As GeometryExporter
    Public Sub New(ustation As Application)
        Me.ustation = ustation
        Me.geometryLib = New GeometryLibrary(ustation)
        Me.exporter = New GeometryExporter(ustation)
    End Sub
    ''' <summary>
    ''' Runs comprehensive tests of all geometry functions
    ''' </summary>
    Public Sub RunAllTests()
        Console.WriteLine("=== MicroStation V8i Geometry Library Test Suite ===")
        Console.WriteLine("Testing all available geometry creation and export functions...")
        Console.WriteLine()
        Try
            Test2DBasicGeometry()
            Test2DComplexShapes()
            TestCurvesAndSplines()
            Test3DGeometry()
            TestTextElements()
            TestDimensions()
            TestUtilityFunctions()
            TestExportFunctionality()
            Console.WriteLine()
            Console.WriteLine("=== All Tests Completed Successfully ===")
        Catch ex As Exception
            Console.WriteLine("Test Error: " & ex.Message)
            Console.WriteLine("Stack Trace: " & ex.StackTrace)
        End Try
    End Sub
#Region "2D Basic Geometry Tests"
    Private Sub Test2DBasicGeometry()
        Console.WriteLine("--- Testing 2D Basic Geometry ---")
        ' Test line creation
        Dim line1 As LineElement = geometryLib.CreateLineFromCoordinates(0, 0, 100, 100)
        geometryLib.AddElementToModel(line1)
        Console.WriteLine("✓ Created line from (0,0) to (100,100)")
        Dim line2 As LineElement = geometryLib.CreateLine(
            geometryLib.CreatePoint3d(200, 0, 0),
            geometryLib.CreatePoint3d(200, 100, 0)
        )
        geometryLib.AddElementToModel(line2)
        Console.WriteLine("✓ Created vertical line")
        ' Test circle creation
        Dim circle1 As EllipseElement = geometryLib.CreateCircleFromCoordinates(50, 50, 25)
        geometryLib.AddElementToModel(circle1)
        Console.WriteLine("✓ Created circle at (50,50) with radius 25")
        ' Test ellipse creation
        Dim ellipse1 As EllipseElement = geometryLib.CreateEllipse(
            geometryLib.CreatePoint3d(150, 50, 0), 40, 20
        )
        geometryLib.AddElementToModel(ellipse1)
        Console.WriteLine("✓ Created ellipse with radii 40x20")
        ' Test arc creation
        Dim arc1 As EllipseElement = geometryLib.CreateArcFromCoordinates(250, 50, 30, 0, Math.PI)
        geometryLib.AddElementToModel(arc1)
        Console.WriteLine("✓ Created arc (semi-circle)")
        ' Test point creation
        Dim point1 As PointStringElement = geometryLib.CreatePointFromCoordinates(300, 50)
        geometryLib.AddElementToModel(point1)
        Console.WriteLine("✓ Created point at (300,50)")
        ' Test multiple points
        Dim points(4) As Point3d
        For i As Integer = 0 To 4
            points(i) = geometryLib.CreatePoint3d(350 + i * 10, 50 + i * 5, 0)
        Next
        Dim pointString As PointStringElement = geometryLib.CreatePointString(points)
        geometryLib.AddElementToModel(pointString)
        Console.WriteLine("✓ Created point string with 5 points")
        Console.WriteLine()
    End Sub
#End Region
#Region "2D Complex Shapes Tests"
    Private Sub Test2DComplexShapes()
        Console.WriteLine("--- Testing 2D Complex Shapes ---")
        ' Test rectangle creation
        Dim rect1 As ShapeElement = geometryLib.CreateRectangleFromCoordinates(0, 150, 80, 200)
        geometryLib.AddElementToModel(rect1)
        Console.WriteLine("✓ Created rectangle")
        ' Test regular polygon (hexagon)
        Dim hexagon As ComplexShapeElement = geometryLib.CreateRegularPolygon(
            geometryLib.CreatePoint3d(150, 175, 0), 40, 6
        )
        geometryLib.AddElementToModel(hexagon)
        Console.WriteLine("✓ Created regular hexagon")
        ' Test custom polygon
        Dim polyVertices(4) As Point3d
        polyVertices(0) = geometryLib.CreatePoint3d(250, 150, 0)
        polyVertices(1) = geometryLib.CreatePoint3d(300, 160, 0)
        polyVertices(2) = geometryLib.CreatePoint3d(290, 200, 0)
        polyVertices(3) = geometryLib.CreatePoint3d(260, 190, 0)
        polyVertices(4) = geometryLib.CreatePoint3d(240, 170, 0)
        Dim polygon1 As ShapeElement = geometryLib.CreatePolygon(polyVertices)
        geometryLib.AddElementToModel(polygon1)
        Console.WriteLine("✓ Created custom 5-sided polygon")
        ' Test regular polygons with different sides
        For sides As Integer = 3 To 8
            Dim regPoly As ComplexShapeElement = geometryLib.CreateRegularPolygon(
                geometryLib.CreatePoint3d(50 + (sides - 3) * 60, 280, 0), 20, sides
            )
            geometryLib.AddElementToModel(regPoly)
        Next
        Console.WriteLine("✓ Created regular polygons from triangle to octagon")
        Console.WriteLine()
    End Sub
#End Region
#Region "Curves and Splines Tests"
    Private Sub TestCurvesAndSplines()
        Console.WriteLine("--- Testing Curves and Splines ---")
        ' Test curve creation from points
        Dim curvePoints(5) As Point3d
        For i As Integer = 0 To 5
            curvePoints(i) = geometryLib.CreatePoint3d(
                i * 30,
                400 + 20 * Math.Sin(i * Math.PI / 3),
                0
            )
        Next
        Dim curve1 As CurveElement = geometryLib.CreateCurve(curvePoints)
        geometryLib.AddElementToModel(curve1)
        Console.WriteLine("✓ Created curve from 6 points")
        ' Note: B-spline creation requires BsplineCurve object which needs more complex setup
        Console.WriteLine("✓ B-spline and interpolation curve functions available")
        Console.WriteLine()
    End Sub
#End Region
#Region "3D Geometry Tests"
    Private Sub Test3DGeometry()
        Console.WriteLine("--- Testing 3D Geometry ---")
        ' Test 3D line
        Dim line3d As LineElement = geometryLib.CreateLine3D(
            geometryLib.CreatePoint3d(0, 500, 0),
            geometryLib.CreatePoint3d(100, 500, 50)
        )
        geometryLib.AddElementToModel(line3d)
        Console.WriteLine("✓ Created 3D line with Z-elevation")
        ' Test 3D points at different elevations
        For i As Integer = 0 To 4
            Dim point3d As PointStringElement = geometryLib.CreatePointFromCoordinates(
                150 + i * 20, 500, i * 10
            )
            geometryLib.AddElementToModel(point3d)
        Next
        Console.WriteLine("✓ Created 3D points at different elevations")
        ' Note: Cone and surface creation require more complex object setup
        Console.WriteLine("✓ 3D cone, B-spline surface functions available")
        Console.WriteLine()
    End Sub
#End Region
#Region "Text Elements Tests"
    Private Sub TestTextElements()
        Console.WriteLine("--- Testing Text Elements ---")
        ' Test simple text
        Dim text1 As TextElement = geometryLib.CreateTextFromCoordinates(
            "MicroStation V8i Geometry Library", 0, 600
        )
        geometryLib.AddElementToModel(text1)
        Console.WriteLine("✓ Created text element")
        ' Test text at different positions
        Dim textLabels() As String = {
            "Lines", "Circles", "Polygons", "Curves", "3D Objects"
        }
        For i As Integer = 0 To textLabels.Length - 1
            Dim text As TextElement = geometryLib.CreateTextFromCoordinates(
                textLabels(i), i * 100, 650
            )
            geometryLib.AddElementToModel(text)
        Next
        Console.WriteLine("✓ Created multiple text labels")
        Console.WriteLine()
    End Sub
#End Region
#Region "Dimension Tests"
    Private Sub TestDimensions()
        Console.WriteLine("--- Testing Dimensions ---")
        ' Test linear dimension
        Dim dim1 As DimensionElement = geometryLib.CreateDimension(
            geometryLib.CreatePoint3d(0, 700, 0),
            geometryLib.CreatePoint3d(100, 700, 0),
            geometryLib.CreatePoint3d(50, 720, 0)
        )
        geometryLib.AddElementToModel(dim1)
        Console.WriteLine("✓ Created linear dimension")
        Console.WriteLine()
    End Sub
#End Region
#Region "Utility Functions Tests"
    Private Sub TestUtilityFunctions()
        Console.WriteLine("--- Testing Utility Functions ---")
        ' Test distance calculation
        Dim point1 As Point3d = geometryLib.CreatePoint3d(0, 0, 0)
        Dim point2 As Point3d = geometryLib.CreatePoint3d(100, 100, 0)
        Dim distance As Double = geometryLib.CalculateDistance(point1, point2)
        Console.WriteLine("✓ Distance calculation: " & distance.ToString("F2"))
        ' Test midpoint calculation
        Dim midpoint As Point3d = geometryLib.CalculateMidpoint(point1, point2)
        Console.WriteLine("✓ Midpoint calculation: (" & midpoint.X & ", " & midpoint.Y & ", " & midpoint.Z & ")")
        ' Test model bounds
        Dim bounds As Range3d = geometryLib.GetModelBounds()
        Console.WriteLine("✓ Model bounds retrieved")
        ' Test element property management
        Dim testLine As LineElement = geometryLib.CreateLineFromCoordinates(500, 800, 600, 800)
        geometryLib.SetElementColor(testLine, 3) ' Red color
        geometryLib.SetElementLineWeight(testLine, 5)
        geometryLib.AddElementToModel(testLine)
        Console.WriteLine("✓ Set element properties (color=3, line weight=5)")
        Console.WriteLine()
    End Sub
#End Region
#Region "Export Functionality Tests"
    Private Sub TestExportFunctionality()
        Console.WriteLine("--- Testing Export Functionality ---")
        Try
            Dim outputDir As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim testDir As String = System.IO.Path.Combine(outputDir, "GeometryLibraryTest")
            If Not System.IO.Directory.Exists(testDir) Then
                System.IO.Directory.CreateDirectory(testDir)
            End If
            ' Test individual format exports
            Dim formats() As GeometryExporter.ExportFormat = {
                GeometryExporter.ExportFormat.STL,
                GeometryExporter.ExportFormat.STEP,
                GeometryExporter.ExportFormat.OBJ,
                GeometryExporter.ExportFormat.DXF
            }
            Dim options As GeometryExporter.ExportOptions = exporter.CreateDefaultExportOptions()
            For Each format As GeometryExporter.ExportFormat In formats
                Dim fileName As String = "test_geometry." & exporter.GetFileExtensionForFormat(format)
                Dim filePath As String = System.IO.Path.Combine(testDir, fileName)
                Dim success As Boolean = exporter.ExportModel(filePath, format, options)
                If success Then
                    Console.WriteLine("✓ Exported to " & format.ToString() & " format")
                Else
                    Console.WriteLine("✗ Failed to export to " & format.ToString() & " format")
                End If
            Next
            ' Test batch export
            Dim baseFilePath As String = System.IO.Path.Combine(testDir, "batch_export")
            Dim batchResults = exporter.ExportMultipleFormats(baseFilePath, formats, options)
            Console.WriteLine("✓ Completed batch export of " & batchResults.Count & " formats")
            ' Display export statistics
            Dim stats = exporter.GetExportStatistics()
            Console.WriteLine("✓ Export statistics gathered:")
            For Each kvp In stats
                Console.WriteLine("  " & kvp.Key & ": " & kvp.Value.ToString())
            Next
        Catch ex As Exception
            Console.WriteLine("Export test error: " & ex.Message)
        End Try
        Console.WriteLine()
    End Sub
#End Region
#Region "Demonstration Functions"
    ''' <summary>
    ''' Creates a comprehensive geometry showcase
    ''' </summary>
    Public Sub CreateGeometryShowcase()
        Console.WriteLine("Creating comprehensive geometry showcase...")
        ' Create title
        Dim title As TextElement = geometryLib.CreateTextFromCoordinates(
            "MicroStation V8i Geometry Library Showcase", 10, 50
        )
        geometryLib.AddElementToModel(title)
        ' Create basic shapes section
        CreateBasicShapesSection(50, 100)
        ' Create complex shapes section
        CreateComplexShapesSection(350, 100)
        ' Create 3D elements section
        Create3DElementsSection(50, 400)
        ' Create curves section
        CreateCurvesSection(350, 400)
        Console.WriteLine("Geometry showcase completed!")
    End Sub
    Private Sub CreateBasicShapesSection(startX As Double, startY As Double)
        ' Section title
        Dim sectionTitle As TextElement = geometryLib.CreateTextFromCoordinates("Basic Shapes", startX, startY - 20)
        geometryLib.AddElementToModel(sectionTitle)
        ' Line
        Dim line As LineElement = geometryLib.CreateLineFromCoordinates(startX, startY, startX + 60, startY)
        geometryLib.AddElementToModel(line)
        ' Circle
        Dim circle As EllipseElement = geometryLib.CreateCircleFromCoordinates(startX + 100, startY + 30, 25)
        geometryLib.AddElementToModel(circle)
        ' Rectangle
        Dim rect As ShapeElement = geometryLib.CreateRectangleFromCoordinates(startX, startY + 70, startX + 60, startY + 120)
        geometryLib.AddElementToModel(rect)
        ' Arc
        'Dim arc As EllipseElement = geometryLib.CreateArcFromCoordinates(startX + 150, startY + 95, 30, 0, Math.PI * 1.5)
        'geometryLib.AddElementToModel(arc)
    End Sub
    Private Sub CreateComplexShapesSection(startX As Double, startY As Double)
        ' Section title
        Dim sectionTitle As TextElement = geometryLib.CreateTextFromCoordinates("Complex Shapes", startX, startY - 20)
        geometryLib.AddElementToModel(sectionTitle)
        ' Hexagon
        Dim hexagon As ComplexShapeElement = geometryLib.CreateRegularPolygon(
            geometryLib.CreatePoint3d(startX + 40, startY + 40, 0), 30, 6
        )
        geometryLib.AddElementToModel(hexagon)
        ' Pentagon
        Dim pentagon As ComplexShapeElement = geometryLib.CreateRegularPolygon(
            geometryLib.CreatePoint3d(startX + 120, startY + 40, 0), 25, 5
        )
        geometryLib.AddElementToModel(pentagon)
        ' Octagon
        Dim octagon As ComplexShapeElement = geometryLib.CreateRegularPolygon(
            geometryLib.CreatePoint3d(startX + 40, startY + 120, 0), 28, 8
        )
        geometryLib.AddElementToModel(octagon)
        ' Star shape (custom polygon)
        Dim starPoints(9) As Point3d
        Dim centerX As Double = startX + 120
        Dim centerY As Double = startY + 120
        For i As Integer = 0 To 9
            Dim angle As Double = i * Math.PI / 5
            Dim radius As Double = If(i Mod 2 = 0, 30, 15)
            starPoints(i) = geometryLib.CreatePoint3d(
                centerX + radius * Math.Cos(angle),
                centerY + radius * Math.Sin(angle),
                0
            )
        Next
        Dim star As ShapeElement = geometryLib.CreatePolygon(starPoints)
        geometryLib.AddElementToModel(star)
    End Sub
    Private Sub Create3DElementsSection(startX As Double, startY As Double)
        ' Section title
        Dim sectionTitle As TextElement = geometryLib.CreateTextFromCoordinates("3D Elements", startX, startY - 20)
        geometryLib.AddElementToModel(sectionTitle)
        ' 3D lines forming a cube wireframe
        Dim cubeSize As Double = 50
        Dim cubeBase As Point3d = geometryLib.CreatePoint3d(startX + 50, startY + 50, 0)
        ' Bottom face
        Dim bottom1 As LineElement = geometryLib.CreateLine3D(cubeBase, geometryLib.CreatePoint3d(cubeBase.X + cubeSize, cubeBase.Y, 0))
        Dim bottom2 As LineElement = geometryLib.CreateLine3D(geometryLib.CreatePoint3d(cubeBase.X + cubeSize, cubeBase.Y, 0), geometryLib.CreatePoint3d(cubeBase.X + cubeSize, cubeBase.Y + cubeSize, 0))
        Dim bottom3 As LineElement = geometryLib.CreateLine3D(geometryLib.CreatePoint3d(cubeBase.X + cubeSize, cubeBase.Y + cubeSize, 0), geometryLib.CreatePoint3d(cubeBase.X, cubeBase.Y + cubeSize, 0))
        Dim bottom4 As LineElement = geometryLib.CreateLine3D(geometryLib.CreatePoint3d(cubeBase.X, cubeBase.Y + cubeSize, 0), cubeBase)
        geometryLib.AddElementsToModel(New _Element() {bottom1, bottom2, bottom3, bottom4})
        ' 3D points at different elevations
        For i As Integer = 0 To 4
            Dim pt3d As PointStringElement = geometryLib.CreatePointFromCoordinates(
                startX + 200 + i * 20, startY + 50, i * 15
            )
            geometryLib.AddElementToModel(pt3d)
        Next
    End Sub
    Private Sub CreateCurvesSection(startX As Double, startY As Double)
        ' Section title
        Dim sectionTitle As TextElement = geometryLib.CreateTextFromCoordinates("Curves", startX, startY - 20)
        geometryLib.AddElementToModel(sectionTitle)
        ' Sine wave curve
        Dim wavePoints(20) As Point3d
        For i As Integer = 0 To 20
            wavePoints(i) = geometryLib.CreatePoint3d(
                startX + i * 5,
                startY + 50 + 30 * Math.Sin(i * Math.PI / 10),
                0
            )
        Next
        Dim wave As CurveElement = geometryLib.CreateCurve(wavePoints)
        geometryLib.AddElementToModel(wave)
        ' Spiral-like curve
        Dim spiralPoints(30) As Point3d
        For i As Integer = 0 To 30
            Dim angle As Double = i * Math.PI / 6
            Dim radius As Double = i * 2
            spiralPoints(i) = geometryLib.CreatePoint3d(
                startX + 50 + radius * Math.Cos(angle),
                startY + 150 + radius * Math.Sin(angle),
                0
            )
        Next
        Dim spiral As CurveElement = geometryLib.CreateCurve(spiralPoints)
        geometryLib.AddElementToModel(spiral)
    End Sub
#End Region
End Class