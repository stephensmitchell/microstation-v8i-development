Imports System
Imports System.Collections.Generic
Imports System.Math
Imports MicroStationDGN
''' <summary>
''' Comprehensive geometry function library for MicroStation V8i
''' Provides complete 2D and 3D CAD geometry creation and manipulation functionality
''' Based on Bentley MicroStation DGN 8.9 Object Library
''' </summary>
Public Class GeometryLibrary
    Private ustation As Application
    Public Sub New(ustation As Application)
        Me.ustation = ustation
    End Sub
#Region "2D Basic Geometry Creation"
    ''' <summary>
    ''' Creates a line element from two 3D points
    ''' </summary>
    Public Function CreateLine(startPoint As Point3d, endPoint As Point3d, Optional template As _Element = Nothing) As LineElement
        Dim startPt As Point3d = startPoint
        Dim endPt As Point3d = endPoint
        Return ustation.CreateLineElement2(template, startPt, endPt)
    End Function
    ''' <summary>
    ''' Creates a line element from coordinates
    ''' </summary>
    Public Function CreateLineFromCoordinates(x1 As Double, y1 As Double, x2 As Double, y2 As Double, Optional z1 As Double = 0.0, Optional z2 As Double = 0.0, Optional template As _Element = Nothing) As LineElement
        Dim startPoint As Point3d = ustation.Point3dFromXYZ(x1, y1, z1)
        Dim endPoint As Point3d = ustation.Point3dFromXYZ(x2, y2, z2)
        Return CreateLine(startPoint, endPoint, template)
    End Function
    ''' <summary>
    ''' Creates a line element from array of points
    ''' </summary>
    Public Function CreateLineString(points() As Point3d, Optional template As _Element = Nothing) As LineElement
        Dim pointArray As Array = points
        Return ustation.CreateLineElement1(template, pointArray)
    End Function
    ''' <summary>
    ''' Creates a circle (ellipse with equal radii)
    ''' </summary>
    Public Function CreateCircle(centerPoint As Point3d, radius As Double, Optional template As _Element = Nothing) As EllipseElement
        Dim center As Point3d = centerPoint
        Dim rotation As Matrix3d = ustation.Matrix3dIdentity()
        Return ustation.CreateEllipseElement2(template, center, radius, radius, rotation)
    End Function
    ''' <summary>
    ''' Creates a circle from coordinates
    ''' </summary>
    Public Function CreateCircleFromCoordinates(centerX As Double, centerY As Double, radius As Double, Optional centerZ As Double = 0.0, Optional template As _Element = Nothing) As EllipseElement
        Dim centerPoint As Point3d = ustation.Point3dFromXYZ(centerX, centerY, centerZ)
        Return CreateCircle(centerPoint, radius, template)
    End Function
    ''' <summary>
    ''' Creates an ellipse with different primary and secondary radii
    ''' </summary>
    Public Function CreateEllipse(centerPoint As Point3d, primaryRadius As Double, secondaryRadius As Double, Optional rotationAngle As Double = 0.0, Optional template As _Element = Nothing) As EllipseElement
        Dim center As Point3d = centerPoint
        Dim rotation As Matrix3d = ustation.Matrix3dIdentity()
        Return ustation.CreateEllipseElement2(template, center, primaryRadius, secondaryRadius, rotation)
    End Function
    ''' <summary>
    ''' Creates an elliptical arc
    ''' </summary>
    Public Function CreateEllipticalArc(centerPoint As Point3d, primaryRadius As Double, secondaryRadius As Double, startAngle As Double, sweepAngle As Double, Optional rotationAngle As Double = 0.0, Optional template As _Element = Nothing) As EllipseElement
        Dim origin As Point3d = centerPoint
        Dim rotation As Matrix3d = ustation.Matrix3dFromRotateZ(rotationAngle)
        Return ustation.CreateEllipseElement2(template, origin, primaryRadius, secondaryRadius, rotation)
    End Function
    ''' <summary>
    ''' Creates a circular arc
    ''' </summary>
    Public Function CreateArc(centerPoint As Point3d, radius As Double, startAngle As Double, sweepAngle As Double, Optional template As _Element = Nothing) As ArcElement
        Dim center As Point3d = centerPoint
        Dim rotation As Matrix3d = ustation.Matrix3dIdentity()
        Return ustation.CreateArcElement2(template, center, radius, radius, rotation, startAngle, sweepAngle)
    End Function
    ''' <summary>
    ''' Creates an arc by center, start point and end point
    ''' </summary>
    Public Function CreateArcByThreePoints(startPoint As Point3d, middlePoint As Point3d, endPoint As Point3d, Optional template As _Element = Nothing) As ArcElement
        Return ustation.CreateArcElement1(template, startPoint, middlePoint, endPoint)
    End Function
    ''' <summary>
    ''' Creates an arc by center, radius and two angles
    ''' </summary>
    Public Function CreateArcByCenterRadius(centerPoint As Point3d, radius As Double, startAngle As Double, sweepAngle As Double, Optional template As _Element = Nothing) As ArcElement
        Dim center As Point3d = centerPoint
        Dim rotation As Matrix3d = ustation.Matrix3dIdentity()
        Return ustation.CreateArcElement2(template, center, radius, radius, rotation, startAngle, sweepAngle)
    End Function
    ''' <summary>
    ''' Creates an arc from coordinates (helper method for tests)
    ''' </summary>
    Public Function CreateArcFromCoordinates(centerX As Double, centerY As Double, centerZ As Double, radius As Double, startAngle As Double, sweepAngle As Double, Optional template As _Element = Nothing) As ArcElement
        Dim centerPoint As Point3d = ustation.Point3dFromXYZ(centerX, centerY, centerZ)
        Return CreateArcByCenterRadius(centerPoint, radius, startAngle, sweepAngle, template)
    End Function
    ''' <summary>
    ''' Creates a point element
    ''' </summary>
    Public Function CreatePoint(point As Point3d, Optional template As _Element = Nothing) As PointStringElement
        Dim points(0) As Point3d
        points(0) = point
        Dim pointArray As Array = points
        Return ustation.CreatePointStringElement1(template, pointArray, False)
    End Function
    ''' <summary>
    ''' Creates a point from coordinates
    ''' </summary>
    Public Function CreatePointFromCoordinates(x As Double, y As Double, Optional z As Double = 0.0, Optional template As _Element = Nothing) As PointStringElement
        Dim point As Point3d = ustation.Point3dFromXYZ(x, y, z)
        Return CreatePoint(point, template)
    End Function
    ''' <summary>
    ''' Creates multiple points
    ''' </summary>
    Public Function CreatePointString(points() As Point3d, Optional template As _Element = Nothing) As PointStringElement
        Return ustation.CreatePointStringElement1(template, points)
    End Function
#End Region
#Region "2D Complex Shapes"
    ''' <summary>
    ''' Creates a polygon (closed shape) from vertices
    ''' </summary>
    Public Function CreatePolygon(vertices() As Point3d, Optional template As _Element = Nothing, Optional fillMode As MsdFillMode = MsdFillMode.msdFillModeUseActive) As ShapeElement
        Dim vertexArray As Array = vertices
        Return ustation.CreateShapeElement1(template, vertexArray, fillMode)
    End Function
    ''' <summary>
    ''' Creates a regular polygon
    ''' </summary>
    Public Function CreateRegularPolygon(centerPoint As Point3d, radius As Double, numSides As Integer, Optional rotationAngle As Double = 0.0, Optional template As _Element = Nothing) As ComplexShapeElement
        If numSides < 3 Then
            Throw New ArgumentException("Polygon must have at least 3 sides")
        End If
        Dim vertices(numSides - 1) As Point3d
        For i As Integer = 0 To numSides - 1
            Dim angle As Double = 2 * PI * i / numSides + rotationAngle
            vertices(i) = ustation.Point3dFromXYZ(
                centerPoint.X + radius * Cos(angle),
                centerPoint.Y + radius * Sin(angle),
                centerPoint.Z
            )
        Next
        Dim chainableElements(numSides - 1) As ChainableElement
        For i As Integer = 0 To numSides - 1
            Dim nextIndex As Integer = (i + 1) Mod numSides
            chainableElements(i) = ustation.CreateLineElement2(template, vertices(i), vertices(nextIndex))
        Next
        Return ustation.CreateComplexShapeElement1(chainableElements)
    End Function
    ''' <summary>
    ''' Creates a rectangle
    ''' </summary>
    Public Function CreateRectangle(corner1 As Point3d, corner2 As Point3d, Optional template As _Element = Nothing) As ShapeElement
        Dim vertices(3) As Point3d
        vertices(0) = corner1
        vertices(1) = ustation.Point3dFromXYZ(corner2.X, corner1.Y, corner1.Z)
        vertices(2) = corner2
        vertices(3) = ustation.Point3dFromXYZ(corner1.X, corner2.Y, corner1.Z)
        Return CreatePolygon(vertices, template)
    End Function
    ''' <summary>
    ''' Creates a rectangle from coordinates
    ''' </summary>
    Public Function CreateRectangleFromCoordinates(x1 As Double, y1 As Double, x2 As Double, y2 As Double, Optional z As Double = 0.0, Optional template As _Element = Nothing) As ShapeElement
        Dim corner1 As Point3d = ustation.Point3dFromXYZ(x1, y1, z)
        Dim corner2 As Point3d = ustation.Point3dFromXYZ(x2, y2, z)
        Return CreateRectangle(corner1, corner2, template)
    End Function
    ''' <summary>
    ''' Creates a complex shape from chainable elements
    ''' </summary>
    Public Function CreateComplexShape(chainableElements() As ChainableElement, Optional fillMode As MsdFillMode = MsdFillMode.msdFillModeUseActive) As ComplexShapeElement
        Return ustation.CreateComplexShapeElement1(chainableElements, fillMode)
    End Function
    ''' <summary>
    ''' Creates a complex string from chainable elements
    ''' </summary>
    Public Function CreateComplexString(chainableElements() As ChainableElement, Optional gapTolerance As Double = -1.0) As ComplexStringElement
        If gapTolerance = -1.0 Then
            Return ustation.CreateComplexStringElement1(chainableElements)
        Else
            Return ustation.CreateComplexStringElement2(chainableElements, gapTolerance)
        End If
    End Function
#End Region
#Region "Curves and Splines"
    ''' <summary>
    ''' Creates a B-spline curve element from control points
    ''' </summary>
    Public Function CreateBSplineCurve(curve As BsplineCurve, Optional template As _Element = Nothing) As BsplineCurveElement
        Return ustation.CreateBsplineCurveElement1(template, curve)
    End Function
    ''' <summary>
    ''' Creates a B-spline curve from interpolation curve
    ''' </summary>
    Public Function CreateInterpolationCurve(curve As InterpolationCurve, Optional template As _Element = Nothing) As BsplineCurveElement
        Return ustation.CreateBsplineCurveElement2(template, curve)
    End Function
    ''' <summary>
    ''' Creates a curve element from points
    ''' </summary>
    Public Function CreateCurve(points() As Point3d, Optional template As _Element = Nothing) As CurveElement
        Dim pointArray As Array = points
        Return ustation.CreateCurveElement1(template, pointArray)
    End Function
#End Region
#Region "3D Geometry Creation"
    ''' <summary>
    ''' Creates a 3D line
    ''' </summary>
    Public Function CreateLine3D(startPoint As Point3d, endPoint As Point3d, Optional template As _Element = Nothing) As LineElement
        Return ustation.CreateLineElement2(template, startPoint, endPoint)
    End Function
    ''' <summary>
    ''' Creates a cone element
    ''' </summary>
    Public Function CreateCone(baseCenter As Point3d, topCenter As Point3d, baseRadius As Double, topRadius As Double, Optional template As _Element = Nothing) As ConeElement
        Dim baseCtr As Point3d = baseCenter
        Dim topCtr As Point3d = topCenter
        Dim rotation As Matrix3d = ustation.Matrix3dIdentity()
        Return ustation.CreateConeElement1(template, baseRadius, baseCtr, topRadius, topCtr, rotation)
    End Function
    ''' <summary>
    ''' Creates a simple cone (cylinder tapered to point)
    ''' </summary>
    Public Function CreateSimpleCone(baseCenter As Point3d, topCenter As Point3d, radius As Double, Optional template As _Element = Nothing) As ConeElement
        Dim baseCtr As Point3d = baseCenter
        Dim topCtr As Point3d = topCenter
        Return ustation.CreateConeElement2(template, radius, baseCtr, topCtr)
    End Function
    ''' <summary>
    ''' Creates a B-spline surface
    ''' </summary>
    Public Function CreateBSplineSurface(surface As BsplineSurface, Optional template As _Element = Nothing) As BsplineSurfaceElement
        Return ustation.CreateBsplineSurfaceElement1(template, surface)
    End Function
#End Region
#Region "Text Elements"
    ''' <summary>
    ''' Creates a text element
    ''' </summary>
    Public Function CreateText(text As String, origin As Point3d, Optional template As _Element = Nothing) As TextElement
        Dim textOrigin As Point3d = origin
        Dim rotation As Matrix3d = ustation.Matrix3dIdentity()
        Return ustation.CreateTextElement1(template, text, textOrigin, rotation)
    End Function
    ''' <summary>
    ''' Creates text from coordinates
    ''' </summary>
    Public Function CreateTextFromCoordinates(text As String, x As Double, y As Double, Optional z As Double = 0.0, Optional template As _Element = Nothing) As TextElement
        Dim origin As Point3d = ustation.Point3dFromXYZ(x, y, z)
        Return CreateText(text, origin, template)
    End Function
    ''' <summary>
    ''' Creates a multi-line text node
    ''' </summary>
    Public Function CreateTextNode(text As String, origin As Point3d, Optional template As _Element = Nothing) As MultiLineElement
        Dim textOrigin As Point3d = origin
        Dim rotation As Matrix3d = ustation.Matrix3dIdentity()
        Return ustation.CreateMultiLineElement1(template, text, textOrigin, rotation)
    End Function
#End Region
#Region "Dimension Elements"
    ''' <summary>
    ''' Creates a dimension element
    ''' </summary>
    Public Function CreateDimension(point1 As Point3d, point2 As Point3d, dimensionPoint As Point3d, Optional template As _Element = Nothing) As DimensionElement
        Dim rotation As Matrix3d = ustation.Matrix3dIdentity()
        'Dim dimType As MsdDimType = MsdDimType.msdDimTypeLength
        Return ustation.CreateDimensionElement1(template, rotation, dimType)
    End Function
#End Region
#Region "Multi-Line Elements"
    ''' <summary>
    ''' Creates a multi-line element
    ''' </summary>
    Public Function CreateMultiLine(points() As Point3d, Optional template As _Element = Nothing) As MultiLineElement
        Return ustation.CreateMultiLineElement1(template, points)
    End Function
#End Region
#Region "Cell Elements"
    ''' <summary>
    ''' Creates a cell element from component elements
    ''' </summary>
    Public Function CreateCell(componentElements() As _Element, origin As Point3d, Optional template As _Element = Nothing) As CellElement
        Return ustation.CreateCellElement1(template, componentElements, origin)
    End Function
    ''' <summary>
    ''' Creates a shared cell element
    ''' </summary>
    Public Function CreateSharedCell(cellName As String, origin As Point3d, Optional trueScale As Boolean = True, Optional template As _Element = Nothing) As SharedCellElement
        Return ustation.CreateSharedCellElement3(cellName, origin, trueScale)
    End Function
#End Region
#Region "Utility Functions"
    ''' <summary>
    ''' Adds an element to the active model
    ''' </summary>
    Public Function AddElementToModel(element As _Element) As Boolean
        Try
            ustation.ActiveModelReference.AddElement(element)
            Return True
        Catch ex As Exception
            Console.WriteLine("Error adding element to model: " & ex.Message)
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Adds multiple elements to the active model
    ''' </summary>
    Public Function AddElementsToModel(elements() As _Element) As Integer
        Dim successCount As Integer = 0
        For Each element As _Element In elements
            If AddElementToModel(element) Then
                successCount += 1
            End If
        Next
        Return successCount
    End Function
    ''' <summary>
    ''' Gets the bounds of the active model
    ''' </summary>
    Public Function GetModelBounds() As Range3d
        Return ustation.ActiveModelReference.Range(True)
    End Function
    ''' <summary>
    ''' Creates a 3D point from coordinates
    ''' </summary>
    Public Function CreatePoint3d(x As Double, y As Double, z As Double) As Point3d
        Return ustation.Point3dFromXYZ(x, y, z)
    End Function
    ''' <summary>
    ''' Creates a 2D point from coordinates
    ''' </summary>
    Public Function CreatePoint2d(x As Double, y As Double) As Point2d
        Return ustation.Point2dFromXY(x, y)
    End Function
    ''' <summary>
    ''' Creates a transformation matrix for translation
    ''' </summary>
    ''' <summary>
    ''' Creates an identity transformation matrix
    ''' </summary>
    Public Function CreateIdentityTransform() As Transform3d
        Return ustation.Transform3dIdentity()
    End Function
    ''' <summary>
    ''' Creates a transformation matrix from translation vector
    ''' </summary>
    Public Function CreateTranslationTransform(translation As Point3d) As Transform3d
        Dim trans As Point3d = translation
        Return ustation.Transform3dFromPoint3d(trans)
    End Function
    ''' <summary>
    ''' Copies an element
    ''' </summary>
    Public Function CopyElement(element As _Element) As _Element
        Return element.Clone()
    End Function
    ''' <summary>
    ''' Transforms an element using a transformation matrix
    ''' </summary>
    Public Sub TransformElement(element As _Element, transform As Transform3d)
        Dim trans As Transform3d = transform
        element.Transform(trans)
    End Sub
    ''' <summary>
    ''' Moves an element by offset
    ''' </summary>
    Public Sub MoveElement(element As _Element, offset As Point3d)
        Dim offsetPt As Point3d = offset
        element.Move(offsetPt)
    End Sub
    ''' <summary>
    ''' Gets element range/bounding box
    ''' </summary>
    Public Function GetElementRange(element As _Element) As Range3d
        Return element.Range(True)
    End Function
#End Region
#Region "Element Property Management"
    ''' <summary>
    ''' Sets element color
    ''' </summary>
    Public Sub SetElementColor(element As _Element, color As Integer)
        element.Color = color
    End Sub
    ''' <summary>
    ''' Sets element line weight
    ''' </summary>
    Public Sub SetElementLineWeight(element As _Element, lineWeight As Integer)
        element.LineWeight = lineWeight
    End Sub
    ''' <summary>
    ''' Sets element level
    ''' </summary>
    Public Sub SetElementLevel(element As _Element, level As Level)
        element.Level = level
    End Sub
    ''' <summary>
    ''' Sets element line style
    ''' </summary>
    Public Sub SetElementLineStyle(element As _Element, lineStyle As LineStyle)
        element.LineStyle = lineStyle
    End Sub
    ''' <summary>
    ''' Locks/unlocks an element
    ''' </summary>
    Public Sub SetElementLocked(element As _Element, locked As Boolean)
        element.IsLocked = locked
    End Sub
    ''' <summary>
    ''' Shows/hides an element
    ''' </summary>
    Public Sub SetElementHidden(element As _Element, hidden As Boolean)
        element.IsHidden = hidden
    End Sub
#End Region
#Region "Element Query Functions"
    ''' <summary>
    ''' Checks if element is a line
    ''' </summary>
    Public Function IsLineElement(element As _Element) As Boolean
        Return element.IsLineElement
    End Function
    ''' <summary>
    ''' Checks if element is an arc
    ''' </summary>
    Public Function IsArcElement(element As _Element) As Boolean
        Return element.IsArcElement
    End Function
    ''' <summary>
    ''' Checks if element is an ellipse
    ''' </summary>
    Public Function IsEllipseElement(element As _Element) As Boolean
        Return element.IsEllipseElement
    End Function
    ''' <summary>
    ''' Checks if element is text
    ''' </summary>
    Public Function IsTextElement(element As _Element) As Boolean
        Return element.IsTextElement
    End Function
    ''' <summary>
    ''' Checks if element is a dimension
    ''' </summary>
    Public Function IsDimensionElement(element As _Element) As Boolean
        Return element.IsDimensionElement
    End Function
    ''' <summary>
    ''' Gets element type
    ''' </summary>
    Public Function GetElementType(element As _Element) As MsdElementType
        Return element.Type
    End Function
#End Region
#Region "Geometric Calculations"
    ''' <summary>
    ''' Calculates distance between two points
    ''' </summary>
    Public Function CalculateDistance(point1 As Point3d, point2 As Point3d) As Double
        Dim pt1 As Point3d = point1
        Dim pt2 As Point3d = point2
        Return ustation.Point3dDistance(pt1, pt2)
    End Function
    ''' <summary>
    ''' Calculates angle between two vectors
    ''' </summary>
    Public Function CalculateAngle(vector1 As Point3d, vector2 As Point3d) As Double
        Dim vec1 As Point3d = vector1
        Dim vec2 As Point3d = vector2
        Return ustation.Point3dDotProduct(vec1, vec2)
    End Function
    ''' <summary>
    ''' Calculates midpoint between two points
    ''' </summary>
    Public Function CalculateMidpoint(point1 As Point3d, point2 As Point3d) As Point3d
        Dim pt1 As Point3d = point1
        Dim pt2 As Point3d = point2
        Return ustation.Point3dInterpolate(pt1, 0.5, pt2)
    End Function
    ''' <summary>
    ''' Normalizes a vector
    ''' </summary>
    Public Function NormalizeVector(vector As Point3d) As Point3d
        Dim vec As Point3d = vector
        Return ustation.Point3dNormalize(vec)
    End Function
    ''' <summary>
    ''' Calculates cross product of two vectors
    ''' </summary>
    Public Function CrossProduct(vector1 As Point3d, vector2 As Point3d) As Point3d
        Dim vec1 As Point3d = vector1
        Dim vec2 As Point3d = vector2
        Return ustation.Point3dCrossProduct(vec1, vec2)
    End Function
    ''' <summary>
    ''' Calculates dot product of two vectors
    ''' </summary>
    Public Function DotProduct(vector1 As Point3d, vector2 As Point3d) As Double
        Dim vec1 As Point3d = vector1
        Dim vec2 As Point3d = vector2
        Return ustation.Point3dDotProduct(vec1, vec2)
    End Function
#End Region
End Class