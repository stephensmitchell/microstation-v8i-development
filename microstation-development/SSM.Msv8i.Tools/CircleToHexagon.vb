Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports MicroStationDGN
Public Class CircleToHexagon
    Private ustation As Application
    Public Sub New(ustation As Application)
        Me.ustation = ustation
    End Sub
    Public Sub Run()
        Dim scanCriteria As New ElementScanCriteria
        scanCriteria.IncludeType(MsdElementType.msdElementTypeEllipse)
        Dim enumerator As ElementEnumerator = ustation.ActiveModelReference.Scan(scanCriteria)
        Dim circle As EllipseElement = Nothing
        While enumerator.MoveNext
            Dim element As Element = enumerator.Current
            If element.IsEllipseElement Then
                Dim ellipse As EllipseElement = CType(element, EllipseElement)
                If ellipse.PrimaryRadius = ellipse.SecondaryRadius Then
                    circle = ellipse
                    Exit While
                End If
            End If
        End While
        If circle Is Nothing Then
            Console.WriteLine("No circles found in the active model.")
            Return
        End If
        Dim numSides As Integer = 6
        Dim radius As Double = circle.PrimaryRadius
        Dim rotationAngleRadians As Double = 0
        Dim center As Point3d = circle.CenterPoint
        Dim vertices(numSides) As Point3d
        For i As Integer = 0 To numSides - 1
            Dim angle As Double = 2 * Math.PI * i / numSides + rotationAngleRadians
            vertices(i) = ustation.Point3dFromXYZ(center.X + radius * Math.Cos(angle), center.Y + radius * Math.Sin(angle), center.Z)
        Next
        vertices(numSides) = vertices(0)
        Dim chainableElements(numSides - 1) As ChainableElement
        For i As Integer = 0 To numSides - 1
            chainableElements(i) = ustation.CreateLineElement2(Nothing, vertices(i), vertices(i + 1))
        Next
        Dim hexagon As ComplexShapeElement = ustation.CreateComplexShapeElement1(chainableElements)
        ustation.ActiveModelReference.AddElement(hexagon)
        ustation.ActiveModelReference.RemoveElement(circle)
        Console.WriteLine("First circle found was converted to a hexagon.")
    End Sub
End Class
