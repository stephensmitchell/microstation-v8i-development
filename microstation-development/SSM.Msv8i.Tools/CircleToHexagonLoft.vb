Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports MicroStationDGN
Public Class CircleToHexagonLoft
    Private ustation As Application
    Public Sub New(ustation As Application)
        Me.ustation = ustation
    End Sub
    Public Sub Run()
        ' Define circle parameters
        Dim circleCenter As Point3d = ustation.Point3dFromXYZ(0, 0, 0)
        Dim circleRadius As Double = 5.0
        ' Create circle profile
        Dim circleArcElement As ArcElement = ustation.CreateArcElement2(Nothing, circleCenter, circleRadius, circleRadius, ustation.Matrix3dIdentity(), 0, 2 * Math.PI)
        Dim circleCurve As New BsplineCurve
        circleCurve.FromElement(circleArcElement)
        ' Define hexagon parameters
        Dim hexagonCenter As Point3d = ustation.Point3dFromXYZ(0, 0, 10)
        Dim hexagonRadius As Double = 5.0
        Dim numSides As Integer = 6
        ' Calculate hexagon vertices
        Dim vertices(numSides) As Point3d
        For i As Integer = 0 To numSides - 1
            Dim angle As Double = 2 * Math.PI * i / numSides
            vertices(i) = ustation.Point3dFromXYZ(hexagonCenter.X + hexagonRadius * Math.Cos(angle), hexagonCenter.Y + hexagonRadius * Math.Sin(angle), hexagonCenter.Z)
        Next
        vertices(numSides) = vertices(0)
        ' Create hexagon profile
        Dim hexagonLineElement As LineElement = ustation.CreateLineElement1(Nothing, vertices)
        Dim hexagonCurve As New BsplineCurve
        hexagonCurve.FromElement(hexagonLineElement)
        ' Create loft surface
        Dim profiles(1) As Object
        profiles(0) = circleCurve
        profiles(1) = hexagonCurve
        Dim loftSurface As New BsplineSurface
        loftSurface.FromCrossSections(profiles)
        Dim loftElement As BsplineSurfaceElement = ustation.CreateBsplineSurfaceElement1(Nothing, loftSurface)
        ' Add the loft to the model
        ustation.ActiveModelReference.AddElement(loftElement)
        Console.WriteLine("Created a lofted surface between a circle and a hexagon.")
    End Sub
End Class
