Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports MicroStationDGN
Public Class PipeTools
    Private ustation As Application
    Public Sub New(ustation As Application)
        Me.ustation = ustation
    End Sub
    Public Sub Run()
        ' Pipe parameters
        Dim outerRadius As Double = 1.0
        Dim innerRadius As Double = 0.8
        Dim length1 As Double = 10.0
        Dim length2 As Double = 8.0
        ' Create the four pipes
        Dim pipe1 As SmartSolidElement = ustation.SmartSolid.CreateCylinder(Nothing, outerRadius, length1)
        Dim pipe2 As SmartSolidElement = ustation.SmartSolid.CreateCylinder(Nothing, outerRadius, length2)
        Dim pipe3 As SmartSolidElement = ustation.SmartSolid.CreateCylinder(Nothing, outerRadius, length1)
        Dim pipe4 As SmartSolidElement = ustation.SmartSolid.CreateCylinder(Nothing, outerRadius, length2)
        ' Position the pipes
        pipe2.Rotate(ustation.Point3dZero(), 0, 0, ustation.Radians(90))
        pipe2.Move(ustation.Point3dFromXYZ(length1, 0, 0))
        pipe3.Move(ustation.Point3dFromXYZ(length1, length2, 0))
        pipe4.Rotate(ustation.Point3dZero(), 0, 0, ustation.Radians(90))
        pipe4.Move(ustation.Point3dFromXYZ(0, length2, 0))
        ' Create mitering planes
        Dim plane1 As ShapeElement = CreateMiterPlane(ustation.Point3dFromXYZ(0, 0, 0), ustation.Radians(45), 5)
        Dim plane2 As ShapeElement = CreateMiterPlane(ustation.Point3dFromXYZ(length1, 0, 0), ustation.Radians(-45), 5)
        Dim plane3 As ShapeElement = CreateMiterPlane(ustation.Point3dFromXYZ(length1, length2, 0), ustation.Radians(135), 5)
        Dim plane4 As ShapeElement = CreateMiterPlane(ustation.Point3dFromXYZ(0, length2, 0), ustation.Radians(225), 5)
        ' Miter the pipes
        Dim miteredPipe1Enum As ElementEnumerator = ustation.SmartSolid.TrimSolidWithSurface(pipe1, plane1, ustation.Point3dFromXYZ(0, 0, 0))
        If miteredPipe1Enum Is Nothing OrElse Not miteredPipe1Enum.MoveNext() Then
            Console.WriteLine("Failed to miter pipe 1")
            Return
        End If
        Dim miteredPipe1 As SmartSolidElement = CType(miteredPipe1Enum.Current, SmartSolidElement)
        Dim miteredPipe2Enum As ElementEnumerator = ustation.SmartSolid.TrimSolidWithSurface(pipe2, plane2, ustation.Point3dFromXYZ(length1, 0, 0))
        If miteredPipe2Enum Is Nothing OrElse Not miteredPipe2Enum.MoveNext() Then
            Console.WriteLine("Failed to miter pipe 2")
            Return
        End If
        Dim miteredPipe2 As SmartSolidElement = CType(miteredPipe2Enum.Current, SmartSolidElement)
        Dim miteredPipe3Enum As ElementEnumerator = ustation.SmartSolid.TrimSolidWithSurface(pipe3, plane3, ustation.Point3dFromXYZ(length1, length2, 0))
        If miteredPipe3Enum Is Nothing OrElse Not miteredPipe3Enum.MoveNext() Then
            Console.WriteLine("Failed to miter pipe 3")
            Return
        End If
        Dim miteredPipe3 As SmartSolidElement = CType(miteredPipe3Enum.Current, SmartSolidElement)
        Dim miteredPipe4Enum As ElementEnumerator = ustation.SmartSolid.TrimSolidWithSurface(pipe4, plane4, ustation.Point3dFromXYZ(0, length2, 0))
        If miteredPipe4Enum Is Nothing OrElse Not miteredPipe4Enum.MoveNext() Then
            Console.WriteLine("Failed to miter pipe 4")
            Return
        End If
        Dim miteredPipe4 As SmartSolidElement = CType(miteredPipe4Enum.Current, SmartSolidElement)
        ' Union the mitered pipes
        Dim assembly As SmartSolidElement = miteredPipe1
        assembly = ustation.SmartSolid.SolidUnion(assembly, miteredPipe2)
        assembly = ustation.SmartSolid.SolidUnion(assembly, miteredPipe3)
        assembly = ustation.SmartSolid.SolidUnion(assembly, miteredPipe4)
        ' Add the assembly to the model
        ustation.ActiveModelReference.AddElement(assembly)
        Console.WriteLine("Created a four-member pipe assembly with mitered ends.")
    End Sub
    Private Function CreateMiterPlane(center As Point3d, angle As Double, size As Double) As ShapeElement
        Dim vertices(4) As Point3d
        vertices(0) = ustation.Point3dFromXYZ(center.X - size, center.Y - size, 0)
        vertices(1) = ustation.Point3dFromXYZ(center.X + size, center.Y - size, 0)
        vertices(2) = ustation.Point3dFromXYZ(center.X + size, center.Y + size, 0)
        vertices(3) = ustation.Point3dFromXYZ(center.X - size, center.Y + size, 0)
        vertices(4) = vertices(0)
        Dim plane As ShapeElement = ustation.CreateShapeElement1(Nothing, vertices)
        plane.Rotate(center, 0, 0, angle)
        Return plane
    End Function
End Class
