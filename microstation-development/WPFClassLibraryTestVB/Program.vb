Imports System
Imports System.Runtime.InteropServices
Imports MicroStationDGN
Imports WPFClassLibrary.MstnCom

Module Program

    Sub Main(args As String())

        Dim app = Attach(True)
        Dim p0 = app.Point3dFromXY(0, 0)
        Dim p1 = app.Point3dFromXY(50, 25)
        Dim ln = app.CreateLineElement2(Nothing, p0, p1)
        app.ActiveModelReference.AddElement(ln)

    End Sub

End Module