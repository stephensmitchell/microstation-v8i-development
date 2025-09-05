Imports System.Runtime.InteropServices
Imports BCOM = MicroStationDGN
Public Module WBeamGenerator
    Public Function CreateWBeam(app As BCOM.Application,length As Double,
                                Optional d As Double = 4.125,
                                Optional bf As Double = 4.0,
                                Optional tf As Double = 0.375,
                                Optional tw As Double = 0.280) As BCOM.SmartSolidElement
        Dim half_d = d / 2
        Dim half_bf = bf / 2
        Dim half_tw = tw / 2
        Dim points(11) As BCOM.Point3d
        points(0) = app.Point3dFromXYZ(-half_bf, half_d, 0)
        points(1) = app.Point3dFromXYZ(half_bf, half_d, 0)
        points(2) = app.Point3dFromXYZ(half_bf, half_d - tf, 0)
        points(3) = app.Point3dFromXYZ(half_tw, half_d - tf, 0)
        points(4) = app.Point3dFromXYZ(half_tw, -half_d + tf, 0)
        points(5) = app.Point3dFromXYZ(half_bf, -half_d + tf, 0)
        points(6) = app.Point3dFromXYZ(half_bf, -half_d, 0)
        points(7) = app.Point3dFromXYZ(-half_bf, -half_d, 0)
        points(8) = app.Point3dFromXYZ(-half_bf, -half_d + tf, 0)
        points(9) = app.Point3dFromXYZ(-half_tw, -half_d + tf, 0)
        points(10) = app.Point3dFromXYZ(-half_tw, half_d - tf, 0)
        points(11) = app.Point3dFromXYZ(-half_bf, half_d - tf, 0)
        Dim profile As BCOM.ShapeElement = app.CreateShapeElement1(Nothing, points)
        Dim solid As BCOM.SmartSolidElement = app.SmartSolid.ExtrudeClosedPlanarCurve(profile, length, 0, True)
        Return solid
    End Function
End Module