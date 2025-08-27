Imports System.Runtime.InteropServices
Imports BCOM = MicroStationDGN
Public Module Init

End Module

Public Module MstnCom
    Private Const ProgId As String = "MicroStationDGN.Application"

    Public Function AttachOrLaunch(Optional visible As Boolean = True,
                                   Optional dgnToOpen As String = Nothing) As BCOM.Application
        Try
            Return Attach(visible)
        Catch ex As COMException
            Return Launch(visible, dgnToOpen)
        End Try
    End Function

    Public Function Attach(Optional visible As Boolean = True) As BCOM.Application
        Dim obj = Marshal.GetActiveObject(ProgId)
        Dim app = CType(obj, BCOM.Application)
        app.Visible = visible
        Return app
    End Function

    Public Function Launch(Optional visible As Boolean = True,
                           Optional dgnToOpen As String = Nothing) As BCOM.Application
        Dim t = Type.GetTypeFromProgID(ProgId, True)
        Dim app = CType(Activator.CreateInstance(t), BCOM.Application)
        app.Visible = visible
        If Not String.IsNullOrEmpty(dgnToOpen) Then app.OpenDesignFile(dgnToOpen, False)
        Return app
    End Function
End Module
