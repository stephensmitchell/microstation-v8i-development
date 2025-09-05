Imports System.Runtime.InteropServices
Imports BCOM = MicroStationDGN
''' <summary>
''' Provides COM interop utilities for integrating and automating MicroStation V8i operations.
''' </summary>
''' <remarks>
''' This module allows you to attach to a running instance of MicroStation, and safely release COM objects.
''' </remarks>
Public Module MsCom
    ''' <summary>
    ''' The programmatic identifier (ProgId) for MicroStationDGN.Application.
    ''' </summary>
    Private Const ProgId As String = "MicroStationDGN.Application"
    ''' <summary>
    ''' Attaches to a running instance of MicroStation V8i.
    ''' </summary>
    ''' <param name="visible">If True, makes the MicroStation application window visible.</param>
    ''' <returns>
    ''' Returns the <see cref="MicroStationDGN.Application"/> object if successful.
    ''' Throws <see cref="InvalidOperationException"/> if MicroStation is not running or cannot be attached.
    ''' </returns>
    ''' <exception cref="InvalidOperationException">Thrown when MicroStation is not running, cannot be attached, or the COM object cannot be cast.</exception>
    Public Function Attach(Optional visible As Boolean = True) As BCOM.Application
        ' Declare 'app' here so it's accessible throughout the function.
        Dim app As BCOM.Application = Nothing
        Try
            Dim obj = Marshal.GetActiveObject(ProgId)
            app = CType(obj, BCOM.Application)
            app.Visible = visible
            Return app ' Return the successfully attached application object.
        Catch ex As COMException
            ' Could not find a running instance of MicroStation
            Throw New InvalidOperationException("MicroStation is not running or cannot be attached.", ex)
        Catch ex As InvalidCastException
            ' The object retrieved is not of the expected type
            Throw New InvalidOperationException("Failed to cast the COM object to MicroStationDGN.Application.", ex)
        Catch ex As Exception
            ' General error
            Throw New InvalidOperationException("An unexpected error occurred while attaching to MicroStation.", ex)
        End Try
        ' The 'Finally' block was removed. The caller is now responsible 
        ' for releasing the COM object using one of the release methods below.
    End Function
    ''' <summary>
    ''' Releases the COM object for <see cref="MicroStationDGN.Application"/>.
    ''' </summary>
    ''' <param name="app">The MicroStationDGN.Application instance to release.</param>
    ''' <remarks>
    ''' Use this method to release COM resources when you are done with the MicroStation application object.
    ''' </remarks>
    Public Sub ReleaseComObject(app As BCOM.Application)
        If app IsNot Nothing Then
            Try
                Marshal.ReleaseComObject(app)
            Catch ex As Exception
                ' Log or handle cleanup error if needed
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Forces final release of the COM object for <see cref="MicroStationDGN.Application"/>.
    ''' </summary>
    ''' <param name="app">The MicroStationDGN.Application instance to release.</param>
    ''' <remarks>
    ''' Use this method to ensure all references to the COM object are released.
    ''' </remarks>
    Public Sub FinalReleaseComObject(app As BCOM.Application)
        If app IsNot Nothing Then
            Try
                Marshal.FinalReleaseComObject(app)
            Catch ex As Exception
                ' Log or handle cleanup error if needed
            End Try
        End If
    End Sub
End Module