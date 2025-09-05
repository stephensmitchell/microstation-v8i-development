Imports System
Imports System.IO
Imports MicroStationDGN
Public Class ExportTool
    Private ustation As Application
    Public Sub New(ustation As Application)
        Me.ustation = ustation
    End Sub
    Public Sub Run()
        Dim outputDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim formats() As String = {"stl", "step", "sat", "obj"}
        For Each format As String In formats
            Dim filePath As String = Path.Combine(outputDir, "export." & format)
            Dim keyin As String
            Select Case format
                Case "stl"
                    keyin = "export stl """ & filePath & """ > NUL"
                Case "step"
                    keyin = "export step """ & filePath & """ > NUL"
                Case "sat"
                    keyin = "export acis """ & filePath & """ > NUL"
                Case "obj"
                    keyin = "export obj """ & filePath & """ > NUL"
            End Select
            Console.WriteLine("Exporting to " & filePath)
            ustation.CadInputQueue.SendCommand(keyin)
            If File.Exists(filePath) Then
                Console.WriteLine("Export successful.")
            Else
                Console.WriteLine("Export failed.")
            End If
        Next
        Console.WriteLine("Finished exporting all formats.")
    End Sub
End Class
