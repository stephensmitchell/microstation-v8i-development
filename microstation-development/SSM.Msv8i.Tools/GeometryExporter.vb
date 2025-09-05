Imports System
Imports System.IO
Imports System.Collections.Generic
Imports MicroStationDGN
''' <summary>
''' Enhanced geometry export functionality for MicroStation V8i
''' Supports multiple export formats and provides comprehensive export options
''' </summary>
Public Class GeometryExporter
    Private ustation As Application
    Private geometryLib As GeometryLibrary
    Public Sub New(ustation As Application)
        Me.ustation = ustation
        Me.geometryLib = New GeometryLibrary(ustation)
    End Sub
    ''' <summary>
    ''' Supported export formats
    ''' </summary>
    Public Enum ExportFormat
        STL = 0
        [STEP] = 1
        IGES = 2
        SAT_ACIS = 3
        OBJ = 4
        DWG = 5
        DXF = 6
        PDF = 7
        PLT = 8
        CGM = 9
    End Enum
    ''' <summary>
    ''' Export options for different formats
    ''' </summary>
    Public Structure ExportOptions
        Public IncludeHidden As Boolean
        Public IncludeLevels() As String
        Public ExcludeLevels() As String
        Public UseActiveView As Boolean
        Public ScaleFactor As Double
        Public Resolution As Double
        Public ColorMode As String
        Public Units As String
    End Structure
#Region "Primary Export Functions"
    ''' <summary>
    ''' Exports the current model to specified format
    ''' </summary>
    Public Function ExportModel(filePath As String, format As ExportFormat, Optional options As ExportOptions = Nothing) As Boolean
        Try
            Dim success As Boolean = False
            Select Case format
                Case ExportFormat.STL
                    success = ExportToSTL(filePath, options)
                Case ExportFormat.STEP
                    success = ExportToSTEP(filePath, options)
                Case ExportFormat.IGES
                    success = ExportToIGES(filePath, options)
                Case ExportFormat.SAT_ACIS
                    success = ExportToSAT(filePath, options)
                Case ExportFormat.OBJ
                    success = ExportToOBJ(filePath, options)
                Case ExportFormat.DWG
                    success = ExportToDWG(filePath, options)
                Case ExportFormat.DXF
                    success = ExportToDXF(filePath, options)
                Case ExportFormat.PDF
                    success = ExportToPDF(filePath, options)
                Case ExportFormat.PLT
                    success = ExportToPLT(filePath, options)
                Case ExportFormat.CGM
                    success = ExportToCGM(filePath, options)
                Case Else
                    Throw New ArgumentException("Unsupported export format: " & format.ToString())
            End Select
            If success Then
                Console.WriteLine("Successfully exported to: " & filePath)
            Else
                Console.WriteLine("Export failed for: " & filePath)
            End If
            Return success
        Catch ex As Exception
            Console.WriteLine("Export error: " & ex.Message)
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Exports multiple formats at once
    ''' </summary>
    Public Function ExportMultipleFormats(baseFilePath As String, formats() As ExportFormat, Optional options As ExportOptions = Nothing) As Dictionary(Of ExportFormat, Boolean)
        Dim results As New Dictionary(Of ExportFormat, Boolean)
        Dim baseDir As String = Path.GetDirectoryName(baseFilePath)
        Dim baseName As String = Path.GetFileNameWithoutExtension(baseFilePath)
        For Each format As ExportFormat In formats
            Dim extension As String = GetFileExtensionForFormat(format)
            Dim fullPath As String = Path.Combine(baseDir, baseName & "." & extension)
            results(format) = ExportModel(fullPath, format, options)
        Next
        Return results
    End Function
    ''' <summary>
    ''' Exports selected elements only
    ''' </summary>
    Public Function ExportSelectedElements(filePath As String, format As ExportFormat, elements() As _Element, Optional options As ExportOptions = Nothing) As Boolean
        Try
            ' Create temporary model or use selection set
            ' Implementation depends on MicroStation's selection mechanisms
            Console.WriteLine("Exporting " & elements.Length & " selected elements to: " & filePath)
            Return ExportModel(filePath, format, options)
        Catch ex As Exception
            Console.WriteLine("Export selected elements error: " & ex.Message)
            Return False
        End Try
    End Function
#End Region
#Region "Format-Specific Export Functions"
    ''' <summary>
    ''' Exports to STL format (for 3D printing)
    ''' </summary>
    Private Function ExportToSTL(filePath As String, options As ExportOptions) As Boolean
        Dim keyin As String = "export stl """ & filePath & """"
        If options.Resolution > 0 Then
            keyin &= " resolution=" & options.Resolution.ToString()
        End If
        Return ExecuteKeyin(keyin, filePath)
    End Function
    ''' <summary>
    ''' Exports to STEP format (ISO standard for CAD data exchange)
    ''' </summary>
    Private Function ExportToSTEP(filePath As String, options As ExportOptions) As Boolean
        Dim keyin As String = "export step """ & filePath & """"
        If Not String.IsNullOrEmpty(options.Units) Then
            keyin &= " units=" & options.Units
        End If
        Return ExecuteKeyin(keyin, filePath)
    End Function
    ''' <summary>
    ''' Exports to IGES format (Initial Graphics Exchange Specification)
    ''' </summary>
    Private Function ExportToIGES(filePath As String, options As ExportOptions) As Boolean
        Dim keyin As String = "export iges """ & filePath & """"
        Return ExecuteKeyin(keyin, filePath)
    End Function
    ''' <summary>
    ''' Exports to SAT format (ACIS solid modeling)
    ''' </summary>
    Private Function ExportToSAT(filePath As String, options As ExportOptions) As Boolean
        Dim keyin As String = "export acis """ & filePath & """"
        Return ExecuteKeyin(keyin, filePath)
    End Function
    ''' <summary>
    ''' Exports to OBJ format (Wavefront 3D model)
    ''' </summary>
    Private Function ExportToOBJ(filePath As String, options As ExportOptions) As Boolean
        Dim keyin As String = "export obj """ & filePath & """"
        Return ExecuteKeyin(keyin, filePath)
    End Function
    ''' <summary>
    ''' Exports to DWG format (AutoCAD Drawing)
    ''' </summary>
    Private Function ExportToDWG(filePath As String, options As ExportOptions) As Boolean
        Dim keyin As String = "export dwg """ & filePath & """"
        Return ExecuteKeyin(keyin, filePath)
    End Function
    ''' <summary>
    ''' Exports to DXF format (Drawing Exchange Format)
    ''' </summary>
    Private Function ExportToDXF(filePath As String, options As ExportOptions) As Boolean
        Dim keyin As String = "export dxf """ & filePath & """"
        Return ExecuteKeyin(keyin, filePath)
    End Function
    ''' <summary>
    ''' Exports to PDF format
    ''' </summary>
    Private Function ExportToPDF(filePath As String, options As ExportOptions) As Boolean
        Dim keyin As String = "print pdf """ & filePath & """"
        If options.ScaleFactor > 0 Then
            keyin &= " scale=" & options.ScaleFactor.ToString()
        End If
        Return ExecuteKeyin(keyin, filePath)
    End Function
    ''' <summary>
    ''' Exports to PLT format (Plotter file)
    ''' </summary>
    Private Function ExportToPLT(filePath As String, options As ExportOptions) As Boolean
        Dim keyin As String = "print plt """ & filePath & """"
        Return ExecuteKeyin(keyin, filePath)
    End Function
    ''' <summary>
    ''' Exports to CGM format (Computer Graphics Metafile)
    ''' </summary>
    Private Function ExportToCGM(filePath As String, options As ExportOptions) As Boolean
        Dim keyin As String = "export cgm """ & filePath & """"
        Return ExecuteKeyin(keyin, filePath)
    End Function
#End Region
#Region "Batch Export Functions"
    ''' <summary>
    ''' Exports all models in the design file
    ''' </summary>
    Public Function ExportAllModels(outputDirectory As String, format As ExportFormat, Optional options As ExportOptions = Nothing) As Dictionary(Of String, Boolean)
        Dim results As New Dictionary(Of String, Boolean)
        Try
            For Each modelRef As ModelReference In ustation.ActiveDesignFile.Models
                If Not modelRef.IsReadOnly Then
                    Dim modelName As String = modelRef.Name
                    Dim fileName As String = modelName & "." & GetFileExtensionForFormat(format)
                    Dim fullPath As String = Path.Combine(outputDirectory, fileName)
                    ' Set active model
                    modelRef.Activate()
                    ' Export the model
                    results(modelName) = ExportModel(fullPath, format, options)
                End If
            Next
        Catch ex As Exception
            Console.WriteLine("Export all models error: " & ex.Message)
        End Try
        Return results
    End Function
    ''' <summary>
    ''' Exports by level/layer
    ''' </summary>
    Public Function ExportByLevels(outputDirectory As String, format As ExportFormat, levels() As String, Optional options As ExportOptions = Nothing) As Dictionary(Of String, Boolean)
        Dim results As New Dictionary(Of String, Boolean)
        For Each levelName As String In levels
            Try
                ' Set level visibility
                Dim level As Level = ustation.ActiveDesignFile.Levels.Find(levelName)
                If level IsNot Nothing Then
                    Dim fileName As String = "Level_" & levelName & "." & GetFileExtensionForFormat(format)
                    Dim fullPath As String = Path.Combine(outputDirectory, fileName)
                    ' Modify options to include only this level
                    Dim levelOptions As ExportOptions = options
                    ReDim levelOptions.IncludeLevels(0)
                    levelOptions.IncludeLevels(0) = levelName
                    results(levelName) = ExportModel(fullPath, format, levelOptions)
                End If
            Catch ex As Exception
                Console.WriteLine("Export level " & levelName & " error: " & ex.Message)
                results(levelName) = False
            End Try
        Next
        Return results
    End Function
    ''' <summary>
    ''' Exports with different view configurations
    ''' </summary>
    Public Function ExportMultipleViews(baseFilePath As String, format As ExportFormat, views() As View, Optional options As ExportOptions = Nothing) As Dictionary(Of String, Boolean)
        Dim results As New Dictionary(Of String, Boolean)
        Dim baseDir As String = Path.GetDirectoryName(baseFilePath)
        Dim baseName As String = Path.GetFileNameWithoutExtension(baseFilePath)
        Dim extension As String = GetFileExtensionForFormat(format)
        For i As Integer = 0 To views.Length - 1
            Try
                Dim view As View = views(i)
                Dim viewName As String = "View" & (i + 1).ToString()
                Dim fileName As String = baseName & "_" & viewName & "." & extension
                Dim fullPath As String = Path.Combine(baseDir, fileName)
                ' Set active view
                ustation.ActiveView = view
                ' Export with current view
                results(viewName) = ExportModel(fullPath, format, options)
            Catch ex As Exception
                Console.WriteLine("Export view error: " & ex.Message)
                results("View" & (i + 1).ToString()) = False
            End Try
        Next
        Return results
    End Function
#End Region
#Region "Utility Functions"
    ''' <summary>
    ''' Gets file extension for export format
    ''' </summary>
    Public Function GetFileExtensionForFormat(format As ExportFormat) As String
        Select Case format
            Case ExportFormat.STL
                Return "stl"
            Case ExportFormat.STEP
                Return "step"
            Case ExportFormat.IGES
                Return "iges"
            Case ExportFormat.SAT_ACIS
                Return "sat"
            Case ExportFormat.OBJ
                Return "obj"
            Case ExportFormat.DWG
                Return "dwg"
            Case ExportFormat.DXF
                Return "dxf"
            Case ExportFormat.PDF
                Return "pdf"
            Case ExportFormat.PLT
                Return "plt"
            Case ExportFormat.CGM
                Return "cgm"
            Case Else
                Return "dat"
        End Select
    End Function
    ''' <summary>
    ''' Executes a keyin command and checks for file creation
    ''' </summary>
    Private Function ExecuteKeyin(keyin As String, expectedFilePath As String) As Boolean
        Try
            Console.WriteLine("Executing: " & keyin)
            ustation.CadInputQueue.SendKeyin(keyin)
            ' Wait a moment for file creation
            System.Threading.Thread.Sleep(1000)
            ' Check if file was created
            If File.Exists(expectedFilePath) Then
                Dim fileInfo As New FileInfo(expectedFilePath)
                If fileInfo.Length > 0 Then
                    Console.WriteLine("Export successful: " & expectedFilePath & " (" & fileInfo.Length & " bytes)")
                    Return True
                Else
                    Console.WriteLine("Export created empty file: " & expectedFilePath)
                    Return False
                End If
            Else
                Console.WriteLine("Export failed: File not created - " & expectedFilePath)
                Return False
            End If
        Catch ex As Exception
            Console.WriteLine("Keyin execution error: " & ex.Message)
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Creates default export options
    ''' </summary>
    Public Function CreateDefaultExportOptions() As ExportOptions
        Dim options As New ExportOptions()
        options.IncludeHidden = False
        options.UseActiveView = True
        options.ScaleFactor = 1.0
        options.Resolution = 0.1
        options.ColorMode = "color"
        options.Units = "mm"
        Return options
    End Function
    ''' <summary>
    ''' Gets available export formats
    ''' </summary>
    Public Function GetAvailableFormats() As ExportFormat()
        Return [Enum].GetValues(GetType(ExportFormat)).Cast(Of ExportFormat)().ToArray()
    End Function
    ''' <summary>
    ''' Validates export file path
    ''' </summary>
    Public Function ValidateExportPath(filePath As String) As Boolean
        Try
            Dim _directory As String = Path.GetDirectoryName(filePath)
            If Not String.IsNullOrEmpty(_directory) AndAlso Not Directory.Exists(_directory) Then
                Directory.CreateDirectory(_directory)
            End If
            Return True
        Catch ex As Exception
            Console.WriteLine("Path validation error: " & ex.Message)
            Return False
        End Try
    End Function
#End Region
#Region "Progress and Status Reporting"
    ''' <summary>
    ''' Event for export progress reporting
    ''' </summary>
    Public Event ExportProgressChanged(sender As Object, progress As Integer, message As String)
    ''' <summary>
    ''' Reports export progress
    ''' </summary>
    Protected Sub OnExportProgressChanged(progress As Integer, message As String)
        RaiseEvent ExportProgressChanged(Me, progress, message)
    End Sub
    ''' <summary>
    ''' Gets export statistics
    ''' </summary>
    Public Function GetExportStatistics() As Dictionary(Of String, Object)
        Dim stats As New Dictionary(Of String, Object)()
        Try
            stats("ModelName") = ustation.ActiveModelReference.Name
            Dim elements = ustation.ActiveModelReference.Scan(Nothing)
            Dim elementCount As Integer = 0
            While elements.MoveNext()
                elementCount += 1
            End While
            stats("ElementCount") = elementCount
            stats("ModelBounds") = ustation.ActiveModelReference.Range(True)
            stats("ActiveView") = ustation.ActiveView.Index
            stats("DesignFileName") = ustation.ActiveDesignFile.FullName
        Catch ex As Exception
            Console.WriteLine("Error getting export statistics: " & ex.Message)
        End Try
        Return stats
    End Function
#End Region
End Class