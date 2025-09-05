Imports MicroStationDGN
Imports SSM.Msv8i.Functions
Module Program
    Sub Main(args As String())
        Dim ustation As Application = Nothing
        Try
            ustation = MsCom.Attach()
            Console.WriteLine("Successfully attached to MicroStation.")
            Dim tool As New PipeTools(ustation)
            tool.Run()
            Console.WriteLine("Tool finished successfully.")
        Catch ex As Exception
            Console.WriteLine($"An error occurred: {ex.Message}")
        Finally
            Console.WriteLine("Cleaning up COM object...")
            MsCom.ReleaseComObject(ustation)
        End Try
        Console.WriteLine("Application finished.")
    End Sub
End Module