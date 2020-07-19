Private Sub SubTemplate()

    ' Used for lower level routines
    
    Const RoutineName As String = Module_Name & "SubTemplate"
    On Error GoTo ErrorHandler
    
    ' Code goes here
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' SubTemplate
