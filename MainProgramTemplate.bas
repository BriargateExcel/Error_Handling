Public Sub MainProgram()

    ' Used as the top level routine
    
    Const RoutineName As String = Module_Name & "MainProgram"
    On Error GoTo ErrorHandler
    
    ' Code goes here
    
Done:
    MsgBox "Normal exit", vbOKOnly
    GoTo Done2
Halted:
    ' Use the Halted exit point after giving the user a message
    '   describing why processing did not run to completion
    MsgBox "Abnormal exit"
Done2:
    CloseErrorFile
    TurnOnAutomaticProcessing
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    TurnOnAutomaticProcessing
    CloseErrorFile
End Sub ' MainProgram
