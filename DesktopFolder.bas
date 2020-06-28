Attribute VB_Name = "DesktopFolder"
Option Explicit

Private Const Module_Name As String = "DesktopFolder."

Public Function DesktopFolder() As String

    ' This routine returns the full pathname to the Windows desktop folder

    Const RoutineName As String = Module_Name & "DesktopFolder"
    On Error GoTo ErrorHandler

    Dim objSFolders As Object
    Set objSFolders = CreateObject("WScript.Shell").specialfolders
    DesktopFolder = objSFolders("desktop")

Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' DesktopFolder
