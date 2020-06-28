**Error Handling for Excel VBA Code**

I have an error handling scheme I’ve devised over years of writing VBA code for Excel.

**The structure:**

- Code module called `Error_WarningRoutines`
- Class module called `MessageFileClass`
- I have a module I include in every application called `CommonRoutines`. This error handling scheme uses one routine in `CommonRoutines` called `DesktopFolder`. `DesktopFolder` returns a String containing the path to the user’s Desktop folder.

**The overall design concept:**

- Include `ErrorWarning_Routines`, `MessageFileClass`, and `DesktopFolder` in your VBA code
- Use the `MainProgram` template for your top-level routines
- Use the `SubTemplate` as the pattern for your routines
- When your code raises an exception, execution jumps to `ErrorHandler`. The error handler code calls `ReportError` in `Error_WarningRoutines`.
- ReportError` starts an error file if this is the first error. More on error files later.
- ReportError` writes an error message to the error file.
- Execution returns to the error handler code in your routine.
- The error handler calls `RaiseError` in `Error_WarningRoutines`
-  `RaiseError` raises the error to the next higher routine in the call stack
- The error is eventually raised to the top-level routine (`MainProgram` template)
- The `MainProgram` error handler calls `CloseError` in `Error_WarningRoutines`
- `CloseError` closes the error file by setting the file’s class instance to Nothing. This triggers the `Class_Terminate` routine in `MessageFileClass`.
- `Class_Terminate` closes the file and alerts the user that they should look on their desktop for an Error Messages folder where the user will find the error file

**What the user sees:**

- When the program raises an error, the user sees a `MsgBox` alerting them to look in the Error Messages folder on their desktop
- In the Error Messages folder, the user sees .txt error file(s)
- The user can open the most recent .txt error file to see the error messages

**Options:**

- There is a `ReportError` and a `ReportWarning` routine in `Error_WarningRoutines`. You, as the programmer, have an option to use either routine depending on the severity of the error. My practice is to put detailed “what happened in the code” messages in the error files for the programmer. I put application-oriented messages in the messages in the warning files. As the programmer I lean on the error files. The user should rely on the warning files. You, as the programmer, must make that distinction in your code.
- `MessageFileClass` is a general-purpose class you can use any time you need to write to a .txt file