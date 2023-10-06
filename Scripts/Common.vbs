Option Explicit

Sub TaskNotification( _
    TaskName As String, _
    Msg As String _
)

    Call WScript.StdOut.WriteLine(Msg)

End Sub

Sub TaskSuccessNotification( _
    TaskName As String _
)

    Call TaskNotification(TaskName, "Task " & TaskName & " completed successfully.")

End Sub

Sub ExecuteShell( _
    TaskName As String, _
    Command As String, _
    StdInputOutput As Boolean, _
    ContinueOnError As Boolean _
)

    Dim StdOutputText As String
    Dim StdErrorText As String
    Dim ExitCode As Long

    Dim ExecuteShell

    Set ExecuteShell = CreateObject("Scripting.Dictionary")

    With WScriptShell.Exec(Command)
        Do While .Status <> 1
            Call WScript.Sleep(100)
        Loop

        StdOutputText = .StdOut.ReadAll()
        StdErrorText = .StdErr.ReadAll()

        If IsStdInputOutput Then
            With ExecuteShell
                Call .Add("StandardOutput", StdOutputText)
                Call .Add("StandardError", StdErrorText)
            End With
        Else
            Call WScript.StdOut.Write(StdOutputText)
            Call WScript.StdErr.Write(StdErrorText)
        End If

        ExitCode = .ExitCode

        Call ExecuteShell.Add("ExitCode", ExitCode)

        If Not ContinueOnError And (ExitCode <> 0) Then
            With WScript.StdErr
                Call .WriteLine("The Child Process exited with status code: " & CStr(ExitCode))
                Call .WriteLine()
