Option Explicit

Dim TaskName
Dim ProjectPath
Dim BuildConfig

Dim WScriptShell
Dim FSO

TaskName = "BUILD"
ProjectPath = GetLocalProjectPath()

Set WScriptShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

If IsMainWorkbookOpen(ProjectPath) Then
    Call TaskNotification(TaskName, "The main workbook is open. Please close it before running this task.")
    Call WScript.Quit(-1)
End If

Set BuildConfig = LoadBuildConfig(FSO.BuildPath(ProjectPath, "Build.xml"))

Call CreateMainWorkbook(ProjectPath, BuildConfig)
Call CreateExecutionScript(ProjectPath, BuildConfig)
Call TaskSuccessNotification(TaskName)
