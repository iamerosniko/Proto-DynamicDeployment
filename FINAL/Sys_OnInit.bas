Attribute VB_Name = "Sys_OnInit"
Option Compare Database

Sub ChangeTitle()
    Dim obj As Object
    Const conPropNotFoundError = 3270
    
    On Error GoTo ErrorHandler
    ' Return Database object variable pointing to
    ' the current database.
    Set dbs = CurrentDb
    ' Change title bar.
    dbs.Properties!AppTitle = "Contacts Databaseasdfasdf"
    ' Update title bar on screen.
    Application.RefreshTitleBar
    Exit Sub
 
ErrorHandler:
' If err.Number = conPropNotFoundError Then
' Set obj = dbs.CreateProperty("AppTitle", dbText, "Contacts Database")
' dbs.Properties.Append obj
' Else
' MsgBox "Error: " & err.Number & vbCrLf & err.Description
' End If
' Resume Next
    Exit Sub
End Sub

Sub sampleloc()
    Dim Str As String
    Str = Application.CurrentProject.path
    If InStr(LCase(Str), "dev\apps") > 1 Then
        MsgBox "under dev"
    ElseIf InStr(LCase(Str), "uat\apps") > 1 Then
        MsgBox "under stage"
    ElseIf InStr(LCase(Str), "ops\apps") > 1 Then
        MsgBox "under production"
    Else
        MsgBox "default: under dev"
    End If
End Sub
