Attribute VB_Name = "ChuckNorrisJokes"
Sub GetJoke()
    Dim json As String
    Dim sc As Object
    Const jokeurl As String = "https://api.icndb.com/jokes/random?exclude=[nerdy,explicit]"
    
    Dim request As Object
    Dim result As String
    
On Error GoTo ErrorJob
    
    Set sc = CreateObject("scriptcontrol")
    sc.Language = "JScript"
    
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    With request
        .Open "GET", jokeurl, True
        .SetRequestHeader "Content-Type", "text/json; charset=UTF-8"
        .Send
        .WaitForResponse
        json = .ResponseText
    End With
        
    sc.Eval "var obj=(" & json & ")"
    sc.AddCode "function parse(){return obj.value.joke}"
    
    num = sc.Run("parse")
    
    MsgBox num

Exit_DoSomeJob:
    On Error Resume Next
    Set request = Nothing
    Exit Sub
    
ErrorJob:
    MsgBox Err.Description, vbExclamation, Err.Number
    Resume Exit_DoSomeJob

End Sub

