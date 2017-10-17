Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Const defaultIndent = 7

Public Function defaultSnippet() As String
Dim s As String
s = ""
s = s & "<!DOCTYPE html>" & vbCrLf
s = s & "<html lang=""en"">" & vbCrLf
s = s & "<head>" & vbCrLf
s = s & Space(defaultIndent) & "<title>Bootstrap Example</title>" & vbCrLf
s = s & Space(defaultIndent) & "<meta charset=""utf-8"">" & vbCrLf
s = s & Space(defaultIndent) & "<meta name=""viewport"" content=""width=device-width, initial-scale=1"">" & vbCrLf
s = s & Space(defaultIndent) & "<link rel=""stylesheet"" href=""https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css"">" & vbCrLf
s = s & Space(defaultIndent) & "<script src=""https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js""></script>" & vbCrLf
s = s & Space(defaultIndent) & "<script src=""https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js""></script>" & vbCrLf
s = s & "</head>" & vbCrLf
s = s & "<body>" & vbCrLf
s = s & vbCrLf
s = s & vbCrLf
s = s & "</body>" & vbCrLf
s = s & "</html>" & vbCrLf
defaultSnippet = s
End Function

Public Function containerSnippet() As String
Dim s As String
s = ""
s = s & "<div class=""container"">" & vbCrLf
s = s & vbCrLf
s = s & "</div>" & vbCrLf
containerSnippet = s
End Function

Public Function Uglify(s As String) As String
s = Replace(s, vbCrLf, "")
Dim k As Integer, m As Integer, s2 As String
Do While InStr(s, ">") <> 0
s2 = s2 & Trim(Left(s, InStr(s, ">")))
s = Mid(s, InStr(s, ">") + 1)
Loop
Uglify = s2
End Function

Public Function InsertStr(s As String, substr As String, pos As Integer) As String
InsertStr = Left(s, pos) & substr & Mid(s, pos + 1)
End Function
Public Function CountStr(ByVal s As String, substr As String)
Dim k As Integer
k = 0
Do While InStr(s, substr) <> 0
s = Mid(s, InStr(s, substr) + 1)
k = k + 1
Loop
CountStr = k
End Function

Public Function Beautify(ByVal s As String, mode As Integer) As String
Dim s2 As String, s3 As String
'Add verify html
s = Uglify(s)
If mode = 1 Then
Do While InStr(s, "</") > 0
    s2 = Left(s, InStr(InStr(s, "</"), s, ">"))
        Do While CountStr(s2, ">") > 2
            s3 = s3 & Mid(s2, 1, InStr(s2, ">")) & vbCrLf
            s2 = Mid(s2, InStr(s2, ">") + 1)
        Loop
        s3 = s3 & s2 & vbCrLf
    s = Mid(s, InStr(InStr(s, "</"), s, ">") + 1)
Loop
End If

If mode = 2 Then
Do While InStr(s, "</") > 0
    s2 = Left(s, InStr(InStr(s, "</"), s, ">"))
        Do While CountStr(s2, ">") > 1
            s3 = s3 & Mid(s2, 1, InStr(s2, ">")) & vbCrLf
            s2 = Mid(s2, InStr(s2, ">") + 1)
        Loop
        s3 = s3 & Space(defaultIndent) & Left(s2, InStr(s2, "</") - 1) & vbCrLf
        s2 = Mid(s2, InStr(s2, "</")) & vbCrLf
        s3 = s3 & s2
    s = Mid(s, InStr(InStr(s, "</"), s, ">") + 1)
Loop
End If

Dim LastIndent As Long, LoopIndent As Long, FirstIndent As Long, s4 As String
If mode = 3 Then
LastIndent = 0: LoopIndent = 0
s4 = Beautify(Left(s, InStr(s, "</head>") + 10), 1)
s = Mid(s, InStr(s, "<body>"))
FirstIndent = CountStr(Left(s, InStr(InStr(s, "</"), s, ">")), ">") - 1
Do While InStr(s, "</") > 0
    s2 = Left(s, InStr(InStr(s, "</"), s, ">"))
    LoopIndent = CountStr(s2, ">")
    If LoopIndent > LastIndent Then LastIndent = LoopIndent
        Do While LoopIndent > 2
            s3 = s3 & Space(IIf((LastIndent - 2) > 0, defaultIndent * (LastIndent - FirstIndent), 0)) & Mid(s2, 1, InStr(s2, ">")) & vbCrLf
            s2 = Mid(s2, InStr(s2, ">") + 1)
            LastIndent = LastIndent + 1
            LoopIndent = LoopIndent - 1
        Loop
        If LoopIndent = 2 Then
            s3 = s3 & Space(IIf(LastIndent > 2, defaultIndent * (LastIndent - FirstIndent), 0)) & s2 & vbCrLf
            s2 = ""
        End If
        If LoopIndent = 1 Then
            LastIndent = LastIndent - 1
            s3 = s3 & Space(IIf(LastIndent > 2, defaultIndent * (LastIndent - FirstIndent), 0)) & s2 & vbCrLf
            s2 = ""
        End If
    s = Mid(s, InStr(InStr(s, "</"), s, ">") + 1)
Loop
s3 = s4 & s3
End If


Beautify = s3
End Function








